/**
 * Calculated sheet pipeline: classify mode/variant, compute per-mode running ratings for a user,
 * and enrich with opponent stats using Chess.com public API.
 */

const CALC_SETTINGS = {
  userPropertyKey: 'PRIMARY_USERNAME',
  cacheMinutes: 60
};

function setPrimaryUsername(username) {
  if (!username) throw new Error('Username required');
  PropertiesService.getDocumentProperties().setProperty(CALC_SETTINGS.userPropertyKey, String(username));
}

function getPrimaryUsername_() {
  return PropertiesService.getDocumentProperties().getProperty(CALC_SETTINGS.userPropertyKey) || '';
}

function classifyModeAndVariant_(rules, timeClass) {
  const rulesNorm = String(rules || 'chess').toLowerCase();
  const timeNorm = String(timeClass || '').toLowerCase();
  // Mode: bullet/blitz/rapid/daily when rules == chess; otherwise variant name
  // Chess960 also has daily960 as a separate combined label
  if (rulesNorm === 'chess') {
    if (['bullet','blitz','rapid','daily'].includes(timeNorm)) return { mode: timeNorm, variant: 'standard' };
    return { mode: 'unknown', variant: 'standard' };
  }
  if (rulesNorm === 'chess960') {
    if (timeNorm === 'daily') return { mode: 'daily960', variant: 'chess960' };
    if (['bullet','blitz','rapid'].includes(timeNorm)) return { mode: 'chess960', variant: 'chess960' };
    return { mode: 'chess960', variant: 'chess960' };
  }
  // For other variants, mode is simply the variant name
  return { mode: rulesNorm, variant: rulesNorm };
}

function resultToScore_(winnerColor, perspectiveColor) {
  if (!winnerColor) return 0.5; // draw or unknown
  if (winnerColor === 'draw') return 0.5;
  return winnerColor === perspectiveColor ? 1 : 0;
}

function winnerColorFromResult_(resultRaw) {
  const r = String(resultRaw || '').trim();
  if (r === '1-0') return 'white';
  if (r === '0-1') return 'black';
  if (r === '1/2-1/2') return 'draw';
  return '';
}

function fetchOpponentStats_(username) {
  if (!username) return null;
  const cache = CacheService.getDocumentCache();
  const key = `oppstats:${username.toLowerCase()}`;
  const cached = cache.get(key);
  if (cached) {
    try { return JSON.parse(cached); } catch (e) {}
  }
  const url = `https://api.chess.com/pub/player/${encodeURIComponent(username)}/stats`;
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, headers: { 'Accept-Encoding': 'gzip' } });
  if (res.getResponseCode() !== 200) return null;
  const data = JSON.parse(res.getContentText());
  cache.put(key, JSON.stringify(data), CALC_SETTINGS.cacheMinutes * 60);
  return data;
}

function extractOpponentSummary_(stats) {
  if (!stats || typeof stats !== 'object') return {};
  function getRating(path) {
    const obj = path.reduce((acc, k) => (acc && acc[k] ? acc[k] : null), stats);
    if (!obj) return '';
    // prefer last.rating else best.rating else highest.rating
    return (obj.last && obj.last.rating) || (obj.best && obj.best.rating) || (obj.highest && obj.highest.rating) || '';
  }
  return {
    opp_bullet_rating: getRating(['chess_bullet']),
    opp_blitz_rating: getRating(['chess_blitz']),
    opp_rapid_rating: getRating(['chess_rapid']),
    opp_daily_rating: getRating(['chess_daily']),
    opp_daily960_rating: getRating(['chess_daily960']) || '',
    opp_chess960_rating: getRating(['chess960']) || ''
  };
}

function buildCalculatedFromRaw() {
  const username = getPrimaryUsername_();
  if (!username) throw new Error('Set your username with setPrimaryUsername("yourname") first.');
  const raw = getSheetByNameOrThrow_(HEADER_CONFIG.rawDataSheetName);
  const calc = getSheetByNameOrThrow_(HEADER_CONFIG.calcDataSheetName);

  const rawHeaders = readHeaderRow_(raw);
  if (!rawHeaders.length) throw new Error('Raw sheet has no headers.');

  const idx = new Map(rawHeaders.map((h, i) => [h, i]));

  // Define calculated headers
  const calcHeaders = [
    'url', 'date', 'time', 'white.username', 'black.username', 'white.rating', 'black.rating', 'result', 'termination',
    'rules', 'time_class', 'mode', 'variant',
    'my_color', 'my_score', 'my_result',
    'my_bullet_rating', 'my_blitz_rating', 'my_rapid_rating', 'my_daily_rating', 'my_chess960_rating', 'my_daily960_rating',
    'opp_bullet_rating', 'opp_blitz_rating', 'opp_rapid_rating', 'opp_daily_rating', 'opp_chess960_rating', 'opp_daily960_rating'
  ];

  // Initialize output
  calc.clear();
  calc.getRange(1, 1, 1, calcHeaders.length).setValues([calcHeaders]);
  calc.setFrozenRows(1);

  const lastRow = raw.getLastRow();
  if (lastRow <= 1) return;
  const values = raw.getRange(2, 1, lastRow - 1, raw.getLastColumn()).getValues();

  let running = {
    bullet: null,
    blitz: null,
    rapid: null,
    daily: null,
    chess960: null,
    daily960: null
  };

  const out = [];
  for (const row of values) {
    function val(h) { const i = idx.get(h); return i != null ? row[i] : ''; }
    const whiteUser = String(val('white.username') || val('white') || '').trim();
    const blackUser = String(val('black.username') || val('black') || '').trim();
    const myColor = whiteUser.toLowerCase() === username.toLowerCase() ? 'white' : (blackUser.toLowerCase() === username.toLowerCase() ? 'black' : '');
    const rules = val('rules');
    const timeClass = val('time_class') || val('time-class') || '';
    const { mode, variant } = classifyModeAndVariant_(rules, timeClass);

    const resultRaw = val('result') || val('Result') || '';
    const winner = winnerColorFromResult_(resultRaw);
    const myScore = myColor ? resultToScore_(winner, myColor) : '';
    const myResult = myScore === '' ? '' : (myScore === 1 ? 'Win' : (myScore === 0.5 ? 'Draw' : 'Loss'));

    // Ratings
    const whiteRating = Number(val('white.rating') || val('WhiteElo') || '') || '';
    const blackRating = Number(val('black.rating') || val('BlackElo') || '') || '';

    // Update running ratings only for the mode being played; others carry forward previous value
    const modes = ['bullet','blitz','rapid','daily','chess960','daily960'];
    const runningNext = { ...running };
    let playedRating = '';
    if (myColor === 'white') playedRating = whiteRating; else if (myColor === 'black') playedRating = blackRating;
    if (modes.includes(mode)) {
      runningNext[mode] = playedRating !== '' ? Number(playedRating) : running[mode];
    }
    running = runningNext;

    // Opponent stats (cached per username)
    const oppUser = myColor === 'white' ? blackUser : (myColor === 'black' ? whiteUser : '');
    const oppStats = oppUser ? extractOpponentSummary_(fetchOpponentStats_(oppUser)) : {};

    out.push([
      val('url') || '',
      val('UTCDate') || val('Date') || '',
      val('UTCTime') || '',
      whiteUser || '',
      blackUser || '',
      whiteRating,
      blackRating,
      resultRaw || '',
      val('termination') || val('Termination') || '',
      rules || '',
      timeClass || '',
      mode,
      variant,
      myColor || '',
      myScore,
      myResult,
      running.bullet,
      running.blitz,
      running.rapid,
      running.daily,
      running.chess960,
      running.daily960,
      oppStats.opp_bullet_rating || '',
      oppStats.opp_blitz_rating || '',
      oppStats.opp_rapid_rating || '',
      oppStats.opp_daily_rating || '',
      oppStats.opp_chess960_rating || '',
      oppStats.opp_daily960_rating || ''
    ]);
  }

  if (out.length) calc.getRange(2, 1, out.length, calcHeaders.length).setValues(out);
}

