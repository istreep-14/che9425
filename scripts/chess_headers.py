#!/usr/bin/env python3
import json
import sys
import time
import re
import os
from urllib.request import Request, urlopen
from urllib.error import URLError, HTTPError


LEADERBOARDS_URL = "https://api.chess.com/pub/leaderboards"
ERIK_SAMPLE_URL = "https://api.chess.com/pub/player/erik/games/2009/10"
PLAYER_ARCHIVES_TEMPLATE = "https://api.chess.com/pub/player/{username}/games/archives"


def fetch_url(url: str, timeout_seconds: float = 20.0, max_attempts: int = 3) -> bytes:
    last_error: Exception | None = None
    for attempt_index in range(1, max_attempts + 1):
        try:
            req = Request(url, headers={
                "User-Agent": "chess-headers-script/1.0 (contact: dev)"
            })
            with urlopen(req, timeout=timeout_seconds) as resp:
                return resp.read()
        except (URLError, HTTPError) as exc:
            last_error = exc
            time.sleep(min(2.0 * attempt_index, 5.0))
    if last_error is not None:
        raise last_error
    raise RuntimeError(f"Failed to fetch URL after {max_attempts} attempts: {url}")


def fetch_json(url: str) -> dict:
    raw = fetch_url(url)
    try:
        return json.loads(raw.decode("utf-8"))
    except json.JSONDecodeError as exc:
        raise RuntimeError(f"Invalid JSON from {url}: {exc}") from exc


def collect_usernames_from_leaderboards() -> set[str]:
    data = fetch_json(LEADERBOARDS_URL)
    usernames: set[str] = set()

    def walk(obj):
        if isinstance(obj, dict):
            # If this dictionary looks like a player entry, capture the username
            if "username" in obj and isinstance(obj["username"], str):
                usernames.add(obj["username"])  # Case as provided by API
            # Recurse
            for value in obj.values():
                walk(value)
        elif isinstance(obj, list):
            for item in obj:
                walk(item)

    walk(data)
    return usernames


def fetch_player_archives(user: str) -> list[str]:
    url = PLAYER_ARCHIVES_TEMPLATE.format(username=user)
    try:
        data = fetch_json(url)
    except HTTPError as exc:
        # Some users may be missing or closed; skip gracefully
        sys.stderr.write(f"Warning: cannot fetch archives for {user}: {exc}\n")
        return []
    archives = data.get("archives", [])
    # Filter to well-formed URLs only
    return [u for u in archives if isinstance(u, str) and u.startswith("http")]


def fetch_games_from_archive_url(archive_url: str) -> list[dict]:
    try:
        data = fetch_json(archive_url)
        games = data.get("games", [])
        if isinstance(games, list):
            return [g for g in games if isinstance(g, dict)]
        return []
    except Exception as exc:
        sys.stderr.write(f"Warning: failed to fetch archive {archive_url}: {exc}\n")
        return []


def flatten_json_keys(obj: dict, prefix: str = "") -> set[str]:
    keys: set[str] = set()
    for k, v in obj.items():
        if not isinstance(k, str):
            continue
        path = f"{prefix}.{k}" if prefix else k
        keys.add(path)
        if isinstance(v, dict):
            keys |= flatten_json_keys(v, path)
        # If value is a list of dicts with consistent shape, include subkeys
        elif isinstance(v, list) and v and all(isinstance(item, dict) for item in v):
            # Use numeric placeholder for arrays of objects
            path_array = f"{path}[]"
            keys.add(path_array)
            # Merge keys from first few items to avoid huge expansion
            for item in v[:3]:
                keys |= flatten_json_keys(item, path_array)
    return keys


PGN_TAG_PATTERN = re.compile(r"^\[(?P<tag>[A-Za-z0-9_\-]+)\s+\"", re.MULTILINE)


def extract_pgn_tags(pgn_text: str) -> set[str]:
    if not isinstance(pgn_text, str) or not pgn_text:
        return set()
    return set(m.group("tag") for m in PGN_TAG_PATTERN.finditer(pgn_text))


def ensure_dir(path: str) -> None:
    os.makedirs(path, exist_ok=True)


def main():
    # Configuration via env vars
    max_users_env = os.getenv("CHESS_HEADERS_MAX_USERS", "120")
    max_months_env = os.getenv("CHESS_HEADERS_MAX_MONTHS_PER_USER", "12")
    sleep_between_requests_env = os.getenv("CHESS_HEADERS_REQUEST_SLEEP_SECONDS", "0.15")
    include_erik_sample_env = os.getenv("CHESS_HEADERS_INCLUDE_ERIK_SAMPLE", "1")

    try:
        max_users = max(1, int(max_users_env))
    except ValueError:
        max_users = 120
    try:
        max_months_per_user = max(1, int(max_months_env))
    except ValueError:
        max_months_per_user = 12
    try:
        sleep_seconds = max(0.0, float(sleep_between_requests_env))
    except ValueError:
        sleep_seconds = 0.15
    include_erik_sample = include_erik_sample_env.strip() not in ("0", "false", "False")

    # Seed usernames from leaderboards
    usernames: set[str] = set()
    try:
        usernames |= collect_usernames_from_leaderboards()
    except Exception as exc:
        sys.stderr.write(f"Warning: failed to collect leaderboards usernames: {exc}\n")

    # Always include 'erik' per requirement
    usernames.add("erik")

    # Truncate to configured limit in a stable order
    usernames_list = sorted(usernames, key=lambda s: s.lower())[:max_users]
    sys.stderr.write(f"Info: collected {len(usernames_list)} usernames (limited to {max_users}).\n")

    # Optionally include the provided Erik month for maximum PGN variety across time
    json_key_union: set[str] = set()
    pgn_tag_union: set[str] = set()

    if include_erik_sample:
        erik_data = fetch_json(ERIK_SAMPLE_URL)
        games = erik_data.get("games", [])
        for game in games:
            if not isinstance(game, dict):
                continue
            json_key_union |= flatten_json_keys(game)
            pgn_tag_union |= extract_pgn_tags(game.get("pgn", ""))
        time.sleep(sleep_seconds)

    # Iterate users and a limited number of archive months per user
    for username in usernames_list:
        archives = fetch_player_archives(username)
        if not archives:
            continue
        # Prefer most recent months
        for archive_url in sorted(archives)[-max_months_per_user:]:
            games = fetch_games_from_archive_url(archive_url)
            for game in games:
                json_key_union |= flatten_json_keys(game)
                pgn_tag_union |= extract_pgn_tags(game.get("pgn", ""))
            time.sleep(sleep_seconds)

    # Prepare outputs
    json_keys_sorted = sorted(json_key_union)
    pgn_tags_sorted = sorted(pgn_tag_union)

    # Write to outputs
    ensure_dir("/workspace/outputs")
    result = {
        "json_game_keys": json_keys_sorted,
        "pgn_tag_keys": pgn_tags_sorted,
        "metadata": {
            "user_count": len(usernames_list),
            "max_months_per_user": max_months_per_user,
            "included_erik_sample": include_erik_sample,
        },
    }
    with open("/workspace/outputs/chess_headers.json", "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    # Also print a concise summary to stdout
    print("JSON keys (union):")
    for key in json_keys_sorted:
        print(key)
    print("")
    print("PGN tags (union):")
    for tag in pgn_tags_sorted:
        print(tag)


if __name__ == "__main__":
    main()

