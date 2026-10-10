"""
Get-LoLSkins.py

Lists all League of Legends skins you own, grouped by champion, as one
self-contained HTML file.

REQUIREMENTS
    - Windows, with the League Client (LeagueClientUx.exe) running and logged in
    - Python 3.8+, stdlib only (no pip installs needed)
    - A Riot developer API key (required, command-line only - see USAGE).
      Get one free from https://developer.riotgames.com/ (log in, copy the
      "Development API Key" on your dashboard). It expires ~24h after being
      generated, so you'll need a fresh one each day you run this.

HOW IT WORKS
    Riot's public API does NOT expose skin ownership for any account - that
    data only exists locally, through the League Client's own local API (the
    "LCU"), while the client is open and you're logged in. This script finds
    the running client, reads its port + auth token off its own process
    command line, pulls your owned skins over that local API, and writes one
    HTML file with thumbnails embedded as base64 (so it still opens fine
    after you close the client). Play Rate/Win Rate come from Riot's
    official match-v5 API instead (see --riot-api-key below), since the
    LCU's own local match history only goes back ~200 games.

USAGE
    --riot-api-key is REQUIRED on every run (no config file, no env var -
    it must be passed on the command line each time):
      python Get-LoLSkins.py --riot-api-key RGAPI-xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
    Get a key (free, valid ~24h) from https://developer.riotgames.com/ -
    log in, then copy the "Development API Key" on your dashboard.

    python Get-LoLSkins.py --riot-api-key RGAPI-xxx --include-base       # also list free/default skins
    python Get-LoLSkins.py --riot-api-key RGAPI-xxx --shards-only        # only champs with an available skin shard
    python Get-LoLSkins.py --riot-api-key RGAPI-xxx --out C:\\temp\\skins.html
    python Get-LoLSkins.py --riot-api-key RGAPI-xxx --no-open
    python Get-LoLSkins.py --riot-api-key RGAPI-xxx --cache-dir C:\\temp\\lol-match-cache  # custom cache location

KNOWN WEAKNESSES (read before assuming this keeps working forever)
    - LCU endpoints are undocumented/unofficial. Riot can change them without
      notice on any patch. If this breaks, check the endpoint paths / field
      names in fetch_champions() and fetch_skins() first.
    - Only reads whichever account is currently logged into the client. No
      client open = no data. Not a way to check an account you're not
      signed into.
    - Disables TLS certificate validation for requests to 127.0.0.1 only
      (the LCU uses a self-signed cert) - scoped to this script's own HTTP
      calls, not a system-wide change.
    - Chroma ownership is counted per-skin but not broken into its own
      section.
    - Skin shards are read from the loot API (type "CHAMPION_SKIN_RENTAL"),
      which is just as undocumented/unofficial as the rest of the LCU - if
      shard counts look wrong, check that type string first.
    - Play Rate and Win Rate normally come from Riot's official match-v5
      API (fetch_full_match_history()), which needs --riot-api-key (see
      USAGE/REQUIREMENTS) and goes back as far as Riot's backend has
      match-v5 data indexed for that account - NOT guaranteed to be every
      game ever played. Results are cached on disk per-account (by puuid,
      under --cache-dir) so repeat runs only fetch new matches - delete
      the account's cache file to force a full re-fetch. A free
      "Development API Key" is rate-limited (roughly 20 req/sec, 100
      req/2min, throttled here conservatively) and expires ~24h after
      being generated, so a first run on an account with thousands of
      games can take a long time and may need the key regenerated and the
      run repeated; the platform->continent routing map
      (PLATFORM_TO_CONTINENT) is also something Riot could change without
      notice. If the key is rejected or a request fails outright at
      runtime, this silently falls back to the local client's own
      match-history API (fetch_match_history()) instead, which only goes
      back ~200 games (a hard ceiling on what the LCU caches locally, not
      something pagination can get past) - so a surprisingly low Play
      Rate denominator usually means that fallback kicked in, not that
      your real history is actually that short. Win Rate in that fallback
      path additionally depends on a "win" field being present on the
      local player's participant entry (checked both flattened and nested
      under "stats") - if a client version drops that field, wins will
      undercount to zero without any other symptom.
"""

import argparse
import base64
import json
import re
import ssl
import subprocess
import sys
import time
import webbrowser
from collections import Counter
from datetime import datetime
from pathlib import Path
from urllib.parse import quote, urlencode
from urllib.request import Request, urlopen
from urllib.error import URLError, HTTPError

RIOT_DEV_PORTAL_URL = "https://developer.riotgames.com/"


def find_client():
    """Return (port, token, region) parsed from the running
    LeagueClientUx.exe command line. region is the client's own
    "--region=" launch flag (e.g. "NA1") if present, else None - it's the
    most reliable source for a real platform id, used for Riot API
    continent routing.
    """
    try:
        raw = subprocess.check_output(
            [
                "powershell", "-NoProfile", "-Command",
                "Get-CimInstance Win32_Process -Filter \"Name='LeagueClientUx.exe'\" "
                "| Select-Object -ExpandProperty CommandLine",
            ],
            stderr=subprocess.DEVNULL,
        ).decode("utf-8", errors="ignore")
    except (subprocess.CalledProcessError, FileNotFoundError):
        raw = ""

    if not raw.strip():
        sys.exit("League Client isn't running. Open it and log in, then re-run this script.")

    port_match = re.search(r"--app-port=(\d+)", raw)
    token_match = re.search(r"--remoting-auth-token=([\w-]+)", raw)
    if not port_match or not token_match:
        sys.exit(
            "Could not read port/token from the client process. Riot may have "
            "changed the launch flags - this is the first thing to check if the "
            "script stops working."
        )
    region_match = re.search(r"--region=(\w+)", raw)
    region = region_match.group(1) if region_match else None
    return port_match.group(1), token_match.group(1), region


class Lcu:
    """Thin client for the local League Client API."""

    def __init__(self, port, token):
        self.base_url = f"https://127.0.0.1:{port}"
        creds = base64.b64encode(f"riot:{token}".encode("ascii")).decode("ascii")
        self.headers = {"Authorization": f"Basic {creds}"}
        # LCU uses a self-signed cert; only trust it for these local calls.
        self.ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
        self.ctx.check_hostname = False
        self.ctx.verify_mode = ssl.CERT_NONE

    def get_json(self, path):
        req = Request(self.base_url + path, headers=self.headers)
        with urlopen(req, context=self.ctx) as resp:
            return json.loads(resp.read().decode("utf-8"))

    def get_bytes(self, path):
        req = Request(self.base_url + path, headers=self.headers)
        with urlopen(req, context=self.ctx) as resp:
            return resp.read()


def fetch_region_locale(lcu):
    """Fallback-only region lookup, off an undocumented LCU endpoint - same
    caveats as everything else in this file. Used only when the client's
    own "--region=" launch flag (read in find_client(), a more reliable
    source) isn't available. Confirmed by hand to return a short regional
    code here (e.g. "NA") rather than a full platform id (e.g. "NA1") -
    see SHORT_REGION_TO_CONTINENT below. Returns None on any
    failure/unexpected shape so callers can fall back further.
    """
    try:
        resp = lcu.get_json("/riotclient/region-locale")
    except (URLError, HTTPError):
        return None
    return (resp or {}).get("region")


# Riot's match-v5 API routes by continent, not platform - this is Riot's
# documented platform -> continent grouping. If Riot adds/moves a platform
# this will need updating. Keyed by the full platform id (e.g. "NA1"), as
# read from the client's own "--region=" launch flag.
PLATFORM_TO_CONTINENT = {
    "NA1": "americas", "BR1": "americas", "LA1": "americas", "LA2": "americas", "OC1": "americas",
    "EUW1": "europe", "EUN1": "europe", "TR1": "europe", "RU": "europe",
    "KR": "asia", "JP1": "asia",
    "PH2": "sea", "SG2": "sea", "TH2": "sea", "TW2": "sea", "VN2": "sea",
}

# Same continent grouping, but keyed by the shorter regional code the LCU's
# /riotclient/region-locale endpoint returns instead (e.g. "NA" rather than
# "NA1") - only used as a fallback when --region= wasn't on the command line.
SHORT_REGION_TO_CONTINENT = {
    "NA": "americas", "BR": "americas", "LAN": "americas", "LAS": "americas", "OCE": "americas",
    "EUW": "europe", "EUNE": "europe", "TR": "europe", "RU": "europe",
    "KR": "asia", "JP": "asia",
    "PH": "sea", "SG": "sea", "TH": "sea", "TW": "sea", "VN": "sea",
}


class RiotApiError(Exception):
    """Raised when the Riot API key is missing/invalid, or retries are exhausted."""


class RiotApi:
    """Thin client for Riot's public match-v5 API, with basic rate-limit
    handling tuned for a personal/development API key.

    KNOWN WEAKNESSES
        - The throttle limits below are conservative guesses at a
          Development API Key's published limits (~20 req/sec, 100
          req/2min) - a production key has different (higher) limits that
          aren't plumbed in here.
        - On a 429 this sleeps for the server's Retry-After header, or 2s
          if that header's missing, and retries a bounded number of times
          before giving up - a very large account's first full sync can
          still take a long time even with that working correctly.
    """

    MAX_REQUESTS_PER_SECOND = 18
    MAX_REQUESTS_PER_2MIN = 95
    MAX_RETRIES = 5

    def __init__(self, api_key, continent):
        self.api_key = api_key
        self.continent = continent
        self._request_times = []  # rolling log of request timestamps, for throttling

    def _throttle(self):
        now = time.monotonic()
        self._request_times = [t for t in self._request_times if now - t < 120]
        while len(self._request_times) >= self.MAX_REQUESTS_PER_2MIN:
            time.sleep(0.25)
            now = time.monotonic()
            self._request_times = [t for t in self._request_times if now - t < 120]
        recent = [t for t in self._request_times if now - t < 1]
        if len(recent) >= self.MAX_REQUESTS_PER_SECOND:
            time.sleep(max(0.0, 1 - (now - recent[0])))
        self._request_times.append(time.monotonic())

    def get_json(self, path, params=None):
        query = f"?{urlencode(params)}" if params else ""
        url = f"https://{self.continent}.api.riotgames.com{path}{query}"
        # Riot's API sits behind Cloudflare, which has been confirmed (via
        # a Cloudflare "error code: 1010" block) to reject urllib's default
        # "Python-urllib/x.y" User-Agent outright as bot traffic - any
        # other value works, so set a plain descriptive one.
        req = Request(url, headers={"X-Riot-Token": self.api_key, "User-Agent": "lol-skins-script/1.0"})

        for _attempt in range(self.MAX_RETRIES):
            self._throttle()
            try:
                with urlopen(req) as resp:
                    return json.loads(resp.read().decode("utf-8"))
            except HTTPError as e:
                if e.code == 429:
                    try:
                        # Spec also allows an HTTP-date here, which int()
                        # can't parse - fall back to a flat 2s in that case.
                        retry_after = int(e.headers.get("Retry-After", "2"))
                    except (TypeError, ValueError):
                        retry_after = 2
                    print(f"  Riot API rate limit hit, waiting {retry_after}s...", file=sys.stderr)
                    time.sleep(retry_after)
                    continue
                if e.code in (401, 403):
                    raise RiotApiError(
                        "Riot API key rejected (expired or invalid) - generate a "
                        f"fresh one at {RIOT_DEV_PORTAL_URL} and pass it via "
                        "--riot-api-key."
                    ) from e
                if e.code == 404:
                    return None
                if e.code in (500, 502, 503, 504):
                    # Transient server-side error - worth a retry rather than
                    # aborting a long sync over one bad request.
                    print(f"  Riot API returned HTTP {e.code}, retrying...", file=sys.stderr)
                    time.sleep(2)
                    continue
                # Anything else (unexpected status code): give up on this
                # request, but as RiotApiError so callers can still fall
                # back to the local client instead of crashing outright.
                raise RiotApiError(f"Riot API request failed: HTTP {e.code} for {path}") from e
            except URLError as e:
                # Dropped connection, DNS blip, etc. - also worth a retry
                # over thousands of requests in a long first sync.
                print(f"  Riot API network error ({e.reason}), retrying...", file=sys.stderr)
                time.sleep(2)
                continue
        raise RiotApiError("Riot API retries exhausted (rate limit or network issues) - try again later.")


def fetch_riot_puuid(riot_api, game_name, tag_line):
    """Resolve a Riot ID (gameName#tagLine) to the real account-wide puuid
    Riot's public API expects.

    This is NOT the same value as the "puuid" field on the LCU's
    current-summoner response - that one is a 36-character locally-scoped
    identifier, confirmed (via a "Bad Request - Exception decrypting ..."
    400 from match-v5) to be unusable with the public API, which expects a
    much longer (~78-char) opaque string. This is the one source of truth
    for that.
    """
    resp = riot_api.get_json(f"/riot/account/v1/accounts/by-riot-id/{quote(game_name)}/{quote(tag_line)}")
    if not resp or not resp.get("puuid"):
        raise RiotApiError(
            f"Couldn't resolve a Riot API puuid for {game_name}#{tag_line} via "
            "account-v1 - double check the Riot ID is correct."
        )
    return resp["puuid"]


def _load_match_cache(cache_path):
    if not cache_path.exists():
        return {}
    try:
        return json.loads(cache_path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return {}


def _save_match_cache(cache_path, cache):
    cache_path.parent.mkdir(parents=True, exist_ok=True)
    cache_path.write_text(json.dumps(cache), encoding="utf-8")


def _summarize_match(match, puuid):
    """Pull just the fields this script needs out of one full match-v5
    match payload, for the local player identified by puuid."""
    info = match.get("info") or {}
    for p in info.get("participants") or []:
        if p.get("puuid") == puuid:
            return {
                "championId": p.get("championId"),
                "win": bool(p.get("win")),
                "queueId": info.get("queueId"),
            }
    return None


def fetch_full_match_history(riot_api, puuid, cache_path, max_matches=10000):
    """Return (games-played Counter, wins Counter, total games counted),
    each Counter keyed by championId - same shape as fetch_match_history(),
    but sourced from Riot's match-v5 API instead of the local client.

    Caches one compact record per match on disk (keyed by match id) so a
    repeat run only has to fetch matches it hasn't seen before: match ids
    come back most-recent-first, so once a whole page of ids is already in
    the cache, everything older must be too, and paging stops there.
    A fresh account (empty cache) instead pages all the way to Riot's
    actual end of history (an empty ids page) or the max_matches safety
    cap, whichever comes first - that first run can be slow (one request
    per previously-unseen match) but every run after it is fast.
    """
    cache = _load_match_cache(cache_path)
    page_size = 100
    start = 0
    new_matches = 0

    while start < max_matches:
        match_ids = riot_api.get_json(
            f"/lol/match/v5/matches/by-puuid/{puuid}/ids",
            params={"start": start, "count": page_size},
        )
        if not match_ids:
            break

        page_all_cached = True
        for match_id in match_ids:
            if match_id in cache:
                continue
            page_all_cached = False
            match = riot_api.get_json(f"/lol/match/v5/matches/{match_id}")
            if match is None:
                continue
            record = _summarize_match(match, puuid)
            if record is not None:
                cache[match_id] = record
                new_matches += 1
                if new_matches % 50 == 0:
                    print(f"  ...fetched {new_matches} new match(es) so far", file=sys.stderr)
                if new_matches % 200 == 0:
                    _save_match_cache(cache_path, cache)  # checkpoint a long first sync

        if page_all_cached:
            break  # everything on this page (and so everything older) was already cached
        if len(match_ids) < page_size:
            break
        start += page_size

    if new_matches:
        _save_match_cache(cache_path, cache)
        print(f"Cached {new_matches} new match(es) ({len(cache)} total cached).")
    else:
        print(f"No new matches since last run - using {len(cache)} cached match(es).")

    champion_counts = Counter()
    champion_wins = Counter()
    for record in cache.values():
        if record.get("queueId") == 0:
            continue  # custom/practice-tool games - same reasoning as fetch_match_history()
        champ_id = record.get("championId")
        if champ_id is None:
            continue
        champion_counts[champ_id] += 1
        if record.get("win"):
            champion_wins[champ_id] += 1

    return champion_counts, champion_wins, sum(champion_counts.values())


def _find_participant(game, puuid):
    """Return the local player's participant entry from one match-history
    game, or None if it can't be identified.

    Three participant shapes have been observed in the wild for this
    endpoint, checked in order of how common they are on a current client:
      1. Modern payloads embed "puuid" directly on each participant.
      2. Older payloads instead map participantId -> puuid via a separate
         "participantIdentities" list.
      3. Some responses pre-filter "participants" down to just the local
         player (one entry) with no puuid field needed at all.
    """
    participants = game.get("participants") or []

    for p in participants:
        if p.get("puuid") == puuid:
            return p

    participant_id = None
    for identity in game.get("participantIdentities") or []:
        if (identity.get("player") or {}).get("puuid") == puuid:
            participant_id = identity.get("participantId")
            break
    if participant_id is not None:
        for p in participants:
            if p.get("participantId") == participant_id:
                return p

    if len(participants) == 1:
        return participants[0]

    return None


def fetch_match_history(lcu, puuid):
    """Return (games-played Counter, wins Counter, total games counted),
    each Counter keyed by championId.

    Pages through /lol-match-history/v1/products/lol/{puuid}/matches, which
    is undocumented/unofficial like the rest of the LCU - it only covers
    whatever history this client has synced locally, not the account's
    full career. In practice this has only ever been observed to serve
    back ~200 games before returning an empty page, no matter how many
    games are actually on the account - that looks like a hard ceiling on
    the LCU's local cache, not something pagination can get around.
    Pagination is driven off the server's own "gameIndexEnd" field rather
    than assuming a fixed page size, since the client has been seen
    silently clamping how many games it returns per request regardless of
    the requested begIndex/endIndex range. Stops when the server returns
    an empty page (real end of available history) or the max_games safety
    cap below is hit, whichever comes first.
    """
    champion_counts = Counter()
    champion_wins = Counter()
    total_games = 0
    page_size = 100
    max_games = 2000  # hard cap: so a very long history doesn't hang the script
    beg_index = 0

    while beg_index < max_games:
        end_index = beg_index + page_size - 1
        try:
            resp = lcu.get_json(
                f"/lol-match-history/v1/products/lol/{puuid}/matches"
                f"?begIndex={beg_index}&endIndex={end_index}"
            )
        except (URLError, HTTPError):
            break

        games_obj = (resp or {}).get("games") or {}
        games = games_obj.get("games") or []
        if not games:
            break

        for game in games:
            # Custom games (queueId 0, incl. the practice tool) commonly
            # have only bots as "participants" - the single-participant
            # fallback in _find_participant can't tell a bot from the
            # local player there, so skip these rather than risk
            # attributing a bot's champion/result to you. They also aren't
            # meaningful "games played" for a play-rate/win-rate stat.
            if game.get("queueId") == 0:
                continue
            participant = _find_participant(game, puuid)
            if participant is None:
                continue
            champ_id = participant.get("championId")
            if champ_id is None:
                continue
            champion_counts[champ_id] += 1
            total_games += 1
            # "win" has been seen both flattened onto the participant
            # directly and nested under a "stats" sub-object - check both.
            win = participant.get("win")
            if win is None:
                win = (participant.get("stats") or {}).get("win")
            if win:
                champion_wins[champ_id] += 1

        # Advance past whatever the server actually returned, not what we
        # asked for - the client has been seen clamping each page well
        # below the requested range (e.g. always ~20 games even when 100
        # are requested), so a short page does NOT mean history ended.
        # Only an empty page (handled above) means that.
        last_index = games_obj.get("gameIndexEnd")
        if last_index is None:
            last_index = beg_index + len(games) - 1
        next_beg_index = last_index + 1
        if next_beg_index <= beg_index:
            break  # server isn't advancing - bail instead of looping forever
        beg_index = next_beg_index

    return champion_counts, champion_wins, total_games


def build_html(display_name, sections, total_skins, total_shards, total_champs):
    generated = datetime.now().strftime("%Y-%m-%d %H:%M")
    shards_note = f" (+{total_shards} shard{'s' if total_shards != 1 else ''})" if total_shards else ""
    return f"""<!DOCTYPE html>
<html><head><meta charset='utf-8'>
<title>LoL Skins - {display_name}</title>
<style>
  body {{ font-family: -apple-system, Segoe UI, sans-serif; background:#0b0e14; color:#e6e6e6; margin:0; padding:24px; }}
  header {{ margin-bottom: 24px; }}
  header h1 {{ margin:0 0 4px 0; }}
  header .meta {{ color:#9aa0a6; font-size:14px; }}
  section {{ margin-bottom: 32px; }}
  h2 {{ border-bottom: 1px solid #2a2f3a; padding-bottom: 6px; }}
  h2 a {{ color:#e6e6e6; text-decoration:none; }}
  h2 a:hover {{ text-decoration:underline; }}
  .count {{ color:#9aa0a6; font-weight:normal; font-size:15px; }}
  .playrate {{ color:#9aa0a6; font-size:13px; margin:2px 0 0 0; }}
  .winrate {{ color:#9aa0a6; font-size:13px; margin:2px 0 0 0; }}
  .grid {{ display:flex; flex-wrap:wrap; gap:16px; margin-top:12px; }}
  .card {{ width:260px; background:#151a24; border-radius:8px; overflow:hidden; text-align:center; padding-bottom:6px; border:2px solid #4ad9e0; }}
  .card img {{ width:260px; height:153px; object-fit:cover; display:block; cursor:zoom-in; }}
  .card.shard {{ background:transparent; border:2px dashed #888; opacity:0.7; }}
  .noimg {{ width:260px; height:153px; display:flex; align-items:center; justify-content:center; color:#666; font-size:11px; }}
  .name {{ font-size:12px; padding:6px 4px 2px 4px; }}
  .chromas {{ font-size:10px; color:#9aa0a6; padding-bottom:4px; }}
  .shard-badge {{ font-size:10px; color:#e0b84b; padding-bottom:4px; text-transform:uppercase; letter-spacing:0.03em; }}
  #search {{ width:100%; max-width:360px; margin-top:14px; padding:8px 12px; font-size:14px; background:#151a24; border:1px solid #2a2f3a; border-radius:6px; color:#e6e6e6; box-sizing:border-box; }}
  #search:focus {{ outline:none; border-color:#4a90e2; }}
  #no-results {{ display:none; color:#9aa0a6; }}
  .shards-toggle {{ display:flex; align-items:center; gap:6px; margin-top:10px; font-size:13px; color:#9aa0a6; }}
  .shards-toggle input {{ margin:0; }}
  #hover-preview {{ position:fixed; display:none; z-index:1000; pointer-events:none; }}
  #hover-preview img {{ display:block; max-width:min(90vw, 900px); max-height:85vh; border-radius:10px; box-shadow:0 8px 32px rgba(0,0,0,0.6); border:2px solid #2a2f3a; }}
</style>
</head><body>
<header>
  <h1>League of Legends - Owned Skins</h1>
  <div class='meta'>{display_name} - {total_skins} skins{shards_note} across {total_champs} champions - generated {generated}</div>
  <input id='search' type='text' placeholder='Filter by champion or skin name...' oninput='filterSkins()'>
  <label class='shards-toggle'><input id='shards-only' type='checkbox' oninput='filterSkins()'> Only champs with an available skin shard</label>
</header>
<div id='no-results'>No matches.</div>
{sections}
<div id='hover-preview'><img alt=''></div>
<script>
function filterSkins() {{
  var query = document.getElementById('search').value.trim().toLowerCase();
  var shardsOnly = document.getElementById('shards-only').checked;
  var anyVisible = false;
  document.querySelectorAll('section').forEach(function(section) {{
    if (shardsOnly && !section.querySelector('.card.shard')) {{
      section.style.display = 'none';
      return;
    }}
    var champMatch = section.dataset.champ.indexOf(query) !== -1;
    var sectionHasCard = false;
    section.querySelectorAll('.card').forEach(function(card) {{
      var show = champMatch || card.dataset.name.indexOf(query) !== -1;
      card.style.display = show ? '' : 'none';
      if (show) sectionHasCard = true;
    }});
    section.style.display = sectionHasCard ? '' : 'none';
    if (sectionHasCard) anyVisible = true;
  }});
  document.getElementById('no-results').style.display = anyVisible ? 'none' : 'block';
}}

(function() {{
  var preview = document.getElementById('hover-preview');
  var previewImg = preview.querySelector('img');
  var pad = 20;

  function position(e) {{
    var w = preview.offsetWidth, h = preview.offsetHeight;
    var x = e.clientX + pad;
    var y = e.clientY + pad;
    if (x + w > window.innerWidth) x = e.clientX - w - pad;
    if (y + h > window.innerHeight) y = e.clientY - h - pad;
    preview.style.left = Math.max(pad, x) + 'px';
    preview.style.top = Math.max(pad, y) + 'px';
  }}

  document.querySelectorAll('.card img').forEach(function(img) {{
    img.addEventListener('mouseenter', function(e) {{
      previewImg.src = img.src;
      preview.style.display = 'block';
      position(e);
    }});
    img.addEventListener('mousemove', position);
    img.addEventListener('mouseleave', function() {{
      preview.style.display = 'none';
    }});
  }});
}})();
</script>
</body></html>
"""


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--include-base", action="store_true", help="also list free/default skins")
    parser.add_argument(
        "--shards-only", action="store_true",
        help="only show champions that have at least one available skin shard",
    )
    parser.add_argument("--out", default="lol-skins.html", help="output HTML path")
    parser.add_argument("--no-open", action="store_true", help="don't auto-open the result")
    parser.add_argument(
        "--riot-api-key", default=None,
        help="REQUIRED. Riot developer API key, for full match history beyond "
             "the local client's ~200-game ceiling. Get one (free, valid ~24h) "
             "from https://developer.riotgames.com/",
    )
    parser.add_argument(
        "--cache-dir", default="match_cache",
        help="directory to cache full match history in (one file per account, "
             "named by puuid)",
    )
    argv = sys.argv[1:]
    # Shorthand: a single bare argument (e.g. `python lol-skins.py
    # RGAPI-xxxx`) is treated as the API key - that's the overwhelmingly
    # common thing anyone passes here, so don't make --riot-api-key
    # mandatory boilerplate for the common case.
    if len(argv) == 1 and not argv[0].startswith("-"):
        argv = ["--riot-api-key", argv[0]]
    args = parser.parse_args(argv)

    if not args.riot_api_key:
        print(f"No --riot-api-key given - opening {RIOT_DEV_PORTAL_URL} so you can grab one...", flush=True)
        webbrowser.open(RIOT_DEV_PORTAL_URL)
        parser.error(
            "--riot-api-key is required. Get a free key (valid ~24h) from "
            f"{RIOT_DEV_PORTAL_URL} (log in, then copy the "
            "\"Development API Key\" on your dashboard) and pass it as:\n"
            "  python lol-skins.py --riot-api-key RGAPI-xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
        )
    riot_api_key = args.riot_api_key

    print("Finding League Client...")
    port, token, region = find_client()
    lcu = Lcu(port, token)

    print("Reading current summoner...")
    summoner = lcu.get_json("/lol-summoner/v1/current-summoner")
    summoner_id = summoner["summonerId"]
    puuid = summoner.get("puuid", "")
    display_name = (
        f"{summoner['gameName']}#{summoner['tagLine']}"
        if summoner.get("gameName")
        else summoner.get("displayName", "Unknown")
    )

    print("Reading champion list...")
    champions = lcu.get_json(f"/lol-champions/v1/inventories/{summoner_id}/champions-minimal")
    champ_lookup = {c["id"]: c["name"] for c in champions}

    def champ_name_for(champ_id):
        return champ_lookup.get(champ_id, f"Unknown ({champ_id})")

    print("Reading skin inventory...")
    skins = lcu.get_json(f"/lol-champions/v1/inventories/{summoner_id}/skins-minimal")

    # skins-minimal has been observed to list the same skin more than once
    # (identical id) and/or split one champion's skins across more than one
    # championId. Collapse by id first, then group by resolved champion name
    # rather than raw id so those don't surface as duplicate cards/sections.
    skins = list({s["id"]: s for s in skins}.values())

    # Each chroma is also listed both nested under its parent skin's
    # "chromas" array AND as its own standalone entry (sharing the parent's
    # name). Drop the standalone entries so chromas don't show up as
    # duplicate skin cards - they're already counted via owned_chromas below.
    chroma_ids = {c["id"] for s in skins for c in s.get("chromas", [])}

    owned = [
        s for s in skins
        if s.get("ownership", {}).get("owned")
        and (args.include_base or not s.get("isBase"))
        and s["id"] not in chroma_ids
    ]

    if not owned:
        print(
            "WARNING: no owned (non-base) skins found. Re-run with --include-base "
            "if you expected base skins to show up too.",
            file=sys.stderr,
        )

    print("Reading loot (skin shards)...")
    try:
        loot = lcu.get_json("/lol-loot/v1/player-loot")
    except (URLError, HTTPError):
        loot = []

    skins_by_id = {s["id"]: s for s in skins}
    owned_ids = {s["id"] for s in owned}
    shard_skin_ids = set()
    for item in loot:
        item_type = (item.get("type") or "").upper()
        if "SKIN" not in item_type or ("RENTAL" not in item_type and "SHARD" not in item_type):
            continue
        if item.get("count", 0) <= 0:
            continue
        # lootId has been seen as either the bare skin id ("81003") or a
        # prefixed form ("CHAMPION_SKIN_RENTAL_81003") - pull the trailing digits.
        digits = re.search(r"(\d+)$", str(item.get("lootId", "")))
        if not digits:
            continue
        shard_id = int(digits.group(1))
        if shard_id in skins_by_id and shard_id not in owned_ids:
            shard_skin_ids.add(shard_id)
    shards = [skins_by_id[sid] for sid in shard_skin_ids]

    if not shards and loot:
        skin_like_types = sorted({
            item.get("type") for item in loot
            if item.get("type") and "SKIN" in item["type"].upper()
        })
        if skin_like_types:
            print(
                f"No shards matched. Skin-related loot types seen: {skin_like_types} "
                "- update the matching logic in main() if these don't look right.",
                file=sys.stderr,
            )

    champion_play_counts = champion_win_counts = None

    if riot_api_key:
        # Prefer the client's own "--region=" launch flag (full platform
        # id, e.g. "NA1") - only fall back to the LCU's region-locale
        # endpoint (shorter codes, e.g. "NA") if that flag wasn't present.
        platform = region
        continent = PLATFORM_TO_CONTINENT.get((platform or "").upper())
        if not continent:
            platform = fetch_region_locale(lcu)
            continent = SHORT_REGION_TO_CONTINENT.get((platform or "").upper())
        if not continent:
            print(
                f"WARNING: couldn't map platform/region '{region}'/'{platform}' to a "
                "Riot API continent - falling back to the local client's match "
                "history instead.",
                file=sys.stderr,
            )
        elif not summoner.get("gameName"):
            print(
                "WARNING: no Riot ID (gameName#tagLine) on this account, needed to "
                "resolve the Riot API puuid - falling back to the local client's "
                "match history instead.",
                file=sys.stderr,
            )
        else:
            riot_api = RiotApi(riot_api_key, continent)
            try:
                # The LCU's own "puuid" field on current-summoner is NOT the
                # same puuid the public Riot API expects (confirmed: the
                # LCU one is a 36-char local identifier, the real one is a
                # ~78-char opaque string) - resolve the real one via the
                # Riot ID instead.
                riot_puuid = fetch_riot_puuid(riot_api, summoner["gameName"], summoner["tagLine"])
            except RiotApiError as e:
                print(f"WARNING: {e}", file=sys.stderr)
                print("Falling back to the local client's match history instead.", file=sys.stderr)
            else:
                print(f"Reading full match history via Riot API ({platform} -> {continent})...")
                cache_path = Path(args.cache_dir) / f"{riot_puuid}.json"
                try:
                    champion_play_counts, champion_win_counts, total_games_played = (
                        fetch_full_match_history(riot_api, riot_puuid, cache_path)
                    )
                except RiotApiError as e:
                    print(f"WARNING: {e}", file=sys.stderr)
                    print("Falling back to the local client's match history instead.", file=sys.stderr)

    if champion_play_counts is None:
        print("Reading match history (local client only, ~200 game ceiling)...")
        champion_play_counts, champion_win_counts, total_games_played = fetch_match_history(lcu, puuid)

    if not total_games_played:
        print(
            "WARNING: no match history found - play rate/win rate won't be shown.",
            file=sys.stderr,
        )
    play_rate_by_champ = {
        champ_name_for(champ_id): count
        for champ_id, count in champion_play_counts.items()
    }
    win_rate_by_champ = {
        champ_name_for(champ_id): count
        for champ_id, count in champion_win_counts.items()
    }

    print(f"Found {len(owned)} owned skin(s) and {len(shards)} shard(s). Downloading thumbnails...")

    by_champ = {}
    for skin in owned:
        champ_name = champ_name_for(skin["championId"])
        by_champ.setdefault(champ_name, []).append(skin)

    by_champ_shards = {}
    for skin in shards:
        champ_name = champ_name_for(skin["championId"])
        by_champ_shards.setdefault(champ_name, []).append(skin)

    def render_card(skin, is_shard):
        # Prefer the full splash art (much higher-res than the champ-select
        # tile), constructed directly from the LCU's game-data asset server -
        # same local client, same undocumented-path caveat as tilePath, just
        # a bigger asset. Fall back to tilePath if that path 404s on an
        # older/newer client, and to a placeholder if both fail.
        splash_path = f"/lol-game-data/assets/v1/champion-splashes/{skin['championId']}/{skin['id']}.jpg"
        img_tag = None
        for path in (splash_path, skin.get("tilePath")):
            if not path:
                continue
            try:
                img_bytes = lcu.get_bytes(path)
                b64 = base64.b64encode(img_bytes).decode("ascii")
                img_tag = f"<img src='data:image/jpeg;base64,{b64}' alt='{skin['name']}'>"
                break
            except (URLError, HTTPError, KeyError):
                continue
        if img_tag is None:
            img_tag = "<div class='noimg'>no image</div>"

        if is_shard:
            note = "<div class='shard-badge'>Skin Shard</div>"
            card_class = "card shard"
        else:
            owned_chromas = [
                c for c in skin.get("chromas", []) if c.get("ownership", {}).get("owned")
            ]
            note = (
                f"<div class='chromas'>+{len(owned_chromas)} chroma(s)</div>"
                if owned_chromas
                else ""
            )
            card_class = "card"

        return (
            f"<div class='{card_class}' data-name='{skin['name'].lower()}'>{img_tag}"
            f"<div class='name'>{skin['name']}</div>{note}</div>"
        )

    sections = []
    total_skins = 0
    total_shards = 0
    all_champ_names = (
        sorted(by_champ_shards) if args.shards_only
        else sorted(set(by_champ) | set(by_champ_shards))
    )
    for champ_name in all_champ_names:
        champ_skins = list({s["name"]: s for s in by_champ.get(champ_name, [])}.values())
        champ_skins.sort(key=lambda s: s["name"])
        total_skins += len(champ_skins)

        owned_names = {s["name"] for s in champ_skins}
        champ_shards = [
            s for s in {s["name"]: s for s in by_champ_shards.get(champ_name, [])}.values()
            if s["name"] not in owned_names
        ]
        champ_shards.sort(key=lambda s: s["name"])
        total_shards += len(champ_shards)

        cards = [render_card(skin, False) for skin in champ_skins]
        cards += [render_card(skin, True) for skin in champ_shards]

        count_parts = []
        if champ_skins:
            count_parts.append(f"{len(champ_skins)} owned")
        if champ_shards:
            count_parts.append(f"{len(champ_shards)} shard{'s' if len(champ_shards) != 1 else ''}")
        count_label = ", ".join(count_parts)

        wiki_url = (
            f"https://wiki.leagueoflegends.com/en-us/{quote(champ_name.replace(' ', '_'))}/Cosmetics"
        )
        games_played = play_rate_by_champ.get(champ_name, 0)
        play_rate_html = ""
        if total_games_played:
            play_rate_html = (
                f"<div class='playrate'>Play Rate: {games_played} / {total_games_played} games</div>"
            )
        win_rate_html = ""
        if games_played:
            wins = win_rate_by_champ.get(champ_name, 0)
            win_rate_html = f"<div class='winrate'>Win Rate: {wins} / {games_played} games</div>"
        sections.append(
            f"<section data-champ='{champ_name.lower()}'>"
            f"<h2><a href='{wiki_url}' target='_blank' rel='noopener'>{champ_name}</a> "
            f"<span class='count'>({count_label})</span></h2>"
            f"{play_rate_html}{win_rate_html}"
            f"<div class='grid'>{''.join(cards)}</div></section>"
        )

    html = build_html(display_name, "\n".join(sections), total_skins, total_shards, len(all_champ_names))

    out_path = Path(args.out)
    out_path.write_text(html, encoding="utf-8")
    print(f"Wrote {out_path.resolve()}")

    if not args.no_open:
        webbrowser.open(out_path.resolve().as_uri())


if __name__ == "__main__":
    main()