"""
Get-LoLSkins.py

Lists all League of Legends skins you own, grouped by champion, as one
self-contained HTML file.

REQUIREMENTS
    - Windows, with the League Client (LeagueClientUx.exe) running and logged in
    - Python 3.8+, stdlib only (no pip installs needed)

HOW IT WORKS
    Riot's public API does NOT expose skin ownership for any account - that
    data only exists locally, through the League Client's own local API (the
    "LCU"), while the client is open and you're logged in. This script finds
    the running client, reads its port + auth token off its own process
    command line, pulls your owned skins over that local API, and writes one
    HTML file with thumbnails embedded as base64 (so it still opens fine
    after you close the client).

USAGE
    python Get-LoLSkins.py
    python Get-LoLSkins.py --include-base       # also list free/default skins
    python Get-LoLSkins.py --shards-only        # only champs with an available skin shard
    python Get-LoLSkins.py --out C:\\temp\\skins.html
    python Get-LoLSkins.py --no-open

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
    - Play Rate and Win Rate come from paging through the local
      match-history API (fetch_match_history()), which has only ever been
      observed to serve back ~200 games before returning an empty page,
      regardless of how many games are actually on the account - this
      looks like a hard ceiling on what the LCU caches locally, not
      something this script can page past. Full career history would
      require Riot's online Match-V5 API with a developer API key
      instead, which this script intentionally does not use (it's
      LCU-only, zero-setup by design). There's a 2000-game safety cap in
      the code, but in practice the ~200-game server ceiling is hit first.
      Win Rate additionally depends on a "win" field being present on the
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
import webbrowser
from collections import Counter
from datetime import datetime
from pathlib import Path
from urllib.parse import quote
from urllib.request import Request, urlopen
from urllib.error import URLError, HTTPError


def find_client():
    """Return (port, token) parsed from the running LeagueClientUx.exe command line."""
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
    return port_match.group(1), token_match.group(1)


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
    args = parser.parse_args()

    print("Finding League Client...")
    port, token = find_client()
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

    print("Reading match history...")
    champion_play_counts, champion_win_counts, total_games_played = fetch_match_history(lcu, puuid)
    if not total_games_played:
        print(
            "WARNING: no match history found via the client - play rate/win rate won't be shown.",
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