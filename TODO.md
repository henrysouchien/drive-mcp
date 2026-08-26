# drive-mcp TODO

## ~~gdrive_search: surface file IDs in results~~ DONE

`gdrive_search` currently formats results with name + link but omits the file ID:

```
📄 Investment-newsletters
   Modified: 2026-01-28T18:17:29.653Z
   Link: https://docs.google.com/spreadsheets/d/1hTgORVAftkCr9s-.../edit?usp=drivesdk
```

The underlying `search_files()` in `google_drive.py` already returns `id` in each result dict — it's just dropped during formatting in `server.py:103-109`.

**Problem:** Consumers that need the file ID (e.g. taskflow's `tf_pin_add` requires the raw ID for gdrive pins) have to parse it out of the URL. LLM agents frequently get this wrong — grabbing the file name instead, or mangling the ID extraction. This caused 7 bad pins during live testing that had to be manually fixed.

**Fix:** Add an `ID: {f['id']}` line to each search result in `server.py`. The data is already there — just needs to be included in the formatted output. Example:

```
📄 Investment-newsletters
   ID: 1hTgORVAftkCr9s-P9QqHOmyhH-kWEW4kj-w-UbGzKIk
   Modified: 2026-01-28T18:17:29.653Z
   Link: https://docs.google.com/spreadsheets/d/1hTgORVAftkCr9s-.../edit?usp=drivesdk
```

**Where:** `src/server.py` lines 103-109, `gdrive_search` function.

**Resolved:** `gdrive_search` emits a dedicated `ID:` line immediately after
each result name. A behavioral regression test now proves the raw ID is exposed
before modification/link metadata and that the requested result limit reaches
the Drive adapter. The repository test bootstrap also now pins this repo's
generic `src` package, preventing shared development environments from silently
importing another MCP repository's `src.server`. Full suite: 11 passed; Ruff
passes for the affected files (2026-07-15).
