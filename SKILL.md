---
name: "feishu-sheet-to-anki"
description: "Generate an Anki .apkg deck (with embedded images) from a Feishu/Lark spreadsheet or a local .xlsx via lark-cli + genanki. Invoke when user wants to turn a Feishu sheet or Excel file into Anki cards/apkg, or batch export sheet images into Anki."
---

# Feishu Sheet / Excel → Anki APKG

Builds an importable Anki `.apkg` from a Feishu/Lark spreadsheet or a local
`.xlsx`: text fields, nested decks, tags, and **embedded in-cell images**
(e.g. brand photos on the card front) are all handled automatically.

Typical workflow: photograph brands while street-scouting → insert photos
into Feishu sheet cells on the phone → register brand info → run this tool
→ import the apkg into Brandki or any Anki app for spaced-repetition review.
CSV cannot be used because it drops embedded images.

## When to Use

- User asks to generate / build / export an Anki deck or `.apkg` from a
  Feishu sheet URL or an Excel file
- Rows of a sheet are cards and the user wants to import them into Anki
- Card fronts/backs contain embedded images that must travel with the deck

## Prerequisites

1. Python 3.9+ with dependencies:
   ```bash
   python3 -m pip install --user genanki openpyxl
   ```
2. Feishu mode only: authenticated `lark-cli`. Auto-detected at
   `~/.trae-cn/plugins/trae-remote-official/lark/*/bin/lark-cli`, or set
   `LARK_CLI=/path/to/lark-cli`. Images download through
   `docs +media-download` with the cell `image_token` — workbook-export
   scopes are NOT required.

## Run

```bash
# Feishu online sheet
python3 scripts/feishu_to_anki.py --url "<FEISHU_SHEET_URL>" -o deck.apkg

# Local Excel (offline; images pulled from xl/media)
python3 scripts/feishu_to_anki.py --xlsx ./brands.xlsx -o deck.apkg
```

Flags: `--sheet-name` (default first sheet), `--deck` (fallback deck name),
`--media-dir` / `--keep-media` (inspect prepared images).

## Table Format (header row = row 1)

| Header pattern | Meaning |
|---|---|
| `Deck1`, `Deck2`, … / `牌组*` | Deck hierarchy levels; non-empty values join with `::` (e.g. `零售::女装::中淑装`) |
| `Front` / `正面` | Card front — text, embedded image, or both |
| `Back` / `背面` | Card back/answer |
| `Note Type` / `卡片类型` | `Basic` (default) or `Cloze` / `填空` |
| `Tags1`, … / `标签*` | Tags split on whitespace / `,` `，` / `;` `；`; leading `#` stripped |
| `Fields1_英文名`, `字段2_x`, … | Extra Anki fields; name = part after the first `_`/`:`/space |

- Unrecognized columns are ignored; fully empty rows are skipped.
- Floating (non-cell) images are not picked up.
- Note GUIDs are stable per (source, row) and deck ids per deck path, so
  re-importing a regenerated apkg updates cards instead of duplicating them.
- Output JSON summary reports notes, decks, image count and skipped rows.

## Verification Checklist

1. Exit code 0; `notes` matches non-empty card rows.
2. `decks` shows the intended `::`-nested paths.
3. `images` matches image cells; `unzip -l deck.apkg` shows non-empty media.
4. Deliver the `.apkg` path; user imports via Anki (File → Import),
   AnkiDroid, AnkiMobile, or Brandki.
