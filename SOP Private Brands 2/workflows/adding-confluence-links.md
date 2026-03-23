# Adding/Updating Confluence Links

[Back to Overview](sop_excel_generation.md)

## Steps

1. Open `Documents/PB SOP Links Confluence.xlsx`
2. Add or modify entries:
   - **Column A**: SOP name (should roughly match the document title)
   - **Column B**: Full Confluence URL (e.g., `https://jswone.atlassian.net/wiki/x/...`)
3. Re-run the generator - links are auto-matched to documents by fuzzy title matching

## How Matching Works

The tool uses a 4-level matching strategy:
1. **Exact normalized match** (lowercase, stripped prefixes)
2. **Substring containment** (title contains link name or vice versa)
3. **Keyword overlap** (3+ matching words)
4. **Fallback map** (hardcoded in `tools/parsers/confluence_links.py` for known mismatches)

If a link isn't matching, add a fallback entry in the `_FALLBACK_MAP` dictionary in `confluence_links.py`.
