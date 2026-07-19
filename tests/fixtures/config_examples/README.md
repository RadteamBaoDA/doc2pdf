# Config-driven Excel examples

Generate the workbooks with:

```powershell
.\.venv\Scripts\python.exe tests\fixtures\build_config_excel_examples.py
```

The files intentionally match the active `config.yml` selectors:

| Workbook | Coverage |
|---|---|
| `config_auto_layout.xlsx` | `*Summary*` one-page attempt/path label; `*Data*` ten-row chunks; auto orientation; repeating titles; full-width horizontal pagination |
| `config_print_area_trim.xlsx` | Preserved authored print area, landscape printing, margins, and Excel PDF whitespace trimming |
| `appendix_oversized.xlsx` | `*appendix*` path rule and oversized content with `oversized_action: error` |
| `config_empty_visible.xlsx` | Explicit failure for a workbook with no visible content |

To run the Office-dependent conversion on an interactive Windows session:

```powershell
doc2pdf convert tests\fixtures\config_examples --output tests\fixtures\config_examples\output --config config.yml --trim --verbose
```
