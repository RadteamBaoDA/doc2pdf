# doc2pdf

[Tiếng Việt](docs/README.vi.md)

A Windows CLI that uses Microsoft Office to convert Word, Excel, and PowerPoint documents to PDF. It supports single files, recursive directory processing, path-pattern configuration, error reports, and PDF post-processing.

In addition to PDF conversion, the `convert-macros` command creates macro-free copies:

- `.docm` → `.docx`
- `.pptm` → `.pptx`
- `.xlsm` → `.xlsx`

> Converting to a macro-free format removes the VBA project. The source file is not modified.

## Requirements

- Windows
- Python 3.12+
- Microsoft Word, Excel, and/or PowerPoint installed for the corresponding file types

## Installation

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install .
```

Development installation:

```powershell
python -m pip install -e ".[dev]"
```

## Commands and Run Modes

### Help and version

```powershell
doc2pdf --help
doc2pdf --version
doc2pdf convert --help
doc2pdf convert-macros --help
```

Run directly from source without installing the command:

```powershell
python -m src.cli --help
```

### Mode 1: Office to PDF

Convert one file. The default output directory is `output`:

```powershell
doc2pdf convert "input\report.docx"
```

Specify the output PDF file:

```powershell
doc2pdf convert "input\report.docx" --output "output\report.pdf"
```

Recursively convert a directory while preserving its subdirectory structure:

```powershell
doc2pdf convert "input" --output "output"
```

Use a different configuration file:

```powershell
doc2pdf convert "input" --output "output" --config "config.yml"
```

Enable verbose logs, enable or disable whitespace trimming, or override the trim margin:

```powershell
doc2pdf convert "input" --output "output" --verbose
doc2pdf convert "input" --output "output" --trim
doc2pdf convert "input" --output "output" --no-trim
doc2pdf convert "input" --output "output" --trim --trim-margin 10
```

Recognized inputs are `.doc`, `.docx`, `.xls`, `.xlsx`, `.xlsm`, `.xlsb`, `.ppt`, `.pptx`, and `.pdf`. The `pdf_handling` section in `config.yml` controls how existing PDF files in the input are handled.

### Mode 2: Remove macros without creating PDFs

Convert one file. The default output directory is `output`:

```powershell
doc2pdf convert-macros "input\report.docm"
doc2pdf convert-macros "input\slides.pptm"
doc2pdf convert-macros "input\workbook.xlsm"
```

Specify the exact output file:

```powershell
doc2pdf convert-macros "input\report.docm" --output "clean\report.docx"
```

Convert all three supported types in a directory while preserving its subdirectory structure:

```powershell
doc2pdf convert-macros "input" --output "clean"
```

This command processes only `.docm`, `.pptm`, and `.xlsm` files. Other files in the directory are ignored. Macros are disabled when Office opens the files through automation and are not saved in the output files.

## Configuration

`config.yml` is the default configuration file. Its main sections include:

- `timeout`: document conversion and Excel trimming timeouts.
- `logging`: log level, log file, rotation, and retention.
- `post_processing`: whitespace trimming after PDF generation.
- `suffix`: PDF filename suffixes for Word, Excel, and PowerPoint.
- `reporting`: result and error reports, plus copying failed source files.
- `pdf_handling`: handling of existing PDF files in the input.
- PDF and layout settings for Word, Excel, and PowerPoint, including pattern-based rules.

`convert-macros` does not use the PDF settings in `config.yml`.

## Testing

```powershell
python -m pytest
```

## Operational Notes

- Close open Office documents before running a batch to avoid file locks and COM dialogs.
- Do not open an output file while it is being converted.
- Logs, summaries, and error reports are written according to `config.yml`.
- If PowerShell cannot find `doc2pdf`, activate the virtual environment or replace `doc2pdf` in the examples with `python -m src.cli`.
