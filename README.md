# markdown-to-docx

Convert Markdown files to Word (`.docx`) documents. Format heavily inspired by [Salesforce's Quip](https://quip.com/).

<!-- IMAGE -->

## Usage

<!-- BEGIN USAGE -->

```console
Usage: markdown-to-docx <input.md> [output.docx] [options]

Arguments:
  input.md        Path to the input Markdown file
  output.docx     Path for the output Word document (default: same name as input)

Options:
  --font-size <n>    Base font size in pt; all styles scale from this (default: 12)
  --footer <text>    Text to display in the bottom-left footer
  --header <text>    Text to display left-aligned in the header (skipped on the first page)
  --bookmarks        Enable automatic bookmark generation for headings
  --line-numbers     Enable Word's built-in document line numbering
  --page-numbers     Add a right-aligned page number to the footer
  --template <path>  Path to a .dotx template file to load styles from

Global Options:
  -h, --help         Show this help message
  -v, --version      Print the version number
```

<!-- END USAGE -->

## Install

### Prebuilt binary

```sh
curl -fsSL https://github.com/josefaidt/markdown-to-docx/releases/latest/download/install.sh | sh
```

### Agent skill

Install the `convert-markdown-to-docx` skill into your coding agent:

```sh
bunx skills add josefaidt/markdown-to-docx
```

## Known Limitations

- **Footnote superscripts in Word Online / SharePoint**: The `FootnoteRef` character style used for raised footnote numbers renders correctly in the Word desktop app but is ignored by Word Online. Superscripts will appear as normal-sized inline text when viewed in a browser.
