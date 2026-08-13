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
  --page-size <size> Page size, named or custom (default: a4)
  --bookmarks        Enable automatic bookmark generation for headings
  --landscape        Turn the page on its side, so its longer edge runs horizontally
  --line-numbers     Enable Word's built-in document line numbering
  --page-numbers     Add a right-aligned page number to the footer
  --template <path>  Path to a .dotx template file to load styles from

Page Sizes:
  Named             a3, a4, a5, b5, executive, legal, letter, tabloid
  Custom            <width>x<height> with an optional unit: in (default), mm, cm, pt
                    e.g. --page-size 9x12, --page-size 9x12in, --page-size 210x297mm

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

The following features render correctly in the Word desktop app but behave differently in Word Online (SharePoint):

- **Footnote superscripts**: The raised superscript style on footnote reference numbers is ignored by Word Online — numbers appear inline at normal size.
- **Footer page numbers**: The right-aligned tab stop used to position page numbers in the footer is not honored by Word Online, so the page number does not align to the right margin.
- **Line number styles**: Word Online always renders line numbers in Times New Roman regardless of the font set in the `lineNumber` character style.
