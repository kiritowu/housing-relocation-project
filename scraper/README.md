# Scraper

`download_nlb_articles.py` downloads the scanned images of historical newspaper articles from the [NLB NewspaperSG](https://eresources.nlb.gov.sg/newspapers/) digital archive. It reads the article links from the source spreadsheets in `data/sources/`, one worksheet at a time.

It uses only the Python standard library, so there is nothing to install.

## How it works

1. Opens the `.xlsx` workbook and reads the chosen worksheet. Hyperlinks hidden behind cell text are picked up too.
2. Collects every unique NLB article URL (`/newspapers/digitised/article/...`).
3. For each article, downloads its image areas (the crops NLB serves for the article) as `.webp` files.
4. Writes `image_paths.txt`, a list of every downloaded image, for the OCR step.

## Usage

Run from the repository root.

**1. List the worksheets in a workbook** (with the number of article URLs in each):

```bash
python scraper/download_nlb_articles.py "data/sources/祖屋 1970-1979.xlsx"
```

**2. Download one worksheet.** The sheet name is used as the folder name, so use a name from the list above (here, a year from `1970` to `1979`):

```bash
python scraper/download_nlb_articles.py "data/sources/祖屋 1970-1979.xlsx" --sheet 1977
```

### Options

| Argument | Description |
|----------|-------------|
| `workbook` | Path to the `.xlsx` file (required). Only `.xlsx` is supported. |
| `--sheet` | Worksheet to download. If omitted, the worksheets are listed and nothing is downloaded. |
| `--output-root` | Where to save images. Default: `data/images`. |
| `--url-column` | Column letter or header holding the URLs. Default: `Permalink`, so only that column is read. |

## Output

```
data/images/
└── <sheet name>/
    ├── image_paths.txt
    └── <article_id>/
        ├── <article_id>-area-1.webp
        ├── <article_id>-area-2.webp
        └── ...
```

`data/images/` is gitignored, so downloaded images are not committed.

## Known limitations

- A failed request stops the whole run, and re-running downloads everything again.
- Article metadata (title, date, remarks) in the spreadsheets is not saved alongside the images.
- The same article can appear in more than one worksheet and will be downloaded once per sheet.
