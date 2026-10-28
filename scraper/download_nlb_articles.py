from __future__ import annotations

import argparse
import re
import zipfile
from pathlib import Path
from time import sleep
from urllib.parse import urljoin, urlencode, urlsplit
from urllib.request import Request, urlopen
from xml.etree import ElementTree


OUTPUT_ROOT = Path(__file__).resolve().parent.parent / "data" / "images"
NLB_ARTICLE_PATH = "/newspapers/digitised/article/"
URL_PATTERN = re.compile(r"https?://[^\s<>\"']+", re.IGNORECASE)
RELATIONSHIPS_NAMESPACE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
)


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def column_number(column_letters: str) -> int:
    number = 0
    for letter in column_letters.upper():
        if not "A" <= letter <= "Z":
            raise ValueError(f"Invalid Excel column: {column_letters}")
        number = number * 26 + ord(letter) - ord("A") + 1
    return number


def cell_column(cell_reference: str) -> int:
    match = re.match(r"([A-Z]+)", cell_reference.upper())
    if not match:
        raise ValueError(f"Invalid cell reference: {cell_reference}")
    return column_number(match.group(1))


def element_text(element: ElementTree.Element | None) -> str:
    if element is None:
        return ""
    return "".join(element.itertext())


def load_shared_strings(workbook: zipfile.ZipFile) -> list[str]:
    try:
        shared_strings = ElementTree.fromstring(
            workbook.read("xl/sharedStrings.xml")
        )
    except KeyError:
        return []

    return [
        element_text(string_item)
        for string_item in shared_strings
        if local_name(string_item.tag) == "si"
    ]


def relationship_targets(
    workbook: zipfile.ZipFile, relationships_path: str
) -> dict[str, str]:
    try:
        relationships = ElementTree.fromstring(workbook.read(relationships_path))
    except KeyError:
        return {}

    return {
        relationship.attrib["Id"]: relationship.attrib["Target"]
        for relationship in relationships
        if local_name(relationship.tag) == "Relationship"
    }


def resolve_relationship_path(source_path: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")

    source_directory = source_path.rsplit("/", 1)[0]
    parts = (source_directory + "/" + target).split("/")
    resolved = []
    for part in parts:
        if part in ("", "."):
            continue
        if part == "..":
            if resolved:
                resolved.pop()
        else:
            resolved.append(part)
    return "/".join(resolved)


def read_xlsx_rows(workbook_path: Path) -> dict[str, list[dict[int, str]]]:
    """Return workbook rows without requiring an external Excel package."""
    with zipfile.ZipFile(workbook_path) as workbook:
        workbook_xml = ElementTree.fromstring(workbook.read("xl/workbook.xml"))
        workbook_relationships = relationship_targets(
            workbook, "xl/_rels/workbook.xml.rels"
        )
        shared_strings = load_shared_strings(workbook)
        sheets = {}

        for sheet in workbook_xml.iter():
            if local_name(sheet.tag) != "sheet":
                continue

            sheet_name = sheet.attrib["name"]
            relationship_id = sheet.attrib[f"{{{RELATIONSHIPS_NAMESPACE}}}id"]
            sheet_path = resolve_relationship_path(
                "xl/workbook.xml", workbook_relationships[relationship_id]
            )
            sheet_xml = ElementTree.fromstring(workbook.read(sheet_path))
            sheet_relationships = relationship_targets(
                workbook,
                resolve_relationship_path(
                    sheet_path,
                    f"_rels/{sheet_path.rsplit('/', 1)[-1]}.rels",
                ),
            )

            hyperlink_targets = {}
            for hyperlink in sheet_xml.iter():
                if local_name(hyperlink.tag) != "hyperlink":
                    continue
                hyperlink_relationship_id = hyperlink.attrib.get(
                    f"{{{RELATIONSHIPS_NAMESPACE}}}id"
                )
                if hyperlink_relationship_id in sheet_relationships:
                    hyperlink_targets[hyperlink.attrib["ref"]] = urljoin(
                        "https://eresources.nlb.gov.sg/",
                        sheet_relationships[hyperlink_relationship_id],
                    )

            rows = []
            for row_element in sheet_xml.iter():
                if local_name(row_element.tag) != "row":
                    continue

                row = {}
                for cell in row_element:
                    if local_name(cell.tag) != "c":
                        continue

                    reference = cell.attrib.get("r", "")
                    if not reference:
                        continue

                    value_element = next(
                        (
                            child
                            for child in cell
                            if local_name(child.tag) in {"v", "is"}
                        ),
                        None,
                    )
                    value = element_text(value_element)
                    cell_type = cell.attrib.get("t")

                    if cell_type == "s" and value:
                        value = shared_strings[int(value)]
                    elif cell_type == "b":
                        value = "TRUE" if value == "1" else "FALSE"

                    hyperlink = hyperlink_targets.get(reference)
                    if hyperlink and hyperlink not in value:
                        value = f"{value} {hyperlink}".strip()

                    row[cell_column(reference)] = value

                rows.append(row)

            sheets[sheet_name] = rows

    return sheets


def normalise_header(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", value.strip().lower())


def find_url_column(rows: list[dict[int, str]], selector: str) -> int:
    wanted = normalise_header(selector)
    for row in rows[:10]:
        for column, value in row.items():
            if normalise_header(value) == wanted:
                return column

    if re.fullmatch(r"[A-Za-z]+", selector):
        return column_number(selector)

    raise ValueError(
        f"Could not find a column named {selector!r} in the first 10 rows"
    )


def article_urls_from_value(value: str) -> list[str]:
    urls = []
    for match in URL_PATTERN.findall(value):
        url = match.rstrip(".,;)]}")
        if NLB_ARTICLE_PATH in urlsplit(url).path.lower():
            urls.append(url)
    return urls


def read_article_urls(
    workbook_path: Path,
    sheet_name: str | None = None,
    url_column: str | None = None,
) -> list[str]:
    if workbook_path.suffix.lower() != ".xlsx":
        raise ValueError(
            "Only .xlsx workbooks are supported. Save the Excel file as .xlsx "
            "before running this script."
        )

    sheets = read_xlsx_rows(workbook_path)
    if sheet_name is not None:
        if sheet_name not in sheets:
            available = ", ".join(sheets) or "none"
            raise ValueError(
                f"Worksheet {sheet_name!r} was not found. Available sheets: {available}"
            )
        sheets = {sheet_name: sheets[sheet_name]}

    article_urls = []
    seen = set()
    for rows in sheets.values():
        selected_column = find_url_column(rows, url_column) if url_column else None
        for row in rows:
            values = (
                [row.get(selected_column, "")]
                if selected_column is not None
                else row.values()
            )
            for value in values:
                for url in article_urls_from_value(str(value)):
                    if url not in seen:
                        seen.add(url)
                        article_urls.append(url)

    if not article_urls:
        scope = f" in worksheet {sheet_name!r}" if sheet_name else ""
        raise ValueError(f"No NLB article URLs were found{scope} in {workbook_path}")

    return article_urls


def download_article(
    article_url: str,
    output_root: Path,
    max_areas: int | None = None,
) -> list[Path]:
    if max_areas is not None and max_areas <= 0:
        raise ValueError("max_areas must be a positive integer")

    article_id = urlsplit(article_url).path.rstrip("/").rsplit("/", 1)[-1]
    if not article_id:
        raise ValueError(f"Could not determine an article ID from {article_url}")

    article_dir = output_root / article_id
    article_dir.mkdir(parents=True, exist_ok=True)
    endpoint = (
        "https://eservice.nlb.gov.sg/newspapercontent/digitised/article/"
        f"{article_id}.webp"
    )
    headers = {
        "User-Agent": "Mozilla/5.0",
        "Referer": article_url,
    }
    saved = []

    area_limit = 99 if max_areas is None else min(max_areas, 99)
    for area in range(1, area_limit + 1):
        output_path = article_dir / f"{article_id}-area-{area}.webp"
        if output_path.exists():
            saved.append(output_path)
            continue

        query = urlencode(
            {
                "area": area,
                "width": 2000,
                "ct": "ARTICLE",
                "ns": "",
                "coord": "",
            }
        )
        request = Request(f"{endpoint}?{query}", headers=headers)

        with urlopen(request, timeout=120) as response:
            content_type = response.headers.get_content_type()
            image_data = response.read()

        if content_type != "image/webp":
            break

        output_path.write_bytes(image_data)
        saved.append(output_path)
        print(output_path)
        sleep(0.5)

    if not saved:
        raise RuntimeError(f"No images found for {article_url}")

    return saved


def parse_arguments() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Download NLB article images for the URLs in one worksheet of an "
            ".xlsx workbook. Run without --sheet to list the worksheets."
        )
    )
    parser.add_argument(
        "workbook",
        type=Path,
        help="Path to the .xlsx workbook containing NLB article URLs",
    )
    parser.add_argument(
        "--sheet",
        help="Worksheet to download. Also used as the folder name under --output-root.",
    )
    parser.add_argument(
        "--url-column",
        default="Permalink",
        help=(
            "Excel column letter or header containing article URLs "
            "(default: %(default)s)."
        ),
    )
    parser.add_argument(
        "--output-root",
        type=Path,
        default=OUTPUT_ROOT,
        help=f"Folder for downloaded images (default: {OUTPUT_ROOT})",
    )
    return parser.parse_args()


def main() -> None:
    arguments = parse_arguments()
    workbook_path = arguments.workbook.expanduser().resolve()
    if arguments.sheet is None:
        for name in read_xlsx_rows(workbook_path):
            try:
                count = len(
                    read_article_urls(workbook_path, name, arguments.url_column)
                )
                print(f"{count:5d} URLs  {name}")
            except ValueError:  # no URL column, or no URLs in this sheet
                print(f"  n/a URLs  {name}")
        return

    # Excel forbids / \ : ? * [ ] in sheet names, so this is a safe folder name.
    output_root = arguments.output_root.expanduser().resolve() / arguments.sheet
    article_urls = read_article_urls(
        workbook_path,
        sheet_name=arguments.sheet,
        url_column=arguments.url_column,
    )
    print(f"Found {len(article_urls)} unique NLB article URL(s).")

    all_paths = []
    for article_url in article_urls:
        all_paths.extend(download_article(article_url, output_root))

    manifest = output_root / "image_paths.txt"
    manifest.write_text(
        "\n".join(str(path.resolve()) for path in all_paths) + "\n",
        encoding="utf-8",
    )
    print(f"Manifest: {manifest}")


if __name__ == "__main__":
    main()
