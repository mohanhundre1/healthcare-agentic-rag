"""Read-only OOXML inspection for formula-cache validation and output checks."""

import math
import posixpath
import xml.etree.ElementTree as ET

from pathlib import Path
from zipfile import ZipFile


NS = {
    "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
}

RID = (
    "{http://schemas.openxmlformats.org/"
    "officeDocument/2006/relationships}id"
)


class Workbook:
    """Small read-only OOXML reader for this supplied workbook layout."""

    def __init__(self, path: Path):
        self.sheets = {}

        with ZipFile(path) as archive:

            # Read workbook metadata
            book = ET.fromstring(
                archive.read("xl/workbook.xml")
            )

            # Validate Excel date system
            props = book.find("m:workbookPr", NS)

            if (
                props is not None
                and props.get("date1904") in ("1", "true")
            ):
                raise ValueError(
                    "This seed adapter requires the Excel 1900 date system"
                )

            # Read shared strings
            shared = []

            if "xl/sharedStrings.xml" in archive.namelist():
                shared_xml = ET.fromstring(
                    archive.read("xl/sharedStrings.xml")
                )

                for item in shared_xml:
                    shared.append(
                        "".join(
                            t.text or ""
                            for t in item.findall(".//m:t", NS)
                        )
                    )

            # Read workbook relationships
            relationships_xml = ET.fromstring(
                archive.read(
                    "xl/_rels/workbook.xml.rels"
                )
            )

            rels = {
                relationship.get("Id"): relationship.get("Target")
                for relationship in relationships_xml
            }

            # Read every worksheet
            sheets = book.find("m:sheets", NS)

            if sheets is None:
                raise ValueError(
                    "Workbook does not contain worksheets"
                )

            for sheet in sheets:

                target = rels[sheet.get(RID)]

                if target.startswith("/"):
                    location = target.lstrip("/")
                else:
                    location = posixpath.normpath(
                        "xl/" + target
                    )

                cells = {}

                worksheet_xml = ET.fromstring(
                    archive.read(location)
                )

                worksheet_cells = worksheet_xml.findall(
                    ".//m:sheetData/m:row/m:c",
                    NS,
                )

                for cell in worksheet_cells:

                    value = cell.find("m:v", NS)

                    text = (
                        value.text
                        if value is not None
                        else None
                    )

                    kind = cell.get("t")

                    # Shared string
                    if kind == "s" and text is not None:
                        parsed = shared[int(text)]

                    # Inline string
                    elif kind == "inlineStr":
                        parsed = "".join(
                            t.text or ""
                            for t in cell.findall(
                                ".//m:t",
                                NS,
                            )
                        )

                    # Formula string / Excel error
                    elif kind in ("str", "e"):
                        parsed = text

                    # Empty cell
                    elif text is None:
                        parsed = None

                    # Numeric cell
                    else:
                        parsed = float(text)

                    cells[cell.get("r")] = (
                        parsed,
                        cell.find("m:f", NS) is not None,
                        kind,
                    )

                self.sheets[sheet.get("name")] = cells

    def get(self, sheet: str, address: str):
        """Return the saved value for a workbook cell."""

        if sheet not in self.sheets:
            raise ValueError(
                f"Worksheet not found: {sheet}"
            )

        value, formula, kind = self.sheets[sheet].get(
            address,
            (None, False, None),
        )

        if kind == "e":
            raise ValueError(
                f"Excel error at {sheet}!{address}: {value}"
            )

        if formula and value is None and kind != "str":
            raise ValueError(
                f"Missing saved formula value at "
                f"{sheet}!{address}; recalculate in Excel"
            )

        return value


def number(value, reference: str) -> float:
    """Validate and return a finite numeric value."""

    if (
        isinstance(value, bool)
        or not isinstance(value, (int, float))
        or not math.isfinite(value)
    ):
        raise ValueError(
            f"Expected finite number at "
            f"{reference}, got {value!r}"
        )

    return value


def numeric(
    workbook: Workbook,
    sheet: str,
    address: str,
) -> float:
    """Read and validate a numeric workbook value."""

    return number(
        workbook.get(sheet, address),
        f"{sheet}!{address}",
    )


def excel_text(value) -> str:
    """Convert an Excel value to its expected text representation."""

    if (
        isinstance(value, float)
        and value.is_integer()
    ):
        return str(int(value))

    return str(value)