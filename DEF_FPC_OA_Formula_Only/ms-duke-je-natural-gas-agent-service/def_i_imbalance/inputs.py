"""Read calculated FPC JSON or this service's fixed-layout I_FPC_OA.xlsx."""

import json
from datetime import date, datetime, timedelta
from hashlib import sha256
from pathlib import Path

from def_fpc_oa.calculations import PLANTS, BTU
from def_fpc_oa.seed_reference import Workbook


def _find_calculated_payload(payload):
    """Accept direct JSON, known envelopes, or a single A2A result list."""

    found = []

    def visit(value, depth=0):
        if depth > 8:
            raise ValueError(
                "I_FPC_OA result nesting is too deep"
            )

        if isinstance(value, list):
            for item in value:
                visit(item, depth + 1)

        elif isinstance(value, dict):
            if value.get("success") is False:
                raise ValueError(
                    "Upstream agent reported failure"
                )

            if isinstance(value.get("records"), list):
                found.append(value)
                return

            for key in (
                "output",
                "result",
                "data",
                "calculated_result",
                "i_fpc_oa",
                "I_FPC_OA",
            ):
                candidate = value.get(key)

                if isinstance(candidate, (dict, list)):
                    visit(candidate, depth + 1)

    visit(payload)

    if len(found) != 1:
        raise ValueError(
            "Expected exactly one calculated I_FPC_OA result; "
            "use calculated_I_FPC_OA.json or I_FPC_OA.xlsx"
        )

    if not found[0]["records"]:
        raise ValueError(
            "I_FPC_OA input contains no records"
        )

    return found[0]


def _read_workbook(
    path,
    *,
    workbook_opener=None,
):
    """
    Read the fixed-layout I_FPC_OA worksheet.

    The supported worksheet area is A1:AF54.
    """

    if workbook_opener is None:
        from ms_duke_je_common.normalization.extraction.workbook_reader import (
            open_workbook,
        )

        workbook_opener = open_workbook

    metadata = Workbook(path)
    wb = workbook_opener(path)

    try:
        if "I_FPC OA" not in wb.sheetnames:
            raise ValueError(
                "Workbook must contain I_FPC OA"
            )

        sheet = wb["I_FPC OA"]

        # Materialize only the service's fixed A1:AF54 block once.
        values = {
            cell.coordinate: cell.value
            for row in sheet.iter_rows(
                min_row=1,
                max_row=54,
                max_col=32,
            )
            for cell in row
            if cell.value is not None
        }

    finally:
        wb.close()

    def get(address):
        """
        Return a worksheet value after validating the source cell.
        """

        # Reject Excel errors and uncached formulas.
        metadata.get("I_FPC OA", address)

        return values.get(address)

    def day_text(value):
        """Convert an Excel or Python date value to ISO format."""

        if isinstance(value, datetime):
            return value.date().isoformat()

        if isinstance(value, date):
            return value.isoformat()

        if (
            isinstance(value, (int, float))
            and not isinstance(value, bool)
        ):
            return (
                date(1899, 12, 30)
                + timedelta(days=value)
            ).isoformat()

        if isinstance(value, str):
            try:
                return date.fromisoformat(
                    value
                ).isoformat()

            except ValueError:
                try:
                    return datetime.strptime(
                        value,
                        "%m/%d/%Y",
                    ).date().isoformat()

                except ValueError:
                    pass

        raise ValueError(
            f"Invalid I_FPC OA daily date: {value!r}"
        )

    records = []

    for row in range(13, 44):
        records.append(
            {
                "date": day_text(
                    get(f"B{row}")
                ),
                "burns_dt": {
                    plant: get(f"{col}{row}")
                    for col, plant in PLANTS.items()
                },
                "btu_factors": {
                    PLANTS[source]: get(f"{col}{row}")
                    for col, source in BTU.items()
                },
                "total_plant_burns_dt": get(
                    f"N{row}"
                ),
                "source": {
                    "file": path.name,
                    "sheet": "I_FPC OA",
                    "row": row,
                },
            }
        )

    if get("N45") is None:
        raise ValueError(
            "Missing I_FPC OA!N45 monthly burn total"
        )

    return {
        "schema_version": "1",
        "records": records,
        "source": {
            "file": path.name,
            "sha256": sha256(
                path.read_bytes()
            ).hexdigest(),
            "adapter": "service-I_FPC-OA-A1-AF54",
        },
        "summary": {
            "total_plant_burns_dt": get("N45"),
            "ending_imbalance_dt": get("L47"),
        },
    }


def read_i_fpc_oa(
    path: Path,
    *,
    workbook_opener=None,
) -> dict:
    """
    Read calculated I_FPC_OA data from JSON or XLSX.

    Supported inputs:
    - calculated I_FPC_OA JSON
    - service-generated I_FPC_OA XLSX
    """

    path = Path(path)

    if not path.is_file():
        raise FileNotFoundError(
            f"I_FPC_OA input file not found: {path}"
        )

    if path.suffix.lower() == ".xlsx":
        return _read_workbook(
            path,
            workbook_opener=workbook_opener,
        )

    if path.suffix.lower() != ".json":
        raise ValueError(
            "Expected calculated I_FPC_OA .json "
            "or service-generated .xlsx"
        )

    try:
        payload = json.loads(
            path.read_text(
                encoding="utf-8-sig"
            )
        )

    except (json.JSONDecodeError, UnicodeError) as exc:
        raise ValueError(
            f"Invalid I_FPC_OA JSON: {path}"
        ) from exc

    return _find_calculated_payload(payload)
