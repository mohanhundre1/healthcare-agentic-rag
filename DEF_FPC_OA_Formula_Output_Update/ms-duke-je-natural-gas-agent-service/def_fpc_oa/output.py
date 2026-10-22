"""
Creates I_FPC_OA file.
"""

import math
import re

from datetime import date
from decimal import Decimal
from pathlib import Path

from openpyxl import Workbook

from .calculations import PLANTS, BTU, jsonable
from .seed_reference import Workbook as SeedWorkbook
from .client_layout import apply_client_layout, apply_value_formats
from .formula_workbook import populate_dependencies, finalize_formula_workbook


def cell_updates(result):
    """Build calculated I_FPC OA cell values."""

    updates = {
        "Q12": result["opening_imbalance_dt"]
    }

    for col, station in result["station_ids"].items():
        updates[f"{col}6"] = int(station)

    for row, record in enumerate(result["records"], 13):

        day = date.fromisoformat(record["date"])

        cells = {
            "A": (day.weekday() + 1) % 7 + 1,
            "B": day,
            "C": record["scheduled_volume_dt"],
            "D": record["alert_day"],
            "E": record["alert_direction"],
            "N": record["total_plant_burns_dt"],
            "O": record["daily_imbalance_dt"],
            "P": record["tolerance_volume_dt"],
            "Q": record["cumulative_imbalance_dt"],
            "AA": day,
            "AD": record["fgt_peaker_burn_dt"],
            "AE": record["gulfstream_peaker_burn_dt"],
            "AF": record["fpc_peaker_burn_dt"],
        }

        cells.update({
            col: record["burns_dt"][plant]
            for col, plant in PLANTS.items()
        })

        cells.update({
            col: record["btu_factors"][PLANTS[source_col]]
            for col, source_col in BTU.items()
        })

        updates.update({
            f"{col}{row}": value
            for col, value in cells.items()
        })

    summary = result["summary"]

    updates.update({
        "C45": summary["scheduled_volume_dt"],
        "N45": summary["total_plant_burns_dt"],
        "F47": summary["gru_therms"],
        "L47": summary["ending_imbalance_dt"],
        "P53": "FGT Rolling imb",
        "Q53": summary["fgt_rolling_imbalance_dt"],
        "Q54": summary["rolling_less_opening_dt"],
    })

    updates.update({
        f"{col}45": summary["plant_totals_dt"][plant]
        for col, plant in PLANTS.items()
    })

    updates.update({
        f"{col}45": summary["average_btu_factors"][PLANTS[source_col]]
        for col, source_col in BTU.items()
        if col != "X"
    })

    return updates


def compare_seed(result, seed_path):
    """Compare calculated results against saved FPC OA seed values."""

    source = SeedWorkbook(seed_path)

    differences = []
    checked = 0

    for address, value in cell_updates(result).items():

        if re.fullmatch(r"(A|B|D|E|AA)\d+", address) or address == "P53":
            continue

        expected = source.get("FPC OA", address)

        if value is None and expected in (None, ""):
            equal = True

        elif (
            isinstance(value, (int, float, Decimal))
            and isinstance(expected, (int, float))
        ):
            equal = math.isclose(
                float(value),
                expected,
                rel_tol=1e-10,
                abs_tol=1e-6,
            )

        else:
            equal = value == expected

        checked += 1

        if not equal:
            differences.append({
                "cell": address,
                "calculated": jsonable(value),
                "saved": expected,
            })

    return {
        "cells_checked": checked,
        "differences": differences,
        "matches_saved_seed": not differences,
    }


def write_output(output: Path, result: dict, *, validated_input: dict = None,
                 recalc_engine: str = None):
    """Create a formula workbook and verify native saved calculation results.

    Raw validated input is required because rounded output BTUs cannot recreate
    source lookup factors. Never publish a workbook with missing formula caches.
    """
    output = Path(output)
    if output.exists():
        raise FileExistsError(f"Output already exists: {output}")
    if output.suffix.lower() != ".xlsx":
        raise ValueError("Formula output must use .xlsx")
    if len(result["records"]) != 31:
        raise ValueError("Current FPC OA adapter requires 31 daily records")
    if validated_input is None:
        raise ValueError("validated_input is required for formula dependencies")

    # Protect against pairing source records with a different calculated result.
    from .calculations import calculate
    if calculate(validated_input, strict=True) != result:
        raise ValueError("Calculated result does not match validated_input")

    from .client_formulas import formula_updates
    formulas = formula_updates()
    workbook = Workbook()
    try:
        worksheet = workbook.active
        worksheet.title = "I_FPC OA"
        apply_client_layout(worksheet, result)
        for address, value in cell_updates(result).items():
            if re.fullmatch(r"D(?:1[3-9]|[23][0-9]|4[0-3])", address):
                value = None if value in (None, "", " ") else float(Decimal(str(value)))
            worksheet[address] = float(value) if isinstance(value, Decimal) else value
        apply_value_formats(worksheet, result)
        populate_dependencies(workbook, validated_input)
        # Formulas are written LAST: formats must never overwrite formulas.
        for address, formula in formulas.items():
            worksheet[address] = formula
        return finalize_formula_workbook(
            workbook, output, result, formulas, recalc_engine=recalc_engine)
    finally:
        workbook.close()
