"""Calculations and mappings for I_FPC_OA -> I_Imbalance."""

from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation
from pathlib import Path


STATION_CROSSWALK = {
    "gainesville": "Univ of FL",
    "florida (gainesville)": "Univ of FL",
    "univ of fl": "Univ of FL",
    "university of florida": "Univ of FL",

    "bartow b-peaker": "Bartow",
    "bartow": "Bartow",

    "suwannee": "Suwannee",
    "hines": "Hines",

    "debary": "DeBary",
    "de bary": "DeBary",

    "anclote": "Anclote",

    "tiger bay": "Tiger Bay",
    "tiger_bay": "Tiger Bay",

    "vandolah": "Vandolah",
}


CANONICAL_STATIONS = (
    "Univ of FL",
    "Bartow",
    "Suwannee",
    "Hines",
    "DeBary",
    "Anclote",
    "Tiger Bay",
    "Vandolah",
)


def jsonable(value):
    """Convert computation values into JSON-safe values."""

    if isinstance(value, Decimal):
        return str(value)

    if isinstance(value, (date, datetime)):
        return value.isoformat()

    if isinstance(value, Path):
        return str(value)

    if isinstance(value, dict):
        return {
            str(key): jsonable(item)
            for key, item in value.items()
        }

    if isinstance(value, (list, tuple)):
        return [
            jsonable(item)
            for item in value
        ]

    return value


def normalize_station_name(station_name: str) -> str:
    """Convert an I_FPC_OA station name to its canonical name."""

    if not isinstance(station_name, str):
        raise ValueError(
            "UNMAPPED_FGT_STATION: station name must be text."
        )

    normalized_key = " ".join(
        station_name.strip().lower().split()
    )

    canonical_station = STATION_CROSSWALK.get(
        normalized_key
    )

    if canonical_station is None:
        raise ValueError(
            "UNMAPPED_FGT_STATION: "
            f"no mapping found for '{station_name}'."
        )

    return canonical_station


def to_decimal(value, field_name: str) -> Decimal:
    """Convert a source value to Decimal and reject invalid values."""

    if isinstance(value, bool):
        raise ValueError(
            f"{field_name} cannot be Boolean."
        )

    if value is None or value == "":
        raise ValueError(
            f"{field_name} is missing."
        )

    try:
        result = Decimal(str(value))
    except (InvalidOperation, ValueError, TypeError) as exc:
        raise ValueError(
            f"{field_name} is not numeric: {value!r}"
        ) from exc

    if not result.is_finite():
        raise ValueError(
            f"{field_name} must be a finite number."
        )

    return result


def normalize_station_dictionary(
    values: dict,
    field_name: str,
    strict: bool,
    exceptions: list,
    allow_blank: bool = False,
) -> dict:
    """
    Normalize station-keyed values into canonical station names.
    """

    if values is None:
        values = {}

    if not isinstance(values, dict):
        raise ValueError(
            f"{field_name} must be a station-keyed object."
        )

    canonical_values = {}
    seen = set()

    for source_station, source_value in values.items():
        try:
            canonical_station = normalize_station_name(
                source_station
            )
        except ValueError as exc:
            if strict:
                raise

            exceptions.append(
                {
                    "kind": "unmapped_station",
                    "exception_code": "UNMAPPED_FGT_STATION",
                    "station": str(source_station),
                    "field": field_name,
                    "message": str(exc),
                }
            )
            continue

        if canonical_station in seen:
            raise ValueError(
                "Duplicate source mapping for canonical station "
                f"'{canonical_station}' in {field_name}."
            )

        seen.add(canonical_station)

        try:
            if allow_blank and source_value in (None, ""):
                canonical_values[canonical_station] = None
                continue

            canonical_values[canonical_station] = to_decimal(
                source_value,
                f"{field_name}.{source_station}",
            )

        except ValueError as exc:
            exception_code = (
                "MISSING_FGT_HEAT_RATE"
                if field_name == "btu_factors"
                else "INVALID_FGT_MMBTU"
            )

            if strict:
                raise

            exceptions.append(
                {
                    "kind": "invalid_value",
                    "exception_code": exception_code,
                    "station": canonical_station,
                    "field": field_name,
                    "source_value": source_value,
                    "message": str(exc),
                }
            )

    return canonical_values


def get_first_available(record: dict, names: tuple):
    """Return the first source dictionary present in a record."""

    for name in names:
        value = record.get(name)

        if isinstance(value, dict):
            return value

    return {}


def build_station_rows(
    record: dict,
    strict: bool,
    exceptions: list,
) -> list:
    """
    Build one I_Imbalance row per canonical FGT station.

    Current draft mapping:
    - burns_dt -> MMBTU-FGT
    - btu_factors -> HEAT RATE-FGT

    Values and signs are preserved at this mapping stage.
    """

    burns = get_first_available(
        record,
        (
            "burns_dt",
            "plant_burns_dt",
            "station_burns_dt",
        ),
    )

    heat_rates = get_first_available(
        record,
        (
            "btu_factors",
            "heat_rates",
            "heat_rate_fgt",
        ),
    )

    canonical_burns = normalize_station_dictionary(
        burns,
        "burns_dt",
        strict,
        exceptions,
    )

    canonical_heat_rates = normalize_station_dictionary(
        heat_rates,
        "btu_factors",
        strict,
        exceptions,
        allow_blank=True,
    )

    station_rows = []

    for canonical_station in CANONICAL_STATIONS:
        mmbtu_value = canonical_burns.get(
            canonical_station
        )

        heat_rate_value = canonical_heat_rates.get(
            canonical_station
        )

        if mmbtu_value is None:
            exception = {
                "kind": "missing_mapping",
                "exception_code": "UNMAPPED_FGT_STATION",
                "station": canonical_station,
                "field": "MMBTU-FGT",
                "message": (
                    "No I_FPC_OA station value was found."
                ),
            }

            if strict:
                raise ValueError(
                    "UNMAPPED_FGT_STATION: "
                    f"missing MMBTU-FGT value for {canonical_station}."
                )

            exceptions.append(exception)

        # calculate() intentionally emits None for a zero-burn day's factor.
        # Preserve that blank; only a nonzero burn requires a heat factor.
        if (
            heat_rate_value is None
            and mmbtu_value != Decimal(0)
        ):
            exception = {
                "kind": "missing_heat_rate",
                "exception_code": "MISSING_FGT_HEAT_RATE",
                "station": canonical_station,
                "field": "HEAT RATE-FGT",
                "message": (
                    "No I_FPC_OA heat-rate value was found."
                ),
            }

            if strict:
                raise ValueError(
                    "MISSING_FGT_HEAT_RATE: "
                    f"missing heat rate for {canonical_station}."
                )

            exceptions.append(exception)

        station_rows.append(
            {
                "station": canonical_station,
                "pipeline": "FGT",
                "MMBTU-FGT": mmbtu_value,
                "HEAT RATE-FGT": heat_rate_value,
            }
        )

    return station_rows


def calculate_i_imbalance(
    i_fpc_oa: dict,
    strict: bool = True,
) -> dict:
    """
    Transform calculated I_FPC_OA records into I_Imbalance records.
    """

    if not isinstance(i_fpc_oa, dict):
        raise ValueError(
            "I_FPC_OA payload must be an object."
        )

    if i_fpc_oa.get("schema_version") != "1":
        raise ValueError(
            "Unsupported I_FPC_OA schema_version"
        )

    if i_fpc_oa.get("target", {}).get("type") == "I_Imbalance":
        raise ValueError(
            "Expected I_FPC_OA input, not an I_Imbalance result"
        )

    upstream_exceptions = i_fpc_oa.get("exceptions", [])

    if not isinstance(upstream_exceptions, list):
        raise ValueError(
            "I_FPC_OA exceptions must be a list"
        )

    if upstream_exceptions and strict:
        raise ValueError(
            "I_FPC_OA has upstream exceptions; review the first flow"
        )

    source_records = i_fpc_oa.get("records")

    if not isinstance(source_records, list):
        raise ValueError(
            "I_FPC_OA payload must contain a records list."
        )

    if not source_records:
        raise ValueError(
            "I_FPC_OA records list is empty."
        )

    exceptions = [
        {
            "kind": "upstream_exception",
            "source_exception": item,
        }
        for item in upstream_exceptions
    ]

    imbalance_records = []
    previous_date = None

    for source_index, source_record in enumerate(
        source_records,
        start=1,
    ):
        if not isinstance(source_record, dict):
            raise ValueError(
                f"I_FPC_OA record {source_index} must be an object."
            )

        record_date = source_record.get("date")

        if record_date is None:
            raise ValueError(
                f"I_FPC_OA record {source_index} has no date."
            )

        if isinstance(record_date, datetime):
            record_day = record_date.date()

        elif isinstance(record_date, date):
            record_day = record_date

        else:
            try:
                record_day = date.fromisoformat(
                    str(record_date)
                )
            except ValueError as exc:
                raise ValueError(
                    f"Invalid I_FPC_OA date: {record_date!r}"
                ) from exc

        record_date_text = record_day.isoformat()

        if (
            previous_date is not None
            and record_day != previous_date + timedelta(days=1)
        ):
            raise ValueError(
                "Dates must be unique, ascending and contiguous."
            )

        if (
            imbalance_records
            and record_date_text[:7]
            != imbalance_records[0]["date"][:7]
        ):
            raise ValueError(
                "Dates must belong to one reporting month"
            )

        exception_start = len(exceptions)

        station_rows = build_station_rows(
            source_record,
            strict,
            exceptions,
        )

        for item in exceptions[exception_start:]:
            item["date"] = record_date_text
            item["source_record_number"] = source_index

        saved_daily_total = source_record.get(
            "total_plant_burns_dt"
        )

        mapped_burns = [
            row["MMBTU-FGT"]
            for row in station_rows
        ]

        if (
            saved_daily_total is not None
            and all(v is not None for v in mapped_burns)
        ):
            if (
                sum(mapped_burns, Decimal(0))
                != to_decimal(
                    saved_daily_total,
                    "daily plant total",
                )
            ):
                if strict:
                    raise ValueError(
                        "FGT daily burns do not reconcile "
                        f"on {record_date_text}"
                    )

                exceptions.append(
                    {
                        "kind": "daily_total_mismatch",
                        "date": record_date_text,
                    }
                )

        imbalance_records.append(
            {
                "date": record_date_text,
                "pipeline": "FGT",
                "stations": station_rows,
                "source_record_number": source_index,
                "source": source_record.get("source"),
                "lookup_sources": source_record.get(
                    "lookup_sources",
                    {},
                ),
            }
        )

        previous_date = record_day

    ending_imbalance = (
        i_fpc_oa.get("summary", {}).get(
            "ending_imbalance_dt"
        )
    )

    if ending_imbalance is not None:
        ending_imbalance = to_decimal(
            ending_imbalance,
            "summary.ending_imbalance_dt",
        )

    monthly_stations = []

    for station_index, station in enumerate(
        CANONICAL_STATIONS
    ):
        rows = [
            record["stations"][station_index]
            for record in imbalance_records
        ]

        burns = [
            row["MMBTU-FGT"]
            for row in rows
        ]

        factors = [
            row["HEAT RATE-FGT"]
            for row in rows
            if row["HEAT RATE-FGT"] is not None
        ]

        total = (
            sum(burns, Decimal(0))
            if all(v is not None for v in burns)
            else None
        )

        invalid_heat = any(
            row["MMBTU-FGT"] != Decimal(0)
            and row["HEAT RATE-FGT"] is None
            for row in rows
        )

        average = (
            sum(factors, Decimal(0)) / len(factors)
            if factors
            else None
        )

        if invalid_heat:
            average = None

        monthly_stations.append(
            {
                "station": station,
                "pipeline": "FGT",
                "MMBTU-FGT": total,
                "HEAT RATE-FGT": average,
            }
        )

    totals = [
        row["MMBTU-FGT"]
        for row in monthly_stations
    ]

    total_fgt = (
        sum(totals, Decimal(0))
        if all(v is not None for v in totals)
        else None
    )

    source_total = (
        i_fpc_oa.get("summary", {}).get(
            "total_plant_burns_dt"
        )
    )

    reconciled = None

    if source_total is not None:
        source_total = to_decimal(
            source_total,
            "summary.total_plant_burns_dt",
        )

        reconciled = (
            total_fgt is not None
            and total_fgt == source_total
        )

        if not reconciled:
            if strict:
                raise ValueError(
                    "FGT monthly burns do not reconcile "
                    "to I_FPC_OA summary"
                )

            exceptions.append(
                {
                    "kind": "source_total_mismatch",
                    "source_total": source_total,
                    "mapped_total": total_fgt,
                }
            )

    upstream_controls = i_fpc_oa.get("control", {})

    if upstream_controls.get("rollforward_balanced") is False:
        if strict:
            raise ValueError(
                "I_FPC_OA rollforward control failed"
            )

        exceptions.append(
            {
                "kind": "upstream_control_failed",
            }
        )

    return {
        "schema_version": "1",
        "rule_version": "i-imbalance-fgt-v2",
        "scope": "FGT_ONLY",
        "complete_fgt": not exceptions,
        "complete_workflow": False,
        "unavailable_sections": [
            "Gulfstream",
            "Osprey",
            "Citrus / Sabal Trail",
        ],
        "source": {
            "type": "I_FPC_OA",
            "upstream": i_fpc_oa.get(
                "source",
                {},
            ),
        },
        "target": {
            "type": "I_Imbalance",
        },
        "records": imbalance_records,
        "monthly_stations": monthly_stations,
        "reporting_period": (
            imbalance_records[0]["date"][:7]
        ),
        "summary": {
            "source_records": len(source_records),
            "target_date_records": len(
                imbalance_records
            ),
            "target_station_records": (
                len(imbalance_records)
                * len(CANONICAL_STATIONS)
            ),
            "ending_imbalance_dt": ending_imbalance,
            "exception_count": len(exceptions),
            "total_fgt_mmbtu": total_fgt,
        },
        "control": {
            "mapping_stage": (
                "I_FPC_OA -> I_Imbalance"
            ),
            "sign_rule": "preserve_source_value",
            "station_count_per_date": len(
                CANONICAL_STATIONS
            ),
            "strict": strict,
            "source_total_reconciled": reconciled,
            "heat_rate_rule": (
                "Arithmetic mean of available daily factors; "
                "zero-burn blanks excluded"
            ),
        },
        "exceptions": exceptions,
    }