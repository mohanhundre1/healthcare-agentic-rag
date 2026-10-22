
import argparse
import asyncio
import json

from datetime import date, datetime
from decimal import Decimal
from pathlib import Path

from ms_duke_je_common.context.execution import (
    ExecutionContext,
    RequestContext,
    ResponseContext,
    ServiceContainer,
)
from ms_duke_je_common.context.security import SecurityContext
from ms_duke_je_common.executor.request_pipeline import RequestPipeline

from def_fpc_oa.calculations import calculate, jsonable
from def_fpc_oa.inputs import read_seed
from def_fpc_oa.logic import FpcOaLogic
from def_fpc_oa.output import compare_seed, write_output


async def execute(payload):
    """
    Execute FPC OA logic through the common RequestPipeline.
    """

    request_context = RequestContext(
        "local",
        "local",
        "def-fpc-oa-agent",
        "fpc_oa",
        "",
        {
            "validated_input": payload
        },
        SecurityContext(
            "local",
            "local",
            "Local development",
        ),
        "local",
    )

    context = ExecutionContext(
        request_context,
        ResponseContext(),
        ServiceContainer(),
    )

    pipeline = RequestPipeline(
        FpcOaLogic(),
        [],
    )

    await pipeline.startup(
        context.services
    )

    try:
        await pipeline.run(
            context
        )
    finally:
        await pipeline.shutdown()

    return context.response.result


def json_serializer(value):
    """
    Convert Python objects that are not directly JSON serializable.
    """

    if isinstance(value, (datetime, date)):
        return value.isoformat()

    if isinstance(value, Decimal):
        return str(value)

    if isinstance(value, Path):
        return str(value)

    raise TypeError(
        f"Object of type {type(value).__name__} "
        "is not JSON serializable"
    )


def write_json(path: Path, value):
    """
    Write formatted JSON output safely.
    """

    path.write_text(
        json.dumps(
            jsonable(value),
            indent=2,
            allow_nan=False,
            default=json_serializer,
        ),
        encoding="utf-8",
    )


def main():
    parser = argparse.ArgumentParser(
        description=(
            "Generate I_FPC_OA calculation output "
            "from FPC OA input without an Excel template."
        )
    )

    # ---------------------------------------------------------
    # Input arguments
    # ---------------------------------------------------------

    parser.add_argument(
        "--seed",
        type=Path,
        required=True,
        help="Path to the FPC OA seed/input Excel file.",
    )

    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path("output"),
        help="Directory where generated files will be written.",
    )

    args = parser.parse_args()

    output_names = (
        "I_FPC_OA.xlsx",
        "I_FPC_OA.json",
        "calculated_I_FPC_OA.json",
        "validated_input.json",
        "validation_report.json",
        "validation_failed.json",
    )

    for name in output_names:
        output_path = args.output_dir / name

        if output_path.exists():
            parser.error(
                f"Output already exists: {output_path}; "
                "choose a new folder"
            )

    # ---------------------------------------------------------
    # 1. Validate input file
    # ---------------------------------------------------------

    if not args.seed.exists():
        raise SystemExit(
            f"Seed/input file not found: {args.seed}"
        )

    if not args.seed.is_file():
        raise SystemExit(
            f"Seed/input path is not a file: {args.seed}"
        )

    # ---------------------------------------------------------
    # 2. Create output directory
    # ---------------------------------------------------------

    args.output_dir.mkdir(
        parents=True,
        exist_ok=True,
    )

    # ---------------------------------------------------------
    # 3. Read FPC OA input
    # ---------------------------------------------------------

    print(
        f"Reading FPC OA input: {args.seed}"
    )

    payload = read_seed(
        args.seed
    )

    # ---------------------------------------------------------
    # 4. Execute through common RequestPipeline
    # ---------------------------------------------------------

    print(
        "Running FPC OA RequestPipeline..."
    )

    pipeline_result = asyncio.run(
        execute(payload)
    )

    if not isinstance(pipeline_result, dict):
        raise RuntimeError(
            "Common pipeline returned no FPC OA result"
        )

    # ---------------------------------------------------------
    # 5. Run calculations
    # ---------------------------------------------------------

    print(
        "Running FPC OA calculations..."
    )

    calculated_result = calculate(
        payload
    )

    if (
        pipeline_result.get("records")
        != jsonable(calculated_result["records"])
    ):
        raise RuntimeError(
            "Pipeline records differ from the direct calculation"
        )

    # ---------------------------------------------------------
    # 6. Validate against saved seed/reference
    # ---------------------------------------------------------

    print(
        "Running seed/reference validation..."
    )

    reference = compare_seed(
        calculated_result,
        args.seed,
    )

    if not reference["matches_saved_seed"]:
        validation_failed_path = (
            args.output_dir
            / "validation_failed.json"
        )

        write_json(
            validation_failed_path,
            reference,
        )

        raise SystemExit(
            "Saved seed comparison failed. "
            "See output/validation_failed.json."
        )

    # ---------------------------------------------------------
    # 7. Generate template-free workbook
    # ---------------------------------------------------------

    output_info = write_output(
        args.output_dir / "I_FPC_OA.xlsx",
        calculated_result,
        validated_input=payload,
    )

    pipeline_result["output"].update(
        output_info
    )

    # ---------------------------------------------------------
    # Build validation report
    # ---------------------------------------------------------

    report = {
        "status": "success",
        "output": output_info,
        "reference_comparison": reference,
        "control": calculated_result.get(
            "control"
        ),
        "exceptions": calculated_result.get(
            "exceptions",
            [],
        ),
        "summary": calculated_result.get(
            "summary",
            {},
        ),
        "runtime_tested": (
            "common RequestPipeline; "
            "HTTP A2A not tested"
        ),
    }

    # ---------------------------------------------------------
    # 8. Write generated output files
    # ---------------------------------------------------------

    print(
        "Writing output files..."
    )

    write_json(
        args.output_dir / "I_FPC_OA.json",
        pipeline_result,
    )

    write_json(
        args.output_dir / "calculated_I_FPC_OA.json",
        calculated_result,
    )

    write_json(
        args.output_dir / "validated_input.json",
        payload,
    )

    write_json(
        args.output_dir / "validation_report.json",
        report,
    )

    # ---------------------------------------------------------
    # 9. Print final execution summary
    # ---------------------------------------------------------

    console_result = {
        "status": "success",
        "summary": calculated_result.get(
            "summary",
            {},
        ),
        "validation": reference,
        "output_directory": str(
            args.output_dir
        ),
    }

    print()
    print(
        "FPC OA processing completed successfully."
    )
    print()

    print(
        json.dumps(
            jsonable(console_result),
            indent=2,
            allow_nan=False,
            default=json_serializer,
        )
    )


if __name__ == "__main__":
    main()