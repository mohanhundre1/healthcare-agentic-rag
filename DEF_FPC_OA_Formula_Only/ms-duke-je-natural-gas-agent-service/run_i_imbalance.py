"""Run I_FPC_OA JSON/XLSX -> FGT-only I_Imbalance through common RequestPipeline."""

import argparse
import asyncio
import base64
import json
from pathlib import Path

from def_i_imbalance.inputs import read_i_fpc_oa


async def execute(payload, *, strict=True):
    """
    Execute the I_FPC_OA -> I_Imbalance transformation
    through the common RequestPipeline.
    """

    from ms_duke_je_common.context.execution import (
        ExecutionContext,
        RequestContext,
        ResponseContext,
        ServiceContainer,
    )
    from ms_duke_je_common.context.security import SecurityContext
    from ms_duke_je_common.executor.request_pipeline import RequestPipeline

    from def_i_imbalance.logic import IImbalanceLogic

    ctx = ExecutionContext(
        RequestContext(
            "local",
            "local",
            "def-i-imbalance-agent",
            "i_imbalance_fgt",
            "",
            {
                "i_fpc_oa": payload,
            },
            SecurityContext(
                "local",
                "local",
                "Local development",
            ),
            "local",
        ),
        ResponseContext(),
        ServiceContainer(),
    )

    pipeline = RequestPipeline(
        IImbalanceLogic(strict=strict),
        [],
    )

    await pipeline.startup(ctx.services)

    try:
        await pipeline.run(ctx)

    finally:
        await pipeline.shutdown()

    if not isinstance(ctx.response.result, dict):
        raise RuntimeError(
            "Common pipeline returned no I_Imbalance result"
        )

    return ctx.response.result


def main():
    """Run the local I_Imbalance generation process."""

    parser = argparse.ArgumentParser(
        description=__doc__
    )

    parser.add_argument(
        "--input",
        type=Path,
        required=True,
        help=(
            "calculated_I_FPC_OA.json, "
            "I_FPC_OA.json, or I_FPC_OA.xlsx"
        ),
    )

    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path("output_imbalance"),
    )

    parser.add_argument(
        "--non-strict",
        action="store_true",
        help=(
            "Write review artifacts with exceptions; "
            "do not treat as approved output"
        ),
    )

    args = parser.parse_args()

    names = (
        "I_Imbalance.json",
        "I_Imbalance.xlsx",
        "I_Imbalance_validation_report.json",
    )

    if args.output_dir.resolve() == args.input.resolve():
        parser.error(
            "Output directory cannot be the input file"
        )

    for name in names:
        output_path = args.output_dir / name

        if output_path.exists():
            parser.error(
                f"Output already exists: {output_path}; "
                "choose a new folder"
            )

    try:
        payload = read_i_fpc_oa(
            args.input
        )

        result = asyncio.run(
            execute(
                payload,
                strict=not args.non_strict,
            )
        )

    except (ValueError, OSError) as exc:
        parser.exit(
            1,
            f"I_Imbalance processing failed: {exc}\n",
        )

    workbook = base64.b64decode(
        result["output"].pop("workbook_base64"),
        validate=True,
    )

    report = {
        "status": (
            "success_fgt_only"
            if result["complete_fgt"]
            else "requires_review"
        ),
        "source_file": str(args.input),
        "scope": "FGT_ONLY",
        "complete_workflow": False,
        "summary": result["summary"],
        "control": result["control"],
        "exceptions": result["exceptions"],
        "output": result["output"],
    }

    args.output_dir.mkdir(
        parents=True,
        exist_ok=True,
    )

    (
        args.output_dir / names[1]
    ).write_bytes(workbook)

    for name, content in (
        (names[0], result),
        (names[2], report),
    ):
        (
            args.output_dir / name
        ).write_text(
            json.dumps(
                content,
                indent=2,
                allow_nan=False,
            ),
            encoding="utf-8",
        )

    print(
        json.dumps(
            report,
            indent=2,
            allow_nan=False,
        )
    )

    return (
        0
        if result["complete_fgt"]
        else 2
    )


if __name__ == "__main__":
    raise SystemExit(main())