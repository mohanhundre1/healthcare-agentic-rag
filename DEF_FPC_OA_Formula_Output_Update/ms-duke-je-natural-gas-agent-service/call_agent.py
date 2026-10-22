"""Call the FPC OA agent using the common A2A client wrapper."""

import argparse
import asyncio
import base64
import json
import os

from pathlib import Path
from uuid import uuid4

from ms_duke_je_common.transport.a2a.a2a_client import (
    DownstreamAgent,
    extract_data_parts,
)


def encode_file(path: Path) -> str:
    """Read a file and return its Base64 encoded content."""

    return base64.b64encode(
        path.read_bytes()
    ).decode("ascii")


def write_json(path: Path, value) -> None:
    """Write formatted JSON output."""

    path.write_text(
        json.dumps(
            value,
            indent=2,
            allow_nan=False,
            default=str,
        ),
        encoding="utf-8",
    )


async def main():
    parser = argparse.ArgumentParser(
        description="Call the FPC OA agent through A2A."
    )

    parser.add_argument(
        "--url",
        required=True,
        help="URL of the running FPC OA agent.",
    )

    input_group = parser.add_mutually_exclusive_group(
        required=True
    )

    input_group.add_argument(
        "--validated-input",
        type=Path,
        help="Path to validated_input.json.",
    )

    input_group.add_argument(
        "--seed",
        type=Path,
        help="Path to the FPC OA seed/input file.",
    )

    input_group.add_argument(
        "--i-fpc-oa",
        type=Path,
        help=(
            "Calculated I_FPC_OA JSON or service-generated "
            "XLSX for FGT I_Imbalance."
        ),
    )

    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path("output"),
        help="Directory for generated A2A output files.",
    )

    args = parser.parse_args()

    # ---------------------------------------------------------
    # 1. Validate token
    # ---------------------------------------------------------

    token = os.environ.get(
        "A2A_BEARER_TOKEN"
    )

    if not token:
        raise SystemExit(
            "A2A_BEARER_TOKEN is not set. "
            "Set the approved development bearer token "
            "before calling the agent."
        )

    # ---------------------------------------------------------
    # 2. Validate input files
    # ---------------------------------------------------------

    if args.seed:
        if not args.seed.exists():
            raise SystemExit(
                f"Seed file not found: {args.seed}"
            )

        if not args.seed.is_file():
            raise SystemExit(
                f"Seed path is not a file: {args.seed}"
            )

    if args.validated_input:
        if not args.validated_input.exists():
            raise SystemExit(
                "Validated input file not found: "
                f"{args.validated_input}"
            )

        if not args.validated_input.is_file():
            raise SystemExit(
                "Validated input path is not a file: "
                f"{args.validated_input}"
            )

    if (
        args.i_fpc_oa
        and not args.i_fpc_oa.is_file()
    ):
        parser.error(
            f"I_FPC_OA file not found: {args.i_fpc_oa}"
        )

    workbook_name = (
        "I_Imbalance.xlsx"
        if args.i_fpc_oa
        else "I_FPC_OA.xlsx"
    )

    for name in (
        workbook_name,
        "a2a_result.json",
    ):
        output_path = args.output_dir / name

        if output_path.exists():
            parser.error(
                f"Output already exists: {output_path}"
            )

    # ---------------------------------------------------------
    # 3. Create output directory automatically
    # ---------------------------------------------------------

    args.output_dir.mkdir(
        parents=True,
        exist_ok=True,
    )

    # ---------------------------------------------------------
    # 4. Create A2A client
    # ---------------------------------------------------------

    client = DownstreamAgent(
        args.url,
        auth_headers={
            "Authorization": f"Bearer {token}"
        },
        context_id=str(uuid4()),
    )

    # ---------------------------------------------------------
    # 5. Build request payload
    # ---------------------------------------------------------

    if args.seed:
        data = {
            "seed_base64": encode_file(
                args.seed
            ),
        }

    elif args.i_fpc_oa:
        if args.i_fpc_oa.suffix.lower() == ".xlsx":
            data = {
                "i_fpc_oa_base64": encode_file(
                    args.i_fpc_oa
                )
            }

        else:
            from def_i_imbalance.inputs import (
                read_i_fpc_oa,
            )

            data = {
                "i_fpc_oa": read_i_fpc_oa(
                    args.i_fpc_oa
                )
            }

    else:
        try:
            data = {
                "validated_input": json.loads(
                    args.validated_input.read_text(
                        encoding="utf-8"
                    )
                )
            }

        except json.JSONDecodeError as exc:
            raise SystemExit(
                f"Invalid JSON file: "
                f"{args.validated_input}\n{exc}"
            ) from exc

    # ---------------------------------------------------------
    # 6. Send request to FPC OA agent
    # ---------------------------------------------------------

    task = await client.send_data(
        data
    )

    results = extract_data_parts(
        task
    )

    if not results:
        raise RuntimeError(
            "A2A task returned no data artifacts. "
            "Inspect the agent/server logs."
        )

    # ---------------------------------------------------------
    # 7. Process returned artifacts
    # ---------------------------------------------------------

    workbook_saved = False
    review_required = False

    for result in results:
        if not isinstance(result, dict):
            continue

        if result.get("complete_fgt") is False:
            review_required = True

        output = result.get("output")

        if not isinstance(output, dict):
            continue

        workbook_base64 = output.get(
            "workbook_base64"
        )

        if not workbook_base64:
            continue

        workbook_path = (
            args.output_dir / workbook_name
        )

        try:
            workbook_bytes = base64.b64decode(
                workbook_base64,
                validate=True,
            )

        except (
            ValueError,
            base64.binascii.Error,
        ) as exc:
            raise RuntimeError(
                "Agent returned invalid workbook Base64 data."
            ) from exc

        workbook_path.write_bytes(
            workbook_bytes
        )

        workbook_saved = True

        print(
            f"Saved generated workbook to: "
            f"{workbook_path}"
        )

    # ---------------------------------------------------------
    # 8. Save complete A2A response
    # ---------------------------------------------------------

    response_path = (
        args.output_dir
        / "a2a_result.json"
    )

    write_json(
        response_path,
        results,
    )

    print(
        f"Saved A2A response to: "
        f"{response_path}"
    )

    # ---------------------------------------------------------
    # 9. Final status
    # ---------------------------------------------------------

    if review_required:
        raise SystemExit(
            "FGT artifacts saved for review; "
            "mapping has exceptions"
        )

    if workbook_saved:
        print(
            "FPC OA agent execution completed successfully."
        )
    else:
        print(
            "A2A execution completed, but no workbook "
            "was returned by the agent."
        )


if __name__ == "__main__":
    asyncio.run(
        main()
    )