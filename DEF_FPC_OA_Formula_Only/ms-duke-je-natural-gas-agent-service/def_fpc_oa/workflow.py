"""
First-flow composition.

Reads the FPC OA seed workbook, performs calculations,
validates the calculated values against the saved seed,
and generates I_FPC_OA without requiring an Excel template.
"""

import base64
import binascii
import tempfile

from pathlib import Path

from .inputs import read_seed
from .calculations import calculate, jsonable
from .output import compare_seed, write_output, cell_updates


def process_workbooks(data, *, workbook_opener=None):
    """
    Process the FPC OA seed workbook.

    Expected input:
        seed_base64

    No template_base64 is required.
    """

    def decode(key):
        value = data.get(key)

        if not isinstance(value, str):
            raise ValueError(
                f"{key}: expected base64 workbook"
            )

        if len(value) > 8_000_000:
            raise ValueError(
                f"{key}: expected base64 workbook "
                f"of at most 8MB encoded"
            )

        try:
            return base64.b64decode(
                value,
                validate=True,
            )

        except (binascii.Error, ValueError) as exc:
            raise ValueError(
                f"Invalid {key}"
            ) from exc

    with tempfile.TemporaryDirectory(
        prefix="def-fpc-oa-"
    ) as directory:

        root = Path(directory)

        # -------------------------------------------------
        # Temporary files
        # -------------------------------------------------

        seed = root / "seed.xlsx"

        output = root / "I_FPC_OA.xlsx"

        # -------------------------------------------------
        # Decode seed workbook
        # -------------------------------------------------

        seed.write_bytes(
            decode("seed_base64")
        )

        # -------------------------------------------------
        # Read seed workbook
        # -------------------------------------------------

        payload = read_seed(
            seed,
            workbook_opener=workbook_opener,
        )

        # -------------------------------------------------
        # Perform FPC OA calculations
        # -------------------------------------------------

        result = calculate(
            payload,
            strict=True,
        )

        # -------------------------------------------------
        # Validate against saved FPC OA values
        # -------------------------------------------------

        validation = compare_seed(
            result,
            seed,
        )

        if not validation["matches_saved_seed"]:
            raise ValueError(
                "Calculated values differ from saved seed; "
                "output withheld for review"
            )

        # -------------------------------------------------
        # Generate NEW I_FPC_OA workbook
        #
        # No Actual_data.xlsm/template is required
        # -------------------------------------------------

        info = write_output(
            output,
            result,
            validated_input=payload,
        )

        info["file"] = "I_FPC_OA.xlsx"

        # -------------------------------------------------
        # Return output
        # -------------------------------------------------

        result["output"] = {
            **info,
            "cell_updates": cell_updates(result),
            "workbook_base64": base64.b64encode(
                output.read_bytes()
            ).decode("ascii"),
        }

        result["validation"] = validation

        return jsonable(result)