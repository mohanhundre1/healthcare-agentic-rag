"""Shared second-flow composition for local and A2A execution."""

import base64
import binascii
import tempfile

from pathlib import Path

from .calculations import calculate_i_imbalance, jsonable
from .inputs import _find_calculated_payload, read_i_fpc_oa
from .output import write_output


def process_i_imbalance(
    data,
    *,
    strict=True,
    workbook_opener=None,
):
    """
    Process I_FPC_OA input and generate the FGT-only I_Imbalance output.
    """

    payload = data.get("i_fpc_oa")

    # ---------------------------------------------------------
    # 1. Read I_FPC_OA input
    # ---------------------------------------------------------

    if payload is None:
        encoded = data.get(
            "i_fpc_oa_base64"
        )

        if (
            not isinstance(encoded, str)
            or len(encoded) > 8_000_000
        ):
            raise ValueError(
                "Expected i_fpc_oa JSON or bounded "
                "i_fpc_oa_base64 XLSX"
            )

        try:
            contents = base64.b64decode(
                encoded,
                validate=True,
            )

        except (
            ValueError,
            binascii.Error,
        ) as exc:
            raise ValueError(
                "Invalid i_fpc_oa_base64"
            ) from exc

        # -----------------------------------------------------
        # XLSX input flow
        # -----------------------------------------------------

        with tempfile.TemporaryDirectory(
            prefix="def-i-imbalance-input-"
        ) as directory:

            source = (
                Path(directory)
                / "I_FPC_OA.xlsx"
            )

            source.write_bytes(
                contents
            )

            payload = read_i_fpc_oa(
                source,
                workbook_opener=workbook_opener,
            )

    else:
        # -----------------------------------------------------
        # JSON / calculated payload flow
        # -----------------------------------------------------

        payload = _find_calculated_payload(
            payload
        )

    # ---------------------------------------------------------
    # 2. Calculate I_Imbalance
    # ---------------------------------------------------------

    result = calculate_i_imbalance(
        payload,
        strict=strict,
    )

    # ---------------------------------------------------------
    # 3. Generate temporary I_Imbalance workbook
    # ---------------------------------------------------------

    with tempfile.TemporaryDirectory(
        prefix="def-i-imbalance-output-"
    ) as directory:

        path = (
            Path(directory)
            / "I_Imbalance.xlsx"
        )

        info = write_output(
            path,
            result,
        )

        # -----------------------------------------------------
        # 4. Encode generated workbook for local/A2A response
        # -----------------------------------------------------

        info["workbook_base64"] = (
            base64.b64encode(
                path.read_bytes()
            ).decode("ascii")
        )

    # ---------------------------------------------------------
    # 5. Attach output information to result
    # ---------------------------------------------------------

    result["output"] = info

    # ---------------------------------------------------------
    # 6. Return JSON-safe response
    # ---------------------------------------------------------

    return jsonable(
        result
    )