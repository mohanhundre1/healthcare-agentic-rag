"""Common-library agent adapter. Infrastructure stays in the common library."""

import asyncio

from ms_duke_je_common.executor.agent_logic import BaseAgentLogic
from ms_duke_je_common.errors import InputValidationError

from .calculations import calculate, jsonable
from .output import cell_updates
from .workflow import process_workbooks


class FpcOaLogic(BaseAgentLogic):

    async def execute_business_logic(self, ctx):

        # ---------------------------------------------------------
        # Route I_FPC_OA input to the I_Imbalance flow
        # ---------------------------------------------------------

        if any(
            key in ctx.request.input_data
            for key in (
                "i_fpc_oa",
                "i_fpc_oa_base64",
            )
        ):
            from def_i_imbalance.logic import IImbalanceLogic

            return await IImbalanceLogic().execute_business_logic(
                ctx
            )

        payload = ctx.request.input_data.get(
            "validated_input"
        )

        try:

            # -----------------------------------------------------
            # Flow 1:
            # Already validated input data
            # -----------------------------------------------------

            if isinstance(payload, dict):

                result = calculate(
                    payload,
                    strict=True,
                )

                result["output"] = {
                    "target_sheet": "I_FPC OA",
                    "cell_updates": cell_updates(
                        result
                    ),
                }

            # -----------------------------------------------------
            # Flow 2:
            # Seed workbook supplied directly
            # -----------------------------------------------------

            elif "seed_base64" in ctx.request.input_data:

                result = await asyncio.to_thread(
                    process_workbooks,
                    ctx.request.input_data,
                )

            else:

                raise ValueError(
                    "Expected validated_input or seed_base64"
                )

        except (ValueError, KeyError, TypeError) as exc:

            raise InputValidationError(
                str(exc)
            ) from exc

        # ---------------------------------------------------------
        # Prepare common pipeline response
        # ---------------------------------------------------------

        ctx.response.ui_kind = "result"

        ctx.response.success = (
            result["control"]["rollforward_balanced"]
        )

        ctx.response.status_message = (
            "FPC OA mapped, calculated and ready "
            "for I_FPC_OA output"
        )

        return jsonable(
            result
        )