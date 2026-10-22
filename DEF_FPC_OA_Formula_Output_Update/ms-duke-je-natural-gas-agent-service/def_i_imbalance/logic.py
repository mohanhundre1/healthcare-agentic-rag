"""Common RequestPipeline adapter for the FGT I_Imbalance handoff."""

import asyncio

from ms_duke_je_common.executor.agent_logic import BaseAgentLogic
from ms_duke_je_common.errors import InputValidationError

from .workflow import process_i_imbalance


class IImbalanceLogic(BaseAgentLogic):
    """Common pipeline logic for the FGT I_Imbalance flow."""

    def __init__(self, *, strict=True):
        super().__init__()

        self.strict = strict

    async def execute_business_logic(self, ctx):
        """
        Execute the I_FPC_OA -> I_Imbalance workflow.
        """

        try:
            result = await asyncio.to_thread(
                process_i_imbalance,
                ctx.request.input_data,
                strict=self.strict,
            )

        except (ValueError, KeyError, TypeError) as exc:
            raise InputValidationError(
                str(exc)
            ) from exc

        # ---------------------------------------------------------
        # Prepare common RequestPipeline response
        # ---------------------------------------------------------

        ctx.response.ui_kind = "result"

        # Success covers the advertised FGT-only skill,
        # not the full JE workflow.
        ctx.response.success = result[
            "complete_fgt"
        ]

        ctx.response.status_message = (
            "FGT mapping completed; other pipelines unavailable"
            if result["complete_fgt"]
            else "FGT mapping requires review"
        )

        return result
