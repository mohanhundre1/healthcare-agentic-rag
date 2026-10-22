"""Run the REAL common pipeline when installed; never substitute framework mocks."""

import asyncio
import importlib.util
import unittest

from test_calculations import example
from test_i_imbalance_calculations import example_i_fpc_oa


@unittest.skipUnless(
    importlib.util.find_spec("ms_duke_je_common"),
    "Approved common runtime is not installed",
)
class CommonRuntimeTests(unittest.TestCase):

    def test_fpc_local_pipeline(self):
        from run_local import execute

        result = asyncio.run(
            execute(
                example()
            )
        )

        self.assertTrue(
            result["control"]["rollforward_balanced"]
        )

        self.assertEqual(
            result["summary"]["ending_imbalance_dt"],
            "100",
        )

    def test_imbalance_local_pipeline_zero_burn_blank_heat(self):
        from run_i_imbalance import execute

        payload = example_i_fpc_oa()

        payload["records"][0]["burns_dt"]["tiger bay"] = "0"

        payload["records"][0]["btu_factors"]["tiger bay"] = None

        result = asyncio.run(
            execute(
                payload
            )
        )

        self.assertTrue(
            result["complete_fgt"]
        )

        self.assertFalse(
            result["complete_workflow"]
        )

        self.assertTrue(
            result["output"]["workbook_base64"]
        )

    def test_combined_agent_dispatches_second_flow(self):
        from ms_duke_je_common.context.execution import (
            ExecutionContext,
            RequestContext,
            ResponseContext,
            ServiceContainer,
        )
        from ms_duke_je_common.context.security import (
            SecurityContext,
        )
        from ms_duke_je_common.executor.request_pipeline import (
            RequestPipeline,
        )

        from def_fpc_oa.logic import FpcOaLogic

        ctx = ExecutionContext(
            RequestContext(
                "local",
                "local",
                "def-fpc-oa-agent",
                "i_imbalance_fgt",
                "",
                {
                    "i_fpc_oa": example_i_fpc_oa()
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

        async def run():
            pipeline = RequestPipeline(
                FpcOaLogic(),
                [],
            )

            await pipeline.startup(
                ctx.services
            )

            try:
                await pipeline.run(
                    ctx
                )
            finally:
                await pipeline.shutdown()

        asyncio.run(
            run()
        )

        self.assertEqual(
            ctx.response.result["target"]["type"],
            "I_Imbalance",
        )

        self.assertTrue(
            ctx.response.success
        )