"""Start the FPC OA agent using the common-library A2A server."""

import os

from ms_duke_je_common.agent_descriptor import (
    AgentDescriptor,
    SkillDescriptor,
)
from ms_duke_je_common.entrypoints.pipeline_server import (
    run_pipeline_agent,
)
from ms_duke_je_common.transport.a2a.agent_card_adapter import (
    agent_card_from_descriptor,
)

from def_fpc_oa.logic import FpcOaLogic


AGENT_ID = "def-fpc-oa-agent"
CAPABILITY = "fpc_oa"
SERVICE_NAME = "duke-def-fpc-oa"


def create_agent_card():
    """Create the A2A agent card for the FPC OA service."""

    return agent_card_from_descriptor(
        AgentDescriptor(
            name="DEF FPC OA calculation agent",
            description=(
                "Maps and calculates validated FPC OA records"
            ),
            path_prefix="",
            base_url=os.environ["SERVICE_BASE_URL"],
            skills=(
                SkillDescriptor(
                    id=CAPABILITY,
                    name="FPC OA calculations",
                    description=(
                        "Validated records to I_FPC_OA cell values"
                    ),
                    tags=(
                        "def",
                        "gas",
                        "calculations",
                    ),
                ),
                SkillDescriptor(
                    id="i_imbalance_fgt",
                    name="FGT I_Imbalance mapping",
                    description=(
                        "Calculated I_FPC_OA JSON/XLSX "
                        "to FGT-only I_Imbalance"
                    ),
                    tags=(
                        "def",
                        "gas",
                        "fgt",
                        "mapping",
                    ),
                ),
            ),
        )
    )


def require_environment():
    """Validate environment required to start the A2A server."""

    required_variables = (
        "SERVICE_BASE_URL",
        "FPC_OA_AGENT_PORT",
        "AUTH_PROVIDER",
    )

    missing_variables = [
        name
        for name in required_variables
        if not os.getenv(name)
    ]

    if missing_variables:
        raise RuntimeError(
            "Missing required environment variable(s): "
            + ", ".join(missing_variables)
        )

    if (
        os.environ["AUTH_PROVIDER"] in (
            "local",
            "local_dev",
        )
        and not os.getenv("JWT_SECRET")
    ):
        raise RuntimeError(
            "JWT_SECRET is required when AUTH_PROVIDER "
            "is local or local_dev."
        )

    # Validate authentication provider during startup.
    from ms_duke_je_common.providers.auth.factory import (
        AuthFactory,
    )

    AuthFactory.create()


def main():
    """Start the FPC OA A2A server."""

    require_environment()

    run_pipeline_agent(
        logic_factory=FpcOaLogic,
        agent_card_factory=create_agent_card,
        agent_id=AGENT_ID,
        default_capability=CAPABILITY,
        service_name=SERVICE_NAME,
        description="DEF FPC OA",
        default_port=int(
            os.environ["FPC_OA_AGENT_PORT"]
        ),
        start_banner="Starting DEF FPC OA",
        enable_services=(
            "auth",
        ),
        optional_middleware=(),
        pre_start_hooks=(
            require_environment,
        ),
    )


if __name__ == "__main__":
    main()