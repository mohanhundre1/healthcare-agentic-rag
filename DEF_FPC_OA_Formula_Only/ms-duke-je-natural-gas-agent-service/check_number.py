"""Report actual readiness without importing/stubbing an unavailable server."""
import importlib
import sys

failed=sys.version_info < (3,13)
print(f'Python {sys.version.split()[0]}: '+('OK' if not failed else 'requires Python >=3.13'))
for module in ('ms_duke_je_common','ms_duke_je_common.entrypoints.pipeline_server',
                'ms_duke_je_common.transport.a2a.agent_card_adapter',
                'ms_duke_je_common.transport.a2a.a2a_client','uvicorn'):
    try:
        m=importlib.import_module(module)
        print(f'OK {module}: {getattr(m,"__file__","")}')
    except Exception as exc:
        failed=True
        print(f'FAIL {module}: {type(exc).__name__}: {exc}')
raise SystemExit(1 if failed else 0)