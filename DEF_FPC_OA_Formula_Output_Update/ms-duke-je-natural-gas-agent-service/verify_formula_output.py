"""Diagnostic formula/output checks without running the agent HTTP server.

Normal use imports the installed common reader. --common-root explicitly loads
only that real source module when the full common package cannot import.
Neither mode is a RequestPipeline or live A2A acceptance test.
"""
import argparse
import importlib.util
import json
from pathlib import Path

from def_fpc_oa.calculations import calculate, jsonable
from def_fpc_oa.inputs import read_seed
from def_fpc_oa.output import compare_seed, write_output
from def_i_imbalance.inputs import read_i_fpc_oa
from def_i_imbalance.calculations import calculate_i_imbalance


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--seed', type=Path, required=True)
    parser.add_argument('--output-dir', type=Path, required=True)
    parser.add_argument('--common-root', type=Path)
    parser.add_argument('--engine', choices=('auto', 'excel', 'libreoffice'), default='auto')
    args = parser.parse_args()
    if args.output_dir.exists():
        parser.error('Choose a new diagnostic output directory')
    opener = None
    if args.common_root:
        source = args.common_root / 'src/ms_duke_je_common/normalization/extraction/workbook_reader.py'
        spec = importlib.util.spec_from_file_location('actual_common_reader', source)
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        opener = module.open_workbook
    payload = read_seed(args.seed, workbook_opener=opener)
    result = calculate(payload)
    reference = compare_seed(result, args.seed)
    args.output_dir.mkdir(parents=True)
    output = args.output_dir / 'I_FPC_OA.xlsx'
    info = write_output(output, result, validated_input=payload, recalc_engine=args.engine)
    reread = read_i_fpc_oa(output, workbook_opener=opener)
    direct = calculate_i_imbalance(result)
    downstream = calculate_i_imbalance(reread)
    if direct['summary'] != downstream['summary']:
        raise ValueError('Next-stage summary differs after native workbook readback')
    report = {'status': 'diagnostic_checks_passed', 'output': info,
              'reference_comparison': reference, 'downstream_records': len(reread['records']),
              'next_stage_summary_matches': True, 'request_pipeline_tested': False,
              'http_a2a_tested': False, 'reader_mode': 'actual common source module' if args.common_root
              else 'installed common reader'}
    for filename, data in [('validated_input.json', payload),
                           ('calculated_I_FPC_OA.json', result),
                           ('validation_report.json', report)]:
        (args.output_dir / filename).write_text(json.dumps(jsonable(data), indent=2,
                                                         allow_nan=False), encoding='utf-8')
    print(json.dumps(report, indent=2))


if __name__ == '__main__':
    main()
