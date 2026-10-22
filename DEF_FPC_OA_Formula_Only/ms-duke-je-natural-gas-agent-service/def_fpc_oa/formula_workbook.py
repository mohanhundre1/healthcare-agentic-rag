"""Embedded dependencies and formula-only output. No calculation engine runs."""
import os
import tempfile
from datetime import date
from decimal import Decimal
from pathlib import Path

def populate_dependencies(workbook, payload):
    """Embed the validated inputs at the addresses used by client formulas.

    Raw rows begin at 4 to match the CLIENT VLOOKUP range. The supplied seed
    starts at row 3; every validated record is copied, including its first row.
    The helper sheets are visible and editable for formula traceability.
    """
    raw = workbook.create_sheet('FGT Data Download - New')
    for address, value in {'A3': 'Lookup key', 'B3': 'StartDate', 'E3': 'LOC',
                           'N3': 'EnergyQty', 'Q3': 'BTUFctr'}.items():
        raw[address] = value
    for row, record in enumerate(payload['raw_records'], 4):
        raw[f'A{row}'] = f'=B{row}&E{row}'
        raw[f'B{row}'] = date.fromisoformat(record['date'])
        raw[f'B{row}'].number_format = 'mm/dd/yyyy'
        raw[f'E{row}'] = int(record['station_id'])
        raw[f'N{row}'] = float(Decimal(str(record['energy_qty_dt'])))
        raw[f'Q{row}'] = float(Decimal(str(record['btu_factor'])))
    total = workbook.create_sheet('TOTAL FGT OA')
    first = date.fromisoformat(payload['days'][0]['date'])
    total['A4'] = payload['source'].get('period_label') or first.strftime('%B %Y')
    gulf = workbook.create_sheet('GSNG Imbalance')
    gulf['Y10'] = 'Validated Gulfstream peaker burn (DT)'
    for index, day in enumerate(payload['days']):
        row = index + 12
        total[f'B{row}'] = date.fromisoformat(day['date'])
        total[f'B{row}'].number_format = 'mm/dd/yyyy'
        # D/E are supporting snapshots. The client's I_FPC OA D/E cells
        # themselves are inputs rather than formulas.
        total[f'D{row}'] = day.get('alert_day')
        total[f'E{row}'] = day.get('alert_direction')
        gulf[f'Y{index + 11}'] = float(Decimal(str(day['gulfstream_peaker_burn_dt'])))


def finalize_formula_workbook(workbook, output, result, formulas):
    """Save formula expressions without running or manufacturing calculation caches."""
    from .client_formulas import CLIENT_FORMULAS
    from openpyxl import load_workbook
    from openpyxl.workbook.properties import CalcProperties

    # Do not request calculation when opening/saving the generated workbook.
    # Excel application settings and any later user actions control calculation.
    workbook.calculation = CalcProperties(calcMode='manual', fullCalcOnLoad=False,
                                          forceFullCalc=False, calcOnSave=False)
    output.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix='fpc-formulas-', dir=output.parent) as temp:
        generated = Path(temp) / 'I_FPC_OA.xlsx'
        workbook.save(generated)
        check = load_workbook(generated, read_only=True, data_only=False)
        try:
            saved = {c.coordinate: c.value for row in check['I_FPC OA']
                     for c in row if c.data_type == 'f'}
            if saved != formulas:
                raise ValueError('Saved formula text differs from client formula contract')
        finally:
            check.close()
        os.link(generated, output)
    return {'file': str(output), 'target_sheet': 'I_FPC OA', 'template_used': False,
            'vba_executed': False, 'downstream_recalculated': False,
            'formulas_written': len(formulas), 'recalculation_performed': False,
            'recalculation_engine': None, 'native_formula_caches': 0,
            'formula_validation': {'formula_text_checked': len(formulas),
                                   'matches_client_with_internal_references': True,
                                   'formula_results_checked': 0},
            'external_references_internalized': sum('[17]' in f for f in CLIENT_FORMULAS.values()),
            'downstream_input': 'calculated_I_FPC_OA.json',
            'note': 'Formula cells have no saved results. Use calculated JSON for I_Imbalance.'}
