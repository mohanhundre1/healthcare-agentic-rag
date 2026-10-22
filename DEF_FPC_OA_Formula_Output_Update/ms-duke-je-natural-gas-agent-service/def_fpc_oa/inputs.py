
"""Temporary Rahul-side seed adapter using the supplied common workbook reader."""
from datetime import datetime, timedelta
from hashlib import sha256
from pathlib import Path

from .calculations import PLANTS
from .seed_reference import Workbook, numeric, excel_text


def read_seed(path: Path, *, workbook_opener=None) -> dict:
    if workbook_opener is None:
        from ms_duke_je_common.normalization.extraction.workbook_reader import open_workbook
        workbook_opener = open_workbook
    # Inspect OOXML caches as well: openpyxl alone cannot distinguish a missing
    # saved formula value from a genuinely blank source cell.
    metadata = Workbook(path)
    wb = workbook_opener(path)
    try:
        values = {s.title: {c.coordinate: c.value for row in s for c in row
                            if c.value is not None} for s in wb}
    finally:
        wb.close()
    fpc, raw, total, gsng = 'FPC OA', 'FGT Data Download - New', 'TOTAL FGT OA', 'GSNG Imbalance'
    for sheet in (fpc, raw, total, gsng):
        if sheet not in values:
            raise ValueError(f'Missing required sheet: {sheet}')

    def get(sheet, cell):
        metadata.get(sheet, cell) # Fail on Excel errors or unavailable caches.
        value = values[sheet].get(cell)
        if isinstance(value, datetime):
            return (value - datetime(1899,12,30)).days
        return value

    def ref(sheet, row):
        return {'file': path.name, 'sheet': sheet, 'row': row}

    def iso(serial):
        return (datetime(1899,12,30) + timedelta(days=float(serial))).date().isoformat()

    for cell, header in {'B2':'StartDate','E2':'LOC','N2':'EnergyQty','Q2':'BTUFctr'}.items():
        if get(raw, cell) != header:
            raise ValueError(f'Unsupported seed header: {raw}!{cell}')
    raw_records = []
    for row in range(3, 251):
        day, station = get(raw, f'B{row}'), get(raw, f'E{row}')
        if day is None or station is None:
            continue
        raw_records.append({'date': iso(day), 'station_id': excel_text(station),
                            'energy_qty_dt': get(raw, f'N{row}'), 'btu_factor': get(raw, f'Q{row}'),
                            'source': ref(raw, row)})
    days = []
    for row in range(13,44):
        days.append({'date': iso(get(total, f'B{row-1}')),
                     'scheduled_volume_dt': get(fpc, f'C{row}'),
                     'alert_day': get(total, f'D{row-1}'),
                     'alert_direction': get(total, f'E{row-1}'),
                     'gulfstream_peaker_burn_dt': get(gsng, f'Y{row-2}'),
                     'source': ref(fpc,row),
                     'dependencies': [ref(total,row-1),ref(gsng,row-2)]})
    return {'schema_version':'1','source': {'file': path.name,'sha256':sha256(path.read_bytes()).hexdigest(),
                                            'period_label':get(total,'A4'),
                                            'adapter':'supplied-August-seed-layout','uses_saved_cross_sheet_values':True},
            'station_ids': {col:excel_text(get(fpc,f'{col}6')) for col in PLANTS},
            'opening_imbalance_dt':get(fpc,'Q12'), 'fgt_rolling_imbalance_dt':get(fpc,'Q53'),
            'raw_records':raw_records,'days':days}