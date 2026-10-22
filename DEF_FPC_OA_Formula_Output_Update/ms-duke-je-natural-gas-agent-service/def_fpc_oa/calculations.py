
"""
No Excel, network, A2A SDK, or infrastructure dependencies.
Rules reproduce the supplied August seed, not unspecified future allocations.
"""
from datetime import date, timedelta
from decimal import Decimal, ROUND_HALF_UP
from math import isfinite

PLANTS = {'F': 'gainesville', 'G': 'bartow', 'H': 'suwannee', 'I': 'hines',
          'J': 'debary', 'K': 'anclote', 'L': 'tiger_bay', 'M': 'vandolah'}
BTU = {'S': 'I', 'T': 'M', 'U': 'H', 'V': 'G', 'W': 'K', 'X': 'L', 'Y': 'J', 'Z': 'F'}
RULE_VERSION = 'fpc-oa-seed-v1'


def decimal(value, field):
  if isinstance(value, bool) or not isinstance(value, (str, int, float, Decimal)):
      raise ValueError(f'{field}: expected a finite decimal')
  try:
      result = Decimal(str(value))
  except Exception as exc:
      raise ValueError(f'{field}: invalid decimal') from exc
  if not result.is_finite():
      raise ValueError(f'{field}: expected finite decimal')
  return result


def calculate(payload: dict, *, strict: bool = True) -> dict:
    """
    strict=True rejects unknown/duplicate lookups. strict=False explicitly mirrors
    the seed's IFERROR-zero / VLOOKUP-first behavior and records every exception.
    """
    if payload.get('schema_version') != '1':
        raise ValueError('Unsupported input schema_version')
    mappings = payload['station_ids']
    if set(mappings) != set(PLANTS):
        raise ValueError('station_ids must contain F:M')
    ids = [str(mappings[c]) for c in PLANTS]
    if len(set(ids)) != len(ids):
        raise ValueError('Duplicate station mapping IDs')
    opening = decimal(payload['opening_imbalance_dt'], 'opening imbalance')
    cumulative = opening
    lookup, exceptions = {}, []
    for raw in payload['raw_records']:
        day = date.fromisoformat(raw['date']).isoformat()
        key = (day, str(raw['station_id']))
        if key in lookup:
            exceptions.append({'kind': 'duplicate_lookup', 'key': list(key),
                                'source': raw['source']})
        lookup.setdefault(key, raw)
    if strict and exceptions:
        raise ValueError(f'Duplicate source lookups: {len(exceptions)}')
    days = payload['days']
    if not days:
        raise ValueError('No daily records')
    dates = [date.fromisoformat(d['date']) for d in days]
    if dates != sorted(set(dates)) or any(
        b-a != timedelta(days=1)
        for a,b in zip(dates,dates[1:])):
        raise ValueError('Dates must be unique, sorted and contiguous')
    if any((d.year, d.month) != (dates[0].year, dates[0].month) for d in dates):
        raise ValueError('Daily records must belong to one reporting month')
    records = []
    for day in days:
        scheduled = decimal(day['scheduled_volume_dt'], 'scheduled volume')
        burns, factors, refs = {}, {}, {}
        for col, plant in PLANTS.items():
            match = lookup.get((day['date'], str(mappings[col])))
            if match is None:
                exc = {'kind': 'missing_lookup', 'date': day['date'], 'plant': plant,
                        'station_id': mappings[col], 'source': day['source']}
                if strict:
                    raise ValueError(f'Missing mapping input: {exc}')
                exceptions.append(exc)
            burn = decimal(match['energy_qty_dt'], 'energy_qty_dt') if match else Decimal(0)
            burns[plant] = burn
            refs[plant] = match['source'] if match else None
            factors[plant] = (
                int(
                    (decimal(match['btu_factor'], 'btu_factor') * 1000)
                    .quantize(Decimal('1'), rounding=ROUND_HALF_UP)
                )
                if match and burn
                else None
            )
        total = sum(burns.values(), Decimal(0))
        daily = scheduled + total
        cumulative += daily
        alert = day.get('alert_day')
        tolerance = 'NA' if alert in (None, '', ' ') else scheduled * decimal(alert, 'alert_day')
        fgt = sum((burns[p] for p in ('hines', 'anclote', 'tiger_bay')), Decimal(0))
        gulf = decimal(day['gulfstream_peaker_burn_dt'], 'gulfstream peaker')
        records.append({'date': day['date'], 'scheduled_volume_dt': scheduled,
                        'alert_day': alert, 'alert_direction': day.get('alert_direction'),
                        'burns_dt': burns, 'btu_factors': factors,
                        'total_plant_burns_dt': total, 'daily_imbalance_dt': daily,
                        'tolerance_volume_dt': tolerance, 'cumulative_imbalance_dt': cumulative,
                        'fgt_peaker_burn_dt': fgt, 'gulfstream_peaker_burn_dt': gulf,
                        'fpc_peaker_burn_dt': fgt + gulf, 'source': day['source'],
                        'lookup_sources': refs, 'rule_version': RULE_VERSION})
    scheduled_total = sum((r['scheduled_volume_dt'] for r in records), Decimal(0))
    burn_total = sum((r['total_plant_burns_dt'] for r in records), Decimal(0))
    totals = {p: sum((r['burns_dt'][p] for r in records), Decimal(0)) for p in PLANTS.values()}
    averages = {}
    for plant in PLANTS.values():
        values = [r['btu_factors'][plant] for r in records if r['btu_factors'][plant] is not None]
        averages[plant] = sum(map(Decimal, values)) / len(values) if values else None
    rolling = decimal(payload['fgt_rolling_imbalance_dt'], 'rolling imbalance')
    return {'schema_version': '1', 'rule_version': RULE_VERSION, 'source': payload['source'],
            'station_ids': mappings, 'opening_imbalance_dt': opening, 'records': records,
            'summary': {'scheduled_volume_dt': scheduled_total, 'total_plant_burns_dt': burn_total,
            'plant_totals_dt': totals, 'average_btu_factors': averages,
            'ending_imbalance_dt': cumulative, 'gru_therms': totals['gainesville'] * 10,
            'fgt_rolling_imbalance_dt': rolling, 'rolling_less_opening_dt': rolling-opening},
        'exceptions': exceptions, 'lookup_policy': 'reject' if strict else 'seed_compatible',
        'control': {'rollforward_balanced': cumulative == opening + scheduled_total + burn_total}}


def jsonable(value):
    """Decimal values travel as strings to preserve precision across handoffs."""
    if isinstance(value, Decimal):
        return str(value)
    if isinstance(value, dict):
        return {k: jsonable(v) for k,v in value.items()}
    if isinstance(value, list):
        return [jsonable(v) for v in value]
    return value