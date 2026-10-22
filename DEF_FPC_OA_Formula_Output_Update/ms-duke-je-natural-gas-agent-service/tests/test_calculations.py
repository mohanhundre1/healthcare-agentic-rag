import copy
import unittest
from decimal import Decimal

from def_fpc_oa.calculations import calculate, PLANTS


def example():
    """
    Create synthetic input data for testing the FPC OA calculations.

    No JSON file is required for these unit tests.
    """

    return {
        "schema_version": "1",

        "source": {
            "file": "synthetic"
        },

        "station_ids": {
            plant: str(index)
            for index, plant in enumerate(PLANTS, 1)
        },

        "opening_imbalance_dt": "100",
        "fgt_rolling_imbalance_dt": "200",

        "days": [
            {
                "date": "2026-08-01",
                "scheduled_volume_dt": "80",
                "alert_day": "0.05",
                "gulfstream_peaker_burn_dt": "-3",
                "source": {
                    "row": 13
                }
            }
        ],

        "raw_records": [
            {
                "date": "2026-08-01",
                "station_id": str(index),
                "energy_qty_dt": "-10",
                "btu_factor": "1.0225",
                "source": {
                    "row": index
                }
            }
            for index in range(1, 9)
        ]
    }


class CalculationTests(unittest.TestCase):

    def test_signed_burn_rollforward_and_tolerance(self):
        result = calculate(example())
        record = result["records"][0]

        self.assertEqual(
            record["total_plant_burns_dt"],
            Decimal("-80")
        )

        self.assertEqual(
            record["daily_imbalance_dt"],
            0
        )

        self.assertEqual(
            record["cumulative_imbalance_dt"],
            100
        )

        self.assertEqual(
            record["tolerance_volume_dt"],
            4
        )

        self.assertEqual(
            record["fpc_peaker_burn_dt"],
            -33
        )

    def test_round_half_up_matches_excel(self):
        result = calculate(example())

        self.assertEqual(
            result["records"][0]["btu_factors"]["hines"],
            1023
        )

    def test_missing_station_rejected(self):
        payload = example()

        # Remove one station record.
        payload["raw_records"].pop()

        with self.assertRaisesRegex(
            ValueError,
            "Missing mapping"
        ):
            calculate(payload)

    def test_duplicate_rejected(self):
        payload = example()

        # Add duplicate source record.
        payload["raw_records"].append(
            payload["raw_records"][0]
        )

        with self.assertRaisesRegex(
            ValueError,
            "Duplicate source"
        ):
            calculate(payload)

    def test_seed_compatible_reports_zero_fallback(self):
        payload = example()

        # Remove one station.
        payload["raw_records"].pop()

        result = calculate(
            payload,
            strict=False
        )

        self.assertEqual(
            result["records"][0]["burns_dt"]["vandolah"],
            0
        )

        self.assertEqual(
            result["exceptions"][0]["kind"],
            "missing_lookup"
        )

    def test_nan_rejected(self):
        payload = example()

        payload["days"][0]["scheduled_volume_dt"] = "NaN"

        with self.assertRaises(ValueError):
            calculate(payload)

    def test_boolean_rejected(self):
        payload = example()

        payload["raw_records"][0]["energy_qty_dt"] = True

        with self.assertRaises(ValueError):
            calculate(payload)

    def test_date_order_rejected(self):
        payload = example()

        # Add duplicate date.
        payload["days"].append(
            payload["days"][0]
        )

        with self.assertRaisesRegex(
            ValueError,
            "Dates"
        ):
            calculate(payload)

    def test_changed_input_changes_result_without_cached_formulas(self):
        payload = example()

        before = calculate(payload)

        payload["days"][0]["scheduled_volume_dt"] = "90"

        after = calculate(payload)

        difference = (
            after["summary"]["ending_imbalance_dt"]
            - before["summary"]["ending_imbalance_dt"]
        )

        self.assertEqual(
            difference,
            10
        )

    def test_inputs_not_mutated(self):
        payload = example()

        original = copy.deepcopy(payload)

        calculate(payload)

        self.assertEqual(
            payload,
            original
        )


if __name__ == "__main__":
    unittest.main()