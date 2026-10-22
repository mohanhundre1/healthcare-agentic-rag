import copy
import unittest
from decimal import Decimal

from def_i_imbalance.calculations import (
    calculate_i_imbalance,
)


def example_i_fpc_oa():
    return {
        "schema_version": "1",
        "records": [
            {
                "date": "2026-08-01",
                "burns_dt": {
                    "florida (gainesville)": "-10",
                    "bartow b-peaker": "-20",
                    "suwannee": "-30",
                    "hines": "-40",
                    "debary": "-50",
                    "anclote": "-60",
                    "tiger bay": "-70",
                    "vandolah": "-80",
                },
                "btu_factors": {
                    "florida (gainesville)": "1023",
                    "bartow b-peaker": "1024",
                    "suwannee": "1025",
                    "hines": "1026",
                    "debary": "1027",
                    "anclote": "1028",
                    "tiger bay": "1029",
                    "vandolah": "1030",
                },
            }
        ],
        "summary": {
            "ending_imbalance_dt": "100"
        },
    }


class IImbalanceCalculationTests(unittest.TestCase):

    def test_creates_eight_station_rows(self):
        result = calculate_i_imbalance(
            example_i_fpc_oa()
        )

        stations = result["records"][0]["stations"]

        self.assertEqual(
            len(stations),
            8,
        )

    def test_station_crosswalk(self):
        result = calculate_i_imbalance(
            example_i_fpc_oa()
        )

        first_station = (
            result["records"][0]["stations"][0]
        )

        self.assertEqual(
            first_station["station"],
            "Univ of FL",
        )

    def test_mmbtu_value_is_preserved(self):
        result = calculate_i_imbalance(
            example_i_fpc_oa()
        )

        first_station = (
            result["records"][0]["stations"][0]
        )

        self.assertEqual(
            first_station["MMBTU-FGT"],
            Decimal("-10"),
        )

    def test_heat_rate_is_preserved(self):
        result = calculate_i_imbalance(
            example_i_fpc_oa()
        )

        hines = result["records"][0]["stations"][3]

        self.assertEqual(
            hines["station"],
            "Hines",
        )

        self.assertEqual(
            hines["HEAT RATE-FGT"],
            Decimal("1026"),
        )

    def test_missing_station_rejected(self):
        payload = example_i_fpc_oa()

        del payload["records"][0]["burns_dt"]["vandolah"]

        with self.assertRaisesRegex(
            ValueError,
            "UNMAPPED_FGT_STATION",
        ):
            calculate_i_imbalance(payload)

    def test_missing_heat_rate_rejected(self):
        payload = example_i_fpc_oa()

        del payload["records"][0]["btu_factors"]["hines"]

        with self.assertRaisesRegex(
            ValueError,
            "MISSING_FGT_HEAT_RATE",
        ):
            calculate_i_imbalance(payload)

    def test_non_strict_records_exceptions(self):
        payload = example_i_fpc_oa()

        del payload["records"][0]["burns_dt"]["vandolah"]

        result = calculate_i_imbalance(
            payload,
            strict=False,
        )

        self.assertGreater(
            len(result["exceptions"]),
            0,
        )

    def test_source_input_not_modified(self):
        payload = example_i_fpc_oa()
        original = copy.deepcopy(payload)

        calculate_i_imbalance(payload)

        self.assertEqual(
            payload,
            original,
        )


if __name__ == "__main__":
    unittest.main()