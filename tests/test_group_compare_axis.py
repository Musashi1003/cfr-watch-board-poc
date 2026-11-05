from unittest import TestCase, main
from unittest.mock import MagicMock, patch

import pandas as pd

import streamlit_app as app


def trend(weeks):
    return pd.DataFrame({
        "end_week": weeks,
        "weekly_failure": [1] * len(weeks),
        "cumulative_failure": list(range(1, len(weeks) + 1)),
        "cumulative_activation": [1000] * len(weeks),
        "cumulative_cfr": [i / 1000 for i in range(1, len(weeks) + 1)],
    })


class GroupComparisonAxisTests(TestCase):
    def render(self, frames, metric="Cumulative CFR"):
        groups = [{"label": f"G{i + 1}", "filters": {}} for i in range(len(frames))]
        ui = MagicMock()
        ui.session_state = {"group_compare_groups": groups}
        ui.columns.side_effect = lambda spec: [MagicMock() for _ in range(spec if isinstance(spec, int) else len(spec))]
        ui.button.return_value = False
        ui.text_input.return_value = "Group"
        ui.radio.return_value = metric
        with (
            patch.object(app, "st", ui),
            patch.object(app, "filter_controls", return_value={}),
            patch.object(app, "render_group_insight_cards"),
            patch.object(app, "cumulative_cfr_trend_data", side_effect=[(f, "no data") for f in frames]),
        ):
            app.render_group_compare([])
        return ui, ui.altair_chart.call_args.args[0].to_dict()

    def test_two_groups_align_from_first_point_for_both_metrics(self):
        first = trend(["W2526", "W2527", "W2618"])
        second = trend(["W2606", "W2607", "W2643"])
        for metric in ("Cumulative CFR", "Weekly Failure Qty"):
            with self.subTest(metric=metric):
                ui, chart = self.render([first, second], metric)
                self.assertEqual("relative_week_label", chart["encoding"]["x"]["field"])
                self.assertEqual("Comparison Week", chart["encoding"]["x"]["axis"]["title"])
                data = next(iter(chart["datasets"].values()))
                for label, original in (("G1", first), ("G2", second)):
                    rows = [r for r in data if r["group"] == label]
                    self.assertEqual("WK01", rows[0]["relative_week_label"])
                    self.assertEqual(original["end_week"].tolist(), [r["end_week"] for r in rows])
                    self.assertEqual(original["cumulative_cfr"].tolist(), [r["cumulative_cfr"] for r in rows])
                summary = ui.dataframe.call_args.args[0]
                self.assertEqual(["W2618", "W2643"], summary["Latest Week"].tolist())

    def test_single_group_and_single_nonempty_group_use_natural_weeks(self):
        for frames in ([trend(["W2606", "W2607"])], [pd.DataFrame(), trend(["W2606", "W2607"])]):
            with self.subTest(groups=len(frames)):
                _, chart = self.render(frames)
                self.assertEqual("end_week", chart["encoding"]["x"]["field"])
                self.assertNotIn("relative_week_label", [t["field"] for t in chart["encoding"]["tooltip"]])

    def test_same_year_and_mixed_year_groups_also_align(self):
        self.assertTrue(app.should_use_relative_week_axis([], [{"filters": {}}, {"filters": {}}]))
        _, chart = self.render([trend(["W2606"]), trend(["W2626"])])
        self.assertEqual("relative_week_label", chart["encoding"]["x"]["field"])

    def test_elapsed_weeks_preserve_gaps_and_53_week_year(self):
        result = app.add_relative_week_axis(trend(["W2702", "W2652", "W2653"]))
        self.assertEqual(["W2652", "W2653", "W2702"], result["end_week"].tolist())
        self.assertEqual([1, 2, 4], result["relative_week_index"].tolist())


if __name__ == "__main__":
    main()
