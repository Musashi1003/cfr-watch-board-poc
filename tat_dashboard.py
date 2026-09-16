from __future__ import annotations

import argparse
import calendar
import html
import importlib.util
import json
import tempfile
from collections import Counter
from dataclasses import dataclass
from datetime import date, datetime
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

from openpyxl import load_workbook
from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parent
STUDY_ROOT = ROOT.parent
EEPROM_ROOT = STUDY_ROOT / "20260813-EEPROM Data Study"
EEPROM_SCRIPT = EEPROM_ROOT / "generate_ppt_images.py"
OUTPUT_DIR = ROOT / "outputs"
ASSET_DIR = ROOT / "assets"
LOGO_PATH = ASSET_DIR / "inventec_logo.png"
DEFAULT_HOST = "127.0.0.1"
DEFAULT_PORT = 8500
HPSS_URL = "http://localhost:8501/"
DOA_URL = "http://localhost:8502/"
SERVER_HOST = DEFAULT_HOST
SERVER_PORT = DEFAULT_PORT

PALETTE = {
    "bg": "#F8E7E7",
    "panel": "#FFF8F8",
    "card": "#FFFFFF",
    "track": "#ECD0B6",
    "line": "#D8BDBD",
    "text": "#000000",
    "muted": "#4A4A4A",
    "brown": "#B3742F",
    "blue": "#96BFD2",
    "green": "#B0CD8D",
    "purple": "#B3A9B8",
    "good": "#5F9E43",
    "bad": "#C00000",
}


@dataclass
class DashboardResult:
    image_path: Path
    json_path: Path
    month_label: str
    total_cases: int
    input_name: str
    used_fallback_month: bool


def previous_full_month(today: date | None = None) -> tuple[int, int]:
    today = today or date.today()
    if today.month == 1:
        return today.year - 1, 12
    return today.year, today.month - 1


def previous_month(year: int, month: int) -> tuple[int, int]:
    if month == 1:
        return year - 1, 12
    return year, month - 1


def excel_weeknum_monday(value: date) -> int:
    return int(value.strftime("%U")) + 1


def month_label(year: int, month: int) -> str:
    return f"{year}-{month:02d}"


def dashboard_month_label(month_info: dict) -> str:
    year = int(month_info.get("year", 0))
    month = int(month_info.get("month", 0))
    if year and month:
        return f"{year % 100:02d}年{month}月"
    return safe_text(month_info.get("label", ""))


def default_feedback_month_value() -> str:
    year, month = previous_full_month()
    return month_label(year, month)


def parse_feedback_month(value: str) -> tuple[int, int] | None:
    value = safe_text(value)
    if not value:
        return None
    try:
        year_text, month_text = value.split("-", 1)
        year = int(year_text)
        month = int(month_text)
    except ValueError as exc:
        raise ValueError("Invalid feedback month format. Use YYYY-MM.") from exc
    if month < 1 or month > 12:
        raise ValueError("Invalid feedback month. Month must be 01-12.")
    return year, month


def safe_text(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip()


MONTH_LABELS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
MONTH_ALIASES = {name.lower(): index + 1 for index, name in enumerate(MONTH_LABELS)}


def parse_workbook_month(value: object) -> tuple[int, int] | None:
    text = safe_text(value)
    if not text:
        return None
    for separator in ("/", "-", "."):
        if separator in text:
            year_text, month_text = text.split(separator, 1)
            year_text = year_text.strip()
            month_text = month_text.strip()[:2]
            if year_text.isdigit() and month_text.isdigit():
                month = int(month_text)
                if 1 <= month <= 12:
                    return int(year_text), month
    return None


def parse_report_month(report_year: object, mom: object) -> tuple[int, int] | None:
    year_text = safe_text(report_year).replace("Y", "")
    mom_text = safe_text(mom)
    if not year_text.isdigit():
        return None
    if mom_text.isdigit():
        month = int(mom_text)
    else:
        month = MONTH_ALIASES.get(mom_text[:3].lower(), 0)
    if 1 <= month <= 12:
        return int(year_text), month
    return None


def row_month_key(row: dict) -> tuple[int, int]:
    return row.get("feedback_month") or (row["date"].year, row["date"].month)


def axis_ticks(max_value: int) -> tuple[int, list[int]]:
    if max_value <= 0:
        return 1, [0, 1]
    axis_max = max_value if max_value <= 10 else ((max_value + 9) // 10) * 10
    midpoint = axis_max // 2
    return axis_max, sorted(set([0, midpoint, axis_max]))


def darken_hex(hex_color: str, factor: float = 0.72) -> str:
    color = hex_color.strip().lstrip("#")
    red = int(int(color[0:2], 16) * factor)
    green = int(int(color[2:4], 16) * factor)
    blue = int(int(color[4:6], 16) * factor)
    return f"#{red:02X}{green:02X}{blue:02X}"


def load_rows(workbook_path: Path) -> list[dict]:
    wb = load_workbook(workbook_path, read_only=True, data_only=True)
    sheet_name = next((name for name in wb.sheetnames if name.startswith("New Template")), None)
    if not sheet_name:
        raise ValueError("Could not find a worksheet starting with 'New Template'.")

    ws = wb[sheet_name]
    rows: list[dict] = []
    for values in ws.iter_rows(min_row=2, max_col=31, values_only=True):
        if all(value in (None, "") for value in values):
            continue
        raw_date = values[8]
        if not hasattr(raw_date, "year"):
            continue
        feedback_month = parse_workbook_month(values[27] if len(values) > 27 else None)
        report_month = parse_report_month(values[28] if len(values) > 28 else None, values[29] if len(values) > 29 else None)
        rows.append(
            {
                "family": safe_text(values[1]),
                "model": safe_text(values[2]),
                "issue": safe_text(values[3]),
                "region": safe_text(values[5]),
                "date": raw_date,
                "week": values[9] if isinstance(values[9], (int, float)) else safe_text(values[9]),
                "case_status": safe_text(values[12]),
                "category": safe_text(values[13]),
                "part": safe_text(values[14]),
                "first_tat": values[19] if isinstance(values[19], (int, float)) else None,
                "third_tat": values[21] if isinstance(values[21], (int, float)) else None,
                "fifth_tat": values[23] if isinstance(values[23], (int, float)) else None,
                "tat_status": safe_text(values[24]),
                "problem_sorting": normalize_problem_sorting(safe_text(values[26])),
                "feedback_month": feedback_month or report_month or (raw_date.year, raw_date.month),
            }
        )
    wb.close()
    if not rows:
        raise ValueError("No usable data rows were found in 'New Template'.")
    return rows


def normalize_problem_sorting(value: str) -> str:
    replacements = {
        "": "Unspecified",
        "Wrong Cofig": "Wrong Config",
    }
    return replacements.get(value, value or "Unspecified")


def select_month_rows(rows: list[dict], target_year: int, target_month: int, allow_fallback: bool = True) -> tuple[list[dict], int, int, bool]:
    selected = [row for row in rows if row_month_key(row) == (target_year, target_month)]
    if selected:
        return selected, target_year, target_month, False
    if not allow_fallback:
        raise ValueError(f"No data found for feedback month {month_label(target_year, target_month)}.")
    latest_year, latest_month = max(row_month_key(row) for row in rows)
    fallback = [row for row in rows if row_month_key(row) == (latest_year, latest_month)]
    return fallback, latest_year, latest_month, True


def top_items(rows: list[dict], key: str, limit: int) -> list[tuple[str, int]]:
    counter = Counter(row[key] or "Unspecified" for row in rows)
    return counter.most_common(limit)


def is_long_tail_case(row: dict) -> bool:
    return isinstance(row.get("third_tat"), (int, float)) or isinstance(row.get("fifth_tat"), (int, float))


def month_metrics(rows: list[dict]) -> dict:
    total = len(rows)
    tat_pass = sum(1 for row in rows if row["tat_status"].upper() == "YES")
    confirmed = sum(1 for row in rows if row["case_status"].upper() == "YES")
    tat_values = [float(row["first_tat"]) for row in rows if isinstance(row["first_tat"], (int, float))]
    third_cases = sum(1 for row in rows if isinstance(row.get("third_tat"), (int, float)))
    fifth_cases = sum(1 for row in rows if isinstance(row.get("fifth_tat"), (int, float)))
    long_tail_cases = sum(1 for row in rows if is_long_tail_case(row))
    return {
        "total": total,
        "tat_pass": tat_pass,
        "tat_fail": total - tat_pass,
        "tat_rate": round(tat_pass / total * 100, 1) if total else 0,
        "confirmed": confirmed,
        "unconfirmed": total - confirmed,
        "avg_tat": round(sum(tat_values) / len(tat_values), 2) if tat_values else None,
        "max_tat": round(max(tat_values), 2) if tat_values else None,
        "third_cases": third_cases,
        "fifth_cases": fifth_cases,
        "long_tail_cases": long_tail_cases,
        "regions": len(set(row["region"] for row in rows if row["region"])),
    }

def trend_marker(current: float | None, previous: float | None, higher_is_better: bool) -> tuple[str, str]:
    if current is None or previous is None:
        return "", PALETTE["muted"]
    if current == previous:
        return "-", PALETTE["muted"]
    improved = current > previous if higher_is_better else current < previous
    if improved:
        return ("\u25b2" if higher_is_better else "\u25bc"), PALETTE["good"]
    return ("\u25bc" if higher_is_better else "\u25b2"), PALETTE["bad"]



def build_summary(rows: list[dict], year: int, month: int, input_name: str, used_fallback_month: bool) -> dict:
    current = [row for row in rows if row_month_key(row) == (year, month)]
    prev_year, prev_month = previous_month(year, month)
    previous = [row for row in rows if row_month_key(row) == (prev_year, prev_month)]
    current_metrics = month_metrics(current)
    previous_metrics = month_metrics(previous) if previous else None
    long_tail_rows = [row for row in current if is_long_tail_case(row)]
    month_counts = Counter(row_month_key(row) for row in rows)
    yoy_previous_year = year - 1
    yoy_current = [month_counts.get((year, item), 0) if item <= month else None for item in range(1, 13)]
    yoy_previous = [month_counts.get((yoy_previous_year, item), 0) for item in range(1, 13)]
    yoy_current_value = month_counts.get((year, month), 0)
    yoy_previous_value = month_counts.get((yoy_previous_year, month), 0)
    yoy_delta = yoy_current_value - yoy_previous_value
    yoy_delta_rate = round(yoy_delta / yoy_previous_value * 100) if yoy_previous_value else None
    quarter_confirmed_counts = Counter(
        (row_month_key(row)[0], ((row_month_key(row)[1] - 1) // 3) + 1)
        for row in rows
        if row["case_status"].upper() == "YES"
    )
    quarter_yoy = [
        {
            "label": f"Q{quarter}",
            "current": quarter_confirmed_counts.get((year, quarter), 0),
            "previous": quarter_confirmed_counts.get((yoy_previous_year, quarter), 0),
        }
        for quarter in range(1, 5)
    ]

    week_range = {
        excel_weeknum_monday(date(year, month, day))
        for day in range(1, calendar.monthrange(year, month)[1] + 1)
    }
    weeks = []
    for row in current:
        raw_week = row.get("week")
        if isinstance(raw_week, (int, float)):
            weeks.append(int(raw_week))
        else:
            text_week = safe_text(raw_week).upper().replace("WK", "").replace("W", "")
            if text_week.isdigit():
                weeks.append(int(text_week))
    week_counts = Counter(weeks)
    min_week = min(week_range) if week_range else min(weeks)
    max_week = max(week_range) if week_range else max(weeks)
    weekly_items = [{"week": week, "label": f"WK{week}", "count": week_counts.get(week, 0)} for week in range(min_week, max_week + 1)]
    return {
        "input_name": input_name,
        "used_fallback_month": used_fallback_month,
        "month": {"year": year, "month": month, "label": month_label(year, month)},
        "previous_month": {"year": prev_year, "month": prev_month, "label": month_label(prev_year, prev_month)},
        "metrics": current_metrics,
        "previous_metrics": previous_metrics,
        "charts": {
            "problem_sorting": top_items(current, "problem_sorting", 7),
            "category": top_items(current, "category", 5),
            "region": top_items(current, "region", 5),
            "family": top_items(current, "family", 5),
            "part": top_items(current, "part", 5),
            "long_tail_part": top_items(long_tail_rows, "part", 5),
            "weekly": weekly_items,
            "yoy": {
                "target_year": year,
                "previous_year": yoy_previous_year,
                "target_month": month,
                "current": yoy_current,
                "previous": yoy_previous,
                "current_value": yoy_current_value,
                "previous_value": yoy_previous_value,
                "delta": yoy_delta,
                "delta_rate": yoy_delta_rate,
            },
            "quarter_yoy": {
                "target_year": year,
                "previous_year": yoy_previous_year,
                "items": quarter_yoy,
            },
        },
    }


def find_font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    candidates = []
    if bold:
        candidates.extend([r"C:\Windows\Fonts\msjhbd.ttc", r"C:\Windows\Fonts\msjhb.ttc", r"C:\Windows\Fonts\segoeuib.ttf"])
    candidates.extend([r"C:\Windows\Fonts\msjh.ttc", r"C:\Windows\Fonts\segoeui.ttf"])
    for item in candidates:
        if Path(item).exists():
            return ImageFont.truetype(item, size=size)
    return ImageFont.load_default()


def rounded_panel(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int], fill: str, outline: str | None = None, radius: int = 20) -> None:
    draw.rounded_rectangle(box, radius=radius, fill=fill, outline=outline, width=2 if outline else 1)


def draw_text(draw: ImageDraw.ImageDraw, xy: tuple[int, int], text: str, font: ImageFont.ImageFont, fill: str, anchor: str | None = None) -> None:
    draw.text(xy, text, font=font, fill=fill, anchor=anchor)


def draw_logo(image: Image.Image, logo_path: Path, box: tuple[int, int, int, int]) -> None:
    if not logo_path.exists():
        return
    logo = Image.open(logo_path).convert("RGBA")
    alpha_bbox = logo.getbbox()
    if alpha_bbox:
        logo = logo.crop(alpha_bbox)
    left, top, right, bottom = box
    logo.thumbnail((right - left, bottom - top), Image.Resampling.LANCZOS)
    image.paste(logo, (right - logo.width, top + (bottom - top - logo.height) // 2), logo)


def draw_kpi_card(
    draw: ImageDraw.ImageDraw,
    box: tuple[int, int, int, int],
    title: str,
    value: str,
    subtitle: str,
    accent: str,
    status_color: str,
    previous_note: str,
    arrow: str,
    arrow_color: str,
    fonts: dict,
) -> None:
    rounded_panel(draw, box, fill=PALETTE["card"], outline=PALETTE["line"], radius=22)
    left, top, right, bottom = box
    draw.rounded_rectangle((left + 18, top + 18, left + 34, bottom - 18), radius=8, fill=accent)
    draw_text(draw, (left + 56, top + 20), title, fonts["label"], PALETTE["muted"])
    draw_text(draw, (left + 56, top + 58), value, fonts["kpi"], status_color)
    draw_text(draw, (left + 56, bottom - 34), subtitle, fonts["small"], PALETTE["muted"])
    if previous_note:
        draw_text(draw, (right - 22, bottom - 34), previous_note, fonts["small"], PALETTE["muted"], anchor="ra")
    if arrow:
        value_width = draw.textbbox((0, 0), value, font=fonts["kpi"])[2]
        arrow_x = min(left + 72 + value_width, right - 34)
        draw_text(draw, (arrow_x, top + 56), arrow, fonts["arrow"], arrow_color, anchor="la")


def draw_bar_chart(
    draw: ImageDraw.ImageDraw,
    box: tuple[int, int, int, int],
    title: str,
    items: list[tuple[str, int]],
    accent: str,
    fonts: dict,
    show_ratio: bool = False,
) -> None:
    rounded_panel(draw, box, fill=PALETTE["panel"], outline=PALETTE["line"], radius=22)
    left, top, right, bottom = box
    draw_text(draw, (left + 24, top + 18), title, fonts["title"], PALETTE["text"])
    if not items:
        draw_text(draw, ((left + right) // 2, (top + bottom) // 2), "No data this month", fonts["body"], PALETTE["muted"], anchor="mm")
        return
    total = sum(value for _, value in items)
    chart_top = top + 64
    row_height = max(31, (bottom - chart_top - 18) // len(items))
    label_width = int((right - left) * 0.33)
    bar_left = left + label_width + 18
    suffixes = [
        f"{value}  {value / total:.0%}" if total else str(value)
        for _, value in items
    ]
    suffix_width = max(draw.textbbox((0, 0), suffix, font=fonts["body"])[2] for suffix in suffixes)
    bar_right = max(bar_left + 40, right - 24 - suffix_width - 28)
    max_value = max(value for _, value in items) or 1
    for index, ((label, value), suffix) in enumerate(zip(items, suffixes)):
        y = chart_top + index * row_height
        label_text = label if len(label) <= 20 else f"{label[:18]}.."
        draw_text(draw, (left + 24, y + 3), label_text, fonts["body"], PALETTE["text"])
        draw.rounded_rectangle((bar_left, y + 8, bar_right, y + 24), radius=8, fill=PALETTE["track"])
        fill_width = int((bar_right - bar_left) * value / max_value)
        draw.rounded_rectangle((bar_left, y + 8, bar_left + max(fill_width, 14), y + 24), radius=8, fill=accent)
        draw_text(draw, (right - 24, y + 2), suffix, fonts["body"], PALETTE["text"], anchor="ra")


def draw_weekly_timeline(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int], title: str, items: list[dict], fonts: dict) -> None:
    rounded_panel(draw, box, fill=PALETTE["panel"], outline=PALETTE["line"], radius=22)
    left, top, right, bottom = box
    draw_text(draw, (left + 24, top + 18), title, fonts["title"], PALETTE["text"])
    if not items:
        return
    chart_left, chart_right = left + 30, right - 28
    chart_top, chart_bottom = top + 84, bottom - 50
    max_value = max(item["count"] for item in items) or 1
    axis_max, ticks = axis_ticks(max_value)
    for tick in ticks:
        y = chart_bottom - ((chart_bottom - chart_top) * tick / axis_max)
        draw.line((chart_left, y, chart_right, y), fill=PALETTE["line"], width=1)
        draw_text(draw, (chart_left - 10, y - 8), str(tick), fonts["tiny"], PALETTE["muted"], anchor="ra")
    points = []
    span = max(len(items) - 1, 1)
    for index, item in enumerate(items):
        x = chart_left + ((chart_right - chart_left) * index / span)
        y = chart_bottom - ((chart_bottom - chart_top) * item["count"] / axis_max)
        points.append((x, y))
    if len(points) > 1:
        draw.line(points, fill=PALETTE["blue"], width=5)
    for point, item in zip(points, items):
        x, y = point
        draw.ellipse((x - 6, y - 6, x + 6, y + 6), fill=PALETTE["blue"])
        draw_text(draw, (x, chart_bottom + 14), item["label"], fonts["tiny"], PALETTE["muted"], anchor="ma")


def draw_yoy_trend(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int], items: dict, fonts: dict) -> None:
    rounded_panel(draw, box, fill=PALETTE["panel"], outline=PALETTE["line"], radius=22)
    left, top, right, bottom = box
    draw_text(draw, (left + 24, top + 18), "YoY Monthly Case Trend", fonts["title"], PALETTE["text"])
    if not items:
        return

    current_color = PALETTE["brown"]
    previous_color = PALETTE["blue"]
    current_highlight = darken_hex(current_color, 0.62)
    previous_highlight = darken_hex(previous_color, 0.76)

    legend_y = top + 32
    legend_x = right - (230 if right - left < 1300 else 270)
    draw.line((legend_x, legend_y, legend_x + 42, legend_y), fill=current_color, width=5)
    draw_text(draw, (legend_x + 52, legend_y - 10), str(items["target_year"]), fonts["small"], PALETTE["muted"])
    draw.line((legend_x + 130, legend_y, legend_x + 172, legend_y), fill=previous_color, width=5)
    draw_text(draw, (legend_x + 182, legend_y - 10), str(items["previous_year"]), fonts["small"], PALETTE["muted"])

    chart_left, chart_right = left + 76, right - 56
    chart_top, chart_bottom = top + 76, bottom - 52
    values = [value for value in items["current"] if value is not None] + items["previous"]
    axis_max, ticks = axis_ticks(max(values or [1]))
    for tick in ticks:
        y = chart_bottom - ((chart_bottom - chart_top) * tick / axis_max)
        draw.line((chart_left, y, chart_right, y), fill=PALETTE["line"], width=1)
        draw_text(draw, (chart_left - 12, y - 8), str(tick), fonts["tiny"], PALETTE["muted"], anchor="ra")

    month_points = []
    for index, label in enumerate(MONTH_LABELS):
        x = chart_left + ((chart_right - chart_left) * index / 11)
        month_points.append(x)
        draw_text(draw, (x, chart_bottom + 18), label, fonts["tiny"], PALETTE["muted"], anchor="ma")

    def series_points(series: list[int | None]) -> list[tuple[float, float, int, int]]:
        points = []
        for index, value in enumerate(series):
            if value is None:
                continue
            x = month_points[index]
            y = chart_bottom - ((chart_bottom - chart_top) * value / axis_max)
            points.append((x, y, index + 1, value))
        return points

    previous_points = series_points(items["previous"])
    current_points = series_points(items["current"])
    if len(previous_points) > 1:
        draw.line([(x, y) for x, y, _, _ in previous_points], fill=previous_color, width=4)
    if len(current_points) > 1:
        draw.line([(x, y) for x, y, _, _ in current_points], fill=current_color, width=5)

    target_month = items["target_month"]
    for x, y, month, _ in previous_points:
        radius = 10 if month == target_month else 5
        fill = previous_highlight if month == target_month else previous_color
        draw.ellipse((x - radius, y - radius, x + radius, y + radius), fill=fill)
    for x, y, month, _ in current_points:
        radius = 14 if month == target_month else 6
        fill = current_highlight if month == target_month else current_color
        draw.ellipse((x - radius, y - radius, x + radius, y + radius), fill=fill, outline=current_color, width=4 if month == target_month else 1)

    target_index = target_month - 1
    if items.get("delta_rate") is None:
        delta_prefix = f"{MONTH_LABELS[target_index]}: {items['target_year']} {items['current_value']} vs {items['previous_year']} {items['previous_value']} | diff "
        delta_value = f"{items['delta']:+d}"
    else:
        delta_prefix = f"{MONTH_LABELS[target_index]} YoY: {items['target_year']} {items['current_value']} vs {items['previous_year']} {items['previous_value']} | "
        delta_value = f"{items['delta']:+d} ({items['delta_rate']:+.0f}%)"
    prefix_width = draw.textbbox((0, 0), delta_prefix, font=fonts["body"])[2]
    delta_width = draw.textbbox((0, 0), delta_value, font=fonts["body"])[2]
    text_width = prefix_width + delta_width
    pill_left = min(max(right - 720, chart_left), chart_right - text_width - 36)
    pill_top = top + 16
    draw.rounded_rectangle((pill_left, pill_top, pill_left + text_width + 36, pill_top + 38), radius=18, fill="#FFFFFF", outline=PALETTE["line"], width=1)
    draw_text(draw, (pill_left + 18, pill_top + 8), delta_prefix, fonts["body"], PALETTE["text"])
    delta_color = PALETTE["good"] if items["delta"] < 0 else PALETTE["bad"] if items["delta"] > 0 else PALETTE["muted"]
    draw_text(draw, (pill_left + 18 + prefix_width, pill_top + 8), delta_value, fonts["body"], delta_color)


def draw_quarter_yoy_bars(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int], items: dict, fonts: dict) -> None:
    rounded_panel(draw, box, fill=PALETTE["panel"], outline=PALETTE["line"], radius=22)
    left, top, right, bottom = box
    draw_text(draw, (left + 24, top + 18), "QoQ Confirmed Cases", fonts["title"], PALETTE["text"])
    if not items:
        return

    current_color = PALETTE["brown"]
    previous_color = PALETTE["blue"]
    legend_y = top + 32
    legend_x = right - 230
    draw.rectangle((legend_x, legend_y - 5, legend_x + 36, legend_y + 5), fill=current_color)
    draw_text(draw, (legend_x + 46, legend_y - 11), str(items["target_year"]), fonts["small"], PALETTE["muted"])
    draw.rectangle((legend_x + 120, legend_y - 5, legend_x + 156, legend_y + 5), fill=previous_color)
    draw_text(draw, (legend_x + 166, legend_y - 11), str(items["previous_year"]), fonts["small"], PALETTE["muted"])

    chart_left, chart_right = left + 52, right - 28
    chart_top, chart_bottom = top + 76, bottom - 48
    values = [item["current"] for item in items["items"]] + [item["previous"] for item in items["items"]]
    axis_max, ticks = axis_ticks(max(values or [1]))
    for tick in ticks:
        y = chart_bottom - ((chart_bottom - chart_top) * tick / axis_max)
        draw.line((chart_left, y, chart_right, y), fill=PALETTE["line"], width=1)
        draw_text(draw, (chart_left - 10, y - 8), str(tick), fonts["tiny"], PALETTE["muted"], anchor="ra")

    group_count = max(len(items["items"]), 1)
    group_width = (chart_right - chart_left) / group_count
    bar_width = min(32, max(18, int(group_width * 0.18)))
    for index, item in enumerate(items["items"]):
        center_x = chart_left + group_width * index + group_width / 2
        current_height = (chart_bottom - chart_top) * item["current"] / axis_max
        previous_height = (chart_bottom - chart_top) * item["previous"] / axis_max
        current_box = (
            int(center_x - bar_width - 4),
            int(chart_bottom - current_height),
            int(center_x - 4),
            chart_bottom,
        )
        previous_box = (
            int(center_x + 4),
            int(chart_bottom - previous_height),
            int(center_x + bar_width + 4),
            chart_bottom,
        )
        draw.rounded_rectangle(current_box, radius=5, fill=current_color)
        draw.rounded_rectangle(previous_box, radius=5, fill=previous_color)
        draw_text(draw, (center_x - bar_width / 2 - 4, current_box[1] - 18), str(item["current"]), fonts["tiny"], PALETTE["text"], anchor="ma")
        draw_text(draw, (center_x + bar_width / 2 + 4, previous_box[1] - 18), str(item["previous"]), fonts["tiny"], PALETTE["text"], anchor="ma")
        draw_text(draw, (center_x, chart_bottom + 16), item["label"], fonts["tiny"], PALETTE["muted"], anchor="ma")


def render_dashboard(summary: dict, image_path: Path) -> None:
    image = Image.new("RGB", (1920, 1080), PALETTE["bg"])
    draw = ImageDraw.Draw(image)
    fonts = {
        "hero": find_font(44, bold=True),
        "section": find_font(30, bold=True),
        "title": find_font(24, bold=True),
        "kpi": find_font(52, bold=True),
        "arrow": find_font(26, bold=True),
        "label": find_font(20, bold=True),
        "body": find_font(19),
        "small": find_font(17),
        "tiny": find_font(14),
    }
    metrics = summary["metrics"]
    prev = summary["previous_metrics"] or {}
    tat_color = PALETTE["good"] if metrics["tat_rate"] >= 80 else PALETTE["bad"]
    case_color = PALETTE["good"] if metrics["confirmed"] == 0 else PALETTE["bad"]

    total_arrow, total_arrow_color = trend_marker(metrics["total"], prev.get("total"), higher_is_better=False)
    tat_arrow, tat_arrow_color = trend_marker(metrics["tat_rate"], prev.get("tat_rate"), higher_is_better=True)

    draw_logo(image, LOGO_PATH, (1510, 38, 1848, 116))
    draw_text(draw, (72, 54), f"\u5e02\u5834\u5ba2\u8a34TAT\u7ba1\u7406 - {dashboard_month_label(summary['month'])}", fonts["hero"], PALETTE["text"])
    card_y, card_w, gap = 132, 554, 22
    cards = [
        ("案件總數", str(metrics["total"]), "", PALETTE["blue"], PALETTE["text"], f"較上期 {prev.get('total', '-')}" , total_arrow, total_arrow_color),
        ("TAT 達成率", f"{metrics['tat_rate']:.1f}%", f"達成 {metrics['tat_pass']} / 未達 {metrics['tat_fail']}", PALETTE["green"], tat_color, f"較上期 {prev.get('tat_rate', '-')}%", tat_arrow, tat_arrow_color),
        ("成案狀態", f"成案 {metrics['confirmed']}", f"未成案 {metrics['unconfirmed']}", PALETTE["purple"], case_color, "", "", PALETTE["muted"]),
    ]
    for index, card in enumerate(cards):
        left = 72 + index * (card_w + gap)
        draw_kpi_card(draw, (left, card_y, left + card_w, card_y + 158), *card, fonts=fonts)
    draw_yoy_trend(draw, (72, 316, 1200, 516), summary["charts"]["yoy"], fonts)
    draw_quarter_yoy_bars(draw, (1228, 316, 1778, 516), summary["charts"]["quarter_yoy"], fonts)
    draw_bar_chart(draw, (72, 532, 720, 815), "Problem Sorting", summary["charts"]["problem_sorting"], PALETTE["brown"], fonts)
    draw_bar_chart(draw, (748, 532, 1268, 815), "Cases by Category", summary["charts"]["category"], PALETTE["green"], fonts)
    draw_bar_chart(draw, (1296, 532, 1778, 815), "Cases by Region", summary["charts"]["region"], PALETTE["blue"], fonts, show_ratio=True)
    draw_bar_chart(draw, (72, 835, 622, 1050), "Cases by Family", summary["charts"]["family"], PALETTE["purple"], fonts)
    draw_weekly_timeline(draw, (650, 835, 1200, 1050), "Weekly Case Trend", summary["charts"]["weekly"], fonts)
    draw_bar_chart(draw, (1228, 835, 1778, 1050), "Parts Sorting", summary["charts"]["part"], PALETTE["brown"], fonts, show_ratio=True)

    image_path.parent.mkdir(parents=True, exist_ok=True)
    image.save(image_path, dpi=(144, 144))


def generate_dashboard(
    workbook_path: Path,
    output_dir: Path = OUTPUT_DIR,
    today: date | None = None,
    target_year: int | None = None,
    target_month: int | None = None,
) -> DashboardResult:
    rows = load_rows(workbook_path)
    explicit_month = target_year is not None and target_month is not None
    if target_year is None or target_month is None:
        target_year, target_month = previous_full_month(today)
    _, used_year, used_month, used_fallback = select_month_rows(rows, target_year, target_month, allow_fallback=not explicit_month)
    summary = build_summary(rows, used_year, used_month, workbook_path.name, used_fallback)

    run_dir = output_dir / f"{summary['month']['label']}_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    run_dir.mkdir(parents=True, exist_ok=True)
    image_path = run_dir / "dashboard.png"
    json_path = run_dir / "dashboard_summary.json"
    render_dashboard(summary, image_path)
    json_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    (output_dir / "latest_dashboard.png").write_bytes(image_path.read_bytes())
    (output_dir / "latest_summary.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    return DashboardResult(image_path, json_path, summary["month"]["label"], summary["metrics"]["total"], workbook_path.name, used_fallback)



def load_eeprom_module():
    if not EEPROM_SCRIPT.exists():
        raise FileNotFoundError(f"Cannot find EEPROM script: {EEPROM_SCRIPT}")
    spec = importlib.util.spec_from_file_location("eeprom_ppt_images", EEPROM_SCRIPT)
    if spec is None or spec.loader is None:
        raise RuntimeError("Cannot load EEPROM image generator.")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def generate_eeprom_images(workbook_path: Path) -> dict:
    module = load_eeprom_module()
    run_dir = OUTPUT_DIR / "eeprom" / datetime.now().strftime("%Y%m%d_%H%M%S")
    run_dir.mkdir(parents=True, exist_ok=True)
    module.OUTDIR = run_dir
    sheets = module.read_workbook(workbook_path)
    records, week_labels = module.build_records(sheets)
    gaming_path = module.render_combined_image(workbook_path, records, week_labels, "Gaming")
    pc_path = module.render_combined_image(workbook_path, records, week_labels, "PC")
    summary = {
        "input_name": workbook_path.name,
        "weeks": week_labels,
        "gaming": {
            "open": len(records.get("Gaming", {}).get("OPEN", [])),
            "action": len(records.get("Gaming", {}).get("ACTION", [])),
        },
        "pc": {
            "open": len(records.get("PC", {}).get("OPEN", [])),
            "action": len(records.get("PC", {}).get("ACTION", [])),
        },
        "unmapped": {
            "open": len(records.get("Unmapped", {}).get("OPEN", [])),
            "action": len(records.get("Unmapped", {}).get("ACTION", [])),
        },
        "gaming_image": gaming_path.name,
        "pc_image": pc_path.name,
    }
    summary_path = run_dir / "eeprom_summary.json"
    summary_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    latest_dir = OUTPUT_DIR / "eeprom_latest"
    latest_dir.mkdir(parents=True, exist_ok=True)
    (latest_dir / "Gaming_OPEN_ACTION.png").write_bytes(gaming_path.read_bytes())
    (latest_dir / "PC_OPEN_ACTION.png").write_bytes(pc_path.read_bytes())
    (latest_dir / "eeprom_summary.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    return {
        "run_dir": run_dir,
        "gaming_path": gaming_path,
        "pc_path": pc_path,
        "summary_path": summary_path,
        "summary": summary,
    }


def latest_eeprom_result() -> dict | None:
    latest_dir = OUTPUT_DIR / "eeprom_latest"
    summary_path = latest_dir / "eeprom_summary.json"
    gaming_path = latest_dir / "Gaming_OPEN_ACTION.png"
    pc_path = latest_dir / "PC_OPEN_ACTION.png"
    if not summary_path.exists() or not gaming_path.exists() or not pc_path.exists():
        return None
    try:
        summary = json.loads(summary_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None
    return {
        "gaming_path": gaming_path,
        "pc_path": pc_path,
        "summary_path": summary_path,
        "summary": summary,
    }


def eeprom_body(message: str = "", result: dict | None = None) -> str:
    result = result or latest_eeprom_result()
    message_html = f"<p>{html.escape(message)}</p>" if message else ""
    result_html = ""
    if result:
        summary = result["summary"]
        gaming_rel = "/" + result["gaming_path"].relative_to(ROOT).as_posix()
        pc_rel = "/" + result["pc_path"].relative_to(ROOT).as_posix()
        json_rel = "/" + result["summary_path"].relative_to(ROOT).as_posix()
        weeks = " / ".join(summary.get("weeks", []))
        result_html = f"""
      <div class="links">
        <a class="button" href="{gaming_rel}" download>Download Gaming PNG</a>
        <a class="button" href="{pc_rel}" download>Download PC PNG</a>
        <a class="button" href="{json_rel}" download>Download Summary JSON</a>
      </div>
      <div class="muted-card">Weeks: {html.escape(weeks)} | Gaming OPEN {summary.get('gaming', {}).get('open', 0)} / ACTION {summary.get('gaming', {}).get('action', 0)} | PC OPEN {summary.get('pc', {}).get('open', 0)} / ACTION {summary.get('pc', {}).get('action', 0)} | Unmapped OPEN {summary.get('unmapped', {}).get('open', 0)} / ACTION {summary.get('unmapped', {}).get('action', 0)}</div>
      <h2 style="margin-top:18px;">Gaming</h2>
      <div class="preview"><img src="{gaming_rel}" alt="Gaming EEPROM summary"></div>
      <h2 style="margin-top:18px;">PC</h2>
      <div class="preview"><img src="{pc_rel}" alt="PC EEPROM summary"></div>
"""
    return nav("eeprom") + f"""<main class="panel">
      <h1>EEPROM Item Summary</h1>
      <p>Upload an EEPROM Excel file to generate Gaming / PC OPEN + ACTION PPT images.</p>
      {message_html}
      <form method="post" enctype="multipart/form-data" action="/eeprom-upload">
        <input type="file" name="workbook" accept=".xlsm,.xlsx" required>
        <button type="submit">Generate EEPROM Images</button>
      </form>
      {result_html}
    </main>"""


def html_page(title: str, body: str) -> bytes:
    page = f"""<!doctype html>
<html lang="zh-Hant">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{html.escape(title)}</title>
  <style>
    :root {{ --bg:#f8e7e7; --panel:#fff8f8; --line:#d8bdbd; --text:#000; --muted:#4a4a4a; --accent:#b3742f; --soft:#ecd0b6; }}
    * {{ box-sizing:border-box; }}
    body {{ margin:0; font-family:"Microsoft JhengHei","Segoe UI",sans-serif; background:var(--bg); color:var(--text); }}
    .wrap {{ width:min(1120px, calc(100vw - 40px)); margin:32px auto; }}
    .shell {{ display:grid; grid-template-columns:240px 1fr; gap:22px; align-items:start; }}
    .nav, .panel {{ background:var(--panel); border:1px solid var(--line); border-radius:8px; padding:20px; }}
    .brand {{ font-size:28px; font-weight:800; margin-bottom:20px; }}
    .nav a {{ display:block; padding:12px 14px; border-radius:8px; text-decoration:none; color:var(--text); margin-bottom:8px; }}
    .nav a.active {{ background:var(--soft); font-weight:700; }}
    h1 {{ margin:0 0 8px; font-size:30px; }}
    h2 {{ margin:0 0 10px; font-size:22px; }}
    p {{ color:var(--muted); line-height:1.6; }}
    form {{ margin-top:18px; display:flex; gap:12px; flex-wrap:wrap; align-items:center; }}
    input[type=file], input[type=month] {{ background:#fff; border:1px dashed var(--line); border-radius:8px; color:var(--text); padding:14px; min-width:220px; flex:1 1 260px; }}
    button, .button {{ background:var(--accent); color:#fff; border:none; border-radius:8px; padding:13px 18px; font-weight:700; text-decoration:none; display:inline-block; cursor:pointer; }}
    .muted-card {{ margin-top:16px; border:1px solid var(--line); border-radius:8px; padding:16px; color:var(--muted); }}
    .preview {{ margin-top:18px; border:1px solid var(--line); border-radius:8px; overflow:hidden; background:#fff; }}
    .preview img {{ display:block; width:100%; height:auto; }}
    .links {{ display:flex; gap:12px; flex-wrap:wrap; margin-top:14px; }}
    .panel-head {{ display:flex; justify-content:space-between; gap:20px; align-items:flex-start; flex-wrap:wrap; }}
    .panel-head-main {{ flex:1 1 380px; min-width:0; }}
    .toolbar-form {{ margin-top:0; flex:0 1 480px; justify-content:flex-end; align-items:flex-end; }}
    .toolbar-field {{ min-width:160px; flex:1 1 180px; }}
    .toolbar-field span {{ display:block; color:var(--muted); font-size:13px; margin-bottom:6px; }}
    .toolbar-form input[type=file], .toolbar-form input[type=month] {{ min-width:0; width:100%; padding:12px 14px; }}
    .toolbar-form button {{ white-space:nowrap; }}
    .portal-grid {{ display:grid; grid-template-columns:repeat(auto-fit, minmax(260px, 1fr)); gap:16px; margin-top:18px; }}
    .portal-card {{ background:#fff; border:1px solid var(--line); border-radius:12px; padding:18px; }}
    .portal-card p {{ margin:0 0 14px; min-height:72px; }}
    .portal-card .button {{ margin-right:10px; margin-bottom:10px; }}
    .status-line {{ margin-top:8px; font-size:14px; color:var(--muted); }}
    .chip {{ display:inline-block; padding:4px 10px; border-radius:999px; background:var(--soft); color:var(--text); font-size:13px; font-weight:700; margin-bottom:10px; }}
  </style>
</head>
<body><div class="wrap"><div class="shell">{body}</div></div></body></html>"""
    return page.encode("utf-8")


def nav(active: str) -> str:
    return f"""<aside class="nav">
      <div class="brand">Local Tools</div>
      <a class="{ 'active' if active == 'portal' else '' }" href="/">總入口</a>
      <a class="{ 'active' if active == 'tat' else '' }" href="/tat">客訴 TAT</a>
      <a class="{ 'active' if active == 'eeprom' else '' }" href="/eeprom">EEPROM Item Summary</a>
      <a class="{ 'active' if active == 'hpss' else '' }" href="/hpss">HPSS 檢查</a>
      <a class="{ 'active' if active == 'doa' else '' }" href="/doa">FIR5 DOA Data Review</a>
    </aside>"""


def portal_body() -> str:
    latest = latest_result()
    latest_line = ""
    if latest:
        latest_line = (
            f"<div class=\"status-line\">最新 TAT：{html.escape(latest.month_label)} | "
            f"案件數 {latest.total_cases} | 檔名 {html.escape(latest.input_name)}</div>"
        )
    return nav("portal") + f"""<main class="panel">
      <h1>功能總入口</h1>
      <p>依不同 function 整理入口，先選工具，再進入對應頁面。</p>
      <div class="portal-grid">
        <section class="portal-card">
          <div class="chip">Dashboard</div>
          <h2>客訴 TAT</h2>
          <p>上傳 ASUS 客訴 Excel，產出 dashboard PNG 與 JSON，適合月報與簡報使用。</p>
          <a class="button" href="/tat">開啟 TAT 工具</a>
          {latest_line}
        </section>
        <section class="portal-card">
          <div class="chip">Score Check</div>
          <h2>HPSS Score</h2>
          <p>連到 HPSS 驗算工具。這個入口只負責導向，真正的 HPSS 服務需要另外啟動在 `localhost:8501`。</p>
          <a class="button" href="/hpss">開啟 HPSS 入口</a>
          <div class="status-line">如果打不開，代表 HPSS 服務本體還沒啟動。</div>
        </section>
        <section class="portal-card">
          <div class="chip">EEPROM</div>
          <h2>EEPROM Item Summary</h2>
          <p>Upload an EEPROM Excel file and generate PPT-ready PNG summaries for Gaming / PC OPEN and ACTION items.</p>
          <a class="button" href="/eeprom">?? EEPROM Item Summary</a>
          <div class="status-line">Supports .xlsm / .xlsx. Outputs one Gaming PNG and one PC PNG.</div>
        </section>
        <section class="portal-card">
          <div class="chip">DOA Review</div>
          <h2>FIR5 DOA Data Review</h2>
          <p>上傳 FIR5 DOA original Excel，依 Check 歷史經驗與 Judgment 規則新增 TYPE X Sorting 欄位。</p>
          <a class="button" href="/doa">開啟 DOA 入口</a>
          <div class="status-line">DOA 服務本體運行在 localhost:8502。</div>
        </section>
      </div>
    </main>"""


def home_body(message: str = "") -> str:
    message_html = f"<p>{html.escape(message)}</p>" if message else ""
    default_month = default_feedback_month_value()
    return nav("tat") + f"""<main class="panel">
      <h1>客訴 TAT Dashboard</h1>
      <p>上傳 ASUS 客訴 Excel，產出可用於 PPT 的 dashboard PNG。</p>
      {message_html}
      <form method="post" enctype="multipart/form-data" action="/upload">
        <input type="file" name="workbook" accept=".xlsx" required>
        <label>
          <span style="display:block; color:var(--muted); font-size:13px; margin-bottom:6px;">回饋月份</span>
          <input type="month" name="feedback_month" value="{default_month}">
        </label>
        <button type="submit">產生 Dashboard</button>
      </form>
      <div class="muted-card">規則：TAT 達成率 >= 90% 顯示綠色；TAT狀態為 0 最佳；成案件數 > 0 顯示紅色。</div>
    </main>"""


def latest_result() -> DashboardResult | None:
    summary_path = OUTPUT_DIR / "latest_summary.json"
    image_path = OUTPUT_DIR / "latest_dashboard.png"
    if not summary_path.exists() or not image_path.exists():
        return None
    try:
        summary = json.loads(summary_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None
    month = summary.get("month", {}).get("label", "-")
    metrics = summary.get("metrics", {})
    return DashboardResult(
        image_path=image_path,
        json_path=summary_path,
        month_label=month,
        total_cases=int(metrics.get("total", 0) or 0),
        input_name=safe_text(summary.get("input_name")) or image_path.name,
        used_fallback_month=bool(summary.get("used_fallback_month")),
    )


def hpss_body() -> str:
    return nav("hpss") + f"""<main class="panel">
      <h1>HPSS 檢查</h1>
      <p>HPSS 請使用 <strong>localhost:8501</strong>，這個 Local Tools 頁面運行在 http://{SERVER_HOST}:{SERVER_PORT}/ 。</p>
      <div class="links">
        <a class="button" href="{HPSS_URL}" target="_blank" rel="noreferrer">開啟 HPSS</a>
        <a class="button" href="/tat">回到客訴 TAT</a>
        <a class="button" href="/">回總入口</a>
      </div>
      <div class="muted-card">如果 HPSS 無法開啟，代表 localhost:8501 沒有啟動。這個頁面只是總入口，不是 HPSS 本體。</div>
    </main>"""


def doa_body() -> str:
    return nav("doa") + f"""<main class="panel">
      <h1>FIR5 DOA Data Review</h1>
      <p>FIR5 DOA Data Review 請使用 <strong>localhost:8502</strong>，這個 Local Tools 頁面運行在 http://{SERVER_HOST}:{SERVER_PORT}/ 。</p>
      <div class="links">
        <a class="button" href="{DOA_URL}" target="_blank" rel="noreferrer">開啟 FIR5 DOA Data Review</a>
        <a class="button" href="/">回總入口</a>
      </div>
      <div class="muted-card">如果 DOA 無法開啟，代表 localhost:8502 沒有啟動。請從 Start_Field_Tools.bat 或 DOA 專案的 Start_DOA.bat 啟動。</div>
    </main>"""


def result_body(result: DashboardResult) -> str:
    image_rel = "/" + result.image_path.relative_to(ROOT).as_posix()
    json_rel = "/" + result.json_path.relative_to(ROOT).as_posix()
    return nav("tat") + f"""<main class="panel">
      <h1>Dashboard 已完成</h1>
      <p>檔案：<strong>{html.escape(result.input_name)}</strong> | 月份：<strong>{html.escape(result.month_label)}</strong> | 案件數：<strong>{result.total_cases}</strong></p>
      <div class="links">
        <a class="button" href="{image_rel}" download>下載簡報尺寸 PNG</a>
        <a class="button" href="{json_rel}" download>下載 JSON</a>
        <a class="button" href="/tat">再上傳一個檔案</a>
        <a class="button" href="/">回總入口</a>
      </div>
      <div class="preview"><img src="{image_rel}" alt="Dashboard 預覽"></div>
    </main>"""


def result_body_compact(result: DashboardResult) -> str:
    image_rel = "/" + result.image_path.relative_to(ROOT).as_posix()
    json_rel = "/" + result.json_path.relative_to(ROOT).as_posix()
    default_month = result.month_label if result.month_label else default_feedback_month_value()
    return nav("tat") + f"""<main class="panel">
      <div class="panel-head">
        <div class="panel-head-main">
          <h1>Dashboard</h1>
          <p>檔案：<strong>{html.escape(result.input_name)}</strong> | 月份：<strong>{html.escape(result.month_label)}</strong> | 案件數：<strong>{result.total_cases}</strong></p>
        </div>
        <form class="toolbar-form" method="post" enctype="multipart/form-data" action="/upload">
          <label class="toolbar-field">
            <span>上傳檔案</span>
            <input type="file" name="workbook" accept=".xlsx" required>
          </label>
          <label class="toolbar-field">
            <span>回饋月份</span>
            <input type="month" name="feedback_month" value="{html.escape(default_month)}">
          </label>
          <button type="submit">更新 Dashboard</button>
        </form>
      </div>
      <div class="links">
        <a class="button" href="{image_rel}" download>下載簡報尺寸 PNG</a>
        <a class="button" href="{json_rel}" download>下載 JSON</a>
        <a class="button" href="/">回總入口</a>
      </div>
      <div class="preview"><img src="{image_rel}" alt="Dashboard 預覽"></div>
    </main>"""
class DashboardHandler(BaseHTTPRequestHandler):
    server_version = "LocalTools/1.0"

    def do_GET(self) -> None:
        if self.path in ("/", ""):
            self.respond(HTTPStatus.OK, html_page("Local Tools", portal_body()))
        elif self.path == "/tat":
            latest = latest_result()
            self.respond(HTTPStatus.OK, html_page("Local Tools", result_body_compact(latest) if latest else home_body()))
        elif self.path == "/hpss":
            self.respond(HTTPStatus.OK, html_page("HPSS", hpss_body()))
        elif self.path == "/doa":
            self.respond(HTTPStatus.OK, html_page("FIR5 DOA", doa_body()))
        elif self.path == "/eeprom":
            self.respond(HTTPStatus.OK, html_page("EEPROM", eeprom_body()))
        elif self.path.startswith("/outputs/"):
            self.serve_file(ROOT / self.path.lstrip("/"))
        else:
            self.respond(HTTPStatus.NOT_FOUND, html_page("Not Found", home_body("找不到頁面。")))

    def do_POST(self) -> None:
        import cgi

        if self.path == "/eeprom-upload":
            self.handle_eeprom_upload()
            return
        if self.path != "/upload":
            self.respond(HTTPStatus.NOT_FOUND, html_page("Not Found", home_body("找不到頁面。")))
            return
        form = cgi.FieldStorage(
            fp=self.rfile,
            headers=self.headers,
            environ={"REQUEST_METHOD": "POST", "CONTENT_TYPE": self.headers.get("Content-Type", "")},
        )
        file_item = form["workbook"] if "workbook" in form else None
        if file_item is None or not getattr(file_item, "filename", ""):
            self.respond(HTTPStatus.BAD_REQUEST, html_page("Upload Error", home_body("請先選擇 Excel 檔案。")))
            return
        selected_month = parse_feedback_month(form.getfirst("feedback_month", ""))
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            tmp.write(file_item.file.read())
            temp_path = Path(tmp.name)
        try:
            if selected_month:
                result = generate_dashboard(temp_path, target_year=selected_month[0], target_month=selected_month[1])
            else:
                result = generate_dashboard(temp_path)
        except Exception as exc:
            self.respond(HTTPStatus.BAD_REQUEST, html_page("Export Failed", home_body(f"產出失敗：{exc}")))
        else:
            self.respond(HTTPStatus.OK, html_page("Dashboard Ready", result_body_compact(result)))
        finally:
            try:
                temp_path.unlink(missing_ok=True)
            except PermissionError:
                pass


    def handle_eeprom_upload(self) -> None:
        import cgi

        form = cgi.FieldStorage(
            fp=self.rfile,
            headers=self.headers,
            environ={"REQUEST_METHOD": "POST", "CONTENT_TYPE": self.headers.get("Content-Type", "")},
        )
        file_item = form["workbook"] if "workbook" in form else None
        if file_item is None or not getattr(file_item, "filename", ""):
            self.respond(HTTPStatus.BAD_REQUEST, html_page("Upload Error", eeprom_body("Please choose an EEPROM Excel file.")))
            return
        suffix = Path(file_item.filename).suffix.lower()
        if suffix not in (".xlsm", ".xlsx"):
            self.respond(HTTPStatus.BAD_REQUEST, html_page("Upload Error", eeprom_body("Please upload a .xlsm or .xlsx file.")))
            return
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            tmp.write(file_item.file.read())
            temp_path = Path(tmp.name)
        display_path = temp_path.with_name(Path(file_item.filename).name)
        try:
            temp_path.replace(display_path)
            result = generate_eeprom_images(display_path)
        except Exception as exc:
            self.respond(HTTPStatus.BAD_REQUEST, html_page("EEPROM Export Failed", eeprom_body(f"EEPROM Item Summary???{exc}")))
        else:
            self.respond(HTTPStatus.OK, html_page("EEPROM Ready", eeprom_body("EEPROM Item Summary??", result)))
        finally:
            try:
                display_path.unlink(missing_ok=True)
            except (PermissionError, UnboundLocalError):
                pass
            try:
                temp_path.unlink(missing_ok=True)
            except PermissionError:
                pass

    def serve_file(self, path: Path) -> None:
        if not path.exists() or not path.is_file():
            self.respond(HTTPStatus.NOT_FOUND, html_page("Not Found", home_body("找不到指定檔案。")))
            return
        content_type = "image/png" if path.suffix.lower() == ".png" else "application/json; charset=utf-8"
        data = path.read_bytes()
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def respond(self, status: HTTPStatus, payload: bytes) -> None:
        self.send_response(status)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)



def run_server(host: str, port: int) -> None:
    global SERVER_HOST, SERVER_PORT
    SERVER_HOST = host
    SERVER_PORT = port
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    server = ThreadingHTTPServer((host, port), DashboardHandler)
    print(f"Local tools server running at http://localhost:{port}/")
    server.serve_forever()


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Local dashboard tools.")
    parser.add_argument("--input", type=Path, help="Path to the ASUS workbook.")
    parser.add_argument("--output-dir", type=Path, default=OUTPUT_DIR, help="Directory for exported dashboard files.")
    parser.add_argument("--month", help="Feedback month to export, in YYYY-MM format.")
    parser.add_argument("--host", default=DEFAULT_HOST)
    parser.add_argument("--port", default=DEFAULT_PORT, type=int)
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    if args.input:
        selected_month = parse_feedback_month(args.month or "")
        if selected_month:
            result = generate_dashboard(args.input, args.output_dir, target_year=selected_month[0], target_month=selected_month[1])
        else:
            result = generate_dashboard(args.input, args.output_dir)
        print(f"Dashboard exported: {result.image_path}")
        print(f"Summary exported: {result.json_path}")
        return
    run_server(args.host, args.port)


if __name__ == "__main__":
    main()
