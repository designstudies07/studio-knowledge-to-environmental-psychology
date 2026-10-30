from __future__ import annotations

import csv
import math
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = ROOT / "data" / "derived"
FIGURE_DIR = ROOT / "figures"

INK = "#17212B"
NAVY = "#17324D"
TEAL = "#0F766E"
AMBER = "#D99018"
MID = "#66727D"
GRID = "#D8E1E8"
PALE = "#EEF4F7"
WHITE = "#FFFFFF"


def font(size: int, bold: bool = False, italic: bool = False) -> ImageFont.FreeTypeFont:
    names = []
    if bold and italic:
        names = ["arialbi.ttf", "DejaVuSans-BoldOblique.ttf"]
    elif bold:
        names = ["arialbd.ttf", "DejaVuSans-Bold.ttf"]
    elif italic:
        names = ["ariali.ttf", "DejaVuSans-Oblique.ttf"]
    else:
        names = ["arial.ttf", "DejaVuSans.ttf"]
    roots = [Path("C:/Windows/Fonts"), Path("/usr/share/fonts/truetype/dejavu")]
    for root in roots:
        for name in names:
            candidate = root / name
            if candidate.exists():
                return ImageFont.truetype(str(candidate), size=size)
    return ImageFont.load_default(size=size)


def centered(draw: ImageDraw.ImageDraw, xy: tuple[float, float], text: str, text_font, fill: str) -> None:
    box = draw.textbbox((0, 0), text, font=text_font)
    draw.text((xy[0] - (box[2] - box[0]) / 2, xy[1] - (box[3] - box[1]) / 2), text, font=text_font, fill=fill)


def right(draw: ImageDraw.ImageDraw, xy: tuple[float, float], text: str, text_font, fill: str) -> None:
    box = draw.textbbox((0, 0), text, font=text_font)
    draw.text((xy[0] - (box[2] - box[0]), xy[1]), text, font=text_font, fill=fill)


def read_rows(name: str) -> list[dict[str, str]]:
    with (DATA_DIR / name).open(newline="", encoding="utf-8") as handle:
        return list(csv.DictReader(handle))


def build_composite_summary() -> None:
    rows = read_rows("composite_trajectory.csv")
    width, height = 2600, 1500
    image = Image.new("RGB", (width, height), WHITE)
    draw = ImageDraw.Draw(image, "RGBA")
    draw.text((150, 80), "Composite person-space self-report across three stages", font=font(60, bold=True), fill=INK)
    draw.text((150, 155), "Matched sample  n = 34  |  aggregate means and 95% confidence intervals", font=font(34), fill=MID)

    left, right_edge, top, bottom = 320, 2300, 300, 1240
    x_positions = [490, 1300, 2110]
    labels = [("T0", "Before lecture"), ("T1", "After Lecture 1"), ("T2", "After assignment")]
    colors = [NAVY, TEAL, AMBER]

    def y_map(value: float) -> float:
        return bottom - (value - 1) / 4 * (bottom - top)

    for tick in range(1, 6):
        y = y_map(float(tick))
        draw.line((left, y, right_edge, y), fill=GRID, width=2)
        right(draw, (left - 35, y - 20), str(tick), font(30), MID)
    draw.text((65, 640), "Mean score", font=font(34, bold=True), fill=INK)

    means = [float(row["mean"]) for row in rows]
    n_values = [int(row["n"]) for row in rows]
    sds = [float(row["sd"]) for row in rows]
    points = [(x_positions[index], y_map(means[index])) for index in range(3)]
    draw.line(points, fill=INK, width=12)

    t_critical = 2.0345
    for index, (row, mean, n_value, sd, color) in enumerate(zip(rows, means, n_values, sds, colors)):
        margin = t_critical * sd / math.sqrt(n_value)
        low, high = mean - margin, mean + margin
        x = x_positions[index]
        draw.line((x, y_map(low), x, y_map(high)), fill=color, width=8)
        draw.line((x - 28, y_map(low), x + 28, y_map(low)), fill=color, width=8)
        draw.line((x - 28, y_map(high), x + 28, y_map(high)), fill=color, width=8)
        y = y_map(mean)
        draw.ellipse((x - 29, y - 29, x + 29, y + 29), fill=color, outline=WHITE, width=5)
        centered(draw, (x, y - 78), f"{mean:.2f}", font(37, bold=True), color)
        centered(draw, (x, bottom + 85), labels[index][0], font(32, bold=True), INK)
        centered(draw, (x, bottom + 130), labels[index][1], font(31), INK)

    draw.text(
        (320, 1405),
        "Points show means; bars show 95% confidence intervals computed from the published n, mean, and SD.",
        font=font(28, italic=True),
        fill=MID,
    )
    image.save(FIGURE_DIR / "figure_1_composite_trajectory.png", dpi=(300, 300), optimize=True)


def build_exploratory_summary() -> None:
    rows = {row["analysis"]: row for row in read_rows("exploratory_associations.csv")}
    studio = rows["Studio progression and assignment total"]
    detail = rows["Assignment detail and T1-T2 change"]

    width, height = 2750, 1600
    image = Image.new("RGB", (width, height), WHITE)
    draw = ImageDraw.Draw(image, "RGBA")
    draw.text((120, 65), "Exploratory associations with studio progression and assignment detail", font=font(56, bold=True), fill=INK)
    draw.text((120, 140), "Aggregate estimates for matched cases; no participant-level points are displayed", font=font(34), fill=MID)

    # Left panel: one aggregate association estimate.
    left, right_edge = 180, 1320
    draw.text((left, 255), "Studio progression and assignment score", font=font(35, bold=True), fill=INK)
    draw.text((left, 320), f"Spearman association  n = {studio['n']}", font=font(29), fill=MID)
    axis_y = 760
    axis_left, axis_right = left + 70, right_edge - 70
    draw.line((axis_left, axis_y, axis_right, axis_y), fill=GRID, width=7)
    for tick in [-1.0, -0.5, 0.0, 0.5, 1.0]:
        x = axis_left + (tick + 1) / 2 * (axis_right - axis_left)
        draw.line((x, axis_y - 18, x, axis_y + 18), fill=INK if tick == 0 else MID, width=4)
        centered(draw, (x, axis_y + 58), f"{tick:+.1f}", font(26), MID)
    estimate = float(studio["estimate"])
    estimate_x = axis_left + (estimate + 1) / 2 * (axis_right - axis_left)
    draw.ellipse((estimate_x - 30, axis_y - 30, estimate_x + 30, axis_y + 30), fill=TEAL, outline=WHITE, width=5)
    centered(draw, (estimate_x, axis_y - 70), f"rho = {estimate:.2f}", font(31, bold=True), TEAL)
    centered(draw, ((axis_left + axis_right) / 2, 930), f"Holm-adjusted p = {float(studio['p_value']):.3f}", font(30, bold=True), NAVY)
    centered(draw, ((axis_left + axis_right) / 2, 1010), "Higher studio progression was associated with higher assignment scores.", font(27), INK)

    # Right panel: two aggregate group means.
    panel_left, panel_right = 1510, 2580
    draw.text((panel_left, 255), "T1 to T2 change by assignment detail", font=font(35, bold=True), fill=INK)
    draw.text((panel_left, 320), f"Matched assignments  n = {detail['n']}", font=font(29), fill=MID)
    plot_left, plot_right = panel_left + 110, panel_right - 90
    row_y = [650, 880]
    x_min, x_max = -0.5, 0.1
    for tick in [-0.5, -0.3, -0.1, 0.1]:
        x = plot_left + (tick - x_min) / (x_max - x_min) * (plot_right - plot_left)
        draw.line((x, 500, x, 1030), fill=GRID, width=2)
        centered(draw, (x, 1085), f"{tick:+.1f}", font(26), MID)
    zero_x = plot_left + (0 - x_min) / (x_max - x_min) * (plot_right - plot_left)
    draw.line((zero_x, 500, zero_x, 1030), fill=MID, width=4)
    group_specs = [
        (detail["group_1"], int(detail["group_1_n"]), float(detail["group_1_mean"]), TEAL),
        (detail["group_2"], int(detail["group_2_n"]), float(detail["group_2_mean"]), AMBER),
    ]
    for y, (label, n_value, mean, color) in zip(row_y, group_specs):
        draw.text((panel_left, y - 25), f"{label}  n = {n_value}", font=font(29, bold=True), fill=INK)
        x = plot_left + (mean - x_min) / (x_max - x_min) * (plot_right - plot_left)
        draw.polygon([(x, y - 33), (x + 33, y), (x, y + 33), (x - 33, y)], fill=color)
        draw.text((x + 48, y - 20), f"mean {mean:+.2f}", font=font(28, bold=True), fill=color)
    centered(draw, ((plot_left + plot_right) / 2, 1160), "Mean T2 minus T1 change", font(28, bold=True), INK)
    centered(draw, ((plot_left + plot_right) / 2, 1250), f"Mann-Whitney p = {float(detail['p_value']):.3f}  |  Cliff delta = +0.12", font(29, bold=True), NAVY)

    draw.text((120, 1505), "The displayed values are aggregate summaries and do not permit reconstruction of individual response patterns.", font=font(27, italic=True), fill=MID)
    image.save(FIGURE_DIR / "figure_4_studio_and_subgroup.png", dpi=(300, 300), optimize=True)


if __name__ == "__main__":
    FIGURE_DIR.mkdir(parents=True, exist_ok=True)
    build_composite_summary()
    build_exploratory_summary()
    print("Updated privacy-safe aggregate figures 1 and 4.")
