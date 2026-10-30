from __future__ import annotations

import hashlib
import json
import math
import os
import re
import unicodedata
from collections import Counter
from pathlib import Path

import numpy as np
import openpyxl
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor


ROOT = Path(__file__).resolve().parents[1]
PRIVATE_DIR = ROOT / "data" / "private"
SOURCE_WORKBOOK = Path(
    os.environ.get(
        "STUDY_ANALYSIS_WORKBOOK",
        PRIVATE_DIR / "Three_Stage_Analysis_Single_Researcher.xlsx",
    )
)
T0_FILE = Path(
    os.environ.get(
        "STUDY_T0_WORKBOOK",
        PRIVATE_DIR / "Pre-Lecture_Test_Responses.xlsx",
    )
)
OUTPUT_DIR = ROOT / "outputs"
ASSET_DIR = ROOT / "figures"
OUTPUT_DOCX = OUTPUT_DIR / "Environmental_Psychology_Manuscript_Through_Results.docx"

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
ASSET_DIR.mkdir(parents=True, exist_ok=True)

for required_file in (SOURCE_WORKBOOK, T0_FILE):
    if not required_file.exists():
        raise FileNotFoundError(
            f"Required restricted input is missing: {required_file}. "
            "See data/private/README.md for the expected local-only files."
        )


# -----------------------------------------------------------------------------
# Data loading
# -----------------------------------------------------------------------------


def worksheet_frame(ws, header_row: int, start_row: int, end_row: int | None = None) -> pd.DataFrame:
    headers = [cell.value for cell in ws[header_row]]
    last = end_row or ws.max_row
    rows = []
    for row in ws.iter_rows(min_row=start_row, max_row=last, values_only=True):
        if not any(value is not None for value in row):
            continue
        rows.append(list(row[: len(headers)]))
    return pd.DataFrame(rows, columns=headers)


wb = openpyxl.load_workbook(SOURCE_WORKBOOK, data_only=True)

ws_long = wb["Longitudinal"]
flow = []
for row in range(7, 10):
    flow.append(
        {
            "stage": ws_long.cell(row, 1).value,
            "raw": int(ws_long.cell(row, 2).value),
            "unique": int(ws_long.cell(row, 3).value),
            "duplicates": int(ws_long.cell(row, 4).value),
            "alpha_all": float(ws_long.cell(row, 8).value),
            "matched_n": int(ws_long.cell(row, 9).value),
            "alpha_matched": float(ws_long.cell(row, 10).value),
        }
    )

longitudinal = []
for row in range(13, 22):
    longitudinal.append(
        {
            "measure": ws_long.cell(row, 1).value,
            "t0": float(ws_long.cell(row, 2).value),
            "sd0": float(ws_long.cell(row, 3).value),
            "t1": float(ws_long.cell(row, 4).value),
            "sd1": float(ws_long.cell(row, 5).value),
            "t2": float(ws_long.cell(row, 6).value),
            "sd2": float(ws_long.cell(row, 7).value),
            "delta02": float(ws_long.cell(row, 8).value),
            "q": float(ws_long.cell(row, 9).value),
            "p": float(ws_long.cell(row, 10).value),
            "p01": float(ws_long.cell(row, 11).value),
            "r01": float(ws_long.cell(row, 12).value),
            "p12": float(ws_long.cell(row, 13).value),
            "r12": float(ws_long.cell(row, 14).value),
            "p02": float(ws_long.cell(row, 15).value),
            "r02": float(ws_long.cell(row, 16).value),
        }
    )

attrition = {
    "u": float(ws_long.cell(26, 13).value),
    "p": float(ws_long.cell(26, 14).value),
    "matched_mean": float(ws_long.cell(26, 12).value),
    "not_complete_n": int(ws_long.cell(27, 11).value),
    "not_complete_mean": float(ws_long.cell(27, 12).value),
    "earliest_p02": float(ws_long.cell(27, 8).value),
}

ws_assignment = wb["Assignment"]
assignment_rows = []
assignment_headers = [ws_assignment.cell(11, col).value for col in range(1, 10)]
for row in range(12, 56):
    values = [ws_assignment.cell(row, col).value for col in range(1, 10)]
    assignment_rows.append(dict(zip(assignment_headers, values)))
assignments = pd.DataFrame(assignment_rows)

assignment_dimensions = []
for row in range(7, 12):
    assignment_dimensions.append(
        {
            "dimension": ws_assignment.cell(row, 11).value,
            "mean": float(ws_assignment.cell(row, 12).value),
            "median": float(ws_assignment.cell(row, 13).value),
            "sd": float(ws_assignment.cell(row, 14).value),
        }
    )

assignment_associations = {
    "t2_n": int(ws_assignment.cell(33, 12).value),
    "t2_rho": float(ws_assignment.cell(33, 13).value),
    "t2_p": float(ws_assignment.cell(33, 14).value),
    "change_n": int(ws_assignment.cell(34, 12).value),
    "change_rho": float(ws_assignment.cell(34, 13).value),
    "change_p": float(ws_assignment.cell(34, 14).value),
}

ws_data = wb["Deidentified Data"]
data_headers = [cell.value for cell in ws_data[5]]
data_rows = []
for row in ws_data.iter_rows(min_row=6, values_only=True):
    if row[0] is None:
        continue
    data_rows.append(list(row[: len(data_headers)]))
data = pd.DataFrame(data_rows, columns=data_headers)
matched = data[(data["T0 present"] == True) & (data["T1 present"] == True) & (data["T2 present"] == True)].copy()  # noqa: E712

ws_qual = wb["Qualitative"]
common_themes = []
for row in range(7, 13):
    common_themes.append(
        {
            "theme": ws_qual.cell(row, 1).value,
            "T0": float(ws_qual.cell(row, 2).value),
            "T1": float(ws_qual.cell(row, 3).value),
            "T2": float(ws_qual.cell(row, 4).value),
        }
    )

lecture_themes = []
for row in range(18, 24):
    lecture_themes.append(
        {
            "theme": ws_qual.cell(row, 1).value,
            "count": int(ws_qual.cell(row, 2).value),
            "n": int(ws_qual.cell(row, 3).value),
            "percent": float(ws_qual.cell(row, 4).value),
        }
    )

contribution_counts = []
for row in range(29, 33):
    contribution_counts.append(
        {
            "category": ws_qual.cell(row, 1).value,
            "count": int(ws_qual.cell(row, 2).value),
            "percent": float(ws_qual.cell(row, 3).value),
        }
    )

changed_counts = []
for row in range(29, 32):
    changed_counts.append(
        {
            "category": ws_qual.cell(row, 5).value,
            "count": int(ws_qual.cell(row, 6).value),
            "percent": float(ws_qual.cell(row, 7).value),
        }
    )


# -----------------------------------------------------------------------------
# Studio progression linkage and exploratory statistics
# -----------------------------------------------------------------------------


ALIAS_FILE = PRIVATE_DIR / "name_aliases.json"
ALIASES = json.loads(ALIAS_FILE.read_text(encoding="utf-8")) if ALIAS_FILE.exists() else {}

HASH_SALT = os.environ.get("STUDY_LINKAGE_SALT")
if not HASH_SALT:
    salt_file = PRIVATE_DIR / "hash_salt.txt"
    if salt_file.exists():
        HASH_SALT = salt_file.read_text(encoding="utf-8").strip()
if not HASH_SALT:
    raise RuntimeError(
        "Set STUDY_LINKAGE_SALT or create data/private/hash_salt.txt. "
        "Never commit the salt to a public repository."
    )


def norm_name(value: object) -> str:
    text = str(value or "").strip().lower()
    text = text.translate(
        str.maketrans(
            {
                "ı": "i",
                "İ": "i",
                "ş": "s",
                "Ş": "s",
                "ğ": "g",
                "Ğ": "g",
                "ü": "u",
                "Ü": "u",
                "ö": "o",
                "Ö": "o",
                "ç": "c",
                "Ç": "c",
            }
        )
    )
    text = "".join(char for char in unicodedata.normalize("NFKD", text) if not unicodedata.combining(char))
    return re.sub(r"[^a-z0-9]", "", text)


def canonical(value: object) -> str:
    key = norm_name(value)
    return ALIASES.get(key, key)


def pid_for(key: str) -> str:
    return "P-" + hashlib.sha256((HASH_SALT + "|" + key).encode("utf-8")).hexdigest()[:6].upper()


t0_raw = pd.read_excel(T0_FILE)
t0_raw = t0_raw.copy()
t0_raw["_key"] = t0_raw.iloc[:, 1].map(canonical)
t0_raw["_pid"] = t0_raw["_key"].map(pid_for)
t0_raw["_timestamp"] = pd.to_datetime(t0_raw.iloc[:, 0])
t0_latest = t0_raw.sort_values("_timestamp", ascending=False).drop_duplicates("_key", keep="first")
studio_map = dict(zip(t0_latest["_pid"], t0_latest.iloc[:, 2].astype(str).str.strip()))
matched["Prior studio"] = matched["Participant ID"].map(studio_map)

STUDIO_ORDER = {"IAED 1001": 1, "IAED 1002": 2, "IAED 2001": 3, "IAED 2002": 4, "IAED 3001": 5}
matched["Studio rank"] = matched["Prior studio"].map(STUDIO_ORDER)

dimension_cols = [
    "Human–space relationship",
    "Psychological needs",
    "Multisensory / arousal / adaptation",
    "Affordances",
    "Spatial Gestalt",
]

matched_assign = matched[matched["Assignment ID"].notna()].copy()
matched_assign["High detail"] = (matched_assign["Assignment total /15"] >= 13) & (matched_assign[dimension_cols].min(axis=1) >= 2)
matched_assign["T1 to T2 change"] = matched_assign["T2 mean"] - matched_assign["T1 mean"]


def rankdata(values: np.ndarray) -> np.ndarray:
    values = np.asarray(values, dtype=float)
    order = np.argsort(values, kind="mergesort")
    ranks = np.empty(len(values), dtype=float)
    index = 0
    while index < len(values):
        end = index + 1
        while end < len(values) and values[order[end]] == values[order[index]]:
            end += 1
        ranks[order[index:end]] = (index + 1 + end) / 2.0
        index = end
    return ranks


def betacf(a: float, b: float, x: float) -> float:
    qab, qap, qam = a + b, a + 1.0, a - 1.0
    c = 1.0
    d = 1.0 - qab * x / qap
    if abs(d) < 3e-14:
        d = 3e-14
    d = 1.0 / d
    h = d
    for m in range(1, 201):
        m2 = 2 * m
        aa = m * (b - m) * x / ((qam + m2) * (a + m2))
        d = 1.0 + aa * d
        if abs(d) < 3e-14:
            d = 3e-14
        c = 1.0 + aa / c
        if abs(c) < 3e-14:
            c = 3e-14
        d = 1.0 / d
        h *= d * c
        aa = -(a + m) * (qab + m) * x / ((a + m2) * (qap + m2))
        d = 1.0 + aa * d
        if abs(d) < 3e-14:
            d = 3e-14
        c = 1.0 + aa / c
        if abs(c) < 3e-14:
            c = 3e-14
        d = 1.0 / d
        delta = d * c
        h *= delta
        if abs(delta - 1.0) < 3e-12:
            break
    return h


def betainc_reg(a: float, b: float, x: float) -> float:
    if x <= 0:
        return 0.0
    if x >= 1:
        return 1.0
    bt = math.exp(math.lgamma(a + b) - math.lgamma(a) - math.lgamma(b) + a * math.log(x) + b * math.log1p(-x))
    if x < (a + 1.0) / (a + b + 2.0):
        return bt * betacf(a, b, x) / a
    return 1.0 - bt * betacf(b, a, 1.0 - x) / b


def spearman(values_x, values_y) -> tuple[float, float]:
    x = np.asarray(values_x, dtype=float)
    y = np.asarray(values_y, dtype=float)
    mask = np.isfinite(x) & np.isfinite(y)
    x = x[mask]
    y = y[mask]
    rx = rankdata(x)
    ry = rankdata(y)
    rho = float(np.corrcoef(rx, ry)[0, 1])
    n = len(x)
    if n <= 2 or abs(rho) >= 1:
        return rho, 0.0 if abs(rho) >= 1 else math.nan
    t_value = abs(rho) * math.sqrt((n - 2) / (1 - rho * rho))
    p_value = betainc_reg((n - 2) / 2.0, 0.5, (n - 2) / ((n - 2) + t_value * t_value))
    return rho, float(p_value)


def normal_sf(z: float) -> float:
    return 0.5 * math.erfc(z / math.sqrt(2.0))


def wilcoxon_p(before, after) -> tuple[float, float]:
    differences = np.asarray(after, dtype=float) - np.asarray(before, dtype=float)
    differences = differences[differences != 0]
    if not len(differences):
        return 1.0, 0.0
    ranks = rankdata(np.abs(differences))
    positive = float(ranks[differences > 0].sum())
    negative = float(ranks[differences < 0].sum())
    n = len(differences)
    _, counts = np.unique(np.abs(differences), return_counts=True)
    tie = float(sum(count**3 - count for count in counts))
    mean = n * (n + 1) / 4
    variance = (n * (n + 1) * (2 * n + 1) - 0.5 * tie) / 24
    z_value = max(0.0, abs(min(positive, negative) - mean) - 0.5) / math.sqrt(variance)
    p_value = min(1.0, 2 * normal_sf(z_value))
    rank_biserial = (positive - negative) / (positive + negative)
    return p_value, float(rank_biserial)


def mann_whitney(values_x, values_y) -> tuple[float, float, float]:
    x = np.asarray(values_x, dtype=float)
    y = np.asarray(values_y, dtype=float)
    all_values = np.concatenate([x, y])
    ranks = rankdata(all_values)
    nx, ny, total_n = len(x), len(y), len(all_values)
    u1 = float(ranks[:nx].sum() - nx * (nx + 1) / 2)
    mean = nx * ny / 2
    _, counts = np.unique(all_values, return_counts=True)
    tie = float(sum(count**3 - count for count in counts))
    variance = nx * ny / 12 * ((total_n + 1) - tie / (total_n * (total_n - 1)))
    z_value = max(0.0, abs(u1 - mean) - 0.5) / math.sqrt(variance)
    p_value = min(1.0, 2 * normal_sf(z_value))
    comparisons = [np.sign(a - b) for a in x for b in y]
    cliff_delta = float(np.mean(comparisons))
    return u1, p_value, cliff_delta


def holm(values: list[float]) -> list[float]:
    order = np.argsort(values)
    adjusted = np.empty(len(values), dtype=float)
    running = 0.0
    total = len(values)
    for rank, index in enumerate(order):
        running = max(running, min(1.0, (total - rank) * values[index]))
        adjusted[index] = running
    return adjusted.tolist()


def fisher_exact_two_sided(table: list[list[int]]) -> float:
    a, b = table[0]
    c, d = table[1]
    row1 = a + b
    row2 = c + d
    col1 = a + c
    total = row1 + row2

    def probability(value_a: int) -> float:
        return math.comb(row1, value_a) * math.comb(row2, col1 - value_a) / math.comb(total, col1)

    observed = probability(a)
    lower = max(0, col1 - row2)
    upper = min(row1, col1)
    return float(sum(probability(candidate) for candidate in range(lower, upper + 1) if probability(candidate) <= observed + 1e-12))


studio_t0_rho, studio_t0_p = spearman(matched["Studio rank"], matched["T0 mean"])
studio_total_rho, studio_total_p = spearman(matched_assign["Studio rank"], matched_assign["Assignment total /15"])
studio_dim_correlations = []
raw_studio_ps = []
for column in dimension_cols:
    rho, p_value = spearman(matched_assign["Studio rank"], matched_assign[column])
    studio_dim_correlations.append({"dimension": column, "rho": rho, "p": p_value})
    raw_studio_ps.append(p_value)
all_studio_adjusted = holm([studio_total_p] + raw_studio_ps)
studio_total_p_holm = all_studio_adjusted[0]
for record, adjusted in zip(studio_dim_correlations, all_studio_adjusted[1:]):
    record["p_holm"] = adjusted

high_group = matched_assign[matched_assign["High detail"]].copy()
other_scored_group = matched_assign[~matched_assign["High detail"]].copy()
high_wilcoxon_p, high_rank_biserial = wilcoxon_p(high_group["T1 mean"], high_group["T2 mean"])
other_wilcoxon_p, other_rank_biserial = wilcoxon_p(other_scored_group["T1 mean"], other_scored_group["T2 mean"])
subgroup_u, subgroup_p, subgroup_cliff = mann_whitney(high_group["T1 to T2 change"], other_scored_group["T1 to T2 change"])

early = matched_assign[matched_assign["Studio rank"] <= 3]
later = matched_assign[matched_assign["Studio rank"] >= 4]
early_high = int(early["High detail"].sum())
later_high = int(later["High detail"].sum())
studio_high_fisher_p = fisher_exact_two_sided(
    [[early_high, len(early) - early_high], [later_high, len(later) - later_high]]
)


# -----------------------------------------------------------------------------
# Figure drawing with Pillow
# -----------------------------------------------------------------------------


NAVY = "#17324D"
TEAL = "#0F766E"
AMBER = "#D99A2B"
CORAL = "#C65B5B"
INK = "#17212B"
MID = "#66727D"
GRID = "#D9E2E8"
PALE = "#F4F7F9"
WHITE = "#FFFFFF"

FONT_REG = r"C:\Windows\Fonts\arial.ttf"
FONT_BOLD = r"C:\Windows\Fonts\arialbd.ttf"
FONT_ITALIC = r"C:\Windows\Fonts\ariali.ttf"


def font(size: int, bold: bool = False, italic: bool = False):
    path = FONT_BOLD if bold else FONT_ITALIC if italic else FONT_REG
    return ImageFont.truetype(path, size=size)


def hex_rgb(value: str) -> tuple[int, int, int]:
    value = value.lstrip("#")
    return tuple(int(value[index : index + 2], 16) for index in (0, 2, 4))


def draw_centered(draw: ImageDraw.ImageDraw, xy: tuple[float, float], text: str, text_font, fill=INK):
    box = draw.textbbox((0, 0), text, font=text_font)
    width = box[2] - box[0]
    height = box[3] - box[1]
    draw.text((xy[0] - width / 2, xy[1] - height / 2), text, font=text_font, fill=fill)


def draw_right(draw: ImageDraw.ImageDraw, xy: tuple[float, float], text: str, text_font, fill=INK):
    box = draw.textbbox((0, 0), text, font=text_font)
    draw.text((xy[0] - (box[2] - box[0]), xy[1]), text, font=text_font, fill=fill)


def wrapped_lines(draw: ImageDraw.ImageDraw, text: str, text_font, max_width: int) -> list[str]:
    words = text.split()
    lines: list[str] = []
    current = ""
    for word in words:
        candidate = word if not current else current + " " + word
        if draw.textbbox((0, 0), candidate, font=text_font)[2] <= max_width:
            current = candidate
        else:
            if current:
                lines.append(current)
            current = word
    if current:
        lines.append(current)
    return lines


def bootstrap_ci(values, samples=4000, seed=20261028):
    array = np.asarray(values, dtype=float)
    rng = np.random.default_rng(seed)
    indices = rng.integers(0, len(array), size=(samples, len(array)))
    means = array[indices].mean(axis=1)
    return float(np.quantile(means, 0.025)), float(np.quantile(means, 0.975))


def save_image(image: Image.Image, path: Path):
    image.save(path, format="PNG", optimize=True, dpi=(300, 300))


def figure_composite(path: Path):
    width, height = 2600, 1500
    image = Image.new("RGB", (width, height), WHITE)
    draw = ImageDraw.Draw(image, "RGBA")
    draw.text((150, 80), "Composite person-space self-report across three stages", font=font(60, bold=True), fill=INK)
    draw.text((150, 155), "Matched sample  n = 34  |  aggregate means and 95% confidence intervals", font=font(34), fill=MID)

    left, right, top, bottom = 320, 2300, 300, 1240
    stages = ["T0\nBefore lecture", "T1\nAfter Lecture 1", "T2\nAfter assignment"]
    x_positions = np.linspace(left + 170, right - 170, 3)

    def y_map(value):
        return bottom - (value - 1) / 4 * (bottom - top)

    for tick in range(1, 6):
        y = y_map(tick)
        draw.line((left, y, right, y), fill=hex_rgb(GRID) + (190,), width=2)
        draw_right(draw, (left - 35, y - 20), str(tick), font(30), MID)
    draw.text((65, 640), "Mean score", font=font(34, bold=True), fill=INK)

    stage_columns = ["T0 mean", "T1 mean", "T2 mean"]
    stage_colors = [NAVY, TEAL, AMBER]
    means = [float(matched[column].mean()) for column in stage_columns]
    mean_points = [(x_positions[index], y_map(means[index])) for index in range(3)]
    draw.line(mean_points, fill=hex_rgb(INK) + (255,), width=12)
    for index, column in enumerate(stage_columns):
        low, high = bootstrap_ci(matched[column], seed=20261028 + index)
        x = x_positions[index]
        draw.line((x, y_map(low), x, y_map(high)), fill=hex_rgb(stage_colors[index]) + (255,), width=8)
        draw.line((x - 24, y_map(low), x + 24, y_map(low)), fill=hex_rgb(stage_colors[index]) + (255,), width=8)
        draw.line((x - 24, y_map(high), x + 24, y_map(high)), fill=hex_rgb(stage_colors[index]) + (255,), width=8)
        y = y_map(means[index])
        draw.ellipse((x - 25, y - 25, x + 25, y + 25), fill=stage_colors[index], outline=WHITE, width=5)
        draw_centered(draw, (x, y - 70), f"{means[index]:.2f}", font(36, bold=True), stage_colors[index])
        label_lines = stages[index].split("\n")
        for line_index, line in enumerate(label_lines):
            draw_centered(draw, (x, bottom + 80 + line_index * 40), line, font(31, bold=(line_index == 0)), INK)

    draw.text(
        (320, 1405),
        "Points show means; bars show 95% bootstrap confidence intervals. Individual trajectories are not displayed.",
        font=font(28, italic=True),
        fill=MID,
    )
    save_image(image, path)


def diverging_color(value: float, limit=0.75) -> tuple[int, int, int]:
    value = max(-limit, min(limit, value)) / limit
    white = np.array([255, 255, 255], dtype=float)
    target = np.array(hex_rgb(TEAL if value >= 0 else CORAL), dtype=float)
    mixed = white * (1 - abs(value)) + target * abs(value)
    return tuple(int(channel) for channel in mixed)


def figure_items(path: Path):
    width, height = 2700, 1800
    image = Image.new("RGB", (width, height), WHITE)
    draw = ImageDraw.Draw(image, "RGBA")
    draw.text((130, 70), "Concept-specific profiles and stage-to-stage change", font=font(58, bold=True), fill=INK)
    draw.text((130, 145), "Eight matched self-report items  n = 34", font=font(34), fill=MID)

    items = longitudinal[:-1]
    left_label = 130
    plot_left, plot_right = 710, 1660
    heat_left, heat_right = 1850, 2560
    top, row_gap = 310, 160
    x_min, x_max = 3.4, 4.5

    def x_map(value):
        return plot_left + (value - x_min) / (x_max - x_min) * (plot_right - plot_left)

    for tick in np.arange(3.4, 4.51, 0.2):
        x = x_map(float(tick))
        draw.line((x, top - 40, x, top + row_gap * len(items) - 50), fill=hex_rgb(GRID) + (160,), width=2)
        draw_centered(draw, (x, top - 85), f"{tick:.1f}", font(26), MID)
    draw_centered(draw, ((plot_left + plot_right) / 2, top - 145), "Mean score", font(31, bold=True), INK)

    changes = ["T0 to T1", "T1 to T2", "T0 to T2"]
    heat_width = (heat_right - heat_left) / 3
    for index, label in enumerate(changes):
        draw_centered(draw, (heat_left + heat_width * (index + 0.5), top - 120), label, font(27, bold=True), INK)
    draw_centered(draw, ((heat_left + heat_right) / 2, top - 175), "Mean change", font(31, bold=True), INK)

    colors = [NAVY, TEAL, AMBER]
    for row_index, record in enumerate(items):
        y = top + row_index * row_gap
        label_lines = wrapped_lines(draw, str(record["measure"]), font(29, bold=True), 500)
        label_y = y - 18 * (len(label_lines) - 1)
        for line_index, line in enumerate(label_lines):
            draw.text((left_label, label_y + line_index * 36), line, font=font(29, bold=True), fill=INK)
        draw.line((plot_left, y, plot_right, y), fill=hex_rgb(GRID) + (125,), width=2)
        values = [record["t0"], record["t1"], record["t2"]]
        point_positions = [(x_map(value), y) for value in values]
        draw.line(point_positions, fill=hex_rgb(MID) + (135,), width=4)
        for stage_index, (x, point_y) in enumerate(point_positions):
            draw.ellipse((x - 17, point_y - 17, x + 17, point_y + 17), fill=colors[stage_index], outline=WHITE, width=4)

        delta_values = [record["t1"] - record["t0"], record["t2"] - record["t1"], record["t2"] - record["t0"]]
        for column_index, delta in enumerate(delta_values):
            x0 = heat_left + heat_width * column_index + 8
            x1 = heat_left + heat_width * (column_index + 1) - 8
            y0, y1 = y - 55, y + 55
            draw.rounded_rectangle((x0, y0, x1, y1), radius=14, fill=diverging_color(delta), outline=GRID, width=2)
            draw_centered(draw, ((x0 + x1) / 2, y), f"{delta:+.2f}", font(28, bold=True), INK)

    legend_y = 1645
    for index, (label, color_value) in enumerate(zip(["T0", "T1", "T2"], colors)):
        x = 720 + index * 200
        draw.ellipse((x, legend_y, x + 28, legend_y + 28), fill=color_value)
        draw.text((x + 42, legend_y - 5), label, font=font(28, bold=True), fill=INK)
    draw.text(
        (130, 1730),
        "The score axis is limited to 3.4-4.5 to resolve item-level differences. The response scale ranged from 1 to 5.",
        font=font(27, italic=True),
        fill=MID,
    )
    save_image(image, path)


def figure_assignment(path: Path):
    width, height = 2700, 1650
    image = Image.new("RGB", (width, height), WHITE)
    draw = ImageDraw.Draw(image, "RGBA")
    draw.text((120, 65), "Assignment evidence across five environmental psychology dimensions", font=font(56, bold=True), fill=INK)
    draw.text((120, 140), "Researcher-scored submissions  n = 44", font=font(34), fill=MID)

    draw.text((120, 250), "Distribution of rubric scores", font=font(36, bold=True), fill=INK)
    bar_left, bar_right = 700, 1640
    top, row_gap = 390, 190
    score_colors = ["#E7ECF0", CORAL, AMBER, TEAL]
    short_names = [
        "Human-space relationship",
        "Psychological needs",
        "Multisensory arousal adaptation",
        "Affordances",
        "Spatial Gestalt",
    ]
    for row_index, (column, short_name) in enumerate(zip(dimension_cols, short_names)):
        y = top + row_index * row_gap
        label_lines = wrapped_lines(draw, short_name, font(28, bold=True), 500)
        for line_index, line in enumerate(label_lines):
            draw.text((120, y - 28 + line_index * 34), line, font=font(28, bold=True), fill=INK)
        counts = Counter(int(value) for value in assignments[column])
        x = bar_left
        for score in range(4):
            proportion = counts.get(score, 0) / len(assignments)
            width_segment = proportion * (bar_right - bar_left)
            draw.rectangle((x, y - 42, x + width_segment, y + 42), fill=score_colors[score])
            if proportion >= 0.09:
                draw_centered(draw, (x + width_segment / 2, y), f"{100 * proportion:.0f}%", font(26, bold=True), WHITE if score in (1, 3) else INK)
            x += width_segment
        draw.rounded_rectangle((bar_left, y - 42, bar_right, y + 42), radius=7, outline=GRID, width=3)

    legend_y = 1370
    for score in range(4):
        x = 180 + score * 270
        draw.rectangle((x, legend_y, x + 34, legend_y + 34), fill=score_colors[score], outline=GRID)
        draw.text((x + 50, legend_y - 3), f"Score {score}", font=font(27), fill=INK)

    panel_left, panel_right = 1810, 2570
    panel_top, panel_bottom = 365, 1300
    draw.text((panel_left, 250), "Total score distribution", font=font(36, bold=True), fill=INK)
    totals = assignments["Total /15"].astype(int)
    frequencies = Counter(totals)
    max_count = max(frequencies.values())

    def y_map(count):
        return panel_bottom - count / max_count * (panel_bottom - panel_top)

    for count in range(0, max_count + 1, 3):
        y = y_map(count)
        draw.line((panel_left, y, panel_right, y), fill=hex_rgb(GRID) + (170,), width=2)
        draw_right(draw, (panel_left - 25, y - 15), str(count), font(25), MID)
    total_values = list(range(5, 16))
    x_positions = np.linspace(panel_left + 35, panel_right - 35, len(total_values))
    for total, x in zip(total_values, x_positions):
        count = frequencies.get(total, 0)
        y = y_map(count)
        draw.line((x, panel_bottom, x, y), fill=hex_rgb(NAVY) + (220,), width=14)
        draw.ellipse((x - 15, y - 15, x + 15, y + 15), fill=NAVY)
        if count:
            draw_centered(draw, (x, y - 45), str(count), font(24, bold=True), NAVY)
        draw_centered(draw, (x, panel_bottom + 45), str(total), font(25), INK)
    median_total = float(np.median(totals))
    median_x = np.interp(median_total, [5, 15], [x_positions[0], x_positions[-1]])
    draw.line((median_x, panel_top - 10, median_x, panel_bottom), fill=hex_rgb(AMBER) + (255,), width=6)
    draw.text((median_x + 15, panel_top + 5), f"Median {median_total:.0f}", font=font(26, bold=True), fill=AMBER)
    draw_centered(draw, ((panel_left + panel_right) / 2, panel_bottom + 105), "Assignment total  0-15", font(29, bold=True), INK)

    draw.text((120, 1550), "Each dimension was scored from 0 to 3; the five dimensions produced a total from 0 to 15.", font=font(27, italic=True), fill=MID)
    save_image(image, path)


def deterministic_jitter(label: str, width: float) -> float:
    digest = hashlib.sha256(label.encode("utf-8")).digest()
    number = int.from_bytes(digest[:4], "big") / (2**32 - 1)
    return (number - 0.5) * width


def figure_studio_and_subgroup(path: Path):
    width, height = 2750, 1600
    image = Image.new("RGB", (width, height), WHITE)
    draw = ImageDraw.Draw(image, "RGBA")
    draw.text((120, 65), "Exploratory associations with studio progression and assignment detail", font=font(56, bold=True), fill=INK)
    draw.text((120, 140), "Aggregate estimates for matched cases; no participant-level points are displayed", font=font(34), fill=MID)

    # Left panel: aggregate rank association only.
    left, right = 180, 1320
    draw.text((left, 255), "Studio progression and assignment score", font=font(35, bold=True), fill=INK)
    draw.text((left, 320), f"Spearman association  n = {len(matched_assign)}", font=font(29), fill=MID)
    axis_y = 760
    axis_left, axis_right = left + 70, right - 70
    draw.line((axis_left, axis_y, axis_right, axis_y), fill=GRID, width=7)
    for tick in [-1.0, -0.5, 0.0, 0.5, 1.0]:
        x = axis_left + (tick + 1) / 2 * (axis_right - axis_left)
        draw.line((x, axis_y - 18, x, axis_y + 18), fill=INK if tick == 0 else MID, width=4)
        draw_centered(draw, (x, axis_y + 58), f"{tick:+.1f}", font(26), MID)
    estimate_x = axis_left + (studio_total_rho + 1) / 2 * (axis_right - axis_left)
    draw.ellipse((estimate_x - 30, axis_y - 30, estimate_x + 30, axis_y + 30), fill=TEAL, outline=WHITE, width=5)
    draw_centered(draw, (estimate_x, axis_y - 70), f"rho = {studio_total_rho:.2f}", font(31, bold=True), TEAL)
    draw_centered(draw, ((axis_left + axis_right) / 2, 930), f"Holm-adjusted p = {studio_total_p_holm:.3f}", font(30, bold=True), NAVY)
    draw_centered(draw, ((axis_left + axis_right) / 2, 1010), "Higher studio progression was associated with higher assignment scores.", font(27), INK)

    # Right panel: aggregate group means only.
    panel_left, panel_right = 1510, 2580
    draw.text((panel_left, 255), "T1 to T2 change by assignment detail", font=font(35, bold=True), fill=INK)
    draw.text((panel_left, 320), f"Matched assignments  n = {len(high_group) + len(other_scored_group)}", font=font(29), fill=MID)
    plot_left, plot_right = panel_left + 110, panel_right - 90
    x_min, x_max = -0.5, 0.1
    for tick in [-0.5, -0.3, -0.1, 0.1]:
        x = plot_left + (tick - x_min) / (x_max - x_min) * (plot_right - plot_left)
        draw.line((x, 500, x, 1030), fill=GRID, width=2)
        draw_centered(draw, (x, 1085), f"{tick:+.1f}", font(26), MID)
    zero_x = plot_left + (0 - x_min) / (x_max - x_min) * (plot_right - plot_left)
    draw.line((zero_x, 500, zero_x, 1030), fill=MID, width=4)
    group_specs = [(high_group, "High-detail", TEAL), (other_scored_group, "Other scored", AMBER)]
    for y, (group, label, color_value) in zip([650, 880], group_specs):
        mean_change = float(group["T1 to T2 change"].mean())
        draw.text((panel_left, y - 25), f"{label}  n = {len(group)}", font=font(29, bold=True), fill=INK)
        x = plot_left + (mean_change - x_min) / (x_max - x_min) * (plot_right - plot_left)
        draw.polygon([(x, y - 33), (x + 33, y), (x, y + 33), (x - 33, y)], fill=color_value)
        draw.text((x + 48, y - 20), f"mean {mean_change:+.2f}", font=font(28, bold=True), fill=color_value)
    draw_centered(draw, ((plot_left + plot_right) / 2, 1160), "Mean T2 minus T1 change", font(28, bold=True), INK)
    draw_centered(draw, ((plot_left + plot_right) / 2, 1250), f"Mann-Whitney p = {subgroup_p:.3f}  |  Cliff delta = {subgroup_cliff:+.2f}", font(29, bold=True), NAVY)
    draw.text((120, 1505), "The displayed values are aggregate summaries and do not permit reconstruction of individual response patterns.", font=font(27, italic=True), fill=MID)
    save_image(image, path)


def figure_themes(path: Path):
    width, height = 2700, 1700
    image = Image.new("RGB", (width, height), WHITE)
    draw = ImageDraw.Draw(image, "RGBA")
    draw.text((120, 70), "Themes in the common open-ended question", font=font(58, bold=True), fill=INK)
    draw.text((120, 145), "Keyword-assisted multi-label indicators at T0, T1, and T2", font=font(34), fill=MID)

    label_left, plot_left, plot_right = 120, 820, 2500
    top, row_gap = 340, 190
    for tick in range(0, 101, 20):
        x = plot_left + tick / 100 * (plot_right - plot_left)
        draw.line((x, top - 70, x, top + row_gap * len(common_themes) - 70), fill=hex_rgb(GRID) + (170,), width=2)
        draw_centered(draw, (x, top - 120), f"{tick}%", font(27), MID)
    stage_colors = [NAVY, TEAL, AMBER]
    for row_index, record in enumerate(common_themes):
        y = top + row_index * row_gap
        label_lines = wrapped_lines(draw, record["theme"], font(29, bold=True), 620)
        for line_index, line in enumerate(label_lines):
            draw.text((label_left, y - 25 + line_index * 34), line, font=font(29, bold=True), fill=INK)
        values = [record["T0"] * 100, record["T1"] * 100, record["T2"] * 100]
        points = [(plot_left + value / 100 * (plot_right - plot_left), y) for value in values]
        draw.line(points, fill=hex_rgb(MID) + (145,), width=5)
        for stage_index, (x, point_y) in enumerate(points):
            draw.ellipse((x - 18, point_y - 18, x + 18, point_y + 18), fill=stage_colors[stage_index], outline=WHITE, width=4)
            offset_y = [-43, 37, 0][stage_index]
            draw.text((x + 20, point_y + offset_y - 14), f"{values[stage_index]:.0f}%", font=font(24, bold=True), fill=stage_colors[stage_index])
    legend_y = 1540
    for index, (label, color_value) in enumerate(zip(["T0  n=60", "T1  n=53", "T2  n=40"], stage_colors)):
        x = 850 + index * 430
        draw.ellipse((x, legend_y, x + 30, legend_y + 30), fill=color_value)
        draw.text((x + 45, legend_y - 3), label, font=font(28, bold=True), fill=INK)
    draw.text((120, 1640), "Responses could receive more than one indicator; percentages therefore do not sum to 100%.", font=font(27, italic=True), fill=MID)
    save_image(image, path)


FIGURE_PATHS = {
    "composite": ASSET_DIR / "figure_1_composite_trajectory.png",
    "items": ASSET_DIR / "figure_2_item_profiles.png",
    "assignment": ASSET_DIR / "figure_3_assignment_evidence.png",
    "studio": ASSET_DIR / "figure_4_studio_and_subgroup.png",
    "themes": ASSET_DIR / "figure_5_open_ended_themes.png",
}

figure_composite(FIGURE_PATHS["composite"])
figure_items(FIGURE_PATHS["items"])
figure_assignment(FIGURE_PATHS["assignment"])
figure_studio_and_subgroup(FIGURE_PATHS["studio"])
figure_themes(FIGURE_PATHS["themes"])


# -----------------------------------------------------------------------------
# Word helpers
# -----------------------------------------------------------------------------


DOC_TEXT = RGBColor(23, 33, 43)
DOC_NAVY = RGBColor(23, 50, 77)
DOC_TEAL = RGBColor(15, 118, 110)
DOC_GRAY = RGBColor(102, 114, 125)
DOC_LIGHT = "F4F7F9"
DOC_PALE_BLUE = "EEF4F7"
DOC_BORDER = "D9D9D9"


def set_run_font(run, size=10.5, bold=False, italic=False, color=DOC_TEXT, name="Arial"):
    run.font.name = name
    run._element.get_or_add_rPr().rFonts.set(qn("w:ascii"), name)
    run._element.get_or_add_rPr().rFonts.set(qn("w:hAnsi"), name)
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color


def set_cell_shading(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shading = tc_pr.find(qn("w:shd"))
    if shading is None:
        shading = OxmlElement("w:shd")
        tc_pr.append(shading)
    shading.set(qn("w:fill"), fill)


def set_cell_border(cell, color=DOC_BORDER, size="4"):
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = tc_pr.first_child_found_in("w:tcBorders")
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        node = borders.find(qn(f"w:{edge}"))
        if node is None:
            node = OxmlElement(f"w:{edge}")
            borders.append(node)
        node.set(qn("w:val"), "single")
        node.set(qn("w:sz"), size)
        node.set(qn("w:space"), "0")
        node.set(qn("w:color"), color)


def set_cell_margins(cell, top=90, start=110, bottom=90, end=110):
    tc_pr = cell._tc.get_or_add_tcPr()
    margins = tc_pr.first_child_found_in("w:tcMar")
    if margins is None:
        margins = OxmlElement("w:tcMar")
        tc_pr.append(margins)
    for margin, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = margins.find(qn(f"w:{margin}"))
        if node is None:
            node = OxmlElement(f"w:{margin}")
            margins.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def set_cell_width(cell, width_in):
    tc_pr = cell._tc.get_or_add_tcPr()
    width = tc_pr.find(qn("w:tcW"))
    if width is None:
        width = OxmlElement("w:tcW")
        tc_pr.append(width)
    width.set(qn("w:w"), str(int(width_in * 1440)))
    width.set(qn("w:type"), "dxa")


def repeat_header(row):
    tr_pr = row._tr.get_or_add_trPr()
    header = OxmlElement("w:tblHeader")
    header.set(qn("w:val"), "true")
    tr_pr.append(header)


def prevent_row_split(row):
    tr_pr = row._tr.get_or_add_trPr()
    cant_split = OxmlElement("w:cantSplit")
    tr_pr.append(cant_split)


def style_table(table, widths, left_columns=None, font_size=8.4):
    left_columns = left_columns or {0}
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    for row_index, row in enumerate(table.rows):
        prevent_row_split(row)
        if row_index == 0:
            repeat_header(row)
        for column_index, cell in enumerate(row.cells):
            set_cell_width(cell, widths[column_index])
            set_cell_margins(cell)
            set_cell_border(cell)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            if row_index == 0:
                set_cell_shading(cell, "17324D")
            elif row_index % 2 == 0:
                set_cell_shading(cell, DOC_PALE_BLUE)
            else:
                set_cell_shading(cell, "FFFFFF")
            for paragraph in cell.paragraphs:
                paragraph.paragraph_format.space_before = Pt(0)
                paragraph.paragraph_format.space_after = Pt(0)
                paragraph.paragraph_format.line_spacing = 1.05
                paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT if column_index in left_columns else WD_ALIGN_PARAGRAPH.CENTER
                for run in paragraph.runs:
                    set_run_font(
                        run,
                        size=font_size,
                        bold=row_index == 0,
                        color=RGBColor(255, 255, 255) if row_index == 0 else DOC_TEXT,
                    )


def add_table_caption(doc, number: int, title: str):
    paragraph = doc.add_paragraph()
    paragraph.style = doc.styles["Caption"]
    paragraph.paragraph_format.keep_with_next = True
    paragraph.paragraph_format.space_before = Pt(8)
    paragraph.paragraph_format.space_after = Pt(4)
    run = paragraph.add_run(f"Table {number}  ")
    set_run_font(run, size=9.2, bold=True)
    run = paragraph.add_run(title)
    set_run_font(run, size=9.2, italic=True)


def add_table_note(doc, text: str):
    paragraph = doc.add_paragraph()
    paragraph.paragraph_format.space_before = Pt(3)
    paragraph.paragraph_format.space_after = Pt(7)
    paragraph.paragraph_format.line_spacing = 1.0
    run = paragraph.add_run("Note  ")
    set_run_font(run, size=8.2, bold=True, italic=True, color=DOC_GRAY)
    run = paragraph.add_run(text)
    set_run_font(run, size=8.2, italic=True, color=DOC_GRAY)


def set_image_alt(inline_shape, text: str):
    inline_shape._inline.docPr.set("descr", text)
    inline_shape._inline.docPr.set("title", text[:80])


def add_figure(doc, path: Path, number: int, caption: str, alt_text: str, width=6.65):
    paragraph = doc.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.space_before = Pt(7)
    paragraph.paragraph_format.space_after = Pt(2)
    paragraph.paragraph_format.keep_with_next = True
    shape = paragraph.add_run().add_picture(str(path), width=Inches(width))
    set_image_alt(shape, alt_text)
    caption_paragraph = doc.add_paragraph()
    caption_paragraph.style = doc.styles["Caption"]
    caption_paragraph.paragraph_format.space_before = Pt(2)
    caption_paragraph.paragraph_format.space_after = Pt(8)
    run = caption_paragraph.add_run(f"Figure {number}  ")
    set_run_font(run, size=9.2, bold=True)
    run = caption_paragraph.add_run(caption)
    set_run_font(run, size=9.2, italic=True)


def add_page_number(paragraph):
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = paragraph.add_run()
    set_run_font(run, size=8.5, color=DOC_GRAY)
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instruction = OxmlElement("w:instrText")
    instruction.set(qn("xml:space"), "preserve")
    instruction.text = " PAGE "
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.extend([begin, instruction, end])


def p_value(value: float) -> str:
    if value < 0.001:
        return "< .001"
    return f"= {value:.3f}".replace("0.", ".")


def number(value: float, digits=2) -> str:
    return f"{value:.{digits}f}"


def add_body_paragraph(doc, text: str):
    paragraph = doc.add_paragraph(text)
    paragraph.paragraph_format.keep_together = False
    return paragraph


def add_reference(doc, text: str):
    paragraph = doc.add_paragraph()
    paragraph.paragraph_format.left_indent = Inches(0.28)
    paragraph.paragraph_format.first_line_indent = Inches(-0.28)
    paragraph.paragraph_format.space_after = Pt(5)
    paragraph.paragraph_format.line_spacing = 1.05
    run = paragraph.add_run(text)
    set_run_font(run, size=9.4)


# -----------------------------------------------------------------------------
# Build the manuscript
# -----------------------------------------------------------------------------


doc = Document()
section = doc.sections[0]
section.page_width = Inches(8.5)
section.page_height = Inches(11)
section.top_margin = Inches(0.75)
section.bottom_margin = Inches(0.75)
section.left_margin = Inches(0.82)
section.right_margin = Inches(0.82)
section.header_distance = Inches(0.28)
section.footer_distance = Inches(0.3)

styles = doc.styles
normal = styles["Normal"]
normal.font.name = "Arial"
normal._element.rPr.rFonts.set(qn("w:ascii"), "Arial")
normal._element.rPr.rFonts.set(qn("w:hAnsi"), "Arial")
normal.font.size = Pt(10.4)
normal.font.color.rgb = DOC_TEXT
normal.paragraph_format.line_spacing = 1.12
normal.paragraph_format.space_after = Pt(6)

title_style = styles["Title"]
title_style.font.name = "Arial"
title_style._element.rPr.rFonts.set(qn("w:ascii"), "Arial")
title_style._element.rPr.rFonts.set(qn("w:hAnsi"), "Arial")
title_style.font.size = Pt(17)
title_style.font.bold = True
title_style.font.color.rgb = RGBColor(0, 0, 0)
title_style.paragraph_format.space_after = Pt(7)

for style_name, size, before, after in (
    ("Heading 1", 13.5, 14, 6),
    ("Heading 2", 11.5, 11, 4),
    ("Heading 3", 10.7, 9, 3),
):
    style = styles[style_name]
    style.font.name = "Arial"
    style._element.rPr.rFonts.set(qn("w:ascii"), "Arial")
    style._element.rPr.rFonts.set(qn("w:hAnsi"), "Arial")
    style.font.size = Pt(size)
    style.font.bold = True
    style.font.color.rgb = RGBColor(0, 0, 0)
    style.paragraph_format.space_before = Pt(before)
    style.paragraph_format.space_after = Pt(after)
    style.paragraph_format.keep_with_next = True

caption_style = styles["Caption"]
caption_style.font.name = "Arial"
caption_style._element.rPr.rFonts.set(qn("w:ascii"), "Arial")
caption_style._element.rPr.rFonts.set(qn("w:hAnsi"), "Arial")
caption_style.font.size = Pt(9.2)
caption_style.font.italic = True
caption_style.font.color.rgb = DOC_TEXT

header = section.header.paragraphs[0]
header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
run = header.add_run("Environmental Psychology in Interior Architecture Education")
set_run_font(run, size=8.3, color=DOC_GRAY)
add_page_number(section.footer.paragraphs[0])

doc.core_properties.title = "Learning Environmental Psychology: From Intuitive Human-Centred Understanding to Spatial Analysis"
doc.core_properties.subject = "Introduction Methods and Results"
doc.core_properties.author = ""
doc.core_properties.keywords = "environmental psychology, interior architecture, design studio, person-environment relationship"

title = doc.add_paragraph("Learning Environmental Psychology: From Intuitive Human-Centred Understanding to Spatial Analysis", style="Title")
title.alignment = WD_ALIGN_PARAGRAPH.LEFT
subtitle = doc.add_paragraph()
subtitle.paragraph_format.space_after = Pt(13)
run = subtitle.add_run("A Three Stage Study of Third Year Interior Architecture Students")
set_run_font(run, size=12, italic=True, color=DOC_GRAY)

doc.add_paragraph("Introduction", style="Heading 1")
add_body_paragraph(
    doc,
    "Interior architecture operates through continuous relations between people and environments. Spatial layout, material character, light, sound, temperature, scale, and circulation do not function as neutral background conditions. They participate in how occupants perceive a setting, regulate comfort and privacy, move, interact, and assign meaning to place. Environmental psychology provides a framework for examining these reciprocal relations without reducing either the person or the environment to an isolated variable. Its transactional tradition treats people and environments as mutually constitutive rather than as separate entities linked by one-way effects (Uzzell & Räthzel, 2009).",
)
add_body_paragraph(
    doc,
    "Recent work has made this relationship increasingly concrete for design. Experimental studies show that built-environment features such as privacy, biophilia, clutter, and color temperature can shape perceived autonomy, competence, relatedness, and approach tendencies (Jutzi et al., 2025). Architectural experience is also multisensory. Visual form interacts with acoustics, touch, smell, temperature, proprioception, and bodily movement, although architectural representation and evaluation often privilege vision (Spence, 2020). The affordance concept offers a complementary account by linking environmental properties to the actions that a particular user can perform in a setting (Maier et al., 2009). Research at the intersection of neuroscience and architecture similarly describes how environmental features can guide attention and sensorimotor behavior, often before occupants formulate an explicit verbal account of the experience (Djebbara et al., 2022).",
)
add_body_paragraph(
    doc,
    "These perspectives are directly relevant to interior architecture education. Students are expected to design for users from the beginning of studio training, yet the theoretical language needed to distinguish psychological needs, sensory conditions, action possibilities, arousal, adaptation, and spatial perception may be introduced later or in separate courses. Studies in architecture and interior architecture education have therefore argued for stronger integration of human-space relations and environmental psychology within undergraduate curricula (Uzunoğlu & Özer, 2014; Aygenç & Özburak, 2024).",
)
add_body_paragraph(
    doc,
    "The design studio complicates a simple distinction between knowing and not knowing. Studio education develops professional judgment through repeated project work, critique, representation, and reflection in action (Schön, 1987). Students may account for a user, adjust circulation, or organise social and private zones without naming the psychological or perceptual principle that informs the decision. Observational research on studio interaction likewise shows that design knowledge can remain embedded in teacher-student dialogue and visual work rather than stated as formal theory (Ferreira et al., 2016). A relatively high baseline score before environmental psychology instruction may therefore reflect an intuitive human-centred understanding shaped by accumulated studio experience rather than prior mastery of the relevant terminology.",
)
add_body_paragraph(
    doc,
    "The present study examined environmental psychology learning among third-year interior architecture students at three successive points in one course unit. T0 documented intuitive human-centred understanding before a concept-focused lecture. T1 followed the lecture, in which selected environmental psychology concepts were named and explained. T2 followed an assignment in which students used the same concepts to analyse a personally experienced interior. The study did not test whether a lecture or an assignment was the superior educational method. It examined movement from initial understanding, through concept-focused teaching, to spatial analysis across self-report, written reflection, and submitted work.",
)

doc.add_paragraph("Research Questions", style="Heading 2")
questions = [
    "What intuitive human-centred understanding did students report before the lecture, and how was this profile related to their prior studio progression?",
    "What concept-specific changes appeared immediately after selected environmental psychology concepts were introduced and explained?",
    "Which changes remained, weakened, or changed direction after students used the concepts in a spatial analysis of a real interior?",
    "How did self-reports compare with assignment evidence and open-ended responses?",
]
for index, question in enumerate(questions, start=1):
    paragraph = doc.add_paragraph(style="List Number")
    paragraph.add_run(question)

doc.add_page_break()
doc.add_paragraph("Methods", style="Heading 1")

doc.add_paragraph("Research Design", style="Heading 2")
add_body_paragraph(
    doc,
    "The study used a course-embedded, single-group, three-wave mixed-methods design. Quantitative self-report data were interpreted alongside open-ended responses and rubric-based evidence from student assignments. T0 represented students' intuitive human-centred understanding before the concept-focused lecture. T1 represented the immediate profile after the lecture. T2 represented the profile after students used the concepts in a spatial analysis of a real interior. The stages were treated as connected stages of environmental psychology learning rather than as a randomised comparison of instructional methods. The design therefore supports within-person description and association, not causal attribution.",
)

doc.add_paragraph("Context and Participants", style="Heading 2")
add_body_paragraph(
    doc,
    "The study was conducted during the 2026-2027 autumn semester in a third-year Human Factors in Interior Space course within an interior architecture and environmental design programme at a private university in Türkiye. The course is positioned in the fifth academic semester. Students had therefore completed several studio-based design projects and had repeatedly designed spaces for defined user groups. Their most recently completed studio varied because of individual academic progression and course repetition. This variation was retained as contextual information rather than treated as a deficit indicator.",
)
add_body_paragraph(
    doc,
    "Students enrolled in the course were invited to participate voluntarily. Eligibility required course enrolment, an age of at least 18 years, and informed consent. Removing repeated submissions left 60 T0 records, 53 T1 records, and 40 T2 records. Thirty-four students could be matched across all three stages and formed the complete-case longitudinal sample. The assignment corpus contained 44 submissions.",
)

doc.add_paragraph("Procedure", style="Heading 2")
add_body_paragraph(
    doc,
    "T0 was administered before Lecture 1. The lecture introduced environmental psychology as the study of reciprocal person-environment relations and addressed user needs, multisensory experience, emotion and arousal, adaptation, affordances, and spatial Gestalt principles. T1 was administered immediately after the lecture.",
)
add_body_paragraph(
    doc,
    "Students then completed Assignment 1 Part I, Reading a Space Through Environmental Psychology. Each student selected a familiar interior that they had personally experienced and used photographs or other visual evidence to analyse how the environment and its users interacted. The assignment asked students to examine the space and its users, relevant psychological needs, multisensory conditions, pleasure and arousal, adaptation, at least four affordances, and at least three Gestalt principles operating spatially. T2 was administered after the assignment. Lecture 1 was the first unit in a broader course; the concepts examined here do not represent the full course content.",
)

doc.add_paragraph("Measures", style="Heading 2")
add_body_paragraph(
    doc,
    "The same eight study-specific Likert items were administered at T0, T1, and T2. Responses ranged from 1 strongly disagree to 5 strongly agree. The items assessed perceived ability to identify users and needs, address psychological needs, consider nonvisual senses, explain emotional and arousal effects, identify action possibilities, understand spatial organization, connect physical characteristics to human responses, and design for people with confidence. The equally weighted item mean formed a 1-5 composite. The measure assessed perceived understanding and capability, not objective achievement. Internal consistency was estimated separately at each stage using Cronbach's alpha.",
)
add_body_paragraph(
    doc,
    "All three forms included the same open-ended question about what should be understood regarding people who use a space. T1 also asked which lecture idea most changed or expanded students' thinking. T2 included three contextual evaluation items and two reflections on real-space analysis and the relative contribution of the lecture and assignment. The three contextual items were reported separately and were not included in the eight-item composite.",
)

doc.add_paragraph("Assignment Evidence", style="Heading 2")
add_body_paragraph(
    doc,
    "The assignments were re-reviewed for research purposes using five dimensions scored from 0 to 3: human-space relationship, psychological needs, multisensory experience with arousal and adaptation, affordances, and spatial Gestalt. A score of 0 indicated absent or incorrect use, 1 indicated descriptive or superficial use, 2 indicated correct application to spatial evidence, and 3 indicated interpretive application linking spatial features to perception, experience, or behavior. The five scores produced a total from 0 to 15. One researcher scored and re-reviewed all 44 submissions. The scores were separate from course grades and were used as descriptive supporting evidence.",
)
add_body_paragraph(
    doc,
    "For an exploratory subgroup analysis, high-detail assignments were defined as totals of at least 13 of 15 with no dimension below 2. This rule required broad, correct application rather than a high total produced by one strong dimension. It was an analytical threshold, not a validated mastery standard.",
)

doc.add_paragraph("Data Preparation", style="Heading 2")
add_body_paragraph(
    doc,
    "Names were standardized only to link repeated surveys and assignments. The analytic dataset used pseudonymous participant identifiers and excluded direct identifiers and verbatim free-text responses. When a participant submitted more than one response at the same stage, the latest complete response was retained as the primary record. An earliest-response sensitivity analysis assessed whether this choice changed the longitudinal conclusions. Missing values were not imputed, and the primary repeated-measures tests used the 34 complete cases.",
)

doc.add_paragraph("Quantitative Analysis", style="Heading 2")
add_body_paragraph(
    doc,
    "Means, standard deviations, medians, and interquartile ranges described the survey and assignment scores. Overall T0-T1-T2 differences were tested with the Friedman test. Paired Wilcoxon signed-rank tests examined T0-T1, T1-T2, and T0-T2 contrasts, with Holm correction within each outcome. Paired rank-biserial correlations described effect size, oriented so that positive values indicated higher scores at the later stage. Item-level tests were treated as exploratory. Baseline T0 scores of complete and incomplete longitudinal cases were compared with the Mann-Whitney U test.",
)
add_body_paragraph(
    doc,
    "Spearman correlations examined assignment totals in relation to T2 self-report, T1-T2 change, and prior studio progression. The studio correlations for the assignment total and five dimensions were Holm-adjusted as one exploratory family. High-detail and other scored assignments were compared on T1-T2 change using the Mann-Whitney U test and Cliff's delta. Students without a matched assignment record were not classified as low-performing. Analyses were implemented in Python 3.12 using pandas 3.0.1, NumPy 2.3.5, and openpyxl 3.1.5. Figures were generated with Pillow 12.3.0.",
)

doc.add_paragraph("Open Ended Responses", style="Heading 2")
add_body_paragraph(
    doc,
    "Open-ended responses were summarized with concept-guided, keyword-assisted, multi-label indicators. The categories were needs comfort and function, emotion and psychology, individuality and accessibility, sensory experience and atmosphere, behavior movement and use, and safety privacy and social life. One response could receive more than one indicator. This procedure was used to map the visibility of concepts across stages and was not treated as an independent, fully interpretive thematic analysis.",
)

doc.add_paragraph("Ethics", style="Heading 2")
add_body_paragraph(
    doc,
    "The study received approval from the relevant institutional ethics committee. Participation was voluntary, and students were informed that participation or withdrawal would not affect grades, assignment evaluation, access to the course, or their relationship with the instructor. Research assignment scores were kept separate from course grading, and results were reported in aggregate form.",
)

doc.add_page_break()
doc.add_paragraph("Results", style="Heading 1")

doc.add_paragraph("Participant Flow and Measurement Quality", style="Heading 2")
add_body_paragraph(
    doc,
    "After duplicate submissions were removed, the dataset contained 60 students at T0, 53 at T1, and 40 at T2. The complete-case longitudinal sample included 34 students. The eight-item scale showed acceptable to high internal consistency in the full stage-specific samples and in the matched sample. Cronbach's alpha ranged from .825 at T0 to .923 at T2 in the available-stage data and from .824 to .921 in the matched sample.",
)

add_table_caption(doc, 1, "Participant flow and internal consistency")
table = doc.add_table(rows=1, cols=7)
headers = ["Stage", "Raw rows", "Unique students", "Duplicates removed", "Alpha all", "Matched n", "Alpha matched"]
for index, value in enumerate(headers):
    table.rows[0].cells[index].text = value
for record in flow:
    cells = table.add_row().cells
    values = [
        record["stage"],
        str(record["raw"]),
        str(record["unique"]),
        str(record["duplicates"]),
        f"{record['alpha_all']:.3f}",
        str(record["matched_n"]),
        f"{record['alpha_matched']:.3f}",
    ]
    for index, value in enumerate(values):
        cells[index].text = value
style_table(table, [0.58, 0.75, 0.95, 1.05, 0.75, 0.75, 0.92], left_columns={0}, font_size=8.1)
add_table_note(doc, "Alpha denotes Cronbach's alpha. Longitudinal inference used the same 34 students at all three stages.")

add_body_paragraph(
    doc,
    f"The baseline composite did not differ statistically between students retained in the complete-case analysis and students without all three measurements, U = {attrition['u']:.0f}, p {p_value(attrition['p'])}. Their T0 means were {attrition['matched_mean']:.2f} and {attrition['not_complete_mean']:.2f}, respectively.",
)

doc.add_paragraph("Baseline Person Environment Profile", style="Heading 2")
add_body_paragraph(
    doc,
    "The T0 composite mean was 3.90 on the five-point scale. Baseline item means were already above 3.60. Psychological needs had the highest T0 mean at 4.15, followed by identifying users and needs at 4.12. Emotion and arousal had the lowest mean at 3.62, closely followed by multisensory design at 3.65. This pattern documents a comparatively strong initial self-assessment alongside greater room for development in sensory and affective aspects of person-environment reasoning.",
)
studio_counts = matched["Prior studio"].value_counts().reindex(STUDIO_ORDER.keys(), fill_value=0)
add_body_paragraph(
    doc,
    "The most recent successfully completed studio varied within the matched sample: IAED 1001 n = "
    f"{studio_counts['IAED 1001']}, IAED 1002 n = {studio_counts['IAED 1002']}, IAED 2001 n = {studio_counts['IAED 2001']}, "
    f"IAED 2002 n = {studio_counts['IAED 2002']}, and IAED 3001 n = {studio_counts['IAED 3001']}. "
    f"Studio rank was not detectably associated with the T0 composite, Spearman rho = {studio_t0_rho:.2f}, p {p_value(studio_t0_p)}. The distribution was strongly concentrated in IAED 2002, so this estimate should be read cautiously.",
)

doc.add_paragraph("Overall Three Stage Trajectory", style="Heading 2")
composite = longitudinal[-1]
add_body_paragraph(
    doc,
    f"The composite increased from T0 M = {composite['t0']:.2f}, SD = {composite['sd0']:.2f}, to T1 M = {composite['t1']:.2f}, SD = {composite['sd1']:.2f}, and then decreased to T2 M = {composite['t2']:.2f}, SD = {composite['sd2']:.2f}. The overall stage difference was statistically significant, Friedman Q(2) = {composite['q']:.2f}, p {p_value(composite['p'])}. Holm-adjusted comparisons showed an increase from T0 to T1, p {p_value(composite['p01'])}, rank-biserial r = {composite['r01']:.2f}; a decrease from T1 to T2, p {p_value(composite['p12'])}, r = {composite['r12']:.2f}; and a higher T2 score than T0, p {p_value(composite['p02'])}, r = {composite['r02']:.2f}.",
)
add_body_paragraph(
    doc,
    f"The retained T0-T2 difference depended on the duplicate-response rule. When the earliest rather than latest response was retained, the T0-T2 comparison was not statistically significant after Holm correction, p {p_value(attrition['earliest_p02'])}. The immediate T0-T1 increase and the T1-T2 decline remained statistically detectable under both rules.",
)
add_figure(
    doc,
    FIGURE_PATHS["composite"],
    1,
    "Individual and mean trajectories for the eight-item composite. Bars show 95% bootstrap confidence intervals. The figure displays the 34 participants with data at all three stages.",
    "Thirty-four individual lines rise on average from T0 to T1 and then decline at T2 while remaining above T0. Mean values are 3.90, 4.32, and 4.09.",
)

doc.add_paragraph("Concept Specific Change", style="Heading 2")
add_table_caption(doc, 2, "Item-level matched results")
table = doc.add_table(rows=1, cols=7)
headers = ["Measure", "T0 M SD", "T1 M SD", "T2 M SD", "Friedman Q", "p", "Holm-significant contrasts"]
for index, value in enumerate(headers):
    table.rows[0].cells[index].text = value
contrast_labels = []
for record in longitudinal[:-1]:
    contrasts = []
    if record["p01"] < 0.05:
        contrasts.append("T0-T1")
    if record["p12"] < 0.05:
        contrasts.append("T1-T2")
    if record["p02"] < 0.05:
        contrasts.append("T0-T2")
    contrast_labels.append(", ".join(contrasts) if contrasts else "None")
    cells = table.add_row().cells
    values = [
        record["measure"],
        f"{record['t0']:.2f}  {record['sd0']:.2f}",
        f"{record['t1']:.2f}  {record['sd1']:.2f}",
        f"{record['t2']:.2f}  {record['sd2']:.2f}",
        f"{record['q']:.2f}",
        "< .001" if record["p"] < 0.001 else f"{record['p']:.3f}".lstrip("0"),
        contrast_labels[-1],
    ]
    for index, value in enumerate(values):
        cells[index].text = value
style_table(table, [1.75, 0.75, 0.75, 0.75, 0.75, 0.55, 1.15], left_columns={0, 6}, font_size=7.6)
add_table_note(doc, "M and SD are shown for interpretability. Inference used rank-based tests. Pairwise p values were Holm-adjusted within each item.")

add_body_paragraph(
    doc,
    "Multisensory design showed the clearest retained change. Its mean increased from 3.65 at T0 to 4.35 at T1 and remained at 4.15 at T2. Both T0-T1 and T0-T2 comparisons were significant after Holm adjustment, p = .002 for each. Emotion and arousal increased from 3.62 to 4.29 after the lecture, p = .003, and decreased to 3.88 after the assignment, p = .033 for T1-T2. The T0-T2 contrast for this item was not significant.",
)
add_body_paragraph(
    doc,
    "The physical-environment to human-response item increased from T0 to T1, p = .039. Overall confidence also increased from T0 to T1, p = .021, and decreased from T1 to T2, p = .014, returning to the T0 mean. Spatial perception showed a significant omnibus difference, p = .015, although none of its Holm-adjusted pairwise contrasts crossed the .05 threshold. User identification, psychological needs, and affordances did not show statistically significant omnibus differences.",
)
add_figure(
    doc,
    FIGURE_PATHS["items"],
    2,
    "Item means at T0, T1, and T2 with mean changes for each stage contrast. The score panel is limited to the observed range to make item differences legible; the underlying response scale ranged from 1 to 5.",
    "Eight rows show item means across three stages and a heatmap of T0 to T1, T1 to T2, and T0 to T2 mean changes. Multisensory design has the largest retained increase.",
)

doc.add_paragraph("Evidence from Real Space Assignments", style="Heading 2")
assignment_mean = float(assignments["Total /15"].mean())
assignment_sd = float(assignments["Total /15"].std(ddof=1))
assignment_median = float(assignments["Total /15"].median())
assignment_q1 = float(assignments["Total /15"].quantile(0.25))
assignment_q3 = float(assignments["Total /15"].quantile(0.75))
add_body_paragraph(
    doc,
    f"The 44 assignment totals averaged {assignment_mean:.2f} of 15, SD = {assignment_sd:.2f}, with a median of {assignment_median:.0f} and an interquartile range of {assignment_q1:.0f}-{assignment_q3:.0f}. Affordances had the highest dimension mean at 2.80 of 3. Multisensory experience, arousal, and adaptation followed at 2.75. Psychological needs had the lowest dimension mean at 2.43. Spatial Gestalt showed the largest dispersion, SD = 0.98, including several submissions in which the concept was absent or treated only as a two-dimensional graphic principle.",
)

add_table_caption(doc, 3, "Research assignment scores by rubric dimension")
table = doc.add_table(rows=1, cols=5)
for index, value in enumerate(["Dimension", "Mean", "Median", "SD", "Scale"]):
    table.rows[0].cells[index].text = value
for record in assignment_dimensions:
    cells = table.add_row().cells
    for index, value in enumerate(
        [record["dimension"], f"{record['mean']:.2f}", f"{record['median']:.0f}", f"{record['sd']:.2f}", "0-3"]
    ):
        cells[index].text = value
style_table(table, [3.45, 0.75, 0.75, 0.75, 0.75], left_columns={0}, font_size=8.4)
add_table_note(doc, "Scores were assigned by one researcher and are used as descriptive supporting evidence. The total score ranged from 0 to 15.")

add_figure(
    doc,
    FIGURE_PATHS["assignment"],
    3,
    "Rubric-score distributions for five assignment dimensions and the distribution of total scores. Percentages in the left panel describe the share of 44 submissions receiving each score.",
    "Stacked bars show scores of zero to three for five assignment dimensions. A second panel shows assignment totals concentrated between 12 and 15, with a median of 14.",
)

add_body_paragraph(
    doc,
    f"Assignment total was not significantly associated with the T2 composite, n = {assignment_associations['t2_n']}, Spearman rho = {assignment_associations['t2_rho']:.2f}, p {p_value(assignment_associations['t2_p'])}. It was also not associated with T1-T2 composite change among the {assignment_associations['change_n']} students with both longitudinal and assignment data, rho = {assignment_associations['change_rho']:.2f}, p {p_value(assignment_associations['change_p'])}. Self-reported capability and demonstrated use of the concepts therefore did not align as a single continuous measure in this sample.",
)

doc.add_paragraph("Studio Progression and Assignment Detail", style="Heading 2")
add_body_paragraph(
    doc,
    f"Among the 28 complete-case students with a matched assignment, later studio progression was associated with higher assignment totals, Spearman rho = {studio_total_rho:.2f}, raw p {p_value(studio_total_p)}, Holm-adjusted p {p_value(studio_total_p_holm)}. The strongest dimension-specific association was for affordances, rho = {next(item['rho'] for item in studio_dim_correlations if item['dimension'] == 'Affordances'):.2f}, Holm-adjusted p {p_value(next(item['p_holm'] for item in studio_dim_correlations if item['dimension'] == 'Affordances'))}. Multisensory experience, arousal, and adaptation was also associated with studio rank, rho = {next(item['rho'] for item in studio_dim_correlations if item['dimension'] == 'Multisensory / arousal / adaptation'):.2f}, Holm-adjusted p {p_value(next(item['p_holm'] for item in studio_dim_correlations if item['dimension'] == 'Multisensory / arousal / adaptation'))}. The remaining dimension-specific adjusted p values exceeded .05.",
)
add_body_paragraph(
    doc,
    f"Seventeen of the 28 students with a matched assignment met the high-detail rule; 11 had scored work below that rule, and 6 additional complete-case students had no matched assignment record. The high-detail group changed from M = {high_group['T1 mean'].mean():.2f} at T1 to M = {high_group['T2 mean'].mean():.2f} at T2, mean change = {high_group['T1 to T2 change'].mean():+.2f}, Wilcoxon p {p_value(high_wilcoxon_p)}, rank-biserial r = {high_rank_biserial:.2f}. The other scored group changed from M = {other_scored_group['T1 mean'].mean():.2f} to M = {other_scored_group['T2 mean'].mean():.2f}, mean change = {other_scored_group['T1 to T2 change'].mean():+.2f}, Wilcoxon p {p_value(other_wilcoxon_p)}, r = {other_rank_biserial:.2f}. The between-group difference in change was not significant, Mann-Whitney U = {subgroup_u:.1f}, p {p_value(subgroup_p)}, Cliff's delta = {subgroup_cliff:+.2f}.",
)
add_body_paragraph(
    doc,
    f"High-detail work was more frequent among students whose latest completed studio was IAED 2002 or IAED 3001 than among those at IAED 1001, IAED 1002, or IAED 2001, Fisher's exact p {p_value(studio_high_fisher_p)}. This exploratory comparison combines small and unequal studio groups and cannot separate studio progression from broader academic progression.",
)
add_figure(
    doc,
    FIGURE_PATHS["studio"],
    4,
    "Assignment totals by most recent completed studio and T1-T2 self-report change by assignment-detail group. Six complete-case students without a matched assignment record are excluded from the right panel.",
    "The left panel shows generally higher assignment totals at later completed studio levels. The right panel shows overlapping T1 to T2 changes for 17 high-detail and 11 other scored assignments.",
)

doc.add_paragraph("Open Ended Indicators", style="Heading 2")
add_body_paragraph(
    doc,
    "Needs, comfort, and function remained the most frequent indicator in the common open-ended question, appearing in 76.7% of T0 responses, 88.7% of T1 responses, and 87.5% of T2 responses. Behavior, movement, and use increased from 16.7% at T0 to 26.4% at T1 and 27.5% at T2. Safety, privacy, and social life increased from 10.0% to 17.5% between T0 and T2. Emotion and psychology rose from 43.3% at T0 to 54.7% at T1 and then declined to 40.0% at T2. Sensory and atmosphere terms remained relatively infrequent in this broad question despite the retained increase in the corresponding self-report item.",
)
add_figure(
    doc,
    FIGURE_PATHS["themes"],
    5,
    "Prevalence of six concept-guided indicators in the common open-ended question. Coding was multi-label, so a response could contribute to more than one indicator.",
    "A dot and line plot shows needs comfort and function as the most frequent theme at all stages. Behavior movement and use and safety privacy and social life rise by T2.",
)
add_body_paragraph(
    doc,
    "In the T1 lecture reflection, psychology, emotion, and Maslow-related ideas were identified in 34 of 53 responses, 64.2%. Human-centered thinking beyond appearance was identified in 16 responses, 30.2%. Sensory experience appeared in 7 responses, behavior and affordances in 6, Gestalt and perception in 3, and arousal and adaptation in 1. These categories were not mutually exclusive.",
)
add_body_paragraph(
    doc,
    "At T2, 29 of 40 students, 72.5%, identified the combination of lecture and assignment as the most helpful component. Seven selected the assignment or application, three selected the lecture, and one response was unclear. Thirty-two students, 80.0%, gave an affirmative response that analyzing a real space changed their thinking; six responses were nuanced or unclear, and two indicated no or limited change. The three T2 evaluation items were also favorable: concepts clearer M = 4.28 with 90.0% selecting 4 or 5; real-space application more helpful than the lecture alone M = 4.28 with 85.0% selecting 4 or 5; and benefit justified time and effort M = 4.23 with 77.5% selecting 4 or 5.",
)

doc.add_paragraph("Integrated Results", style="Heading 2")
add_body_paragraph(
    doc,
    "The three evidence sources did not produce a uniform pattern. Multisensory design showed the largest retained self-report increase and a high assignment mean, although sensory language remained uncommon in the broad open-ended question. Affordances had the highest assignment score and behavior-related language became more frequent, but the corresponding self-report item did not change significantly. Psychological needs were highly rated at baseline and dominated open-ended responses, yet this was the lowest-scoring assignment dimension. These differences indicate that familiarity, confidence, spontaneous language, and demonstrated application captured related but distinct aspects of environmental psychology understanding.",
)

add_table_caption(doc, 4, "Joint display of concept-specific evidence")
table = doc.add_table(rows=1, cols=4)
for index, value in enumerate(["Domain", "Self-report trajectory", "Assignment evidence", "Open-ended evidence"]):
    table.rows[0].cells[index].text = value
joint_rows = [
    [
        "Psychological needs",
        "High at T0 and no significant stage difference",
        "Lowest dimension mean 2.43 of 3",
        "Needs comfort and function remained dominant",
    ],
    [
        "Multisensory experience",
        "Largest retained T0-T2 increase plus 0.50",
        "High mean 2.75 of 3",
        "Sensory terms remained uncommon in the broad prompt",
    ],
    [
        "Emotion and arousal",
        "Increase at T1 followed by T2 decline",
        "Included within multisensory arousal adaptation scoring",
        "Emotion terms peaked at T1 and declined at T2",
    ],
    [
        "Affordances",
        "No significant omnibus stage difference",
        "Highest dimension mean 2.80 of 3",
        "Behavior movement and use increased by T2",
    ],
    [
        "Spatial perception and Gestalt",
        "Omnibus difference without significant pairwise contrasts",
        "Mean 2.55 with the greatest dispersion",
        "Gestalt appeared infrequently in the T1 reflection",
    ],
]
for row_values in joint_rows:
    cells = table.add_row().cells
    for index, value in enumerate(row_values):
        cells[index].text = value
style_table(table, [1.3, 1.75, 1.75, 1.75], left_columns={0, 1, 2, 3}, font_size=7.7)
add_table_note(doc, "The joint display summarizes convergence and divergence across self-report, submitted work, and keyword-assisted open-ended indicators.")

doc.add_page_break()
doc.add_paragraph("References", style="Heading 1")
references = [
    "Aygenç, B., & Özburak, C. (2024). The role and significance of human-space relationship and psychology in interior architecture and architectural education: A proposed training model. Prostor, 32(1), 84-97. https://doi.org/10.31522/p.32.1(67).8",
    "Djebbara, Z., Jensen, O. B., Parada, F. J., & Gramann, K. (2022). Neuroscience and architecture: Modulating behavior through sensorimotor responses to the built environment. Neuroscience & Biobehavioral Reviews, 138, 104715. https://doi.org/10.1016/j.neubiorev.2022.104715",
    "Ferreira, J., Christiaans, H., & Almendra, R. (2016). A visual tool for analysing teacher and student interactions in a design studio setting. CoDesign, 12(1-2), 112-131. https://doi.org/10.1080/15710882.2015.1135246",
    "Jutzi, C. A., Möller, J., Hansen, J., Klackl, J., & Jonas, E. (2025). Psychological needs in the built environment. Journal of Environmental Psychology, 101, 102419. https://doi.org/10.1016/j.jenvp.2024.102419",
    "Maier, J. R. A., Fadel, G. M., & Battisto, D. G. (2009). An affordance-based approach to architectural theory, design, and practice. Design Studies, 30(4), 393-414. https://doi.org/10.1016/j.destud.2009.01.002",
    "Schön, D. A. (1987). Educating the reflective practitioner. Jossey-Bass.",
    "Spence, C. (2020). Senses of place: Architectural design for the multisensory mind. Cognitive Research: Principles and Implications, 5, 46. https://doi.org/10.1186/s41235-020-00243-4",
    "Uzzell, D., & Räthzel, N. (2009). Transforming environmental psychology. Journal of Environmental Psychology, 29(3), 340-350. https://doi.org/10.1016/j.jenvp.2008.11.005",
    "Uzunoğlu, S. S., & Özer, H. (2014). A model for the development of architectural psychology formation in architectural education. Megaron, 9(2), 143-165. https://doi.org/10.5505/megaron.2014.39974",
]
for reference in references:
    add_reference(doc, reference)

# Widow control and orphan protection.
for paragraph in doc.paragraphs:
    properties = paragraph._p.get_or_add_pPr()
    widow = properties.find(qn("w:widowControl"))
    if widow is None:
        widow = OxmlElement("w:widowControl")
        properties.append(widow)
    widow.set(qn("w:val"), "1")

doc.save(OUTPUT_DOCX)

print(OUTPUT_DOCX)
print(
    {
        "paragraphs": len(doc.paragraphs),
        "tables": len(doc.tables),
        "figures": len(doc.inline_shapes),
        "matched_n": len(matched),
        "matched_assignment_n": len(matched_assign),
        "high_detail_n": len(high_group),
        "other_scored_n": len(other_scored_group),
        "studio_total_rho": studio_total_rho,
        "studio_total_p": studio_total_p,
        "studio_total_p_holm": studio_total_p_holm,
        "subgroup_p": subgroup_p,
    }
)
