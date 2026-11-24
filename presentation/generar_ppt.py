"""
Generador automático de la presentación grupal del TPI de Análisis de Datos
(Telco Customer Churn).

Acciones que realiza el script:
1. Extrae los gráficos PNG embebidos en notebooks/TPI_Analisis_de_Datos.ipynb.
2. Genera gráficos ejecutivos de alta resolución (300 DPI) a partir de
   data/processed/telco_churn_clean.csv con la paleta corporativa.
3. Construye presentation/Presentacion_TPI_Telco_Churn.pptx (12 diapositivas,
   formato 16:9) con tarjetas KPI, tablas, gráficos embebidos y notas del orador.
"""

from __future__ import annotations

import base64
import json
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt


# =============================================================================
# RUTAS Y PALETA CORPORATIVA
# =============================================================================
PRES_DIR = Path(__file__).resolve().parent
REPO_DIR = PRES_DIR.parent
CHARTS_DIR = PRES_DIR / "charts"
NB_PATH = REPO_DIR / "notebooks" / "TPI_Analisis_de_Datos.ipynb"
CLEAN_CSV = REPO_DIR / "data" / "processed" / "telco_churn_clean.csv"
OUTPUT_PPTX = PRES_DIR / "Presentacion_TPI_Telco_Churn.pptx"

# Paleta corporativa (Azul oscuro / Cian / Gris claro + Coral alerta / Turquesa retención)
HEX_NAVY = "#0F2942"
HEX_SLATE = "#1E4E79"
HEX_CYAN = "#00A8E8"
HEX_TEAL = "#00A896"
HEX_CORAL = "#E05A2B"
HEX_AMBER = "#E69F00"
HEX_BG = "#F4F7FA"
HEX_WHITE = "#FFFFFF"
HEX_BORDER = "#D8E2EC"
HEX_TEXT = "#1A2B3C"
HEX_MUTED = "#5A6E82"
HEX_SOFT_BLUE = "#EAF4FB"
HEX_SOFT_CORAL = "#FDF0EB"
HEX_SOFT_TEAL = "#E8F7F5"

C_NAVY = RGBColor(0x0F, 0x29, 0x42)
C_SLATE = RGBColor(0x1E, 0x4E, 0x79)
C_CYAN = RGBColor(0x00, 0xA8, 0xE8)
C_TEAL = RGBColor(0x00, 0xA8, 0x96)
C_CORAL = RGBColor(0xE0, 0x5A, 0x2B)
C_AMBER = RGBColor(0xE6, 0x9F, 0x00)
C_BG = RGBColor(0xF4, 0xF7, 0xFA)
C_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
C_BORDER = RGBColor(0xD8, 0xE2, 0xEC)
C_TEXT = RGBColor(0x1A, 0x2B, 0x3C)
C_MUTED = RGBColor(0x5A, 0x6E, 0x82)
C_SOFT_BLUE = RGBColor(0xEA, 0xF4, 0xFB)
C_SOFT_CORAL = RGBColor(0xFD, 0xF0, 0xEB)
C_SOFT_TEAL = RGBColor(0xE8, 0xF7, 0xF5)
C_DARK_CARD = RGBColor(0x17, 0x3A, 0x5C)

FONT_MAIN = "Segoe UI"
SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)
TOTAL_SLIDES = 12


# =============================================================================
# 1. EXTRACCIÓN DE GRÁFICOS DEL NOTEBOOK
# =============================================================================
def exportar_graficos_notebook() -> dict[str, Path]:
    """Extrae los gráficos PNG embebidos en el notebook oficial."""
    CHARTS_DIR.mkdir(parents=True, exist_ok=True)
    with open(NB_PATH, encoding="utf-8") as f:
        nb = json.load(f)

    nombres = [
        "01_nb_boxplots_outliers.png",
        "02_nb_heatmap_pearson.png",
        "03_nb_heatmap_ampliado.png",
        "04_nb_scatter_relaciones.png",
    ]
    rutas: dict[str, Path] = {}
    idx = 0
    for cell in nb["cells"]:
        if cell.get("cell_type") == "code":
            for out in cell.get("outputs", []):
                data = out.get("data", {})
                if "image/png" in data:
                    b64 = data["image/png"]
                    if isinstance(b64, list):
                        b64 = "".join(b64)
                    fname = nombres[idx] if idx < len(nombres) else f"nb_chart_{idx + 1}.png"
                    out_path = CHARTS_DIR / fname
                    out_path.write_bytes(base64.b64decode(b64))
                    rutas[fname] = out_path
                    idx += 1
    return rutas


# =============================================================================
# 2. GENERACIÓN DE GRÁFICOS EJECUTIVOS (300 DPI) DESDE EL DATASET DEPURADO
# =============================================================================
def configurar_matplotlib() -> None:
    plt.rcParams.update({
        "font.family": "sans-serif",
        "font.sans-serif": ["Segoe UI", "DejaVu Sans", "Arial"],
        "axes.edgecolor": "#CBD5E1",
        "axes.linewidth": 0.8,
        "text.color": HEX_TEXT,
        "axes.labelcolor": HEX_TEXT,
        "xtick.color": HEX_MUTED,
        "ytick.color": HEX_TEXT,
        "figure.facecolor": "white",
        "axes.facecolor": "white",
    })


def fmt_es(num: float, dec: int = 1) -> str:
    """Formatea un número con coma decimal estilo español (ej. 26,54)."""
    return f"{num:,.{dec}f}".replace(",", "_").replace(".", ",").replace("_", ".")


def generar_graficos_ejecutivos(df: pd.DataFrame) -> dict[str, Path]:
    """Genera los gráficos de soporte para las diapositivas de EDA y hallazgos."""
    configurar_matplotlib()
    rutas: dict[str, Path] = {}
    global_churn = df["Churn_num"].mean() * 100
    label_pill = dict(facecolor="white", edgecolor="none", pad=1.4, alpha=0.92)

    # -------------------------------------------------------------------------
    # Gráfico 05: Panorama General (Donut de Churn + Impacto MRR)
    # -------------------------------------------------------------------------
    activos = int((df["Churn_num"] == 0).sum())
    bajas = int((df["Churn_num"] == 1).sum())
    mrr_ret = df.loc[df["Churn_num"] == 0, "MonthlyCharges"].sum()
    mrr_loss = df.loc[df["Churn_num"] == 1, "MonthlyCharges"].sum()
    mrr_tot = mrr_ret + mrr_loss

    fig, (ax1, ax2) = plt.subplots(
        1, 2, figsize=(6.2, 3.1), gridspec_kw={"width_ratios": [1.1, 1.0]}
    )

    # Donut
    wedges, _ = ax1.pie(
        [activos, bajas],
        startangle=90,
        colors=[HEX_NAVY, HEX_CORAL],
        wedgeprops=dict(width=0.38, edgecolor="white", linewidth=2.5),
        counterclock=False,
    )
    ax1.text(0, 0.08, f"{fmt_es(global_churn, 2)}%", ha="center", va="center",
             fontsize=18.5, fontweight="bold", color=HEX_CORAL)
    ax1.text(0, -0.18, "Tasa de Churn", ha="center", va="center",
             fontsize=9.5, color=HEX_MUTED, fontweight="bold")
    ax1.set_title("Clientes (n = 7.043)", fontsize=11.5, fontweight="bold",
                  color=HEX_NAVY, pad=6)
    ax1.legend(
        wedges,
        [f"Activos: {activos:,}".replace(",", ".") + " (73,5 %)",
         f"Bajas: {bajas:,}".replace(",", ".") + " (26,5 %)"],
        loc="lower center", bbox_to_anchor=(0.5, -0.14), frameon=False, fontsize=8.8
    )

    # Barras de impacto MRR
    cats = ["MRR\nRetenido", "MRR\nPerdido"]
    vals = [mrr_ret / 1000, mrr_loss / 1000]
    pcts = [mrr_ret / mrr_tot * 100, mrr_loss / mrr_tot * 100]
    bars = ax2.bar(cats, vals, color=[HEX_NAVY, HEX_CORAL], width=0.52, edgecolor="none")
    ax2.set_ylim(0, max(vals) * 1.30)
    ax2.set_title("Impacto en Ingreso Mensual", fontsize=11.5, fontweight="bold",
                  color=HEX_NAVY, pad=6)
    ax2.set_ylabel("Miles de USD / mes", fontsize=9, color=HEX_MUTED)
    ax2.spines["top"].set_visible(False)
    ax2.spines["right"].set_visible(False)
    ax2.grid(axis="y", linestyle="--", alpha=0.35)
    for bar, val, pct in zip(bars, vals, pcts):
        ax2.text(
            bar.get_x() + bar.get_width() / 2, val + 8,
            f"US$ {fmt_es(val, 1)}k\n({fmt_es(pct, 1)} %)",
            ha="center", va="bottom", fontsize=9.2, fontweight="bold",
            color=HEX_CORAL if pct < 50 else HEX_NAVY
        )

    plt.tight_layout(pad=1.0)
    p5 = CHARTS_DIR / "05_churn_overview_donut.png"
    fig.savefig(p5, dpi=300, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    rutas[p5.name] = p5

    # -------------------------------------------------------------------------
    # Gráfico 06: Contratos y Métodos de Pago (Slide 7)
    # -------------------------------------------------------------------------
    fig, (ax1, ax2) = plt.subplots(
        2, 1, figsize=(7.6, 5.1), gridspec_kw={"height_ratios": [1.0, 1.24]}
    )

    # Panel A: Tipo de Contrato
    orden_c = ["Two year", "One year", "Month-to-month"]
    labels_c = ["Bianual (Two year)", "Anual (One year)", "Mes a mes (Month-to-month)"]
    g_c = df.groupby("Contract", observed=False)["Churn_num"].agg(["mean", "count"]).loc[orden_c]
    rates_c = g_c["mean"] * 100
    counts_c = g_c["count"]
    colors_c = [HEX_TEAL, HEX_SLATE, HEX_CORAL]

    ax1.axvline(global_churn, color=HEX_NAVY, linestyle="--", linewidth=1.2, alpha=0.75, zorder=1)
    bars1 = ax1.barh(labels_c, rates_c, color=colors_c, height=0.58, zorder=2)
    ax1.text(global_churn + 0.9, 0.08, f"Promedio global: {fmt_es(global_churn, 1)} %",
             fontsize=8.5, color=HEX_NAVY, fontweight="bold", bbox=label_pill, zorder=4)
    ax1.set_xlim(0, 62)
    ax1.set_title("Tasa de Churn por Tipo de Contrato (%)", fontsize=11.5,
                  fontweight="bold", color=HEX_NAVY, loc="left", pad=8)
    ax1.spines["top"].set_visible(False)
    ax1.spines["right"].set_visible(False)
    ax1.grid(axis="x", linestyle="--", alpha=0.32, zorder=0)
    for bar, r, n in zip(bars1, rates_c, counts_c):
        n_txt = f"{int(n):,}".replace(",", ".")
        ax1.text(
            r + 1.0, bar.get_y() + bar.get_height() / 2,
            f"{fmt_es(r, 2)} %  (n = {n_txt})",
            va="center", fontsize=9.4, fontweight="bold",
            color=HEX_CORAL if r > 30 else HEX_TEXT,
            bbox=label_pill, zorder=5
        )

    # Panel B: Método de Pago
    orden_p = [
        "Credit card (automatic)",
        "Bank transfer (automatic)",
        "Mailed check",
        "Electronic check",
    ]
    labels_p = [
        "Tarjeta de crédito (auto)",
        "Transferencia bancaria (auto)",
        "Cheque por correo (manual)",
        "Cheque electrónico (manual)",
    ]
    g_p = df.groupby("PaymentMethod", observed=False)["Churn_num"].agg(["mean", "count"]).loc[orden_p]
    rates_p = g_p["mean"] * 100
    counts_p = g_p["count"]
    colors_p = [HEX_TEAL, HEX_TEAL, HEX_SLATE, HEX_CORAL]

    ax2.axvline(global_churn, color=HEX_NAVY, linestyle="--", linewidth=1.2, alpha=0.75, zorder=1)
    bars2 = ax2.barh(labels_p, rates_p, color=colors_p, height=0.58, zorder=2)
    ax2.set_xlim(0, 65)
    ax2.set_xlabel("Tasa de abandono (%)", fontsize=9.5, color=HEX_MUTED)
    ax2.set_title("Tasa de Churn por Método de Pago (%)", fontsize=11.5,
                  fontweight="bold", color=HEX_NAVY, loc="left", pad=8)
    ax2.spines["top"].set_visible(False)
    ax2.spines["right"].set_visible(False)
    ax2.grid(axis="x", linestyle="--", alpha=0.32, zorder=0)
    for bar, r, n in zip(bars2, rates_p, counts_p):
        n_txt = f"{int(n):,}".replace(",", ".")
        ax2.text(
            r + 1.0, bar.get_y() + bar.get_height() / 2,
            f"{fmt_es(r, 2)} %  (n = {n_txt})",
            va="center", fontsize=9.4, fontweight="bold",
            color=HEX_CORAL if r > 30 else HEX_TEXT,
            bbox=label_pill, zorder=5
        )

    plt.tight_layout(h_pad=1.6)
    p6 = CHARTS_DIR / "06_contratos_y_pagos.png"
    fig.savefig(p6, dpi=300, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    rutas[p6.name] = p6

    # -------------------------------------------------------------------------
    # Gráfico 07: Antigüedad (Tenure) y Servicios Protectores (Slide 8)
    # -------------------------------------------------------------------------
    fig, (ax1, ax2) = plt.subplots(
        1, 2, figsize=(8.2, 4.9), gridspec_kw={"width_ratios": [1.02, 1.25]}
    )

    # Panel Izquierdo: Churn por tramo de antigüedad (sin línea diagonal que tape etiquetas)
    bands = pd.cut(df["tenure"], bins=[-1, 12, 24, 48, 72],
                   labels=["0–12 m", "13–24 m", "25–48 m", "49–72 m"])
    g_t = df.groupby(bands, observed=False)["Churn_num"].agg(["mean", "count"])
    t_rates = g_t["mean"] * 100
    t_counts = g_t["count"]
    t_colors = [HEX_CORAL, HEX_AMBER, HEX_SLATE, HEX_TEAL]

    x_pos = np.arange(len(t_rates))
    ax1.axhline(global_churn, color=HEX_NAVY, linestyle="--", linewidth=1.15, alpha=0.75, zorder=1)
    bars_t = ax1.bar(x_pos, t_rates, color=t_colors, width=0.56, zorder=2)
    ax1.text(2.55, global_churn + 1.5, f"Global: {fmt_es(global_churn, 1)} %",
             fontsize=8.3, color=HEX_NAVY, fontweight="bold", ha="center",
             bbox=label_pill, zorder=4)
    ax1.set_xticks(x_pos)
    ax1.set_xticklabels(g_t.index, fontsize=9.5, fontweight="bold")
    ax1.set_ylim(0, 58)
    ax1.set_ylabel("Tasa de Churn (%)", fontsize=9.5, color=HEX_MUTED)
    ax1.set_title("Churn por Antigüedad (tenure)", fontsize=11.5,
                  fontweight="bold", color=HEX_NAVY, pad=10)
    ax1.spines["top"].set_visible(False)
    ax1.spines["right"].set_visible(False)
    ax1.grid(axis="y", linestyle="--", alpha=0.32, zorder=0)
    for bar, r, n in zip(bars_t, t_rates, t_counts):
        n_txt = f"{int(n):,}".replace(",", ".")
        ax1.text(
            bar.get_x() + bar.get_width() / 2, r + 1.2,
            f"{fmt_es(r, 1)} %\n(n={n_txt})",
            ha="center", va="bottom", fontsize=8.7, fontweight="bold",
            color=HEX_CORAL if r > 40 else HEX_TEXT,
            bbox=dict(facecolor="white", edgecolor="none", pad=0.8, alpha=0.88), zorder=5
        )

    # Panel Derecho: Internet y Servicios Protectores
    serv_labels = [
        "Internet: Sin servicio",
        "Internet: DSL",
        "Internet: Fibra Óptica",
        "Seguridad Online: SÍ",
        "Seguridad Online: NO",
        "Soporte Técnico: SÍ",
        "Soporte Técnico: NO",
    ]
    serv_vals = [
        df.loc[df["InternetService"] == "No", "Churn_num"].mean() * 100,
        df.loc[df["InternetService"] == "DSL", "Churn_num"].mean() * 100,
        df.loc[df["InternetService"] == "Fiber optic", "Churn_num"].mean() * 100,
        df.loc[df["OnlineSecurity"] == "Yes", "Churn_num"].mean() * 100,
        df.loc[df["OnlineSecurity"] == "No", "Churn_num"].mean() * 100,
        df.loc[df["TechSupport"] == "Yes", "Churn_num"].mean() * 100,
        df.loc[df["TechSupport"] == "No", "Churn_num"].mean() * 100,
    ]
    serv_colors = [
        HEX_TEAL, HEX_SLATE, HEX_CORAL,
        HEX_TEAL, HEX_CORAL,
        HEX_TEAL, HEX_CORAL,
    ]
    y_pos = [0, 0.85, 1.7, 3.0, 3.85, 5.15, 6.0]
    ax2.axvline(global_churn, color=HEX_NAVY, linestyle="--", linewidth=1.15, alpha=0.75, zorder=1)
    bars_s = ax2.barh(y_pos, serv_vals, color=serv_colors, height=0.62, zorder=2)
    ax2.set_yticks(y_pos)
    ax2.set_yticklabels(serv_labels, fontsize=9.2)
    ax2.set_xlim(0, 54)
    ax2.set_xlabel("Tasa de Churn (%)", fontsize=9.5, color=HEX_MUTED)
    ax2.set_title("Impacto de Tecnología y Servicios Protectores", fontsize=11.5,
                  fontweight="bold", color=HEX_NAVY, pad=10)
    ax2.spines["top"].set_visible(False)
    ax2.spines["right"].set_visible(False)
    ax2.grid(axis="x", linestyle="--", alpha=0.32, zorder=0)
    for bar, r in zip(bars_s, serv_vals):
        ax2.text(
            r + 1.0, bar.get_y() + bar.get_height() / 2,
            f"{fmt_es(r, 1)} %", va="center", fontsize=9.3, fontweight="bold",
            color=HEX_CORAL if r > 35 else HEX_TEXT,
            bbox=label_pill, zorder=5
        )

    plt.tight_layout(w_pad=1.8)
    p7 = CHARTS_DIR / "07_servicios_y_antiguedad.png"
    fig.savefig(p7, dpi=300, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    rutas[p7.name] = p7

    # -------------------------------------------------------------------------
    # Gráfico 08: Escalada de Riesgo del Cliente Crítico (Slide 10)
    # -------------------------------------------------------------------------
    s0 = df
    s1 = df[df["Contract"] == "Month-to-month"]
    s2 = s1[s1["InternetService"] == "Fiber optic"]
    s3 = s2[s2["PaymentMethod"] == "Electronic check"]
    s4 = s3[s3["tenure"] <= 12]
    s5 = s4[(s4["OnlineSecurity"] == "No") & (s4["TechSupport"] == "No")]
    low = df[(df["Contract"] == "Two year") & (df["ModalidadCobro"] == "Automático")]

    etapas = [
        "Benchmark bajo riesgo\n(2 años + Débito auto)",
        "Cartera global\n(Todos los clientes)",
        "Contrato mes a mes\n(Month-to-month)",
        "+ Internet Fibra Óptica",
        "+ Pago Cheque Electrónico",
        "+ Antigüedad ≤ 12 meses",
        "+ Sin Seguridad ni Soporte\n(Perfil Crítico)",
    ]
    grupos = [low, s0, s1, s2, s3, s4, s5]
    tasas = [g["Churn_num"].mean() * 100 for g in grupos]
    tamanos = [len(g) for g in grupos]
    colores = [HEX_TEAL, HEX_SLATE, "#4E79A7", "#D9822B", "#E06D2B", "#D94E28", "#B82E1F"]

    fig, ax = plt.subplots(figsize=(7.2, 4.9))
    y = np.arange(len(etapas))
    ax.axvline(global_churn, color=HEX_NAVY, linestyle="--", linewidth=1.2, alpha=0.72, zorder=1)
    bars = ax.barh(y, tasas, color=colores, height=0.62, zorder=2)
    ax.set_yticks(y)
    ax.set_yticklabels(etapas, fontsize=9.3, fontweight="bold")
    ax.set_xlim(0, 98)
    ax.set_xlabel("Probabilidad observada de abandono (% Churn)", fontsize=9.5, color=HEX_MUTED)
    ax.set_title("Escalada del Riesgo por Acumulación de Factores", fontsize=12,
                 fontweight="bold", color=HEX_NAVY, loc="left", pad=10)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.grid(axis="x", linestyle="--", alpha=0.32, zorder=0)
    for bar, r, n in zip(bars, tasas, tamanos):
        n_txt = f"{n:,}".replace(",", ".")
        ax.text(
            r + 1.3, bar.get_y() + bar.get_height() / 2,
            f"{fmt_es(r, 1)} %  (n = {n_txt})",
            va="center", fontsize=9.4, fontweight="bold",
            color="#B82E1F" if r >= 65 else (HEX_TEAL if r < 10 else HEX_TEXT),
            bbox=label_pill, zorder=5
        )

    plt.tight_layout()
    p8 = CHARTS_DIR / "08_escalada_riesgo.png"
    fig.savefig(p8, dpi=300, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    rutas[p8.name] = p8

    return rutas


# =============================================================================
# 3. UTILIDADES DE DISEÑO EN PYTHON-PPTX
# =============================================================================
def set_slide_bg(slide, color: RGBColor) -> None:
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_rect(
    slide, left, top, width, height,
    fill_color: RGBColor,
    border_color: RGBColor | None = None,
    border_pt: float = 0.75,
    rounded: bool = False,
    radius: float = 0.05,
):
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE
    shape = slide.shapes.add_shape(shape_type, left, top, width, height)
    if rounded:
        shape.adjustments[0] = radius
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if border_color is not None:
        shape.line.color.rgb = border_color
        shape.line.width = Pt(border_pt)
    else:
        shape.line.fill.background()
    return shape


def add_textbox(
    slide, left, top, width, height, text: str,
    font_size: float = 13, bold: bool = False, italic: bool = False,
    color: RGBColor = C_TEXT, align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.TOP,
    line_spacing: float = 1.15,
):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = valign
    tf.margin_left = Inches(0.04)
    tf.margin_right = Inches(0.04)
    tf.margin_top = Inches(0.02)
    tf.margin_bottom = Inches(0.02)
    p = tf.paragraphs[0]
    p.alignment = align
    p.line_spacing = line_spacing
    run = p.add_run()
    run.text = text
    run.font.name = FONT_MAIN
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return tb


def add_bullets(
    slide, left, top, width, height,
    items: list[tuple[str, str]],
    font_size: float = 12.5,
    space_after: float = 7,
    bullet_char: str = "▸ ",
    bullet_color: RGBColor = C_CYAN,
    text_color: RGBColor = C_TEXT,
    line_spacing: float = 1.15,
):
    """Agrega lista de viñetas con prefijo en negrita opcional: [(negrita, resto), ...]."""
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.04)
    tf.margin_right = Inches(0.04)
    tf.margin_top = Inches(0.02)
    tf.margin_bottom = Inches(0.02)
    for idx, (lead, body) in enumerate(items):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        p.space_after = Pt(space_after)
        p.line_spacing = line_spacing

        r_b = p.add_run()
        r_b.text = bullet_char
        r_b.font.name = FONT_MAIN
        r_b.font.size = Pt(font_size)
        r_b.font.bold = True
        r_b.font.color.rgb = bullet_color

        if lead:
            r_l = p.add_run()
            r_l.text = lead + " "
            r_l.font.name = FONT_MAIN
            r_l.font.size = Pt(font_size)
            r_l.font.bold = True
            r_l.font.color.rgb = C_NAVY if text_color == C_TEXT else text_color

        r_t = p.add_run()
        r_t.text = body
        r_t.font.name = FONT_MAIN
        r_t.font.size = Pt(font_size)
        r_t.font.color.rgb = text_color
    return tb


def add_header_footer(slide, slide_num: int, kicker: str, title: str, subtitle: str) -> None:
    set_slide_bg(slide, C_BG)
    # Barra superior delgada Navy + segmento Cyan
    add_rect(slide, Inches(0), Inches(0), SLIDE_W, Inches(0.08), C_NAVY)
    add_rect(slide, Inches(0), Inches(0), Inches(3.2), Inches(0.08), C_CYAN)

    # Kicker, Título y Subtítulo
    add_textbox(slide, Inches(0.55), Inches(0.18), Inches(10.5), Inches(0.24),
                kicker.upper(), font_size=9.5, bold=True, color=C_CYAN)
    add_textbox(slide, Inches(0.55), Inches(0.39), Inches(12.2), Inches(0.42),
                title, font_size=23, bold=True, color=C_NAVY)
    add_textbox(slide, Inches(0.55), Inches(0.82), Inches(12.2), Inches(0.30),
                subtitle, font_size=12, color=C_MUTED)

    # Línea separadora sutil
    add_rect(slide, Inches(0.55), Inches(1.16), Inches(12.23), Inches(0.015), C_BORDER)

    # Footer
    add_rect(slide, Inches(0.55), Inches(7.05), Inches(12.23), Inches(0.01), C_BORDER)
    add_textbox(
        slide, Inches(0.55), Inches(7.10), Inches(9.5), Inches(0.28),
        "TPI Introducción al Análisis de Datos (UTN)  ·  Telco Customer Churn  ·  Grupo: Toledo · Sosa · Antuñez · Safatle",
        font_size=8.8, color=C_MUTED
    )
    add_textbox(
        slide, Inches(11.3), Inches(7.10), Inches(1.48), Inches(0.28),
        f"{slide_num:02d} / {TOTAL_SLIDES:02d}",
        font_size=9.5, bold=True, color=C_NAVY, align=PP_ALIGN.RIGHT
    )


def add_kpi_card(
    slide, left, top, width, height,
    value: str, label: str, sublabel: str = "",
    accent: RGBColor = C_CYAN, bg: RGBColor = C_WHITE,
    val_size: float = 28,
):
    add_rect(slide, left, top, width, height, bg, border_color=C_BORDER, rounded=True, radius=0.06)
    add_rect(slide, left, top, width, Inches(0.07), accent)
    add_textbox(slide, left + Inches(0.18), top + Inches(0.14), width - Inches(0.36), Inches(0.52),
                value, font_size=val_size, bold=True, color=accent if accent != C_CYAN else C_NAVY)
    add_textbox(slide, left + Inches(0.18), top + Inches(0.68), width - Inches(0.36), Inches(0.32),
                label, font_size=11.5, bold=True, color=C_TEXT)
    if sublabel:
        add_textbox(slide, left + Inches(0.18), top + Inches(0.98), width - Inches(0.36), Inches(0.35),
                    sublabel, font_size=9.8, color=C_MUTED)


def add_picture_fit(slide, img_path: Path, left, top, max_w, max_h) -> None:
    """Inserta una imagen respetando estrictamente su relación de aspecto."""
    with Image.open(img_path) as im:
        iw, ih = im.size
    box_w = float(max_w)
    box_h = float(max_h)
    scale = min(box_w / iw, box_h / ih)
    final_w = int(iw * scale)
    final_h = int(ih * scale)
    offset_x = int(left + (box_w - final_w) / 2)
    offset_y = int(top + (box_h - final_h) / 2)
    slide.shapes.add_picture(str(img_path), offset_x, offset_y, width=final_w, height=final_h)


def add_styled_table(
    slide, left, top, width, height,
    headers: list[str], rows: list[list[str]],
    col_widths: list[float] | None = None,
    font_size: float = 10.5,
    header_bg: RGBColor = C_NAVY,
    highlight_col: int | None = None,
):
    n_rows = len(rows) + 1
    n_cols = len(headers)
    shape = slide.shapes.add_table(n_rows, n_cols, left, top, width, height)
    tbl = shape.table
    if col_widths:
        for i, w in enumerate(col_widths):
            tbl.columns[i].width = Inches(w)

    # Encabezado
    for c_idx, h in enumerate(headers):
        cell = tbl.cell(0, c_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = header_bg
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER if c_idx > 0 else PP_ALIGN.LEFT
        r = p.add_run()
        r.text = h
        r.font.name = FONT_MAIN
        r.font.size = Pt(font_size)
        r.font.bold = True
        r.font.color.rgb = C_WHITE

    # Filas
    for r_idx, row in enumerate(rows):
        row_bg = C_WHITE if r_idx % 2 == 0 else RGBColor(0xF1, 0xF5, 0xF9)
        for c_idx, val in enumerate(row):
            cell = tbl.cell(r_idx + 1, c_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = row_bg
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER if c_idx > 0 else PP_ALIGN.LEFT
            r = p.add_run()
            r.text = str(val)
            r.font.name = FONT_MAIN
            r.font.size = Pt(font_size)
            is_bold = (c_idx == 0) or (highlight_col is not None and c_idx == highlight_col)
            r.font.bold = is_bold
            r.font.color.rgb = C_CORAL if (highlight_col == c_idx) else (C_NAVY if c_idx == 0 else C_TEXT)
    return shape


def set_notes(slide, text: str) -> None:
    notes_slide = slide.notes_slide
    tf = notes_slide.notes_text_frame
    tf.text = text.strip()


# =============================================================================
# 4. CONSTRUCCIÓN DE LAS 12 DIAPOSITIVAS
# =============================================================================
def build_presentation(nb_charts: dict[str, Path], exec_charts: dict[str, Path]) -> None:
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    blank_layout = prs.slide_layouts[6]

    # =========================================================================
    # SLIDE 1 — PORTADA
    # =========================================================================
    s1 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s1, C_NAVY)

    # Elementos geométricos modernos
    add_rect(s1, Inches(0), Inches(0), Inches(0.22), SLIDE_H, C_CYAN)
    add_rect(s1, Inches(8.85), Inches(0), Inches(4.483), SLIDE_H, C_DARK_CARD)
    add_rect(s1, Inches(8.85), Inches(0), Inches(0.05), SLIDE_H, C_CYAN)

    # Kicker
    add_rect(s1, Inches(0.75), Inches(0.95), Inches(4.35), Inches(0.38), C_DARK_CARD,
             border_color=C_CYAN, border_pt=1.0, rounded=True, radius=0.2)
    add_textbox(s1, Inches(0.90), Inches(1.01), Inches(4.10), Inches(0.28),
                "UTN · INTRODUCCIÓN AL ANÁLISIS DE DATOS", font_size=10.5, bold=True, color=C_CYAN)

    # Título principal
    add_textbox(
        s1, Inches(0.75), Inches(1.60), Inches(7.7), Inches(1.95),
        "Retención de Clientes y\nAnálisis de Fuga (Churn)\nen Telecomunicaciones",
        font_size=36, bold=True, color=C_WHITE, line_spacing=1.12
    )
    add_textbox(
        s1, Inches(0.75), Inches(3.85), Inches(7.5), Inches(0.85),
        "Trabajo Práctico Integrador — Diagnóstico estadístico, patrones de abandono\ny recomendaciones accionables basadas en Telco Customer Churn",
        font_size=14, color=RGBColor(0xC7, 0xD8, 0xE8), line_spacing=1.25
    )

    # Separador y bloque de integrantes
    add_rect(s1, Inches(0.75), Inches(4.92), Inches(6.8), Inches(0.025), C_CYAN)
    add_textbox(s1, Inches(0.75), Inches(5.12), Inches(7.5), Inches(0.28),
                "INTEGRANTES DEL GRUPO", font_size=10, bold=True, color=C_CYAN)
    add_textbox(
        s1, Inches(0.75), Inches(5.42), Inches(7.6), Inches(0.72),
        "Santiago Valentín Toledo  ·  Javier Sosa\nJuan Antuñez  ·  Sebastian Safatle",
        font_size=15, bold=True, color=C_WHITE, line_spacing=1.35
    )
    add_textbox(
        s1, Inches(0.75), Inches(6.48), Inches(7.6), Inches(0.35),
        "Noviembre 2026   |   github.com/Janegro09/tpi-analisis-de-datos   |   Google Colab",
        font_size=10.5, color=RGBColor(0x94, 0xB0, 0xC8)
    )

    # Columna derecha: 3 KPIs de impacto en portada
    add_textbox(s1, Inches(9.28), Inches(0.95), Inches(3.6), Inches(0.35),
                "RADIOGRAFÍA DEL ESTUDIO", font_size=10.5, bold=True, color=C_CYAN)
    portada_kpis = [
        ("7.043", "Clientes analizados", "100 % conservados tras depuración ETL", C_CYAN),
        ("26,54 %", "Tasa global de Churn", "1.869 abonados cancelaron el servicio", C_CORAL),
        ("US$ 1,67 M", "Ingreso anual en riesgo", "US$ 139.131/mes (30,5 % del MRR total)", C_TEAL),
    ]
    for i, (val, lbl, sub, col) in enumerate(portada_kpis):
        y = Inches(1.55 + i * 1.75)
        add_rect(s1, Inches(9.28), y, Inches(3.55), Inches(1.48), C_NAVY,
                 border_color=col, border_pt=1.2, rounded=True, radius=0.08)
        add_textbox(s1, Inches(9.50), y + Inches(0.15), Inches(3.15), Inches(0.55),
                    val, font_size=30, bold=True, color=col)
        add_textbox(s1, Inches(9.50), y + Inches(0.74), Inches(3.15), Inches(0.28),
                    lbl, font_size=12, bold=True, color=C_WHITE)
        add_textbox(s1, Inches(9.50), y + Inches(1.04), Inches(3.15), Inches(0.32),
                    sub, font_size=9.5, color=RGBColor(0xB0, 0xC6, 0xDA))

    set_notes(s1, """[Orador: Santiago Valentín Toledo · 0:30]
Buenas noches a todos. Somos Santiago Toledo, Javier Sosa, Juan Antuñez y Sebastian Safatle, y hoy presentamos nuestro Trabajo Práctico Integrador de Introducción al Análisis de Datos sobre la fuga de clientes —Churn— en una empresa de telecomunicaciones.
En los próximos minutos vamos a recorrer el problema de negocio, la depuración técnica de los 7.043 registros, los hallazgos clave del EDA, la validación por correlaciones y cuatro acciones concretas para revertir la pérdida de ingresos.""")

    # =========================================================================
    # SLIDE 2 — CONTEXTO Y PROBLEMA DE NEGOCIO
    # =========================================================================
    s2 = prs.slides.add_slide(blank_layout)
    add_header_footer(
        s2, 2, "01 · Contexto y Problema de Negocio",
        "¿Por qué el Churn es Crítico en Telecomunicaciones?",
        "La pérdida de abonados erosiona los ingresos recurrentes y encarece la adquisición de nuevos clientes"
    )

    # 3 KPI Cards superiores
    add_kpi_card(s2, Inches(0.55), Inches(1.38), Inches(3.88), Inches(1.42),
                 "US$ 456.117", "Facturación mensual total (MRR)",
                 "Cartera histórica de 7.043 clientes (US$ 16,06 M acumulados)", accent=C_SLATE)
    add_kpi_card(s2, Inches(4.72), Inches(1.38), Inches(3.88), Inches(1.42),
                 "US$ 139.131 / mes", "MRR perdido por Churn (30,5 %)",
                 "1.869 clientes dados de baja (26,54 % del padrón)", accent=C_CORAL, bg=C_SOFT_CORAL)
    add_kpi_card(s2, Inches(8.90), Inches(1.38), Inches(3.88), Inches(1.42),
                 "US$ 1,67 M / año", "Ingreso anualizado en riesgo",
                 "Retener solo el 20 % recuperaría ~US$ 334.000/año", accent=C_TEAL)

    # Panel Izquierdo: Problema de Negocio
    add_rect(s2, Inches(0.55), Inches(3.05), Inches(5.95), Inches(3.75), C_WHITE,
             border_color=C_BORDER, rounded=True, radius=0.04)
    add_rect(s2, Inches(0.55), Inches(3.05), Inches(0.09), Inches(3.75), C_NAVY)
    add_textbox(s2, Inches(0.85), Inches(3.22), Inches(5.4), Inches(0.36),
                "El Problema de Negocio (Fundamentación)", font_size=15, bold=True, color=C_NAVY)
    add_bullets(s2, Inches(0.85), Inches(3.68), Inches(5.4), Inches(2.95), [
        ("Rentabilidad en riesgo:", "La retención de clientes es crítica en telecomunicaciones."),
        ("Doble impacto financiero:", "Se pierden ingresos recurrentes y crece el costo de adquisición."),
        ("Objetivo analítico:", "Identificar qué perfiles, contratos y servicios disparan el abandono."),
        ("Variable objetivo (Churn):", "Clasificación binaria (Yes/No) eje de todo el análisis estadístico."),
    ], font_size=13, space_after=11, bullet_color=C_CYAN)

    # Panel Derecho: La Paradoja Financiera
    add_rect(s2, Inches(6.80), Inches(3.05), Inches(5.98), Inches(3.75), C_WHITE,
             border_color=C_BORDER, rounded=True, radius=0.04)
    add_rect(s2, Inches(6.80), Inches(3.05), Inches(0.09), Inches(3.75), C_CORAL)
    add_textbox(s2, Inches(7.10), Inches(3.22), Inches(5.4), Inches(0.36),
                "La Paradoja Financiera Detectada", font_size=15, bold=True, color=C_CORAL)
    add_bullets(s2, Inches(7.10), Inches(3.68), Inches(5.45), Inches(1.75), [
        ("Pérdida desproporcionada:", "El 26,5 % de clientes genera el 30,5 % de pérdida de MRR."),
        ("Se van los clientes caros:", "Quienes cancelan pagan US$ 79,65/mes vs. US$ 64,43 activos."),
        ("Fuga temprana:", "La mitad abandona antes de cumplir 10 meses de antigüedad."),
    ], font_size=13, space_after=10, bullet_color=C_CORAL)

    # Caja destacada inferior derecha
    add_rect(s2, Inches(7.10), Inches(5.55), Inches(5.38), Inches(1.02), C_SOFT_BLUE,
             border_color=C_CYAN, border_pt=0.9, rounded=True, radius=0.08)
    add_textbox(
        s2, Inches(7.28), Inches(5.66), Inches(5.05), Inches(0.82),
        "Pregunta rectora de negocio:\n¿Qué factores explican el abandono y dónde intervenir primero para maximizar el retorno?",
        font_size=12, bold=True, color=C_NAVY, line_spacing=1.22
    )

    set_notes(s2, """[Orador: Santiago Valentín Toledo · 1:15]
En el mercado de telecomunicaciones, la retención es clave para la rentabilidad porque captar un cliente nuevo resulta mucho más costoso que conservar uno actual.
Al cuantificar el problema sobre nuestra cartera de 7.043 clientes, la empresa factura 456 mil dólares mensuales, pero pierde más de 139 mil dólares por mes debido a las 1.869 bajas —eso representa el 30,5 % de la facturación mensual, o 1,67 millones de dólares anualizados—.
Aquí aparece una paradoja importante: se va el 26,5 % de los abonados, pero se llevan más del 30 % de los ingresos mensuales, porque quienes abandonan tienen tarifas significativamente más altas (mediana de 79,65 dólares contra 64,43). Por eso definimos Churn como nuestra variable objetivo central.""")

    # =========================================================================
    # SLIDE 3 — OBJETIVOS E HIPÓTESIS INICIALES
    # =========================================================================
    s3 = prs.slides.add_slide(blank_layout)
    add_header_footer(
        s3, 3, "02 · Objetivos e Hipótesis",
        "Objetivos del Estudio e Hipótesis Analíticas",
        "Tres hipótesis preliminares formuladas desde el negocio y contrastadas con evidencia estadística"
    )

    # Banner de Objetivo General
    add_rect(s3, Inches(0.55), Inches(1.36), Inches(12.23), Inches(0.92), C_NAVY, rounded=True, radius=0.08)
    add_rect(s3, Inches(0.55), Inches(1.36), Inches(0.12), Inches(0.92), C_CYAN)
    add_textbox(s3, Inches(0.85), Inches(1.44), Inches(2.2), Inches(0.24),
                "OBJETIVO CENTRAL", font_size=9.5, bold=True, color=C_CYAN)
    add_textbox(
        s3, Inches(0.85), Inches(1.69), Inches(11.7), Inches(0.50),
        "Analizar el comportamiento histórico de los clientes (demografía, contratos, tarifas y servicios) para identificar los principales detonantes del Churn y formular recomendaciones estratégicas de retención.",
        font_size=13, color=C_WHITE, line_spacing=1.18
    )

    # 3 Tarjetas de Hipótesis (H1, H2, H3)
    hipotesis = [
        (
            "H1 · TIPO DE CONTRATO",
            "Contratos mes a mes elevan la fuga",
            "Los clientes con contrato Month-to-month presentan una tasa de abandono significativamente mayor que los planes anuales o bianuales.",
            "Variables clave:\nContract vs. Churn",
            "42,71 % vs. 2,83 %",
            "Month-to-month concentra el 88,6 % del total de bajas.",
            C_CORAL,
        ),
        (
            "H2 · TARIFAS Y ANTIGÜEDAD",
            "Alto costo inicial impulsa bajas",
            "Los clientes con cargos mensuales elevados y baja antigüedad muestran una propensión mucho mayor a cancelar el servicio.",
            "Variables clave:\nMonthlyCharges, tenure, RatioCostoAntiguedad",
            "10 m  ·  US$ 79,65",
            "Medianas en bajas vs. 38 meses y US$ 64,43 en activos.",
            C_CYAN,
        ),
        (
            "H3 · SERVICIOS DE VALOR AÑADIDO",
            "Soporte y seguridad retienen",
            "La contratación de servicios adicionales de soporte técnico y seguridad en línea reduce la probabilidad de cancelación.",
            "Variables clave:\nOnlineSecurity, TechSupport vs. Churn",
            "~42 %  →  ~15 %",
            "Contar con estos servicios reduce la fuga a un tercio.",
            C_TEAL,
        ),
    ]
    for i, (tag, htitle, desc, vkey, stat, note, accent) in enumerate(hipotesis):
        x = Inches(0.55 + i * 4.16)
        w = Inches(3.91)
        add_rect(s3, x, Inches(2.55), w, Inches(4.25), C_WHITE, border_color=C_BORDER, rounded=True, radius=0.04)
        add_rect(s3, x, Inches(2.55), w, Inches(0.08), accent)

        # Tag y badge VALIDADA
        add_textbox(s3, x + Inches(0.20), Inches(2.74), Inches(2.35), Inches(0.26),
                    tag, font_size=9.5, bold=True, color=accent)
        add_rect(s3, x + Inches(2.60), Inches(2.72), Inches(1.12), Inches(0.28), C_SOFT_TEAL,
                 border_color=C_TEAL, border_pt=0.8, rounded=True, radius=0.25)
        add_textbox(s3, x + Inches(2.60), Inches(2.74), Inches(1.12), Inches(0.24),
                    "✔ VALIDADA", font_size=8.8, bold=True, color=C_TEAL, align=PP_ALIGN.CENTER)

        add_textbox(s3, x + Inches(0.20), Inches(3.08), w - Inches(0.40), Inches(0.42),
                    htitle, font_size=14.5, bold=True, color=C_NAVY)
        add_textbox(s3, x + Inches(0.20), Inches(3.52), w - Inches(0.40), Inches(1.18),
                    desc, font_size=11.5, color=C_TEXT, line_spacing=1.22)
        add_textbox(s3, x + Inches(0.20), Inches(4.74), w - Inches(0.40), Inches(0.54),
                    vkey, font_size=10, italic=True, color=C_MUTED, line_spacing=1.18)

        # Caja de evidencia al pie
        add_rect(s3, x + Inches(0.18), Inches(5.38), w - Inches(0.36), Inches(1.22), C_BG,
                 border_color=C_BORDER, rounded=True, radius=0.08)
        add_textbox(s3, x + Inches(0.28), Inches(5.46), w - Inches(0.56), Inches(0.44),
                    stat, font_size=19, bold=True, color=accent)
        add_textbox(s3, x + Inches(0.28), Inches(5.96), w - Inches(0.56), Inches(0.55),
                    note, font_size=10.2, color=C_TEXT, line_spacing=1.15)

    set_notes(s3, """[Orador: Santiago Valentín Toledo · 1:00]
Antes de tocar los datos definimos tres hipótesis concretas en nuestra fundamentación de negocio.
H1 postula que los contratos mes a mes sufren mucha más fuga que los anuales o bianuales. H2 plantea que los cargos mensuales elevados combinados con baja antigüedad aumentan la probabilidad de abandono. Y H3 sostiene que los servicios de valor añadido —como seguridad en línea y soporte técnico— actúan como factores protectores.
Como adelanta el tablero, las tres hipótesis fueron validadas estadísticamente durante el EDA y la correlación. Ahora le doy la palabra a Javier para presentar la arquitectura del dataset y la limpieza de datos.""")

    # =========================================================================
    # SLIDE 4 — ORIGEN Y ARQUITECTURA DEL DATASET
    # =========================================================================
    s4 = prs.slides.add_slide(blank_layout)
    add_header_footer(
        s4, 4, "03 · Datos y Metodología",
        "Origen del Dataset y Arquitectura Reproducible",
        "Datos estructurados tabulares (Telco Customer Churn) bajo flujo colaborativo versionado"
    )

    # Columna Izquierda: Arquitectura del Dataset
    add_rect(s4, Inches(0.55), Inches(1.36), Inches(6.20), Inches(5.46), C_WHITE,
             border_color=C_BORDER, rounded=True, radius=0.04)
    add_textbox(s4, Inches(0.80), Inches(1.52), Inches(5.7), Inches(0.36),
                "Dataset Original: 7.043 filas × 21 columnas", font_size=15, bold=True, color=C_NAVY)
    add_textbox(s4, Inches(0.80), Inches(1.88), Inches(5.7), Inches(0.32),
                "WA_Fn-UseC_-Telco-Customer-Churn.csv  ·  1 fila = 1 cliente único",
                font_size=10.5, color=C_MUTED)

    grupos_vars = [
        ("Sociodemográficas (4)", "gender · SeniorCitizen · Partner · Dependents",
         "Cualitativas nominales y binaria (0/1)", C_SLATE),
        ("Servicios Contratados (9)", "PhoneService · MultipleLines · InternetService · OnlineSecurity · OnlineBackup · DeviceProtection · TechSupport · StreamingTV · StreamingMovies",
         "Cualitativas nominales", C_CYAN),
        ("Contrato y Permanencia (4)", "tenure (meses) · Contract · PaperlessBilling · PaymentMethod",
         "Cuantitativa discreta + cualitativa ordinal/nominal", C_TEAL),
        ("Facturación y Objetivo (4)", "customerID · MonthlyCharges · TotalCharges · Churn",
         "ID único + cuantitativas continuas + Target binario", C_CORAL),
    ]
    for idx, (gname, cols, tipo, col) in enumerate(grupos_vars):
        y = Inches(2.30 + idx * 1.08)
        add_rect(s4, Inches(0.80), y, Inches(5.70), Inches(0.96), C_BG,
                 border_color=C_BORDER, rounded=True, radius=0.06)
        add_rect(s4, Inches(0.80), y, Inches(0.08), Inches(0.96), col)
        add_textbox(s4, Inches(1.00), y + Inches(0.06), Inches(3.1), Inches(0.25),
                    gname, font_size=11.2, bold=True, color=C_NAVY)
        add_textbox(s4, Inches(3.75), y + Inches(0.07), Inches(2.65), Inches(0.24),
                    tipo, font_size=8.8, italic=True, color=col, align=PP_ALIGN.RIGHT)
        add_textbox(s4, Inches(1.00), y + Inches(0.34), Inches(5.35), Inches(0.56),
                    cols, font_size=9.8, color=C_TEXT, line_spacing=1.14)

    # Columna Derecha: Estructura del Repositorio y Metodología
    add_rect(s4, Inches(7.00), Inches(1.36), Inches(5.78), Inches(2.85), C_WHITE,
             border_color=C_BORDER, rounded=True, radius=0.04)
    add_textbox(s4, Inches(7.25), Inches(1.52), Inches(5.3), Inches(0.34),
                "Estructura Reproducible del Repositorio", font_size=15, bold=True, color=C_NAVY)

    # Pipeline visual de carpetas
    pasos_repo = [
        ("data/raw/", "CSV crudo\nintacto", C_SLATE),
        ("notebooks/", "TPI Colab /\nJupyter", C_CYAN),
        ("data/processed/", "CSV limpio\n7.043×27", C_TEAL),
        ("docs/", "Negocio &\nGuion", C_NAVY),
    ]
    for i, (folder, desc, col) in enumerate(pasos_repo):
        px = Inches(7.25 + i * 1.35)
        add_rect(s4, px, Inches(2.02), Inches(1.22), Inches(1.05), col, rounded=True, radius=0.08)
        add_textbox(s4, px + Inches(0.05), Inches(2.12), Inches(1.12), Inches(0.30),
                    folder, font_size=8.8, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
        add_textbox(s4, px + Inches(0.05), Inches(2.44), Inches(1.12), Inches(0.52),
                    desc, font_size=8.5, color=RGBColor(0xE2, 0xEE, 0xF8), align=PP_ALIGN.CENTER)
        if i < 3:
            add_textbox(s4, px + Inches(1.20), Inches(2.35), Inches(0.18), Inches(0.30),
                        "→", font_size=12, bold=True, color=C_MUTED, align=PP_ALIGN.CENTER)

    add_bullets(s4, Inches(7.25), Inches(3.22), Inches(5.30), Inches(0.90), [
        ("Carga dual reproducible:", "Lee ruta local o URL raw de GitHub en Google Colab."),
        ("Stack Python:", "Pandas · NumPy · Seaborn · Matplotlib · SciPy · Colab."),
    ], font_size=11.2, space_after=5, bullet_color=C_CYAN)

    # Tarjeta inferior derecha: Metodología y Roles
    add_rect(s4, Inches(7.00), Inches(4.42), Inches(5.78), Inches(2.40), C_WHITE,
             border_color=C_BORDER, rounded=True, radius=0.04)
    add_rect(s4, Inches(7.00), Inches(4.42), Inches(5.78), Inches(0.07), C_CYAN)
    add_textbox(s4, Inches(7.25), Inches(4.58), Inches(5.3), Inches(0.32),
                "Metodología Colaborativa (Planificación TPI)", font_size=14, bold=True, color=C_NAVY)
    add_bullets(s4, Inches(7.25), Inches(4.96), Inches(5.35), Inches(1.78), [
        ("Santiago Toledo:", "README, fundamentación de negocio y clasificación de variables."),
        ("Javier Sosa:", "Setup del repositorio, limpieza ETL y exportación a processed/."),
        ("Juan Antuñez:", "Transformaciones, análisis IQR/outliers y EDA descriptivo."),
        ("Sebastian Safatle:", "Matriz/heatmap de correlación y auditoría en Colab."),
    ], font_size=11.2, space_after=6, bullet_color=C_SLATE)

    set_notes(s4, """[Orador: Javier Sosa · 1:30]
Trabajamos con el conjunto de datos estructurados Telco Customer Churn: 7.043 filas —un cliente único por registro— y 21 atributos originales agrupados en cuatro bloques: variables demográficas, nueve servicios contratados, modalidad contractual y antigüedad, y facturación junto con el target Churn.
Desde el inicio organizamos el proyecto en GitHub con una arquitectura reproducible: el archivo crudo vive en data/raw/ y jamás se sobrescribe; todo el flujo corre en un único notebook que detecta automáticamente si se ejecuta en local o en Google Colab y descarga el CSV desde GitHub si no existe localmente. Además, dividimos la planificación en roles bien delimitados que luego se integraron en el pipeline final.""")

    # =========================================================================
    # SLIDE 5 — DATA WRANGLING Y LIMPIEZA (ETL)
    # =========================================================================
    s5 = prs.slides.add_slide(blank_layout)
    add_header_footer(
        s5, 5, "04 · Data Wrangling y Calidad de Datos",
        "Limpieza ETL: de data/raw/ a telco_churn_clean.csv",
        "Depuración metodológica justificada, tipado formal, ingeniería de atributos y control de outliers"
    )

    # Flujo ETL horizontal superior (4 pasos)
    etl_steps = [
        ("1. DETECCIÓN", "11 espacios ' ' en TotalCharges\n(ocultos a isnull())", C_CORAL),
        ("2. IMPUTACIÓN", "pd.to_numeric → float64\nfillna(0) por tenure = 0", C_CYAN),
        ("3. INGENIERÍA", "+6 variables analíticas\n(Churn_num, Ratio, etc.)", C_SLATE),
        ("4. TIPADO E IQR", "category + Contract ordinal\n0 outliers globales", C_TEAL),
    ]
    for i, (stitle, sdesc, col) in enumerate(etl_steps):
        x = Inches(0.55 + i * 3.12)
        add_rect(s5, x, Inches(1.36), Inches(2.88), Inches(1.16), C_WHITE,
                 border_color=col, border_pt=1.3, rounded=True, radius=0.08)
        add_rect(s5, x, Inches(1.36), Inches(2.88), Inches(0.34), col)
        add_textbox(s5, x + Inches(0.12), Inches(1.40), Inches(2.64), Inches(0.26),
                    stitle, font_size=10.5, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
        add_textbox(s5, x + Inches(0.12), Inches(1.78), Inches(2.64), Inches(0.68),
                    sdesc, font_size=10.5, color=C_TEXT, align=PP_ALIGN.CENTER, line_spacing=1.2)
        if i < 3:
            add_textbox(s5, x + Inches(2.89), Inches(1.74), Inches(0.22), Inches(0.36),
                        "→", font_size=16, bold=True, color=C_NAVY, align=PP_ALIGN.CENTER)

    # Tabla comparativa Antes vs Después (izquierda)
    add_textbox(s5, Inches(0.55), Inches(2.74), Inches(7.2), Inches(0.32),
                "Auditoría de Calidad: Dataset Crudo vs. Dataset Depurado",
                font_size=13.5, bold=True, color=C_NAVY)
    add_styled_table(
        s5, Inches(0.55), Inches(3.10), Inches(7.35), Inches(3.72),
        headers=["Dimensión de Calidad", "Raw (Original)", "Processed (Limpio)"],
        rows=[
            ["Dimensiones (filas × cols)", "7.043 × 21", "7.043 × 27 (+6 derivadas)"],
            ["TotalCharges (dtype)", "object / texto", "float64 (numérica)"],
            ["Espacios vacíos / Nulos", "11 strings ' ' (0,16 %)", "0 nulos (imputado con 0.0)"],
            ["Duplicados (filas / ID)", "0 / 0", "0 / 0 (verificado con assert)"],
            ["Categóricas y Contract", "17 columnas object", "category + Contract ordinal"],
            ["Outliers univariados (IQR)", "Sin auditar", "0 globales (100 % retenido)"],
        ],
        col_widths=[2.55, 2.15, 2.65],
        font_size=10.5,
    )

    # Columna derecha: Decisiones metodológicas + KPI de memoria
    add_rect(s5, Inches(8.15), Inches(2.74), Inches(4.63), Inches(2.68), C_WHITE,
             border_color=C_BORDER, rounded=True, radius=0.04)
    add_rect(s5, Inches(8.15), Inches(2.74), Inches(0.09), Inches(2.68), C_CYAN)
    add_textbox(s5, Inches(8.40), Inches(2.88), Inches(4.25), Inches(0.32),
                "Decisiones Metodológicas Clave", font_size=13.5, bold=True, color=C_NAVY)
    add_bullets(s5, Inches(8.40), Inches(3.26), Inches(4.25), Inches(2.08), [
        ("Patrón no aleatorio:", "Las 11 celdas vacías eran el 100 % de clientes con tenure = 0."),
        ("Imputación con 0:", "Con 0 meses aún no facturaron; usar la media sería erróneo."),
        ("Nuevas variables:", "Churn_num, RatioCostoAntiguedad, ModalidadCobro y Addons."),
        ("Sin recorte de filas:", "Los puntos altos en Churn=Yes son bajas reales premium."),
    ], font_size=11.0, space_after=6, bullet_color=C_CYAN)

    # Tarjeta optimización de memoria
    add_rect(s5, Inches(8.15), Inches(5.60), Inches(4.63), Inches(1.22), C_NAVY, rounded=True, radius=0.06)
    add_textbox(s5, Inches(8.40), Inches(5.70), Inches(1.95), Inches(0.55),
                "−86,1 %", font_size=28, bold=True, color=C_CYAN)
    add_textbox(s5, Inches(10.35), Inches(5.76), Inches(2.25), Inches(0.45),
                "Reducción de memoria\npor tipado formal", font_size=11, bold=True, color=C_WHITE)
    add_textbox(s5, Inches(8.40), Inches(6.35), Inches(4.20), Inches(0.32),
                "De 7.375,5 KB → 1.022,4 KB en Pandas  ·  Exportado a data/processed/",
                font_size=9.5, color=RGBColor(0xB8, 0xCE, 0xE3))

    set_notes(s5, """[Orador: Javier Sosa · 1:30]
En la depuración detectamos un caso muy instructivo: df.info() indicaba cero nulos, pero TotalCharges figuraba como texto. Con .str.strip() == "" encontramos 11 cadenas con espacios en blanco.
Antes de borrarlas o imputar a ciegas, investigamos el patrón de no respuesta: los 11 casos coincidían exactamente con los 11 clientes de tenure igual a cero, recién dados de alta sin factura emitida. Por eso convertimos la columna con pd.to_numeric(errors='coerce') a float64 e imputamos con 0 —usar la media les habría asignado más de 2.200 dólares que nunca pagaron—.
Luego confirmamos cero duplicados, agregamos Churn_num para correlaciones y cinco columnas calculadas de negocio (como RatioCostoAntiguedad y ModalidadCobro), optimizamos tipos a category y Contract ordinal —reduciendo un 86,1 % la memoria— y validamos por IQR que no existen outliers globales. Así obtuvimos telco_churn_clean.csv con 7.043 filas y 27 columnas sin perder clientes. Le paso la palabra a Juan.""")

    # =========================================================================
    # SLIDE 6 — EDA: PANORAMA GENERAL
    # =========================================================================
    s6 = prs.slides.add_slide(blank_layout)
    add_header_footer(
        s6, 6, "05 · Análisis Exploratorio de Datos (EDA)",
        "Panorama General: Tasa de Churn y Distribuciones",
        "Uno de cada cuatro clientes abandona; las bajas se concentran en baja antigüedad y tarifas altas"
    )

    # Panel Izquierdo Superior: Donut + MRR chart
    add_rect(s6, Inches(0.55), Inches(1.34), Inches(5.15), Inches(2.85), C_WHITE,
             border_color=C_BORDER, rounded=True, radius=0.04)
    add_picture_fit(s6, exec_charts["05_churn_overview_donut.png"],
                    Inches(0.65), Inches(1.42), Inches(4.95), Inches(2.70))

    # Panel Izquierdo Inferior: Tabla de Medianas (Activos vs Bajas)
    add_textbox(s6, Inches(0.55), Inches(4.32), Inches(5.15), Inches(0.30),
                "Comparativa de Medianas (Activos vs. Bajas)", font_size=12.5, bold=True, color=C_NAVY)
    add_styled_table(
        s6, Inches(0.55), Inches(4.66), Inches(5.15), Inches(2.15),
        headers=["Variable Numérica", "Activos (No)", "Bajas (Yes)", "Brecha"],
        rows=[
            ["tenure (meses)", "38,0 m", "10,0 m", "−73,7 %"],
            ["MonthlyCharges (USD)", "US$ 64,43", "US$ 79,65", "+23,6 %"],
            ["TotalCharges (USD)", "US$ 1.679,52", "US$ 703,55", "−58,1 %"],
        ],
        col_widths=[1.85, 1.15, 1.15, 1.00],
        font_size=10.2,
        highlight_col=2,
    )

    # Panel Derecho: Gráfico REAL del Notebook (Boxplots globales y por Churn)
    add_rect(s6, Inches(5.92), Inches(1.34), Inches(6.86), Inches(5.47), C_WHITE,
             border_color=C_BORDER, rounded=True, radius=0.04)
    add_textbox(s6, Inches(6.12), Inches(1.46), Inches(6.45), Inches(0.28),
                "Gráfico del Notebook (Sección 3.2) — Boxplots Globales y por Churn",
                font_size=11.5, bold=True, color=C_NAVY)
    add_picture_fit(s6, nb_charts["01_nb_boxplots_outliers.png"],
                    Inches(6.02), Inches(1.80), Inches(6.66), Inches(3.60))

    # Nota analítica debajo de los boxplots
    add_rect(s6, Inches(6.12), Inches(5.52), Inches(6.46), Inches(1.12), C_SOFT_BLUE,
             border_color=C_CYAN, border_pt=0.8, rounded=True, radius=0.06)
    add_bullets(s6, Inches(6.24), Inches(5.60), Inches(6.22), Inches(0.98), [
        ("Sin outliers globales (IQR):", "Máximos coherentes (72 m, US$ 118,75, US$ 8.684,80)."),
        ("Asimetría en TotalCharges (+0,96):", "Mediana e IQR son más representativos que la media."),
        ("Puntos altos en Churn=Yes:", "Clientes antiguos premium que se fugaron (alto impacto)."),
    ], font_size=10.3, space_after=3, bullet_color=C_NAVY)

    set_notes(s6, """[Orador: Juan Antuñez · 1:30]
Iniciando el Análisis Exploratorio, vemos que el 26,54 % de la cartera —1.869 clientes— abandonó el servicio, arrastrando el 30,5 % de los ingresos mensuales.
A la derecha mostramos los boxplots reales generados en la Sección 3.2 de nuestro notebook. En la fila superior se confirma visualmente la ausencia de outliers univariados globales según el rango intercuartílico. En la fila inferior, al abrir por Churn, el contraste es nítido: la mediana de antigüedad de quienes cancelan cae a 10 meses frente a 38 meses en los clientes retenidos, mientras que su factura mensual mediana trepa a 79,65 dólares frente a 64,43. Además, como TotalCharges tiene fuerte asimetría positiva (0,96), elegimos mediana e IQR para evitar sesgos de la media.""")

    # =========================================================================
    # SLIDE 7 — HALLAZGOS CLAVE 1: CONTRATOS Y FINANZAS
    # =========================================================================
    s7 = prs.slides.add_slide(blank_layout)
    add_header_footer(
        s7, 7, "06 · Hallazgos Clave 1",
        "Factores Contractuales, Medios de Pago y Facturación",
        "El contrato mes a mes y el cheque electrónico concentran la gran mayoría de la fuga"
    )

    # Panel Izquierdo: Gráfico dual Contratos y Pagos
    add_rect(s7, Inches(0.55), Inches(1.34), Inches(7.85), Inches(5.48), C_WHITE,
             border_color=C_BORDER, rounded=True, radius=0.04)
    add_picture_fit(s7, exec_charts["06_contratos_y_pagos.png"],
                    Inches(0.68), Inches(1.44), Inches(7.58), Inches(5.28))

    # Columna Derecha: 3 Tarjetas de Hallazgos
    insights_7 = [
        (
            "88,6 % DEL CHURN",
            "El contrato mensual es el foco crítico",
            [
                ("Month-to-month:", "42,71 % de fuga (1.655 bajas) y 18 m de vida media."),
                ("Two year retiene:", "97,17 % activos y US$ 6,28 M facturados."),
            ],
            C_CORAL, C_SOFT_CORAL
        ),
        (
            "FRICCIÓN DE PAGO",
            "Cheque electrónico triplica el riesgo",
            [
                ("Electronic check:", "45,29 % de churn (57,3 % del total de bajas)."),
                ("Débito automático:", "Baja a 15,98 % (vs. 34,67 % pagos manuales)."),
            ],
            C_CYAN, C_WHITE
        ),
        (
            "PRESIÓN TARIFARIA",
            "Facturas > US$ 70 disparan bajas",
            [
                ("Rango US$ 70–90:", "Tasa de fuga alcanza 37,8 % (vs. 10,9 % < US$ 35)."),
                ("PaperlessBilling:", "33,57 % de churn en factura digital vs. 16,33 %."),
            ],
            C_SLATE, C_WHITE
        ),
    ]
    for idx, (kicker, title, blts, accent, bg) in enumerate(insights_7):
        y = Inches(1.34 + idx * 1.88)
        add_rect(s7, Inches(8.62), y, Inches(4.16), Inches(1.72), bg,
                 border_color=C_BORDER, rounded=True, radius=0.05)
        add_rect(s7, Inches(8.62), y, Inches(0.09), Inches(1.72), accent)
        add_textbox(s7, Inches(8.84), y + Inches(0.10), Inches(3.8), Inches(0.24),
                    kicker, font_size=9.5, bold=True, color=accent)
        add_textbox(s7, Inches(8.84), y + Inches(0.32), Inches(3.8), Inches(0.30),
                    title, font_size=12.5, bold=True, color=C_NAVY)
        add_bullets(s7, Inches(8.84), y + Inches(0.68), Inches(3.82), Inches(0.98),
                    blts, font_size=10.6, space_after=4, bullet_color=accent)

    set_notes(s7, """[Orador: Juan Antuñez · 1:30]
Al segmentar mediante groupby y tablas de contingencia, encontramos que dos decisiones comerciales explican casi todo el abandono.
Primero, el tipo de contrato: los clientes mes a mes aportan el mayor ingreso mensual (257 mil dólares), pero presentan un 42,7 % de fuga y concentran el 88,6 % de todas las cancelaciones. En cambio, los contratos a dos años retienen al 97,2 % y acumulan 6,28 millones de dólares históricos.
Segundo, el método de pago: quienes usan cheque electrónico alcanzan un 45,3 % de abandono —57 % del churn total—, frente a solo 15–16 % en tarjeta o transferencia automática. El cobro manual genera fricción cada mes, y ese efecto se agrava en facturas superiores a 70 dólares.""")

    # =========================================================================
    # SLIDE 8 — HALLAZGOS CLAVE 2: SERVICIOS Y ANTIGÜEDAD
    # =========================================================================
    s8 = prs.slides.add_slide(blank_layout)
    add_header_footer(
        s8, 8, "07 · Hallazgos Clave 2",
        "Impacto de la Antigüedad (Tenure) y Servicios de Valor",
        "El primer año y la fibra óptica sin servicios protectores concentran el mayor riesgo operativo"
    )

    # Panel Izquierdo: Gráfico dual Tenure + Servicios
    add_rect(s8, Inches(0.55), Inches(1.34), Inches(8.05), Inches(5.48), C_WHITE,
             border_color=C_BORDER, rounded=True, radius=0.04)
    add_picture_fit(s8, exec_charts["07_servicios_y_antiguedad.png"],
                    Inches(0.65), Inches(1.44), Inches(7.85), Inches(5.28))

    # Columna Derecha: 3 Tarjetas de Hallazgos
    insights_8 = [
        (
            "BARRERA DE LOS 12 MESES",
            "55,5 % de bajas ocurre el primer año",
            [
                ("0–12 meses:", "47,4 % de churn (1.037 bajas); moda global = 1 mes."),
                ("Mes a mes ≤ 12 m:", "Sube a 51,35 % (1.994 clientes)."),
            ],
            C_CORAL, C_SOFT_CORAL
        ),
        (
            "ALERTA EN FIBRA ÓPTICA",
            "Producto premium con fuga elevada",
            [
                ("Fiber optic:", "41,89 % de churn (vs. 18,96 % en DSL)."),
                ("Fibra + Mensual:", "54,61 % de abandono (US$ 87,02/mes promedio)."),
            ],
            C_AMBER, C_WHITE
        ),
        (
            "ESCUDO DE RETENCIÓN (H3)",
            "Seguridad y soporte bajan churn a 1/3",
            [
                ("OnlineSecurity:", "41,8 % sin servicio vs. 14,6 % con servicio."),
                ("TechSupport:", "41,6 % sin soporte vs. 15,2 % con soporte."),
            ],
            C_TEAL, C_SOFT_TEAL
        ),
    ]
    for idx, (kicker, title, blts, accent, bg) in enumerate(insights_8):
        y = Inches(1.34 + idx * 1.88)
        add_rect(s8, Inches(8.80), y, Inches(3.98), Inches(1.72), bg,
                 border_color=C_BORDER, rounded=True, radius=0.05)
        add_rect(s8, Inches(8.80), y, Inches(0.09), Inches(1.72), accent)
        add_textbox(s8, Inches(9.02), y + Inches(0.10), Inches(3.65), Inches(0.24),
                    kicker, font_size=9.3, bold=True, color=accent)
        add_textbox(s8, Inches(9.02), y + Inches(0.32), Inches(3.65), Inches(0.30),
                    title, font_size=12.2, bold=True, color=C_NAVY)
        add_bullets(s8, Inches(9.02), y + Inches(0.68), Inches(3.65), Inches(0.98),
                    blts, font_size=10.5, space_after=4, bullet_color=accent)

    set_notes(s8, """[Orador: Juan Antuñez · 1:30]
El segundo bloque de hallazgos muestra cuándo ocurre la fuga y qué rol juegan los servicios.
En el panel izquierdo vemos una caída abrupta del riesgo con el paso del tiempo: el 47,4 % de los clientes con hasta 12 meses abandona —más de la mitad de todas las bajas—, mientras que después de 4 años la tasa cae al 9,5 %. Si combinamos contrato mensual con el primer año, la fuga trepa al 51,35 %.
En el panel derecho destaca la alerta de Fibra Óptica: pese a ser un servicio de alto ingreso medio, pierde al 41,9 % de sus usuarios y al 54,6 % en plan mensual. Pero también aparece la solución: quienes contratan Seguridad Online o Soporte Técnico bajan su tasa de abandono del 42 % al 15 %, validando H3. Ahora Sebastian presentará la matriz de correlación.""")

    # =========================================================================
    # SLIDE 9 — CORRELACIÓN LINEAL Y HEATMAP (SECCIÓN 6 DEL NOTEBOOK)
    # =========================================================================
    s9 = prs.slides.add_slide(blank_layout)
    add_header_footer(
        s9, 9, "08 · Evidencia Estadística (Sección 6)",
        "Matriz de Correlación, Heatmap y Control de Robustez",
        "Cuantificación con Pearson (r), significancia con SciPy (p < 0,001) y validación con Spearman (ρ)"
    )

    # Panel Izquierdo: Heatmap REAL exportado de la Sección 6.3 del Notebook
    add_rect(s9, Inches(0.55), Inches(1.34), Inches(6.15), Inches(5.48), C_WHITE,
             border_color=C_BORDER, rounded=True, radius=0.04)
    add_textbox(s9, Inches(0.75), Inches(1.46), Inches(5.8), Inches(0.28),
                "Gráfico del Notebook (Sección 6.3) — Heatmap Ampliado de Pearson",
                font_size=11.5, bold=True, color=C_NAVY)
    add_picture_fit(s9, nb_charts["03_nb_heatmap_ampliado.png"],
                    Inches(0.68), Inches(1.80), Inches(5.88), Inches(4.90))

    # Panel Derecho Superior: Tabla Pearson vs Spearman
    add_textbox(s9, Inches(6.92), Inches(1.36), Inches(5.85), Inches(0.30),
                "Significancia y Robustez (n = 7.043 · todos p < 0,001)",
                font_size=12.5, bold=True, color=C_NAVY)
    add_styled_table(
        s9, Inches(6.92), Inches(1.70), Inches(5.86), Inches(2.05),
        headers=["Par de Variables", "Pearson (r)", "R²", "Spearman (ρ)", "Fuerza"],
        rows=[
            ["tenure – TotalCharges", "+0,826", "68,3 %", "+0,890", "Fuerte"],
            ["Monthly – TotalCharges", "+0,651", "42,4 %", "+0,638", "Moderada"],
            ["tenure – MonthlyCharges", "+0,248", "6,1 %", "+0,276", "Débil"],
            ["RatioCostoAntig. – Churn", "+0,412", "17,0 %", "—", "Mayor vs Churn"],
        ],
        col_widths=[1.96, 0.96, 0.78, 1.06, 1.10],
        font_size=9.8,
        highlight_col=1,
    )

    # Panel Derecho Inferior: Interpretación de negocio y rigor
    add_rect(s9, Inches(6.92), Inches(3.96), Inches(5.86), Inches(2.86), C_WHITE,
             border_color=C_BORDER, rounded=True, radius=0.04)
    add_rect(s9, Inches(6.92), Inches(3.96), Inches(0.09), Inches(2.86), C_NAVY)
    add_textbox(s9, Inches(7.16), Inches(4.10), Inches(5.45), Inches(0.32),
                "Implicancias Analíticas y de Negocio", font_size=13.5, bold=True, color=C_NAVY)
    add_bullets(s9, Inches(7.16), Inches(4.48), Inches(5.48), Inches(2.24), [
        ("Retener > subir precios:", "tenure explica 68 % del ingreso total (vs. 42 % la tarifa)."),
        ("Mejor predictor numérico:", "RatioCostoAntiguedad (+0,41) supera a tenure (−0,35)."),
        ("Cantidad ≠ protección:", "TotalServiciosAddons (−0,09) prueba que importa qué servicio."),
        ("Robustez y cautela:", "Spearman confirma hallazgos; correlación no implica causalidad."),
        ("Alerta para ML:", "Multicolinealidad entre TotalCharges y tenure × MonthlyCharges."),
    ], font_size=10.9, space_after=5, bullet_color=C_CYAN)

    set_notes(s9, """[Orador: Sebastian Safatle · 1:30]
Para consolidar el análisis con rigor estadístico calculamos la matriz de correlación de Pearson y verificamos la significancia con scipy.stats.pearsonr: todos los pares arrojaron p-valores menores a 0,001.
En el heatmap ampliado de la Sección 6.3 —que incluye Churn_num y las columnas calculadas— surgen tres conclusiones fuertes. Primero, tenure y TotalCharges tienen una correlación fuerte de 0,83 (68 % de varianza compartida), superando a MonthlyCharges (0,65): el valor de vida del cliente depende más de cuánto tiempo se queda que de cobrarle caro. Segundo, RatioCostoAntiguedad logra la asociación más alta con el churn (+0,41), sintetizando la presión de tarifas altas en clientes nuevos. Tercero, validamos con Spearman para controlar asimetrías y las magnitudes no cambiaron. También destacamos dos recaudos clave: correlación no prueba causalidad, y en un futuro modelo de Machine Learning habrá que tratar la multicolinealidad de TotalCharges.""")

    # =========================================================================
    # SLIDE 10 — PERFIL DEL CLIENTE EN RIESGO
    # =========================================================================
    s10 = prs.slides.add_slide(blank_layout)
    add_header_footer(
        s10, 10, "09 · Segmentación de Riesgo",
        "Radiografía del Cliente en Alto Riesgo vs. Cliente Fiel",
        "La superposición de factores eleva la probabilidad observada de abandono del 26,5 % al 74,4 %"
    )

    # Panel Izquierdo: Gráfico de Escalada de Riesgo
    add_rect(s10, Inches(0.55), Inches(1.34), Inches(7.35), Inches(5.48), C_WHITE,
             border_color=C_BORDER, rounded=True, radius=0.04)
    add_picture_fit(s10, exec_charts["08_escalada_riesgo.png"],
                    Inches(0.66), Inches(1.44), Inches(7.13), Inches(5.28))

    # Derecha Arriba: Tarjeta Perfil Crítico
    add_rect(s10, Inches(8.12), Inches(1.34), Inches(4.66), Inches(3.25), C_SOFT_CORAL,
             border_color=C_CORAL, border_pt=1.3, rounded=True, radius=0.05)
    add_rect(s10, Inches(8.12), Inches(1.34), Inches(4.66), Inches(0.42), C_CORAL)
    add_textbox(s10, Inches(8.30), Inches(1.41), Inches(4.30), Inches(0.28),
                "PERFIL CRÍTICO  ·  HASTA 74,4 % DE CHURN", font_size=11, bold=True, color=C_WHITE)
    add_bullets(s10, Inches(8.30), Inches(1.88), Inches(4.32), Inches(2.58), [
        ("Contrato:", "Mes a mes (Month-to-month)."),
        ("Servicio:", "Internet por Fibra Óptica sin Seguridad ni Soporte."),
        ("Pago:", "Cheque electrónico manual (US$ 68k/mes perdidos con Fibra+M2M)."),
        ("Ciclo y tarifa:", "Primeros 12 meses de vida y factura > US$ 70/mes."),
        ("Demografía:", "Adulto mayor (41,7 % churn) o sin dependientes (31,3 %)."),
    ], font_size=11.0, space_after=6, bullet_color=C_CORAL)

    # Derecha Abajo: Tarjeta Perfil Retenido
    add_rect(s10, Inches(8.12), Inches(4.78), Inches(4.66), Inches(2.04), C_SOFT_TEAL,
             border_color=C_TEAL, border_pt=1.3, rounded=True, radius=0.05)
    add_rect(s10, Inches(8.12), Inches(4.78), Inches(4.66), Inches(0.40), C_TEAL)
    add_textbox(s10, Inches(8.30), Inches(4.84), Inches(4.30), Inches(0.28),
                "PERFIL FIEL  ·  SOLO 2,8 % DE CHURN", font_size=11, bold=True, color=C_WHITE)
    add_bullets(s10, Inches(8.30), Inches(5.30), Inches(4.32), Inches(1.40), [
        ("Compromiso:", "Contrato a 2 años + débito automático."),
        ("Protección:", "Soporte técnico y seguridad online activos."),
        ("Estabilidad:", "Antigüedad > 3 años; género es totalmente neutro (~26,5 %)."),
    ], font_size=11.0, space_after=5, bullet_color=C_TEAL)

    set_notes(s10, """[Orador: Sebastian Safatle · 1:00]
Al cruzar todos los hallazgos podemos perfilar exactamente quién se está yendo.
El gráfico de escalada muestra cómo se acumula el riesgo: desde el 26,5 % global pasamos a 42,7 % en contrato mensual, a 54,6 % al sumar Fibra Óptica, a 60,4 % con cheque electrónico —un segmento que por sí solo representa más de 68 mil dólares perdidos cada mes—, a 71,2 % en el primer año y hasta un 74,4 % si además carece de seguridad y soporte (405 bajas en 544 clientes).
En lo sociodemográfico, el género no influye, pero los adultos mayores alcanzan 41,7 % de fuga y los hogares sin dependientes 31,3 %. En contraste, el cliente con contrato a dos años y débito automático apenas registra un 2,8 % de churn.""")

    # =========================================================================
    # SLIDE 11 — RECOMENDACIONES ACCIONABLES DE NEGOCIO
    # =========================================================================
    s11 = prs.slides.add_slide(blank_layout)
    add_header_footer(
        s11, 11, "10 · Estrategia y Retención",
        "Plan de Acción: 4 Palancas para Reducir el Churn Hoy",
        "Iniciativas concretas y medibles enfocadas en los segmentos de mayor pérdida de ingresos"
    )

    acciones = [
        (
            "01", "Migración Contractual (M2M → Anual)",
            "SEGMENTO: 3.875 clientes mes a mes (88,6 % del churn)",
            [
                ("Acción:", "Ofrecer 10–15 % OFF o tarifa congelada al pasar a plan de 1 o 2 años."),
                ("Momento clave:", "Activar la oferta entre el mes 3 y el mes 6 de permanencia."),
                ("Impacto esperado:", "Mover la retención desde 57 % hacia el 88–97 %."),
            ],
            C_NAVY
        ),
        (
            "02", "Incentivos a Débito Automático",
            "SEGMENTO: 2.365 usuarios de Electronic check (45,3 % churn)",
            [
                ("Acción:", "Bonificación única o 5 % de descuento por adherir tarjeta/transferencia."),
                ("Simplificación:", "Flujo de adhesión en un clic dentro de la factura digital."),
                ("Impacto esperado:", "Reducir la fricción mensual y bajar la fuga hacia ~16 %."),
            ],
            C_CYAN
        ),
        (
            "03", "Bundle de Retención en Fibra Óptica",
            "SEGMENTO: 2.722 clientes fibra sin seguridad o soporte (45,7 % churn)",
            [
                ("Acción:", "Incluir OnlineSecurity y TechSupport gratis 6 meses en altas de fibra."),
                ("Diagnóstico técnico:", "Auditar calidad de red y relación precio/velocidad en fibra."),
                ("Impacto esperado:", "Proteger el segmento de mayor ARPU (US$ 87/mes)."),
            ],
            C_TEAL
        ),
        (
            "04", "Onboarding 90 Días y Atención Senior",
            "SEGMENTO: Primer año (47,4 % churn) y SeniorCitizen (41,7 %)",
            [
                ("Acción:", "Seguimiento proactivo en días 30/60/90 y canal simplificado para seniors."),
                ("Alerta temprana:", "Disparar retención si RatioCostoAntiguedad supera umbral crítico."),
                ("Impacto esperado:", "Cruzar la barrera crítica de los 12 meses de permanencia."),
            ],
            C_CORAL
        ),
    ]

    positions = [
        (Inches(0.55), Inches(1.38)),
        (Inches(6.78), Inches(1.38)),
        (Inches(0.55), Inches(4.20)),
        (Inches(6.78), Inches(4.20)),
    ]
    for (num, title, seg, blts, accent), (x, y) in zip(acciones, positions):
        w, h = Inches(6.00), Inches(2.60)
        add_rect(s11, x, y, w, h, C_WHITE, border_color=C_BORDER, rounded=True, radius=0.05)
        add_rect(s11, x, y, Inches(0.09), h, accent)

        # Círculo/cuadro numerado
        add_rect(s11, x + Inches(0.24), y + Inches(0.18), Inches(0.46), Inches(0.42), accent,
                 rounded=True, radius=0.2)
        add_textbox(s11, x + Inches(0.24), y + Inches(0.21), Inches(0.46), Inches(0.36),
                    num, font_size=13, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)

        add_textbox(s11, x + Inches(0.82), y + Inches(0.16), Inches(4.95), Inches(0.34),
                    title, font_size=14.5, bold=True, color=C_NAVY)
        add_textbox(s11, x + Inches(0.82), y + Inches(0.50), Inches(4.95), Inches(0.24),
                    seg, font_size=9.5, bold=True, color=accent)
        add_bullets(s11, x + Inches(0.26), y + Inches(0.86), Inches(5.52), Inches(1.62),
                    blts, font_size=11.3, space_after=6, bullet_color=accent)

    set_notes(s11, """[Orador: Santiago Valentín Toledo o Sebastian Safatle · 1:15]
Traducimos toda la evidencia en cuatro palancas concretas que la compañía puede ejecutar desde mañana.
Primera: incentivar la migración de contratos mes a mes hacia planes anuales a partir del tercer mes, atacando casi el 89 % del abandono.
Segunda: bonificar el traspaso desde cheque electrónico hacia débito automático con adhesión simple desde la factura digital.
Tercera: empaquetar seguridad online y soporte técnico sin cargo durante seis meses en cada alta de Fibra Óptica, acompañándolo con una auditoría técnica de ese servicio.
Y cuarta: desplegar un programa de onboarding activo durante el primer año con atención prioritaria para adultos mayores y alertas automáticas basadas en nuestro indicador RatioCostoAntiguedad.""")

    # =========================================================================
    # SLIDE 12 — CONCLUSIONES, REPRODUCIBILIDAD Y PRÓXIMOS PASOS
    # =========================================================================
    s12 = prs.slides.add_slide(blank_layout)
    add_header_footer(
        s12, 12, "11 · Cierre y Escalabilidad",
        "Conclusiones, Reproducibilidad y Próximos Pasos",
        "Evidencia validada en entorno limpio de Colab, lista para evolucionar a Machine Learning"
    )

    # 3 Columnas limpias
    cols_12 = [
        (
            "CONCLUSIONES CLAVE",
            "Evidencia de Negocio",
            C_NAVY,
            [
                ("3/3 hipótesis validadas:", "Contrato, tarifa/antigüedad y servicios explican la fuga."),
                ("Churn concentrado:", "No es aleatorio; 4 palancas atacan el núcleo del riesgo."),
                ("Retorno directo:", "Bajar el churn un 20 % recupera ~US$ 334.000/año."),
                ("Foco en CLV:", "Retener supera a aumentar precios en impacto histórico."),
            ],
            "US$ 334k / año", "Potencial de recupero (meta −20 % bajas)", C_SOFT_BLUE
        ),
        (
            "AUDITORÍA COLAB (SEC. 7)",
            "100 % Reproducible",
            C_TEAL,
            [
                ("12/12 controles assert:", "Validados en kernel limpio y simulación Colab."),
                ("Sin fricción técnica:", "Rutas relativas + respaldo GitHub; sin pip install."),
                ("Trazabilidad total:", "Raw intacto; processed/ actualizado (7.043 × 27)."),
                ("Consistencia verificada:", "Cifras exactas en notebook, CSV y tablas."),
            ],
            "12 / 12 ✔", "Auditoría automática superada sin errores", C_SOFT_TEAL
        ),
        (
            "ESCALABILIDAD (ML)",
            "Próximos Pasos",
            C_CYAN,
            [
                ("Modelos supervisados:", "Regresión Logística, Random Forest y XGBoost."),
                ("Ingeniería lista:", "Tratar multicolinealidad de TotalCharges y desbalance 73/27."),
                ("Evaluación rigurosa:", "Priorizar Recall, F1-Score y ROC-AUC en clase Churn."),
                ("Despliegue CRM:", "Scoring mensual de riesgo para retención preventiva."),
            ],
            "Churn Scoring", "De análisis descriptivo a acción predictiva", C_BG
        ),
    ]
    for i, (kicker, title, accent, blts, badge_val, badge_lbl, badge_bg) in enumerate(cols_12):
        x = Inches(0.55 + i * 4.16)
        w = Inches(3.91)
        add_rect(s12, x, Inches(1.36), w, Inches(4.56), C_WHITE,
                 border_color=C_BORDER, rounded=True, radius=0.04)
        add_rect(s12, x, Inches(1.36), w, Inches(0.08), accent)
        add_textbox(s12, x + Inches(0.20), Inches(1.54), w - Inches(0.40), Inches(0.24),
                    kicker, font_size=9.3, bold=True, color=accent)
        add_textbox(s12, x + Inches(0.20), Inches(1.78), w - Inches(0.40), Inches(0.36),
                    title, font_size=15, bold=True, color=C_NAVY)
        add_bullets(s12, x + Inches(0.20), Inches(2.22), w - Inches(0.40), Inches(2.55),
                    blts, font_size=11.0, space_after=7, bullet_color=accent)

        # Sello inferior de columna
        add_rect(s12, x + Inches(0.18), Inches(4.88), w - Inches(0.36), Inches(0.88), badge_bg,
                 border_color=accent, border_pt=0.9, rounded=True, radius=0.08)
        add_textbox(s12, x + Inches(0.28), Inches(4.95), w - Inches(0.56), Inches(0.40),
                    badge_val, font_size=18, bold=True, color=accent, align=PP_ALIGN.CENTER)
        add_textbox(s12, x + Inches(0.28), Inches(5.38), w - Inches(0.56), Inches(0.30),
                    badge_lbl, font_size=9.5, color=C_TEXT, align=PP_ALIGN.CENTER)

    # Banda inferior de cierre / Q&A
    add_rect(s12, Inches(0.55), Inches(6.12), Inches(12.23), Inches(0.76), C_NAVY, rounded=True, radius=0.10)
    add_textbox(
        s12, Inches(0.85), Inches(6.22), Inches(7.2), Inches(0.54),
        "¡Muchas gracias por su atención!  Quedamos abiertos a sus preguntas.",
        font_size=15, bold=True, color=C_WHITE, valign=MSO_ANCHOR.MIDDLE
    )
    add_textbox(
        s12, Inches(8.10), Inches(6.22), Inches(4.45), Inches(0.54),
        "Santiago Toledo · Javier Sosa · Juan Antuñez · Sebastian Safatle",
        font_size=10.5, color=C_CYAN, align=PP_ALIGN.RIGHT, valign=MSO_ANCHOR.MIDDLE
    )

    set_notes(s12, """[Orador: Sebastian Safatle / Cierre Grupal · 1:00]
Para cerrar, resumimos tres grandes resultados. En el negocio, validamos nuestras tres hipótesis y comprobamos que el churn responde a patrones accionables: reducir apenas un 20 % las bajas permitiría recuperar alrededor de 334 mil dólares anuales.
En calidad técnica, auditamos la reproducibilidad completa del notebook en Google Colab desde un entorno limpio, superando 12 controles automáticos sin errores y dejando el dataset limpio versionado.
Finalmente, el siguiente paso natural es escalar a un modelo predictivo de Machine Learning —como Regresión Logística, Random Forest o XGBoost—, tratando la multicolinealidad de TotalCharges y el desbalance de clases para alimentar un puntaje mensual de riesgo en el CRM. Muchas gracias por su atención y quedamos a disposición para las preguntas.""")

    prs.save(str(OUTPUT_PPTX))
    print(f"Presentación generada correctamente: {OUTPUT_PPTX}")
    print(f"Total de diapositivas: {len(prs.slides)}")


def main() -> None:
    print("1. Extrayendo gráficos del notebook...")
    nb_charts = exportar_graficos_notebook()
    for name, path in nb_charts.items():
        print(f"   ✔ {name} ({path.stat().st_size // 1024} KB)")

    print("2. Generando gráficos ejecutivos desde telco_churn_clean.csv...")
    df = pd.read_csv(CLEAN_CSV)
    exec_charts = generar_graficos_ejecutivos(df)
    for name, path in exec_charts.items():
        print(f"   ✔ {name} ({path.stat().st_size // 1024} KB)")

    print("3. Construyendo presentación PowerPoint...")
    build_presentation(nb_charts, exec_charts)


if __name__ == "__main__":
    main()
