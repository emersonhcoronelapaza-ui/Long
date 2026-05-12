# ═══════════════════════════════════════════════════════════════════════
#  SIMULADOR DE DILUCIÓN — TALADROS LARGOS SIMBA S7D
#  Streamlit + PDF gerencial (WeasyPrint)
# ═══════════════════════════════════════════════════════════════════════

import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from scipy.stats import truncnorm
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import io
import base64
import warnings
from datetime import datetime

warnings.filterwarnings("ignore")
import streamlit as st

# ── Configuración de página ──────────────────────────────────────────
st.set_page_config(
    page_title="Simulador de Dilución — Simba S7D",
    page_icon="⛏️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Tema oscuro matplotlib ───────────────────────────────────────────
plt.rcParams.update({
    "figure.facecolor": "#161b22",
    "axes.facecolor": "#161b22",
    "axes.edgecolor": "#30363d",
    "axes.labelcolor": "#c9d1d9",
    "text.color": "#c9d1d9",
    "xtick.color": "#8b949e",
    "ytick.color": "#8b949e",
    "grid.color": "#21262d",
    "figure.dpi": 150,
    "savefig.bbox": "tight",
    "font.size": 10,
})

# ── Estilos Excel ────────────────────────────────────────────────────
HEADER_FILL = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
HEADER_FONT = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
SUB_FILL = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
SUB_FONT = Font(name="Calibri", bold=True, color="000000", size=10)
CELL_FONT = Font(name="Calibri", size=10)
THIN_BORDER = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin"),
)
FILL_GREEN = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
FILL_YELLOW = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
FILL_ORANGE = PatternFill(start_color="FCD5B4", end_color="FCD5B4", fill_type="solid")
FILL_RED = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

# ── Parámetros de entrada ───────────────────────────────────────────
PARAM_INFO = {
    "H": {"label": "Altura entre niveles (H)", "unit": "m", "hint": "Duro/Típico: 15-25 | Blando/Inestable: 10-15", "default": 20.0, "std_frac": 0.0, "section": "Geometría del Sistema"},
    "W_sup": {"label": "Ancho sección nivel superior/collar (W_sup)", "unit": "m", "hint": "Típico manto: 3.0-5.0", "default": 4.0, "std_frac": 0.0, "section": "Geometría del Sistema"},
    "W_inf": {"label": "Ancho sección nivel inferior/ampliada (W_inf)", "unit": "m", "hint": "Típico ampliación: 7.0-10.0", "default": 9.0, "std_frac": 0.0, "section": "Geometría del Sistema"},
    "A_hast": {"label": "Ampliación por hastial (A_hast)", "unit": "m", "hint": "ENTER = (W_inf - W_sup)/2", "default": 2.5, "std_frac": 0.0, "section": "Geometría del Sistema"},
    "dx_fondo": {"label": "Desplazamiento horizontal del fondo (dx_fondo)", "unit": "m", "hint": "Duro: 1.0-2.0 | Intermedio: 2.0-3.5 | Blando: 3.5-5.0", "default": 3.5, "std_frac": 0.0, "section": "Geometría del Sistema"},
    "B": {"label": "Burden de diseño entre anillos (B)", "unit": "m", "hint": "Duro: 1.2-1.5 | Intermedio: 1.5-2.0 | Blando: 0.8-1.2", "default": 1.5, "std_frac": 0.0, "section": "Geometría del Sistema"},
    "S": {"label": "Espaciado entre fondos de taladros (S)", "unit": "m", "hint": "Duro: 1.5-2.0 | Intermedio: 1.2-1.5 | Blando: 0.8-1.2", "default": 1.4, "std_frac": 0.0, "section": "Geometría del Sistema"},
    "delta_manto": {"label": "Buzamiento del manto (delta_manto)", "unit": "°", "hint": "Sub-horizontal: 0-15 | Inclinado: 15-45 | Empinado: 45-90", "default": 20.0, "std_frac": 0.0, "section": "Geometría del Sistema"},
    "r_taladro": {"label": "Radio del taladro (r_taladro)", "unit": "m", "hint": "Simba S7D Ø64mm: 0.032", "default": 0.032, "std_frac": 0.0, "section": "Geometría del Sistema"},
    "GSI": {"label": "Geological Strength Index (GSI)", "unit": "[0-100]", "hint": "Duro/Masivo: 65-85 | Intermedio: 40-65 | Blando/Falla: 15-35", "default": 45.0, "std_frac": 0.10, "section": "Parámetros Geomecánicos"},
    "RQD": {"label": "Rock Quality Designation (RQD)", "unit": "[%]", "hint": "Excelente: >90 | Bueno: 75-90 | Regular: 50-75 | Malo: 25-50", "default": 60.0, "std_frac": 0.08, "section": "Parámetros Geomecánicos"},
    "sigma_ci": {"label": "Resistencia uniaxial roca intacta (sigma_ci)", "unit": "MPa", "hint": "Duro: >100 | Intermedio: 50-100 | Blando/Alterado: 10-50", "default": 75.0, "std_frac": 0.15, "section": "Parámetros Geomecánicos"},
    "mi": {"label": "Constante de roca Hoek-Brown (mi)", "unit": "[-]", "hint": "Duro/Ígneo: 25-35 | Intermedio: 10-20 | Blando/Sedimento: 5-10", "default": 15.0, "std_frac": 0.05, "section": "Parámetros Geomecánicos"},
    "D_blast": {"label": "Factor de daño por voladura (D)", "unit": "[0-1]", "hint": "Favorable: 0.0-0.2 | Típico: 0.5-0.7 | Severo: 0.8-1.0", "default": 0.7, "std_frac": 0.10, "section": "Parámetros Geomecánicos"},
    "sigma_v": {"label": "Esfuerzo principal vertical in situ (sigma_v)", "unit": "MPa", "hint": "Poco profundo: 2-5 | Intermedio: 5-15 | Profundo: >15", "default": 8.0, "std_frac": 0.12, "section": "Parámetros Geomecánicos"},
    "sigma_h": {"label": "Esfuerzo principal horizontal (sigma_h)", "unit": "MPa", "hint": "Isotrópico: sigma_v | Anisotrópico: 0.5-2.5 * sigma_v", "default": 10.0, "std_frac": 0.12, "section": "Parámetros Geomecánicos"},
    "RMR": {"label": "Rock Mass Rating base (RMR)", "unit": "[0-100]", "hint": "Buena: >65 | Intermedia: 40-65 | Mala: <40", "default": 50.0, "std_frac": 0.08, "section": "Clasificación Geomecánica (Laubscher)"},
    "Q_prime": {"label": "Índice Q modificado (Q') Mathews-Potvin", "unit": "[-]", "hint": "Bueno: >10 | Intermedio: 1-10 | Malo: 0.1-1.0", "default": 2.5, "std_frac": 0.15, "section": "Clasificación Geomecánica (Laubscher)"},
    "f_weather": {"label": "Factor meteorización Laubscher (f)", "unit": "[-]", "hint": "Fresco: 1.0 | Leve: 0.90 | Severo: 0.70", "default": 0.90, "std_frac": 0.05, "section": "Clasificación Geomecánica (Laubscher)"},
    "w_stress": {"label": "Factor orientación esfuerzos (w)", "unit": "[-]", "hint": "Favorable: 1.20 | Típico: 1.00 | Desfavorable: 0.70", "default": 0.85, "std_frac": 0.05, "section": "Clasificación Geomecánica (Laubscher)"},
    "s_orient": {"label": "Factor inducción de esfuerzos (s_orient)", "unit": "[-]", "hint": "Bajo: 1.00 | Medio: 0.80 | Alto/Esq. irreg.: 0.60", "default": 0.80, "std_frac": 0.05, "section": "Clasificación Geomecánica (Laubscher)"},
    "d_explosive": {"label": "Factor daño explosivos Laubscher (d_explosive)", "unit": "[-]", "hint": "Leve: 1.00 | Típico: 0.90 | Severo/Repetido: 0.80", "default": 0.85, "std_frac": 0.05, "section": "Clasificación Geomecánica (Laubscher)"},
    "A_mp": {"label": "Factor resistencia roca Mathews-Potvin (A)", "unit": "[-]", "hint": "Roca fuerte: 1.0 | Intermedia: 0.5-0.8 | Débil: 0.1-0.4", "default": 0.6, "std_frac": 0.0, "section": "Mathews-Potvin"},
    "B_mp": {"label": "Factor orientación disc. Mathews-Potvin (B)", "unit": "[-]", "hint": "Favorable: 0.8-1.0 | Intermedio: 0.5-0.7 | Desfavorable: 0.2-0.4", "default": 0.3, "std_frac": 0.0, "section": "Mathews-Potvin"},
    "C_mp": {"label": "Factor gravedad Mathews-Potvin (C)", "unit": "[-]", "hint": "Techo: 8.0 | Sub-vertical: 5.0-6.0 | Vertical: 2.0", "default": 5.5, "std_frac": 0.0, "section": "Mathews-Potvin"},
    "VOD": {"label": "Velocidad de detonación (VOD)", "unit": "m/s", "hint": "ANFO: 3000-3500 | Emulsión: 4500-5500 | Alta energía: >5500", "default": 5000.0, "std_frac": 0.03, "section": "Parámetros de Voladura"},
    "rho_exp": {"label": "Densidad del explosivo (rho_exp)", "unit": "kg/m³", "hint": "ANFO: 800-850 | Emulsión: 1100-1250", "default": 1150.0, "std_frac": 0.02, "section": "Parámetros de Voladura"},
    "WSR": {"label": "Potencia relativa explosivo (WSR vs ANFO)", "unit": "[-]", "hint": "ANFO: 1.0 | Emulsión: 1.1-1.3 | Heavy ANFO: 1.05", "default": 1.15, "std_frac": 0.0, "section": "Parámetros de Voladura"},
    "A_rock": {"label": "Factor de roca Kuz-Ram (A_rock)", "unit": "[-]", "hint": "Duro/Masivo: 10-12 | Intermedio: 7-9 | Blando/Fracturado: 4-6", "default": 8.0, "std_frac": 0.0, "section": "Parámetros de Voladura"},
    "D_geo": {"label": "Dilución geométrica base (D_geo)", "unit": "[%]", "hint": "Diseño óptimo: 1-3 | Típico: 3-5 | Deficiente: >5", "default": 3.0, "std_frac": 0.20, "section": "Estimación de Dilución"},
    "N_sim": {"label": "Número de simulaciones (Monte Carlo)", "unit": "[-]", "hint": "Mayor = más precisión pero más lento", "default": 10000, "std_frac": 0.0, "section": "Simulación Probabilística"},
}


# ═══════════════════════════════════════════════════════════════════════
#  FUNCIONES DE CÁLCULO DETERMINÍSTICO (ESCALARES PUROS)
# ═══════════════════════════════════════════════════════════════════════

def calc_geometry(p):
    L = np.sqrt(p["H"]**2 + p["dx_fondo"]**2)
    alpha = np.degrees(np.arctan(p["dx_fondo"] / p["H"]))
    beta = 90.0 - p["delta_manto"] - alpha
    P_efectiva = L * np.cos(np.radians(alpha))
    f_corr = 1.0 / np.cos(np.radians(alpha)) if alpha < 89 else 10.0
    return {"L": L, "alpha": alpha, "beta": beta, "P_efectiva": P_efectiva, "f_corr": f_corr}

def calc_blasting(p, g):
    f_lang = 1.0 - ((100.0 - p["RMR"]) / 100.0)
    S_tecnico = p["S"] * f_lang
    q = 0.09 * p["WSR"] + 0.32
    pl = q * p["B"] * S_tecnico
    pl_corr = pl * g["f_corr"]
    Q_taladro = pl_corr * g["L"]
    return {"f_lang": f_lang, "S_tecnico": S_tecnico, "q": q, "pl": pl, "pl_corr": pl_corr, "Q_taladro": Q_taladro}

def calc_holmberg(p, g):
    k = 0.020 if g["alpha"] > 45 else 0.005
    if p["RMR"] < 40: k = 0.015
    exp_holm = 2.0 if p["RMR"] < 40 else 1.5
    deviation = k * (g["L"] ** exp_holm)
    return {"k": k, "exp_holm": exp_holm, "deviation": deviation}

def calc_kirsch(p):
    sigma_r = 0.0
    sigma_theta = 3.0 * p["sigma_v"] - p["sigma_h"]
    return {"sigma_r": sigma_r, "sigma_theta": sigma_theta}

def calc_hoek_brown(p):
    mb = p["mi"] * np.exp((p["GSI"] - 100.0) / 28.0)
    s = np.exp((p["GSI"] - 100.0) / 9.0)
    a_hb = 0.5 + (p["GSI"] - 100.0) / 200.0
    sigma_cm = p["sigma_ci"] * ((mb + 4.0 * s - a_hb * (mb - 8.0 * s)) * (s / (mb + 4.0 * s))**(a_hb - 1.0))
    return {"mb": mb, "s": s, "a_hb": a_hb, "sigma_cm": sigma_cm}

def calc_mathews_potvin(p):
    N_prime = p["Q_prime"] * p["A_mp"] * p["B_mp"] * p["C_mp"]
    hydraulic_radius = p["A_hast"] * p["H"] / (2.0 * (p["A_hast"] + p["H"]))
    return {"N_prime": N_prime, "hydraulic_radius": hydraulic_radius}

def calc_laubscher(p):
    mrmr = p["RMR"] * p["f_weather"] * p["w_stress"] * p["s_orient"] * p["d_explosive"]
    d_est = max(0.0, np.exp(5.5 - 0.1 * mrmr)) if mrmr < 60 else 0.0
    return {"mrmr": mrmr, "d_est": d_est}

def calc_perimetral_dilation(p, b):
    rho_t = p["rho_exp"] / 1000.0
    Pd = (p["VOD"]**2 * rho_t) * 1e-6
    r_sobre = 0.04 * (Pd / p["sigma_ci"])**0.5
    safe_pl = b["pl_corr"] if b["pl_corr"] > 0 else 1e-10
    V_peri = np.pi * (r_sobre**2 - p["r_taladro"]**2) * b["Q_taladro"] / safe_pl
    V_roca = p["B"] * p["S"] * (b["Q_taladro"] / safe_pl)
    d_peri = (V_peri / V_roca) * 100 if V_roca > 0 else 0
    return {"Pd": Pd, "r_sobre": r_sobre, "d_peri": d_peri}

def calc_kuzram(p, b, g):
    V0 = p["B"] * p["S"] * g["L"] if g["L"] > 0 else 1.0
    x50 = p["A_rock"] * (b["Q_taladro"] / V0)**0.8 * b["q"]**-0.8 * (115.0 / p["VOD"])**(1.0 / 3.0)
    return {"V0": V0, "x50": x50}

def calc_total_dilation(p, d_peri, d_est):
    d_total = p["D_geo"] + d_peri + d_est
    T_rec = 100.0 / (1.0 + d_total / 100.0)
    return {"d_total": d_total, "T_rec": T_rec}


# ═══════════════════════════════════════════════════════════════════════
#  MONTE CARLO VECTORIZADO (NUMPY PURO - SIN LOOPS)
# ═══════════════════════════════════════════════════════════════════════

def monte_carlo_vectorized(p):
    N = int(p["N_sim"])
    stoch_keys = [k for k, v in PARAM_INFO.items() if v["std_frac"] > 0 and k != "N_sim"]
    
    samples = {}
    for key in stoch_keys:
        mean = p[key]
        std = mean * PARAM_INFO[key]["std_frac"]
        if std <= 0: std = 1e-4
        if key == "D_geo": low, high = 0.0, mean * 3.0
        elif key == "GSI": low, high = 5.0, 95.0
        elif key in ("RQD", "RMR"): low, high = 0.0, 100.0
        else: low, high = mean - 4 * std, mean + 4 * std
        
        a_clip = (low - mean) / std
        b_clip = (high - mean) / std
        samples[key] = truncnorm.rvs(a_clip, b_clip, loc=mean, scale=std, size=N)
        
    for key in [k for k, v in PARAM_INFO.items() if v["std_frac"] == 0 and k != "N_sim"]:
        samples[key] = np.full(N, p[key])

    H = samples["H"]; dx = samples["dx_fondo"]
    L = np.sqrt(H**2 + dx**2)
    alpha = np.degrees(np.arctan(dx / H))
    f_corr = np.where(alpha < 89, 1.0 / np.cos(np.radians(alpha)), 10.0)

    f_lang = 1.0 - ((100.0 - samples["RMR"]) / 100.0)
    S_tec = samples["S"] * f_lang
    q = 0.09 * samples["WSR"] + 0.32
    pl = q * samples["B"] * S_tec
    pl_corr = pl * f_corr
    Q_tal = pl_corr * L

    k_h = np.where(alpha > 45, 0.020, 0.005)
    k_h = np.where(samples["RMR"] < 40, 0.015, k_h)
    exp_h = np.where(samples["RMR"] < 40, 2.0, 1.5)
    deviation = k_h * (L ** exp_h)

    sigma_theta = 3.0 * samples["sigma_v"] - samples["sigma_h"]

    mb = samples["mi"] * np.exp((samples["GSI"] - 100.0) / 28.0)
    s_hb = np.exp((samples["GSI"] - 100.0) / 9.0)
    a_hb = 0.5 + (samples["GSI"] - 100.0) / 200.0
    base_hb = np.maximum(s_hb / (mb + 4.0 * s_hb), 1e-12)
    sigma_cm = samples["sigma_ci"] * ((mb + 4.0 * s_hb - a_hb * (mb - 8.0 * s_hb)) * base_hb**(a_hb - 1.0))

    N_prime = samples["Q_prime"] * samples["A_mp"] * samples["B_mp"] * samples["C_mp"]
    hr = samples["A_hast"] * samples["H"] / (2.0 * (samples["A_hast"] + samples["H"]))

    mrmr = samples["RMR"] * samples["f_weather"] * samples["w_stress"] * samples["s_orient"] * samples["d_explosive"]
    d_est = np.where(mrmr < 60, np.maximum(np.exp(5.5 - 0.1 * mrmr), 0.0), 0.0)

    rho_t = samples["rho_exp"] / 1000.0
    Pd = (samples["VOD"]**2 * rho_t) * 1e-6
    r_sob = 0.04 * np.maximum(Pd / samples["sigma_ci"], 0)**0.5
    safe_pl = np.where(pl_corr > 0, pl_corr, 1e-10)
    V_peri = np.pi * (r_sob**2 - samples["r_taladro"]**2) * Q_tal / safe_pl
    V_roca = samples["B"] * samples["S"] * (Q_tal / safe_pl)
    d_peri = np.where(V_roca > 0, (V_peri / V_roca) * 100, 0)

    V0 = np.where(L > 0, samples["B"] * samples["S"] * L, 1.0)
    x50 = samples["A_rock"] * (Q_tal / V0)**0.8 * np.maximum(q, 1e-6)**-0.8 * (115.0 / samples["VOD"])**(1.0 / 3.0)

    d_total = samples["D_geo"] + d_peri + d_est
    T_rec = 100.0 / (1.0 + d_total / 100.0)

    res = {
        "x50_mc": np.nan_to_num(x50, nan=0, posinf=0, neginf=0),
        "d_total_mc": np.nan_to_num(d_total, nan=0),
        "N_prime_mc": N_prime,
        "deviation_mc": deviation,
        "sigma_theta_mc": sigma_theta,
        "mrmr_mc": mrmr,
    }
    for key in stoch_keys:
        res[key] = samples[key]
    return res


# ═══════════════════════════════════════════════════════════════════════
#  FUNCIONES DE INTERPRETACIÓN
# ═══════════════════════════════════════════════════════════════════════

def interpret_kirsch(st, sci):
    if st > 0.5 * sci: return f"Esfuerzo tangencial ({st:.2f} MPa) supera 50% de sigma_ci. Riesgo ALTO de spalling. Se requiere encapsulado obligatorio del collar con resina epoxi.", "red"
    elif st > 0: return f"Esfuerzo tangencial ({st:.2f} MPa) de compresión moderada. Riesgo BAJO de falla en collar con macizo competente.", "green"
    else: return f"Esfuerzo tangencial negativo ({st:.2f} MPa) indica tracción. Riesgo CRÍTICO de desprendimiento. Atasco inminente del barreno.", "red"

def interpret_holmberg(dev, al):
    if al > 45: return f"Desviación: {dev:.3f} m (alfa={al:.1f}°). CRÍTICO. Plantilla rígida + inclinómetro + reducción avance 30-40%.", "red"
    elif al > 25: return f"Desviación: {dev:.3f} m (alfa={al:.1f}°). MODERADO. Verificar con inclinómetro + registrar desviación real.", "yellow"
    else: return f"Desviación: {dev:.3f} m (alfa={al:.1f}°). BAJO. Control estándar de collar.", "green"

def interpret_mathews(np_, hr):
    if np_ < 1.2: return f"N'={np_:.2f}, RH={hr:.2f}m. COLAPSO. Sostenimiento intensivo previo al disparo. Reducir span.", "red"
    elif np_ < 4.0: return f"N'={np_:.2f}, RH={hr:.2f}m. INESTABLE. Sostenimiento sistemático + monitoreo convergencia obligatorio.", "orange"
    elif np_ < 8.0: return f"N'={np_:.2f}, RH={hr:.2f}m. ESTABLE TRANSITORIO. Pernos puntuales + shotcrete si water ingress.", "yellow"
    else: return f"N'={np_:.2f}, RH={hr:.2f}m. ESTABLE. Generalmente sin sostenimiento post-disparo.", "green"

def interpret_kuzram(x50):
    if x50 > 600: return f"x50={x50:.1f}mm. MUY GRUESA. Riesgo extremo de tacos. Aumentar carga específica inmediatamente.", "red"
    elif x50 > 300: return f"x50={x50:.1f}mm. GRUESA. Alta probabilidad de atascos. Aumentar WSR o reducir espaciado.", "orange"
    elif x50 > 150: return f"x50={x50:.1f}mm. ACEPTABLE. Compatible con LHD 3-4 yd³ estándar.", "green"
    else: return f"x50={x50:.1f}mm. FINA. Riesgo de exceso de finos y sobreconsumo de explosivos.", "yellow"

def interpret_dilution(dt, dg, dp, de, mr):
    if dt > 25: return f"Dilución Total: {dt:.2f}% (Geo:{dg:.2f}%|Peri:{dp:.2f}%|Est:{de:.2f}%). CRÍTICA. MRMR:{mr:.1f}. Rediseño total del patrón o cambio de método.", "red"
    elif dt > 15: return f"Dilución Total: {dt:.2f}% (Geo:{dg:.2f}%|Peri:{dp:.2f}%|Est:{de:.2f}%). SEVERA. MRMR:{mr:.1f}. Revisar carga específica + reforzar sostenimiento.", "orange"
    elif dt > 8: return f"Dilución Total: {dt:.2f}% (Geo:{dg:.2f}%|Peri:{dp:.2f}%|Est:{de:.2f}%). MODERADA. MRMR:{mr:.1f}. Margen de mejora en geometría de malla.", "yellow"
    else: return f"Dilución Total: {dt:.2f}% (Geo:{dg:.2f}%|Peri:{dp:.2f}%|Est:{de:.2f}%). FAVORABLE. MRMR:{mr:.1f}. Diseño adecuado.", "green"


# ═══════════════════════════════════════════════════════════════════════
#  GRÁFICOS DARK THEME (PARA PANTALLA STREAMLIT)
# ═══════════════════════════════════════════════════════════════════════

def plot_mc_histograms(mc):
    fig, axes = plt.subplots(2, 2, figsize=(13, 10))
    vp = [
        ("d_total_mc", "Dilución Total (%)", "Distribución — Dilución Total"),
        ("x50_mc", "Tamaño Medio x50 (mm)", "Distribución — Fragmentación Kuz-Ram"),
        ("N_prime_mc", "Número de Estabilidad N'", "Distribución — Mathews-Potvin N'"),
        ("deviation_mc", "Desviación Holmberg (m)", "Distribución — Desviación Taladros")
    ]
    for ax, (k, xl, t) in zip(axes.flatten(), vp):
        d = mc[k]
        ax.hist(d, bins=50, density=True, alpha=0.65, color="#58a6ff", edgecolor="#30363d", linewidth=0.5)
        ax2 = ax.twinx()
        sd = np.sort(d)
        cdf = np.arange(1, len(sd) + 1) / len(sd)
        ax2.plot(sd, cdf, color="#f85149", linewidth=2)
        ax2.set_ylabel("CDF", color="#f85149", fontsize=9)
        ax2.tick_params(axis="y", colors="#f85149")
        ax.set_xlabel(xl, fontsize=9)
        ax.set_title(t, fontsize=10, fontweight="bold", pad=8)
        ax.grid(True, linestyle="--", alpha=0.3)
    plt.tight_layout(pad=2.0)
    return fig

def plot_tornado(mc):
    tgt = mc["d_total_mc"]
    sv = ["GSI", "sigma_ci", "RMR", "Q_prime", "D_geo", "sigma_v", "RQD"]
    cr = {v: np.corrcoef(mc[v], tgt)[0, 1] for v in sv}
    sc = dict(sorted(cr.items(), key=lambda x: x[1]))
    fig, ax = plt.subplots(figsize=(10, 5))
    cols = ["#f85149" if v < 0 else "#58a6ff" for v in sc.values()]
    bars = ax.barh(list(sc.keys()), list(sc.values()), color=cols, edgecolor="#30363d", height=0.6)
    for b, v in zip(bars, sc.values()):
        ax.text(v + (0.01 if v >= 0 else -0.01), b.get_y() + b.get_height() / 2,
                f"{v:.3f}", va="center", ha="left" if v >= 0 else "right", fontsize=9, color="#c9d1d9")
    ax.set_xlabel("Correlación Pearson vs Dilución Total", fontsize=10)
    ax.set_title("Diagrama de Tornado — Sensibilidad", fontsize=11, fontweight="bold", pad=10)
    ax.axvline(0, color="#8b949e", linewidth=0.8)
    ax.grid(True, axis="x", linestyle="--", alpha=0.3)
    ax.set_xlim(-1.1, 1.1)
    plt.tight_layout()
    return fig

def plot_scatters(mc):
    fig, axes = plt.subplots(1, 2, figsize=(14, 6))
    axes[0].scatter(mc["GSI"], mc["d_total_mc"], alpha=0.08, s=10, c="#58a6ff")
    z = np.polyfit(mc["GSI"], mc["d_total_mc"], 1)
    axes[0].plot(sorted(mc["GSI"]), np.poly1d(z)(sorted(mc["GSI"])), color="#f85149", linewidth=2)
    axes[0].set_xlabel("GSI"); axes[0].set_ylabel("Dilución Total (%)")
    axes[0].set_title("GSI vs Dilución Total", fontsize=10, fontweight="bold")
    axes[0].grid(True, linestyle="--", alpha=0.3)

    axes[1].scatter(mc["sigma_ci"], mc["x50_mc"], alpha=0.08, s=10, c="#3fb950")
    z2 = np.polyfit(mc["sigma_ci"], mc["x50_mc"], 1)
    axes[1].plot(sorted(mc["sigma_ci"]), np.poly1d(z2)(sorted(mc["sigma_ci"])), color="#f85149", linewidth=2)
    axes[1].set_xlabel("sigma_ci (MPa)"); axes[1].set_ylabel("x50 (mm)")
    axes[1].set_title("Resistencia vs Fragmentación", fontsize=10, fontweight="bold")
    axes[1].grid(True, linestyle="--", alpha=0.3)
    plt.tight_layout(pad=2.0)
    return fig

def plot_polar(p):
    ang = np.linspace(0, 60, 100)
    dxr = p["H"] * np.tan(np.radians(ang))
    Lr = np.sqrt(p["H"]**2 + dxr**2)
    Pr = Lr * np.cos(np.radians(ang))
    fig, ax = plt.subplots(subplot_kw={"projection": "polar"}, figsize=(8, 8))
    ax.plot(np.radians(ang), Lr, label="Longitud Real L (m)", color="#58a6ff", linewidth=2)
    ax.plot(np.radians(ang), Pr * 2, label="Proyección Efectiva x2 (m)", color="#f85149", linestyle="--", linewidth=2)
    ax.set_theta_zero_location("N"); ax.set_theta_direction(-1)
    ax.set_title("Geometría Polar (0°=Vertical)", va="bottom", fontsize=11, fontweight="bold", pad=15)
    ax.legend(loc="lower right", bbox_to_anchor=(1.25, -0.05), fontsize=9)
    plt.tight_layout()
    return fig


# ═══════════════════════════════════════════════════════════════════════
#  EXPORTAR EXCEL A BUFFER
# ═══════════════════════════════════════════════════════════════════════

def export_excel_buffer(p, dr, mc):
    buf = io.BytesIO()
    wb = openpyxl.Workbook()

    def sw(ws, w):
        for i, wi in enumerate(w, 1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = wi

    # Hoja 1
    ws1 = wb.active; ws1.title = "1_Datos_Entrada"
    ws1.append(["Parámetro", "Valor", "Unidades", "Valores típicos", "CoV MC (%)"])
    for c in ws1[1]: c.font = HEADER_FONT; c.fill = HEADER_FILL; c.alignment = Alignment(horizontal="center", wrap_text=True); c.border = THIN_BORDER
    sw(ws1, [45, 18, 15, 60, 20])
    cs = None
    for k, inf in PARAM_INFO.items():
        if k == "N_sim": continue
        if inf["section"] != cs:
            cs = inf["section"]; ws1.append([cs, "", "", "", ""])
            for c in ws1[ws1.max_row]: c.font = SUB_FONT; c.fill = SUB_FILL; c.border = THIN_BORDER
        cv = f"{inf['std_frac']*100:.1f}%" if inf["std_frac"] > 0 else "Determinístico"
        ws1.append([inf["label"], p[k], inf["unit"], inf["hint"], cv])
        for c in ws1[ws1.max_row]: c.font = CELL_FONT; c.border = THIN_BORDER

    # Hoja 2
    ws2 = wb.create_sheet("2_Interpretacion")
    ws2.append(["Análisis", "Interpretación Ingenieril"])
    for c in ws2[1]: c.font = HEADER_FONT; c.fill = HEADER_FILL; c.alignment = Alignment(wrap_text=True); c.border = THIN_BORDER
    sw(ws2, [35, 100])
    rf = {"green": FILL_GREEN, "yellow": FILL_YELLOW, "orange": FILL_ORANGE, "red": FILL_RED}
    il = [
        ("Kirsch", interpret_kirsch(dr["kirsch"]["sigma_theta"], p["sigma_ci"])),
        ("Holmberg", interpret_holmberg(dr["holmberg"]["deviation"], dr["geom"]["alpha"])),
        ("Mathews-Potvin", interpret_mathews(dr["mathews"]["N_prime"], dr["mathews"]["hydraulic_radius"])),
        ("Kuz-Ram", interpret_kuzram(dr["kuzram"]["x50"])),
        ("Dilución", interpret_dilution(dr["dilution"]["d_total"], p["D_geo"], dr["peri"]["d_peri"], dr["laubscher"]["d_est"], dr["laubscher"]["mrmr"]))
    ]
    for t, (tx, rk) in il:
        ws2.append([t, tx])
        for c in ws2[ws2.max_row]: c.font = CELL_FONT; c.fill = rf[rk]; c.alignment = Alignment(wrap_text=True, vertical="top"); c.border = THIN_BORDER
        ws2.row_dimensions[ws2.max_row].height = 60

    # Hoja 3
    ws3 = wb.create_sheet("3_Calculos")
    ws3.append(["Variable", "Valor", "Unidades"])
    for c in ws3[1]: c.font = HEADER_FONT; c.fill = HEADER_FILL; c.border = THIN_BORDER
    sw(ws3, [45, 20, 15])
    rd = [
        ("GEOMETRÍA", None, None), ("Longitud real L", dr["geom"]["L"], "m"),
        ("Inclinación alfa", dr["geom"]["alpha"], "°"), ("Proyección efectiva P", dr["geom"]["P_efectiva"], "m"),
        ("VOLADURA", None, None), ("Espaciado técnico S_tec", dr["blast"]["S_tecnico"], "m"),
        ("Carga específica q", dr["blast"]["q"], "kg/m³"), ("Carga lineal corregida", dr["blast"]["pl_corr"], "kg/m"),
        ("GEOMECÁNICA", None, None), ("Esfuerzo tangencial collar", dr["kirsch"]["sigma_theta"], "MPa"),
        ("s Hoek-Brown", dr["hoek"]["s"], "[-]"), ("sigma_cm macizo", dr["hoek"]["sigma_cm"], "MPa"),
        ("DILUCIÓN", None, None), ("Dilución geométrica", p["D_geo"], "%"),
        ("Dilución perimetral", dr["peri"]["d_peri"], "%"), ("Dilución estructural", dr["laubscher"]["d_est"], "%"),
        ("DILUCIÓN TOTAL", dr["dilution"]["d_total"], "%"), ("TONELAJE RECUPERADO", dr["dilution"]["T_rec"], "%")
    ]
    for v, val, u in rd:
        if val is None:
            ws3.append([v, "", ""])
            for c in ws3[ws3.max_row]: c.font = SUB_FONT; c.fill = SUB_FILL; c.border = THIN_BORDER
        else:
            ws3.append([v, round(val, 4), u])
            for c in ws3[ws3.max_row]: c.font = CELL_FONT; c.border = THIN_BORDER

    # Hoja 4
    ws4 = wb.create_sheet("4_MC_Estadisticas")
    ws4.append(["Variable", "Media", "Desv.Est.", "P5", "P50", "P90", "P95"])
    for c in ws4[1]: c.font = HEADER_FONT; c.fill = HEADER_FILL; c.border = THIN_BORDER
    sw(ws4, [30, 15, 15, 15, 15, 15, 15])
    mv = {"Dilución Total (%)": "d_total_mc", "x50 (mm)": "x50_mc", "N' Mathews": "N_prime_mc", "Desviación (m)": "deviation_mc", "MRMR": "mrmr_mc"}
    for n, k in mv.items():
        d = mc[k]
        ws4.append([n, round(float(np.mean(d)), 2), round(float(np.std(d)), 2), round(float(np.percentile(d, 5)), 2), round(float(np.percentile(d, 50)), 2), round(float(np.percentile(d, 90)), 2), round(float(np.percentile(d, 95)), 2)])
        for c in ws4[ws4.max_row]: c.font = CELL_FONT; c.border = THIN_BORDER; c.number_format = "0.00"
    r = 3
    ws4.cell(row=r, column=1, value="Correlaciones vs Dilución Total").font = SUB_FONT
    ws4.cell(row=r, column=1).fill = SUB_FILL
    for v in ["GSI", "sigma_ci", "RMR", "Q_prime", "D_geo", "sigma_v", "RQD"]:
        r += 1
        cr_val = np.corrcoef(mc[v], mc["d_total_mc"])[0, 1]
        ws4.cell(row=r, column=1, value=v).font = CELL_FONT
        ws4.cell(row=r, column=2, value=round(float(cr_val), 4)).font = CELL_FONT
        ws4.cell(row=r, column=2).number_format = "0.0000"
        for c in range(1, 3): ws4.cell(row=r, column=c).border = THIN_BORDER

    # Hoja 5
    ws5 = wb.create_sheet("5_MC_Muestra")
    sk = [k for k, v in PARAM_INFO.items() if v["std_frac"] > 0 and k != "N_sim"]
    cm = sk + ["d_total_mc", "x50_mc"]
    hm = sk + ["Dilución Total (%)", "x50 (mm)"]
    ws5.append(hm)
    for c in ws5[1]: c.font = HEADER_FONT; c.fill = HEADER_FILL; c.border = THIN_BORDER
    sw(ws5, [15] * len(hm))
    for i in range(min(500, int(p["N_sim"]))):
        ws5.append([round(float(mc[k][i]), 2) for k in cm])
        for c in ws5[ws5.max_row]: c.font = CELL_FONT; c.border = THIN_BORDER; c.number_format = "0.00"

    wb.save(buf)
    buf.seek(0)
    return buf


# ═══════════════════════════════════════════════════════════════════════
#  GENERACIÓN DE PDF GERENCIAL (WEASYPRINT)
# ═══════════════════════════════════════════════════════════════════════

MESES = {1:"Enero",2:"Febrero",3:"Marzo",4:"Abril",5:"Mayo",6:"Junio",
         7:"Julio",8:"Agosto",9:"Septiembre",10:"Octubre",11:"Noviembre",12:"Diciembre"}
RC = {"green":("#e8f8e8","#1b5e20","#27ae60"),"yellow":("#fff8e1","#f57f17","#f39c12"),
      "orange":("#fff3e0","#e65100","#e67e22"),"red":("#ffebee","#b71c1c","#e74c3c")}
RL = {"green":"FAVORABLE","yellow":"MODERADO","orange":"SEVERO","red":"CRÍTICO"}

def _fig_b64(fig):
    b = io.BytesIO()
    fig.savefig(b, format="png", dpi=150, bbox_inches="tight", facecolor=fig.get_facecolor())
    plt.close(fig)
    b.seek(0)
    return base64.b64encode(b.read()).decode()

def _pdf_figs(p, mc):
    old = plt.rcParams.copy()
    plt.rcParams.update({"figure.facecolor":"white","axes.facecolor":"white","axes.edgecolor":"#bbb",
        "axes.labelcolor":"#333","text.color":"#333","xtick.color":"#333","ytick.color":"#333",
        "grid.color":"#e0e0e0","font.size":9,"figure.dpi":150})
    imgs = {
        "hist": _fig_b64(plot_mc_histograms(mc)),
        "tornado": _fig_b64(plot_tornado(mc)),
        "scatter": _fig_b64(plot_scatters(mc)),
        "polar": _fig_b64(plot_polar(p))
    }
    plt.rcParams.update(old)
    return imgs

def generate_pdf(p, dr, mc, op_name, author):
    try:
        from weasyprint import HTML
    except Exception:
        return None
        
    imgs = _pdf_figs(p, mc)
    now = datetime.now()
    fecha = f"{now.day} de {MESES[now.month]} de {now.year}"
    
    txt_k, rk_k = interpret_kirsch(dr["kirsch"]["sigma_theta"], p["sigma_ci"])
    txt_h, rk_h = interpret_holmberg(dr["holmberg"]["deviation"], dr["geom"]["alpha"])
    txt_m, rk_m = interpret_mathews(dr["mathews"]["N_prime"], dr["mathews"]["hydraulic_radius"])
    txt_kr, rk_kr = interpret_kuzram(dr["kuzram"]["x50"])
    txt_d, rk_d = interpret_dilution(dr["dilution"]["d_total"], p["D_geo"], dr["peri"]["d_peri"], dr["laubscher"]["d_est"], dr["laubscher"]["mrmr"])
    
    mc_rows = ""
    for nm, k in [("Dilución Total (%)","d_total_mc"),("x50 (mm)","x50_mc"),("N' Mathews","N_prime_mc"),("Desviación (m)","deviation_mc"),("MRMR","mrmr_mc")]:
        d = mc[k]
        mc_rows += f"<tr><td>{nm}</td><td>{np.mean(d):.2f}</td><td>{np.std(d):.2f}</td><td>{np.percentile(d,5):.2f}</td><td>{np.percentile(d,50):.2f}</td><td>{np.percentile(d,90):.2f}</td><td>{np.percentile(d,95):.2f}</td></tr>"
    
    corr_rows = ""
    for v in ["GSI","sigma_ci","RMR","Q_prime","D_geo","sigma_v","RQD"]:
        cr_val = np.corrcoef(mc[v], mc["d_total_mc"])[0,1]
        color = "#c0392b" if cr_val < -0.3 else "#27ae60" if cr_val > 0.3 else "#555"
        corr_rows += f'<tr><td>{v}</td><td style="color:{color};font-weight:600">{cr_val:.4f}</td></tr>'
    
    param_rows = ""
    cs = None
    for k, inf in PARAM_INFO.items():
        if k == "N_sim": continue
        if inf["section"] != cs:
            cs = inf["section"]
            param_rows += f'<tr class="sec"><td colspan="3">{cs}</td></tr>'
        param_rows += f'<tr><td>{inf["label"]}</td><td style="text-align:center;font-weight:600">{p[k]:.4g}</td><td>{inf["unit"]}</td></tr>'
    
    def ibox(title, txt, rk):
        bg, fg, bc = RC[rk]
        return f'<div style="background:{bg};border-left:4px solid {bc};padding:10px 14px;margin-bottom:8px;border-radius:0 6px 6px 0;"><div style="font-weight:700;color:{fg};font-size:0.85rem;margin-bottom:3px;">{title} — {RL[rk]}</div><div style="color:#333;font-size:0.82rem;line-height:1.45;">{txt}</div></div>'
    
    worst = "green"
    order = {"green":0,"yellow":1,"orange":2,"red":3}
    for rk in [rk_k, rk_h, rk_m, rk_kr, rk_d]:
        if order[rk] > order[worst]: worst = rk
    bg_ov, fg_ov, bc_ov = RC[worst]
    
    if worst == "red":
        concl = '<li>Condiciones actuales presentan riesgo <strong>CRÍTICO</strong>. Se requiere intervención inmediata.</li><li>Recomendar conformar comité técnico para evaluar rediseño del patrón o cambio de método de explotación.</li><li>No avanzar sin sostenimiento intensivo previo al disparo.</li>'
    elif worst == "orange":
        concl = '<li>Riesgos <strong>SEVEROS</strong> que requieren atención prioritaria. Diseño marginalmente aceptable.</li><li>Optimizar carga específica de voladura y reforzar sostenimiento preventivo.</li><li>Implementar monitoreo de convergencia y fragmentación.</li>'
    elif worst == "yellow":
        concl = '<li>Condiciones <strong>MODERADAS</strong> con margen de mejora. Diseño operacionalmente viable.</li><li>Optimizar geometría de malla (burden/espaciado) y verificar secuencia de disparos.</li><li>Establecer controles periódicos de desviación como protocolo estándar.</li>'
    else:
        concl = '<li>Condiciones <strong>FAVORABLES</strong>. Diseño de malla y sostenimiento adecuado.</li><li>Mantener parámetros actuales como línea base operacional.</li><li>Implementar monitoreo visual periódico estándar.</li>'
    
    html_str = f'''<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8">
<style>
@page {{ size:A4; margin:20mm 18mm 25mm 18mm; }}
@page cover {{ margin:0; }}
@page:first {{ @top-center {{ content:none; }} @bottom-center {{ content:none; }} }}
@top-center {{ content:"INFORME DE SIMULACIÓN — {op_name}"; font-size:7pt; color:#888; font-family:Helvetica,Arial,sans-serif; }}
@bottom-center {{ content:"Página " counter(page) " de " counter(pages); font-size:7pt; color:#888; font-family:Helvetica,Arial,sans-serif; }}
* {{ margin:0; padding:0; box-sizing:border-box; }}
body {{ font-family:Helvetica,Arial,sans-serif; font-size:9.5pt; color:#222; line-height:1.5; }}
.cover {{ page:cover; page-break-after:always; width:210mm; height:297mm; background:linear-gradient(160deg,#0d1b2a 0%,#1b2d45 40%,#1a3a5c 100%); color:white; text-align:center; display:flex; flex-direction:column; justify-content:center; align-items:center; }}
.cover-line {{ width:80mm; height:1px; background:rgba(255,255,255,0.3); margin:18px auto; }}
.cover h1 {{ font-size:26pt; font-weight:300; letter-spacing:3px; text-transform:uppercase; }}
.cover h2 {{ font-size:14pt; font-weight:400; color:#7eb8da; margin-top:6px; }}
.cover .sub {{ font-size:10pt; color:#8899aa; margin-top:30px; }}
.cover .meta {{ font-size:9pt; color:#667788; margin-top:40px; line-height:2; }}
.cover .conf {{ font-size:8pt; color:#e74c3c; margin-top:50px; letter-spacing:4px; font-weight:600; }}
.section-title {{ background:#1a2744; color:white; padding:8px 14px; font-size:11pt; font-weight:600; margin:22px 0 12px 0; border-radius:4px; letter-spacing:0.5px; }}
table {{ width:100%; border-collapse:collapse; margin:8px 0; font-size:8.5pt; }}
th {{ background:#1a2744; color:white; padding:7px 10px; text-align:left; font-weight:600; font-size:8pt; }}
td {{ padding:6px 10px; border-bottom:1px solid #ddd; }}
tr:nth-child(even) td {{ background:#f7f9fb; }}
tr.sec td {{ background:#e8edf2; font-weight:700; font-size:8.5pt; color:#1a2744; border-bottom:2px solid #1a2744; }}
.kpi-table {{ border-collapse:separate; border-spacing:8px; margin:12px 0; }}
.kpi-table td {{ background:#f7f9fb; border:1px solid #ddd; border-radius:8px; padding:14px 10px; text-align:center; vertical-align:top; width:25%; }}
.kpi-table .kpi-label {{ font-size:7.5pt; color:#666; text-transform:uppercase; letter-spacing:0.5px; }}
.kpi-table .kpi-val {{ font-size:22pt; font-weight:700; margin:4px 0; }}
.kpi-table .kpi-unit {{ font-size:7.5pt; color:#888; }}
.kpi-table .kpi-tag {{ font-size:7.5pt; font-weight:700; margin-top:6px; letter-spacing:1px; }}
.fig {{ text-align:center; margin:10px 0; page-break-inside:avoid; }}
.fig img {{ max-width:100%; border:1px solid #ddd; border-radius:4px; }}
.fig-caption {{ font-size:7.5pt; color:#888; font-style:italic; margin-top:4px; }}
.page-break {{ page-break-before:always; }}
.disclaimer {{ font-size:7pt; color:#999; border-top:1px solid #ddd; padding-top:10px; margin-top:30px; line-height:1.6; font-style:italic; }}
.concl-list {{ padding-left:18px; margin:10px 0; }}
.concl-list li {{ margin-bottom:8px; line-height:1.5; }}
</style></head><body>

<div class="cover">
    <h1>Simulación de Dilución</h1>
    <div class="cover-line"></div>
    <h2>Taladros Largos — Simba S7D</h2>
    <div class="sub">{op_name}</div>
    <div class="meta">
        Fecha: {fecha}<br>
        Preparado por: {author}<br>
        Simulaciones Monte Carlo: {int(p['N_sim']):,}
    </div>
    <div class="conf">CONFIDENCIAL</div>
</div>

<div class="section-title">1. RESUMEN EJECUTIVO</div>
<table class="kpi-table"><tr>
    <td style="border-left:4px solid {RC[rk_d][2]}">
        <div class="kpi-label">Dilución Total</div>
        <div class="kpi-val" style="color:{RC[rk_d][1]}">{dr['dilution']['d_total']:.2f}%</div>
        <div class="kpi-unit">Geom: {p['D_geo']:.1f}% | Peri: {dr['peri']['d_peri']:.2f}% | Est: {dr['laubscher']['d_est']:.2f}%</div>
        <div class="kpi-tag" style="color:{RC[rk_d][2]}">{RL[rk_d]}</div>
    </td>
    <td style="border-left:4px solid {RC[rk_kr][2]}">
        <div class="kpi-label">Fragmentación x50</div>
        <div class="kpi-val" style="color:{RC[rk_kr][1]}">{dr['kuzram']['x50']:.1f}</div>
        <div class="kpi-unit">mm — Kuz-Ram</div>
        <div class="kpi-tag" style="color:{RC[rk_kr][2]}">{RL[rk_kr]}</div>
    </td>
    <td style="border-left:4px solid {RC[rk_m][2]}">
        <div class="kpi-label">N Prime Mathews</div>
        <div class="kpi-val" style="color:{RC[rk_m][1]}">{dr['mathews']['N_prime']:.2f}</div>
        <div class="kpi-unit">RH = {dr['mathews']['hydraulic_radius']:.2f} m</div>
        <div class="kpi-tag" style="color:{RC[rk_m][2]}">{RL[rk_m]}</div>
    </td>
    <td style="border-left:4px solid {RC[rk_d][2]}">
        <div class="kpi-label">Tonelaje Recuperado</div>
        <div class="kpi-val" style="color:{RC[rk_d][1]}">{dr['dilution']['T_rec']:.1f}%</div>
        <div class="kpi-unit">de mineral in-situ</div>
        <div class="kpi-tag" style="color:{RC[rk_d][2]}">{RL[rk_d]}</div>
    </td>
</tr></table>

<div style="background:{RC[worst][0]};border:1px solid {RC[worst][2]};border-radius:6px;padding:12px 16px;margin:14px 0;">
    <div style="font-weight:700;font-size:10pt;color:{RC[worst][1]};">EVALUACION GLOBAL: {RL[worst]}</div>
    <div style="font-size:8.5pt;color:#333;margin-top:4px;">La evaluacion integral clasifica el escenario como <strong>{RL[worst].lower()}</strong>. {'Se requiere accion correctiva inmediata.' if worst in ('red','orange') else 'Diseno viable con ajustes menores.' if worst=='yellow' else 'Diseno adecuado para condiciones actuales.'}</div>
</div>

<div class="section-title page-break">2. PARAMETROS DE DISENO</div>
<table><tr><th>Parametro</th><th style="text-align:center">Valor</th><th>Unidad</th></tr>
{param_rows}</table>

<div class="section-title page-break">3. RESULTADOS DETERMINISTAS</div>
<table>
<tr class="sec"><td colspan="3">Geometria del Taladro</td></tr>
<tr><td>Longitud real L</td><td style="text-align:center;font-weight:600">{dr['geom']['L']:.3f}</td><td>m</td></tr>
<tr><td>Inclinacion alfa</td><td style="text-align:center;font-weight:600">{dr['geom']['alpha']:.2f}</td><td>grados</td></tr>
<tr><td>Proyeccion efectiva P</td><td style="text-align:center;font-weight:600">{dr['geom']['P_efectiva']:.3f}</td><td>m</td></tr>
<tr class="sec"><td colspan="3">Diseno de Voladura</td></tr>
<tr><td>Carga especifica q</td><td style="text-align:center;font-weight:600">{dr['blast']['q']:.3f}</td><td>kg/m3</td></tr>
<tr><td>Carga lineal corregida</td><td style="text-align:center;font-weight:600">{dr['blast']['pl_corr']:.3f}</td><td>kg/m</td></tr>
<tr><td>Carga total por taladro</td><td style="text-align:center;font-weight:600">{dr['blast']['Q_taladro']:.2f}</td><td>kg</td></tr>
<tr class="sec"><td colspan="3">Geomecanica</td></tr>
<tr><td>Esfuerzo tangencial collar</td><td style="text-align:center;font-weight:600">{dr['kirsch']['sigma_theta']:.2f}</td><td>MPa</td></tr>
<tr><td>Resistencia macizo sigma_cm</td><td style="text-align:center;font-weight:600">{dr['hoek']['sigma_cm']:.2f}</td><td>MPa</td></tr>
<tr><td>MRMR (Laubscher)</td><td style="text-align:center;font-weight:600">{dr['laubscher']['mrmr']:.1f}</td><td>[-]</td></tr>
<tr class="sec"><td colspan="3">Desglose de Dilucion</td></tr>
<tr><td>Dilucion geometrica</td><td style="text-align:center;font-weight:600">{p['D_geo']:.2f}</td><td>%</td></tr>
<tr><td>Dilucion perimetral</td><td style="text-align:center;font-weight:600">{dr['peri']['d_peri']:.3f}</td><td>%</td></tr>
<tr><td>Dilucion estructural</td><td style="text-align:center;font-weight:600">{dr['laubscher']['d_est']:.3f}</td><td>%</td></tr>
<tr><td style="font-weight:700;background:#e8edf2;">DILUCION TOTAL</td><td style="text-align:center;font-weight:700;font-size:11pt;background:#e8edf2;color:{RC[rk_d][1]}">{dr['dilution']['d_total']:.2f}</td><td style="font-weight:700;background:#e8edf2;">%</td></tr>
<tr><td style="font-weight:700;background:#e8edf2;">TONELAJE RECUPERADO</td><td style="text-align:center;font-weight:700;font-size:11pt;background:#e8edf2;color:{RC[rk_d][1]}">{dr['dilution']['T_rec']:.2f}</td><td style="font-weight:700;background:#e8edf2;">%</td></tr>
</table>

<div class="section-title page-break">4. INTERPRETACION DE RESULTADOS</div>
{ibox("Esfuerzos en Collar (Kirsch)", txt_k, rk_k)}
{ibox("Desviacion de Taladros (Holmberg)", txt_h, rk_h)}
{ibox("Estabilidad de Hastiales (Mathews-Potvin)", txt_m, rk_m)}
{ibox("Fragmentacion (Kuz-Ram)", txt_kr, rk_kr)}
{ibox("Dilucion Total y Recuperacion", txt_d, rk_d)}

<div class="section-title page-break">5. ANALISIS PROBABILISTICO — MONTE CARLO ({int(p['N_sim']):,} simulaciones)</div>
<div class="fig"><img src="data:image/png;base64,{imgs['hist']}" /><div class="fig-caption">Figura 1: Distribuciones probabilisticas y CDF.</div></div>

<div class="page-break"></div>
<div class="fig"><img src="data:image/png;base64,{imgs['tornado']}" /><div class="fig-caption">Figura 2: Diagrama de Tornado — Sensibilidad de parametros.</div></div>
<div style="height:10mm;"></div>
<div class="fig"><img src="data:image/png;base64,{imgs['scatter']}" /><div class="fig-caption">Figura 3: Diagramas de dispersion GSI vs Dilucion y Resistencia vs Fragmentacion.</div></div>

<div class="section-title page-break">6. ESTADISTICAS DE SIMULACION</div>
<table><tr><th>Variable</th><th>Media</th><th>Desv. Est.</th><th>P5</th><th>P50</th><th>P90</th><th>P95</th></tr>
{mc_rows}</table>

<div style="margin-top:16px;"><div style="font-weight:600;font-size:9.5pt;margin-bottom:6px;">Correlaciones de Pearson vs Dilucion Total</div>
<table><tr><th>Parametro</th><th>Correlacion</th></tr>{corr_rows}</table></div>

<div class="section-title page-break">7. CONCLUSIONES Y RECOMENDACIONES</div>
<ol class="concl-list">
{concl}
</ol>

<div class="disclaimer">
DESCARGO DE RESPONSABILIDAD: Este informe fue generado automaticamente mediante un modelo de simulacion computacional. Los resultados son indicativos y deben ser validados por un ingeniero geomecanico competente antes de su uso en decisiones operacionales. Las correlaciones empiricas utilizadas (Kuz-Ram, Holmberg, Mathews-Potvin, Laubscher) tienen rangos de aplicabilidad definidos que deben ser verificados. La simulacion Monte Carlo asume distribuciones normales truncadas independientes. Este documento es CONFIDENCIAL y su distribucion esta restringida al personal autorizado de {op_name}.
</div>

</body></html>'''
    
    return HTML(string=html_str).write_pdf()


# ═══════════════════════════════════════════════════════════════════════
#  HELPERS DE INTERFAZ
# ═══════════════════════════════════════════════════════════════════════

def _step(d):
    if d == 0: return 0.01
    m = 10**np.floor(np.log10(abs(d)))
    return float(m / 10)

def _bounds(inf):
    u = inf["unit"]; d = inf["default"]
    if u in ("[%]", "[0-100]"): return 0.0, 100.0
    if u == "[0-1]": return 0.0, 1.0
    if u == "°": return 0.0, 90.0
    if u == "m" and d < 0.1: return 0.001, 0.5
    return 0.0, None

RISK_C = {"green": "#3fb950", "yellow": "#d29922", "orange": "#f78166", "red": "#f85149"}
RISK_L = {"green": "FAVORABLE", "yellow": "MODERADO", "orange": "SEVERO", "red": "CRÍTICO"}

def _kpi(label, val, unit, ck):
    c = RISK_C[ck]; t = RISK_L[ck]
    return f'<div style="background:#161b22;border:1px solid #30363d;border-left:4px solid {c};border-radius:8px;padding:1.2rem 1rem;text-align:center;"><div style="color:#8b949e;font-size:0.8rem;margin-bottom:0.4rem;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">{label}</div><div style="color:#e6edf3;font-size:1.8rem;font-weight:700;">{val}</div><div style="color:#8b949e;font-size:0.75rem;margin-top:0.15rem;">{unit}</div><div style="color:{c};font-size:0.7rem;font-weight:600;margin-top:0.35rem;letter-spacing:0.05em;">{t}</div></div>'

def _ibox_s(title, icon, txt, ck):
    c = RISK_C[ck]
    return f'<div style="background:#161b22;border:1px solid #30363d;border-left:4px solid {c};border-radius:8px;padding:1rem 1.2rem;margin-bottom:0.6rem;"><div style="font-weight:700;font-size:0.95rem;margin-bottom:0.4rem;">{icon} {title}</div><div style="color:#c9d1d9;font-size:0.85rem;line-height:1.5;">{txt}</div></div>'


# ═══════════════════════════════════════════════════════════════════════
#  APLICACIÓN PRINCIPAL
# ═══════════════════════════════════════════════════════════════════════

def main():
    st.markdown("""<style>
    .block-container{padding-top:1.5rem;max-width:1420px;}
    section[data-testid="stSidebar"]{background:#0d1117;border-right:1px solid #21262d;}
    section[data-testid="stSidebar"] .stMarkdown{color:#c9d1d9;}
    .stTabs [data-baseweb="tab-list"]{gap:2px;background:#0d1117;border-radius:8px;padding:4px;}
    .stTabs [data-baseweb="tab"]{background:#161b22;color:#8b949e;border-radius:6px;padding:8px 16px;font-size:0.85rem;}
    .stTabs [aria-selected="true"]{background:#21262d;color:#58a6ff;font-weight:600;}
    .stDownloadButton>button{background:#238636;color:#fff;border:none;border-radius:6px;font-weight:600;padding:0.6rem 1.5rem;}
    .stDownloadButton>button:hover{background:#2ea043;}
    h1,h2,h3{color:#e6edf3 !important;}
    </style>""", unsafe_allow_html=True)

    # ── Sidebar ──
    with st.sidebar:
        st.markdown("### ⛏️ Simulador de Dilución")
        st.markdown("---")
        op_name = st.text_input("Nombre de la Operación", value="Operación Minera", key="op_name")
        author = st.text_input("Preparado por", value="Ing. Geomecánica", key="author")
        st.markdown("---")
        
        if st.button("Restaurar Valores por Defecto", use_container_width=True):
            for k in PARAM_INFO:
                if f"p_{k}" in st.session_state:
                    del st.session_state[f"p_{k}"]
            st.rerun()
            
        st.markdown("---")
        
        params = {}
        cs = None
        for k, inf in PARAM_INFO.items():
            if inf["section"] != cs:
                cs = inf["section"]
                st.markdown(f"**{cs}**")
            
            mn, mx = _bounds(inf)
            if k == "N_sim":
                params[k] = st.selectbox(
                    inf["label"], [1000, 5000, 10000, 25000, 50000],
                    index=2, help=inf["hint"], key=f"p_{k}"
                )
            else:
                ct = f"  _(CoV: {inf['std_frac']*100:.0f}%)_" if inf["std_frac"] > 0 else ""
                params[k] = st.number_input(
                    f"{inf['label']} [{inf['unit']}]{ct}",
                    value=inf["default"], min_value=mn, max_value=mx,
                    step=_step(inf["default"]), help=inf["hint"], key=f"p_{k}",
                    format="%.4f" if _step(inf["default"]) < 0.01 else "%.3f" if _step(inf["default"]) < 0.1 else "%g"
                )
        
        st.markdown("---")
        run_btn = st.button("EJECUTAR SIMULACIÓN", type="primary", use_container_width=True)

    # ── Ejecutar cálculos ──
    if run_btn:
        with st.spinner("Modelo determinístico..."):
            g = calc_geometry(params)
            b = calc_blasting(params, g)
            h = calc_holmberg(params, g)
            ki = calc_kirsch(params)
            hb = calc_hoek_brown(params)
            mp = calc_mathews_potvin(params)
            lb = calc_laubscher(params)
            pe = calc_perimetral_dilation(params, b)
            kr = calc_kuzram(params, b, g)
            dl = calc_total_dilation(params, pe["d_peri"], lb["d_est"])
            
            dr = {
                "geom": g, "blast": b, "holmberg": h, "kirsch": ki, "hoek": hb,
                "mathews": mp, "laubscher": lb, "peri": pe, "kuzram": kr, "dilution": dl
            }
            
        with st.spinner("Monte Carlo..."):
            mc = monte_carlo_vectorized(params)
            
        with st.spinner("Generando reporte PDF..."):
            pdf_bytes = generate_pdf(params, dr, mc, op_name, author)
            
        st.session_state.update({
            "det_res": dr, "mc_res": mc, "params": params,
            "calculated": True, "pdf_bytes": pdf_bytes
        })

    # ── Pantalla de bienvenida ──
    if not st.session_state.get("calculated"):
        st.markdown("""<div style="text-align:center;padding:4rem 2rem;">
        <div style="font-size:3rem;margin-bottom:1rem;">⛏️</div>
        <h1 style="font-size:2rem;margin-bottom:0.5rem;">Simulador de Dilución — Taladros Largos</h1>
        <h2 style="font-size:1.1rem;color:#8b949e;font-weight:400;margin-bottom:2rem;">Análisis determinístico + Monte Carlo | PDF gerencial incluido</h2>
        <div style="max-width:640px;margin:0 auto;text-align:left;background:#161b22;border:1px solid #30363d;border-radius:10px;padding:1.5rem 2rem;color:#c9d1d9;font-size:0.9rem;line-height:1.7;">
        <p><strong style="color:#58a6ff;">Instrucciones:</strong></p>
        <ol style="padding-left:1.2rem;margin:0.5rem 0;">
        <li>Configure parámetros en la barra lateral.</li>
        <li>Indique <strong>nombre de la operación</strong> y <strong>autor</strong> (aparecen en el PDF).</li>
        <li>Presione <strong>EJECUTAR SIMULACIÓN</strong>.</li>
        <li>Revise resultados en pantalla y descargue el <strong>PDF gerencial</strong> o el <strong>Excel</strong>.</li></ol>
        <p style="color:#8b949e;font-size:0.8rem;margin-bottom:0;">Módulos: Kirsch · Holmberg · Hoek-Brown · Mathews-Potvin · Laubscher · Kuz-Ram</p>
        </div></div>""", unsafe_allow_html=True)
        return

    # ── Recuperar resultados ──
    dr = st.session_state["det_res"]
    mc = st.session_state["mc_res"]
    p = st.session_state["params"]
    
    st.markdown(f"""<div style="margin-bottom:1.5rem;">
    <h1 style="margin:0 0 0.2rem 0;">Resultados de la Simulación</h1>
    <p style="color:#8b949e;margin:0;font-size:0.9rem;">{op_name} — {int(p['N_sim']):,} simulaciones Monte Carlo</p></div>""", unsafe_allow_html=True)

    # ── Interpretaciones para KPIs ──
    txt_k, rk_k = interpret_kirsch(dr["kirsch"]["sigma_theta"], p["sigma_ci"])
    txt_h, rk_h = interpret_holmberg(dr["holmberg"]["deviation"], dr["geom"]["alpha"])
    txt_m, rk_m = interpret_mathews(dr["mathews"]["N_prime"], dr["mathews"]["hydraulic_radius"])
    txt_kr, rk_kr = interpret_kuzram(dr["kuzram"]["x50"])
    txt_d, rk_d = interpret_dilution(dr["dilution"]["d_total"], p["D_geo"], dr["peri"]["d_peri"], dr["laubscher"]["d_est"], dr["laubscher"]["mrmr"])

    # ── KPI Cards ──
    kpis = [
        ("Dilución Total", f"{dr['dilution']['d_total']:.2f} %", "%", rk_d),
        ("Fragmentación x50", f"{dr['kuzram']['x50']:.1f}", "mm", rk_kr),
        ("N' Mathews-Potvin", f"{dr['mathews']['N_prime']:.2f}", "[-]", rk_m),
        ("Tonelaje Recuperado", f"{dr['dilution']['T_rec']:.1f} %", "%", rk_d)
    ]
    cols_kpi = st.columns(4)
    for c, (l, v, u, ck) in zip(cols_kpi, kpis):
        c.markdown(_kpi(l, v, u, ck), unsafe_allow_html=True)
        
    st.markdown("<div style='height:1.2rem;'></div>", unsafe_allow_html=True)

    # ── Interpretaciones ──
    st.markdown("### Interpretación Ingenieril")
    st.markdown(_ibox_s("Esfuerzos Collar (Kirsch)", "🔵", txt_k, rk_k), unsafe_allow_html=True)
    st.markdown(_ibox_s("Desviación Taladros (Holmberg)", "📐", txt_h, rk_h), unsafe_allow_html=True)
    st.markdown(_ibox_s("Estabilidad Hastiales (Mathews-Potvin)", "⛰️", txt_m, rk_m), unsafe_allow_html=True)
    st.markdown(_ibox_s("Fragmentación (Kuz-Ram)", "💥", txt_kr, rk_kr), unsafe_allow_html=True)
    st.markdown(_ibox_s("Dilución y Recuperación", "📊", txt_d, rk_d), unsafe_allow_html=True)
    
    st.markdown("<div style='height:1rem;'></div>", unsafe_allow_html=True)

    # ── Tabs ──
    t1, t2, t3, t4 = st.tabs(["📊 Cálculos Detallados", "🎲 Monte Carlo", "📈 Sensibilidad", "💾 Exportar Reportes"])

    with t1:
        cl, cr2 = st.columns([1, 1])
        with cl:
            st.markdown("#### Resumen de Cálculos")
            calc_list = [
                ("GEOMETRÍA", None, None), ("Longitud real L", f"{dr['geom']['L']:.3f}", "m"),
                ("Inclinación alfa", f"{dr['geom']['alpha']:.2f}", "°"),
                ("Proyección efectiva P", f"{dr['geom']['P_efectiva']:.3f}", "m"),
                ("VOLADURA", None, None), ("Factor Langefors", f"{dr['blast']['f_lang']:.3f}", "[-]"),
                ("Espaciado técnico S_tec", f"{dr['blast']['S_tecnico']:.3f}", "m"),
                ("Carga específica q", f"{dr['blast']['q']:.3f}", "kg/m³"),
                ("Carga lineal corregida", f"{dr['blast']['pl_corr']:.3f}", "kg/m"),
                ("Carga total Q", f"{dr['blast']['Q_taladro']:.2f}", "kg"),
                ("GEOMECÁNICA", None, None), ("Esfuerzo tangencial", f"{dr['kirsch']['sigma_theta']:.2f}", "MPa"),
                ("mb Hoek-Brown", f"{dr['hoek']['mb']:.4f}", "[-]"),
                ("s Hoek-Brown", f"{dr['hoek']['s']:.6f}", "[-]"),
                ("sigma_cm", f"{dr['hoek']['sigma_cm']:.2f}", "MPa"),
                ("DILUCIÓN", None, None), ("Geométrica", f"{p['D_geo']:.2f}", "%"),
                ("Perimetral", f"{dr['peri']['d_peri']:.3f}", "%"),
                ("Estructural", f"{dr['laubscher']['d_est']:.3f}", "%"),
                ("MRMR", f"{dr['laubscher']['mrmr']:.1f}", "[-]"),
                ("DILUCIÓN TOTAL", f"{dr['dilution']['d_total']:.2f}", "%"),
                ("TONELAJE RECUPERADO", f"{dr['dilution']['T_rec']:.2f}", "%")
            ]
            for v, val, u in calc_list:
                if val is None:
                    st.markdown(f"**{v}**")
                else:
                    st.markdown(f"<span style='color:#8b949e'>{v}:</span>  **{val}**  `{u}`", unsafe_allow_html=True)
        with cr2:
            st.markdown("#### Geometría Polar del Taladro")
            fp = plot_polar(p)
            st.pyplot(fp, use_container_width=True)
            plt.close(fp)

    with t2:
        st.markdown("#### Distribuciones Probabilísticas + CDF")
        fh = plot_mc_histograms(mc)
        st.pyplot(fh, use_container_width=True)
        plt.close(fh)

        st.markdown("#### Estadísticas de Simulación")
        sd = []
        for nm, k in [("Dilución Total (%)", "d_total_mc"), ("x50 (mm)", "x50_mc"), ("N' Mathews", "N_prime_mc"), ("Desviación (m)", "deviation_mc"), ("MRMR", "mrmr_mc")]:
            d = mc[k]
            sd.append({
                "Variable": nm, "Media": f"{np.mean(d):.2f}", "Desv.Est.": f"{np.std(d):.2f}",
                "P5": f"{np.percentile(d, 5):.2f}", "P50": f"{np.percentile(d, 50):.2f}",
                "P90": f"{np.percentile(d, 90):.2f}", "P95": f"{np.percentile(d, 95):.2f}"
            })
        st.dataframe(sd, use_container_width=True, hide_index=True)

    with t3:
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("#### Diagrama de Tornado")
            ft = plot_tornado(mc)
            st.pyplot(ft, use_container_width=True)
            plt.close(ft)
        with c2:
            st.markdown("#### Diagramas de Dispersión")
            fs = plot_scatters(mc)
            st.pyplot(fs, use_container_width=True)
            plt.close(fs)

    with t4:
        st.markdown("### Exportar Reportes")
        cp, cx = st.columns(2)
        
        with cp:
            st.markdown("""<div style="background:#161b22;border:1px solid #30363d;border-radius:10px;padding:1.5rem;">
            <div style="font-size:2rem;margin-bottom:0.5rem;">📄</div>
            <div style="font-size:1.1rem;font-weight:700;color:#e6edf3;margin-bottom:0.3rem;">Reporte PDF — Presentación Gerencial</div>
            <div style="color:#8b949e;font-size:0.82rem;line-height:1.6;margin-bottom:1rem;">
            Portada corporativa, Resumen ejecutivo, Tablas, Gráficos integrados, Sensibilidad, Conclusiones.</div></div>""", unsafe_allow_html=True)
            
            if st.session_state.get("pdf_bytes") is not None:
                st.download_button(
                    "DESCARGAR PDF GERENCIAL", st.session_state["pdf_bytes"],
                    file_name=f"Informe_Dilucion_{op_name.replace(' ', '_')}.pdf",
                    mime="application/pdf", use_container_width=True
                )
            else:
                st.error("**WeasyPrint no está instalado.** No se puede generar el PDF.\n\n"
                         "Para instalar en Windows: `pip install weasyprint`\n\n"
                         "Luego instalar GTK3 desde el enlace oficial de WeasyPrint.\n\n"
                         "Para macOS: `brew install weasyprint`\n"
                         "Para Linux (Ubuntu): `sudo apt install weasyprint`")
        
        with cx:
            st.markdown("""<div style="background:#161b22;border:1px solid #30363d;border-radius:10px;padding:1.5rem;">
            <div style="font-size:2rem;margin-bottom:0.5rem;">📊</div>
            <div style="font-size:1.1rem;font-weight:700;color:#e6edf3;margin-bottom:0.3rem;">Archivo Excel — Datos Completos</div>
            <div style="color:#8b949e;font-size:0.82rem;line-height:1.6;margin-bottom:1rem;">
            5 hojas: Datos de entrada, Interpretaciones, Cálculos detallados, Estadísticas MC, Muestra de 500 iteraciones.</div></div>""", unsafe_allow_html=True)
            
            ebuf = export_excel_buffer(p, dr, mc)
            st.download_button(
                "DESCARGAR EXCEL", ebuf, file_name="Resultados_Simba_S7D.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

if __name__ == "__main__":
    main()