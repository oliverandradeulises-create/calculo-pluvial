import io
import math
import re
import matplotlib.pyplot as plt
from matplotlib.patches import PathPatch, Rectangle
from matplotlib.path import Path
import numpy as np
import openpyxl
import requests
import streamlit as st
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Inches, Pt, RGBColor

# Configuración de la página Web
st.set_page_config(
    page_title="Generador de Memorias Pluviales",
    page_icon="🌧️",
    layout="wide",
)

# Ajuste CSS para reducir el tamaño de fuente en métricas y evitar texto truncado (...)
st.markdown(
    """
    <style>
    [data-testid="stMetricValue"] {
        font-size: 1.20rem !important;
        white-space: normal !important;
        overflow: visible !important;
        text-overflow: clip !important;
        line-height: 1.25 !important;
        word-break: break-word !important;
    }
    [data-testid="stMetricLabel"] {
        font-size: 0.86rem !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ==============================================================================
# TABLA DE COEFICIENTES DE ESCURRIMIENTO (C) SEGÚN TR
# ==============================================================================
TABLA_COEFICIENTES_C = {
    2: {"Azotea / Techumbre": 0.80, "Patios / Vialidad": 0.87},
    5: {"Azotea / Techumbre": 0.85, "Patios / Vialidad": 0.88},
    10: {"Azotea / Techumbre": 0.90, "Patios / Vialidad": 0.90},
    20: {"Azotea / Techumbre": 0.90, "Patios / Vialidad": 0.90},
    25: {"Azotea / Techumbre": 0.90, "Patios / Vialidad": 0.90},
    50: {"Azotea / Techumbre": 0.90, "Patios / Vialidad": 0.92},
    100: {"Azotea / Techumbre": 0.90, "Patios / Vialidad": 0.97},
    200: {"Azotea / Techumbre": 0.90, "Patios / Vialidad": 1.00},
    500: {"Azotea / Techumbre": 0.90, "Patios / Vialidad": 1.00},
}

# ==============================================================================
# CATÁLOGOS DE MATERIALES (Rugosidad y Secciones Comerciales)
# ==============================================================================
CATALOGO_MATERIALES = {
    "PVC Alcantarillado S-20": {
        "n": 0.009,
        "v_max": 5.0,
        "v_min": 0.60,
        "tuberias": [
            {"nom_in": '160 mm (6.3")', "d_int_mm": 152.20},
            {"nom_in": '200 mm (7.9")', "d_int_mm": 190.20},
            {"nom_in": '250 mm (9.8")', "d_int_mm": 237.80},
            {"nom_in": '315 mm (12.4")', "d_int_mm": 299.60},
            {"nom_in": '355 mm (14.0")', "d_int_mm": 337.60},
            {"nom_in": '400 mm (15.7")', "d_int_mm": 380.40},
            {"nom_in": '450 mm (17.7")', "d_int_mm": 428.00},
            {"nom_in": '500 mm (19.7")', "d_int_mm": 475.60},
            {"nom_in": '630 mm (24.8")', "d_int_mm": 599.20},
            {"nom_in": '800 mm (31.5")', "d_int_mm": 761.00},
        ],
    },
    "PEAD Corrugado": {
        "n": 0.010,
        "v_max": 5.0,
        "v_min": 0.60,
        "tuberias": [
            {"nom_in": '6"', "d_int_mm": 145.00},
            {"nom_in": '8"', "d_int_mm": 195.00},
            {"nom_in": '10"', "d_int_mm": 245.00},
            {"nom_in": '12"', "d_int_mm": 294.00},
            {"nom_in": '15"', "d_int_mm": 369.00},
            {"nom_in": '18"', "d_int_mm": 450.00},
            {"nom_in": '24"', "d_int_mm": 588.00},
            {"nom_in": '30"', "d_int_mm": 751.00},
            {"nom_in": '36"', "d_int_mm": 902.00},
            {"nom_in": '42"', "d_int_mm": 1051.00},
            {"nom_in": '48"', "d_int_mm": 1185.00},
            {"nom_in": '60"', "d_int_mm": 1501.00},
        ],
    },
    "PVC Sanitario Norma": {
        "n": 0.009,
        "v_max": 5.0,
        "v_min": 0.60,
        "tuberias": [
            {"nom_in": '2"', "d_int_mm": 46.40},
            {"nom_in": '3"', "d_int_mm": 71.40},
            {"nom_in": '4"', "d_int_mm": 105.40},
            {"nom_in": '6"', "d_int_mm": 153.40},
            {"nom_in": '8"', "d_int_mm": 192.00},
        ],
    },
    "Lámina Galvanizada": {
        "n": 0.012,
        "v_max": 4.5,
        "v_min": 0.50,
        "tuberias": [
            {"nom_in": "3 in x 3 in (76x76 mm)", "ancho_in": 3.0, "largo_in": 3.0, "d_int_mm": 76.20},
            {"nom_in": "3 in x 4 in (76x102 mm)", "ancho_in": 3.0, "largo_in": 4.0, "d_int_mm": 87.09},
            {"nom_in": "4 in x 4 in (102x102 mm)", "ancho_in": 4.0, "largo_in": 4.0, "d_int_mm": 101.60},
            {"nom_in": "4 in x 5 in (102x127 mm)", "ancho_in": 4.0, "largo_in": 5.0, "d_int_mm": 112.89},
            {"nom_in": "4 in x 6 in (102x152 mm)", "ancho_in": 4.0, "largo_in": 6.0, "d_int_mm": 121.92},
            {"nom_in": "5 in x 5 in (127x127 mm)", "ancho_in": 5.0, "largo_in": 5.0, "d_int_mm": 127.00},
            {"nom_in": "5 in x 6 in (127x152 mm)", "ancho_in": 5.0, "largo_in": 6.0, "d_int_mm": 138.55},
            {"nom_in": "6 in x 6 in (152x152 mm)", "ancho_in": 6.0, "largo_in": 6.0, "d_int_mm": 152.40},
            {"nom_in": "6 in x 8 in (152x203 mm)", "ancho_in": 6.0, "largo_in": 8.0, "d_int_mm": 174.17},
            {"nom_in": "8 in x 8 in (203x203 mm)", "ancho_in": 8.0, "largo_in": 8.0, "d_int_mm": 203.20},
        ],
    },
    "Concreto": {
        "n": 0.015,
        "v_max": 6.0,
        "v_min": 0.60,
        "tuberias": [],
    },
}

# ==============================================================================
# CATÁLOGO DE ESTACIONES POR DEFECTO
# ==============================================================================
CATALOGO_ESTACIONES = {
    "AGUASCALIENTES": {
        "lat": 21.8853,
        "lon": -102.2916,
        "estacion": "San Francisco de los Romo (DGE) - Aguascalientes",
        "I_10_5min": 206.49,
    },
    "PUEBLA": {
        "lat": 19.0414,
        "lon": -98.2063,
        "estacion": "Puebla (DGE) - Puebla",
        "I_10_5min": 214.69,
    },
    "CDMX": {
        "lat": 19.4326,
        "lon": -99.1332,
        "estacion": "Tacubaya / Central (DGE) - CDMX",
        "I_10_5min": 178.50,
    },
    "GUADALAJARA": {
        "lat": 20.6597,
        "lon": -103.3496,
        "estacion": "Guadalajara (DGE) - Jalisco",
        "I_10_5min": 210.30,
    },
    "MONTERREY": {
        "lat": 25.6866,
        "lon": -100.3161,
        "estacion": "Monterrey (DGE) - Nuevo León",
        "I_10_5min": 225.80,
    },
    "QUERETARO": {
        "lat": 20.5888,
        "lon": -100.3899,
        "estacion": "Querétaro (DGE) - Querétaro",
        "I_10_5min": 195.40,
    },
    "LEON": {
        "lat": 21.1221,
        "lon": -101.6826,
        "estacion": "León (DGE) - Guanajuato",
        "I_10_5min": 198.20,
    },
    "SAN_LUIS_POTOSI": {
        "lat": 22.1565,
        "lon": -100.9855,
        "estacion": "San Luis Potosí (DGE) - SLP",
        "I_10_5min": 188.60,
    },
}


@st.cache_data(show_spinner=False)
def obtener_estacion_mas_cercana(lat, lon):
    estacion_cercana = None
    distancia_minima_km = float("inf")
    R_tierra = 6371.0

    for clave, datos in CATALOGO_ESTACIONES.items():
        dlat = math.radians(datos["lat"] - lat)
        dlon = math.radians(datos["lon"] - lon)
        a = (
            math.sin(dlat / 2.0) ** 2
            + math.cos(math.radians(lat))
            * math.cos(math.radians(datos["lat"]))
            * math.sin(dlon / 2.0) ** 2
        )
        c = 2.0 * math.atan2(math.sqrt(a), math.sqrt(1.0 - a))
        dist_km = R_tierra * c

        if dist_km < distancia_minima_km:
            distancia_minima_km = dist_km
            estacion_cercana = datos

    nombre_con_dist = (
        f"{estacion_cercana['estacion']} (a {distancia_minima_km:.1f} km)"
    )
    return nombre_con_dist, estacion_cercana["I_10_5min"]


def calcular_tiempo_concentracion(
    longitud_m, diferencia_altura_m=None, pendiente_tramo_pct=None
):
    L = float(longitud_m)
    if L <= 0:
        return 5.0

    if diferencia_altura_m is not None and float(diferencia_altura_m) > 0:
        S = float(diferencia_altura_m) / L
    elif pendiente_tramo_pct is not None and float(pendiente_tramo_pct) > 0:
        S = float(pendiente_tramo_pct) / 100.0
    else:
        return 5.0

    if S <= 0:
        return 5.0

    tc_minutos = 0.000325 * (L**0.77) * (S**-0.385) * 60.0
    return round(tc_minutos, 2)


def aplicar_sombreado_celda(celda, color_hex):
    tcPr = celda._element.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    tcPr.append(shd)


def fijar_margenes_celda(celda, top=120, bottom=120, left=150, right=150):
    tcPr = celda._element.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for m_name, m_val in [
        ("top", top),
        ("bottom", bottom),
        ("left", left),
        ("right", right),
    ]:
        node = OxmlElement(f"w:{m_name}")
        node.set(qn("w:w"), str(m_val))
        node.set(qn("w:type"), "dxa")
        tcMar.append(node)
    tcPr.append(tcMar)


@st.cache_data(show_spinner=False)
def extraer_coordenadas(texto):
    texto = texto.strip()
    m_iframe = re.search(r"!2d(-?\d+\.\d+)!3d(-?\d+\.\d+)", texto)
    if m_iframe:
        return float(m_iframe.group(2)), float(m_iframe.group(1))

    m_directo = re.search(r"(-?\d{1,2}\.\d+)\s*,\s*(-?\d{2,3}\.\d+)", texto)
    if m_directo:
        lat, lon = float(m_directo.group(1)), float(m_directo.group(2))
        if -90 <= lat <= 90 and -180 <= lon <= 180:
            return lat, lon

    if "maps" in texto or "goo.gl" in texto:
        headers = {"User-Agent": "Mozilla/5.0"}
        try:
            session = requests.Session()
            session.cookies.set(
                "CONSENT",
                "YES+cb.20210328-17-p0.es+FX+667",
                domain=".google.com",
            )
            res = session.get(
                texto, headers=headers, allow_redirects=True, timeout=10
            )
            url_p, html_p = res.url, res.text
        except Exception:
            url_p, html_p = texto, ""

        m_pin = re.search(r"!3d(-?\d+\.\d+)!4d(-?\d+\.\d+)", url_p)
        if m_pin:
            return float(m_pin.group(1)), float(m_pin.group(2))

        m_cam = re.search(r"@(-?\d+\.\d+),(-?\d+\.\d+)", url_p)
        if m_cam:
            return float(m_cam.group(1)), float(m_cam.group(2))

        if html_p:
            m_arr = re.findall(
                r"\[\s*(-?\d{1,2}\.\d{4,15})\s*,\s*(-?\d{2,3}\.\d{4,15})\s*\]",
                html_p,
            )
            for lat_s, lon_s in m_arr:
                l1, l2 = float(lat_s), float(lon_s)
                if -90 <= l1 <= 90 and -180 <= l2 <= 180:
                    return l1, l2

    raise ValueError("Coordenadas no encontradas.")


@st.cache_data(show_spinner=False)
def procesar_excel_estacion(
    archivo_bytes, lat, lon, tr_seleccionado=10, tc_adoptado=5.0
):
    nombre_estacion_default, I_ref_default = obtener_estacion_mas_cercana(
        lat, lon
    )
    nombre_estacion = nombre_estacion_default
    I_ref = None
    tabla_tr = []

    factor_tc_adoptado = (
        (5.0 / tc_adoptado) ** 0.45 if tc_adoptado > 0 else 1.0
    )

    if archivo_bytes is not None:
        try:
            wb = openpyxl.load_workbook(
                io.BytesIO(archivo_bytes), data_only=True
            )
            if "Curvas IDTR" in wb.sheetnames:
                ws = wb["Curvas IDTR"]
            elif len(wb.sheetnames) >= 5:
                ws = wb.worksheets[4]
            else:
                ws = wb.active

            nom_val = ws.cell(row=4, column=5).value
            clave_val = ws.cell(row=5, column=5).value
            if nom_val:
                nombre_estacion = (
                    f"{nom_val} (Clave {clave_val})"
                    if clave_val
                    else str(nom_val)
                )

            duraciones_min = []
            duraciones_header = []
            for col in range(3, 12):
                val_dur = ws.cell(row=24, column=col).value
                if val_dur is not None:
                    d_min = float(val_dur)
                    duraciones_min.append(d_min)
                    duraciones_header.append(f"{int(d_min)}m")

            periodos_interes = [2, 5, 10, 20, 25, 50, 100, 200, 500]
            for row in range(25, 36):
                tr_val = ws.cell(row=row, column=2).value
                if tr_val is not None:
                    try:
                        tr_num = int(float(tr_val))
                        if tr_num in periodos_interes:
                            fila_dict = {"Tr": f"{tr_num} Años"}
                            valores_fila = []

                            for idx, col in enumerate(range(3, 12)):
                                val_i = ws.cell(row=row, column=col).value
                                if val_i is not None and idx < len(
                                    duraciones_header
                                ):
                                    val_float = float(val_i)
                                    valores_fila.append(val_float)
                                    col_name = duraciones_header[idx]
                                    fila_dict[col_name] = round(val_float, 2)

                            if tr_num == tr_seleccionado and valores_fila:
                                d_validas = duraciones_min[: len(valores_fila)]
                                if (
                                    d_validas
                                    and d_validas[0]
                                    <= tc_adoptado
                                    <= d_validas[-1]
                                ):
                                    I_interp = float(
                                        np.interp(
                                            tc_adoptado,
                                            d_validas,
                                            valores_fila,
                                        )
                                    )
                                    I_ref = round(I_interp, 2)
                                else:
                                    I_ref = round(
                                        valores_fila[0] * factor_tc_adoptado, 2
                                    )

                            tabla_tr.append(fila_dict)
                    except ValueError:
                        continue
        except Exception as e:
            st.warning(
                f"No se pudo procesar el formato del Excel cargado: {e}"
            )

    if not tabla_tr:
        duraciones_fijas = [5, 10, 15, 20, 25, 30, 35, 40, 45, 50]
        periodos_tr = [2, 5, 10, 20, 25, 50, 100, 200, 500]
        for Tr in periodos_tr:
            I_5min_tr = I_ref_default * ((Tr / 10.0) ** 0.22)
            if Tr == tr_seleccionado:
                I_ref = round(I_5min_tr * factor_tc_adoptado, 2)

            fila_sintetica = {"Tr": f"{Tr} Años"}
            for d in duraciones_fijas:
                fila_sintetica[f"{d}m"] = round(
                    I_5min_tr * ((5.0 / d) ** 0.45), 2
                )
            tabla_tr.append(fila_sintetica)

    if I_ref is None:
        I_ref = round(I_ref_default * factor_tc_adoptado, 2)

    return nombre_estacion, I_ref, tabla_tr


# ==============================================================================
# SOLUCIONADORES HIDRÁULICOS EXACTOS (BISECCIÓN) Y MÉTODO SMACNA
# ==============================================================================
def resolver_tirante_canal(Q_m3_s, b_m, talud_z, n_manning, S, y_max_m):
    """Calcula el tirante normal exacto (y_n) y velocidad real por bisección."""
    if Q_m3_s <= 0 or S <= 0:
        return 0.0, 0.0

    def caudal_manning(y):
        if y <= 0:
            return 0.0
        A = (b_m + talud_z * y) * y
        P = b_m + 2.0 * y * math.sqrt(1.0 + talud_z**2)
        Rh = A / P if P > 0 else 0.0
        return (1.0 / n_manning) * A * (Rh ** (2.0 / 3.0)) * (S**0.5)

    if caudal_manning(y_max_m) <= Q_m3_s:
        A_max = (b_m + talud_z * y_max_m) * y_max_m
        v_max_util = Q_m3_s / A_max if A_max > 0 else 0.0
        return y_max_m, v_max_util

    low, high = 0.0, y_max_m
    for _ in range(40):
        mid = (low + high) / 2.0
        if caudal_manning(mid) < Q_m3_s:
            low = mid
        else:
            high = mid

    y_real = (low + high) / 2.0
    A_real = (b_m + talud_z * y_real) * y_real
    v_real = Q_m3_s / A_real if A_real > 0 else 0.0
    return y_real, v_real


def resolver_tirante_circular(Q_m3_s, D_m, n_manning, S):
    """Calcula el ángulo theta, tirante real (m) y velocidad real (m/s) en tubería circular."""
    if Q_m3_s <= 0 or S <= 0 or D_m <= 0:
        return 0.0, 0.0, 0.0

    theta_max_q = 5.2781
    low, high = 0.0001, theta_max_q

    def caudal_parcial(th):
        A_p = (D_m**2 / 8.0) * (th - math.sin(th))
        P_p = (D_m / 2.0) * th
        Rh_p = A_p / P_p if P_p > 0 else 0.0
        return (1.0 / n_manning) * A_p * (Rh_p ** (2.0 / 3.0)) * (S**0.5)

    if caudal_parcial(theta_max_q) <= Q_m3_s:
        A_lleno = math.pi * (D_m**2) / 4.0
        return 2.0 * math.pi, D_m, (Q_m3_s / A_lleno)

    for _ in range(40):
        mid = (low + high) / 2.0
        if caudal_parcial(mid) < Q_m3_s:
            low = mid
        else:
            high = mid

    theta_sol = (low + high) / 2.0
    tirante_m = (D_m / 2.0) * (1.0 - math.cos(theta_sol / 2.0))
    A_real = (D_m**2 / 8.0) * (theta_sol - math.sin(theta_sol))
    vel_m_s = Q_m3_s / A_real if A_real > 0 else 0.0
    return theta_sol, tirante_m, vel_m_s


def calcular_smacna_canalon_bajante(
    W_m,
    I_mm_h,
    ancho_canalon_in,
    peralte_canalon_in,
    apoyo_polin_in,
    apoyo_muro_in,
    geometria_bajada,
    seccion_bajada_info,
    dist_ejes_m,
    largo_nave_m,
    coeficiente_C=0.90,
):
    """Método acoplado SMACNA / Ternium para separación máxima entre bajantes y modulación."""
    I_in_hr = I_mm_h / 25.4
    W_ft = W_m * 3.28084

    mu_ft = ancho_canalon_in / 12.0
    peralte_ft = peralte_canalon_in / 12.0
    m_ratio = (
        peralte_canalon_in / ancho_canalon_in if ancho_canalon_in > 0 else 1.0
    )

    if "Cuadrada" in geometria_bajada or "Rectangular" in geometria_bajada:
        a_in = seccion_bajada_info["ancho_in"]
        b_in = seccion_bajada_info["largo_in"]
        area_bajada_in2 = a_in * b_in
        es_rectangular = True
    else:
        d_in_real = seccion_bajada_info["d_int_mm"] / 25.4
        area_bajada_in2 = math.pi * (d_in_real**2) / 4.0
        a_in = d_in_real
        b_in = d_in_real
        es_rectangular = False

    # 1. Distancia máxima por capacidad del canalón (Fórmula de Beij / SMACNA)
    if W_ft > 0 and I_in_hr > 0:
        num_can = (mu_ft ** (28.0 / 13.0)) * (m_ratio ** (16.0 / 13.0))
        den_can = (0.481 ** (28.0 / 13.0)) * (
            ((W_ft * I_in_hr) / 43200.0) ** (10.0 / 13.0)
        )
        DB_canalon_ft = num_can / den_can if den_can > 0 else 0.0
    else:
        DB_canalon_ft = 0.0
    DB_canalon_m = DB_canalon_ft / 3.28084

    # 2. Distancia máxima por capacidad de la embocadura de la bajante (SMACNA)
    if W_ft > 0 and I_in_hr > 0:
        DB_bajante_ft = (1200.0 * area_bajada_in2) / (W_ft * I_in_hr)
    else:
        DB_bajante_ft = 0.0
    DB_bajante_m = DB_bajante_ft / 3.28084

    # 3. Elemento gobernante
    if DB_bajante_ft <= DB_canalon_ft:
        DB_max_ft = DB_bajante_ft
        DB_max_m = DB_bajante_m
        gobierna = "Gobierna la capacidad de la bajante"
    else:
        DB_max_ft = DB_canalon_ft
        DB_max_m = DB_canalon_m
        gobierna = "Gobierna la capacidad del canalón"

    # 4. Desarrollo de lámina del canalón
    desarrollo_in = (
        ancho_canalon_in
        + (2.0 * peralte_canalon_in)
        + apoyo_polin_in
        + apoyo_muro_in
    )
    desarrollo_cm = desarrollo_in * 2.54

    # 5. Modulación con entre-ejes (columnas) y total de bajantes en la vertiente
    alertas = []
    if dist_ejes_m > 0:
        if dist_ejes_m > DB_max_m:
            ejes_por_bajante = 0
            sep_real_m = dist_ejes_m
            n_bajantes_linea = (
                max(1, math.ceil(largo_nave_m / DB_max_m))
                if DB_max_m > 0
                else 1
            )
            alertas.append(
                f"⚠️ **SEPARACIÓN ENTRE COLUMNAS EXCEDE EL MÁXIMO ADMISIBLE:** La distancia entre ejes (**{dist_ejes_m:.2f} m**) es mayor que la distancia máxima permitida (**{DB_max_m:.2f} m**). **{gobierna}**. Se recomienda aumentar la sección de la bajante/canalón o colocar bajante intermedia."
            )
        else:
            ejes_por_bajante = max(1, int(math.floor(DB_max_m / dist_ejes_m)))
            sep_real_m = ejes_por_bajante * dist_ejes_m
            n_bajantes_linea = (
                max(1, int(math.ceil(largo_nave_m / sep_real_m)))
                if sep_real_m > 0
                else 1
            )
    else:
        ejes_por_bajante = 1
        sep_real_m = DB_max_m
        n_bajantes_linea = (
            max(1, int(math.ceil(largo_nave_m / DB_max_m)))
            if DB_max_m > 0
            else 1
        )

    area_total_vertiente_m2 = W_m * largo_nave_m
    Q_total_lps = (
        coeficiente_C * (I_mm_h / 3600.0) * area_total_vertiente_m2
    )
    q_por_bajante_lps = (
        Q_total_lps / n_bajantes_linea if n_bajantes_linea > 0 else Q_total_lps
    )

    if not alertas:
        mensaje_estado = (
            f"✅ **DISEÑO SMACNA ÓPTIMO ({gobierna}):** Distancia máx. admisible $DB_{{\\text{{máx}}}} = {DB_max_m:.2f}\\text{{ m}}$ ({DB_max_ft:.2f} ft). "
            f"Con columnas a cada **{dist_ejes_m:.2f} m**, se recomienda bajar cada **{ejes_por_bajante} eje(s)** (cada **{sep_real_m:.2f} m**), "
            f"requiriendo **{n_bajantes_linea} bajantes** para los {largo_nave_m:.1f} m de vertiente."
        )
    else:
        mensaje_estado = " \n\n".join(alertas)

    return {
        "es_smacna": True,
        "W_m": round(W_m, 2),
        "W_ft": round(W_ft, 2),
        "I_in_hr": round(I_in_hr, 2),
        "ancho_canalon_in": ancho_canalon_in,
        "peralte_canalon_in": peralte_canalon_in,
        "mu_ft": round(mu_ft, 2),
        "peralte_ft": round(peralte_ft, 2),
        "m_ratio": round(m_ratio, 2),
        "apoyo_polin_in": apoyo_polin_in,
        "apoyo_muro_in": apoyo_muro_in,
        "desarrollo_in": round(desarrollo_in, 1),
        "desarrollo_cm": round(desarrollo_cm, 1),
        "area_bajada_in2": round(area_bajada_in2, 2),
        "DB_canalon_ft": round(DB_canalon_ft, 2),
        "DB_canalon_m": round(DB_canalon_m, 2),
        "DB_bajante_ft": round(DB_bajante_ft, 2),
        "DB_bajante_m": round(DB_bajante_m, 2),
        "DB_max_ft": round(DB_max_ft, 2),
        "DB_max_m": round(DB_max_m, 2),
        "gobierna": gobierna,
        "dist_ejes_m": round(dist_ejes_m, 2),
        "largo_nave_m": round(largo_nave_m, 2),
        "ejes_por_bajante": ejes_por_bajante,
        "sep_real_m": round(sep_real_m, 2),
        "n_bajantes": n_bajantes_linea,
        "n_bajantes_recomendadas": n_bajantes_linea,
        "Q_lps": round(Q_total_lps, 2),
        "q_unit_lps": round(q_por_bajante_lps, 2),
        "nom_in": seccion_bajada_info["nom_in"],
        "d_int_mm": seccion_bajada_info["d_int_mm"],
        "es_bajada_rect": es_rectangular,
        "ancho_bajada_cm": round(a_in * 2.54, 1),
        "largo_bajada_cm": round(b_in * 2.54, 1),
        "b_canal_cm": round(ancho_canalon_in * 2.54, 1),
        "h_canal_cm": round(peralte_canalon_in * 2.54, 1),
        "borde_libre_cm": round(peralte_canalon_in * 2.54 * 0.20, 1),
        "y_cm": round(peralte_canalon_in * 2.54 * 0.75, 1),
        "espesor_mm": 8.0,
        "porcentaje_llenado": round(
            min(100.0, (sep_real_m / DB_max_m) * 100.0) if DB_max_m > 0 else 100.0,
            1,
        ),
        "cap_lps": round(
            Q_total_lps * (DB_max_m / sep_real_m) if sep_real_m > 0 else Q_total_lps,
            2,
        ),
        "vel_m_s": 1.50,
        "v_max_normativa": 4.5,
        "cumple_velocidad": True,
        "estado_ocupacion": (
            f"CUMPLE (c/{sep_real_m:.1f}m ≤ {DB_max_m:.1f}m)"
            if dist_ejes_m <= DB_max_m
            else f"EXCEDIDO ({dist_ejes_m:.1f}m > {DB_max_m:.1f}m)"
        ),
        "n_manning": 0.012,
        "mensaje_estado": mensaje_estado,
        "tiene_alertas": len(alertas) > 0,
        "es_bajante": True,
        "es_canaleta": True,
        "es_canal_abierto": False,
    }


# ==============================================================================
# CÁLCULO HIDRÁULICO GENERAL (CON BAJANTES CIRCULARES Y CUADRADAS)
# ==============================================================================
def calcular_hidraulica(
    area_m2,
    I_mm_h,
    pendiente_pct,
    material_nombre,
    coeficiente_C=0.90,
    diametro_manual_idx=None,
    n_bajantes_manual=None,
    tipo_elemento="Colector Horizontal (Gravedad / Manning)",
    b_canal_cm=30.0,
    h_canal_cm=30.0,
    borde_libre_cm=4.0,
    talud_z=0.0,
):
    material_info = CATALOGO_MATERIALES[material_nombre]
    n_manning = material_info["n"]
    v_max_normativa = material_info.get("v_max", 5.0)
    v_min_normativa = material_info.get("v_min", 0.60)

    I_m_s = (I_mm_h / 1000.0) / 3600.0
    Q_m3_s = coeficiente_C * I_m_s * area_m2
    Q_lps = Q_m3_s * 1000.0

    es_bajante = "Bajante" in tipo_elemento
    es_canaleta = "Canaleta" in tipo_elemento
    es_canal_abierto = "Canal Abierto" in tipo_elemento

    if es_canal_abierto:
        b_m = b_canal_cm / 100.0
        h_total_m = h_canal_cm / 100.0
        bl_m = borde_libre_cm / 100.0
        y_util_m = max(0.001, h_total_m - bl_m)
        S = pendiente_pct / 100.0
        z_eff = talud_z if "Trapezoidal" in tipo_elemento else 0.0

        if "Trapezoidal" in tipo_elemento:
            area_util = (b_m + (z_eff * y_util_m)) * y_util_m
            perimetro_util = b_m + (
                2.0 * y_util_m * math.sqrt(1.0 + (z_eff**2))
            )
            nom_geometria = (
                f"Trap. b={b_canal_cm:.0f}, H={h_canal_cm:.0f} cm (z={z_eff:.1f})"
            )
        else:
            area_util = b_m * y_util_m
            perimetro_util = b_m + (2.0 * y_util_m)
            nom_geometria = f"Rect. {b_canal_cm:.0f}x{h_canal_cm:.0f} cm"

        Rh_util = area_util / perimetro_util if perimetro_util > 0 else 0.0
        cap_m3s = (
            (1.0 / n_manning) * area_util * (Rh_util ** (2.0 / 3.0)) * (S**0.5)
        )
        cap_lps = cap_m3s * 1000.0

        y_real_m, vel_m_s = resolver_tirante_canal(
            Q_m3_s, b_m, z_eff, n_manning, S, y_util_m
        )
        tirante_y_cm = y_real_m * 100.0

        if Q_lps <= cap_lps and y_util_m > 0:
            porcentaje_llenado = (y_real_m / y_util_m) * 100.0
        else:
            porcentaje_llenado = (
                (Q_lps / cap_lps) * 100.0 if cap_lps > 0 else 100.0
            )

        alertas = []
        if Q_lps > cap_lps:
            alertas.append(
                f"⚠️ **CAPACIDAD DEL CANAL EXCEDIDA:** El caudal generado ({Q_lps:.2f} L/s) supera la capacidad útil ({cap_lps:.2f} L/s)."
            )
        elif porcentaje_llenado < 70.0:
            alertas.append(
                f"💡 **OPCIÓN DE OPTIMIZACIÓN ({porcentaje_llenado:.1f}% del tirante útil):** El nivel de agua ocupa menos del 70% de la altura útil de la sección. Se sugiere reducir la geometría (ancho o alto) para optimizar espacio y costos."
            )

        cumple_velocidad = vel_m_s <= v_max_normativa
        if not cumple_velocidad:
            alertas.append(
                f"🚨 **VELOCIDAD EXCEDIDA ({vel_m_s:.2f} m/s > {v_max_normativa:.1f} m/s):** Riesgo de erosión en el canal."
            )
        elif vel_m_s < v_min_normativa and Q_lps > 0:
            alertas.append(
                f"⚠️ **VELOCIDAD BAJA ({vel_m_s:.2f} m/s < {v_min_normativa:.2f} m/s):** Riesgo de sedimentación y azolve."
            )

        estado_ocupacion = (
            "CUMPLE (≤ 100% Útil)"
            if Q_lps <= cap_lps
            else "EXCEDIDO (> 100% Útil)"
        )

        mensaje_estado = (
            " \n\n".join(alertas)
            if alertas
            else "✅ **CANAL ÓPTIMO:** Sección y pendiente adecuadas para el caudal de diseño."
        )

        return {
            "es_smacna": False,
            "Q_lps": round(Q_lps, 2),
            "nom_in": nom_geometria,
            "d_int_mm": b_canal_cm * 10.0,
            "cap_lps": round(cap_lps, 2),
            "vel_m_s": round(vel_m_s, 2),
            "y_cm": round(tirante_y_cm, 2),
            "porcentaje_llenado": round(porcentaje_llenado, 1),
            "estado_ocupacion": estado_ocupacion,
            "n_manning": n_manning,
            "v_max_normativa": v_max_normativa,
            "cumple_velocidad": cumple_velocidad,
            "idx_sugerido": 0,
            "sugerido_nom_in": nom_geometria,
            "es_manual": False,
            "mensaje_estado": mensaje_estado,
            "tiene_alertas": len(alertas) > 0,
            "es_bajante": False,
            "es_canaleta": False,
            "es_canal_abierto": True,
            "b_canal_cm": b_canal_cm,
            "h_canal_cm": h_canal_cm,
            "borde_libre_cm": borde_libre_cm,
            "talud_z": z_eff,
        }

    elif es_canaleta:
        b_m = b_canal_cm / 100.0
        h_total_m = h_canal_cm / 100.0
        bl_m = borde_libre_cm / 100.0
        h_util_m = max(0.001, h_total_m - bl_m)

        area_util = b_m * h_util_m
        perimetro_util = b_m + (2.0 * h_util_m)
        Rh_util = area_util / perimetro_util if perimetro_util > 0 else 0.0
        S = pendiente_pct / 100.0

        cap_m3s = (
            (1.0 / n_manning) * area_util * (Rh_util ** (2.0 / 3.0)) * (S**0.5)
        )
        cap_lps = cap_m3s * 1000.0

        y_real_m, vel_m_s = resolver_tirante_canal(
            Q_m3_s, b_m, 0.0, n_manning, S, h_util_m
        )
        tirante_y_cm = y_real_m * 100.0

        if Q_lps <= cap_lps and h_util_m > 0:
            porcentaje_llenado = (y_real_m / h_util_m) * 100.0
        else:
            porcentaje_llenado = (
                (Q_lps / cap_lps) * 100.0 if cap_lps > 0 else 100.0
            )

        alertas = []
        if Q_lps > cap_lps:
            alertas.append(
                f"⚠️ **CAPACIDAD DE CANALETA EXCEDIDA:** El caudal generado ({Q_lps:.2f} L/s) supera la capacidad máxima útil ({cap_lps:.2f} L/s)."
            )
        elif porcentaje_llenado < 50.0:
            alertas.append(
                f"💡 **OPCIÓN DE OPTIMIZACIÓN ({porcentaje_llenado:.1f}% del tirante útil):** La canaleta opera por debajo del 50% de su altura útil. Se recomienda ajustar la sección para reducir costos de fabricación e instalación."
            )

        cumple_velocidad = vel_m_s <= v_max_normativa
        if not cumple_velocidad:
            alertas.append(
                f"🚨 **VELOCIDAD EXCEDIDA ({vel_m_s:.2f} m/s > {v_max_normativa:.1f} m/s):** Riesgo de desborde por turbulencia y oleaje."
            )

        estado_ocupacion = (
            "CUMPLE (≤ 100% Útil)"
            if Q_lps <= cap_lps
            else "EXCEDIDO (> 100% Útil)"
        )

        mensaje_estado = (
            " \n\n".join(alertas)
            if alertas
            else "✅ **CANALETA ÓPTIMA:** Dimensiones y pendiente adecuadas para el caudal de diseño."
        )

        return {
            "es_smacna": False,
            "Q_lps": round(Q_lps, 2),
            "nom_in": f"Rect. {b_canal_cm:.0f}x{h_canal_cm:.0f} cm",
            "d_int_mm": b_canal_cm * 10.0,
            "cap_lps": round(cap_lps, 2),
            "vel_m_s": round(vel_m_s, 2),
            "y_cm": round(tirante_y_cm, 2),
            "porcentaje_llenado": round(porcentaje_llenado, 1),
            "estado_ocupacion": estado_ocupacion,
            "n_manning": n_manning,
            "v_max_normativa": v_max_normativa,
            "cumple_velocidad": cumple_velocidad,
            "idx_sugerido": 0,
            "sugerido_nom_in": f"Rect. {b_canal_cm:.0f}x{h_canal_cm:.0f} cm",
            "es_manual": False,
            "mensaje_estado": mensaje_estado,
            "tiene_alertas": len(alertas) > 0,
            "es_bajante": False,
            "es_canaleta": True,
            "es_canal_abierto": False,
            "b_canal_cm": b_canal_cm,
            "h_canal_cm": h_canal_cm,
            "borde_libre_cm": borde_libre_cm,
        }

    elif es_bajante:
        lista_tuberias = material_info["tuberias"]
        K_BAJANTE = 0.000052
        es_diam_manual = diametro_manual_idx is not None
        es_cant_manual = n_bajantes_manual is not None

        if not es_diam_manual and not es_cant_manual:
            idx_sugerido = None
            for idx, tub in enumerate(lista_tuberias):
                cap_u = K_BAJANTE * (tub["d_int_mm"] ** 2.6667)
                if cap_u >= Q_lps:
                    idx_sugerido = idx
                    break
            if idx_sugerido is not None:
                idx_evaluar = idx_sugerido
                n_bajantes = 1
            else:
                idx_6in = next(
                    (
                        i
                        for i, t in enumerate(lista_tuberias)
                        if "6" in t["nom_in"]
                    ),
                    len(lista_tuberias) - 1,
                )
                idx_sugerido = idx_6in
                idx_evaluar = idx_sugerido
                cap_u = K_BAJANTE * (
                    lista_tuberias[idx_evaluar]["d_int_mm"] ** 2.6667
                )
                n_bajantes = (
                    max(1, math.ceil(Q_lps / cap_u)) if cap_u > 0 else 1
                )

        elif es_diam_manual and not es_cant_manual:
            idx_evaluar = diametro_manual_idx
            idx_sugerido = diametro_manual_idx
            cap_u = K_BAJANTE * (
                lista_tuberias[idx_evaluar]["d_int_mm"] ** 2.6667
            )
            n_bajantes = max(1, math.ceil(Q_lps / cap_u)) if cap_u > 0 else 1

        elif not es_diam_manual and es_cant_manual:
            n_bajantes = max(1, int(n_bajantes_manual))
            q_por_baj = Q_lps / n_bajantes
            idx_sugerido = len(lista_tuberias) - 1
            for idx, tub in enumerate(lista_tuberias):
                cap_u = K_BAJANTE * (tub["d_int_mm"] ** 2.6667)
                if cap_u >= q_por_baj:
                    idx_sugerido = idx
                    break
            idx_evaluar = idx_sugerido

        else:
            n_bajantes = max(1, int(n_bajantes_manual))
            q_por_baj = Q_lps / n_bajantes
            idx_sugerido = len(lista_tuberias) - 1
            for idx, tub in enumerate(lista_tuberias):
                cap_u = K_BAJANTE * (tub["d_int_mm"] ** 2.6667)
                if cap_u >= q_por_baj:
                    idx_sugerido = idx
                    break
            idx_evaluar = diametro_manual_idx

        tuberia_elegida = lista_tuberias[idx_evaluar]
        D_mm = tuberia_elegida["d_int_mm"]
        es_bajada_rect = "ancho_in" in tuberia_elegida
        ancho_bajada_cm = (
            tuberia_elegida["ancho_in"] * 2.54 if es_bajada_rect else D_mm / 10.0
        )
        largo_bajada_cm = (
            tuberia_elegida["largo_in"] * 2.54 if es_bajada_rect else D_mm / 10.0
        )

        cap_unit_lps = K_BAJANTE * (D_mm**2.6667)
        cap_total_lps = cap_unit_lps * n_bajantes
        n_bajantes_recomendadas = (
            max(1, math.ceil(Q_lps / cap_unit_lps)) if cap_unit_lps > 0 else 1
        )
        q_unit_lps = Q_lps / n_bajantes

        if cap_unit_lps > 0 and q_unit_lps > 0:
            r_s = min(1.0, (1.0 / 3.0) * ((q_unit_lps / cap_unit_lps) ** 0.6))
        else:
            r_s = 0.0

        porcentaje_llenado = r_s * 100.0
        R_mm = D_mm / 2.0
        espesor_anillo_mm = R_mm * (1.0 - math.sqrt(max(0.0, 1.0 - r_s)))
        tirante_y_cm = espesor_anillo_mm / 10.0

        vel_m_s = (
            2.85 * ((q_unit_lps / D_mm) ** (1.0 / 3.0))
            if (D_mm > 0 and q_unit_lps > 0)
            else 0.0
        )

        tabla_comparativa_bap = []
        for tub in lista_tuberias:
            d_i = tub["d_int_mm"]
            c_u = K_BAJANTE * (d_i**2.6667)
            n_req = max(1, math.ceil(Q_lps / c_u)) if c_u > 0 else 1
            q_u = Q_lps / n_req
            tabla_comparativa_bap.append(
                {
                    "Sección Comercial": tub["nom_in"],
                    "D_int / D_hid (mm)": f"{d_i:.1f}",
                    "Cap. por Bajante (L/s)": f"{c_u:.2f}",
                    "Bajantes Recomendadas (BAP)": n_req,
                    "Gasto por Bajante (L/s)": f"{q_u:.2f}",
                }
            )

        alertas = []
        if q_unit_lps > cap_unit_lps * 1.001 or porcentaje_llenado > 33.4:
            alertas.append(
                f"⚠️ **CAPACIDAD DE BAJANTE EXCEDIDA:** Con **{n_bajantes} bajante(s) de {tuberia_elegida['nom_in']}** la capacidad total es de **{cap_total_lps:.2f} L/s** (Gasto de diseño: **{Q_lps:.2f} L/s**). Se recomiendan al menos **{n_bajantes_recomendadas} bajantes de {tuberia_elegida['nom_in']}**."
            )
        elif es_cant_manual and n_bajantes > n_bajantes_recomendadas:
            alertas.append(
                f"💡 **CANTIDAD DE BAJANTES HOLGADA:** Definiste **{n_bajantes} bajantes de {tuberia_elegida['nom_in']}**, mientras que el mínimo requerido por capacidad hidráulica es de **{n_bajantes_recomendadas} bajante(s)**."
            )
        elif (
            es_diam_manual
            and es_cant_manual
            and idx_evaluar > idx_sugerido
            and n_bajantes == n_bajantes_recomendadas
        ):
            alertas.append(
                f"💡 **SECCIÓN HOLGADA:** Para {n_bajantes} bajante(s), la sección mínima requerida es **{lista_tuberias[idx_sugerido]['nom_in']}**."
            )

        cumple_velocidad = vel_m_s <= v_max_normativa
        if not cumple_velocidad:
            alertas.append(
                f"🚨 **VELOCIDAD TERMINAL EXCEDIDA ({vel_m_s:.2f} m/s > {v_max_normativa:.1f} m/s):** Disipar energía en el codo/registro inferior."
            )

        estado_ocupacion = (
            "CUMPLE (Anular ≤ 33.3%)"
            if (q_unit_lps <= cap_unit_lps * 1.001 and porcentaje_llenado <= 33.4)
            else "EXCEDIDO (> 33.3%)"
        )

        mensaje_estado = (
            " \n\n".join(alertas)
            if alertas
            else f"✅ **DISEÑO DE BAJANTES ÓPTIMO:** {n_bajantes} bajante(s) de {tuberia_elegida['nom_in']} ({q_unit_lps:.2f} L/s por bajada) cumplen con el límite de flujo anular (≤ 33.3%)."
        )

        return {
            "es_smacna": False,
            "Q_lps": round(Q_lps, 2),
            "q_unit_lps": round(q_unit_lps, 2),
            "n_bajantes": n_bajantes,
            "n_bajantes_recomendadas": n_bajantes_recomendadas,
            "cap_unit_lps": round(cap_unit_lps, 2),
            "cap_lps": round(cap_total_lps, 2),
            "nom_in": tuberia_elegida["nom_in"],
            "d_int_mm": tuberia_elegida["d_int_mm"],
            "es_bajada_rect": es_bajada_rect,
            "ancho_bajada_cm": round(ancho_bajada_cm, 1),
            "largo_bajada_cm": round(largo_bajada_cm, 1),
            "vel_m_s": round(vel_m_s, 2),
            "y_cm": round(tirante_y_cm, 2),
            "espesor_mm": round(espesor_anillo_mm, 1),
            "porcentaje_llenado": round(porcentaje_llenado, 1),
            "estado_ocupacion": estado_ocupacion,
            "n_manning": n_manning,
            "v_max_normativa": v_max_normativa,
            "cumple_velocidad": cumple_velocidad,
            "idx_sugerido": idx_sugerido,
            "sugerido_nom_in": lista_tuberias[idx_sugerido]["nom_in"],
            "es_manual": es_diam_manual or es_cant_manual,
            "mensaje_estado": mensaje_estado,
            "tiene_alertas": len(alertas) > 0,
            "es_bajante": True,
            "es_canaleta": False,
            "es_canal_abierto": False,
            "tabla_comparativa_bap": tabla_comparativa_bap,
        }

    else:
        lista_tuberias = material_info["tuberias"]
        S = pendiente_pct / 100.0
        LIMITE_LLENADO_PCT = 80.0

        idx_sugerido = len(lista_tuberias) - 1
        for idx, tub in enumerate(lista_tuberias):
            D = tub["d_int_mm"] / 1000.0
            A = math.pi * (D**2) / 4.0
            Rh = D / 4.0
            Q_cap = (
                (1.0 / n_manning) * A * (Rh ** (2.0 / 3.0)) * (S**0.5) * 1000.0
            )

            if Q_cap >= Q_lps:
                _, tirante_m, _ = resolver_tirante_circular(
                    Q_m3_s, D, n_manning, S
                )
                if (tirante_m / D) * 100.0 <= LIMITE_LLENADO_PCT:
                    idx_sugerido = idx
                    break

        es_manual = diametro_manual_idx is not None
        idx_evaluar = diametro_manual_idx if es_manual else idx_sugerido
        tuberia_elegida = lista_tuberias[idx_evaluar]

        D = tuberia_elegida["d_int_mm"] / 1000.0
        A = math.pi * (D**2) / 4.0
        Rh = D / 4.0
        cap_lps = (
            (1.0 / n_manning) * A * (Rh ** (2.0 / 3.0)) * (S**0.5) * 1000.0
        )

        if Q_lps > cap_lps:
            porcentaje_llenado = 100.0
            tirante_y_cm = D * 100.0
            vel_m_s = Q_m3_s / A
        else:
            _, tirante_m, vel_m_s = resolver_tirante_circular(
                Q_m3_s, D, n_manning, S
            )
            tirante_y_cm = tirante_m * 100.0
            porcentaje_llenado = (tirante_m / D) * 100.0

        alertas = []
        if porcentaje_llenado > LIMITE_LLENADO_PCT or Q_lps > cap_lps:
            alertas.append(
                "⚠️ **CAPACIDAD EXCEDIDA:** Llenado > 80%. Se recomienda aumentar el diámetro."
            )
        elif idx_evaluar > idx_sugerido:
            alertas.append(
                f"💡 **DIÁMETRO SOBRADO:** El diámetro seleccionado es mayor al sugerido ({lista_tuberias[idx_sugerido]['nom_in']})."
            )

        cumple_velocidad = vel_m_s <= v_max_normativa
        if not cumple_velocidad:
            alertas.append(
                f"🚨 **VELOCIDAD EXCEDIDA ({vel_m_s:.2f} m/s > {v_max_normativa:.1f} m/s):** Disminuir la pendiente."
            )
        elif vel_m_s < v_min_normativa and Q_lps > 0:
            alertas.append(
                f"⚠️ **VELOCIDAD BAJA DE AUTOLIMPIEZA ({vel_m_s:.2f} m/s < {v_min_normativa:.2f} m/s):** Riesgo de sedimentación."
            )

        estado_ocupacion = (
            "CUMPLE (≤ 80.0%)"
            if (porcentaje_llenado <= LIMITE_LLENADO_PCT and Q_lps <= cap_lps)
            else "EXCEDIDO (> 80.0%)"
        )

        mensaje_estado = (
            " \n\n".join(alertas)
            if alertas
            else "✅ **DISEÑO ÓPTIMO:** Cumple con llenado (≤ 80%) y velocidad permisible."
        )

        return {
            "es_smacna": False,
            "Q_lps": round(Q_lps, 2),
            "nom_in": tuberia_elegida["nom_in"],
            "d_int_mm": tuberia_elegida["d_int_mm"],
            "cap_lps": round(cap_lps, 2),
            "vel_m_s": round(vel_m_s, 2),
            "y_cm": round(tirante_y_cm, 2),
            "porcentaje_llenado": round(porcentaje_llenado, 1),
            "estado_ocupacion": estado_ocupacion,
            "n_manning": n_manning,
            "v_max_normativa": v_max_normativa,
            "cumple_velocidad": cumple_velocidad,
            "idx_sugerido": idx_sugerido,
            "sugerido_nom_in": lista_tuberias[idx_sugerido]["nom_in"],
            "es_manual": es_manual,
            "mensaje_estado": mensaje_estado,
            "tiene_alertas": len(alertas) > 0,
            "es_bajante": False,
            "es_canaleta": False,
            "es_canal_abierto": False,
        }


# ==============================================================================
# ESQUEMA GRÁFICO CON ACOTACIONES POSITIVAS TIPO CAD (SIN EJES NEGATIVOS)
# ==============================================================================
def dibujar_cota_cad(
    ax,
    p1,
    p2,
    texto,
    offset=0.0,
    orientacion="H",
    color="#263238",
    fontsize=7.5,
):
    """Dibuja una línea de acotación estilo CAD (<--->) con texto paralelo a la cota sin empalmes."""
    x1, y1 = p1
    x2, y2 = p2

    if orientacion == "H":
        y_cota = y1 + offset
        ax.plot(
            [x1, x1], [y1, y_cota], color=color, linestyle=":", lw=0.8, alpha=0.7
        )
        ax.plot(
            [x2, x2], [y2, y_cota], color=color, linestyle=":", lw=0.8, alpha=0.7
        )
        ax.annotate(
            "",
            xy=(x1, y_cota),
            xytext=(x2, y_cota),
            arrowprops=dict(arrowstyle="<->", color=color, lw=1.1),
        )
        ax.text(
            (x1 + x2) / 2.0,
            y_cota,
            texto,
            ha="center",
            va="center",
            fontsize=fontsize,
            fontweight="bold",
            color=color,
            bbox=dict(
                boxstyle="round,pad=0.20",
                facecolor="white",
                edgecolor="none",
                alpha=0.95,
            ),
        )
    else:
        x_cota = x1 + offset
        ax.plot(
            [x1, x_cota], [y1, y1], color=color, linestyle=":", lw=0.8, alpha=0.7
        )
        ax.plot(
            [x2, x_cota], [y2, y2], color=color, linestyle=":", lw=0.8, alpha=0.7
        )
        ax.annotate(
            "",
            xy=(x_cota, y1),
            xytext=(x_cota, y2),
            arrowprops=dict(arrowstyle="<->", color=color, lw=1.1),
        )
        # Texto rotado 90° paralelo a la línea de cota vertical para no invadir el dibujo
        ax.text(
            x_cota,
            (y1 + y2) / 2.0,
            texto,
            ha="center",
            va="center",
            rotation=90,
            fontsize=fontsize,
            fontweight="bold",
            color=color,
            bbox=dict(
                boxstyle="round,pad=0.20",
                facecolor="white",
                edgecolor="none",
                alpha=0.95,
            ),
        )


def generar_grafico_seccion(
    tipo_elemento,
    b,
    H,
    y,
    talud_z=0.0,
    D_cm=None,
    borde_libre_cm=0.0,
    es_bajada_rect=False,
    ancho_bajada_cm=None,
    largo_bajada_cm=None,
):
    fig, ax = plt.subplots(figsize=(4.6, 3.1), dpi=135)

    if "Canal" in tipo_elemento:
        es_canaleta = "Canaleta" in tipo_elemento
        z_eff = talud_z if "Trapezoidal" in tipo_elemento else 0.0
        lbl_estructura = (
            f"Canaleta ({b:.0f}x{H:.0f} cm)"
            if es_canaleta
            else f"Canal (b={b:.0f}, H={H:.0f} cm)"
        )

        x_sup_der = b / 2.0 + (z_eff * H)
        x_canal = [-x_sup_der, -b / 2.0, b / 2.0, x_sup_der]
        y_canal = [H, 0, 0, H]

        y_agua_val = min(y, H)
        x_agua_der = b / 2.0 + (z_eff * y_agua_val)
        x_agua = [-x_agua_der, -b / 2.0, b / 2.0, x_agua_der]
        y_agua = [y_agua_val, 0, 0, y_agua_val]

        ax.fill(
            x_agua,
            y_agua,
            color="#4FC3F7",
            alpha=0.65,
            label=f"Tirante y = {y:.1f} cm",
        )
        ax.plot(
            x_canal, y_canal, color="#37474F", lw=2.8, label=lbl_estructura
        )
        ax.plot(
            [-x_sup_der, x_sup_der],
            [H, H],
            color="#D32F2F",
            linestyle="--",
            lw=1.1,
            label=f"Altura H = {H:.1f} cm",
        )

        y_util_cota = H - borde_libre_cm
        if borde_libre_cm > 0 and y_util_cota > 0:
            x_bl_lim = b / 2.0 + (z_eff * y_util_cota)
            ax.plot(
                [-x_bl_lim, x_bl_lim],
                [y_util_cota, y_util_cota],
                color="#2E7D32",
                linestyle="-.",
                lw=1.5,
                label=f"B. Libre = {borde_libre_cm:.1f} cm",
            )

        # Acotaciones CAD positivas con holgura limpia
        off_x = max(x_sup_der * 0.28, 3.5)
        off_y = max(H * 0.20, 3.0)
        dibujar_cota_cad(
            ax,
            (-b / 2.0, 0),
            (b / 2.0, 0),
            f"b = {b:.1f} cm",
            offset=-off_y,
            orientacion="H",
        )
        if z_eff > 0:
            B_sup = b + 2.0 * z_eff * H
            dibujar_cota_cad(
                ax,
                (-x_sup_der, H),
                (x_sup_der, H),
                f"B = {B_sup:.1f} cm",
                offset=off_y * 0.85,
                orientacion="H",
            )
        dibujar_cota_cad(
            ax,
            (-x_sup_der, 0),
            (-x_sup_der, H),
            f"H = {H:.1f} cm",
            offset=-off_x,
            orientacion="V",
        )
        if y_agua_val > 0.05 * H:
            dibujar_cota_cad(
                ax,
                (x_sup_der, 0),
                (x_sup_der, y_agua_val),
                f"y = {y:.1f} cm",
                offset=off_x,
                orientacion="V",
                color="#0277BD",
            )

        ax.set_xlim(-x_sup_der - off_x * 1.8, x_sup_der + off_x * 1.8)
        ax.set_ylim(-off_y * 1.7, H + off_y * 1.5)

    elif "Bajante" in tipo_elemento:
        if es_bajada_rect and ancho_bajada_cm and largo_bajada_cm:
            W_b = ancho_bajada_cm
            L_b = largo_bajada_cm
            espesor_cm = min(y, min(W_b, L_b) * 0.42)

            rect_ext = Rectangle(
                (-W_b / 2.0, 0),
                W_b,
                L_b,
                facecolor="#4FC3F7",
                edgecolor="#37474F",
                lw=2.8,
                alpha=0.70,
                label=f"Anillo e = {espesor_cm * 10.0:.1f} mm",
            )
            ax.add_patch(rect_ext)

            W_int = max(0.2, W_b - 2.0 * espesor_cm)
            L_int = max(0.2, L_b - 2.0 * espesor_cm)
            rect_int = Rectangle(
                (-W_int / 2.0, espesor_cm),
                W_int,
                L_int,
                facecolor="#FFFFFF",
                edgecolor="#0288D1",
                linestyle="--",
                lw=1.3,
                label="Núcleo de Aire",
            )
            ax.add_patch(rect_int)

            off_x = max(W_b * 0.28, 2.5)
            off_y = max(L_b * 0.22, 2.5)
            dibujar_cota_cad(
                ax,
                (-W_b / 2.0, 0),
                (W_b / 2.0, 0),
                f"a = {W_b:.1f} cm",
                offset=-off_y,
                orientacion="H",
            )
            dibujar_cota_cad(
                ax,
                (W_b / 2.0, 0),
                (W_b / 2.0, L_b),
                f"b = {L_b:.1f} cm",
                offset=off_x,
                orientacion="V",
            )

            ax.set_xlim(-W_b / 2.0 - off_x * 1.5, W_b / 2.0 + off_x * 1.8)
            ax.set_ylim(-off_y * 1.7, L_b + off_y * 1.3)
        else:
            R = (D_cm / 2.0) if D_cm else 5.0
            D_total = 2.0 * R
            espesor_cm = min(y, R * 0.90)
            R_aire = max(0.1, R - espesor_cm)

            theta = np.linspace(0, 2 * np.pi, 200)
            x_ext = R * np.cos(theta)
            y_ext = R * np.sin(theta) + R
            x_int = R_aire * np.cos(theta)
            y_int = R_aire * np.sin(theta) + R

            ax.fill(
                x_ext,
                y_ext,
                color="#4FC3F7",
                alpha=0.65,
                label=f"Anillo e = {espesor_cm * 10.0:.1f} mm",
            )
            ax.fill(
                x_int,
                y_int,
                color="#FFFFFF",
                alpha=1.0,
                label="Núcleo de Aire",
            )
            ax.plot(
                x_ext,
                y_ext,
                color="#37474F",
                lw=2.8,
                label=f"Bajante Ø = {D_total:.1f} cm",
            )
            ax.plot(x_int, y_int, color="#0288D1", linestyle="--", lw=1.3)

            off = max(D_total * 0.22, 2.2)
            dibujar_cota_cad(
                ax,
                (-R, 0),
                (R, 0),
                f"Ø = {D_total:.1f} cm",
                offset=-off,
                orientacion="H",
            )
            dibujar_cota_cad(
                ax,
                (-R_aire, R),
                (R_aire, R),
                f"Aire = {2.0 * R_aire:.1f} cm",
                offset=R + off * 0.75,
                orientacion="H",
                color="#0277BD",
                fontsize=7.2,
            )

            ax.set_xlim(-R - off * 1.5, R + off * 1.5)
            ax.set_ylim(-off * 1.7, D_total + off * 1.4)

    else:  # Colector Horizontal Circular
        R = (D_cm / 2.0) if D_cm else 10.0
        D_total = 2.0 * R
        theta = np.linspace(0, 2 * np.pi, 200)
        x_tubo = R * np.cos(theta)
        y_tubo = R * np.sin(theta) + R

        ax.plot(
            x_tubo,
            y_tubo,
            color="#37474F",
            lw=2.8,
            label=f"Tubería Ø = {D_total:.1f} cm",
        )

        y_agua_val = min(y, D_total)
        y_fill = np.linspace(0, y_agua_val, 120)
        x_lims = np.sqrt(np.maximum(0.0, R**2 - (y_fill - R) ** 2))
        ax.fill_betweenx(
            y_fill,
            -x_lims,
            x_lims,
            color="#4FC3F7",
            alpha=0.65,
            label=f"Tirante y = {y:.1f} cm",
        )

        if y_agua_val > 0:
            x_sup_agua = math.sqrt(max(0.0, R**2 - (y_agua_val - R) ** 2))
            ax.plot(
                [-x_sup_agua, x_sup_agua],
                [y_agua_val, y_agua_val],
                color="#0288D1",
                linestyle="-",
                lw=1.5,
            )

        y_80_cm = 0.80 * D_total
        x_80 = math.sqrt(max(0.0, R**2 - (y_80_cm - R) ** 2))
        ax.plot(
            [-x_80, x_80],
            [y_80_cm, y_80_cm],
            color="#E65100",
            linestyle="--",
            lw=1.6,
            label=f"Límite 80% ({y_80_cm:.1f} cm)",
        )

        # Acotaciones CAD positivas con texto vertical rotado
        off_x = max(D_total * 0.25, 3.0)
        off_y = max(D_total * 0.20, 2.5)
        dibujar_cota_cad(
            ax,
            (-R, 0),
            (R, 0),
            f"Ø = {D_total:.1f} cm",
            offset=-off_y,
            orientacion="H",
        )
        dibujar_cota_cad(
            ax,
            (-R, 0),
            (-R, y_80_cm),
            f"0.80Ø = {y_80_cm:.1f} cm",
            offset=-off_x,
            orientacion="V",
            color="#E65100",
        )
        if y_agua_val > 0.05 * D_total:
            dibujar_cota_cad(
                ax,
                (R, 0),
                (R, y_agua_val),
                f"y = {y:.1f} cm",
                offset=off_x,
                orientacion="V",
                color="#0277BD",
            )

        ax.set_xlim(-R - off_x * 1.8, R + off_x * 1.8)
        ax.set_ylim(-off_y * 1.7, D_total + off_y * 1.3)

    ax.set_aspect("equal")
    ax.axis("off")

    ax.legend(
        loc="upper center",
        bbox_to_anchor=(0.5, -0.04),
        ncol=2,
        fontsize=7.5,
        frameon=True,
        facecolor="#F8F9FA",
        edgecolor="#E0E0E0",
    )
    plt.tight_layout()
    return fig
# ==============================================================================
# GENERADOR DEL DOCUMENTO DOCX
# ==============================================================================
def generar_memoria_docx(
    nombre_p,
    lat,
    lon,
    area_m2,
    pendiente_pct,
    tr_seleccionado,
    tc_adoptado,
    tipo_superficie,
    coeficiente_C,
    material_nombre,
    estacion_nom,
    I_ref,
    tabla_tr,
    hidro,
    tipo_elemento,
    img_seccion_bytes=None,
    img_secundaria_bytes=None,
):
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.98)
        section.bottom_margin = Inches(0.98)
        section.left_margin = Inches(0.98)
        section.right_margin = Inches(0.98)

    p_title = doc.add_paragraph()
    r_title = p_title.add_run(
        "MEMORIA DE CÁLCULO HIDRÁULICO Y DRENAJE PLUVIAL"
    )
    r_title.font.name = "Arial"
    r_title.font.size = Pt(16)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(27, 73, 131)

    p_sub = doc.add_paragraph()
    r_sub = p_sub.add_run(
        f"Proyecto: {nombre_p}  |  Coordenadas: Lat {lat:.6f}, Lon {lon:.6f}"
    )
    r_sub.font.name = "Arial"
    r_sub.font.size = Pt(10)
    r_sub.font.italic = True
    r_sub.font.color.rgb = RGBColor(80, 80, 80)

    doc.add_paragraph()

    h1 = doc.add_paragraph()
    r_h1 = h1.add_run("1. Bases de Diseño Hidrológico y Criterios Normativos")
    r_h1.font.name = "Arial"
    r_h1.font.size = Pt(13)
    r_h1.font.bold = True
    r_h1.font.color.rgb = RGBColor(27, 73, 131)

    p_exp = doc.add_paragraph()
    p_exp.paragraph_format.line_spacing = 1.15
    p_exp.add_run("El diseño pluvial se calcula mediante el Método Racional: ")
    p_exp.add_run("Q = (C · I · A) / 3600\n").bold = True

    p_vars = doc.add_paragraph()
    p_vars.paragraph_format.line_spacing = 1.15
    p_vars.add_run("• Tipo de Elemento Evaluado: ").bold = True
    p_vars.add_run(f"{tipo_elemento}\n")
    p_vars.add_run(
        f"• Periodo de Retorno Adoptado (T = {tr_seleccionado} años): "
    ).bold = True
    p_vars.add_run(
        "Criterio de diseño seleccionado para la infraestructura.\n"
    )
    p_vars.add_run(
        f"• Tiempo de Concentración (tc = {tc_adoptado:.2f} min): "
    ).bold = True
    p_vars.add_run(
        "Calculado por la Fórmula de Kirpich / Valor normativo adoptado.\n"
    )
    p_vars.add_run(
        f"• Coeficiente de Escurrimiento (C = {coeficiente_C:.2f}): "
    ).bold = True
    p_vars.add_run(
        f"Tipo de superficie: {tipo_superficie} (Tr={tr_seleccionado} años).\n"
    )
    p_vars.add_run("• Material Propuesto y Rugosidad: ").bold = True
    p_vars.add_run(f"{material_nombre} (n = {hidro['n_manning']:.3f}).\n")
    p_vars.add_run("• Criterio Hidráulico: ").bold = True
    if hidro.get("es_smacna", False):
        p_vars.add_run(
            "Método SMACNA / Beij para sistema acoplado Canalón-Bajante y "
            "modulación estructural por entre-ejes de columnas.\n"
        )
    elif hidro.get("es_canal_abierto", False):
        p_vars.add_run(
            f"Canal Abierto de Concreto (Manning n = {hidro['n_manning']:.3f}).\n"
        )
    elif hidro.get("es_canaleta", False):
        p_vars.add_run(
            "Canal Abierto Rectangular en Lámina Galvanizada (Manning n ="
            f" {hidro['n_manning']:.3f}).\n"
        )
    elif hidro["es_bajante"]:
        p_vars.add_run(
            "Flujo Anular Vertical (Núcleo libre de aire <= 33.3%"
            " ocupación).\n"
        )
    else:
        p_vars.add_run(
            f"Fórmula de Manning (n = {hidro['n_manning']:.3f}) en flujo por"
            " gravedad.\n"
        )
    p_vars.add_run("• Fuente Meteorológica: ").bold = True
    p_vars.add_run(f"{estacion_nom}.")

    area_reporte = (
        hidro["W_m"] * hidro["largo_nave_m"]
        if hidro.get("es_smacna", False)
        else area_m2
    )

    t1_data = [
        ["Parámetro", "Símbolo", "Valor", "Unidad / Criterio"],
        [
            "Elemento de Diseño",
            "Elemento",
            tipo_elemento.split("(")[0].strip(),
            "Estructura",
        ],
        ["Área de Aportación", "A", f"{area_reporte:.2f}", "m²"],
        ["Periodo de Retorno", "T", f"{tr_seleccionado}", "Años"],
        ["Tiempo de Concentración", "tc", f"{tc_adoptado:.2f}", "Minutos"],
        ["Tipo de Superficie", "Superficie", tipo_superficie, "Superficie"],
        [
            "Coeficiente de Escurrimiento",
            "C",
            f"{coeficiente_C:.2f}",
            "Adimensional",
        ],
        [
            "Intensidad Adoptada",
            "I",
            f"{I_ref:.2f} ({I_ref / 25.4:.2f})",
            "mm/h (in/hr)",
        ],
        ["Material de Conducto", "Mat", material_nombre, "Tubería / Sección"],
        [
            "Coeficiente de Rugosidad",
            "n",
            f"{hidro['n_manning']:.3f}",
            "Manning",
        ],
    ]
    t1 = doc.add_table(rows=len(t1_data), cols=4)
    t1.alignment = WD_TABLE_ALIGNMENT.CENTER
    for ir, row in enumerate(t1.rows):
        for ic, cell in enumerate(row.cells):
            cell.text = t1_data[ir][ic]
            fijar_margenes_celda(cell)
            p = cell.paragraphs[0]
            r = p.runs[0]
            r.font.name = "Arial"
            r.font.size = Pt(9.5)
            if ir == 0:
                aplicar_sombreado_celda(cell, "1B4983")
                r.font.bold = True
                r.font.color.rgb = RGBColor(255, 255, 255)
            elif ir % 2 == 1:
                aplicar_sombreado_celda(cell, "F2F4F8")

    doc.add_paragraph()

    if tabla_tr:
        h_tr = doc.add_paragraph()
        r_htr = h_tr.add_run(
            "Tabla de Consulta: Matriz de Intensidades (mm/h) por Duración y Tr"
        )
        r_htr.bold = True
        r_htr.font.color.rgb = RGBColor(27, 73, 131)

        headers_tr = list(tabla_tr[0].keys())
        t_tr = doc.add_table(rows=len(tabla_tr) + 1, cols=len(headers_tr))
        t_tr.alignment = WD_TABLE_ALIGNMENT.CENTER

        for ic, head in enumerate(headers_tr):
            cell = t_tr.cell(0, ic)
            cell.text = head
            fijar_margenes_celda(cell)
            p = cell.paragraphs[0]
            r = p.runs[0]
            r.font.name = "Arial"
            r.font.size = Pt(8.5)
            r.font.bold = True
            r.font.color.rgb = RGBColor(255, 255, 255)
            aplicar_sombreado_celda(cell, "1B4983")

        texto_tr_resaltar = f"{tr_seleccionado} Años"

        for ir, fila in enumerate(tabla_tr):
            for ic, head in enumerate(headers_tr):
                cell = t_tr.cell(ir + 1, ic)
                cell.text = str(fila[head])
                fijar_margenes_celda(cell)
                p = cell.paragraphs[0]
                r = p.runs[0]
                r.font.name = "Arial"
                r.font.size = Pt(8.5)

                if texto_tr_resaltar == str(fila.get("Tr", "")):
                    aplicar_sombreado_celda(cell, "D9E1F2")
                    r.font.bold = True
                elif ir % 2 == 1:
                    aplicar_sombreado_celda(cell, "F2F4F8")

    doc.add_paragraph()

    h2 = doc.add_paragraph()
    r_h2 = h2.add_run(
        "2. Dimensionamiento Hidráulico y Revisión de Capacidad"
    )
    r_h2.font.name = "Arial"
    r_h2.font.size = Pt(13)
    r_h2.font.bold = True
    r_h2.font.color.rgb = RGBColor(27, 73, 131)

    if hidro.get("es_smacna", False):
        t2_data = [
            ["Parámetro Calculado (SMACNA)", "Símbolo", "Resultado", "Unidad / Criterio"],
            [
                "Longitud de Vertiente por Drenar",
                "W",
                f"{hidro['W_m']:.2f} ({hidro['W_ft']:.2f})",
                "m (ft)",
            ],
            [
                "Ancho Promedio del Canalón",
                "μ",
                f"{hidro['ancho_canalon_in']:.1f} ({hidro['mu_ft']:.2f})",
                "in (ft)",
            ],
            [
                "Peralte del Canalón y Relación",
                "P / m",
                f"{hidro['peralte_canalon_in']:.1f} in (m = {hidro['m_ratio']:.2f})",
                "Peralte / Ancho",
            ],
            [
                "Desarrollo Total de Lámina",
                "Des.",
                f"{hidro['desarrollo_in']:.1f} ({hidro['desarrollo_cm']:.1f})",
                "in (cm) inc. apoyos",
            ],
            [
                "Sección y Área de Bajante",
                "A",
                f"{hidro['nom_in']} ({hidro['area_bajada_in2']:.2f})",
                "in²",
            ],
            [
                "Dist. entre Bajantes (Cap. Canalón)",
                "DB_can",
                f"{hidro['DB_canalon_m']:.2f} ({hidro['DB_canalon_ft']:.2f})",
                "m (ft)",
            ],
            [
                "Dist. entre Bajantes (Cap. Bajante)",
                "DB_baj",
                f"{hidro['DB_bajante_m']:.2f} ({hidro['DB_bajante_ft']:.2f})",
                "m (ft)",
            ],
            [
                "Distancia Máxima Admisible",
                "DB_máx",
                f"{hidro['DB_max_m']:.2f} ({hidro['DB_max_ft']:.2f})",
                hidro["gobierna"],
            ],
            [
                "Modulación por Ejes de Columnas",
                "Ejes",
                f"Cada {hidro['ejes_por_bajante']} eje(s) = {hidro['sep_real_m']:.2f} m",
                f"Entre-eje: {hidro['dist_ejes_m']:.2f} m",
            ],
            [
                "Cantidad de Bajantes en la Línea",
                "N_BAP",
                f"{hidro['n_bajantes']} bajantes",
                f"Largo nave: {hidro['largo_nave_m']:.1f} m",
            ],
            [
                "Gasto Pluvial Total / Por Bajada",
                "Q",
                f"{hidro['Q_lps']:.2f} ({hidro['q_unit_lps']:.2f} c/u)",
                "L/s",
            ],
        ]
    else:
        v_max = hidro.get("v_max_normativa", 5.0)
        estado_v = (
            f"CUMPLE (≤ {v_max:.1f} m/s)"
            if hidro["cumple_velocidad"]
            else f"EXCEDIDO (> {v_max:.1f} m/s)"
        )

        if hidro.get("es_canal_abierto", False):
            lbl_vel = "Velocidad Real de Flujo"
            lbl_cap = "Capacidad Máxima Útil"
            lbl_tirante = f"{hidro['y_cm']:.2f} cm"
            dim_label = f"{hidro['nom_in']} (B.L. {hidro['borde_libre_cm']:.1f} cm)"
        elif hidro.get("es_canaleta", False):
            lbl_vel = "Velocidad Real de Flujo"
            lbl_cap = "Capacidad Máxima Útil"
            lbl_tirante = f"{hidro['y_cm']:.2f} cm"
            dim_label = (
                f"{hidro['b_canal_cm']:.0f}x{hidro['h_canal_cm']:.0f} cm (B.L."
                f" {hidro['borde_libre_cm']:.1f} cm)"
            )
        elif hidro["es_bajante"]:
            lbl_vel = "Velocidad Terminal"
            lbl_cap = "Capacidad Anular Total"
            lbl_tirante = (
                f"{hidro.get('espesor_mm', hidro['y_cm'] * 10.0):.1f} mm (Anillo)"
            )
            dim_label = (
                f"{hidro.get('n_bajantes', 1)} BAP de {hidro['nom_in']} "
                f"(D_hid: {hidro['d_int_mm']} mm)"
            )
        else:
            lbl_vel = "Velocidad Real de Flujo"
            lbl_cap = "Capacidad a Tubo Lleno"
            lbl_tirante = f"{hidro['y_cm']:.2f} cm"
            dim_label = f"{hidro['nom_in']}"

        t2_data = [
            ["Parámetro Calculado", "Símbolo", "Resultado", "Unidad / Estado"],
            ["Gasto Pluvial Total", "Q_diseño", f"{hidro['Q_lps']:.2f}", "L/s"],
        ]

        if hidro["es_bajante"]:
            t2_data.append(
                [
                    "Cantidad de Bajantes Pluviales",
                    "N_BAP",
                    f"{hidro.get('n_bajantes', 1)}",
                    f"Recomendadas: {hidro.get('n_bajantes_recomendadas', 1)}",
                ]
            )
            t2_data.append(
                [
                    "Gasto por Bajante Individual",
                    "Q_unit",
                    f"{hidro.get('q_unit_lps', hidro['Q_lps']):.2f}",
                    f"Cap. Unit.: {hidro.get('cap_unit_lps', hidro['cap_lps']):.2f} L/s",
                ]
            )

        t2_data.extend(
            [
                [
                    "Pendiente / Orientación",
                    "S",
                    (
                        "90° (Vertical)"
                        if hidro["es_bajante"]
                        else f"{pendiente_pct:.2f}%"
                    ),
                    "% / Grados",
                ],
                ["Geometría / Sección", "Sección", dim_label, "Dimensiones"],
                [lbl_cap, "Q_máx", f"{hidro['cap_lps']:.2f}", "L/s"],
                ["Tirante / Espesor Hidráulico", "y", lbl_tirante, "Tirante Normal"],
                [
                    "Llenado / Ocupación",
                    "y / D",
                    f"{hidro['porcentaje_llenado']:.1f}%",
                    hidro.get("estado_ocupacion", "OK"),
                ],
                [lbl_vel, "V", f"{hidro['vel_m_s']:.2f}", "m/s"],
                [
                    "Revisión de Velocidad",
                    "V_eval",
                    f"{hidro['vel_m_s']:.2f} m/s",
                    estado_v,
                ],
            ]
        )

    t2 = doc.add_table(rows=len(t2_data), cols=4)
    t2.alignment = WD_TABLE_ALIGNMENT.CENTER
    for ir, row in enumerate(t2.rows):
        for ic, cell in enumerate(row.cells):
            cell.text = t2_data[ir][ic]
            fijar_margenes_celda(cell)
            p = cell.paragraphs[0]
            r = p.runs[0]
            r.font.name = "Arial"
            r.font.size = Pt(9.5)
            if ir == 0:
                aplicar_sombreado_celda(cell, "1B4983")
                r.font.bold = True
                r.font.color.rgb = RGBColor(255, 255, 255)
            elif ir % 2 == 1:
                aplicar_sombreado_celda(cell, "F2F4F8")

    # Dictamen técnico en el documento
    doc.add_paragraph()
    p_dictamen = doc.add_paragraph()
    r_dic_lbl = p_dictamen.add_run("Dictamen Técnico: ")
    r_dic_lbl.font.name = "Arial"
    r_dic_lbl.font.size = Pt(10)
    r_dic_lbl.font.bold = True
    texto_limpio = (
        hidro["mensaje_estado"]
        .replace("**", "")
        .replace("$DB_{\\text{máx}} =", "DB_máx =")
        .replace("\\text{ m}$", " m")
    )
    r_dic_txt = p_dictamen.add_run(texto_limpio)
    r_dic_txt.font.name = "Arial"
    r_dic_txt.font.size = Pt(10)

    # Inserción del esquema geométrico acotado en el Word
    if img_seccion_bytes is not None and img_secundaria_bytes is not None:
        doc.add_paragraph()
        t_img = doc.add_table(rows=1, cols=2)
        t_img.alignment = WD_TABLE_ALIGNMENT.CENTER
        p_left = t_img.cell(0, 0).paragraphs[0]
        p_left.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_left.add_run().add_picture(
            io.BytesIO(img_seccion_bytes), width=Inches(3.0)
        )
        p_right = t_img.cell(0, 1).paragraphs[0]
        p_right.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_right.add_run().add_picture(
            io.BytesIO(img_secundaria_bytes), width=Inches(3.0)
        )
    elif img_seccion_bytes is not None:
        doc.add_paragraph()
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_img = p_img.add_run()
        run_img.add_picture(io.BytesIO(img_seccion_bytes), width=Inches(3.6))

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


# ==============================================================================
# INTERFAZ GRÁFICA (STREAMLIT)
# ==============================================================================

st.title("🌧️ Generador de Memorias de Cálculo Pluvial")
st.markdown(
    "Plataforma interactiva para el dimensionamiento de colectores, bajantes"
    " (circulares y cuadradas), canaletas y canales abiertos."
)
st.divider()

# ==============================================================================
# PANEL LATERAL DE ENTRADA DE DATOS (st.sidebar)
# ==============================================================================
with st.sidebar:
    st.header("📋 Datos de Entrada")

    nombre_p = st.text_input(
        "Nombre del Proyecto", value="Nave Industrial - Planta Bajío"
    )
    maps_input = st.text_area(
        "Enlace, <iframe> o Coordenadas de Google Maps",
        height=100,
        value=(
            '<iframe'
            ' src="https://www.google.com/maps/embed?pb=!1m18!1m12!1m3!1d3702.481!2d-102.290017!3d22.024815"></iframe>'
        ),
    )

    # --- FUENTE METEOROLÓGICA ---
    st.markdown("**📊 Fuente Meteorológica**")
    excel_file = st.file_uploader(
        "Cargar archivo Excel de la Estación (Opcional)",
        type=["xlsx", "xls"],
    )

    # --- DATOS HIDROLÓGICOS ---
    st.markdown("---")
    area_m2 = st.number_input(
        "Área de Aportación / Cubierta (m²)",
        min_value=1.0,
        value=1200.0,
        step=50.0,
    )

    opciones_tr = [2, 5, 10, 20, 25, 50, 100, 200, 500]
    tr_seleccionado = st.selectbox(
        "Periodo de Retorno de Diseño (Tr en Años)",
        options=opciones_tr,
        index=2,
    )

    # --- SELECCIÓN DE COEFICIENTES DE ESCURRIMIENTO (C) ---
    st.markdown("---")
    st.markdown("**🌊 Coeficiente de Escurrimiento ($C$)**")
    tipo_superficie = st.selectbox(
        "Tipo de Superficie:",
        options=["Azotea / Techumbre", "Patios / Vialidad"],
        index=0,
    )

    coeficiente_C = TABLA_COEFICIENTES_C[tr_seleccionado][tipo_superficie]
    st.info(
        f"💡 **Coeficiente $C$ asignado:** `{coeficiente_C:.2f}` (para $T_r ="
        f" {tr_seleccionado}$ Años - {tipo_superficie})"
    )

    # --- TIEMPO DE CONCENTRACIÓN (tc) ---
    st.markdown("---")
    st.markdown("**⏱️ Tiempo de Concentración ($t_c$)**")
    modo_tc = st.radio(
        "Opción de cálculo para tc:",
        options=[
            "Predeterminado (5.0 min - Cubiertas/Áreas pequeñas)",
            "Calcular tc Real (Fórmula de Kirpich)",
        ],
        index=0,
        key="modo_tc_radio",
    )

    tc_adoptado = 5.0
    longitud_cauce_m = 35.0

    if "Kirpich" in modo_tc:
        longitud_cauce_m = st.number_input(
            "Longitud del cauce / vertiente (L en metros)",
            value=35.0,
            step=5.0,
            min_value=1.0,
            key="input_L",
        )

        tipo_pendiente = st.radio(
            "Definir desnivel por:",
            options=[
                "Diferencia de altura / Desnivel (H en m)",
                "Pendiente del cauce (S en %)",
            ],
            index=0,
            key="tipo_pend_radio",
        )

        if "Diferencia" in tipo_pendiente:
            desnivel_m = st.number_input(
                "Diferencia de altura (H en metros)",
                value=0.50,
                step=0.05,
                min_value=0.01,
                key="input_H",
            )
            tc_adoptado = calcular_tiempo_concentracion(
                longitud_m=longitud_cauce_m, diferencia_altura_m=desnivel_m
            )
        else:
            pendiente_cauce_pct = st.number_input(
                "Pendiente del cauce (S en %)",
                value=2.00,
                step=0.10,
                min_value=0.01,
                key="input_S",
            )
            tc_adoptado = calcular_tiempo_concentracion(
                longitud_m=longitud_cauce_m,
                pendiente_tramo_pct=pendiente_cauce_pct,
            )

        tc_adoptado = max(5.0, tc_adoptado)
        st.info(
            "⏱️ **Tiempo de concentración calculado ($t_c$):**"
            f" `{tc_adoptado:.2f} min`"
        )

    # --- DATOS HIDRÁULICOS Y SELECCIÓN DE SISTEMA ---
    st.markdown("---")
    st.markdown("**🛠️ Configuración del Drenaje**")
    tipo_elemento = st.selectbox(
        "Tipo de Sistema:",
        options=[
            "Colector Horizontal (Gravedad / Manning)",
            "Bajante Pluvial Vertical (Flujo Anular / BAP)",
            "Canaleta Rectangular (Lámina Galvanizada)",
            "Canal Abierto Rectangular (Concreto)",
            "Canal Abierto Trapezoidal (Concreto)",
        ],
        index=0,
    )

    # Selector de método con SMACNA / Ternium como primera opción predeterminada
    metodo_diseno_bap = "Criterio Hidráulico Tradicional (Manning / Flujo Anular)"
    if "Bajante" in tipo_elemento or "Canaleta" in tipo_elemento:
        metodo_diseno_bap = st.radio(
            "Método de Cálculo:",
            options=[
                "Método SMACNA / Ternium (Distancia Máx. entre Bajantes y Ejes)",
                "Criterio Hidráulico Tradicional (Manning / Flujo Anular)",
            ],
            index=0,
        )

    es_modo_smacna = "SMACNA" in metodo_diseno_bap

    # Inicialización de variables por defecto
    b_canal_cm, h_canal_cm, borde_libre_cm, talud_z = 30.0, 30.0, 4.0, 0.0
    diametro_manual_idx = None
    n_bajantes_manual = None
    pendiente_pct = 1.0

    # Variables para SMACNA
    W_smacna_m = longitud_cauce_m
    ancho_canalon_in = 10.0
    peralte_canalon_in = 10.0
    apoyo_polin_in = 4.0
    apoyo_muro_in = 2.0
    geometria_bajada_smacna = "Cuadrada / Rectangular (Lámina Galvanizada)"
    seccion_bajada_smacna = CATALOGO_MATERIALES["Lámina Galvanizada"]["tuberias"][7]
    dist_ejes_m = 9.0
    largo_nave_m = round(area_m2 / max(1.0, W_smacna_m), 1)

    if es_modo_smacna:
        material_nombre = "Lámina Galvanizada"
        st.info(
            "🧱 **Coeficiente de rugosidad ($n$):** `0.012`  | "
            " **Sistema:** `Canalón + Bajante (SMACNA)`"
        )
        st.markdown("---")
        st.markdown("**📐 Datos de Cubierta y Modulación (SMACNA)**")
        W_smacna_m = st.number_input(
            "Longitud de vertiente por drenar (W) [m]",
            min_value=1.0,
            max_value=300.0,
            value=float(longitud_cauce_m),
            step=1.0,
            help="Toma por defecto la longitud del cauce/escurrimiento.",
        )
        dist_ejes_m = st.number_input(
            "Distancia entre ejes de columnas [m]",
            min_value=1.0,
            max_value=50.0,
            value=9.0,
            step=0.5,
            help="Separación entre marcos/columnas donde se alojan las bajantes.",
        )
        largo_sugerido = round(area_m2 / max(1.0, W_smacna_m), 1)
        largo_nave_m = st.number_input(
            "Longitud total de la nave / canalón [m]",
            min_value=1.0,
            max_value=1000.0,
            value=max(9.0, largo_sugerido),
            step=3.0,
        )

        st.markdown("---")
        st.markdown("**📏 Geometría del Canalón (Pulgadas)**")
        ancho_canalon_in = st.number_input(
            "Ancho promedio del canalón (μ) [in]",
            min_value=3.0,
            max_value=48.0,
            value=10.0,
            step=1.0,
        )
        peralte_canalon_in = st.number_input(
            "Peralte del canalón [in]",
            min_value=3.0,
            max_value=48.0,
            value=10.0,
            step=1.0,
        )
        col_ap1, col_ap2 = st.columns(2)
        with col_ap1:
            apoyo_polin_in = st.number_input(
                "Apoyo en polín (a) [in]",
                min_value=0.0,
                max_value=20.0,
                value=4.0,
                step=0.5,
            )
        with col_ap2:
            apoyo_muro_in = st.number_input(
                "Apoyo en muro (b) [in]",
                min_value=0.0,
                max_value=20.0,
                value=2.0,
                step=0.5,
            )

        st.markdown("---")
        st.markdown("**🔽 Geometría de la Bajada Pluvial**")
        geometria_bajada_smacna = st.radio(
            "Tipo de Sección de Bajante:",
            options=[
                "Cuadrada / Rectangular (Lámina Galvanizada)",
                "Circular (Tubería PVC Sanitario)",
            ],
            index=0,
        )
        if "Cuadrada" in geometria_bajada_smacna:
            lista_sec_smacna = CATALOGO_MATERIALES["Lámina Galvanizada"]["tuberias"]
            idx_def = 7 if len(lista_sec_smacna) > 7 else 0
        else:
            lista_sec_smacna = CATALOGO_MATERIALES["PVC Sanitario Norma"]["tuberias"]
            idx_def = 3 if len(lista_sec_smacna) > 3 else 0
            material_nombre = "PVC Sanitario Norma"

        opciones_sec_smacna = [t["nom_in"] for t in lista_sec_smacna]
        idx_sec_smacna = st.selectbox(
            "Seleccionar Sección de Bajada Pluvial:",
            options=range(len(opciones_sec_smacna)),
            format_func=lambda x: opciones_sec_smacna[x],
            index=idx_def,
        )
        seccion_bajada_smacna = lista_sec_smacna[idx_sec_smacna]

    else:
        # --- MODO TRADICIONAL ---
        if "Bajante" in tipo_elemento:
            geometria_bap_trad = st.radio(
                "Geometría de Bajante Pluvial:",
                options=[
                    "Circular (PVC Sanitario Norma)",
                    "Cuadrada / Rectangular (Lámina Galvanizada)",
                ],
                index=0,
            )
            if "Circular" in geometria_bap_trad:
                materiales_disponibles = ["PVC Sanitario Norma"]
            else:
                materiales_disponibles = ["Lámina Galvanizada"]

            st.info(
                "📐 **Flujo Vertical:** Pendiente fija a 90° (sin inclinación"
                " horizontal)."
            )
            pendiente_pct = 100.0
        elif "Canaleta" in tipo_elemento:
            materiales_disponibles = ["Lámina Galvanizada"]
            pendiente_pct = st.number_input(
                "Pendiente de la Canaleta (%)",
                min_value=0.10,
                max_value=2.50,
                value=1.00,
                step=0.10,
                format="%.2f",
            )
        elif "Canal Abierto" in tipo_elemento:
            materiales_disponibles = ["Concreto"]
            pendiente_pct = st.number_input(
                "Pendiente del Canal (%)",
                min_value=0.01,
                max_value=10.00,
                value=0.50,
                step=0.10,
                format="%.2f",
            )
        else:
            materiales_disponibles = ["PVC Alcantarillado S-20", "PEAD Corrugado"]
            pendiente_pct = st.number_input(
                "Pendiente de la Tubería (%)",
                min_value=0.01,
                max_value=20.00,
                value=0.50,
                step=0.10,
                format="%.2f",
            )

        material_nombre = st.selectbox(
            "Material de Tubería / Sección",
            options=materiales_disponibles,
            index=0,
        )

        n_mat_actual = CATALOGO_MATERIALES[material_nombre]["n"]
        vmax_mat_actual = CATALOGO_MATERIALES[material_nombre]["v_max"]
        st.info(
            f"🧱 **Coeficiente de rugosidad ($n$):** `{n_mat_actual:.3f}`  | "
            f" **Vel. máx:** `{vmax_mat_actual:.1f} m/s`"
        )

        if "Canal Abierto" in tipo_elemento:
            st.markdown("---")
            st.markdown(f"**📐 Geometría del {tipo_elemento.split('(')[0]}**")
            b_canal_cm = st.number_input(
                "Ancho de base / plantilla (b) [cm]",
                min_value=10.0,
                max_value=500.0,
                value=50.0,
                step=5.0,
            )
            h_canal_cm = st.number_input(
                "Altura total (H) [cm]",
                min_value=10.0,
                max_value=300.0,
                value=40.0,
                step=5.0,
            )
            borde_libre_cm = st.number_input(
                "Borde libre de seguridad [cm]",
                min_value=0.0,
                max_value=float(max(0.0, h_canal_cm - 2.0)),
                value=min(5.0, float(max(0.0, h_canal_cm - 2.0))),
                step=1.0,
            )
            if "Trapezoidal" in tipo_elemento:
                talud_z = st.number_input(
                    "Talud lateral (z = H:V) [m/m]",
                    min_value=0.1,
                    max_value=3.0,
                    value=1.0,
                    step=0.1,
                    help=(
                        "Relación de inclinación horizontal por unidad vertical."
                    ),
                )
                angulo_deg = math.degrees(math.atan(1.0 / talud_z))
                st.caption(
                    f"📐 **Ángulo equivalente:** `{angulo_deg:.1f}°` con respecto"
                    " a la horizontal."
                )

        elif "Canaleta" in tipo_elemento:
            st.markdown("---")
            st.markdown("**📐 Geometría de Canaleta Rectangular**")
            b_canal_cm = st.number_input(
                "Ancho de base (b) [cm]",
                min_value=5.0,
                max_value=150.0,
                value=30.0,
                step=1.0,
            )
            h_canal_cm = st.number_input(
                "Altura total (H) [cm]",
                min_value=5.0,
                max_value=150.0,
                value=30.0,
                step=1.0,
            )
            borde_libre_cm = st.number_input(
                "Borde Libre de seguridad [cm]",
                min_value=0.0,
                max_value=float(max(0.0, h_canal_cm - 1.0)),
                value=min(4.0, float(max(0.0, h_canal_cm - 1.0))),
                step=0.5,
            )

        elif "Bajante" in tipo_elemento:
            st.markdown("---")
            st.markdown("**🔢 Cantidad de Bajantes Pluviales (BAP)**")
            modo_cantidad_bap = st.radio(
                "Modo de Cantidad de Bajantes:",
                options=[
                    "Automático (Sugerir N° de bajadas)",
                    "Manual (Definir N° de bajadas)",
                ],
                index=0,
            )
            if modo_cantidad_bap == "Manual (Definir N° de bajadas)":
                n_bajantes_manual = st.number_input(
                    "Número de Bajantes Pluviales (BAP):",
                    min_value=1,
                    max_value=200,
                    value=4,
                    step=1,
                )

            st.markdown("---")
            st.markdown("**⭕ Sección / Diámetro de Bajante**")
            modo_diametro = st.radio(
                "Modo de Selección de Sección:",
                options=[
                    "Sugerido Automático",
                    "Seleccionar Sección Manualmente",
                ],
                index=0,
            )
            if modo_diametro == "Seleccionar Sección Manualmente":
                opciones_tuberias = [
                    f"{tub['nom_in']} (D_hid: {tub['d_int_mm']} mm)"
                    for tub in CATALOGO_MATERIALES[material_nombre]["tuberias"]
                ]
                idx_seleccionado = st.selectbox(
                    "Elige la Sección Comercial a Evaluar:",
                    options=range(len(opciones_tuberias)),
                    format_func=lambda x: opciones_tuberias[x],
                    index=3 if len(opciones_tuberias) > 3 else 0,
                )
                diametro_manual_idx = idx_seleccionado

        else:
            st.markdown("---")
            modo_diametro = st.radio(
                "Modo de Selección de Diámetro",
                options=[
                    "Sugerido Automático",
                    "Seleccionar Diámetro Manualmente",
                ],
                index=0,
            )
            if modo_diametro == "Seleccionar Diámetro Manualmente":
                opciones_tuberias = [
                    f"{tub['nom_in']} (D_int: {tub['d_int_mm']} mm)"
                    for tub in CATALOGO_MATERIALES[material_nombre]["tuberias"]
                ]
                idx_seleccionado = st.selectbox(
                    "Elige el Diámetro Comercial a Evaluar:",
                    options=range(len(opciones_tuberias)),
                    format_func=lambda x: opciones_tuberias[x],
                    index=0,
                )
                diametro_manual_idx = idx_seleccionado

# ==============================================================================
# ÁREA PRINCIPAL DE RESULTADOS Y VISUALIZACIÓN
# ==============================================================================

st.subheader("🔍 Previsualización y Análisis en Tiempo Real")

lat, lon = 22.024815, -102.290017
try:
    lat, lon = extraer_coordenadas(maps_input)
    st.success(
        f"📍 Coordenadas detectadas: **Lat:** {lat:.6f} | **Lon:** {lon:.6f}"
    )
except Exception:
    st.warning(
        "⚠️ Coordenadas no detectadas. Usando ubicación predeterminada."
    )

archivo_bytes = excel_file.getvalue() if excel_file is not None else None
estacion_nom, I_ref, tabla_tr = procesar_excel_estacion(
    archivo_bytes,
    lat,
    lon,
    tr_seleccionado=tr_seleccionado,
    tc_adoptado=tc_adoptado,
)

if excel_file is not None:
    st.info(f"📄 **Estación desde Excel:** {estacion_nom}")
else:
    st.info(f"📍 **Estación por cercanía:** {estacion_nom}")

st.markdown(
    f"**Matriz de Intensidades (mm/h) — [$T_r = {tr_seleccionado}$ Años / $t_c ="
    f" {tc_adoptado:.2f}$ min / $I_{{\\text{{diseño}}}} = {I_ref}$ mm/h"
    f" ({I_ref / 25.4:.2f} in/hr)]:**"
)
st.dataframe(tabla_tr, use_container_width=True)

if es_modo_smacna:
    hidro = calcular_smacna_canalon_bajante(
        W_m=W_smacna_m,
        I_mm_h=I_ref,
        ancho_canalon_in=ancho_canalon_in,
        peralte_canalon_in=peralte_canalon_in,
        apoyo_polin_in=apoyo_polin_in,
        apoyo_muro_in=apoyo_muro_in,
        geometria_bajada=geometria_bajada_smacna,
        seccion_bajada_info=seccion_bajada_smacna,
        dist_ejes_m=dist_ejes_m,
        largo_nave_m=largo_nave_m,
        coeficiente_C=coeficiente_C,
    )
else:
    hidro = calcular_hidraulica(
        area_m2,
        I_ref,
        pendiente_pct,
        material_nombre,
        coeficiente_C=coeficiente_C,
        diametro_manual_idx=diametro_manual_idx,
        n_bajantes_manual=n_bajantes_manual,
        tipo_elemento=tipo_elemento,
        b_canal_cm=b_canal_cm,
        h_canal_cm=h_canal_cm,
        borde_libre_cm=borde_libre_cm,
        talud_z=talud_z,
    )

st.subheader("⚡ Resultados del Cálculo")

if hidro["tiene_alertas"]:
    st.markdown(hidro["mensaje_estado"])
else:
    st.success(hidro["mensaje_estado"])

m1, m2, m3, m4, m5 = st.columns(5)

if hidro.get("es_smacna", False):
    m1.metric(
        "Dist. Máx. ($DB_{\\text{máx}}$)",
        f"{hidro['DB_max_m']} m ({hidro['DB_max_ft']} ft)",
        hidro["gobierna"],
    )
    m2.metric(
        "Modulación en Columnas",
        (
            f"Cada {hidro['ejes_por_bajante']} eje(s) ({hidro['sep_real_m']} m)"
            if hidro["ejes_por_bajante"] > 0
            else f"Excede eje ({hidro['dist_ejes_m']} m)"
        ),
        delta=(
            f"Ejes a {hidro['dist_ejes_m']} m (OK)"
            if hidro["ejes_por_bajante"] > 0
            else f"Eje > {hidro['DB_max_m']} m"
        ),
        delta_color="normal" if hidro["ejes_por_bajante"] > 0 else "inverse",
    )
    m3.metric(
        "Total Bajantes Línea",
        f"{hidro['n_bajantes']} BAP ({hidro['nom_in'].split('(')[0].strip()})",
        f"Área BAP: {hidro['area_bajada_in2']} in²",
    )
    m4.metric(
        "Canalón y Desarrollo",
        f"{hidro['ancho_canalon_in']:.0f}\" x {hidro['peralte_canalon_in']:.0f}\"",
        f"Des. Lámina: {hidro['desarrollo_in']}\" ({hidro['desarrollo_cm']} cm)",
    )
    m5.metric(
        "Gasto Vertiente ($Q$)",
        f"{hidro['Q_lps']} L/s",
        f"Q/bajada: {hidro['q_unit_lps']} L/s",
    )

    with st.expander(
        "📋 Ver Desglose Completo de Cálculos SMACNA / Ternium", expanded=True
    ):
        tabla_resumen_smacna = [
            {
                "Concepto": "Longitud de vertiente por drenar (W)",
                "Unidades Inglesas": f"{hidro['W_ft']:.2f} ft",
                "Sistema Métrico": f"{hidro['W_m']:.2f} m",
            },
            {
                "Concepto": "Intensidad de lluvia de diseño (I)",
                "Unidades Inglesas": f"{hidro['I_in_hr']:.2f} in/hr",
                "Sistema Métrico": f"{I_ref:.2f} mm/h",
            },
            {
                "Concepto": "Ancho promedio del canalón (μ)",
                "Unidades Inglesas": f"{hidro['mu_ft']:.2f} ft ({hidro['ancho_canalon_in']:.1f} in)",
                "Sistema Métrico": f"{hidro['b_canal_cm']:.1f} cm",
            },
            {
                "Concepto": "Peralte del canalón y relación (m = P/μ)",
                "Unidades Inglesas": f"{hidro['peralte_ft']:.2f} ft (m = {hidro['m_ratio']:.2f})",
                "Sistema Métrico": f"{hidro['h_canal_cm']:.1f} cm",
            },
            {
                "Concepto": "Desarrollo total de lámina (incluye apoyos a y b)",
                "Unidades Inglesas": f"{hidro['desarrollo_in']:.1f} in",
                "Sistema Métrico": f"{hidro['desarrollo_cm']:.1f} cm",
            },
            {
                "Concepto": f"Área de la bajada pluvial ({hidro['nom_in']})",
                "Unidades Inglesas": f"{hidro['area_bajada_in2']:.2f} in²",
                "Sistema Métrico": f"{hidro['area_bajada_in2'] * 6.4516:.1f} cm²",
            },
            {
                "Concepto": "Distancia entre bajantes (Capacidad del canalón)",
                "Unidades Inglesas": f"{hidro['DB_canalon_ft']:.2f} ft",
                "Sistema Métrico": f"{hidro['DB_canalon_m']:.2f} m",
            },
            {
                "Concepto": "Distancia entre bajantes (Capacidad de la bajante)",
                "Unidades Inglesas": f"{hidro['DB_bajante_ft']:.2f} ft",
                "Sistema Métrico": f"{hidro['DB_bajante_m']:.2f} m",
            },
            {
                "Concepto": f"DISTANCIA ENTRE BAJANTES MÁXIMA ({hidro['gobierna'].upper()})",
                "Unidades Inglesas": f"{hidro['DB_max_ft']:.2f} ft",
                "Sistema Métrico": f"{hidro['DB_max_m']:.2f} m",
            },
        ]
        st.dataframe(
            tabla_resumen_smacna, use_container_width=True, hide_index=True
        )

elif hidro.get("es_canal_abierto", False):
    m1.metric(
        "Gasto ($Q$)",
        f"{hidro['Q_lps']} L/s",
        f"Cap. útil: {hidro['cap_lps']} L/s",
    )
    m2.metric(
        "Sección Canal",
        f"{hidro['nom_in']}",
        f"B. Libre: {hidro['borde_libre_cm']} cm",
    )
    m3.metric(
        "Ocupación Útil",
        f"{hidro['porcentaje_llenado']}%",
        f"y = {hidro['y_cm']} cm",
    )
    m4.metric(
        "Velocidad Real",
        f"{hidro['vel_m_s']} m/s",
        delta=(
            f" OK (≤ {hidro['v_max_normativa']:.1f} m/s)"
            if hidro["cumple_velocidad"]
            else f" EXCEDIDA (> {hidro['v_max_normativa']:.1f} m/s)"
        ),
        delta_color="normal" if hidro["cumple_velocidad"] else "inverse",
    )
    m5.metric(
        "Rugosidad ($n$)",
        f"{hidro['n_manning']:.3f}",
        material_nombre,
    )
elif hidro.get("es_canaleta", False):
    m1.metric(
        "Gasto ($Q$)",
        f"{hidro['Q_lps']} L/s",
        f"Cap. útil: {hidro['cap_lps']} L/s",
    )
    m2.metric(
        "Sección Canaleta",
        f"Rect. {hidro['b_canal_cm']:.0f}x{hidro['h_canal_cm']:.0f} cm",
        f"B. Libre: {hidro['borde_libre_cm']} cm",
    )
    m3.metric(
        "Ocupación Útil",
        f"{hidro['porcentaje_llenado']}%",
        f"y = {hidro['y_cm']} cm",
    )
    m4.metric(
        "Velocidad Real",
        f"{hidro['vel_m_s']} m/s",
        delta=(
            f" OK (≤ {hidro['v_max_normativa']:.1f} m/s)"
            if hidro["cumple_velocidad"]
            else f" EXCEDIDA (> {hidro['v_max_normativa']:.1f} m/s)"
        ),
        delta_color="normal" if hidro["cumple_velocidad"] else "inverse",
    )
    m5.metric(
        "Rugosidad ($n$)",
        f"{hidro['n_manning']:.3f}",
        material_nombre,
    )
elif hidro["es_bajante"]:
    m1.metric(
        "Gasto Total ($Q$)",
        f"{hidro['Q_lps']} L/s",
        f"Q/bajada: {hidro['q_unit_lps']} L/s",
    )
    m2.metric(
        "Bajantes y Sección",
        f"{hidro['n_bajantes']} BAP de {hidro['nom_in'].split('(')[0].strip()}",
        f"D_hid: {hidro['d_int_mm']} mm | Sug: {hidro['n_bajantes_recomendadas']} BAP",
    )
    m3.metric(
        "Ocupación Anular",
        f"{hidro['porcentaje_llenado']}%",
        f"Anillo: {hidro.get('espesor_mm', round(hidro['y_cm'] * 10.0, 1))} mm (≤ 33.3%)",
    )
    m4.metric(
        "Velocidad Terminal",
        f"{hidro['vel_m_s']} m/s",
        delta=(
            f" OK (≤ {hidro['v_max_normativa']:.1f} m/s)"
            if hidro["cumple_velocidad"]
            else " ALTA VELOCIDAD"
        ),
        delta_color="normal" if hidro["cumple_velocidad"] else "inverse",
    )
    m5.metric(
        "Rugosidad ($n$)",
        f"{hidro['n_manning']:.3f}",
        material_nombre,
    )

    with st.expander(
        "📊 Ver tabla comparativa de bajantes recomendadas por sección comercial",
        expanded=False,
    ):
        st.dataframe(
            hidro.get("tabla_comparativa_bap", []),
            use_container_width=True,
            hide_index=True,
        )
else:
    m1.metric(
        "Gasto ($Q$)",
        f"{hidro['Q_lps']} L/s",
        f"Cap. lleno: {hidro['cap_lps']} L/s",
    )
    m2.metric(
        "Tubería Eval.",
        f"{hidro['nom_in']}",
        f"D_int: {hidro['d_int_mm']} mm",
    )
    m3.metric(
        "Llenado ($y/D$)",
        f"{hidro['porcentaje_llenado']}%",
        f"y = {hidro['y_cm']} cm",
    )
    m4.metric(
        "Velocidad Real",
        f"{hidro['vel_m_s']} m/s",
        delta=(
            f" OK (≤ {hidro['v_max_normativa']:.1f} m/s)"
            if hidro["cumple_velocidad"]
            else f" EXCEDIDA (> {hidro['v_max_normativa']:.1f} m/s)"
        ),
        delta_color="normal" if hidro["cumple_velocidad"] else "inverse",
    )
    m5.metric(
        "Rugosidad ($n$)",
        f"{hidro['n_manning']:.3f}",
        material_nombre,
    )

# --- VISUALIZACIÓN GRÁFICA ACOTADA (ESTILO CAD CON VALORES POSITIVOS) ---
st.markdown("### 📊 Croquis Acotado de la Sección Hidráulica")

img_secundaria_bytes = None

if hidro.get("es_smacna", False):
    fig_canalon = generar_grafico_seccion(
        tipo_elemento="Canaleta Rectangular",
        b=hidro["b_canal_cm"],
        H=hidro["h_canal_cm"],
        y=hidro["y_cm"],
        borde_libre_cm=hidro["borde_libre_cm"],
    )
    buf_can = io.BytesIO()
    fig_canalon.savefig(buf_can, format="png", dpi=160, bbox_inches="tight")
    img_seccion_bytes = buf_can.getvalue()
    plt.close(fig_canalon)

    fig_bap = generar_grafico_seccion(
        tipo_elemento="Bajante Pluvial Vertical",
        b=hidro["ancho_bajada_cm"],
        H=hidro["largo_bajada_cm"],
        y=0.8,
        D_cm=hidro["d_int_mm"] / 10.0,
        es_bajada_rect=hidro.get("es_bajada_rect", False),
        ancho_bajada_cm=hidro.get("ancho_bajada_cm"),
        largo_bajada_cm=hidro.get("largo_bajada_cm"),
    )
    buf_bap = io.BytesIO()
    fig_bap.savefig(buf_bap, format="png", dpi=160, bbox_inches="tight")
    img_secundaria_bytes = buf_bap.getvalue()
    plt.close(fig_bap)

    c_esp1, c_graf1, c_graf2, c_esp2 = st.columns([0.4, 1.2, 1.2, 0.4])
    with c_graf1:
        st.caption("**Sección del Canalón Acotada (cm)**")
        st.image(img_seccion_bytes, use_container_width=True)
    with c_graf2:
        st.caption("**Sección de la Bajante Pluvial Acotada (cm)**")
        st.image(img_secundaria_bytes, use_container_width=True)
else:
    fig_seccion = generar_grafico_seccion(
        tipo_elemento=tipo_elemento,
        b=b_canal_cm,
        H=h_canal_cm,
        y=hidro["y_cm"],
        talud_z=talud_z,
        D_cm=(
            hidro["d_int_mm"] / 10.0
            if not (hidro.get("es_canaleta") or hidro.get("es_canal_abierto"))
            else None
        ),
        borde_libre_cm=borde_libre_cm,
        es_bajada_rect=hidro.get("es_bajada_rect", False),
        ancho_bajada_cm=hidro.get("ancho_bajada_cm"),
        largo_bajada_cm=hidro.get("largo_bajada_cm"),
    )

    img_buffer = io.BytesIO()
    fig_seccion.savefig(img_buffer, format="png", dpi=160, bbox_inches="tight")
    img_seccion_bytes = img_buffer.getvalue()
    plt.close(fig_seccion)

    col_izq, col_centro, col_der = st.columns([1, 1.3, 1])
    with col_centro:
        st.image(img_seccion_bytes, use_container_width=True)

st.divider()

doc_buffer = generar_memoria_docx(
    nombre_p,
    lat,
    lon,
    area_m2,
    pendiente_pct,
    tr_seleccionado,
    tc_adoptado,
    tipo_superficie,
    coeficiente_C,
    material_nombre,
    estacion_nom,
    I_ref,
    tabla_tr,
    hidro,
    tipo_elemento,
    img_seccion_bytes=img_seccion_bytes,
    img_secundaria_bytes=img_secundaria_bytes,
)

st.download_button(
    label="📄 Descargar Memoria en Word (.docx)",
    data=doc_buffer,
    file_name=f"Memoria_Calculo_{nombre_p.replace(' ', '_')}.docx",
    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    type="primary",
)