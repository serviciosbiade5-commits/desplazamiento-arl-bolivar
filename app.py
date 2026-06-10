import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, time
import os

# ─── CONFIG ───────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Relación de Desplazamiento | ARL Bolívar",
    page_icon="🗺️",
    layout="wide"
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
.main { background-color: #F7F8FA; }
.header-bar {
    background: linear-gradient(135deg, #003087 0%, #0052CC 100%);
    padding: 2rem 2.5rem; border-radius: 12px;
    margin-bottom: 2rem; color: white;
}
.header-bar h1 { margin: 0; font-size: 1.8rem; font-weight: 700; }
.header-bar p  { margin: 0.3rem 0 0; font-size: 0.95rem; opacity: 0.85; }
.section-card {
    background: white; border-radius: 10px;
    padding: 1.5rem 2rem; margin-bottom: 1.5rem;
    border: 1px solid #E5E9F0;
    box-shadow: 0 1px 4px rgba(0,0,0,0.05);
}
.section-title {
    font-size: 1rem; font-weight: 600; color: #003087;
    border-left: 4px solid #0052CC;
    padding-left: 0.75rem; margin-bottom: 1.2rem;
}
.stButton > button {
    background: linear-gradient(135deg, #003087, #0052CC);
    color: white; border: none;
    padding: 0.75rem 2.5rem; border-radius: 8px;
    font-weight: 600; font-size: 1rem; width: 100%;
}
.val-aprobado { background:#E6F4EA; border:1px solid #34A853; border-radius:8px; padding:0.75rem 1rem; color:#1E6E3A; font-weight:500; margin:0.3rem 0; }
.val-nocumple { background:#FDE8E8; border:1px solid #EA4335; border-radius:8px; padding:0.75rem 1rem; color:#9B1C1C; font-weight:500; margin:0.3rem 0; }
.val-info     { background:#E8F0FE; border:1px solid #4285F4; border-radius:8px; padding:0.75rem 1rem; color:#1A56DB; font-weight:500; margin:0.3rem 0; }
</style>
""", unsafe_allow_html=True)

# ─── HEADER ───────────────────────────────────────────────────────────────────
st.markdown("""
<div class="header-bar">
    <h1>🗺️ Relación de Desplazamiento</h1>
    <p>ARL Bolívar · Outsourcing Adecco — Registro y validación de desplazamientos</p>
</div>
""", unsafe_allow_html=True)

# ─── CONFIGURACIÓN ────────────────────────────────────────────────────────────
TOLERANCIA = 500
ARCHIVO_RESPUESTAS = "respuestas_desplazamiento.xlsx"

# Tarifas alimentación
ALIMENTACION = {
    "desayuno":    {"municipio": 10289.71, "ciudad": 16890},
    "almuerzo":    {"municipio": 14406.05, "ciudad": 22736},
    "cena":        {"municipio": 14406.05, "ciudad": 22736},
    "hospedaje":   {"municipio": 97743,    "ciudad": 126120},
}

# ─── CARGAR DATOS DE VALIDACIÓN ───────────────────────────────────────────────
@st.cache_data
def cargar_tarifas():
    ruta = "DESPLAZAMIENTOS_2026.xlsx"
    if not os.path.exists(ruta):
        return pd.DataFrame(), pd.DataFrame()
    wb_desp = pd.read_excel(ruta, sheet_name="DESPLAZAMIENTOS 2026",
                            skiprows=1, engine="openpyxl")
    wb_desp.columns = wb_desp.columns.str.strip().str.upper()
    wb_desp = wb_desp.rename(columns={
        wb_desp.columns[0]: "ORIGEN",
        wb_desp.columns[1]: "DESTINO",
        wb_desp.columns[2]: "VALOR_IDA",
        wb_desp.columns[3]: "VALOR_IDA_VUELTA",
    })
    wb_desp = wb_desp[["ORIGEN","DESTINO","VALOR_IDA","VALOR_IDA_VUELTA"]].dropna(subset=["ORIGEN","DESTINO"])
    wb_desp["ORIGEN"]  = wb_desp["ORIGEN"].astype(str).str.strip().str.upper()
    wb_desp["DESTINO"] = wb_desp["DESTINO"].astype(str).str.strip().str.upper()

    wb_taxi = pd.read_excel(ruta, sheet_name="TAXIS 2026", engine="openpyxl")
    wb_taxi.columns = wb_taxi.columns.str.strip().str.upper()
    wb_taxi = wb_taxi.rename(columns={
        wb_taxi.columns[0]: "CIUDAD",
        wb_taxi.columns[1]: "VALOR_IDA",
        wb_taxi.columns[2]: "VALOR_IDA_VUELTA",
    })
    wb_taxi = wb_taxi[["CIUDAD","VALOR_IDA","VALOR_IDA_VUELTA"]].dropna(subset=["CIUDAD"])
    wb_taxi["CIUDAD"] = wb_taxi["CIUDAD"].astype(str).str.strip().str.upper()
    return wb_desp, wb_taxi

@st.cache_data
def cargar_programacion():
    ruta = "Programacion_Adecco.xlsx"
    if not os.path.exists(ruta):
        return pd.DataFrame()
    df = pd.read_excel(ruta, sheet_name="GLOBAL", engine="openpyxl")
    df.columns = df.columns.str.strip().str.upper()
    return df

def cargar_respuestas():
    if os.path.exists(ARCHIVO_RESPUESTAS):
        return pd.read_excel(ARCHIVO_RESPUESTAS)
    return pd.DataFrame()

def guardar_respuesta(datos: dict):
    df_existente = cargar_respuestas()
    df_nuevo = pd.DataFrame([datos])
    df_final = pd.concat([df_existente, df_nuevo], ignore_index=True)
    df_final.to_excel(ARCHIVO_RESPUESTAS, index=False)

# ─── FUNCIONES DE VALIDACIÓN ──────────────────────────────────────────────────

def validar_tarifa(df_desp, origen, destino, recorrido, val_pasajes, transporte_interno):
    """Validación 3: Cumple/No cumple con la tarifa establecida"""
    o = origen.strip().upper()
    d = destino.strip().upper()

    fila = df_desp[(df_desp["ORIGEN"] == o) & (df_desp["DESTINO"] == d)]
    if fila.empty:
        fila = df_desp[(df_desp["ORIGEN"] == d) & (df_desp["DESTINO"] == o)]

    if fila.empty:
        return "NO CUMPLE", None, "Ruta no encontrada en tabla de tarifas"

    fila = fila.iloc[0]
    if "ida y vuelta" in recorrido.lower():
        tarifa = fila["VALOR_IDA_VUELTA"]
    else:
        tarifa = fila["VALOR_IDA"]

    try:
        tarifa = float(tarifa)
    except:
        return "NO CUMPLE", None, "Tarifa no disponible para esta ruta"

    valor_solicitado = val_pasajes + transporte_interno

    if valor_solicitado <= tarifa + TOLERANCIA:
        return "APROBADO", tarifa, f"Valor solicitado ${valor_solicitado:,.0f} ≤ Tarifa ${tarifa:,.0f}"
    else:
        diferencia = valor_solicitado - tarifa
        return "NO CUMPLE", tarifa, f"Valor solicitado ${valor_solicitado:,.0f} supera tarifa ${tarifa:,.0f} en ${diferencia:,.0f}"

def validar_global_autorizacion(df_prog, cronograma):
    """Validación 1: Cumple/No cumple con la autorización en el global"""
    if df_prog.empty or not cronograma:
        return "NO CUMPLE", "No se pudo consultar el global (archivo no cargado o cronograma vacío)"

    try:
        crono_num = int(str(cronograma).strip())
    except:
        return "NO CUMPLE", "Cronograma inválido"

    if "NUMERO_CRONOGRAMA" not in df_prog.columns:
        return "NO CUMPLE", "Columna NUMERO_CRONOGRAMA no encontrada"

    fila = df_prog[df_prog["NUMERO_CRONOGRAMA"] == crono_num]
    if fila.empty:
        return "NO CUMPLE", f"Cronograma {crono_num} no encontrado en el global"

    autoriza = str(fila.iloc[0].get("AUTORIZA_DESPLAZAMIENTO", "N")).strip().upper()
    if autoriza in ["S", "SI", "SÍ", "Y", "YES"]:
        return "CUMPLE", f"Cronograma {crono_num} tiene autorización de desplazamiento"
    else:
        return "NO CUMPLE", f"Cronograma {crono_num} NO tiene autorización de desplazamiento (valor: {autoriza})"

def validar_global_transporte(df_prog, cronograma, val_pasajes, transporte_interno):
    """Validación 2: Cumple/No cumple con el valor de transporte en el global"""
    if df_prog.empty or not cronograma:
        return "NO CUMPLE", "No se pudo consultar el global"

    try:
        crono_num = int(str(cronograma).strip())
    except:
        return "NO CUMPLE", "Cronograma inválido"

    fila = df_prog[df_prog["NUMERO_CRONOGRAMA"] == crono_num]
    if fila.empty:
        return "NO CUMPLE", f"Cronograma {crono_num} no encontrado"

    try:
        valor_global = float(fila.iloc[0].get("VALOR_TRANSPORTE", 0) or 0)
    except:
        valor_global = 0

    valor_solicitado = val_pasajes + transporte_interno

    if valor_global <= 0:
        return "NO CUMPLE", f"El global no tiene valor de transporte registrado para cronograma {crono_num}"

    if valor_solicitado <= valor_global + TOLERANCIA:
        return "CUMPLE", f"Valor solicitado ${valor_solicitado:,.0f} ≤ Valor global ${valor_global:,.0f}"
    else:
        return "NO CUMPLE", f"Valor solicitado ${valor_solicitado:,.0f} supera valor global ${valor_global:,.0f}"

def decision_final(val1, val2, val3):
    if val1 == "CUMPLE" and val2 == "CUMPLE" and val3 == "APROBADO":
        return "✅ APROBAR"
    elif val1 == "NO CUMPLE":
        return "❌ NO APROBAR — Sin autorización en el global"
    elif val2 == "NO CUMPLE":
        return "❌ NO APROBAR — Valor supera lo autorizado en el global"
    elif val3 == "NO CUMPLE":
        return "❌ NO APROBAR — Valor supera tarifa establecida"
    else:
        return "⚠️ REVISAR MANUALMENTE"

# ─── CARGAR ARCHIVOS ──────────────────────────────────────────────────────────
df_desp, df_taxi = cargar_tarifas()
df_prog = cargar_programacion()

FRECUENCIAS = ["Única visita", "Diaria", "Semanal", "Quincenal", "Mensual"]
RECORRIDOS  = ["Ida", "Ida y vuelta"]

# Ciudades desde la tabla de tarifas
if not df_desp.empty:
    ciudades_origen  = sorted(df_desp["ORIGEN"].dropna().unique().tolist())
    ciudades_destino = sorted(df_desp["DESTINO"].dropna().unique().tolist())
    CIUDADES = sorted(set(ciudades_origen + ciudades_destino))
else:
    CIUDADES = sorted(["ACACÍAS","BOGOTÁ","MEDELLÍN","CALI","BARRANQUILLA",
        "BUCARAMANGA","CARTAGENA","VILLAVICENCIO","PUERTO GAITÁN","YOPAL"])

# ─── TABS ─────────────────────────────────────────────────────────────────────
tab_form, tab_historial = st.tabs(["📝 Nuevo Registro", "📋 Historial"])

with tab_form:
    with st.form("form_desplazamiento", clear_on_submit=True):

        # ── CARGUE MASIVO ─────────────────────────────────────────────────
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        st.markdown('<div class="section-title">📤 Cargue de Información</div>', unsafe_allow_html=True)
        cargue_masivo = st.radio("¿Desea realizar un cargue masivo de información?",
                                  ["No", "Sí"], horizontal=True)
        archivo_masivo = None
        if cargue_masivo == "Sí":
            archivo_masivo = st.file_uploader("Subir archivo",
                type=["xlsx","xls","csv"],
                help="Sube el archivo con los registros masivos")
        st.markdown('</div>', unsafe_allow_html=True)

        # ── DATOS PGR ─────────────────────────────────────────────────────
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        st.markdown('<div class="section-title">👤 Datos del PGR</div>', unsafe_allow_html=True)

        c1, c2, c3 = st.columns(3)
        with c1: documento    = st.text_input("Documento de Identidad", placeholder="Ej: 1012345678")
        with c2: codigo_sipab = st.text_input("Código SIPAB", placeholder="Ej: 7643")
        with c3: nombre_pgr   = st.text_input("Nombre PGR", placeholder="Nombre completo")

        c4, c5, c6 = st.columns(3)
        with c4: fecha          = st.date_input("Fecha", value=datetime.today())
        with c5: nombre_agr     = st.text_input("Nombre AGR", placeholder="Nombre del AGR")
        with c6: empresa_cliente = st.text_input("Empresa Cliente", placeholder="Nombre de la empresa")

        c7, c8 = st.columns(2)
        with c7: cronograma = st.text_input("Cronograma", placeholder="Número de cronograma")
        with c8: secuencia  = st.text_input("Secuencia",  placeholder="Número de secuencia")

        st.markdown('</div>', unsafe_allow_html=True)

        # ── DESPLAZAMIENTO ────────────────────────────────────────────────
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        st.markdown('<div class="section-title">🚗 Datos del Desplazamiento</div>', unsafe_allow_html=True)

        c9, c10 = st.columns(2)
        with c9:  empresa   = st.text_input("Empresa", placeholder="Empresa de transporte")
        with c10: recorrido = st.selectbox("Especifique Recorrido", RECORRIDOS)

        c11, c12 = st.columns(2)
        with c11: origen  = st.selectbox("Origen",  ["Seleccione..."] + CIUDADES)
        with c12: destino = st.selectbox("Destino", ["Seleccione..."] + CIUDADES)

        c13, c14, c15 = st.columns(3)
        with c13: hora_inicio = st.time_input("Hora Inicio", value=time(6, 0))
        with c14: hora_fin    = st.time_input("Hora Fin",    value=time(10, 0))
        with c15: frecuencia  = st.selectbox("Frecuencia", FRECUENCIAS)

        st.markdown('</div>', unsafe_allow_html=True)

        # ── VALORES ───────────────────────────────────────────────────────
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        st.markdown('<div class="section-title">💰 Valores</div>', unsafe_allow_html=True)

        c16, c17, c18, c19 = st.columns(4)
        with c16: val_pasajes        = st.number_input("Valor Pasajes Intermunicipal ($)", min_value=0, step=1000, value=0, format="%d")
        with c17: transporte_interno = st.number_input("Transporte Interno ($)",           min_value=0, step=1000, value=0, format="%d")
        with c18: desayuno_val       = st.number_input("Desayuno ($)",                     min_value=0, step=1000, value=0, format="%d")
        with c19: almuerzo_cena_val  = st.number_input("Almuerzo/Cena ($)",                min_value=0, step=1000, value=0, format="%d")

        c20, c21 = st.columns([1, 3])
        with c20: hospedaje_val = st.number_input("Hospedaje ($)", min_value=0, step=1000, value=0, format="%d")
        with c21: detalles_rutas = st.text_area("Detalles de las Rutas",
                    placeholder="Describa el detalle de las rutas tomadas...", height=80)

        total = val_pasajes + transporte_interno + desayuno_val + almuerzo_cena_val + hospedaje_val
        st.info(f"💵 **Total estimado: ${total:,.0f}**")
        st.markdown('</div>', unsafe_allow_html=True)

        submitted = st.form_submit_button("✅ Guardar y Validar Registro")

    # ── VALIDACIONES (fuera del form para mostrar resultados) ─────────────
    if submitted:
        if origen == "Seleccione..." or destino == "Seleccione...":
            st.error("⚠️ Por favor selecciona el Origen y Destino.")
        elif not nombre_pgr or not documento:
            st.error("⚠️ El Documento de Identidad y Nombre PGR son obligatorios.")
        else:
            st.markdown("---")
            st.markdown("### 🔍 Resultados de Validación")

            # Validación 1
            res1, msg1 = validar_global_autorizacion(df_prog, cronograma)
            css1 = "val-aprobado" if res1 == "CUMPLE" else "val-nocumple"
            st.markdown(f'<div class="{css1}">✔️ <b>Val. 1 — Autorización en el Global:</b> {res1}<br><small>{msg1}</small></div>', unsafe_allow_html=True)

            # Validación 2
            res2, msg2 = validar_global_transporte(df_prog, cronograma, val_pasajes, transporte_interno)
            css2 = "val-aprobado" if res2 == "CUMPLE" else "val-nocumple"
            st.markdown(f'<div class="{css2}">✔️ <b>Val. 2 — Valor de transporte en el Global:</b> {res2}<br><small>{msg2}</small></div>', unsafe_allow_html=True)

            # Validación 3
            res3, tarifa, msg3 = validar_tarifa(df_desp, origen, destino, recorrido, val_pasajes, transporte_interno)
            css3 = "val-aprobado" if res3 == "APROBADO" else "val-nocumple"
            tarifa_str = f"${tarifa:,.0f}" if tarifa else "N/A"
            st.markdown(f'<div class="{css3}">✔️ <b>Val. 3 — Tarifa establecida:</b> {res3}<br><small>{msg3} | Tarifa permitida: {tarifa_str}</small></div>', unsafe_allow_html=True)

            # Decisión final
            decision = decision_final(res1, res2, res3)
            st.markdown(f'<div class="val-info" style="margin-top:1rem; font-size:1.05rem;">🏁 <b>Decisión recomendada:</b> {decision}</div>', unsafe_allow_html=True)

            # Guardar
            datos = {
                "Marca temporal":                           datetime.now(),
                "Cargue masivo":                            cargue_masivo,
                "Documento de identidad":                   documento,
                "Código Sipab":                             codigo_sipab,
                "Nombre PGR":                               nombre_pgr,
                "Fecha":                                    fecha,
                "Nombre AGR":                               nombre_agr,
                "Empresa Cliente":                          empresa_cliente,
                "Cronograma":                               cronograma,
                "Secuencia":                                secuencia,
                "Empresa":                                  empresa,
                "Especifique Recorrido":                    recorrido,
                "Origen":                                   origen,
                "Destino":                                  destino,
                "Hora Inicio desplazamiento":               str(hora_inicio),
                "Hora Fin desplazamiento":                  str(hora_fin),
                "Frecuencia":                               frecuencia,
                "Valor Pasajes Intermunicipal":             val_pasajes,
                "Transporte Interno":                       transporte_interno,
                "Desayuno":                                 desayuno_val,
                "Almuerzo/Cena":                            almuerzo_cena_val,
                "Hospedaje":                                hospedaje_val,
                "Detalles de las rutas":                    detalles_rutas,
                "Total":                                    total,
                "Cumple autorización global":               res1,
                "Cumple valor transporte global":           res2,
                "Cumple tarifa establecida":                res3,
                "Tarifa permitida":                         tarifa,
                "Decisión recomendada por automatización":  decision,
                "Aprueba/No aprueba":                       "",
            }
            guardar_respuesta(datos)
            st.success("✅ Registro guardado exitosamente.")

# ─── HISTORIAL ────────────────────────────────────────────────────────────────
with tab_historial:
    df = cargar_respuestas()
    if df.empty:
        st.info("📭 Aún no hay registros guardados.")
    else:
        st.markdown(f"**{len(df)} registro(s) encontrado(s)**")

        cf1, cf2 = st.columns(2)
        with cf1:
            if "Nombre PGR" in df.columns:
                pgr_f = st.multiselect("Filtrar por PGR",
                    options=sorted(df["Nombre PGR"].dropna().unique()))
        with cf2:
            if "Decisión recomendada por automatización" in df.columns:
                dec_f = st.multiselect("Filtrar por Decisión",
                    options=sorted(df["Decisión recomendada por automatización"].dropna().unique()))

        df_f = df.copy()
        if pgr_f: df_f = df_f[df_f["Nombre PGR"].isin(pgr_f)]
        if dec_f: df_f = df_f[df_f["Decisión recomendada por automatización"].isin(dec_f)]

        st.dataframe(df_f, use_container_width=True, hide_index=True)

        if os.path.exists(ARCHIVO_RESPUESTAS):
            with open(ARCHIVO_RESPUESTAS, "rb") as f:
                st.download_button("⬇️ Descargar Excel", f.read(),
                    file_name="respuestas_desplazamiento.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
