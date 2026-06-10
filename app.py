import streamlit as st
import pandas as pd
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

.login-card {
    background: white; border-radius: 12px;
    padding: 2.5rem; max-width: 420px;
    margin: 4rem auto;
    border: 1px solid #E5E9F0;
    box-shadow: 0 4px 20px rgba(0,0,0,0.08);
}
.login-title {
    text-align: center; font-size: 1.3rem;
    font-weight: 700; color: #003087;
    margin-bottom: 1.5rem;
}

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

.badge-registrador {
    display:inline-block; background:#E8F0FE; color:#1A56DB;
    padding:0.3rem 0.9rem; border-radius:20px;
    font-size:0.85rem; font-weight:600; margin-bottom:1rem;
}
.badge-validador {
    display:inline-block; background:#FEF3C7; color:#92400E;
    padding:0.3rem 0.9rem; border-radius:20px;
    font-size:0.85rem; font-weight:600; margin-bottom:1rem;
}

.val-aprobado { background:#E6F4EA; border:1px solid #34A853; border-radius:8px; padding:0.75rem 1rem; color:#1E6E3A; font-weight:500; margin:0.3rem 0; }
.val-nocumple { background:#FDE8E8; border:1px solid #EA4335; border-radius:8px; padding:0.75rem 1rem; color:#9B1C1C; font-weight:500; margin:0.3rem 0; }
.val-info     { background:#E8F0FE; border:1px solid #4285F4; border-radius:8px; padding:0.75rem 1rem; color:#1A56DB; font-weight:500; margin:0.3rem 0; }

.stButton > button {
    background: linear-gradient(135deg, #003087, #0052CC);
    color: white; border: none;
    padding: 0.75rem 2.5rem; border-radius: 8px;
    font-weight: 600; font-size: 1rem; width: 100%;
}
.btn-aprobar > button { background: linear-gradient(135deg, #1E6E3A, #34A853) !important; }
.btn-rechazar > button { background: linear-gradient(135deg, #9B1C1C, #EA4335) !important; }
</style>
""", unsafe_allow_html=True)

# ─── USUARIOS ────────────────────────────────────────────────────────────────
USUARIOS = {
    # cedula : (contraseña, rol, nombre)
    "123456789": ("123456789", "registrador", "Registrador"),
    "validador": ("validador2026", "validador", "Validador General"),
}

ARCHIVO_RESPUESTAS = "respuestas_desplazamiento.xlsx"
TOLERANCIA = 500

# ─── GENERAR PLANTILLA EXCEL ──────────────────────────────────────────────────
def generar_plantilla():
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    import io

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "PLANTILLA DESPLAZAMIENTO"

    campos = [
        "Documento de identidad", "Codigo Sipab", "Nombre PGR",
        "Fecha (DD/MM/AAAA)", "Nombre AGR", "Empresa Cliente",
        "Cronograma", "Secuencia", "Empresa de transporte",
        "Recorrido (Ida / Ida y vuelta)", "Origen", "Destino",
        "Hora Inicio (HH:MM)", "Hora Fin (HH:MM)", "Frecuencia",
        "Valor Pasajes Intermunicipal", "Transporte Interno",
        "Desayuno", "Almuerzo/Cena", "Hospedaje",
        "Detalles de las rutas",
    ]
    ejemplos = [
        "1012345678", "7643", "JUAN PEREZ", "04/06/2026",
        "MARIA GOMEZ", "BOLIVAR S.A.", "100001", "1",
        "EXPRESO BOGOTA", "Ida y vuelta", "BOGOTA", "MEDELLIN",
        "06:00", "10:00", "Unica visita",
        "150000", "30000", "10000", "15000", "0",
        "Ruta por autopista principal",
    ]

    hf  = PatternFill(start_color="003087", end_color="003087", fill_type="solid")
    ef  = PatternFill(start_color="EBF2FF", end_color="EBF2FF", fill_type="solid")
    brd = Border(left=Side(style="thin"), right=Side(style="thin"),
                 top=Side(style="thin"),  bottom=Side(style="thin"))

    for i, campo in enumerate(campos, 1):
        c = ws.cell(row=1, column=i, value=campo)
        c.fill      = hf
        c.font      = Font(color="FFFFFF", bold=True, size=11)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border    = brd
        ws.column_dimensions[get_column_letter(i)].width = max(len(campo) + 4, 20)

    for i, ej in enumerate(ejemplos, 1):
        c = ws.cell(row=2, column=i, value=ej)
        c.fill      = ef
        c.font      = Font(color="555555", italic=True, size=10)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border    = brd

    ws.row_dimensions[1].height = 40
    ws.row_dimensions[2].height = 25
    ws.freeze_panes = "A2"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ─── SESIÓN ───────────────────────────────────────────────────────────────────
if "autenticado" not in st.session_state:
    st.session_state.autenticado = False
    st.session_state.rol = None
    st.session_state.nombre_usuario = None
    st.session_state.usuario_id = None

# ─── FUNCIONES ────────────────────────────────────────────────────────────────
def cargar_respuestas():
    if os.path.exists(ARCHIVO_RESPUESTAS):
        return pd.read_excel(ARCHIVO_RESPUESTAS)
    return pd.DataFrame()

def guardar_respuesta(datos: dict):
    df_existente = cargar_respuestas()
    df_nuevo = pd.DataFrame([datos])
    df_final = pd.concat([df_existente, df_nuevo], ignore_index=True)
    df_final.to_excel(ARCHIVO_RESPUESTAS, index=False)

def actualizar_decision(idx, decision, observacion, validador):
    df = cargar_respuestas()
    df.at[idx, "Aprueba/No aprueba"]       = decision
    df.at[idx, "Observación validador"]     = observacion
    df.at[idx, "Validado por"]              = validador
    df.at[idx, "Fecha validación"]          = datetime.now()
    df.to_excel(ARCHIVO_RESPUESTAS, index=False)

def validar_tarifa(df_desp, origen, destino, recorrido, val_pasajes, transporte_interno):
    if df_desp.empty:
        return "SIN DATOS", None, "Tabla de tarifas no cargada"
    o = origen.strip().upper()
    d = destino.strip().upper()
    fila = df_desp[(df_desp["ORIGEN"] == o) & (df_desp["DESTINO"] == d)]
    if fila.empty:
        fila = df_desp[(df_desp["ORIGEN"] == d) & (df_desp["DESTINO"] == o)]
    if fila.empty:
        return "NO CUMPLE", None, "Ruta no encontrada en tabla de tarifas"
    fila = fila.iloc[0]
    tarifa = fila["VALOR_IDA_VUELTA"] if "ida y vuelta" in recorrido.lower() else fila["VALOR_IDA"]
    try:
        tarifa = float(tarifa)
    except:
        return "NO CUMPLE", None, "Tarifa no disponible"
    valor = val_pasajes + transporte_interno
    if valor <= tarifa + TOLERANCIA:
        return "APROBADO", tarifa, f"Valor ${valor:,.0f} ≤ Tarifa ${tarifa:,.0f}"
    else:
        return "NO CUMPLE", tarifa, f"Valor ${valor:,.0f} supera tarifa ${tarifa:,.0f} en ${valor-tarifa:,.0f}"

@st.cache_data
def cargar_tarifas():
    ruta = "DESPLAZAMIENTOS_2026.xlsx"
    if not os.path.exists(ruta):
        return pd.DataFrame(), pd.DataFrame()
    try:
        wb = pd.read_excel(ruta, sheet_name="DESPLAZAMIENTOS 2026", skiprows=1, engine="openpyxl")
        wb.columns = wb.columns.str.strip().str.upper()
        wb = wb.rename(columns={wb.columns[0]:"ORIGEN", wb.columns[1]:"DESTINO",
                                  wb.columns[2]:"VALOR_IDA", wb.columns[3]:"VALOR_IDA_VUELTA"})
        wb = wb[["ORIGEN","DESTINO","VALOR_IDA","VALOR_IDA_VUELTA"]].dropna(subset=["ORIGEN","DESTINO"])
        wb["ORIGEN"]  = wb["ORIGEN"].astype(str).str.strip().str.upper()
        wb["DESTINO"] = wb["DESTINO"].astype(str).str.strip().str.upper()
        CIUDADES_DEF = sorted(set(wb["ORIGEN"].tolist() + wb["DESTINO"].tolist()))
        return wb, CIUDADES_DEF
    except:
        return pd.DataFrame(), []

df_desp, CIUDADES = cargar_tarifas()
if not CIUDADES:
    CIUDADES = ["BOGOTÁ","MEDELLÍN","CALI","BARRANQUILLA","BUCARAMANGA",
                "CARTAGENA","VILLAVICENCIO","YOPAL","ACACÍAS","PUERTO GAITÁN"]

FRECUENCIAS = ["Única visita","Diaria","Semanal","Quincenal","Mensual"]
RECORRIDOS  = ["Ida","Ida y vuelta"]

# ─── PANTALLA LOGIN ───────────────────────────────────────────────────────────
def pantalla_login():
    st.markdown("""
    <div class="header-bar">
        <h1>🗺️ Relación de Desplazamiento</h1>
        <p>ARL Bolívar · Outsourcing Adecco</p>
    </div>
    """, unsafe_allow_html=True)

    st.markdown('<div class="login-card">', unsafe_allow_html=True)
    st.markdown('<div class="login-title">🔐 Iniciar Sesión</div>', unsafe_allow_html=True)

    usuario  = st.text_input("Usuario (Cédula)", placeholder="Ej: 123456789")
    password = st.text_input("Contraseña", type="password", placeholder="Ingresa tu contraseña")

    if st.button("Ingresar"):
        if usuario in USUARIOS:
            pwd, rol, nombre = USUARIOS[usuario]
            if password == pwd:
                st.session_state.autenticado    = True
                st.session_state.rol            = rol
                st.session_state.nombre_usuario = nombre
                st.session_state.usuario_id     = usuario
                st.rerun()
            else:
                st.error("❌ Contraseña incorrecta")
        else:
            st.error("❌ Usuario no encontrado")

    st.markdown('</div>', unsafe_allow_html=True)

# ─── VISTA REGISTRADOR ────────────────────────────────────────────────────────
def vista_registrador():
    st.markdown("""
    <div class="header-bar">
        <h1>🗺️ Relación de Desplazamiento</h1>
        <p>ARL Bolívar · Outsourcing Adecco — Registro de desplazamientos</p>
    </div>
    """, unsafe_allow_html=True)

    col_badge, col_logout = st.columns([8, 1])
    with col_badge:
        st.markdown(f'<span class="badge-registrador">👤 Registrador — {st.session_state.nombre_usuario}</span>',
                    unsafe_allow_html=True)
    with col_logout:
        if st.button("Salir"):
            st.session_state.autenticado = False
            st.rerun()

    tab_form, tab_masivo, tab_mis_registros = st.tabs(["📝 Nuevo Registro", "📤 Cargue Masivo", "📋 Mis Registros"])

    with tab_form:
        with st.form("form_desplazamiento", clear_on_submit=True):



            # DATOS PGR
            st.markdown('<div class="section-card">', unsafe_allow_html=True)
            st.markdown('<div class="section-title">👤 Datos del PGR</div>', unsafe_allow_html=True)
            c1, c2, c3 = st.columns(3)
            with c1: documento    = st.text_input("Documento de Identidad", placeholder="Ej: 1012345678")
            with c2: codigo_sipab = st.text_input("Código SIPAB")
            with c3: nombre_pgr   = st.text_input("Nombre PGR")
            c4, c5, c6 = st.columns(3)
            with c4: fecha           = st.date_input("Fecha", value=datetime.today())
            with c5: nombre_agr      = st.text_input("Nombre AGR")
            with c6: empresa_cliente = st.text_input("Empresa Cliente")
            c7, c8 = st.columns(2)
            with c7: cronograma = st.text_input("Cronograma")
            with c8: secuencia  = st.text_input("Secuencia")
            st.markdown('</div>', unsafe_allow_html=True)

            # DESPLAZAMIENTO
            st.markdown('<div class="section-card">', unsafe_allow_html=True)
            st.markdown('<div class="section-title">🚗 Datos del Desplazamiento</div>', unsafe_allow_html=True)
            c9, c10 = st.columns(2)
            with c9:  empresa   = st.text_input("Empresa de transporte")
            with c10: recorrido = st.selectbox("Recorrido", RECORRIDOS)
            c11, c12 = st.columns(2)
            with c11: origen  = st.selectbox("Origen",  ["Seleccione..."] + CIUDADES)
            with c12: destino = st.selectbox("Destino", ["Seleccione..."] + CIUDADES)
            c13, c14, c15 = st.columns(3)
            with c13: hora_inicio = st.time_input("Hora Inicio", value=time(6,0))
            with c14: hora_fin    = st.time_input("Hora Fin",    value=time(10,0))
            with c15: frecuencia  = st.selectbox("Frecuencia", FRECUENCIAS)
            st.markdown('</div>', unsafe_allow_html=True)

            # VALORES
            st.markdown('<div class="section-card">', unsafe_allow_html=True)
            st.markdown('<div class="section-title">💰 Valores</div>', unsafe_allow_html=True)
            c16, c17, c18, c19 = st.columns(4)
            with c16: val_pasajes        = st.number_input("Pasajes Intermunicipal ($)", min_value=0, step=1000, value=0, format="%d")
            with c17: transporte_interno = st.number_input("Transporte Interno ($)",     min_value=0, step=1000, value=0, format="%d")
            with c18: desayuno_val       = st.number_input("Desayuno ($)",               min_value=0, step=1000, value=0, format="%d")
            with c19: almuerzo_cena_val  = st.number_input("Almuerzo/Cena ($)",          min_value=0, step=1000, value=0, format="%d")
            c20, c21 = st.columns([1,3])
            with c20: hospedaje_val  = st.number_input("Hospedaje ($)", min_value=0, step=1000, value=0, format="%d")
            with c21: detalles_rutas = st.text_area("Detalles de las Rutas", height=80)
            total = val_pasajes + transporte_interno + desayuno_val + almuerzo_cena_val + hospedaje_val
            st.info(f"💵 **Total estimado: ${total:,.0f}**")
            st.markdown('</div>', unsafe_allow_html=True)

            submitted = st.form_submit_button("✅ Enviar Registro")

        st.success("✅ Registro enviado. Quedará pendiente de aprobación por el validador.")


    with tab_masivo:
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        st.markdown('<div class="section-title">📤 Cargue Masivo de Registros</div>', unsafe_allow_html=True)

        st.markdown("**Paso 1 — Descarga la plantilla, diligénciala y vuelve aquí para subirla.**")
        st.download_button(
            label="⬇️ Descargar Plantilla Excel",
            data=generar_plantilla(),
            file_name="plantilla_desplazamiento.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_plantilla"
        )

        st.markdown("---")
        st.markdown("**Paso 2 — Sube el archivo diligenciado.**")
        archivo_masivo = st.file_uploader(
            "Selecciona el archivo Excel",
            type=["xlsx", "xls"],
            key="uploader_masivo"
        )

        if archivo_masivo is not None:
            st.markdown("**Paso 3 — Haz clic en Subir para guardar los registros.**")
            if st.button("📤 Subir Registros", key="btn_subir_masivo"):
                try:
                    import io as _io
                    df_masivo = pd.read_excel(_io.BytesIO(archivo_masivo.read()), engine="openpyxl")
                    df_masivo.columns = df_masivo.columns.str.strip()
                    # Quitar fila de ejemplo
                    df_masivo = df_masivo[df_masivo.iloc[:, 0].astype(str).str.strip() != "1012345678"]
                    df_masivo = df_masivo.dropna(how="all")

                    errores  = []
                    guardados = 0

                    for i, fila in df_masivo.iterrows():
                        doc = str(fila.get("Documento de identidad", "")).strip()
                        pgr = str(fila.get("Nombre PGR", "")).strip()

                        if not doc or doc == "nan" or not pgr or pgr == "nan":
                            errores.append(f"Fila {i+2}: Documento o Nombre PGR vacío")
                            continue

                        orig = str(fila.get("Origen", "")).strip().upper()
                        dest = str(fila.get("Destino", "")).strip().upper()
                        rec  = str(fila.get("Recorrido (Ida / Ida y vuelta)", "")).strip()

                        try:
                            vp  = float(str(fila.get("Valor Pasajes Intermunicipal", 0) or 0).replace(",","").replace("$",""))
                            ti  = float(str(fila.get("Transporte Interno", 0) or 0).replace(",","").replace("$",""))
                            des = float(str(fila.get("Desayuno", 0) or 0).replace(",","").replace("$",""))
                            alm = float(str(fila.get("Almuerzo/Cena", 0) or 0).replace(",","").replace("$",""))
                            hos = float(str(fila.get("Hospedaje", 0) or 0).replace(",","").replace("$",""))
                        except:
                            errores.append(f"Fila {i+2}: Error en valores numéricos")
                            continue

                        total_fila = vp + ti + des + alm + hos
                        res3, tarifa, msg3 = validar_tarifa(df_desp, orig, dest, rec, vp, ti)

                        datos = {
                            "Marca temporal":           datetime.now(),
                            "Registrado por":           st.session_state.usuario_id,
                                    "Documento de identidad":   doc,
                            "Codigo Sipab":             str(fila.get("Codigo Sipab", "")).strip(),
                            "Nombre PGR":               pgr,
                            "Fecha":                    fila.get("Fecha (DD/MM/AAAA)", ""),
                            "Nombre AGR":               str(fila.get("Nombre AGR", "")).strip(),
                            "Empresa Cliente":          str(fila.get("Empresa Cliente", "")).strip(),
                            "Cronograma":               str(fila.get("Cronograma", "")).strip(),
                            "Secuencia":                str(fila.get("Secuencia", "")).strip(),
                            "Empresa":                  str(fila.get("Empresa de transporte", "")).strip(),
                            "Recorrido":                rec,
                            "Origen":                   orig,
                            "Destino":                  dest,
                            "Hora Inicio":              str(fila.get("Hora Inicio (HH:MM)", "")).strip(),
                            "Hora Fin":                 str(fila.get("Hora Fin (HH:MM)", "")).strip(),
                            "Frecuencia":               str(fila.get("Frecuencia", "")).strip(),
                            "Valor Pasajes":            vp,
                            "Transporte Interno":       ti,
                            "Desayuno":                 des,
                            "Almuerzo/Cena":            alm,
                            "Hospedaje":                hos,
                            "Detalles de las rutas":    str(fila.get("Detalles de las rutas", "")).strip(),
                            "Total":                    total_fila,
                            "Validacion tarifa":        res3,
                            "Tarifa permitida":         tarifa,
                            "Detalle validacion":       msg3,
                            "Estado":                   "PENDIENTE",
                            "Aprueba/No aprueba":       "",
                            "Observacion validador":    "",
                            "Validado por":             "",
                            "Fecha validacion":         "",
                        }
                        guardar_respuesta(datos)
                        guardados += 1

                    if guardados > 0:
                        st.success(f"✅ {guardados} registro(s) cargados exitosamente. Quedan pendientes de aprobación.")
                    if errores:
                        st.warning("⚠️ Las siguientes filas tuvieron errores y no se cargaron:")
                        for e in errores:
                            st.write(f"- {e}")

                except Exception as e:
                    st.error(f"❌ Error procesando el archivo: {e}")

        st.markdown('</div>', unsafe_allow_html=True)

    with tab_mis_registros:
        df = cargar_respuestas()
        if df.empty:
            st.info("📭 Aún no tienes registros.")
        else:
            mis = df[df["Registrado por"] == st.session_state.usuario_id].copy() if "Registrado por" in df.columns else df.copy()
            st.markdown(f"**{len(mis)} registro(s)**")

            for col in ["Estado","Aprueba/No aprueba"]:
                if col in mis.columns:
                    opciones = ["Todos"] + sorted(mis[col].dropna().unique().tolist())
                    filtro = st.selectbox(f"Filtrar por {col}", opciones, key=f"f_{col}")
                    if filtro != "Todos":
                        mis = mis[mis[col] == filtro]

            st.dataframe(mis, use_container_width=True, hide_index=True)

# ─── VISTA VALIDADOR ──────────────────────────────────────────────────────────
def vista_validador():
    st.markdown("""
    <div class="header-bar">
        <h1>🗺️ Relación de Desplazamiento</h1>
        <p>ARL Bolívar · Outsourcing Adecco — Panel de Validación</p>
    </div>
    """, unsafe_allow_html=True)

    col_badge, col_logout = st.columns([8,1])
    with col_badge:
        st.markdown(f'<span class="badge-validador">🔍 Validador — {st.session_state.nombre_usuario}</span>',
                    unsafe_allow_html=True)
    with col_logout:
        if st.button("Salir"):
            st.session_state.autenticado = False
            st.rerun()

    tab_pendientes, tab_historial = st.tabs(["⏳ Pendientes de Validación", "📋 Historial Completo"])

    df = cargar_respuestas()

    with tab_pendientes:
        if df.empty:
            st.info("📭 No hay registros aún.")
        else:
            pendientes = df[df["Aprueba/No aprueba"].fillna("") == ""].copy()
            st.markdown(f"**{len(pendientes)} registro(s) pendiente(s) de validación**")

            if pendientes.empty:
                st.success("✅ No hay registros pendientes.")
            else:
                for idx, row in pendientes.iterrows():
                    with st.expander(f"📄 {row.get('Nombre PGR','—')} | {row.get('Origen','—')} → {row.get('Destino','—')} | ${row.get('Total',0):,.0f} | {row.get('Fecha','')}"):

                        c1, c2, c3 = st.columns(3)
                        with c1:
                            st.markdown(f"**Documento:** {row.get('Documento de identidad','—')}")
                            st.markdown(f"**SIPAB:** {row.get('Código Sipab','—')}")
                            st.markdown(f"**AGR:** {row.get('Nombre AGR','—')}")
                        with c2:
                            st.markdown(f"**Empresa Cliente:** {row.get('Empresa Cliente','—')}")
                            st.markdown(f"**Cronograma:** {row.get('Cronograma','—')}")
                            st.markdown(f"**Recorrido:** {row.get('Recorrido','—')}")
                        with c3:
                            st.markdown(f"**Pasajes:** ${row.get('Valor Pasajes',0):,.0f}")
                            st.markdown(f"**Transporte interno:** ${row.get('Transporte Interno',0):,.0f}")
                            st.markdown(f"**Total:** **${row.get('Total',0):,.0f}**")

                        # Resultado validación automática
                        res3 = row.get("Validación tarifa","—")
                        css  = "val-aprobado" if res3 == "APROBADO" else "val-nocumple"
                        st.markdown(f'<div class="{css}">🔍 <b>Validación automática tarifa:</b> {res3} — {row.get("Detalle validación","")}</div>',
                                    unsafe_allow_html=True)

                        # Decisión del validador
                        st.markdown("---")
                        obs = st.text_area("Observación (opcional)", key=f"obs_{idx}", placeholder="Escribe una observación...")

                        bc1, bc2 = st.columns(2)
                        with bc1:
                            st.markdown('<div class="btn-aprobar">', unsafe_allow_html=True)
                            if st.button("✅ Aprobar", key=f"apr_{idx}"):
                                actualizar_decision(idx, "APROBADO", obs, st.session_state.nombre_usuario)
                                st.success("✅ Registro aprobado.")
                                st.rerun()
                            st.markdown('</div>', unsafe_allow_html=True)
                        with bc2:
                            st.markdown('<div class="btn-rechazar">', unsafe_allow_html=True)
                            if st.button("❌ Rechazar", key=f"rec_{idx}"):
                                actualizar_decision(idx, "RECHAZADO", obs, st.session_state.nombre_usuario)
                                st.error("❌ Registro rechazado.")
                                st.rerun()
                            st.markdown('</div>', unsafe_allow_html=True)

    with tab_historial:
        if df.empty:
            st.info("📭 No hay registros.")
        else:
            st.markdown(f"**{len(df)} registro(s) en total**")

            cf1, cf2 = st.columns(2)
            with cf1:
                if "Nombre PGR" in df.columns:
                    pgr_f = st.multiselect("Filtrar por PGR",
                        options=sorted(df["Nombre PGR"].dropna().unique()))
            with cf2:
                if "Aprueba/No aprueba" in df.columns:
                    est_f = st.multiselect("Filtrar por Estado",
                        options=sorted(df["Aprueba/No aprueba"].dropna().unique()))

            df_f = df.copy()
            if pgr_f: df_f = df_f[df_f["Nombre PGR"].isin(pgr_f)]
            if est_f: df_f = df_f[df_f["Aprueba/No aprueba"].isin(est_f)]

            st.dataframe(df_f, use_container_width=True, hide_index=True)

            if os.path.exists(ARCHIVO_RESPUESTAS):
                with open(ARCHIVO_RESPUESTAS,"rb") as f:
                    st.download_button("⬇️ Descargar Excel", f.read(),
                        file_name="respuestas_desplazamiento.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ─── ROUTER ───────────────────────────────────────────────────────────────────
if not st.session_state.autenticado:
    pantalla_login()
elif st.session_state.rol == "registrador":
    vista_registrador()
elif st.session_state.rol == "validador":
    vista_validador()
