import datetime
import io
import os
from google.oauth2.service_account import Credentials
import gspread
import openpyxl
from openpyxl.drawing.image import Image as OpenPyxlImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
from PIL import Image as PILImage
import streamlit as st

st.set_page_config(
    page_title="Control de Proceso - Cocina Dulce",
    page_icon="📋",
    layout="wide",
)


# --- CONEXIÓN A GOOGLE SHEETS ---
@st.cache_resource
def conectar_google_sheets():
    try:
        scope = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ]

        if "gcp_service_account" in st.secrets:
            creds_dict = dict(st.secrets["gcp_service_account"])

            if "private_key" in creds_dict:
                pk = str(creds_dict["private_key"])
                pk = pk.strip().strip('"').strip("'")
                pk = pk.replace("\\n", "\n")
                creds_dict["private_key"] = pk

            creds = Credentials.from_service_account_info(
                creds_dict, scopes=scope
            )
            client = gspread.authorize(creds)
            spreadsheet_id = st.secrets["SPREADSHEET_ID"]
            sheet = client.open_by_key(spreadsheet_id).sheet1
            return sheet
        else:
            st.error(
                "No se encontraron las credenciales 'gcp_service_account' en st.secrets."
            )
            return None

    except Exception as e:
        st.error(f"Error al conectar con Google Sheets: {e}")
        return None


sheet = conectar_google_sheets()

FILE_PATH = "CONTROL PROCESOS COCINA DULCE.xlsx"

PRODUCTOS_DEFAULT = [
    "GLACE DE KEKE LIMÓN",
    "COMPOTA DE MUFFINS DE MANZANA",
    "COMPOTA DE KEKE GINGER",
    "COMPOTA DE CROCANTE DE MANZANA",
    "RELLENO PIE DE LIMON STBX",
    "ALMIBAR TORTA DE CHOCOLATE",
    "CREMA DE QUESO",
    "CREMA DE QUESO DE KEKES",
    "FUDGE",
    "FUDGE SILVESTRE",
    "FUDGE LIQUIDO PARA TIENDA",
    "GANACHE AGENDA",
    "MERMELADA MIXTA",
    "ZANAHORIA RALLADA",
    "ALMIBAR RED VELVET",
    "TRUFA DE CHOCOLATE 2023",
    "MANÁ",
    "MANJAR LIQUIDO",
    "MANJAR SILVESTRE",
    "MANJAR CASERO",
    "CHOCOLATE RALLADO BITTER",
    "CHOCOLATE RALLADO BLANCO MA",
    "RELLENO DE PIE DE LIMON MARIA",
    "ALMIBAR TRES LECHES VAINILLA",
    "ALMIBAR TRES LECHES CHOCOLATE",
]

EQUIPOS_DEFAULT = [
    "BATIDORA",
    "MARMITA",
    "HORNO RATIONAL",
    "ROBOTCOUPE",
    "NO APLICA",
    "OTRO",
]

RESPONSABLES_DEFAULT = [
    "CINDY VILLA",
    "JHOJAN VENTURA",
    "DANIEL CUADROS",
    "CESAR VILCHEZ",
    "LUCIA CHIQUIRENI",
    "JENNY RIVAS",
    "OTRO",
]


@st.cache_data(ttl=60)
def cargar_catalogos():
    productos = PRODUCTOS_DEFAULT
    equipos = EQUIPOS_DEFAULT
    responsables = RESPONSABLES_DEFAULT

    if os.path.exists(FILE_PATH):
        try:
            xls = pd.ExcelFile(FILE_PATH)
            for sheet_name in xls.sheet_names:
                df = pd.read_excel(FILE_PATH, sheet_name=sheet_name)
                for idx, row in df.iterrows():
                    row_vals = [
                        str(v).strip() for v in row.values if pd.notna(v)
                    ]
                    if (
                        "PRODUCTO" in row_vals
                        and "EQUIPO UTILIZADO" in row_vals
                    ):
                        header_idx = idx
                        df_data = pd.read_excel(
                            FILE_PATH,
                            sheet_name=sheet_name,
                            skiprows=header_idx + 1,
                        )
                        df_data.columns = [
                            str(c).strip() for c in df.iloc[header_idx].values
                        ]

                        if "PRODUCTO" in df_data.columns:
                            prods = (
                                df_data["PRODUCTO"]
                                .dropna()
                                .astype(str)
                                .str.strip()
                                .unique()
                                .tolist()
                            )
                            if prods:
                                productos = prods

                        if "EQUIPO UTILIZADO" in df_data.columns:
                            eqs = (
                                df_data["EQUIPO UTILIZADO"]
                                .dropna()
                                .astype(str)
                                .str.strip()
                                .unique()
                                .tolist()
                            )
                            if eqs:
                                equipos = list(dict.fromkeys(eqs + ["OTRO"]))

                        if "RESPONSABLE" in df_data.columns:
                            resps = (
                                df_data["RESPONSABLE"]
                                .dropna()
                                .astype(str)
                                .str.strip()
                                .unique()
                                .tolist()
                            )
                            if resps:
                                responsables = list(
                                    dict.fromkeys(resps + ["OTRO"])
                                )
                        break
        except Exception:
            pass

    return productos, equipos, responsables


lista_productos, lista_equipos, lista_responsables = cargar_catalogos()

st.title("📋 FORMATO CONTROL DE PROCESO COCINA DULCE")
st.markdown("---")

tab1, tab2 = st.tabs(["📝 Nuevo Registro", "⚙️ Gestionar / Modificar Registros"])


# --- FUNCIÓN DE RESETEO SEGURO ---
def resetear_claves_form():
    st.session_state["form_fecha_p"] = datetime.date.today()
    st.session_state["form_lote"] = ""
    st.session_state["form_batch"] = ""
    st.session_state["form_responsable_sel"] = (
        lista_responsables[0] if lista_responsables else ""
    )
    st.session_state["form_responsable_text"] = ""
    st.session_state["form_producto"] = (
        lista_productos[0] if lista_productos else ""
    )
    st.session_state["form_equipo_sel"] = (
        lista_equipos[0] if lista_equipos else ""
    )
    st.session_state["form_equipo_text"] = ""
    st.session_state["form_brix"] = ""
    st.session_state["form_hora_inicio"] = datetime.time(8, 0)
    st.session_state["form_hora_termino"] = datetime.time(8, 15)
    st.session_state["form_temp_equipo"] = 0.0
    st.session_state["form_vel_agitador"] = ""
    st.session_state["form_estado_obs"] = "CONFORME"
    st.session_state["form_obs_detalle"] = ""
    st.session_state["form_obs_adicional"] = ""
    st.session_state["form_estado_area"] = "CONFORME"
    st.session_state["form_obs_area"] = ""


# Si se activó la bandera de reseteo en la ejecución anterior, limpiamos antes de instanciar widgets
if st.session_state.get("necesita_reset", False):
    resetear_claves_form()
    st.session_state["necesita_reset"] = False

# Inicialización primaria si es primera carga
if "form_fecha_p" not in st.session_state:
    resetear_claves_form()

# ---------------------------------------------------------
# PESTAÑA 1: NUEVO REGISTRO
# ---------------------------------------------------------
with tab1:
    with st.form("form_control_proceso", clear_on_submit=False):
        st.subheader("1. Información General del Proceso")
        col1, col2, col3, col4 = st.columns(4)

        with col1:
            fecha_p = st.date_input(
                "F.P (Fecha de Producción)", key="form_fecha_p"
            )
        with col2:
            lote = st.text_input(
                "LOTE", placeholder="Ej. L-20260901", key="form_lote"
            )
        with col3:
            batch = st.text_input(
                "BATCH", placeholder="Ej. B-01", key="form_batch"
            )
        with col4:
            responsable_sel = st.selectbox(
                "RESPONSABLE",
                options=lista_responsables,
                key="form_responsable_sel",
            )
            if responsable_sel == "OTRO":
                responsable_text = st.text_input(
                    "Especifique Responsable",
                    placeholder="Nombre completo",
                    key="form_responsable_text",
                )
                responsable_final = responsable_text
            else:
                responsable_final = responsable_sel

        st.subheader("2. Selección de Producto, Equipo y Parámetros Iniciales")
        col5, col6, col_brix = st.columns([2, 2, 1])

        with col5:
            producto = st.selectbox(
                "PRODUCTO", options=lista_productos, key="form_producto"
            )
        with col6:
            equipo_sel = st.selectbox(
                "EQUIPO UTILIZADO",
                options=lista_equipos,
                key="form_equipo_sel",
            )
            if equipo_sel == "OTRO":
                equipo_text = st.text_input(
                    "Especifique Equipo",
                    placeholder="Nombre del equipo",
                    key="form_equipo_text",
                )
                equipo_final = equipo_text
            else:
                equipo_final = equipo_sel
        with col_brix:
            brix = st.text_input(
                "BRIX (°Bx)", placeholder="Ej. 65.5", key="form_brix"
            )

        st.subheader("3. Parámetros de Control y Tiempos")
        col7, col8, col9 = st.columns(3)

        with col7:
            hora_inicio = st.time_input("HORA INICIO", key="form_hora_inicio")
        with col8:
            hora_termino = st.time_input(
                "HORA TÉRMINO", key="form_hora_termino"
            )

        dt_inicio = datetime.datetime.combine(
            datetime.date.today(), hora_inicio
        )
        dt_termino = datetime.datetime.combine(
            datetime.date.today(), hora_termino
        )

        if dt_termino < dt_inicio:
            dt_termino += datetime.timedelta(days=1)

        minutos_totales = int((dt_termino - dt_inicio).total_seconds() / 60)
        tiempo_calculado = f"{minutos_totales} min"

        with col9:
            tiempo_proceso = st.text_input(
                "TIEMPO (COCCIÓN/BATIDO/ HORNEADO)",
                value=tiempo_calculado,
                disabled=True,
            )

        col10, col11 = st.columns(2)

        with col10:
            temp_equipo = st.number_input(
                "TEMPERATURA  EQUIPO (°C)",
                step=0.5,
                format="%.1f",
                key="form_temp_equipo",
            )
        with col11:
            vel_agitador = st.text_input(
                "VELOCIDAD DEL AGITADOR (hz)/ BATIDORA/OTROS",
                placeholder="Ej. 50 Hz / V2",
                key="form_vel_agitador",
            )

        st.subheader("4. Conformidad y Observaciones")
        col12, col13 = st.columns([1, 2])

        with col12:
            estado_obs = st.radio(
                "ESTADO",
                options=["CONFORME", "NO CONFORME", "OTRO (Texto Libre)"],
                key="form_estado_obs",
            )

        with col13:
            if estado_obs == "OTRO (Texto Libre)":
                obs_detalle = st.text_area(
                    "OBSERVACIÓN",
                    placeholder="Escriba aquí la observación personalizada...",
                    key="form_obs_detalle",
                )
                observacion_final = obs_detalle
            else:
                obs_adicional = st.text_input(
                    "Detalle / Comentario adicional (Opcional)",
                    placeholder="Escriba detalles si aplica...",
                    key="form_obs_adicional",
                )
                observacion_final = (
                    f"{estado_obs} - {obs_adicional}".strip(" -")
                    if obs_adicional
                    else estado_obs
                )

        st.subheader("5. Condiciones del área de trabajo")
        col14, col15 = st.columns([1, 2])

        with col14:
            estado_area = st.radio(
                "ESTADO ÁREA",
                options=["CONFORME", "NO CONFORME"],
                key="form_estado_area",
            )

        with col15:
            obs_area = st.text_input(
                "Detalle / Comentario adicional (Opcional)",
                placeholder="Escriba detalles si aplica...",
                key="form_obs_area",
            )

        condicion_area_final = (
            f"{estado_area} - {obs_area}".strip(" -")
            if obs_area
            else estado_area
        )

        st.markdown("---")
        btn_guardar = st.form_submit_button(
            "💾 Guardar Registro en Google Sheets", use_container_width=True
        )

    if btn_guardar:
        if sheet is not None:
            try:
                nueva_fila = [
                    producto,
                    fecha_p.strftime("%Y-%m-%d"),
                    lote,
                    batch,
                    equipo_final,
                    brix,
                    hora_inicio.strftime("%H:%M"),
                    float(temp_equipo),
                    tiempo_calculado,
                    vel_agitador,
                    hora_termino.strftime("%H:%M"),
                    responsable_final,
                    observacion_final,
                    condicion_area_final,
                ]

                registros_existentes = sheet.get_all_values()
                siguiente_fila = len(registros_existentes) + 1

                rango_insercion = f"A{siguiente_fila}:N{siguiente_fila}"
                sheet.update(
                    range_name=rango_insercion,
                    values=[nueva_fila],
                    value_input_option="USER_ENTERED",
                )

                st.success(
                    f"✅ ¡Registro guardado exitosamente en la fila {siguiente_fila} de Google Sheets!"
                )
                st.cache_data.clear()

                # Marcamos bandera para resetear en el próximo refresco
                st.session_state["necesita_reset"] = True
                st.rerun()

            except Exception as err:
                st.error(f"Error al registrar en la hoja: {err}")
        else:
            st.error("No se pudo conectar a Google Sheets.")

# ---------------------------------------------------------
# PESTAÑA 2: MODIFICAR / ELIMINAR REGISTROS
# ---------------------------------------------------------
with tab2:
    st.subheader("✏️ Editar o Eliminar un Registro Existente")
    if sheet is not None:
        try:
            raw_data = sheet.get_all_values()
            if len(raw_data) > 1:
                headers = raw_data[0]
                rows = raw_data[1:]

                df_edit = pd.DataFrame(rows, columns=headers)
                df_edit["Fila GS"] = list(range(2, len(rows) + 2))

                opciones_registro = [
                    f"Fila {r['Fila GS']} | {r.get('PRODUCTO', '')} | Lote: {r.get('LOTE', '')} | FP: {r.get('F.P', '')}"
                    for _, r in df_edit.iterrows()
                ]

                registro_sel = st.selectbox(
                    "Seleccione el registro que desea modificar o eliminar:",
                    options=opciones_registro,
                )

                idx_seleccionado = opciones_registro.index(registro_sel)
                fila_gs = df_edit.iloc[idx_seleccionado]["Fila GS"]
                datos_fila = df_edit.iloc[idx_seleccionado].to_dict()

                col_mod, col_del = st.columns(2)

                # --- ACCIÓN 1: MODIFICAR ---
                with col_mod:
                    with st.expander(
                        f"📝 Editar datos de la Fila {fila_gs}", expanded=True
                    ):
                        with st.form(f"form_editar_{fila_gs}"):
                            e_prod = st.text_input(
                                "PRODUCTO", value=datos_fila.get("PRODUCTO", "")
                            )
                            e_fp = st.text_input(
                                "F.P", value=datos_fila.get("F.P", "")
                            )
                            e_lote = st.text_input(
                                "LOTE", value=datos_fila.get("LOTE", "")
                            )
                            e_batch = st.text_input(
                                "BATCH", value=datos_fila.get("BATCH", "")
                            )
                            e_equipo = st.text_input(
                                "EQUIPO UTILIZADO",
                                value=datos_fila.get("EQUIPO UTILIZADO", ""),
                            )
                            e_brix = st.text_input(
                                "BRIX (°Bx)",
                                value=datos_fila.get("BRIX (°Bx)", ""),
                            )
                            e_h_ini = st.text_input(
                                "HORA INICIO",
                                value=datos_fila.get("HORA INICIO", ""),
                            )
                            e_temp = st.text_input(
                                "TEMPERATURA EQUIPO (°C)",
                                value=datos_fila.get(
                                    "TEMPERATURA EQUIPO (°C)", ""
                                ),
                            )
                            e_tiempo = st.text_input(
                                "TIEMPO", value=datos_fila.get("TIEMPO", "")
                            )
                            e_vel = st.text_input(
                                "VELOCIDAD AGITADOR",
                                value=datos_fila.get(
                                    "VELOCIDAD AGITADOR/ BATIDORA", ""
                                ),
                            )
                            e_h_term = st.text_input(
                                "HORA TÉRMINO",
                                value=datos_fila.get("HORA TÉRMINO", ""),
                            )
                            e_resp = st.text_input(
                                "RESPONSABLE",
                                value=datos_fila.get("RESPONSABLE", ""),
                            )
                            e_obs = st.text_area(
                                "OBSERVACIÓN",
                                value=datos_fila.get("OBSERVACIÓN", ""),
                            )
                            e_cond_area = st.text_input(
                                "CONDICIONES ÁREA TRABAJO",
                                value=datos_fila.get(
                                    "CONDICIONES DEL AREA DE TRABAJO", ""
                                ),
                            )

                            btn_actualizar = st.form_submit_button(
                                "💾 Actualizar Fila en Google Sheets"
                            )

                        if btn_actualizar:
                            valores_actualizados = [
                                e_prod,
                                e_fp,
                                e_lote,
                                e_batch,
                                e_equipo,
                                e_brix,
                                e_h_ini,
                                e_temp,
                                e_tiempo,
                                e_vel,
                                e_h_term,
                                e_resp,
                                e_obs,
                                e_cond_area,
                            ]
                            sheet.update(
                                range_name=f"A{fila_gs}:N{fila_gs}",
                                values=[valores_actualizados],
                                value_input_option="USER_ENTERED",
                            )
                            st.success(
                                f"✅ ¡Fila {fila_gs} actualizada correctamente!"
                            )
                            st.cache_data.clear()
                            st.rerun()

                # --- ACCIÓN 2: ELIMINAR ---
                with col_del:
                    with st.expander(
                        f"🗑️ Eliminar Fila {fila_gs}", expanded=True
                    ):
                        st.warning(
                            f"⚠️ ¿Estás seguro de que deseas eliminar permanentemente el registro de la fila {fila_gs}?"
                        )
                        btn_eliminar = st.button(
                            f"❌ Sí, Eliminar Fila {fila_gs}",
                            type="primary",
                            key=f"btn_del_{fila_gs}",
                        )

                        if btn_eliminar:
                            sheet.delete_rows(int(fila_gs))
                            st.success(
                                f"🗑️ ¡Fila {fila_gs} eliminada correctamente!"
                            )
                            st.cache_data.clear()
                            st.rerun()

            else:
                st.info("No hay registros suficientes para modificar.")
        except Exception as e:
            st.error(f"Error al cargar registros para edición: {e}")


# --- GENERACIÓN DE EXCEL ---
def generar_excel_formateado(df_datos):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Control Procesos"

    thin_border = Border(
        left=Side(style="thin", color="000000"),
        right=Side(style="thin", color="000000"),
        top=Side(style="thin", color="000000"),
        bottom=Side(style="thin", color="000000"),
    )
    fill_header = PatternFill(
        start_color="F2F2F2", end_color="F2F2F2", fill_type="solid"
    )
    font_bold = Font(name="Calibri", size=11, bold=True)
    font_title = Font(name="Calibri", size=13, bold=True)

    ws.merge_cells("A1:B3")
    ws.merge_cells("C1:L3")
    ws.merge_cells("M1:N3")

    ws["C1"] = "FORMATO CONTROL DE PROCESO COCINA DULCE"
    ws["C1"].font = font_title
    ws["C1"].alignment = Alignment(horizontal="center", vertical="center")

    ws["M1"] = "MA-FR- 109"
    ws["M1"].font = font_bold
    ws["M1"].alignment = Alignment(horizontal="center", vertical="center")

    logo_path = None
    for posible_logo in [
        "logo.jfif",
        "logo.png",
        "LOGOMA.jfif",
        "logo.jpg",
        "logo.jpeg",
        "LOGOMA.png",
    ]:
        if os.path.exists(posible_logo):
            logo_path = posible_logo
            break

    if logo_path:
        try:
            img = PILImage.open(logo_path).convert("RGB")
            img.thumbnail((110, 50))
            img_temp_path = "temp_logo.png"
            img.save(img_temp_path, "PNG")

            excel_img = OpenPyxlImage(img_temp_path)
            excel_img.anchor = "A1"
            ws.add_image(excel_img)
        except Exception:
            ws["A1"] = "MA"
            ws["A1"].alignment = Alignment(
                horizontal="center", vertical="center"
            )
    else:
        ws["A1"] = "MA"
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

    for row in ws["A1:N3"]:
        for cell in row:
            cell.border = thin_border

    headers = list(df_datos.columns)
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=5, column=col_idx, value=header)
        cell.font = font_bold
        cell.fill = fill_header
        cell.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )
        cell.border = thin_border

    for row_idx, row_data in enumerate(df_datos.values, start=6):
        for col_idx, value in enumerate(row_data, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border

    for col in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


st.markdown("---")
st.subheader("📊 Historial de Registros (Google Sheets)")

if sheet is not None:
    try:
        data = sheet.get_all_records()
        if data:
            df_registros = pd.DataFrame(data)
            st.dataframe(df_registros, use_container_width=True)

            excel_data = generar_excel_formateado(df_registros)

            st.download_button(
                label="📥 Descargar Historial Formateado (Excel MA-FR-109)",
                data=excel_data,
                file_name=f"Control_Cocina_Dulce_MA-FR-109_{datetime.date.today()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        else:
            st.info("La hoja de Google Sheets está vacía por el momento.")
    except Exception as e:
        st.warning(f"No se pudieron cargar los registros: {e}")
