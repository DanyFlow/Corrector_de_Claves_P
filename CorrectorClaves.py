import base64
from pathlib import Path

import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Control de Asistencia",
    page_icon="📋",
    layout="wide",
)

SLOTS = ("Q1", "Q2", "ACTIVOS")
KEY_HEADER = "CLAVE DE PAGO"


def read_first_sheet(uploaded_file):
    """Lee la primera hoja del archivo cargado, sin depender de su nombre."""
    suffix = Path(uploaded_file.name).suffix.lower()
    if suffix == ".xls":
        engine = "xlrd"
    elif suffix == ".xlsb":
        engine = "pyxlsb"
    else:
        engine = "openpyxl"

    uploaded_file.seek(0)
    df = pd.read_excel(
        uploaded_file,
        sheet_name=0,
        header=0,
        dtype=object,
        engine=engine,
    )
    # Normalizar encabezados sin cambiar los valores de las filas.
    df.columns = [str(c).strip() if not pd.isna(c) else "" for c in df.columns]
    return df


def is_blank(value):
    if value is None:
        return True
    try:
        if pd.isna(value):
            return True
    except (TypeError, ValueError):
        pass
    return str(value).strip() == ""


def as_text(value):
    if is_blank(value):
        return ""
    # Excel suele importar identificadores numéricos como enteros o floats.
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def find_key_column(df):
    for idx, name in enumerate(df.columns):
        if str(name).strip().casefold() == KEY_HEADER.casefold():
            return idx
    return None


def transform_key(value):
    """Reproduce la transformación VBA:

    R0142951103 M020... -> 295001R0141103 M020...
    Solo cambia el primer bloque y conserva el espacio/resto tal cual.
    """
    if is_blank(value):
        return value

    original = as_text(value)
    if len(original) < 8 or not original.startswith("R014"):
        return value

    space_pos = original.find(" ")
    if space_pos >= 0:
        first_block = original[:space_pos]
        remainder = original[space_pos:]
    else:
        first_block = original
        remainder = ""

    if not first_block.startswith("R014") or len(first_block) < 7:
        return value

    remainder_of_block = first_block[4:]
    transformed_block = (
        remainder_of_block[:3] + "001R014" + remainder_of_block[3:]
    )
    return transformed_block + remainder


def transform_sheet(df, sheet_name):
    col_idx = find_key_column(df)
    if col_idx is None:
        return df.copy(), 0, (
            f"No se encontró el encabezado '{KEY_HEADER}' en la fila 1 de {sheet_name}."
        )

    result = df.copy()
    col_name = result.columns[col_idx]
    modified = 0
    for idx, value in result[col_name].items():
        if is_blank(value):
            continue
        text_value = as_text(value)
        if len(text_value) >= 8 and text_value.startswith("R014"):
            new_value = transform_key(value)
            if as_text(new_value) != text_value:
                result.at[idx, col_name] = new_value
                modified += 1
    return result, modified, None


def ensure_column_count(df, count):
    result = df.copy()
    while result.shape[1] < count:
        result[f"__col_{result.shape[1] + 1}"] = pd.NA
    return result


def find_employee_key(q_df, employee_id, key_col_idx):
    """Busca en columna A; devuelve (encontrado, clave), respetando primera coincidencia."""
    if q_df.empty or q_df.shape[1] == 0:
        return False, None

    for _, row in q_df.iterrows():
        candidate = as_text(row.iloc[0]).strip()
        if candidate.casefold() == employee_id.casefold():
            if key_col_idx < len(row):
                return True, row.iloc[key_col_idx]
            return True, None
    return False, None


def build_employee_lookup(q_df, key_col_idx):
    """Crea un índice por tarjeta/empleado, conservando la primera coincidencia."""
    lookup = {}
    if q_df.empty or q_df.shape[1] == 0:
        return lookup

    # Un diccionario evita volver a recorrer Q1/Q2 por cada fila de ACTIVOS.
    for employee_id, key in zip(q_df.iloc[:, 0], q_df.iloc[:, key_col_idx]):
        normalized_id = as_text(employee_id).strip().casefold()
        if normalized_id and normalized_id not in lookup:
            lookup[normalized_id] = key
    return lookup


def run_lookup(q1, q2, activos):
    q1_key_idx = find_key_column(q1)
    q2_key_idx = find_key_column(q2)
    if q1_key_idx is None or q2_key_idx is None:
        missing = []
        if q1_key_idx is None:
            missing.append("Q1")
        if q2_key_idx is None:
            missing.append("Q2")
        raise ValueError(
            "No se encontró 'CLAVE DE PAGO' en la primera fila de: "
            + ", ".join(missing)
            + ". Ejecuta primero la transformación o revisa los encabezados."
        )

    # Prioridad Q1 sobre Q2, igual que en la versión anterior.
    q1_lookup = build_employee_lookup(q1, q1_key_idx)
    q2_lookup = build_employee_lookup(q2, q2_key_idx)

    result = ensure_column_count(activos, 10)
    # La columna J equivale al índice 9.
    j_name = result.columns[9]
    found_count = 0
    not_found_count = 0
    skipped_count = 0
    # Mantener J sin cambios cuando B está vacía, como hacía la versión original.
    output_keys = result[j_name].tolist()

    employee_ids = result.iloc[:, 1] if result.shape[1] > 1 else [None] * len(result)
    for position, raw_id in enumerate(employee_ids):
        if is_blank(raw_id):
            skipped_count += 1
            continue

        employee_id = as_text(raw_id).strip().casefold()
        if employee_id in q1_lookup:
            output_keys[position] = q1_lookup[employee_id]
            found_count += 1
        elif employee_id in q2_lookup:
            output_keys[position] = q2_lookup[employee_id]
            found_count += 1
        else:
            # Evita exportar un código viejo de una ejecución anterior.
            output_keys[position] = pd.NA
            not_found_count += 1

    result[j_name] = output_keys
    return result, found_count, not_found_count, skipped_count


def generate_codes_csv(activos):
    if activos.shape[1] < 10:
        raise ValueError("La hoja ACTIVOS no tiene suficientes columnas para leer A, G y J.")

    rows = []
    seen = set()
    for _, row in activos.iterrows():
        employee_card = row.iloc[0]
        correct_code = row.iloc[9]
        if is_blank(correct_code):
            continue

        # Dedupe de tarjeta de empleado: conservar la primera fila con J no vacía,
        # tal como sucede en el VBA (primero filtra J y luego elimina duplicados).
        card_key = as_text(employee_card)
        if card_key in seen:
            continue
        seen.add(card_key)

        incorrect_code = row.iloc[6] if len(row) > 6 else None
        rows.append({
            "TARJETAEMPLEADO": employee_card,
            "CODIGO_INCORRECTO": incorrect_code,
            "CODIGO_CORRECTO": correct_code,
        })

    output = pd.DataFrame(
        rows,
        columns=["TARJETAEMPLEADO", "CODIGO_INCORRECTO", "CODIGO_CORRECTO"],
    )
    # UTF-8 con BOM, comas y encabezados para facilitar la apertura en Excel.
    csv_bytes = output.to_csv(index=False, sep=",", encoding="utf-8-sig").encode("utf-8-sig")
    return output, csv_bytes


def load_uploads():
    loaded = {}
    for slot in SLOTS:
        file = st.session_state.get(f"upload_{slot}")
        if file is None:
            raise ValueError(f"Falta cargar el archivo correspondiente a {slot}.")
        loaded[slot] = read_first_sheet(file)
    st.session_state["data"] = loaded
    st.session_state["lookup_done"] = False
    st.session_state["csv_bytes"] = None
    st.session_state["csv_preview"] = None
    st.session_state["messages"] = [f"{slot}: {len(df):,} filas cargadas." for slot, df in loaded.items()]


def transform_all(show_messages=True):
    data = st.session_state["data"]
    messages = []
    for slot in ("Q1", "Q2"):
        transformed, count, warning = transform_sheet(data[slot], slot)
        data[slot] = transformed
        if warning:
            messages.append(("warning", warning))
        else:
            messages.append(("success", f"{slot}: {count:,} claves transformadas."))
    st.session_state["data"] = data
    if show_messages:
        for level, message in messages:
            getattr(st, level)(message)
    return messages


def lookup_all(show_messages=True):
    data = st.session_state["data"]
    result, found, not_found, skipped = run_lookup(
        data["Q1"], data["Q2"], data["ACTIVOS"]
    )
    data["ACTIVOS"] = result
    st.session_state["data"] = data
    st.session_state["lookup_done"] = True
    if show_messages:
        c1, c2, c3 = st.columns(3)
        c1.metric("Coincidencias", f"{found:,}")
        c2.metric("Sin coincidencia", f"{not_found:,}")
        c3.metric("Sin tarjeta en B", f"{skipped:,}")
    return found, not_found, skipped


def make_csv(show_messages=True):
    data = st.session_state["data"]
    preview, csv_bytes = generate_codes_csv(data["ACTIVOS"])
    st.session_state["csv_preview"] = preview
    st.session_state["csv_bytes"] = csv_bytes
    if show_messages:
        st.success(f"CSV generado: {len(preview):,} filas.")
    return preview, csv_bytes


st.title("Control de Asistencia")
st.caption("Procesamiento de Q1, Q2 y ACTIVOS · Aplicación local con Streamlit")
st.markdown("[📘 Abrir guía de uso](https://frolicking-crepe-61b29a.netlify.app/)")

with st.expander("Cómo usar la aplicación", expanded=True):
    st.markdown(
        """
        1. Carga los tres archivos Excel en sus espacios correspondientes. Los archivos pueden tener cualquier nombre.
        2. Ejecuta los módulos por separado o usa **Ejecutar todo**.
        3. Descarga el CSV generado al finalizar.

        Se lee únicamente la primera hoja de cada archivo. No se modifica ni sobrescribe ningún archivo original.
        """
    )

st.subheader("1. Cargar los archivos")
col1, col2, col3 = st.columns(3)
with col1:
    st.file_uploader(
        "Archivo que corresponde a Q1",
        type=["xlsx", "xlsm", "xls", "xlsb"],
        key="upload_Q1",
        help="El nombre del archivo puede ser cualquiera.",
    )
with col2:
    st.file_uploader(
        "Archivo que corresponde a Q2",
        type=["xlsx", "xlsm", "xls", "xlsb"],
        key="upload_Q2",
        help="El nombre del archivo puede ser cualquiera.",
    )
with col3:
    st.file_uploader(
        "Archivo que corresponde a ACTIVOS",
        type=["xlsx", "xlsm", "xls", "xlsb"],
        key="upload_ACTIVOS",
        help="El nombre del archivo puede ser cualquiera.",
    )

if st.button("Cargar los tres archivos", type="primary", use_container_width=True):
    try:
        load_uploads()
        st.success("Archivos cargados correctamente.")
        for message in st.session_state.get("messages", []):
            st.write("• " + message)
    except Exception as exc:
        st.error(f"No se pudieron cargar los archivos: {exc}")

if "data" in st.session_state:
    data = st.session_state["data"]
    st.success("Hay datos cargados en esta sesión.")
    metric_cols = st.columns(3)
    for i, slot in enumerate(SLOTS):
        metric_cols[i].metric(f"Filas en {slot}", f"{len(data[slot]):,}")

    st.divider()
    st.subheader("2. Ejecutar operaciones")
    left, right = st.columns(2)

    with left:
        if st.button("Módulo 2 · Transformar claves", use_container_width=True):
            messages = transform_all(show_messages=False)
            for level, message in messages:
                getattr(st, level)(message)

        if st.button("Módulo 3 · Buscar claves", use_container_width=True):
            try:
                lookup_all(show_messages=True)
            except Exception as exc:
                st.error(str(exc))

    with right:
        if st.button("Módulo 4 · Generar CSV", use_container_width=True):
            try:
                preview, _ = make_csv()
                st.dataframe(preview, use_container_width=True)
            except Exception as exc:
                st.error(str(exc))

        if st.button("Ejecutar todo", type="primary", use_container_width=True):
            try:
                with st.spinner("Procesando archivos..."):
                    transform_messages = transform_all(show_messages=False)
                    for level, message in transform_messages:
                        getattr(st, level)(message)
                    found, not_found, skipped = lookup_all(show_messages=False)
                    preview, _ = make_csv(show_messages=False)
                st.success("Proceso completo.")
                a, b, c, d = st.columns(4)
                a.metric("Coincidencias", f"{found:,}")
                b.metric("Sin coincidencia", f"{not_found:,}")
                c.metric("Sin tarjeta en B", f"{skipped:,}")
                d.metric("Filas en CSV", f"{len(preview):,}")
            except Exception as exc:
                st.error(f"No se pudo completar el proceso: {exc}")

    if st.session_state.get("lookup_done"):
        st.subheader("Vista previa de ACTIVOS")
        st.caption("La columna J contiene la clave encontrada.")
        st.dataframe(data["ACTIVOS"], use_container_width=True)

    if st.session_state.get("csv_bytes") is not None:
        st.subheader("3. Descargar resultado")
        st.download_button(
            label="Descargar CODIGOS.csv",
            data=st.session_state["csv_bytes"],
            file_name="CODIGOS.csv",
            mime="text/csv",
            type="primary",
            use_container_width=True,
        )
        preview = st.session_state.get("csv_preview")
        if preview is not None:
            st.caption(f"El archivo contiene {len(preview):,} filas de datos, más el encabezado.")
            st.dataframe(preview, use_container_width=True)

    with st.expander("Diagnóstico de encabezados"):
        for slot in ("Q1", "Q2"):
            key_col = find_key_column(data[slot])
            if key_col is None:
                st.warning(f"{slot}: no se encontró el encabezado exacto '{KEY_HEADER}'.")
            else:
                st.write(f"{slot}: '{KEY_HEADER}' se encuentra en la columna {key_col + 1}.")
else:
    st.info("Carga los tres archivos para habilitar las operaciones.")

# ------- Firma flotante centrada
url_imagen_drive = "https://lh3.googleusercontent.com/d/1uqXeCGgBghAjmqj3pXqK-HdCrRybQiQL"
enlace_perfil = "https://share.streamlit.io/user/danyflowt"

firma_html = f"""
<style>
.firma-flotante-link {{
    position: fixed;
    bottom: 20px;
    left: 50%;
    transform: translateX(-50%);
    z-index: 999999;
    text-decoration: none;
}}

.firma-flotante {{
    display: flex;
    align-items: center;
    background: #06b6d4;
    color: #ffffff;
    width: 44px;
    height: 44px;
    padding: 6px;
    border-radius: 50px;
    box-shadow: 0 4px 12px rgba(0, 0, 0, 0.25);
    border: 1px solid #0891b2;
    font-family: var(--font, sans-serif);
    font-size: 13px;
    overflow: hidden;
    transition: width 0.4s cubic-bezier(0.4, 0, 0.2, 1), box-shadow 0.3s ease, background 0.3s ease, padding 0.3s ease;
    cursor: pointer;
    white-space: nowrap;
}}

.firma-flotante:hover {{
    width: max-content; 
    background: #0891b2; 
    box-shadow: 0 6px 18px rgba(0, 0, 0, 0.35);
    padding-right: 16px; 
}}

.firma-flotante img {{
    width: 30px;
    height: 30px;
    border-radius: 50%;
    object-fit: cover;
    flex-shrink: 0;
    border: 1.5px solid rgba(255, 255, 255, 0.9);
}}

.firma-texto {{
    margin-left: 12px;
    opacity: 0;
    transition: opacity 0.25s ease 0.1s;
    display: flex;
    flex-direction: column;
    justify-content: center;
    text-align: left;
    line-height: 1.3;
}}

.firma-flotante:hover .firma-texto {{
    opacity: 1;
}}

.firma-titulo {{
    font-weight: 600;
    font-size: 12px;
    color: #ffffff;
}}

.firma-sub {{
    font-size: 10px;
    color: #ecfeff;
}}
</style>

<a href="{enlace_perfil}" target="_blank" class="firma-flotante-link" title="Visitar perfil de Streamlit">
    <div class="firma-flotante">
        <img src="{url_imagen_drive}" alt="Dany FloVel">
        <div class="firma-texto">
            <span class="firma-titulo">Dany FloVel</span>
            <span class="firma-sub">⚡ Ver perfil ⚡</span>
        </div>
    </div>
</a>
"""

st.markdown(firma_html, unsafe_allow_html=True)
