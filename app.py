import io
from typing import Optional, List, Dict

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Búsqueda de empleado en Excel", page_icon="🔎", layout="wide")


EXPECTED_HEADERS = [
    "Número de empleado",
    "Nombre",
    "Mes",
    "Periodo de nómina",
    "UUID Vigente",
    "/559 Transferencia",
]


def normalize_text(value) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if text.endswith(".0"):
        try:
            float_val = float(text)
            if float_val.is_integer():
                return str(int(float_val))
        except Exception:
            pass
    return text


def get_excel_engine(filename: str) -> Optional[str]:
    name = filename.lower()
    if name.endswith(".xlsx") or name.endswith(".xlsm"):
        return "openpyxl"
    if name.endswith(".xls"):
        return "xlrd"
    return None


def find_column(headers: List[str], exact: Optional[str] = None, contains_all: Optional[List[str]] = None) -> Optional[int]:
    for idx, header in enumerate(headers):
        current = normalize_text(header)
        if exact and current == exact:
            return idx
        if contains_all and all(token.lower() in current.lower() for token in contains_all):
            return idx
    return None


@st.cache_data(show_spinner=False)
def search_employee_in_workbook(file_bytes: bytes, file_name: str, employee_number: str) -> pd.DataFrame:
    engine = get_excel_engine(file_name)
    if engine is None:
        raise ValueError("Formato de archivo no soportado. Usa .xls, .xlsx o .xlsm")

    employee_number = normalize_text(employee_number)
    excel_file = pd.ExcelFile(io.BytesIO(file_bytes), engine=engine)
    results: List[Dict[str, str]] = []

    for sheet_name in excel_file.sheet_names:
        try:
            df = pd.read_excel(
                io.BytesIO(file_bytes),
                sheet_name=sheet_name,
                engine=engine,
                dtype=object,
            )
        except Exception as exc:
            results.append(
                {
                    "Hoja": sheet_name,
                    "Número de empleado": "",
                    "Nombre": "",
                    "Mes": "",
                    "Periodo de nómina": "",
                    "UUID Vigente": "",
                    "/559 Transferencia": "",
                    "Estado": f"No se pudo leer la hoja: {exc}",
                }
            )
            continue

        if df.empty:
            continue

        headers = [normalize_text(col) for col in df.columns.tolist()]

        idx_num_empleado = 0
        idx_nombre = find_column(headers, exact="Nombre")
        idx_mes = find_column(headers, exact="Mes Acumulación")
        idx_periodo = find_column(headers, exact="Periodo Nómina")
        idx_uuid = find_column(headers, exact="UUID Vigente")
        idx_transfer = find_column(headers, contains_all=["559", "Transferencia"])

        if idx_nombre is None or idx_mes is None or idx_periodo is None or idx_uuid is None or idx_transfer is None:
            continue

        first_col_series = df.iloc[:, idx_num_empleado].map(normalize_text)
        matches = df[first_col_series == employee_number]

        if matches.empty:
            continue

        for _, row in matches.iterrows():
            results.append(
                {
                    "Hoja": sheet_name,
                    "Número de empleado": normalize_text(row.iloc[idx_num_empleado]),
                    "Nombre": normalize_text(row.iloc[idx_nombre]),
                    "Mes": normalize_text(row.iloc[idx_mes]),
                    "Periodo de nómina": normalize_text(row.iloc[idx_periodo]),
                    "UUID Vigente": normalize_text(row.iloc[idx_uuid]),
                    "/559 Transferencia": normalize_text(row.iloc[idx_transfer]),
                    "Estado": "Encontrado",
                }
            )

    if not results:
        return pd.DataFrame(columns=["Hoja", *EXPECTED_HEADERS, "Estado"])

    return pd.DataFrame(results)


st.title("🔎 Búsqueda de empleado en archivo Excel")
st.write(
    "Sube tu archivo de nómina y escribe el número de empleado. "
    "La app buscará en todas las hojas y devolverá los datos clave."
)

with st.sidebar:
    st.header("Parámetros")
    uploaded_file = st.file_uploader(
        "Sube el archivo Excel",
        type=["xls", "xlsx", "xlsm"],
        help="La app revisa todas las hojas del archivo.",
    )
    employee_number = st.text_input("Número de empleado", placeholder="Ej. 10001175")
    search_clicked = st.button("Buscar", type="primary", use_container_width=True)

st.markdown(
    """
    **La app devuelve estas columnas:**
    - Número de empleado
    - Nombre
    - Mes
    - Periodo de nómina
    - UUID Vigente
    - /559 Transferencia
    - Hoja
    """
)

if search_clicked:
    if uploaded_file is None:
        st.error("Primero sube un archivo Excel.")
    elif not employee_number.strip():
        st.error("Escribe un número de empleado.")
    else:
        with st.spinner("Buscando en todas las hojas..."):
            file_bytes = uploaded_file.getvalue()
            result_df = search_employee_in_workbook(file_bytes, uploaded_file.name, employee_number)

        if result_df.empty:
            st.warning("No se encontraron coincidencias para ese número de empleado.")
        else:
            st.success(f"Se encontraron {len(result_df)} coincidencia(s).")
            show_df = result_df[[
                "Hoja",
                "Número de empleado",
                "Nombre",
                "Mes",
                "Periodo de nómina",
                "UUID Vigente",
                "/559 Transferencia",
            ]]
            st.dataframe(show_df, use_container_width=True, hide_index=True)

            csv_data = show_df.to_csv(index=False).encode("utf-8-sig")
            st.download_button(
                "Descargar resultados en CSV",
                data=csv_data,
                file_name=f"resultado_empleado_{normalize_text(employee_number)}.csv",
                mime="text/csv",
            )

with st.expander("Notas técnicas"):
    st.write(
        "La búsqueda se hace en la primera columna de cada hoja. "
        "La columna de /559 Transferencia se identifica por encabezado, así que aunque cambie de posición, la app la encuentra."
    )
