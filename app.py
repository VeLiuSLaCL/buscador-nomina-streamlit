import io
from typing import Optional, List

import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Búsqueda de empleado en múltiples archivos Excel",
    page_icon="🔎",
    layout="wide",
)

EXPECTED_HEADERS = [
    "Archivo",
    "Hoja",
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
    if text.lower() == "nan":
        return ""
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


def find_column(
    headers: List[str],
    exact: Optional[str] = None,
    contains_all: Optional[List[str]] = None,
) -> Optional[int]:
    for idx, header in enumerate(headers):
        current = normalize_text(header)

        if exact and current.lower() == exact.lower():
            return idx

        if contains_all and all(token.lower() in current.lower() for token in contains_all):
            return idx

    return None


@st.cache_data(show_spinner=False)
def build_search_index(file_bytes: bytes, file_name: str) -> pd.DataFrame:
    engine = get_excel_engine(file_name)
    if engine is None:
        raise ValueError(f"Formato de archivo no soportado: {file_name}")

    results = []

    with pd.ExcelFile(io.BytesIO(file_bytes), engine=engine) as excel_file:
        for sheet_name in excel_file.sheet_names:
            try:
                headers_df = excel_file.parse(sheet_name=sheet_name, nrows=0)
                headers = [normalize_text(col) for col in headers_df.columns.tolist()]

                if not headers:
                    continue

                idx_num_empleado = 0
                idx_nombre = find_column(headers, exact="Nombre")

                idx_mes = find_column(headers, exact="Mes Acumulación")
                if idx_mes is None:
                    idx_mes = find_column(headers, exact="Mes")

                idx_periodo = find_column(headers, exact="Periodo Nómina")
                if idx_periodo is None:
                    idx_periodo = find_column(headers, exact="Periodo de nómina")

                idx_uuid = find_column(headers, exact="UUID Vigente")
                idx_transfer = find_column(headers, contains_all=["559", "Transferencia"])

                if (
                    idx_nombre is None
                    or idx_mes is None
                    or idx_periodo is None
                    or idx_uuid is None
                    or idx_transfer is None
                ):
                    continue

                needed_cols = sorted(
                    set(
                        [
                            idx_num_empleado,
                            idx_nombre,
                            idx_mes,
                            idx_periodo,
                            idx_uuid,
                            idx_transfer,
                        ]
                    )
                )

                df = excel_file.parse(
                    sheet_name=sheet_name,
                    usecols=needed_cols,
                    dtype=object,
                )

                if df.empty:
                    continue

                rel_num = needed_cols.index(idx_num_empleado)
                rel_nombre = needed_cols.index(idx_nombre)
                rel_mes = needed_cols.index(idx_mes)
                rel_periodo = needed_cols.index(idx_periodo)
                rel_uuid = needed_cols.index(idx_uuid)
                rel_transfer = needed_cols.index(idx_transfer)

                selected = pd.DataFrame(
                    {
                        "Archivo": file_name,
                        "Hoja": sheet_name,
                        "Número de empleado": df.iloc[:, rel_num].map(normalize_text),
                        "Nombre": df.iloc[:, rel_nombre].map(normalize_text),
                        "Mes": df.iloc[:, rel_mes].map(normalize_text),
                        "Periodo de nómina": df.iloc[:, rel_periodo].map(normalize_text),
                        "UUID Vigente": df.iloc[:, rel_uuid].map(normalize_text),
                        "/559 Transferencia": df.iloc[:, rel_transfer].map(normalize_text),
                    }
                )

                selected = selected[selected["Número de empleado"] != ""].copy()

                if not selected.empty:
                    results.append(selected)

            except Exception:
                continue

    if not results:
        return pd.DataFrame(columns=EXPECTED_HEADERS)

    final_df = pd.concat(results, ignore_index=True)
    final_df["Número de empleado"] = final_df["Número de empleado"].map(normalize_text)
    return final_df


def search_employee(index_df: pd.DataFrame, employee_number: str) -> pd.DataFrame:
    employee_number = normalize_text(employee_number)
    result = index_df[index_df["Número de empleado"] == employee_number].copy()
    return result


st.title("🔎 Búsqueda de empleado en múltiples archivos Excel")
st.write(
    "Sube uno o varios archivos Excel y escribe el número de empleado para buscarlo en todas las hojas."
)

with st.sidebar:
    st.header("Parámetros")
    uploaded_files = st.file_uploader(
        "Sube uno o varios archivos Excel",
        type=["xls", "xlsx", "xlsm"],
        accept_multiple_files=True,
        help="La app revisa todas las hojas de todos los archivos cargados.",
    )
    employee_number = st.text_input(
        "Número de empleado",
        placeholder="Ej. 10001175",
    )
    search_clicked = st.button(
        "Buscar",
        type="primary",
        use_container_width=True,
    )

if search_clicked:
    if not uploaded_files:
        st.error("Primero sube al menos un archivo Excel.")
    elif not employee_number.strip():
        st.error("Escribe un número de empleado.")
    else:
        all_results = []

        with st.spinner("Procesando archivos..."):
            for uploaded_file in uploaded_files:
                try:
                    file_bytes = uploaded_file.getvalue()
                    index_df = build_search_index(file_bytes, uploaded_file.name)
                    result_df = search_employee(index_df, employee_number)

                    if not result_df.empty:
                        all_results.append(result_df)

                except Exception as e:
                    st.warning(f"No se pudo procesar el archivo {uploaded_file.name}: {e}")

        if not all_results:
            st.warning("No se encontraron coincidencias para ese número de empleado.")
        else:
            final_result = pd.concat(all_results, ignore_index=True)

            final_result["Origen"] = (
                final_result["Archivo"].astype(str) + " | " + final_result["Hoja"].astype(str)
            )

            final_result = final_result[
                [
                    "Número de empleado",
                    "Nombre",
                    "Mes",
                    "Periodo de nómina",
                    "UUID Vigente",
                    "/559 Transferencia",
                    "Hoja",
                    "Archivo",
                    "Origen",
                ]
            ]

            st.success(f"Se encontraron {len(final_result)} coincidencia(s).")
            st.dataframe(final_result, use_container_width=True, hide_index=True)

            csv_data = final_result.to_csv(index=False).encode("utf-8-sig")
            st.download_button(
                "Descargar resultados en CSV",
                data=csv_data,
                file_name=f"resultado_empleado_{normalize_text(employee_number)}.csv",
                mime="text/csv",
            )
