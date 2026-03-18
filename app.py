"""KF-PDFLiberator — Extract tables from PDFs into Excel/CSV."""

import io
import streamlit as st

st.set_page_config(
    page_title="KF-PDFLiberator",
    page_icon="📄",
    layout="wide",
)

import pdfplumber
import openpyxl
import pandas as pd
from components.header import render_header
from components.footer import render_footer
from components.i18n import t


def extract_tables(pdf_bytes: bytes, method: str) -> list[dict]:
    """Extract tables from PDF bytes.

    Returns a list of dicts with keys: page, table_index, dataframe.
    """
    results = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            settings = {}
            if method == "lattice":
                settings = {"vertical_strategy": "lines", "horizontal_strategy": "lines"}
            elif method == "stream":
                settings = {"vertical_strategy": "text", "horizontal_strategy": "text"}

            tables = page.extract_tables(table_settings=settings)
            for table_idx, table in enumerate(tables):
                if not table or len(table) < 2:
                    continue
                header = table[0]
                rows = table[1:]
                df = pd.DataFrame(rows, columns=header)
                results.append({
                    "page": page_num,
                    "table_index": table_idx + 1,
                    "dataframe": df,
                })
    return results


def df_to_excel_bytes(dfs: list[pd.DataFrame], sheet_names: list[str]) -> bytes:
    """Convert multiple DataFrames to a single Excel file with multiple sheets."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for df, name in zip(dfs, sheet_names):
            df.to_excel(writer, sheet_name=name[:31], index=False)
    return output.getvalue()


# --- Header ---
render_header()
st.info("💻 " + t("desktop_recommended"))

# --- Sidebar: extraction method ---
st.sidebar.markdown(f"### {t('extraction_method')}")
method_options = {
    t("method_auto"): "auto",
    t("method_lattice"): "lattice",
    t("method_stream"): "stream",
}
selected_label = st.sidebar.radio(
    t("extraction_method"),
    options=list(method_options.keys()),
    index=0,
    label_visibility="collapsed",
)
method = method_options[selected_label]

# --- Main Content ---
uploaded_file = st.file_uploader(t("upload_prompt"), type=["pdf"])

if uploaded_file is not None:
    pdf_bytes = uploaded_file.read()

    with st.spinner(t("processing")):
        tables = extract_tables(pdf_bytes, method)

    if not tables:
        st.warning(t("no_tables"))
    else:
        st.success(f"{len(tables)} {t('tables_found')}")

        # Download all as single Excel
        all_dfs = [tbl["dataframe"] for tbl in tables]
        all_names = [f"P{tbl['page']}_T{tbl['table_index']}" for tbl in tables]
        excel_bytes = df_to_excel_bytes(all_dfs, all_names)

        st.download_button(
            label=f"📥 {t('download_all_excel')}",
            data=excel_bytes,
            file_name=f"{uploaded_file.name.rsplit('.', 1)[0]}_tables.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        # Show each table with individual downloads
        for tbl in tables:
            page_num = tbl["page"]
            table_idx = tbl["table_index"]
            df = tbl["dataframe"]

            st.markdown(f"#### {t('page')} {page_num} — {t('table')} {table_idx}")

            # Preview
            st.dataframe(df, use_container_width=True)

            # Individual downloads
            col1, col2 = st.columns(2)
            with col1:
                single_excel = df_to_excel_bytes([df], [f"P{page_num}_T{table_idx}"])
                st.download_button(
                    label=f"📥 {t('download_excel')}",
                    data=single_excel,
                    file_name=f"P{page_num}_T{table_idx}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"xlsx_{page_num}_{table_idx}",
                )
            with col2:
                csv_bytes = df.to_csv(index=False).encode("utf-8-sig")
                st.download_button(
                    label=f"📥 {t('download_csv')}",
                    data=csv_bytes,
                    file_name=f"P{page_num}_T{table_idx}.csv",
                    mime="text/csv",
                    key=f"csv_{page_num}_{table_idx}",
                )
else:
    st.info(t("no_file"))

# --- Footer ---
render_footer(libraries=[
    "pdfplumber — PDF table extraction with cell-level coordinate detection",
    "openpyxl — Excel file generation",
    "pandas — Data manipulation and CSV export",
])
