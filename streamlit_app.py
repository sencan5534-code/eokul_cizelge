import streamlit as st
import tempfile
from excel_compressor import process_excel

st.title("E-Okul Excel Oluşturucu")

uploaded_file = st.file_uploader(
    "Excel dosyasını yükle",
    type=["xlsx"]
)

if uploaded_file is not None:
    if st.button("Excel Oluştur"):

        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_input:
            temp_input.write(uploaded_file.read())
            input_path = temp_input.name

        output_path = "E_okul_template_doldurulmus.xlsx"

        process_excel(input_path, output_path)

        with open(output_path, "rb") as file:
            st.download_button(
                label="Çıktı Excel'i İndir",
                data=file,
                file_name="E_okul_template_doldurulmus.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
