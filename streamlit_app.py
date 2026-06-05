import streamlit as st
import tempfile
import os
import time
from excel_compressor import process_excel


st.set_page_config(
    page_title="E-Okul Çizelge Oluşturucu",
    page_icon="📘",
    layout="centered"
)


# =========================================
# SESSION STATE
# =========================================

if "output_file_path" not in st.session_state:
    st.session_state.output_file_path = None

if "output_file_name" not in st.session_state:
    st.session_state.output_file_name = None

if "uploaded_file_name" not in st.session_state:
    st.session_state.uploaded_file_name = None

if "processed" not in st.session_state:
    st.session_state.processed = False

if "downloaded" not in st.session_state:
    st.session_state.downloaded = False

if "process_time" not in st.session_state:
    st.session_state.process_time = None


def reset_app():
    st.session_state.output_file_path = None
    st.session_state.output_file_name = None
    st.session_state.uploaded_file_name = None
    st.session_state.processed = False
    st.session_state.downloaded = False
    st.session_state.process_time = None


def format_file_size(size_bytes):
    if size_bytes < 1024:
        return f"{size_bytes} B"
    elif size_bytes < 1024 * 1024:
        return f"{round(size_bytes / 1024, 2)} KB"
    else:
        return f"{round(size_bytes / (1024 * 1024), 2)} MB"


# =========================================
# HEADER
# =========================================

st.title("📘 E-Okul Performans Çizelge Oluşturucu")

st.markdown(
    """
    Bu uygulama, yüklediğiniz E-Okul Excel dosyasındaki sınıf sayfalarını işler.  
    Her sınıf için **1. Perf** ve **2. Perf** çizelgelerini otomatik oluşturur.
    """
)

st.divider()


# =========================================
# INFO BOX
# =========================================

with st.expander("ℹ️ Kullanım bilgisi", expanded=False):
    st.markdown(
        """
        **Nasıl kullanılır?**

        1. `.xlsx` uzantılı E-Okul Excel dosyasını yükleyin.
        2. **Excel Oluştur** butonuna basın.
        3. Oluşan çıktı dosyasını indirin.

        **Not:** Dosyada `Template` isimli sayfa bulunmalıdır.
        """
    )


# =========================================
# FILE UPLOAD
# =========================================

uploaded_file = st.file_uploader(
    "Excel dosyasını buraya sürükleyip bırak veya seç",
    type=["xlsx"],
    help="Sadece .xlsx formatındaki dosyalar desteklenir."
)

if uploaded_file is not None:
    st.session_state.uploaded_file_name = uploaded_file.name

    file_size = format_file_size(uploaded_file.size)

    st.success("Dosya başarıyla yüklendi.")

    col1, col2 = st.columns(2)

    with col1:
        st.metric("Yüklenen Dosya", uploaded_file.name)

    with col2:
        st.metric("Dosya Boyutu", file_size)

else:
    st.info("Henüz bir Excel dosyası yüklenmedi.")


# =========================================
# ACTION BUTTONS
# =========================================

st.divider()

button_col1, button_col2 = st.columns(2)

with button_col1:
    create_clicked = st.button(
        "⚙️ Excel Oluştur",
        type="primary",
        use_container_width=True,
        disabled=uploaded_file is None
    )

with button_col2:
    reset_clicked = st.button(
        "🔄 Yeni Dosya / Sıfırla",
        use_container_width=True
    )

if reset_clicked:
    reset_app()
    st.rerun()


# =========================================
# PROCESS EXCEL
# =========================================

if create_clicked and uploaded_file is not None:
    try:
        start_time = time.time()

        with st.spinner("Excel dosyası işleniyor, lütfen bekleyin..."):

            with tempfile.NamedTemporaryFile(
                delete=False,
                suffix=".xlsx"
            ) as temp_input:
                temp_input.write(uploaded_file.read())
                input_path = temp_input.name

            base_name = uploaded_file.name.replace(".xlsx", "")
            output_file_name = f"{base_name}_doldurulmus.xlsx"

            with tempfile.NamedTemporaryFile(
                delete=False,
                suffix=".xlsx"
            ) as temp_output:
                output_path = temp_output.name

            process_excel(
                input_path,
                output_path
            )

            elapsed_time = round(time.time() - start_time, 2)

            st.session_state.output_file_path = output_path
            st.session_state.output_file_name = output_file_name
            st.session_state.processed = True
            st.session_state.downloaded = False
            st.session_state.process_time = elapsed_time

        st.success("Excel başarıyla oluşturuldu.")

    except Exception as e:
        st.session_state.processed = False
        st.session_state.output_file_path = None
        st.session_state.output_file_name = None

        st.error("Excel oluşturulurken bir hata oluştu.")
        st.exception(e)


# =========================================
# DOWNLOAD AREA
# =========================================

if st.session_state.processed and st.session_state.output_file_path:

    st.divider()

    st.subheader("✅ Çıktı Hazır")

    if st.session_state.process_time is not None:
        st.caption(
            f"İşlem süresi: {st.session_state.process_time} saniye"
        )

    with open(st.session_state.output_file_path, "rb") as file:
        downloaded = st.download_button(
            label="📥 Çıktı Excel'i İndir",
            data=file,
            file_name=st.session_state.output_file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    if downloaded:
        st.session_state.downloaded = True

    if st.session_state.downloaded:
        st.success("Dosya indirildi. Yeni bir dosya işlemek için sıfırlayabilirsiniz.")


# =========================================
# FOOTER
# =========================================

st.divider()

st.caption(
    "E-Okul çizelge oluşturucu | Excel dosyanız yalnızca işlem sırasında kullanılır."
)
