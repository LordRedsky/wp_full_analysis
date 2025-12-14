import streamlit as st
from exact_match import process_exact_match
from deep_analysis import process_deep_analysis
from analysis_sert import process_certificate_analysis
from final_analysis import process_final_analysis


st.set_page_config(page_title="Deep Analysis", page_icon="📊", layout="centered")
st.title("📊 Deep Analysis - Data Wajib Pajak")
st.markdown(
    """
    <style>
    .stButton > button {
      background: linear-gradient(90deg, #FF7F50, #FF5E57);
      color: #ffffff; border: none; border-radius: 8px; padding: 8px 16px; font-weight: 600;
    }
    .stDownloadButton > button {
      background: linear-gradient(90deg, #4B7BEC, #3867D6);
      color: #ffffff; border: none; border-radius: 8px; padding: 8px 16px; font-weight: 600;
    }
    .stButton > button:hover, .stDownloadButton > button:hover { filter: brightness(1.05); }
    </style>
    """,
    unsafe_allow_html=True,
)

def _make_name(orig: str, suffix: str):
    import os
    base, ext = os.path.splitext(orig or "")
    if not base:
        base = "file"
    if not ext:
        ext = ".xlsx"
    return f"{base}_{suffix}{ext}"

# Certificate Analysis Section (should be done first)
st.subheader(" Tahap 00 - Analysis Sertifikat")
st.markdown(
    """
    **Deskripsi:** Melakukan Analisa Untuk Mengecek Data NIB yang Sama.
    """
)
cert_file = st.file_uploader("Upload file Sertifikat", type=["xlsx", "xls"], key="cert_analysis")

if st.button("Proses Analysis Sertifikat"):
    if not cert_file:
        st.error("Upload file Sertifikat terlebih dahulu")
    else:
        result = process_certificate_analysis(
            sertifikat_file=cert_file.getvalue(),
        )
        st.session_state["cert_analyzed_bytes"] = result['workbook_bytes']
        st.session_state["cert_orig_name"] = getattr(cert_file, "name", "sertifikat.xlsx")
        st.session_state["cert_name_analyzed"] = _make_name(st.session_state["cert_orig_name"], "analysed")
        # Menyimpan informasi statistik ke session_state untuk ditampilkan
        st.session_state["data_diproses"] = result['data_diproses']
        st.session_state["nib_duplikat"] = result['nib_duplikat']
        st.success("Analysis Sertifikat selesai")

if "cert_analyzed_bytes" in st.session_state:
    # Tampilkan statistik hasil analisis sertifikat
    if "data_diproses" in st.session_state and "nib_duplikat" in st.session_state:
        col1, col2 = st.columns(2)
        with col1:
            st.metric(label="Data Diproses", value=st.session_state["data_diproses"])
        with col2:
            st.metric(label="NIB Duplikat", value=st.session_state["nib_duplikat"])

    st.download_button(
        label="Download Sertifikat (analysed)",
        data=st.session_state["cert_analyzed_bytes"],
        file_name=st.session_state.get("cert_name_analyzed", "sertifikat_analysed.xlsx"),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download_cert_analyzed",
    )

# Exact Match Section
st.subheader(" Tahap 01 - Exact Match")
st.markdown(
    """
    **Deskripsi:** Melakukan pencocokan data antara file Wajib Pajak, file Sertifikat dan file Form Kendali dengan Prinsip Exact Match.
    """
)
wp_file = st.file_uploader("Upload file Wajib Pajak", type=["xlsx", "xls"], key="wp")
cert_file_exact = st.file_uploader("Upload file Sertifikat (optional, default: analyzed certificate)", type=["xlsx", "xls"], key="cert_exact")
kendali_file = st.file_uploader("Upload file Form Kendali", type=["xlsx", "xls"], key="kendali")

if st.button("Proses Exact Match"):
    # Use analyzed certificate if available, otherwise use uploaded one for exact match
    if "cert_analyzed_bytes" in st.session_state:
        cert_source = st.session_state["cert_analyzed_bytes"]
    elif cert_file_exact:
        cert_source = cert_file_exact.getvalue()
    else:
        cert_source = cert_file.getvalue() if cert_file else None

    if not (wp_file and cert_source and kendali_file):
        st.error("Lengkapi ketiga file terlebih dahulu")
    else:
        wp_bytes, cert_bytes, stats = process_exact_match(
            wajib_pajak_file=wp_file.getvalue(),
            sertifikat_file=cert_source,  # Use certificate that was analyzed or uploaded
            kendali_file=kendali_file.getvalue(),
        )
        st.session_state["wp_bytes"] = wp_bytes
        st.session_state["cert_bytes"] = cert_bytes
        st.session_state["kendali_bytes"] = kendali_file.getvalue()
        st.session_state["kendali_orig_name"] = getattr(kendali_file, "name", "FORM KENDALI VILLAGE.xlsx")
        st.session_state["stats"] = stats
        st.session_state["wp_orig_name"] = getattr(wp_file, "name", "wajib_pajak.xlsx")
        st.session_state["cert_orig_name"] = getattr(cert_file_exact, "name", "sertifikat.xlsx") if cert_file_exact else (
            getattr(cert_file, "name", "sertifikat.xlsx") if cert_file else "sertifikat.xlsx"
        )
        st.session_state["wp_name_update"] = _make_name(st.session_state["wp_orig_name"], "update")
        st.session_state["cert_name_update"] = _make_name(st.session_state["cert_orig_name"], "update")
        st.success("Proses selesai")

if "wp_bytes" in st.session_state and "cert_bytes" in st.session_state:
    if "stats" in st.session_state:
        s = st.session_state["stats"]
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Diproses", s.get("processed", 0))
        c2.metric("Matched", s.get("matched", 0))
        c3.metric("NIB Not Found", s.get("nib_not_found", 0))
        c4.metric("Nama Not Found", s.get("name_not_found", 0))
        c5.metric("Ahli waris", s.get("ahli_waris", 0))
    st.download_button(
        label="Download Wajib Pajak (updated)",
        data=st.session_state["wp_bytes"],
        file_name=st.session_state.get("wp_name_update", "wajib_pajak_update.xlsx"),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download_wp",
    )
    st.download_button(
        label="Download Sertifikat (updated)",
        data=st.session_state["cert_bytes"],
        file_name=st.session_state.get("cert_name_update", "sertifikat_update.xlsx"),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download_cert",
    )

# Deep Analysis Section
st.subheader(" Tahap 02 - Deep Analysis")
st.markdown(
    """
    **Deskripsi:** Melakukan Analisa Mendalam pada file WP yang sudah diupdate oleh Exact Match.
    """
)

if st.button("Proses Deep Analysis"):
    if not ("wp_bytes" in st.session_state and "cert_bytes" in st.session_state and "kendali_bytes" in st.session_state):
        st.error("Jalankan Exact Match terlebih dahulu")
    else:
        wp_bytes2, team_bytes2, cert_bytes2, stats2 = process_deep_analysis(
            wajib_pajak_file=st.session_state["wp_bytes"],
            sertifikat_file=st.session_state["cert_bytes"],
            kendali_file=st.session_state["kendali_bytes"],
        )
        st.session_state["wp_bytes_deep"] = wp_bytes2
        st.session_state["wp_bytes_team"] = team_bytes2
        st.session_state["cert_bytes_deep"] = cert_bytes2
        st.session_state["stats_deep"] = stats2
        st.session_state["wp_name_deep"] = _make_name(st.session_state.get("wp_orig_name", "wajib_pajak.xlsx"), "deep analysis")
        st.session_state["wp_name_team"] = _make_name(st.session_state.get("wp_orig_name", "wajib_pajak.xlsx"), "tim pemetaan")
        st.session_state["cert_name_deep"] = _make_name(st.session_state.get("cert_orig_name", "sertifikat.xlsx"), "deep analysis")
        st.success("Deep Analysis selesai")

if "wp_bytes_deep" in st.session_state and "cert_bytes_deep" in st.session_state:
    if "stats_deep" in st.session_state:
        sd = st.session_state["stats_deep"]
        s = st.session_state.get("stats", {})
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Data diproses", sd.get("processed_grey", 0))
        c2.metric("NIB Match", sd.get("matched", sd.get("updated", 0)))
        c3.metric("NIB Not Found", sd.get("nib_not_found", 0))
        c4.metric("Nama Not Found", sd.get("name_not_found", 0))
        c5.metric("Ahli waris", s.get("ahli_waris", 0))
    st.download_button(
        label="Download Wajib Pajak (deep updated)",
        data=st.session_state["wp_bytes_deep"],
        file_name=st.session_state.get("wp_name_deep", "wajib_pajak_deep analysis.xlsx"),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download_wp_deep",
    )
    if "wp_bytes_team" in st.session_state:
        st.download_button(
            label="Download Wajib Pajak (tim pemetaan)",
            data=st.session_state["wp_bytes_team"],
            file_name=st.session_state.get("wp_name_team", _make_name(st.session_state.get("wp_orig_name", "wajib_pajak.xlsx"), "tim pemetaan")),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_wp_team",
        )
    st.download_button(
        label="Download Sertifikat (deep updated)",
        data=st.session_state["cert_bytes_deep"],
        file_name=st.session_state.get("cert_name_deep", "sertifikat_deep analysis.xlsx"),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download_cert_deep",
    )

# Final Analysis Section
st.subheader(" Tahap 03 - Final Analysis")
st.markdown(
    """
    **Deskripsi:** Proses final analysis menggunakan file WP kosong dari folder 'empty',
    file sertifikat hasil deep analysis, dan file form kendali awal. Proses ini akan mengisi
    file WP kosong dengan data dari sertifikat, lalu melakukan analisis seperti exact match.
    """
)

if st.button("Proses Final Analysis"):
    # Check if deep analysis results are available to use as cert_file
    if not ("cert_bytes_deep" in st.session_state and "kendali_bytes" in st.session_state):
        st.error("Jalankan Deep Analysis terlebih dahulu untuk mendapatkan file sertifikat dan kendali")
    else:
        # Load the empty WP file from the empty folder
        import os
        empty_wp_path = os.path.join("empty", "WP NAMA DESA.xlsx")

        if not os.path.exists(empty_wp_path):
            st.error(f"File WP kosong tidak ditemukan di {empty_wp_path}")
        else:
            with open(empty_wp_path, "rb") as f:
                empty_wp_bytes = f.read()

            # Extract village name from kendali filename for naming the output files
            from final_analysis import extract_village_name_from_filename
            # Get kendali filename from session state (stored during exact match)
            kendali_orig_name = st.session_state.get("kendali_orig_name", "FORM KENDALI VILLAGE.xlsx")
            village_name_full = extract_village_name_from_filename(kendali_orig_name)
            # Extract just the village name part (after "DESA ")
            if village_name_full.startswith("DESA "):
                village_name = village_name_full[5:]  # Remove "DESA " prefix
            else:
                village_name = village_name_full

            with st.spinner("Sedang memproses final analysis..."):
                try:
                    wp_bytes_final, team_bytes_final, cert_bytes_final, stats_final = process_final_analysis(
                        wajib_pajak_file=empty_wp_bytes,
                        sertifikat_file=st.session_state["cert_bytes_deep"],  # Use certificate from deep analysis
                        kendali_file=st.session_state["kendali_bytes"],  # Use kendali from exact match
                    )
                    st.session_state["wp_bytes_final"] = wp_bytes_final
                    st.session_state["wp_bytes_team_final"] = team_bytes_final
                    st.session_state["cert_bytes_final"] = cert_bytes_final
                    st.session_state["stats_final"] = stats_final
                    st.session_state["wp_orig_name_final"] = f"WP_{village_name}_final_analysis.xlsx"
                    st.session_state["cert_orig_name_final"] = f"SERT_{village_name}_final_analysis.xlsx"
                    st.session_state["wp_team_orig_name_final"] = f"WP_{village_name}_pemetaan_final_analysis.xlsx"
                    st.session_state["wp_name_final"] = st.session_state["wp_orig_name_final"]
                    st.session_state["wp_name_team_final"] = st.session_state["wp_team_orig_name_final"]
                    st.session_state["cert_name_final"] = st.session_state["cert_orig_name_final"]
                    st.success("Final Analysis selesai")
                except Exception as e:
                    st.error(f"Terjadi kesalahan saat memproses final analysis: {str(e)}")

if "wp_bytes_final" in st.session_state and "cert_bytes_final" in st.session_state:
    if "stats_final" in st.session_state:
        sf = st.session_state["stats_final"]
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Diproses", sf.get("processed", 0))
        c2.metric("Matched", sf.get("matched", 0))
        c3.metric("NIB Not Found", sf.get("nib_not_found", 0))
        c4.metric("Nama Not Found", sf.get("name_not_found", 0))
    st.download_button(
        label="Download Wajib Pajak (final result)",
        data=st.session_state["wp_bytes_final"],
        file_name=st.session_state.get("wp_name_final", "wp_final_result.xlsx"),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download_wp_final",
    )
    if "wp_bytes_team_final" in st.session_state:
        st.download_button(
            label="Download Wajib Pajak (team result)",
            data=st.session_state["wp_bytes_team_final"],
            file_name=st.session_state.get("wp_name_team_final", "wp_team_result.xlsx"),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_wp_team_final",
        )
    st.download_button(
        label="Download Sertifikat (final result)",
        data=st.session_state["cert_bytes_final"],
        file_name=st.session_state.get("cert_name_final", "sertifikat_final_result.xlsx"),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download_cert_final",
    )
