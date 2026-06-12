import streamlit as st
import tempfile
import os
import compare  # ton fichier compare.py

# ===== CONFIG APP =====
st.set_page_config(
    page_title="Comparateur XML ↔ EDI",
    page_icon="📄",
    layout="centered"
)

# ===== HEADER =====
st.title("📄 Comparateur XML ↔ EDI")
st.write("Chargez vos fichiers XML et EDI pour générer un Excel automatiquement.")

st.markdown("---")

# ===== UPLOAD =====
xml_file = st.file_uploader("📥 Importer fichier XML", type=["xml"])
edi_file = st.file_uploader("📥 Importer fichier EDI", type=["txt", "edi"])

st.markdown("---")

# ===== ACTION =====
if st.button("🚀 Lancer la comparaison"):

    if xml_file is None or edi_file is None:
        st.warning("⚠️ Merci de charger les deux fichiers.")
    else:
        with st.spinner("⏳ Traitement en cours..."):

            # Sauvegarde temporaire des fichiers
            tmp_xml = tempfile.NamedTemporaryFile(delete=False, suffix=".xml")
            tmp_xml.write(xml_file.read())
            tmp_xml.close()

            tmp_edi = tempfile.NamedTemporaryFile(delete=False, suffix=".edi")
            tmp_edi.write(edi_file.read())
            tmp_edi.close()

            try:
                # === APPEL DE TON SCRIPT ===
                edi_segments = compare.extract_edi_segments(tmp_edi.name)
                rows = compare.extract_xml_pairs(tmp_xml.name)

                output_file = compare.write_excel(rows, edi_segments)

                st.success("✅ Fichier généré avec succès !")

                # === BOUTON TELECHARGEMENT ===
                with open(output_file, "rb") as f:
                    st.download_button(
                        label="📥 Télécharger le fichier Excel",
                        data=f,
                        file_name="compare.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

            except Exception as e:
                st.error(f"❌ Une erreur est survenue : {e}")

            finally:
                # Nettoyage
                os.remove(tmp_xml.name)
                os.remove(tmp_edi.name)

# ===== FOOTER =====
st.markdown("---")
st.caption("Outil interne - Comparateur XML / EDI")
