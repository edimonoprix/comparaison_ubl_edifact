import streamlit as st
import tempfile
import os
import compare

# ===== CONFIG =====
st.set_page_config(page_title="Comparateur XML ↔ EDI")

# ===== UI =====
st.title("📄 Comparateur XML ↔ EDI")
st.write("Chargez vos fichiers XML et EDI pour générer un Excel.")

# Upload fichiers
xml_file = st.file_uploader("📥 Fichier XML", type=["xml"], key="xml")
edi_file = st.file_uploader("📥 Fichier EDI", type=["txt", "edi"], key="edi")

# Bouton (clé unique très important)
if st.button("🚀 Lancer la comparaison", key="run"):

    if xml_file is None or edi_file is None:
        st.warning("⚠️ Merci de charger les deux fichiers.")
    else:
        with st.spinner("Traitement en cours..."):

            # fichiers temporaires
            tmp_xml = tempfile.NamedTemporaryFile(delete=False, suffix=".xml")
            tmp_xml.write(xml_file.read())
            tmp_xml.close()

            tmp_edi = tempfile.NamedTemporaryFile(delete=False, suffix=".edi")
            tmp_edi.write(edi_file.read())
            tmp_edi.close()

            try:
                edi_segments = compare.extract_edi_segments(tmp_edi.name)
                rows = compare.extract_xml_pairs(tmp_xml.name)

                output_file = compare.write_excel(rows, edi_segments)

                st.success("✅ Fichier généré !")

                with open(output_file, "rb") as f:
                    st.download_button(
                        "📥 Télécharger le fichier Excel",
                        f,
                        file_name="compare.xlsx",
                        key="download"
                    )

            except Exception as e:
                st.error(f"Erreur : {e}")

            finally:
                os.remove(tmp_xml.name)
                os.remove(tmp_edi.name)
