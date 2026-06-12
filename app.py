import streamlit as st
import tempfile
import os
import compare

st.set_page_config(page_title="Comparateur XML ↔ EDI")

st.title("📄 Comparateur XML ↔ EDI")
st.write("Chargez vos fichiers XML et EDI pour générer un Excel.")

# Upload
xml_file = st.file_uploader("Fichier XML", type=["xml"], key="xml")
edi_file = st.file_uploader("Fichier EDI", type=["txt", "edi"], key="edi")

# Bouton
if st.button("Lancer la comparaison", key="run"):

    if not xml_file or not edi_file:
        st.warning("Merci de charger les deux fichiers.")
    else:
        with st.spinner("Traitement..."):

            tmp_xml = tempfile.NamedTemporaryFile(delete=False, suffix=".xml")
            tmp_xml.write(xml_file.read())
            tmp_xml.close()

            tmp_edi = tempfile.NamedTemporaryFile(delete=False, suffix=".edi")
            tmp_edi.write(edi_file.read())
            tmp_edi.close()

            try:
                edi_segments = compare.extract_edi_segments(tmp_edi.name)
                rows = compare.extract_xml_pairs(tmp_xml.name)

                output = compare.write_excel(rows, edi_segments)

                st.success("Fichier généré ✅")

                with open(output, "rb") as f:
                    st.download_button(
                        "Télécharger Excel",
                        f,
                        file_name="compare.xlsx",
                        key="dl"
                    )

            except Exception as e:
                st.error(str(e))

            finally:
                os.remove(tmp_xml.name)
                os.remove(tmp_edi.name)
