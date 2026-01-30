import streamlit as st
import uuid
import os
from generate_ppt import generate_ppt

# ===== CONFIG STREAMLIT =====
st.set_page_config(
    page_title="Gerador de PPT",
    layout="centered"
)

# ===== PASTAS =====
os.makedirs("uploads", exist_ok=True)
os.makedirs("outputs", exist_ok=True)

# ===== UI =====
st.image("assets/logo.jpg", width=180)
st.title("Gerador de Apresentações")

template_file = st.file_uploader(
    "Template PowerPoint (.pptx)",
    type=["pptx"]
)

content_file = st.file_uploader(
    "Arquivo de conteúdo (.txt)",
    type=["txt"]
)

# ===== AÇÃO =====
if st.button("Gerar apresentação"):
    if not template_file or not content_file:
        st.error("Envie o template e o conteúdo.")
    else:
        with st.spinner("Gerando apresentação..."):
            uid = str(uuid.uuid4())

            template_path = f"uploads/{uid}_template.pptx"
            content_path = f"uploads/{uid}_conteudo.txt"
            output_path = f"outputs/{uid}_output.pptx"

            with open(template_path, "wb") as f:
                f.write(template_file.read())

            with open(content_path, "wb") as f:
                f.write(content_file.read())

            generate_ppt(
                template_path=template_path,
                content_path=content_path,
                output_path=output_path
            )

        st.success("Apresentação gerada com sucesso!")

        with open(output_path, "rb") as f:
            st.download_button(
                "📥 Baixar apresentação",
                f,
                file_name="apresentacao.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )