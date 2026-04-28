"""
Streamlit UI – PDF → PowerPoint via Gemini
"""

import os
import tempfile
import streamlit as st
from main import extract_pdf_text, call_gemini, build_pptx

st.set_page_config(page_title="PDF → PPT", page_icon="📊", layout="centered")

st.title("📄 → 📊  PDF to PowerPoint")
st.caption("Powered by Google Gemini AI")

api_key = st.text_input(
    "Gemini API Key",
    type="password",
    value=os.getenv("GEMINI_API_KEY", ""),
    placeholder="Paste your Gemini API key here",
)

uploaded = st.file_uploader("Upload PDF", type=["pdf"])

if uploaded and api_key:
    if st.button("🚀 Generate Presentation", use_container_width=True):
        os.environ["GEMINI_API_KEY"] = api_key

        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as f:
            f.write(uploaded.read())
            pdf_path = f.name

        output_path = pdf_path.replace(".pdf", ".pptx")

        try:
            progress = st.progress(0, text="Extracting PDF text…")
            text = extract_pdf_text(pdf_path)
            progress.progress(33, text=f"Extracted {len(text):,} characters.")

            progress.progress(40, text="Gemini is analysing…")
            data = call_gemini(text)
            n = len(data.get("slides", []))
            progress.progress(80, text=f"Got {n} slides from Gemini.")

            progress.progress(90, text="Building PowerPoint…")
            build_pptx(data, output_path)
            progress.progress(100, text="Done!")

            st.success(f"✅ **{data['title']}** — {n} slides")

            with open(output_path, "rb") as f:
                st.download_button(
                    "⬇️  Download PowerPoint",
                    data=f.read(),
                    file_name="final_presentation.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True,
                )

        except Exception as e:
            st.error(f"❌ {e}")
        finally:
            for p in [pdf_path, output_path]:
                try:
                    os.unlink(p)
                except Exception:
                    pass

elif uploaded and not api_key:
    st.warning("Enter your Gemini API key above.")
