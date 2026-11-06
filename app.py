from pathlib import Path
import streamlit as st
import streamlit.components.v1 as components

st.set_page_config(page_title="Star Builders", page_icon="🏗️",
                   layout="wide", initial_sidebar_state="collapsed")

# Remove Streamlit's default padding/header so the app fills the screen
st.markdown(
    """
    <style>
      header, footer, #MainMenu {visibility: hidden;}
      .block-container {padding: 0 !important; max-width: 100% !important;}
    </style>
    """,
    unsafe_allow_html=True,
)

# index.html must be in the SAME folder as this app.py
html = Path(__file__).with_name("index.html").read_text(encoding="utf-8")
components.html(html, height=1400, scrolling=True)
