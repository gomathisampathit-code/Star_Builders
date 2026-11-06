from pathlib import Path
import streamlit as st
import streamlit.components.v1 as components

st.set_page_config(
    page_title="Star Builders",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
      header, footer, #MainMenu {visibility: hidden;}
      .block-container {padding: 0 !important; max-width: 100% !important;}

      /* Remove the outer Streamlit page scrollbar */
      html, body, .stApp {
          height: 100% !important;
          overflow: hidden !important;
          margin: 0 !important;
      }

      /* Make the component iframe fill the viewport */
      iframe {
          height: 100vh !important;
          width: 100% !important;
          border: none !important;
      }
    </style>
    """,
    unsafe_allow_html=True,
)

html = Path(__file__).with_name("index.html").read_text(encoding="utf-8")

components.html(
    html,
    height=0,           # CSS above forces the iframe to 100vh
    scrolling=True,     # keep scrolling only inside your HTML
)
