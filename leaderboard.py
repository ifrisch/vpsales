import streamlit as st
from PIL import Image
import base64
from io import BytesIO
import os
from datetime import datetime
import pandas as pd
from fuzzywuzzy import fuzz
from st_aggrid import AgGrid, GridOptionsBuilder

# --- HELPER: get base64 of image ---
def get_base64_image(image_path):
    img = Image.open(image_path)
    buffered = BytesIO()
    img.save(buffered, format="JPEG")
    return base64.b64encode(buffered.getvalue()).decode()

logo_path = "0005.jpg"
logo_base64 = get_base64_image(logo_path)

# --- Inject CSS to remove all spacing around main container, image, and title ---
st.markdown("""
    <style>
    /* Reset Streamlit's default app padding and margins */
    .stApp {
        margin: 0 !important;
        padding: 0 !important;
    }
    /* Remove padding/margin from main container and block container */
    .main, .block-container, .css-18e3th9, .css-1d391kg {
        padding: 0 !important;
        margin: 0 !important;
    }
    /* Ensure no spacing between elements */
    div.block-container > div {
        padding-top: 0 !important;
        padding-bottom: 0 !important;
        margin-top: 0 !important;
        margin-bottom: 0 !important;
    }
    /* Style the logo-title container */
    .logo-title-container {
        text-align: center;
        margin: 0 !important;
        padding: 0 !important;
        line-height: 1 !important;
    }
    /* Remove all spacing for the image */
    .logo-title-container img {
        display: block;
        margin: 0 !important;
        padding: 0 !important;
        max-width: 300px;
        height: auto;
        border: none !important;
    }
    /* Remove spacing for the title */
    .logo-title-container h1 {
        margin: 0 !important;
        padding: 0 !important;
        line-height: 1.1 !important;
        font-size: 2rem !important;
    }
    /* Ensure no extra spacing around markdown elements */
    .stMarkdown, .stMarkdown > div {
        margin: 0 !important;
        padding: 0 !important;
    }
    </style>
""", unsafe_allow_html=True)

# --- Render logo and title with zero spacing ---
st.markdown(f"""
    <div class="logo-title-container">
        <img src="data:image/jpeg;base64,{logo_base64}" alt="Logo" />
        <h1>🏆 Salesrep Leaderboard</h1>
    </div>
""", unsafe_allow_html=True)

# --- Rest of your leaderboard code ---
# (Append your existing leaderboard code here, starting from excel_path = "leaderboardexport.xlsx")