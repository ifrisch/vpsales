import streamlit as st
import pandas as pd
from PIL import Image
import base64
from io import BytesIO
from datetime import datetime
import os

# --- CSS RESET TO REMOVE EXCESS SPACE ---
st.markdown(
    """
    <style>
    /* Remove padding and margin from main content container */
    .block-container {
        padding-top: 0 !important;
        padding-bottom: 0 !important;
        margin: 0 !important;
    }
    /* Remove margin and padding on html and body */
    html, body {
        margin: 0 !important;
        padding: 0 !important;
        height: 100%;
    }
    /* Also remove margin/padding from header and footer */
    header, footer {
        margin: 0 !important;
        padding: 0 !important;
        height: 0 !important;
        min-height: 0 !important;
        display: none !important;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# --- LOGO CENTERED ---
logo_path = "0005.jpg"  # Make sure this image is in the same folder as your script

def get_base64_image(image_path):
    img = Image.open(image_path)
    buffered = BytesIO()
    img.save(buffered, format="JPEG")
    return base64.b64encode(buffered.getvalue()).decode()

logo_base64 = get_base64_image(logo_path)

st.markdown(
    f"""
    <div style="text-align:center; margin: 0; padding: 0;">
        <img src="data:image/jpeg;base64,{logo_base64}" style="max-width: 400px; height: auto; margin: 0; padding: 0;" />
    </div>
    """,
    unsafe_allow_html=True,
)

# --- TITLE ---
st.markdown("<h1 style='text-align:center; margin-top: 0.5rem;'>📊 Salesrep Leaderboard</h1>", unsafe_allow_html=True)

# --- LOAD DATA ---
excel_path = "leaderboardexport.xlsx"  # Make sure this Excel file is in the same folder

try:
    df = pd.read_excel(excel_path, usecols="A:D")
    df.columns = ["New Customer", "Salesrep", "Ignore", "Last Invoice Date"]
    df = df.dropna(subset=["New Customer", "Salesrep"])
    df = df[df["Salesrep"].str.strip().str.lower() != "house account"]
    df["Last Invoice Date"] = pd.to_datetime(df["Last Invoice Date"], errors="coerce")

    leaderboard = (
        df.groupby("Salesrep")["New Customer"]
        .nunique()
        .reset_index()
        .rename(columns={"New Customer": "Number of New Customers"})
    )
    leaderboard = leaderboard.sort_values(by="Number of New Customers", ascending=False).reset_index(drop=True)

    def ordinal(n):
        suffixes = {1: "st", 2: "nd", 3: "rd"}
        if 10 <= n % 100 <= 20:
            suffix = "th"
        else:
            suffix = suffixes.get(n % 10, "th")
        return f"{n}{suffix}"

    leaderboard.insert(0, "Rank", [ordinal(i + 1) for i in range(len(leaderboard))])
    leaderboard = leaderboard.set_index("Rank")

    def highlight_first_salesrep(s):
        styles = pd.DataFrame("", index=s.index, columns=s.columns)
        if "1st" in s.index:
            styles.loc["1st", "Salesrep"] = "background-color: yellow; font-weight: bold;"
        return styles

    styled_leaderboard = leaderboard.style.apply(highlight_first_salesrep, axis=None)

    st.write(styled_leaderboard)

    last_updated = datetime.fromtimestamp(os.path.getmtime(excel_path))
    st.markdown(
        f"<div style='text-align: center; margin-top: 30px; color: gray;'>Last updated: {last_updated.strftime('%B %d, %Y at %I:%M %p')}</div>",
        unsafe_allow_html=True,
    )

except FileNotFoundError:
    st.error(f"File not found: {excel_path}")
except Exception as e:
    st.error(f"An error occurred: {e}")
