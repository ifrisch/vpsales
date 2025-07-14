import streamlit as st
import pandas as pd
from PIL import Image
from io import BytesIO
import base64
import os
from datetime import datetime
from fuzzywuzzy import fuzz
from st_aggrid import AgGrid, GridOptionsBuilder

# --- Aggressive CSS to remove padding/margin and center logo ---
st.markdown("""
    <style>
    /* Remove all padding and margin in Streamlit main container */
    .block-container {
        padding-top: 0 !important;
        padding-bottom: 0 !important;
        margin-top: 0 !important;
        margin-bottom: 0 !important;
    }
    /* Remove header and footer */
    header, footer {
        display: none !important;
    }
    /* Remove margins/padding from body and html */
    html, body {
        margin: 0 !important;
        padding: 0 !important;
        height: 100%;
    }
    /* Center container for logo */
    .logo-container {
        margin: 0 !important;
        padding: 0 !important;
        text-align: center;
        line-height: 0;
    }
    .logo-container img {
        max-width: 400px;
        height: auto;
        margin: 0 !important;
        padding: 0 !important;
        display: inline-block;
        vertical-align: top;
    }
    h1 {
        text-align: center;
        margin-top: 1rem;
        margin-bottom: 1rem;
        font-size: clamp(1.5rem, 5vw, 2.5rem);
    }
    .leaderboard-container {
        max-width: 600px;
        margin-left: auto;
        margin-right: auto;
        padding-left: 0.5rem;
        padding-right: 0.5rem;
    }
    .ag-theme-streamlit {
        border: 1px solid white !important;
        box-shadow: none !important;
    }
    </style>
""", unsafe_allow_html=True)

# --- Load and display logo ---
logo_path = r"C:\Users\Isaac\Downloads\0005.jpg"  # Make sure this path is correct

def get_base64_image(image_path):
    img = Image.open(image_path)
    buffer = BytesIO()
    img.save(buffer, format="JPEG")
    return base64.b64encode(buffer.getvalue()).decode()

logo_base64 = get_base64_image(logo_path)

st.markdown(f"""
    <div class="logo-container">
        <img src="data:image/jpeg;base64,{logo_base64}" alt="Logo" />
    </div>
""", unsafe_allow_html=True)

# --- Title ---
st.markdown("<h1>📊 Salesrep Leaderboard</h1>", unsafe_allow_html=True)

# --- Load Excel data ---
excel_path = r"C:\Users\Isaac\Downloads\leaderboardexport.xlsx"

try:
    df = pd.read_excel(excel_path, usecols="A:D", dtype={"A": str, "B": str})
    df.columns = ["New Customer", "Salesrep", "Ignore", "Last Invoice Date"]
    df = df.dropna(subset=["New Customer", "Salesrep"])
    df = df[df["Salesrep"].str.strip().str.lower() != "house account"]
    df["Last Invoice Date"] = pd.to_datetime(df["Last Invoice Date"], errors="coerce")
    df["Cleaned Customer"] = df["New Customer"].str.strip().str.lower()

    used_customers = set()
    kept_rows = []
    pending_rows = []

    for i, row in df.iterrows():
        cust_name = row["Cleaned Customer"]
        if cust_name in used_customers:
            continue

        matches = df[df["Cleaned Customer"].apply(lambda x: fuzz.token_sort_ratio(x, cust_name) >= 90)].copy()
        used_customers.update(matches["Cleaned Customer"].tolist())

        matches_with_invoice = matches[~matches["Last Invoice Date"].isna()]
        if not matches_with_invoice.empty:
            best_match = matches_with_invoice.sort_values(by="Last Invoice Date", ascending=False).iloc[0]
            kept_rows.append(best_match)
        else:
            pending_rows.append(matches.iloc[0])

    df_cleaned = pd.DataFrame(kept_rows)
    df_pending = pd.DataFrame(pending_rows)

    # --- LEADERBOARD ---
    leaderboard = df_cleaned.groupby("Salesrep")["New Customer"].nunique().reset_index()
    leaderboard = leaderboard.rename(columns={"New Customer": "Number of New Customers"})
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

    st.markdown('<div class="leaderboard-container">', unsafe_allow_html=True)
    st.write(styled_leaderboard)
    st.markdown('</div>', unsafe_allow_html=True)

    last_updated = datetime.fromtimestamp(os.path.getmtime(excel_path))
    st.markdown(
        f"<div style='text-align: center; margin-top: 30px; color: gray;'>Last updated: {last_updated.strftime('%B %d, %Y at %I:%M %p')}</div>",
        unsafe_allow_html=True
    )

    # --- PENDING CUSTOMERS ---
    st.markdown("<h2>⏲ Pending Customers</h2>", unsafe_allow_html=True)

    if not df_pending.empty:
        for salesrep, group_df in df_pending.groupby("Salesrep"):
            st.markdown(f"<h4>{salesrep}</h4>", unsafe_allow_html=True)
            rows = len(group_df)
            grid_height = 40 + rows * 35

            gb = GridOptionsBuilder.from_dataframe(group_df[["New Customer", "Last Invoice Date"]].reset_index(drop=True))
            gb.configure_grid_options(domLayout='normal')
            gb.configure_default_column(resizable=True, filter=True, sortable=True)
            gb.configure_column("Last Invoice Date", hide=True)
            gridOptions = gb.build()

            AgGrid(
                group_df[["New Customer", "Last Invoice Date"]].reset_index(drop=True),
                gridOptions=gridOptions,
                fit_columns_on_grid_load=True,
                enable_enterprise_modules=False,
                height=grid_height,
                theme='streamlit',
            )
    else:
        st.info("No pending customers! 🎉")

except FileNotFoundError:
    st.error(f"File not found: {excel_path}")
except Exception as e:
    st.error(f"An error occurred: {e}")
