import streamlit as st
import pandas as pd
from PIL import Image
import base64
from io import BytesIO
import os
from datetime import datetime
from zoneinfo import ZoneInfo
from fuzzywuzzy import fuzz
from st_aggrid import AgGrid, GridOptionsBuilder

# --- CSS ---
st.markdown("""
<style>
    .stApp > div:first-child {
        padding-top: 0rem !important;
    }
    .block-container {
        padding-top: 0rem !important;
        padding-bottom: 0rem !important;
    }

    #logo-block {
        margin-top: -6rem;
        margin-bottom: -10rem;
        text-align: center;
    }

    #main-block {
        margin-top: -12rem;
    }

    @media (max-width: 768px) {
        #logo-block img {
            max-width: 280px !important;
            width: 90% !important;
        }
        #main-block {
            margin-top: -20rem !important;
        }
    }

    img.logo-img {
        max-width: 480px;
        width: 60%;
        height: auto;
        display: block;
        margin: 0 auto;
    }
</style>
""", unsafe_allow_html=True)

# --- LOGO BLOCK ---
logo_path = "0005.jpg"
encoded_logo = base64.b64encode(open(logo_path, "rb").read()).decode()
st.markdown(f"""
<div id="logo-block">
    <img src="data:image/jpeg;base64,{encoded_logo}" class="logo-img" />
</div>
""", unsafe_allow_html=True)

# --- MAIN CONTENT BLOCK ---
st.markdown('<div id="main-block">', unsafe_allow_html=True)

# --- TITLE ---
st.markdown("<h1 style='margin-top: 0rem; margin-bottom: 1rem;'>🏆 Leaderboard</h1>", unsafe_allow_html=True)

# --- LOAD DATA ---
excel_path = "leaderboard.xlsx"

try:
    df = pd.read_excel(excel_path, usecols="A:D", dtype={"A": str, "B": str})
    df.columns = ["New Customer", "Salesrep", "Ignore", "Last Invoice Date"]
    df = df.dropna(subset=["New Customer", "Salesrep"])
    df = df[df["Salesrep"].str.strip().str.lower() != "house account"]
    df["Last Invoice Date"] = pd.to_datetime(df["Last Invoice Date"], errors="coerce")

    # Clean customer names
    df["Cleaned Customer"] = df["New Customer"].str.lower()
    df["Cleaned Customer"] = df["Cleaned Customer"].str.replace(r'[^\w\s]', '', regex=True)
    df["Cleaned Customer"] = df["Cleaned Customer"].str.replace(r'\s+', ' ', regex=True).str.strip()

    used_customers = set()
    kept_rows = []
    pending_rows = []

    # Fuzzy matching to group similar customers
    for i, row in df.iterrows():
        cust_name = row["Cleaned Customer"]
        if cust_name in used_customers:
            continue

        matches = df[df["Cleaned Customer"].apply(lambda x: fuzz.token_set_ratio(x, cust_name) >= 85)].copy()
        used_customers.update(matches["Cleaned Customer"].tolist())

        matches_with_invoice = matches[~matches["Last Invoice Date"].isna()]
        if not matches_with_invoice.empty:
            best_match = matches_with_invoice.sort_values(by="Last Invoice Date", ascending=False).iloc[0]
            kept_rows.append(best_match)
        else:
            pending_rows.append(matches.iloc[0])

    df_cleaned = pd.DataFrame(kept_rows)
    df_pending = pd.DataFrame(pending_rows)

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

    # Prize Logic
    top_count = leaderboard["Number of New Customers"].max()
    top_reps = leaderboard[leaderboard["Number of New Customers"] == top_count]
    split_prize = 100 / len(top_reps)

    def calculate_prize(row):
        prize = 0
        if row["Number of New Customers"] >= 3:
            prize += 50
        if row["Number of New Customers"] == top_count:
            prize += split_prize
        return f"${int(prize)}"

    leaderboard["Prize"] = leaderboard.apply(calculate_prize, axis=1)

    # Set Rank as index
    leaderboard = leaderboard.set_index("Rank")

    # --- Highlight all first-place reps ---
    def highlight_first_place(s):
        styles = pd.DataFrame("", index=s.index, columns=s.columns)
        for idx in s.index:
            if "1st" in idx:
                styles.loc[idx, "Salesrep"] = "background-color: yellow; font-weight: bold;"
        return styles

    styled_leaderboard = leaderboard.style.apply(highlight_first_place, axis=None)
    st.write(styled_leaderboard)

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

# --- Close MAIN BLOCK ---
st.markdown('</div>', unsafe_allow_html=True)

# --- LAST UPDATED TIMESTAMP (based on file modification) ---
timestamp_path = excel_path
if os.path.exists(timestamp_path):
    central = ZoneInfo("America/Chicago")
    last_modified = datetime.fromtimestamp(os.path.getmtime(timestamp_path), tz=central)
    st.markdown(
        f"<div style='text-align: center; margin-top: 30px; color: gray;'>Last updated: {last_modified.strftime('%B %d, %Y at %I:%M %p')}</div>",
        unsafe_allow_html=True
    )
