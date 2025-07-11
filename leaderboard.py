import streamlit as st
import pandas as pd
from PIL import Image
import base64
from io import BytesIO
import os
from datetime import datetime
from fuzzywuzzy import fuzz
from st_aggrid import AgGrid, GridOptionsBuilder

# --- CSS FIX: LOGO fixed in viewport top-left, content padded to avoid overlap ---
st.markdown(
    """
    <style>
    .fixed-logo {
        position: fixed;
        top: 1rem;
        left: 1rem;
        width: min(150px, 25vw);
        height: auto;
        z-index: 9999;
    }
    main > div:first-child {
        margin-left: 170px;
        margin-top: 20px;
    }
    @media (max-width: 600px) {
        .fixed-logo {
            width: 80px;
        }
        main > div:first-child {
            margin-left: 90px;
        }
    }
    h1 {
        font-size: clamp(1.5rem, 5vw, 2.5rem);
        text-align: center;
        margin-top: 2rem;
    }
    h4 {
        font-size: clamp(0.9rem, 3vw, 1.2rem);
        margin-bottom: 0.3rem;
        color: #555;
        font-style: italic;
    }
    .leaderboard-container {
        max-width: 600px;
        margin-left: auto;
        margin-right: auto;
        padding-left: 0.5rem;
        padding-right: 0.5rem;
    }
    h2 {
        margin-top: 3rem !important;
        font-size: clamp(1.2rem, 4vw, 1.8rem);
    }
    .ag-theme-streamlit {
        border: 1px solid white !important;
        box-shadow: none !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# --- LOGO ---
logo_path = "logo2.png"

def get_base64_image(image_path):
    img = Image.open(image_path)
    buffered = BytesIO()
    img.save(buffered, format="PNG")
    return base64.b64encode(buffered.getvalue()).decode()

logo_base64 = get_base64_image(logo_path)

st.markdown(
    f"""
    <img src="data:image/png;base64,{logo_base64}" class="fixed-logo" />
    """,
    unsafe_allow_html=True
)

# --- TITLE ---
st.markdown("<h1>📊 Salesrep Leaderboard</h1>", unsafe_allow_html=True)

# --- LOAD DATA ---
excel_path = r"C:\\Users\\Isaac\\Downloads\\leaderboardexport.xlsx"

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
