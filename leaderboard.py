import streamlit as st
import pandas as pd
from PIL import Image
import base64
from io import BytesIO
import os
from datetime import datetime
from fuzzywuzzy import fuzz
from st_aggrid import AgGrid, GridOptionsBuilder

# --- UPDATED CSS with different logo sizes for desktop and mobile ---
st.markdown("""
<style>
    .stApp > div:first-child {
        padding-top: 0rem !important;
    }
    .block-container {
        padding-top: 0rem !important;
        padding-bottom: 0rem !important;
        margin-top: -2rem !important;
    }
    .element-container {
        margin: 0px !important;
        padding: 0px !important;
    }
    div[data-testid="stMarkdownContainer"] {
        margin: 0px !important;
        padding: 0px !important;
    }
    div[data-testid="stImage"] {
        margin: -1rem 0px !important;
        padding: 0px !important;
        text-align: center !important;
    }
    div[data-testid="stImage"] > div {
        margin: 0px auto !important;
        padding: 0px !important;
        text-align: center !important;
    }
    div[data-testid="stImage"] img {
        margin: 0 auto !important;
        display: block !important;
        max-width: 480px !important;  /* larger max width on desktop */
        width: 60% !important;        /* responsive width on desktop */
        height: auto !important;
    }

    /* Mobile responsiveness */
    @media (max-width: 768px) {
        .stApp {
            text-align: center !important;
        }
        .block-container {
            padding-left: 1rem !important;
            padding-right: 1rem !important;
            text-align: center !important;
        }
        div[data-testid="stImage"] {
            text-align: center !important;
            width: 100% !important;
        }
        div[data-testid="stImage"] > div {
            text-align: center !important;
            margin: 0 auto !important;
            width: 100% !important;
        }
        div[data-testid="stImage"] img {
            margin: 0 auto !important;
            display: block !important;
            max-width: 280px !important;  /* smaller max width on mobile */
            width: 90% !important;
        }
        div[data-testid="stVerticalBlock"] {
            text-align: center !important;
        }
        div[data-testid="column"] {
            text-align: center !important;
        }

        /* Prevent Streamlit columns from stacking on mobile */
        .css-1lcbmhc.e1fqkh3o3 {
            flex-direction: row !important;
        }
    }
</style>
""", unsafe_allow_html=True)

# --- UPDATED LOGO DISPLAY ---
logo_path = "0005.jpg"  # Make sure this file is in your repo folder

st.markdown(f"""
<div style="display:flex; justify-content:center; margin-top: -10rem; margin-bottom: 1rem;">
    <img src="data:image/jpeg;base64,{base64.b64encode(open(logo_path, "rb").read()).decode()}"
        style="height:auto;"/>
</div>
""", unsafe_allow_html=True)

# --- TITLE ---
st.markdown("<h1 style='margin-top: -7rem; margin-bottom: -10rem;'>🏆 Salesrep Leaderboard</h1>", unsafe_allow_html=True)

# --- LOAD DATA ---
excel_path = "leaderboardexport.xlsx"  # relative path inside repo

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
