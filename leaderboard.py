import streamlit as st
import pandas as pd
import base64
import os
from datetime import datetime
from fuzzywuzzy import fuzz
from st_aggrid import AgGrid, GridOptionsBuilder
import re

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

st.markdown('<div id="main-block">', unsafe_allow_html=True)

st.markdown("<h1 style='margin-top: 0rem; margin-bottom: 1rem;'>🏆 Leaderboard</h1>", unsafe_allow_html=True)

excel_path = "leaderboardexport.xlsx"

def clean_customer_name(name):
    name = str(name).lower()
    name = re.sub(r'#\d+', '', name)            # remove #1, #2, etc.
    name = re.sub(r'\bgrill\b', '', name)       # remove "grill"
    name = re.sub(r'[^a-z0-9\s]', '', name)     # remove punctuation
    name = re.sub(r'\s+', ' ', name).strip()    # normalize whitespace
    return name

def assign_clusters(names, threshold=85):
    """
    Assign cluster IDs so that names with fuzzy ratio >= threshold get same cluster.
    Simple hierarchical clustering approach: 
    - Initialize clusters empty
    - For each name, assign to first cluster with matching representative or create new
    """
    clusters = []
    cluster_ids = {}
    current_id = 0

    for name in names:
        assigned = False
        for idx, cluster in enumerate(clusters):
            rep = cluster[0]
            score = fuzz.token_sort_ratio(name, rep)
            if score >= threshold:
                cluster.append(name)
                cluster_ids[name] = idx
                assigned = True
                break
        if not assigned:
            clusters.append([name])
            cluster_ids[name] = current_id
            current_id += 1
    return cluster_ids, clusters

try:
    df = pd.read_excel(excel_path, usecols="A:D", dtype={"A": str, "B": str})
    df.columns = ["New Customer", "Salesrep", "Ignore", "Last Invoice Date"]
    df = df.dropna(subset=["New Customer", "Salesrep"])
    df = df[df["Salesrep"].str.strip().str.lower() != "house account"]
    df["Last Invoice Date"] = pd.to_datetime(df["Last Invoice Date"], errors="coerce")

    df["Cleaned Customer"] = df["New Customer"].apply(clean_customer_name)

    unique_names = df["Cleaned Customer"].unique()
    cluster_ids, clusters = assign_clusters(unique_names, threshold=85)

    # Map cluster ID to each cleaned customer
    df["Cluster ID"] = df["Cleaned Customer"].map(cluster_ids)

    # Optional: show clusters (for debugging)
    # st.write("Clusters found:")
    # for i, cluster in enumerate(clusters):
    #     st.write(f"Cluster {i}: {cluster}")

    # Aggregate to latest invoice date per salesrep per cluster
    agg_df = (
        df.groupby(["Salesrep", "Cluster ID"])
        .agg({"Last Invoice Date": "max", "New Customer": "first"})
        .reset_index()
    )

    leaderboard = (
        agg_df.groupby("Salesrep")["Cluster ID"]
        .nunique()
        .reset_index()
        .rename(columns={"Cluster ID": "Number of New Customers"})
        .sort_values("Number of New Customers", ascending=False)
        .reset_index(drop=True)
    )

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

    # Pending customers without invoice date
    pending = agg_df[agg_df["Last Invoice Date"].isna()]
    st.markdown("<h2>⏲ Pending Customers</h2>", unsafe_allow_html=True)

    if not pending.empty:
        for salesrep, group_df in pending.groupby("Salesrep"):
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

    # Show last updated time in Central Time zone
    from pytz import timezone
    central = timezone('US/Central')
    now_central = datetime.now(central)
    st.markdown(
        f"<div style='text-align: center; margin-top: 30px; color: gray;'>Last updated: {now_central.strftime('%B %d, %Y at %I:%M %p %Z')}</div>",
        unsafe_allow_html=True
    )

except FileNotFoundError:
    st.error(f"File not found: {excel_path}")
except Exception as e:
    st.error(f"An error occurred: {e}")

st.markdown('</div>', unsafe_allow_html=True)
