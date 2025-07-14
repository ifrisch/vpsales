import streamlit as st
import pandas as pd
import base64
import os
from datetime import datetime
from fuzzywuzzy import fuzz
from st_aggrid import AgGrid, GridOptionsBuilder
import re
from pytz import timezone

# CSS styling as before
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

# Logo block
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
    name = re.sub(r'[#\d]+', '', name)
    common_words = ['grill', 'restaurant', 'inc', 'llc', 'the', 'and', 'co', 'company', 'corp', 'corporation']
    for w in common_words:
        name = re.sub(r'\b' + re.escape(w) + r'\b', '', name)
    name = re.sub(r'[^a-z\s]', '', name)
    name = re.sub(r'\s+', ' ', name).strip()
    return name

def assign_clusters(names, threshold=80):
    clusters = []
    cluster_ids = {}
    current_id = 0
    for name in names:
        assigned = False
        for idx, cluster in enumerate(clusters):
            if any(fuzz.token_sort_ratio(name, member) >= threshold for member in cluster):
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

    # Clean customer names
    df["Cleaned Customer"] = df["New Customer"].apply(clean_customer_name)

    # Show raw vs cleaned for debugging
    st.markdown("### Raw vs Cleaned Customer Names (Debug)")
    for idx, row in df.iterrows():
        st.write(f"Salesrep: {row['Salesrep']} | Raw: {row['New Customer']} | Cleaned: {row['Cleaned Customer']}")

    # Cluster customers within each salesrep
    cluster_results = []
    for salesrep, group_df in df.groupby("Salesrep"):
        unique_names = group_df["Cleaned Customer"].unique()
        cluster_ids, clusters = assign_clusters(unique_names, threshold=80)
        # Map back cluster ids
        group_df = group_df.copy()
        group_df["Cluster ID"] = group_df["Cleaned Customer"].map(cluster_ids)
        cluster_results.append(group_df)

        st.markdown(f"### Clusters for Salesrep: {salesrep}")
        for i, cluster in enumerate(clusters):
            st.write(f"Cluster {i} ({len(cluster)} names): {cluster}")

    # Combine all clustered data back
    df_clustered = pd.concat(cluster_results)

    # Aggregate to count unique clusters per salesrep
    leaderboard = (
        df_clustered.groupby("Salesrep")["Cluster ID"]
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

    # Pending customers
    pending = df_clustered[df_clustered["Last Invoice Date"].isna()]
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

    # Show current central time as last updated
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
