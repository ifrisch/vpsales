import streamlit as st
import pandas as pd
from PIL import Image
import base64
from datetime import datetime
from zoneinfo import ZoneInfo  # For Central Time
from fuzzywuzzy import fuzz
from st_aggrid import AgGrid, GridOptionsBuilder
import time

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

    /* Global font family */
    .stApp, .stMarkdown, .stSelectbox, .stTabs, .stExpander {
        font-family: 'Futura', 'Trebuchet MS', 'Arial', sans-serif !important;
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

# --- OVERVIEW SECTION ---
st.markdown("""
<div style='
    background-color: #F8F9FA; 
    border-left: 4px solid #6C757D; 
    padding: 20px; 
    margin: 20px 0; 
    border-radius: 5px;
    font-family: Futura, sans-serif;
'>
    <h3 style='margin-top: 0; color: #495057;'>Incentive Overview</h3>
    <ul style='margin-bottom: 15px;'>
        <li><strong>$50 Bonus:</strong> For every rep who secures 3+ new accounts</li>
        <li><strong>$100 Top Performer:</strong> Additional prize for the rep with the most new accounts</li>
    </ul>
    <div style='font-size: 11px; color: #999; border-top: 1px solid #ddd; padding-top: 12px; margin-top: 12px; line-height: 1.4;'>
        <strong>Note:</strong> New ownership, new management, and name changes do not count as new accounts. Accounts must have at least 1 order to qualify. Additional ship-tos or separate customers associated with the same primary business do not count as additional customers.
    </div>
    <div style='margin-top: 15px; text-align: center; font-weight: bold; color: #666;'>
        📅 Incentive Period: June 27th - September 19th
    </div>
</div>
""", unsafe_allow_html=True)

# --- LEADERBOARD SECTION ---
st.markdown("<h3 style='margin-bottom: 0.5rem; color: #333; font-family: Futura, sans-serif;'>🏆 Current Standings</h3>", unsafe_allow_html=True)

# --- LOAD DATA ---
excel_path = "leaderboard_new.xlsx"  # Using fresh Van Paper data from 8:55 AM email

try:
    # Read the Excel file with the correct column names
    df = pd.read_excel(excel_path, usecols="A:E", dtype={"A": str, "B": str, "E": str})
    
    # Use the actual column names from your Excel file
    df.columns = ["Customer Name", "Salesperson", "Prospect", "Last Invoice Date", "Customer Number"]
    
    # Rename to match our internal naming convention
    df = df.rename(columns={
        "Customer Name": "New Customer",
        "Salesperson": "Salesrep",
        "Customer Number": "Customer Number",
        "Prospect": "Rule Violation"
    })
    
    df = df.dropna(subset=["New Customer", "Salesrep"])
    
    if len(df) == 0:
        st.error("No valid data found after removing empty rows")
        st.stop()
    
    df = df[df["Salesrep"].str.strip().str.lower() != "house account"]
    
    df["Last Invoice Date"] = pd.to_datetime(df["Last Invoice Date"], errors="coerce")
    
    # Convert Rule Violation to string and handle NaN values  
    df["Rule Violation"] = df["Rule Violation"].astype(str)
    df["Rule Violation"] = df["Rule Violation"].replace("nan", "")

    # Clean customer names
    df["Cleaned Customer"] = df["New Customer"].str.lower()
    df["Cleaned Customer"] = df["Cleaned Customer"].str.replace(r'[^\w\s]', '', regex=True)
    df["Cleaned Customer"] = df["Cleaned Customer"].str.replace(r'\s+', ' ', regex=True).str.strip()

    used_customers = set()
    kept_rows = []
    pending_rows = []
    violation_rows = []

    # We'll iterate through all rows, grouping customers by fuzzy matches (token_set_ratio) >= 85
    for i, row in df.iterrows():
        cust_name = row["Cleaned Customer"]
        if cust_name in used_customers:
            continue

        # Find all rows with fuzzy token_set_ratio >= 85
        matches = df[df["Cleaned Customer"].apply(lambda x: fuzz.token_set_ratio(x, cust_name) >= 85)].copy()

        # Mark all matched cleaned customers as used
        used_customers.update(matches["Cleaned Customer"].tolist())

        # Check if any matches have rule violations
        # For now, treat "Prospect" values as valid customers, not violations
        # Real violations would be specific text like "Duplicate", "Invalid", etc.
        matches_with_violations = matches[
            (matches["Rule Violation"] != "") & 
            (matches["Rule Violation"] != "nan") & 
            (matches["Rule Violation"] != "Prospect") &
            (matches["Rule Violation"].str.contains("violation|duplicate|invalid|exclude", case=False, na=False))
        ]
        if not matches_with_violations.empty:
            # If there's a rule violation, add to violation list
            best_violation = matches_with_violations.iloc[0]
            violation_rows.append(best_violation)
        else:
            # If no explicit violations, process normally
            # If any matched rows have an invoice date, pick the latest one for keeping
            matches_with_invoice = matches[~matches["Last Invoice Date"].isna()]
            if not matches_with_invoice.empty:
                best_match = matches_with_invoice.sort_values(by="Last Invoice Date", ascending=False).iloc[0]
                kept_rows.append(best_match)
                
                # Add any remaining matches as duplicates/violations
                remaining_matches = matches[matches.index != best_match.name]
                for _, duplicate_row in remaining_matches.iterrows():
                    violation_rows.append(duplicate_row)
            else:
                # If none have invoice dates, just take the first match row
                best_match = matches.iloc[0]
                pending_rows.append(best_match)
                
                # Add any remaining matches as duplicates/violations
                remaining_matches = matches[matches.index != best_match.name]
                for _, duplicate_row in remaining_matches.iterrows():
                    violation_rows.append(duplicate_row)

    df_cleaned = pd.DataFrame(kept_rows)
    df_pending = pd.DataFrame(pending_rows)
    df_violations = pd.DataFrame(violation_rows)

    if len(df_cleaned) == 0:
        st.warning("No customers with invoices found for leaderboard")
        leaderboard = pd.DataFrame(columns=["Rank", "Salesrep", "Number of New Customers", "Prize"])
        max_customers = 0
    else:
        leaderboard = df_cleaned.groupby("Salesrep")["New Customer"].nunique().reset_index()
        leaderboard = leaderboard.rename(columns={"New Customer": "Number of New Customers"})
        leaderboard = leaderboard.sort_values(by="Number of New Customers", ascending=False).reset_index(drop=True)

        # Calculate prizes
        max_customers = leaderboard["Number of New Customers"].max()
        first_place_winners = leaderboard[leaderboard["Number of New Customers"] == max_customers]
        num_first_place = len(first_place_winners)

        # Prize per first place winner (split $100 among ties)
        first_place_prize_each = 100 / num_first_place if num_first_place > 0 else 0

        def calc_prize(row):
            prize = 0
            if row["Number of New Customers"] >= 3:
                prize += 50
            if row["Number of New Customers"] == max_customers:
                prize += first_place_prize_each
            return prize

        leaderboard["Prize"] = leaderboard.apply(calc_prize, axis=1)

        # Format Prize column with $ symbol and no decimals if whole number
        leaderboard["Prize"] = leaderboard["Prize"].apply(lambda x: f"${int(x)}" if x.is_integer() else f"${x:.2f}")

        # Create rank labels with ties
        ranks_numeric = leaderboard["Number of New Customers"].rank(method='min', ascending=False).astype(int)
        suffixes = {1: "st", 2: "nd", 3: "rd"}

        def rank_label(n):
            if 10 <= n % 100 <= 20:
                return f"{n}th"
            return f"{n}{suffixes.get(n % 10, 'th')}"

        ranks = ranks_numeric.apply(rank_label)
        
        # Only insert Rank column if it doesn't already exist
        if "Rank" not in leaderboard.columns:
            leaderboard.insert(0, "Rank", ranks)

    # Streamlined Leaderboard Display
    if len(leaderboard) > 0:
        for i, row in leaderboard.iterrows():
            rank = row["Rank"]
            salesrep = row["Salesrep"]
            customers = row["Number of New Customers"]
            prize = row["Prize"]
            
            # Special styling for first place
            is_first_place = customers == max_customers
            
            if is_first_place:
                # First place gets special styling
                if rank == "1st":
                    emoji = "🥇"
                    name_color = "#DAA520"
                elif rank == "2nd":
                    emoji = "🥈"
                    name_color = "#C0C0C0"
                elif rank == "3rd":
                    emoji = "🥉"
                    name_color = "#CD7F32"
                else:
                    emoji = "🏆"
                    name_color = "#DAA520"
                name_weight = "bold"
            else:
                emoji = ""
                name_color = "#333"
                name_weight = "normal"
            
            # Create compact row
            st.markdown(f"""
            <div style="
                display: flex; 
                justify-content: space-between; 
                align-items: center;
                padding: 8px 12px;
                margin: 4px 0;
                background-color: {'#FFF9E6' if is_first_place else '#FAFAFA'};
                border-left: 4px solid {'#FFD700' if is_first_place else '#E0E0E0'};
                border-radius: 4px;
            ">
                <div style="display: flex; align-items: center; flex: 1;">
                    <span style="font-size: 16px; margin-right: 8px; width: 20px;">{emoji}</span>
                    <span style="font-size: 16px; font-weight: bold; color: #666; margin-right: 12px; min-width: 30px;">{rank}</span>
                    <span style="font-size: 18px; font-weight: {name_weight}; color: {name_color};">{salesrep}</span>
                </div>
                <div style="display: flex; align-items: center; gap: 20px;">
                    <div style="text-align: center;">
                        <span style="font-size: 18px; font-weight: bold; color: #2E8B57;">{customers}</span>
                        <span style="font-size: 12px; color: #666; margin-left: 4px;">customers</span>
                    </div>
                    <div style="text-align: right; min-width: 60px;">
                        <span style="font-size: 16px; font-weight: bold; color: #228B22;">{prize}</span>
                    </div>
                </div>
            </div>
            """, unsafe_allow_html=True)

    # Add spacing between leaderboard and customer details
    st.markdown("<br><br>", unsafe_allow_html=True)
    st.markdown("---")
    st.markdown("<br>", unsafe_allow_html=True)

    # --- TABBED DATA SECTION ---
    tab1, tab2, tab3 = st.tabs(["🏆 New Customers", "⏲ Pending Customers", "❌ Rule Violations"])
    
    with tab1:
        st.markdown("### Customers Counted Toward New Customer Goals")
        
        if not df_cleaned.empty:
            for salesrep, group_df in df_cleaned.groupby("Salesrep"):
                with st.expander(f"**{salesrep}** ({len(group_df)} customers)", expanded=False):
                    for _, row in group_df.iterrows():
                        customer_num = row["Customer Number"] if pd.notna(row["Customer Number"]) else "N/A"
                        customer_display = f"{row['New Customer']} ({customer_num})"
                        invoice_date = row["Last Invoice Date"]
                        if pd.notna(invoice_date):
                            date_str = invoice_date.strftime("%m/%d/%Y")
                            st.markdown(f"• **{customer_display}** - *Invoice: {date_str}*")
                        else:
                            st.markdown(f"• **{customer_display}**")
        else:
            st.info("No new customers found.")
    
    with tab2:
        st.markdown("### Customers Not Yet Counted")
        
        if not df_pending.empty:
            for salesrep, group_df in df_pending.groupby("Salesrep"):
                with st.expander(f"**{salesrep}** ({len(group_df)} customers)", expanded=False):
                    for _, row in group_df.iterrows():
                        customer_num = row["Customer Number"] if pd.notna(row["Customer Number"]) else "N/A"
                        customer_display = f"{row['New Customer']} ({customer_num})"
                        st.markdown(f"• **{customer_display}** - *Awaiting first invoice*")
        else:
            st.info("No pending customers! 🎉")
    
    with tab3:
        st.markdown("### Customers Excluded Due to Rule Violations")
        
        if not df_violations.empty:
            for salesrep, group_df in df_violations.groupby("Salesrep"):
                with st.expander(f"**{salesrep}** ({len(group_df)} customers)", expanded=False):
                    for _, row in group_df.iterrows():
                        customer_num = row["Customer Number"] if pd.notna(row["Customer Number"]) else "N/A"
                        customer_display = f"{row['New Customer']} ({customer_num})"
                        # Just show the customer name and number, no violation reason
                        st.markdown(f"• **{customer_display}**")
        else:
            st.info("No rule violations found! ✅")

except FileNotFoundError:
    st.error(f"File not found: {excel_path}")
except Exception as e:
    st.error(f"An error occurred: {e}")

# --- Close MAIN BLOCK ---
st.markdown('</div>', unsafe_allow_html=True)

# --- TIMESTAMP ---
central = ZoneInfo("America/Chicago")
LAST_SYNC_TIMESTAMP = "2025-08-28 13:51:16"  # AUTO-UPDATED BY BATCH FILE

# Display sync timestamp
sync_time = datetime.strptime(LAST_SYNC_TIMESTAMP, '%Y-%m-%d %H:%M:%S')
last_updated = sync_time.replace(tzinfo=central)

st.markdown(
    f"<div style='text-align: center; margin-top: 30px; color: gray; font-family: Futura, sans-serif;'>App last synced: {last_updated.strftime('%B %d, %Y at %I:%M %p')}</div>",
    unsafe_allow_html=True
)
