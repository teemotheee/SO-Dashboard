import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import pytz
from datetime import datetime, timedelta
import os
from streamlit_autorefresh import st_autorefresh
st_autorefresh(interval=60000)
# =========================================================
# PARAMETERS (GitHub-Deploy Friendly)
# =========================================================
# Repo folder where app and Excel files are located
folder_path = os.path.dirname(os.path.abspath(__file__))

col_start = 2  # column C
col_end = 25
rounder = lambda x: round(x, 2) if isinstance(x, (int, float)) else x

# Hour labels for 24 hours
hour_labels = [f"{str(h).zfill(2)}00H" for h in range(24)]

# ========================================================= 
# STREAMLIT PAGE CONFIG
# =========================================================
st.set_page_config(layout="wide", page_title="Palawan Grid Dashboard")


# =========================================================
# HELPER FUNCTION TO LOAD EXCEL FILE
# =========================================================
# @st.cache_data
def load_month_file(month_name):
    excel_files = [f for f in os.listdir(folder_path)
                   if month_name.lower() in f.lower() and f.endswith((".xls", ".xlsx", ".xlsm"))]

    if not excel_files:
        st.warning(f"No Excel file found for {month_name}")
        return None, None, None

    file_mod_times = {f: os.path.getmtime(os.path.join(folder_path, f)) for f in excel_files}
    latest_file = max(file_mod_times, key=file_mod_times.get)
    file_path = os.path.join(folder_path, latest_file)

    df_full = pd.read_excel(file_path, sheet_name=0, header=None)

    df_freq = df_full.iloc[2:33, col_start:col_end+1].map(rounder)
    df_demand = df_full.iloc[34:65, col_start:col_end+1].map(rounder)

    # Replace zero values with NaN → they will not plot in charts
    df_freq = df_freq.replace(0, np.nan)
    df_demand = df_demand.replace(0, np.nan)

    df_freq.columns = hour_labels
    df_demand.columns = hour_labels

    
    return df_freq, df_demand, latest_file

# @st.cache_data
def load_online_units(latest_file):
    file_path = os.path.join(folder_path, latest_file)

    df_log = pd.read_excel(file_path, sheet_name="GEN SWITCHING LOGS", header=None)

    # Filter rows where columns D(3), E(4), G(6) are empty
    df_online = df_log[
        df_log[3].isna() &
        df_log[4].isna() &
        df_log[6].isna()
    ]

    # Keep only Unit Name (Column A)
    df_online_units = df_online[[0]].copy()
    df_online_units.columns = ["Unit"]

    return df_online_units


# =========================================================
# SELECT VIEW
# =========================================================
page = st.sidebar.selectbox("Select View", ["Current Day", "Full Month"])

# =========================================================
# CURRENT DATE AND MONTH
# =========================================================
today = datetime.today()
current_month_name = today.strftime("%B")
df_freq, df_demand, latest_file = load_month_file(current_month_name)

# Load and show online units in sidebar immediately
df_online_units = load_online_units(latest_file)
units = df_online_units["Unit"].tolist()  # just the unit names

# Build HTML table with navy blue background and white text, centered
table_html = '''
<div style="max-height:400px; overflow-y:auto; background-color:#001f4d; border-radius:5px; padding:5px;">
    <table style="border-collapse: collapse; width: 100%; text-align: center; color: white;">
'''
for unit in units:
    table_html += f'<tr><td style="padding: 8px; border-bottom: 1px solid #004080;">{unit}</td></tr>'
table_html += '</table></div>'

st.sidebar.subheader("Online Units")
st.sidebar.markdown(table_html, unsafe_allow_html=True)



if df_freq is None:
    st.stop()

# =========================================================
# CURRENT DAY VIEW
# =========================================================
if page == "Current Day":
    st.set_page_config(layout="wide", page_title="Palawan Grid Dashboard")

    # Fixed header styling + glow animation for latest point
    st.markdown("""
    <style>
    .fixed-header {
        position: fixed;
        top: 0;
        left: 0;
        width: 100%;
        background-color: #001f4d;
        color: white;
        font-size: 32px;
        font-weight: bold;
        text-align: left;
        line-height: 1.2;
        padding-top: 60px;
        padding-bottom: 15px;
        padding-left: 350px;
        z-index: 9999;
    }
    .main-content {
        padding-top: 30px;
    }

    /* Glow pulse animation for latest point */
    @keyframes glowPulse {
        0%   { transform: scale(1); opacity: 0.6; box-shadow: 0 0 0px red; }
        50%  { transform: scale(1.5); opacity: 1; box-shadow: 0 0 15px red; }
        100% { transform: scale(1); opacity: 0.6; box-shadow: 0 0 0px red; }
    }

    /* Applies the glow to HTML span elements with class 'glow-point' */
    .glow-point {
        display: inline-block;
        color: red;
        font-weight: bold;
        font-size: 20px;
        animation: glowPulse 5s infinite;
    }
    </style>
    """, unsafe_allow_html=True)

    # Header
    st.markdown('<div class="fixed-header">IGSOD PCC Dashboard</div>', unsafe_allow_html=True)
    st.markdown('<div class="main-content">', unsafe_allow_html=True)

    # Get today's data
    row_idx = today.day - 1
    df_freq_today = df_freq.iloc[[row_idx]].copy().reset_index(drop=True)
    df_demand_today = df_demand.iloc[[row_idx]].copy().reset_index(drop=True)

    # Previous day data
    prev_day = today - timedelta(days=1)
    prev_month_name = prev_day.strftime("%B")
    if prev_month_name == current_month_name:
        prev_row_idx = prev_day.day - 1
        df_demand_prev = df_demand.iloc[[prev_row_idx]].copy().reset_index(drop=True)
    else:
        df_freq_prev_dummy, df_demand_prev, _ = load_month_file(prev_month_name)
        prev_row_idx = prev_day.day - 1
        df_demand_prev = df_demand_prev.iloc[[prev_row_idx]].copy().reset_index(drop=True)

    # Demand Table
    st.subheader("Demand – " + today.strftime("%Y-%m-%d"))
    st.dataframe(df_demand_today, hide_index=True, use_container_width=True)

    # Current hour PH time
    ph_time = datetime.now(pytz.timezone("Asia/Manila"))
    current_hour = ph_time.hour + 1  # include current hour

    # Prepare demand data
    y_today = df_demand_today.values.flatten().copy()
    y_today[current_hour:] = np.nan
    y_prev = df_demand_prev.values.flatten().copy()

    # Demand Plot
    fig, ax = plt.subplots(figsize=(12, 3))
    ax.plot(hour_labels, y_prev, marker="x", linestyle="--", alpha=0.4, label="Previous Day")
    ax.plot(hour_labels, y_today, marker="o", label="Today")

    # Latest hour index
    latest_hour = current_hour - 1
    if latest_hour >= 0:
        latest_x = hour_labels[latest_hour]
        latest_y = y_today[latest_hour]
        # Large red marker
        ax.plot(latest_x, latest_y, marker="o", markersize=14,
                markerfacecolor="red", markeredgecolor="red")

        # Overlay HTML glow using Streamlit
        st.markdown(f"<span class='glow-point'>⬤</span>", unsafe_allow_html=True)

    ax.set_xlabel("Hour")
    ax.set_ylabel("Demand (MW)")
    ax.set_title("Demand (MW)")
    ax.grid(True)
    ax.legend()
    plt.xticks(rotation=45)
    st.pyplot(fig)

    # Frequency Table
    st.subheader("Frequency – Current Day")
    st.dataframe(df_freq_today, hide_index=True, use_container_width=True)

    # Prepare frequency data
    y_freq_today = df_freq_today.values.flatten().copy()
    y_freq_today[y_freq_today == 0] = np.nan
    y_freq_plot = np.full(24, np.nan)
    y_freq_plot[:len(y_freq_today)] = y_freq_today

    # Frequency Plot
    fig2, ax2 = plt.subplots(figsize=(12, 3))
    ax2.plot(range(1, 25), y_freq_plot, marker="o", label="Today")

    if latest_hour >= 0 and not np.isnan(y_freq_plot[latest_hour]):
        ax2.plot(latest_hour + 1, y_freq_plot[latest_hour],
                 marker="o", markersize=14,
                 markerfacecolor="red", markeredgecolor="red")
        # HTML glow overlay
        st.markdown(f"<span class='glow-point'>⬤</span>", unsafe_allow_html=True)

    ax2.set_xticks(range(1, 25))
    ax2.set_xticklabels([f"{str(h-1).zfill(2)}00H" for h in range(1, 25)])
    ax2.set_xlabel("Hour")
    ax2.set_ylabel("Frequency (Hz)")
    ax2.set_title("Frequency (Hz)")
    ax2.set_ylim(50, 62)
    ax2.grid(True)
    ax2.legend()
    plt.xticks(rotation=45)
    st.pyplot(fig2)



# =========================================================
# FULL MONTH VIEW
# =========================================================
else:
    st.title(f"Full Month – {current_month_name} Dashboard")

    st.subheader("Full Month – Demand")
    st.dataframe(df_demand, hide_index=True, use_container_width=True)

    st.subheader("Full Month – Frequency")
    st.dataframe(df_freq, hide_index=True, use_container_width=True)
