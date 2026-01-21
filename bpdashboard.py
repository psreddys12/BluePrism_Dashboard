import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
from io import BytesIO
from datetime import datetime

# ---------------- Page Config ---------------- #
st.set_page_config(
    page_title="RPA Metrics Dashboard",
    page_icon="🖥️",
    layout="wide"
)

# ---------------- Custom CSS ---------------- #
st.markdown("""
<style>
.main { background-color: #f8f9f9; }
.stMetric {
    background-color: lightblue;
    padding: 10px;
    border-radius: 18px;
    text-align: center;
}
h1 { color: #1e3a8a; text-align: center; }
h2 { color: #3b82f6; text-align: center; }
</style>
""", unsafe_allow_html=True)

# ---------------- Load MAIN RPA Data ---------------- #
@st.cache_data
def load_rpa_data():
    df = pd.read_excel("RPA_Metrics_Dashboard.xlsx")

    df['Date'] = pd.to_datetime(
        df['Run_Year'].astype(str) + "-" + df['Run_Month'].astype(str) + "-01"
    )
    df['Quarter'] = df['Date'].dt.quarter
    df['Year_Quarter'] = df['Run_Year'].astype(str) + " Q" + df['Quarter'].astype(str)
    df['Year_Month'] = df['Run_Year'].astype(str) + "-" + df['Month']

    df['Cost_Savings_Dollars'] = pd.to_numeric(df['Cost_Savings_Dollars'], errors='coerce')
    df['Manual_Hours_Saved'] = pd.to_numeric(df['Manual_Hours_Saved'], errors='coerce')

    return df

# ---------------- Load FUNCTIONAL AREA Data ---------------- #
@st.cache_data
def load_functional_area_data():
    fa_df = pd.read_excel("Automation Savings New.xlsx")

    fa_df['Year'] = pd.to_numeric(fa_df['Year'], errors='coerce')
    fa_df['Cost_Savings_Dollars'] = pd.to_numeric(
        fa_df['Cost_Savings_Dollars'], errors='coerce'
    )

    return fa_df

df = load_rpa_data()
fa_df = load_functional_area_data()

# ---------------- Sidebar Filters ---------------- #
st.sidebar.title("🎯 Dashboard Filters")

years = sorted(df['Run_Year'].unique())
selected_years = st.sidebar.multiselect(
    "📅 Select Year(s)", years, default=years
)

business_areas = ['All'] + sorted(df['Business_Area'].unique())
selected_business_area = st.sidebar.selectbox("🏢 Business Area", business_areas)

# ---------------- Apply Filters ---------------- #
filtered_df = df.copy()

if selected_years:
    filtered_df = filtered_df[filtered_df['Run_Year'].isin(selected_years)]

if selected_business_area != 'All':
    filtered_df = filtered_df[filtered_df['Business_Area'] == selected_business_area]

# ---------------- KPI Section ---------------- #
st.title("🖥️ RPA Automation Metrics Dashboard")
st.markdown("### Executive Performance Overview")

c1, c2, c3, c4 = st.columns(4)

total_exec = filtered_df['Total_Executions'].sum()
total_hours = filtered_df['Manual_Hours_Saved'].sum()
total_savings = filtered_df['Cost_Savings_Dollars'].sum()
success_rate = (
    filtered_df['Successful_Executions'].sum() / total_exec * 100
    if total_exec > 0 else 0
)

c1.metric("📊 Total Executions", f"{total_exec:,}")
c2.metric("⏱️ Hours Saved", f"{total_hours:,.0f}")
c3.metric("💰 Cost Savings", f"${total_savings:,.0f}")
c4.metric("✅ Success Rate", f"{success_rate:.1f}%")

st.markdown("---")

# ---------------- Trend Chart ---------------- #
trend_df = filtered_df.groupby('Year_Month', as_index=False).agg({
    'Total_Executions': 'sum',
    'Cost_Savings_Dollars': 'sum'
})

fig_trend = px.line(
    trend_df,
    x='Year_Month',
    y='Cost_Savings_Dollars',
    title="Cost Savings Trend",
    markers=True,
    template="plotly_white"
)
st.plotly_chart(fig_trend, use_container_width=True)

st.markdown("---")

# ---------------- Existing Pie Charts ---------------- #
st.markdown("## 🥧 Distribution Analysis")
col1, col2 = st.columns(2)

with col1:
    pie_exec = filtered_df.groupby('Business_Area', as_index=False)['Total_Executions'].sum()
    fig_exec = px.pie(
        pie_exec,
        names='Business_Area',
        values='Total_Executions',
        title="Executions by Business Area"
    )
    st.plotly_chart(fig_exec, use_container_width=True)

with col2:
    pie_app = filtered_df.groupby('Application', as_index=False)['Cost_Savings_Dollars'].sum()
    pie_app = pie_app.sort_values('Cost_Savings_Dollars', ascending=False).head(8)
    fig_app = px.pie(
        pie_app,
        names='Application',
        values='Cost_Savings_Dollars',
        title="Cost Savings by Application"
    )
    st.plotly_chart(fig_app, use_container_width=True)

# ======================================================
# 🔥 NEW SECTION – FUNCTIONAL AREA PIE (SECOND EXCEL)
# ======================================================
st.markdown("---")
st.markdown("## 🥧 Functional Area Cost Savings")

fa_years = sorted(fa_df['Year'].dropna().unique())
selected_fa_year = st.selectbox(
    "Select Year (Functional Area Savings)",
    fa_years,
    index=len(fa_years) - 1
)

fa_filtered = fa_df[fa_df['Year'] == selected_fa_year]

fa_pie_data = (
    fa_filtered
    .groupby('Functional_Area', as_index=False)['Cost_Savings_Dollars']
    .sum()
)

fig_fa = px.pie(
    fa_pie_data,
    names='Functional_Area',
    values='Cost_Savings_Dollars',
    title=f"Cost Savings by Functional Area – {selected_fa_year}",
    template='plotly_white',
    color_discrete_sequence=px.colors.qualitative.Bold
)

fig_fa.update_traces(
    textinfo='percent+label',
    hovertemplate="<b>%{label}</b><br>$%{value:,.0f}<extra></extra>"
)

st.plotly_chart(fig_fa, use_container_width=True)

# ---------------- Export ---------------- #
st.markdown("---")
st.markdown("### 📥 Export Filtered Data")

csv_data = filtered_df.to_csv(index=False).encode("utf-8")

excel_buffer = BytesIO()
with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
    filtered_df.to_excel(writer, index=False)

st.download_button(
    "⬇️ Download CSV",
    csv_data,
    file_name="rpa_metrics.csv",
    mime="text/csv"
)

st.download_button(
    "⬇️ Download Excel",
    excel_buffer.getvalue(),
    file_name="rpa_metrics.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.markdown("""
<div style='text-align:center; color:#6b7280; padding:20px'>
RPA Metrics Dashboard | Streamlit & Plotly
</div>
""", unsafe_allow_html=True)
