import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
from io import BytesIO
from datetime import datetime

# Page configuration
st.set_page_config(
    page_title="RPA Metrics Dashboard",
    page_icon="🖥️",
    layout="wide",
    initial_sidebar_state="auto"
    
)

# Custom CSS for dashboard appearance
st.markdown("""
    <style>
    .main {
        background-color: #f8f9fg;
    }
    .stMetric {
        background-color: lightblue;
        padding: 10px;
        border-radius: 18px;
        text-align: center;
        box-shadow: 0 2px 2px rgba(0,0,0,0.1);
    }
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 25px;
        border-radius: 30px;
        color: white;
        text-align: center;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    h1 {
        color: #1e3a8a;
        font-weight: 700;
        text-align: center;
        padding: 20px 0;
    }
    h2 {
        color: #3b82f6;
        font-weight: 600;
        margin-top: 30px;
        text-align: center;
    }
    h3 {
        color: #60a5fa;
        font-weight: 500;
        text-align: center;
    }
    </style>
""", unsafe_allow_html=True)

# Load data
@st.cache_data
def load_data():
    df = pd.read_excel('RPA_Metrics_Dashboard.xlsx')
    # GitHub Release download URL
    # url = 'https://github.com/sarkarbikram90/RPA-file-path/releases/download/V1/RPA_Metrics_Dashboard.xlsx'
    
    # df = pd.read_excel(url)
    
            
    # Data cleaning and preparation
    df['Date'] = pd.to_datetime(df['Run_Year'].astype(str) + '-' + df['Run_Month'].astype(str) + '-01')
    df['Quarter'] = df['Date'].dt.quarter
    df['Year_Quarter'] = df['Run_Year'].astype(str) + ' Q' + df['Quarter'].astype(str)
    df['Year_Month'] = df['Run_Year'].astype(str) + '-' + df['Month']
    df['Week'] = df['Date'].dt.isocalendar().week
    
    # Clean numeric columns
    df['Cost_Savings_Dollars'] = pd.to_numeric(df['Cost_Savings_Dollars'], errors='coerce')
    df['Manual_Hours_Saved'] = pd.to_numeric(df['Manual_Hours_Saved'], errors='coerce')
    
    return df

df = load_data()
# df = load_data_from_sql()

# Sidebar Filters
# st.sidebar.image("🖥️", width=80)
st.sidebar.title("🎯 Dashboard Filters")

# Time period selector
time_period = st.sidebar.radio(
    "📊 Time Aggregation",
    ["Daily", "Weekly", "Monthly", "Quarterly", "Yearly"],
    index=0
)

# Year filter
all_years = sorted(df['Run_Year'].unique())
selected_years = st.sidebar.multiselect(
    "📅 Select Year(s)",
    options=all_years,
    default=all_years
)

# Month filter
if selected_years:
    available_months = sorted(df[df['Run_Year'].isin(selected_years)]['Month'].unique(),
                             key=lambda x: datetime.strptime(x, '%b').month)
    selected_months = st.sidebar.multiselect(
        "📆 Select Month(s)",
        options=available_months,
        default=available_months
    )
else:
    selected_months = []

# Business Area filter
business_areas = ['All'] + sorted(df['Business_Area'].unique().tolist())
selected_business_area = st.sidebar.selectbox(
    "🏢 Business Area",
    options=business_areas
)

# Process Name filter
if selected_business_area != 'All':
    process_names = ['All'] + sorted(df[df['Business_Area'] == selected_business_area]['Process_Name'].unique().tolist())
else:
    process_names = ['All'] + sorted(df['Process_Name'].unique().tolist())

selected_process = st.sidebar.selectbox(
    "⚙️ Process Name",
    options=process_names
)

# Machine Name filter
machine_names = ['All'] + sorted(df['Machine_Name'].unique().tolist())
selected_machine = st.sidebar.selectbox(
    "🖥️ Machine Name",
    options=machine_names
)

# Apply filters
filtered_df = df.copy()

if selected_years:
    filtered_df = filtered_df[filtered_df['Run_Year'].isin(selected_years)]

if selected_months:
    filtered_df = filtered_df[filtered_df['Month'].isin(selected_months)]

if selected_business_area != 'All':
    filtered_df = filtered_df[filtered_df['Business_Area'] == selected_business_area]

if selected_process != 'All':
    filtered_df = filtered_df[filtered_df['Process_Name'] == selected_process]

if selected_machine != 'All':
    filtered_df = filtered_df[filtered_df['Machine_Name'] == selected_machine]

# Main Dashboard
st.title("🖥️ RPA Automation Metrics Dashboard")
st.markdown("### Executive Performance Overview")

# Key Metrics Row
col1, col2, col3, col4, col5 = st.columns(5)

with col1:
    total_executions = filtered_df['Total_Executions'].sum()
    st.metric(
        label="📊 Total Executions",
        value=f"{total_executions:,}",
        # delta=f"{len(filtered_df)} records"
    )

with col2:
    total_hours_saved = filtered_df['Manual_Hours_Saved'].sum()
    st.metric(
        label="⏱️ Hours Saved",
        value=f"{total_hours_saved:,.0f}",
        # delta=f"{total_hours_saved/8760*100:.1f}% of year"
    )

with col3:
    total_cost_savings = filtered_df['Cost_Savings_Dollars'].sum()
    st.metric(
        label="💰 Cost Savings (USD)",
        value=f"${total_cost_savings:,.0f}",
        # delta="USD"
    )

with col4:
    success_rate = (filtered_df['Successful_Executions'].sum() / total_executions * 100) if total_executions > 0 else 0
    st.metric(
        label="✅ Success Rate",
        value=f"{success_rate:.1f}%",
        # delta=f"{filtered_df['Successful_Executions'].sum():,} successful"
    )

with col5:
    unique_processes = filtered_df['Process_Name'].nunique()
    st.metric(
        label="⚙️ Active Processes",
        value=f"{unique_processes}",
        # delta=f"{filtered_df['Machine_Name'].nunique()} machines"
    )

st.markdown("---")

# Aggregate data based on time period
def aggregate_data(df, period):
    if period == "Daily":
        group_by = 'Date'
    elif period == "Weekly":
        group_by = 'Week'
        df = df.groupby(['Run_Year', 'Week']).agg({
            'Total_Executions': 'sum',
            'Manual_Hours_Saved': 'sum',
            'Cost_Savings_Dollars': 'sum',
            'Successful_Executions': 'sum',
            'Exception_Executions': 'sum'
        }).reset_index()
        df['Period'] = 'W' + df['Week'].astype(str) + ' ' + df['Run_Year'].astype(str)
        return df
    elif period == "Monthly":
        group_by = 'Year_Month'
    elif period == "Quarterly":
        group_by = 'Year_Quarter'
    else:  # Yearly
        group_by = 'Run_Year'
    
    agg_df = df.groupby(group_by).agg({
        'Total_Executions': 'sum',
        'Manual_Hours_Saved': 'sum',
        'Cost_Savings_Dollars': 'sum',
        'Successful_Executions': 'sum',
        'Exception_Executions': 'sum'
    }).reset_index()
    
    agg_df.columns = ['Period', 'Total_Executions', 'Manual_Hours_Saved', 
                      'Cost_Savings_Dollars', 'Successful_Executions', 'Exception_Executions']
    return agg_df

time_series_data = aggregate_data(filtered_df, time_period)

# Row 1: Line Charts - Trend Analysis
st.markdown("## 📈 Trend Analysis")
col1, col2 = st.columns(2)

with col1:
    fig_executions = px.line(
        time_series_data,
        x='Period',
        y='Total_Executions',
        title=f'Total Executions Over Time ({time_period})',
        markers=True,
        template='plotly_white'
    )
    fig_executions.update_traces(line_color='#3b82f6', line_width=3, marker=dict(size=8))
    fig_executions.update_layout(
        title_font_size=18,
        title_font_color='#1e3a8a',
        height=400,
        hovermode='x unified'
    )
    st.plotly_chart(fig_executions, use_container_width=True)

with col2:
    fig_hours = px.line(
        time_series_data,
        x='Period',
        y='Manual_Hours_Saved',
        title=f'Manual Hours Saved Over Time ({time_period})',
        markers=True,
        template='plotly_white'
    )
    fig_hours.update_traces(line_color='#10b981', line_width=3, marker=dict(size=8))
    fig_hours.update_layout(
        title_font_size=18,
        title_font_color='#1e3a8a',
        height=400,
        hovermode='x unified'
    )
    st.plotly_chart(fig_hours, use_container_width=True)

# Row 2: Cost Savings Line Chart + Area Chart
col1, col2 = st.columns(2)

with col1:
    fig_cost = px.line(
        time_series_data,
        x='Period',
        y='Cost_Savings_Dollars',
        title=f'Cost Savings Trend ({time_period})',
        markers=True,
        template='plotly_white'
    )
    fig_cost.update_traces(line_color='#f59e0b', line_width=3, marker=dict(size=8))
    fig_cost.update_layout(
        title_font_size=18,
        title_font_color='#1e3a8a',
        height=400,
        hovermode='x unified'
    )
    fig_cost.update_yaxes(tickprefix='$')
    st.plotly_chart(fig_cost, use_container_width=True)

with col2:
    # Cumulative savings
    time_series_data_sorted = time_series_data.sort_values('Period')
    time_series_data_sorted['Cumulative_Savings'] = time_series_data_sorted['Cost_Savings_Dollars'].cumsum()
    
    fig_cumulative = px.area(
        time_series_data_sorted,
        x='Period',
        y='Cumulative_Savings',
        title=f'Cumulative Cost Savings ({time_period})',
        template='plotly_white'
    )
    fig_cumulative.update_traces(fillcolor='rgba(124, 58, 237, 0.3)', line_color='#7c3aed', line_width=3)
    fig_cumulative.update_layout(
        title_font_size=18,
        title_font_color='#1e3a8a',
        height=400,
        hovermode='x unified'
    )
    fig_cumulative.update_yaxes(tickprefix='$')
    st.plotly_chart(fig_cumulative, use_container_width=True)

st.markdown("---")

# Row 3: Bar Charts
st.markdown("## 📊 Performance Breakdown")
col1, col2 = st.columns(2)

with col1:
    # Top processes by executions
    top_processes = filtered_df.groupby('Process_Name').agg({
        'Total_Executions': 'sum'
    }).reset_index().sort_values('Total_Executions', ascending=False).head(10)
    
    fig_top_processes = px.bar(
        top_processes,
        x='Total_Executions',
        y='Process_Name',
        orientation='h',
        title='Top 10 Processes by Executions',
        template='plotly_white',
        color='Total_Executions',
        color_continuous_scale='Blues'
    )
    fig_top_processes.update_layout(
        title_font_size=18,
        title_font_color='#1e3a8a',
        height=450,
        showlegend=False,
        yaxis={'categoryorder': 'total ascending'}
    )
    st.plotly_chart(fig_top_processes, use_container_width=True)

with col2:
    # Hours saved by business area
    business_hours = filtered_df.groupby('Business_Area').agg({
        'Manual_Hours_Saved': 'sum'
    }).reset_index().sort_values('Manual_Hours_Saved', ascending=False)
    
    fig_business_hours = px.bar(
        business_hours,
        x='Business_Area',
        y='Manual_Hours_Saved',
        title='Hours Saved by Business Area',
        template='plotly_white',
        color='Manual_Hours_Saved',
        color_continuous_scale='Greens'
    )
    fig_business_hours.update_layout(
        title_font_size=18,
        title_font_color='#1e3a8a',
        height=450,
        showlegend=False
    )
    fig_business_hours.update_xaxes(tickangle=45)
    st.plotly_chart(fig_business_hours, use_container_width=True)

# Row 4: Column Chart and Stacked Bar
col1, col2 = st.columns(2)

with col1:
    # Monthly cost savings comparison
    monthly_savings = filtered_df.groupby('Month').agg({
        'Cost_Savings_Dollars': 'sum'
    }).reset_index()
    
    # Order months correctly
    month_order = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
    monthly_savings['Month'] = pd.Categorical(monthly_savings['Month'], categories=month_order, ordered=True)
    monthly_savings = monthly_savings.sort_values('Month')
    
    fig_monthly = go.Figure(data=[
        go.Bar(
            x=monthly_savings['Month'],
            y=monthly_savings['Cost_Savings_Dollars'],
            marker_color='rgba(99, 102, 241, 0.8)',
            text=monthly_savings['Cost_Savings_Dollars'].apply(lambda x: f'${x:,.0f}'),
            textposition='outside'
        )
    ])
    fig_monthly.update_layout(
        title='Cost Savings by Month',
        title_font_size=18,
        title_font_color='#1e3a8a',
        height=450,
        template='plotly_white',
        yaxis_tickprefix='$'
    )
    st.plotly_chart(fig_monthly, use_container_width=True)

with col2:
    # Success vs Exception by Process (Top 10)
    process_success = filtered_df.groupby('Process_Name').agg({
        'Successful_Executions': 'sum',
        'Exception_Executions': 'sum'
    }).reset_index()
    process_success['Total'] = process_success['Successful_Executions'] + process_success['Exception_Executions']
    process_success = process_success.sort_values('Total', ascending=False).head(10)
    
    fig_success = go.Figure(data=[
        go.Bar(name='Successful', x=process_success['Process_Name'], y=process_success['Successful_Executions'], 
               marker_color='#10b981'),
        go.Bar(name='Exceptions', x=process_success['Process_Name'], y=process_success['Exception_Executions'], 
               marker_color='#ef4444')
    ])
    fig_success.update_layout(
        title='Success vs Exceptions (Top 10 Processes)',
        title_font_size=18,
        title_font_color='#1e3a8a',
        height=450,
        template='plotly_white',
        barmode='stack',
        xaxis_tickangle=45
    )
    st.plotly_chart(fig_success, use_container_width=True)

st.markdown("---")

# Row 5: Pie Charts
st.markdown("## 🥧 Distribution Analysis")
col1, col2, col3 = st.columns(3)

# ---------------- Shared Layout Config ---------------- #
COMMON_PIE_LAYOUT = dict(
    height=400,
    margin=dict(l=20, r=20, t=60, b=90),
    legend=dict(
        orientation="h",
        yanchor="bottom",
        y=-0.35,
        xanchor="center",
        x=0.5
    )
)

COMMON_PIE_DOMAIN = dict(domain=dict(x=[0, 1], y=[0, 1]))

# ---------------- Column 1 ---------------- #
with col1:
    # Executions by Business Area
    business_exec = (
        filtered_df.groupby('Business_Area')['Total_Executions']
        .sum()
        .reset_index()
    )

    fig_pie_business = px.pie(
        business_exec,
        values='Total_Executions',
        names='Business_Area',
        title='Executions by Business Area',
        template='plotly_white',
        color_discrete_sequence=px.colors.qualitative.Set3
    )

    fig_pie_business.update_traces(
        textinfo='percent+label',
        textposition='inside',
        **COMMON_PIE_DOMAIN
    )

    fig_pie_business.update_layout(
        title_font_size=16,
        title_font_color='#1e3a8a',
        **COMMON_PIE_LAYOUT
    )

    st.plotly_chart(fig_pie_business, use_container_width=True)

# ---------------- Column 2 ---------------- #
with col2:
    # Cost Savings by Application (Top N – data-driven)
    TOP_N = 10

    app_savings = (
        filtered_df
        .groupby('Application', as_index=False)['Cost_Savings_Dollars']
        .sum()
        .query("Cost_Savings_Dollars > 0")  # remove zero / null contributors
        .sort_values('Cost_Savings_Dollars', ascending=False)
        .head(TOP_N)
    )

    actual_n = len(app_savings)

    fig_pie_app = px.pie(
        app_savings,
        values='Cost_Savings_Dollars',
        names='Application',
        title=f'Cost Savings by Application (Top {actual_n})',
        template='plotly_white',
        color_discrete_sequence=px.colors.qualitative.Pastel
    )

    fig_pie_app.update_traces(
        textinfo='percent',
        textposition='inside',
        textfont=dict(size=14),                 # enlarged text
        insidetextorientation='radial',
        pull=[0.05 if i == 0 else 0 for i in range(actual_n)],  # highlight dominant app
        hovertemplate=(
            "<b>%{label}</b><br>"
            "Cost Savings: $%{value:,.0f}<br>"
            "%{percent}<extra></extra>"
        ),
        domain=dict(x=[0, 1], y=[0, 1])
    )

    fig_pie_app.update_layout(
        title_font_size=16,
        title_font_color='#1e3a8a',
        showlegend=True,
        height=400,
        margin=dict(l=20, r=20, t=60, b=100),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=-0.35,
            xanchor="center",
            x=0.5
        )
    )

    st.plotly_chart(fig_pie_app, use_container_width=True)

# ---------------- Column 3 ---------------- #
with col3:
    # Hours Saved by Machine (Show All Machines)
    machine_hours = (
        filtered_df.groupby('Machine_Name')['Manual_Hours_Saved']
        .sum()
        .reset_index()
    )

    fig_pie_machine = px.pie(
        machine_hours,
        values='Manual_Hours_Saved',
        names='Machine_Name',
        title='Hours Saved by Machine',
        template='plotly_white',
        color_discrete_sequence=px.colors.qualitative.Safe
    )

    fig_pie_machine.update_traces(
        textinfo='label+percent',
        textposition='outside',  # ensures small slices are visible
        textfont=dict(size=13),
        hovertemplate=(
            "<b>%{label}</b><br>"
            "Hours Saved: %{value:,.0f}<br>"
            "%{percent}<extra></extra>"
        ),
        domain=dict(x=[0, 1], y=[0, 1])
    )

    fig_pie_machine.update_layout(
        title_font_size=16,
        title_font_color='#1e3a8a',
        height=400,
        margin=dict(l=20, r=20, t=60, b=100),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=-0.35,
            xanchor="center",
            x=0.5
        )
    )

    st.plotly_chart(fig_pie_machine, use_container_width=True)

st.markdown("---")


# Row 6: Treemap Visualizations
st.markdown("## 🗺️ Hierarchical Analysis")
col1, col2 = st.columns(2)

with col1:
    # Treemap: Business Area > Process > Executions
    treemap_data = filtered_df.groupby(['Business_Area', 'Process_Name']).agg({
        'Total_Executions': 'sum',
        'Cost_Savings_Dollars': 'sum'
    }).reset_index()
    
    fig_treemap1 = px.treemap(
        treemap_data,
        path=['Business_Area', 'Process_Name'],
        values='Total_Executions',
        color='Cost_Savings_Dollars',
        color_continuous_scale='RdYlGn',
        title='Process Execution Hierarchy (Size: Executions, Color: Cost Savings)',
        template='plotly_white'
    )
    fig_treemap1.update_layout(
        title_font_size=18,
        title_font_color='#1e3a8a',
        height=500
    )
    st.plotly_chart(fig_treemap1, use_container_width=True)

with col2:
    # Treemap: Business Area > SubArea > Application
    treemap_data2 = filtered_df.groupby(['Business_Area', 'Business_SubArea', 'Application']).agg({
        'Manual_Hours_Saved': 'sum',
        'Total_Executions': 'sum'
    }).reset_index()
    
    fig_treemap2 = px.treemap(
        treemap_data2,
        path=['Business_Area', 'Business_SubArea', 'Application'],
        values='Manual_Hours_Saved',
        color='Total_Executions',
        color_continuous_scale='Viridis',
        title='Application Hierarchy (Size: Hours Saved, Color: Executions)',
        template='plotly_white'
    )
    fig_treemap2.update_layout(
        title_font_size=18,
        title_font_color='#1e3a8a',
        height=500
    )
    st.plotly_chart(fig_treemap2, use_container_width=True)

st.markdown("---")

# Row 7: Advanced Visualizations
st.markdown("## 🎯 Advanced Analytics")
col1, col2 = st.columns(2)

with col1:
    # Sunburst Chart
    sunburst_data = filtered_df.groupby(['Business_Area', 'Business_SubArea', 'Process_Name']).agg({
        'Cost_Savings_Dollars': 'sum'
    }).reset_index()
    
    fig_sunburst = px.sunburst(
        sunburst_data,
        path=['Business_Area', 'Business_SubArea', 'Process_Name'],
        values='Cost_Savings_Dollars',
        title='Cost Savings Hierarchy (Sunburst)',
        color='Cost_Savings_Dollars',
        color_continuous_scale='thermal',
        template='plotly_white'
    )
    fig_sunburst.update_layout(
        title_font_size=18,
        title_font_color='#1e3a8a',
        height=500
    )
    st.plotly_chart(fig_sunburst, use_container_width=True)

with col2:
    # Bubble Chart: Process efficiency
    bubble_data = filtered_df.groupby('Process_Name').agg({
        'Total_Executions': 'sum',
        'Manual_Hours_Saved': 'sum',
        'Cost_Savings_Dollars': 'sum'
    }).reset_index()
    bubble_data['Avg_Savings_Per_Execution'] = bubble_data['Cost_Savings_Dollars'] / bubble_data['Total_Executions']
    bubble_data = bubble_data.nlargest(15, 'Total_Executions')
    
    fig_bubble = px.scatter(
        bubble_data,
        x='Total_Executions',
        y='Manual_Hours_Saved',
        size='Cost_Savings_Dollars',
        color='Avg_Savings_Per_Execution',
        hover_name='Process_Name',
        title='Process Efficiency Analysis (Bubble Chart)',
        template='plotly_white',
        color_continuous_scale='Plasma',
        size_max=60
    )
    fig_bubble.update_layout(
        title_font_size=18,
        title_font_color='#1e3a8a',
        height=500
    )
    st.plotly_chart(fig_bubble, use_container_width=True)

st.markdown("---")

# Row 8: Heatmap and Comparison
col1, col2 = st.columns(2)

with col1:
    # Heatmap: Month vs Business Area
    heatmap_data = filtered_df.pivot_table(
        values='Cost_Savings_Dollars',
        index='Business_Area',
        columns='Month',
        aggfunc='sum',
        fill_value=0
    )
    
    # Reorder columns by month
    month_order = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
    heatmap_data = heatmap_data.reindex(columns=[m for m in month_order if m in heatmap_data.columns])
    
    fig_heatmap = px.imshow(
        heatmap_data,
        labels=dict(x="Month", y="Business Area", color="Cost Savings ($)"),
        title="Cost Savings Heatmap: Business Area × Month",
        template='plotly_white',
        color_continuous_scale='YlOrRd',
        aspect='auto'
    )
    fig_heatmap.update_layout(
        title_font_size=18,
        title_font_color='#1e3a8a',
        height=450
    )
    st.plotly_chart(fig_heatmap, use_container_width=True)

with col2:
    # Year-over-Year comparison
    yoy_data = filtered_df.groupby(['Run_Year', 'Month']).agg({
        'Cost_Savings_Dollars': 'sum'
    }).reset_index()
    
    fig_yoy = px.line(
        yoy_data,
        x='Month',
        y='Cost_Savings_Dollars',
        color='Run_Year',
        title='Year-over-Year Cost Savings Comparison',
        template='plotly_white',
        markers=True,
        category_orders={"Month": month_order}
    )
    fig_yoy.update_layout(
        title_font_size=18,
        title_font_color='#1e3a8a',
        height=450,
        hovermode='x unified'
    )
    fig_yoy.update_yaxes(tickprefix='$')
    st.plotly_chart(fig_yoy, use_container_width=True)

st.markdown("---")

# Row 9: Summary Table and Funnel
col1, col2 = st.columns(2)

with col1:
    # Top Performers Summary Table
    st.markdown("### 🏆 Top Performing Processes")
    
    summary_table = filtered_df.groupby('Process_Name').agg({
        'Total_Executions': 'sum',
        'Manual_Hours_Saved': 'sum',
        'Cost_Savings_Dollars': 'sum',
        'Successful_Executions': 'sum'
    }).reset_index()
    
    summary_table['Success_Rate'] = (summary_table['Successful_Executions'] / summary_table['Total_Executions'] * 100).round(1)
    summary_table = summary_table.sort_values('Cost_Savings_Dollars', ascending=False).head(10)
    
    summary_table['Cost_Savings_Dollars'] = summary_table['Cost_Savings_Dollars'].apply(lambda x: f'${x:,.0f}')
    summary_table['Manual_Hours_Saved'] = summary_table['Manual_Hours_Saved'].apply(lambda x: f'{x:,.0f}')
    summary_table['Total_Executions'] = summary_table['Total_Executions'].apply(lambda x: f'{x:,}')
    summary_table['Success_Rate'] = summary_table['Success_Rate'].apply(lambda x: f'{x}%')
    
    summary_table.columns = ['Process', 'Executions', 'Hours Saved', 'Cost Savings', 'Successful', 'Success Rate']
    
    st.dataframe(summary_table, use_container_width=True, hide_index=True)

with col2:
    # Funnel Chart: Process execution stages
    total_exec = filtered_df['Total_Executions'].sum()
    successful_exec = filtered_df['Successful_Executions'].sum()
    # exception_exec = filtered_df['Exception_Executions'].sum()
    
    funnel_data = pd.DataFrame({
        'Stage': ['Total Initiated', 'Successful'],
         #With Exceptions
        'Count': [total_exec, successful_exec, ] 
        # exception_exec
    })
    
    fig_funnel = go.Figure(go.Funnel(
        y=funnel_data['Stage'],
        x=funnel_data['Count'],
        textinfo="value+percent initial",
        marker=dict(color=["#1559c5", "#15cf34", '#10b981', '#ef4444'])
    ))
    fig_funnel.update_layout(
        title='Execution Funnel Analysis',
        title_font_size=20,
        title_font_color='#1e3a8a',
        height=450,
        template='plotly_white'
    )
    st.plotly_chart(fig_funnel, use_container_width=True)

st.markdown("---")

# Footer with additional insights
st.markdown("## 📋 Data Summary")
col1, col2, col3, col4 = st.columns(4)

with col1:
    st.info(f"""
    **Data Period**  
    {filtered_df['Date'].min().strftime('%B %Y')} to {filtered_df['Date'].max().strftime('%B %Y')}
    """)

with col2:
    avg_hourly_rate = filtered_df['Cost_Savings_Dollars'].sum() / filtered_df['Manual_Hours_Saved'].sum() if filtered_df['Manual_Hours_Saved'].sum() > 0 else 0
    st.info(f"""
    **Avg. Hourly Rate**  
    ${avg_hourly_rate:.2f} per hour
    """)

with col3:
    st.info(f"""
    **Total Processes**  
    {filtered_df['Process_Name'].nunique()} unique processes
    """)

with col4:
    st.info(f"""
    **Business Coverage**  
    {filtered_df['Business_Area'].nunique()} business areas
    """)


# Download filtered data
st.markdown("---")
st.markdown("### 📥 Export Data", unsafe_allow_html=True)
# ---------- CSV ----------
csv_data = filtered_df.to_csv(index=False).encode("utf-8")

# ---------- Excel ----------
excel_buffer = BytesIO()
with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
    filtered_df.to_excel(writer, index=False, sheet_name="RPA Metrics")

excel_data = excel_buffer.getvalue()

# ---------- Buttons (centered) ----------
col1, col2, col3 = st.columns([1, 2, 1])

with col2:
    col_csv, col_xlsx = st.columns(2)

    with col_csv:
        st.download_button(
            label="⬇️ Download CSV",
            data=csv_data,
            file_name=f"rpa_metrics_{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv",
            use_container_width=True,
        )

    with col_xlsx:
        st.download_button(
            label="⬇️ Download Excel",
            data=excel_data,
            file_name=f"rpa_metrics_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

# Footer
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #6b7280; padding: 20px;'>
    <p>RPA Metrics Dashboard | Built with Streamlit & Plotly</p>
</div>
""", unsafe_allow_html=True)
