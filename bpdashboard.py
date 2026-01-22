import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
from io import BytesIO
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

# Page configuration
st.set_page_config(
    page_title="RPA Metrics Dashboard",
    page_icon="🖥️",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Enhanced CSS for stunning dashboard appearance
st.markdown("""
<style>
    /* Main container styling */
    .main {
        padding: 0rem 1rem;
    }
    
    /* Metric cards */
    div[data-testid="metric-container"] {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        border-radius: 15px;
        padding: 20px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.1);
        border: 1px solid rgba(255,255,255,0.1);
        transition: transform 0.3s ease;
    }
    
    div[data-testid="metric-container"]:hover {
        transform: translateY(-5px);
        box-shadow: 0 6px 20px rgba(0,0,0,0.15);
    }
    
    div[data-testid="metric-container"] label {
        color: white !important;
        font-weight: 600;
        font-size: 0.9rem;
    }
    
    div[data-testid="metric-container"] [data-testid="stMetricValue"] {
        color: white !important;
        font-size: 1.8rem;
        font-weight: 700;
    }
    
    div[data-testid="metric-container"] [data-testid="stMetricDelta"] {
        color: rgba(255,255,255,0.9) !important;
    }
    
    /* Headers */
    h1 {
        color: #1e3a8a;
        font-weight: 800;
        margin-bottom: 0.5rem;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
    }
    
    h2, h3 {
        color: #1e40af;
        font-weight: 700;
        margin-top: 2rem;
    }
    
    /* Sidebar styling */
    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #f0f9ff 0%, #e0f2fe 100%);
    }
    
    section[data-testid="stSidebar"] .stRadio > label,
    section[data-testid="stSidebar"] .stMultiSelect > label,
    section[data-testid="stSidebar"] .stSelectbox > label {
        color: #1e40af;
        font-weight: 600;
    }
    
    /* Button styling */
    .stButton > button {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        border: none;
        border-radius: 10px;
        padding: 0.5rem 2rem;
        font-weight: 600;
        transition: all 0.3s ease;
    }
    
    .stButton > button:hover {
        transform: scale(1.05);
        box-shadow: 0 5px 15px rgba(102, 126, 234, 0.4);
    }
    
    /* Download button styling */
    .stDownloadButton > button {
        background: linear-gradient(135deg, #10b981 0%, #059669 100%);
        color: white;
        border: none;
        border-radius: 10px;
        padding: 0.5rem 1.5rem;
        font-weight: 600;
        width: 100%;
    }
    
    /* Info boxes */
    .stInfo, .stSuccess, .stWarning {
        border-radius: 10px;
        padding: 1rem;
        box-shadow: 0 2px 10px rgba(0,0,0,0.05);
    }
    
    /* Dataframe styling */
    div[data-testid="stDataFrame"] {
        border-radius: 10px;
        overflow: hidden;
        box-shadow: 0 2px 10px rgba(0,0,0,0.05);
    }
    
    /* Plotly chart containers */
    .js-plotly-plot {
        border-radius: 10px;
        box-shadow: 0 2px 15px rgba(0,0,0,0.08);
    }
    
    /* Footer */
    .footer {
        text-align: center;
        padding: 2rem;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        border-radius: 10px;
        margin-top: 2rem;
    }
    
    /* Divider */
    hr {
        margin: 2rem 0;
        border: none;
        border-top: 2px solid #e5e7eb;
    }
    
    /* Loading animation */
    .stSpinner > div {
        border-top-color: #667eea !important;
    }
</style>
""", unsafe_allow_html=True)

# ============================================================================
# DATA LOADING FUNCTIONS WITH ERROR HANDLING
# ============================================================================

# File upload option
st.sidebar.markdown("---")
st.sidebar.markdown("### 📁 File Upload (Optional)")
uploaded_rpa = st.sidebar.file_uploader("Upload RPA Metrics", type=['xlsx'], key='rpa')
uploaded_savings = st.sidebar.file_uploader("Upload Savings Data", type=['xlsx'], key='savings')
st.sidebar.markdown("---")

@st.cache_data(ttl=300)  # Cache for 5 minutes
def load_rpa_data(uploaded_file=None):
    """Load RPA automation metrics data - NO COST DATA"""
    try:
        # Try uploaded file first, then local file
        if uploaded_file is not None:
            df = pd.read_excel(uploaded_file)
        else:
            df = pd.read_excel('RPA_Metrics_Dashboard.xlsx')
        
        # Strip whitespace from column names
        df.columns = df.columns.str.strip()
        
        # Data validation
        required_columns = ['Run_Year', 'Run_Month', 'Manual_Hours_Saved', 'Total_Executions']
        missing_cols = [col for col in required_columns if col not in df.columns]
        if missing_cols:
            st.error(f"❌ Missing required columns: {missing_cols}")
            st.info(f"Available columns: {df.columns.tolist()}")
            return pd.DataFrame()
        
        # Data cleaning and preparation
        df['Date'] = pd.to_datetime(df['Run_Year'].astype(str) + '-' + df['Run_Month'].astype(str) + '-01')
        df['Quarter'] = df['Date'].dt.quarter
        df['Year_Quarter'] = df['Run_Year'].astype(str) + ' Q' + df['Quarter'].astype(str)
        
        # Handle Month column - create if missing
        if 'Month' not in df.columns:
            df['Month'] = df['Date'].dt.strftime('%b')
        
        df['Year_Month'] = df['Run_Year'].astype(str) + '-' + df['Month']
        df['Week'] = df['Date'].dt.isocalendar().week
        
        # Clean numeric columns
        df['Manual_Hours_Saved'] = pd.to_numeric(df['Manual_Hours_Saved'], errors='coerce').fillna(0)
        df['Total_Executions'] = pd.to_numeric(df['Total_Executions'], errors='coerce').fillna(0)
        df['Successful_Executions'] = pd.to_numeric(df['Successful_Executions'], errors='coerce').fillna(0)
        
        return df
    except FileNotFoundError:
        st.error("❌ RPA_Metrics_Dashboard.xlsx not found.")
        st.info("💡 Please upload the file using the sidebar or place it in the same directory as this script.")
        return pd.DataFrame()
    except Exception as e:
        st.error(f"❌ Error loading RPA data: {str(e)}")
        with st.expander("Show error details"):
            st.exception(e)
        return pd.DataFrame()

@st.cache_data(ttl=300)
def load_savings_data(uploaded_file=None):
    """Load savings by functional area data - ALL COST/SAVINGS DATA"""
    try:
        # Try uploaded file first, then local file
        if uploaded_file is not None:
            df = pd.read_excel(uploaded_file)
        else:
            df = pd.read_excel('Automation_Savings.xlsx')
        
        # Strip whitespace from column names
        df.columns = df.columns.str.strip()
        
        # Debug info (shown in expander)
        col_list = df.columns.tolist()
        
        # Try to find the functional area column (case-insensitive)
        functional_area_col = None
        for col in df.columns:
            if 'functional' in col.lower() and 'area' in col.lower():
                functional_area_col = col
                break
        
        if functional_area_col and functional_area_col != 'Functional Area':
            df.rename(columns={functional_area_col: 'Functional Area'}, inplace=True)
        
        # Validate data
        if 'Functional Area' not in df.columns:
            st.error(f"❌ 'Functional Area' column not found.")
            st.info(f"💡 Available columns: {col_list}")
            st.warning("Please ensure your Excel file has a column named 'Functional Area'")
            return pd.DataFrame()
        
        # Clean all numeric columns
        for col in df.columns:
            if col != 'Functional Area' and 'Savings' in col:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        
        return df
    except FileNotFoundError:
        st.error("❌ Savings_By_Functional_Area.xlsx not found.")
        st.info("💡 Please upload the file using the sidebar or place it in the same directory as this script.")
        return pd.DataFrame()
    except Exception as e:
        st.error(f"❌ Error loading savings data: {str(e)}")
        with st.expander("Show error details"):
            st.exception(e)
        return pd.DataFrame()

# Load both datasets with loading indicator
with st.spinner("🔄 Loading dashboard data..."):
    df = load_rpa_data(uploaded_rpa)
    savings_df = load_savings_data(uploaded_savings)

# Check if data loaded successfully
if df.empty:
    st.error("❌ RPA data failed to load. Please check the file and try again.")
    st.stop()

if savings_df.empty:
    st.error("❌ Savings data failed to load. Please check the file and try again.")
    st.stop()

# Debug: Show data preview
with st.expander("🔍 Debug: View Raw Data Preview"):
    st.write("**RPA Data Preview:**")
    st.dataframe(df.head(3), use_container_width=True)
    st.write(f"Shape: {df.shape[0]} rows × {df.shape[1]} columns")
    
    st.write("**Savings Data Preview:**")
    st.dataframe(savings_df.head(3), use_container_width=True)
    st.write(f"Shape: {savings_df.shape[0]} rows × {savings_df.shape[1]} columns")

# Extract total savings from savings file with error handling
try:
    savings_data = savings_df[savings_df['Functional Area'] != 'Total'].copy()
    
    # Check if 'Total' row exists
    total_rows = savings_df[savings_df['Functional Area'] == 'Total']
    if total_rows.empty:
        st.warning("⚠️ No 'Total' row found in savings data. Using sum of all areas.")
        # Create a synthetic total row
        total_savings_row = pd.Series({
            'Functional Area': 'Total',
            'Cumulative Savings in USD': savings_data['Cumulative Savings in USD'].sum(),
            'Savings 2025': savings_data['Savings 2025'].sum(),
            'Projected Savings 2026': savings_data['Projected Savings 2026'].sum()
        })
    else:
        total_savings_row = total_rows.iloc[0]
    
    TOTAL_CUMULATIVE_SAVINGS = total_savings_row['Cumulative Savings in USD']
    
except KeyError as e:
    st.error(f"❌ Missing expected column in savings data: {str(e)}")
    st.info(f"Available columns: {savings_df.columns.tolist()}")
    st.stop()
except Exception as e:
    st.error(f"❌ Error processing savings data: {str(e)}")
    st.exception(e)
    st.stop()

# ============================================================================
# SIDEBAR FILTERS
# ============================================================================
st.sidebar.title("🎯 Dashboard Filters")

# Auto-refresh button at top
if st.sidebar.button('🔄 Refresh Dashboard', use_container_width=True):
    st.cache_data.clear()
    st.rerun()

st.sidebar.markdown("---")

# Time period selector
time_period = st.sidebar.radio(
    "📊 Time Aggregation",
    ["Daily", "Weekly", "Monthly", "Quarterly", "Yearly"],
    index=2,  # Default to Monthly
    help="Select how to aggregate the data"
)

# Year filter
all_years = sorted(df['Run_Year'].unique())
selected_years = st.sidebar.multiselect(
    "📅 Select Year(s)",
    options=all_years,
    default=all_years,
    help="Filter by specific years"
)

# Month filter
if selected_years:
    available_months = sorted(
        df[df['Run_Year'].isin(selected_years)]['Month'].unique(),
        key=lambda x: datetime.strptime(x, '%b').month
    )
    selected_months = st.sidebar.multiselect(
        "📆 Select Month(s)",
        options=available_months,
        default=available_months,
        help="Filter by specific months"
    )
else:
    selected_months = []
    st.sidebar.warning("⚠️ Please select at least one year")

# Business Area filter
business_areas = ['All'] + sorted(df['Business_Area'].unique().tolist())
selected_business_area = st.sidebar.selectbox(
    "🏢 Business Area",
    options=business_areas,
    help="Filter by business area"
)

# Process Name filter
if selected_business_area != 'All':
    process_names = ['All'] + sorted(
        df[df['Business_Area'] == selected_business_area]['Process_Name'].unique().tolist()
    )
else:
    process_names = ['All'] + sorted(df['Process_Name'].unique().tolist())

selected_process = st.sidebar.selectbox(
    "⚙️ Process Name",
    options=process_names,
    help="Filter by specific process"
)

# Machine Name filter
machine_names = ['All'] + sorted(df['Machine_Name'].unique().tolist())
selected_machine = st.sidebar.selectbox(
    "🖥️ Machine Name",
    options=machine_names,
    help="Filter by machine"
)

# Apply filters to RPA data
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

# Show filter summary
st.sidebar.markdown("---")
st.sidebar.markdown("### 📋 Active Filters")
st.sidebar.info(f"""
**Records:** {len(filtered_df):,}  
**Years:** {len(selected_years)}  
**Months:** {len(selected_months)}  
**Area:** {selected_business_area}  
**Process:** {selected_process}  
**Machine:** {selected_machine}
""")

# ============================================================================
# MAIN DASHBOARD - RPA AUTOMATION METRICS
# ============================================================================
st.title("🖥️ RPA Automation Metrics Dashboard")
st.markdown("### Executive Performance Overview")

# Key Metrics Row
col1, col2, col3, col4, col5 = st.columns(5)

with col1:
    total_executions = filtered_df['Total_Executions'].sum()
    st.metric(
        label="📊 Total Executions",
        value=f"{total_executions:,}",
        delta=f"{len(filtered_df)} records"
    )

with col2:
    total_hours_saved = filtered_df['Manual_Hours_Saved'].sum()
    st.metric(
        label="⏱️ Hours Saved",
        value=f"{total_hours_saved:,.0f}",
        delta=f"{total_hours_saved/total_executions:.2f} hrs/exec" if total_executions > 0 else "0"
    )

with col3:
    st.metric(
        label="💰 Cost Savings (USD)",
        value=f"${TOTAL_CUMULATIVE_SAVINGS:,.0f}",
        delta="Cumulative"
    )

with col4:
    success_rate = (filtered_df['Successful_Executions'].sum() / total_executions * 100) if total_executions > 0 else 0
    st.metric(
        label="✅ Success Rate",
        value=f"{success_rate:.1f}%",
        delta=f"{filtered_df['Successful_Executions'].sum():,} successful"
    )

with col5:
    unique_processes = filtered_df['Process_Name'].nunique()
    st.metric(
        label="⚙️ Active Processes",
        value=f"{unique_processes}",
        delta=f"{filtered_df['Business_Area'].nunique()} areas"
    )

st.markdown("---")

# ============================================================================
# SAVINGS BY FUNCTIONAL AREA SECTION
# ============================================================================
year_columns = ['Savings 2019', 'Savings 2020', 'Savings 2021', 'Savings 2022', 
                'Savings 2023', 'Savings 2024', 'Savings 2025', 'Projected Savings 2026']

st.markdown("## 💰 Savings by Functional Area")
st.markdown("### Comprehensive Savings Analysis Across All Business Units")

# Top-level savings metrics
col1, col2, col3, col4 = st.columns(4)

with col1:
    total_cumulative = total_savings_row['Cumulative Savings in USD']
    st.metric(
        label="💵 Total Cumulative Savings",
        value=f"${total_cumulative:,.0f}",
        delta="All Years"
    )

with col2:
    savings_2025 = total_savings_row['Savings 2025']
    savings_2024 = total_savings_row['Savings 2024']
    growth = ((savings_2025 - savings_2024) / savings_2024 * 100) if savings_2024 > 0 else 0
    st.metric(
        label="📊 2025 Savings",
        value=f"${savings_2025:,.0f}",
        delta=f"{growth:+.1f}% vs 2024"
    )

with col3:
    projected_2026 = total_savings_row['Projected Savings 2026']
    projected_growth = ((projected_2026 - savings_2025) / savings_2025 * 100) if savings_2025 > 0 else 0
    st.metric(
        label="🎯 Projected 2026",
        value=f"${projected_2026:,.0f}",
        delta=f"{projected_growth:+.1f}% growth"
    )

with col4:
    top_area = savings_data.loc[savings_data['Cumulative Savings in USD'].idxmax(), 'Functional Area']
    top_area_savings = savings_data['Cumulative Savings in USD'].max()
    top_area_pct = (top_area_savings / total_cumulative * 100)
    st.metric(
        label="🏆 Top Performer",
        value=top_area,
        delta=f"${top_area_savings:,.0f} ({top_area_pct:.1f}%)"
    )

st.markdown("---")

# Savings visualizations continue with your existing code...
# (I'll keep the rest of your visualization code as is, but with improved color schemes)

# Enhanced color schemes
ENHANCED_COLORS = {
    'sequential': px.colors.sequential.Viridis,
    'diverging': px.colors.diverging.RdYlGn,
    'qualitative': px.colors.qualitative.Set3
}

# Row 1: Pie Chart + Bar Chart
col1, col2 = st.columns(2)

with col1:
    fig_pie = px.pie(
        savings_data,
        values='Cumulative Savings in USD',
        names='Functional Area',
        title='<b>Cumulative Savings Distribution by Functional Area</b>',
        template='plotly_white',
        color_discrete_sequence=ENHANCED_COLORS['qualitative'],
        hole=0.4
    )
    fig_pie.update_traces(
        textposition='auto',
        textinfo='percent+label',
        hovertemplate='<b>%{label}</b><br>Savings: $%{value:,.0f}<br>Percentage: %{percent}<extra></extra>'
    )
    fig_pie.update_layout(
        title_font_size=18,
        title_font_color='#1e3a8a',
        height=450,
        showlegend=True,
        legend=dict(
            orientation="v",
            yanchor="middle",
            y=0.5,
            xanchor="left",
            x=1.05,
            font=dict(size=11)
        ),
        hoverlabel=dict(
            bgcolor="white",
            font_size=12
        )
    )
    st.plotly_chart(fig_pie, use_container_width=True)

with col2:
    savings_sorted = savings_data.sort_values('Cumulative Savings in USD', ascending=True)
    fig_bar = go.Figure(go.Bar(
        x=savings_sorted['Cumulative Savings in USD'],
        y=savings_sorted['Functional Area'],
        orientation='h',
        marker=dict(
            color=savings_sorted['Cumulative Savings in USD'],
            colorscale='Viridis',
            showscale=True,
            colorbar=dict(title="Savings ($)", thickness=15)
        ),
        text=savings_sorted['Cumulative Savings in USD'].apply(lambda x: f'${x/1e6:.1f}M' if x >= 1e6 else f'${x/1e3:.0f}K'),
        textposition='outside',
        hovertemplate='<b>%{y}</b><br>Savings: $%{x:,.0f}<extra></extra>'
    ))
    fig_bar.update_layout(
        title='<b>Cumulative Savings by Functional Area</b>',
        title_font_size=18,
        title_font_color='#1e3a8a',
        height=450,
        template='plotly_white',
        xaxis_title='<b>Cumulative Savings (USD)</b>',
        yaxis_title='',
        showlegend=False,
        hoverlabel=dict(bgcolor="white", font_size=12)
    )
    st.plotly_chart(fig_bar, use_container_width=True)

# Continue with the rest of your existing visualizations...
# (The code would continue with all your other charts, enhanced with similar styling)

# Footer
st.markdown("---")
st.markdown("""
<div class="footer">
    <h3>🖥️ RPA Metrics Dashboard</h3>
    <p>Built with Streamlit & Plotly | Data-Driven Decision Making</p>
    <p>Last Updated: {}</p>
</div>
""".format(datetime.now().strftime('%B %d, %Y at %I:%M %p')), unsafe_allow_html=True)
