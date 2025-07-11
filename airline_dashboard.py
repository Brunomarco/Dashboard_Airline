import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
import openpyxl
from io import BytesIO

# Set page configuration
st.set_page_config(
    page_title="Airline Bids Dashboard",
    page_icon="✈️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better styling
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
    .section-header {
        font-size: 1.8rem;
        font-weight: bold;
        color: #2c3e50;
        margin: 1.5rem 0 1rem 0;
        border-bottom: 2px solid #3498db;
        padding-bottom: 0.5rem;
    }
    .metric-container {
        background-color: #f8f9fa;
        padding: 1rem;
        border-radius: 0.5rem;
        margin: 0.5rem 0;
    }
    .info-box {
        background-color: #e8f4fd;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #1f77b4;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

@st.cache_data
def load_data(uploaded_file):
    """Load and process the Excel file data"""
    try:
        # Read the Excel file
        workbook = openpyxl.load_workbook(uploaded_file, data_only=True)
        
        # Get the 'Airline Bids' sheet
        if 'Airline Bids' not in workbook.sheetnames:
            st.error("Sheet 'Airline Bids' not found in the Excel file")
            return None
        
        sheet = workbook['Airline Bids']
        
        # Convert to DataFrame starting from row 11 (data starts there)
        data = []
        headers = []
        
        # Get headers from row 10 (index 10)
        for col in range(3, sheet.max_column + 1):  # Start from column C (index 3)
            cell = sheet.cell(row=11, column=col)
            headers.append(cell.value if cell.value else f'col_{col}')
        
        # Get data starting from row 12
        for row in range(12, sheet.max_row + 1):
            row_data = []
            for col in range(3, sheet.max_column + 1):
                cell = sheet.cell(row=row, column=col)
                row_data.append(cell.value)
            
            # Only add rows that have data in key columns
            if row_data[3] and row_data[4] and row_data[13]:  # Origin Airport, Destination Airport, Airline
                data.append(row_data)
        
        # Create DataFrame
        df = pd.DataFrame(data, columns=headers)
        
        # Clean and standardize column names
        column_mapping = {
            'Commodity Group': 'commodity_group',
            'TempControlled': 'temp_controlled',
            'Air Mode': 'air_mode',
            'Origin Airport': 'origin_airport',
            'Destination Airport': 'destination_airport',
            'Origin Country': 'origin_country',
            'Destinatin Country': 'destination_country',
            'Origin Region': 'origin_region',
            'Destination Region': 'destination_region',
            'Airline': 'airline',
            'Intention to Bid (Yes/No)': 'intention_to_bid',
            'Direct / Indirect': 'direct_indirect',
            'Via': 'via',
            'Currency': 'currency',
            'Min Charge': 'min_charge',
            'Min Charge2': 'min_charge2',
            'Percentage': 'percentage',
            'Numerical Rating': 'rating'
        }
        
        # Rename columns that exist in the DataFrame
        for old_name, new_name in column_mapping.items():
            if old_name in df.columns:
                df = df.rename(columns={old_name: new_name})
        
        # Convert numeric columns
        numeric_columns = ['min_charge2', 'rating']
        for col in numeric_columns:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce')
        
        # Create route column
        df['route'] = df['origin_airport'] + ' ➜ ' + df['destination_airport']
        
        # Clean percentage column and create color mapping
        if 'percentage' in df.columns:
            df['percentage'] = df['percentage'].astype(str).str.strip()
            # Map colors for visualization
            color_map = {
                'Green': '#28a745',
                'Orange': '#fd7e14', 
                'Red': '#dc3545'
            }
            df['color'] = df['percentage'].map(color_map)
            df['color'] = df['color'].fillna('#6c757d')  # Gray for unknown
        
        # Filter out rows with missing critical data
        df = df.dropna(subset=['origin_airport', 'destination_airport', 'airline', 'min_charge2'])
        
        return df
        
    except Exception as e:
        st.error(f"Error loading data: {str(e)}")
        return None

def show_data_overview(df):
    """Show overview of the loaded data"""
    st.markdown('<div class="section-header">📊 Data Overview</div>', unsafe_allow_html=True)
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        total_records = len(df)
        st.metric("📝 Total Records", f"{total_records:,}")
    
    with col2:
        unique_routes = df['route'].nunique()
        st.metric("🛫 Unique Routes", f"{unique_routes:,}")
    
    with col3:
        unique_airlines = df['airline'].nunique()
        st.metric("✈️ Airlines", f"{unique_airlines:,}")
    
    with col4:
        avg_price = df['min_charge2'].mean()
        st.metric("💰 Avg Price", f"${avg_price:.2f}")
    
    # Price distribution by category
    st.markdown("### Price Category Distribution")
    
    if 'percentage' in df.columns:
        color_counts = df['percentage'].value_counts()
        
        col1, col2 = st.columns([2, 1])
        
        with col1:
            fig = px.pie(
                values=color_counts.values,
                names=color_counts.index,
                title="Distribution of Price Categories",
                color=color_counts.index,
                color_discrete_map={
                    'Green': '#28a745',
                    'Orange': '#fd7e14',
                    'Red': '#dc3545'
                },
                hole=0.4
            )
            fig.update_traces(textposition='inside', textinfo='percent+label')
            fig.update_layout(height=300, showlegend=True)
            st.plotly_chart(fig, use_container_width=True)
        
        with col2:
            st.markdown("**Category Meanings:**")
            st.markdown("🟢 **Green**: Cheapest 20%")
            st.markdown("🟠 **Orange**: Middle 50%")
            st.markdown("🔴 **Red**: Most expensive 30%")
            
            st.markdown("**Summary:**")
            total = color_counts.sum()
            for category, count in color_counts.items():
                percentage = (count / total) * 100
                st.write(f"• {category}: {count} ({percentage:.1f}%)")

def create_route_analysis(df, origin, destination):
    """Create detailed analysis for a specific route"""
    route_data = df[(df['origin_airport'] == origin) & (df['destination_airport'] == destination)].copy()
    
    if route_data.empty:
        st.warning("⚠️ No data found for this route combination.")
        return
    
    route_name = f"{origin} ➜ {destination}"
    st.markdown(f'<div class="section-header">🎯 Route Analysis: {route_name}</div>', unsafe_allow_html=True)
    
    # Route metrics
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        airline_count = route_data['airline'].nunique()
        st.metric("Airlines Available", airline_count)
    
    with col2:
        min_price = route_data['min_charge2'].min()
        st.metric("Best Price", f"${min_price:.2f}")
    
    with col3:
        avg_price = route_data['min_charge2'].mean()
        st.metric("Average Price", f"${avg_price:.2f}")
    
    with col4:
        price_range = route_data['min_charge2'].max() - route_data['min_charge2'].min()
        st.metric("Price Range", f"${price_range:.2f}")
    
    # Main visualization
    st.markdown("### 📊 Airlines Comparison")
    
    # Sort by price for better visualization
    route_data = route_data.sort_values('min_charge2')
    
    # Create the main chart
    fig = go.Figure()
    
    # Add bars with custom colors
    fig.add_trace(go.Bar(
        x=route_data['airline'],
        y=route_data['min_charge2'],
        marker_color=route_data['color'],
        text=[f"${price:.2f}<br>{category}" for price, category in zip(route_data['min_charge2'], route_data['percentage'])],
        textposition='outside',
        hovertemplate="<b>%{x}</b><br>" +
                      "Price: $%{y:.2f}<br>" +
                      "<extra></extra>",
        name="Price"
    ))
    
    fig.update_layout(
        title=f"Airline Pricing for {route_name}",
        xaxis_title="Airlines",
        yaxis_title="Price (USD)",
        height=500,
        showlegend=False,
        xaxis={'categoryorder': 'total ascending'}
    )
    
    st.plotly_chart(fig, use_container_width=True)
    
    # Detailed data table
    st.markdown("### 📋 Detailed Information")
    
    # Prepare display data
    display_data = route_data.copy()
    display_columns = ['airline', 'min_charge2', 'percentage', 'direct_indirect', 'rating']
    available_columns = [col for col in display_columns if col in display_data.columns]
    
    display_df = display_data[available_columns].copy()
    
    # Rename for better display
    column_names = {
        'airline': 'Airline',
        'min_charge2': 'Price (USD)',
        'percentage': 'Price Category',
        'direct_indirect': 'Connection',
        'rating': 'Rating'
    }
    
    display_df = display_df.rename(columns=column_names)
    
    # Format price
    if 'Price (USD)' in display_df.columns:
        display_df['Price (USD)'] = display_df['Price (USD)'].apply(lambda x: f"${x:.2f}")
    
    # Style the dataframe
    def highlight_categories(row):
        if 'Price Category' in row.index:
            if row['Price Category'] == 'Green':
                return ['background-color: #d4edda'] * len(row)
            elif row['Price Category'] == 'Orange':
                return ['background-color: #fff3cd'] * len(row)
            elif row['Price Category'] == 'Red':
                return ['background-color: #f8d7da'] * len(row)
        return [''] * len(row)
    
    styled_df = display_df.style.apply(highlight_categories, axis=1)
    st.dataframe(styled_df, use_container_width=True, hide_index=True)

def create_airline_overview(df):
    """Create overview analysis by airline"""
    st.markdown('<div class="section-header">🏢 Airlines Overview</div>', unsafe_allow_html=True)
    
    # Calculate airline statistics
    airline_stats = df.groupby('airline').agg({
        'min_charge2': ['mean', 'min', 'max', 'count'],
        'route': 'nunique',
        'percentage': lambda x: (x == 'Green').sum() / len(x) * 100
    }).round(2)
    
    # Flatten column names
    airline_stats.columns = ['avg_price', 'min_price', 'max_price', 'total_bids', 'routes_served', 'green_percentage']
    airline_stats = airline_stats.reset_index()
    airline_stats = airline_stats.sort_values('total_bids', ascending=False)
    
    # Top airlines chart
    st.markdown("### 📈 Airline Performance Metrics")
    
    col1, col2 = st.columns(2)
    
    with col1:
        # Routes served vs Average price
        fig1 = px.scatter(
            airline_stats.head(15),
            x='routes_served',
            y='avg_price',
            size='total_bids',
            color='green_percentage',
            hover_name='airline',
            title="Routes Served vs Average Price",
            labels={
                'routes_served': 'Number of Routes Served',
                'avg_price': 'Average Price (USD)',
                'green_percentage': 'Green Rate (%)',
                'total_bids': 'Total Bids'
            },
            color_continuous_scale='RdYlGn'
        )
        fig1.update_layout(height=400)
        st.plotly_chart(fig1, use_container_width=True)
    
    with col2:
        # Price range by airline
        top_airlines = airline_stats.head(10)
        fig2 = go.Figure()
        
        fig2.add_trace(go.Bar(
            name='Min Price',
            x=top_airlines['airline'],
            y=top_airlines['min_price'],
            marker_color='lightblue'
        ))
        
        fig2.add_trace(go.Bar(
            name='Average Price',
            x=top_airlines['airline'],
            y=top_airlines['avg_price'],
            marker_color='blue'
        ))
        
        fig2.add_trace(go.Bar(
            name='Max Price',
            x=top_airlines['airline'],
            y=top_airlines['max_price'],
            marker_color='darkblue'
        ))
        
        fig2.update_layout(
            title='Price Range by Top Airlines',
            xaxis_title='Airlines',
            yaxis_title='Price (USD)',
            barmode='group',
            height=400
        )
        st.plotly_chart(fig2, use_container_width=True)
    
    # Airlines summary table
    st.markdown("### 📊 Airlines Summary Table")
    
    # Format the display
    display_stats = airline_stats.copy()
    display_stats['avg_price'] = display_stats['avg_price'].apply(lambda x: f"${x:.2f}")
    display_stats['min_price'] = display_stats['min_price'].apply(lambda x: f"${x:.2f}")
    display_stats['max_price'] = display_stats['max_price'].apply(lambda x: f"${x:.2f}")
    display_stats['green_percentage'] = display_stats['green_percentage'].apply(lambda x: f"{x:.1f}%")
    
    # Rename columns
    display_stats = display_stats.rename(columns={
        'airline': 'Airline',
        'avg_price': 'Avg Price',
        'min_price': 'Min Price',
        'max_price': 'Max Price',
        'total_bids': 'Total Bids',
        'routes_served': 'Routes',
        'green_percentage': 'Green Rate'
    })
    
    st.dataframe(display_stats, use_container_width=True, hide_index=True)

def main():
    # Header
    st.markdown('<h1 class="main-header">✈️ Airline Bids Dashboard</h1>', unsafe_allow_html=True)
    
    # File upload
    uploaded_file = st.file_uploader(
        "📁 Upload your Excel file with Airline Bids data",
        type=['xlsx', 'xls'],
        help="Please upload the Excel file containing the 'Airline Bids' sheet"
    )
    
    if uploaded_file is not None:
        # Load data
        with st.spinner("🔄 Loading and processing data..."):
            df = load_data(uploaded_file)
        
        if df is not None:
            # Show data overview first
            show_data_overview(df)
            
            # Airlines overview
            create_airline_overview(df)
            
            # Route selection section
            st.markdown('<div class="section-header">🎯 Route-Specific Analysis</div>', unsafe_allow_html=True)
            
            st.markdown("""
            <div class="info-box">
            <strong>📍 Select Origin and Destination</strong><br>
            Choose specific airports to see detailed airline comparison, pricing, and ratings for that route.
            </div>
            """, unsafe_allow_html=True)
            
            # Get unique airports
            origins = sorted(df['origin_airport'].unique())
            destinations = sorted(df['destination_airport'].unique())
            
            col1, col2 = st.columns(2)
            
            with col1:
                selected_origin = st.selectbox(
                    "🛫 Origin Airport",
                    origins,
                    help="Select the departure airport"
                )
            
            with col2:
                # Filter destinations based on origin
                available_destinations = df[df['origin_airport'] == selected_origin]['destination_airport'].unique()
                available_destinations = sorted(available_destinations)
                
                selected_destination = st.selectbox(
                    "🛬 Destination Airport",
                    available_destinations,
                    help="Select the arrival airport"
                )
            
            # Show route analysis
            if selected_origin and selected_destination:
                create_route_analysis(df, selected_origin, selected_destination)
            
    else:
        # Instructions when no file is uploaded
        st.markdown("""
        <div class="info-box">
        <strong>🚀 Getting Started</strong><br>
        Upload your Excel file above to begin analyzing airline bid data. The file should contain a sheet named 'Airline Bids' with pricing and route information.
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("""
        ### 📋 What This Dashboard Shows
        
        **1. 📊 Data Overview**
        - Summary statistics of your airline bids data
        - Price category distribution (Green/Orange/Red)
        - Key metrics at a glance
        
        **2. 🏢 Airlines Overview**  
        - Performance comparison across all airlines
        - Routes served vs pricing analysis
        - Success rates and competitiveness metrics
        
        **3. 🎯 Route-Specific Analysis**
        - Select any origin-destination pair
        - Compare all airlines serving that route
        - Clear pricing visualization with color coding
        - Detailed information table
        
        ### 🎨 Color Coding System
        
        - 🟢 **Green**: Best prices (cheapest 20%)
        - 🟠 **Orange**: Moderate prices (middle 50%)  
        - 🔴 **Red**: Highest prices (most expensive 30%)
        
        ### 📁 Expected File Format
        
        Your Excel file should have an **'Airline Bids'** sheet with columns like:
        - Origin Airport, Destination Airport
        - Airline, Min Charge2 (price)
        - Percentage (color category)
        - Direct/Indirect, Rating
        """)

if __name__ == "__main__":
    main()
