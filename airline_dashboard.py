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
    .info-box {
        background-color: #e8f4fd;
        padding: 1.5rem;
        border-radius: 0.5rem;
        border-left: 4px solid #1f77b4;
        margin: 1rem 0;
    }
    .explanation-box {
        background-color: #f8f9fa;
        padding: 1rem;
        border-radius: 0.5rem;
        margin: 1rem 0;
        border: 1px solid #dee2e6;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 24px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 50px;
        padding-left: 20px;
        padding-right: 20px;
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
                'Red': '#dc3545',
                'green': '#28a745',
                'orange': '#fd7e14', 
                'red': '#dc3545'
            }
            df['color'] = df['percentage'].map(color_map)
            df['color'] = df['color'].fillna('#6c757d')  # Gray for unknown
        
        # Filter out rows with missing critical data
        df = df.dropna(subset=['origin_airport', 'destination_airport', 'airline', 'min_charge2'])
        
        return df
        
    except Exception as e:
        st.error(f"Error loading data: {str(e)}")
        return None

def show_simple_overview(df):
    """Show simple overview of the loaded data"""
    st.markdown('<div class="section-header">📊 Your Data Summary</div>', unsafe_allow_html=True)
    
    # Simple metrics
    col1, col2, col3 = st.columns(3)
    
    with col1:
        unique_routes = df['route'].nunique()
        st.metric("🛫 Available Routes", f"{unique_routes}")
    
    with col2:
        unique_airlines = df['airline'].nunique()
        st.metric("✈️ Airlines Competing", f"{unique_airlines}")
    
    with col3:
        total_records = len(df)
        st.metric("📝 Total Bids", f"{total_records}")
    
    # Explanation of what this data means
    st.markdown("""
    <div class="explanation-box">
    <h4>📋 What This Data Shows</h4>
    <p><strong>This dashboard analyzes airline pricing bids for cargo shipments.</strong></p>
    <ul>
        <li><strong>Routes:</strong> Different origin-destination airport pairs where cargo can be shipped</li>
        <li><strong>Airlines:</strong> Different carriers competing to transport your cargo</li>
        <li><strong>Bids:</strong> Price quotes from airlines for specific routes</li>
        <li><strong>Color Rating:</strong> How competitive each price is compared to others</li>
    </ul>
    </div>
    """, unsafe_allow_html=True)

def create_route_analysis(df, origin, destination):
    """Create detailed analysis for a specific route"""
    route_data = df[(df['origin_airport'] == origin) & (df['destination_airport'] == destination)].copy()
    
    if route_data.empty:
        st.warning("⚠️ No data found for this route combination.")
        return
    
    route_name = f"{origin} ➜ {destination}"
    
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
    
    # Main visualization with proper colors
    st.markdown("### 📊 Airlines Comparison")
    
    # Sort by price for better visualization
    route_data = route_data.sort_values('min_charge2')
    
    # Create the main chart with proper color mapping
    fig = go.Figure()
    
    # Add bars with colors based on percentage rating
    fig.add_trace(go.Bar(
        x=route_data['airline'],
        y=route_data['min_charge2'],
        marker_color=route_data['color'],
        text=[f"${price:.2f}<br>{category}" for price, category in zip(route_data['min_charge2'], route_data['percentage'])],
        textposition='outside',
        hovertemplate="<b>%{x}</b><br>" +
                      "Price: $%{y:.2f}<br>" +
                      "Category: %{customdata}<br>" +
                      "<extra></extra>",
        customdata=route_data['percentage'],
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
    
    # Color explanation
    st.markdown("""
    <div class="explanation-box">
    <h4>🎨 Price Color Guide</h4>
    <p><strong>🟢 Green:</strong> Most competitive prices (cheapest 20%)</p>
    <p><strong>🟠 Orange:</strong> Moderate prices (middle 50%)</p>
    <p><strong>🔴 Red:</strong> Highest prices (most expensive 30%)</p>
    </div>
    """, unsafe_allow_html=True)
    
    return route_data

def show_detailed_route_info(route_data):
    """Show detailed information table for the route"""
    st.markdown("### 📋 Detailed Information")
    
    # Prepare display data
    display_columns = ['airline', 'min_charge2', 'percentage', 'direct_indirect', 'rating']
    available_columns = [col for col in display_columns if col in route_data.columns]
    
    display_df = route_data[available_columns].copy()
    
    # Rename for better display
    column_names = {
        'airline': 'Airline',
        'min_charge2': 'Price (USD)',
        'percentage': 'Price Category',
        'direct_indirect': 'Connection Type',
        'rating': 'Internal Rating'
    }
    
    display_df = display_df.rename(columns=column_names)
    
    # Format price
    if 'Price (USD)' in display_df.columns:
        display_df['Price (USD)'] = display_df['Price (USD)'].apply(lambda x: f"${x:.2f}")
    
    # Style the dataframe
    def highlight_categories(row):
        if 'Price Category' in row.index:
            category = str(row['Price Category']).lower()
            if category in ['green']:
                return ['background-color: #d4edda'] * len(row)
            elif category in ['orange']:
                return ['background-color: #fff3cd'] * len(row)
            elif category in ['red']:
                return ['background-color: #f8d7da'] * len(row)
        return [''] * len(row)
    
    styled_df = display_df.style.apply(highlight_categories, axis=1)
    st.dataframe(styled_df, use_container_width=True, hide_index=True)
    
    # Additional route insights
    if len(route_data) > 1:
        cheapest = route_data.loc[route_data['min_charge2'].idxmin()]
        most_expensive = route_data.loc[route_data['min_charge2'].idxmax()]
        savings = most_expensive['min_charge2'] - cheapest['min_charge2']
        savings_percent = (savings / most_expensive['min_charge2']) * 100
        
        st.markdown(f"""
        <div class="info-box">
        <h4>💡 Route Insights</h4>
        <p><strong>Best Deal:</strong> {cheapest['airline']} at ${cheapest['min_charge2']:.2f}</p>
        <p><strong>Most Expensive:</strong> {most_expensive['airline']} at ${most_expensive['min_charge2']:.2f}</p>
        <p><strong>Potential Savings:</strong> ${savings:.2f} ({savings_percent:.1f}%) by choosing the best option</p>
        </div>
        """, unsafe_allow_html=True)

def create_airline_overview(df):
    """Create overview analysis by airline"""
    st.markdown('<div class="section-header">🏢 Airlines Performance Overview</div>', unsafe_allow_html=True)
    
    # Calculate airline statistics
    airline_stats = df.groupby('airline').agg({
        'min_charge2': ['mean', 'min', 'max', 'count'],
        'route': 'nunique',
        'percentage': lambda x: (x.astype(str).str.lower() == 'green').sum() / len(x) * 100
    }).round(2)
    
    # Flatten column names
    airline_stats.columns = ['avg_price', 'min_price', 'max_price', 'total_bids', 'routes_served', 'green_percentage']
    airline_stats = airline_stats.reset_index()
    airline_stats = airline_stats.sort_values('total_bids', ascending=False)
    
    # Airlines summary table
    st.markdown("### 📊 Airlines Performance Summary")
    
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
        'min_price': 'Lowest Price',
        'max_price': 'Highest Price',
        'total_bids': 'Total Bids',
        'routes_served': 'Routes Served',
        'green_percentage': 'Best Price Rate'
    })
    
    st.dataframe(display_stats, use_container_width=True, hide_index=True)
    
    # Performance charts
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
                'green_percentage': 'Best Price Rate (%)',
                'total_bids': 'Total Bids'
            },
            color_continuous_scale='RdYlGn'
        )
        fig1.update_layout(height=400)
        st.plotly_chart(fig1, use_container_width=True)
    
    with col2:
        # Top 10 airlines by competitiveness (green percentage)
        top_competitive = airline_stats.nlargest(10, 'green_percentage')
        
        fig2 = px.bar(
            top_competitive,
            x='airline',
            y='green_percentage',
            title='Most Competitive Airlines (% of Best Prices)',
            labels={'green_percentage': 'Best Price Rate (%)', 'airline': 'Airlines'},
            color='green_percentage',
            color_continuous_scale='RdYlGn'
        )
        fig2.update_layout(height=400, showlegend=False)
        st.plotly_chart(fig2, use_container_width=True)

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
            # Show simple overview
            show_simple_overview(df)
            
            # Create tabs for different sections
            tab1, tab2 = st.tabs(["📊 Route Analysis", "🏢 Airlines Overview"])
            
            with tab1:
                # Route selection section
                st.markdown('<div class="section-header">🎯 Select Your Route</div>', unsafe_allow_html=True)
                
                st.markdown("""
                <div class="info-box">
                <strong>📍 Choose Origin and Destination Airports</strong><br>
                Select airports to compare all airlines serving that specific route, see their prices, and understand which offers the best value.
                </div>
                """, unsafe_allow_html=True)
                
                # Get unique airports
                origins = sorted(df['origin_airport'].unique())
                
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
                    route_data = create_route_analysis(df, selected_origin, selected_destination)
                    
                    if route_data is not None and not route_data.empty:
                        # Sub-tabs for summary and detailed info
                        sub_tab1, sub_tab2 = st.tabs(["📈 Summary", "📋 Detailed Info"])
                        
                        with sub_tab1:
                            st.markdown("### 💡 Quick Summary")
                            
                            best_airline = route_data.loc[route_data['min_charge2'].idxmin()]
                            worst_airline = route_data.loc[route_data['min_charge2'].idxmax()]
                            
                            col1, col2 = st.columns(2)
                            
                            with col1:
                                st.markdown(f"""
                                **🏆 Best Option:**
                                - **Airline:** {best_airline['airline']}
                                - **Price:** ${best_airline['min_charge2']:.2f}
                                - **Type:** {best_airline.get('direct_indirect', 'N/A')}
                                - **Rating:** {best_airline['percentage']}
                                """)
                            
                            with col2:
                                if len(route_data) > 1:
                                    savings = worst_airline['min_charge2'] - best_airline['min_charge2']
                                    st.markdown(f"""
                                    **💰 Savings Opportunity:**
                                    - **Most Expensive:** ${worst_airline['min_charge2']:.2f} ({worst_airline['airline']})
                                    - **Potential Savings:** ${savings:.2f}
                                    - **Savings %:** {(savings/worst_airline['min_charge2']*100):.1f}%
                                    """)
                        
                        with sub_tab2:
                            show_detailed_route_info(route_data)
            
            with tab2:
                create_airline_overview(df)
            
    else:
        # Instructions when no file is uploaded
        st.markdown("""
        <div class="info-box">
        <strong>🚀 Welcome to the Airline Bids Dashboard!</strong><br>
        Upload your Excel file above to start analyzing airline pricing data and find the best shipping options.
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("""
        ### 🎯 What This Dashboard Does
        
        **Helps you make smart shipping decisions by:**
        
        - 🔍 **Comparing Airlines:** See all carriers serving your route
        - 💰 **Finding Best Prices:** Identify the most cost-effective options  
        - 📊 **Understanding Value:** Color-coded ratings show price competitiveness
        - 📈 **Analyzing Trends:** Overview of airline performance across all routes
        
        ### 🎨 Easy Color System
        
        - 🟢 **Green = Great Deal** (Top 20% best prices)
        - 🟠 **Orange = Fair Price** (Middle 50% range)
        - 🔴 **Red = Premium Price** (Top 30% most expensive)
        
        ### 📁 File Requirements
        
        Upload an Excel file with an **'Airline Bids'** sheet containing:
        - Airport codes (origin/destination)
        - Airline names and pricing
        - Route and connection information
        """)

if __name__ == "__main__":
    main()
