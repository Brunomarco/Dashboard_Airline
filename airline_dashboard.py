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
    .metric-container {
        background-color: #f8f9fa;
        padding: 1rem;
        border-radius: 0.5rem;
        margin: 0.5rem 0;
    }
    .stSelectbox > div > div > div > div {
        background-color: white;
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
        df['route'] = df['origin_airport'] + ' → ' + df['destination_airport']
        
        # Clean percentage column and create color mapping
        if 'percentage' in df.columns:
            df['percentage'] = df['percentage'].astype(str).str.strip()
            df['color'] = df['percentage'].map({
                'Green': '#28a745',
                'Orange': '#fd7e14', 
                'Red': '#dc3545'
            })
        
        # Filter out rows with missing critical data
        df = df.dropna(subset=['origin_airport', 'destination_airport', 'airline', 'min_charge2'])
        
        return df
        
    except Exception as e:
        st.error(f"Error loading data: {str(e)}")
        return None

def create_route_comparison_chart(df, selected_route):
    """Create a comparison chart for airlines serving a specific route"""
    route_data = df[df['route'] == selected_route].copy()
    
    if route_data.empty:
        return None
    
    # Sort by price
    route_data = route_data.sort_values('min_charge2')
    
    # Create bar chart
    fig = px.bar(
        route_data,
        x='airline',
        y='min_charge2',
        color='percentage',
        color_discrete_map={
            'Green': '#28a745',
            'Orange': '#fd7e14',
            'Red': '#dc3545'
        },
        title=f'Airline Pricing Comparison for {selected_route}',
        labels={'min_charge2': 'Min Charge (USD)', 'airline': 'Airline'},
        text='min_charge2'
    )
    
    # Update layout
    fig.update_traces(texttemplate='$%{text:.2f}', textposition='outside')
    fig.update_layout(
        xaxis_title="Airline",
        yaxis_title="Min Charge (USD)",
        showlegend=True,
        height=400
    )
    
    return fig

def create_price_distribution_chart(df):
    """Create a price distribution chart by color category"""
    color_counts = df['percentage'].value_counts()
    
    fig = px.pie(
        values=color_counts.values,
        names=color_counts.index,
        title="Price Distribution by Category",
        color=color_counts.index,
        color_discrete_map={
            'Green': '#28a745',
            'Orange': '#fd7e14',
            'Red': '#dc3545'
        }
    )
    
    fig.update_layout(height=400)
    return fig

def create_route_overview_chart(df):
    """Create an overview chart of all routes with average prices"""
    route_summary = df.groupby('route').agg({
        'min_charge2': 'mean',
        'airline': 'count'
    }).round(2)
    
    route_summary.columns = ['avg_price', 'airline_count']
    route_summary = route_summary.reset_index()
    route_summary = route_summary.sort_values('avg_price', ascending=True)
    
    # Take top 20 routes by number of airlines
    top_routes = route_summary.nlargest(20, 'airline_count')
    
    fig = px.scatter(
        top_routes,
        x='airline_count',
        y='avg_price',
        size='airline_count',
        hover_data=['route'],
        title="Route Overview: Average Price vs Number of Airlines",
        labels={
            'airline_count': 'Number of Airlines',
            'avg_price': 'Average Price (USD)'
        }
    )
    
    fig.update_layout(height=500)
    return fig

def main():
    # Header
    st.markdown('<h1 class="main-header">✈️ Airline Bids Dashboard</h1>', unsafe_allow_html=True)
    
    # File upload
    uploaded_file = st.file_uploader(
        "Upload your Excel file with Airline Bids data",
        type=['xlsx', 'xls'],
        help="Please upload the Excel file containing the 'Airline Bids' sheet"
    )
    
    if uploaded_file is not None:
        # Load data
        with st.spinner("Loading data..."):
            df = load_data(uploaded_file)
        
        if df is not None:
            # Sidebar filters
            st.sidebar.header("Filters")
            
            # Route selection
            routes = sorted(df['route'].unique())
            selected_route = st.sidebar.selectbox(
                "Select Route",
                routes,
                help="Choose a route to see detailed airline comparison"
            )
            
            # Airline filter
            airlines = sorted(df['airline'].unique())
            selected_airlines = st.sidebar.multiselect(
                "Filter by Airlines",
                airlines,
                default=airlines,
                help="Select airlines to include in the analysis"
            )
            
            # Price range filter
            if 'min_charge2' in df.columns:
                min_price = float(df['min_charge2'].min())
                max_price = float(df['min_charge2'].max())
                price_range = st.sidebar.slider(
                    "Price Range (USD)",
                    min_value=min_price,
                    max_value=max_price,
                    value=(min_price, max_price),
                    step=0.1
                )
            
            # Apply filters
            filtered_df = df[
                (df['airline'].isin(selected_airlines)) &
                (df['min_charge2'] >= price_range[0]) &
                (df['min_charge2'] <= price_range[1])
            ]
            
            # Main dashboard
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.metric("Total Routes", len(filtered_df['route'].unique()))
            
            with col2:
                st.metric("Total Airlines", len(filtered_df['airline'].unique()))
            
            with col3:
                avg_price = filtered_df['min_charge2'].mean()
                st.metric("Average Price", f"${avg_price:.2f}")
            
            with col4:
                green_percentage = (filtered_df['percentage'] == 'Green').mean() * 100
                st.metric("Green Rates %", f"{green_percentage:.1f}%")
            
            # Charts
            st.markdown("---")
            
            # Route-specific analysis
            col1, col2 = st.columns([2, 1])
            
            with col1:
                st.subheader(f"Airline Comparison for {selected_route}")
                route_chart = create_route_comparison_chart(filtered_df, selected_route)
                if route_chart:
                    st.plotly_chart(route_chart, use_container_width=True)
                else:
                    st.info("No data available for the selected route with current filters.")
            
            with col2:
                st.subheader("Price Category Distribution")
                price_dist_chart = create_price_distribution_chart(filtered_df)
                st.plotly_chart(price_dist_chart, use_container_width=True)
            
            # Route overview
            st.markdown("---")
            st.subheader("Routes Overview")
            route_overview_chart = create_route_overview_chart(filtered_df)
            st.plotly_chart(route_overview_chart, use_container_width=True)
            
            # Detailed data table
            st.markdown("---")
            st.subheader("Detailed Data")
            
            # Table filters
            col1, col2 = st.columns(2)
            with col1:
                show_route = st.selectbox(
                    "Show data for route:",
                    ["All routes"] + routes,
                    key="table_route_filter"
                )
            
            with col2:
                sort_by = st.selectbox(
                    "Sort by:",
                    ["min_charge2", "airline", "route", "percentage"],
                    key="table_sort"
                )
            
            # Filter and sort data for table
            table_df = filtered_df.copy()
            if show_route != "All routes":
                table_df = table_df[table_df['route'] == show_route]
            
            table_df = table_df.sort_values(sort_by)
            
            # Display columns for the table
            display_columns = [
                'route', 'airline', 'min_charge2', 'percentage', 
                'direct_indirect', 'currency'
            ]
            
            # Filter columns that exist in the dataframe
            available_columns = [col for col in display_columns if col in table_df.columns]
            
            # Rename columns for display
            column_names = {
                'route': 'Route',
                'airline': 'Airline',
                'min_charge2': 'Price (USD)',
                'percentage': 'Price Category',
                'direct_indirect': 'Connection Type',
                'currency': 'Currency'
            }
            
            display_df = table_df[available_columns].copy()
            display_df = display_df.rename(columns=column_names)
            
            # Format price column
            if 'Price (USD)' in display_df.columns:
                display_df['Price (USD)'] = display_df['Price (USD)'].apply(lambda x: f"${x:.2f}" if pd.notna(x) else "N/A")
            
            # Color code the rows based on price category
            def highlight_price_category(row):
                if 'Price Category' in row.index:
                    if row['Price Category'] == 'Green':
                        return ['background-color: #d4edda'] * len(row)
                    elif row['Price Category'] == 'Orange':
                        return ['background-color: #fff3cd'] * len(row)
                    elif row['Price Category'] == 'Red':
                        return ['background-color: #f8d7da'] * len(row)
                return [''] * len(row)
            
            styled_df = display_df.style.apply(highlight_price_category, axis=1)
            st.dataframe(styled_df, use_container_width=True, height=400)
            
            # Download filtered data
            csv = table_df.to_csv(index=False)
            st.download_button(
                label="Download filtered data as CSV",
                data=csv,
                file_name=f"airline_bids_filtered_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.csv",
                mime="text/csv"
            )
            
            # Summary statistics
            st.markdown("---")
            st.subheader("Summary Statistics")
            
            col1, col2, col3 = st.columns(3)
            
            with col1:
                st.markdown("**Price Statistics**")
                if 'min_charge2' in filtered_df.columns:
                    st.write(f"• Minimum Price: ${filtered_df['min_charge2'].min():.2f}")
                    st.write(f"• Maximum Price: ${filtered_df['min_charge2'].max():.2f}")
                    st.write(f"• Median Price: ${filtered_df['min_charge2'].median():.2f}")
                    st.write(f"• Standard Deviation: ${filtered_df['min_charge2'].std():.2f}")
            
            with col2:
                st.markdown("**Category Distribution**")
                if 'percentage' in filtered_df.columns:
                    category_counts = filtered_df['percentage'].value_counts()
                    total = len(filtered_df)
                    for category, count in category_counts.items():
                        percentage = (count / total) * 100
                        st.write(f"• {category}: {count} ({percentage:.1f}%)")
            
            with col3:
                st.markdown("**Route Statistics**")
                routes_count = len(filtered_df['route'].unique())
                airlines_count = len(filtered_df['airline'].unique())
                avg_airlines_per_route = filtered_df.groupby('route')['airline'].nunique().mean()
                
                st.write(f"• Total Routes: {routes_count}")
                st.write(f"• Total Airlines: {airlines_count}")
                st.write(f"• Avg Airlines per Route: {avg_airlines_per_route:.1f}")
            
    else:
        # Instructions when no file is uploaded
        st.info("👆 Please upload an Excel file to get started")
        
        st.markdown("""
        ### Expected File Format
        
        Your Excel file should contain a sheet named **'Airline Bids'** with the following columns:
        
        - **Origin Airport**: Source airport code
        - **Destination Airport**: Destination airport code  
        - **Airline**: Airline code or name
        - **Min Charge2**: Price in USD
        - **Percentage**: Price category (Green/Orange/Red)
        - **Direct / Indirect**: Connection type
        - And other relevant columns...
        
        ### Color Coding System
        
        - 🟢 **Green**: Cheapest 20% of prices
        - 🟠 **Orange**: Next 50% of prices  
        - 🔴 **Red**: Most expensive 30% of prices
        
        ### Dashboard Features
        
        - **Route Comparison**: Compare airlines serving the same route
        - **Price Analysis**: Distribution of prices by category
        - **Interactive Filters**: Filter by route, airline, and price range
        - **Data Export**: Download filtered results as CSV
        """)

if __name__ == "__main__":
    main()
