# Airline Bids Dashboard ✈️

A Streamlit dashboard for visualizing and analyzing airline bid data with origin-destination pairs, carrier pricing, and color-coded price rankings.

## Features

- **Interactive Route Analysis**: Compare airlines serving specific routes
- **Price Category Visualization**: Color-coded pricing tiers (Green/Orange/Red)
- **Dynamic Filtering**: Filter by route, airline, and price range
- **Data Export**: Download filtered results as CSV
- **Summary Statistics**: Key metrics and distributions

## Color Coding System

- 🟢 **Green**: Cheapest 20% of prices
- 🟠 **Orange**: Next 50% of prices  
- 🔴 **Red**: Most expensive 30% of prices

## Installation & Setup

### Prerequisites
- Python 3.8 or higher
- pip package manager

### Local Setup

1. **Clone the repository**
   ```bash
   git clone <your-repo-url>
   cd airline-bids-dashboard
   ```

2. **Create a virtual environment** (recommended)
   ```bash
   python -m venv venv
   
   # On Windows
   venv\Scripts\activate
   
   # On macOS/Linux
   source venv/bin/activate
   ```

3. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

4. **Run the application**
   ```bash
   streamlit run airline_dashboard.py
   ```

5. **Open your browser**
   The app will automatically open at `http://localhost:8501`

### GitHub Deployment

To deploy on Streamlit Cloud:

1. **Push to GitHub**
   ```bash
   git add .
   git commit -m "Initial commit"
   git push origin main
   ```

2. **Deploy on Streamlit Cloud**
   - Go to [share.streamlit.io](https://share.streamlit.io)
   - Connect your GitHub account
   - Select your repository
   - Choose `airline_dashboard.py` as the main file
   - Click "Deploy"

## File Structure

```
airline-bids-dashboard/
├── airline_dashboard.py    # Main Streamlit application
├── requirements.txt        # Python dependencies
├── README.md              # This file
└── sample_data/           # (Optional) Sample data files
```

## Data Format

Your Excel file should contain a sheet named **'Airline Bids'** with these columns:

| Column | Description |
|--------|-------------|
| Origin Airport | Source airport code (e.g., JFK, LAX) |
| Destination Airport | Destination airport code |
| Airline | Airline code or name |
| Min Charge2 | Price in USD |
| Percentage | Price category (Green/Orange/Red) |
| Direct / Indirect | Connection type |
| Via | Connection airport (if indirect) |
| Currency | Currency code |

## Usage

1. **Upload Data**: Click "Browse files" and select your Excel file
2. **Select Route**: Use the sidebar to choose a specific origin-destination pair
3. **Apply Filters**: Filter by airlines and price range
4. **Analyze**: View charts and tables showing:
   - Airline price comparison for selected route
   - Price category distribution
   - Route overview with airline counts
   - Detailed data table with sorting options

## Troubleshooting

### Common Issues

**File Upload Error**
- Ensure your Excel file contains a sheet named 'Airline Bids'
- Check that required columns exist
- Verify the data starts around row 11 (after headers)

**Missing Data**
- Check that Origin Airport, Destination Airport, and Airline columns have data
- Ensure Min Charge2 column contains numeric values

**Performance Issues**
- Large files (>10MB) may take longer to load
- Consider filtering data before upload if possible

## Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/new-feature`)
3. Commit your changes (`git commit -am 'Add new feature'`)
4. Push to the branch (`git push origin feature/new-feature`)
5. Create a Pull Request

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Support

If you encounter any issues or have questions:
1. Check the troubleshooting section above
2. Create an issue on GitHub
3. Review the data format requirements

## Screenshots

*Add screenshots of your dashboard here once deployed*

---

Built with ❤️ using [Streamlit](https://streamlit.io/)
