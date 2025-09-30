# Excel Viewer & Editor Pro 📊

A powerful, feature-rich Python-based Excel viewer and editor built with Streamlit. This application provides an intuitive interface for viewing, editing, and analyzing Excel files with advanced data manipulation capabilities.

## 🚀 Features

### Core Functionality
- **📂 Multi-Format Support**: Load Excel (.xlsx, .xls) and CSV files
- **📑 Multiple Sheets**: View and manage multiple sheets within a workbook
- **✏️ Interactive Editing**: Edit cell values directly in an interactive data grid
- **➕ Dynamic Rows/Columns**: Add, delete, insert rows and columns on the fly
- **💾 Export Options**: Save as Excel (.xlsx) or CSV with all modifications

### Advanced Features
- **🔍 Search & Filter**: 
  - Global search across all cells
  - Column-based filtering with multiple values
  - Save filtered results as new sheets
  
- **📊 Data Analytics**:
  - Statistical summaries (mean, median, min, max)
  - Distribution histograms
  - Correlation matrices for numeric data
  - Categorical data analysis with bar charts
  
- **🔧 Data Operations**:
  - Sort by any column (ascending/descending)
  - Rename columns
  - Dynamic row insertion (start/end)
  - Bulk deletion of rows and columns
  
- **📈 Real-time Statistics**:
  - Row/column counts
  - Filled/empty cell tracking
  - Memory usage monitoring
  - Data type analysis
  - Unique value counts

### User Experience
- **🎨 Modern UI**: Clean, intuitive interface with custom styling
- **📱 Responsive Layout**: Wide layout optimized for data viewing
- **🔄 Live Updates**: Changes reflected immediately
- **📊 Visual Analytics**: Interactive charts using Plotly
- **💡 Sample Data**: Built-in sample datasets for testing

## 📋 Requirements

```
streamlit==1.29.0
pandas==2.1.3
openpyxl==3.1.2
xlrd==2.0.1
plotly==5.18.0
numpy==1.26.2
```

## 🛠️ Installation

1. **Clone or download this repository**

2. **Install dependencies**:
```bash
pip install -r requirements.txt
```

## 🎯 Usage

### Starting the Application

Run the application using Streamlit:

```bash
streamlit run excel_app.py
```

The application will open in your default web browser at `http://localhost:8501`

### Basic Workflow

1. **Upload a File**:
   - Click "Browse files" in the sidebar
   - Select an Excel (.xlsx, .xls) or CSV file
   - The file will be loaded automatically

2. **Select a Sheet**:
   - Choose from available sheets in the dropdown
   - View sheet statistics at the top

3. **Edit Data**:
   - Go to "Edit Data" tab
   - Click any cell to edit
   - Click "Save Changes" to apply modifications

4. **Perform Operations**:
   - Use "Operations" tab for structural changes
   - Add/delete rows and columns
   - Sort data by any column

5. **Search & Filter**:
   - Use "Search & Filter" tab
   - Search across all cells or filter by column
   - Save filtered results as new sheets

6. **Analyze Data**:
   - Go to "Analytics" tab
   - View statistics and visualizations
   - Explore correlations and distributions

7. **Export Results**:
   - Click "Download Excel" in sidebar
   - Or download individual sheets as CSV

### Advanced Features

#### Sheet Management
- **Add New Sheet**: Enter a name and click "Add Sheet"
- **Delete Sheet**: Remove unwanted sheets (must have at least one)
- **Switch Sheets**: Use the dropdown to navigate between sheets

#### Row Operations
- Add rows at the beginning or end
- Delete multiple rows by selecting indices
- Rows can be added dynamically in the data editor

#### Column Operations
- Add new columns with custom names
- Delete multiple columns at once
- Rename existing columns
- Automatic naming for new columns

#### Sorting
- Sort by any column
- Choose ascending or descending order
- Changes persist until saved

#### Search Functionality
- Search for any text across the entire sheet
- Case-insensitive matching
- Results show matching rows only

#### Filtering
- Filter by specific column values
- Multi-select filtering
- Save filtered data as new sheets

#### Analytics
- **Numeric Columns**: Mean, median, min, max statistics
- **Distributions**: Interactive histograms
- **Correlations**: Heatmap for numeric columns
- **Categorical Data**: Top 10 value counts with bar charts

## 📊 Sample Data

The application includes a sample dataset feature:
- Click "Load Sample Dataset" on the welcome screen
- Includes "Sales" and "Employees" sheets
- Pre-populated with realistic data for testing

## 🎨 User Interface

### Sidebar
- File upload
- Sheet selector
- Sheet operations
- Export buttons

### Main Tabs
1. **Edit Data**: Interactive data grid for cell editing
2. **Operations**: Row/column management and sorting
3. **Search & Filter**: Find and filter data
4. **Analytics**: Statistics and visualizations
5. **Info**: Detailed sheet information

### Metrics Dashboard
- Total rows and columns
- Filled vs empty cells
- Real-time updates

## 💡 Tips & Best Practices

1. **Save Changes**: Always click "Save Changes" after editing in the data editor
2. **Backup Original**: Download a copy before making major changes
3. **Use Filters**: Save filtered views as new sheets for reference
4. **Check Statistics**: Use the Info tab to understand your data structure
5. **Explore Analytics**: Visualizations can reveal data patterns and insights

## 🔧 Technical Details

### Architecture
- **Frontend**: Streamlit (Python web framework)
- **Data Processing**: Pandas DataFrames
- **Excel I/O**: openpyxl (read/write Excel files)
- **Visualizations**: Plotly Express (interactive charts)
- **State Management**: Streamlit session state

### Key Components
- `load_excel_file()`: Reads Excel files into dictionary of DataFrames
- `save_to_excel()`: Exports DataFrames back to Excel format
- `get_dataframe_statistics()`: Calculates comprehensive data statistics
- `search_in_dataframe()`: Implements global search functionality
- Session state: Maintains data across user interactions

### Data Flow
1. File upload → Excel parsing
2. DataFrame creation per sheet
3. User interactions → State updates
4. State changes → UI refresh
5. Export → Excel file generation

## 🐛 Troubleshooting

### File Won't Load
- Ensure file format is .xlsx, .xls, or .csv
- Check file isn't corrupted
- Verify file isn't password protected

### Changes Not Saving
- Click "Save Changes" button after editing
- Check browser console for errors
- Ensure sufficient memory available

### Performance Issues
- Large files (>100MB) may be slow
- Consider filtering data to reduce size
- Close unused browser tabs

### Export Issues
- Ensure download permissions in browser
- Check available disk space
- Try CSV export if Excel export fails

## 🚀 Future Enhancements

Potential features for future versions:
- Formula support and calculation
- Cell formatting (colors, fonts, borders)
- Charts and pivot tables
- Undo/redo functionality
- Collaborative editing
- Data validation rules
- Conditional formatting
- Import from databases
- Automated data cleaning
- Export to multiple formats (JSON, Parquet)

## 📝 License

This project is open source and available for personal and commercial use.

## 🤝 Contributing

Contributions are welcome! Feel free to:
- Report bugs
- Suggest new features
- Submit pull requests
- Improve documentation

## 📧 Support

For issues or questions:
- Check this README for common solutions
- Review the code comments for implementation details
- Test with sample data to isolate issues

## 🙏 Acknowledgments

Built with:
- **Streamlit**: Modern web framework for data apps
- **Pandas**: Powerful data manipulation library
- **openpyxl**: Excel file handling
- **Plotly**: Interactive visualizations

---

**Happy Data Editing! 📊✨**
