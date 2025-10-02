import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
import plotly.express as px
import numpy as np
from datetime import datetime
import json
import os

# Page configuration
st.set_page_config(
    page_title="Excel Viewer & Editor",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better UI
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
    .stat-box {
        background-color: #f0f2f6;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #1f77b4;
    }
    .success-box {
        background-color: #d4edda;
        border-color: #c3e6cb;
        color: #155724;
        padding: 1rem;
        border-radius: 0.5rem;
        margin: 1rem 0;
    }
    .warning-box {
        background-color: #fff3cd;
        border-color: #ffeaa7;
        color: #856404;
        padding: 1rem;
        border-radius: 0.5rem;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

def get_recent_files_path():
    """Get the path to the recent files storage file"""
    return os.path.join(os.path.dirname(__file__), 'recent_files.json')

def get_file_paths_path():
    """Get the path to the file paths storage file"""
    return os.path.join(os.path.dirname(__file__), 'file_paths.json')

def load_recent_files_from_storage():
    """Load recent files from persistent storage"""
    try:
        recent_files_path = get_recent_files_path()
        if os.path.exists(recent_files_path):
            with open(recent_files_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        return []
    except Exception:
        return []

def save_recent_files_to_storage(recent_files):
    """Save recent files to persistent storage"""
    try:
        recent_files_path = get_recent_files_path()
        with open(recent_files_path, 'w', encoding='utf-8') as f:
            json.dump(recent_files, f)
    except Exception:
        pass  # Silently fail if we can't save

def load_file_paths_from_storage():
    """Load file paths from persistent storage"""
    try:
        file_paths_path = get_file_paths_path()
        if os.path.exists(file_paths_path):
            with open(file_paths_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        return {}
    except Exception:
        return {}

def save_file_paths_to_storage(file_paths):
    """Save file paths to persistent storage"""
    try:
        file_paths_path = get_file_paths_path()
        with open(file_paths_path, 'w', encoding='utf-8') as f:
            json.dump(file_paths, f)
    except Exception:
        pass  # Silently fail if we can't save

def get_file_path_by_name(file_name):
    """Get file path by file name from storage"""
    file_paths = load_file_paths_from_storage()
    file_info = file_paths.get(file_name)
    if file_info and isinstance(file_info, dict):
        # Handle old format with type/value structure
        return file_info.get("value") if file_info.get("type") == "path" else None
    elif file_info and isinstance(file_info, str):
        # Handle new simplified format
        return file_info
    return None

# Initialize session state
if 'df_dict' not in st.session_state:
    st.session_state.df_dict = {}
if 'current_sheet' not in st.session_state:
    st.session_state.current_sheet = None
if 'file_name' not in st.session_state:
    st.session_state.file_name = None
if 'edited_data' not in st.session_state:
    st.session_state.edited_data = None
if 'history' not in st.session_state:
    st.session_state.history = []
if 'redo_stack' not in st.session_state:
    st.session_state.redo_stack = []
if 'recent_files' not in st.session_state:
    st.session_state.recent_files = []

def migrate_old_data():
    """Migrate old format data to new format"""
    try:
        recent_files = load_recent_files_from_storage()
        file_paths = load_file_paths_from_storage()
        
        # Check if we need to migrate
        needs_migration = any(isinstance(item, dict) for item in recent_files)
        
        if needs_migration:
            # Migrate recent files
            migrated_files = []
            for item in recent_files:
                if isinstance(item, dict):
                    # Old format: {"name": "file.xlsx", "path": "path"}
                    file_name = item.get("name", "")
                    file_path = item.get("path")
                    if file_name:
                        migrated_files.append(file_name)
                        if file_path:
                            file_paths[file_name] = file_path
                elif isinstance(item, str):
                    # New format: "file.xlsx"
                    migrated_files.append(item)
            
            # Save migrated data
            save_recent_files_to_storage(migrated_files)
            save_file_paths_to_storage(file_paths)
            
    except Exception:
        pass  # Silently fail if migration doesn't work

def initialize_recent_files():
    """Initialize recent files from storage if not already loaded"""
    if not st.session_state.recent_files:
        # First, try to migrate old data
        migrate_old_data()
        
        # Then load the data
        recent_files = load_recent_files_from_storage()
        st.session_state.recent_files = recent_files

def add_to_recent_files(file_name, file_path):
    """Add a file to the recent files list and store its path (only for file system files)"""
    if not file_path:
        return  # Don't add files without paths
    
    recent_files = load_recent_files_from_storage()
    
    # Check if file already exists (by name)
    if file_name in recent_files:
        # Remove existing entry
        recent_files.remove(file_name)
    
    # Add to top
    recent_files.insert(0, file_name)
    
    # Keep only the last 10 files
    if len(recent_files) > 10:
        recent_files = recent_files[:10]
    
    # Save recent files
    save_recent_files_to_storage(recent_files)
    
    # Store file path
    file_paths = load_file_paths_from_storage()
    file_paths[file_name] = file_path
    save_file_paths_to_storage(file_paths)
    
    # Update session state
    st.session_state.recent_files = recent_files

def remove_from_recent_files(file_name):
    """Remove a file from the recent files list"""
    recent_files = load_recent_files_from_storage()
    if file_name in recent_files:
        recent_files.remove(file_name)
        save_recent_files_to_storage(recent_files)
        st.session_state.recent_files = recent_files

def clear_recent_files():
    """Clear all recent files"""
    save_recent_files_to_storage([])
    st.session_state.recent_files = []

def load_file_from_path(file_path):
    """Load Excel file from file path"""
    try:
        if file_path.endswith('.csv'):
            df_dict = {'Sheet1': pd.read_csv(file_path)}
            sheet_names = ['Sheet1']
        else:
            excel_file = pd.ExcelFile(file_path)
            df_dict = {}
            for sheet_name in excel_file.sheet_names:
                df_dict[sheet_name] = pd.read_excel(file_path, sheet_name=sheet_name)
            sheet_names = excel_file.sheet_names
        
        return df_dict, sheet_names
    except Exception as e:
        st.error(f"Error loading file from path: {str(e)}")
        return None, None

def load_excel_file(uploaded_file):
    """Load Excel file and return dictionary of DataFrames for each sheet"""
    try:
        # Read all sheets
        excel_file = pd.ExcelFile(uploaded_file)
        df_dict = {}
        
        for sheet_name in excel_file.sheet_names:
            df_dict[sheet_name] = pd.read_excel(uploaded_file, sheet_name=sheet_name)
        
        return df_dict, excel_file.sheet_names
    except Exception as e:
        st.error(f"Error loading file: {str(e)}")
        return None, None

def save_to_excel(df_dict, filename="edited_file.xlsx"):
    """Save dictionary of DataFrames to Excel file"""
    output = BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name, dataframe in df_dict.items():
            dataframe.to_excel(writer, sheet_name=sheet_name, index=False)
    
    output.seek(0)
    return output

def add_row(df, position='end'):
    """Add a new row to DataFrame"""
    new_row = pd.DataFrame([{col: None for col in df.columns}])
    
    if position == 'end':
        df = pd.concat([df, new_row], ignore_index=True)
    elif position == 'start':
        df = pd.concat([new_row, df], ignore_index=True)
    
    return df

def add_column(df, col_name=None):
    """Add a new column to DataFrame"""
    if col_name is None:
        col_name = f"New_Column_{len(df.columns) + 1}"
    
    df[col_name] = None
    return df

def delete_rows(df, indices):
    """Delete specified rows from DataFrame"""
    if indices:
        df = df.drop(indices).reset_index(drop=True)
    return df

def delete_columns(df, columns):
    """Delete specified columns from DataFrame"""
    if columns:
        df = df.drop(columns=columns)
    return df

def search_in_dataframe(df, search_term):
    """Search for a term in the entire DataFrame"""
    mask = df.astype(str).apply(lambda row: row.str.contains(search_term, case=False, na=False).any(), axis=1)
    return df[mask]

def get_dataframe_statistics(df):
    """Calculate statistics for the DataFrame"""
    stats = {
        'Total Rows': len(df),
        'Total Columns': len(df.columns),
        'Total Cells': len(df) * len(df.columns),
        'Memory Usage': f"{df.memory_usage(deep=True).sum() / 1024:.2f} KB"
    }
    
    # Count data types
    numeric_cols = df.select_dtypes(include=[np.number]).columns
    if len(numeric_cols) > 0:
        stats['Numeric Columns'] = len(numeric_cols)
        stats['Text Columns'] = len(df.columns) - len(numeric_cols)
    
    # Count null values
    null_count = df.isnull().sum().sum()
    stats['Empty Cells'] = null_count
    stats['Filled Cells'] = (len(df) * len(df.columns)) - null_count
    
    return stats

def create_pivot_table(df, index_col, values_col, aggfunc='sum'):
    """Create a simple pivot table"""
    try:
        pivot = df.pivot_table(values=values_col, index=index_col, aggfunc=aggfunc)
        return pivot
    except Exception as e:
        st.error(f"Error creating pivot table: {str(e)}")
        return None

# Main UI
st.markdown('<div class="main-header">📊 Excel Viewer & Editor Pro</div>', unsafe_allow_html=True)

# Sidebar
with st.sidebar:
    st.header("📁 File Management")
    
    # Initialize recent files
    initialize_recent_files()
    
    # Recent Files Section
    if st.session_state.recent_files:
        st.subheader("📋 Recent Files")
        
        # Create display names for selectbox
        recent_file_names = []
        for item in st.session_state.recent_files:
            # Handle both old and new data formats
            if isinstance(item, dict):
                # Old format: extract name
                file_name = item.get("name", "Unknown")
            elif isinstance(item, str):
                # New format: use directly
                file_name = item
            else:
                # Skip invalid items
                continue
                
            # Get file path from storage
            file_path = get_file_path_by_name(file_name)
            
            if file_path:
                recent_file_names.append(f"📁 {file_name} ({file_path})")
            else:
                # Skip files without paths (shouldn't happen now, but for safety)
                continue
        
        # Display recent files with selection
        selected_recent_file = st.selectbox(
            "Select from recent files",
            ["None"] + recent_file_names,
            help="Choose a recently opened file"
        )
        
        # Legend
        st.caption("📁 = File system file (can be opened directly)")
        
        if selected_recent_file != "None":
            col1, col2 = st.columns([3, 1])
            with col1:
                if st.button("📂 Open Selected File", key="open_recent"):
                    # Extract file name from selected option
                    selected_index = recent_file_names.index(selected_recent_file)
                    selected_item = st.session_state.recent_files[selected_index]
                    
                    # Handle both old and new data formats
                    if isinstance(selected_item, dict):
                        selected_file_name = selected_item.get("name", "Unknown")
                    elif isinstance(selected_item, str):
                        selected_file_name = selected_item
                    else:
                        st.error("Invalid file data format")
                        st.rerun()
                    
                    # Get file path from storage
                    file_path = get_file_path_by_name(selected_file_name)
                    
                    if file_path:
                        # Try to open the file from path
                        if os.path.exists(file_path):
                            try:
                                df_dict, sheet_names = load_file_from_path(file_path)
                                if df_dict:
                                    st.session_state.df_dict = df_dict
                                    st.session_state.current_sheet = sheet_names[0]
                                    st.session_state.file_name = selected_file_name
                                    st.session_state.history = []
                                    st.session_state.redo_stack = []
                                    st.success(f"✅ Opened: {selected_file_name}")
                                    st.rerun()
                                else:
                                    st.error("Failed to load file from path")
                            except Exception as e:
                                st.error(f"Error opening file: {str(e)}")
                        else:
                            st.error(f"File not found: {file_path}")
                    else:
                        st.error("File path not available")
            with col2:
                if st.button("🗑️", help="Remove from recent files", key="remove_recent"):
                    selected_index = recent_file_names.index(selected_recent_file)
                    selected_item = st.session_state.recent_files[selected_index]
                    
                    # Handle both old and new data formats
                    if isinstance(selected_item, dict):
                        selected_file_name = selected_item.get("name", "Unknown")
                    elif isinstance(selected_item, str):
                        selected_file_name = selected_item
                    else:
                        st.error("Invalid file data format")
                        st.rerun()
                        remove_from_recent_files(selected_file_name)
                        st.rerun()
        
        # Clear all recent files button
        if st.button("🧹 Clear All Recent Files", key="clear_recent"):
            clear_recent_files()
            st.rerun()
        
        st.divider()
    
    # File upload
    uploaded_file = st.file_uploader(
        "Upload Excel File",
        type=['xlsx', 'xls', 'csv'],
        help="Upload an Excel or CSV file to view and edit"
    )
    
    
    
    if uploaded_file is not None:
        # Load file if new or different
        if st.session_state.file_name != uploaded_file.name:
            df_dict, sheet_names = load_excel_file(uploaded_file)
            
            if df_dict:
                st.session_state.df_dict = df_dict
                st.session_state.current_sheet = sheet_names[0]
                st.session_state.file_name = uploaded_file.name
                st.session_state.history = []
                st.session_state.redo_stack = []
                # Don't add uploaded files to recent files since they can't be reopened
                st.success(f"✅ Loaded: {uploaded_file.name}")
    
    # File path input for opening files from file system
    st.divider()
    st.markdown("**💡 Open from file system:**")
    
    file_path = st.text_input(
        "Enter file path",
        placeholder="C:\\path\\to\\your\\file.xlsx",
        help="Enter the full path to an Excel or CSV file. Files opened this way will be added to recent files.",
        value=st.session_state.get('quick_file_path', '')
    )
    
    # Clear quick file path after use
    if 'quick_file_path' in st.session_state:
        del st.session_state.quick_file_path
    
    if file_path and st.button("📂 Open from Path", key="open_path"):
        if os.path.exists(file_path):
            try:
                df_dict, sheet_names = load_file_from_path(file_path)
                if df_dict:
                    st.session_state.df_dict = df_dict
                    st.session_state.current_sheet = sheet_names[0]
                    st.session_state.file_name = os.path.basename(file_path)
                    st.session_state.history = []
                    st.session_state.redo_stack = []
                    # Add to recent files with path (only file system files)
                    add_to_recent_files(os.path.basename(file_path), file_path)
                    st.success(f"✅ Opened: {os.path.basename(file_path)}")
                    st.rerun()
                else:
                    st.error("Failed to load file from path")
            except Exception as e:
                st.error(f"Error opening file: {str(e)}")
        else:
            st.error(f"File not found: {file_path}")
    
    st.divider()
    
    # Sheet selector
    if st.session_state.df_dict:
        st.divider()
        st.header("📑 Sheets")
        
        sheet_names = list(st.session_state.df_dict.keys())
        st.session_state.current_sheet = st.selectbox(
            "Select Sheet",
            sheet_names,
            index=sheet_names.index(st.session_state.current_sheet) if st.session_state.current_sheet in sheet_names else 0
        )
        
        # Sheet management
        with st.expander("Sheet Operations"):
            new_sheet_name = st.text_input("New Sheet Name")
            col1, col2 = st.columns(2)
            
            with col1:
                if st.button("➕ Add Sheet"):
                    if new_sheet_name and new_sheet_name not in st.session_state.df_dict:
                        st.session_state.df_dict[new_sheet_name] = pd.DataFrame()
                        st.session_state.current_sheet = new_sheet_name
                        st.rerun()
                    elif new_sheet_name in st.session_state.df_dict:
                        st.error("Sheet already exists!")
                    else:
                        st.error("Please enter a sheet name")
            
            with col2:
                if st.button("🗑️ Delete Sheet"):
                    if len(st.session_state.df_dict) > 1:
                        del st.session_state.df_dict[st.session_state.current_sheet]
                        st.session_state.current_sheet = list(st.session_state.df_dict.keys())[0]
                        st.rerun()
                    else:
                        st.error("Cannot delete the only sheet!")
        
        st.divider()
        
        # Download options - only show for file system files
        file_path = get_file_path_by_name(st.session_state.file_name)
        if file_path:
            st.header("💾 Export")
            
            output_file = save_to_excel(st.session_state.df_dict, st.session_state.file_name)
            
            st.download_button(
                label="📥 Download Excel",
                data=output_file,
                file_name=f"edited_{st.session_state.file_name}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
            # CSV export for current sheet
            if st.session_state.current_sheet:
                csv = st.session_state.df_dict[st.session_state.current_sheet].to_csv(index=False)
                st.download_button(
                    label="📄 Download as CSV",
                    data=csv,
                    file_name=f"{st.session_state.current_sheet}.csv",
                    mime="text/csv",
                    use_container_width=True
                )
        else:
            # For uploaded files, show a message that export is not available
            st.info("💡 **Export not available for uploaded files.** To export your changes, please open the file from your file system using the file path input above.")

# Main content area
if st.session_state.df_dict and st.session_state.current_sheet:
    df = st.session_state.df_dict[st.session_state.current_sheet].copy()
    
    # Statistics overview
    st.header(f"📊 Current Sheet: {st.session_state.current_sheet}")
    
    stats = get_dataframe_statistics(df)
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Total Rows", stats['Total Rows'])
    with col2:
        st.metric("Total Columns", stats['Total Columns'])
    with col3:
        st.metric("Filled Cells", stats.get('Filled Cells', 0))
    with col4:
        st.metric("Empty Cells", stats.get('Empty Cells', 0))
    
    # Tabs for different operations
    tab1, tab2, tab3, tab4, tab5 = st.tabs(["📝 Edit Data", "🔧 Operations", "🔍 Search & Filter", "📈 Analytics", "ℹ️ Info"])
    
    with tab1:
        st.subheader("Data Editor")
        
        # Data editor with editing capabilities
        if not df.empty:
            # Convert datetime columns to strings for data editor compatibility
            df_for_editor = df.copy()
            for col in df_for_editor.columns:
                if df_for_editor[col].dtype == 'datetime64[ns]':
                    df_for_editor[col] = df_for_editor[col].dt.strftime('%Y-%m-%d %H:%M:%S')
            
            edited_df = st.data_editor(
                df_for_editor,
                use_container_width=True,
                num_rows="dynamic",
                key=f"editor_{st.session_state.current_sheet}",
                height=500
            )
            
            # Update button
            col1, col2 = st.columns([1, 5])
            with col1:
                if st.button("💾 Save Changes", type="primary"):
                    # Convert string dates back to datetime if original column was datetime
                    edited_df_final = edited_df.copy()
                    for col in edited_df_final.columns:
                        if df[col].dtype == 'datetime64[ns]':
                            try:
                                edited_df_final[col] = pd.to_datetime(edited_df_final[col])
                            except:
                                # If conversion fails, keep as string
                                pass
                    
                    st.session_state.df_dict[st.session_state.current_sheet] = edited_df_final
                    st.success("✅ Changes saved!")
                    st.rerun()
        else:
            st.info("📝 This sheet is empty. Add columns and rows using the Operations tab.")
    
    with tab2:
        st.subheader("Row & Column Operations")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("**Row Operations**")
            
            row_action = st.selectbox(
                "Select Action",
                ["Add Row at End", "Add Row at Start", "Delete Selected Rows"]
            )
            
            if row_action == "Delete Selected Rows":
                if not df.empty:
                    rows_to_delete = st.multiselect(
                        "Select row indices to delete",
                        options=df.index.tolist(),
                        help="Select multiple rows to delete"
                    )
                else:
                    st.info("No rows to delete")
            
            if st.button("Execute Row Operation", key="row_op"):
                if row_action == "Add Row at End":
                    df = add_row(df, 'end')
                    st.session_state.df_dict[st.session_state.current_sheet] = df
                    st.success("✅ Row added at end!")
                    st.rerun()
                elif row_action == "Add Row at Start":
                    df = add_row(df, 'start')
                    st.session_state.df_dict[st.session_state.current_sheet] = df
                    st.success("✅ Row added at start!")
                    st.rerun()
                elif row_action == "Delete Selected Rows":
                    if 'rows_to_delete' in locals() and rows_to_delete:
                        df = delete_rows(df, rows_to_delete)
                        st.session_state.df_dict[st.session_state.current_sheet] = df
                        st.success(f"✅ Deleted {len(rows_to_delete)} row(s)!")
                        st.rerun()
                    else:
                        st.warning("Please select rows to delete")
        
        with col2:
            st.markdown("**Column Operations**")
            
            col_action = st.selectbox(
                "Select Action",
                ["Add New Column", "Delete Selected Columns", "Rename Column"]
            )
            
            if col_action == "Add New Column":
                new_col_name = st.text_input("New Column Name")
            elif col_action == "Delete Selected Columns":
                if not df.empty:
                    cols_to_delete = st.multiselect(
                        "Select columns to delete",
                        options=df.columns.tolist(),
                        help="Select multiple columns to delete"
                    )
                else:
                    st.info("No columns to delete")
            elif col_action == "Rename Column":
                if not df.empty:
                    old_col_name = st.selectbox("Select column to rename", df.columns.tolist())
                    new_col_name = st.text_input("New name")
                else:
                    st.info("No columns to rename")
            
            if st.button("Execute Column Operation", key="col_op"):
                if col_action == "Add New Column":
                    if new_col_name:
                        df = add_column(df, new_col_name)
                        st.session_state.df_dict[st.session_state.current_sheet] = df
                        st.success(f"✅ Column '{new_col_name}' added!")
                        st.rerun()
                    else:
                        st.warning("Please enter a column name")
                elif col_action == "Delete Selected Columns":
                    if 'cols_to_delete' in locals() and cols_to_delete:
                        df = delete_columns(df, cols_to_delete)
                        st.session_state.df_dict[st.session_state.current_sheet] = df
                        st.success(f"✅ Deleted {len(cols_to_delete)} column(s)!")
                        st.rerun()
                    else:
                        st.warning("Please select columns to delete")
                elif col_action == "Rename Column":
                    if 'old_col_name' in locals() and 'new_col_name' in locals() and new_col_name:
                        df = df.rename(columns={old_col_name: new_col_name})
                        st.session_state.df_dict[st.session_state.current_sheet] = df
                        st.success(f"✅ Column renamed from '{old_col_name}' to '{new_col_name}'!")
                        st.rerun()
                    else:
                        st.warning("Please enter a new name")
        
        st.divider()
        
        # Sorting
        st.markdown("**Sorting**")
        if not df.empty:
            col1, col2, col3 = st.columns([2, 2, 1])
            with col1:
                sort_column = st.selectbox("Select column to sort by", df.columns.tolist())
            with col2:
                sort_order = st.selectbox("Order", ["Ascending", "Descending"])
            with col3:
                if st.button("Sort"):
                    ascending = sort_order == "Ascending"
                    df = df.sort_values(by=sort_column, ascending=ascending)
                    st.session_state.df_dict[st.session_state.current_sheet] = df
                    st.success(f"✅ Sorted by {sort_column}!")
                    st.rerun()
    
    with tab3:
        st.subheader("Search & Filter")
        
        if not df.empty:
            # Search
            st.markdown("**Search**")
            search_term = st.text_input("🔍 Search for text in all cells", placeholder="Enter search term...")
            
            if search_term:
                filtered_df = search_in_dataframe(df, search_term)
                st.write(f"Found {len(filtered_df)} matching rows:")
                st.dataframe(filtered_df, use_container_width=True)
            
            st.divider()
            
            # Filter by column
            st.markdown("**Filter by Column**")
            filter_col = st.selectbox("Select column to filter", ["None"] + df.columns.tolist())
            
            if filter_col != "None":
                unique_values = df[filter_col].dropna().unique()
                
                if len(unique_values) > 0:
                    filter_values = st.multiselect(
                        f"Select values from '{filter_col}'",
                        options=sorted(unique_values.astype(str))
                    )
                    
                    if filter_values:
                        filtered_df = df[df[filter_col].astype(str).isin(filter_values)]
                        st.write(f"Filtered to {len(filtered_df)} rows:")
                        st.dataframe(filtered_df, use_container_width=True)
                        
                        # Option to save filtered data
                        if st.button("💾 Save Filtered Data as New Sheet"):
                            new_sheet_name = f"{st.session_state.current_sheet}_filtered_{datetime.now().strftime('%H%M%S')}"
                            st.session_state.df_dict[new_sheet_name] = filtered_df
                            st.success(f"✅ Saved as new sheet: {new_sheet_name}")
                            st.rerun()
        else:
            st.info("No data to search or filter")
    
    with tab4:
        st.subheader("Data Analytics")
        
        if not df.empty:
            numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
            
            if numeric_cols:
                st.markdown("**Numeric Column Statistics**")
                
                selected_numeric_col = st.selectbox("Select numeric column", numeric_cols)
                
                if selected_numeric_col:
                    col1, col2, col3, col4 = st.columns(4)
                    
                    with col1:
                        st.metric("Mean", f"{df[selected_numeric_col].mean():.2f}")
                    with col2:
                        st.metric("Median", f"{df[selected_numeric_col].median():.2f}")
                    with col3:
                        st.metric("Min", f"{df[selected_numeric_col].min():.2f}")
                    with col4:
                        st.metric("Max", f"{df[selected_numeric_col].max():.2f}")
                    
                    # Histogram
                    st.markdown("**Distribution**")
                    fig = px.histogram(df, x=selected_numeric_col, title=f"Distribution of {selected_numeric_col}")
                    st.plotly_chart(fig, use_container_width=True)
                
                st.divider()
                
                # Correlation matrix
                if len(numeric_cols) > 1:
                    st.markdown("**Correlation Matrix**")
                    corr_matrix = df[numeric_cols].corr()
                    fig = px.imshow(
                        corr_matrix,
                        labels=dict(color="Correlation"),
                        x=corr_matrix.columns,
                        y=corr_matrix.columns,
                        title="Correlation Heatmap"
                    )
                    st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("No numeric columns found for analytics")
            
            st.divider()
            
            # Value counts for categorical columns
            st.markdown("**Categorical Analysis**")
            text_cols = df.select_dtypes(include=['object']).columns.tolist()
            
            if text_cols:
                selected_text_col = st.selectbox("Select categorical column", text_cols)
                
                if selected_text_col:
                    value_counts = df[selected_text_col].value_counts().head(10)
                    
                    fig = px.bar(
                        x=value_counts.values,
                        y=value_counts.index,
                        orientation='h',
                        title=f"Top 10 Values in {selected_text_col}",
                        labels={'x': 'Count', 'y': selected_text_col}
                    )
                    st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("No categorical columns found")
        else:
            st.info("No data available for analytics")
    
    with tab5:
        st.subheader("Sheet Information")
        
        if not df.empty:
            # Display detailed statistics
            st.markdown("**Detailed Statistics**")
            
            for key, value in stats.items():
                st.write(f"**{key}:** {value}")
            
            st.divider()
            
            # Column information
            st.markdown("**Column Details**")
            
            col_info = pd.DataFrame({
                'Column': df.columns,
                'Data Type': df.dtypes.values,
                'Non-Null Count': df.count().values,
                'Null Count': df.isnull().sum().values,
                'Unique Values': [df[col].nunique() for col in df.columns]
            })
            
            st.dataframe(col_info, use_container_width=True)
            
            st.divider()
            
            # Preview data
            st.markdown("**Data Preview**")
            preview_rows = st.slider("Number of rows to preview", 5, 50, 10)
            st.dataframe(df.head(preview_rows), use_container_width=True)
        else:
            st.info("Sheet is empty")

    # Welcome screen
    st.markdown("""
    ## Welcome to Excel Viewer & Editor Pro! 🎉
    
    ### Features:
    - 📂 **Load Excel & CSV files** with multiple sheets
    - ✏️ **Edit cells directly** in an interactive table
    - ➕ **Add/Delete rows and columns** dynamically
    - 🔍 **Search and filter** data across sheets
    - 📊 **Analytics and visualizations** for numeric and categorical data
    - 📑 **Multiple sheet management** (add, delete, switch between sheets)
    - 💾 **Export to Excel or CSV** with all your changes
    - 📈 **Statistics and insights** about your data
    - 🎨 **Modern, intuitive UI** for better user experience
    
    ### How to Start:
    1. Upload an Excel or CSV file using the sidebar
    2. Select a sheet to view and edit
    3. Use the tabs to perform various operations
    4. Download your edited file when done
    
    ---
    
    ### Quick Tips:
    - Use the **Edit Data** tab to modify cell values directly
    - Use **Operations** tab for structural changes (rows/columns)
    - **Search & Filter** helps you find specific data quickly
    - **Analytics** provides insights and visualizations
    - All changes are saved in memory until you download the file
    
    **👈 Start by uploading a file in the sidebar!**
    """)
    
    # Sample data option
    st.divider()
    st.subheader("Or try with sample data")
    
    if st.button("📊 Load Sample Dataset"):
        # Create sample data
        sample_data = {
            'Sales': pd.DataFrame({
                'Date': pd.date_range('2024-01-01', periods=50, freq='D'),
                'Product': np.random.choice(['Product A', 'Product B', 'Product C'], 50),
                'Revenue': np.random.randint(100, 1000, 50),
                'Quantity': np.random.randint(1, 20, 50),
                'Region': np.random.choice(['North', 'South', 'East', 'West'], 50)
            }),
            'Employees': pd.DataFrame({
                'Name': [f'Employee {i}' for i in range(1, 21)],
                'Department': np.random.choice(['Sales', 'Marketing', 'Engineering', 'HR'], 20),
                'Salary': np.random.randint(50000, 150000, 20),
                'Experience': np.random.randint(1, 15, 20)
            })
        }
        
        st.session_state.df_dict = sample_data
        st.session_state.current_sheet = 'Sales'
        st.session_state.file_name = 'sample_data.xlsx'
        # Don't add sample data to recent files since it's not a real file
        st.success("✅ Sample data loaded!")
        st.rerun()

# Footer
st.divider()
st.markdown("""
<div style='text-align: center; color: #666; padding: 1rem;'>
    <small>Excel Viewer & Editor Pro | Built with Streamlit 📊</small>
</div>
""", unsafe_allow_html=True)
