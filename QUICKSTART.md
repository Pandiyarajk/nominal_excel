# 🚀 Quick Start Guide

## Getting Started in 3 Steps

### 1️⃣ Install Dependencies
```bash
pip install -r requirements.txt
```

Or if you're on Linux/Ubuntu with externally managed Python:
```bash
pip install --break-system-packages -r requirements.txt
```

### 2️⃣ Run the Application
```bash
streamlit run excel_app.py
```

Or use the provided script:
```bash
./run_app.sh
```

### 3️⃣ Open Your Browser
The app will automatically open at: **http://localhost:8501**

---

## 📊 Try Sample Data

The repository includes 3 sample Excel files ready to use:

1. **sales_data.xlsx** - 100 days of sales transactions
2. **employee_data.xlsx** - 50 employee records with department info
3. **inventory_data.xlsx** - 200 products with stock levels

Upload any of these files to explore all features!

---

## 🎯 First-Time Usage

### Upload a File
1. Click **"Browse files"** in the left sidebar
2. Select an Excel (.xlsx, .xls) or CSV file
3. Wait for it to load (you'll see a success message)

### View & Edit Data
1. Select a **sheet** from the dropdown (if multiple sheets exist)
2. Go to **"Edit Data"** tab
3. Click any cell to edit it
4. Click **"Save Changes"** to apply your edits

### Add/Remove Rows & Columns
1. Go to **"Operations"** tab
2. Choose your action (add/delete rows or columns)
3. Click the execute button
4. Changes apply immediately

### Search & Filter
1. Go to **"Search & Filter"** tab
2. Enter text to search across all cells
3. Or filter by specific column values
4. Save filtered results as a new sheet

### Analyze Your Data
1. Go to **"Analytics"** tab
2. Select a numeric column to see statistics
3. View distribution histograms
4. Explore correlation heatmaps (if multiple numeric columns)
5. Analyze categorical data with bar charts

### Export Your Work
1. In the left sidebar, click **"Download Excel"**
2. Or download individual sheets as CSV
3. Your file includes ALL sheets with ALL changes

---

## 💡 Pro Tips

- **Undo Changes**: If you haven't clicked "Save Changes" yet, just refresh the page
- **Multiple Edits**: Make all your edits first, then click "Save Changes" once
- **Large Files**: Files over 10MB may take a few seconds to load
- **Keyboard Navigation**: Use Tab and Enter to navigate cells in the editor
- **Quick Sort**: Use the Operations tab to sort by any column instantly
- **Data Backup**: Always download a copy before making major changes

---

## 🐛 Troubleshooting

**App won't start?**
- Make sure all dependencies are installed
- Check that port 8501 is not already in use
- Try `streamlit run excel_app.py --server.port 8502`

**File won't upload?**
- Verify the file format is .xlsx, .xls, or .csv
- Check the file isn't corrupted
- Try with one of the sample files first

**Changes not saving?**
- Make sure you clicked "Save Changes" button
- Check browser console for errors (F12)
- Try refreshing and re-uploading the file

**Performance issues?**
- Large files (>50,000 rows) may be slow
- Try filtering data to work with smaller subsets
- Close other browser tabs to free up memory

---

## 📚 Learn More

For detailed documentation, see [README.md](README.md)

For technical details about the implementation, check the comments in `excel_app.py`

---

## 🆘 Need Help?

Common questions answered:

**Q: Can I edit formulas?**  
A: Currently, formulas are calculated when you load the file, but you can edit the resulting values.

**Q: Does it preserve formatting?**  
A: The app focuses on data editing. Some formatting may be lost when saving.

**Q: Can I work with very large files?**  
A: Yes, but performance depends on your system. Files with 10,000+ rows may be slower.

**Q: Is my data secure?**  
A: Everything runs locally on your machine. No data is sent to external servers.

**Q: Can multiple people edit simultaneously?**  
A: Not currently. This is a single-user desktop app.

---

**Happy Editing! 📊✨**
