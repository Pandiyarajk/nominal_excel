#!/bin/bash
# Quick start script for Excel Viewer & Editor Pro

echo "🚀 Starting Excel Viewer & Editor Pro..."
echo "================================================"
echo ""

# Check if streamlit is installed
if ! command -v streamlit &> /dev/null && ! python3 -c "import streamlit" &> /dev/null
then
    echo "⚠️  Streamlit is not installed."
    echo "📦 Installing required packages..."
    echo ""
    pip3 install -q -r requirements.txt --break-system-packages 2>&1 | grep -v "WARNING"
    echo "✅ Installation complete!"
    echo ""
fi

echo "🌐 Starting application..."
echo "📊 Open your browser at: http://localhost:8501"
echo ""
echo "Press Ctrl+C to stop the server"
echo "================================================"
echo ""

# Run streamlit with custom config
python3 -m streamlit run excel_app.py \
    --server.headless=true \
    --server.address=0.0.0.0 \
    --server.port=8501 \
    --browser.gatherUsageStats=false
