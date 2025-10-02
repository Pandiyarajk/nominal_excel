#!/usr/bin/env python3
"""
Main entry point for Excel Viewer & Editor Pro
Starts the Streamlit application with proper configuration
"""

import subprocess
import sys
import os
from pathlib import Path

def check_requirements():
    """Check if required packages are installed"""
    try:
        import streamlit
        import pandas
        import openpyxl
        import plotly
        import numpy
        return True
    except ImportError as e:
        print(f"❌ Missing required package: {e}")
        print("📦 Installing required packages...")
        return False

def install_requirements():
    """Install required packages from requirements.txt"""
    try:
        requirements_file = Path(__file__).parent / "requirements.txt"
        if requirements_file.exists():
            subprocess.check_call([
                sys.executable, "-m", "pip", "install", "-r", str(requirements_file)
            ])
            print("✅ Installation complete!")
            return True
        else:
            print("❌ requirements.txt not found!")
            return False
    except subprocess.CalledProcessError as e:
        print(f"❌ Failed to install requirements: {e}")
        return False

def start_streamlit():
    """Start the Streamlit application"""
    try:
        print("🚀 Starting Excel Viewer & Editor Pro...")
        print("================================================")
        print("")
        print("🌐 Starting application...")
        print("📊 Open your browser at: http://localhost:8501")
        print("")
        print("Press Ctrl+C to stop the server")
        print("================================================")
        print("")
        
        # Run streamlit with custom config
        subprocess.run([
            sys.executable, "-m", "streamlit", "run", "excel_app.py",
            "--server.headless=true",
            "--server.address=0.0.0.0",
            "--server.port=8501",
            "--browser.gatherUsageStats=false"
        ])
    except KeyboardInterrupt:
        print("\n👋 Application stopped by user")
    except Exception as e:
        print(f"❌ Error starting application: {e}")

def main():
    """Main function"""
    print("🚀 Starting Excel Viewer & Editor Pro...")
    print("================================================")
    print("")
    
    # Check if requirements are installed
    if not check_requirements():
        if not install_requirements():
            print("❌ Failed to install requirements. Please install manually:")
            print("   pip install -r requirements.txt")
            sys.exit(1)
    
    # Start the Streamlit application
    start_streamlit()

if __name__ == "__main__":
    main()

