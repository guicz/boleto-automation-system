#!/bin/bash
# Final Boleto Solution Setup Script

set -e

echo "🚀 Setting up Final Boleto Solution Package..."
echo "🔧 This package includes the COMPLETE FIX for popup content loading"

# Check Python version
python_version=$(python3 --version 2>&1 | awk '{print $2}' | cut -d. -f1,2)
required_version="3.8"

if [ "$(printf '%s\n' "$required_version" "$python_version" | sort -V | head -n1)" != "$required_version" ]; then
    echo "❌ Python 3.8+ is required. Current version: $python_version"
    exit 1
fi

echo "✅ Python version check passed: $python_version"

# Create virtual environment
echo "📦 Creating virtual environment..."
if [ -d "venv" ]; then
    echo "Virtual environment already exists, removing old one..."
    rm -rf venv
fi

python3 -m venv venv
source venv/bin/activate

# Upgrade pip
echo "⬆️ Upgrading pip..."
pip install --upgrade pip

# Install Python dependencies
echo "📚 Installing Python dependencies..."
pip install -r requirements.txt

# Install Playwright browsers
echo "🌐 Installing Playwright browsers..."
playwright install chromium

# Install system dependencies for Playwright
echo "🔧 Installing system dependencies..."
if command -v apt-get &> /dev/null; then
    echo "Detected apt-get, installing Debian/Ubuntu dependencies..."
    sudo apt-get update -qq
    sudo apt-get install -y -qq \
        libnss3 \
        libnspr4 \
        libatk-bridge2.0-0 \
        libdrm2 \
        libxkbcommon0 \
        libxcomposite1 \
        libxdamage1 \
        libxrandr2 \
        libgbm1 \
        libxss1 \
        libasound2 \
        fonts-liberation \
        libappindicator3-1 \
        xdg-utils
elif command -v yum &> /dev/null; then
    echo "Detected yum, installing RHEL/CentOS dependencies..."
    sudo yum install -y -q \
        nss \
        nspr \
        at-spi2-atk \
        libdrm \
        libxkbcommon \
        libXcomposite \
        libXdamage \
        libXrandr \
        mesa-libgbm \
        libXScrnSaver \
        alsa-lib
else
    echo "⚠️ Could not detect package manager, skipping system dependencies"
    echo "You may need to install browser dependencies manually"
fi

# Create necessary directories
echo "📁 Creating directories..."
mkdir -p downloads reports screenshots temp logs

# Set permissions
echo "🔐 Setting permissions..."
chmod +x final_working_boleto_processor.py
chmod +x debug_submitfunction_test.py
chmod +x test_final_solution.py
chmod +x run_final_solution.sh

echo ""
echo "✅ Setup completed successfully!"
echo ""
echo "🎯 FINAL SOLUTION READY!"
echo ""
echo "📋 Next steps:"
echo "1. 🔍 Debug test: python debug_submitfunction_test.py"
echo "2. 🧪 Single test: python test_final_solution.py" 
echo "3. 🚀 Production:  ./run_final_solution.sh --max-records 5"
echo ""
echo "🔧 Manual activation:"
echo "source venv/bin/activate"
echo "python final_working_boleto_processor.py controle_boletos_hs.xlsx --max-records 5"
echo ""
echo "🎉 The popup content loading issue has been FIXED!"
