#!/bin/bash
# Final Boleto Solution Runner

# Check if virtual environment exists
if [ ! -d "venv" ]; then
    echo "❌ Virtual environment not found. Please run ./setup.sh first"
    exit 1
fi

# Activate virtual environment
source venv/bin/activate

# Check if Excel file exists
if [ ! -f "controle_boletos_hs.xlsx" ]; then
    echo "❌ Excel file not found: controle_boletos_hs.xlsx"
    exit 1
fi

echo "🚀 Starting Final Boleto Solution..."
echo "🔧 INCLUDES COMPLETE FIX for popup content loading!"
echo "📊 Processing file: controle_boletos_hs.xlsx"
echo "⏰ Started at: $(date)"
echo ""

# Check if arguments provided
if [ $# -eq 0 ]; then
    echo "📋 No arguments provided. Running with default settings:"
    echo "   --max-records 10 --batch-size 1"
    echo ""
    python final_working_boleto_processor.py controle_boletos_hs.xlsx --max-records 10 --batch-size 1
else
    echo "📋 Running with custom arguments: $@"
    echo ""
    python final_working_boleto_processor.py controle_boletos_hs.xlsx "$@"
fi

echo ""
echo "⏰ Completed at: $(date)"
echo ""
echo "📁 Check results:"
echo "   - downloads/     → Your boleto files"
echo "   - reports/       → Processing results"
echo "   - screenshots/   → Popup captures"
echo "   - complete_fixed_automation.log → Detailed logs"
echo ""
echo "🎯 Quick commands for next runs:"
echo "   - Small test:  ./run_final_solution.sh --max-records 5 --batch-size 1"
echo "   - Full run:    ./run_final_solution.sh --batch-size 5"
echo "   - Debug mode:  ./run_final_solution.sh --max-records 2 --debug"
