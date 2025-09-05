# Quick Start Guide - 5 Minutes to Running

🚀 **Get the boleto automation working in 5 minutes!**

## Step 1: Setup (2 minutes)
```bash
chmod +x setup.sh
./setup.sh
```

## Step 2: Test (1 minute)  
```bash
python debug_submitfunction_test.py
```
*Should show "✅ POPUP NAVIGATED TO" and real content*

## Step 3: Run (2 minutes)
```bash
./run_final_solution.sh --max-records 2 --batch-size 1
```

## ✅ Success Indicators
```
🔧 FINAL SOLUTION: Loading boleto content via main page submitFunction
✅ POPUP NAVIGATED TO: https://consweb.hsconsorcios.com.br/Slip/Slip.asp
✅ POPUP CONTENT LENGTH: 15234 characters
✅ PDF SUCCESS: boleto_file.pdf (45678 bytes)
```

## ❌ If Still Not Working
1. Check your internet connection
2. Verify credentials in `config.yaml` (3090A / SuzanaAdm25@)
3. Check `complete_fixed_automation.log` for errors
4. Look at screenshots in `screenshots/` folder

## 🎯 Production Commands
```bash
# Small test
./run_final_solution.sh --max-records 10

# Full automation  
./run_final_solution.sh --batch-size 5
```

**That's it! The popup content loading issue should be FIXED!** 🎉
