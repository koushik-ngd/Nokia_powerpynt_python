# 🐍 Python Installation Guide for Nokia Presentation Generator

## Step 1: Install Python

### Download Python
1. Go to [https://www.python.org/downloads/](https://www.python.org/downloads/)
2. Click "Download Python 3.x.x" (latest version)
3. Run the downloaded installer

### Important Installation Settings
⚠️ **CRITICAL**: During installation, make sure to:
- ✅ Check "Add Python to PATH" (very important!)
- ✅ Check "Install pip" 
- ✅ Choose "Install for all users" (recommended)

### Verify Installation
Open Command Prompt (cmd) and type:
```bash
python --version
```
You should see something like: `Python 3.11.x`

If you get "Python was not found", restart your computer and try again.

## Step 2: Install Required Libraries

Open Command Prompt and run:
```bash
pip install python-pptx matplotlib numpy Pillow
```

Or use the requirements file:
```bash
pip install -r requirements.txt
```

## Step 3: Run the Presentation Generator

### Option A: Full Version (with charts)
```bash
python nokia_presentation_generator.py
```

### Option B: Simple Version (text only)
```bash
python nokia_simple_generator.py
```

### Option C: Use Batch File (Windows)
Double-click: `run_presentation_generator.bat`

## 🔧 Troubleshooting

### Problem: "Python was not found"
**Solution**: 
1. Reinstall Python with "Add to PATH" checked
2. Restart your computer
3. Try again

### Problem: "pip is not recognized"
**Solution**:
```bash
python -m pip install python-pptx matplotlib numpy Pillow
```

### Problem: Permission denied
**Solution**: Run Command Prompt as Administrator

### Problem: Module not found
**Solution**: Make sure all dependencies are installed:
```bash
pip list
```
Should show: python-pptx, matplotlib, numpy, Pillow

## 📁 What Gets Created

After running successfully, you'll get:
- `Nokia_Failure_Analysis_PowerPynt.pptx` (full version with charts)
- `Nokia_Failure_Analysis_Simple.pptx` (simple text version)

## 🎯 Quick Start for Beginners

1. **Install Python** from python.org (check "Add to PATH"!)
2. **Restart computer**
3. **Open Command Prompt**
4. **Navigate to this folder**:
   ```bash
   cd "C:\Users\aayan\OneDrive\Documents\Python PPT"
   ```
5. **Install libraries**:
   ```bash
   pip install python-pptx
   ```
6. **Run simple version**:
   ```bash
   python nokia_simple_generator.py
   ```

## 🆘 Still Having Issues?

### Alternative: Use Online Python
If local installation fails, you can:
1. Copy the code to [replit.com](https://replit.com)
2. Create a new Python project
3. Install dependencies in the shell
4. Run the script online

### Alternative: Use Anaconda
1. Download [Anaconda](https://www.anaconda.com/products/distribution)
2. Install Anaconda (includes Python + libraries)
3. Open Anaconda Prompt
4. Run the scripts from there

---

**Need Help?** 
- Check Python version: `python --version`
- Check pip: `pip --version`
- List installed packages: `pip list`
- Upgrade pip: `python -m pip install --upgrade pip`