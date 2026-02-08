# Quick Start Guide

## Step 1: Install Dependencies
```bash
pip install -r requirements.txt
```

## Step 2: Configure SAP Connection
```bash
# Copy the template
cp config_template.ini config.ini

# Edit config.ini with your SAP credentials
nano config.ini  # or use any text editor
```

## Step 3: Prepare Your Data
Edit `sample_data/material_master_template.csv` or create your own CSV file with the required columns.

## Step 4: Validate Your Data
```bash
python examples/validate_data.py
```

## Step 5: Run the Automation

### Option A: Using SAP GUI Scripting (Windows only)
```bash
# Make sure SAP GUI is running and you're logged in
python src/material_master.py sample_data/material_master_template.csv --method gui
```

### Option B: Using RFC API
```bash
python src/material_master.py sample_data/material_master_template.csv --method rfc
```

## Step 6: Check Results
- View console output for summary
- Check `logs/material_master.log` for detailed logs
- Review `logs/results_*.csv` for individual record results

## Common Issues

### Issue: "win32com not available"
**Solution**: Install pywin32
```bash
pip install pywin32
```

### Issue: "pyrfc not available"
**Solution**: Install SAP RFC SDK and pyrfc
1. Download SAP RFC SDK from SAP website
2. Install according to your OS
3. Install pyrfc: `pip install pyrfc`

### Issue: "SAP GUI session not connected"
**Solution**: 
1. Open SAP GUI and log in manually
2. Keep the session open
3. Enable scripting in SAP GUI options
4. Run the script again

### Issue: "Validation failed"
**Solution**: 
1. Run the validation script first: `python examples/validate_data.py`
2. Fix the errors reported
3. Try processing again

## Next Steps

1. Test with a small subset of data first
2. Always test in a development/sandbox environment
3. Review logs after each run
4. Customize field mappings in the code as needed for your SAP system
5. Consider adding more validations specific to your business rules

## Support

For more information, see the main [README.md](README.md) file.
