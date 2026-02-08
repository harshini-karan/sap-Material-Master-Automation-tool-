# SAP Material Master Automation Tool

A Python-based automation tool for creating, updating, and validating SAP Material Master records in bulk using SAP GUI Scripting or RFC API.

## 🚀 Features

- **Bulk Material Creation**: Process multiple materials from Excel or CSV files
- **Dual Methods**: Support for both SAP GUI Scripting and RFC API
- **Data Validation**: Automatic validation of mandatory fields before posting
- **Error Handling**: Comprehensive logging of success and failures
- **Flexible Configuration**: Easy-to-use configuration file for SAP connection settings
- **Detailed Logging**: Track all operations with timestamps and status

## 📋 Prerequisites

- Python 3.7 or higher
- SAP GUI (for GUI scripting method)
- SAP RFC SDK (for RFC API method)
- Valid SAP credentials with MM01 authorization

## 🔧 Installation

1. **Clone the repository:**
```bash
git clone https://github.com/harshini-karan/sap-Material-Master-Automation-tool-.git
cd sap-Material-Master-Automation-tool-
```

2. **Install required Python packages:**
```bash
pip install -r requirements.txt
```

3. **Configure SAP connection:**
```bash
# Copy the template configuration file
cp config_template.ini config.ini

# Edit config.ini with your SAP credentials
# Note: config.ini is in .gitignore for security
```

## ⚙️ Configuration

Edit `config.ini` with your SAP system details:

```ini
[SAP]
sap_system = PRD
sap_client = 100
sap_user = YOUR_USERNAME
sap_password = YOUR_PASSWORD
sap_language = EN

[RFC]
ashost = sap-server.company.com
sysnr = 00
client = 100
user = YOUR_USERNAME
passwd = YOUR_PASSWORD
lang = EN

[Automation]
transaction_code = MM01
delay_between_actions = 0.5
screenshot_on_error = True
max_retries = 3
```

## 📊 Input File Format

The tool accepts CSV or Excel files with the following columns:

| Column Name | Required | Description | Example |
|------------|----------|-------------|---------|
| Material_Number | No | Material number (leave empty for auto-generation) | |
| Material_Type | Yes | Type of material | FERT, ROH, HALB |
| Industry_Sector | Yes | Industry sector | M, C, P, A |
| Description | Yes | Material description | Finished Product Example |
| Base_Unit | Yes | Base unit of measure | EA, KG, PC |
| Material_Group | No | Material group | 1000 |
| Plant | No | Plant code | 1000 |
| Storage_Location | No | Storage location | 0001 |
| Valuation_Class | No | Valuation class | 3000 |
| Price | No | Material price | 100.00 |
| Currency | No | Currency code | USD |

See `sample_data/material_master_template.csv` for a complete example.

## 🎯 Usage

### Method 1: Using SAP GUI Scripting

```bash
# Basic usage
python src/material_master.py sample_data/material_master_template.csv --method gui

# With custom config file
python src/material_master.py your_data.csv --method gui --config your_config.ini
```

### Method 2: Using RFC API

```bash
python src/material_master.py sample_data/material_master_template.csv --method rfc
```

### Example Scripts

The `examples/` directory contains ready-to-use scripts:

1. **Validate Data Before Processing:**
```bash
python examples/validate_data.py
```

2. **Bulk Create with GUI:**
```bash
python examples/bulk_create_gui.py
```

3. **Bulk Create with RFC:**
```bash
python examples/bulk_create_rfc.py
```

## 📝 Validation Rules

The tool automatically validates:

- ✅ Mandatory fields (Material_Type, Industry_Sector, Description, Base_Unit)
- ✅ Valid material types (FERT, ROH, HALB, HAWA, VERP)
- ✅ Valid industry sectors (M, C, P, A)
- ✅ Base unit length (max 3 characters)
- ✅ Price values (must be numeric and non-negative)

## 📈 Output and Logging

### Console Output
The tool provides real-time progress updates and a summary:
```
==================================================
PROCESSING SUMMARY
==================================================
Total Records: 4
Successful: 3
Failed: 1
Status: completed
==================================================
```

### Log Files
- **Main Log**: `logs/material_master.log` - Detailed operation log
- **Results CSV**: `logs/results_YYYYMMDD_HHMMSS.csv` - Processing results for each record

## 🏗️ Project Structure

```
sap-Material-Master-Automation-tool-/
├── src/
│   └── material_master.py       # Main automation module
├── examples/
│   ├── bulk_create_gui.py       # GUI method example
│   ├── bulk_create_rfc.py       # RFC method example
│   └── validate_data.py         # Validation example
├── sample_data/
│   └── material_master_template.csv  # Sample input file
├── logs/                        # Log files (auto-created)
├── config_template.ini          # Configuration template
├── requirements.txt             # Python dependencies
├── .gitignore                   # Git ignore rules
└── README.md                    # This file
```

## 🔒 Security Notes

- **Never commit** `config.ini` with real credentials to version control
- The `.gitignore` file is configured to exclude sensitive files
- Use environment variables or secure vaults for production deployments
- Regularly rotate SAP passwords

## 🛠️ Troubleshooting

### SAP GUI Scripting Issues

1. **Enable SAP GUI Scripting:**
   - In SAP GUI: Options → Accessibility & Scripting → Scripting
   - Check "Enable scripting"
   - Uncheck "Notify when a script opens a connection"

2. **COM Object Error:**
   - Ensure SAP GUI is running before executing the script
   - Verify win32com is installed: `pip install pywin32`

### RFC Connection Issues

1. **pyrfc not found:**
   ```bash
   pip install pyrfc
   ```
   
2. **Connection Failed:**
   - Verify SAP RFC SDK is installed
   - Check network connectivity to SAP server
   - Validate credentials in config.ini

### General Issues

- Check log files in `logs/` directory for detailed error messages
- Ensure input file encoding is UTF-8
- Verify SAP user has MM01 authorization

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📄 License

This project is provided as-is for educational and automation purposes.

## 👥 Authors

- Harshini Karan

## 🙏 Acknowledgments

- SAP Community for GUI scripting documentation
- pyrfc library maintainers
- All contributors to this project

## 📞 Support

For issues, questions, or contributions, please open an issue on GitHub.

---

**Note:** This tool is designed for SAP MM module automation. Always test in a development/sandbox environment before using in production.