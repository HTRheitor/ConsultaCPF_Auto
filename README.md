# Payment Verification Automation

## Important Notice
**WARNING:** The spreadsheet files included in this repository contain **FICTIONAL** data (names, CPF/ID numbers, amounts, and dates) created for demonstration purposes only. All real client data has been removed to preserve privacy.

## About the Project
This project automates the process of verifying customer payment status through integration with a corporate web system. The automation reads customer data from an Excel spreadsheet, individually queries each ID number in the web system, and records the results in a closing report spreadsheet.

### Features
- Automatic reading of customer data from an Excel spreadsheet
- Automated form filling in the web system
- Payment status verification (paid or pending)
- For confirmed payments, collection of:
  - Payment date
  - Payment method (credit card, bank slip, etc.)
- Report generation in a closing spreadsheet

## Requirements
- Python 3.6 or higher
- Libraries:
  - openpyxl
  - selenium
  - Chrome browser and compatible ChromeDriver

## File Structure
- `app.py` - Main automation script
- `dados_clientes.xlsx` - Spreadsheet with customer information (FICTIONAL DATA)
- `planilha fechamento.xlsx` - Output spreadsheet with verification results

## How to Use
1. Make sure the Excel files are in the same directory as the script
2. Run the script:
```
python app.py
```
3. The script will open Chrome, query each customer and fill in the closing spreadsheet
