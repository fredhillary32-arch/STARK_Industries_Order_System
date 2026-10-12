# STARK Industries Order Support

STARK Industries Order Support is a Windows desktop application with a browser-based interface for handling customer order inquiries. It checks order status and delivery dates, applies the configured support rules for late or missing orders, refunds, exchanges, damaged or incorrect items, and cancellations, then prepares a customer-ready email. It validates edited email text against the calculated order details and previews replacement orders before saving them to the workbook.

**Run the app**

1. Install Python 3.10 or newer from https://www.python.org/downloads/ and enable **Add Python to PATH** during setup.
2. Download or clone this repository.
3. Double-click `Start_Order_Support_App.bat`. On first launch, it creates a local `.venv` environment and installs the required packages. The app opens at `http://localhost:8501`.
4. Enter the order number, inquiry type, and customer contact date. Review the status, recommended action, and email. Saving a replacement or exchange requires confirmation.

**Sample data and privacy**

The repository includes `data/stark_orders_demo.xlsx`, which uses demonstration customer details. The app uses `data/stark_orders_clean.xlsx` instead when that local file is present. Keep real customer workbooks and secrets out of GitHub; `.gitignore` excludes local `.xlsx` files other than the demo workbook.

**Tests**

Double-click `Run_Tests.bat` to run the engine tests, then run the interface tests with:

```bat
.venv\Scripts\python.exe tests\test_ui_adapter.py
```

The optional browser check uses Microsoft Edge or Google Chrome:

```bat
.venv\Scripts\python.exe tests\browser_check.py
```
