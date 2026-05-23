# google-sheet-automated-code-for-syncing-inventory-and-also-have-all-avlidation
# Google Sheets Order Management System (OMS)

A powerful Google Apps Script designed to turn Google Sheets into a fully automated, multi-stage Order Management System. It streamlines order intake, validates data in real-time, calculates multi-currency financials, maps SKUs to inventory, and pushes orders through a custom fulfillment workflow.

## 🚀 Key Features

* **Real-Time Data Validation**: Automatically highlights missing mandatory fields in red and logs specific errors to guide data entry.
* **Automated Internal Order IDs**: Generates sequential order numbers based on dynamic brand prefixes configured by the user (fills gaps upon row deletion).
* **Inventory & SKU Management**: Cross-references Portal SKUs with Vendor SKUs via a Matching Table to pull live product costs and available stock.
* **Financial Calculations**: Automatically fetches live currency conversion rates (via API) to calculate Price in INR, Shipping Charges, Maximum/Actual Expenses, and Estimated Profit.
* **Multi-Stage Fulfillment Workflow**: Safely pushes verified orders from `New_Orders` to `All Orders`, and syncs pending orders through fulfillment queues (`stockcheckRTS` -> `CreateLabels` -> `Dispatched`).
* **Duplicate Protection**: Prevents duplicate order entries into the master database and provides safe, partial-push overrides via styled HTML modal popups.
* **Bulk-Edit Aware**: The `onEdit` trigger is heavily optimized to safely handle copy/pasting multiple rows at once without timing out.
* **Secure UI Alerts**: Custom-styled HTML dialogs built with XSS protection for safe user feedback.

## 📑 Sheet Structure Requirements

For this script to function correctly, your Google Spreadsheet ecosystem must contain the following sheets (exact names required unless altered in the `SHEET_NAMES` constant):

1. **`New_Orders`**: The primary intake sheet for manual entry or pasted orders.
2. **`All Orders`**: The master database where valid orders are permanently pushed.
3. **`Config`**: Defines dropdown options, Brand Prefixes, and Country Magic Numbers.
4. **`stockcheckRTS`**: Queue for checking physical stock status.
5. **`CreateLabels`**: Queue for generating shipping labels (skipped automatically for domestic/India orders).
6. **`Dispached`**: Final queue marking orders as shipped.
7. **`MatchingTable`**: Maps `PortalSKU` (sales channel SKU) to `VendorSKU` (internal SKU). *(Can be in a separate Catalog spreadsheet).*
8. **`Inventory`**: Master stock sheet tracking `Stock`, `Product_Cost`, and `Location` by `VendorSKU`. *(Can be in a separate Catalog spreadsheet).*

## 🛠️ Setup & Installation

1. Open your main Google Spreadsheet.
2. Navigate to **Extensions > Apps Script**.
3. Clear any existing code and paste the provided `new_order.gs` script.
4. **Configure Spreadsheet IDs**:
   At the top of the script, locate the following variables and replace them with your actual Google Sheet IDs (found in the URL of your sheets):
```javascript
   var SHIPPING_TARGET_SPREADSHEET_ID = "YOUR_SHIPPING_SPREADSHEET_ID_HERE";
   var CATALOG_SPREADSHEET_ID = "YOUR_CATALOG_SPREADSHEET_ID_HERE";
