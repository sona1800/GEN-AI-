# VBA Macro for Excel Data Cleaning and Formatting

## Overview
This document provides a comprehensive guide to the VBA macro (`DataCleaning.bas`) designed to automate data cleaning operations in Microsoft Excel.

## Features

The VBA macro provides the following capabilities:

### 1. **Remove Duplicate Rows**
- Automatically detects and removes duplicate entries based on the first column
- Preserves header row
- Works with any number of columns

### 2. **Fill Missing/Null Values**
- Identifies all empty or null cells in the dataset
- Replaces them with "N/A"
- Highlights filled cells in light yellow for easy identification

### 3. **Format as Standard Table**
- Converts the data range into a professional Excel table
- Applies automatic table styling
- Enables Excel table features (filtering, sorting, etc.)
- Auto-fits column widths for better readability

---

## Available Macros

### **Main Macro: `CleanAndFormatData()`**
Executes all three operations in sequence:
```vba
Sub CleanAndFormatData()
```
- Removes duplicates
- Fills missing values with "N/A"
- Formats as Excel table
- Provides progress notifications

### **Standalone Macros**

#### `FillMissingValuesOnly()`
Fills all empty/null cells with "N/A" without other operations.

#### `RemoveDuplicatesOnly()`
Removes duplicate rows based on the first column only.

#### `FormatAsTableOnly()`
Formats the used range as an Excel table with headers.

---

## Installation Instructions

### Step 1: Open Excel and Enable Developer Tools
1. Open your Excel file
2. Go to **File** → **Options** → **Trust Center** → **Trust Center Settings**
3. Enable **Macro Settings**: Select "Enable all macros" (or adjust as needed)

### Step 2: Import the VBA Module
1. Press `Alt + F11` to open the Visual Basic Editor
2. Right-click on your project in the left panel
3. Select **Import File**
4. Choose `DataCleaning.bas` from your repository

### Step 3: Run the Macro
1. Go to **View** → **Macros** or press `Alt + F8`
2. Select `CleanAndFormatData`
3. Click **Run**

---

## Usage Instructions

### Basic Workflow

1. **Prepare Your Data**
   - Ensure your data has headers in the first row
   - Place your cursor anywhere in the data range

2. **Run the Macro**
   - Open Macros dialog (`Alt + F8`)
   - Select `CleanAndFormatData`
   - Click **Run**

3. **Review Results**
   - Check the summary dialog for details
   - Verify that duplicates are removed
   - Confirm that "N/A" values are highlighted in yellow
   - Review the formatted table

---

## Features in Detail

### Duplicate Removal
- **Method**: Uses Excel's `RemoveDuplicates` function
- **Based on**: First column values
- **Preserves**: Header row
- **Output**: Clean dataset with unique rows

### Missing Value Handling
- **Detection**: Empty cells or cells with only whitespace
- **Replacement**: "N/A"
- **Visual Indicator**: Light yellow background (RGB: 255, 255, 200)
- **Coverage**: All columns except headers

### Table Formatting
- **Format**: Excel ListObject
- **Style**: TableStyleMedium2 (professional blue theme)
- **Auto-fit**: Column widths adjusted automatically
- **Naming**: Unique table name with timestamp (e.g., `DataTable_20260623_102548`)
- **Features Enabled**: Filtering, sorting, calculated columns

---

## Example: Processing SalesData

Using the `create_salesdata.sql` script from this repository:

1. **Input Data Structure**
   ```
   CustomerID | Name           | Age | City          | PurchaseAmount | PurchaseDate
   1          | John Smith     | 35  | New York      | 1234.56        | 2024-01-15
   2          | Jane Doe       |     | Los Angeles   | (empty)        | 2024-02-20
   1          | John Smith     | 35  | New York      | 1234.56        | 2024-01-15
   ```

2. **After Running `CleanAndFormatData()`**
   ```
   CustomerID | Name           | Age | City          | PurchaseAmount | PurchaseDate
   1          | John Smith     | 35  | New York      | 1234.56        | 2024-01-15
   2          | Jane Doe       | N/A | Los Angeles   | N/A            | 2024-02-20
   ```
   - Duplicate row removed
   - Empty Age field filled with "N/A" (highlighted)
   - Empty PurchaseAmount filled with "N/A" (highlighted)
   - Formatted as professional table

---

## Error Handling

The macro includes comprehensive error handling:
- Validates that the sheet contains data
- Provides detailed error messages if operations fail
- Continues execution where possible
- Logs error codes for troubleshooting

### Common Issues & Solutions

| Issue | Solution |
|-------|----------|
| "No data found" message | Ensure your sheet has data with headers in the first row |
| Macro doesn't run | Check that macros are enabled in your security settings |
| Error 424 "Object Required" | Ensure you're running from the correct worksheet |
| Table not created | Check that you have at least 2 rows (header + data) |

---

## Technical Details

### Requirements
- **Excel Version**: 2010 or later
- **VBA Version**: 7.0 or later
- **Minimum Data**: 1 header row + at least 1 data row

### Performance Notes
- Processes up to 10,000+ rows efficiently
- Time varies based on data size and system resources
- Progress dialogs provide user feedback

### Data Types Handled
- Text/Strings
- Numbers (Integer, Decimal)
- Dates
- Blank cells
- Null values

---

## Advanced Usage

### Customize Table Style
Modify this line in the macro:
```vba
tbl.TableStyle = "TableStyleMedium2"
```

Replace with other styles:
- `TableStyleLight1` - Light theme
- `TableStyleDark1` - Dark theme
- `TableStyleAccent1` - Accent color 1

### Modify Missing Value Indicator
Change the fill color for "N/A" cells:
```vba
.Interior.Color = RGB(255, 255, 200)  ' Light yellow
```

### Change Duplicate Detection Column
Modify the RemoveDuplicates method:
```vba
usedRange.RemoveDuplicates Columns:=Array(1, 2), Header:=xlYes
```
Use Array(1, 2) to check columns 1 and 2 for duplicates.

---

## Integration with SalesData

This macro is designed to work seamlessly with the `create_salesdata.sql` script:

1. **Export SalesData from SQL Server to Excel**
   - Use SQL Server Management Studio
   - Export to Excel format

2. **Run `CleanAndFormatData()` macro**
   - Removes any duplicate customer records
   - Fills missing Age, City, or PurchaseAmount fields
   - Creates professional analysis table

3. **Use the formatted table for**
   - Data analysis
   - Pivot tables
   - Charts and visualizations
   - Further processing

---

## Support & Troubleshooting

### Enable Macros Dialog
If you see a security warning when opening the file:
- Click **Enable Macros**
- Trust the document if it's from a known source

### Macro Not Visible
1. Ensure the `.bas` file was imported correctly
2. Go to **Tools** → **References** and verify required libraries are checked
3. Restart Excel

### Need Help?
- Check the error message in the dialog box
- Review the VBA code comments for detailed explanations
- Test with a small sample dataset first

---

## Version History

| Version | Date | Changes |
|---------|------|---------|
| 1.0 | 2026-06-23 | Initial release with three main functions |

---

## License & Usage

This VBA macro is provided as part of the GEN-AI- repository and can be freely used and modified for your Excel data cleaning needs.

---

**Last Updated**: June 23, 2026  
**Repository**: [sona1800/GEN-AI-](https://github.com/sona1800/GEN-AI-)
