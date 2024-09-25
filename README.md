# nr-housing-data-validation
### Overview

The script is designed to validate data in an Excel worksheet for housing dashboard from BC housing dataset. It ensures data integrity before the data is used for further processing into a database.

### Getting Started

#### Adding the Script to Excel

1. Open your workbook in Excel on the web.
2. Click on the **Automate** tab in the ribbon.
3. Click on **All Scripts** to open the Code Editor.
4. Click **New Script** to create a new Office Script.
5. Copy the entire script from the `script.ts` file in this repository.
6. Paste the script into the Code Editor, replacing any existing code.
7. Make sure the tab name is `Sheet1` or you may need to change `"Sheet1"` (const sheet = workbook.getWorksheet("Sheet1")) to your Excel tab name.

#### Running the Script

1. Click **Run**.

### Output

- If errors are found: The console will display detailed error messages.
- If no errors are found: The console will display "Great! Data validated. Ready to use."
