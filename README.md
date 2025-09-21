# krutidev-sheets-helper
Many official documents from government offices in India are still prepared and circulated in **Kruti Dev** format. In practice, this creates a challenge because while most modern systems and applications (including Google Sheets) use **Unicode (Mangal/Arial Unicode MS)**, government offices often expect the data to remain in Kruti Dev. Personally, I have faced this issue many times — whenever filling cells in Google Sheets, I had to manually convert content from Unicode to Kruti Dev or vice versa, just to match the required format. This not only wastes time but also increases the risk of mistakes.

This Google Sheets Add-on was built to solve that problem. It provides a robust, automated solution for converting text between **Kruti Dev** (a popular legacy non-Unicode Hindi font) and **Unicode**, directly within Google Sheets. With it, you can seamlessly type, edit, and export Hindi text in whichever format is required — whether for modern usage (Unicode) or for legacy compliance (Kruti Dev).

**This Add-on solves this incompatibility by:**

1.  **Converting Kruti Dev text (from your legacy Excel files) into readable Unicode Hindi within Google Sheets.** This allows you to view and edit the text correctly.
2.  **Converting Unicode Hindi text back into Kruti Dev format.** This is useful for exporting data that needs to be used in systems or applications that still rely on Kruti Dev fonts.

## Features

*   **Bidirectional Conversion:** Convert text from Kruti Dev to Unicode and vice-versa.
*   **Selected Cell Conversion:** Easily convert text in selected cells or ranges within your active Google Sheet.
*   **Direct Text Conversion:** A sidebar interface to quickly type or paste text for on-the-fly conversion.
*   **Kruti Dev CSV Download:** Export your sheet data as a CSV file with text converted to Kruti Dev, which can then be opened in Excel and formatted with a Kruti Dev font.
*   **Integrated Print Functionality:** Print your sheet directly from the sidebar (note: for Kruti Dev font appearance in print, convert to Unicode first, or print the exported Kruti Dev CSV from Excel).

## How to Install and Use

### Installation Steps

1.  **Open your Google Sheet:** Go to [sheets.google.com](https://sheets.google.com/) and open an existing sheet or create a new one.
2.  **Open Apps Script Editor:** Go to `Extensions > Apps Script`. This will open the Apps Script editor in a new browser tab.
3.  **Create New Script Files:**
    *   In the Apps Script editor, you'll see a default `Code.gs` file.
    *   Click `File > New > Script file` and name it `Sidebar.html`.
    *   Click `File > New > Script file` again and name it `KrutiDevConversionLibrary.gs`.
4.  **Copy and Paste Code:**
    *   **`Code.gs`:** Copy the content from the `Code.gs` file in this repository and paste it into your `Code.gs` file in the Apps Script editor.
    *   **`Sidebar.html`:** Copy the content from the `Sidebar.html` file in this repository and paste it into your `Sidebar.html` file.
    *   **`KrutiDevConversionLibrary.gs`:** Copy the content from the `KrutiDevConversionLibrary.gs` file in this repository and paste it into your `KrutiDevConversionLibrary.gs` file.
5.  **Save Project:** Click the floppy disk icon (Save project) or go to `File > Save project` in the Apps Script editor. You can name your project (e.g., "Kruti Dev Helper").
6.  **Refresh Google Sheet:** Go back to your Google Sheet tab and refresh the page.
7.  **Authorize Script:**
    *   Go to `Extensions > Kruti Dev Tools > Open Helper Sidebar`.
    *   A dialog "Authorization required" will appear. Click "Review permissions."
    *   Select your Google account.
    *   Click "Allow" on the next screen to grant the necessary permissions (e.g., to view and manage your spreadsheets).
    *   You may need to go to `Extensions > Kruti Dev Tools > Open Helper Sidebar` again after authorization.

### Using the Add-on

1.  **Open the Helper Sidebar:** In your Google Sheet, go to `Extensions > Kruti Dev Tools > Open Helper Sidebar`. The sidebar will appear on the right.
2.  **Convert Typed Text:**
    *   Use the "Type Kruti Dev here:" textarea to paste or type Kruti Dev text, then click "Convert to Unicode" to see the Unicode output.
    *   Use the "Type Unicode here:" textarea to paste or type Unicode text, then click "Convert to Kruti Dev" to see the Kruti Dev output.
3.  **Convert Selected Cells:**
    *   **To Read/Edit (Kruti Dev to Unicode):** If you have data in your sheet that appears as gibberish (because it's Kruti Dev text being displayed in a Unicode environment), select the cells containing this text. Then, click "Convert Selected to Unicode" in the sidebar. The text in the selected cells will be converted to readable Hindi.
    *   **To Prepare for Export (Unicode to Kruti Dev):** If you have Unicode Hindi text in your sheet and need to convert it back to Kruti Dev format (e.g., for use in a legacy application), select the cells. Then, click "Convert Selected to Kruti Dev".
4.  **Download as Kruti Dev CSV:**
    *   Click "Download as Kruti Dev CSV". This will download a `.csv` file where all string content has been converted to Kruti Dev.
    *   **Important:** Open this CSV file in Microsoft Excel or a similar spreadsheet program that allows you to apply the Kruti Dev font to the cells for correct display. CSVs do not retain font information.
5.  **Print Sheet:**
    *   Click "Print Sheet (as current view)". This will trigger your browser's print dialog for the current view of your Google Sheet.
    *   **Note on Printing Kruti Dev:** For Hindi text to print correctly *directly from Google Sheets*, it should be in Unicode format. If you need a printed output specifically in a Kruti Dev font, your best approach is to:
        1.  Convert the data in your sheet to Kruti Dev using the sidebar.
        2.  Download it as a Kruti Dev CSV.
        3.  Open the CSV in Excel, apply the Kruti Dev font, and then print from Excel.
