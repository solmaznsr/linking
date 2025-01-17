# linking
Generate Google Doc Links from Google Sheets
This script automates the process of generating Google Docs from rows in a Google Sheet. Each row's data is used to populate a document, and a link to the generated document is stored back in the sheet. The script ensures only the latest updates are reflected in the documents, avoiding duplication.
##Features
1.Batch Processing: Processes up to 100 rows at a time to avoid performance issues.
2.Dynamic Updates: Updates existing documents if they already exist, ensuring only one document per unique ID.
3.Customizable Columns: Specify which columns from the sheet should be included in the document.
4.Styling Support: Apply custom styles to the generated documents, including font size, font family, alignment, and bold text.
5.Error Handling: Logs errors and skips rows with missing data or pre-existing links.
