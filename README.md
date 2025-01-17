# linking
Generate Google Doc Links from Google Sheets
This script automates the process of generating Google Docs from rows in a Google Sheet. Each row's data is used to populate a document, and a link to the generated document is stored back in the sheet. The script ensures only the latest updates are reflected in the documents, avoiding duplication.
## Features
1. Batch Processing: Processes up to 100 rows at a time to avoid performance issues.
2. Dynamic Updates: Updates existing documents if they already exist, ensuring only one document per unique ID.
3. Customizable Columns: Specify which columns from the sheet should be included in the document.
4. Styling Support: Apply custom styles to the generated documents, including font size, font family, alignment, and bold text.
5. Error Handling: Logs errors and skips rows with missing data or pre-existing links.

## Setup Instructions
1. Clone or Download the Repository
2. Prepare Your Google Sheet:
Ensure your sheet contains a unique identifier column (e.g., ID).
Add a column to store the document links (e.g., "Doc Link").0
3. Create a Google Doc Template:
Design a Google Doc to use as the template for generated documents.
4. Configure the Script:
Open the script in Google Apps Script.
Replace placeholders with your actual data:
TEMPLATE_DOC_ID_HERE: Google Doc template ID.
TARGET_FOLDER_ID_HERE: Target folder ID for saving the documents.
5. Authorize the Script:
Run the script in Apps Script and provide necessary permissions
6. Run the Script:
Open the Google Sheet, go to Extensions > Apps Script, and execute the generateDocLinksBatch function.