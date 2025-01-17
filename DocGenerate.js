//Batch generates Google Docsfor rows in a Google Sheet and inserts document links.

function generateDocLinksBatch() {
    const CONFIG = {
        TEMPLATE_DOC_ID: "TEMPLATE_DOC_ID_HERE", // Replace with your template document ID
        TARGET_FOLDER_ID: "TARGET_FOLDER_ID_HERE", // Replace with your target folder ID
        BATCH_SIZE: 100, // Maximum rows to process per execution
        LINK_COLUMN_INDEX: 0, // Column index for storing document links (0-based)
        COLUMN_NAMES: [
            "ID", "Name", "Category", "City", "Address",
            "Status", "Priority", "Assigned To"
            // Add more generic column names as needed
        ],
        DEFAULT_STYLES: { // Default document styles
            fontSize: 12, // Font size in points
            fontFamily: "Arial", // Font family
            alignment: DocumentApp.HorizontalAlignment.LEFT, // Text alignment
        }
    };

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const dataRange = sheet.getDataRange();
    const data = dataRange.getValues();
    const headers = data[0];

    // Check if "ID" column exists in headers
    const idIndex = headers.indexOf("ID");
    if (idIndex === -1) {
        Logger.log('Error: "ID" column not found in the sheet headers.');
        return;
    }

    const templateDoc = DriveApp.getFileById(CONFIG.TEMPLATE_DOC_ID);
    const targetFolder = DriveApp.getFolderById(CONFIG.TARGET_FOLDER_ID);

    let rowsProcessed = 0; // Track the number of rows processed

    // Loop through each row and process it
    for (let i = 1; i < data.length && rowsProcessed < CONFIG.BATCH_SIZE; i++) {
        const row = data[i];
        const id = row[idIndex];
        const linkCell = sheet.getRange(i + 1, CONFIG.LINK_COLUMN_INDEX + 1);

        // Skip rows with no ID or already processed rows
        if (!id || linkCell.getValue()) {
            continue;
        }

        // Get or create a document for the current ID
        const doc = getOrCreateDoc(targetFolder, templateDoc, id);

        // Update the document content with row data
        updateDocWithRowData(doc, row, headers, CONFIG.COLUMN_NAMES, CONFIG.DEFAULT_STYLES, sheet, i);

        // Save the link to the document in the Google Sheet
        linkCell.setValue(`https://docs.google.com/document/d/${doc.getId()}`);
        rowsProcessed++;
    }

    // Log processing results
    Logger.log(rowsProcessed > 0 ?
        `Processed ${rowsProcessed} rows.` :
        "No new rows to process.");
}

/**
 * Retrieve an existing document or create a new one for an ID.
 * @param {Folder} folder - Target folder for the document.
 * @param {File} templateDoc - Template document file.
 * @param {string} id - Unique ID for the row.
 * @returns {Document} - Google Document object.
 */
function getOrCreateDoc(folder, templateDoc, id) {
    const existingDocs = folder.getFilesByName(`Data - ${id}`);
    return existingDocs.hasNext() ?
        DocumentApp.openById(existingDocs.next().getId()) :
        DocumentApp.openById(templateDoc.makeCopy(`Data - ${id}`, folder).getId());
}

/**
 * Update the document with data from the current row and apply styling.
 * @param {Document} doc - Google Document object.
 * @param {Array} row - Row data from the Google Sheet.
 * @param {Array} headers - Column headers from the Google Sheet.
 * @param {Array} columnNames - List of columns to include in the document.
 * @param {Object} styles - Default text styles (font, size, alignment).
 * @param {Sheet} sheet - Google Sheet object.
 * @param {number} rowIndex - Row index (1-based) in the Google Sheet.
 */
function updateDocWithRowData(doc, row, headers, columnNames, styles, sheet, rowIndex) {
    const body = doc.getBody();
    body.clear(); // Clear the document before adding new content

    // Add document title
    const titleParagraph = body.appendParagraph("Data Summary");
    titleParagraph.setHeading(DocumentApp.ParagraphHeading.HEADING1);
    applyStyles(titleParagraph, {...styles, fontSize: 16 }); // Apply custom title style

    // Add row data to the document
    columnNames.forEach(columnName => {
        const colIndex = headers.indexOf(columnName);
        if (colIndex !== -1) {
            const cell = sheet.getRange(rowIndex + 1, colIndex + 1);
            const fontSize = cell.getFontSize();
            const fontStyle = cell.getFontFamily();
            const isBold = cell.getFontWeight() === "bold";

            const paragraph = body.appendParagraph(`${columnName}: ${row[colIndex] || "N/A"}`);
            applyStyles(paragraph, {
                fontSize: fontSize || styles.fontSize,
                fontFamily: fontStyle || styles.fontFamily,
                alignment: styles.alignment,
                bold: isBold,
            });
        }
    });
}

/**
 * Apply styles to a document paragraph.
 * @param {Paragraph} paragraph - Paragraph to style.
 * @param {Object} styles - Text styles (font, size, alignment, bold).
 */
function applyStyles(paragraph, styles) {
    if (styles.fontSize) paragraph.setFontSize(styles.fontSize);
    if (styles.fontFamily) paragraph.setFontFamily(styles.fontFamily);
    if (styles.alignment) paragraph.setAlignment(styles.alignment);
    if (styles.bold !== undefined) paragraph.setBold(styles.bold);
}