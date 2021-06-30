/*
 * --- List to Template ---
 *
 * A simple 'mail merge' that populates a document template with the values in
 *  columns.
 * The first column of data should contain the template document ID.
 * 
 * Creates a new document for each row in the selected sheet
 * 
 */

function onOpen() {
    var menuEntries = [{ name: "Create Report from Template", functionName: "AutofillDocFromTemplate" }];
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.addMenu("Report", menuEntries);
}

TAG_SEPARATOR = "%";

function AutofillDocFromTemplate() {
    var folder_id = "" // Output folder id
    var folder = DriveApp.getFolderById(folder_id);

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();

    // Get Filter Range
    var filter = sheet.getFilter();
    if (filter == null) {
        SpreadsheetApp.getUi().alert("No filter found in curent sheet. Create a filter and try again.")
        return;
    }

    data = filter.getRange().getDisplayValues();
    num_data_rows = filter.getRange().getNumRows() - 1;
    headings = data[0];
    num_cols = headings.length;

    for (var row = 0; row < num_data_rows; row++) {
        this_row = data[row + 1]; // skip header
        template_id = this_row[0];

        // Create Copy of Template in Destination
        var doc_title = this_row[1];
        var new_doc = DriveApp.getFileById(template_id).makeCopy();
        new_doc.setName("Report for " + doc_title);
        folder.addFile(new_doc);

        // Get Document Body
        var doc = DocumentApp.openById(new_doc.getId());
        var body = doc.getActiveSection();

        // Populate Template Fields
        for (var col = 1; col < num_cols; col++) { // Skip first col (template_id)
            key = TAG_SEPARATOR + headings[col] + TAG_SEPARATOR;
            body.replaceText(key, this_row[col])
        }

        doc.saveAndClose();

        // Convert to Docx
        exportAsDocx(doc_title, new_doc.getId(), folder_id);
        DriveApp.getFileById(new_doc.getId()).setTrashed(true);
    }
    ss.toast("Template has been compiled.");
}

function exportAsDocx(doc_name, doc_id, folder_id) {
    token = ScriptApp.getOAuthToken();
    var blob = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + doc_id + '/export?mimeType=application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        {
            headers: {
                'Authorization': 'Bearer ' + token
            }
        }).getBlob();
    var file = DriveApp.createFile(blob).setName(doc_name + '.docx');
    DriveApp.getFolderById(folder_id).addFile(file);
}