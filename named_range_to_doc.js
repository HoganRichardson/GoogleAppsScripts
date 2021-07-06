/*
 * --- Named Ranges to Template ---
 *
 * A spreadsheet to document report that populates a document template with 
 *  the values of named ranges in a spreadsheet.
 * 
 * Creates a single document for each report generated
 * 
 */

function onOpen() {
  var menuEntries = [ {name: "Create Report from Template", functionName: "AutofillDocFromTemplate"}]
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  ss.addMenu("Report", menuEntries)
}

TAG_SEPARATOR = "%"
OUTPUT_FOLDER = ""

function AutofillDocFromTemplate() {
  var namedRanges = SpreadsheetApp.getActiveSpreadsheet().getNamedRanges()
  
  // Create Copy of Template
  var doc_filename = get_nr_value(namedRanges, "filename");
  var template_id = get_nr_value(namedRanges, "template_id");
  
  var new_doc = copyDocument(template_id, OUTPUT_FOLDER, doc_filename);
  var doc = DocumentApp.openById(new_doc.getId());
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Fill in all named ranges
  for (var i = 0; i < namedRanges.length; i++) {
    n_range = namedRanges[i]

    if (n_range.getRange().getNumRows() == 1 && n_range.getRange().getNumColumns() == 1) {
      /* Single text string to replace */
      replaceText(n_range, doc)
    } else {
      /* Insert table range */
      appendNamedRange(n_range, doc)    
    }
  }

  // Save change
  doc.saveAndClose()

  // Convert to DOCX and delete Google Doc
  exportAsDocx(doc_filename, new_doc.getId(), OUTPUT_FOLDER)
  DriveApp.getFileById(new_doc.getId()).setTrashed(true)
  
  ss.toast("Report has been generated.")
}

function get_nr_value(namedRanges, name) {
  for (var i = 0; i < namedRanges.length; i++) { 
    if (namedRanges[i].getName() == name) {
      return namedRanges[i].getRange().getValue()
    }
  }
}

function copyDocument(source_id, destination_id, filename) {
  // Create Copy of Template in Destination
  var folder = DriveApp.getFolderById(destination_id) 
  var new_doc = DriveApp.getFileById(source_id).makeCopy()
  new_doc.setName(filename)
  folder.addFile(new_doc)
  return new_doc
}

function appendNamedRange(n_range, doc) {
  tables = doc.getBody().getTables()

  // Find the insert point
  for (var t = 0; t < tables.length; t++) {
    curr_table = tables[t]

    // If this table contains tag
    tag = TAG_SEPARATOR + n_range.getName() + TAG_SEPARATOR

    if (curr_table.findText(tag) != null) {
      // Remove last row (with tag)
      curr_table.removeRow(curr_table.getNumRows() - 1)

      // Append range to table
      newtable = doc.getBody().appendTable(n_range.getRange().getValues())
      newtable.removeFromParent()

      while (newtable.getNumRows() > 0) {
        var newrow = newtable.getRow(0).removeFromParent()
        // Skip empty strings
        if (newrow.getText().replace(/\s/g, "").length != 0) { 
          curr_table.appendTableRow(newrow)
        }
      }
    }
  }
}

function replaceText(n_range, doc) {
  var tag = TAG_SEPARATOR + n_range.getName() + TAG_SEPARATOR
  var body = doc.getActiveSection();
  body.replaceText(tag, n_range.getRange().getValue())
}

function exportAsDocx(doc_name, doc_id, folder_id) {
  token = ScriptApp.getOAuthToken()
  var blob = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + doc_id + '/export?mimeType=application/vnd.openxmlformats-officedocument.wordprocessingml.document', 
      {
        headers: {
          'Authorization': 'Bearer ' + token
        }
      }).getBlob()
  var file = DriveApp.createFile(blob).setName(doc_name + '.docx')
  DriveApp.getFolderById(folder_id).addFile(file)
}