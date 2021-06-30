# A collection of Google Apps Scripts

## `list_to_template`

A simple 'mail merge' that populates a document template with the values in columns.
* The first column of data should contain the document ID of the template.

Example Spreadsheet (must have filter view on the data):

template_id | columname1 | columname2 | ...
---|---|---|---
asdf_1234|data1|data2|...
asdf_1234|doc2data1|doc2data2|...

The Document should contain tags with the format `%columnname%`. All instances of such tags will be replaced with the cell contents. 

## `named_range_to_doc`