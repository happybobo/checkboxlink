checkboxlink
============

Checkboxlink

A proof of concept that check box can be manipulated from LibXL (with help of Excel Macros)

# Usage

1. Launch app
2. Tap 'Edit XLS' or 'Edit XLSX' buttons
   Every time you tap, it is going to toggle check mark of existing excel sheet (initial form is bundled in the app as sample.xls and sample.xlsx).
NOTE : Excel does not reload automatically, so you need to close/open Excel Book after tapping 'Edit ...' to check latest status.

# Idea

If you look at the Row 33-36 of sample.xls and sample.xlsx, there are boolean text entries. Those entries are linked to checkboxes located up 32 rows.
So, this idea only works "if there is one checkbox corresponiding to one cell". So far from templates, it seems to be the case.
