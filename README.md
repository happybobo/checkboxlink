checkboxlink
============

Checkboxlink

A proof of concept that check box can be manipulated from LibXL with help of Excel Macros

# Usage

1. Launch app
2. Tap 'Edit XLS' or 'Edit XLSX' buttons.
3. Every time you tap the button, the app toggles check mark of bundled excel sheet (sample.xls and sample.xlsx).
NOTE : Excel does not reload automatically, so you need to close/open Excel Book after tapping 'Edit ...' to check latest status.

# Idea

If you look at the Row 33-36 of sample.xls and sample.xlsx, there are boolean text entries. Those entries are `linked` to checkboxes located up 32 rows. LibXL needs to only interact with the boolean text fields, and Excel automatically reflect the result on checkboxes.

This idea only works "if there is one checkbox corresponiding to one cell". (So far from templates, it seems to be the case.)

# Key part to look
* `ExcelHandleViewController.m` line 88-108. Here it toggle boolean values
* The macro to link checkboxes to the cells below 32 rows (it just needs to run once) : 
```
Sub LinkCheckBoxes()
  Dim chk As CheckBox
  Dim nRow As Long
  nRow = 32 'number of columns to the right for link

  For Each chk In ActiveSheet.CheckBoxes
    With chk
      .LinkedCell = _
        .TopLeftCell.Offset(nRow, 0).Address
    End With
  Next chk
End Sub
```
