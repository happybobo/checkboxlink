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
Excel allows you to *__link__* a checkbox to a value of the cell.

If you look at the Row 33-36 of sample.xls and sample.xlsx, there are boolean text entries. Those entries are *__linked__* to checkboxes located up 32 rows. LibXL needs to only interact with the boolean text fields, and Excel automatically reflect the result to corresponding checkboxes.

This idea only works "if there is one checkbox corresponiding to one cell". But so far from templates, it seems to be the case.

# Important part of the Code
* `ExcelHandleViewController.m` line 88-108. Here app toggles boolean entries via libXL.
* The macro to link checkboxes to the cells below 32 rows. It just needs to run once before form gets bundled into the app. (Thus, not in the repo) : 
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
