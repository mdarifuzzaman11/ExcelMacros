Sub ChangeReport()
'
' ChangeReport Macro
'
' Keyboard Shortcut: Ctrl+Shift+C
'
Dim lastRow As Long
Dim lastCol As Long
' Find the last row with data in column F
lastRow = Cells(Rows.Count, "F").End(xlUp).Row

' Find the last column with data in row 1
lastCol = Cells(1, Columns.Count).End(xlToLeft).Column
Range("A1:F1").Select
Selection.Style = "Accent1"
Range(Cells(1, 1), Cells(lastRow, lastCol)).Select
Selection.ColumnWidth = 23
Selection.ColumnWidth = 31.29
Selection.ColumnWidth = 36.29

' Bold Row 1
Range("A1:" & Cells(1, lastCol).Address).Font.Bold = True

With Selection.Borders
.LineStyle = xlContinuous
.ColorIndex = 0
.TintAndShade = 0
.Weight = xlThin
End With

Cells.Select
With Selection
.HorizontalAlignment = xlCenter
.Orientation = 0
.AddIndent = False
.IndentLevel = 0
.ShrinkToFit = False
.ReadingOrder = xlContext
.MergeCells = False
End With

Range("A1").Select

' Copy the data to the clipboard
Range(Cells(1, 1), Cells(lastRow, lastCol)).Select
Selection.Copy

End Sub
