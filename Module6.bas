Attribute VB_Name = "Module6"
Sub CopyReport()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim rangeToCopy As Range

    ' Get the currently active worksheet
    Set ws = ActiveSheet

    ' Find the row and column numbers of the last row and last column
    lastRow = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
    lastCol = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column

    ' Set the range to copy, starting from E2 to the last row and last column
    Set rangeToCopy = ws.Range(ws.Cells(2, 5), ws.Cells(lastRow, lastCol))

    ' Copy the range to the clipboard
    rangeToCopy.Copy
End Sub
