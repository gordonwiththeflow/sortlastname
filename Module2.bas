Attribute VB_Name = "Module2"
Sub CombineMacros()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim outputRow As Long
    Dim startCell As Range
    Dim rng As Range
    Dim borderType As Variant
    Dim lastCol As Long
    
    ' Set the worksheet
    Set ws = ActiveSheet
    
    ' --- Macro1: CreateTableAndMergeValues ---
    
    ' count last row
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' initiate the row
    outputRow = 2
    
    ' Combine value and set font
    For i = 2 To lastRow Step 4
    ' Combine values
    Dim combinedValue5 As String
    Dim combinedValue6 As String
    Dim combinedValue7 As String
    Dim combinedValue8 As String
    
    combinedValue5 = IIf(Trim(ws.Cells(i, 3).Value) <> "", ws.Cells(i, 3).Value & " " & ws.Cells(i, 2).Value & " " & ws.Cells(i, 1).Value, ws.Cells(i, 2).Value & " " & ws.Cells(i, 1).Value)
    combinedValue6 = IIf(Trim(ws.Cells(i + 1, 3).Value) <> "", ws.Cells(i + 1, 3).Value & " " & ws.Cells(i + 1, 2).Value & " " & ws.Cells(i + 1, 1).Value, ws.Cells(i + 1, 2).Value & " " & ws.Cells(i + 1, 1).Value)
    combinedValue7 = IIf(Trim(ws.Cells(i + 2, 3).Value) <> "", ws.Cells(i + 2, 3).Value & " " & ws.Cells(i + 2, 2).Value & " " & ws.Cells(i + 2, 1).Value, ws.Cells(i + 2, 2).Value & " " & ws.Cells(i + 2, 1).Value)
    combinedValue8 = IIf(Trim(ws.Cells(i + 3, 3).Value) <> "", ws.Cells(i + 3, 3).Value & " " & ws.Cells(i + 3, 2).Value & " " & ws.Cells(i + 3, 1).Value, ws.Cells(i + 3, 2).Value & " " & ws.Cells(i + 3, 1).Value)
    
    ' Set the combined value and apply font
    With ws.Cells(outputRow, 5)
        .Value = combinedValue5
        ' Set font for the specific portions
        .Characters(Start:=1, Length:=Len(ws.Cells(i, 3).Value)).Font.Name = "標楷體" ' Replace with the desired font
        .Characters(Start:=Len(ws.Cells(i, 3).Value) + 1, Length:=Len(ws.Cells(i, 2).Value)).Font.Name = "Times New Roman" ' Replace with the desired font
        .Characters(Start:=Len(ws.Cells(i, 3).Value) + Len(ws.Cells(i, 2).Value) + 2, Length:=Len(ws.Cells(i, 1).Value)).Font.Name = "Times New Roman" ' Replace with the desired font
    End With
    
    With ws.Cells(outputRow, 6)
        .Value = combinedValue6
        ' Set font for the specific portions
        .Characters(Start:=1, Length:=Len(ws.Cells(i + 1, 3).Value)).Font.Name = "標楷體" ' Replace with the desired font
        .Characters(Start:=Len(ws.Cells(i + 1, 3).Value) + 1, Length:=Len(ws.Cells(i + 1, 2).Value)).Font.Name = "Times New Roman" ' Replace with the desired font
        .Characters(Start:=Len(ws.Cells(i + 1, 3).Value) + Len(ws.Cells(i + 1, 2).Value) + 2, Length:=Len(ws.Cells(i + 1, 1).Value)).Font.Name = "Times New Roman" ' Replace with the desired font
    End With
    
    With ws.Cells(outputRow, 7)
        .Value = combinedValue7
        ' Set font for the specific portions
        .Characters(Start:=1, Length:=Len(ws.Cells(i + 2, 3).Value)).Font.Name = "標楷體" ' Replace with the desired font
        .Characters(Start:=Len(ws.Cells(i + 2, 3).Value) + 1, Length:=Len(ws.Cells(i + 2, 2).Value)).Font.Name = "Times New Roman" ' Replace with the desired font
        .Characters(Start:=Len(ws.Cells(i + 2, 3).Value) + Len(ws.Cells(i + 2, 2).Value) + 2, Length:=Len(ws.Cells(i + 2, 1).Value)).Font.Name = "Times New Roman" ' Replace with the desired font
    End With
    
    With ws.Cells(outputRow, 8)
        .Value = combinedValue8
        ' Set font for the specific portions
        .Characters(Start:=1, Length:=Len(ws.Cells(i + 3, 3).Value)).Font.Name = "標楷體" ' Replace with the desired font
        .Characters(Start:=Len(ws.Cells(i + 3, 3).Value) + 1, Length:=Len(ws.Cells(i + 3, 2).Value)).Font.Name = "Times New Roman" ' Replace with the desired font
        .Characters(Start:=Len(ws.Cells(i + 3, 3).Value) + Len(ws.Cells(i + 3, 2).Value) + 2, Length:=Len(ws.Cells(i + 3, 1).Value)).Font.Name = "Times New Roman" ' Replace with the desired font
    End With

    outputRow = outputRow + 1 ' move to the next
    Next i
    
    ' --- Macro2: Macro4 ---
    
    ' Set the starting cell (E2)
    Set startCell = ws.Range("E2")
    
    ' Find the last row and last column in the specified range
    lastRow = ws.Cells(ws.Rows.Count, startCell.Column).End(xlUp).Row
    lastCol = ws.Cells(startCell.Row, ws.Columns.Count).End(xlToLeft).Column
    
    ' Set the range to format
    Set rng = ws.Range(startCell, ws.Cells(lastRow, lastCol))
    
    ' Loop through each border type and set formatting
    For Each borderType In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, xlInsideVertical, xlInsideHorizontal)
        With rng.Borders(borderType)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin ' Use xlThin for normal line weight
        End With
    Next borderType
    
    ' Clear diagonal borders
    rng.Borders(xlDiagonalDown).LineStyle = xlNone
    rng.Borders(xlDiagonalUp).LineStyle = xlNone
    
    ' Autofit columns
    Columns("E:h").EntireColumn.AutoFit
End Sub

