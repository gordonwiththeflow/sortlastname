Attribute VB_Name = "Module1"
Sub SortByLastName()
    Dim ws As Worksheet
    Dim lastRow As Long
    
    ' 設置工作表
    Set ws = ActiveSheet
    
    ' 計算最後一行
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' 排序 A 欄的數據
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add Key:=ws.Range("A2:A" & lastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    
    With ws.Sort
        .SetRange ws.Range("A1:Z" & lastRow) ' 假設 Z 列為最大列
        .Header = xlYes ' 如果有表頭，使用 xlYes；沒有表頭使用 xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin ' 這裡用的是拼音，但仍然可以正確排序英文
        .Apply
    End With
End Sub

