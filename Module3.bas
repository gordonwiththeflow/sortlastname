Attribute VB_Name = "Module3"
Sub CountValuesInColumnA()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim countA As Long
    
    ' 設置工作表
    Set ws = ActiveSheet
    
    ' 計算最後一行
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' 計算 A 欄中數值的總數（不包括 A1 表頭）
    countA = Application.WorksheetFunction.countA(ws.Range("A2:A" & lastRow))
    
    ' 顯示 MessageBox
    MsgBox "目前FCCD 總共有 " & countA & " 個會員", vbInformation
End Sub
