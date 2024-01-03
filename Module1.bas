Attribute VB_Name = "Module1"
Sub SortByLastName()
    Dim ws As Worksheet
    Dim lastRow As Long
    
    ' �]�m�u�@��
    Set ws = ActiveSheet
    
    ' �p��̫�@��
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' �Ƨ� A �檺�ƾ�
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add Key:=ws.Range("A2:A" & lastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    
    With ws.Sort
        .SetRange ws.Range("A1:Z" & lastRow) ' ���] Z �C���̤j�C
        .Header = xlYes ' �p�G�����Y�A�ϥ� xlYes�F�S�����Y�ϥ� xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin ' �o�̥Ϊ��O�����A�����M�i�H���T�Ƨǭ^��
        .Apply
    End With
End Sub

