Attribute VB_Name = "Module3"
Sub CountValuesInColumnA()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim countA As Long
    
    ' �]�m�u�@��
    Set ws = ActiveSheet
    
    ' �p��̫�@��
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' �p�� A �椤�ƭȪ��`�ơ]���]�A A1 ���Y�^
    countA = Application.WorksheetFunction.countA(ws.Range("A2:A" & lastRow))
    
    ' ��� MessageBox
    MsgBox "�ثeFCCD �`�@�� " & countA & " �ӷ|��", vbInformation
End Sub
