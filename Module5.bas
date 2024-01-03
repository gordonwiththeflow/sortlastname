Attribute VB_Name = "Module5"
Sub ChangeFont()
    Dim ws As Worksheet
    Dim cell As Range

    ' Get the currently active worksheet
    Set ws = ActiveSheet

    ' Iterate through each cell in the worksheet
    For Each cell In ws.UsedRange
        ' Check if the content of the cell is Chinese
        If IsChineseText(cell.Value) Then
            cell.Font.Name = "¼Ð·¢Åé" ' Set the font to SimKai for Chinese
        ElseIf IsEnglishText(cell.Value) Then
            cell.Font.Name = "Times New Roman" ' Set the font to Times New Roman for English
        End If
    Next cell
End Sub

Function IsChineseText(text As String) As Boolean
    ' Use regular expression to check if it's Chinese
    IsChineseText = Len(text) > 0 And Not RegExTest(text, "[^\u4e00-\u9fa5]")
End Function

Function IsEnglishText(text As String) As Boolean
    ' Use regular expression to check if it's English
    IsEnglishText = Len(text) > 0 And Not RegExTest(text, "[^\x00-\x7F]")
End Function

Function RegExTest(inputText As String, pattern As String) As Boolean
    ' Use regular expression for matching test
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True
    regEx.pattern = pattern
    RegExTest = regEx.Test(inputText)
End Function
