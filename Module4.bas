Attribute VB_Name = "Module4"
Sub CleanAndMakeSpace()
    ' Clean table and make space for EFG columns
    
    Columns("E:h").Delete
    Columns("E:h").Insert Shift:=xlToRight
End Sub
