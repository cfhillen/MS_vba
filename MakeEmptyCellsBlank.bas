Attribute VB_Name = "MakeEmptyCellsBlank"
Sub MakeEmptyCellsBlank()
Dim cell As Range
For Each cell In Range("A1:Z1000") ' <<<< to be changed
  If IsEmpty(cell) Then
    cell.Value = " "
  End If
Next
End Sub
