Attribute VB_Name = "HighlightAlternateRows"
'This code would highlight alternate rows in the selection
Sub HighlightAlternateRows()
Dim Myrange As Range
Dim Myrow As Range
Set Myrange = Selection
For Each Myrow In Myrange.Rows
   If Myrow.Row Mod 2 = 1 Then
      Myrow.Interior.Color = RGB(255, 128, 128)
   End If
Next Myrow
End Sub
