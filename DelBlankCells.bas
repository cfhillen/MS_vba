Attribute VB_Name = "DelBlankCells"
Sub DeleteBlankCells()
    Dim intCol As Integer
     
    For intCol = 1 To 26 'cols A to Z
        Range(Cells(2, intCol), Cells(1000, intCol)). _
        SpecialCells(xlCellTypeBlanks).Delete Shift:=xlUp
    Next intCol
End Sub

