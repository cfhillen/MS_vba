Attribute VB_Name = "DelBlankRows"
Sub DeleteBlankRows()
Dim x As Long

With ActiveSheet
    For x = .Cells.SpecialCells(xlCellTypeLastCell).Row _
        To 1 Step -1

        If WorksheetFunction.CountA(.Rows(x)) = 0 Then
            ActiveSheet.Rows(x).Delete
        End If

    Next
End With

End Sub


