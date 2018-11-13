Attribute VB_Name = "DelDupesInString"
Option Explicit
 
Sub Remove_DupesInString()
Attribute Remove_DupesInString.VB_Description = "Removes duplicates within a cell that are separated by a semicolon"
Attribute Remove_DupesInString.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim starval As String
    Dim finval As String
    Dim strarray() As String
    Dim x As Long
    Dim k As Long
    Dim cell As Range
    Dim rw As Long
     
     ' step through each cell in range
    For Each cell In ActiveSheet.Range("F1:I8377")
        Erase strarray ' erase array
        finval = "" ' erase final value"
        starval = cell.Value
        On Error Resume Next
         
        strarray = Split(starval, ";")
         
         'Step through length of string and look for duplicate
        For rw = 0 To UBound(strarray)
             
            For k = rw + 1 To UBound(strarray)
                If Trim(strarray(k)) = Trim(strarray(rw)) Then
                    strarray(k) = "" 'if duplicate clear array value
                End If
            Next k
        Next rw
         
         ' combine all value in string less duplicate
        For x = 0 To UBound(strarray)
            If strarray(x) <> "" Then
                 
                finval = finval & Trim(strarray(x)) & "; "
            End If
             
        Next x
         ' remove last space and comma
        finval = Trim(finval)
        finval = Left(finval, Len(finval) - 1)
         ' output value to Column J
        cell.Offset(0, 0).Value = finval
         
    Next cell
     
End Sub

