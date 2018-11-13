Attribute VB_Name = "ExtractEmailsFromWord"
Sub ExtractEmailsFromWord()
    Dim strFolder       As String
    Dim vntWord         As Variant
    Dim flDialog        As FileDialog
    
    Set flDialog = Application.FileDialog(msoFileDialogFilePicker)
    
    With flDialog
        .Filters.Clear
        .Filters.Add "Word Documents", "*.docx,*.docm,*.dotx,*.dotm,*.doc,*.dot"
        .Title = "Open"
        
        On Error Resume Next
        
        '/* The user pressed the button .*/
        If .Show = -1 Then
            If .SelectedItems.Count >= 1 Then
                Dim objWordDoc  As Object
                Dim tbRange     As Object
                Dim objRegExp   As Object
                Dim objMatch    As Object
                Dim objMatches  As Object
                Dim lngLoop     As Long
                
                lngLoop = 1
                
                ' Create a Regular Expression object.
                Set objRegExp = CreateObject("VBScript.RegExp")
                
                objRegExp.IgnoreCase = True  ' Ignore case.
                objRegExp.MultiLine = False  ' Cancel multiline mode.
                objRegExp.Global = True      ' Global match.
                
                ' To match the Email address.
                objRegExp.Pattern = "\w+([-+.]\w+)*@\w+([-.]\w+)*\.*"
                
                ' /* Step through each string in the FileDialogSelectedItems collection. */
                For Each vntWord In .SelectedItems
                    Set objWordDoc = GetObject(vntWord)
                    
                    ' /* Step through each table in the objWordDoc. */
                    For Each tbRange In objWordDoc.Tables
                        Set objMatches = objRegExp.Execute(tbRange.Range.Text)
                        
                        ' /* Step through Match object in the MatchCollection object. */
                        For Each objMatch In objMatches
                            ' Put the email address into column A.
                            Range("A" & CStr(lngLoop)).Value = objMatch.Value
                            ' Change the counter.
                            lngLoop = lngLoop + 1
                        Next
                    Next
                Next
            End If
        End If
    End With
    
    ' /* Release memory. */
    If Not (objWordDoc Is Nothing) Then Set objWordDoc = Nothing
    If Not (tbRange Is Nothing) Then Set tbRange = Nothing
    If Not (objRegExp Is Nothing) Then Set objRegExp = Nothing
    If Not (objMatch Is Nothing) Then Set objMatch = Nothing
    If Not (objMatches Is Nothing) Then Set objMatches = Nothing
End Sub
