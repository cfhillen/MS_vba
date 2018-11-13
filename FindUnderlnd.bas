Attribute VB_Name = "FindUnderlnd"
Sub Macro1()

    Selection.Find.ClearFormatting
    Selection.Find.Font.Underline = wdUnderlineSingle
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
End Sub
