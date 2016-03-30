Attribute VB_Name = "Marker"
Sub YellowMarker()
Attribute YellowMarker.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro2"
    Selection.Range.HighlightColorIndex = wdYellow
End Sub
Sub RedMarker()
    Selection.Range.HighlightColorIndex = wdRed
End Sub
Sub nonMarker()
    Selection.Range.HighlightColorIndex = wdNoHighlight
End Sub
Sub count_words()
    Dim wd As String
    wd = "検索数:"
    If WordAssist.WordList.ListCount <> 0 Then
        wd = wd & WordAssist.WordList.ListCount
       
        If changed = False Then
            WordAssist.words_num.ForeColor = &H0&
        Else
            WordAssist.words_num.ForeColor = &HFF&
        End If
    Else
        wd = ""
    End If
     WordAssist.words_num.Caption = wd
End Sub
Sub Sentence()
    On Error GoTo myError
    If Selection.Information(wdFirstCharacterLineNumber) <> -1 Then
        Selection.Expand unit:=wdSentence
        backup_Sentence = Selection.Text
        With Selection
            page = .Information(wdActiveEndPageNumber)
            row = .Information(wdFirstCharacterLineNumber)
        End With
    Else
        Selection.Expand unit:=wdSentence
    End If
    Selection.Comments.Add Range:=Selection.Range
    ActiveWindow.ActivePane.Close
    If row <> -1 Then
        WordAssist.next_sentence.Visible = True
    End If
Exit Sub
myError:
    MsgBox "カーソルを本文に合わせてください"
End Sub
Sub next_sentence()
    If row <> -1 Then
        Selection.GoTo what:=wdGoToPage, count:=page
        Call next_select
        Call Sentence
    End If
End Sub
Sub next_select()
    With Selection.Find
        .Forward = False
        .ClearFormatting
        .MatchWholeWord = True
        .MatchCase = False
        .Wrap = wdFindContinue
        
        If .Execute(findText:=backup_Sentence, Forward:=True, Format:=True) = True Then
            Selection.HomeKey unit:=wdLine, Extend:=wdExtend
            Selection.Expand unit:=wdSentence
            Selection.MoveRight unit:=wdCharacter, count:=3
        End If
    End With
    Selection.Find.ClearFormatting
End Sub
