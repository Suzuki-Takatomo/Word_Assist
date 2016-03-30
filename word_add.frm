VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} word_add 
   Caption         =   "単語の追加"
   ClientHeight    =   1524
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "word_add.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "word_add"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub add_B_Click()
    Dim word, meanings As String
    
    word = Me.word_text.Text
    meanings = Me.meaning_text.Text
    word = Replace(word, ",", "")
    meanigs = Replace(meaning, ",", "/")
    
    If Replace(word, " ", "") <> "" Then
        mean(Index) = meanings
        Index = Index + 1
        WordAssist.WordList.AddItem word
        changed = True
        
        If Index <> 0 Then
            WordAssist.WordList.selected(Index - 1) = True           '新しく追加された単語の選択。
        End If
        Unload Me
        Call count_words
    Else
        Unload Me
    End If
End Sub

Private Sub cancel_B_Click()
    Unload Me
End Sub

Private Sub meaning_clear_Click()
    Me.meaning_text = ""
End Sub

Private Sub UserForm_Activate()
    On Error Resume Next
    Dim word As String
    Dim CB As New DataObject
    CB.GetFromClipboard
    
    word = Selection.Text
    word_add.word_text.Value = word
    word_add.meaning_text.Value = CB.GetText
End Sub

Private Sub word_clear_Click()
    Me.word_text = ""
End Sub
