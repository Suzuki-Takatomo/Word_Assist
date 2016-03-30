VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} WordAssist 
   Caption         =   "Word Assist"
   ClientHeight    =   4572
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   2340
   OleObjectBlob   =   "WordAssist.frx":0000
   StartUpPosition =   3  'Windows の既定値
End
Attribute VB_Name = "WordAssist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub add_words_Click()
    word_add.Show vbModeless
End Sub

Private Sub CommandButton1_Click()
    Call excludex
End Sub

Private Sub CommandButton2_Click()
    Call nonMarker
End Sub

Private Sub CommandButton3_Click()
    Call Marker.next_sentence
End Sub

Private Sub Delete_Click()
    Call DeleteList
End Sub

Private Sub next_sentence_Click()
    Call Marker.next_sentence
End Sub

Private Sub RED_Click()
    Call RedMarker
End Sub

Private Sub SearchButtom_Click()
    Me.information1.Visible = True
    Me.Delete.Visible = False
    Call Search
    Me.Delete.Visible = True
    Me.information1.Visible = False
End Sub
Private Sub SearchWeb_Click()
    Call OpenURL
End Sub
Private Sub SettingButtom_Click()
    Setting.Show vbModeless
End Sub

Private Sub this_sentence_Click()
    Call Marker.Sentence
End Sub



Private Sub UserForm_Initialize()                           '既定値の設定
    Searchengine = "http://ejje.weblio.jp/content/"
    Searchtag = "meaning"
    SearchIndex = 1
End Sub
Private Sub UserForm_Terminate()
    If changed = True Then
        Dim rc As Long
        rc = MsgBox("単語帳の内容が変更されています。変更を保存しますか?", vbYesNo + vbInformation)
        If rc = vbYes Then
            Call excludex
        End If
    End If
    Index = 0
    changed = False
    saved = False
    Dim i As Integer
    For i = 0 To wordnumber
        mean(i) = ""
    Next i
End Sub
Private Sub WordList_AfterUpdate()
    WordAssist.meaningLabel.Caption = mean(Me.WordList.ListIndex)
End Sub
Private Sub WordList_Change()
    WordAssist.meaningLabel.Caption = mean(Me.WordList.ListIndex)
End Sub
Private Sub Yellow_Click()
    Call YellowMarker
End Sub
