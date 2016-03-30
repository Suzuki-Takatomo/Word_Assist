VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Setting 
   Caption         =   "Setting"
   ClientHeight    =   2664
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "Setting.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "Setting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function SearchSet(SE As String, ST As String, SI As Integer)
    Me.searchE.Text = SE
    Me.searchT.Text = ST
    Me.SearchIndexBox.Text = SI
End Function
Private Sub ComboBox_Change()
    Select Case ComboBox.Text
        Case "weblio英和辞典"
            Call SearchSet("http://ejje.weblio.jp/content/", "meaning", 1)
        Case "コトバンク"
            Call SearchSet("https://kotobank.jp/ejword/", "word_foreign", 2)
        Case "WordReference"
            Call SearchSet("http://www.wordreference.com/enja/", "ToWrd", 1)
        Case Else
            Call SearchSet("", "", 1)
    End Select
End Sub

Private Sub include_Click()
    Call includex
End Sub

Private Sub OK_Click()
    Searchengine = Me.searchE.Text
    Searchtag = Me.searchT.Text
    SearchIndex = Me.SearchIndexBox.Text
    Unload Me
End Sub

Private Sub UserForm_Initialize()
        Me.ComboBox.AddItem ("weblio英和辞典")
        Me.ComboBox.AddItem ("コトバンク")
        Me.ComboBox.AddItem ("WordReference")
        Select Case Searchengine
            Case "http://ejje.weblio.jp/content/"
                Me.ComboBox.Text = "weblio英和辞典"
            Case "https://kotobank.jp/ejword/"
                Me.ComboBox.Text = "コトバンク"
            Case "http://www.wordreference.com/enja/"
                Me.ComboBox.Text = "WordReference"
            Case Else
                Me.ComboBox.Text = ""
        End Select
        
        For i = 1 To 5
            Me.SearchIndexBox.AddItem (i)
        Next i
        Call SearchSet(Searchengine, Searchtag, SearchIndex)
End Sub
