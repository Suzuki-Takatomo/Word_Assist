Attribute VB_Name = "NewMacros1"
Option Explicit
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Const wordnumber As Integer = 200                               '登録できる単語数
Public Searchengine As String                                   '検索対象URL
Public Searchtag As String                                      'URL内の意味の取得するタグ
Public SearchIndex As Integer                                   'タグの何番目の意味を取得するかのインデックス
Public Index As Integer                                         '登録した数
Public mean(wordnumber) As String                               '登録された単語の意味
Public changed, saved As Boolean
Public DefaultFilePath As String
Public page, row As Integer
Public backup_Sentence As String
Sub Search()
Attribute Search.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro1"
    Dim objIE As Object                                         'IEのオブジェクト
    Dim i As Long
    Dim word As String                                          '検索対象ワード
    Dim StartTime, StopTime As Variant
    StartTime = Timer
    Application.StatusBar = "単語を取得しています。"
    Index = WordAssist.WordList.ListCount
         word = Selection.Text
         word = LCase(word)
         word = Replace(word, vbCr, "")
         word = Replace(word, vbCrLf, "")
         word = Replace(word, " ", "")
         word = Replace(word, "-", "")
         
         If WordAssist.WordList.ListIndex = -1 Then GoTo jump
         If WordAssist.WordList.List(WordAssist.WordList.ListIndex) <> word Then
jump:
            For i = 0 To WordAssist.WordList.ListCount - 1
               If WordAssist.WordList.List(i) = word Then
                   WordAssist.WordList.selected(i) = True
                   GoTo Next_step
               End If
            Next i
         End If
         Dim objXML As New MSHTML.HTMLDocument
         Dim htmlDoc As New MSHTML.HTMLDocument
         Dim objITEM As Object
         Dim URL As String
         URL = Searchengine & word
         Application.StatusBar = "webページの応答を待っています。"
         Set htmlDoc = objXML.createDocumentFromUrl(URL, vbNullString)
        Do While htmlDoc.readyState <> "complete"  '読み込み待ち
        htmlDoc.designMode = "on"
        DoEvents

        Loop
        Application.StatusBar = "意味を取得しています。"
        
        Dim IndexCount As Integer
        IndexCount = 0
        
            For Each objITEM In htmlDoc.getElementsByClassName(Searchtag)
                   Dim meaning As String
                   If objITEM.innerText <> "" Then                  '意味が空白になっていなければ、
                       IndexCount = IndexCount + 1                  '意味を格納し、空白を除去、カンマも除去(CSV対策)
                       If IndexCount = SearchIndex Then
                           meaning = objITEM.innerText
                           meaning = Replace(meaning, vbCrLf, "")
                           meaning = Replace(meaning, ",", "/")
                           mean(Index) = meaning
                           Index = Index + 1                         '単語数を追加
                           WordAssist.WordList.AddItem word
                           
                           changed = True
                           Exit For
                       End If
                   End If
second:
             Next
                       
                       If meaning = "" Then
                         MsgBox ("見つかりませんでした")             '見つからなかった場合のメッセージボックスの表示
                         GoTo Next_step
                       End If
         
        If Index <> 0 Then
            WordAssist.WordList.selected(Index - 1) = True           '新しく追加された単語の選択。
        End If
Next_step:
    Set objITEM = Nothing
    Set htmlDoc = Nothing
    Set objXML = Nothing
    StopTime = Timer
    Application.StatusBar = "検索時間:" & Str(StopTime - StartTime) & "秒"
    Call count_words
    Sleep 750
    Application.StatusBar = ""
    
    Exit Sub
myError:
    MsgBox "インターネットオブジェクト処理時にエラーが発生しました。", vbExclamation
End Sub
Sub Assist()
    WordAssist.Show vbModeless
    changed = False
    saved = False
    row = -1
    DefaultFilePath = ActiveDocument.Path & "\" & ActiveDocument.name
End Sub
Sub DeleteList()                                                'リスト項目の削除の処理
    If WordAssist.WordList.ListCount <> WordAssist.WordList.ListIndex + 1 Then
        changed = True
        WordAssist.WordList.RemoveItem (WordAssist.WordList.ListIndex)
        Index = Index - 1
        Dim i As Integer
            For i = WordAssist.WordList.ListIndex To wordnumber
                mean(i) = mean(i + 1)
                If mean(i) = "" Then
                    Exit For
                End If
            Next i
            WordAssist.meaningLabel.Caption = mean(WordAssist.WordList.ListIndex)
    Else
        If WordAssist.WordList.ListIndex <> -1 Then
            WordAssist.WordList.RemoveItem (WordAssist.WordList.ListIndex)
            Index = Index - 1
            If WordAssist.WordList.ListCount = 0 Then
                mean(0) = ""
                WordAssist.meaningLabel.Caption = ""
                changed = True
            End If
        End If
        
    End If
    Call count_words
End Sub
Sub excludex()
    WordAssist.CommandButton1.Visible = False
    Dim cnsFILENAME, Fname, tmpname As String
    cnsFILENAME = "\" & ActiveDocument.name & ".csv"
    tmpname = ActiveDocument.Path & "\" & ActiveDocument.name
    Dim intFF As Integer            ' FreeFile値
    Dim X(1 To 2) As Variant        ' 書き出すレコード内容
    Dim COL As Long                 ' カラム(Work)
'    Application.StatusBar = "ファイルを保存しています。"
    If saved = False Then
        Dim xlApp As Object
        Set xlApp = CreateObject("Excel.Application")
        Fname = xlApp.Application.GetSaveAsFilename(DefaultFilePath & ".csv", "csvファイル(*.csv),*.csv")
        If Fname = False Then GoTo labelE
    Else
        Fname = DefaultFilePath
    End If
    
    ' FreeFile値の取得(以降この値で入出力する)
    intFF = FreeFile
    ' 指定ファイルをOPEN(出力モード)
    If Dir(Fname) <> "" And saved = False Then
        Dim rc As Long
        rc = MsgBox("同名のファイルがすでに存在します。ファイルを上書きしますか？", vbYesNo + vbExclamation, "上書きの確認")
    End If

    If rc = vbYes Or Dir(Fname) = "" Or changed = True Then
        DefaultFilePath = Fname
        Open Fname For Output As #intFF
        ' 最終行まで繰り返す
        Dim i As Integer
        For i = 0 To WordAssist.WordList.ListCount - 1
            If mean(i) = "" Then Exit For
            X(1) = WordAssist.WordList.List(i)
            X(2) = mean(i)
            ' レコードを出力
            Print #intFF, X(1); ","; X(2)
            ' 行を加算
            Application.StatusBar = "ファイルを保存しています" & (i + 1) & "/" & WordAssist.WordList.ListCount
        Next i
        ' 指定ファイルをCLOSE
        Close #intFF
        If saved = False Then
            MsgBox Fname & "にcsvファイルを保存しました。", vbInformation
        Else
            Application.StatusBar = "csvファイルを上書き保存しました。"
'            Sleep 750
            Application.StatusBar = ""
        End If
        changed = False
        saved = True
    Else
    End If
labelE:
    If saved = True Then
        WordAssist.CommandButton1.Caption = "上書き保存"
    Else
        WordAssist.CommandButton1.Caption = "csv出力"
    End If
    WordAssist.CommandButton1.Visible = True
    Call count_words
End Sub
Sub includex()
    Dim OpenFileName, buf As String
    Dim tmp1 As Variant
    Dim xlApp As Object
    If Index <> 0 Then changed = True
    Setting.include.Visible = False
    Setting.OK.Visible = False
    Setting.Comment.Visible = True
    Set xlApp = CreateObject("Excel.Application")
'    ChDir ActiveDocument.Path & "\"
    OpenFileName = xlApp.GetOpenFilename("csvファイル,*.csv")
    
    DefaultFilePath = OpenFileName
    If OpenFileName = False Then GoTo labelEND
    saved = True
    If WordAssist.WordList.ListCount <> 0 Then
        changed = True
    Else
        changed = False
    End If
    Open OpenFileName For Input As #1
        Do Until EOF(1)
            Line Input #1, buf
            buf = Replace(buf, Chr(34), "")
            tmp1 = Split(buf, ",")
            WordAssist.WordList.AddItem (tmp1(0))
            mean(Index) = tmp1(1)
            Index = Index + 1
        Loop
    Close #1
    
labelEND:
    Setting.include.Visible = True
    Setting.OK.Visible = True
    Setting.Comment.Visible = False
    If saved = True Then
        WordAssist.CommandButton1.Caption = "上書き保存"
    Else
        WordAssist.CommandButton1.Caption = "csv出力"
    End If
    Call count_words
End Sub
Sub OpenURL()
    Dim WSH As Object
    Dim URL As String
    Dim word As String
    word = Selection.Text
    word = LCase(word)
    word = Replace(word, vbCr, "")
    word = Replace(word, " ", "+")
    Set WSH = CreateObject("Wscript.shell")
    URL = "https://www.google.co.jp/webhp?sourceid=chrome-instant&ion=1&espv=2&ie=UTF-8#q=" & word
    WSH.Run URL, 3
End Sub

