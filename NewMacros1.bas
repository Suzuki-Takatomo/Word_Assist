Attribute VB_Name = "NewMacros1"
Option Explicit
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Const wordnumber As Integer = 200                               '�o�^�ł���P�ꐔ
Public Searchengine As String                                   '�����Ώ�URL
Public Searchtag As String                                      'URL���̈Ӗ��̎擾����^�O
Public SearchIndex As Integer                                   '�^�O�̉��Ԗڂ̈Ӗ����擾���邩�̃C���f�b�N�X
Public Index As Integer                                         '�o�^������
Public mean(wordnumber) As String                               '�o�^���ꂽ�P��̈Ӗ�
Public changed, saved As Boolean
Public DefaultFilePath As String
Public page, row As Integer
Public backup_Sentence As String
Sub Search()
Attribute Search.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro1"
    Dim objIE As Object                                         'IE�̃I�u�W�F�N�g
    Dim i As Long
    Dim word As String                                          '�����Ώۃ��[�h
    Dim StartTime, StopTime As Variant
    StartTime = Timer
    Application.StatusBar = "�P����擾���Ă��܂��B"
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
         Application.StatusBar = "web�y�[�W�̉�����҂��Ă��܂��B"
         Set htmlDoc = objXML.createDocumentFromUrl(URL, vbNullString)
        Do While htmlDoc.readyState <> "complete"  '�ǂݍ��ݑ҂�
        htmlDoc.designMode = "on"
        DoEvents

        Loop
        Application.StatusBar = "�Ӗ����擾���Ă��܂��B"
        
        Dim IndexCount As Integer
        IndexCount = 0
        
            For Each objITEM In htmlDoc.getElementsByClassName(Searchtag)
                   Dim meaning As String
                   If objITEM.innerText <> "" Then                  '�Ӗ����󔒂ɂȂ��Ă��Ȃ���΁A
                       IndexCount = IndexCount + 1                  '�Ӗ����i�[���A�󔒂������A�J���}������(CSV�΍�)
                       If IndexCount = SearchIndex Then
                           meaning = objITEM.innerText
                           meaning = Replace(meaning, vbCrLf, "")
                           meaning = Replace(meaning, ",", "/")
                           mean(Index) = meaning
                           Index = Index + 1                         '�P�ꐔ��ǉ�
                           WordAssist.WordList.AddItem word
                           
                           changed = True
                           Exit For
                       End If
                   End If
second:
             Next
                       
                       If meaning = "" Then
                         MsgBox ("������܂���ł���")             '������Ȃ������ꍇ�̃��b�Z�[�W�{�b�N�X�̕\��
                         GoTo Next_step
                       End If
         
        If Index <> 0 Then
            WordAssist.WordList.selected(Index - 1) = True           '�V�����ǉ����ꂽ�P��̑I���B
        End If
Next_step:
    Set objITEM = Nothing
    Set htmlDoc = Nothing
    Set objXML = Nothing
    StopTime = Timer
    Application.StatusBar = "��������:" & Str(StopTime - StartTime) & "�b"
    Call count_words
    Sleep 750
    Application.StatusBar = ""
    
    Exit Sub
myError:
    MsgBox "�C���^�[�l�b�g�I�u�W�F�N�g�������ɃG���[���������܂����B", vbExclamation
End Sub
Sub Assist()
    WordAssist.Show vbModeless
    changed = False
    saved = False
    row = -1
    DefaultFilePath = ActiveDocument.Path & "\" & ActiveDocument.name
End Sub
Sub DeleteList()                                                '���X�g���ڂ̍폜�̏���
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
    Dim intFF As Integer            ' FreeFile�l
    Dim X(1 To 2) As Variant        ' �����o�����R�[�h���e
    Dim COL As Long                 ' �J����(Work)
'    Application.StatusBar = "�t�@�C����ۑ����Ă��܂��B"
    If saved = False Then
        Dim xlApp As Object
        Set xlApp = CreateObject("Excel.Application")
        Fname = xlApp.Application.GetSaveAsFilename(DefaultFilePath & ".csv", "csv�t�@�C��(*.csv),*.csv")
        If Fname = False Then GoTo labelE
    Else
        Fname = DefaultFilePath
    End If
    
    ' FreeFile�l�̎擾(�ȍ~���̒l�œ��o�͂���)
    intFF = FreeFile
    ' �w��t�@�C����OPEN(�o�̓��[�h)
    If Dir(Fname) <> "" And saved = False Then
        Dim rc As Long
        rc = MsgBox("�����̃t�@�C�������łɑ��݂��܂��B�t�@�C�����㏑�����܂����H", vbYesNo + vbExclamation, "�㏑���̊m�F")
    End If

    If rc = vbYes Or Dir(Fname) = "" Or changed = True Then
        DefaultFilePath = Fname
        Open Fname For Output As #intFF
        ' �ŏI�s�܂ŌJ��Ԃ�
        Dim i As Integer
        For i = 0 To WordAssist.WordList.ListCount - 1
            If mean(i) = "" Then Exit For
            X(1) = WordAssist.WordList.List(i)
            X(2) = mean(i)
            ' ���R�[�h���o��
            Print #intFF, X(1); ","; X(2)
            ' �s�����Z
            Application.StatusBar = "�t�@�C����ۑ����Ă��܂�" & (i + 1) & "/" & WordAssist.WordList.ListCount
        Next i
        ' �w��t�@�C����CLOSE
        Close #intFF
        If saved = False Then
            MsgBox Fname & "��csv�t�@�C����ۑ����܂����B", vbInformation
        Else
            Application.StatusBar = "csv�t�@�C�����㏑���ۑ����܂����B"
'            Sleep 750
            Application.StatusBar = ""
        End If
        changed = False
        saved = True
    Else
    End If
labelE:
    If saved = True Then
        WordAssist.CommandButton1.Caption = "�㏑���ۑ�"
    Else
        WordAssist.CommandButton1.Caption = "csv�o��"
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
    OpenFileName = xlApp.GetOpenFilename("csv�t�@�C��,*.csv")
    
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
        WordAssist.CommandButton1.Caption = "�㏑���ۑ�"
    Else
        WordAssist.CommandButton1.Caption = "csv�o��"
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

