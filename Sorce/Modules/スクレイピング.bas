Attribute VB_Name = "�X�N���C�s���O"
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Dim driver As New Selenium.WebDriver
'Dim driver As Object


'**************************************************************************************************
' * �f�[�^����
' *
' * @author Bunpei.Koizumi<koizumi.bunpei@trans-cosmos.co.jp>
'**************************************************************************************************
Function SheetDataDelete()
  
  Call Library.delSheetData(5)
  Cells.FormatConditions.Delete
  
  Range("E2") = "������������������������������������"
  Range("E3") = ""
  
  endLine = Cells(3, Columns.count).End(xlToLeft).Column - 3
  Range(Cells(4, 10), Cells(4, endLine)).ClearContents

End Function


'**************************************************************************************************
' * ���[�U�[�t�H�[���̕\��
' *
'**************************************************************************************************
Function dispUserForm()
  With KOETOL_Form
    .StartUpPosition = 0
    .Top = Application.Top + (ActiveWindow.Width / 4)
    .Left = Application.Left + (ActiveWindow.Height / 2)
    .Caption = "KOETOL"
  End With
  KOETOL_Form.Show vbModeless
End Function


'**************************************************************************************************
' * Web�f�[�^�擾
' *
' * @author Bunpei.Koizumi<koizumi.bunpei@trans-cosmos.co.jp>
'**************************************************************************************************
Function webScraping()
  
  Dim rc As Long

'  On Error GoTo catchError
  
  Call init.setting
  Call Library.startScript
  Call ProgressBar.showStart

  Set driver = CreateObject("Selenium.WebDriver")

  'Call �����e�i���X.���p���O�擾("KOETOL�F" & setVal("SiteName"))
  Call ProgressBar.showCount("Web�f�[�^�擾", 0, 10, "")
  
  'chromeDriver �̍X�V�m�F-------------------------------------------------------------------------
  rc = Shell(binPath & "\SeleniumBasic\updateChromeDriver.bat", vbNormalFocus)
  
  '�����f�[�^�폜----------------------------------------------------------------------------------
  sheetKoetol.Select
  
  Call SheetDataDelete
  
'  Call ProgressBar.showCount("Web�f�[�^�擾", 0, 10, setVal("SiteName"))
  Select Case setVal("SiteName")
    Case "@�R�X��"
        Call site_cosme
        
    Case "Amazon"
        Call site_Amazon
        
    Case "�y�V"
        Call site_rakuten
        
    Case "���i.com"
        Call site_kakakuCom
        
    Case "�R�X���f�l�b�g"
        Call site_cosmedeNet
  End Select
    
  Call �����ݒ�

  Application.Goto Reference:=Range("I5"), Scroll:=True

  '�I����ʕ\��
  With driver
    .Get "file:///" & openingHTML("Koetol") & "\end.html"
    .Wait 7000
    .Close
    .Quit
  End With
  Set driver = Nothing
  
  Call Library.endScript
  Call ProgressBar.showEnd
  
  Exit Function

'�G���[������=====================================================================================
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)

End Function


'**************************************************************************************************
' * �O���t�p�f�[�^�W�v
' *
' * @author Bunpei.Koizumi<koizumi.bunpei@trans-cosmos.co.jp>
'**************************************************************************************************
Function �O���t�p�f�[�^�W�v()
 
  Dim endRowLine As Long, endColLine As Long
  Dim line As Long, rowLine As Long, colLine As Long
  
  Call init.setting
  
  ' �v���O���X�o�[�̕\���J�n
'  Call ProgressBar.showStart
  
  '�����f�[�^�폜
  sheetGraf.Range("E7:E9").ClearContents
  sheetGraf.Range("O7:O8").ClearContents
  sheetGraf.Range("D27:G33").ClearContents
  sheetGraf.Range("C42:G71").ClearContents
  sheetGraf.Range("D85:J91").ClearContents
  sheetGraf.Range("C101:K130").ClearContents
  
  endRowLine = sheetKoetol.Cells(Rows.count, 3).End(xlUp).Row
  
  '�j���䗦
'  ProgressBar.showCount "�j���䗦�W�v", 1, 10, "�������E�E�E"
  sheetGraf.Range("E7") = WorksheetFunction.CountIf(sheetKoetol.Range("F5:F" & endRowLine), "�j��")
  sheetGraf.Range("E8") = WorksheetFunction.CountIf(sheetKoetol.Range("F5:F" & endRowLine), "����")
  sheetGraf.Range("E9") = WorksheetFunction.CountIf(sheetKoetol.Range("F5:F" & endRowLine), "�s��")
  For i = 7 To 9
    If sheetGraf.Cells(i, 5) = 0 Then
      sheetGraf.Cells(i, 5) = ""
    End If
  Next
  
  
  '�l�K�|�W�䗦
'  ProgressBar.showCount "�l�K�|�W�䗦�W�v", 1, 10, "�������E�E�E"
  sheetGraf.Range("O7") = WorksheetFunction.Sum(sheetKoetol.Range("AY5:AY" & endRowLine))
  sheetGraf.Range("O8") = WorksheetFunction.Sum(sheetKoetol.Range("AZ5:AZ" & endRowLine))
  For i = 7 To 8
    If sheetGraf.Cells(i, 15) = 0 Then
      sheetGraf.Cells(i, 15) = " "
    End If
  Next
  
  '�N��E����
'  ProgressBar.showCount "�N��䗦�W�v", 1, 10, "�������E�E�E"
  sheetGraf.Range("D27") = WorksheetFunction.CountIf(sheetKoetol.Range("G5:G" & endRowLine), "10��")
  sheetGraf.Range("D28") = WorksheetFunction.CountIf(sheetKoetol.Range("G5:G" & endRowLine), "20��")
  sheetGraf.Range("D29") = WorksheetFunction.CountIf(sheetKoetol.Range("G5:G" & endRowLine), "30��")
  sheetGraf.Range("D30") = WorksheetFunction.CountIf(sheetKoetol.Range("G5:G" & endRowLine), "40��")
  sheetGraf.Range("D31") = WorksheetFunction.CountIf(sheetKoetol.Range("G5:G" & endRowLine), "50��")
  sheetGraf.Range("D32") = WorksheetFunction.CountIf(sheetKoetol.Range("G5:G" & endRowLine), "60��") + _
                           WorksheetFunction.CountIf(sheetKoetol.Range("G5:G" & endRowLine), "70��")
  sheetGraf.Range("D33") = WorksheetFunction.CountIf(sheetKoetol.Range("G5:G" & endRowLine), "�s��")
  
  sheetGraf.Range("E27") = WorksheetFunction.CountIfs(sheetKoetol.Range("F5:F" & endRowLine), "�j��", sheetKoetol.Range("G5:G" & endRowLine), "10��")
  sheetGraf.Range("E28") = WorksheetFunction.CountIfs(sheetKoetol.Range("F5:F" & endRowLine), "�j��", sheetKoetol.Range("G5:G" & endRowLine), "20��")
  sheetGraf.Range("E29") = WorksheetFunction.CountIfs(sheetKoetol.Range("F5:F" & endRowLine), "�j��", sheetKoetol.Range("G5:G" & endRowLine), "30��")
  sheetGraf.Range("E30") = WorksheetFunction.CountIfs(sheetKoetol.Range("F5:F" & endRowLine), "�j��", sheetKoetol.Range("G5:G" & endRowLine), "40��")
  sheetGraf.Range("E31") = WorksheetFunction.CountIfs(sheetKoetol.Range("F5:F" & endRowLine), "�j��", sheetKoetol.Range("G5:G" & endRowLine), "50��")
  sheetGraf.Range("E32") = WorksheetFunction.CountIfs(sheetKoetol.Range("F5:F" & endRowLine), "�j��", sheetKoetol.Range("G5:G" & endRowLine), "60��")
  sheetGraf.Range("E33") = WorksheetFunction.CountIfs(sheetKoetol.Range("F5:F" & endRowLine), "�j��", sheetKoetol.Range("G5:G" & endRowLine), "�s��")
  
  sheetGraf.Range("F27") = WorksheetFunction.CountIfs(sheetKoetol.Range("F5:F" & endRowLine), "����", sheetKoetol.Range("G5:G" & endRowLine), "10��")
  sheetGraf.Range("F28") = WorksheetFunction.CountIfs(sheetKoetol.Range("F5:F" & endRowLine), "����", sheetKoetol.Range("G5:G" & endRowLine), "20��")
  sheetGraf.Range("F29") = WorksheetFunction.CountIfs(sheetKoetol.Range("F5:F" & endRowLine), "����", sheetKoetol.Range("G5:G" & endRowLine), "30��")
  sheetGraf.Range("F30") = WorksheetFunction.CountIfs(sheetKoetol.Range("F5:F" & endRowLine), "����", sheetKoetol.Range("G5:G" & endRowLine), "40��")
  sheetGraf.Range("F31") = WorksheetFunction.CountIfs(sheetKoetol.Range("F5:F" & endRowLine), "����", sheetKoetol.Range("G5:G" & endRowLine), "50��")
  sheetGraf.Range("F32") = WorksheetFunction.CountIfs(sheetKoetol.Range("F5:F" & endRowLine), "����", sheetKoetol.Range("G5:G" & endRowLine), "60��")
  sheetGraf.Range("F33") = WorksheetFunction.CountIfs(sheetKoetol.Range("F5:F" & endRowLine), "����", sheetKoetol.Range("G5:G" & endRowLine), "�s��")
  
  sheetGraf.Range("G27") = WorksheetFunction.CountIfs(sheetKoetol.Range("F5:F" & endRowLine), "�s��", sheetKoetol.Range("G5:G" & endRowLine), "10��")
  sheetGraf.Range("G28") = WorksheetFunction.CountIfs(sheetKoetol.Range("F5:F" & endRowLine), "�s��", sheetKoetol.Range("G5:G" & endRowLine), "20��")
  sheetGraf.Range("G29") = WorksheetFunction.CountIfs(sheetKoetol.Range("F5:F" & endRowLine), "�s��", sheetKoetol.Range("G5:G" & endRowLine), "30��")
  sheetGraf.Range("G30") = WorksheetFunction.CountIfs(sheetKoetol.Range("F5:F" & endRowLine), "�s��", sheetKoetol.Range("G5:G" & endRowLine), "40��")
  sheetGraf.Range("G31") = WorksheetFunction.CountIfs(sheetKoetol.Range("F5:F" & endRowLine), "�s��", sheetKoetol.Range("G5:G" & endRowLine), "50��")
  sheetGraf.Range("G32") = WorksheetFunction.CountIfs(sheetKoetol.Range("F5:F" & endRowLine), "�s��", sheetKoetol.Range("G5:G" & endRowLine), "60��")
  sheetGraf.Range("G33") = WorksheetFunction.CountIfs(sheetKoetol.Range("F5:F" & endRowLine), "�s��", sheetKoetol.Range("G5:G" & endRowLine), "�s��")
  
  For i = 27 To 33
    For j = 4 To 7
      If sheetGraf.Cells(i, j) = 0 Then
        sheetGraf.Cells(i, j) = ""
      End If
    Next
  Next
  
  '���R�~���e
  line = 42
  colLine = 5
  endColLine = sheetKoetol.Cells(3, Columns.count).End(xlToLeft).Column - 3
  
  For rowLine = 20 To endColLine
    
    If sheetKoetol.Cells(3, rowLine) <> "" Then
      sheetGraf.Range("C" & line) = sheetKoetol.Cells(3, rowLine)
      sheetGraf.Range("D" & line) = WorksheetFunction.Sum(sheetKoetol.Range(sheetKoetol.Cells(colLine, rowLine), sheetKoetol.Cells(endRowLine, rowLine)))
      
      count01 = 0
      count02 = 0
      count03 = 0
        
      For i = 3 To endRowLine
        If sheetKoetol.Cells(i, rowLine) = 1 Then
          If sheetKoetol.Range("F" & i) = "�j��" Then
            count01 = count01 + 1
          ElseIf sheetKoetol.Range("F" & i) = "����" Then
            count02 = count02 + 1
          ElseIf sheetKoetol.Range("F" & i) = "�s��" Then
            count03 = count03 + 1
          End If
        End If
      Next i
      sheetGraf.Range("E" & line) = count01
      sheetGraf.Range("F" & line) = count02
      sheetGraf.Range("G" & line) = count03
      
      line = line + 1
    End If
  Next rowLine
  
  For i = 42 To 71
    For j = 4 To 7
      If sheetGraf.Cells(i, j) = 0 Then
        sheetGraf.Cells(i, j) = ""
      End If
    Next
  Next
  
  '�N��ʕ]��
  line = 85
  For Each colLineName In Split("10��,20��,30��,40��,50��,60��,�s��", ",")
    sheetGraf.Range("D" & line) = WorksheetFunction.CountIfs(sheetKoetol.Range("H5:H" & endRowLine), "1", sheetKoetol.Range("G5:G" & endRowLine), colLineName)
    sheetGraf.Range("E" & line) = WorksheetFunction.CountIfs(sheetKoetol.Range("H5:H" & endRowLine), "2", sheetKoetol.Range("G5:G" & endRowLine), colLineName)
    sheetGraf.Range("F" & line) = WorksheetFunction.CountIfs(sheetKoetol.Range("H5:H" & endRowLine), "3", sheetKoetol.Range("G5:G" & endRowLine), colLineName)
    sheetGraf.Range("G" & line) = WorksheetFunction.CountIfs(sheetKoetol.Range("H5:H" & endRowLine), "4", sheetKoetol.Range("G5:G" & endRowLine), colLineName)
    sheetGraf.Range("H" & line) = WorksheetFunction.CountIfs(sheetKoetol.Range("H5:H" & endRowLine), "5", sheetKoetol.Range("G5:G" & endRowLine), colLineName)
    sheetGraf.Range("I" & line) = WorksheetFunction.CountIfs(sheetKoetol.Range("H5:H" & endRowLine), "6", sheetKoetol.Range("G5:G" & endRowLine), colLineName)
    sheetGraf.Range("J" & line) = WorksheetFunction.CountIfs(sheetKoetol.Range("H5:H" & endRowLine), "7", sheetKoetol.Range("G5:G" & endRowLine), colLineName)

    If sheetGraf.Range("D" & line) = 0 Then
      sheetGraf.Range("D" & line) = ""
    End If
    If sheetGraf.Range("E" & line) = 0 Then
      sheetGraf.Range("E" & line) = ""
    End If
    If sheetGraf.Range("F" & line) = 0 Then
      sheetGraf.Range("F" & line) = ""
    End If
    If sheetGraf.Range("G" & line) = 0 Then
      sheetGraf.Range("G" & line) = ""
    End If
    If sheetGraf.Range("H" & line) = 0 Then
      sheetGraf.Range("H" & line) = ""
    End If
    If sheetGraf.Range("I" & line) = 0 Then
      sheetGraf.Range("I" & line) = ""
    End If
    If sheetGraf.Range("J" & line) = 0 Then
      sheetGraf.Range("J" & line) = ""
    End If
    
    line = line + 1
  Next colLineName


  '�N��ʌ��R�~���e
  Dim serachWord As String
  Dim count As Long
  
  line = 101
  
  For colLine = 20 To endColLine
    serachWord = sheetKoetol.Cells(3, colLine)
    sheetGraf.Range("C" & line) = serachWord
    count = 5
    
    For Each colLineName In Split("10��,20��,30��,40��,50��,60��,�s��", ",")
      sheetGraf.Cells(line, count) = WorksheetFunction.CountIfs(sheetKoetol.Range(sheetKoetol.Cells(5, colLine), sheetKoetol.Cells(endRowLine, colLine)), "1", sheetKoetol.Range("G5:G" & endRowLine), colLineName)
      
      If sheetGraf.Cells(line, count) = 0 Then
        sheetGraf.Cells(line, count) = ""
      End If
      sheetGraf.Range("D" & line) = sheetGraf.Range("D" & line) + sheetGraf.Cells(line, count)
      
      If sheetGraf.Range("D" & line) = 0 Then
        sheetGraf.Range("D" & line) = ""
      End If
      
      count = count + 1
    Next colLineName
    
    line = line + 1
  Next
  
  '�I������----------------------------------------------------------------------------------------------------
  sheetGraf.Activate
  Application.Goto Reference:=Range("A1"), Scroll:=True
  
'  Call ProgressBar.ProgShowClose

End Function


'**************************************************************************************************
' * �����񌟍�
' *
' * @author Bunpei.Koizumi<koizumi.bunpei@trans-cosmos.co.jp>
'**************************************************************************************************
Function setWordAnalyze()
  
  Dim endRowLine As Long, line As Long
  Dim tmp As Variant
  Dim searchWord As Variant
  Dim result As Range
  Dim colLineName As Variant
  Dim firstAddress  As String
  
  Set sheetKoetol = sheetKoetol
  endRowLine = sheetKoetol.Cells(Rows.count, 3).End(xlUp).Row
  
  '�����f�[�^�폜
  sheetKoetol.Range("J5:AW" & endRowLine).ClearContents

  For Each colLineName In Split("H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z,AA,AB,AC,AD,AE,AF", ",")
  
    sheetKoetol.Range(colLineName & "1") = Replace(sheetKoetol.Range(colLineName & "1"), "�@", " ")
    sheetKoetol.Range(colLineName & "1") = Replace(sheetKoetol.Range(colLineName & "1"), vbCrLf, " ")
    
    
    tmp = Split(sheetKoetol.Range(colLineName & "1"), " ")
    For Each searchWord In tmp
      With sheetKoetol.Range("G5:G" & endRowLine)
      
        Set result = .Find(What:=searchWord, LookIn:=xlValues, LookAt:=xlPart)
        If Not result Is Nothing Then
          firstAddress = result.Address
  
          Do
            sheetKoetol.Range(colLineName & result.Row) = 1
            Set result = .FindNext(result)
            If result Is Nothing Then Exit Do
          Loop While result.Address <> firstAddress
  
        End If
      
      End With
    Next searchWord
  Next colLineName

  Application.Goto Reference:=Range("A1"), Scroll:=True

End Function



'**************************************************************************************************
' * �d���E�w��L�[���[�h�폜�p���[�U�[�t�H�[���̕\��
' *
'**************************************************************************************************
Function �w���叜�O�t�H�[���\��()
  
  Dim dbCon As ADODB.Connection
  Dim DBRecordset As ADODB.Recordset
  Dim count As Integer
  Dim QueryString As String
  
  Call init.setting
  
  With keywordForm
    .StartUpPosition = 0
    .Top = Application.Top + (ActiveWindow.Width / 4)
    .Left = Application.Left + (ActiveWindow.Height / 2)
    .Caption = "KOETOL:�d���E�w��L�[���[�h�폜"
    .keywordList = Library.getRegistry("koetol_delKeyword")
  End With
  
  keywordForm.Show

End Function


'**************************************************************************************************
' * �d���E�w��L�[���[�h�폜
' *
' * @author Bunpei.Koizumi<koizumi.bunpei@trans-cosmos.co.jp>
'**************************************************************************************************
Function �w���叜�O()

  Dim line As Long, endLine As Long
  Dim delWord As Variant

  Call init.setting
  Call Library.startScript
  Call ProgressBar.showStart
  
  sheetKoetol.Select
  endLine = sheetKoetol.Cells(Rows.count, 9).End(xlUp).Row
  
  '�d���폜
  ActiveSheet.Range("$C$5:$AZ$" & endLine).RemoveDuplicates Columns:=7, Header:=xlNo
  '��U�w�i�F���O���[�ɕύX
  Rows("5:" & Rows.count).Interior.Color = RGB(242, 242, 242)
  Call �����ݒ�
  
  '�w�胏�[�h�폜
  For line = 5 To endLine
    For Each delWord In Split(delStringList, vbCrLf)
      Call ProgressBar.showCount("�w��L�[���[�h�폜", line, endLine, CStr(delWord))
      
      If sheetKoetol.Range("I" & line) Like "*" & delWord & "*" And delWord <> "" Then
        sheetKoetol.Range("I" & line).Interior.Color = RGB(255, 199, 206)
        sheetKoetol.Range("I" & line).Select
        Call Library.setFontClor(delWord, RGB(156, 0, 6), True)
      End If
    Next
  Next
  
  Call ProgressBar.showEnd
  Call Library.endScript(True)
End Function



'**************************************************************************************************
' * �L�[���[�h���
' *
' * @author Bunpei.Koizumi<koizumi.bunpei@trans-cosmos.co.jp>
'**************************************************************************************************
Function �L�[���[�h���()
  
  Dim endRowLine As Long, line As Long, count As Long
  Dim tmp As Variant
  Dim searchWord As Variant
  Dim result As Range
  Dim colLineName As Variant
  Dim firstAddress  As String
  
  Call init.setting
  count = 1
  endRowLine = sheetKoetol.Cells(Rows.count, 3).End(xlUp).Row
  
  If MsgBox("�ݒ�ς݃f�[�^���N���A���܂����H", vbYesNo + vbExclamation + vbDefaultButton2) = vbYes Then
    Range("J5:AW" & endRowLine).ClearContents
  End If
  
  Call ProgressBar.showStart
  For Each colLineName In Split("J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z,AA,AB,AC,AD,AE,AF,AG,AH,AI,AJ,AK,AL,AM,AN,AO,AP,AQ,AR,AS,AT,AU,AV,AW", ",")
    Call ProgressBar.showCount("�L�[���[�h���", count, 40, colLineName & "��@�������E�E�E")
    count = count + 1

  
    '��؂蕶���𓝈�
    sheetKoetol.Range(colLineName & "1") = Replace(sheetKoetol.Range(colLineName & "1"), "�@", " ")
    sheetKoetol.Range(colLineName & "1") = Replace(sheetKoetol.Range(colLineName & "1"), vbCrLf, " ")
    
    tmp = Split(sheetKoetol.Range(colLineName & "1"), " ")
    For Each searchWord In tmp
      With sheetKoetol.Range("I5:I" & endRowLine)
      
        Set result = .Find(What:=searchWord, LookIn:=xlValues, LookAt:=xlPart)
        If Not result Is Nothing Then
          firstAddress = result.Address
  
          Do
            sheetKoetol.Range(colLineName & result.Row) = 1
            Set result = .FindNext(result)
            If result Is Nothing Then Exit Do
          Loop While result.Address <> firstAddress
  
        End If
      
      End With
    Next searchWord
  Next colLineName

  Call ProgressBar.showEnd
  Application.Goto Reference:=Range("B1"), Scroll:=True

End Function

'**************************************************************************************************
' * �����ݒ�
' *
' * @author Bunpei.Koizumi<koizumi.bunpei@trans-cosmos.co.jp>
'**************************************************************************************************
Function �����ݒ�()
  Dim endRowLine As Long, line As Long
  
  Call ProgressBar.showCount("�����ݒ�", 0, 10, "")
  
  '�Z���̏����ݒ�
  endRowLine = sheetKoetol.Cells(Rows.count, 3).End(xlUp).Row
  
  '�w�i�F���Ȃ��ɂ���
  sheetKoetol.Range("C5:AZ" & endRowLine).Interior.Pattern = xlNone
  
  '���e���̏����ݒ�
  Range("E5:E" & endRowLine).NumberFormatLocal = "yyyy/mm/dd"
    
  '�s�̍�������
  Rows("5:" & endRowLine).RowHeight = 30
  
  
  '4�s�ڂ̍��v�l�Z�o
  For Each colLineName In Split("J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z,AA,AB,AC,AD,AE,AF,AG,AH,AI,AJ,AK,AL,AM,AN,AO,AP,AQ,AR,AS,AT,AU,AV,AW", ",")
    Range(colLineName & "4") = "=SUM(" & colLineName & "5:" & colLineName & endRowLine & ")"
  Next colLineName
  
  Range("J4:AH4").NumberFormatLocal = "G/�W��"
  
  
  'AG:AI��̃J�E���g
  For line = 5 To endRowLine
    sheetKoetol.Range("AX" & line).FormulaR1C1 = "=COUNTIF(RC[-40]:RC[-31],1)"
    sheetKoetol.Range("AY" & line).FormulaR1C1 = "=COUNTIF(RC[-31]:RC[-17],1)"
    sheetKoetol.Range("AZ" & line).FormulaR1C1 = "=COUNTIF(RC[-17]:RC[-3],1)"
  Next line
  
  
  '�蓮���͕����̕\���`���ݒ�
  sheetKoetol.Range("J5" & ":AW" & endRowLine).Select
  With Selection
      .NumberFormatLocal = "[=1]""��"""
      .Validation.Delete
      .Validation.Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="1", Formula2:="1"
      .Validation.IgnoreBlank = True
      .Validation.InCellDropdown = True
      .Validation.IMEMode = xlIMEModeDisable
      .Validation.ShowInput = True
      .Validation.ShowError = True
      .Validation.ErrorTitle = ""
      .Validation.ErrorMessage = "1�̂ݓ��͉\�ł�"
      .HorizontalAlignment = xlCenter
      .VerticalAlignment = xlCenter
  End With
   
    
  '�r���ݒ�
  Range("C4:AZ" & endRowLine).Select
  With Selection.Borders
      .LineStyle = xlContinuous
      .Color = RGB(128, 128, 128)
  End With
  
  Range("J2:J" & endRowLine).Select
  With Selection.Borders(xlEdgeLeft)
    .LineStyle = xlDouble
    .Color = RGB(128, 128, 128)
  End With
  
  Range("T2:T" & endRowLine).Select
  With Selection.Borders(xlEdgeLeft)
    .LineStyle = xlDouble
    .Color = RGB(128, 128, 128)
  End With
  
  Range("AI2:AI" & endRowLine).Select
  With Selection.Borders(xlEdgeLeft)
    .LineStyle = xlDouble
    .Color = RGB(128, 128, 128)
  End With
  Range("AX2:AX" & endRowLine).Select
  With Selection.Borders(xlEdgeLeft)
    .LineStyle = xlDouble
    .Color = RGB(128, 128, 128)
  End With
  
  
End Function


'**************************************************************************************************
' * �����_�����o
' *
' * @author Bunpei.Koizumi<koizumi.bunpei@trans-cosmos.co.jp>
'**************************************************************************************************
Function �����_�����o()

  Dim endRowLine As Long, line As Long, lineCount As Long, randomLine As Long
  Dim count As Long, minCnt As Long, maxCnt As Long, setRowLine As Long
  
  Call Library.startScript
  Call init.setting
  
  minCnt = 5
  maxCnt = 5
  line = sheetKoetol.Cells(Rows.count, 3).End(xlUp).Row
  
'Re�����_�����o:
  sheetKoetol.Select
  maxCnt = Library.getRegistry("koetol_DataCount")
  
  If (line - minCnt + 1) <= maxCnt Then
    Call Library.showNotice(410, , True)
  End If
  
  Do While WorksheetFunction.CountA(Range("A5:A" & line)) < maxCnt
    Call ProgressBar.showCount("�����_�����o", line, maxCnt, "�@�������E�E�E")

    randomLine = Library.makeRandomNo(minCnt, line)
    Range("A" & randomLine) = 1
  Loop
  
  For count = minCnt To line
    Call ProgressBar.showCount("�����_�����o", count, line, "�@�������E�E�E")
    
    If Range("A" & count) <> 1 And Range("D" & count) <> "" Then
      Rows(count & ":" & count).Select
      Selection.Delete Shift:=xlUp
      'Rows(count & ":" & count).Delete Shift:=xlUp
      count = count - 1
    End If
  Next
  Range("A5:A" & count) = ""
  
  Application.Goto Reference:=Range("A1"), Scroll:=True
  
  Call Library.endScript
End Function

'*******************************************************************************************************
' * Chrome�N��
' *
' * @author Bunpei.Koizumi<Koizumi.Bunpei@trans-cosmos.co.jp>
'*******************************************************************************************************
Function Chrome�N��()
  
  'chrome �g���@�\�ǂݍ���
  With driver
    .AddArgument ("--user-data-dir=" & BrowserProfiles("default") & "\BrowserProfile")
    .AddArgument ("--disable-session-crashed-bubble")
    
    If setVal("InstNetwork") = True Then
      .AddArgument ("--proxy-server=tci-proxy.trans-cosmos.co.jp:8080")
    End If
    .AddArgument ("--window-size=1200,600")
    .AddArgument ("--app=file:///" & openingHTML("Koetol") & "\index.html")
    
    .start "chrome"
    .Wait 6000
    .Get setVal("URL")
  End With

End Function


'*******************************************************************************************************
' * www.amazon.co.jp
' *
' * @author Bunpei.Koizumi<Koizumi.Bunpei@trans-cosmos.co.jp>
'*******************************************************************************************************
Function site_Amazon()
  Dim elements As WebElements
  Dim line As Long, endLine As Long
  Dim exitFlg As Boolean
  Dim getInnerHtmlString As Variant
  Dim myBy As New By
  Dim commentTitle As String
  Dim Red As Long, Green As Long, Blue As Long
  
  
  exitFlg = False
'  On Error GoTo catchError
  Call init.setting
  Call Chrome�N��
  
  '�^�C�g���擾
  sheetKoetol.Range("E3") = driver.title

reCrawl:
  driver.Wait 1000
  line = sheetKoetol.Cells(Rows.count, 3).End(xlUp).Row + 1

  '���R�~���擾
  Set elements = driver.FindElementById("cm_cr-review_list").FindElementsByClass("aok-relative")
  For i = 1 To elements.count
    Call ProgressBar.showCount("Web�f�[�^�擾", line - 4, CLng(setVal("DataCount") * setVal("DataRate")), "")
    
    getInnerHtmlString = ""
    
    '�T�C�g------------------------------------------------------------------------------------------
    sheetKoetol.Range("C" & line) = line - 4
    sheetKoetol.Range("D" & line) = setVal("SiteName")
    
    '���e��----------------------------------------------------------------------------------------
    getInnerHtmlString = elements.Item(i).FindElementByClass("review-date").Text
    getInnerHtmlString = Replace(getInnerHtmlString, "�ɓ��{�Ń��r���[�ς�", "")
    getInnerHtmlString = Format(DateValue(getInnerHtmlString), "YYYY/MM/DD")
    sheetKoetol.Range("E" & line) = getInnerHtmlString
    sheetKoetol.Range("E" & line).NumberFormatLocal = "yyyy/mm/dd"
    
    '����------------------------------------------------------------------------------------------
    sheetKoetol.Range("F" & line) = "�s��"
    
    '�N��------------------------------------------------------------------------------------------
    sheetKoetol.Range("G" & line) = "�s��"
    
    '�]��------------------------------------------------------------------------------------------
    getInnerHtmlString = elements.Item(i).FindElementByClass("a-link-normal").Attribute("title")
    getInnerHtmlString = Replace(getInnerHtmlString, "5���̂���", "")
    sheetKoetol.Range("H" & line) = getInnerHtmlString
    sheetKoetol.Range("H" & line).NumberFormatLocal = "0_ ;[��]-0 "
    
    '�R�����g--------------------------------------------------------------------------------------
    commentTitle = ""
    getInnerHtmlString = ""
    If elements.Item(i).IsElementPresent(myBy.Class("review-title")) Then
      commentTitle = setVal("commentTitleStart") & elements.Item(i).FindElementByClass("review-title").Text & setVal("commentTitleEnd")
      getInnerHtmlString = commentTitle & vbCrLf
      
    End If
    getInnerHtmlString = getInnerHtmlString & Library.delCellLinefeed(elements.Item(i).FindElementByClass("review-text").Text)
    sheetKoetol.Range("I" & line) = getInnerHtmlString
    
    '�^�C�g�����{�[���h�ɂ���
    If commentTitle <> "" Then
      Call Library.getRGB(setVal("commentTitleColor"), Red, Green, Blue)
      With sheetKoetol.Range("I" & line).Characters(start:=1, Length:=Len(commentTitle)).Font
        .Bold = True
        .Color = RGB(Red, Green, Blue)
      End With
    End If
    
    '���y�[�W�J�ڔ���------------------------------------------------------------------------------
    If (Cells(Rows.count, 5).End(xlUp).Row - 4) >= setVal("DataCount") * setVal("DataRate") Then
      exitFlg = True
      Exit For
    Else
      exitFlg = False
    End If
    
    line = line + 1
  Next
  
  '���y�[�W�J��------------------------------------------------------------------------------------
  If exitFlg = True Then
    Call ProgressBar.showCount("Web�f�[�^�擾", 9, 10, "�I��")
    driver.Get "file:///" & openingHTML("Koetol") & "\pause.html"

  Else
    If driver.FindElementByClass("a-last").FindElementsByTag("a").count > 0 Then
      driver.FindElementByClass("a-last").FindElementByTag("a").Click
      driver.Wait 2000
      
      GoTo reCrawl
    End If
  End If

  Exit Function
'--------------------------------------------------------------------------------------------------
'�G���[�������̏���
'--------------------------------------------------------------------------------------------------
catchError:
   
  If Not driver Is Nothing Then
    driver.Close
    driver.Quit

    Set driver = Nothing
  End If

  Call Library.showNotice(Err.Number, Err.Description, True)
End Function


'*******************************************************************************************************
' * www.cosme.net
' *
' * @author Bunpei.Koizumi<Koizumi.Bunpei@trans-cosmos.co.jp>
'*******************************************************************************************************
Function site_cosme()
  Dim elements As WebElements
  Dim line As Long, endLine As Long
  Dim exitFlg As Boolean
  Dim getInnerHtmlString As Variant
  
'  On Error GoTo catchError
  
  exitFlg = False
  Call init.setting
  Call Chrome�N��
  
  
  '�^�C�g���擾
  sheetKoetol.Range("E3") = driver.title

  '�ŏ��̋L���̑�����ǂނ��N���b�N����
  Set elements = driver.FindElementsByClass("read-more")
  For i = 1 To elements.count
    getInnerHtmlString = elements.Item(i).Text
    If getInnerHtmlString Like "*[������ǂ�]*" Then
      elements.Item(i).FindElementByTag("a").Click
      driver.Wait 1000
      Exit For
    End If
  Next

reCrawl:
  driver.Wait 1000
  line = sheetKoetol.Cells(Rows.count, 3).End(xlUp).Row + 1

  '���R�~���擾
  DoEvents
  getInnerHtmlString = ""
  
  '�T�C�g------------------------------------------------------------------------------------------
  sheetKoetol.Range("C" & line) = line - 4
  sheetKoetol.Range("D" & line) = setVal("SiteName")
  
  '���e��----------------------------------------------------------------------------------------
  getInnerHtmlString = driver.FindElementByXPath("//*[@id='product-review-list']/div/div/div[2]/div[1]/p[2]").Text
'  getInnerHtmlString = Format(DateValue(getInnerHtmlString), "YYYY/MM/DD")
  sheetKoetol.Range("E" & line) = getInnerHtmlString
  sheetKoetol.Range("E" & line).NumberFormatLocal = "yyyy/mm/dd"
  
  '����------------------------------------------------------------------------------------------
  sheetKoetol.Range("F" & line) = "�s��"
  
  '�N��------------------------------------------------------------------------------------------
  getInnerHtmlString = driver.FindElementByXPath("//*[@id='product-review-list']/div/div/div[1]/div[2]/ul/li[1]").Text
  sheetKoetol.Range("G" & line) = getInnerHtmlString
  
  getInnerHtmlString = Replace(getInnerHtmlString, "��", "")
  getInnerHtmlString = Application.WorksheetFunction.RoundDown(getInnerHtmlString, -1)
  init.sheetKoetol.Range("G" & line) = getInnerHtmlString & "��"
    
  '�]��------------------------------------------------------------------------------------------
  getInnerHtmlString = driver.FindElementByClass("reviewer-rating").Text
  getInnerHtmlString = Replace(getInnerHtmlString, "�w���i", "")
  getInnerHtmlString = Replace(getInnerHtmlString, "���s�[�g", "")
  
  If getInnerHtmlString = "�]�����Ȃ�" Then
    sheetKoetol.Range("H" & line) = "�s��"
  Else
    sheetKoetol.Range("H" & line) = getInnerHtmlString
    sheetKoetol.Range("H" & line).NumberFormatLocal = "0_ ;[��]-0 "
  End If
  
  '�R�����g--------------------------------------------------------------------------------------
  Range("I" & line) = Library.delCellLinefeed(driver.FindElementByClass("read").Text)
  
  '���y�[�W�J�ڔ���------------------------------------------------------------------------------
  If (Cells(Rows.count, 5).End(xlUp).Row - 4) >= setVal("DataCount") * setVal("DataRate") Then
    exitFlg = True
  Else
    exitFlg = False
  End If
  
  line = line + 1
  
  '���y�[�W�J��------------------------------------------------------------------------------------
  If exitFlg = True Then
    Call ProgressBar.showCount("Web�f�[�^�擾", 9, 10, "�I��")
    driver.Get "file:///" & openingHTML("Koetol") & "\pause.html"

  Else
    getInnerHtmlString = driver.FindElementByClass("next").Text
    
    If driver.FindElementByClass("next").FindElementsByTag("a").count > 0 Then
      driver.FindElementByClass("next").FindElementByTag("a").Click
      driver.Wait 1000
      
      GoTo reCrawl
    End If
  End If

  Exit Function
'--------------------------------------------------------------------------------------------------
'�G���[�������̏���
'--------------------------------------------------------------------------------------------------
catchError:
   
  If Not driver Is Nothing Then
    driver.Close
    driver.Quit
    Set driver = Nothing
  End If

  Call Library.showNotice(Err.Number, Err.Description, True)

End Function


'*******************************************************************************************************
' * review.rakuten.co.jp
' *
' * @author Bunpei.Koizumi<Koizumi.Bunpei@trans-cosmos.co.jp>
'*******************************************************************************************************
Function site_rakuten()
  Dim elements As WebElements
  Dim line As Long, endLine As Long
  Dim exitFlg As Boolean
  Dim getInnerHtmlString As Variant
  Dim myBy As New By
  Dim commentTitle As String
  Dim Red As Long, Green As Long, Blue As Long

'  On Error GoTo catchError
  
  exitFlg = False
  Call init.setting
  Call Chrome�N��
  
  '�^�C�g���擾
  sheetKoetol.Range("E3") = driver.title

reCrawl:
  driver.Wait 1000
  line = sheetKoetol.Cells(Rows.count, 3).End(xlUp).Row + 1

  '���R�~���擾
  Set elements = driver.FindElementByClass("revRvwUserSecCnt").FindElementsByClass("revRvwUserSec")
  For i = 1 To elements.count
    Call ProgressBar.showCount("Web�f�[�^�擾", line - 4, CLng(setVal("DataCount") * setVal("DataRate")), "")

    getInnerHtmlString = ""
    
    '�T�C�g------------------------------------------------------------------------------------------
    sheetKoetol.Range("C" & line) = line - 4
    sheetKoetol.Range("D" & line) = setVal("SiteName")
    
    '���e��----------------------------------------------------------------------------------------
    getInnerHtmlString = elements.Item(i).FindElementByClass("revUserEntryDate").Text
    getInnerHtmlString = Format(DateValue(getInnerHtmlString), "YYYY/MM/DD")
    sheetKoetol.Range("E" & line) = getInnerHtmlString
    sheetKoetol.Range("E" & line).NumberFormatLocal = "yyyy/mm/dd"
    
    '���ʁE�N��------------------------------------------------------------------------------------------
    getInnerHtmlString = elements.Item(i).FindElementByClass("revUserFaceDtlTxt").FindElementByTag("span").Text
    
    If getInnerHtmlString = "" Then
      sheetKoetol.Range("F" & line) = "�s��"
      sheetKoetol.Range("G" & line) = "�s��"
    Else
      getInnerHtmlString = Split(getInnerHtmlString, " ")
      sheetKoetol.Range("F" & line) = getInnerHtmlString(1)
      sheetKoetol.Range("G" & line) = getInnerHtmlString(0)
    End If
    
    '�]��------------------------------------------------------------------------------------------
    getInnerHtmlString = elements.Item(i).FindElementByClass("revUserRvwerNum").Text
    sheetKoetol.Range("H" & line) = getInnerHtmlString
    sheetKoetol.Range("H" & line).NumberFormatLocal = "0_ ;[��]-0 "

    '�R�����g--------------------------------------------------------------------------------------
    commentTitle = ""
    getInnerHtmlString = ""
    If elements.Item(i).IsElementPresent(myBy.Class("revRvwUserEntryTtl")) Then
      commentTitle = setVal("commentTitleStart") & elements.Item(i).FindElementByClass("revRvwUserEntryTtl").Text & setVal("commentTitleEnd")
      getInnerHtmlString = commentTitle & vbCrLf
    End If
    getInnerHtmlString = getInnerHtmlString & Library.delCellLinefeed(elements.Item(i).FindElementByClass("revRvwUserEntryCmt").Text)
    sheetKoetol.Range("I" & line) = getInnerHtmlString
    
    '�^�C�g�����{�[���h�ɂ���
    If commentTitle <> "" Then
      Call Library.getRGB(setVal("commentTitleColor"), Red, Green, Blue)
      With sheetKoetol.Range("I" & line).Characters(start:=1, Length:=Len(commentTitle)).Font
        .Bold = True
        .Color = RGB(Red, Green, Blue)
      End With
    End If
              
    '���y�[�W�J�ڔ���------------------------------------------------------------------------------
    If (Cells(Rows.count, 5).End(xlUp).Row - 4) >= setVal("DataCount") * setVal("DataRate") Then
      exitFlg = True
      Exit For
    Else
      exitFlg = False
    End If
    
    line = line + 1
  Next
  
  '���y�[�W�J��------------------------------------------------------------------------------------
  If exitFlg = True Then
    Call ProgressBar.showCount("Web�f�[�^�擾", 9, 10, "�I��")
    driver.Get "file:///" & openingHTML("Koetol") & "\pause.html"

  Else
    getInnerHtmlString = driver.FindElementByClass("revThisPage").Text
    If getInnerHtmlString Like "*[����15��]*" Then
      driver.FindElementByXPath("//*[@id='revRvwSec']/div[1]/div/div[3]/div[16]/div/div/a[5]").Click
      driver.Wait 1000
      
      GoTo reCrawl
    End If
  End If

  Exit Function
'--------------------------------------------------------------------------------------------------
'�G���[�������̏���
'--------------------------------------------------------------------------------------------------
catchError:
   
  If Not driver Is Nothing Then
    driver.Close
    driver.Quit

    Set driver = Nothing
  End If

    Call Library.showNotice(Err.Number, Err.Description, True)
End Function



'*******************************************************************************************************
' * review.kakaku.com
' *
' * @author Bunpei.Koizumi<Koizumi.Bunpei@trans-cosmos.co.jp>
'*******************************************************************************************************
Function site_kakakuCom()
  Dim elements As WebElements
  Dim line As Long, endLine As Long
  Dim exitFlg As Boolean
  Dim getInnerHtmlString As Variant
  Dim myBy As New By
  Dim commentTitle As String
  Dim Red As Long, Green As Long, Blue As Long
  
  
'  On Error GoTo catchError
 
  exitFlg = False
  Call init.setting
  Call Chrome�N��
  
  '�^�C�g���擾
  sheetKoetol.Range("E3") = driver.title

reCrawl:
  driver.Wait 1000
  line = sheetKoetol.Cells(Rows.count, 3).End(xlUp).Row + 1

  '���R�~���擾
  Set elements = driver.FindElementById("mainLeft").FindElementsByClass("reviewBox")
  For i = 1 To elements.count
    DoEvents
    getInnerHtmlString = ""
    
    '�T�C�g------------------------------------------------------------------------------------------
    sheetKoetol.Range("C" & line) = line - 4
    sheetKoetol.Range("D" & line) = setVal("SiteName")
    
    '���e��----------------------------------------------------------------------------------------
    getInnerHtmlString = elements.Item(i).FindElementByClass("entryDate").Text
    getInnerHtmlString = Split(getInnerHtmlString, " ")
    getInnerHtmlString = Format(DateValue(getInnerHtmlString(0)), "YYYY/MM/DD")
    sheetKoetol.Range("E" & line) = getInnerHtmlString
    sheetKoetol.Range("E" & line).NumberFormatLocal = "yyyy/mm/dd"
    
    '���ʁE�N��------------------------------------------------------------------------------------------
    sheetKoetol.Range("F" & line) = "�s��"
    sheetKoetol.Range("G" & line) = "�s��"
    
    '�]��------------------------------------------------------------------------------------------
    getInnerHtmlString = elements.Item(i).FindElementByClass("total").Text
    sheetKoetol.Range("H" & line) = Replace(getInnerHtmlString, "�����x ", "")
    sheetKoetol.Range("H" & line).NumberFormatLocal = "0_ ;[��]-0 "

    '�R�����g--------------------------------------------------------------------------------------
    commentTitle = ""
    getInnerHtmlString = ""
    If elements.Item(i).IsElementPresent(myBy.Class("reviewTitle")) Then
      commentTitle = setVal("commentTitleStart") & elements.Item(i).FindElementByClass("reviewTitle").Text & setVal("commentTitleEnd")
      getInnerHtmlString = commentTitle & vbCrLf
      
    End If
    getInnerHtmlString = getInnerHtmlString & Library.delCellLinefeed(elements.Item(i).FindElementByClass("revEntryCont").Text)
    sheetKoetol.Range("I" & line) = getInnerHtmlString
    
    '�^�C�g�����{�[���h�ɂ���
    If commentTitle <> "" Then
      Call Library.getRGB(setVal("commentTitleColor"), Red, Green, Blue)
      With sheetKoetol.Range("I" & line).Characters(start:=1, Length:=Len(commentTitle)).Font
        .Bold = True
        .Color = RGB(Red, Green, Blue)
      End With
    End If
    
    '���y�[�W�J�ڔ���------------------------------------------------------------------------------
    If (Cells(Rows.count, 5).End(xlUp).Row - 4) >= setVal("DataCount") * setVal("DataRate") Then
      exitFlg = True
      Exit For
    Else
      exitFlg = False
    End If
    
    line = line + 1
  Next
  
  '���y�[�W�J��------------------------------------------------------------------------------------
  If exitFlg = True Then
    Call ProgressBar.showCount("Web�f�[�^�擾", 9, 10, "�I��")
    driver.Get "file:///" & openingHTML("Koetol") & "\pause.html"

  Else
    If driver.FindElementsByClass("arrowNext01").count > 0 Then
      driver.FindElementByClass("arrowNext01").Click
      driver.Wait 1000
      
      GoTo reCrawl
    End If
  End If

  Exit Function
'--------------------------------------------------------------------------------------------------
'�G���[�������̏���
'--------------------------------------------------------------------------------------------------
catchError:
   
  If Not driver Is Nothing Then
    driver.Close
    driver.Quit
    Set driver = Nothing
  End If

  Call Library.showNotice(Err.Number, Err.Description, True)

End Function


'*******************************************************************************************************
' * cosme-de.net
' *
' * @author Bunpei.Koizumi<Koizumi.Bunpei@trans-cosmos.co.jp>
'*******************************************************************************************************
Function site_cosmedeNet()
  Dim elements As WebElements, elements01 As WebElements
  Dim line As Long, endLine As Long
  Dim exitFlg As Boolean
  Dim getInnerHtmlString As Variant
  Dim myBy As New By
  Dim commentTitle As String
  Dim Red As Long, Green As Long, Blue As Long
  
'  On Error GoTo catchError

  exitFlg = False
  Call init.setting
  Call Chrome�N��
  
  '�^�C�g���擾
  sheetKoetol.Range("E3") = driver.title

reCrawl:
  driver.Wait 1000
  line = sheetKoetol.Cells(Rows.count, 3).End(xlUp).Row + 1

  '���R�~���擾
  Set elements = driver.FindElementByClass("review_list").FindElementsByClass("block")
  For i = 1 To elements.count
    DoEvents
    getInnerHtmlString = ""
    
    '�T�C�g------------------------------------------------------------------------------------------
    sheetKoetol.Range("C" & line) = line - 4
    sheetKoetol.Range("D" & line) = setVal("SiteName")
    
    '���e��----------------------------------------------------------------------------------------
    Set elements01 = elements.Item(i).FindElementsByClass("text_03")
    For j = 1 To elements01.count
      If elements01.Item(j).Text Like "���e���F*" Then
        getInnerHtmlString = elements01.Item(j).Text
        sheetKoetol.Range("E" & line) = Replace(getInnerHtmlString, "���e���F", "")
        sheetKoetol.Range("E" & line).NumberFormatLocal = "yyyy/mm/dd"
        Exit For
      End If
    Next
    
    '����------------------------------------------------------------------------------------------
    sheetKoetol.Range("F" & line) = "�s��"
    
    '�N��------------------------------------------------------------------------------------------
    getInnerHtmlString = elements.Item(i).FindElementByClass("text_02").Text
    Select Case getInnerHtmlString
      Case "�N��F20�Ζ���"
        sheetKoetol.Range("G" & line) = "10��"
      Case "�N��F20-25��"
        sheetKoetol.Range("G" & line) = "20��"
      Case "�N��F26-30��"
        sheetKoetol.Range("G" & line) = "20��"
      Case "�N��F31-35��"
        sheetKoetol.Range("G" & line) = "30��"
      Case "�N��F36-40��"
        sheetKoetol.Range("G" & line) = "30��"
      Case "�N��F41-50��"
        sheetKoetol.Range("G" & line) = "40��"
      Case "�N��F51-60��"
        sheetKoetol.Range("G" & line) = "50��"
      Case "�N��F61�Έȏ�"
        sheetKoetol.Range("G" & line) = "60��"
      Case "�N��F���o�^"
        sheetKoetol.Range("G" & line) = "�s��"
      Case Else
        sheetKoetol.Range("G" & line) = "�s��"
    End Select
        
    '�]��------------------------------------------------------------------------------------------
    getInnerHtmlString = elements.Item(i).FindElementByClass("star_area").FindElementByClass("star").Attribute("class")

    Select Case getInnerHtmlString
      Case "star star_00h"
        sheetKoetol.Range("H" & line) = "0.5"
      Case "star star_01"
        sheetKoetol.Range("H" & line) = "1.0"
      Case "star star_01h"
        sheetKoetol.Range("H" & line) = "1.5"
      Case "star star_02"
        sheetKoetol.Range("H" & line) = "2.0"
      Case "star star_02h"
        sheetKoetol.Range("H" & line) = "2.5"
      Case "star star_03"
        sheetKoetol.Range("H" & line) = "3.0"
      Case "star star_03h"
        sheetKoetol.Range("H" & line) = "3.5"
      Case "star star_04"
        sheetKoetol.Range("H" & line) = "4.0"
      Case "star star_04h"
        sheetKoetol.Range("H" & line) = "4.5"
      Case "star star_05"
        sheetKoetol.Range("H" & line) = "5.0"
      Case Else
        sheetKoetol.Range("H" & line) = ""
    End Select
    sheetKoetol.Range("H" & line).NumberFormatLocal = "0.0_ "

    '�R�����g--------------------------------------------------------------------------------------
    commentTitle = ""
    getInnerHtmlString = ""
    If elements.Item(i).IsElementPresent(myBy.Class("ttl")) Then
      commentTitle = setVal("commentTitleStart") & elements.Item(i).FindElementByClass("ttl").Text & setVal("commentTitleEnd")
      getInnerHtmlString = commentTitle & vbCrLf
      
    End If
    getInnerHtmlString = getInnerHtmlString & Library.delCellLinefeed(elements.Item(i).FindElementByClass("js-review_comment").Text)
    
    '�J�e�S���[------------------------------------------------------------------------------------
    Set elements01 = elements.Item(i).FindElementsByClass("text_03")
    For j = 1 To elements01.count
      If elements01.Item(j).Text Like "�����F*" Then
        skinType = elements01.Item(j).Text
        
      ElseIf elements01.Item(j).Text Like "���ʁF*" Then
        effect = elements01.Item(j).Text
      End If
    Next j
    sheetKoetol.Range("I" & line) = getInnerHtmlString & vbCrLf & vbCrLf & skinType & vbCrLf & effect
    
    
    '�^�C�g�����{�[���h�ɂ���
    If commentTitle <> "" Then
      Call Library.getRGB(setVal("commentTitleColor"), Red, Green, Blue)
      With sheetKoetol.Range("I" & line).Characters(start:=1, Length:=Len(commentTitle)).Font
        .Bold = True
        .Color = RGB(Red, Green, Blue)
      End With
    End If
     
    '���y�[�W�J�ڔ���------------------------------------------------------------------------------
    If (Cells(Rows.count, 5).End(xlUp).Row - 4) >= setVal("DataCount") * setVal("DataRate") Then
      exitFlg = True
      Exit For
    Else
      exitFlg = False
    End If
    line = line + 1
  Next
  
  '���y�[�W�J��------------------------------------------------------------------------------------
  If exitFlg = True Then
    Call ProgressBar.showCount("Web�f�[�^�擾", 9, 10, "�I��")
    driver.Get "file:///" & openingHTML("Koetol") & "\pause.html"

  Else
    If driver.FindElementsByClass("next").count > 0 Then
      driver.FindElementByClass("next").Click
      driver.Wait 1000
      
      GoTo reCrawl
    End If
  End If

  Exit Function
'--------------------------------------------------------------------------------------------------
'�G���[�������̏���
'--------------------------------------------------------------------------------------------------
catchError:
   
  If Not driver Is Nothing Then
    driver.Close
    driver.Quit
    Set driver = Nothing
  End If

  Call Library.showNotice(Err.Number, Err.Description, True)

End Function



