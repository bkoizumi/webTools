Attribute VB_Name = "�����N���o"
Public Const startLine As Long = 2


Dim activLine As Long
Dim pageCount As Long
Dim baseURL As String
Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
Dim targetFileDir As String

'*******************************************************************************************************
' * �擾�J�n
' *
' * @author Bunpei.Koizumi<Koizumi.Bunpei@trans-cosmos.co.jp>
'*******************************************************************************************************
Function �擾�J�n()
  Dim elements As WebElements
  
  Dim c As Range
  Dim URLColumn As Integer
  Dim IDPW As String
  
'  On Error GoTo catchError
  
  Call init.setting
  Call ProgressBar.showStart
  Call WebToolLib.Chrome�N��
  
  For colLine = 3 To Cells(1, Columns.count).End(xlToLeft).Column

    endLine = sheetLinkExtract.Cells(Rows.count, colLine).End(xlUp).Row
    
    For line = startLine To endLine
      sheetTmp.Select
      Call Library.delSheetData
      sheetTmp.Columns("A:B").ColumnWidth = 100
      sheetLinkExtract.Select
      baseURL = sheetLinkExtract.Cells(1, colLine)
      
      
      'CSV�t�@�C���ۑ��ꏊ����
      targetFileDir = setVal("workPath") & "\�����N���o\" & sheetLinkExtract.Cells(1, colLine)
      Call Library.makeDir(targetFileDir)
      
      targetURL = sheetLinkExtract.Cells(line, colLine)
      
      If targetURL = "" Then
        GoTo nextForLabel
      End If
      
      driver.Get targetURL
      Call a�^�O����
      Call CSV�쐬
    Next
nextForLabel:
  Next

  
  Call WebToolLib.Chrome�I��
  
  Call ProgressBar.showEnd
  Call Library.endScript
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
' * a�^�O����
' *
' * @author Bunpei.Koizumi<Koizumi.Bunpei@trans-cosmos.co.jp>
'*******************************************************************************************************
Function a�^�O����()
  Dim line As Long, endLine As Long
  Dim elements As WebElements, targetElements As WebElements
  
  Dim exitFlg As Boolean
  Dim getInnerHtmlString As Variant
  Dim myBy As New By
  Dim URL As String
  Dim i As Long
  
'  On Error GoTo catchError
  
  Call Library.showDebugForm("�����͈́F" & sheetLinkExtract.Cells(4, colLine))
  
  Set elements = driver.FindElementsByTag("a")
'  Set targetElements = driver.FindElementsByClass(sheetLinkExtract.Cells(4, colLine))
'  Set elements = targetElements(1).FindElementsByTag("a")
  
  line = 1
  sheetTmp.Select
  
  For i = 1 To elements.count
    URL = elements.Item(i).Attribute("href")
    attText = elements.Item(i).Text
    attText = Replace(attText, vbCrLf, "")
    attText = Replace(attText, vbLf, "")
    
    Call ProgressBar.showCount("�����N�`�F�b�N", i, elements.count, "a�^�O�����F" & URL)
    
    
    '���OURL���ǂ���
    For setLine = 3 To sheetSetting.Cells(Rows.count, 7).End(xlUp).Row
      If URL <> "" And URL Like sheetSetting.Range("G" & setLine) & "*" Then
        URL = ""
        Exit For
      End If
    Next
    
    
    '���O�������ǂ���
    Call Library.showDebugForm("attText�F" & attText)
    For setLine = 3 To sheetSetting.Cells(Rows.count, 9).End(xlUp).Row
      If attText <> "" And attText Like sheetSetting.Range("I" & setLine) & "*" Then
        attText = ""
        Exit For
      End If
    Next
    
    If URL <> "" And attText <> "" Then
      URL = Replace(URL, baseURL, "")
      sheetTmp.Range("A" & line) = URL
      sheetTmp.Range("B" & line) = attText
      
      line = line + 1
    End If
    
    
  Next
  
  Exit Function
  
'�G���[������====================================
catchError:
   
  If Not driver Is Nothing Then
    WebToolLib.Chrome�I��
  End If

  Call Library.showNotice(Err.Number, Err.Description, True)
End Function


'*******************************************************************************************************
' * CSV�쐬
' *
' * @author Bunpei.Koizumi<Koizumi.Bunpei@trans-cosmos.co.jp>
'*******************************************************************************************************
Function CSV�쐬()
  Dim csvLine As Long, csvEndLine As Long


'  On Error GoTo catchError
  
  targetFilePath = targetFileDir & "\" & sheetLinkExtract.Cells(line, 2)
  
  sheetTmp.Select
  csvEndLine = sheetTmp.Cells(Rows.count, 1).End(xlUp).Row
  
  Call Library.showDebugForm("targetFilePath�F" & targetFilePath)
  Open targetFilePath For Output As #1
  For csvLine = 1 To csvEndLine
    Print #1, Replace(sheetTmp.Range("B" & csvLine), vbLf, "") & "," & sheetTmp.Range("A" & csvLine)
  Next
  Close #1

  Exit Function
'�G���[������====================================
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function
































