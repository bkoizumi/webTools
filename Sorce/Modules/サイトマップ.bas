Attribute VB_Name = "�T�C�g�}�b�v"

Dim driver As New Selenium.WebDriver
Dim targetURL As String
Dim activLine As Long
Dim pageCount As Long

'*******************************************************************************************************
' * Chrome�N��
' *
' * @author Bunpei.Koizumi<Koizumi.Bunpei@trans-cosmos.co.jp>
'*******************************************************************************************************
Function Chrome�N��()
  
  'chrome �g���@�\�ǂݍ���
  With driver
    .AddArgument ("--lang=ja")
    .AddArgument ("--user-data-dir=" & BrowserProfiles("noScript"))
    .AddArgument ("--window-size=1200,600")
    .AddArgument ("--disable-gpu")
    .AddArgument ("--incognito")                   '�V�[�N���b�g���[�h
    
    If setVal("InstNetwork") = "TCI" Then
      .AddArgument ("--proxy-server=" & setVal("ProxyURL") & ":" & setVal("ProxyPort"))
    End If
    .AddArgument ("--app=" & targetURL)
    .start "chrome"
    .Wait 1000
    '.Get targetURL
  End With

End Function


'*******************************************************************************************************
' * �擾�J�n
' *
' * @author Bunpei.Koizumi<Koizumi.Bunpei@trans-cosmos.co.jp>
'*******************************************************************************************************
Function �擾�J�n()
  Dim elements As WebElements
  Dim line As Long, endLine As Long
  Dim c As Range
  
'  On Error GoTo catchError
  
  baseURL = "https://news.yahoo.co.jp/"
  baseURL = "https://www.trans-cosmos.co.jp/"
  exitFlg = False
  Call init.setting
  Call ProgressBar.showStart
  
  targetURL = baseURL
  pageCount = 1
  
  sheetSitemapTmp.Select
  Call Library.delSheetData
  
  sheetSitemap.Select
  Call Library.delSheetData(3)
  
  Call Chrome�N��
  
  Call �y�[�W���擾
  Call a�^�O����

  Do While sheetSitemapTmp.Range("A1") <> ""
    pageCount = pageCount + 1
    targetURL = sheetSitemapTmp.Range("A1")
    sheetSitemapTmp.Rows("1:1").Delete Shift:=xlUp
    
    '�����ς�URL���m�F
    Set c = sheetSitemap.UsedRange.Columns(4).Find(What:=targetURL, LookIn:=xlValues, LookAt:=xlPart)
    If c Is Nothing Then
      Call ProgressBar.showCount("�T�C�g�}�b�v����", 0, 10, "�y�[�W�J�ځF" & targetURL)
      Call Library.waitTime(2000)
      DoEvents
      driver.Get targetURL
      driver.Wait 1000
      Call ProgressBar.showCount("�T�C�g�}�b�v����", 10, 10, "�y�[�W�J�ځF" & targetURL)
      
      Call �y�[�W���擾
      Call a�^�O����
'      sheetSitemap.Select

    End If
'    Application.Goto Reference:=sheetSitemapTmp.Range("A1"), Scroll:=True
    ThisWorkbook.Save
  Loop
  
  driver.Close
  driver.Quit
  Set driver = Nothing
  
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
' * �y�[�W���擾
' *
' * @author Bunpei.Koizumi<Koizumi.Bunpei@trans-cosmos.co.jp>
'*******************************************************************************************************
Function �y�[�W���擾()
  
  Dim line As Long, endLine As Long
  Dim descElements As WebElements, keyWdElements As WebElements
  Dim myBy As New By
  Dim pageSourceTag As String
  
  On Error Resume Next
  ThisWorkbook.Save

  line = sheetSitemap.Cells(Rows.count, 1).End(xlUp).Row + 1
  
  Call ProgressBar.showCount("�T�C�g�}�b�v����", 0, 10, "�y�[�W���擾�F" & driver.title)

  sheetSitemap.Range("A" & line).FormulaR1C1 = "=ROW()-2"
  sheetSitemap.Range("C" & line) = driver.title
  sheetSitemap.Range("D" & line) = driver.URL


  If driver.IsElementPresent(myBy.XPath("//meta[@name=""description""]")) Then
    Set descElements = driver.FindElementsByXPath("//meta[@name=""description""]")
    sheetSitemap.Range("E" & line) = descElements.Item(1).Attribute("content")
  End If
    
  If driver.IsElementPresent(myBy.XPath("//meta[@name=""keywords""]")) Then
    Set keyWdElements = driver.FindElementsByXPath("//meta[@name=""keywords""]")
    sheetSitemap.Range("F" & line) = keyWdElements.Item(1).Attribute("content")
  End If
  
  'GMT�^�O�����邩�ǂ�������
  pageSourceTag = driver.PageSource
  If pageSourceTag Like "*googletagmanager*" Then
    sheetSitemap.Range("G" & line) = "��"
  End If
  If pageSourceTag Like "*yjtag.jp*" Then
    sheetSitemap.Range("H" & line) = "��"
  End If
  
  
  pageSourceTag = ""
  Set descElements = Nothing
  Set keyWdElements = Nothing
End Function


'*******************************************************************************************************
' * a�^�O����
' *
' * @author Bunpei.Koizumi<Koizumi.Bunpei@trans-cosmos.co.jp>
'*******************************************************************************************************
Function a�^�O����()
  Dim line As Long, endLine As Long
  Dim elements As WebElements
  
  Dim exitFlg As Boolean
  Dim getInnerHtmlString As Variant
  Dim myBy As New By
  Dim URL As String
  Dim i As Long
  
'  On Error GoTo catchError
  ThisWorkbook.Save
  
  Set elements = driver.FindElementsByTag("a")
  
'  Call Library.showDebugForm("a�^�O����", targetURL)
'  Call Library.showDebugForm("a�^�O����", "�@" & elements.count & "���m�F")

  For i = 1 To elements.count
    URL = elements.Item(i).Attribute("href")
    Call ProgressBar.showCount("�T�C�g�}�b�v����", i, elements.count, "a�^�O�����F" & URL)
    
    If URL Like targetURL & "*" And Not (URL Like targetURL & "*.pdf") Then
      If Not (URL = targetURL Or URL Like targetURL & "[#]*") Then
        endLine = sheetSitemapTmp.Cells(Rows.count, 1).End(xlUp).Row + 1
        sheetSitemapTmp.Range("A" & endLine) = URL
      End If
    End If
  Next
  
  endLine = Cells(Rows.count, 1).End(xlUp).Row
  If endLine = 1 Then
    Exit Function
  End If
  
  '���ёւ��Əd���폜
  sheetSitemapTmp.Sort.SortFields.Clear
  
  sheetSitemapTmp.Sort.SortFields.Add Key:=Range("A1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
  With sheetSitemapTmp.Sort
      .SetRange Range("A1:A" & endLine)
      .Header = xlGuess
      .MatchCase = False
      .Orientation = xlTopToBottom
      .SortMethod = xlPinYin
      .Apply
  End With
  
  
  sheetSitemapTmp.Select
'  Application.Goto Reference:=sheetSitemapTmp.Range("A" & endLine - 20), Scroll:=True
'  sheetSitemapTmp.Range("A2:A" & endLine).Select

  sheetSitemapTmp.Range("A1:A" & endLine).RemoveDuplicates Columns:=1, Header:=xlNo

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





























