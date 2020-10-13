Attribute VB_Name = "WebCapture"

Dim driver As New Selenium.WebDriver
Dim targetURL As String
Dim line As Long
Dim endLine As Long

'*******************************************************************************************************
' * Chrome�N��
' *
' * @author Bunpei.Koizumi<Koizumi.Bunpei@trans-cosmos.co.jp>
'*******************************************************************************************************
Function Chrome�N��()
  
  Call init.setting
  
  'chrome �g���@�\�ǂݍ���
  Call ProgressBar.showCount("WebCapture", 0, 10, "Chrome�N��")
  With driver
    .AddArgument ("--lang=ja")
    .AddArgument ("--user-data-dir=" & BrowserProfiles("default"))
    .AddArgument ("--window-size=1200,600")
    .AddArgument ("--hide-scrollbars")
    .AddArgument ("--disable-gpu")
    
    If setVal("debugMode") = "develop" Then
      Call Library.showDebugForm("WebCapture", "�V�[�N���b�g���[�h")
      
      .AddArgument ("--incognito")                   '�V�[�N���b�g���[�h
    Else
      Call Library.showDebugForm("WebCapture", "headless���[�h")
      .AddArgument ("--headless")
    End If
    
    If setVal("InstNetwork") = "TCI" Then
      .AddArgument ("--proxy-server=" & setVal("ProxyURL") & ":" & setVal("ProxyPort"))
    End If
    
    .start "chrome"
    .Wait 1000
  End With
  
  Call ProgressBar.showCount("WebCapture", 10, 10, "Chrome�N��")
  Call Library.showDebugForm("WebCapture", "Chrome�N������")
  
End Function


'*******************************************************************************************************
' * �擾�J�n
' *
' * @author Bunpei.Koizumi<Koizumi.Bunpei@trans-cosmos.co.jp>
'*******************************************************************************************************
Function �擾�J�n()
  Dim elements As WebElements

  Dim shellProcID As Long
  Dim targetFilePath As String, targetFileName As String
  
  On Error GoTo catchError
  ThisWorkbook.Save
  Call init.setting
  Call Library.startScript
  Call ProgressBar.showStart
  
  Set targetBook = Workbooks.Open(FileName:=AppWebCapturePath & "\�V�KBook.xlsm", ReadOnly:=True)
  Windows(targetBook.Name).WindowState = xlMinimized
  ThisWorkbook.Activate
  
  '���O�����ĕۑ�
  targetFilePath = Library.getRegistry("WebCapturePath")
  If targetFilePath <> "" Then
    targetFileName = Dir(targetFilePath)
    targetFilePath = Replace(targetFilePath, "\" & targetFileName, "")
  End If
  targetFilePath = Library.getSaveFilePath(targetFilePath, targetFileName, 3)
  
  targetBook.SaveAs FileName:=targetFilePath
  
  'chromeDriver �̍X�V�m�F-------------------------------------------------------------------------
  shellProcID = Shell(binPath & "\SeleniumBasic\updateChromeDriver.bat", vbNormalFocus)
  Call Library.chkShellEnd(shellProcID)
  
  Call Chrome�N��
  
  endLine = sheetWebCaptureList.Cells(Rows.count, 1).End(xlUp).Row
  
  For line = 15 To endLine
    targetURL = sheetWebCaptureList.Range("A" & line)
    If targetURL <> "" Then
      Call ProgressBar.showCount("WebCapture", line - 14, endLine - 14, driver.title)
      
reLoad:
      driver.Timeouts.PageLoad = 60000
      driver.Wait 2000
      driver.Get targetURL
      driver.Wait 1000
      
      Call �y�[�W���擾
      Call ��Book�ɕۑ�
      
      If sheetWebCaptureList.Range("F" & line) <> "" Then
        ThisWorkbook.Activate
        sheetWebCapture.Select
        sheetWebCapture.Shapes.Range(Array(sheetWebCapture.Range("O1"))).Delete
        
        Call �y�[�W���擾(True)
        Call ��Book�ɕۑ�
      
      End If

      
      ThisWorkbook.Activate
      sheetWebCapture.Shapes.Range(Array(sheetWebCapture.Range("O1"))).Delete
      ThisWorkbook.Save

    End If
  Next
  Call ProgressBar.showCount("WebCapture", 0, 100, "index�쐬")
  Call index�쐬

  driver.Close
  driver.Quit
  Set driver = Nothing

  '�o�̓t�@�C���̏㏑��
  targetBook.SaveAs
  targetBook.Close
  Set targetBook = Nothing


  sheetWebCapture.Range("B1") = ""
  sheetWebCapture.Range("B2") = ""
  sheetWebCapture.Range("B3") = ""
  sheetWebCapture.Range("L1") = ""
  sheetWebCapture.Range("O1") = ""
  
  Call Library.setRegistry("WebCapturePath", targetFilePath)
  Call Shell("Explorer.exe /select, " & targetFilePath, vbNormalFocus)
  

  Call ProgressBar.showEnd
  Call Library.endScript
  
  
  Exit Function
'--------------------------------------------------------------------------------------------------
'�G���[�������̏���
'--------------------------------------------------------------------------------------------------
catchError:
   
  If Err.Number = 13 And Err.Description Like "*ERR_CONNECTION_TIMED_OUT*" Then
    Err.Clear
    driver.Wait 7000
    GoTo reLoad
  Else
    If Not driver Is Nothing Then
      driver.Close
      driver.Quit
      Set driver = Nothing
    End If
    Call Library.showNotice(Err.Number, Err.Description, True)
  End If
End Function


'*******************************************************************************************************
' * �y�[�W���擾
' *
' * @author Bunpei.Koizumi<Koizumi.Bunpei@trans-cosmos.co.jp>
'*******************************************************************************************************
Function �y�[�W���擾(Optional actionFlg As Boolean)
  Dim elements As WebElements
  Dim getInnerHtmlString As Variant
  Dim myBy As New By
  Dim ks As New Keys
  
  
  Dim page_width As Long, page_height As Long
  Dim userID As String, userPW As String
  
  Dim userIDTagID As String, userPWTagID As String
  Dim userIDTagName As String, userPWTagName As String
  
  Dim btn01TagId As String, btn01TagName As String
  Dim btn02TagId As String, btn02TagName As String
  
  
  ThisWorkbook.Save

  Call ProgressBar.showCount("WebCapture", line - 14, endLine - 14, driver.title)

  '�y�[�W���i�[
  sheetWebCapture.Range("B1") = ""
  sheetWebCapture.Range("O1") = "WC" & Format(line - 14, "00")
  sheetWebCapture.Range("L1") = Format(Now(), "yyyy/mm/dd hh:nn:ss")


  If actionFlg = True Then
    If sheetWebCaptureList.Range("F" & line) = "����1" Then
      sheetWebCapture.Range("B1") = "WC" & Format(line - 14, "00") & "_������"

      searchWord = sheetWebCaptureList.Range("B6")
      searchWordTagName = sheetWebCaptureList.Range("C6")
      searchWordTagID = sheetWebCaptureList.Range("D6")
      searchWordTagClass = sheetWebCaptureList.Range("E6")
      
      searchBtn = sheetWebCaptureList.Range("B7")
      searchBtnTagName = sheetWebCaptureList.Range("C7")
      searchBtnTagID = sheetWebCaptureList.Range("D7")
      searchBtnTagClass = sheetWebCaptureList.Range("E7")
      
      '�����L�[���[�h���̓^�O����
      If driver.IsElementPresent(myBy.Name(searchWordTagName)) And searchWordTagName <> "" Then
        driver.FindElementByName(searchWordTagName).SendKeys (searchWord)
        
      ElseIf driver.IsElementPresent(myBy.Class(searchWordTagClass)) And searchWordTagClass <> "" Then
        driver.FindElementByClass(searchWordTagClass).SendKeys (searchWord)
      
      ElseIf driver.IsElementPresent(myBy.id(searchWordTagID)) And searchWordTagID <> "" Then
        driver.FindElementById(searchWordTagID).SendKeys (searchWord)
      End If
      
      '�����{�^������
      If driver.IsElementPresent(myBy.Name(searchBtnTagName)) And searchBtnTagName <> "" Then
        driver.FindElementByName(searchBtnTagName).Click
      ElseIf driver.IsElementPresent(myBy.Class(searchBtnTagClass)) And searchBtnTagClass <> "" Then
        driver.FindElementByClass(searchBtnTagClass).Click
      
      ElseIf driver.IsElementPresent(myBy.id(searchBtnTagID)) And searchBtnTagID <> "" Then
        driver.FindElementById(searchBtnTagID).Click
      Else
        driver.SendKeys (ks.Enter)
      End If
      driver.Wait 1000
        
    ElseIf sheetWebCaptureList.Range("F" & line) = "����2" Then
      sheetWebCapture.Range("B1") = "WC" & Format(line - 14, "00") & "_������"

      searchWord = sheetWebCaptureList.Range("B8")
      searchWordTagName = sheetWebCaptureList.Range("C8")
      searchWordTagID = sheetWebCaptureList.Range("D8")
      searchWordTagClass = sheetWebCaptureList.Range("E8")
      
      searchBtn = sheetWebCaptureList.Range("B9")
      searchBtnTagName = sheetWebCaptureList.Range("C9")
      searchBtnTagID = sheetWebCaptureList.Range("D9")
      searchBtnTagClass = sheetWebCaptureList.Range("E9")
      
      '�����L�[���[�h���̓^�O����
      If driver.IsElementPresent(myBy.Name(searchWordTagName)) And searchWordTagName <> "" Then
        driver.FindElementByName(searchWordTagName).SendKeys (searchWord)
        
      ElseIf driver.IsElementPresent(myBy.Class(searchWordTagClass)) And searchWordTagClass <> "" Then
        driver.FindElementByClass(searchWordTagClass).SendKeys (searchWord)
      
      ElseIf driver.IsElementPresent(myBy.id(searchWordTagID)) And searchWordTagID <> "" Then
        driver.FindElementById(searchWordTagID).SendKeys (searchWord)
      End If
      
      '�����{�^������
      If driver.IsElementPresent(myBy.Name(searchBtnTagName)) And searchBtnTagName <> "" Then
        driver.FindElementByName(searchBtnTagName).Click
      ElseIf driver.IsElementPresent(myBy.Class(searchBtnTagClass)) And searchBtnTagClass <> "" Then
        driver.FindElementByClass(searchBtnTagClass).Click
      
      ElseIf driver.IsElementPresent(myBy.id(searchBtnTagID)) And searchBtnTagID <> "" Then
        driver.FindElementById(searchBtnTagID).Click
      Else
        driver.SendKeys (ks.Enter)
      End If
      driver.Wait 1000
    
    ElseIf sheetWebCaptureList.Range("F" & line) = "����3" Then
      sheetWebCapture.Range("B1") = "WC" & Format(line - 14, "00") & "_������"

      searchWord = sheetWebCaptureList.Range("B10")
      searchWordTagName = sheetWebCaptureList.Range("C10")
      searchWordTagID = sheetWebCaptureList.Range("D10")
      searchWordTagClass = sheetWebCaptureList.Range("E10")
      
      searchBtn = sheetWebCaptureList.Range("B11")
      searchBtnTagName = sheetWebCaptureList.Range("C11")
      searchBtnTagID = sheetWebCaptureList.Range("D11")
      searchBtnTagClass = sheetWebCaptureList.Range("E11")
      
      '�����L�[���[�h���̓^�O����
      If driver.IsElementPresent(myBy.Name(searchWordTagName)) And searchWordTagName <> "" Then
        driver.FindElementByName(searchWordTagName).SendKeys (searchWord)
        
      ElseIf driver.IsElementPresent(myBy.Class(searchWordTagClass)) And searchWordTagClass <> "" Then
        driver.FindElementByClass(searchWordTagClass).SendKeys (searchWord)
      
      ElseIf driver.IsElementPresent(myBy.id(searchWordTagID)) And searchWordTagID <> "" Then
        driver.FindElementById(searchWordTagID).SendKeys (searchWord)
      End If
      
      '�����{�^������
      If driver.IsElementPresent(myBy.Name(searchBtnTagName)) And searchBtnTagName <> "" Then
        driver.FindElementByName(searchBtnTagName).Click
      ElseIf searchBtnTagClass <> "" Then
        If driver.IsElementPresent(myBy.Class(searchBtnTagClass)) Then
          driver.FindElementByClass(searchBtnTagClass).Click
        End If
      ElseIf searchBtnTagID <> "" Then
        If driver.IsElementPresent(myBy.id(searchBtnTagID)) Then
          driver.FindElementById(searchBtnTagID).Click
        End If
      Else
        driver.SendKeys (ks.Enter)
      End If
      driver.Wait 1000


    ElseIf sheetWebCaptureList.Range("F" & line) = "��i�K���O�C��" Then
      sheetWebCapture.Range("B1") = "WC" & Format(line - 14, "00") & "_�F�،�"
      
      userID = sheetWebCaptureList.Range("B2")
      userIDTagName = sheetWebCaptureList.Range("C2")
      
      userPW = sheetWebCaptureList.Range("B3")
      userPWTagName = sheetWebCaptureList.Range("C3")
      
      btn01TagName = sheetWebCaptureList.Range("C4")
      btn02TagName = sheetWebCaptureList.Range("C5")
      driver.FindElementByName(userIDTagName).SendKeys (userID)
      driver.FindElementByName(btn01TagName).Click
      driver.Wait 1000
      
      driver.FindElementByName(userPWTagName).SendKeys (userPW)
      driver.FindElementByName(btn02TagName).Click
    ElseIf sheetWebCaptureList.Range("F" & line) = "�ʏ탍�O�C��" Then
      sheetWebCapture.Range("B1") = "WC" & Format(line - 14, "00") & "_�F�،�"
      
      userID = sheetWebCaptureList.Range("B2")
      userIDTagName = sheetWebCaptureList.Range("C2")
      
      userPW = sheetWebCaptureList.Range("B3")
      userPWTagName = sheetWebCaptureList.Range("C3")
      
      btn01TagName = sheetWebCaptureList.Range("C4")
      btn02TagName = sheetWebCaptureList.Range("C5")
      driver.FindElementByName(userIDTagName).SendKeys (userID)
      driver.FindElementByName(userPWTagName).SendKeys (userPW)
      driver.FindElementByName(btn02TagName).Click
    End If

    On Error GoTo 0
  End If
  
  '�y�[�W���i�[
  sheetWebCapture.Range("B2") = driver.title
  sheetWebCapture.Range("B3") = driver.URL
  
  
  '��ʃV���b�g����
  page_width = driver.ExecuteScript("return document.body.scrollWidth")
  page_height = driver.ExecuteScript("return document.body.scrollHeight")
  driver.Window.SetSize page_width, page_height

  driver.TakeScreenshot.Copy
  ThisWorkbook.Activate
  sheetWebCapture.Select
  sheetWebCapture.Range("A5").Select
  Call Library.waitTime(2000)
  
  'ActiveSheet.Paste
  ActiveSheet.PasteSpecial Format:="�r�b�g�}�b�v", link:=False, DisplayAsIcon:=False
  
  Selection.Name = sheetWebCapture.Range("O1")

  Call ProgressBar.showCount("WebCapture", line - 14, endLine - 14, driver.title)
  
  Selection.ShapeRange.Width = 480
  Selection.ShapeRange.IncrementLeft 20
  Selection.ShapeRange.IncrementTop 10
    
  Selection.Placement = xlFreeFloating
  With Selection.ShapeRange.line
    .Visible = msoTrue
    .ForeColor.ObjectThemeColor = msoThemeColorBackground1
    .ForeColor.TintAndShade = 0
    .ForeColor.Brightness = -0.5
    .Transparency = 0
  End With
  Application.CutCopyMode = False
  Application.Goto Reference:=Range("A1"), Scroll:=True
  
  Call ProgressBar.showCount("WebCapture", line - 14, endLine - 14, driver.title)

End Function


'**************************************************************************************************
' * ��Book�ɕۑ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function ��Book�ɕۑ�()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long

  Call ProgressBar.showCount("WebCapture", line - 14, endLine - 14, driver.title)
  sheetWebCapture.Copy After:=targetBook.Worksheets(targetBook.Worksheets.count)
  
  targetBook.ActiveSheet.Name = sheetWebCapture.Range("O1")
  Call ProgressBar.showCount("WebCapture", line - 14, endLine - 14, driver.title)


End Function


'**************************************************************************************************
' * index�쐬
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function index�쐬()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim tempSheet As Object
  Dim result As Boolean

  targetBook.Activate
  Windows(targetBook.Name).WindowState = xlNormal
  Worksheets("Index").Select


  line = 2
  For Each tempSheet In Sheets
    DoEvents
    If tempSheet.Name <> "Index" Then
      Worksheets("Index").Range("B" & line) = Worksheets(tempSheet.Name).Range("B2")
      Worksheets("Index").Range("C" & line) = Worksheets(tempSheet.Name).Range("B3")
      Worksheets("Index").Range("D" & line) = Worksheets(tempSheet.Name).Range("L1")
      Worksheets("Index").Range("A" & line) = "=HYPERLINK(""#'" & tempSheet.Name & "'!A1"",""" & ActiveWorkbook.Sheets(tempSheet.Name).Range("O1") & """)"
      line = line + 1
    End If
  Next
  Columns("D:D").Select
  Selection.NumberFormatLocal = "yyyy/mm/dd"
  
  Worksheets("Index").Select
  Application.Goto Reference:=Range("A1"), Scroll:=True

End Function






















