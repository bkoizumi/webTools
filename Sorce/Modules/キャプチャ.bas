Attribute VB_Name = "�L���v�`��"
Dim line As Long
Dim endLine As Long




'*******************************************************************************************************
' * �擾�J�n
' *
' * @author Bunpei.Koizumi<Koizumi.Bunpei@trans-cosmos.co.jp>
'*******************************************************************************************************
Function �擾�J�n()
  Dim elements As WebElements
  Dim shellProcID As Long
  
  
'  On Error GoTo catchError

  
  Set targetBook = Workbooks.Open(FileName:=AppWebCapturePath & "\�V�KBook.xlsm", ReadOnly:=True)
  Windows(targetBook.Name).WindowState = xlMinimized
  ThisWorkbook.Activate
  
  targetFilePath = setVal("workPath") & "\�L���v�`��\" & Format(Date, "yyyymmdd") & ".xlsm"
  
  targetBook.SaveAs FileName:=targetFilePath
  
  'chromeDriver �̍X�V�m�F-------------------------------------------------------------------------
  shellProcID = Shell(setVal("binPath") & "\SeleniumBasic\updateChromeDriver.bat", vbNormalFocus)
  Call Library.chkShellEnd(shellProcID)
  
  Call WebToolLib.Chrome�N��
  
  endLine = sheetWebCaptureList.Cells(Rows.count, 1).End(xlUp).Row
  
  For line = 2 To endLine
    targetURL = sheetWebCaptureList.Range("A" & line)
    If targetURL <> "" Then
      Call ProgressBar.showCount("WebCapture", line - 1, endLine - 1, driver.title)
      
reLoad:
      
      driver.Window.SetSize 1200, 600
      driver.Wait 2000
      driver.Get targetURL
      driver.Wait 1000
      
      Call �y�[�W���擾
      Call ��Book�ɕۑ�
      
      If sheetWebCaptureList.Range("C" & line) <> "" Then
        ThisWorkbook.Activate
        sheetWebCapture.Select
        sheetWebCapture.Shapes.Range(Array(sheetWebCapture.Range("B1").Text)).Delete
        
        Call �y�[�W���擾(True)
        Call ��Book�ɕۑ�
      
      End If

      
      ThisWorkbook.Activate
      sheetWebCapture.Shapes.Range(Array(sheetWebCapture.Range("B1").Text)).Delete

    End If
  Next
  Call ProgressBar.showCount("WebCapture", 0, 100, "index�쐬")
  Call index�쐬


  Call WebToolLib.Chrome�I��
  
  
  '�o�̓t�@�C���̏㏑��
  targetBook.Save
  targetBook.Close
  Set targetBook = Nothing


  sheetWebCapture.Range("B1") = ""
  sheetWebCapture.Range("B2") = ""
  sheetWebCapture.Range("B3") = ""
  sheetWebCapture.Range("L1") = ""
  
 
  Exit Function
'--------------------------------------------------------------------------------------------------
'�G���[�������̏���
'--------------------------------------------------------------------------------------------------
catchError:
  On Error Resume Next
  
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
  Dim setSheetName As String
  
  
  Dim page_width As Long, page_height As Long
  Dim userID As String, userPW As String
  
  Dim htmlType As String, htmlVal As String
  
  
  Call ProgressBar.showCount("WebCapture", line - 1, endLine - 1, driver.title)
'  openingDriver.ExecuteScript ("$('.appName').text('" & driver.title & " �擾�J�n') ")

  '�y�[�W���i�[
  setSheetName = sheetWebCaptureList.Range("B" & line)
  If setSheetName = "" Then
    setSheetName = setVal("sheetName") & Format(line - 1, "000")
  End If
  
  sheetWebCapture.Select
  sheetWebCapture.Range("B1") = setSheetName
  sheetWebCapture.Range("L1") = Format(Now(), "yyyy/mm/dd hh:nn:ss")


  If actionFlg = True Then
    If sheetWebCaptureList.Range("C" & line) = "����1" Then
      sheetWebCapture.Range("B1") = setSheetName & "_������"

      searchWord = sheetSetting.Range("search1Val")
      searchWordTagName = sheetSetting.Range("search1TagName")
      searchWordTagID = sheetSetting.Range("search1TagID")
      searchWordTagClass = sheetSetting.Range("search1TagClass")
      
      searchBtn = sheetSetting.Range("search1BtnVal")
      searchBtnTagName = sheetSetting.Range("search1BtnTagName")
      searchBtnTagID = sheetSetting.Range("search1BtnTagID")
      searchBtnTagClass = sheetSetting.Range("search1BtnTagClass")
      
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
        
    ElseIf sheetWebCaptureList.Range("C" & line) = "����2" Then
      sheetWebCapture.Range("B1") = setSheetName & "_������"

      searchWord = sheetSetting.Range("search2Val")
      searchWordTagName = sheetSetting.Range("search2TagName")
      searchWordTagID = sheetSetting.Range("search2TagID")
      searchWordTagClass = sheetSetting.Range("search2TagClass")
      
      searchBtn = sheetSetting.Range("search2BtnVal")
      searchBtnTagName = sheetSetting.Range("search2BtnTagName")
      searchBtnTagID = sheetSetting.Range("search2BtnTagID")
      searchBtnTagClass = sheetSetting.Range("search2BtnTagClass")
      
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
    
    ElseIf sheetWebCaptureList.Range("C" & line) = "����3" Then
      sheetWebCapture.Range("B1") = setSheetName & "_������"

      searchWord = sheetSetting.Range("search3Val")
      searchWordTagName = sheetSetting.Range("search3TagName")
      searchWordTagID = sheetSetting.Range("search3TagID")
      searchWordTagClass = sheetSetting.Range("search3TagClass")
      
      searchBtn = sheetSetting.Range("search3BtnVal")
      searchBtnTagName = sheetSetting.Range("search3BtnTagName")
      searchBtnTagID = sheetSetting.Range("search3BtnTagID")
      searchBtnTagClass = sheetSetting.Range("search3BtnTagClass")
      
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


    ElseIf sheetWebCaptureList.Range("C" & line) Like "��i�K�F��*" Then
      sheetWebCapture.Range("B1") = setSheetName & "_�F��" & Replace(sheetWebCaptureList.Range("C" & line), "��i�K�F��", "") & "��"
      
      If sheetWebCaptureList.Range("C" & line) = "��i�K�F��1" Then
        sheetWebCapture.Range("B1") = setSheetName & "_cfm1"
        
        userID = setVal("Login1IDVal")
        userPW = setVal("Login1PWVal")
        
        Call ���͔���(setVal("Login1IDTagName"), setVal("Login1IDTagID"), setVal("Login1IDTagClass"), htmlType, htmlVal)
        
        Call �L�[���M(htmlType, htmlVal, userID)
        
        Call ���͔���(setVal("Login1Btn1TagName"), setVal("Login1Btn1TagID"), setVal("Login1Btn1TagClass"), htmlType, htmlVal)
        Call �N���b�N(htmlType, htmlVal)
        driver.Wait 1000
        
        Call ���͔���(setVal("Login1PWTagName"), setVal("Login1PWTagID"), setVal("Login1PWTagClass"), htmlType, htmlVal)
        Call �L�[���M(htmlType, htmlVal, userPW)
        
        Call ���͔���(setVal("Login1Btn2TagName"), setVal("Login1Btn2TagID"), setVal("Login1Btn2TagClass"), htmlType, htmlVal)
        Call �N���b�N(htmlType, htmlVal)
          
        
      ElseIf sheetWebCaptureList.Range("C" & line) = "��i�K�F��2" Then
      ElseIf sheetWebCaptureList.Range("C" & line) = "��i�K�F��3" Then
      End If

      
    ElseIf sheetWebCaptureList.Range("C" & line) = "�ʏ�F��" Then
      sheetWebCapture.Range("B1") = setSheetName & "_cfm1"
      
      userID = sheetSetting.Range("B2")
      userIDTagName = sheetSetting.Range("C2")
      
      userPW = sheetSetting.Range("B3")
      userPWTagName = sheetSetting.Range("C3")
      
      btn01TagName = sheetSetting.Range("C4")
      btn02TagName = sheetSetting.Range("C5")
      
      
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
  Selection.Name = sheetWebCapture.Range("B1")

  Call ProgressBar.showCount("WebCapture", line - 1, endLine - 1, driver.title)
'  openingDriver.ExecuteScript ("$('.appName').text('" & driver.title & " �擾����')")
  
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
  
  Call ProgressBar.showCount("WebCapture", line - 1, endLine - 1, driver.title)

End Function


'**************************************************************************************************
' * ��Book�ɕۑ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function ��Book�ɕۑ�()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long

  Call ProgressBar.showCount("WebCapture", line - 1, endLine - 1, driver.title)
  sheetWebCapture.Copy After:=targetBook.Worksheets(targetBook.Worksheets.count)
  
  targetBook.ActiveSheet.Name = sheetWebCapture.Range("B1")
  Call ProgressBar.showCount("WebCapture", line - 1, endLine - 1, driver.title)


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
    If tempSheet.Name Like "[Index,�X�V����,�\��]*" Then
    Else
      Worksheets("Index").Range("B" & line) = Worksheets(tempSheet.Name).Range("B2")
      Worksheets("Index").Range("C" & line) = Worksheets(tempSheet.Name).Range("B3")
      Worksheets("Index").Range("D" & line) = Worksheets(tempSheet.Name).Range("L1")
      Worksheets("Index").Range("A" & line) = "=HYPERLINK(""#'" & tempSheet.Name & "'!A1"",""" & ActiveWorkbook.Sheets(tempSheet.Name).Range("B1") & """)"
      line = line + 1
    End If
  Next
  Columns("D:D").Select
  Selection.NumberFormatLocal = "yyyy/mm/dd"
  
  Worksheets("Index").Select
  Application.Goto Reference:=Range("A1"), Scroll:=True

End Function



'**************************************************************************************************
' * �ۑ��V�[�g���`�F�b�N
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �ۑ��V�[�g���`�F�b�N()
  Dim line As Long, endLine As Long
  Dim targetSheetName As String, errMeg As String
  
  errMeg = ""
  endLine = sheetWebCaptureList.Cells(Rows.count, 1).End(xlUp).Row
  
  For line = 2 To endLine
    targetSheetName = sheetWebCaptureList.Range("B" & line)
    If Len(targetSheetName) > 30 Then
      sheetWebCaptureList.Range("B" & line).Style = "�G���["
      errMeg = "30�����ȉ�"
    Else
      sheetWebCaptureList.Range("B" & line).Style = "Normal"
    End If
  Next
  If errMeg <> "" Then
    Call Library.showNotice(411, errMeg, True)
  End If
End Function



'**************************************************************************************************
' * ���͔���
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function ���͔���(strName As String, strID As String, strClass As String, htmlType As String, htmlVal As String)
  
  If strName <> "" Then
    htmlType = "Name"
    htmlVal = strName
  
  ElseIf strID <> "" Then
    htmlType = "ID"
    htmlVal = strID
  
  ElseIf strClass <> "" Then
    htmlType = "Class"
    htmlVal = strClass
  End If

End Function


'**************************************************************************************************
' * �L�[���M
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �L�[���M(htmlType As String, htmlVal As String, sendVal As String)
  
  Select Case htmlType
    Case "Name"
      driver.FindElementByName(htmlVal).SendKeys (sendVal)
      
    Case "ID"
      driver.FindElementById(htmlVal).SendKeys (sendVal)
      
    Case "Class"
      driver.FindElementByClass(htmlVal).SendKeys (sendVal)
  End Select

End Function



'**************************************************************************************************
' * �N���b�N
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �N���b�N(htmlType As String, htmlVal As String)
  
  Select Case htmlType
    Case "Name"
      driver.FindElementByName(htmlVal).Click
      
    Case "ID"
      driver.FindElementById(htmlVal).Click
      
    Case "Class"
      driver.FindElementByClass(htmlVal).Click
  End Select

End Function










