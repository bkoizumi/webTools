Attribute VB_Name = "WebCapture"

Dim driver As New Selenium.WebDriver
Dim targetURL As String
Dim line As Long
Dim endLine As Long


'*******************************************************************************************************
' * optionForm表示
' *
' * @author Bunpei.Koizumi<Koizumi.Bunpei@trans-cosmos.co.jp>
'*******************************************************************************************************
Function showOptionForm()
  
  'ダイアログへ表示
  With optionForm
    .StartUpPosition = 0
    .Top = Application.Top + 150
    .Left = Application.Left + (ActiveWindow.Height / 2)
  End With

  optionForm.Show vbModeless

End Function


'*******************************************************************************************************
' * Chrome起動
' *
' * @author Bunpei.Koizumi<Koizumi.Bunpei@trans-cosmos.co.jp>
'*******************************************************************************************************
Function Chrome起動()
  
  Call init.setting
  
  'chrome 拡張機能読み込み
  Call ProgressBar.showCount("WebCapture", 0, 10, "Chrome起動")
  With driver
    .AddArgument ("--lang=ja")
    .AddArgument ("--user-data-dir=" & BrowserProfiles("default"))
    .AddArgument ("--window-size=1200,600")
    .AddArgument ("--hide-scrollbars")
    .AddArgument ("--disable-gpu")
    
    If setVal("debugMode") = "develop" Then
      Call Library.showDebugForm("WebCapture", "シークレットモード")
      
      .AddArgument ("--incognito")                   'シークレットモード
    Else
      Call Library.showDebugForm("WebCapture", "headlessモード")
      .AddArgument ("--headless")
    End If
    
    If setVal("InstNetwork") = "TCI" Then
      .AddArgument ("--proxy-server=" & setVal("ProxyURL") & ":" & setVal("ProxyPort"))
    End If
    
    .start "chrome"
    .Wait 1000
  End With
  
  Call ProgressBar.showCount("WebCapture", 10, 10, "Chrome起動")
  Call Library.showDebugForm("WebCapture", "Chrome起動完了")
  
End Function


'*******************************************************************************************************
' * 取得開始
' *
' * @author Bunpei.Koizumi<Koizumi.Bunpei@trans-cosmos.co.jp>
'*******************************************************************************************************
Function 取得開始()
  Dim elements As WebElements
  Dim shellProcID As Long
  
  
'  On Error GoTo catchError

  
  Set targetBook = Workbooks.Open(FileName:=AppWebCapturePath & "\新規Book.xlsm", ReadOnly:=True)
  Windows(targetBook.Name).WindowState = xlMinimized
  ThisWorkbook.Activate
  
  '名前をつけて保存
  targetFilePath = Library.getRegistry("WebCapturePath")
  If targetFilePath <> "" Then
    targetFileName = Dir(targetFilePath)
    targetFilePath = Replace(targetFilePath, "\" & targetFileName, "")
  End If
  targetFilePath = Library.getFilePath(targetFilePath, targetFileName, "出力ファイルの保存", 1)
  If targetFilePath = "" Then
    targetBook.Close
    Set targetBook = Nothing
    Call Library.showNotice(100, , True)
  End If
  
  targetBook.SaveAs FileName:=targetFilePath
  
  'chromeDriver の更新確認-------------------------------------------------------------------------
  shellProcID = Shell(binPath & "\SeleniumBasic\updateChromeDriver.bat", vbNormalFocus)
  Call Library.chkShellEnd(shellProcID)
  
  Call Chrome起動
  
  endLine = sheetWebCaptureList.Cells(Rows.count, 1).End(xlUp).Row
  
  For line = 2 To endLine
    targetURL = sheetWebCaptureList.Range("A" & line)
    If targetURL <> "" Then
      Call ProgressBar.showCount("WebCapture", line - 1, endLine - 1, driver.title)
      
reLoad:
      driver.Timeouts.PageLoad = 60000
      driver.Window.SetSize 1200, 600
      driver.Wait 2000
      driver.Get targetURL
      driver.Wait 1000
      
      Call ページ情報取得
      Call 別Bookに保存
      
      If sheetWebCaptureList.Range("F" & line) <> "" Then
        ThisWorkbook.Activate
        sheetWebCapture.Select
        sheetWebCapture.Shapes.Range(Array(sheetWebCapture.Range("B1").Text)).Delete
        
        Call ページ情報取得(True)
        Call 別Bookに保存
      
      End If

      
      ThisWorkbook.Activate
      sheetWebCapture.Shapes.Range(Array(sheetWebCapture.Range("B1").Text)).Delete

    End If
  Next
  Call ProgressBar.showCount("WebCapture", 0, 100, "index作成")
  Call index作成

  driver.Close
  driver.Quit
  Set driver = Nothing

  '出力ファイルの上書き
  targetBook.SaveAs
  targetBook.Close
  Set targetBook = Nothing


  sheetWebCapture.Range("B1") = ""
  sheetWebCapture.Range("B2") = ""
  sheetWebCapture.Range("B3") = ""
  sheetWebCapture.Range("L1") = ""
  
  Call Library.setRegistry("WebCapturePath", targetFilePath)


  
  
  Exit Function
'--------------------------------------------------------------------------------------------------
'エラー発生時の処理
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
' * ページ情報取得
' *
' * @author Bunpei.Koizumi<Koizumi.Bunpei@trans-cosmos.co.jp>
'*******************************************************************************************************
Function ページ情報取得(Optional actionFlg As Boolean)
  Dim elements As WebElements
  Dim getInnerHtmlString As Variant
  Dim myBy As New By
  Dim ks As New Keys
  Dim setSheetName As String
  
  
  Dim page_width As Long, page_height As Long
  Dim userID As String, userPW As String
  
  Dim userIDTagID As String, userPWTagID As String
  Dim userIDTagName As String, userPWTagName As String
  
  Dim btn01TagId As String, btn01TagName As String
  Dim btn02TagId As String, btn02TagName As String
  
  
  Call ProgressBar.showCount("WebCapture", line - 1, endLine - 1, driver.title)

  'ページ情報格納
  setSheetName = sheetWebCaptureList.Range("B" & line)
  If setSheetName = "" Then
    setSheetName = setVal("sheetName") & Format(line - 1, "000")
  End If
  
  sheetWebCapture.Select
  sheetWebCapture.Range("B1") = setSheetName
  sheetWebCapture.Range("L1") = Format(Now(), "yyyy/mm/dd hh:nn:ss")


  If actionFlg = True Then
    If sheetWebCaptureList.Range("B" & line) = "検索1" Then
      sheetWebCapture.Range("B1") = setSheetName & "_検索後"

      searchWord = sheetSetting.Range("search1Val")
      searchWordTagName = sheetSetting.Range("search1TagName")
      searchWordTagID = sheetSetting.Range("search1TagID")
      searchWordTagClass = sheetSetting.Range("search1TagClass")
      
      searchBtn = sheetSetting.Range("search1BtnVal")
      searchBtnTagName = sheetSetting.Range("search1BtnTagName")
      searchBtnTagID = sheetSetting.Range("search1BtnTagID")
      searchBtnTagClass = sheetSetting.Range("search1BtnTagClass")
      
      '検索キーワード入力タグ調査
      If driver.IsElementPresent(myBy.Name(searchWordTagName)) And searchWordTagName <> "" Then
        driver.FindElementByName(searchWordTagName).SendKeys (searchWord)
        
      ElseIf driver.IsElementPresent(myBy.Class(searchWordTagClass)) And searchWordTagClass <> "" Then
        driver.FindElementByClass(searchWordTagClass).SendKeys (searchWord)
      
      ElseIf driver.IsElementPresent(myBy.id(searchWordTagID)) And searchWordTagID <> "" Then
        driver.FindElementById(searchWordTagID).SendKeys (searchWord)
      End If
      
      '検索ボタン調査
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
        
    ElseIf sheetWebCaptureList.Range("B" & line) = "検索2" Then
      sheetWebCapture.Range("B1") = setSheetName & "_検索後"

      searchWord = sheetSetting.Range("search2Val")
      searchWordTagName = sheetSetting.Range("search2TagName")
      searchWordTagID = sheetSetting.Range("search2TagID")
      searchWordTagClass = sheetSetting.Range("search2TagClass")
      
      searchBtn = sheetSetting.Range("search2BtnVal")
      searchBtnTagName = sheetSetting.Range("search2BtnTagName")
      searchBtnTagID = sheetSetting.Range("search2BtnTagID")
      searchBtnTagClass = sheetSetting.Range("search2BtnTagClass")
      
      '検索キーワード入力タグ調査
      If driver.IsElementPresent(myBy.Name(searchWordTagName)) And searchWordTagName <> "" Then
        driver.FindElementByName(searchWordTagName).SendKeys (searchWord)
        
      ElseIf driver.IsElementPresent(myBy.Class(searchWordTagClass)) And searchWordTagClass <> "" Then
        driver.FindElementByClass(searchWordTagClass).SendKeys (searchWord)
      
      ElseIf driver.IsElementPresent(myBy.id(searchWordTagID)) And searchWordTagID <> "" Then
        driver.FindElementById(searchWordTagID).SendKeys (searchWord)
      End If
      
      '検索ボタン調査
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
    
    ElseIf sheetWebCaptureList.Range("F" & line) = "検索3" Then
      sheetWebCapture.Range("B1") = setSheetName & "_検索後"

      searchWord = sheetSetting.Range("search3Val")
      searchWordTagName = sheetSetting.Range("search3TagName")
      searchWordTagID = sheetSetting.Range("search3TagID")
      searchWordTagClass = sheetSetting.Range("search3TagClass")
      
      searchBtn = sheetSetting.Range("search3BtnVal")
      searchBtnTagName = sheetSetting.Range("search3BtnTagName")
      searchBtnTagID = sheetSetting.Range("search3BtnTagID")
      searchBtnTagClass = sheetSetting.Range("search3BtnTagClass")
      
      '検索キーワード入力タグ調査
      If driver.IsElementPresent(myBy.Name(searchWordTagName)) And searchWordTagName <> "" Then
        driver.FindElementByName(searchWordTagName).SendKeys (searchWord)
        
      ElseIf driver.IsElementPresent(myBy.Class(searchWordTagClass)) And searchWordTagClass <> "" Then
        driver.FindElementByClass(searchWordTagClass).SendKeys (searchWord)
      
      ElseIf driver.IsElementPresent(myBy.id(searchWordTagID)) And searchWordTagID <> "" Then
        driver.FindElementById(searchWordTagID).SendKeys (searchWord)
      End If
      
      '検索ボタン調査
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


    ElseIf sheetWebCaptureList.Range("F" & line) = "二段階ログイン" Then
      sheetWebCapture.Range("B1") = setSheetName & "_認証後"
      
      userID = sheetSetting.Range("B2")
      userIDTagName = sheetSetting.Range("C2")
      
      userPW = sheetSetting.Range("B3")
      userPWTagName = sheetSetting.Range("C3")
      
      btn01TagName = sheetSetting.Range("C4")
      btn02TagName = sheetSetting.Range("C5")
      
      
      driver.FindElementByName(userIDTagName).SendKeys (userID)
      driver.FindElementByName(btn01TagName).Click
      driver.Wait 1000
      
      driver.FindElementByName(userPWTagName).SendKeys (userPW)
      driver.FindElementByName(btn02TagName).Click
    ElseIf sheetWebCaptureList.Range("F" & line) = "通常ログイン" Then
      sheetWebCapture.Range("B1") = setSheetName & "_認証後"
      
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
  
  'ページ情報格納
  sheetWebCapture.Range("B2") = driver.title
  sheetWebCapture.Range("B3") = driver.URL
  
  
  '画面ショット処理
  page_width = driver.ExecuteScript("return document.body.scrollWidth")
  page_height = driver.ExecuteScript("return document.body.scrollHeight")
  driver.Window.SetSize page_width, page_height

  driver.TakeScreenshot.Copy
  ThisWorkbook.Activate
  sheetWebCapture.Select
  sheetWebCapture.Range("A5").Select
  Call Library.waitTime(2000)
  
  'ActiveSheet.Paste
  ActiveSheet.PasteSpecial Format:="ビットマップ", link:=False, DisplayAsIcon:=False
  Selection.Name = setSheetName

  Call ProgressBar.showCount("WebCapture", line - 1, endLine - 1, driver.title)
  
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
' * 別Bookに保存
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function 別Bookに保存()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long

  Call ProgressBar.showCount("WebCapture", line - 1, endLine - 1, driver.title)
  sheetWebCapture.Copy After:=targetBook.Worksheets(targetBook.Worksheets.count)
  
  targetBook.ActiveSheet.Name = sheetWebCapture.Range("B1")
  Call ProgressBar.showCount("WebCapture", line - 1, endLine - 1, driver.title)


End Function


'**************************************************************************************************
' * index作成
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function index作成()
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
' * 保存シート名チェック
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function 保存シート名チェック()
  Dim line As Long, endLine As Long
  Dim targetSheetName As String, errMeg As String
  
  errMeg = ""
  endLine = sheetWebCaptureList.Cells(Rows.count, 1).End(xlUp).Row
  
  For line = 2 To endLine
    targetSheetName = sheetWebCaptureList.Range("B" & line)
    If Len(targetSheetName) > 30 Then
      sheetWebCaptureList.Range("B" & line).Style = "エラー"
      errMeg = "30文字以下"
    Else
      sheetWebCaptureList.Range("B" & line).Style = "Normal"
    End If
  Next
  If errMeg <> "" Then
    Call Library.showNotice(411, errMeg, True)
  End If
End Function




















