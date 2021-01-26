Attribute VB_Name = "サイトマップ"
Dim activLine As Long
Dim pageCount As Long
Dim baseURL As String

'*******************************************************************************************************
' * 取得開始
' *
' * @author Bunpei.Koizumi<Koizumi.Bunpei@trans-cosmos.co.jp>
'*******************************************************************************************************
Function 取得開始()
  Dim elements As WebElements
  Dim line As Long, endLine As Long
  Dim c As Range
  Dim URLColumn As Integer
  
'  On Error GoTo catchError
  
  exitFlg = False
  Call init.setting
  Call ProgressBar.showStart
  
  baseURL = setVal("siteMapURL") & "/"
  targetURL = baseURL
  pageCount = 1
  
  sheetTmp.Select
  Call Library.delSheetData
  Columns("A:A").ColumnWidth = 100
  
  sheetSitemap.Select
  Call Library.delSheetData(3)
  
  URLColumn = Library.getColumnNo(SitemapCell("url"))
  
  Call WebToolLib.Chrome起動
  
  driver.Get targetURL
  Call ページ情報取得
  Call aタグ調査

  Do While sheetTmp.Range("A1") <> ""
    pageCount = pageCount + 1
    targetURL = sheetTmp.Range("A1")
    sheetTmp.Rows("1:1").Delete Shift:=xlUp
    
    '/で終わっていた場合
    Call Library.showDebugForm("targetURL：" & targetURL)
    If targetURL Like "*index.*" Then
      replaceURL = Mid(targetURL, Len("index") + InStr(1, targetURL, "index"))
      Set c = sheetSitemap.UsedRange.Columns(URLColumn).Find(What:=targetURL, LookIn:=xlValues, LookAt:=xlPart)
      If c Is Nothing Then
        reTargetURL = Replace(targetURL, replaceURL, "")
        reTargetURL = Replace(reTargetURL, "index", "")
        Set c = sheetSitemap.UsedRange.Columns(URLColumn).Find(What:=reTargetURL, LookIn:=xlValues, LookAt:=xlPart)
      End If
    Else
      Set c = sheetSitemap.UsedRange.Columns(URLColumn).Find(What:=targetURL, LookIn:=xlValues, LookAt:=xlPart)
    End If
    
    
    '調査済みURLか確認
    If c Is Nothing Then
      Call ProgressBar.showCount("サイトマップ生成", 0, 10, "ページ遷移：" & targetURL)
      Call Library.waitTime(2000)
      DoEvents
      driver.Get targetURL
      driver.Wait 1000
      Call ProgressBar.showCount("サイトマップ生成", 10, 10, "ページ遷移：" & targetURL)
      
      Call ページ情報取得
      Call aタグ調査
    End If
    
    If sheetSitemap.Cells(Rows.count, 1).End(xlUp).Row = 12 Then
      Exit Do
    End If
  Loop
  Set c = Nothing
  
  Call WebToolLib.Chrome終了
  
  Call 整形
  
  Call ProgressBar.showEnd
  Call Library.endScript
  Exit Function
'--------------------------------------------------------------------------------------------------
'エラー発生時の処理
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
' * ページ情報取得
' *
' * @author Bunpei.Koizumi<Koizumi.Bunpei@trans-cosmos.co.jp>
'*******************************************************************************************************
Function ページ情報取得()
  
  Dim line As Long, endLine As Long
  Dim descElements As WebElements, keyWdElements As WebElements
  Dim myBy As New By
  Dim pageSourceTag As String
  
  On Error Resume Next
  ThisWorkbook.Save

  line = sheetSitemap.Cells(Rows.count, 1).End(xlUp).Row + 1
  
  Call ProgressBar.showCount("サイトマップ生成", 0, 10, "ページ情報取得：" & driver.title)

  sheetSitemap.Range("A" & line).FormulaR1C1 = "=ROW()-2"
  sheetSitemap.Range(SitemapCell("title") & line) = driver.title
  
'  If baseURL = driver.URL Then
'    sheetSitemap.Range("D" & line) = "/"
'  Else
'    sheetSitemap.Range("D" & line) = Replace(driver.URL, baseURL, "")
'  End If
  sheetSitemap.Range(SitemapCell("url") & line) = driver.URL
    
    
  If driver.IsElementPresent(myBy.XPath("//meta[@name=""description""]")) Then
    Set descElements = driver.FindElementsByXPath("//meta[@name=""description""]")
    sheetSitemap.Range(SitemapCell("description") & line) = descElements.Item(1).Attribute("content")
  End If
    
  If driver.IsElementPresent(myBy.XPath("//meta[@name=""keywords""]")) Then
    Set keyWdElements = driver.FindElementsByXPath("//meta[@name=""keywords""]")
    sheetSitemap.Range(SitemapCell("keywords") & line) = keyWdElements.Item(1).Attribute("content")
  End If
  
  'GMTタグがあるかどうか調査
  pageSourceTag = driver.PageSource
  If pageSourceTag Like "*googletagmanager*" Then
    sheetSitemap.Range(SitemapCell("google") & line) = "○"
  End If
  If pageSourceTag Like "*yjtag.jp*" Then
    sheetSitemap.Range(SitemapCell("yahoo") & line) = "○"
  End If
  
  'H1タグ調査
  sheetSitemap.Range(SitemapCell("htmlTag_H1") & line) = driver.FindElementByTag("h1").Text
  
  pageSourceTag = ""
  Set descElements = Nothing
  Set keyWdElements = Nothing
End Function


'*******************************************************************************************************
' * aタグ調査
' *
' * @author Bunpei.Koizumi<Koizumi.Bunpei@trans-cosmos.co.jp>
'*******************************************************************************************************
Function aタグ調査()
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
  line = sheetTmp.Cells(Rows.count, 1).End(xlUp).Row + 1
  
  For i = 1 To elements.count
    URL = elements.Item(i).Attribute("href")
    Call ProgressBar.showCount("サイトマップ生成", i, elements.count, "aタグ調査：" & URL)
    
    If URL = "" Or URL Like "javascript*" Or Not (URL Like targetURL & "*") Then
      GoTo L_nextFor
    Else
      Call Library.showDebugForm("  aタグ調査：" & URL)
    End If
    
    
    '除外URLかどうか
    For setLine = 3 To sheetSetting.Cells(Rows.count, 7).End(xlUp).Row
      If URL Like sheetSetting.Range("G" & setLine) & "*" Then
        GoTo L_nextFor
      End If
    Next
    
    '同一ドメインかどうか
    If Not (URL = baseURL Or URL = targetURL Or URL Like targetURL & "*[#,\?]*") Then
      sheetTmp.Range("A" & line) = URL
      line = line + 1
    
    ElseIf InStr(URL, "#") <> 0 And InStr(URL, "?") <> 0 Then
      URL = Left(URL, InStr(URL, "#") - 1)
      URL = Left(URL, InStr(URL, "?") - 1)
      If URL <> baseURL Then
       sheetTmp.Range("A" & line) = URL
       line = line + 1
      End If
    
    ElseIf InStr(URL, "#") <> 0 Then
      URL = Left(URL, InStr(URL, "#") - 1)
      If URL <> baseURL Then
       sheetTmp.Range("A" & line) = URL
       line = line + 1
      End If
    
    ElseIf InStr(URL, "?") <> 0 Then
      URL = Left(URL, InStr(URL, "?") - 1)
      If URL <> baseURL Then
       sheetTmp.Range("A" & line) = URL
       line = line + 1
      End If
    
    End If
    
L_nextFor:
  Next
  
  endLine = sheetTmp.Cells(Rows.count, 1).End(xlUp).Row
  If endLine = 1 Then
    Exit Function
  End If
  
  '並び替えと重複削除
'  sheetTmp.Select
  sheetTmp.Sort.SortFields.Clear
  
  sheetTmp.Sort.SortFields.Add Key:=Range("A1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
  With sheetTmp.Sort
      .SetRange Range("A1:A" & endLine)
      .Header = xlGuess
      .MatchCase = False
      .Orientation = xlTopToBottom
      .SortMethod = xlPinYin
      .Apply
  End With
  
  sheetTmp.Range("A1:A" & endLine).RemoveDuplicates Columns:=1, Header:=xlNo

  Exit Function
  
'エラー発生時====================================
catchError:
   
  If Not driver Is Nothing Then
    WebToolLib.Chrome終了
  End If

  Call Library.showNotice(Err.Number, Err.Description, True)
End Function


'*******************************************************************************************************
' * 整形
' *
' * @author Bunpei.Koizumi<Koizumi.Bunpei@trans-cosmos.co.jp>
'*******************************************************************************************************
Function 整形()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long, endColName As String
  Dim titleVal As String, titleVals As Variant
  Dim URLVal As String, URLVals As Variant
  Dim count As Long
  Dim startLine As Long
  Dim reg As Object
  
'  On Error GoTo catchError
  
  Call init.setting
  sheetSitemap.Select
  baseURL = setVal("siteMapURL")

  Set reg = CreateObject("VBScript.RegExp")
  
  '正規表現の指定
  With reg
    .Pattern = "( |　)?[\-|\||｜]( |　)?"
    .IgnoreCase = False
    .Global = True
  End With

'  colLine = Library.getColumnNo(SitemapCell("dirName"))
  endLine = sheetSitemap.Cells(Rows.count, 1).End(xlUp).Row
  
  sheetSitemap.Range(SitemapCell("level_1") & "3:" & SitemapCell("level_" & Range("maxTitleLevel")) & endLine).ClearContents
  sheetSitemap.Range(SitemapCell("dirLevel_1") & "3:" & SitemapCell("dirLevel_" & Range("maxDirLevel")) & endLine).ClearContents
  
  For line = 3 To endLine
    count = 1
    titleVal = sheetSitemap.Range(SitemapCell("title") & line)
    URLVal = sheetSitemap.Range(SitemapCell("url") & line)
    
    'タイトルの区切り設定
    titleVal = reg.Replace(titleVal, "<|>")
    titleVals = Split(titleVal, "<|>")


    'For i = 0 To UBound(titleVals)
    For i = UBound(titleVals) To 0 Step -1
      sheetSitemap.Range(SitemapCell("level_" & count) & line) = titleVals(i)
      If count = setVal("maxTitleLevel") Then
        Exit For
      End If
      count = count + 1
    Next
  
    URLVal = Replace(URLVal, baseURL, "")
    URLVals = Split(URLVal, "/")
    count = 1
    
    sheetSitemap.Range(SitemapCell("level") & line) = UBound(URLVals)
    
    For i = 0 To UBound(URLVals)
      If i = 0 And URLVals(0) = "" Then
        sheetSitemap.Range(SitemapCell("dirLevel_1") & line) = "/"
      ElseIf InStr(URLVals(i), ".") = 0 Then
        sheetSitemap.Range(SitemapCell("dirLevel_" & count) & line) = URLVals(i)
      End If
      
      If count = setVal("maxDirLevel") Then
        Exit For
      End If
      count = count + 1
    Next
    
    sheetSitemap.Range(SitemapCell("testURL") & line).FormulaR1C1 = "=siteMapURL_test & ""/"" &  RC[-4] & ""/"" & RC[-3] & ""/"" & RC[-2] & ""/"" & RC[-1]"
    
    
    sheetSitemap.Range(SitemapCell("preURL") & line).FormulaR1C1 = "=siteMapURL_pre  & ""/"" &  RC[-5] & ""/"" & RC[-4] & ""/"" & RC[-3] & ""/"" & RC[-2]"
                                                                    
    sheetSitemap.Range(SitemapCell("proURL") & line).FormulaR1C1 = "=siteMapURL_pro & ""/"" &  RC[-6] & ""/"" & RC[-5] & ""/"" & RC[-4] & ""/"" & RC[-3]"
    
    
  Next
  
  '逆L字罫線設定
  endColName = Library.getColumnName(sheetSitemap.Cells(2, Columns.count).End(xlToLeft).Column)
  Call Library.罫線_クリア(sheetSitemap.Range("A3" & ":" & endColName & endLine))
    
  For count = 1 To 3
    startLine = 2 + count
    For line = 3 To endLine
      If sheetSitemap.Range(SitemapCell("dirLevel_" & count) & line) = "" Then
        startLine = line + 1
      ElseIf sheetSitemap.Range(SitemapCell("dirLevel_" & count) & line) <> sheetSitemap.Range(SitemapCell("dirLevel_" & count) & line + 1) Then
        sheetSitemap.Range(SitemapCell("dirLevel_" & count) & startLine & ":" & endColName & line).Select
        Call Library.罫線_破線_逆L字(sheetSitemap.Range(SitemapCell("dirLevel_" & count) & startLine & ":" & endColName & line), 0)
        startLine = line + 1
      
      End If
    Next
  Next
  Call Library.罫線_実線_垂直(sheetSitemap.Range("A3" & ":" & SitemapCell("proURL") & endLine), 0)
  Call Library.罫線_破線_水平(sheetSitemap.Range("A3" & ":" & SitemapCell("proURL") & endLine), 0)
  Call Library.罫線_実線_垂直(sheetSitemap.Range(SitemapCell("url") & "3" & ":" & endColName & endLine), 0)
  Call Library.罫線_実線_囲み(sheetSitemap.Range("A3" & ":" & endColName & endLine), 0)
  
  Application.Goto Reference:=Range("A1"), Scroll:=True
  Exit Function
'エラー発生時====================================
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function































