Attribute VB_Name = "リンク抽出"
Public Const startLine As Long = 2


Dim activLine As Long
Dim pageCount As Long
Dim baseURL As String
Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
Dim targetFileDir As String

'*******************************************************************************************************
' * 取得開始
' *
' * @author Bunpei.Koizumi<Koizumi.Bunpei@trans-cosmos.co.jp>
'*******************************************************************************************************
Function 取得開始()
  Dim elements As WebElements
  
  Dim c As Range
  Dim URLColumn As Integer
  Dim IDPW As String
  
'  On Error GoTo catchError
  
  Call init.setting
  Call ProgressBar.showStart
  Call WebToolLib.Chrome起動
  
  For colLine = 3 To Cells(1, Columns.count).End(xlToLeft).Column

    endLine = sheetLinkExtract.Cells(Rows.count, colLine).End(xlUp).Row
    
    For line = startLine To endLine
      sheetTmp.Select
      Call Library.delSheetData
      sheetTmp.Columns("A:B").ColumnWidth = 100
      sheetLinkExtract.Select
      baseURL = sheetLinkExtract.Cells(1, colLine)
      
      
      'CSVファイル保存場所生成
      targetFileDir = setVal("workPath") & "\リンク抽出\" & sheetLinkExtract.Cells(1, colLine)
      Call Library.makeDir(targetFileDir)
      
      targetURL = sheetLinkExtract.Cells(line, colLine)
      
      If targetURL = "" Then
        GoTo nextForLabel
      End If
      
      driver.Get targetURL
      Call aタグ調査
      Call CSV作成
    Next
nextForLabel:
  Next

  
  Call WebToolLib.Chrome終了
  
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
' * aタグ調査
' *
' * @author Bunpei.Koizumi<Koizumi.Bunpei@trans-cosmos.co.jp>
'*******************************************************************************************************
Function aタグ調査()
  Dim line As Long, endLine As Long
  Dim elements As WebElements, targetElements As WebElements
  
  Dim exitFlg As Boolean
  Dim getInnerHtmlString As Variant
  Dim myBy As New By
  Dim URL As String
  Dim i As Long
  
'  On Error GoTo catchError
  
  Call Library.showDebugForm("検索範囲：" & sheetLinkExtract.Cells(4, colLine))
  
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
    
    Call ProgressBar.showCount("リンクチェック", i, elements.count, "aタグ調査：" & URL)
    
    
    '除外URLかどうか
    For setLine = 3 To sheetSetting.Cells(Rows.count, 7).End(xlUp).Row
      If URL <> "" And URL Like sheetSetting.Range("G" & setLine) & "*" Then
        URL = ""
        Exit For
      End If
    Next
    
    
    '除外文字かどうか
    Call Library.showDebugForm("attText：" & attText)
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
  
'エラー発生時====================================
catchError:
   
  If Not driver Is Nothing Then
    WebToolLib.Chrome終了
  End If

  Call Library.showNotice(Err.Number, Err.Description, True)
End Function


'*******************************************************************************************************
' * CSV作成
' *
' * @author Bunpei.Koizumi<Koizumi.Bunpei@trans-cosmos.co.jp>
'*******************************************************************************************************
Function CSV作成()
  Dim csvLine As Long, csvEndLine As Long


'  On Error GoTo catchError
  
  targetFilePath = targetFileDir & "\" & sheetLinkExtract.Cells(line, 2)
  
  sheetTmp.Select
  csvEndLine = sheetTmp.Cells(Rows.count, 1).End(xlUp).Row
  
  Call Library.showDebugForm("targetFilePath：" & targetFilePath)
  Open targetFilePath For Output As #1
  For csvLine = 1 To csvEndLine
    Print #1, Replace(sheetTmp.Range("B" & csvLine), vbLf, "") & "," & sheetTmp.Range("A" & csvLine)
  Next
  Close #1

  Exit Function
'エラー発生時====================================
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function
































