Attribute VB_Name = "スクレイピング"
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Dim driver As New Selenium.WebDriver
'Dim driver As Object


'**************************************************************************************************
' * データ消去
' *
' * @author Bunpei.Koizumi<koizumi.bunpei@trans-cosmos.co.jp>
'**************************************************************************************************
Function SheetDataDelete()
  
  Call Library.delSheetData(5)
  Cells.FormatConditions.Delete
  
  Range("E2") = "＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊"
  Range("E3") = ""
  
  endLine = Cells(3, Columns.count).End(xlToLeft).Column - 3
  Range(Cells(4, 10), Cells(4, endLine)).ClearContents

End Function


'**************************************************************************************************
' * ユーザーフォームの表示
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
' * Webデータ取得
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

  'Call メンテナンス.利用ログ取得("KOETOL：" & setVal("SiteName"))
  Call ProgressBar.showCount("Webデータ取得", 0, 10, "")
  
  'chromeDriver の更新確認-------------------------------------------------------------------------
  rc = Shell(binPath & "\SeleniumBasic\updateChromeDriver.bat", vbNormalFocus)
  
  '既存データ削除----------------------------------------------------------------------------------
  sheetKoetol.Select
  
  Call SheetDataDelete
  
'  Call ProgressBar.showCount("Webデータ取得", 0, 10, setVal("SiteName"))
  Select Case setVal("SiteName")
    Case "@コスメ"
        Call site_cosme
        
    Case "Amazon"
        Call site_Amazon
        
    Case "楽天"
        Call site_rakuten
        
    Case "価格.com"
        Call site_kakakuCom
        
    Case "コスメデネット"
        Call site_cosmedeNet
  End Select
    
  Call 書式設定

  Application.Goto Reference:=Range("I5"), Scroll:=True

  '終了画面表示
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

'エラー発生時=====================================================================================
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)

End Function


'**************************************************************************************************
' * グラフ用データ集計
' *
' * @author Bunpei.Koizumi<koizumi.bunpei@trans-cosmos.co.jp>
'**************************************************************************************************
Function グラフ用データ集計()
 
  Dim endRowLine As Long, endColLine As Long
  Dim line As Long, rowLine As Long, colLine As Long
  
  Call init.setting
  
  ' プログレスバーの表示開始
'  Call ProgressBar.showStart
  
  '既存データ削除
  sheetGraf.Range("E7:E9").ClearContents
  sheetGraf.Range("O7:O8").ClearContents
  sheetGraf.Range("D27:G33").ClearContents
  sheetGraf.Range("C42:G71").ClearContents
  sheetGraf.Range("D85:J91").ClearContents
  sheetGraf.Range("C101:K130").ClearContents
  
  endRowLine = sheetKoetol.Cells(Rows.count, 3).End(xlUp).Row
  
  '男女比率
'  ProgressBar.showCount "男女比率集計", 1, 10, "処理中・・・"
  sheetGraf.Range("E7") = WorksheetFunction.CountIf(sheetKoetol.Range("F5:F" & endRowLine), "男性")
  sheetGraf.Range("E8") = WorksheetFunction.CountIf(sheetKoetol.Range("F5:F" & endRowLine), "女性")
  sheetGraf.Range("E9") = WorksheetFunction.CountIf(sheetKoetol.Range("F5:F" & endRowLine), "不明")
  For i = 7 To 9
    If sheetGraf.Cells(i, 5) = 0 Then
      sheetGraf.Cells(i, 5) = ""
    End If
  Next
  
  
  'ネガポジ比率
'  ProgressBar.showCount "ネガポジ比率集計", 1, 10, "処理中・・・"
  sheetGraf.Range("O7") = WorksheetFunction.Sum(sheetKoetol.Range("AY5:AY" & endRowLine))
  sheetGraf.Range("O8") = WorksheetFunction.Sum(sheetKoetol.Range("AZ5:AZ" & endRowLine))
  For i = 7 To 8
    If sheetGraf.Cells(i, 15) = 0 Then
      sheetGraf.Cells(i, 15) = " "
    End If
  Next
  
  '年代・性別
'  ProgressBar.showCount "年齢比率集計", 1, 10, "処理中・・・"
  sheetGraf.Range("D27") = WorksheetFunction.CountIf(sheetKoetol.Range("G5:G" & endRowLine), "10代")
  sheetGraf.Range("D28") = WorksheetFunction.CountIf(sheetKoetol.Range("G5:G" & endRowLine), "20代")
  sheetGraf.Range("D29") = WorksheetFunction.CountIf(sheetKoetol.Range("G5:G" & endRowLine), "30代")
  sheetGraf.Range("D30") = WorksheetFunction.CountIf(sheetKoetol.Range("G5:G" & endRowLine), "40代")
  sheetGraf.Range("D31") = WorksheetFunction.CountIf(sheetKoetol.Range("G5:G" & endRowLine), "50代")
  sheetGraf.Range("D32") = WorksheetFunction.CountIf(sheetKoetol.Range("G5:G" & endRowLine), "60代") + _
                           WorksheetFunction.CountIf(sheetKoetol.Range("G5:G" & endRowLine), "70代")
  sheetGraf.Range("D33") = WorksheetFunction.CountIf(sheetKoetol.Range("G5:G" & endRowLine), "不明")
  
  sheetGraf.Range("E27") = WorksheetFunction.CountIfs(sheetKoetol.Range("F5:F" & endRowLine), "男性", sheetKoetol.Range("G5:G" & endRowLine), "10代")
  sheetGraf.Range("E28") = WorksheetFunction.CountIfs(sheetKoetol.Range("F5:F" & endRowLine), "男性", sheetKoetol.Range("G5:G" & endRowLine), "20代")
  sheetGraf.Range("E29") = WorksheetFunction.CountIfs(sheetKoetol.Range("F5:F" & endRowLine), "男性", sheetKoetol.Range("G5:G" & endRowLine), "30代")
  sheetGraf.Range("E30") = WorksheetFunction.CountIfs(sheetKoetol.Range("F5:F" & endRowLine), "男性", sheetKoetol.Range("G5:G" & endRowLine), "40代")
  sheetGraf.Range("E31") = WorksheetFunction.CountIfs(sheetKoetol.Range("F5:F" & endRowLine), "男性", sheetKoetol.Range("G5:G" & endRowLine), "50代")
  sheetGraf.Range("E32") = WorksheetFunction.CountIfs(sheetKoetol.Range("F5:F" & endRowLine), "男性", sheetKoetol.Range("G5:G" & endRowLine), "60代")
  sheetGraf.Range("E33") = WorksheetFunction.CountIfs(sheetKoetol.Range("F5:F" & endRowLine), "男性", sheetKoetol.Range("G5:G" & endRowLine), "不明")
  
  sheetGraf.Range("F27") = WorksheetFunction.CountIfs(sheetKoetol.Range("F5:F" & endRowLine), "女性", sheetKoetol.Range("G5:G" & endRowLine), "10代")
  sheetGraf.Range("F28") = WorksheetFunction.CountIfs(sheetKoetol.Range("F5:F" & endRowLine), "女性", sheetKoetol.Range("G5:G" & endRowLine), "20代")
  sheetGraf.Range("F29") = WorksheetFunction.CountIfs(sheetKoetol.Range("F5:F" & endRowLine), "女性", sheetKoetol.Range("G5:G" & endRowLine), "30代")
  sheetGraf.Range("F30") = WorksheetFunction.CountIfs(sheetKoetol.Range("F5:F" & endRowLine), "女性", sheetKoetol.Range("G5:G" & endRowLine), "40代")
  sheetGraf.Range("F31") = WorksheetFunction.CountIfs(sheetKoetol.Range("F5:F" & endRowLine), "女性", sheetKoetol.Range("G5:G" & endRowLine), "50代")
  sheetGraf.Range("F32") = WorksheetFunction.CountIfs(sheetKoetol.Range("F5:F" & endRowLine), "女性", sheetKoetol.Range("G5:G" & endRowLine), "60代")
  sheetGraf.Range("F33") = WorksheetFunction.CountIfs(sheetKoetol.Range("F5:F" & endRowLine), "女性", sheetKoetol.Range("G5:G" & endRowLine), "不明")
  
  sheetGraf.Range("G27") = WorksheetFunction.CountIfs(sheetKoetol.Range("F5:F" & endRowLine), "不明", sheetKoetol.Range("G5:G" & endRowLine), "10代")
  sheetGraf.Range("G28") = WorksheetFunction.CountIfs(sheetKoetol.Range("F5:F" & endRowLine), "不明", sheetKoetol.Range("G5:G" & endRowLine), "20代")
  sheetGraf.Range("G29") = WorksheetFunction.CountIfs(sheetKoetol.Range("F5:F" & endRowLine), "不明", sheetKoetol.Range("G5:G" & endRowLine), "30代")
  sheetGraf.Range("G30") = WorksheetFunction.CountIfs(sheetKoetol.Range("F5:F" & endRowLine), "不明", sheetKoetol.Range("G5:G" & endRowLine), "40代")
  sheetGraf.Range("G31") = WorksheetFunction.CountIfs(sheetKoetol.Range("F5:F" & endRowLine), "不明", sheetKoetol.Range("G5:G" & endRowLine), "50代")
  sheetGraf.Range("G32") = WorksheetFunction.CountIfs(sheetKoetol.Range("F5:F" & endRowLine), "不明", sheetKoetol.Range("G5:G" & endRowLine), "60代")
  sheetGraf.Range("G33") = WorksheetFunction.CountIfs(sheetKoetol.Range("F5:F" & endRowLine), "不明", sheetKoetol.Range("G5:G" & endRowLine), "不明")
  
  For i = 27 To 33
    For j = 4 To 7
      If sheetGraf.Cells(i, j) = 0 Then
        sheetGraf.Cells(i, j) = ""
      End If
    Next
  Next
  
  '口コミ内容
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
          If sheetKoetol.Range("F" & i) = "男性" Then
            count01 = count01 + 1
          ElseIf sheetKoetol.Range("F" & i) = "女性" Then
            count02 = count02 + 1
          ElseIf sheetKoetol.Range("F" & i) = "不明" Then
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
  
  '年代別評価
  line = 85
  For Each colLineName In Split("10代,20代,30代,40代,50代,60代,不明", ",")
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


  '年代別口コミ内容
  Dim serachWord As String
  Dim count As Long
  
  line = 101
  
  For colLine = 20 To endColLine
    serachWord = sheetKoetol.Cells(3, colLine)
    sheetGraf.Range("C" & line) = serachWord
    count = 5
    
    For Each colLineName In Split("10代,20代,30代,40代,50代,60代,不明", ",")
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
  
  '終了処理----------------------------------------------------------------------------------------------------
  sheetGraf.Activate
  Application.Goto Reference:=Range("A1"), Scroll:=True
  
'  Call ProgressBar.ProgShowClose

End Function


'**************************************************************************************************
' * 文字列検索
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
  
  '既存データ削除
  sheetKoetol.Range("J5:AW" & endRowLine).ClearContents

  For Each colLineName In Split("H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z,AA,AB,AC,AD,AE,AF", ",")
  
    sheetKoetol.Range(colLineName & "1") = Replace(sheetKoetol.Range(colLineName & "1"), "　", " ")
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
' * 重複・指定キーワード削除用ユーザーフォームの表示
' *
'**************************************************************************************************
Function 指定語句除外フォーム表示()
  
  Dim dbCon As ADODB.Connection
  Dim DBRecordset As ADODB.Recordset
  Dim count As Integer
  Dim QueryString As String
  
  Call init.setting
  
  With keywordForm
    .StartUpPosition = 0
    .Top = Application.Top + (ActiveWindow.Width / 4)
    .Left = Application.Left + (ActiveWindow.Height / 2)
    .Caption = "KOETOL:重複・指定キーワード削除"
    .keywordList = Library.getRegistry("koetol_delKeyword")
  End With
  
  keywordForm.Show

End Function


'**************************************************************************************************
' * 重複・指定キーワード削除
' *
' * @author Bunpei.Koizumi<koizumi.bunpei@trans-cosmos.co.jp>
'**************************************************************************************************
Function 指定語句除外()

  Dim line As Long, endLine As Long
  Dim delWord As Variant

  Call init.setting
  Call Library.startScript
  Call ProgressBar.showStart
  
  sheetKoetol.Select
  endLine = sheetKoetol.Cells(Rows.count, 9).End(xlUp).Row
  
  '重複削除
  ActiveSheet.Range("$C$5:$AZ$" & endLine).RemoveDuplicates Columns:=7, Header:=xlNo
  '一旦背景色をグレーに変更
  Rows("5:" & Rows.count).Interior.Color = RGB(242, 242, 242)
  Call 書式設定
  
  '指定ワード削除
  For line = 5 To endLine
    For Each delWord In Split(delStringList, vbCrLf)
      Call ProgressBar.showCount("指定キーワード削除", line, endLine, CStr(delWord))
      
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
' * キーワード解析
' *
' * @author Bunpei.Koizumi<koizumi.bunpei@trans-cosmos.co.jp>
'**************************************************************************************************
Function キーワード解析()
  
  Dim endRowLine As Long, line As Long, count As Long
  Dim tmp As Variant
  Dim searchWord As Variant
  Dim result As Range
  Dim colLineName As Variant
  Dim firstAddress  As String
  
  Call init.setting
  count = 1
  endRowLine = sheetKoetol.Cells(Rows.count, 3).End(xlUp).Row
  
  If MsgBox("設定済みデータをクリアしますか？", vbYesNo + vbExclamation + vbDefaultButton2) = vbYes Then
    Range("J5:AW" & endRowLine).ClearContents
  End If
  
  Call ProgressBar.showStart
  For Each colLineName In Split("J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z,AA,AB,AC,AD,AE,AF,AG,AH,AI,AJ,AK,AL,AM,AN,AO,AP,AQ,AR,AS,AT,AU,AV,AW", ",")
    Call ProgressBar.showCount("キーワード解析", count, 40, colLineName & "列　処理中・・・")
    count = count + 1

  
    '区切り文字を統一
    sheetKoetol.Range(colLineName & "1") = Replace(sheetKoetol.Range(colLineName & "1"), "　", " ")
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
' * 書式設定
' *
' * @author Bunpei.Koizumi<koizumi.bunpei@trans-cosmos.co.jp>
'**************************************************************************************************
Function 書式設定()
  Dim endRowLine As Long, line As Long
  
  Call ProgressBar.showCount("書式設定", 0, 10, "")
  
  'セルの初期設定
  endRowLine = sheetKoetol.Cells(Rows.count, 3).End(xlUp).Row
  
  '背景色をなしにする
  sheetKoetol.Range("C5:AZ" & endRowLine).Interior.Pattern = xlNone
  
  '投稿日の書式設定
  Range("E5:E" & endRowLine).NumberFormatLocal = "yyyy/mm/dd"
    
  '行の高さ調整
  Rows("5:" & endRowLine).RowHeight = 30
  
  
  '4行目の合計値算出
  For Each colLineName In Split("J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z,AA,AB,AC,AD,AE,AF,AG,AH,AI,AJ,AK,AL,AM,AN,AO,AP,AQ,AR,AS,AT,AU,AV,AW", ",")
    Range(colLineName & "4") = "=SUM(" & colLineName & "5:" & colLineName & endRowLine & ")"
  Next colLineName
  
  Range("J4:AH4").NumberFormatLocal = "G/標準"
  
  
  'AG:AI列のカウント
  For line = 5 To endRowLine
    sheetKoetol.Range("AX" & line).FormulaR1C1 = "=COUNTIF(RC[-40]:RC[-31],1)"
    sheetKoetol.Range("AY" & line).FormulaR1C1 = "=COUNTIF(RC[-31]:RC[-17],1)"
    sheetKoetol.Range("AZ" & line).FormulaR1C1 = "=COUNTIF(RC[-17]:RC[-3],1)"
  Next line
  
  
  '手動入力部分の表示形式設定
  sheetKoetol.Range("J5" & ":AW" & endRowLine).Select
  With Selection
      .NumberFormatLocal = "[=1]""○"""
      .Validation.Delete
      .Validation.Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="1", Formula2:="1"
      .Validation.IgnoreBlank = True
      .Validation.InCellDropdown = True
      .Validation.IMEMode = xlIMEModeDisable
      .Validation.ShowInput = True
      .Validation.ShowError = True
      .Validation.ErrorTitle = ""
      .Validation.ErrorMessage = "1のみ入力可能です"
      .HorizontalAlignment = xlCenter
      .VerticalAlignment = xlCenter
  End With
   
    
  '罫線設定
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
' * ランダム抽出
' *
' * @author Bunpei.Koizumi<koizumi.bunpei@trans-cosmos.co.jp>
'**************************************************************************************************
Function ランダム抽出()

  Dim endRowLine As Long, line As Long, lineCount As Long, randomLine As Long
  Dim count As Long, minCnt As Long, maxCnt As Long, setRowLine As Long
  
  Call Library.startScript
  Call init.setting
  
  minCnt = 5
  maxCnt = 5
  line = sheetKoetol.Cells(Rows.count, 3).End(xlUp).Row
  
'Reランダム抽出:
  sheetKoetol.Select
  maxCnt = Library.getRegistry("koetol_DataCount")
  
  If (line - minCnt + 1) <= maxCnt Then
    Call Library.showNotice(410, , True)
  End If
  
  Do While WorksheetFunction.CountA(Range("A5:A" & line)) < maxCnt
    Call ProgressBar.showCount("ランダム抽出", line, maxCnt, "　処理中・・・")

    randomLine = Library.makeRandomNo(minCnt, line)
    Range("A" & randomLine) = 1
  Loop
  
  For count = minCnt To line
    Call ProgressBar.showCount("ランダム抽出", count, line, "　処理中・・・")
    
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
' * Chrome起動
' *
' * @author Bunpei.Koizumi<Koizumi.Bunpei@trans-cosmos.co.jp>
'*******************************************************************************************************
Function Chrome起動()
  
  'chrome 拡張機能読み込み
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
  Call Chrome起動
  
  'タイトル取得
  sheetKoetol.Range("E3") = driver.title

reCrawl:
  driver.Wait 1000
  line = sheetKoetol.Cells(Rows.count, 3).End(xlUp).Row + 1

  '口コミ情報取得
  Set elements = driver.FindElementById("cm_cr-review_list").FindElementsByClass("aok-relative")
  For i = 1 To elements.count
    Call ProgressBar.showCount("Webデータ取得", line - 4, CLng(setVal("DataCount") * setVal("DataRate")), "")
    
    getInnerHtmlString = ""
    
    'サイト------------------------------------------------------------------------------------------
    sheetKoetol.Range("C" & line) = line - 4
    sheetKoetol.Range("D" & line) = setVal("SiteName")
    
    '投稿日----------------------------------------------------------------------------------------
    getInnerHtmlString = elements.Item(i).FindElementByClass("review-date").Text
    getInnerHtmlString = Replace(getInnerHtmlString, "に日本でレビュー済み", "")
    getInnerHtmlString = Format(DateValue(getInnerHtmlString), "YYYY/MM/DD")
    sheetKoetol.Range("E" & line) = getInnerHtmlString
    sheetKoetol.Range("E" & line).NumberFormatLocal = "yyyy/mm/dd"
    
    '性別------------------------------------------------------------------------------------------
    sheetKoetol.Range("F" & line) = "不明"
    
    '年齢------------------------------------------------------------------------------------------
    sheetKoetol.Range("G" & line) = "不明"
    
    '評価------------------------------------------------------------------------------------------
    getInnerHtmlString = elements.Item(i).FindElementByClass("a-link-normal").Attribute("title")
    getInnerHtmlString = Replace(getInnerHtmlString, "5つ星のうち", "")
    sheetKoetol.Range("H" & line) = getInnerHtmlString
    sheetKoetol.Range("H" & line).NumberFormatLocal = "0_ ;[赤]-0 "
    
    'コメント--------------------------------------------------------------------------------------
    commentTitle = ""
    getInnerHtmlString = ""
    If elements.Item(i).IsElementPresent(myBy.Class("review-title")) Then
      commentTitle = setVal("commentTitleStart") & elements.Item(i).FindElementByClass("review-title").Text & setVal("commentTitleEnd")
      getInnerHtmlString = commentTitle & vbCrLf
      
    End If
    getInnerHtmlString = getInnerHtmlString & Library.delCellLinefeed(elements.Item(i).FindElementByClass("review-text").Text)
    sheetKoetol.Range("I" & line) = getInnerHtmlString
    
    'タイトルをボールドにする
    If commentTitle <> "" Then
      Call Library.getRGB(setVal("commentTitleColor"), Red, Green, Blue)
      With sheetKoetol.Range("I" & line).Characters(start:=1, Length:=Len(commentTitle)).Font
        .Bold = True
        .Color = RGB(Red, Green, Blue)
      End With
    End If
    
    '次ページ遷移判定------------------------------------------------------------------------------
    If (Cells(Rows.count, 5).End(xlUp).Row - 4) >= setVal("DataCount") * setVal("DataRate") Then
      exitFlg = True
      Exit For
    Else
      exitFlg = False
    End If
    
    line = line + 1
  Next
  
  '次ページ遷移------------------------------------------------------------------------------------
  If exitFlg = True Then
    Call ProgressBar.showCount("Webデータ取得", 9, 10, "終了")
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
  Call Chrome起動
  
  
  'タイトル取得
  sheetKoetol.Range("E3") = driver.title

  '最初の記事の続きを読むをクリックする
  Set elements = driver.FindElementsByClass("read-more")
  For i = 1 To elements.count
    getInnerHtmlString = elements.Item(i).Text
    If getInnerHtmlString Like "*[続きを読む]*" Then
      elements.Item(i).FindElementByTag("a").Click
      driver.Wait 1000
      Exit For
    End If
  Next

reCrawl:
  driver.Wait 1000
  line = sheetKoetol.Cells(Rows.count, 3).End(xlUp).Row + 1

  '口コミ情報取得
  DoEvents
  getInnerHtmlString = ""
  
  'サイト------------------------------------------------------------------------------------------
  sheetKoetol.Range("C" & line) = line - 4
  sheetKoetol.Range("D" & line) = setVal("SiteName")
  
  '投稿日----------------------------------------------------------------------------------------
  getInnerHtmlString = driver.FindElementByXPath("//*[@id='product-review-list']/div/div/div[2]/div[1]/p[2]").Text
'  getInnerHtmlString = Format(DateValue(getInnerHtmlString), "YYYY/MM/DD")
  sheetKoetol.Range("E" & line) = getInnerHtmlString
  sheetKoetol.Range("E" & line).NumberFormatLocal = "yyyy/mm/dd"
  
  '性別------------------------------------------------------------------------------------------
  sheetKoetol.Range("F" & line) = "不明"
  
  '年齢------------------------------------------------------------------------------------------
  getInnerHtmlString = driver.FindElementByXPath("//*[@id='product-review-list']/div/div/div[1]/div[2]/ul/li[1]").Text
  sheetKoetol.Range("G" & line) = getInnerHtmlString
  
  getInnerHtmlString = Replace(getInnerHtmlString, "歳", "")
  getInnerHtmlString = Application.WorksheetFunction.RoundDown(getInnerHtmlString, -1)
  init.sheetKoetol.Range("G" & line) = getInnerHtmlString & "代"
    
  '評価------------------------------------------------------------------------------------------
  getInnerHtmlString = driver.FindElementByClass("reviewer-rating").Text
  getInnerHtmlString = Replace(getInnerHtmlString, "購入品", "")
  getInnerHtmlString = Replace(getInnerHtmlString, "リピート", "")
  
  If getInnerHtmlString = "評価しない" Then
    sheetKoetol.Range("H" & line) = "不明"
  Else
    sheetKoetol.Range("H" & line) = getInnerHtmlString
    sheetKoetol.Range("H" & line).NumberFormatLocal = "0_ ;[赤]-0 "
  End If
  
  'コメント--------------------------------------------------------------------------------------
  Range("I" & line) = Library.delCellLinefeed(driver.FindElementByClass("read").Text)
  
  '次ページ遷移判定------------------------------------------------------------------------------
  If (Cells(Rows.count, 5).End(xlUp).Row - 4) >= setVal("DataCount") * setVal("DataRate") Then
    exitFlg = True
  Else
    exitFlg = False
  End If
  
  line = line + 1
  
  '次ページ遷移------------------------------------------------------------------------------------
  If exitFlg = True Then
    Call ProgressBar.showCount("Webデータ取得", 9, 10, "終了")
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
  Call Chrome起動
  
  'タイトル取得
  sheetKoetol.Range("E3") = driver.title

reCrawl:
  driver.Wait 1000
  line = sheetKoetol.Cells(Rows.count, 3).End(xlUp).Row + 1

  '口コミ情報取得
  Set elements = driver.FindElementByClass("revRvwUserSecCnt").FindElementsByClass("revRvwUserSec")
  For i = 1 To elements.count
    Call ProgressBar.showCount("Webデータ取得", line - 4, CLng(setVal("DataCount") * setVal("DataRate")), "")

    getInnerHtmlString = ""
    
    'サイト------------------------------------------------------------------------------------------
    sheetKoetol.Range("C" & line) = line - 4
    sheetKoetol.Range("D" & line) = setVal("SiteName")
    
    '投稿日----------------------------------------------------------------------------------------
    getInnerHtmlString = elements.Item(i).FindElementByClass("revUserEntryDate").Text
    getInnerHtmlString = Format(DateValue(getInnerHtmlString), "YYYY/MM/DD")
    sheetKoetol.Range("E" & line) = getInnerHtmlString
    sheetKoetol.Range("E" & line).NumberFormatLocal = "yyyy/mm/dd"
    
    '性別・年齢------------------------------------------------------------------------------------------
    getInnerHtmlString = elements.Item(i).FindElementByClass("revUserFaceDtlTxt").FindElementByTag("span").Text
    
    If getInnerHtmlString = "" Then
      sheetKoetol.Range("F" & line) = "不明"
      sheetKoetol.Range("G" & line) = "不明"
    Else
      getInnerHtmlString = Split(getInnerHtmlString, " ")
      sheetKoetol.Range("F" & line) = getInnerHtmlString(1)
      sheetKoetol.Range("G" & line) = getInnerHtmlString(0)
    End If
    
    '評価------------------------------------------------------------------------------------------
    getInnerHtmlString = elements.Item(i).FindElementByClass("revUserRvwerNum").Text
    sheetKoetol.Range("H" & line) = getInnerHtmlString
    sheetKoetol.Range("H" & line).NumberFormatLocal = "0_ ;[赤]-0 "

    'コメント--------------------------------------------------------------------------------------
    commentTitle = ""
    getInnerHtmlString = ""
    If elements.Item(i).IsElementPresent(myBy.Class("revRvwUserEntryTtl")) Then
      commentTitle = setVal("commentTitleStart") & elements.Item(i).FindElementByClass("revRvwUserEntryTtl").Text & setVal("commentTitleEnd")
      getInnerHtmlString = commentTitle & vbCrLf
    End If
    getInnerHtmlString = getInnerHtmlString & Library.delCellLinefeed(elements.Item(i).FindElementByClass("revRvwUserEntryCmt").Text)
    sheetKoetol.Range("I" & line) = getInnerHtmlString
    
    'タイトルをボールドにする
    If commentTitle <> "" Then
      Call Library.getRGB(setVal("commentTitleColor"), Red, Green, Blue)
      With sheetKoetol.Range("I" & line).Characters(start:=1, Length:=Len(commentTitle)).Font
        .Bold = True
        .Color = RGB(Red, Green, Blue)
      End With
    End If
              
    '次ページ遷移判定------------------------------------------------------------------------------
    If (Cells(Rows.count, 5).End(xlUp).Row - 4) >= setVal("DataCount") * setVal("DataRate") Then
      exitFlg = True
      Exit For
    Else
      exitFlg = False
    End If
    
    line = line + 1
  Next
  
  '次ページ遷移------------------------------------------------------------------------------------
  If exitFlg = True Then
    Call ProgressBar.showCount("Webデータ取得", 9, 10, "終了")
    driver.Get "file:///" & openingHTML("Koetol") & "\pause.html"

  Else
    getInnerHtmlString = driver.FindElementByClass("revThisPage").Text
    If getInnerHtmlString Like "*[次の15件]*" Then
      driver.FindElementByXPath("//*[@id='revRvwSec']/div[1]/div/div[3]/div[16]/div/div/a[5]").Click
      driver.Wait 1000
      
      GoTo reCrawl
    End If
  End If

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
  Call Chrome起動
  
  'タイトル取得
  sheetKoetol.Range("E3") = driver.title

reCrawl:
  driver.Wait 1000
  line = sheetKoetol.Cells(Rows.count, 3).End(xlUp).Row + 1

  '口コミ情報取得
  Set elements = driver.FindElementById("mainLeft").FindElementsByClass("reviewBox")
  For i = 1 To elements.count
    DoEvents
    getInnerHtmlString = ""
    
    'サイト------------------------------------------------------------------------------------------
    sheetKoetol.Range("C" & line) = line - 4
    sheetKoetol.Range("D" & line) = setVal("SiteName")
    
    '投稿日----------------------------------------------------------------------------------------
    getInnerHtmlString = elements.Item(i).FindElementByClass("entryDate").Text
    getInnerHtmlString = Split(getInnerHtmlString, " ")
    getInnerHtmlString = Format(DateValue(getInnerHtmlString(0)), "YYYY/MM/DD")
    sheetKoetol.Range("E" & line) = getInnerHtmlString
    sheetKoetol.Range("E" & line).NumberFormatLocal = "yyyy/mm/dd"
    
    '性別・年齢------------------------------------------------------------------------------------------
    sheetKoetol.Range("F" & line) = "不明"
    sheetKoetol.Range("G" & line) = "不明"
    
    '評価------------------------------------------------------------------------------------------
    getInnerHtmlString = elements.Item(i).FindElementByClass("total").Text
    sheetKoetol.Range("H" & line) = Replace(getInnerHtmlString, "満足度 ", "")
    sheetKoetol.Range("H" & line).NumberFormatLocal = "0_ ;[赤]-0 "

    'コメント--------------------------------------------------------------------------------------
    commentTitle = ""
    getInnerHtmlString = ""
    If elements.Item(i).IsElementPresent(myBy.Class("reviewTitle")) Then
      commentTitle = setVal("commentTitleStart") & elements.Item(i).FindElementByClass("reviewTitle").Text & setVal("commentTitleEnd")
      getInnerHtmlString = commentTitle & vbCrLf
      
    End If
    getInnerHtmlString = getInnerHtmlString & Library.delCellLinefeed(elements.Item(i).FindElementByClass("revEntryCont").Text)
    sheetKoetol.Range("I" & line) = getInnerHtmlString
    
    'タイトルをボールドにする
    If commentTitle <> "" Then
      Call Library.getRGB(setVal("commentTitleColor"), Red, Green, Blue)
      With sheetKoetol.Range("I" & line).Characters(start:=1, Length:=Len(commentTitle)).Font
        .Bold = True
        .Color = RGB(Red, Green, Blue)
      End With
    End If
    
    '次ページ遷移判定------------------------------------------------------------------------------
    If (Cells(Rows.count, 5).End(xlUp).Row - 4) >= setVal("DataCount") * setVal("DataRate") Then
      exitFlg = True
      Exit For
    Else
      exitFlg = False
    End If
    
    line = line + 1
  Next
  
  '次ページ遷移------------------------------------------------------------------------------------
  If exitFlg = True Then
    Call ProgressBar.showCount("Webデータ取得", 9, 10, "終了")
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
  Call Chrome起動
  
  'タイトル取得
  sheetKoetol.Range("E3") = driver.title

reCrawl:
  driver.Wait 1000
  line = sheetKoetol.Cells(Rows.count, 3).End(xlUp).Row + 1

  '口コミ情報取得
  Set elements = driver.FindElementByClass("review_list").FindElementsByClass("block")
  For i = 1 To elements.count
    DoEvents
    getInnerHtmlString = ""
    
    'サイト------------------------------------------------------------------------------------------
    sheetKoetol.Range("C" & line) = line - 4
    sheetKoetol.Range("D" & line) = setVal("SiteName")
    
    '投稿日----------------------------------------------------------------------------------------
    Set elements01 = elements.Item(i).FindElementsByClass("text_03")
    For j = 1 To elements01.count
      If elements01.Item(j).Text Like "投稿日：*" Then
        getInnerHtmlString = elements01.Item(j).Text
        sheetKoetol.Range("E" & line) = Replace(getInnerHtmlString, "投稿日：", "")
        sheetKoetol.Range("E" & line).NumberFormatLocal = "yyyy/mm/dd"
        Exit For
      End If
    Next
    
    '性別------------------------------------------------------------------------------------------
    sheetKoetol.Range("F" & line) = "不明"
    
    '年齢------------------------------------------------------------------------------------------
    getInnerHtmlString = elements.Item(i).FindElementByClass("text_02").Text
    Select Case getInnerHtmlString
      Case "年代：20歳未満"
        sheetKoetol.Range("G" & line) = "10代"
      Case "年代：20-25歳"
        sheetKoetol.Range("G" & line) = "20代"
      Case "年代：26-30歳"
        sheetKoetol.Range("G" & line) = "20代"
      Case "年代：31-35歳"
        sheetKoetol.Range("G" & line) = "30代"
      Case "年代：36-40歳"
        sheetKoetol.Range("G" & line) = "30代"
      Case "年代：41-50歳"
        sheetKoetol.Range("G" & line) = "40代"
      Case "年代：51-60歳"
        sheetKoetol.Range("G" & line) = "50代"
      Case "年代：61歳以上"
        sheetKoetol.Range("G" & line) = "60代"
      Case "年代：未登録"
        sheetKoetol.Range("G" & line) = "不明"
      Case Else
        sheetKoetol.Range("G" & line) = "不明"
    End Select
        
    '評価------------------------------------------------------------------------------------------
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

    'コメント--------------------------------------------------------------------------------------
    commentTitle = ""
    getInnerHtmlString = ""
    If elements.Item(i).IsElementPresent(myBy.Class("ttl")) Then
      commentTitle = setVal("commentTitleStart") & elements.Item(i).FindElementByClass("ttl").Text & setVal("commentTitleEnd")
      getInnerHtmlString = commentTitle & vbCrLf
      
    End If
    getInnerHtmlString = getInnerHtmlString & Library.delCellLinefeed(elements.Item(i).FindElementByClass("js-review_comment").Text)
    
    'カテゴリー------------------------------------------------------------------------------------
    Set elements01 = elements.Item(i).FindElementsByClass("text_03")
    For j = 1 To elements01.count
      If elements01.Item(j).Text Like "肌質：*" Then
        skinType = elements01.Item(j).Text
        
      ElseIf elements01.Item(j).Text Like "効果：*" Then
        effect = elements01.Item(j).Text
      End If
    Next j
    sheetKoetol.Range("I" & line) = getInnerHtmlString & vbCrLf & vbCrLf & skinType & vbCrLf & effect
    
    
    'タイトルをボールドにする
    If commentTitle <> "" Then
      Call Library.getRGB(setVal("commentTitleColor"), Red, Green, Blue)
      With sheetKoetol.Range("I" & line).Characters(start:=1, Length:=Len(commentTitle)).Font
        .Bold = True
        .Color = RGB(Red, Green, Blue)
      End With
    End If
     
    '次ページ遷移判定------------------------------------------------------------------------------
    If (Cells(Rows.count, 5).End(xlUp).Row - 4) >= setVal("DataCount") * setVal("DataRate") Then
      exitFlg = True
      Exit For
    Else
      exitFlg = False
    End If
    line = line + 1
  Next
  
  '次ページ遷移------------------------------------------------------------------------------------
  If exitFlg = True Then
    Call ProgressBar.showCount("Webデータ取得", 9, 10, "終了")
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



