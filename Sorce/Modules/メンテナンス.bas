Attribute VB_Name = "メンテナンス"

Function アクセス最適化()

  Call init.setting
  Call ProgressBar.showStart
  Call Access.fileOptimisation
  Call ProgressBar.ProgShowClose
  MsgBox "完了"

End Function


Function Slopyデータ取得(typeCode As String)

  Dim dbCon As ADODB.Connection
  Dim DBRecordset As ADODB.Recordset
  Dim line As Long, endLine As Long, targetbookLine As Long
  Dim filePath As String
  Dim targetWorkbook As Workbook, targetWorkSheet As Worksheet
  Dim maxID As Long
  Dim QueryString As String
 
  If filePath = "" Then
    filePath = Library.getFilePath(init.thisWorkbookPath, "", 1)
  Else
    filePath = Library.getFilePath(InStrRev(filePath, "\"), "", 1)
  End If
  
  If filePath = "" Then
    Call Library.errorHandle(400, "ファイルが選択されませんでした")
    End
  End If
  
  Call Library.startScript
  Call init.setting
  Call ProgressBar.showStart
  
  'IDの最大値を取得
  Set dbCon = New ADODB.Connection
  dbCon.Open init.ConnectionString
  Set DBRecordset = New ADODB.Recordset

  QueryString = "SELECT IIF(max(id) IS NULL, 0, max(id))  as maxid from キャッチコピー;"
  DBRecordset.Open QueryString, dbCon, adOpenKeyset, adLockReadOnly
  maxID = DBRecordset.Fields("maxid").Value
  dbCon.Close
  Set DBRecordset = Nothing
  
  
  init.maintenanceSheet.Select
  If typeCode = "new" Then
    line = 0
    Call Library.delSheetData
    Cells.RowHeight = 70
    ActiveWindow.FreezePanes = False
  Else
    maxID = ThisWorkbook.ActiveSheet.Range("A" & Cells(Rows.count, 1).End(xlUp).Row)
    line = Cells(Rows.count, 1).End(xlUp).Row
  End If
  
  
  Set targetWorkbook = Workbooks.Open(filePath)
  Set targetWorkSheet = targetWorkbook.Worksheets("sheet1")
  
  endLine = targetWorkSheet.Cells(Rows.count, 1).End(xlUp).Row
  For targetbookLine = 1 To endLine
    Call ProgressBar.showCount("", targetbookLine, endLine, "処理中・・・")
    
    Select Case targetWorkSheet.Range("A" & targetbookLine).Interior.Color
      Case 10498160
        line = line + 1
        maxID = maxID + 1
        ThisWorkbook.ActiveSheet.Range("A" & line) = maxID
        ThisWorkbook.ActiveSheet.Range("B" & line) = Replace(Dir(filePath), ".xlsx", "")
        ThisWorkbook.ActiveSheet.Range("C" & line) = targetWorkSheet.Range("A" & targetbookLine)
        
      Case 65535
        If ThisWorkbook.ActiveSheet.Range("D" & line) = "" Then
          ThisWorkbook.ActiveSheet.Range("D" & line) = targetWorkSheet.Range("A" & targetbookLine)
        Else
          ThisWorkbook.ActiveSheet.Range("D" & line) = ThisWorkbook.ActiveSheet.Range("D" & line) & targetWorkSheet.Range("A" & targetbookLine)
        End If
        
      Case 12611584
        If ThisWorkbook.ActiveSheet.Range("E" & line) = "" Then
          ThisWorkbook.ActiveSheet.Range("E" & line) = targetWorkSheet.Range("A" & targetbookLine)
        Else
          ThisWorkbook.ActiveSheet.Range("E" & line) = ThisWorkbook.ActiveSheet.Range("E" & line) & vbCrLf & targetWorkSheet.Range("A" & targetbookLine)
        End If
      
      Case 5287936
        If ThisWorkbook.ActiveSheet.Range("F" & line) = "" Then
          ThisWorkbook.ActiveSheet.Range("F" & line) = targetWorkSheet.Range("A" & targetbookLine)
        Else
          ThisWorkbook.ActiveSheet.Range("F" & line) = ThisWorkbook.ActiveSheet.Range("F" & line) & vbCrLf & targetWorkSheet.Range("A" & targetbookLine)
        End If
        
      Case Else
    End Select
  Next
  targetWorkbook.Close

  'データ整形--------------------------------------------------------------------------------------
  endLine = Cells(Rows.count, 1).End(xlUp).Row
  For line = 1 To endLine
    Range("D" & line) = Replace(Range("D" & line), "【例】", "")
'    Range("F" & line) = Replace(Range("F" & line), "類義語_", "")
  Next
  
  Rows("1:" & endLine).RowHeight = 70
  Columns("B:B").ColumnWidth = 25
  Columns("C:E").ColumnWidth = 50

  Columns("C:C").WrapText = True
  
  Call ProgressBar.ProgShowClose
  Call Library.endScript

End Function


Function Slopyデータ登録()
  Dim line As Long, endLine As Long
  
  Call init.setting
  endLine = Cells(Rows.count, 1).End(xlUp).Row
  
  init.maintenanceSheet.Select
  Call ProgressBar.showStart
  Call Access.insertTable("キャッチコピー")
  
  Call Library.delSheetData
  Call ProgressBar.ProgShowClose
  
End Function


Function 全データクリア()

  Dim endLine As Long
  Call init.setting
  
  'コエトルデータ削除
  sheetKoetol.Select
  Call Koetol.SheetDataDelete
  sheetKoetol.Range("J4:AW4").ClearContents
  sheetKoetol.Range("J1:AW1").ClearContents
  Application.Goto Reference:=Range("C5"), Scroll:=True

  'グラフデータ削除
  sheetGraf.Select
  sheetGraf.Range("E7:E9").ClearContents
  sheetGraf.Range("O7:O8").ClearContents
  sheetGraf.Range("D27:G33").ClearContents
  sheetGraf.Range("C42:G71").ClearContents
  sheetGraf.Range("D85:J91").ClearContents
  sheetGraf.Range("C101:K130").ClearContents
  Application.Goto Reference:=Range("A1"), Scroll:=True
  
  sheetjsonList.Select
  Call Json.データクリア
  Application.Goto Reference:=Range("A1"), Scroll:=True
  
  'Slopyシート
  sheetSlopy.Select
  endLine = Cells(Rows.count, 1).End(xlUp).Row + 3
  Rows("2:" & endLine).Select
  Selection.Delete Shift:=xlUp
  Application.Goto Reference:=Range("A1"), Scroll:=True
  
  '形態素解析
  If sheetMecab.Visible = xlSheetVisible Then
    sheetMecabt.Select
    Call Library.delSheetData
    Application.Goto Reference:=Range("A1"), Scroll:=True
    
    sheetMecabResult.Select
    ActiveSheet.PivotTables("形態素解析").PivotCache.Refresh
    Application.Goto Reference:=Range("A1"), Scroll:=True
  
    sheetMecabdic.Select
    endLine = Cells(Rows.count, 1).End(xlUp).Row + 3
    Rows("3:" & endLine).Delete Shift:=xlUp
    Application.Goto Reference:=Range("A1"), Scroll:=True
  End If
  
  'メンテナンス用シート
  sheetMaintenance.Select
  Call Library.delSheetData
  Cells.RowHeight = 20
  Application.Goto Reference:=Range("A1"), Scroll:=True


  ThisWorkbook.Worksheets("Index").Select


End Function

Function 利用者情報取得()

  Dim dbCon As ADODB.Connection
  Dim DBRecordset As ADODB.Recordset
  Dim QueryString As String, category As String, searchWord As String, findString As String
  Dim line As Long, endLine As Long

  Call init.setting
  

  '変数定義
  line = 2
  Call Library.startScript
  maintenanceSheet.Select
  Cells.RowHeight = 20
  
  'データクリア
  Call Library.delSheetData
  Application.Goto Reference:=Range("A1"), Scroll:=True

  Range("A1") = "アカウント"
  Range("B1") = "操作"
  Range("C1") = "操作日時"

  Columns("A:A").ColumnWidth = 10
  Columns("B:B").ColumnWidth = 35
  Columns("C:C").ColumnWidth = 20

  Range("A2").Select
  ActiveWindow.FreezePanes = True
  
  Range("A1:C1").HorizontalAlignment = xlCenter
 
  Call ProgressBar.showStart
  
  'ADODB.Connection生成し、DBに接続
  Set dbCon = New ADODB.Connection
  dbCon.Open init.ConnectionString
  Set DBRecordset = New ADODB.Recordset

  QueryString = "SELECT * from 利用ログ order by create_at desc;"

  DBRecordset.Open QueryString, dbCon, adOpenKeyset, adLockReadOnly
  Do Until DBRecordset.EOF

    Call ProgressBar.showCount("利用者情報取得", DBRecordset.AbsolutePosition, DBRecordset.RecordCount, "　処理中・・・")
    
    Range("A" & line) = DBRecordset.Fields("account").Value
    Range("B" & line) = DBRecordset.Fields("action").Value
    Range("C" & line) = DBRecordset.Fields("create_at").Value
    
    line = line + 1
    DBRecordset.MoveNext
  Loop
  Range("C2:C" & line).NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
  
  
  'DBクローズ
  dbCon.Close
  Set DBRecordset = Nothing
  
  Application.Goto Reference:=Range("A1"), Scroll:=True
  Call ProgressBar.ProgShowClose
  Call Library.endScript
End Function


'***************************************************************************************************************************************************
' * 利用ログ取得
' *
' * @author Bunpei.Koizumi<Koizumi.Bunpei@trans-cosmos.co.jp>
'***************************************************************************************************************************************************
Function 利用ログ取得(logInfo As String)

  Dim oCn As ADODB.Connection
  Dim oRs As ADODB.Recordset
  Dim intL As Long, intMaxF As Integer
  Dim endLine As Long

  On Error GoTo catchError:
'  Call init.setting

'  If Dir(init.accdbPath) <> "" Then
'    Call Library.getMachineInfo
'
'    Set oCn = CreateObject("ADODB.Connection")
'    Set oRs = CreateObject("ADODB.Recordset")
'    oCn.Open ConnectionString
'
'    oRs.Open "insert into 利用ログ  VALUES('" & MachineInfo("UserName") & "','" & logInfo & "',#" & Now() & "#);", oCn
'    oCn.Close
'
'    Set oRs = Nothing
'    Set oCn = Nothing
'
'  End If
  Exit Function

catchError:

End Function


'***************************************************************************************************************************************************
' * 利用ログクリア取得
' *
' * @author Bunpei.Koizumi<Koizumi.Bunpei@trans-cosmos.co.jp>
'***************************************************************************************************************************************************
Function 利用ログクリア取得()

  Dim oCn As ADODB.Connection
  Dim oRs As ADODB.Recordset
  Dim intL As Long, intMaxF As Integer
  Dim endLine As Long

  Call init.setting
    
  Set oCn = CreateObject("ADODB.Connection")
  Set oRs = CreateObject("ADODB.Recordset")
  oCn.Open ConnectionString
  
  oRs.Open "delete from 利用ログ;", oCn
  oCn.Close
  
  Set oRs = Nothing
  Set oCn = Nothing
  
  'メンテナンス用シート
  maintenanceSheet.Select
  Call Library.delSheetData
  Cells.RowHeight = 20
  Application.Goto Reference:=Range("A1"), Scroll:=True
  
  
  Exit Function




Run_Error:

End Function


'***************************************************************************************************************************************************
' * 目次生成
' *
' * @author Bunpei.Koizumi<Koizumi.Bunpei@trans-cosmos.co.jp>
'***************************************************************************************************************************************************
Function 目次生成()

  Dim line As Long, endLine As Long, mline As Long

  On Error GoTo catchError:
  endLine = Cells(Rows.count, 1).End(xlUp).Row
  
  mline = 3
  
  Call Library.startScript
  ThisWorkbook.Worksheets("Help").Select
  
  
  For line = 35 To endLine
    If Range("A" & line) <> "" Then
    
    With Range("B" & mline)
      .Value = Range("A" & line)
      .Select
      .Hyperlinks.Add anchor:=Selection, Address:="", SubAddress:="#" & "A" & line
      .Font.ColorIndex = 1
      .Font.Underline = xlUnderlineStyleNone
      .HorizontalAlignment = xlGeneral
      .VerticalAlignment = xlCenter
      .ShrinkToFit = True
      .Font.Name = "游ゴシック"
      .Font.Size = 11
    End With
    mline = mline + 1
    End If
  
  Next
  Call Library.endScript
  
  Exit Function
'---------------------------------------------------------------------------------------
'エラー発生時の処理
'---------------------------------------------------------------------------------------
catchError:

    Call Library.errorHandle(Err.Number, Err.Description)

End Function

