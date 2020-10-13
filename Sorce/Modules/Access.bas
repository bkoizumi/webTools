Attribute VB_Name = "Access"

'***********************************************************************************************************************************************
' * Accessへデータ登録
' *
' * @author Bunpei.Koizumi<koizumi.bunpei@trans-cosmos.co.jp>
'***********************************************************************************************************************************************
Function insertTable(tableName As String)

  Dim oCn As ADODB.Connection
  Dim oRs As ADODB.Recordset
  Dim intL As Long, intMaxF As Integer
  Dim endLine As Long
  
  'On Error GoTo catchError
  
  '最終行の取得
  endLine = Cells(Rows.count, 1).End(xlUp).Row
  
  Set oCn = CreateObject("ADODB.Connection")
  Set oRs = CreateObject("ADODB.Recordset")
  oCn.Open init.ConnectionString


  Call ProgressBar.showCount("Accessデータ登録", 1, 100, tableName & " データ登録中・・・")
  
  If tableName = "UserDic" Or tableName = "キャッチコピー" Then
  Else
    oRs.Open "DELETE FROM " & tableName & ";", oCn
  End If
  
  'オートナンバー型のリセット
'  QueryString = "select count(*) as rowcount from " & tableName & ";"
'  oRs.Open QueryString, oCn, adOpenKeyset, adLockReadOnly
'  DBRowCount = oRs.Fields("rowcount").Value
'  oRs.Close
'
'  oRs.Open "SELECT * FROM " & tableName & ";", oCn, adOpenDynamic, adLockOptimistic
'  oRs.AddNew
'  oRs![id] = DBRowCount
'  oRs.Update
'  oRs.MoveLast
'  oRs.delete
'  oRs.Close
  
  oRs.Open "SELECT * FROM " & tableName & ";", oCn, adOpenDynamic, adLockOptimistic
  
  intMaxF = Cells(1, 3000).End(xlToLeft).Column - 1
  intL = 1
  
  Do While (Cells(intL, 1).Value <> "")
    ProgressBar.showCount "Accessデータ登録", intL, endLine, tableName & " データ登録中・・・"
    oRs.AddNew
    
    For j = 0 To intMaxF
      If Cells(intL, j + 1).Value <> "" Then
        oRs.Fields(j) = Cells(intL, j + 1).Value
      End If
    Next j
    oRs.Update
    intL = intL + 1
  Loop
  oRs.Close
  oCn.Close
  
  Set oRs = Nothing
  Set oCn = Nothing
  
  Exit Function

'---------------------------------------------------------------------------------------------------------------------------------------------
'エラー発生時の処理
'---------------------------------------------------------------------------------------------------------------------------------------------
catchError:
  Dim Message As String
  
  Message = ""
  Message = Message & "エラー発生------------------------------------------------------------" & vbCrLf
  Message = Message & "    行　数：" & intL & vbCrLf
  Message = Message & "    項目名：" & Replace(Cells(1, j + 1).Value, vbLf, "") & vbCrLf
  Message = Message & "    入力値：" & Cells(intL, j + 1).Value & vbCrLf
  Message = Message & "    ------------------------------------------------------------------------" & vbCrLf

  Resume Next
  
End Function

'***********************************************************************************************************************************************
' * Accessファイル最適化
' *
' * @author Bunpei.Koizumi<koizumi.bunpei@trans-cosmos.co.jp>
'***********************************************************************************************************************************************
Function fileOptimisation()
  Dim res As Integer
  Dim AWP As String

  ' データベースを最適化する
  On Error GoTo DB_Err
  
  ProgressBar.showCount "Accessファイル最適化", 1, 100, "処理中・・・"

  AWP = init.accdbPath & "\"
  DBEngine.CompactDatabase AWP & init.accFileName, AWP & "最適化中.accdb", , , ""
  
  '前回のBackupファイルがあった場合前のファイルは削除する
  If Dir(AWP & "最適元Backup.accdb") <> "" Then
    Kill AWP & "最適元Backup.accdb"
  End If
  
  ProgressBar.showCount "Accessファイル最適化", 50, 100, "処理中・・・"
  
  '最適元はファイル名を変更してBackupとして保存
  Name AWP & init.accFileName As AWP & "最適元Backup.accdb"
  
  ProgressBar.showCount "Accessファイル最適化", 80, 100, "処理中・・・"

  '最適化後のファイルの名前を元に戻します。
  Name AWP & "最適化中.accdb" As AWP & init.accFileName
  
  If Dir(AWP & "最適元Backup.accdb") <> "" Then
    Kill AWP & "最適元Backup.accdb"
  End If
  
  Exit Function

DB_Err:
'  MsgBox Err.Number & Err.Description
End Function


'***********************************************************************************************************************************************
' * クエリ実行
' *
' * @Link https://antonsan.net/vt/excel-db/heading-1/page-003
'***********************************************************************************************************************************************
Function runAccessQuery(queryName As String)

  Dim oCn As ADODB.Connection
  Dim oRs As ADODB.Recordset
  Dim intL As Long, intMaxF As Integer
  Dim endLine As Long
  
  Set oCn = CreateObject("ADODB.Connection")
  Set oRs = CreateObject("ADODB.Recordset")
  oCn.Open init.ConnectionString
  
  If queryName = "コピーメカ用MaxID取得" Then

    oRs.Open "SELECT max(id) as maxid from キャッチコピー;", oCn
  
  Else
    myCon.Execute queryName
  End If

  oCn.Close
  
  Set oRs = Nothing
  Set oCn = Nothing
  
End Function
