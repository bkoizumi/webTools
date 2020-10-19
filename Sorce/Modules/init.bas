Attribute VB_Name = "init"
'ワークブック用変数------------------------------
Public ThisBook As Workbook
Public targetBook As Workbook


'ワークシート用変数------------------------------
Public sheetHelp As Worksheet
Public sheetNotice As Worksheet
Public sheetSetting As Worksheet

Public sheetWebCaptureList As Worksheet
Public sheetWebCapture As Worksheet
Public sheetSitemapTmp As Worksheet
Public sheetSitemap As Worksheet



'グローバル変数----------------------------------
Public Const thisAppName = "WebTools"
Public Const thisAppVersion = "0.0.2.0"

'レジストリ登録用サブキー
Public Const RegistryKey As String = "B.Koizumi"
Public Const RegistrySubKey As String = "WebTools"


Public ConnectionString As String

Public setVal As collection
Public getVal As collection
Public sitesInfo As Object


'パス関連
Public thisWorkbookPath As String
Public CurrentDirPath As String
Public binPath As String
Public logPath As String
Public AppWebCapturePath As String
Public AppSitemapPath As String
Public BrowserProfiles As collection
Public openingHTML As collection

Public targetFilePath As String
Public targetFileName As String

'Public saveDir As String

'ファイル関連
Public logFile As String

'その他
Public StartTime As Date
Public StopTime As Date


'**************************************************************************************************
' * 設定クリア
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function clearSetting()
  Set sheetHelp = Nothing
  Set sheetNotice = Nothing
  Set sheetSetting = Nothing
  
  Set sheetSitemap = Nothing
  Set sheetSitemapTmp = Nothing
  
  Set setVal = Nothing
  Set BrowserProfiles = Nothing
  Set openingHTML = Nothing

  logFile = ""
End Function

'**************************************************************************************************
' * 設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function setting(Optional reCheckFlg As Boolean)
  Dim line As Long
  Dim Message As String
  Dim varPath As String
  
'  On Error GoTo catchError
  ThisWorkbook.Save

  If logFile <> "" And reCheckFlg <> True Then
    Exit Function
  End If

'  Call Library.showDebugForm("setting", CStr(reCheckFlg))

  'ブックの設定
  Set ThisBook = ThisWorkbook
  ThisBook.Activate
  
  'ワークシート名の設定
  Set sheetHelp = ThisBook.Worksheets("Help")
  Set sheetNotice = ThisBook.Worksheets("Notice")
  Set sheetSetting = ThisBook.Worksheets("設定")
  
  Set sheetWebCaptureList = ThisBook.Worksheets("WebCaptureList")
  Set sheetWebCapture = ThisBook.Worksheets("WebCapture")
  Set sheetSitemap = ThisBook.Worksheets("サイトマップ")
  Set sheetSitemapTmp = ThisBook.Worksheets("サイトマップtmp")


  '設定値読み込み
  Set setVal = New collection
  Set sitesInfo = CreateObject("Scripting.Dictionary")
  
  With setVal
    For line = 3 To sheetSetting.Cells(Rows.count, 1).End(xlUp).Row
      If sheetSetting.Range("A" & line) <> "" Then
       .Add Item:=sheetSetting.Range("B" & line), Key:=sheetSetting.Range("A" & line)
      End If
    Next
  End With
  
  'レジストリからインストール情報取得
  With setVal
    .Add Item:=Library.getRegistry("InstDir"), Key:="appInstDir"
    .Add Item:=Library.getRegistry("InstVersion"), Key:="appVersion"
    .Add Item:=Library.getRegistry("InstNetwork"), Key:="InstNetwork"
  End With
  
  'ドライブパス関連
  thisWorkbookPath = ThisWorkbook.Path
  
  CurrentDirPath = setVal("appInstDir") & "\koetol"
  binPath = setVal("appInstDir") & "\bin"
  logPath = setVal("appInstDir") & "\logs"
  varPath = setVal("appInstDir") & "\var"
  
  logFile = logPath & "\ExcelMacro.log"
  
  AppWebCapturePath = varPath & "\WebCapture"
  AppSitemapPath = varPath & "\Sitemap"
  
  Set BrowserProfiles = New collection
  With BrowserProfiles
    .Add Item:=varPath & "\BrowserProfile\noScript", Key:="noScript"
    .Add Item:=varPath & "\BrowserProfile\default", Key:="default"
  End With
  
  Set openingHTML = New collection
  With openingHTML
    .Add Item:=varPath & "\Sitemap\opening", Key:="Sitemap"
    .Add Item:=varPath & "\WebCapture\opening", Key:="WebCapture"
  End With
  
  
  Call 名前定義
  
  Exit Function
  
'エラー発生時=====================================================================================
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
  
End Function


'**************************************************************************************************
' * 名前定義
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function 名前定義()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim Name As Object
  
'  On Error GoTo catchError

  '名前の定義を削除
  For Each Name In Names
    If Name.Visible = False Then
      Name.Visible = True
    End If
    If Not Name.Name Like "*!Print_Area" And Not Name.Name Like "*!Print_Titles" And Not Name.Name Like "スライサー*" Then
      Name.Delete
    End If
  Next
  
  'VBA用の設定
  For line = 3 To sheetSetting.Cells(Rows.count, 1).End(xlUp).Row
    If sheetSetting.Range("A" & line) <> "" Then
      sheetSetting.Range("B" & line).Name = sheetSetting.Range("A" & line)
    End If
  Next
  
  'Book用の設定
  For line = 3 To sheetSetting.Cells(Rows.count, 1).End(xlUp).Row
    If sheetSetting.Range("A" & line) <> "" Then
      sheetSetting.Range("B" & line).Name = sheetSetting.Range("A" & line)
    End If
  Next

  Exit Function
'エラー発生時=====================================================================================
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
  
End Function


'**************************************************************************************************
' * シートの表示/非表示
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function noDispSheet()

  Call init.setting
  Worksheets("Tmp").Visible = xlSheetVeryHidden
  Worksheets("Notice").Visible = xlSheetVeryHidden
  Worksheets("WebCapture").Visible = xlSheetVeryHidden
  Worksheets("サイトマップtmp").Visible = xlSheetVeryHidden
  Worksheets("サイトマップ").Visible = xlSheetVeryHidden
  
  Worksheets("WebCapture").Select
End Function



Function dispSheet()

  Call init.setting
  Worksheets("Notice").Visible = True
  Worksheets("設定").Visible = True
  Worksheets("WebCapture").Visible = True
  Worksheets("サイトマップtmp").Visible = True
  
  
  Worksheets("WebCapture").Select
  
End Function




































