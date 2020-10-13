Attribute VB_Name = "ctl_ribbon"
Public ribbonUI As IRibbonUI ' リボン


'トグルボタン------------------------------------
Public RibbonToggleButton1 As Boolean
Public rbButton_Visible As Boolean
Public rbButton_Enabled As Boolean

'**************************************************************************************************
' * リボンメニュー設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'読み込み時処理------------------------------------------------------------------------------------
Function onLoad(ribbon As IRibbonUI)
  Set ribbonUI = ribbon
  ribbonUI.ActivateTab ("ExcelMethod")
  
  'リボンの表示を更新する
  ribbonUI.Invalidate
End Function

'選択行の色付切替ボタン制御------------------------------------------------------------------------
Function TButton1GetPressed(control As IRibbonControl, ByRef returnValue)
  
  Call init.setting(True)
  
  If setVal("ribbonHighLightFlg") = False Then
    returnedVal = True
  Else
    returnedVal = False
  End If
  
End Function


'Labelの動的表示-----------------------------------------------------------------------------------
Public Sub getLabel(control As IRibbonControl, ByRef setRibbonVal)
  setRibbonVal = getRibbonMenu(control.id, 2)
End Sub

Sub getonAction(control As IRibbonControl)
  Dim setRibbonVal As String

  setRibbonVal = getRibbonMenu(control.id, 3)
  Application.Run setRibbonVal

End Sub


'Supertipの動的表示--------------------------------------------------------------------------------
Public Sub getSupertip(control As IRibbonControl, ByRef setRibbonVal)
  setRibbonVal = getRibbonMenu(control.id, 5)
End Sub

Public Sub getDescription(control As IRibbonControl, ByRef setRibbonVal)
  setRibbonVal = getRibbonMenu(control.id, 6)
End Sub

Public Sub getsize(control As IRibbonControl, ByRef setRibbonVal)
  Dim getVal As String
  getVal = getRibbonMenu(control.id, 4)

  Select Case getVal
    Case "large"
      setRibbonVal = 1
    Case "normal"
      setRibbonVal = 0
    Case Else
  End Select


End Sub

'Ribbonシートから内容を取得------------------------------------------------------------------------
Function getRibbonMenu(menuId As String, offsetVal As Long)

  Dim getString As String
  Dim FoundCell As Range
  Dim ribSheet As Worksheet
  Dim endLine As Long

  On Error GoTo catchError

  Call Library.startScript
  Set ribSheet = ThisWorkbook.Worksheets("Ribbon")

  endLine = ribSheet.Cells(Rows.count, 1).End(xlUp).Row

  getRibbonMenu = Application.VLookup(menuId, ribSheet.Range("A2:F" & endLine), offsetVal, False)
  Call Library.endScript


  Exit Function
'エラー発生時=====================================================================================
catchError:
  getRibbonMenu = "エラー"

End Function



'**************************************************************************************************
' * その他
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function ヘルプ(control As IRibbonControl)
  Call menu.その他_ヘルプ
End Function

Function 選択行色付切替(control As IRibbonControl)
  Call menu.その他_ハイライト
End Function

Function メンテ_データクリア(control As IRibbonControl)
  Call menu.その他_データクリア
End Function



'**************************************************************************************************
' * WebCapture
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'----------------------------------------------------------------------------------------------
Function WebCapture(control As IRibbonControl)
  Call menu.WebCapture_開始
End Function

'----------------------------------------------------------------------------------------------
Function Sitemap(control As IRibbonControl)
  Call menu.サイトマップ_開始
End Function




