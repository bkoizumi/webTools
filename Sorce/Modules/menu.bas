Attribute VB_Name = "menu"



'******************************************************************************************************
' * その他
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'******************************************************************************************************
Function その他_ヘルプ()

  If Worksheets("Help").Visible = 2 Then
    Worksheets("Help").Visible = True
    Worksheets("Help").Select
    Range("B3").Select
    
  ElseIf Worksheets("Help").Visible = True Then
    Worksheets("Help").Visible = xlSheetVeryHidden
  End If
End Function


Function その他_ハイライト()
  
  Dim endRowLine As Long
  Dim line As Long
  Dim SetActiveSheet As String
  
  Call init.setting
  Call Library.startScript
  
  SetActiveCell = Selection.Address
  SetActiveSheet = ActiveSheet.Name
  
  If setVal("ribbonHighLightFlg") = True Then
    sheetKoetol.Select
    endRowLine = sheetKoetol.Cells(Rows.count, 3).End(xlUp).Row
    Call Library.setLineColor("C5:I" & endRowLine, False, RGB(255, 242, 204))
    Call Library.setLineColor("J3:AZ" & endRowLine, True, RGB(255, 242, 204))
  
    Worksheets("Slopy").Select
    endRowLine = Worksheets("Slopy").Cells(Rows.count, 3).End(xlUp).Row
    Call Library.setLineColor("A2:E" & endRowLine, False, RGB(255, 242, 204))
  Else
    sheetKoetol.Select
    endRowLine = sheetKoetol.Cells(Rows.count, 3).End(xlUp).Row
    Call Library.unsetLineColor("C5:I" & endRowLine)
    Call Library.unsetLineColor("J3:AZ" & endRowLine)
  
    Worksheets("Slopy").Select
    endRowLine = Worksheets("Slopy").Cells(Rows.count, 3).End(xlUp).Row
    Call Library.unsetLineColor("A2:E" & endRowLine)
  End If
  
  Worksheets(SetActiveSheet).Select

  Call Library.endScript(True)
  
End Function

Function その他_データクリア()
  
  Call Library.startScript
  
  If MsgBox("すべてのシートのデータを削除しますか？", vbYesNo + vbExclamation + vbDefaultButton2) = vbYes Then
    Call メンテナンス.全データクリア
  End If
  Call Library.endScript
  
End Function



'**************************************************************************************************
' * WebTools
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Sub オプション表示()
  Call Library.startScript
  Call init.setting(True)
  Call Library.endScript(True)
  Call WebToolLib.showOptionForm
End Sub


'--------------------------------------------------------------------------------------------------
Sub WebCapture_開始()

  StartTime = Now

  Call init.setting
  sheetWebCaptureList.Select
  Call Library.startScript
  
'  If MsgBox("リストを実行します。", vbYesNo + vbExclamation) = vbNo Then
'    End
'  End If
  Worksheets("WebCapture").Visible = True
  
  Call ProgressBar.showStart
  
  Call キャプチャ.保存シート名チェック
  Call キャプチャ.取得開始
  
  Worksheets("WebCapture").Visible = xlSheetVeryHidden
  
  StopTime = Now
  StopTime = StopTime - StartTime
  
  sheetWebCaptureList.Select
  Application.Goto Reference:=Range("A1"), Scroll:=True
  Call Library.showNotice(200, "キャプチャ")
  Call Shell("Explorer.exe /select, " & targetFilePath, vbNormalFocus)
  
  Call ProgressBar.showEnd
  Call Library.endScript
  
End Sub

'--------------------------------------------------------------------------------------------------
Sub サイトマップ_開始()

  Call init.setting
  Call Library.startScript
  
  Call init.項目列チェック
  
  Call サイトマップ.取得開始
  
  Call Library.endScript
  sheetSitemap.Select
  Application.Goto Reference:=Range("A1"), Scroll:=True

End Sub



