Attribute VB_Name = "menu"



'***************************************************************************************************************************************************
' * その他
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***************************************************************************************************************************************************
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



'***********************************************************************************************************************************************
' * WebTools
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***********************************************************************************************************************************************
Sub WebCapture_開始()
  Dim StartTime, StopTime As Variant
  StartTime = Now

  Call init.setting
  
  
  sheetWebCaptureList.Select
  Application.Goto Reference:=Range("A1"), Scroll:=True
  
  If MsgBox("リストを実行します。", vbYesNo + vbExclamation) = vbNo Then
    End
  End If
  
  Call WebCapture.取得開始
  
  StopTime = Now
  StopTime = StopTime - StartTime
  
  sheetWebCaptureList.Range("G2") = WorksheetFunction.Text(StopTime, "[h]:mm:ss")
  MsgBox "処理完了：" & WorksheetFunction.Text(StopTime, "[h]:mm:ss")
  
End Sub

Sub サイトマップ_開始()

  Call init.setting
  Call サイトマップ.取得開始
  
End Sub



