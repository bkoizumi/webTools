Attribute VB_Name = "menu"



'******************************************************************************************************
' * ���̑�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'******************************************************************************************************
Function ���̑�_�w���v()

  If Worksheets("Help").Visible = 2 Then
    Worksheets("Help").Visible = True
    Worksheets("Help").Select
    Range("B3").Select
    
  ElseIf Worksheets("Help").Visible = True Then
    Worksheets("Help").Visible = xlSheetVeryHidden
  End If
End Function


Function ���̑�_�n�C���C�g()
  
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

Function ���̑�_�f�[�^�N���A()
  
  Call Library.startScript
  
  If MsgBox("���ׂẴV�[�g�̃f�[�^���폜���܂����H", vbYesNo + vbExclamation + vbDefaultButton2) = vbYes Then
    Call �����e�i���X.�S�f�[�^�N���A
  End If
  Call Library.endScript
  
End Function



'**************************************************************************************************
' * WebTools
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Sub �I�v�V�����\��()

  Call init.setting(True)
  Call WebCapture.showOptionForm
End Sub


'--------------------------------------------------------------------------------------------------
Sub WebCapture_�J�n()

  StartTime = Now

  If MsgBox("���X�g�����s���܂��B", vbYesNo + vbExclamation) = vbNo Then
    End
  End If
  Worksheets("WebCapture").Visible = True
  
  Call init.setting
  Call Library.startScript
  Call ProgressBar.showStart
  
  Call WebCapture.�ۑ��V�[�g���`�F�b�N
  Call WebCapture.�擾�J�n
  
  Worksheets("WebCapture").Visible = xlSheetVeryHidden
  
  StopTime = Now
  StopTime = StopTime - StartTime
  
  Application.Goto Reference:=Range("A1"), Scroll:=True
  Call Library.showNotice(200, "�L���v�`��")
  Call Shell("Explorer.exe /select, " & targetFilePath, vbNormalFocus)
  
  Call ProgressBar.showEnd
  Call Library.endScript
  
End Sub

'--------------------------------------------------------------------------------------------------
Sub �T�C�g�}�b�v_�J�n()

  Worksheets("�T�C�g�}�b�vtmp").Visible = True
  Call init.setting
  
  Call �T�C�g�}�b�v.�擾�J�n
  
  Worksheets("�T�C�g�}�b�vtmp").Visible = xlSheetVeryHidden
  
End Sub



