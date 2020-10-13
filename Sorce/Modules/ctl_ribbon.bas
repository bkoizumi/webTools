Attribute VB_Name = "ctl_ribbon"
Public ribbonUI As IRibbonUI ' ���{��


'�g�O���{�^��------------------------------------
Public RibbonToggleButton1 As Boolean
Public rbButton_Visible As Boolean
Public rbButton_Enabled As Boolean

'**************************************************************************************************
' * ���{�����j���[�ݒ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'�ǂݍ��ݎ�����------------------------------------------------------------------------------------
Function onLoad(ribbon As IRibbonUI)
  Set ribbonUI = ribbon
  ribbonUI.ActivateTab ("ExcelMethod")
  
  '���{���̕\�����X�V����
  ribbonUI.Invalidate
End Function

'�I���s�̐F�t�ؑփ{�^������------------------------------------------------------------------------
Function TButton1GetPressed(control As IRibbonControl, ByRef returnValue)
  
  Call init.setting(True)
  
  If setVal("ribbonHighLightFlg") = False Then
    returnedVal = True
  Else
    returnedVal = False
  End If
  
End Function


'Label�̓��I�\��-----------------------------------------------------------------------------------
Public Sub getLabel(control As IRibbonControl, ByRef setRibbonVal)
  setRibbonVal = getRibbonMenu(control.id, 2)
End Sub

Sub getonAction(control As IRibbonControl)
  Dim setRibbonVal As String

  setRibbonVal = getRibbonMenu(control.id, 3)
  Application.Run setRibbonVal

End Sub


'Supertip�̓��I�\��--------------------------------------------------------------------------------
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

'Ribbon�V�[�g������e���擾------------------------------------------------------------------------
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
'�G���[������=====================================================================================
catchError:
  getRibbonMenu = "�G���["

End Function



'**************************************************************************************************
' * ���̑�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �w���v(control As IRibbonControl)
  Call menu.���̑�_�w���v
End Function

Function �I���s�F�t�ؑ�(control As IRibbonControl)
  Call menu.���̑�_�n�C���C�g
End Function

Function �����e_�f�[�^�N���A(control As IRibbonControl)
  Call menu.���̑�_�f�[�^�N���A
End Function



'**************************************************************************************************
' * WebCapture
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'----------------------------------------------------------------------------------------------
Function WebCapture(control As IRibbonControl)
  Call menu.WebCapture_�J�n
End Function

'----------------------------------------------------------------------------------------------
Function Sitemap(control As IRibbonControl)
  Call menu.�T�C�g�}�b�v_�J�n
End Function




