Attribute VB_Name = "init"
'���[�N�u�b�N�p�ϐ�------------------------------
Public ThisBook   As Workbook
Public targetBook As Workbook


'���[�N�V�[�g�p�ϐ�------------------------------
Public sheetHelp            As Worksheet
Public sheetNotice          As Worksheet
Public sheetSetting         As Worksheet
Public sheetWebCaptureList  As Worksheet
Public sheetWebCapture      As Worksheet
Public sheetTmp             As Worksheet
Public sheetSitemap         As Worksheet
Public sheetLinkExtract       As Worksheet



'�O���[�o���ϐ�----------------------------------
Public Const thisAppName = "WebTools"
Public Const thisAppVersion = "0.0.4.1"

'���W�X�g���o�^�p�L�[
Public Const RegistryKey    As String = "WebTools"
Public Const RegistrySubKey As String = "Main"
Public RegistryRibbonName   As String


Public ConnectionString     As String

'�ݒ�l�ێ�
Public setVal       As Object
Public getVal       As Object
Public SitemapCell  As Object
Public sitesInfo    As Object


'�p�X�֘A
Public thisWorkbookPath   As String


Public AppWebCapturePath  As String
Public AppSitemapPath     As String

Public BrowserProfiles    As Object
Public openingHTML        As Object

Public targetFilePath     As String
Public targetFileName     As String

'Public saveDir As String

'�t�@�C���֘A
Public logFile            As String

'�������Ԍv���p
Public StartTime          As Date
Public StopTime           As Date

'Chrome�֘A
Public driver         As New Selenium.WebDriver
Public openingDriver  As New Selenium.WebDriver
Public targetURL      As String


'**************************************************************************************************
' * �ݒ�N���A
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function clearSetting()
  Set sheetHelp = Nothing
  Set sheetNotice = Nothing
  Set sheetSetting = Nothing
  
  Set sheetSitemap = Nothing
  Set sheetTmp = Nothing
  
  Set setVal = Nothing
  Set BrowserProfiles = Nothing
  Set openingHTML = Nothing

  logFile = ""
End Function

'**************************************************************************************************
' * �ݒ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function setting(Optional reCheckFlg As Boolean)
  Dim line As Long
  Dim Message As String
  
  
'  On Error GoTo catchError
  ThisWorkbook.Save

  If logFile <> "" And reCheckFlg <> True Then
    Exit Function
  End If

'  Call Library.showDebugForm("setting", CStr(reCheckFlg))

  '�u�b�N�̐ݒ�
  Set ThisBook = ThisWorkbook
  ThisBook.Activate
  
  '���[�N�V�[�g���̐ݒ�
  Set sheetHelp = ThisBook.Worksheets("Help")
  Set sheetNotice = ThisBook.Worksheets("Notice")
  Set sheetTmp = ThisBook.Worksheets("Tmp")
  Set sheetSetting = ThisBook.Worksheets("�ݒ�")
  
  Set sheetWebCaptureList = ThisBook.Worksheets("�L���v�`��")
  Set sheetWebCapture = ThisBook.Worksheets("WebCapture")
  Set sheetSitemap = ThisBook.Worksheets("�T�C�g�}�b�v")
  Set sheetLinkExtract = ThisBook.Worksheets("�����N���o")


  '�ݒ�l�ǂݍ���----------------------------------------------------------------------------------
  Set setVal = Nothing
  Set setVal = CreateObject("Scripting.Dictionary")
  With setVal
    For line = 3 To sheetSetting.Cells(Rows.count, 1).End(xlUp).Row
      If sheetSetting.Range("A" & line) <> "" Then
       .Add sheetSetting.Range("A" & line).Text, sheetSetting.Range("B" & line).Text
      End If
    Next
  End With
  
  '���W�X�g������C���X�g�[�����擾
  With setVal
    .Add "appInstDir", Library.getRegistry("InstDir")
    .Add "appVersion", Library.getRegistry("InstVersion")
    .Add "InstNetwork", Library.getRegistry("InstNetwork")
  End With
  RegistryRibbonName = "RP_" & ActiveWorkbook.Name
  
  
  '�h���C�u�p�X�֘A
  thisWorkbookPath = ThisWorkbook.Path
  
  With setVal
    .Add "binPath", setVal("appInstDir") & "\bin"
    .Add "logPath", setVal("appInstDir") & "\logs"
    .Add "varPath", setVal("appInstDir") & "\var"
    .Add "workPath", setVal("appInstDir") & "\work"
  End With
  
 
  logFile = setVal("logPath") & "\ExcelMacro.log"
  
  AppWebCapturePath = setVal("varPath") & "\WebCapture"
  AppSitemapPath = setVal("varPath") & "\Sitemap"
  
  Set BrowserProfiles = Nothing
  Set BrowserProfiles = CreateObject("Scripting.Dictionary")
  With BrowserProfiles
    .Add "noScript", setVal("varPath") & "\BrowserProfile\noScript"
    .Add "default", setVal("varPath") & "\BrowserProfile\default"
  End With
  
  Set openingHTML = Nothing
  Set openingHTML = CreateObject("Scripting.Dictionary")
  With openingHTML
    .Add "Sitemap", setVal("varPath") & "\Sitemap\opening"
    .Add "WebCapture", setVal("varPath") & "\WebCapture\opening"
  End With
  
  
  '�T�C�g�}�b�v�p�̃w�b�_�[�^�C�g���擾
  Set SitemapCell = Nothing
  Set SitemapCell = CreateObject("Scripting.Dictionary")
  
  With SitemapCell
    For line = 3 To sheetSetting.Cells(Rows.count, 4).End(xlUp).Row
      If sheetSetting.Range("D" & line) <> "" Then
       .Add sheetSetting.Range("D" & line).Text, sheetSetting.Range("E" & line).Text
      End If
    Next
  End With
  
  
  Call ���O��`
  
  Exit Function
  
'�G���[������====================================
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
  
End Function


'**************************************************************************************************
' * ���O��`
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function ���O��`()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim Name As Object
  
'  On Error GoTo catchError
  sheetSetting.Select
  
  '���O�̒�`���폜
  For Each Name In Names
    If Name.Visible = False Then
      Name.Visible = True
    End If
    If Not Name.Name Like "*!Print_Area" And Not Name.Name Like "*!Print_Titles" And Not Name.Name Like "�X���C�T�[*" Then
      Name.Delete
    End If
  Next
  
  'VBA�p�̐ݒ�
  For line = 3 To sheetSetting.Cells(Rows.count, 1).End(xlUp).Row
    If sheetSetting.Range("A" & line) <> "" Then
      sheetSetting.Range("B" & line).Name = sheetSetting.Range("A" & line)
    End If
  Next
  
  For line = 3 To sheetSetting.Cells(Rows.count, Library.getColumnNo(setVal("Cell_siteMapName"))).End(xlUp).Row
    If sheetSetting.Range(setVal("Cell_siteMapName") & line) <> "" Then
      sheetSetting.Range(setVal("Cell_siteMapVal") & line).Name = sheetSetting.Range("D" & line)
    End If
  Next
  
  
  'Book�p�̐ݒ�
  For colLine = Library.getColumnNo(setVal("Cell_ExcList")) To Cells(2, Columns.count).End(xlToLeft).Column
    If sheetSetting.Cells(2, colLine) <> "" Then
      endLine = Cells(Rows.count, colLine).End(xlUp).Row
      sheetSetting.Range(Cells(3, colLine), Cells(endLine, colLine)).Select
      sheetSetting.Range(Cells(3, colLine), Cells(endLine, colLine)).Name = sheetSetting.Cells(2, colLine)
    End If
  Next

  Application.Goto Reference:=Range("A1"), Scroll:=True


  Exit Function
'�G���[������====================================
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
  
End Function


'**************************************************************************************************
' * �V�[�g�̕\��/��\��
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function noDispSheet()

  Call init.setting
  sheetHelp.Visible = xlSheetVeryHidden
  sheetWebCapture.Visible = xlSheetVeryHidden
  
  
  sheetWebCaptureList.Select
End Function



Function dispSheet()

  Call init.setting
  
  sheetHelp.Visible = True
  sheetNotice.Visible = True
  sheetTmp.Visible = True
  sheetSetting.Visible = True
  sheetWebCaptureList.Visible = True
  sheetWebCapture.Visible = True
  sheetSitemap.Visible = True
  
  
End Function


'**************************************************************************************************
' * ���ڃ`�F�b�N
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function ���ڗ�`�F�b�N()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim itemName As String
  Dim defaultLine As Long
  
  
  Call init.setting

  sheetSetting.Range("D3:E100").ClearContents
  
  For colLine = 1 To sheetSitemap.Cells(2, Columns.count).End(xlToLeft).Column
    If sheetSitemap.Cells(2, colLine) <> "" Then
      itemName = sheetSitemap.Cells(2, colLine)
    Else
      GoTo Label_nextFor
    End If
    
    line = sheetSetting.Cells(Rows.count, 4).End(xlUp).Row + 1
    If line < defaultLine Then
      line = defaultLine
    End If
    
    
    
    Select Case True
      Case itemName = "#"
        sheetSetting.Range("D" & line) = "No"
      Case itemName = "�Ǘ�No"
        sheetSetting.Range("D" & line) = "ctlNo"
        
      Case itemName = "�K�w"
        sheetSetting.Range("D" & line) = "level"
      
      Case itemName = "�e�X�gURL"
        sheetSetting.Range("D" & line) = "testURL"
      
      Case itemName = "�v��URL"
        sheetSetting.Range("D" & line) = "preURL"
        
      Case itemName = "�{��URL"
        sheetSetting.Range("D" & line) = "proURL"
        
      Case itemName Like "L*"
        itemName = Replace(itemName, "L", "")
        sheetSetting.Range("D" & line) = "dirLevel_" & itemName
        sheetSetting.Range("maxDirLevel") = itemName
        
      Case itemName Like "T*"
        itemName = Replace(itemName, "T", "")
        sheetSetting.Range("D" & line) = "level_" & itemName
        sheetSetting.Range("maxTitleLevel") = itemName
        
      Case itemName Like "new_L*"
        itemName = Replace(itemName, "new_L", "")
        sheetSetting.Range("D" & line) = "newDirLevel_" & itemName
        sheetSetting.Range("maxTitleLevel") = itemName
        
      Case itemName = "�^�C�g��"
        sheetSetting.Range("D" & line) = "title"
      Case itemName = "��URL"
        sheetSetting.Range("D" & line) = "url"
      Case itemName = "description"
        sheetSetting.Range("D" & line) = "description"
      Case itemName = "keywords"
        sheetSetting.Range("D" & line) = "keywords"
      Case itemName = "Google"
        sheetSetting.Range("D" & line) = "google"
      Case itemName = "Yahoo!"
        sheetSetting.Range("D" & line) = "Yahoo"
      
      Case itemName = "H1"
        sheetSetting.Range("D" & line) = "htmlTag_H1"
        
        
      Case Else
    End Select
    If sheetSetting.Range("D" & line) <> "" Then
      sheetSetting.Range("E" & line) = Library.getColumnName(colLine)
    End If
    
Label_nextFor:
  Next

  Call setting(True)
End Function

