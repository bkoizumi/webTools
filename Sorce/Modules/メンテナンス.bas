Attribute VB_Name = "�����e�i���X"

Function �A�N�Z�X�œK��()

  Call init.setting
  Call ProgressBar.showStart
  Call Access.fileOptimisation
  Call ProgressBar.ProgShowClose
  MsgBox "����"

End Function


Function Slopy�f�[�^�擾(typeCode As String)

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
    Call Library.errorHandle(400, "�t�@�C�����I������܂���ł���")
    End
  End If
  
  Call Library.startScript
  Call init.setting
  Call ProgressBar.showStart
  
  'ID�̍ő�l���擾
  Set dbCon = New ADODB.Connection
  dbCon.Open init.ConnectionString
  Set DBRecordset = New ADODB.Recordset

  QueryString = "SELECT IIF(max(id) IS NULL, 0, max(id))  as maxid from �L���b�`�R�s�[;"
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
    Call ProgressBar.showCount("", targetbookLine, endLine, "�������E�E�E")
    
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

  '�f�[�^���`--------------------------------------------------------------------------------------
  endLine = Cells(Rows.count, 1).End(xlUp).Row
  For line = 1 To endLine
    Range("D" & line) = Replace(Range("D" & line), "�y��z", "")
'    Range("F" & line) = Replace(Range("F" & line), "�ދ`��_", "")
  Next
  
  Rows("1:" & endLine).RowHeight = 70
  Columns("B:B").ColumnWidth = 25
  Columns("C:E").ColumnWidth = 50

  Columns("C:C").WrapText = True
  
  Call ProgressBar.ProgShowClose
  Call Library.endScript

End Function


Function Slopy�f�[�^�o�^()
  Dim line As Long, endLine As Long
  
  Call init.setting
  endLine = Cells(Rows.count, 1).End(xlUp).Row
  
  init.maintenanceSheet.Select
  Call ProgressBar.showStart
  Call Access.insertTable("�L���b�`�R�s�[")
  
  Call Library.delSheetData
  Call ProgressBar.ProgShowClose
  
End Function


Function �S�f�[�^�N���A()

  Dim endLine As Long
  Call init.setting
  
  '�R�G�g���f�[�^�폜
  sheetKoetol.Select
  Call Koetol.SheetDataDelete
  sheetKoetol.Range("J4:AW4").ClearContents
  sheetKoetol.Range("J1:AW1").ClearContents
  Application.Goto Reference:=Range("C5"), Scroll:=True

  '�O���t�f�[�^�폜
  sheetGraf.Select
  sheetGraf.Range("E7:E9").ClearContents
  sheetGraf.Range("O7:O8").ClearContents
  sheetGraf.Range("D27:G33").ClearContents
  sheetGraf.Range("C42:G71").ClearContents
  sheetGraf.Range("D85:J91").ClearContents
  sheetGraf.Range("C101:K130").ClearContents
  Application.Goto Reference:=Range("A1"), Scroll:=True
  
  sheetjsonList.Select
  Call Json.�f�[�^�N���A
  Application.Goto Reference:=Range("A1"), Scroll:=True
  
  'Slopy�V�[�g
  sheetSlopy.Select
  endLine = Cells(Rows.count, 1).End(xlUp).Row + 3
  Rows("2:" & endLine).Select
  Selection.Delete Shift:=xlUp
  Application.Goto Reference:=Range("A1"), Scroll:=True
  
  '�`�ԑf���
  If sheetMecab.Visible = xlSheetVisible Then
    sheetMecabt.Select
    Call Library.delSheetData
    Application.Goto Reference:=Range("A1"), Scroll:=True
    
    sheetMecabResult.Select
    ActiveSheet.PivotTables("�`�ԑf���").PivotCache.Refresh
    Application.Goto Reference:=Range("A1"), Scroll:=True
  
    sheetMecabdic.Select
    endLine = Cells(Rows.count, 1).End(xlUp).Row + 3
    Rows("3:" & endLine).Delete Shift:=xlUp
    Application.Goto Reference:=Range("A1"), Scroll:=True
  End If
  
  '�����e�i���X�p�V�[�g
  sheetMaintenance.Select
  Call Library.delSheetData
  Cells.RowHeight = 20
  Application.Goto Reference:=Range("A1"), Scroll:=True


  ThisWorkbook.Worksheets("Index").Select


End Function

Function ���p�ҏ��擾()

  Dim dbCon As ADODB.Connection
  Dim DBRecordset As ADODB.Recordset
  Dim QueryString As String, category As String, searchWord As String, findString As String
  Dim line As Long, endLine As Long

  Call init.setting
  

  '�ϐ���`
  line = 2
  Call Library.startScript
  maintenanceSheet.Select
  Cells.RowHeight = 20
  
  '�f�[�^�N���A
  Call Library.delSheetData
  Application.Goto Reference:=Range("A1"), Scroll:=True

  Range("A1") = "�A�J�E���g"
  Range("B1") = "����"
  Range("C1") = "�������"

  Columns("A:A").ColumnWidth = 10
  Columns("B:B").ColumnWidth = 35
  Columns("C:C").ColumnWidth = 20

  Range("A2").Select
  ActiveWindow.FreezePanes = True
  
  Range("A1:C1").HorizontalAlignment = xlCenter
 
  Call ProgressBar.showStart
  
  'ADODB.Connection�������ADB�ɐڑ�
  Set dbCon = New ADODB.Connection
  dbCon.Open init.ConnectionString
  Set DBRecordset = New ADODB.Recordset

  QueryString = "SELECT * from ���p���O order by create_at desc;"

  DBRecordset.Open QueryString, dbCon, adOpenKeyset, adLockReadOnly
  Do Until DBRecordset.EOF

    Call ProgressBar.showCount("���p�ҏ��擾", DBRecordset.AbsolutePosition, DBRecordset.RecordCount, "�@�������E�E�E")
    
    Range("A" & line) = DBRecordset.Fields("account").Value
    Range("B" & line) = DBRecordset.Fields("action").Value
    Range("C" & line) = DBRecordset.Fields("create_at").Value
    
    line = line + 1
    DBRecordset.MoveNext
  Loop
  Range("C2:C" & line).NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
  
  
  'DB�N���[�Y
  dbCon.Close
  Set DBRecordset = Nothing
  
  Application.Goto Reference:=Range("A1"), Scroll:=True
  Call ProgressBar.ProgShowClose
  Call Library.endScript
End Function


'***************************************************************************************************************************************************
' * ���p���O�擾
' *
' * @author Bunpei.Koizumi<Koizumi.Bunpei@trans-cosmos.co.jp>
'***************************************************************************************************************************************************
Function ���p���O�擾(logInfo As String)

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
'    oRs.Open "insert into ���p���O  VALUES('" & MachineInfo("UserName") & "','" & logInfo & "',#" & Now() & "#);", oCn
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
' * ���p���O�N���A�擾
' *
' * @author Bunpei.Koizumi<Koizumi.Bunpei@trans-cosmos.co.jp>
'***************************************************************************************************************************************************
Function ���p���O�N���A�擾()

  Dim oCn As ADODB.Connection
  Dim oRs As ADODB.Recordset
  Dim intL As Long, intMaxF As Integer
  Dim endLine As Long

  Call init.setting
    
  Set oCn = CreateObject("ADODB.Connection")
  Set oRs = CreateObject("ADODB.Recordset")
  oCn.Open ConnectionString
  
  oRs.Open "delete from ���p���O;", oCn
  oCn.Close
  
  Set oRs = Nothing
  Set oCn = Nothing
  
  '�����e�i���X�p�V�[�g
  maintenanceSheet.Select
  Call Library.delSheetData
  Cells.RowHeight = 20
  Application.Goto Reference:=Range("A1"), Scroll:=True
  
  
  Exit Function




Run_Error:

End Function


'***************************************************************************************************************************************************
' * �ڎ�����
' *
' * @author Bunpei.Koizumi<Koizumi.Bunpei@trans-cosmos.co.jp>
'***************************************************************************************************************************************************
Function �ڎ�����()

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
      .Font.Name = "���S�V�b�N"
      .Font.Size = 11
    End With
    mline = mline + 1
    End If
  
  Next
  Call Library.endScript
  
  Exit Function
'---------------------------------------------------------------------------------------
'�G���[�������̏���
'---------------------------------------------------------------------------------------
catchError:

    Call Library.errorHandle(Err.Number, Err.Description)

End Function

