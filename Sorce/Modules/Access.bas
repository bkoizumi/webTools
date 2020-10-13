Attribute VB_Name = "Access"

'***********************************************************************************************************************************************
' * Access�փf�[�^�o�^
' *
' * @author Bunpei.Koizumi<koizumi.bunpei@trans-cosmos.co.jp>
'***********************************************************************************************************************************************
Function insertTable(tableName As String)

  Dim oCn As ADODB.Connection
  Dim oRs As ADODB.Recordset
  Dim intL As Long, intMaxF As Integer
  Dim endLine As Long
  
  'On Error GoTo catchError
  
  '�ŏI�s�̎擾
  endLine = Cells(Rows.count, 1).End(xlUp).Row
  
  Set oCn = CreateObject("ADODB.Connection")
  Set oRs = CreateObject("ADODB.Recordset")
  oCn.Open init.ConnectionString


  Call ProgressBar.showCount("Access�f�[�^�o�^", 1, 100, tableName & " �f�[�^�o�^���E�E�E")
  
  If tableName = "UserDic" Or tableName = "�L���b�`�R�s�[" Then
  Else
    oRs.Open "DELETE FROM " & tableName & ";", oCn
  End If
  
  '�I�[�g�i���o�[�^�̃��Z�b�g
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
    ProgressBar.showCount "Access�f�[�^�o�^", intL, endLine, tableName & " �f�[�^�o�^���E�E�E"
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
'�G���[�������̏���
'---------------------------------------------------------------------------------------------------------------------------------------------
catchError:
  Dim Message As String
  
  Message = ""
  Message = Message & "�G���[����------------------------------------------------------------" & vbCrLf
  Message = Message & "    �s�@���F" & intL & vbCrLf
  Message = Message & "    ���ږ��F" & Replace(Cells(1, j + 1).Value, vbLf, "") & vbCrLf
  Message = Message & "    ���͒l�F" & Cells(intL, j + 1).Value & vbCrLf
  Message = Message & "    ------------------------------------------------------------------------" & vbCrLf

  Resume Next
  
End Function

'***********************************************************************************************************************************************
' * Access�t�@�C���œK��
' *
' * @author Bunpei.Koizumi<koizumi.bunpei@trans-cosmos.co.jp>
'***********************************************************************************************************************************************
Function fileOptimisation()
  Dim res As Integer
  Dim AWP As String

  ' �f�[�^�x�[�X���œK������
  On Error GoTo DB_Err
  
  ProgressBar.showCount "Access�t�@�C���œK��", 1, 100, "�������E�E�E"

  AWP = init.accdbPath & "\"
  DBEngine.CompactDatabase AWP & init.accFileName, AWP & "�œK����.accdb", , , ""
  
  '�O���Backup�t�@�C�����������ꍇ�O�̃t�@�C���͍폜����
  If Dir(AWP & "�œK��Backup.accdb") <> "" Then
    Kill AWP & "�œK��Backup.accdb"
  End If
  
  ProgressBar.showCount "Access�t�@�C���œK��", 50, 100, "�������E�E�E"
  
  '�œK���̓t�@�C������ύX����Backup�Ƃ��ĕۑ�
  Name AWP & init.accFileName As AWP & "�œK��Backup.accdb"
  
  ProgressBar.showCount "Access�t�@�C���œK��", 80, 100, "�������E�E�E"

  '�œK����̃t�@�C���̖��O�����ɖ߂��܂��B
  Name AWP & "�œK����.accdb" As AWP & init.accFileName
  
  If Dir(AWP & "�œK��Backup.accdb") <> "" Then
    Kill AWP & "�œK��Backup.accdb"
  End If
  
  Exit Function

DB_Err:
'  MsgBox Err.Number & Err.Description
End Function


'***********************************************************************************************************************************************
' * �N�G�����s
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
  
  If queryName = "�R�s�[���J�pMaxID�擾" Then

    oRs.Open "SELECT max(id) as maxid from �L���b�`�R�s�[;", oCn
  
  Else
    myCon.Execute queryName
  End If

  oCn.Close
  
  Set oRs = Nothing
  Set oCn = Nothing
  
End Function
