Attribute VB_Name = "Module1"
Option Compare Database

Public Const NAME_ROW = "[����]"
Public Const ABBREVIATION_LOW = "[��]"
Public Const EXPRIREDDATE_ROW = "[����]"
Public Const MEMBER_LOW = "[�Ώۈꗗ]"

Public Sub updateTables()
    
    Dim Filepath As String
    Dim i As Long
    Dim TableName As String
    Dim dirPath As String
    dirPath = CurrentProject.Path
    Filepath = Dir(dirPath & "\data\" & "*.csv")
    
    MsgBox "�ŐV��CSV���Ńe�[�u�����X�V���܂�"
    
    Do While Filepath <> ""
        TableName = GetTableName(Filepath)
        DoCmd.TransferText acImportDelim, , TableName, Filepath, True
        MsgBox TableName & "���X�V���܂���"
        Filepath = Dir()
    Loop
    
End Sub

Private Function GetTableName(Filepath As String) As String
    Dim temp As Long
    Dim FileName As String
    '��ԍŌ��\�}�[�N�̈ʒu������
    temp = InStrRev(Filepath, "\")
    '\�}�[�N�����̕�����؂���
    FileName = Mid(Filepath, temp + 1)
    '�g���q���폜
    GetTableName = deleteExtension(FileName)
End Function

Function deleteExtension(FileNameWithExtension As String) As String
    Dim lFindPoint As Long
    '������̉E�[����"."���������A���[����̈ʒu���擾����
    lFindPoint = InStrRev(FileNameWithExtension, ".")
    '�g���q���������t�@�C�����̎擾
    deleteExtension = Left(FileNameWithExtension, lFindPoint - 1)
End Function
Public Sub ImportNewDaicho(FileName As String)
    Dim app As Object
    Dim book As Object
    Dim sheet As Object
    Dim Filepath As String
    
    Filepath = Application.CurrentProject.Path & "\irai\" & FileName
    
    Set app = CreateObject("Excel.Application")
    Set book = app.WorkBooks.Open(Filepath)
   
    For i = 1 To book.sheets.Count
       Call CreateData(book.sheets(i))
    Next i
    
    '�t�@�C���̓o�^
    Call ImportFile
    
End Sub

Public Sub CreateData(sheet As Object)
    Dim Daicho As Daicho
    'Dim DaichoShomo As DaichoShomo
    Dim Name As String
    Dim ExpiredDate As String
    Dim List As Collection
    
    Daicho.SetName (SearchValue(sheet, NAME_ROW))
    Daicho.SetAbbreviation (SearchValue(sheet, ABBREVIATION_ROW))
    Daicho.SetExpiredDate (SearchValue(sheet, EXPRIREDDATE_ROW))
    Daicho.SetShomo (SearchCollection(sheet, MEMBER_LOW))
    Daicho.insertTables
End Sub
Public Function SearchValue(ws As Object, searchWord As String)
    Dim FoundCell As Range
    Dim SearchCell As Range
    Dim SearchValue As String
    Set FoundCell = Range("A:A").CurrentRegion.Find(What:=searchWord)
    If FoundCell Is Nothing Then
        MsgBox searchWord & "���V�[�g���F" & ws.Name & "�ŕs�����Ă��܂��B�S���c�Ƃɖ₢���킹�Ă�������"
    Else
        FoundCell.Select
        Set SearchCell = FoundCell.Offset(1, 0).Activate '����Ɉړ�
        Set SearchValue = SearchCell.Value
    End If
End Function

Public Function SearchCollection(ws As Object, searchWord As String)
    Dim FoundCell As Range
    Dim ActiveCell As Range
    Dim SearchValue As String
    Dim SearchCollection As New Collection
    Set FoundCell = Range("A:A").CurrentRegion.Find(What:=searchWord)
    If FoundCell Is Nothing Then
        MsgBox searchWord & "���V�[�g���F" & ws.Name & "�ŕs�����Ă��܂��B�S���c�Ƃɖ₢���킹�Ă�������"
    Else
        FoundCell.Select
        Set ActiveCell = FoundCell.Offset(1, 0).Activate '����Ɉړ�
        Do Until ActiveCell.Value = ""
            SearchCollection.Add item:=ActivateCell.Value
            ActiveCell.Offset(1, 0).Select
        Loop
    End If
End Function

Public Sub ImportFile()
    '�t�@�C����o�^����
    '�P�DSD���ASD�R�[�h�A���́A�\�����A�����A����
    '�Q�DSD�R�[�h�ƑΉ�����S�R�[�h
    MsgBox "ImportFile"
End Sub

Public Sub ExportFiles()
    '�e���v���[�g�t�@�C���R�s�[
    '�R�s�[�����t�@�C�����X�V����
End Sub

