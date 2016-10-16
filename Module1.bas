Attribute VB_Name = "Module1"
Option Compare Database

Public Const NAME_ROW = "[名称]"
Public Const ABBREVIATION_LOW = "[略]"
Public Const EXPRIREDDATE_ROW = "[日程]"
Public Const MEMBER_LOW = "[対象一覧]"

Public Sub updateTables()
    
    Dim Filepath As String
    Dim i As Long
    Dim TableName As String
    Dim dirPath As String
    dirPath = CurrentProject.Path
    Filepath = Dir(dirPath & "\data\" & "*.csv")
    
    MsgBox "最新のCSV情報でテーブルを更新します"
    
    Do While Filepath <> ""
        TableName = GetTableName(Filepath)
        DoCmd.TransferText acImportDelim, , TableName, Filepath, True
        MsgBox TableName & "を更新しました"
        Filepath = Dir()
    Loop
    
End Sub

Private Function GetTableName(Filepath As String) As String
    Dim temp As Long
    Dim FileName As String
    '一番最後の\マークの位置を検索
    temp = InStrRev(Filepath, "\")
    '\マークより後ろの部分を切り取る
    FileName = Mid(Filepath, temp + 1)
    '拡張子を削除
    GetTableName = deleteExtension(FileName)
End Function

Function deleteExtension(FileNameWithExtension As String) As String
    Dim lFindPoint As Long
    '文字列の右端から"."を検索し、左端からの位置を取得する
    lFindPoint = InStrRev(FileNameWithExtension, ".")
    '拡張子を除いたファイル名の取得
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
    
    'ファイルの登録
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
        MsgBox searchWord & "がシート名：" & ws.Name & "で不足しています。担当営業に問い合わせてください"
    Else
        FoundCell.Select
        Set SearchCell = FoundCell.Offset(1, 0).Activate '一つ下に移動
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
        MsgBox searchWord & "がシート名：" & ws.Name & "で不足しています。担当営業に問い合わせてください"
    Else
        FoundCell.Select
        Set ActiveCell = FoundCell.Offset(1, 0).Activate '一つ下に移動
        Do Until ActiveCell.Value = ""
            SearchCollection.Add item:=ActivateCell.Value
            ActiveCell.Offset(1, 0).Select
        Loop
    End If
End Function

Public Sub ImportFile()
    'ファイルを登録する
    '１．SD名、SDコード、略称、表示順、第一日、第二日
    '２．SDコードと対応するSコード
    MsgBox "ImportFile"
End Sub

Public Sub ExportFiles()
    'テンプレートファイルコピー
    'コピーしたファイルを更新する
End Sub

