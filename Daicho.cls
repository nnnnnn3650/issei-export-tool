VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Daicho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private DaichoName As New DaichoName
Private Abbreviation As New Abbreviation
Private ExpiredDate As New ExpiredDate
Private DaichoMember As New Collection

Public Sub Class_Initialize()
        
End Sub

Public Sub SetName(Value As String)
    DaichoName.setValue (Value)
End Sub

Public Sub SetAbbreviation(Value As String)
    Set Abbreviation = Value
End Sub

Public Sub SetExpiredDate(Value As String)
    Set ExpiredDate = Value
End Sub

Public Sub SetShomo(List As Collection)
    Set Member = List
End Sub

Public Sub insertTables()
    Dim isContainsErrors As Boolean
    
    ' 値のチェック
    Set isContainsErrors = CheckValues
    If isContainsErrors Then Exit Sub
    End If
    
    'データをテーブルへ格納する
    InsertDaichoTable
    InsertDaichoMemberTable
    
End Sub

Public Function CheckValues()
    Dim Checker As DaichoChecke
    Dim ErrorCollection As New Collection
    Dim CheckValues As Boolean
    
    ErrorCollection.Add (Checker.ContainsValidationError(DaichoName))
    ErrorCollection.Add (Checker.ContainsValidationError(ExpiredDate))
    ErrorCollection.Add (Checker.ContainsValidationError(Abbreviation))
    
    CheckValues = isMemberOfCollection(ErrorCollection, True)
End Function

Function isMemberOfCollection(col As Collection, query) As Boolean
    Dim item
    
    For Each item In col
        If item = query Then
            isMemberOfCollection = True
            Exit Function
        End If
    Next
    
    isMemberOfCollection = False
End Function

Public Sub InsertDaichoTable()

End Sub

Public Sub InsertDaichoMemberTable()

End Sub
