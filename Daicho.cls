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

Private Name
Private Abbreviation
Private ExpriredDate
Private Member



Public Sub Class_Initialize()
        
End Sub

Public Sub SetName(Value As String)
    Set Name = Value
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

End Sub
