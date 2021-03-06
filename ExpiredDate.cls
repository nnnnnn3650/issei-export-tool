VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExpiredDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Implements DaichoValues

Private Value As String
Const CharacterLimit As Integer = 15

Public Sub setValue(Value As String)
    this.Value = Value
End Sub

Public Function DaichoValues_isBlank() As Boolean
    If Value = "" Then
        MsgBox "値が入っていません"
        DaichoValues_isBlank = True
    Else
        DaichoValues_isBlank = False
    End If
End Function

Public Function DaichoValues_isLongerThanCharacterLimits() As Boolean
    If Len(Value) > CharacterLimit Then
        MsgBox "文字数が超過しています"
        DaichoValues_isLongerThanCharacterLimits = True
    Else
        DaichoValues_isLongerThanCharacterLimits = False
    End If
End Function

Public Function DaichoValues_HasCheckedValue() As Boolean
    Dim CheckValues As New Collection
    CheckValues.Add (DaichoValues_isBlank)
    CheckValues.Add (DaichoValues_isLongerThanCharacterLimits)
    For i = 1 To CheckValues.Count
        If CheckValues.item(i) Is True Then
            DaichoValues_HasCheckedValue = True
        End If
    Next i
End Function

