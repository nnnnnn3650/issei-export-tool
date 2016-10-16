VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DaichoValues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Value As String
Private Error As Collection

Public Function isBlank()
    Dim isBlank As Boolean
    If Value = "" Then
        MsgBox "’l‚ª“ü‚Á‚Ä‚¢‚Ü‚¹‚ñ"
        isBlank = True
    Else
        isBlank = False
    End If
End Function

Public Function isLongerThanCharacterLimits()
    Dim isLongerThanCharacterLimits As Boolean
    If Len(Value) > CharacterLimit Then
        MsgBox "•¶š”‚ª’´‰ß‚µ‚Ä‚¢‚Ü‚·"
        isLongerThanCharacterLimits = True
    Else
        isLongerThanCharacterLimits = False
    End If
End Function

Public Function HasCheckedValue()
    
End Function
