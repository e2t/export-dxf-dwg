Attribute VB_Name = "Settings"
Option Explicit

Const macroName = "ExportDxfDwg"
Const macroSection = "Main"

Sub SaveStrSetting(Key As String, value As String)
    SaveSetting macroName, macroSection, Key, value
End Sub

Sub SaveIntSetting(Key As String, value As Integer)
    SaveStrSetting Key, str(value)
End Sub

Sub SaveBoolSetting(Key As String, value As Boolean)
    SaveStrSetting Key, BoolToStr(value)
End Sub

Function GetStrSetting(Key As String, Optional default As String = "") As String
    GetStrSetting = GetSetting(macroName, macroSection, Key, default)
End Function

Function GetBoolSetting(Key As String) As Boolean
    GetBoolSetting = StrToBool(GetStrSetting(Key, "0"))
End Function

Function GetIntSetting(Key As String) As Integer
    GetIntSetting = StrToInt(GetStrSetting(Key, "0"))
End Function

Function StrToInt(value As String) As Integer
    If IsNumeric(value) Then
        StrToInt = CInt(value)
    Else
        StrToInt = 0
    End If
End Function

Function StrToBool(value As String) As Boolean
    If IsNumeric(value) Then
        StrToBool = CInt(value)
    Else
        StrToBool = False
    End If
End Function

Function BoolToStr(value As Boolean) As String
    BoolToStr = str(CInt(value))
End Function
