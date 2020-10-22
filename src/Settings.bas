Attribute VB_Name = "Settings"
Option Explicit

Const macroName = "ExportDxfDwg"
Const macroSection = "Main"

Sub SaveStrSetting(key As String, value As String)
    SaveSetting macroName, macroSection, key, value
End Sub

Sub SaveIntSetting(key As String, value As Integer)
    SaveStrSetting key, str(value)
End Sub

Sub SaveBoolSetting(key As String, value As Boolean)
    SaveStrSetting key, BoolToStr(value)
End Sub

Function GetStrSetting(key As String, Optional default As String = "") As String
    GetStrSetting = GetSetting(macroName, macroSection, key, default)
End Function

Function GetBoolSetting(key As String) As Boolean
    GetBoolSetting = StrToBool(GetStrSetting(key, "0"))
End Function

Function GetIntSetting(key As String) As Integer
    GetIntSetting = StrToInt(GetStrSetting(key, "0"))
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
