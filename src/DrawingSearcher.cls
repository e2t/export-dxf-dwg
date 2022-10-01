VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DrawingSearcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim Regex As RegExp
Dim ConfDesignation As String
Dim Names(1) As String

Dim FileName As String
Dim Name As String

Private Sub Class_Initialize()
  Set Regex = New RegExp
  Regex.IgnoreCase = True
  Regex.Global = True
End Sub

Sub Init(Doc As ModelDoc2)
  Dim DocExt As ModelDocExtension
  Dim ConfName As String
  
  Set DocExt = Doc.Extension
  ConfName = FindConfiguration(Doc)

  ConfDesignation = GetProperty("Обозначение", DocExt, ConfName)
  Names(0) = GetProperty("Наименование", DocExt, ConfName)
  Names(1) = GetProperty("Наименование EN", DocExt, ConfName)
End Sub

Sub AddFileIfFit(aFileName As String)
  Dim I As Variant
    
  FileName = aFileName
  
  For Each I In Names
    Name = I
    If CheckFile(ConfDesignation) Then
      Exit Sub
    ElseIf CheckFile(GetBaseDesignation(ConfDesignation)) Then
      Exit Sub
    End If
  Next
End Sub

Private Function CheckFile(Designation As String) As Boolean
  Dim NameF As String
  
  Regex.Pattern = RegEscape(Designation) + " (" + RegEscape(Name) + " *(\(.*\))?)" + RegEscape(DotSLDDRW)
  CheckFile = Regex.Test(FileName)
  If CheckFile Then
    NameF = Regex.Execute(FileName)(0).SubMatches(0)
    FileNames.Add ConfDesignation + " " + NameF, gFSO.GetBaseName(FileName)
  End If
End Function

