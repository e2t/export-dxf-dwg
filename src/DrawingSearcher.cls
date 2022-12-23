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
Dim DocName As String

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
    
    DocName = gFSO.GetBaseName(Doc.GetPathName)
    ConfDesignation = Trim(GetProperty("Обозначение", DocExt, ConfName))
    Names(0) = Trim(GetProperty("Наименование", DocExt, ConfName))
    Names(1) = Trim(GetProperty("Наименование EN", DocExt, ConfName))
End Sub

Sub SearchFileNames()
    Dim CurFolder As Folder
    Dim I As Variant
    Dim F As File
    Dim Count As Integer
    Dim Name As String
    Dim Key As String
    
    Set CurFolder = gFSO.GetFolder(FolderPath)
    Count = 0
    For Each I In CurFolder.Files
        Set F = I
        If AddFileIfFit(F.Name) Then
            Count = Count + 1
        End If
    Next
    ' If no matching drawings are found, take the filenames from the properties.
    If Count = 0 Then
        For Each I In Names
            Name = I
            Key = CreateFileName(Name)
            If Key <> "" Then
                If Not FileNames.Exists(Key) Then
                    FileNames.Add Key, ""
                End If
            End If
        Next
        FileNames.Add DocName, ""
    End If
End Sub

Function CreateFileName(Name As String) As String
    CreateFileName = Trim(ConfDesignation + " " + Name)
End Function

Function AddFileIfFit(FileName As String) As Boolean
    Dim I As Variant
    Dim Name As String
        
    AddFileIfFit = True
    For Each I In Names
        Name = I
        If CheckAndAddFile(ConfDesignation, Name, FileName) Then
            Exit Function
        ElseIf CheckAndAddFile(GetBaseDesignation(ConfDesignation), Name, FileName) Then
            Exit Function
        End If
    Next
    AddFileIfFit = False
End Function

Private Function CheckAndAddFile(Designation As String, Name As String, FileName As String) As Boolean
    Dim NameF As String
    
    Regex.Pattern = "^" + RegEscape(Designation) + " (" + RegEscape(Name) + " *(\(.*\))?)" + RegEscape(DotSLDDRW)
    CheckAndAddFile = Regex.Test(FileName)
    If CheckAndAddFile Then
        NameF = Regex.Execute(FileName)(0).SubMatches(0)
        FileNames.Add CreateFileName(NameF), gFSO.BuildPath(FolderPath, FileName)
    End If
End Function

