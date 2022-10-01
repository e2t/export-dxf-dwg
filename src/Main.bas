Attribute VB_Name = "Main"
Option Explicit

#If VBA7 Then
     Private Declare PtrSafe Function SendMessage Lib "User32" Alias "SendMessageA" ( _
     ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
#Else
     Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" ( _
     ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
#End If

Public Const DotSLDDRW = ".SLDDRW"

Dim swApp As SldWorks.SldWorks
Public gFSO As FileSystemObject

Public FileNames As Dictionary
Public FolderPath As String

Dim CurrentDoc As ModelDoc2

Sub Main()
  Set swApp = Application.SldWorks
  Set gFSO = New FileSystemObject
  
  Set CurrentDoc = swApp.ActiveDoc
  If CurrentDoc Is Nothing Then
    Exit Sub
  End If
  If CurrentDoc.GetType <> swDocPART Then
    MsgBox "Только для деталей.", vbCritical
    Exit Sub
  End If
  
  Set FileNames = New Dictionary
  FolderPath = gFSO.GetParentFolderName(CurrentDoc.GetPathName)
  
  SearchFitDrawings
  InitForm
  MainForm.Show
End Sub

Function InitForm() 'hide
  Dim I As Variant
  
  With MainForm.FilenameCmb
    For Each I In FileNames.Keys
      .AddItem I
    Next
    If .ListCount > 0 Then
      .ListIndex = 0
      .SelStart = 0
      .SelLength = Len(.Text)
      .SetFocus
    End If
  End With
End Function

Function SearchFitDrawings() 'hide

  Dim Searcher As DrawingSearcher
  Dim CurFolder As Folder
  Dim I As Variant
  Dim F As File
   
  Set Searcher = New DrawingSearcher
  Searcher.Init CurrentDoc
  
  Set CurFolder = gFSO.GetFolder(FolderPath)
  For Each I In CurFolder.Files
    Set F = I
    Searcher.AddFileIfFit F.Name
  Next
  
End Function

Sub Run(UserFileName As String, IsDxf As Boolean, IsStep As Boolean)
  Dim DrawingPath As String
  Dim ChangeNumber As Integer
  Dim FileExt As String
  Dim NewName As String
  
  If FileNames.Exists(UserFileName) Then
    DrawingPath = gFSO.BuildPath(FolderPath, FileNames(UserFileName) + DotSLDDRW)
    ChangeNumber = GetChangeNumber(DrawingPath)
  Else
    ChangeNumber = 0
  End If
  
  If IsStep Then
    FileExt = ".STEP"
  ElseIf IsDxf Then
    FileExt = ".DXF"
  Else
    FileExt = ".DWG"
  End If
  
  If ChangeNumber = 0 Then
    NewName = gFSO.BuildPath(FolderPath, UserFileName + FileExt)
  Else
    NewName = gFSO.BuildPath(FolderPath, UserFileName + " (rev." + Format(ChangeNumber, "00") + ")" + FileExt)
  End If

  If IsStep Then
    SaveToSTEP CurrentDoc, NewName
  Else
    ExportFlatPattern CurrentDoc, NewName
  End If
End Sub

Function FindConfiguration(Doc As ModelDoc2) As String
   Const FlatPattern = "SM-FLAT-PATTERN"
   Dim CurrentConfName As String
   Dim BaseConfName As String
   
   CurrentConfName = Doc.ConfigurationManager.ActiveConfiguration.Name
   FindConfiguration = CurrentConfName
   
   If CurrentConfName Like "*" + FlatPattern Then
      BaseConfName = Left(CurrentConfName, Len(CurrentConfName) - Len(FlatPattern))
      If Not CurrentDoc.GetConfigurationByName(BaseConfName) Is Nothing Then
         FindConfiguration = BaseConfName
      End If
   End If
End Function

Sub ExportFlatPattern(Part As SldWorks.PartDoc, FileName As String)
                   
   Dim swEvListener As ExportEventsListener
   Set swEvListener = New ExportEventsListener
   
   'Set the file name for the exported DXF/DWG file
   Set swEvListener.Part = Part
   swEvListener.FilePath = FileName
   
   'Call the Export command
   Const WM_COMMAND As Long = &H111
   Const CMD_ExportFlatPattern As Long = 54244
   
   SendMessage swApp.Frame().GetHWnd(), WM_COMMAND, CMD_ExportFlatPattern, 0
   
   'wait for property page to be displayed
   Dim isActive As Boolean
   
   Do
      swApp.GetRunningCommandInfo -1, "", isActive
      DoEvents
   Loop While Not isActive
   
   Set swEvListener.Part = Nothing

   'TODO: call Windows API to set the required options in the property page
   
   'close property page
   'Const swCommands_PmOK As Long = -2
   'swApp.RunCommand swCommands_PmOK, ""
   
End Sub

Function GetChangeNumber(DrawingPath As String) As Integer
  Dim Model As ModelDoc2
  Dim errors As swFileLoadError_e
  Dim Errors2 As swActivateDocError_e
  Dim warnings As swFileLoadWarning_e
  Dim I As Variant
  Dim F As File
   
  GetChangeNumber = 0
  
  If gFSO.FileExists(DrawingPath) Then
    Set Model = swApp.OpenDoc6(DrawingPath, swDocDRAWING, _
       swOpenDocOptions_Silent + swOpenDocOptions_ViewOnly + swOpenDocOptions_ReadOnly, _
       "", errors, warnings)
    GetChangeNumber = ConvertStringToChangeNumber(GetProperty("Изменение", Model.Extension, ""))
    swApp.QuitDoc DrawingPath
    swApp.ActivateDoc3 CurrentDoc.GetPathName, False, swDontRebuildActiveDoc, Errors2
  End If
End Function

Function ConvertStringToChangeNumber(ChangeNumberProperty As String) As Integer
   ConvertStringToChangeNumber = 0
   On Error Resume Next
   ConvertStringToChangeNumber = CInt(ChangeNumberProperty)
End Function

Function GetProperty(propName As String, DocExt As ModelDocExtension, ConfName As String) As String
   Dim resultGetProp As swCustomInfoGetResult_e
   Dim rawProp As String, resolvedValue As String
   Dim wasResolved As Boolean
   
   resultGetProp = DocExt.CustomPropertyManager(ConfName).Get5(propName, True, rawProp, resolvedValue, wasResolved)
   If resultGetProp = swCustomInfoGetResult_NotPresent Then
      DocExt.CustomPropertyManager("").Get5 propName, True, rawProp, resolvedValue, wasResolved
   End If
   GetProperty = resolvedValue
End Function

Function ExitApp() 'mask
   Unload MainForm
   End
End Function

Function GetBaseDesignation(Designation As String) As String
    Dim lastFullstopPosition As Integer
    Dim firstHyphenPosition As Integer
    
    GetBaseDesignation = Designation
    lastFullstopPosition = InStrRev(Designation, ".")
    If lastFullstopPosition > 0 Then
        firstHyphenPosition = InStr(lastFullstopPosition, Designation, "-")
        If firstHyphenPosition > 0 Then
            GetBaseDesignation = Left(Designation, firstHyphenPosition - 1)
        End If
    End If
End Function

Sub SaveToSTEP(Doc As ModelDoc2, FileName As String)
    Dim errors As swFileSaveError_e
    Dim warnings As swFileSaveWarning_e
    
    Doc.Extension.SaveAs FileName, swSaveAsCurrentVersion, swSaveAsOptions_Silent, Nothing, errors, warnings
End Sub

'See: https://docs.microsoft.com/en-us/dotnet/api/System.Text.RegularExpressions.Regex.Escape
Function RegEscape(ByVal Line As String) As String

  Line = Replace(Line, "\", "\\")  'MUST be first!
  Line = Replace(Line, ".", "\.")
  Line = Replace(Line, "[", "\[")
  'line = Replace(line, "]", "\]")
  Line = Replace(Line, "|", "\|")
  Line = Replace(Line, "^", "\^")
  Line = Replace(Line, "$", "\$")
  Line = Replace(Line, "?", "\?")
  Line = Replace(Line, "+", "\+")
  Line = Replace(Line, "*", "\*")
  Line = Replace(Line, "{", "\{")
  'line = Replace(line, "}", "\}")
  Line = Replace(Line, "(", "\(")
  Line = Replace(Line, ")", "\)")
  Line = Replace(Line, "#", "\#")
  'and white space??
  RegEscape = Line
    
End Function
