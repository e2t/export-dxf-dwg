Attribute VB_Name = "Main"
Option Explicit

#If VBA7 Then
     Private Declare PtrSafe Function SendMessage Lib "User32" Alias "SendMessageA" ( _
     ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
#Else
     Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" ( _
     ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
#End If

Dim swApp As SldWorks.SldWorks
Public gFSO As FileSystemObject

Public InitialFileBaseName As String
Private CurrentDoc As ModelDoc2

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
   
   InitialFileBaseName = CreateNewName(CurrentDoc.Extension, FindConfiguration(CurrentDoc))
   MainForm.Show
End Sub

Sub Run(FileBaseName As String, IsDxf As Boolean)
   Dim FileName As String
   Dim FolderPath As String
   Dim FileExt As String
   
   FolderPath = gFSO.GetParentFolderName(CurrentDoc.GetPathName)
   If IsDxf Then
      FileExt = ".DXF"
   Else
      FileExt = ".DWG"
   End If
   FileName = FolderPath + "\" + FileBaseName + FileExt
   
   ExportFlatPattern CurrentDoc, FileName
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

Function CreateNewName(docext As ModelDocExtension, confName As String) As String
    CreateNewName = GetProperty("Обозначение", docext, confName) + " " + _
                    GetProperty("Наименование", docext, confName)
End Function

Function GetProperty(propName As String, docext As ModelDocExtension, confName As String) As String
    Dim resultGetProp As swCustomInfoGetResult_e
    Dim rawProp As String, resolvedValue As String
    Dim wasResolved As Boolean
    
    resultGetProp = docext.CustomPropertyManager(confName).Get5(propName, True, rawProp, resolvedValue, wasResolved)
    If resultGetProp = swCustomInfoGetResult_NotPresent Then
        docext.CustomPropertyManager("").Get5 propName, True, rawProp, resolvedValue, wasResolved
    End If
    GetProperty = resolvedValue
End Function

Function ExitApp() 'mask
   Unload MainForm
   End
End Function
