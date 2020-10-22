VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExportEventsListener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'**********************
'Copyright(C) 2020 Xarial Pty Limited
'Reference: https://www.codestack.net/solidworks-api/import-export/dwg-export-cleanup/
'License: https://www.codestack.net/license/
'**********************

Public WithEvents Part As SldWorks.PartDoc
Attribute Part.VB_VarHelpID = -1
Public FilePath As String

Private Function Part_FileSaveAsNotify2(ByVal FileName As String) As Long
    
    Dim swModel As SldWorks.ModelDoc2
    Set swModel = Part
    
    swModel.SetSaveAsFileName FilePath
    Part_FileSaveAsNotify2 = 1
    
End Function
