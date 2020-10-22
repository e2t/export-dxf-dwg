VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "Экпорт детали в DXF/DWG"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7905
   OleObjectBlob   =   "MainForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const OptionIsDXF = "IsDXF"

Private Sub CancelBtn_Click()
   ExitApp
End Sub

Private Sub DwgRad_Click()
   SaveBoolSetting OptionIsDXF, False
End Sub

Private Sub DxfRad_Click()
   SaveBoolSetting OptionIsDXF, True
End Sub

Private Sub RunBtn_Click()
   Dim FileBaseName As String
   Dim IsDxf As Boolean
   
   FileBaseName = Me.FilenameTxt.Text
   IsDxf = Me.DxfRad.value
   
   Me.Hide
   Run FileBaseName, IsDxf
   ExitApp
End Sub

Private Sub UserForm_Initialize()
   Me.FilenameTxt.Text = InitialFileBaseName
   If GetBoolSetting(OptionIsDXF) Then
      Me.DxfRad.value = True
   Else
      Me.DwgRad.value = True
   End If
End Sub
