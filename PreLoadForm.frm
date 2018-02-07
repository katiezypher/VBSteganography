VERSION 5.00
Begin VB.Form PreLoadForm 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
End
Attribute VB_Name = "PreLoadForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
'Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
'Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
'Private Type SECURITY_ATTRIBUTES
'    nLength As Long
'    lpSecurityDescriptor As Long
'    bInheritHandle As Long
'End Type

Private Sub Form_Load()
'    Dim Security As SECURITY_ATTRIBUTES
'    If Not PathFileExists("c:\tmp\RICHTX32.OCX") Then
'        CreateDirectory "c:\tmp", Security
'        CopyFile App.Path & "\data\RICHTX32.OCX", "c:\tmp\RICHTX32.OCX", False
'        CopyFile App.Path & "\data\readme.txt", "c:\tmp\readme.txt", False
'    End If
    PreLoadForm.Hide
    AppMainForm.Show
End Sub
