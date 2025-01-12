VERSION 5.00
Begin VB.Form AboutFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About VBSteganography"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3735
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   240
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "AboutFrm.frx":0000
      Top             =   240
      Width           =   3255
   End
   Begin VB.CommandButton OKBttn 
      Caption         =   "OK"
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   1680
      Width           =   1215
   End
End
Attribute VB_Name = "AboutFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub OKBttn_Click()
    AboutFrm.Visible = False
End Sub

Private Sub Text1_GotFocus()
    OKBttn.SetFocus
End Sub
