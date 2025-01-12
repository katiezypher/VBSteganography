VERSION 5.00
Begin VB.Form HelpFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VBSteganography Help"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   8145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton OKBttn 
      Caption         =   "&OK"
      Height          =   495
      Left            =   3480
      TabIndex        =   1
      Top             =   6240
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   0
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "HelpFrm.frx":0000
      Top             =   0
      Width           =   8175
   End
End
Attribute VB_Name = "HelpFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub OKBttn_Click()
    HelpFrm.Visible = False
End Sub

Private Sub Text1_GotFocus()
    OKBttn.SetFocus
End Sub
