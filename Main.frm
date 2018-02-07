VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form AppMainForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Red Dog 1.1"
   ClientHeight    =   10995
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10995
   ScaleWidth      =   7980
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox TempBox 
      Height          =   735
      Left            =   9480
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1296
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"Main.frx":0000
   End
   Begin RichTextLib.RichTextBox EncryptedBox 
      Height          =   1815
      Left            =   0
      TabIndex        =   20
      Top             =   5520
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   3201
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"Main.frx":0082
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox KeyBox 
      Height          =   285
      Left            =   2760
      TabIndex        =   19
      Top             =   465
      Width           =   5175
   End
   Begin VB.CommandButton OpenBttn 
      Caption         =   "&Open Text File"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      ToolTipText     =   "Open a text file for encryption."
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton EncryptBttn 
      Caption         =   "E&ncrypt"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      ToolTipText     =   "Encrypt the text file."
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton DecryptBttn 
      Caption         =   "&Decrypt"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      ToolTipText     =   "Decrypt the text file."
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton SaveBttn 
      Caption         =   "&Save Text"
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      ToolTipText     =   "Save the encrypted text file."
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton ClearBttn 
      Caption         =   "&Clear"
      Height          =   375
      Left            =   5280
      TabIndex        =   5
      ToolTipText     =   "Clear all textual contents."
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Quit"
      Height          =   375
      Left            =   6600
      TabIndex        =   6
      ToolTipText     =   "Exit the program."
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton OpenPictureBttn 
      Caption         =   "O&pen Picture"
      Height          =   375
      Left            =   0
      TabIndex        =   7
      ToolTipText     =   "Open a graphic file to embed message into."
      Top             =   7440
      Width           =   1335
   End
   Begin VB.CommandButton EmbedBttn 
      Caption         =   "Em&bed"
      Height          =   375
      Left            =   1320
      TabIndex        =   8
      ToolTipText     =   "Embed the encrypted text into the graphic."
      Top             =   7440
      Width           =   1335
   End
   Begin VB.CommandButton RetrieveBttn 
      Caption         =   "&Retrieve"
      Height          =   375
      Left            =   2640
      TabIndex        =   9
      ToolTipText     =   "Retrieve embedded message from the graphic."
      Top             =   7440
      Width           =   1335
   End
   Begin VB.CommandButton SavePicBttn 
      Caption         =   "S&ave Picture"
      Height          =   375
      Left            =   3960
      TabIndex        =   10
      ToolTipText     =   "Save the graphic out to a file."
      Top             =   7440
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   1440
      Top             =   7680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Height          =   2415
      Left            =   0
      ScaleHeight     =   2355
      ScaleWidth      =   7875
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   8160
      Width           =   7935
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2295
         Left            =   0
         MousePointer    =   1  'Arrow
         ScaleHeight     =   153
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   529
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   0
         Width           =   7935
      End
   End
   Begin VB.CommandButton Dummy 
      Caption         =   "Command5"
      Height          =   495
      Left            =   0
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   14880
      Width           =   1215
   End
   Begin RichTextLib.RichTextBox DecryptedBox 
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   840
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   7646
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"Main.frx":00FE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   10620
      Width           =   7980
      _ExtentX        =   14076
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   158750
            MinWidth        =   158750
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7440
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label4 
      Caption         =   "Encryption Key:"
      Height          =   255
      Left            =   1560
      TabIndex        =   18
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Picture:"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   7920
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Encrypted Text"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Unencrypted Text"
      Height          =   255
      Left            =   0
      TabIndex        =   15
      Top             =   600
      Width           =   1575
   End
   Begin VB.Menu FileMnu 
      Caption         =   "&File"
      Begin VB.Menu OpenText_FileMnu 
         Caption         =   "Open &Text File"
      End
      Begin VB.Menu OpenPic_FileMnu 
         Caption         =   "Open Picture File"
      End
      Begin VB.Menu SaveText_FileMnu 
         Caption         =   "Save Text"
      End
      Begin VB.Menu SavePic_FileMnu 
         Caption         =   "Save Picture"
      End
      Begin VB.Menu Quit_FileMnu 
         Caption         =   "Quit"
      End
   End
   Begin VB.Menu Encryption_Mnu 
      Caption         =   "&Encryption"
      Begin VB.Menu EncText_EncMnu 
         Caption         =   "Encrypt Text"
      End
      Begin VB.Menu DecText_EncMnu 
         Caption         =   "Decrypt Text"
      End
      Begin VB.Menu Clear_EncMnu 
         Caption         =   "Clear Text"
      End
   End
   Begin VB.Menu EmbedMnu 
      Caption         =   "E&mbedding"
      Begin VB.Menu Embed_EmbMnu 
         Caption         =   "Embed Text Into Picture"
      End
      Begin VB.Menu Retrieve_EmbMnu 
         Caption         =   "Retrieve Text From Picture"
      End
   End
   Begin VB.Menu HelpMnu 
      Caption         =   "&Help"
      Begin VB.Menu Help_HelpMnu 
         Caption         =   "Help"
      End
      Begin VB.Menu About_HelpMnu 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "AppMainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub About_HelpMnu_Click()
    AboutFrm.Visible = True
End Sub

Private Sub Clear_EncMnu_Click()
    Call ClearBttn_Click
End Sub

Private Sub ClearBttn_Click()
    DecryptedBox.Text = ""
    EncryptedBox.Text = ""
    hSay ("")
    KeyBox.Text = ""
End Sub


Private Sub Command1_Click()
    End
End Sub

Private Sub DecryptBttn_Click()
    If Len(EncryptedBox.Text) > 0 And Len(KeyBox.Text) > 0 Then
        DecryptedBox.Text = hDecrypt(EncryptedBox.Text, KeyBox.Text)
        EncryptedBox.Text = ""
        hSay ("Text decrypted.")
    Else
        If Len(EncryptedBox.Text) = 0 Then
            MsgBox "There is no text to decrypt!"
        Else
            MsgBox "There is no key!"
        End If
    End If
End Sub



Private Sub DecText_EncMnu_Click()
    Call DecryptBttn_Click
End Sub

Private Sub Embed_EmbMnu_Click()
    Call EmbedBttn_Click
End Sub

Private Sub EmbedBttn_Click()
    Dim whatToEncrypt As String
    If Len(DecryptedBox.Text) > 0 Or Len(EncryptedBox.Text) > 0 Then
        If Picture2.Picture.Width > 0 Then
            If Len(EncryptedBox.Text) > 1 Then
                hSay ("Embedding encrypted text....")
                whatToEncrypt = Chr(250) + EncryptedBox.Text
            Else
                hSay ("Embedding unencrypted text...")
                whatToEncrypt = DecryptedBox.Text
            End If
            If (hPutMessage(whatToEncrypt)) Then
                DecryptedBox.Text = ""
                EncryptedBox.Text = ""
                hSay ("Message embedded.")
            Else
                hSay ("Picture reported error.")
            End If
        Else
            MsgBox "No picture open!"
        End If
    Else
        MsgBox "Nothing to embed into image!"
    End If
End Sub

Private Sub EncryptBttn_Click()
    If Len(DecryptedBox.Text) > 0 And Len(KeyBox.Text) > 0 Then
        EncryptedBox.Text = hEncrypt(DecryptedBox.Text, KeyBox.Text)
        DecryptedBox.Text = ""
        hSay ("Text encrypted.")
    Else
        If Len(DecryptedBox.Text) = 0 Then
            MsgBox "There is no text to encrypt!"
        Else
            MsgBox "There is no key!"
        End If
    End If
End Sub

Private Sub EncryptedBox_GotFocus()
'    Dummy.SetFocus
End Sub

Private Sub EncText_EncMnu_Click()
    Call EncryptBttn_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub Help_HelpMnu_Click()
    HelpFrm.Visible = True
End Sub

Private Sub OpenBttn_Click()
    Dim URLFilePath As String
    Dim fileContents As String
    Dim firstChar As String
    
    CommonDialog1.Filter = "Text File (*.txt)|*.txt"
    CommonDialog1.ShowOpen
    URLFilePath = CommonDialog1.FileName
    If Len(URLFilePath) > 3 Then
        DecryptedBox.Text = ""
        EncryptedBox.Text = ""
        TempBox.LoadFile URLFilePath
        firstChar = Left(TempBox.Text, 1)
        If Asc(firstChar) = 250 Then
            fileContents = Right(TempBox.Text, (Len(TempBox.Text) - 1))
            EncryptedBox.Text = fileContents
        Else
            DecryptedBox.Text = TempBox.Text
        End If
        hSay ("File opened : " + URLFilePath)
    End If
End Sub

Private Sub hSay(whatToSay As String)
    StatusBar1.Panels(1).Text = whatToSay
End Sub

Private Sub OpenPic_FileMnu_Click()
    Call OpenPictureBttn_Click
End Sub

Private Sub OpenPictureBttn_Click()
    CommonDialog2.Filter = "Pictures (*.bmp;*.gif;*.jpg;*.jpeg)|*.bmp;*.gif;*.jpg;*.jpeg"
    CommonDialog2.ShowOpen
    If CommonDialog2.FileName <> "" Then
        Picture2.Picture = LoadPicture(CommonDialog2.FileName)
        hSay ("Graphic file opened : " + CommonDialog2.FileName)
    End If
End Sub

Private Sub OpenText_FileMnu_Click()
    Call OpenBttn_Click
End Sub

Private Sub Quit_FileMnu_Click()
    Call Command1_Click
End Sub

Private Sub Retrieve_EmbMnu_Click()
    Call RetrieveBttn_Click
End Sub

Private Sub RetrieveBttn_Click()
    Dim retMsg As String
    Dim firstChar As String
    If Picture2.Picture.Width > 0 Then
        retMsg = hGetMessage
        firstChar = Left(retMsg, 1)
        If Asc(firstChar) = 250 Then
            retMsg = Right(retMsg, (Len(retMsg) - 1))
            DecryptedBox.Text = ""
            EncryptedBox.Text = ""
            EncryptedBox.Text = retMsg
        Else
            DecryptedBox.Text = ""
            EncryptedBox.Text = ""
            DecryptedBox.Text = retMsg
        End If
        hSay ("Message retrieved.")
    Else
        MsgBox "No picture open!"
    End If
End Sub

Private Sub SaveBttn_Click()
    Dim fileToSave As String
    Dim fileContents As String
    Dim MyValue As Integer
    Randomize   ' Initialize random-number generator
    If Len(DecryptedBox.Text) > 1 Then
        fileContents = DecryptedBox.Text
    Else
        If Len(EncryptedBox.Text) > 1 Then
            fileContents = EncryptedBox.Text
            fileContents = Chr(250) + fileContents
        Else
            MsgBox "No text to save!"
        End If
    End If
    If Len(fileContents) > 1 Then
        CommonDialog1.Filter = "Text File (*.txt)|*.txt"
        CommonDialog1.ShowSave
        fileToSave = CommonDialog1.FileName
        If Len(fileToSave) > 3 Then
            hWriteTextFile fileToSave, fileContents
            hSay ("File saved : " + fileToSave)
        End If
    End If
End Sub

Private Sub SavePic_FileMnu_Click()
    Call SavePicBttn_Click
End Sub

Private Sub SavePicBttn_Click()
If Picture2.Picture.Width > 0 Then
    CommonDialog2.FileName = Left(CommonDialog2.FileName, InStrRev(CommonDialog2.FileName, "."))
    CommonDialog2.FileName = CommonDialog2.FileName + "bmp"
    CommonDialog2.Filter = "Pictures (*.bmp)|*.bmp"
    CommonDialog2.ShowSave
    If CommonDialog2.FileName <> "" Then
        SavePicture Picture2.Image, CommonDialog2.FileName
        hSay ("Graphic file saved : " + CommonDialog2.FileName)
    End If
Else
    MsgBox "No picture open!"
End If
End Sub

Private Sub SaveText_FileMnu_Click()
    Call SaveBttn_Click
End Sub
