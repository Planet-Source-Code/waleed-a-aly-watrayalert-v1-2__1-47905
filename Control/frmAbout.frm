VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6000
   Icon            =   "frmAbout.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEMail 
      Caption         =   "&EMail Me"
      Default         =   -1  'True
      Height          =   375
      Left            =   4800
      TabIndex        =   0
      Top             =   4680
      Width           =   1095
   End
   Begin VB.PictureBox picHeader 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   6000
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   6000
      Begin VB.Label lblCopyright 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright © 2003 Waleed A. Aly, all rights reserved."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   180
         TabIndex        =   3
         Top             =   420
         Width           =   3660
      End
      Begin VB.Label lblRelease 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[Release]"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   180
         TabIndex        =   2
         Top             =   180
         Width           =   1050
      End
   End
   Begin VB.Label lblContact 
      Caption         =   "Any comments and/or suggestions are welcome on my EMail address."
      Height          =   195
      Left            =   300
      TabIndex        =   9
      Top             =   4140
      Width           =   5415
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Information:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   195
      Index           =   2
      Left            =   180
      TabIndex        =   8
      Top             =   3840
      Width           =   1770
   End
   Begin VB.Label lblWarning 
      Caption         =   $"frmAbout.frx":000C
      Height          =   615
      Left            =   300
      TabIndex        =   7
      Top             =   3060
      Width           =   5400
   End
   Begin VB.Label lblLicence 
      Caption         =   $"frmAbout.frx":00C3
      Height          =   1215
      Left            =   300
      TabIndex        =   5
      Top             =   1380
      Width           =   5400
   End
   Begin VB.Line line 
      BorderColor     =   &H80000010&
      Index           =   2
      X1              =   -60
      X2              =   6000
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line line 
      BorderColor     =   &H80000014&
      Index           =   3
      X1              =   -60
      X2              =   6000
      Y1              =   4575
      Y2              =   4575
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Licence Agreement:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   4
      Top             =   1020
      Width           =   1680
   End
   Begin VB.Label lblVote 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HERE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   560
      MouseIcon       =   "frmAbout.frx":0264
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   4740
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Click            for the source code in VB! Please Vote ;)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   180
      TabIndex        =   10
      Top             =   4740
      Visible         =   0   'False
      Width           =   3765
   End
   Begin VB.Line line 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   0
      X2              =   6000
      Y1              =   870
      Y2              =   870
   End
   Begin VB.Line line 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   0
      X2              =   6000
      Y1              =   855
      Y2              =   855
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Warning:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   195
      Index           =   1
      Left            =   180
      TabIndex        =   6
      Top             =   2760
      Width           =   750
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mRelease As String, mCodeURL As String
Private Const sEMail As String = "wa_aly@hotmail.com"
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub cmdEMail_Click()
    ExecuteLink "mailto:" & sEMail & "?subject=" & mRelease
End Sub

Private Sub Form_Load()
    lblRelease = mRelease
    If mCodeURL <> "" Then
        lbl(3).Visible = True
        lblVote.Visible = True
    End If
    cmdEMail.ToolTipText = sEMail
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DeformatURL lblVote
End Sub

Private Sub lbl_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    DeformatURL lblVote
End Sub

Private Sub lblVote_Click()
    ExecuteLink mCodeURL
End Sub

Private Sub lblVote_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormatURL lblVote
End Sub

Private Sub picHeader_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DeformatURL lblVote
End Sub

Private Sub FormatURL(lblURL As Label)
    lblURL.FontUnderline = True
    lblURL.ForeColor = vbBlue
End Sub

Private Sub DeformatURL(lblURL As Label)
    lblURL.FontUnderline = False
    lblURL.ForeColor = vbHighlight
End Sub

Private Sub ExecuteLink(sLink As String)
    ShellExecute Me.hWnd, vbNullString, sLink, vbNullString, "C:\", 0
End Sub

Public Property Let Release(Value As String)
    mRelease = Value
End Property

Public Property Let CodeURL(Value As String)
    mCodeURL = Value
End Property
