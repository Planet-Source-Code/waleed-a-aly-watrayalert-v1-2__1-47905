VERSION 5.00
Begin VB.UserControl NoteBook 
   AutoRedraw      =   -1  'True
   ClientHeight    =   4275
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3660
   LockControls    =   -1  'True
   ScaleHeight     =   4275
   ScaleWidth      =   3660
   Begin VB.PictureBox picNB 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4275
      Left            =   0
      Picture         =   "NoteBook.ctx":0000
      ScaleHeight     =   4275
      ScaleWidth      =   3660
      TabIndex        =   0
      Top             =   0
      Width           =   3660
      Begin VB.CommandButton cmdOK 
         Caption         =   "&Ok"
         Height          =   435
         Left            =   900
         TabIndex        =   2
         Top             =   3540
         Width           =   1935
      End
      Begin VB.Label lblCaption 
         BackStyle       =   0  'Transparent
         Caption         =   "Caption"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   2715
         Left            =   180
         TabIndex        =   1
         Top             =   660
         Width           =   3315
      End
   End
End
Attribute VB_Name = "NoteBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event ButtonClick()

Private Sub cmdOK_Click()
    RaiseEvent ButtonClick
End Sub

Private Sub UserControl_Initialize()
    lblCaption = CStr(Date) & vbCrLf & vbCrLf & _
                "11:00 AM    " & "SCU lab meeting" & vbCrLf & _
                "01:30 PM    " & "CANCELED" & vbCrLf & _
                "04:00 PM    " & "Nancy" & vbCrLf & _
                "09:00 PM    " & "dinner" & vbCrLf & _
                "10:00 PM    " & "the movies" & vbCrLf
End Sub

Private Sub UserControl_Resize()
    UserControl.Width = picNB.Width
    UserControl.Height = picNB.Height
End Sub

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property
