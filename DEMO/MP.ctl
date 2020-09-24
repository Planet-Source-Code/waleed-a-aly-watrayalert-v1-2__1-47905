VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.UserControl MP 
   AutoRedraw      =   -1  'True
   BackStyle       =   0  'Transparent
   ClientHeight    =   3675
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3495
   LockControls    =   -1  'True
   ScaleHeight     =   3675
   ScaleWidth      =   3495
   Begin MSComDlg.CommonDialog Dlg 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open"
      Filter          =   "Video files (all types)|*.mpg;*.mpeg;*.avi;*.asf;*.wmv"
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&Open a Video File"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   3300
      Width           =   3495
   End
   Begin WMPLibCtl.WindowsMediaPlayer P 
      Height          =   3315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3495
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   0   'False
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   -1  'True
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   6165
      _cy             =   5847
   End
End
Attribute VB_Name = "MP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private sFileName As String

Private Sub UserControl_Resize()
    cmdLoad.Top = P.Height
    cmdLoad.Width = P.Width
    UserControl.Width = P.Width
    UserControl.Height = P.Height + cmdLoad.Height
End Sub

Private Sub cmdLoad_Click()
On Error Resume Next
    P.Close
    sFileName = GetMediaFile
    P.URL = sFileName
End Sub

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Sub StartPlayBack()
On Error Resume Next
    P.Close
    If sFileName = "" Then sFileName = GetMediaFile
    P.URL = sFileName
End Sub

Public Sub StopPlayBack()
    P.Close
End Sub

Private Function GetMediaFile() As String
    Dlg.ShowOpen
    GetMediaFile = Dlg.FileName
    Dlg.FileName = ""
End Function
