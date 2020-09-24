VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "*\A..\Control\TrayAlert.vbp"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "waTrayAlert v1.2 Demo Application"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10710
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   10710
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAbout 
      Caption         =   "?"
      Height          =   315
      Left            =   10080
      TabIndex        =   29
      Top             =   6480
      Width           =   315
   End
   Begin VB.Frame fCustom 
      Caption         =   " Execute "
      Height          =   3315
      Left            =   6540
      TabIndex        =   55
      Top             =   2940
      Width           =   3975
      Begin VB.PictureBox pic 
         BorderStyle     =   0  'None
         Height          =   2415
         Index           =   0
         Left            =   180
         ScaleHeight     =   2415
         ScaleWidth      =   3615
         TabIndex        =   74
         Top             =   600
         Width           =   3615
         Begin VB.CommandButton cmdInno 
            Caption         =   "Innovative Alert"
            Height          =   495
            Left            =   1800
            TabIndex        =   75
            Top             =   1920
            Width           =   1815
         End
         Begin VB.CommandButton cmdCounter 
            Caption         =   "Count Down"
            Height          =   495
            Left            =   0
            TabIndex        =   76
            Top             =   1920
            Width           =   1815
         End
         Begin VB.CommandButton cmdStylish 
            Caption         =   "Stylish Alert"
            Height          =   495
            Left            =   1800
            TabIndex        =   77
            Top             =   1440
            Width           =   1815
         End
         Begin VB.CommandButton cmdPoll 
            Caption         =   "Simple Poll"
            Height          =   495
            Left            =   0
            TabIndex        =   78
            Top             =   1440
            Width           =   1815
         End
         Begin VB.CommandButton cmdNoteBook 
            Caption         =   "NoteBook"
            Height          =   495
            Left            =   1800
            TabIndex        =   79
            Top             =   960
            Width           =   1815
         End
         Begin VB.CommandButton cmdNewsBar 
            Caption         =   "NewsBar"
            Height          =   495
            Left            =   0
            TabIndex        =   80
            Top             =   960
            Width           =   1815
         End
         Begin VB.CommandButton cmdUnloadMP 
            Caption         =   "Unload Media Player"
            Enabled         =   0   'False
            Height          =   495
            Left            =   1800
            TabIndex        =   81
            Top             =   480
            Width           =   1815
         End
         Begin VB.CommandButton cmdLoadMP 
            Caption         =   "Load Media Player"
            Height          =   495
            Left            =   0
            TabIndex        =   82
            Top             =   480
            Width           =   1815
         End
         Begin VB.CommandButton cmdTrigger 
            Caption         =   "Read Properties && Trigger Alert"
            Height          =   495
            Left            =   0
            TabIndex        =   83
            Top             =   0
            Width           =   3615
         End
      End
      Begin VB.PictureBox picContainer 
         Height          =   915
         Left            =   240
         ScaleHeight     =   855
         ScaleWidth      =   3375
         TabIndex        =   56
         Top             =   540
         Visible         =   0   'False
         Width           =   3435
         Begin VB.PictureBox picAlert 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFC0FF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   2775
            Left            =   0
            ScaleHeight     =   2775
            ScaleWidth      =   3495
            TabIndex        =   66
            Top             =   0
            Width           =   3495
            Begin VB.PictureBox picCommand 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   1260
               ScaleHeight     =   225
               ScaleWidth      =   1065
               TabIndex        =   70
               Top             =   2400
               Width           =   1095
               Begin VB.Label lblAlert 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "&Ok"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   2
                  Left            =   420
                  TabIndex        =   71
                  Top             =   0
                  Width           =   240
               End
            End
            Begin VB.OptionButton opt 
               BackColor       =   &H00FFC0FF&
               Caption         =   "That Sucks !"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   2
               Left            =   180
               TabIndex        =   69
               Top             =   1980
               Width           =   1395
            End
            Begin VB.OptionButton opt 
               BackColor       =   &H00FFC0FF&
               Caption         =   "It's Okay .."
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   1
               Left            =   180
               TabIndex        =   68
               Top             =   1680
               Width           =   1275
            End
            Begin VB.OptionButton opt 
               BackColor       =   &H00FFC0FF&
               Caption         =   "I love it !"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   0
               Left            =   180
               TabIndex        =   67
               Top             =   1380
               Width           =   1095
            End
            Begin VB.Label lblAlert 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "What do you think of my control ?"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   1
               Left            =   180
               TabIndex        =   73
               Top             =   960
               Width           =   2910
            End
            Begin VB.Label lblAlert 
               BackStyle       =   0  'Transparent
               Caption         =   "This is to show you that you can use this control to 'TAILOR' very custom alerts. This is a simple poll ;)"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   615
               Index           =   0
               Left            =   180
               TabIndex        =   72
               Top             =   120
               Width           =   3135
            End
         End
         Begin DEMO.NoteBook NoteBook 
            Height          =   4275
            Left            =   0
            TabIndex        =   65
            Top             =   0
            Width           =   3660
            _ExtentX        =   6456
            _ExtentY        =   7541
         End
         Begin DEMO.MP MP 
            Height          =   3690
            Left            =   0
            TabIndex        =   64
            Top             =   0
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   6509
         End
         Begin VB.PictureBox picShania 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   2925
            Left            =   3660
            Picture         =   "frmMain.frx":1042
            ScaleHeight     =   2925
            ScaleWidth      =   2970
            TabIndex        =   63
            Top             =   1320
            Width           =   2970
         End
         Begin MSComDlg.CommonDialog Dlg 
            Left            =   3480
            Top             =   2700
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            Flags           =   1
            FontName        =   "MS Sans Serif"
         End
         Begin VB.PictureBox picLogo 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   45
            Left            =   600
            Picture         =   "frmMain.frx":4ED9
            ScaleHeight     =   45
            ScaleWidth      =   2355
            TabIndex        =   59
            Top             =   4080
            Width           =   2355
         End
         Begin VB.PictureBox picCounter 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   2055
            Left            =   3660
            ScaleHeight     =   2055
            ScaleWidth      =   2355
            TabIndex        =   57
            ToolTipText     =   $"frmMain.frx":53AD
            Top             =   0
            Width           =   2355
            Begin VB.Timer Timer 
               Enabled         =   0   'False
               Interval        =   1000
               Left            =   60
               Top             =   60
            End
            Begin VB.Label lblCount 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "5"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   72
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1815
               Left            =   180
               TabIndex        =   58
               Top             =   120
               Width           =   1995
            End
         End
         Begin DEMO.NewsBar NewsBar 
            Height          =   555
            Left            =   0
            Top             =   0
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   979
            News            =   "NewsBar Control Coded by: Waleed A. Aly"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   128
            ForeColor       =   16777215
            Speed           =   5
            Cycles          =   1
         End
      End
   End
   Begin VB.Frame fText 
      Caption         =   " Text "
      Height          =   2775
      Left            =   180
      TabIndex        =   50
      Top             =   60
      Width           =   6195
      Begin VB.PictureBox pic 
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   1
         Left            =   5580
         ScaleHeight     =   315
         ScaleWidth      =   375
         TabIndex        =   61
         Top             =   720
         Width           =   375
         Begin VB.CommandButton cmdWAVE 
            Caption         =   "..."
            Height          =   285
            Left            =   0
            TabIndex        =   27
            Tag             =   "Wave Files|*.wav"
            Top             =   0
            Width           =   375
         End
      End
      Begin VB.TextBox txtCaption 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1425
         Left            =   1020
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         Text            =   "frmMain.frx":545A
         Top             =   1080
         Width           =   4935
      End
      Begin VB.TextBox txtWave 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1020
         TabIndex        =   51
         Top             =   720
         Width           =   4455
      End
      Begin VB.TextBox txtLink 
         Height          =   285
         Left            =   1020
         TabIndex        =   26
         Text            =   "mailto:wa_aly@hotmail.com?subject=waTrayAlert v1.2"
         Top             =   360
         Width           =   4935
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Caption"
         Height          =   195
         Index           =   13
         Left            =   180
         TabIndex        =   54
         Top             =   1140
         Width           =   540
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Wave File"
         Height          =   195
         Index           =   12
         Left            =   180
         TabIndex        =   53
         Top             =   780
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Link"
         Height          =   195
         Index           =   11
         Left            =   180
         TabIndex        =   52
         Top             =   420
         Width           =   300
      End
   End
   Begin VB.Frame fAppearance 
      Caption         =   " Appearance "
      Height          =   3315
      Left            =   180
      TabIndex        =   36
      Top             =   2940
      Width           =   6195
      Begin TrayAlert.waTrayAlert Alert 
         Left            =   3000
         Top             =   600
         _ExtentX        =   847
         _ExtentY        =   847
         Alignment       =   2
         BackColor1      =   12615808
         BackColor2      =   16777215
         Borderless      =   0   'False
         CloseButton     =   -1  'True
         ForeColor       =   0
         GradientAngle   =   270
         Icon            =   "frmMain.frx":546C
         LinkColor       =   12582912
         Margin          =   400
         MaskColor       =   16777215
         Offset          =   400
         Stretch         =   0   'False
         TextInset       =   0   'False
         Transparency    =   255
         UseMask         =   0   'False
         AlertStyle      =   1
         AlwaysOnTop     =   -1  'True
         Animation       =   0
         AutoSize        =   0
         Duration        =   5000
         Persisting      =   0   'False
         Speed           =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "By: Waleed A. Aly"
         Link            =   "mailto:wa_aly@hotmail.com?subject=waTrayAlert v1.2"
         WaveSound       =   "C:\Program Files\MSN Messenger\type.wav"
      End
      Begin VB.PictureBox pic 
         BorderStyle     =   0  'None
         Height          =   1035
         Index           =   2
         Left            =   1740
         ScaleHeight     =   1035
         ScaleWidth      =   1935
         TabIndex        =   62
         Top             =   1620
         Width           =   1935
         Begin VB.CommandButton cmdAddIcon 
            Caption         =   "Change"
            Height          =   285
            Left            =   0
            TabIndex        =   5
            Tag             =   "Icons|*.ico;*.cur"
            Top             =   0
            Width           =   915
         End
         Begin VB.CommandButton cmdRemIcon 
            Caption         =   "Remove"
            Height          =   285
            Left            =   1020
            TabIndex        =   6
            Top             =   0
            Width           =   915
         End
         Begin VB.CommandButton cmdAddPic 
            Caption         =   "Change"
            Height          =   285
            Left            =   0
            TabIndex        =   7
            Tag             =   "All Picture Files|*.bmp;*.dib;*.gif;*.jpg;*.wmf;*.emf;*.ico;*.cur"
            Top             =   360
            Width           =   915
         End
         Begin VB.CommandButton cmdRemPic 
            Caption         =   "Remove"
            Enabled         =   0   'False
            Height          =   285
            Left            =   1020
            TabIndex        =   8
            Top             =   360
            Width           =   915
         End
         Begin VB.CommandButton cmdFONT 
            Caption         =   "Font"
            Height          =   285
            Left            =   0
            TabIndex        =   9
            Top             =   720
            Width           =   1935
         End
      End
      Begin VB.TextBox txtMargin 
         Height          =   285
         Left            =   5280
         MaxLength       =   5
         TabIndex        =   16
         Text            =   "400"
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox txtOffset 
         Height          =   285
         Left            =   5280
         MaxLength       =   5
         TabIndex        =   17
         Text            =   "400"
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox txtTransparency 
         Height          =   285
         Left            =   5280
         MaxLength       =   3
         TabIndex        =   18
         Text            =   "255"
         Top             =   2400
         Width           =   615
      End
      Begin VB.CheckBox chkMask 
         Appearance      =   0  'Flat
         Caption         =   "Use Mask"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1560
         TabIndex        =   2
         Top             =   660
         Width           =   1035
      End
      Begin VB.CheckBox chkInset 
         Appearance      =   0  'Flat
         Caption         =   "Text Inset"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   660
         Width           =   1035
      End
      Begin VB.CheckBox chkStretch 
         Appearance      =   0  'Flat
         Caption         =   "Stretch Background Picture"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   2295
      End
      Begin VB.TextBox txtAngle 
         Height          =   285
         Left            =   5280
         MaxLength       =   3
         TabIndex        =   19
         Text            =   "270"
         Top             =   2760
         Width           =   615
      End
      Begin VB.CheckBox chkClose 
         Appearance      =   0  'Flat
         Caption         =   "Close Button"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1560
         TabIndex        =   4
         Top             =   960
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chkBorderless 
         Appearance      =   0  'Flat
         Caption         =   "Borderless"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   1035
      End
      Begin VB.ComboBox cmbAlignment 
         Height          =   315
         ItemData        =   "frmMain.frx":5971
         Left            =   1740
         List            =   "frmMain.frx":597E
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2700
         Width           =   1935
      End
      Begin VB.PictureBox picLINK 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5280
         ScaleHeight     =   225
         ScaleWidth      =   585
         TabIndex        =   15
         Top             =   1320
         Width           =   615
      End
      Begin VB.PictureBox picMASK 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5280
         ScaleHeight     =   225
         ScaleWidth      =   585
         TabIndex        =   14
         Top             =   1080
         Width           =   615
      End
      Begin VB.PictureBox picFORE 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5280
         ScaleHeight     =   225
         ScaleWidth      =   585
         TabIndex        =   13
         Top             =   840
         Width           =   615
      End
      Begin VB.PictureBox picBACK2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5280
         ScaleHeight     =   225
         ScaleWidth      =   585
         TabIndex        =   12
         Top             =   600
         Width           =   615
      End
      Begin VB.PictureBox picBACK1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C08080&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5280
         ScaleHeight     =   225
         ScaleWidth      =   585
         TabIndex        =   11
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Caption Font"
         Height          =   195
         Index           =   20
         Left            =   240
         TabIndex        =   49
         Top             =   2400
         Width           =   900
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Background Picture"
         Height          =   195
         Index           =   19
         Left            =   240
         TabIndex        =   48
         Top             =   2040
         Width           =   1410
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Icon Picture"
         Height          =   195
         Index           =   18
         Left            =   240
         TabIndex        =   47
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Caption Alignment"
         Height          =   195
         Index           =   17
         Left            =   240
         TabIndex        =   46
         Top             =   2760
         Width           =   1275
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Transparency"
         Height          =   195
         Index           =   8
         Left            =   3960
         TabIndex        =   45
         Top             =   2460
         Width           =   975
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Offset"
         Height          =   195
         Index           =   7
         Left            =   3960
         TabIndex        =   44
         Top             =   2100
         Width           =   420
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Margin"
         Height          =   195
         Index           =   6
         Left            =   3960
         TabIndex        =   43
         Top             =   1740
         Width           =   480
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Gradient Angle"
         Height          =   195
         Index           =   5
         Left            =   3960
         TabIndex        =   42
         Top             =   2820
         Width           =   1050
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "LinkColor"
         Height          =   195
         Index           =   4
         Left            =   3960
         TabIndex        =   41
         Top             =   1335
         Width           =   660
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "MaskColor"
         Height          =   195
         Index           =   3
         Left            =   3960
         TabIndex        =   40
         Top             =   1095
         Width           =   750
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "ForeColor"
         Height          =   195
         Index           =   2
         Left            =   3960
         TabIndex        =   39
         Top             =   855
         Width           =   675
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "BackColor2"
         Height          =   195
         Index           =   1
         Left            =   3960
         TabIndex        =   38
         Top             =   615
         Width           =   825
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "BackColor1"
         Height          =   195
         Index           =   0
         Left            =   3960
         TabIndex        =   37
         Top             =   375
         Width           =   825
      End
   End
   Begin VB.Frame fBehavior 
      Caption         =   " Behavior "
      Height          =   2775
      Left            =   6540
      TabIndex        =   30
      Top             =   60
      Width           =   3975
      Begin VB.TextBox txtDuration 
         Height          =   285
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   24
         Text            =   "5000"
         Top             =   1860
         Width           =   1695
      End
      Begin VB.TextBox txtSpeed 
         Height          =   285
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   25
         Text            =   "0"
         Top             =   2220
         Width           =   1695
      End
      Begin VB.ComboBox cmbStyle 
         Height          =   315
         ItemData        =   "frmMain.frx":59B7
         Left            =   1920
         List            =   "frmMain.frx":59C1
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   720
         Width           =   1695
      End
      Begin VB.ComboBox cmbAutoSize 
         Height          =   315
         ItemData        =   "frmMain.frx":59E1
         Left            =   1920
         List            =   "frmMain.frx":59EB
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   1080
         Width           =   1695
      End
      Begin VB.ComboBox cmbAnimation 
         Height          =   315
         ItemData        =   "frmMain.frx":5A11
         Left            =   1920
         List            =   "frmMain.frx":5A1E
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CheckBox chkOnTop 
         Appearance      =   0  'Flat
         Caption         =   "Always On Top"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   420
         TabIndex        =   20
         Top             =   360
         Value           =   1  'Checked
         Width           =   1395
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Animation"
         Height          =   195
         Index           =   16
         Left            =   420
         TabIndex        =   35
         Top             =   1500
         Width           =   690
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "AutoSizing Method"
         Height          =   195
         Index           =   15
         Left            =   420
         TabIndex        =   34
         Top             =   1140
         Width           =   1335
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Alert Style"
         Height          =   195
         Index           =   14
         Left            =   420
         TabIndex        =   33
         Top             =   780
         Width           =   705
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Speed Delay"
         Height          =   195
         Index           =   10
         Left            =   420
         TabIndex        =   32
         Top             =   2280
         Width           =   915
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Duration"
         Height          =   195
         Index           =   9
         Left            =   420
         TabIndex        =   31
         Top             =   1920
         Width           =   600
      End
   End
   Begin VB.TextBox txtLog 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   1275
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   60
      Text            =   "frmMain.frx":5A46
      Top             =   6420
      Width           =   10695
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'BY: Waleed A. Aly [22 M Egypt]
'This is a freeware. However, all I ask you for is your feedbacks. EMail me
'if you like this control. This will be very encouraging. Thank you.

'NOTE:
'waTrayAlert control raises several errors, review the help file for error
'numbers and descriptions. You should always write error handlers.

Private objICON As StdPicture
Private objPICTURE As StdPicture
Private CustomProps As taProperties
Private Declare Sub InitCommonControls Lib "comctl32" ()

Private Sub cmbAlignment_Click()
    txtCaption.Alignment = cmbAlignment.ListIndex
End Sub

Private Sub cmdAbout_Click()
    Alert.AboutBox
End Sub

Private Sub cmdAddIcon_Click()
On Error Resume Next
    Set objICON = LoadPicture(GetFile(cmdAddIcon.Tag))
    If objICON <> 0 Then cmdRemIcon.Enabled = True Else cmdRemIcon.Enabled = False
End Sub

Private Sub cmdAddPic_Click()
On Error Resume Next
    Set objPICTURE = LoadPicture(GetFile(cmdAddPic.Tag))
    If objPICTURE <> 0 Then cmdRemPic.Enabled = True Else cmdRemPic.Enabled = False
End Sub

Private Sub cmdFONT_Click()

    Dlg.FontName = txtCaption.FontName
    Dlg.ShowFont
    With txtCaption
        .FontName = Dlg.FontName
        .FontBold = Dlg.FontBold
        .FontItalic = Dlg.FontItalic
        .FontSize = Dlg.FontSize
        .FontStrikethru = Dlg.FontStrikethru
        .FontUnderline = Dlg.FontUnderline
    End With

End Sub

Private Sub cmdRemIcon_Click()
    Set objICON = Nothing
    cmdRemIcon.Enabled = False
End Sub

Private Sub cmdRemPic_Click()
    Set objPICTURE = Nothing
    cmdRemPic.Enabled = False
End Sub

Private Sub cmdTrigger_Click()
On Error Resume Next

    Dim Props As taProperties
    
    With Props
        .AlertStyle = cmbStyle.ListIndex
        .Alignment = cmbAlignment.ListIndex
        .AlwaysOnTop = chkOnTop
        .Animation = cmbAnimation.ListIndex
        .AutoSize = cmbAutoSize.ListIndex
        .BackColor1 = picBACK1.BackColor
        .BackColor2 = picBACK2.BackColor
        .Borderless = chkBorderless
        .Caption = txtCaption
        .CloseButton = chkClose
        .duration = CLng(txtDuration)
        Set .Font = txtCaption.Font
        .ForeColor = picFORE.BackColor
        .GradientAngle = CSng(txtAngle)
        Set .Icon = objICON
        .Link = txtLink
        .LinkColor = picLINK.BackColor
        .Margin = CSng(txtMargin)
        .MaskColor = picMASK.BackColor
        .Offset = CSng(txtOffset)
        Set .Picture = objPICTURE
        .Speed = CByte(txtSpeed)
        .Stretch = chkStretch
        .TextInset = chkInset
        .Transparency = CByte(txtTransparency)
        .UseMask = chkMask
        .WaveSound = txtWave
    End With
    
    TriggerAlert Props

End Sub

Private Sub cmdWAVE_Click()
    txtWave = GetFile(cmdWAVE.Tag)
End Sub

Private Sub Form_Initialize()
    InitCommonControls
End Sub

Private Sub Form_Load()

    txtLog = "App Initializing ..." & vbCrLf
    
    cmbAlignment = "2 - vbCenter"
    cmbAnimation = "0 - taRoll"
    cmbAutoSize = "0 - taDefault"
    cmbStyle = "1 - taLink"
    
    If Not Alert.IsLayerSafe Then
        chkMask.Enabled = False
        picMASK.Enabled = False
        cmbAnimation.Enabled = False
        txtTransparency.Enabled = False
        txtLog = txtLog & "Mask & Transparency-Dependent features have been DISABLED. They are supported only on Windows 2000 or later." & vbCrLf
    End If
    
    NewsBar.AutoSize
    NewsBar.News = "BBC News : Israel's security fence in the West Bank is a major obstacle to the implementation of the roadmap peace plan." & vbCrLf & _
                    "At least 11 members of the same family - mostly children - have been killed in a coalition air strike by US forces on a residential district in central Iraq." & vbCrLf & _
                    "Failure to discover chemical and biological weapons in Iraq will be used by many groups to vilify the United States." & vbCrLf & _
                    "US Central Command has reported the deaths of 101 American service personnel in Iraq since 1 May when President Bush declared that major combat was over."
    
    Alert.DrawGradient picCounter, &HFFC0C0, vbWhite, 270
    
    CustomProps = Alert.DefaultProps
    Set objICON = CustomProps.Icon
    CustomProps.duration = 0
    CustomProps.WaveSound = Alert.WavePath(taMSN_Type)
    If CustomProps.WaveSound = "" Then CustomProps.WaveSound = Alert.WavePath(taMS_NewAlert)
    txtWave = CustomProps.WaveSound
    
    txtLog = txtLog & "App Loading Completed ..." & vbCrLf

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Alert.ClearAll True
End Sub

Private Sub Alert_AlertClick(ByVal Key As Integer, ByVal Tag As String)
    LogEvent Key, Tag, "Alert Clicked."
End Sub

Private Sub Alert_CaptionClick(ByVal Key As Integer, ByVal Tag As String)
    LogEvent Key, Tag, "Caption Clicked."
End Sub

Private Sub Alert_IconClick(ByVal Key As Integer, ByVal Tag As String)
    LogEvent Key, Tag, "Icon Clicked."
End Sub

Private Sub Alert_Created(ByVal Key As Integer, ByVal Tag As String, ByVal hWndAlert As Long)

    Select Case Tag
        Case "POLL"
            picAlert.Tag = Key
        Case "PLAYER"
            MP.Tag = Key
        Case "NEWS"
            NewsBar.Tag = Key
        Case "NOTE"
            NoteBook.Tag = Key
        Case "COUNTER"
            picCounter.Tag = Key
    End Select
    
    LogEvent Key, Tag, "Alert Created [HANDLE=" & hWndAlert & "]."

End Sub

Private Sub Alert_Loaded(ByVal Key As Integer, ByVal Tag As String)

    Select Case Tag
        Case "PLAYER"
            MP.StartPlayBack
        Case "NEWS"
            NewsBar.StartScrolling
        Case "COUNTER"
            Timer.Enabled = True
    End Select
    
    LogEvent Key, Tag, "Successfully Loaded."

End Sub

Private Sub Alert_QueryUnload(ByVal Key As Integer, ByVal Tag As String, ByVal UnloadMode As TrayAlert.taUnloadModes, Cancel As Boolean)

    Dim sReason As String
    
    Select Case UnloadMode
        Case taAlertExpired
            sReason = "Alert has expired."
        Case taCloseButtonClicked
            sReason = "Close Button Clicked."
        Case taClearingAllAlerts
            sReason = "Clearing All Alerts."
        Case taCodeInvoked
            sReason = "Invoked by Code."
    End Select
    
    LogEvent Key, Tag, "Asking Permission to Unload, REASON : " & sReason
    
    If Tag = "POLL" Then
        If Not (opt(0) Or opt(1) Or opt(2)) Then
            Cancel = True
            lblAlert(0).Font.Bold = True
            lblAlert(0) = "Please make a choice first :)"
        End If
    End If
    
    If Cancel Then LogEvent Key, Tag, "Cancelling the Unload."

End Sub

Private Sub Alert_Unloaded(ByVal Key As Integer, ByVal Tag As String)

    Select Case Tag
        Case "PLAYER"
            MP.StopPlayBack
        Case "NEWS"
            NewsBar.StopScrolling
        Case "COUNTER"
            Timer.Enabled = False
    End Select
    
    LogEvent Key, Tag, "Successfully Unloaded."

End Sub

Private Sub cmdInno_Click()

    Dim Props As taProperties
    
    Props = CustomProps
    Props.duration = 5000
    Set Props.Icon = Nothing
    Set Props.Picture = picShania.Picture
    Props.AutoSize = taPictureSize
    TriggerAlert Props

End Sub

Private Sub cmdCounter_Click()
    LoadContainerAlert CustomProps, picCounter.hWnd, "COUNTER"
End Sub

Private Sub cmdNewsBar_Click()

    Dim Props As taProperties
    
    If Alert.AlertCount <> 0 Then
        If MsgBox("I'd suggest you close all other alerts first. Would you like to ?", vbInformation Or vbYesNo) = vbYes Then Alert.ClearAll True
    End If
    
    Props = CustomProps
    Props.Borderless = True
    Props.Offset = 0
    LoadContainerAlert Props, NewsBar.hWnd, "NEWS"

End Sub

Private Sub cmdNoteBook_Click()

    Dim Props As taProperties
    
    Props = CustomProps
    Props.Borderless = True
    LoadContainerAlert Props, NoteBook.hWnd, "NOTE"

End Sub

Private Sub cmdPoll_Click()
    LoadContainerAlert CustomProps, picAlert.hWnd, "POLL"
End Sub

Private Sub cmdStylish_Click()
    TriggerAlert CustomProps, picLogo.hWnd
End Sub

Private Sub cmdLoadMP_Click()
    cmdLoadMP.Enabled = False
    LoadContainerAlert CustomProps, MP.hWnd, "PLAYER"
    cmdUnloadMP.Enabled = True
    cmdUnloadMP.SetFocus
End Sub

Private Sub cmdUnloadMP_Click()
    cmdUnloadMP.Enabled = False
    Alert.CloseAlert CInt(MP.Tag), True
    cmdLoadMP.Enabled = True
    cmdLoadMP.SetFocus
End Sub

Private Sub lblAlert_Click(Index As Integer)
    If Index = 2 Then picCommand_Click
End Sub

Private Sub NewsBar_CycleComplete(Cycle As Long)
    If Cycle = NewsBar.Cycles Then Alert.CloseAlert CInt(NewsBar.Tag), True
End Sub

Private Sub NoteBook_ButtonClick()
    Alert.CloseAlert CInt(NoteBook.Tag), True
End Sub

Private Sub picBACK1_Click()
    picBACK1.BackColor = GetColor
End Sub

Private Sub picBACK2_Click()
    picBACK2.BackColor = GetColor
End Sub

Private Sub picFORE_Click()
    picFORE.BackColor = GetColor
End Sub

Private Sub picLINK_Click()
    picLINK.BackColor = GetColor
End Sub

Private Sub picMASK_Click()
    picMASK.BackColor = GetColor
End Sub

Private Sub picCommand_Click()
    Alert.CloseAlert CInt(picAlert.Tag), True
End Sub

Private Sub Timer_Timer()

    Static Count As Integer, bCounting As Boolean
    
    If Not bCounting Then
        Count = 4
        bCounting = True
        Alert.play App.Path & "\Switch.wav"
    Else
        Count = Count - 1
        If Count < 0 Then
            Count = 5
            bCounting = False
            Timer.Enabled = False
            Alert.CloseAlert CInt(picCounter.Tag), True
        Else
            Alert.play App.Path & "\Switch.wav"
        End If
    End If
    
    lblCount = Count

End Sub

Private Function TriggerAlert(Props As taProperties, Optional hWndControl As Long, Optional sTag As String) As Integer
On Error GoTo errTriggerAlert

    Alert.Trigger Props, hWndControl, sTag
    Exit Function

errTriggerAlert:
LogError
End Function

Private Function LoadContainerAlert(Props As taProperties, hWndControl As Long, Optional sTag As String) As Integer
On Error GoTo errLoadContainerAlert

    Alert.LoadContainer Props, hWndControl, sTag
    Exit Function

errLoadContainerAlert:
LogError
End Function

Private Sub LogError()
    Beep
    txtLog = txtLog & "Error # [" & Err.Number & "] : " & Err.Description & vbCrLf
    txtLog.SelStart = Len(txtLog)
    txtLog.Refresh
End Sub

Private Sub LogEvent(Key As Integer, Tag As String, Message As String)
    txtLog = txtLog & "Alert [KEY=" & Key & "]"
    If Tag <> "" Then txtLog = txtLog & " Tagged As """ & Tag & """"
    txtLog = txtLog & " : " & Message & vbCrLf
    txtLog.SelStart = Len(txtLog)
    txtLog.Refresh
End Sub

Private Function GetColor() As OLE_COLOR
    Dlg.ShowColor
    GetColor = Dlg.Color
End Function

Private Function GetFile(sFilter As String) As String
    Dlg.Filter = sFilter
    Dlg.ShowOpen
    GetFile = Dlg.FileName
    Dlg.FileName = ""
End Function

Private Sub txtAngle_Change()

    Dim i As Long
    
    i = txtAngle.SelStart
    If IsNumeric(txtAngle) Then
        txtAngle = CLng(txtAngle)
        If CLng(txtAngle) < 0 Then txtAngle = 0
        If CLng(txtAngle) > 360 Then txtAngle = 360
    Else
        txtAngle = 270
    End If
    txtAngle.SelStart = i

End Sub

Private Sub txtDuration_Change()

    Dim i As Long
    
    i = txtDuration.SelStart
    If IsNumeric(txtDuration) Then
        txtDuration = CLng(txtDuration)
        If CLng(txtDuration) < 0 Then txtDuration = 0
    Else
        txtDuration = 5000
    End If
    txtDuration.SelStart = i

End Sub

Private Sub txtMargin_Change()

    Dim i As Long
    
    i = txtMargin.SelStart
    If IsNumeric(txtMargin) Then
        txtMargin = CLng(txtMargin)
        If CLng(txtMargin) < 0 Then txtMargin = 0
    Else
        txtMargin = 400
    End If
    txtMargin.SelStart = i

End Sub

Private Sub txtOffset_Change()

    Dim i As Long
    
    i = txtOffset.SelStart
    If IsNumeric(txtOffset) Then
        txtOffset = CLng(txtOffset)
        If CLng(txtOffset) < 0 Then txtOffset = 0
    Else
        txtOffset = 400
    End If
    txtOffset.SelStart = i

End Sub

Private Sub txtSpeed_Change()

    Dim i As Long
    
    i = txtSpeed.SelStart
    If IsNumeric(txtSpeed) Then
        txtSpeed = CLng(txtSpeed)
        If CLng(txtSpeed) < 0 Then txtSpeed = 0
        If CLng(txtSpeed) > 25 Then txtSpeed = 25
    Else
        txtSpeed = 0
    End If
    txtSpeed.SelStart = i

End Sub

Private Sub txtTransparency_Change()

    Dim i As Long
    
    i = txtTransparency.SelStart
    If IsNumeric(txtTransparency) Then
        txtTransparency = CLng(txtTransparency)
        If CLng(txtTransparency) < 0 Then txtTransparency = 0
        If CLng(txtTransparency) > 255 Then txtTransparency = 255
    Else
        txtTransparency = 255
    End If
    txtTransparency.SelStart = i

End Sub
