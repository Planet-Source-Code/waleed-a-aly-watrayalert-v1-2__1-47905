VERSION 5.00
Begin VB.UserControl NewsBar 
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   510
   LockControls    =   -1  'True
   ScaleHeight     =   495
   ScaleWidth      =   510
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "NewsBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Property Variables
Private mNews As String
Private mFont As StdFont
Private mSpeed As Byte
Private mBackColor As Long
Private mForeColor As Long
Private mCycles As Long

'Defaults
Private Const defSpeed As Long = 0
Private Const defCycles As Long = 1
Private Const defBackColor As Long = vbBlack
Private Const defForeColor As Long = vbGreen
Private Const defNews As String = "NewsBar Control Coded by: Waleed A. Aly"

'Locals
Private bScroll As Boolean
Private sDisplayText As String

'Events
Public Event CycleComplete(Cycle As Long)

'APIs
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Sub UserControl_InitProperties()
    mNews = defNews
    Set mFont = Ambient.Font
    mBackColor = defBackColor
    mForeColor = defForeColor
    mSpeed = defSpeed
    mCycles = defCycles
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mNews = PropBag.ReadProperty("News", defNews)
    Set mFont = PropBag.ReadProperty("Font", Ambient.Font)
    mBackColor = PropBag.ReadProperty("BackColor", defBackColor)
    mForeColor = PropBag.ReadProperty("ForeColor", defForeColor)
    mSpeed = PropBag.ReadProperty("Speed", defSpeed)
    mCycles = PropBag.ReadProperty("Cycles", defCycles)
    sDisplayText = Replace(mNews, vbCrLf, "  -  ")
    ApplyAppearance
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "News", mNews
    PropBag.WriteProperty "Font", mFont
    PropBag.WriteProperty "BackColor", mBackColor
    PropBag.WriteProperty "ForeColor", mForeColor
    PropBag.WriteProperty "Speed", mSpeed
    PropBag.WriteProperty "Cycles", mCycles
    sDisplayText = Replace(mNews, vbCrLf, "  -  ")
    ApplyAppearance
End Sub

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get News() As String
    News = mNews
End Property

Public Property Let News(Value As String)
    mNews = Value
    PropertyChanged "News"
    sDisplayText = Replace(Value, vbCrLf, "  -  ")
End Property

Public Property Get Speed() As Byte
    Speed = mSpeed
End Property

Public Property Let Speed(Value As Byte)
    mSpeed = Value
    PropertyChanged "Speed"
End Property

Public Property Get Cycles() As Long
    Cycles = mCycles
End Property

Public Property Let Cycles(Value As Long)
    mCycles = Value
    PropertyChanged "Cycles"
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = mForeColor
End Property

Public Property Let ForeColor(Value As OLE_COLOR)
    mForeColor = Value
    PropertyChanged "ForeColor"
    UserControl.ForeColor = mForeColor
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = mBackColor
End Property

Public Property Let BackColor(Value As OLE_COLOR)
    mBackColor = Value
    PropertyChanged "BackColor"
    UserControl.BackColor = mBackColor
End Property

Public Property Get Font() As StdFont
    Set Font = mFont
End Property

Public Property Set Font(Value As StdFont)
    Set mFont = Value
    PropertyChanged "Font"
    Set UserControl.Font = mFont
End Property

Public Sub ApplyAppearance()
    Set UserControl.Font = mFont
    UserControl.ForeColor = mForeColor
    UserControl.BackColor = mBackColor
End Sub

Public Sub AutoSize()
    UserControl.Width = Screen.Width
    UserControl.Height = UserControl.TextHeight(sDisplayText) + 100
End Sub

Public Sub StartScrolling()
    Timer.Enabled = True
    bScroll = True
End Sub

Public Sub StopScrolling()
    UserControl.Cls
    bScroll = False
End Sub

Private Sub Timer_Timer()
    Timer.Enabled = False
    Scroll
End Sub

Private Sub Scroll()

    Dim i As Long
    Dim TC As Long
    Dim CX As Single
    
    For i = 1 To mCycles
        UserControl.CurrentX = UserControl.Width
        Do Until UserControl.CurrentX < -UserControl.TextWidth(sDisplayText)
            If Not bScroll Then Exit Sub
            CX = UserControl.CurrentX
            UserControl.Cls
            UserControl.CurrentY = 50
            UserControl.CurrentX = CX - 20
            UserControl.Print sDisplayText
            UserControl.CurrentX = CX - 20
            TC = GetTickCount
            Do Until GetTickCount > TC + mSpeed
                DoEvents
            Loop
        Loop
        RaiseEvent CycleComplete(i)
    Next

End Sub
