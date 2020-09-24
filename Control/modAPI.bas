Attribute VB_Name = "modAPI"
Option Explicit

Public Enum LayerFlags
    SetAlpha = &H2
    SetMask = &H1
    SetBoth = &H3
End Enum

Private Type apiRECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Const GWL_EXSTYLE As Long = &HFFEC
Private Const WS_EX_LAYERED As Long = &H80000
Private Const WS_EX_TOOLWINDOW As Long = &H80
Private Const HWND_TOPMOST As Long = &HFFFF
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOMOVE As Long = &H2
Private Const BF_RECT As Long = &HF
Private Const EDGE_RAISED As Long = &H5
Private Const EDGE_SUNKEN As Long = &HA
Private Const SPI_GETWORKAREA As Long = &H30
Private Const SW_SHOW As Long = &H5

Private Declare Sub SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function SetParent Lib "User32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "User32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function SetRect Lib "User32" (lpRect As apiRECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetWindowRect Lib "User32" (ByVal hWnd As Long, lpRect As apiRECT) As Long
Private Declare Function GetDesktopWindow Lib "User32" () As Long
Private Declare Function DrawEdge Lib "User32" (ByVal hDC As Long, qrc As apiRECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function IsWindow Lib "User32" (ByVal hWnd As Long) As Long
Private Declare Function ShowWindow Lib "User32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function MoveWindow Lib "User32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function SystemParametersInfo Lib "User32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long

Public Sub CreateAlert(hWndAlert As Long, hWndControl As Long, bLayered As Boolean, bTopMost As Boolean)

    Dim dwStyle As Long
    
    SetParent hWndAlert, GetDesktopWindow
    If hWndControl <> 0 Then SetParent hWndControl, hWndAlert
    
    dwStyle = GetWindowLong(hWndAlert, GWL_EXSTYLE)
    dwStyle = dwStyle Or WS_EX_TOOLWINDOW
    If bLayered Then dwStyle = dwStyle Or WS_EX_LAYERED
    SetWindowLong hWndAlert, GWL_EXSTYLE, dwStyle
    
    If bLayered Then
        SetLayerAttributes hWndAlert, 0, 0, SetAlpha
        ShowWindow hWndAlert, SW_SHOW
    End If
    
    If bTopMost Then SetTopMostWindow hWndAlert

End Sub

Public Sub DrawBorders(picObject As PictureBox)

    Dim R As apiRECT
    
    SetRect R, 0, 0, picObject.Width / Screen.TwipsPerPixelX, picObject.Height / Screen.TwipsPerPixelY
    DrawEdge picObject.hDC, R, EDGE_RAISED, BF_RECT
    SetRect R, 2, 2, picObject.Width / Screen.TwipsPerPixelX - 2, picObject.Height / Screen.TwipsPerPixelY - 2
    DrawEdge picObject.hDC, R, EDGE_SUNKEN, BF_RECT
    picObject.Refresh

End Sub

Public Sub SetLayerAttributes(hWnd As Long, MaskColor As Long, Alpha As Byte, Flags As LayerFlags)
On Error Resume Next
    SetLayeredWindowAttributes hWnd, MaskColor, Alpha, Flags
End Sub

Public Sub SetTopMostWindow(hWnd As Long)
    SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End Sub

Public Sub ExecuteLink(sLink As String)
    ShellExecute 0, vbNullString, sLink, vbNullString, "C:\", 0
End Sub

Public Function IsInvalidWindow(hWnd As Long) As Boolean
    IsInvalidWindow = (IsWindow(hWnd) = 0) Or (hWnd = GetDesktopWindow)
End Function

Public Function TaskBarTop() As Single
    Dim R As apiRECT
    SystemParametersInfo SPI_GETWORKAREA, 0, R, 0
    TaskBarTop = R.Bottom * Screen.TwipsPerPixelY
End Function

Public Function WindowWidth(hWnd As Long) As Single
    Dim R As apiRECT
    GetWindowRect hWnd, R
    WindowWidth = (R.Right - R.Left) * Screen.TwipsPerPixelX
End Function

Public Function WindowHeight(hWnd As Long) As Single
    Dim R As apiRECT
    GetWindowRect hWnd, R
    WindowHeight = (R.Bottom - R.Top) * Screen.TwipsPerPixelY
End Function

Public Sub WindowMove(hWnd As Long, Left As Single, Top As Single)
    Dim R As apiRECT
    GetWindowRect hWnd, R
    MoveWindow hWnd, Left / Screen.TwipsPerPixelX, Top / Screen.TwipsPerPixelY, R.Right - R.Left, R.Bottom - R.Top, True
End Sub

Public Function LayersSupported() As Boolean
On Error GoTo errLayersSupported
    SetLayeredWindowAttributes 0, 0, 0, 0
    LayersSupported = True
errLayersSupported:
End Function
