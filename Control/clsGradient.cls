VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGradient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'----------------------------------------------------
'Class Written by       : Kath-Rock
'Min OS Requirements    : Windows NT 3.1, Windows 95
'----------------------------------------------------

'Property Variables
Private mlColor1 As Long
Private mlColor2 As Long
Private mfAngle  As Single

'Property Defaults
Private Const mlDefColor1 As Long = &HFFFFD0
Private Const mlDefColor2 As Long = &H400000
Private Const mfDefAngle  As Single = 315

'Misc Constants
Private Const PI   As Double = 3.14159265358979
Private Const RADS As Double = PI / 180

'Types
Private Type PointSng
    X As Single
    Y As Single
End Type

Private Type PointAPI
    X As Long
    Y As Long
End Type

Private Type RectAPI
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type

'APIs
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As PointAPI) As Long
Private Declare Function GetSysColor Lib "User32" (ByVal nIndex As Long) As Long
Private Declare Function GetClientRect Lib "User32" (ByVal hWnd As Long, lpRect As RectAPI) As Long
Private Declare Function LockWindowUpdate Lib "User32" (ByVal hWndLock As Long) As Long

Private Sub Class_Initialize()
    mlColor1 = mlDefColor1
    mlColor2 = mlDefColor2
    mfAngle = mfDefAngle
End Sub

Public Function Draw(picObj As Object) As Boolean
On Error GoTo LocalError
'Note: This class uses API functions to draw. If the
'      destination object is in AutoRedraw mode, the
'      Refresh method for that object must be invoked.
'picObj can be a Form or a PictureBox.

    Dim i     As Long
    Dim lRet  As Long
    Dim uRect As RectAPI
    
    lRet = LockWindowUpdate(picObj.hWnd)
    lRet = GetClientRect(picObj.hWnd, uRect)
    
    If lRet <> 0 Then
        If uRect.Right > 1 And uRect.Bottom > 1 Then
            i = DrawGradient(picObj.hDC, uRect.Right, uRect.Bottom)
            Draw = (i > 0)
        End If
    End If

NormalExit:
    lRet = LockWindowUpdate(0)
    Exit Function

LocalError:
    Resume NormalExit

End Function

Public Function BlendColors(ByVal lColor1 As Long, ByVal lColor2 As Long, ByVal lSteps As Long, laRetColors() As Long) As Long

    Dim i       As Long
    Dim lRed    As Long
    Dim lGrn    As Long
    Dim lBlu    As Long
    Dim fRedStp As Single
    Dim fGrnStp As Single
    Dim fBluStp As Single
    
    If lSteps < 2 Then lSteps = 2
    
    'Extract Red, Blue and Green values
    lRed = (lColor1 And &HFF&)
    lGrn = (lColor1 And &HFF00&) / &H100
    lBlu = (lColor1 And &HFF0000) / &H10000
    
    'Find the amount of change for each color element per color change.
    fRedStp = Div(CSng((lColor2 And &HFF&) - lRed), CSng(lSteps))
    fGrnStp = Div(CSng(((lColor2 And &HFF00&) / &H100&) - lGrn), CSng(lSteps))
    fBluStp = Div(CSng(((lColor2 And &HFF0000) / &H10000) - lBlu), CSng(lSteps))
    
    'Create the colors
    ReDim laRetColors(lSteps - 1)
    laRetColors(0) = lColor1            'First Color
    laRetColors(lSteps - 1) = lColor2   'Last Color
    For i = 1 To lSteps - 2             'All Colors between
        laRetColors(i) = CLng(lRed + (fRedStp * i)) + _
            (CLng(lGrn + (fGrnStp * i)) * &H100&) + _
            (CLng(lBlu + (fBluStp * i)) * &H10000)
    Next
    
    BlendColors = lSteps

End Function

Private Function DrawGradient(ByVal hDC As Long, ByVal lWidth As Long, ByVal lHeight As Long) As Long
On Error GoTo LocalError

    Dim i          As Long
    Dim bDone      As Boolean
    Dim iIncX      As Integer
    Dim iIncY      As Integer
    Dim lRet       As Long
    Dim hPen       As Long
    Dim hOldPen    As Long
    Dim lPointCnt  As Long
    Dim laColors() As Long
    Dim fMovX      As Single
    Dim fMovY      As Single
    Dim fDist      As Single
    Dim fAngle     As Single
    Dim fLongSide  As Single
    Dim uTmpPt     As PointAPI
    Dim uaPts()    As PointAPI
    Dim uaTmpPts() As PointSng
    
    'Start with center of rect
    ReDim uaTmpPts(2)
    uaTmpPts(2).X = Int(lWidth / 2)
    uaTmpPts(2).Y = Int(lHeight / 2)
    
    'Calc distance to furthest edge as if rect were square
    fLongSide = IIf(lWidth > lHeight, lWidth, lHeight)
    fDist = (Sqr((fLongSide ^ 2) + (fLongSide ^ 2)) + 2) / 2
    
    'Create points to the left and the right at a 0º angle (horizontal)
    uaTmpPts(0).X = uaTmpPts(2).X - fDist
    uaTmpPts(0).Y = uaTmpPts(2).Y
    uaTmpPts(1).X = uaTmpPts(2).X + fDist
    uaTmpPts(1).Y = uaTmpPts(2).Y
    
    'Lines will be drawn perpendicular to mfAngle so
    'add 90º and correct for 360º wrap
    fAngle = CDbl(mfAngle + 90) - Int(Int(CDbl(mfAngle + 90) / 360#) * 360#)
    
    'Rotate second and third points to fAngle
    Call RotatePoint(uaTmpPts(2), uaTmpPts(0), fAngle)
    Call RotatePoint(uaTmpPts(2), uaTmpPts(1), fAngle)
    
    'We now have a line that crosses the center and
    'two sides of the rect at the correct angle.
    
    'Calc the starting quadrant, direction of and amount of first move
    '(fMovX, fMovY moves line from center to starting edge)
    'and direction of each incremental move (iIncX, iIncY).
    Select Case mfAngle
        Case 0 To 90
            'Left Bottom
            If Abs(uaTmpPts(0).X - uaTmpPts(1).X) <= Abs(uaTmpPts(0).Y - uaTmpPts(1).Y) Then
                'Move line to left edge; Draw left to right
                fMovX = IIf(uaTmpPts(0).X > uaTmpPts(1).X, -uaTmpPts(0).X, -uaTmpPts(1).X)
                fMovY = 0
                iIncX = 1
                iIncY = 0
            Else
                'Move line to bottom edge; Draw bottom to top
                fMovX = 0
                fMovY = IIf(uaTmpPts(0).Y > uaTmpPts(1).Y, lHeight - uaTmpPts(1).Y, lHeight - uaTmpPts(0).Y)
                iIncX = 0
                iIncY = -1
            End If
        Case 90 To 180
            'Right Bottom
            If Abs(uaTmpPts(0).X - uaTmpPts(1).X) <= Abs(uaTmpPts(0).Y - uaTmpPts(1).Y) Then
                'Move line to right edge; Draw right to left
                fMovX = IIf(uaTmpPts(0).X > uaTmpPts(1).X, lWidth - uaTmpPts(1).X, lWidth - uaTmpPts(0).X)
                fMovY = 0
                iIncX = -1
                iIncY = 0
            Else
                'Move line to bottom edge; Draw bottom to top
                fMovX = 0
                fMovY = IIf(uaTmpPts(0).Y > uaTmpPts(1).Y, lHeight - uaTmpPts(1).Y, lHeight - uaTmpPts(0).Y)
                iIncX = 0
                iIncY = -1
            End If
        Case 180 To 270
            'Right Top
            If Abs(uaTmpPts(0).X - uaTmpPts(1).X) <= Abs(uaTmpPts(0).Y - uaTmpPts(1).Y) Then
                'Move line to right edge; Draw right to left
                fMovX = IIf(uaTmpPts(0).X > uaTmpPts(1).X, lWidth - uaTmpPts(1).X, lWidth - uaTmpPts(0).X)
                fMovY = 0
                iIncX = -1
                iIncY = 0
            Else
                'Move line to top edge; Draw top to bottom
                fMovX = 0
                fMovY = IIf(uaTmpPts(0).Y > uaTmpPts(1).Y, -uaTmpPts(0).Y, -uaTmpPts(1).Y)
                iIncX = 0
                iIncY = 1
            End If
        Case Else   '(270 to 360)
            'Left Top
            If Abs(uaTmpPts(0).X - uaTmpPts(1).X) <= Abs(uaTmpPts(0).Y - uaTmpPts(1).Y) Then
                'Move line to left edge; Draw left to right
                fMovX = IIf(uaTmpPts(0).X > uaTmpPts(1).X, -uaTmpPts(0).X, -uaTmpPts(1).X)
                fMovY = 0
                iIncX = 1
                iIncY = 0
            Else
                'Move line to top edge; Draw top to bottom
                fMovX = 0
                fMovY = IIf(uaTmpPts(0).Y > uaTmpPts(1).Y, -uaTmpPts(0).Y, -uaTmpPts(1).Y)
                iIncX = 0
                iIncY = 1
            End If
    End Select
    
    'At this point we could calculate where the lines will cross the rect edges, but
    'this would slow things down. The picObj clipping region will take care of this.
    
    'Start with 1000 points and add more if needed. This increases
    'speed by not re-dimming the array in each loop.
    ReDim uaPts(999)
    
    'Set the first two points in the array
    uaPts(0).X = uaTmpPts(0).X + fMovX
    uaPts(0).Y = uaTmpPts(0).Y + fMovY
    uaPts(1).X = uaTmpPts(1).X + fMovX
    uaPts(1).Y = uaTmpPts(1).Y + fMovY
    
    i = 2
    'Create the rest of the points by incrementing both points
    'on each line iIncX, iIncY from the previous line's points.
    'Where we stop depends on the direction of travel.
    'We'll continue until both points in a set reach the end.
    While Not bDone
        uaPts(i).X = uaPts(i - 2).X + iIncX
        uaPts(i).Y = uaPts(i - 2).Y + iIncY
        i = i + 1
        Select Case True
            Case iIncX > 0  'Moving Left to Right
                bDone = uaPts(i - 1).X > lWidth And uaPts(i - 2).X > lWidth
            Case iIncX < 0  'Moving Right to Left
                bDone = uaPts(i - 1).X < 0 And uaPts(i - 2).X < 0
            Case iIncY > 0  'Moving Top to Bottom
                bDone = uaPts(i - 1).Y > lHeight And uaPts(i - 2).Y > lHeight
            Case iIncY < 0  'Moving Bottom to Top
                bDone = uaPts(i - 1).Y < 0 And uaPts(i - 2).Y < 0
        End Select
        If (i Mod 1000) = 0 Then
            ReDim Preserve uaPts(UBound(uaPts) + 1000)
        End If
    Wend
    
    'Free excess memory (may have 1001 points dimmed to 2000)
    ReDim Preserve uaPts(i - 1)
    
    'Create the array of colors blending from mlColor1 to mlColor2
    lRet = BlendColors(mlColor1, mlColor2, i / 2, laColors)
    
    'Now draw each line in it's own color
    For i = 0 To UBound(uaPts) - 1 Step 2
        'Move to next point
        lRet = MoveToEx(hDC, uaPts(i).X, uaPts(i).Y, uTmpPt)
        'Create the colored pen and select it into the DC
        hPen = CreatePen(0, 1, laColors(Int(i / 2)))
        hOldPen = SelectObject(hDC, hPen)
        'Draw the line
        lRet = LineTo(hDC, uaPts(i + 1).X, uaPts(i + 1).Y)
        'Get the pen back out of the DC and destroy it
        lRet = SelectObject(hDC, hOldPen)
        lRet = DeleteObject(hPen)
    Next i
    
    DrawGradient = i

NormalExit:
    Erase laColors, uaPts, uaTmpPts
    Exit Function

LocalError:
    DrawGradient = 0
    Resume

End Function

Private Sub RotatePoint(uAxisPt As PointSng, uRotatePt As PointSng, fDegrees As Single)

    Dim fDX         As Single
    Dim fDY         As Single
    Dim fRadians    As Single
    
    fRadians = fDegrees * RADS
    fDX = uRotatePt.X - uAxisPt.X
    fDY = uRotatePt.Y - uAxisPt.Y
    uRotatePt.X = uAxisPt.X + ((fDX * Cos(fRadians)) + (fDY * Sin(fRadians)))
    uRotatePt.Y = uAxisPt.Y + -((fDX * Sin(fRadians)) - (fDY * Cos(fRadians)))

End Sub

Private Function Div(ByVal dNumer As Double, ByVal dDenom As Double) As Double
    If dDenom <> 0 Then Div = dNumer / dDenom
End Function

Public Property Let Color1(ByVal lData As Long)

    Dim i As Long
    
    mlColor1 = lData
    If mlColor1 < 0 Then
        i = (mlColor1 And Not &H80000000)
        If i >= 0 And i <= 24 Then
            mlColor1 = GetSysColor(i)
        End If
    End If

End Property

Public Property Get Color1() As Long
    Color1 = mlColor1
End Property

Public Property Let Color2(ByVal lData As Long)

    Dim i As Long
    
    mlColor2 = lData
    If mlColor2 < 0 Then
        i = (mlColor2 And Not &H80000000)
        If i >= 0 And i <= 24 Then
            mlColor2 = GetSysColor(i)
        End If
    End If

End Property

Public Property Get Color2() As Long
    Color2 = mlColor2
End Property

Public Property Let Angle(ByVal fData As Single)
    mfAngle = CDbl(fData) - Int(Int(CDbl(fData) / 360) * 360)
End Property

Public Property Get Angle() As Single
    Angle = mfAngle
End Property
