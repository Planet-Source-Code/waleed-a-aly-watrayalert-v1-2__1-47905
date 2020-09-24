Attribute VB_Name = "modGeneral"
Option Explicit

Public Type RECT
    Left As Single
    Top As Single
    Right As Single
    Bottom As Single
End Type

Public Sub DrawText(picObject As PictureBox, sText As String, Alignment As AlignmentConstants, bInset As Boolean, R As RECT)
On Error Resume Next

    Dim i As Long
    Dim Parsed() As String
    Dim Color As Long, Offset As Long
    
    Color = picObject.ForeColor
    Parsed = Split(sText, vbCrLf)
    
    Select Case Alignment
        Case vbLeftJustify
            If bInset Then
                picObject.ForeColor = vbWhite
                picObject.CurrentY = R.Top + 10
                Offset = 10
                GoSub DrawAtLeft
                picObject.ForeColor = Color
            End If
            Offset = 0
            picObject.CurrentY = R.Top
            GoSub DrawAtLeft
            
        Case vbRightJustify
            If bInset Then
                picObject.ForeColor = vbWhite
                picObject.CurrentY = R.Top + 10
                Offset = 10
                GoSub DrawAtRight
                picObject.ForeColor = Color
            End If
            Offset = 0
            picObject.CurrentY = R.Top
            GoSub DrawAtRight
            
        Case vbCenter
            If bInset Then
                picObject.ForeColor = vbWhite
                picObject.CurrentY = R.Top + 10
                Offset = 10
                GoSub DrawCentered
                picObject.ForeColor = Color
            End If
            Offset = 0
            picObject.CurrentY = R.Top
            GoSub DrawCentered
    End Select
    
    Exit Sub

DrawAtLeft:
    For i = LBound(Parsed) To UBound(Parsed)
        picObject.CurrentX = R.Left + Offset
        picObject.Print Parsed(i)
    Next
    Return

DrawAtRight:
    For i = LBound(Parsed) To UBound(Parsed)
        picObject.CurrentX = R.Right - picObject.TextWidth(Parsed(i)) + Offset
        picObject.Print Parsed(i)
    Next
    Return

DrawCentered:
    For i = LBound(Parsed) To UBound(Parsed)
        picObject.CurrentX = R.Left + (R.Right - R.Left - picObject.TextWidth(Parsed(i))) / 2 + Offset
        picObject.Print Parsed(i)
    Next
    Return

End Sub

Public Sub DrawImage(picObject As PictureBox, Image As StdPicture, R As RECT, bStretch As Boolean)
On Error Resume Next

    If bStretch Then
        picObject.PaintPicture Image, R.Left, R.Top, R.Right - R.Left, R.Bottom - R.Top
    Else
        picObject.PaintPicture Image, R.Left, R.Top
    End If

End Sub

Public Function DrawGrad(picObject As Object, Color1 As Long, Color2 As Long, Angle As Single) As Boolean
On Error Resume Next

    Dim Brush As New clsGradient
    
    Brush.Color1 = Color1
    Brush.Color2 = Color2
    Brush.Angle = Angle
    DrawGrad = Brush.Draw(picObject)
    picObject.Refresh
    Set Brush = Nothing

End Function

Public Function IsInsideRECT(X As Single, Y As Single, R As RECT) As Boolean
    If X > R.Left And X < R.Right And Y > R.Top And Y < R.Bottom Then IsInsideRECT = True
End Function
