Attribute VB_Name = "modAnimations"
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Sub RollUP(picObject As PictureBox, Props As taProperties, bFadeIN As Boolean)
On Error Resume Next

    Dim Step As Single
    Dim Action As LayerFlags
    Dim AlertHeight As Single, AlertTop As Single
    
    AlertHeight = picObject.Height
    picObject.Height = 0
    picObject.Top = picObject.Top + AlertHeight - picObject.Height
    AlertTop = picObject.Top + picObject.Height - AlertHeight
    Step = 5 * Screen.TwipsPerPixelY
    
    With Props
        If bFadeIN Then
            If .UseMask Then Action = SetBoth Else Action = SetAlpha
            Do While picObject.Height < AlertHeight - Step
                picObject.Top = picObject.Top - Step
                picObject.Height = picObject.Height + Step
                SetLayerAttributes picObject.hWnd, .MaskColor, (picObject.Height / AlertHeight) * .Transparency, Action
                picObject.Refresh
                Sleep .Speed
            Loop
            SetLayerAttributes picObject.hWnd, .MaskColor, .Transparency, Action
        Else
            If .UseMask Then
                SetLayerAttributes picObject.hWnd, .MaskColor, .Transparency, SetBoth
            ElseIf .Transparency <> 255 Then
                SetLayerAttributes picObject.hWnd, .MaskColor, .Transparency, SetAlpha
            End If
            picObject.Visible = True
            Do While picObject.Height < AlertHeight - Step
                picObject.Top = picObject.Top - Step
                picObject.Height = picObject.Height + Step
                picObject.Refresh
                Sleep .Speed
            Loop
        End If
    End With
    
    picObject.Top = AlertTop
    picObject.Height = AlertHeight
    picObject.Refresh

End Sub

Public Sub RollDown(picObject As PictureBox, Props As taProperties, bFadeOUT As Boolean)
On Error Resume Next

    Dim Step As Single
    Dim Action As LayerFlags
    Dim AlertHeight As Single
    
    Step = 5 * Screen.TwipsPerPixelY
    AlertHeight = picObject.Height
    
    With Props
        If bFadeOUT Then
            If .UseMask Then Action = SetBoth Else Action = SetAlpha
            Do While picObject.Height > Step
                picObject.Height = picObject.Height - Step
                picObject.Top = picObject.Top + Step
                SetLayerAttributes picObject.hWnd, .MaskColor, (picObject.Height / AlertHeight) * .Transparency, Action
                picObject.Refresh
                Sleep .Speed
            Loop
            SetLayerAttributes picObject.hWnd, .MaskColor, 0, Action
        Else
            Do While picObject.Height > Step
                picObject.Height = picObject.Height - Step
                picObject.Top = picObject.Top + Step
                picObject.Refresh
                Sleep .Speed
            Loop
        End If
    End With
    
    picObject.Visible = False

End Sub

Public Sub FadeIN(picObject As PictureBox, Props As taProperties, bReverse As Boolean)
On Error Resume Next

    Dim hWndAlert As Long
    Dim Action As LayerFlags
    Dim i As Integer, StepValue As Integer
    Dim StartValue As Byte, FinalValue As Byte
    
    hWndAlert = picObject.hWnd
    
    With Props
    
        If .UseMask Then Action = SetBoth Else Action = SetAlpha
        
        If bReverse Then
            StartValue = .Transparency
            FinalValue = 0
            StepValue = -5
        Else
            StartValue = 0
            FinalValue = .Transparency
            StepValue = 5
        End If
        
        For i = StartValue To FinalValue Step StepValue
            SetLayerAttributes hWndAlert, .MaskColor, CByte(i), Action
            Sleep .Speed
        Next
        
        SetLayerAttributes hWndAlert, .MaskColor, FinalValue, Action
    
    End With

End Sub
