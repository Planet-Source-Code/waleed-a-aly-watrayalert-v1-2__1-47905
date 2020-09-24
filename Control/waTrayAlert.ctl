VERSION 5.00
Begin VB.UserControl waTrayAlert 
   ClientHeight    =   1830
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2040
   InvisibleAtRuntime=   -1  'True
   LockControls    =   -1  'True
   Picture         =   "waTrayAlert.ctx":0000
   PropertyPages   =   "waTrayAlert.ctx":04F5
   ScaleHeight     =   1830
   ScaleWidth      =   2040
   Tag             =   "By: Waleed A. Aly"
   ToolboxBitmap   =   "waTrayAlert.ctx":0503
   Begin VB.PictureBox picAlert 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1575
      Index           =   0
      Left            =   120
      MouseIcon       =   "waTrayAlert.ctx":0815
      ScaleHeight     =   1575
      ScaleWidth      =   1815
      TabIndex        =   0
      Tag             =   "By: Waleed A. Aly"
      Top             =   120
      Visible         =   0   'False
      Width           =   1815
      Begin VB.PictureBox picClose 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   180
         Left            =   1560
         Picture         =   "waTrayAlert.ctx":14DF
         ScaleHeight     =   180
         ScaleWidth      =   195
         TabIndex        =   1
         Top             =   60
         Width           =   195
      End
      Begin VB.Timer Timer 
         Index           =   0
         Left            =   60
         Top             =   60
      End
   End
End
Attribute VB_Name = "waTrayAlert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

Public Enum taErrors
    UNEXPECTED_ERROR = vbObjectError + 60
    INVALID_PROPERTY = vbObjectError + 61
    INVALID_KEY = vbObjectError + 62
    INVALID_CONTROL = vbObjectError + 63
    CONTROL_IN_USE = vbObjectError + 64
    WAVE_NOT_FOUND = vbObjectError + 65
End Enum

Public Type taProperties

    '+ Appearance
    Alignment As AlignmentConstants
    BackColor1 As Long
    BackColor2 As Long
    Borderless As Boolean
    CloseButton As Boolean
    ForeColor As Long
    GradientAngle As Single
    Icon As StdPicture
    LinkColor As Long
    Margin As Single
    MaskColor As Long
    Offset As Single
    Picture As StdPicture
    Stretch As Boolean
    TextInset As Boolean
    Transparency As Byte
    UseMask As Boolean
    
    '+ Behavior
    AlertStyle As taAlertStyles
    AlwaysOnTop As Boolean
    Animation As taAnimations
    AutoSize As taAutoSize
    Duration As Long
    Persisting As Boolean
    Speed As Byte
    
    '+ Font
    Font As StdFont
    
    '+ Text
    Caption As String
    Link As String
    WaveSound As String

End Type

Public Enum taAlertStyles
    taOriginal
    taLink
End Enum

Public Enum taAnimations
    taRoll
    taFade
    taBoth
End Enum

Public Enum taAutoSize
    taDefault
    taPictureSize
End Enum

Public Enum taWaveList
    taMS_NewAlert
    taMS_NewEMail
    taMS_Online
    taMS_Type
    taMSN_NewAlert
    taMSN_NewEMail
    taMSN_Online
    taMSN_Ring
    taMSN_Type
End Enum

Public Enum taUnloadModes
    taAlertExpired
    taCloseButtonClicked
    taClearingAllAlerts
    taCodeInvoked
End Enum

Private Type InstSpace
    Caption As RECT
    Client As RECT
    Close As RECT
    Icon As RECT
End Type

Private Type AlertData
    hWndControl As Long
    Props As taProperties
    Space As InstSpace
    Tag As String
End Type

'About form
Private frm As frmAbout

'Control Infos
Private Count As Integer
Private TopLayer() As Integer
Private mProps As taProperties

'Alert Data
Private Alert() As AlertData

'Alert Identifiers
Private Const StandardAlertID As String = "STANDARD"
Private Const ContainerAlertID As String = "CONTAINER"

'Property Defaults
Private Const defBackColor1 As Long = &HC08080
Private Const defBackColor2 As Long = &HFFFFFF
Private Const defBorderless As Boolean = False
Private Const defCloseButton As Boolean = True
Private Const defForeColor As Long = &H0
Private Const defGradientAngle As Single = 270
Private Const defLinkColor As Long = &HC00000
Private Const defMargin As Single = 400
Private Const defMaskColor As Long = &HFFFFFF
Private Const defOffset As Single = 400
Private Const defStretch As Boolean = False
Private Const defTextInset As Boolean = False
Private Const defTransparency As Byte = 255
Private Const defUseMask As Boolean = False
Private Const defAlwaysOnTop As Boolean = True
Private Const defDuration As Long = 5000
Private Const defPersisting As Boolean = False
Private Const defSpeed As Byte = 0
Private Const defCaption As String = "By: Waleed A. Aly"

'Dynamic Property Defaults
Private defLink As String
Private defWaveSound As String
Private Const sEMail As String = "wa_aly@hotmail.com"

'Control Events
Public Event AlertClick(ByVal Key As Integer, ByVal Tag As String)
Attribute AlertClick.VB_Description = "Occurs when the client area of the alert is clicked."
Public Event CaptionClick(ByVal Key As Integer, ByVal Tag As String)
Attribute CaptionClick.VB_Description = "Occurs when the caption area of the alert is clicked."
Public Event Created(ByVal Key As Integer, ByVal Tag As String, ByVal hWndAlert As Long)
Attribute Created.VB_Description = "Occurs when the alert is created and is about to be animated on."
Public Event IconClick(ByVal Key As Integer, ByVal Tag As String)
Attribute IconClick.VB_Description = "Occurs when the icon area of the alert is clicked."
Public Event Loaded(ByVal Key As Integer, ByVal Tag As String)
Attribute Loaded.VB_Description = "Occurs when the alert has finished loading."
Public Event QueryUnload(ByVal Key As Integer, ByVal Tag As String, ByVal UnloadMode As taUnloadModes, ByRef Cancel As Boolean)
Attribute QueryUnload.VB_Description = "Occurs when the alert is about to unload."
Public Event Unloaded(ByVal Key As Integer, ByVal Tag As String)
Attribute Unloaded.VB_Description = "Occurs when the alert has finished unloading."

'APIs
Private Declare Function GetParent Lib "User32" (ByVal hWnd As Long) As Long
Private Declare Function SetParent Lib "User32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Private Sub picAlert_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    With Alert(Index)
    
        If picAlert(Index).Tag = StandardAlertID Then
            If .Props.AlertStyle = taLink And .Props.Caption <> "" Then
                If IsInsideRECT(X, Y, .Space.Caption) Then
                    If picAlert(Index).MousePointer = vbCustom Then Exit Sub
                    picAlert(Index).MousePointer = vbCustom
                    picAlert(Index).ForeColor = .Props.LinkColor
                Else
                    If picAlert(Index).MousePointer = vbDefault Then Exit Sub
                    picAlert(Index).MousePointer = vbDefault
                    picAlert(Index).ForeColor = .Props.ForeColor
                End If
            Else
                Exit Sub
            End If
        Else
            Exit Sub
        End If
        
        DrawText picAlert(Index), .Props.Caption, .Props.Alignment, .Props.TextInset, .Space.Caption
    
    End With

End Sub

Private Sub picAlert_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    Dim bCancel As Boolean
    Dim TempAlertData As AlertData
    
    TempAlertData = Alert(Index)
    
    With TempAlertData
        If IsInsideRECT(X, Y, .Space.Caption) Then
            If .Props.AlertStyle = taLink Then ExecuteLink .Props.Link
            RaiseEvent CaptionClick(Index, .Tag)
            Exit Sub
        End If
        If IsInsideRECT(X, Y, .Space.Close) Then
            RaiseEvent QueryUnload(Index, .Tag, taCloseButtonClicked, bCancel)
            If Not bCancel Then UnloadAlert Index
            Exit Sub
        End If
        If IsInsideRECT(X, Y, .Space.Icon) Then
            RaiseEvent IconClick(Index, .Tag)
            Exit Sub
        End If
        If IsInsideRECT(X, Y, .Space.Client) Then RaiseEvent AlertClick(Index, .Tag)
    End With

End Sub

Private Sub Timer_Timer(Index As Integer)

    Dim bCancel As Boolean
    
    RaiseEvent QueryUnload(Index, Alert(Index).Tag, taAlertExpired, bCancel)
    If Not bCancel Then AnimateOFF Index: UnloadAlert Index

End Sub

Private Sub UserControl_Initialize()
    ReDim Alert(0 To 0) As AlertData
    ReDim TopLayer(0 To 0) As Integer
    defWaveSound = WavePath(taMSN_Type)
    defLink = "mailto:" & sEMail & "?subject=waTrayAlert v" & App.Major & "." & App.Minor
End Sub

Private Sub UserControl_InitProperties()
    Call AboutBox
    mProps = DefaultProps
End Sub

Private Sub UserControl_Resize()
    UserControl.Width = 32 * Screen.TwipsPerPixelX
    UserControl.Height = 32 * Screen.TwipsPerPixelY
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With mProps
        .Alignment = PropBag.ReadProperty("Alignment", vbCenter)
        .BackColor1 = PropBag.ReadProperty("BackColor1", defBackColor1)
        .BackColor2 = PropBag.ReadProperty("BackColor2", defBackColor2)
        .Borderless = PropBag.ReadProperty("Borderless", defBorderless)
        .CloseButton = PropBag.ReadProperty("CloseButton", defCloseButton)
        .ForeColor = PropBag.ReadProperty("ForeColor", defForeColor)
        .GradientAngle = PropBag.ReadProperty("GradientAngle", defGradientAngle)
        Set .Icon = PropBag.ReadProperty("Icon", Nothing)
        .LinkColor = PropBag.ReadProperty("LinkColor", defLinkColor)
        .Margin = PropBag.ReadProperty("Margin", defMargin)
        .MaskColor = PropBag.ReadProperty("MaskColor", defMaskColor)
        .Offset = PropBag.ReadProperty("Offset", defOffset)
        Set .Picture = PropBag.ReadProperty("Picture", Nothing)
        .Stretch = PropBag.ReadProperty("Stretch", defStretch)
        .TextInset = PropBag.ReadProperty("TextInset", defTextInset)
        .Transparency = PropBag.ReadProperty("Transparency", defTransparency)
        .UseMask = PropBag.ReadProperty("UseMask", defUseMask)
        .AlertStyle = PropBag.ReadProperty("AlertStyle", 1)
        .AlwaysOnTop = PropBag.ReadProperty("AlwaysOnTop", defAlwaysOnTop)
        .Animation = PropBag.ReadProperty("Animation", 0)
        .AutoSize = PropBag.ReadProperty("AutoSize", 0)
        .Duration = PropBag.ReadProperty("Duration", defDuration)
        .Persisting = PropBag.ReadProperty("Persisting", defPersisting)
        .Speed = PropBag.ReadProperty("Speed", defSpeed)
        Set .Font = PropBag.ReadProperty("Font", Ambient.Font)
        .Caption = PropBag.ReadProperty("Caption", defCaption)
        .Link = PropBag.ReadProperty("Link", defLink)
        .WaveSound = PropBag.ReadProperty("WaveSound", defWaveSound)
    End With
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With mProps
        PropBag.WriteProperty "Alignment", .Alignment
        PropBag.WriteProperty "BackColor1", .BackColor1
        PropBag.WriteProperty "BackColor2", .BackColor2
        PropBag.WriteProperty "Borderless", .Borderless
        PropBag.WriteProperty "CloseButton", .CloseButton
        PropBag.WriteProperty "ForeColor", .ForeColor
        PropBag.WriteProperty "GradientAngle", .GradientAngle
        PropBag.WriteProperty "Icon", .Icon
        PropBag.WriteProperty "LinkColor", .LinkColor
        PropBag.WriteProperty "Margin", .Margin
        PropBag.WriteProperty "MaskColor", .MaskColor
        PropBag.WriteProperty "Offset", .Offset
        PropBag.WriteProperty "Picture", .Picture
        PropBag.WriteProperty "Stretch", .Stretch
        PropBag.WriteProperty "TextInset", .TextInset
        PropBag.WriteProperty "Transparency", .Transparency
        PropBag.WriteProperty "UseMask", .UseMask
        PropBag.WriteProperty "AlertStyle", .AlertStyle
        PropBag.WriteProperty "AlwaysOnTop", .AlwaysOnTop
        PropBag.WriteProperty "Animation", .Animation
        PropBag.WriteProperty "AutoSize", .AutoSize
        PropBag.WriteProperty "Duration", .Duration
        PropBag.WriteProperty "Persisting", .Persisting
        PropBag.WriteProperty "Speed", .Speed
        PropBag.WriteProperty "Font", .Font
        PropBag.WriteProperty "Caption", .Caption
        PropBag.WriteProperty "Link", .Link
        PropBag.WriteProperty "WaveSound", .WaveSound
    End With
End Sub

Public Property Get Alignment() As AlignmentConstants
Attribute Alignment.VB_Description = "Caption alignment."
Attribute Alignment.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Alignment = mProps.Alignment
End Property

Public Property Let Alignment(Value As AlignmentConstants)
    mProps.Alignment = Value
    PropertyChanged "Alignment"
End Property

Public Property Get BackColor1() As OLE_COLOR
Attribute BackColor1.VB_Description = "The first color of the background gradient."
Attribute BackColor1.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BackColor1 = mProps.BackColor1
End Property

Public Property Let BackColor1(Value As OLE_COLOR)
    mProps.BackColor1 = Value
    PropertyChanged "BackColor1"
End Property

Public Property Get BackColor2() As OLE_COLOR
Attribute BackColor2.VB_Description = "The second color of the background gradient."
Attribute BackColor2.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BackColor2 = mProps.BackColor2
End Property

Public Property Let BackColor2(Value As OLE_COLOR)
    mProps.BackColor2 = Value
    PropertyChanged "BackColor2"
End Property

Public Property Get Borderless() As Boolean
Attribute Borderless.VB_Description = "Whether or not to draw an alert's border."
Attribute Borderless.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Borderless = mProps.Borderless
End Property

Public Property Let Borderless(Value As Boolean)
    mProps.Borderless = Value
    PropertyChanged "Borderless"
End Property

Public Property Get CloseButton() As Boolean
Attribute CloseButton.VB_Description = "Whether or not to add a close button."
Attribute CloseButton.VB_ProcData.VB_Invoke_Property = ";Appearance"
    CloseButton = mProps.CloseButton
End Property

Public Property Let CloseButton(Value As Boolean)
    mProps.CloseButton = Value
    PropertyChanged "CloseButton"
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "The color to be used for the caption displayed on an alert."
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ForeColor = mProps.ForeColor
End Property

Public Property Let ForeColor(Value As OLE_COLOR)
    mProps.ForeColor = Value
    PropertyChanged "ForeColor"
End Property

Public Property Get GradientAngle() As Single
Attribute GradientAngle.VB_Description = "The angle used to paint the gradient background of an alert."
Attribute GradientAngle.VB_ProcData.VB_Invoke_Property = ";Appearance"
    GradientAngle = mProps.GradientAngle
End Property

Public Property Let GradientAngle(Value As Single)
    mProps.GradientAngle = Value
    PropertyChanged "GradientAngle"
End Property

Public Property Get Icon() As StdPicture
Attribute Icon.VB_Description = "The picture to be painted on an alert as its icon."
Attribute Icon.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Set Icon = mProps.Icon
End Property

Public Property Set Icon(Value As StdPicture)
    Set mProps.Icon = Value
    PropertyChanged "Icon"
End Property

Public Property Get LinkColor() As OLE_COLOR
Attribute LinkColor.VB_Description = "The color to be used for the caption of an alert when indicating a link."
Attribute LinkColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    LinkColor = mProps.LinkColor
End Property

Public Property Let LinkColor(Value As OLE_COLOR)
    mProps.LinkColor = Value
    PropertyChanged "LinkColor"
End Property

Public Property Get Margin() As Single
Attribute Margin.VB_Description = "The margin in 'Twips' to keep to the left and to the right of the caption."
Attribute Margin.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Margin = mProps.Margin
End Property

Public Property Let Margin(Value As Single)
    mProps.Margin = Value
    PropertyChanged "Margin"
End Property

Public Property Get MaskColor() As OLE_COLOR
Attribute MaskColor.VB_Description = "The color to use as the mask of an alert."
Attribute MaskColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    MaskColor = mProps.MaskColor
End Property

Public Property Let MaskColor(Value As OLE_COLOR)
    If Value < 0 Then RaiseError INVALID_PROPERTY, "MaskColor": Exit Property
    mProps.MaskColor = Value
    PropertyChanged "MaskColor"
End Property

Public Property Get Offset() As Single
Attribute Offset.VB_Description = "The distance in 'Twips' between the right border of the screen and that of an alert."
Attribute Offset.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Offset = mProps.Offset
End Property

Public Property Let Offset(Value As Single)
    mProps.Offset = Value
    PropertyChanged "Offset"
End Property

Public Property Get Picture() As StdPicture
Attribute Picture.VB_Description = "The picture to be painted on an alert's background."
Attribute Picture.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Set Picture = mProps.Picture
End Property

Public Property Set Picture(Value As StdPicture)
    Set mProps.Picture = Value
    PropertyChanged "Picture"
End Property

Public Property Get Stretch() As Boolean
Attribute Stretch.VB_Description = "Whether or not to stretch the background picture to fit the client area of an alert."
Attribute Stretch.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Stretch = mProps.Stretch
End Property

Public Property Let Stretch(Value As Boolean)
    mProps.Stretch = Value
    PropertyChanged "Stretch"
End Property

Public Property Get TextInset() As Boolean
Attribute TextInset.VB_Description = "Whether or not to add a 'TextInset' effect to the caption of an alert."
Attribute TextInset.VB_ProcData.VB_Invoke_Property = ";Appearance"
    TextInset = mProps.TextInset
End Property

Public Property Let TextInset(Value As Boolean)
    mProps.TextInset = Value
    PropertyChanged "TextInset"
End Property

Public Property Get Transparency() As Byte
Attribute Transparency.VB_Description = "Transparency level of an alert in the range of 0-255."
Attribute Transparency.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Transparency = mProps.Transparency
End Property

Public Property Let Transparency(Value As Byte)
    mProps.Transparency = Value
    PropertyChanged "Transparency"
End Property

Public Property Get UseMask() As Boolean
Attribute UseMask.VB_Description = "Whether or not to activate the mask of an alert."
Attribute UseMask.VB_ProcData.VB_Invoke_Property = ";Appearance"
    UseMask = mProps.UseMask
End Property

Public Property Let UseMask(Value As Boolean)
    mProps.UseMask = Value
    PropertyChanged "UseMask"
End Property

Public Property Get AlertStyle() As taAlertStyles
Attribute AlertStyle.VB_Description = "Sets the style of an alert."
Attribute AlertStyle.VB_ProcData.VB_Invoke_Property = ";Behavior"
    AlertStyle = mProps.AlertStyle
End Property

Public Property Let AlertStyle(Value As taAlertStyles)
    mProps.AlertStyle = Value
    PropertyChanged "AlertStyle"
End Property

Public Property Get AlwaysOnTop() As Boolean
Attribute AlwaysOnTop.VB_Description = "Whether or not to keep an alert on top of all other windows."
Attribute AlwaysOnTop.VB_ProcData.VB_Invoke_Property = ";Behavior"
    AlwaysOnTop = mProps.AlwaysOnTop
End Property

Public Property Let AlwaysOnTop(Value As Boolean)
    mProps.AlwaysOnTop = Value
    PropertyChanged "AlwaysOnTop"
End Property

Public Property Get Animation() As taAnimations
Attribute Animation.VB_Description = "Sets the animation of an alert."
Attribute Animation.VB_ProcData.VB_Invoke_Property = ";Behavior"
    Animation = mProps.Animation
End Property

Public Property Let Animation(Value As taAnimations)
    mProps.Animation = Value
    PropertyChanged "Animation"
End Property

Public Property Get AutoSize() As taAutoSize
Attribute AutoSize.VB_Description = "Sets the auto sizing method of an alert."
Attribute AutoSize.VB_ProcData.VB_Invoke_Property = ";Behavior"
    AutoSize = mProps.AutoSize
End Property

Public Property Let AutoSize(Value As taAutoSize)
    mProps.AutoSize = Value
    PropertyChanged "AutoSize"
End Property

Public Property Get Duration() As Long
Attribute Duration.VB_Description = "The lifetime of an alert in milliseconds."
Attribute Duration.VB_ProcData.VB_Invoke_Property = ";Behavior"
    Duration = mProps.Duration
End Property

Public Property Let Duration(Value As Long)
    mProps.Duration = Value
    PropertyChanged "Duration"
End Property

Public Property Get Persisting() As Boolean
Attribute Persisting.VB_Description = "Currently inactive ..."
Attribute Persisting.VB_ProcData.VB_Invoke_Property = ";Behavior"
    Persisting = mProps.Persisting
End Property

Public Property Let Persisting(Value As Boolean)
    mProps.Persisting = Value
    PropertyChanged "Persisting"
End Property

Public Property Get Speed() As Byte
Attribute Speed.VB_Description = "Speed delay in milliseconds while animating an alert."
Attribute Speed.VB_ProcData.VB_Invoke_Property = ";Behavior"
    Speed = mProps.Speed
End Property

Public Property Let Speed(Value As Byte)
    If Value > 25 Then Value = 25
    mProps.Speed = Value
    PropertyChanged "Speed"
End Property

Public Property Get Font() As StdFont
    Set Font = mProps.Font
End Property

Public Property Set Font(Value As StdFont)
Attribute Font.VB_Description = "The font to be used for the caption displayed on an alert."
Attribute Font.VB_ProcData.VB_Invoke_PropertyPutRef = ";Font"
    Set mProps.Font = Value
    PropertyChanged "Font"
End Property

Public Property Get Caption() As String
Attribute Caption.VB_Description = "The text to be displayed on an alert."
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Text"
    Caption = mProps.Caption
End Property

Public Property Let Caption(Value As String)
    mProps.Caption = Value
    PropertyChanged "Caption"
End Property

Public Property Get Link() As String
Attribute Link.VB_Description = "An internet URL, EMail link, or a file location to be executed for a 'taLink' style alert when the caption is clicked."
Attribute Link.VB_ProcData.VB_Invoke_Property = ";Text"
    Link = mProps.Link
End Property

Public Property Let Link(Value As String)
    mProps.Link = Value
    PropertyChanged "Link"
End Property

Public Property Get WaveSound() As String
Attribute WaveSound.VB_Description = "The .wav file or the ""WAVE"" type resource ID to be played while animating an alert ON."
Attribute WaveSound.VB_ProcData.VB_Invoke_Property = ";Text"
    WaveSound = mProps.WaveSound
End Property

Public Property Let WaveSound(Value As String)
    mProps.WaveSound = Value
    PropertyChanged "WaveSound"
End Property

Public Function ControlProps() As taProperties
Attribute ControlProps.VB_Description = "Returns a 'taProperties' structure containing the control properties."
    ControlProps = mProps
End Function

Public Function DefaultProps() As taProperties
Attribute DefaultProps.VB_Description = "Returns a 'taProperties' structure containing the control's default properties."
    With DefaultProps
        .Alignment = vbCenter
        .BackColor1 = defBackColor1
        .BackColor2 = defBackColor2
        .Borderless = defBorderless
        .CloseButton = defCloseButton
        .ForeColor = defForeColor
        .GradientAngle = defGradientAngle
        Set .Icon = UserControl.Picture
        .LinkColor = defLinkColor
        .Margin = defMargin
        .MaskColor = defMaskColor
        .Offset = defOffset
        Set .Picture = Nothing
        .Stretch = defStretch
        .TextInset = defTextInset
        .Transparency = defTransparency
        .UseMask = defUseMask
        .AlertStyle = taLink
        .AlwaysOnTop = defAlwaysOnTop
        .Animation = taRoll
        .AutoSize = taDefault
        .Duration = defDuration
        .Persisting = defPersisting
        .Speed = defSpeed
        Set .Font = Ambient.Font
        .Font.Name = "Verdana"
        .Caption = defCaption
        .Link = defLink
        .WaveSound = defWaveSound
    End With
End Function

Public Sub Reset()
Attribute Reset.VB_Description = "Closes all currently displayed alerts and resets the control."
    Call ClearAll(True)
    Call UserControl_Initialize
End Sub

Public Sub AboutBox()
Attribute AboutBox.VB_Description = "Displays an About Box for the Control."
Attribute AboutBox.VB_UserMemId = -552
    If frm Is Nothing Then
        Set frm = New frmAbout
        frm.Release = "waTrayAlert v" & App.Major & "." & App.Minor
        frm.CodeURL = "http://planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=47905&lngWId=1"
        frm.Show vbModal, Me
        Set frm = Nothing
    End If
End Sub

Public Function Trigger(InstanceProps As taProperties, Optional hWndControl As Long, Optional sTag As String) As Integer
Attribute Trigger.VB_Description = "Triggers an alert of the specified properties."

    Dim Ret As String
    Dim Key As Integer
    Dim bLayered As Boolean
    Dim Props As taProperties
    Dim ControlLeft As Single, ControlTop As Single
    
    Props = InstanceProps
    Ret = ValidateProps(Props)
    If Ret <> "" Then
        RaiseError INVALID_PROPERTY, Ret
        Exit Function
    End If
    
    If hWndControl <> 0 Then
        If IsInvalidWindow(hWndControl) Then
            RaiseError INVALID_CONTROL
            Exit Function
        End If
        If IsControlInUse(hWndControl) Then
            RaiseError CONTROL_IN_USE
            Exit Function
        End If
    End If
    
    If (Not IsNumeric(Props.WaveSound)) And (Props.WaveSound <> "") Then
        If Dir(Props.WaveSound) = "" Then
            RaiseError WAVE_NOT_FOUND
            Exit Function
        End If
    End If
    
    Key = NewAlert(True)
    If Key = 0 Then
        Reset
        Key = NewAlert(True)
        If Key = 0 Then
            RaiseError UNEXPECTED_ERROR
            Exit Function
        End If
    End If
    
    With Alert(Key)
        .hWndControl = hWndControl
        .Props = Props
        .Tag = sTag
    End With
    
    With Props
    
        Set picAlert(Key).Font = .Font
        picAlert(Key).ForeColor = .ForeColor
        AlertResize Key, True
        
        bLayered = .UseMask Or .Transparency <> 255 Or .Animation <> taRoll
        CreateAlert picAlert(Key).hWnd, hWndControl, bLayered, .AlwaysOnTop
        
        If hWndControl <> 0 Then
            ControlLeft = (picAlert(Key).Width - WindowWidth(hWndControl)) / 2
            ControlTop = Alert(Key).Space.Client.Bottom - WindowHeight(hWndControl)
            WindowMove hWndControl, ControlLeft, ControlTop
        End If
        
        If .BackColor1 = .BackColor2 Then
            picAlert(Key).BackColor = .BackColor1
        Else
            DrawGrad picAlert(Key), .BackColor1, .BackColor2, .GradientAngle
        End If
        
        If Not .Picture Is Nothing Then
            DrawImage picAlert(Key), .Picture, Alert(Key).Space.Client, .Stretch
        End If
        
        If Not .Icon Is Nothing Then
            DrawImage picAlert(Key), .Icon, Alert(Key).Space.Icon, False
        End If
        
        If .CloseButton Then
            DrawImage picAlert(Key), picClose.Picture, Alert(Key).Space.Close, False
        End If
        
        If .Caption <> "" Then
            DrawText picAlert(Key), .Caption, .Alignment, .TextInset, Alert(Key).Space.Caption
        End If
        
        If Not .Borderless Then DrawBorders picAlert(Key)
        
        AlertReposition Key
        RaiseEvent Created(Key, sTag, picAlert(Key).hWnd)
        
        PlayWave .WaveSound
        AnimateON Key
        Timer(Key).Interval = .Duration
    
    End With
    
    Trigger = Key
    Count = Count + 1
    RaiseEvent Loaded(Key, sTag)

End Function

Public Function LoadContainer(InstanceProps As taProperties, hWndControl As Long, Optional sTag As String) As Integer
Attribute LoadContainer.VB_Description = "Loads a control onto an alert container."

    Dim Ret As String
    Dim Key As Integer
    Dim bLayered As Boolean
    Dim Props As taProperties
    
    Props = InstanceProps
    Ret = ValidateProps(Props)
    If Ret <> "" Then
        RaiseError INVALID_PROPERTY, Ret
        Exit Function
    End If
    
    If IsInvalidWindow(hWndControl) Then
        RaiseError INVALID_CONTROL
        Exit Function
    End If
    
    If IsControlInUse(hWndControl) Then
        RaiseError CONTROL_IN_USE
        Exit Function
    End If
    
    If (Not IsNumeric(Props.WaveSound)) And (Props.WaveSound <> "") Then
        If Dir(Props.WaveSound) = "" Then
            RaiseError WAVE_NOT_FOUND
            Exit Function
        End If
    End If
    
    Key = NewAlert(False)
    If Key = 0 Then
        Reset
        Key = NewAlert(False)
        If Key = 0 Then
            RaiseError UNEXPECTED_ERROR
            Exit Function
        End If
    End If
    
    With Alert(Key)
        .hWndControl = hWndControl
        .Props = Props
        .Tag = sTag
    End With
    
    With Props
    
        AlertResize Key, False
        
        bLayered = .UseMask Or .Transparency <> 255 Or .Animation <> taRoll
        CreateAlert picAlert(Key).hWnd, hWndControl, bLayered, .AlwaysOnTop
        
        WindowMove hWndControl, Alert(Key).Space.Client.Left, Alert(Key).Space.Client.Top
        
        If .BackColor1 = .BackColor2 Then
            picAlert(Key).BackColor = .BackColor1
        Else
            DrawGrad picAlert(Key), .BackColor1, .BackColor2, .GradientAngle
        End If
        
        If Not .Picture Is Nothing Then
            DrawImage picAlert(Key), .Picture, Alert(Key).Space.Client, .Stretch
        End If
        
        If Not .Borderless Then DrawBorders picAlert(Key)
        
        AlertReposition Key
        RaiseEvent Created(Key, sTag, picAlert(Key).hWnd)
        
        PlayWave .WaveSound
        AnimateON Key
        Timer(Key).Interval = .Duration
    
    End With
    
    Count = Count + 1
    LoadContainer = Key
    RaiseEvent Loaded(Key, sTag)

End Function

Public Sub CloseAlert(Key As Integer, bAnimate As Boolean, Optional bForcefully As Boolean)
Attribute CloseAlert.VB_Description = "Closes the alert of the specified key."

    Dim bCancel As Boolean
    
    If IsValidKey(Key) Then
        RaiseEvent QueryUnload(Key, Alert(Key).Tag, taCodeInvoked, bCancel)
        If (Not bForcefully) And bCancel Then Exit Sub
        If bAnimate Then AnimateOFF Key
        UnloadAlert Key
    Else
        RaiseError INVALID_KEY
    End If

End Sub

Public Sub ClearAll(Optional bForcefully As Boolean)
Attribute ClearAll.VB_Description = "Closes all currently displayed alerts."

    Dim i As Integer
    Dim bCancel As Boolean
    
    For i = LBound(Alert) To UBound(Alert)
        If IsValidKey(i) Then
            bCancel = False
            RaiseEvent QueryUnload(i, Alert(i).Tag, taClearingAllAlerts, bCancel)
            If bForcefully Or (Not bCancel) Then UnloadAlert i
        End If
    Next

End Sub

Public Function AlertCount() As Integer
Attribute AlertCount.VB_Description = "Returns the number of currently displayed alerts."
    AlertCount = Count
End Function

Public Function AlertHandle(Key As Integer) As Long
Attribute AlertHandle.VB_Description = "Returns the window handle of the alert of the specified key."
    If IsValidKey(Key) Then AlertHandle = picAlert(Key).hWnd Else RaiseError INVALID_KEY
End Function

Public Function WavePath(ID As taWaveList) As String
Attribute WavePath.VB_Description = "Returns a string containing the location of the .wav file identified by a value from a list of common files."

    Dim TempPath As String
    
    Select Case ID
        Case taMS_NewAlert
            TempPath = GetInstallationPath(MessengerService) & "newalert.wav"
        Case taMS_NewEMail
            TempPath = GetInstallationPath(MessengerService) & "newemail.wav"
        Case taMS_Online
            TempPath = GetInstallationPath(MessengerService) & "online.wav"
        Case taMS_Type
            TempPath = GetInstallationPath(MessengerService) & "type.wav"
        Case taMSN_NewAlert
            TempPath = GetInstallationPath(MSNMessenger) & "newalert.wav"
        Case taMSN_NewEMail
            TempPath = GetInstallationPath(MSNMessenger) & "newemail.wav"
        Case taMSN_Online
            TempPath = GetInstallationPath(MSNMessenger) & "online.wav"
        Case taMSN_Ring
            TempPath = GetInstallationPath(MSNMessenger) & "ring.wav"
        Case taMSN_Type
            TempPath = GetInstallationPath(MSNMessenger) & "type.wav"
    End Select
    
    If Dir(TempPath) <> "" Then WavePath = TempPath

End Function

Public Function IsLayerSafe() As Boolean
Attribute IsLayerSafe.VB_Description = "Returns 'True' if the Windows platform of the target computer supports layered windows."
    IsLayerSafe = LayersSupported
End Function

Public Function DrawGradient(picObject As Object, Color1 As Long, Color2 As Long, Angle As Single) As Boolean
Attribute DrawGradient.VB_Description = "Paints a graphical object with a custom gradient."
    DrawGradient = DrawGrad(picObject, Color1, Color2, Angle)
End Function

Public Function Play(WaveSound As String) As Boolean
Attribute Play.VB_Description = "Plays a .wav file or a ""WAVE"" type resource ID."
    Play = PlayWave(WaveSound)
End Function

Private Function NewAlert(bStandard As Boolean) As Integer
On Error GoTo errNewAlert

    Dim NewKey As Integer
    
    NewKey = UBound(Alert) + 1
    ReDim Preserve Alert(LBound(Alert) To NewKey) As AlertData
    
    Load picAlert(NewKey)
    Load Timer(NewKey)
    
    If bStandard Then
        picAlert(NewKey).Tag = StandardAlertID
    Else
        picAlert(NewKey).Tag = ContainerAlertID
    End If
    
    NewAlert = NewKey

errNewAlert:
End Function

Private Sub UnloadAlert(Key As Integer)
On Error GoTo errUnloadAlert

    Count = Count - 1
    
    With Alert(Key)
        If GetParent(.hWndControl) = picAlert(Key).hWnd Then
            SetParent .hWndControl, UserControl.hWnd
        End If
        Unload Timer(Key)
        Unload picAlert(Key)
        Alert(Key).hWndControl = 0
        RaiseEvent Unloaded(Key, .Tag)
    End With
    
    ReduceAlertData
    Exit Sub

errUnloadAlert:
RaiseError UNEXPECTED_ERROR
End Sub

Private Sub AlertResize(Key As Integer, bStandard As Boolean)

    Dim Header As Single
    Dim OffsetX As Single, OffsetY As Single
    Dim PicWidth As Single, PicHeight As Single
    Dim IconWidth As Single, IconHeight As Single
    Dim CloseWidth As Single, CloseHeight As Single
    Dim AlertWidth As Single, AlertHeight As Single
    Dim CaptionWidth As Single, CaptionHeight As Single
    Dim ControlWidth As Single, ControlHeight As Single
    
    With Alert(Key).Props
    
        If Not .Borderless Then
            OffsetX = 4 * Screen.TwipsPerPixelX
            OffsetY = 4 * Screen.TwipsPerPixelY
        End If
        
        If bStandard Then
        
            If Not .Icon Is Nothing Then
                IconWidth = picAlert(Key).ScaleX(.Icon.Width, vbHimetric, vbTwips)
                IconHeight = picAlert(Key).ScaleY(.Icon.Height, vbHimetric, vbTwips)
            End If
            
            If .CloseButton Then
                CloseWidth = picAlert(Key).ScaleX(picClose.Picture.Width, vbHimetric, vbTwips)
                CloseHeight = picAlert(Key).ScaleY(picClose.Picture.Height, vbHimetric, vbTwips)
            End If
            
            If IconHeight <> 0 Or CloseHeight <> 0 Then
                If IconHeight > CloseHeight Then
                    Header = IconHeight + 4 * Screen.TwipsPerPixelY
                Else
                    Header = CloseHeight + 4 * Screen.TwipsPerPixelY
                End If
                CaptionHeight = 4 * Screen.TwipsPerPixelY
            End If
            
            If .Caption <> "" Then
                CaptionWidth = picAlert(Key).TextWidth(.Caption) + 2 * .Margin
                CaptionHeight = picAlert(Key).TextHeight(.Caption) + 20 * Screen.TwipsPerPixelY
            End If
            
            AlertWidth = IconWidth + CloseWidth + 60 * Screen.TwipsPerPixelX
            If AlertWidth < CaptionWidth Then AlertWidth = CaptionWidth
            
            If Alert(Key).hWndControl <> 0 Then
                ControlWidth = WindowWidth(Alert(Key).hWndControl)
                ControlHeight = WindowHeight(Alert(Key).hWndControl)
                If AlertWidth < ControlWidth Then AlertWidth = ControlWidth
            End If
            
            If .Caption <> "" Then
                If Header > ControlHeight Then
                    AlertHeight = 2 * Header + CaptionHeight
                Else
                    AlertHeight = 2 * ControlHeight + CaptionHeight
                End If
            Else
                AlertHeight = Header + CaptionHeight + ControlHeight
            End If
            
            AlertWidth = AlertWidth + 2 * OffsetX
            AlertHeight = AlertHeight + 2 * OffsetY
        Else
            AlertWidth = WindowWidth(Alert(Key).hWndControl) + 2 * OffsetX
            AlertHeight = WindowHeight(Alert(Key).hWndControl) + 2 * OffsetY
        End If
        
        If .AutoSize = taDefault Or .Picture Is Nothing Then GoTo Finalize
        
        PicWidth = picAlert(Key).ScaleX(.Picture.Width, vbHimetric, vbTwips) + 2 * OffsetX
        PicHeight = picAlert(Key).ScaleY(.Picture.Height, vbHimetric, vbTwips) + 2 * OffsetY
        
        If AlertWidth < PicWidth Then AlertWidth = PicWidth
        If AlertHeight < PicHeight Then AlertHeight = PicHeight
    
    End With

Finalize:

    picAlert(Key).Width = AlertWidth
    picAlert(Key).Height = AlertHeight
    
    With Alert(Key).Space
    
        .Client.Left = OffsetX
        .Client.Right = AlertWidth - OffsetX
        .Client.Top = OffsetY
        .Client.Bottom = AlertHeight - OffsetY
        
        If bStandard Then
            If Not Alert(Key).Props.Icon Is Nothing Then
                .Icon.Left = .Client.Left + 4 * Screen.TwipsPerPixelX
                .Icon.Right = .Icon.Left + IconWidth
                .Icon.Top = .Client.Top + 4 * Screen.TwipsPerPixelY
                .Icon.Bottom = .Icon.Top + IconHeight
            End If
            If Alert(Key).Props.CloseButton Then
                .Close.Left = .Client.Right - 4 * Screen.TwipsPerPixelX - CloseWidth
                .Close.Right = .Close.Left + CloseWidth
                .Close.Top = .Client.Top + 4 * Screen.TwipsPerPixelY
                .Close.Bottom = .Close.Top + CloseHeight
            End If
            If Alert(Key).Props.Caption <> "" Then
                CaptionWidth = picAlert(Key).TextWidth(Alert(Key).Props.Caption)
                CaptionHeight = picAlert(Key).TextHeight(Alert(Key).Props.Caption)
                .Caption.Left = (AlertWidth - CaptionWidth) / 2
                .Caption.Right = .Caption.Left + CaptionWidth
                .Caption.Top = (AlertHeight - CaptionHeight) / 2
                .Caption.Bottom = .Caption.Top + CaptionHeight
            End If
        End If
    
    End With

End Sub

Private Sub AlertReposition(Key As Integer)

    Dim i As Integer
    Dim TempTop As Single
    Dim PrimaryTop As Single
    
    TempTop = TaskBarTop
    PrimaryTop = TempTop - picAlert(Key).Height
    
    For i = 0 To UBound(TopLayer)
        If IsValidKey(TopLayer(i)) Then
            If picAlert(TopLayer(i)).Top < TempTop Then TempTop = picAlert(TopLayer(i)).Top
        End If
    Next
    
    TempTop = TempTop - picAlert(Key).Height
    If TempTop < 0 Then TempTop = PrimaryTop
    
    If TempTop = PrimaryTop Then
        ReDim TopLayer(0 To 0)
        TopLayer(0) = Key
    Else
        ReDim Preserve TopLayer(0 To UBound(TopLayer) + 1)
        TopLayer(UBound(TopLayer)) = Key
    End If
    
    picAlert(Key).Move Screen.Width - picAlert(Key).Width - Alert(Key).Props.Offset, TempTop

End Sub

Private Sub AnimateON(Key As Integer)

    With Alert(Key)
        Select Case .Props.Animation
            Case taRoll
                RollUP picAlert(Key), .Props, False
            Case taFade
                FadeIN picAlert(Key), .Props, False
            Case taBoth
                RollUP picAlert(Key), .Props, True
        End Select
    End With

End Sub

Private Sub AnimateOFF(Key As Integer)

    With Alert(Key)
        Select Case .Props.Animation
            Case taRoll
                RollDown picAlert(Key), .Props, False
            Case taFade
                FadeIN picAlert(Key), .Props, True
            Case taBoth
                RollDown picAlert(Key), .Props, True
        End Select
    End With

End Sub

Private Function ValidateProps(Props As taProperties) As String

    With Props
        If Not LayersSupported Then
            .Transparency = 255
            .UseMask = False
            .Animation = taRoll
        End If
        If Not (.Alignment = vbLeftJustify Or .Alignment = vbRightJustify Or .Alignment = vbCenter) Then ValidateProps = "Alignment, "
        If .Margin < 0 Then ValidateProps = ValidateProps & "Margin, "
        If .MaskColor < 0 Then ValidateProps = ValidateProps & "MaskColor, "
        If .Offset < 0 Then ValidateProps = ValidateProps & "Offset, "
        If Not (.AlertStyle = taOriginal Or .AlertStyle = taLink) Then ValidateProps = ValidateProps & "AlertStyle, "
        If Not (.Animation = taRoll Or .Animation = taFade Or .Animation = taBoth) Then ValidateProps = ValidateProps & "Animation, "
        If Not (.AutoSize = taDefault Or .AutoSize = taPictureSize) Then ValidateProps = ValidateProps & "AutoSize, "
        If .Duration < 0 Then ValidateProps = ValidateProps & "Duration, "
        If .Speed > 25 Then .Speed = 25
        If .Font Is Nothing Then ValidateProps = ValidateProps & "Font, "
    End With
    
    If ValidateProps <> "" Then ValidateProps = Left$(ValidateProps, Len(ValidateProps) - 2) & "."

End Function

Private Function IsValidKey(Key As Integer) As Boolean
On Error GoTo errIsValidKey
    IsValidKey = (picAlert(Key).Tag = StandardAlertID) Or (picAlert(Key).Tag = ContainerAlertID)
errIsValidKey:
End Function

Private Function IsControlInUse(hWndControl As Long) As Boolean

    Dim Key As Integer
    
    For Key = LBound(Alert) To UBound(Alert)
        If IsValidKey(Key) Then
            If Alert(Key).hWndControl = hWndControl Then
                IsControlInUse = True
                Exit Function
            End If
        End If
    Next

End Function

Private Sub ReduceAlertData()

    Dim i As Integer
    Dim NewBase As Integer
    Dim Buffer() As AlertData
    
    Buffer = Alert
    For i = LBound(Alert) To UBound(Alert)
        NewBase = i
        If IsValidKey(i) Then Exit For
    Next
    
    ReDim Alert(NewBase To UBound(Alert))
    For i = LBound(Alert) To UBound(Alert)
        Alert(i) = Buffer(i)
    Next
    Erase Buffer

End Sub
