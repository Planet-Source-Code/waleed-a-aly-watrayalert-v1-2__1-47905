Attribute VB_Name = "modError"
Option Explicit

Public Sub RaiseError(ID As taErrors, Optional sExtraInfo As String)

    Dim errDesc As String, errSource As String
    
    Select Case ID
        Case UNEXPECTED_ERROR
            errSource = "TrayAlert"
            errDesc = "Unexpected error was encountered."
            
        Case INVALID_PROPERTY
            errSource = "TrayAlert.waTrayAlert"
            errDesc = "Invalid property value: " & sExtraInfo
            
        Case INVALID_KEY
            errSource = "TrayAlert.waTrayAlert"
            errDesc = "The value used to reference an alert is not a valid key."
            
        Case INVALID_CONTROL
            errSource = "TrayAlert.waTrayAlert"
            errDesc = "The window handle of the control to be loaded onto an alert is not a valid window handle."
            
        Case CONTROL_IN_USE
            errSource = "TrayAlert.waTrayAlert"
            errDesc = "A control to be loaded onto an alert is already in use by another alert."
            
        Case WAVE_NOT_FOUND
            errSource = "TrayAlert.modSoundAPI"
            errDesc = "Wave file cannot be found."
    End Select
    
    Err.Raise ID, errSource, errDesc

End Sub
