Attribute VB_Name = "modRegistry"
Option Explicit

Public Enum AppList
    MessengerService
    MSNMessenger
End Enum

Private Const sMessengerValueName As String = "InstallationDirectory"
Private Const sMessengerKey As String = "SOFTWARE\Microsoft\MessengerService\"

Private Const sMSNMessengerValueName As String = "InstallationDirectory"
Private Const sMSNMessengerKey As String = "SOFTWARE\Microsoft\MSNMessenger\"

Private Const KEY_QUERY_VALUE = &H1
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long

Public Function GetInstallationPath(Dest As AppList) As String
On Error Resume Next

    Dim hKey As Long
    Dim Buffer As String
    Dim SubKey As String, ValueName As String
    
    Buffer = String(255, Chr(0))
    
    Select Case Dest
        Case MessengerService
            SubKey = sMessengerKey
            ValueName = sMessengerValueName
        Case MSNMessenger
            SubKey = sMSNMessengerKey
            ValueName = sMessengerValueName
    End Select
    
    RegOpenKeyEx HKEY_LOCAL_MACHINE, SubKey, 0, KEY_QUERY_VALUE, hKey
    RegQueryValueEx hKey, ValueName, 0, 0, Buffer, 255
    RegCloseKey hKey
    
    GetInstallationPath = Left(Buffer, InStr(1, Buffer, Chr(0)) - 1)
    If GetInstallationPath <> "" And Right(GetInstallationPath, 1) <> "\" Then
        GetInstallationPath = GetInstallationPath & "\"
    End If
    
    Exit Function

End Function
