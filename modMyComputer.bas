Attribute VB_Name = "modMyComputer"
Option Explicit
'==========================================================================================================='
' Local Constant declarations                                                                               '
'==========================================================================================================='
' Version information constants                                                                             '
'-----------------------------------------------------------------------------------------------------------'
Private Const VER_PLATFORM_WIN32_NT = 2
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32s = 0
'-----------------------------------------------------------------------------------------------------------'
' Constants used to query the registry                                                                      '
'-----------------------------------------------------------------------------------------------------------'
' Registry Key open mode
Const KEY_QUERY_VALUE = &H1
' The Registry section we'll be visiting
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_DYN_DATA = &H80000006
' Root to the processor information
Const RK_Processor = "HARDWARE\DESCRIPTION\System\CentralProcessor\0"
' Root to performance statistics
Public Const RK_Performance = "PerfStats\StatData"
' Root to OS information on Win machines
Const RK_WIN32_OS = "SOFTWARE\Microsoft\Windows\CurrentVersion"
' Root to OS information on NT machines
Const RK_WIN32_OS_NT = "SOFTWARE\Microsoft\Windows NT\CurrentVersion"
'==========================================================================================================='
' API Declarations                                                                                          '
'==========================================================================================================='
' System Information API declarations                                                                       '
'-----------------------------------------------------------------------------------------------------------'
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'-----------------------------------------------------------------------------------------------------------'
' Registry queries API declarations                                                                         '
'-----------------------------------------------------------------------------------------------------------'
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
'==========================================================================================================='
' Type Declarations                                                                                         '
'==========================================================================================================='
Private Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128
End Type
Private tmpVersionInfo As OSVERSIONINFO
'==========================================================================================================='
' Local variable declarations                                                                               '
'==========================================================================================================='
Dim tmpRegKey As String
Dim tmpBuffer As String * 255

Sub GetSysInfo()
'==========================================================================================================='
'==========================================================================================================='
    GetComputerName tmpBuffer, 255
    frmMain.lblComputerName.Caption = Trim$(tmpBuffer)
'-----------------------------------------------------------------------------------------------------------'
    GetUserName tmpBuffer, 255
    frmMain.lblUserName.Caption = tmpBuffer
'-----------------------------------------------------------------------------------------------------------'
    tmpVersionInfo.dwOSVersionInfoSize = 148
    GetVersionEx tmpVersionInfo
'-----------------------------------------------------------------------------------------------------------'
    If tmpVersionInfo.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then
        If tmpVersionInfo.dwMinorVersion = 0 Then
            frmMain.lblOSPlatform.Caption = "Microsoft Windows '95"
        Else
            frmMain.lblOSPlatform.Caption = "Microsoft Windows '98"
        End If
    ElseIf tmpVersionInfo.dwPlatformId = VER_PLATFORM_WIN32_NT Then
        If tmpVersionInfo.dwMajorVersion = 4 Then
            frmMain.lblOSPlatform.Caption = "Microsoft Windows NT"
        Else
            frmMain.lblOSPlatform.Caption = "Microsoft Windows 2000"
        End If
    End If
'-----------------------------------------------------------------------------------------------------------'
    frmMain.lblOSVersion.Caption = tmpVersionInfo.dwMajorVersion & "." & _
        Format(tmpVersionInfo.dwMinorVersion, "00") & "." & _
        tmpVersionInfo.dwBuildNumber
    frmMain.lblOSUpdate.Caption = Left(tmpVersionInfo.szCSDVersion, InStr(1, tmpVersionInfo.szCSDVersion, Chr(0)))
'-----------------------------------------------------------------------------------------------------------'
' Retrieve registration information, this is platform specific                                              '
'-----------------------------------------------------------------------------------------------------------'
    If tmpVersionInfo.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then
        tmpRegKey = RK_WIN32_OS
    ElseIf tmpVersionInfo.dwPlatformId = VER_PLATFORM_WIN32_NT Then
        tmpRegKey = RK_WIN32_OS_NT
    End If
    frmMain.lblRegisteredOrganization.Caption = GetKeyValue(HKEY_LOCAL_MACHINE, tmpRegKey, "RegisteredOrganization")
    frmMain.lblRegisteredUser.Caption = GetKeyValue(HKEY_LOCAL_MACHINE, tmpRegKey, "RegisteredOwner")
    frmMain.lblProductID.Caption = GetKeyValue(HKEY_LOCAL_MACHINE, tmpRegKey, "ProductID")
'-----------------------------------------------------------------------------------------------------------'
' Retrieve CPU information from the registry                                                                '
'-----------------------------------------------------------------------------------------------------------'
    frmMain.lblProcessorMake.Caption = GetKeyValue(HKEY_LOCAL_MACHINE, RK_Processor, "VendorIdentifier")
    frmMain.lblProcessorModel.Caption = GetKeyValue(HKEY_LOCAL_MACHINE, RK_Processor, "Identifier")
    tmpBuffer = GetKeyValue(HKEY_LOCAL_MACHINE, RK_Processor, "~MHZ")
    If Len(Trim(tmpBuffer)) > 0 Then
        frmMain.lblProcessorSpeed.Caption = Trim(tmpBuffer) & " MHz"
    End If
    
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String) As String
'==========================================================================================================='
' Returns a specified key value from the registry                                                           '
'==========================================================================================================='
Dim lKey As Long
Dim tmpVal As String
Dim tmpKeySize As Long
Dim tmpKeyType As Long
Dim Counter As Integer
'-----------------------------------------------------------------------------------------------------------'
' Set up needed variables                                                                                   '
'-----------------------------------------------------------------------------------------------------------'
    tmpVal = String(1024, 0)
    tmpKeySize = 1024
'-----------------------------------------------------------------------------------------------------------'
' Open the registry key. Any value other than zero means something went wrong                               '
'-----------------------------------------------------------------------------------------------------------'
    If RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_QUERY_VALUE, lKey) <> 0 Then
        GetKeyValue = ""
        RegCloseKey lKey
        Exit Function
    End If
'-----------------------------------------------------------------------------------------------------------'
' Retrieve the registry value, any value other than zero means something went wrong                         '
'-----------------------------------------------------------------------------------------------------------'
    If RegQueryValueEx(lKey, SubKeyRef, 0, tmpKeyType, tmpVal, tmpKeySize) Then
        GetKeyValue = ""
        RegCloseKey lKey
        Exit Function
    End If
'-----------------------------------------------------------------------------------------------------------'
' Extract the useful string from the garble                                                                 '
'-----------------------------------------------------------------------------------------------------------'
    If (Asc(Mid(tmpVal, tmpKeySize, 1)) = 0) Then
        tmpVal = Left(tmpVal, tmpKeySize - 1)
    Else
        tmpVal = Left(tmpVal, tmpKeySize)
    End If
'-----------------------------------------------------------------------------------------------------------'
' If the returned value is a dword we need to format the value to something meaningful                      '
'-----------------------------------------------------------------------------------------------------------'
    If tmpKeyType = 4 Then
        For Counter = Len(tmpVal) To 1 Step -1
            GetKeyValue = GetKeyValue + Hex(Asc(Mid(tmpVal, Counter, 1)))
        Next
        GetKeyValue = Format("&h" + GetKeyValue)
    Else
        GetKeyValue = tmpVal
    End If
'-----------------------------------------------------------------------------------------------------------'
' Clean up                                                                                                  '
'-----------------------------------------------------------------------------------------------------------'
    RegCloseKey lKey
    
End Function
