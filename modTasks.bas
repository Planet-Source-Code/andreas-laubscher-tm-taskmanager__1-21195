Attribute VB_Name = "modTasks"
Option Explicit
'==========================================================================================================='
' Constant Declarations                                                                                     '
'==========================================================================================================='
Global Const PROCESS_PRIORITY_IDLE = 4
Global Const PROCESS_PRIORITY_NORMAL = 8
Global Const PROCESS_PRIORITY_HIGH = 13
Global Const PROCESS_PRIORITY_REALTIME = 24
' Priority type, when setting with SetPriorityClass
Private Const HIGH_PRIORITY_CLASS = &H80                    ' Hogs CPU over idle and normal classes
Private Const IDLE_PRIORITY_CLASS = &H40                    ' Only runs when the CPU is idle
Private Const NORMAL_PRIORITY_CLASS = &H20                  ' Duh!
Private Const REALTIME_PRIORITY_CLASS = &H100               ' Highest priority. Even pre-empts operating system
                                                            ' processes, so use with discretion
' Access description when opening a handle to a process.
' These codes aren't in the API viewer, had to get them at Microsoft's site.
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const SYNCHRONIZE = &H100000
Public Const PROCESS_TERMINATE = &H1&                       ' Used to kill a process
Public Const PROCESS_CREATE_THREAD = &H2&
Public Const PROCESS_VM_OPERATION = &H8&
Public Const PROCESS_VM_READ = &H10&
Public Const PROCESS_VM_WRITE = &H206
Public Const PROCESS_DUP_HANDLE = &H40&
Public Const PROCESS_CREATE_PROCESS = &H80&
Public Const PROCESS_SET_QUOTA = &H100&
Public Const PROCESS_SET_INFORMATION = &H200&               ' Used to set information on a process (like priority)
Public Const PROCESS_QUERY_INFORMATION = &H400&
Public Const PROCESS_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF)
'==========================================================================================================='
' API Declarations                                                                                          '
'==========================================================================================================='
' Used to return process information                                                                        '
'-----------------------------------------------------------------------------------------------------------'
Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal dwIdProc As Long) As Long
Declare Function Process32First Lib "kernel32" (ByVal hndl As Long, ByRef pstru As ProcessEntry) As Boolean
Declare Function Process32Next Lib "kernel32" (ByVal hndl As Long, ByRef pstru As ProcessEntry) As Boolean
'-----------------------------------------------------------------------------------------------------------'
' Used to change values of a specific process                                                               '
'-----------------------------------------------------------------------------------------------------------'
Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function GetLastError Lib "kernel32" () As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hnd As Long) As Boolean
'==========================================================================================================='
' Type Declarations                                                                                         '
'==========================================================================================================='
Type ProcessEntry
    dwSize As Long
    peUsage As Long
    peProcessID As Long
    peDefaultHeapID As Long
    peModuleID As Long
    peThreads As Long
    peParentProcessID As Long
    pePriority As Long
    dwFlags As Long
    szExeFile As String * 260
End Type
'==========================================================================================================='
' Local variable declarations                                                                               '
'==========================================================================================================='
Dim hnd                             As Long         ' Handle to a process
Dim lRet                            As Long         ' Return value for API calls
Dim lExitCode                       As Long         ' Exit code
Dim lPriority                       As Long         ' Priority

Sub RefreshTasks()
'==========================================================================================================='
' Queries the system and returns process information                                                        '
'==========================================================================================================='
Dim iIdx            As Integer
Dim bRet            As Boolean
Dim lSnapShot       As Long
Dim tmpPE           As ProcessEntry

Dim intProcesses    As Integer
Dim intThreads      As Integer

Dim tmpProcName     As String
Dim tmpPriority     As String

'-----------------------------------------------------------------------------------------------------------'
' Reset display                                                                                             '
'-----------------------------------------------------------------------------------------------------------'
    frmMain.lstTasks.ListItems.Clear
'-----------------------------------------------------------------------------------------------------------'
' Query system                                                                                              '
'-----------------------------------------------------------------------------------------------------------'
    lSnapShot = CreateToolhelp32Snapshot(&H2, 0)
    tmpPE.dwSize = Len(tmpPE)
    bRet = Process32First(lSnapShot, tmpPE)
'-----------------------------------------------------------------------------------------------------------'
' Return information for all processes                                                                      '
'-----------------------------------------------------------------------------------------------------------'
    Do Until bRet = False
        '---------------------------------------------------------------------------------------------------'
        ' Edit the process name to something useful                                                         '
        '---------------------------------------------------------------------------------------------------'
        tmpProcName = LCase(Mid(tmpPE.szExeFile, _
                            InStrRev(tmpPE.szExeFile, "\", Len(tmpPE.szExeFile)) + 1, _
                            Len(tmpPE.szExeFile) - InStrRev(tmpPE.szExeFile, "\", 1)))
        tmpProcName = Left(tmpProcName, InStr(1, tmpProcName, Chr(0)) - 1)
        '---------------------------------------------------------------------------------------------------'
        ' Set up the priority                                                                               '
        '---------------------------------------------------------------------------------------------------'
        Select Case tmpPE.pePriority
        Case PROCESS_PRIORITY_IDLE
            tmpPriority = "Idle"
        Case PROCESS_PRIORITY_NORMAL
            tmpPriority = "Normal"
        Case PROCESS_PRIORITY_REALTIME
            tmpPriority = "Realtime"
        Case PROCESS_PRIORITY_HIGH
            tmpPriority = "High"
        End Select
        '---------------------------------------------------------------------------------------------------'
        ' Add the item to our list                                                                          '
        '---------------------------------------------------------------------------------------------------'
        With frmMain.lstTasks.ListItems.Add(, , tmpProcName)
            .SubItems(1) = tmpPriority
            .SubItems(2) = tmpPE.peProcessID
            .SubItems(3) = tmpPE.peThreads
        End With
        '---------------------------------------------------------------------------------------------------'
        ' Tally the totals                                                                                  '
        '---------------------------------------------------------------------------------------------------'
        intProcesses = intProcesses + 1
        intThreads = intThreads + tmpPE.peThreads

        bRet = Process32Next(lSnapShot, tmpPE)

    Loop
'-----------------------------------------------------------------------------------------------------------'
' Clean up                                                                                                  '
'-----------------------------------------------------------------------------------------------------------'
    bRet = CloseHandle(lSnapShot)
'-----------------------------------------------------------------------------------------------------------'
' Add a blank item to the end of our list, to make it look better                                           '
'-----------------------------------------------------------------------------------------------------------'
    frmMain.lstTasks.ListItems.Add , , ""
'-----------------------------------------------------------------------------------------------------------'
' Set up the tallies' display                                                                               '
'-----------------------------------------------------------------------------------------------------------'
    frmMain.lblProcesses.Caption = intProcesses
    frmMain.lblThreads.Caption = intThreads

End Sub

Sub EndProcess(strProcess As Long)
'==========================================================================================================='
' End the process of the selected Process                                                                   '
'==========================================================================================================='

'-----------------------------------------------------------------------------------------------------------'
' First we need to create a handle to the desired process                                                   '
'-----------------------------------------------------------------------------------------------------------'
    hnd = OpenProcess(PROCESS_TERMINATE, 0, strProcess)
'-----------------------------------------------------------------------------------------------------------'
' Get the process' exit code                                                                                '
'-----------------------------------------------------------------------------------------------------------'
    lRet = GetExitCodeProcess(hnd, lExitCode)
'-----------------------------------------------------------------------------------------------------------'
' Terminate the process! This might lead to screwy results, so be warned                                    '
'-----------------------------------------------------------------------------------------------------------'
    lRet = TerminateProcess(hnd, lExitCode)
'-----------------------------------------------------------------------------------------------------------'
' Close the handle                                                                                          '
'-----------------------------------------------------------------------------------------------------------'
    lRet = CloseHandle(hnd)
    
End Sub

Sub SetProcessPriority(strProcess As Long, strPriority As String)
'==========================================================================================================='
' Set the priority of the currently selected Process                                                        '
'==========================================================================================================='

'-----------------------------------------------------------------------------------------------------------'
' First we need to open a handle to the desired process                                                     '
'-----------------------------------------------------------------------------------------------------------'
    hnd = OpenProcess(PROCESS_SET_INFORMATION, 0, strProcess)
'-----------------------------------------------------------------------------------------------------------'
' Set the priority to the one requested                                                                     '
'-----------------------------------------------------------------------------------------------------------'
    Select Case strPriority
    Case "Realtime"
        lPriority = REALTIME_PRIORITY_CLASS
    Case "High"
        lPriority = HIGH_PRIORITY_CLASS
    Case "Normal"
        lPriority = NORMAL_PRIORITY_CLASS
    Case "Idle"
        lPriority = IDLE_PRIORITY_CLASS
    End Select
    
    lRet = SetPriorityClass(hnd, lPriority)
    
'-----------------------------------------------------------------------------------------------------------'
' Close the handle to the process                                                                           '
'-----------------------------------------------------------------------------------------------------------'
    lRet = CloseHandle(hnd)
    
End Sub
