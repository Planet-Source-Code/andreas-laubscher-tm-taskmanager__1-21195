Attribute VB_Name = "modTicker"
Option Explicit
'==========================================================================================================='
' API Declarations                                                                                          '
'==========================================================================================================='
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
'==========================================================================================================='
' Type declarations                                                                                         '
'==========================================================================================================='
Private Type MEMORYSTATUS
        dwLength As Long
        dwMemoryLoad As Long
        dwTotalPhys As Long
        dwAvailPhys As Long
        dwTotalPageFile As Long
        dwAvailPageFile As Long
        dwTotalVirtual As Long
        dwAvailVirtual As Long
End Type
Dim tMemStat As MEMORYSTATUS
'==========================================================================================================='
' Variable declarations                                                                                     '
'==========================================================================================================='
Public intStoreX            As Integer
Public intStoreY            As Integer

Public lRet                 As Long
Public NewX                 As Long
Public NewY                 As Long
Public tmpStep              As Integer

Public intProcX             As Integer
Public intProcY             As Integer
Public nProcX               As Integer
Public nProcY               As Integer

Dim iLoadCPU                As Integer
Dim iLoadMemory             As Integer
'-----------------------------------------------------------------------------------------------------------'
' Defines Memory graph styles                                                                               '
'-----------------------------------------------------------------------------------------------------------'
Public strColourCPU         As Long
Public strColourMemory      As Long
Public strNotAlwaysOnTop    As Integer

Public Counter As Integer

Sub RefreshMemory()
'==========================================================================================================='
' Queries the system and returns memory information                                                         '
'==========================================================================================================='
Dim var1
'-----------------------------------------------------------------------------------------------------------'
' Query the system                                                                                          '
'-----------------------------------------------------------------------------------------------------------'
    GlobalMemoryStatus tMemStat
'-----------------------------------------------------------------------------------------------------------'
'   Totals                                                                                                  '
'-----------------------------------------------------------------------------------------------------------'
    frmMain.lblTotalPageFile.Caption = Format(tMemStat.dwTotalPageFile / 1024, "###,##0")
    frmMain.lblTotalVirtual.Caption = Format(tMemStat.dwTotalVirtual / 1024, "###,##0")
    frmMain.lblTotPhys.Caption = Format(tMemStat.dwTotalPhys / 1024, "###,##0")
'-----------------------------------------------------------------------------------------------------------'
'   Available                                                                                               '
'-----------------------------------------------------------------------------------------------------------'
    frmMain.lblAvailPageFile.Caption = Format(tMemStat.dwAvailPageFile / 1024, "###,##0")
    frmMain.lblAvailPhys.Caption = Format(tMemStat.dwAvailPhys / 1024, "###,##0")
    frmMain.lblAvailVirtual.Caption = Format(tMemStat.dwAvailVirtual / 1024, "###,##0")
'-----------------------------------------------------------------------------------------------------------'
'   Percentages                                                                                             '
'-----------------------------------------------------------------------------------------------------------'
    frmMain.lblPercPage.Caption = Format(frmMain.lblAvailPageFile.Caption / frmMain.lblTotalPageFile.Caption * 100, "0.00") & "%"
    frmMain.lblPercVirtual.Caption = Format(frmMain.lblAvailVirtual.Caption / frmMain.lblTotalVirtual.Caption * 100, "0.00") & "%"
    frmMain.lblPercPhysical.Caption = Format(frmMain.lblAvailPhys.Caption / frmMain.lblTotPhys.Caption * 100, "0.00") & "%"
'-----------------------------------------------------------------------------------------------------------'
'   Load indicator                                                                                          '
'   For some reason dwMemoryLoad will return zero on NT machines, therefore we can't just plonk it in.      '
'-----------------------------------------------------------------------------------------------------------'
    If tMemStat.dwMemoryLoad = 0 Then                                                                       '
        '---------------------------------------------------------------------------------------------------'
        ' Now set up the caption                                                                            '
        '---------------------------------------------------------------------------------------------------'
        iLoadMemory = 100 - CInt(frmMain.lblAvailPhys.Caption / frmMain.lblTotPhys.Caption * 100)
    Else                        '
        iLoadMemory = tMemStat.dwMemoryLoad
    End If
    
    iLoadCPU = Asc(Mid(GetKeyValue(HKEY_DYN_DATA, RK_Performance, "Kernel\CPUUsage"), 1, 1))
'-----------------------------------------------------------------------------------------------------------'
' Update the load bars                                                                                      '
'-----------------------------------------------------------------------------------------------------------'
'   CPU
    If frmMain.chkShowCPU.Value Then
        frmMain.lblBarCPU.Height = frmMain.picBackCPU.Height * iLoadCPU / 100
        frmMain.lblBarCPU.Top = frmMain.picBackCPU.Height - frmMain.lblBarCPU.Height
        frmMain.lblLoadCPU.Caption = iLoadCPU
        StepUpProgress intProcX, intProcY, iLoadCPU, strColourCPU, "P"
    End If
    
    If frmMain.chkShowMemory.Value Then
        frmMain.lblBarMemory.Height = frmMain.picBackMemory.Height * iLoadMemory / 100
        frmMain.lblBarMemory.Top = frmMain.picBackMemory.Height - frmMain.lblBarMemory.Height
        frmMain.lblLoadMemory.Caption = iLoadMemory
        StepUpProgress intStoreX, intStoreY, iLoadMemory, strColourMemory, "M"
    End If
    
    DoEvents

End Sub

Sub StepUpProgress(X1 As Integer, Y1 As Integer, Percentage As Integer, Colour As Long, PM As String)
'==========================================================================================================='
' Add a line segment to the Memory load graph                                                               '
'==========================================================================================================='
    tmpStep = frmMain.sldStep.Value + 1
    
    NewX = X1 + tmpStep
    NewY = frmMain.picGraph.ScaleHeight - ((Percentage / 100) * frmMain.picGraph.ScaleHeight)
'-----------------------------------------------------------------------------------------------------------'
' When we've reached the right hand side of the picturebox, we widen it and move it left, so the graph      '
' stays on-screen. This will, rather ironically, use physical memory, as well as some of the page file each '
' step onwards, but it's the only practical solution I could come up with. If you can come up with a better '
' way, feel free to e-mail me.                                                                              '
'-----------------------------------------------------------------------------------------------------------'
    If (NewX) > (frmMain.picGraph.ScaleWidth - 5) Then
        frmMain.picGraph.Width = frmMain.picGraph.Width + tmpStep
        frmMain.picGraph.Left = frmMain.picGraph.Left - tmpStep
    End If
'-----------------------------------------------------------------------------------------------------------'
' Draw the segment                                                                                          '
'-----------------------------------------------------------------------------------------------------------'
    frmMain.picGraph.Line (NewX, NewY)-(X1, Y1), Colour
'-----------------------------------------------------------------------------------------------------------'
' Set up the next cycle's source point                                                                      '
'-----------------------------------------------------------------------------------------------------------'
    If PM = "M" Then
        intStoreX = NewX
        intStoreY = NewY
    Else
        intProcX = NewX
        intProcY = NewY
    End If
    DoEvents
    
End Sub

