VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "(TM) Task Manager v1.1"
   ClientHeight    =   5250
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7050
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   350
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   470
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   5115
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   9022
      _Version        =   393216
      Style           =   1
      TabHeight       =   529
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "My Computer"
      TabPicture(0)   =   "frmMain.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame9"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame8"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Processes"
      TabPicture(1)   =   "frmMain.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Timer2"
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(2)=   "Frame7"
      Tab(1).Control(3)=   "Frame6"
      Tab(1).Control(4)=   "Frame3"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Memory"
      TabPicture(2)   =   "frmMain.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Timer1"
      Tab(2).Control(1)=   "cmmDlg"
      Tab(2).Control(2)=   "Frame4"
      Tab(2).Control(3)=   "Frame5"
      Tab(2).Control(4)=   "Frame1"
      Tab(2).ControlCount=   5
      Begin VB.Timer Timer2 
         Interval        =   10000
         Left            =   -68580
         Top             =   360
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   -68580
         Top             =   360
      End
      Begin MSComDlg.CommonDialog cmmDlg 
         Left            =   -69060
         Top             =   360
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame Frame2 
         Height          =   3855
         Left            =   -74880
         TabIndex        =   26
         Top             =   360
         Width           =   6675
         Begin ComctlLib.ListView lstTasks 
            Height          =   3495
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   6435
            _ExtentX        =   11351
            _ExtentY        =   6165
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   327682
            ForeColor       =   49152
            BackColor       =   0
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   4
            BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Process"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   1
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Priority"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   2
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "ID"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   3
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Threads"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.Frame Frame8 
         Height          =   4635
         Left            =   1620
         TabIndex        =   38
         Top             =   360
         Width           =   5175
         Begin VB.PictureBox Picture9 
            BackColor       =   &H00000000&
            Height          =   795
            Left            =   120
            ScaleHeight     =   735
            ScaleWidth      =   4875
            TabIndex        =   58
            Top             =   240
            Width           =   4935
            Begin VB.Label lblUserName 
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               ForeColor       =   &H0000FF00&
               Height          =   195
               Left            =   1620
               TabIndex        =   62
               Top             =   420
               Width           =   3135
            End
            Begin VB.Label lblComputerName 
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               ForeColor       =   &H0000FF00&
               Height          =   195
               Left            =   1620
               TabIndex        =   61
               Top             =   120
               Width           =   3135
            End
            Begin VB.Label Label18 
               BackColor       =   &H00000000&
               Caption         =   "Computer Name:"
               ForeColor       =   &H0000C000&
               Height          =   195
               Left            =   300
               TabIndex        =   60
               Top             =   120
               Width           =   1215
            End
            Begin VB.Label Label17 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00000000&
               Caption         =   "User Name:"
               ForeColor       =   &H0000C000&
               Height          =   195
               Left            =   360
               TabIndex        =   59
               Top             =   420
               Width           =   1095
            End
         End
         Begin VB.PictureBox Picture8 
            BackColor       =   &H00000000&
            Height          =   1155
            Left            =   120
            ScaleHeight     =   1095
            ScaleWidth      =   4875
            TabIndex        =   53
            Top             =   1080
            Width           =   4935
            Begin VB.Label Label21 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Update:"
               ForeColor       =   &H0000C000&
               Height          =   195
               Left            =   300
               TabIndex        =   69
               Top             =   780
               Width           =   1155
            End
            Begin VB.Label lblOSUpdate 
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               ForeColor       =   &H0000FF00&
               Height          =   195
               Left            =   1620
               TabIndex        =   63
               Top             =   780
               Width           =   3135
            End
            Begin VB.Label lblOSPlatform 
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               ForeColor       =   &H0000FF00&
               Height          =   195
               Left            =   1620
               TabIndex        =   57
               Top             =   180
               Width           =   3135
            End
            Begin VB.Label lblOSVersion 
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               ForeColor       =   &H0000FF00&
               Height          =   195
               Left            =   1620
               TabIndex        =   56
               Top             =   480
               Width           =   3135
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00000000&
               Caption         =   "OS Platform:"
               ForeColor       =   &H0000C000&
               Height          =   195
               Left            =   540
               TabIndex        =   55
               Top             =   180
               Width           =   915
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00000000&
               Caption         =   "OS Version:"
               ForeColor       =   &H0000C000&
               Height          =   195
               Left            =   360
               TabIndex        =   54
               Top             =   480
               Width           =   1095
            End
         End
         Begin VB.PictureBox Picture7 
            BackColor       =   &H00000000&
            Height          =   1095
            Left            =   120
            ScaleHeight     =   1035
            ScaleWidth      =   4875
            TabIndex        =   46
            Top             =   2280
            Width           =   4935
            Begin VB.Label lblProcessorMake 
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               ForeColor       =   &H0000FF00&
               Height          =   195
               Left            =   1620
               TabIndex        =   52
               Top             =   120
               Width           =   3135
            End
            Begin VB.Label lblProcessorModel 
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               ForeColor       =   &H0000FF00&
               Height          =   195
               Left            =   1620
               TabIndex        =   51
               Top             =   420
               Width           =   3135
            End
            Begin VB.Label lblProcessorSpeed 
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               ForeColor       =   &H0000FF00&
               Height          =   195
               Left            =   1620
               TabIndex        =   50
               Top             =   720
               Width           =   3135
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00000000&
               Caption         =   "CPU Make:"
               ForeColor       =   &H0000C000&
               Height          =   195
               Left            =   60
               TabIndex        =   49
               Top             =   120
               Width           =   1395
            End
            Begin VB.Label Label13 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00000000&
               Caption         =   "Model:"
               ForeColor       =   &H0000C000&
               Height          =   195
               Left            =   420
               TabIndex        =   48
               Top             =   420
               Width           =   1035
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00000000&
               Caption         =   "Speed:"
               ForeColor       =   &H0000C000&
               Height          =   195
               Left            =   600
               TabIndex        =   47
               Top             =   720
               Width           =   855
            End
         End
         Begin VB.PictureBox Picture6 
            BackColor       =   &H00000000&
            Height          =   1095
            Left            =   120
            ScaleHeight     =   1035
            ScaleWidth      =   4875
            TabIndex        =   39
            Top             =   3420
            Width           =   4935
            Begin VB.Label lblRegisteredUser 
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               ForeColor       =   &H0000FF00&
               Height          =   195
               Left            =   1620
               TabIndex        =   45
               Top             =   120
               Width           =   3135
            End
            Begin VB.Label lblRegisteredOrganization 
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               ForeColor       =   &H0000FF00&
               Height          =   195
               Left            =   1620
               TabIndex        =   44
               Top             =   420
               Width           =   3135
            End
            Begin VB.Label lblProductID 
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               ForeColor       =   &H0000FF00&
               Height          =   195
               Left            =   1620
               TabIndex        =   43
               Top             =   720
               Width           =   3135
            End
            Begin VB.Label Label11 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00000000&
               Caption         =   "Registered User:"
               ForeColor       =   &H0000C000&
               Height          =   195
               Left            =   180
               TabIndex        =   42
               Top             =   120
               Width           =   1275
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00000000&
               Caption         =   "Organization:"
               ForeColor       =   &H0000C000&
               Height          =   195
               Left            =   180
               TabIndex        =   41
               Top             =   420
               Width           =   1275
            End
            Begin VB.Label Label9 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00000000&
               Caption         =   "Product ID:"
               ForeColor       =   &H0000C000&
               Height          =   195
               Left            =   360
               TabIndex        =   40
               Top             =   720
               Width           =   1095
            End
         End
      End
      Begin VB.Frame Frame7 
         Height          =   915
         Left            =   -70200
         TabIndex        =   36
         Top             =   4080
         Width           =   1995
         Begin VB.CommandButton cmdRefresh 
            Caption         =   "Refresh"
            Height          =   375
            Left            =   420
            TabIndex        =   37
            Top             =   360
            Width           =   1155
         End
      End
      Begin VB.Frame Frame6 
         Height          =   915
         Left            =   -72540
         TabIndex        =   32
         Top             =   4080
         Width           =   2415
         Begin VB.PictureBox Picture5 
            BackColor       =   &H00000000&
            Height          =   555
            Left            =   180
            ScaleHeight     =   495
            ScaleWidth      =   1995
            TabIndex        =   33
            Top             =   240
            Width           =   2055
            Begin VB.Label Label8 
               BackColor       =   &H00000000&
               Caption         =   "Threads:"
               ForeColor       =   &H0000C000&
               Height          =   195
               Left            =   180
               TabIndex        =   35
               Top             =   180
               Width           =   795
            End
            Begin VB.Label lblThreads 
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Caption         =   "0"
               ForeColor       =   &H0000FF00&
               Height          =   195
               Left            =   1200
               TabIndex        =   34
               Top             =   180
               Width           =   555
            End
         End
      End
      Begin VB.Frame Frame3 
         Height          =   915
         Left            =   -74880
         TabIndex        =   28
         Top             =   4080
         Width           =   2415
         Begin VB.PictureBox Picture4 
            BackColor       =   &H00000000&
            Height          =   555
            Left            =   120
            ScaleHeight     =   495
            ScaleWidth      =   1995
            TabIndex        =   29
            Top             =   240
            Width           =   2055
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Processes:"
               ForeColor       =   &H0000C000&
               Height          =   195
               Left            =   180
               TabIndex        =   31
               Top             =   180
               Width           =   915
            End
            Begin VB.Label lblProcesses 
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Caption         =   "0"
               ForeColor       =   &H0000FF00&
               Height          =   195
               Left            =   1200
               TabIndex        =   30
               Top             =   180
               Width           =   555
            End
         End
      End
      Begin VB.Frame Frame4 
         Height          =   3015
         Left            =   -74880
         TabIndex        =   23
         Top             =   360
         Width           =   855
         Begin VB.CheckBox chkShowMemory 
            Caption         =   "Check2"
            Height          =   195
            Left            =   180
            TabIndex        =   75
            Top             =   2640
            Width           =   195
         End
         Begin VB.CheckBox chkShowCPU 
            Caption         =   "Check1"
            Height          =   195
            Left            =   480
            TabIndex        =   74
            Top             =   2640
            Width           =   195
         End
         Begin VB.PictureBox picBackCPU 
            BackColor       =   &H00000000&
            Height          =   2115
            Left            =   420
            ScaleHeight     =   2055
            ScaleWidth      =   255
            TabIndex        =   70
            ToolTipText     =   "Processor Load"
            Top             =   480
            Width           =   315
            Begin VB.Label lblBarCPU 
               BackColor       =   &H0000C0C0&
               Height          =   195
               Left            =   60
               TabIndex        =   72
               Top             =   1740
               Width           =   135
            End
         End
         Begin VB.PictureBox picBackMemory 
            BackColor       =   &H00000000&
            Height          =   2115
            Left            =   120
            ScaleHeight     =   2055
            ScaleWidth      =   255
            TabIndex        =   24
            ToolTipText     =   "Memory Load"
            Top             =   480
            Width           =   315
            Begin VB.Label lblBarMemory 
               BackColor       =   &H0000C000&
               Height          =   195
               Left            =   60
               TabIndex        =   71
               Top             =   1740
               Width           =   135
            End
         End
         Begin VB.Label lblLoadCPU 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "100"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C0C0&
            Height          =   255
            Left            =   420
            TabIndex        =   73
            Top             =   240
            Width           =   315
         End
         Begin VB.Label lblLoadMemory 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "100"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   315
         End
      End
      Begin VB.Frame Frame5 
         Height          =   3015
         Left            =   -74100
         TabIndex        =   20
         Top             =   360
         Width           =   5895
         Begin VB.PictureBox picBackRight 
            BackColor       =   &H80000012&
            Height          =   2655
            Left            =   180
            ScaleHeight     =   173
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   369
            TabIndex        =   21
            Top             =   240
            Width           =   5595
            Begin VB.PictureBox picGraph 
               AutoRedraw      =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   2595
               Left            =   0
               ScaleHeight     =   173
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   369
               TabIndex        =   22
               Top             =   0
               Width           =   5535
            End
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1755
         Left            =   -74880
         TabIndex        =   1
         Top             =   3240
         Width           =   6675
         Begin VB.PictureBox Picture2 
            DrawStyle       =   1  'Dash
            DrawWidth       =   32
            FillColor       =   &H8000000F&
            ForeColor       =   &H8000000F&
            Height          =   1395
            Left            =   5160
            ScaleHeight     =   1335
            ScaleWidth      =   1335
            TabIndex        =   18
            Top             =   240
            Width           =   1395
            Begin ComctlLib.Slider sldStep 
               Height          =   135
               Left            =   60
               TabIndex        =   67
               Top             =   1020
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   238
               _Version        =   327682
            End
            Begin ComctlLib.Slider sldUpdate 
               Height          =   195
               Left            =   60
               TabIndex        =   65
               Top             =   420
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   344
               _Version        =   327682
               SelStart        =   5
               Value           =   5
            End
            Begin VB.Label Label20 
               Alignment       =   2  'Center
               Caption         =   "Step Size"
               Height          =   195
               Left            =   0
               TabIndex        =   66
               Top             =   780
               Width           =   1335
            End
            Begin VB.Label Label6 
               Alignment       =   2  'Center
               Caption         =   "Update Speed"
               Height          =   195
               Left            =   0
               TabIndex        =   19
               Top             =   180
               Width           =   1335
            End
         End
         Begin VB.PictureBox Picture3 
            BackColor       =   &H00000000&
            DrawStyle       =   1  'Dash
            DrawWidth       =   32
            Height          =   1395
            Left            =   4080
            ScaleHeight     =   1335
            ScaleWidth      =   975
            TabIndex        =   14
            Top             =   240
            Width           =   1035
            Begin VB.Label lblPercVirtual 
               Alignment       =   2  'Center
               BackColor       =   &H8000000E&
               BackStyle       =   0  'Transparent
               ForeColor       =   &H0000FF00&
               Height          =   195
               Left            =   120
               TabIndex        =   17
               Top             =   1020
               Width           =   675
            End
            Begin VB.Label lblPercPage 
               Alignment       =   2  'Center
               BackColor       =   &H8000000E&
               BackStyle       =   0  'Transparent
               ForeColor       =   &H0000FF00&
               Height          =   195
               Left            =   120
               TabIndex        =   16
               Top             =   720
               Width           =   675
            End
            Begin VB.Label lblPercPhysical 
               Alignment       =   2  'Center
               BackColor       =   &H8000000E&
               BackStyle       =   0  'Transparent
               ForeColor       =   &H0000FF00&
               Height          =   195
               Left            =   120
               TabIndex        =   15
               Top             =   420
               Width           =   675
            End
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00000000&
            DrawStyle       =   1  'Dash
            DrawWidth       =   32
            Height          =   1395
            Left            =   120
            ScaleHeight     =   1335
            ScaleWidth      =   3975
            TabIndex        =   2
            Top             =   240
            Width           =   4035
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Virtual Memory"
               ForeColor       =   &H0000C000&
               Height          =   195
               Left            =   240
               TabIndex        =   13
               Top             =   1020
               Width           =   1155
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Page File"
               ForeColor       =   &H0000C000&
               Height          =   195
               Left            =   600
               TabIndex        =   12
               Top             =   720
               Width           =   795
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Physical Memory"
               ForeColor       =   &H0000C000&
               Height          =   195
               Left            =   120
               TabIndex        =   11
               Top             =   420
               Width           =   1275
            End
            Begin VB.Label lblTotPhys 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               ForeColor       =   &H0000C000&
               Height          =   195
               Left            =   1560
               TabIndex        =   10
               Top             =   420
               Width           =   1035
            End
            Begin VB.Label lblTotalPageFile 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               ForeColor       =   &H0000C000&
               Height          =   195
               Left            =   1560
               TabIndex        =   9
               Top             =   720
               Width           =   1035
            End
            Begin VB.Label lblTotalVirtual 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               ForeColor       =   &H0000C000&
               Height          =   195
               Left            =   1560
               TabIndex        =   8
               Top             =   1020
               Width           =   1035
            End
            Begin VB.Label lblAvailVirtual 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               ForeColor       =   &H0000C000&
               Height          =   195
               Left            =   2820
               TabIndex        =   7
               Top             =   1020
               Width           =   1035
            End
            Begin VB.Label lblAvailPageFile 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               ForeColor       =   &H0000C000&
               Height          =   195
               Left            =   2820
               TabIndex        =   6
               Top             =   720
               Width           =   1035
            End
            Begin VB.Label lblAvailPhys 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               ForeColor       =   &H0000C000&
               Height          =   195
               Left            =   2820
               TabIndex        =   5
               Top             =   420
               Width           =   1035
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Total"
               ForeColor       =   &H0000C000&
               Height          =   195
               Left            =   1560
               TabIndex        =   4
               Top             =   120
               Width           =   1035
            End
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Available"
               ForeColor       =   &H0000C000&
               Height          =   195
               Left            =   2820
               TabIndex        =   3
               Top             =   120
               Width           =   1035
            End
         End
      End
      Begin VB.Frame Frame9 
         Height          =   4635
         Left            =   120
         TabIndex        =   64
         Top             =   360
         Width           =   1575
         Begin VB.PictureBox Picture10 
            BackColor       =   &H00000000&
            Height          =   4275
            Left            =   120
            Picture         =   "frmMain.frx":0496
            ScaleHeight     =   281
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   81
            TabIndex        =   68
            Top             =   240
            Width           =   1275
         End
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuStayOnTop 
         Caption         =   "On Top"
      End
      Begin VB.Menu sepf1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuPopupTasks 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuEndProcess 
         Caption         =   "End Process"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSetPriority 
         Caption         =   "Set Priority"
         Begin VB.Menu mnuPriority 
            Caption         =   "Real Time"
            Index           =   1
         End
         Begin VB.Menu mnuPriority 
            Caption         =   "High"
            Index           =   2
         End
         Begin VB.Menu mnuPriority 
            Caption         =   "Normal"
            Index           =   3
         End
         Begin VB.Menu mnuPriority 
            Caption         =   "Idle"
            Index           =   4
         End
      End
   End
   Begin VB.Menu mnuPopupGraph 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuClear 
         Caption         =   "Clear..."
      End
      Begin VB.Menu sepclear 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPause 
         Caption         =   "Pause"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==========================================================================================================='
' Author                    : Andreas Laubscher                                                             '
' Contact                   : andreaslaubscher@hotmail.com                                                  '
' Date                      : 17 February 2001                                                              '
' Description               : An NT style Task Manager for Windows 95/98                                    '
'==========================================================================================================='
Option Explicit
'==========================================================================================================='
' API Declarations                                                                                          '
'==========================================================================================================='
' Sets window to top.                                                                                       '
'-----------------------------------------------------------------------------------------------------------'
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'==========================================================================================================='
' Variable Declarations                                                                                     '
'==========================================================================================================='
Private fWidth As Integer
Private fHeight As Integer

Private Sub cmdRefresh_Click()

    Call RefreshTasks
    
End Sub

Private Sub Form_Load()

'-----------------------------------------------------------------------------------------------------------'
' Save the form's current size, this is to disable the resize option, while showing the minimize button     '
'-----------------------------------------------------------------------------------------------------------'
    fWidth = Width
    fHeight = Height
'-----------------------------------------------------------------------------------------------------------'
'   Reset our Load graph control variables                                                                  '
'-----------------------------------------------------------------------------------------------------------'
    intStoreY = picBackRight.Height
    intProcY = picBackRight.Height
    
    Timer1.Interval = 500
    strColourMemory = &HC000&
    strColourCPU = &HC0C0&
    
    lblBarCPU.Top = 0
    lblBarCPU.Left = 0
    lblBarCPU.Height = picBackMemory.Height
    lblBarCPU.Width = picBackMemory.Width
    
    lblBarMemory.Top = 0
    lblBarMemory.Left = 0
    lblBarMemory.Height = picBackCPU.Height
    lblBarMemory.Width = picBackCPU.Width
'-----------------------------------------------------------------------------------------------------------'
' Edit the ListView's column widths                                                                         '
'-----------------------------------------------------------------------------------------------------------'
    frmMain.lstTasks.ColumnHeaders(1).Width = frmMain.lstTasks.Width / 4
    frmMain.lstTasks.ColumnHeaders(2).Width = frmMain.lstTasks.Width / 6
    frmMain.lstTasks.ColumnHeaders(3).Width = frmMain.lstTasks.Width / 6
    frmMain.lstTasks.ColumnHeaders(4).Width = frmMain.lstTasks.Width / 6
'-----------------------------------------------------------------------------------------------------------'
' Set the window to permanently display on top (Z-order wise), this is necessary to prevent redraw problems '
' with the load graph. It also makes it easier to monitor the performance of an application                 '
'-----------------------------------------------------------------------------------------------------------'
    Call mnuStayOnTop_Click
'-----------------------------------------------------------------------------------------------------------'
' Trigger a screen update                                                                                   '
'-----------------------------------------------------------------------------------------------------------'
    Call GetSysInfo
    Call RefreshMemory
    Call RefreshTasks
    
End Sub

Private Sub Form_Resize()

    If WindowState <> vbMinimized Then
        Height = fHeight
        Width = fWidth
    End If
    
End Sub

Private Sub lblBarMemory_Click()
    
    Call picBackMemory_Click
    
End Sub

Private Sub lstTasks_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo exitsub
'==========================================================================================================='
' A right-click triggers the popupmenu                                                                      '
'==========================================================================================================='
    If Button = 2 Then
        lstTasks.HitTest(X, Y).Selected = True
        '---------------------------------------------------------------------------------------------------'
        ' Check the appropriate priority                                                                    '
        '---------------------------------------------------------------------------------------------------'
        frmMain.mnuPriority(1).Checked = False
        frmMain.mnuPriority(2).Checked = False
        frmMain.mnuPriority(3).Checked = False
        frmMain.mnuPriority(4).Checked = False
        
        Select Case lstTasks.SelectedItem.SubItems(1)
        Case "Realtime"
            frmMain.mnuPriority(1).Checked = True
        Case "High"
            frmMain.mnuPriority(2).Checked = True
        Case "Normal"
            frmMain.mnuPriority(3).Checked = True
        Case "Idle"
            frmMain.mnuPriority(4).Checked = True
        End Select
        
        PopupMenu mnuPopupTasks
        
    End If

exitsub:
End Sub

Private Sub mnuClear_Click()
'==========================================================================================================='
' Re-initialize tbe graphing variables                                                                      '
'==========================================================================================================='

    intStoreX = 0
    intProcX = 0
    
    picGraph.Cls
    
    picGraph.Width = picBackRight.Width
    picGraph.Left = 0
    
End Sub

Private Sub mnuEndProcess_Click()
'==========================================================================================================='
' Kill the process                                                                                          '
'==========================================================================================================='
    EndProcess lstTasks.SelectedItem.SubItems(2)
    DoEvents: DoEvents
    Call RefreshTasks
    
End Sub

Private Sub mnuExit_Click()

    Unload Me
    End
    
End Sub

Private Sub mnuPause_Click()
'==========================================================================================================='
' Very simply, we only need to disable the timer that handles the refreshing of the memory data.            '
'==========================================================================================================='
    If Timer1.Enabled Then
        mnuPause.Checked = True
        Timer1.Enabled = False
    Else
        mnuPause.Checked = False
        Timer1.Enabled = True
    End If
    
End Sub

Private Sub mnuPriority_Click(Index As Integer)
'==========================================================================================================='
' Change the priority of the chosen process                                                                 '
'==========================================================================================================='
    SetProcessPriority lstTasks.SelectedItem.SubItems(2), mnuPriority(Index).Caption
    DoEvents: DoEvents
    Call RefreshTasks
    
    For Counter = 1 To 4
        If mnuPriority(Counter).Checked Then
            mnuPriority(Counter).Checked = Not mnuPriority(Counter).Checked
        End If
    Next
    
    mnuPriority(Index).Checked = Not mnuPriority(Index).Checked
    
End Sub

Private Sub mnuStayOnTop_Click()
'==========================================================================================================='
' Sets the Z-Order... um... order of the form.                                                              '
'==========================================================================================================='
    mnuStayOnTop.Checked = Not mnuStayOnTop.Checked
    SetWindowPos Me.hwnd, strNotAlwaysOnTop - 1, 0, 0, 0, 0, 2 Or 1
    strNotAlwaysOnTop = mnuStayOnTop.Checked
    
End Sub

Private Sub picBackCPU_Click()
    On Error GoTo CPUColorExit
    
Dim tmpS As OLE_COLOR

    With cmmDlg
        '---------------------------------------------------------------------------------------------------'
        ' Flags: Sets the initial color value, and sets full colour display on initially                    '
        '---------------------------------------------------------------------------------------------------'
        .DialogTitle = "CPU Colour"
        .Flags = cdlCCRGBInit + cdlCCFullOpen
        .Color = strColourCPU
        .CancelError = True
        .ShowColor
    End With
    
    tmpS = cmmDlg.Color
    strColourCPU = CLng(tmpS)
    
    lblLoadCPU.ForeColor = strColourCPU
    lblBarCPU.BackColor = strColourCPU
    
CPUColorExit:
End Sub

Private Sub picBackMemory_Click()
'==========================================================================================================='
' Trigger a colour change                                                                                   '
'==========================================================================================================='
    On Error GoTo MemoryColorExit
    
Dim tmpS As OLE_COLOR

    With cmmDlg
        '---------------------------------------------------------------------------------------------------'
        ' Flags: Sets the initial color value, and sets full colour display on initially                    '
        '---------------------------------------------------------------------------------------------------'
        .DialogTitle = "Memory Colour"
        .Flags = cdlCCRGBInit + cdlCCFullOpen
        .Color = strColourMemory
        .CancelError = True
        .ShowColor
    End With
    
    tmpS = cmmDlg.Color
    strColourMemory = CLng(tmpS)
    
    lblLoadMemory.ForeColor = strColourMemory
    lblBarMemory.BackColor = strColourMemory
    
MemoryColorExit:
End Sub

Private Sub picGraph_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'==========================================================================================================='
' Triggers the graph's popup menu                                                                           '
'==========================================================================================================='
    If Button = 2 Then
        PopupMenu mnuPopupGraph
    End If
    
End Sub

Private Sub sldUpdate_Change()
'==========================================================================================================='
' Changes the update speed of the ticker                                                                    '
'==========================================================================================================='
    Timer1.Interval = (sldUpdate.Value + 1) * 100
    
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
'==========================================================================================================='
' Refresh the tasks ticker                                                                                  '
'==========================================================================================================='
    If SSTab1.Tab = 1 Then
        Call RefreshTasks
    End If
    
End Sub

Private Sub Timer1_Timer()
'==========================================================================================================='
' Queries the system for ticker information, and edits the results back to the form for display             '
'==========================================================================================================='
    RefreshMemory
    
End Sub

Private Sub Timer2_Timer()
'==========================================================================================================='
' Queries the system for process information, and edits the result back to the form for display             '
'==========================================================================================================='
' We only want to do this if the popumenu is not visible, otherwise me might refresh at the wrong moment    '
'-----------------------------------------------------------------------------------------------------------'
    If Not mnuPopupTasks.Visible Then
        RefreshTasks
    End If
    
End Sub
