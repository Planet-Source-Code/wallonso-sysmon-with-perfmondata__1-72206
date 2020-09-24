VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "System Monitor"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   705
   ClientWidth     =   8475
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   5.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   245
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   565
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame Frame5 
      Caption         =   "CPU"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   5760
      TabIndex        =   9
      Top             =   2040
      Width           =   2655
      Begin VB.Label lblCPUValue 
         Alignment       =   1  'Rechts
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   20
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblCPUDisp 
         Caption         =   "Usage"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblCPUValue 
         Alignment       =   1  'Rechts
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1320
         TabIndex        =   18
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblCPUDisp 
         Caption         =   "Usage"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Zentriert
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   10
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   " RAM "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   5535
      Begin VB.Label lblRamValue 
         Alignment       =   1  'Rechts
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   1560
         TabIndex        =   16
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label lblRamDisp 
         Caption         =   "Used"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   15
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblRamValue 
         Alignment       =   1  'Rechts
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   14
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lblRamDisp 
         Caption         =   "Free"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblRamValue 
         Alignment       =   1  'Rechts
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1560
         TabIndex        =   12
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblRamDisp 
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   6720
      TabIndex        =   2
      Top             =   0
      Width           =   1695
      Begin VB.PictureBox Picture1 
         Appearance      =   0  '2D
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   1695
         Left            =   120
         ScaleHeight     =   1665
         ScaleWidth      =   1425
         TabIndex        =   3
         Top             =   240
         Width           =   1455
         Begin VB.Shape RAM 
            BackColor       =   &H00008000&
            BackStyle       =   1  'Undurchsichtig
            Height          =   60
            Index           =   19
            Left            =   840
            Shape           =   4  'Gerundetes Rechteck
            Top             =   360
            Width           =   495
         End
         Begin VB.Shape RAM 
            BackColor       =   &H00008000&
            BackStyle       =   1  'Undurchsichtig
            Height          =   60
            Index           =   18
            Left            =   840
            Shape           =   4  'Gerundetes Rechteck
            Top             =   400
            Width           =   495
         End
         Begin VB.Shape RAM 
            BackColor       =   &H00008000&
            BackStyle       =   1  'Undurchsichtig
            Height          =   60
            Index           =   17
            Left            =   840
            Shape           =   4  'Gerundetes Rechteck
            Top             =   450
            Width           =   495
         End
         Begin VB.Shape RAM 
            BackColor       =   &H00008000&
            BackStyle       =   1  'Undurchsichtig
            Height          =   60
            Index           =   16
            Left            =   840
            Shape           =   4  'Gerundetes Rechteck
            Top             =   500
            Width           =   495
         End
         Begin VB.Shape RAM 
            BackColor       =   &H00008000&
            BackStyle       =   1  'Undurchsichtig
            Height          =   60
            Index           =   15
            Left            =   840
            Shape           =   4  'Gerundetes Rechteck
            Top             =   540
            Width           =   495
         End
         Begin VB.Shape RAM 
            BackColor       =   &H00008000&
            BackStyle       =   1  'Undurchsichtig
            Height          =   60
            Index           =   14
            Left            =   840
            Shape           =   4  'Gerundetes Rechteck
            Top             =   590
            Width           =   495
         End
         Begin VB.Shape RAM 
            BackColor       =   &H00008000&
            BackStyle       =   1  'Undurchsichtig
            Height          =   60
            Index           =   13
            Left            =   840
            Shape           =   4  'Gerundetes Rechteck
            Top             =   630
            Width           =   495
         End
         Begin VB.Shape RAM 
            BackColor       =   &H00008000&
            BackStyle       =   1  'Undurchsichtig
            Height          =   60
            Index           =   12
            Left            =   840
            Shape           =   4  'Gerundetes Rechteck
            Top             =   680
            Width           =   495
         End
         Begin VB.Shape RAM 
            BackColor       =   &H00008000&
            BackStyle       =   1  'Undurchsichtig
            Height          =   60
            Index           =   11
            Left            =   840
            Shape           =   4  'Gerundetes Rechteck
            Top             =   720
            Width           =   495
         End
         Begin VB.Shape RAM 
            BackColor       =   &H00008000&
            BackStyle       =   1  'Undurchsichtig
            Height          =   60
            Index           =   10
            Left            =   840
            Shape           =   4  'Gerundetes Rechteck
            Top             =   760
            Width           =   495
         End
         Begin VB.Shape RAM 
            BackColor       =   &H00008000&
            BackStyle       =   1  'Undurchsichtig
            Height          =   60
            Index           =   9
            Left            =   840
            Shape           =   4  'Gerundetes Rechteck
            Top             =   810
            Width           =   495
         End
         Begin VB.Shape RAM 
            BackColor       =   &H00008000&
            BackStyle       =   1  'Undurchsichtig
            Height          =   60
            Index           =   8
            Left            =   840
            Shape           =   4  'Gerundetes Rechteck
            Top             =   850
            Width           =   495
         End
         Begin VB.Shape RAM 
            BackColor       =   &H00008000&
            BackStyle       =   1  'Undurchsichtig
            Height          =   60
            Index           =   7
            Left            =   840
            Shape           =   4  'Gerundetes Rechteck
            Top             =   900
            Width           =   495
         End
         Begin VB.Shape RAM 
            BackColor       =   &H00008000&
            BackStyle       =   1  'Undurchsichtig
            Height          =   60
            Index           =   6
            Left            =   840
            Shape           =   4  'Gerundetes Rechteck
            Top             =   940
            Width           =   495
         End
         Begin VB.Shape RAM 
            BackColor       =   &H00008000&
            BackStyle       =   1  'Undurchsichtig
            Height          =   60
            Index           =   5
            Left            =   840
            Shape           =   4  'Gerundetes Rechteck
            Top             =   990
            Width           =   495
         End
         Begin VB.Shape RAM 
            BackColor       =   &H00008000&
            BackStyle       =   1  'Undurchsichtig
            Height          =   60
            Index           =   4
            Left            =   840
            Shape           =   4  'Gerundetes Rechteck
            Top             =   1030
            Width           =   495
         End
         Begin VB.Shape RAM 
            BackColor       =   &H00008000&
            BackStyle       =   1  'Undurchsichtig
            Height          =   60
            Index           =   3
            Left            =   840
            Shape           =   4  'Gerundetes Rechteck
            Top             =   1080
            Width           =   495
         End
         Begin VB.Shape RAM 
            BackColor       =   &H00008000&
            BackStyle       =   1  'Undurchsichtig
            Height          =   60
            Index           =   2
            Left            =   840
            Shape           =   4  'Gerundetes Rechteck
            Top             =   1120
            Width           =   495
         End
         Begin VB.Shape RAM 
            BackColor       =   &H00008000&
            BackStyle       =   1  'Undurchsichtig
            Height          =   60
            Index           =   1
            Left            =   840
            Shape           =   4  'Gerundetes Rechteck
            Top             =   1170
            Width           =   495
         End
         Begin VB.Shape RAM 
            BackColor       =   &H00008000&
            BackStyle       =   1  'Undurchsichtig
            Height          =   60
            Index           =   0
            Left            =   840
            Shape           =   4  'Gerundetes Rechteck
            Top             =   1210
            Width           =   495
         End
         Begin VB.Shape CPU 
            BackColor       =   &H00008000&
            BackStyle       =   1  'Undurchsichtig
            Height          =   60
            Index           =   19
            Left            =   120
            Shape           =   4  'Gerundetes Rechteck
            Top             =   400
            Width           =   495
         End
         Begin VB.Shape CPU 
            BackColor       =   &H00008000&
            BackStyle       =   1  'Undurchsichtig
            Height          =   60
            Index           =   18
            Left            =   120
            Shape           =   4  'Gerundetes Rechteck
            Top             =   360
            Width           =   495
         End
         Begin VB.Shape CPU 
            BackColor       =   &H00008000&
            BackStyle       =   1  'Undurchsichtig
            Height          =   60
            Index           =   17
            Left            =   120
            Shape           =   4  'Gerundetes Rechteck
            Top             =   450
            Width           =   495
         End
         Begin VB.Shape CPU 
            BackColor       =   &H00008000&
            BackStyle       =   1  'Undurchsichtig
            Height          =   60
            Index           =   16
            Left            =   120
            Shape           =   4  'Gerundetes Rechteck
            Top             =   500
            Width           =   495
         End
         Begin VB.Shape CPU 
            BackColor       =   &H00008000&
            BackStyle       =   1  'Undurchsichtig
            Height          =   60
            Index           =   15
            Left            =   120
            Shape           =   4  'Gerundetes Rechteck
            Top             =   540
            Width           =   495
         End
         Begin VB.Shape CPU 
            BackColor       =   &H00008000&
            BackStyle       =   1  'Undurchsichtig
            Height          =   60
            Index           =   14
            Left            =   120
            Shape           =   4  'Gerundetes Rechteck
            Top             =   590
            Width           =   495
         End
         Begin VB.Shape CPU 
            BackColor       =   &H00008000&
            BackStyle       =   1  'Undurchsichtig
            Height          =   60
            Index           =   13
            Left            =   120
            Shape           =   4  'Gerundetes Rechteck
            Top             =   630
            Width           =   495
         End
         Begin VB.Shape CPU 
            BackColor       =   &H00008000&
            BackStyle       =   1  'Undurchsichtig
            Height          =   60
            Index           =   12
            Left            =   120
            Shape           =   4  'Gerundetes Rechteck
            Top             =   680
            Width           =   495
         End
         Begin VB.Shape CPU 
            BackColor       =   &H00008000&
            BackStyle       =   1  'Undurchsichtig
            Height          =   60
            Index           =   11
            Left            =   120
            Shape           =   4  'Gerundetes Rechteck
            Top             =   720
            Width           =   495
         End
         Begin VB.Shape CPU 
            BackColor       =   &H00008000&
            BackStyle       =   1  'Undurchsichtig
            Height          =   60
            Index           =   10
            Left            =   120
            Shape           =   4  'Gerundetes Rechteck
            Top             =   760
            Width           =   495
         End
         Begin VB.Shape CPU 
            BackColor       =   &H00008000&
            BackStyle       =   1  'Undurchsichtig
            Height          =   60
            Index           =   0
            Left            =   120
            Shape           =   4  'Gerundetes Rechteck
            Top             =   1210
            Width           =   495
         End
         Begin VB.Shape CPU 
            BackColor       =   &H00008000&
            BackStyle       =   1  'Undurchsichtig
            Height          =   60
            Index           =   1
            Left            =   120
            Shape           =   4  'Gerundetes Rechteck
            Top             =   1170
            Width           =   495
         End
         Begin VB.Shape CPU 
            BackColor       =   &H00008000&
            BackStyle       =   1  'Undurchsichtig
            Height          =   60
            Index           =   2
            Left            =   120
            Shape           =   4  'Gerundetes Rechteck
            Top             =   1120
            Width           =   495
         End
         Begin VB.Shape CPU 
            BackColor       =   &H00008000&
            BackStyle       =   1  'Undurchsichtig
            Height          =   60
            Index           =   3
            Left            =   120
            Shape           =   4  'Gerundetes Rechteck
            Top             =   1080
            Width           =   495
         End
         Begin VB.Shape CPU 
            BackColor       =   &H00008000&
            BackStyle       =   1  'Undurchsichtig
            Height          =   60
            Index           =   4
            Left            =   120
            Shape           =   4  'Gerundetes Rechteck
            Top             =   1030
            Width           =   495
         End
         Begin VB.Shape CPU 
            BackColor       =   &H00008000&
            BackStyle       =   1  'Undurchsichtig
            Height          =   60
            Index           =   5
            Left            =   120
            Shape           =   4  'Gerundetes Rechteck
            Top             =   990
            Width           =   495
         End
         Begin VB.Shape CPU 
            BackColor       =   &H00008000&
            BackStyle       =   1  'Undurchsichtig
            Height          =   60
            Index           =   6
            Left            =   120
            Shape           =   4  'Gerundetes Rechteck
            Top             =   940
            Width           =   495
         End
         Begin VB.Shape CPU 
            BackColor       =   &H00008000&
            BackStyle       =   1  'Undurchsichtig
            Height          =   60
            Index           =   7
            Left            =   120
            Shape           =   4  'Gerundetes Rechteck
            Top             =   900
            Width           =   495
         End
         Begin VB.Shape CPU 
            BackColor       =   &H00008000&
            BackStyle       =   1  'Undurchsichtig
            Height          =   60
            Index           =   8
            Left            =   120
            Shape           =   4  'Gerundetes Rechteck
            Top             =   850
            Width           =   495
         End
         Begin VB.Shape CPU 
            BackColor       =   &H00008000&
            BackStyle       =   1  'Undurchsichtig
            Height          =   60
            Index           =   9
            Left            =   120
            Shape           =   4  'Gerundetes Rechteck
            Top             =   810
            Width           =   495
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Zentriert
            BackColor       =   &H00000000&
            Caption         =   "CPU"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   1320
            Width           =   495
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Zentriert
            BackColor       =   &H00000000&
            Caption         =   "RAM"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   255
            Left            =   840
            TabIndex        =   6
            Top             =   1320
            Width           =   495
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Zentriert
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   255
            Left            =   0
            TabIndex        =   5
            Top             =   120
            Width           =   735
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Zentriert
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   255
            Left            =   720
            TabIndex        =   4
            Top             =   120
            Width           =   735
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Usage"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6495
      Begin VB.PictureBox Graph 
         Appearance      =   0  '2D
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   1695
         Left            =   120
         ScaleHeight     =   111
         ScaleMode       =   0  'Benutzerdefiniert
         ScaleWidth      =   190
         TabIndex        =   1
         Top             =   240
         Width           =   6255
      End
   End
   Begin VB.Timer tmrUpdate 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8520
      Top             =   240
   End
   Begin VB.Menu MnuFile 
      Caption         =   "File"
      Begin VB.Menu newServer 
         Caption         =   "Enter Server"
      End
      Begin VB.Menu mnuOption 
         Caption         =   "Option"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "Quit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private TabCPU() As Long
Private TabRAM() As Long
Private CptCPU As Long
Public GridScale As Integer
Public LineColorCPU
Public LineColorRAM
Public LineColorGrid


Private m_UsedPhysicalMemory As Currency
Private m_TotalPhysicalMemory As Currency
Private m_AvailablePhysicalMemory As Currency
Private m_NumCpus As Long



Private m_clsPerfRAM As ClassPMonRam
Private m_clsPerfCPU As ClassPMonCPU

Private mVarServer As String



Private Sub UpdateGraphs(CPU, RAMTotal, RAMDispo)
Dim ValueCPU As Long
Dim ValueRAM As Long
Dim cpt As Long
    ReDim Preserve TabCPU(CptCPU)
    ReDim Preserve TabRAM(CptCPU)
    If CptCPU > 0 Then
        For cpt = CptCPU To 1 Step -1
            TabCPU(cpt) = TabCPU(cpt - 1)
            TabRAM(cpt) = TabRAM(cpt - 1)
        Next
    End If
    TabCPU(0) = CPU
    If RAMTotal > 0 Then
        TabRAM(0) = CLng(100 - ((RAMDispo * 100) / RAMTotal))
    End If
    ValueCPU = CLng(TabCPU(0) / 5)
    ValueRAM = CLng(TabRAM(0) / 5)
    For cpt = 0 To 19
        If cpt <= ValueRAM Then
            frmMain.RAM(cpt).BackColor = &HFF00&
        Else
            frmMain.RAM(cpt).BackColor = &H8000&
        End If
        If cpt <= ValueCPU Then
            frmMain.CPU(cpt).BackColor = &HFF00&
        Else
            frmMain.CPU(cpt).BackColor = &H8000&
        End If
    Next
    If CptCPU < 200 Then CptCPU = CptCPU + 1
    frmMain.Graph.Cls
    For cpt = 0 To 100 Step GridScale
        frmMain.Graph.Line (0, (cpt * frmMain.Graph.ScaleHeight) / 100)-(frmMain.Graph.ScaleWidth, (cpt * frmMain.Graph.ScaleHeight) / 100), LineColorGrid
    Next
    For cpt = 1 To CptCPU - 1
        frmMain.Graph.Line (frmMain.Graph.ScaleWidth - (cpt - 1), frmMain.Graph.ScaleHeight - (frmMain.Graph.ScaleHeight * TabCPU(cpt - 1)) / 100)-(frmMain.Graph.ScaleWidth - (cpt), frmMain.Graph.ScaleHeight - (frmMain.Graph.ScaleHeight * TabCPU(cpt)) / 100), LineColorCPU
        frmMain.Graph.Line (frmMain.Graph.ScaleWidth - (cpt - 1), frmMain.Graph.ScaleHeight - (frmMain.Graph.ScaleHeight * TabRAM(cpt - 1)) / 100)-(frmMain.Graph.ScaleWidth - (cpt), frmMain.Graph.ScaleHeight - (frmMain.Graph.ScaleHeight * TabRAM(cpt)) / 100), LineColorRAM
    Next

End Sub


Private Sub Form_Load()
    Dim strSys As String
    
    If App.PrevInstance = True Then
        MsgBox "just running", vbInformation, "CPURAM"
        End
    End If
    
    Set m_clsPerfRAM = New ClassPMonRam
    Set m_clsPerfCPU = New ClassPMonCPU



    
    
    CptCPU = 0
    GridScale = 20
    mVarServer = "" 'Local
    
    m_clsPerfRAM.RemoteServer = mVarServer
    m_clsPerfCPU.RemoteServer = mVarServer
    
    m_NumCpus = m_clsPerfCPU.GetCPUCount
    m_TotalPhysicalMemory = m_clsPerfRAM.GetRamAmount
    
    LineColorCPU = RGB(255, 0, 0)
    LineColorRAM = RGB(0, 255, 0)
    LineColorGrid = RGB(0, 0, 255)
    Graph.BackColor = &H0&
    tmrUpdate_Timer
    tmrUpdate_Timer
    tmrUpdate.Enabled = True
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set m_clsPerfCPU = Nothing
    Set m_clsPerfRAM = Nothing
    End
End Sub




Private Sub mnuOption_Click()
    frmOption.Show
End Sub

Private Sub mnuQuit_Click()
    End
End Sub

Private Sub newServer_Click()
Dim sResult As String
    sResult = InputBox("enter servername")
    If Len(sResult) Then
        If sResult = "." Then
            mVarServer = ""
        Else
            mVarServer = sResult
        End If
        m_clsPerfCPU.RemoteServer = mVarServer
        m_clsPerfRAM.RemoteServer = mVarServer
        m_NumCpus = m_clsPerfCPU.GetCPUCount
        m_TotalPhysicalMemory = m_clsPerfRAM.GetRamAmount
    Else
        'No Change
        
    End If
    

End Sub




Private Sub tmrUpdate_Timer()
tmrUpdate.Enabled = False
Dim CPUUsage As Long
Dim RAMTotal As Currency
Dim RAMAvail As Currency
Dim xRamUse As Currency
Dim svar

    DoEvents
    Dim lCPULoad As Long
    Dim lCPUIndex As Long
    
    m_clsPerfRAM.RemoteServer = mVarServer
    m_clsPerfCPU.RemoteServer = mVarServer
    m_clsPerfCPU.CollectCPUData
    
    
    
    For lCPUIndex = 1 To m_NumCpus
        lCPULoad = lCPULoad + m_clsPerfCPU.GetCPUUsage(lCPUIndex)
        CPUUsage = m_clsPerfCPU.GetCPUUsage(lCPUIndex)
    Next lCPUIndex
    
    GetMemoryInfo
    RAMTotal = m_TotalPhysicalMemory / 1024
    RAMAvail = m_AvailablePhysicalMemory / 1024
    
    lblCPUValue(0) = CPUUsage & " %"
    lblRamValue(0).Caption = RAMTotal & " Kb"
    lblRamValue(1).Caption = RAMAvail & " Kb"
    xRamUse = RAMTotal - RAMAvail
    lblRamValue(2) = xRamUse & " Kb"
    
    lblCPUValue(1).Caption = CPUUsage & " %"
    If RAMTotal Then
    Label4.Caption = Int(((xRamUse * 100) / RAMTotal)) & " %"
    End If
    Call UpdateGraphs(CPUUsage, RAMTotal, RAMAvail)


    tmrUpdate.Enabled = True
End Sub

Public Sub GetMemoryInfo()

  'm_TotalPhysicalMemory = m_clsPerfRAM.GetRamAmount
  If m_TotalPhysicalMemory >= 0 Then
    m_AvailablePhysicalMemory = m_clsPerfRAM.GetPerfMonValue(4, 24)  ' MemStatus.dwAvailPhys
    m_UsedPhysicalMemory = m_TotalPhysicalMemory - m_AvailablePhysicalMemory
  Else
    m_TotalPhysicalMemory = 0
    m_AvailablePhysicalMemory = 0
    m_UsedPhysicalMemory = 0
  End If

End Sub



