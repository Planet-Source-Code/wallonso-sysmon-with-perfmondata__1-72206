VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmOption 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Système Moniteur - Option"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   5160
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton Command6 
      Caption         =   "Can"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   13
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   12
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   " Paramètres "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      Begin VB.CommandButton Command4 
         Caption         =   "Change"
         Height          =   375
         Left            =   3720
         TabIndex        =   11
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Change"
         Height          =   375
         Left            =   3720
         TabIndex        =   10
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Change"
         Height          =   375
         Left            =   3720
         TabIndex        =   9
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Zentriert
         Appearance      =   0  '2D
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   2
         Text            =   "1"
         Top             =   2160
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Change"
         Height          =   375
         Left            =   3720
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Color of CPU line"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label Label2 
         Caption         =   "Color of RAM-Line"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   3015
      End
      Begin VB.Label Label3 
         Caption         =   "Color of Font"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1800
         Width           =   3015
      End
      Begin VB.Label Label4 
         Caption         =   "Color of Bars"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   2895
      End
      Begin VB.Label Label5 
         Caption         =   "grid space"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   2280
         Width           =   2535
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Undurchsichtig
         Height          =   375
         Left            =   3120
         Shape           =   1  'Quadrat
         Top             =   240
         Width           =   615
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  'Undurchsichtig
         Height          =   375
         Left            =   3120
         Shape           =   1  'Quadrat
         Top             =   720
         Width           =   615
      End
      Begin VB.Shape Shape3 
         BackStyle       =   1  'Undurchsichtig
         Height          =   375
         Left            =   3120
         Shape           =   1  'Quadrat
         Top             =   1200
         Width           =   615
      End
      Begin VB.Shape Shape4 
         BackStyle       =   1  'Undurchsichtig
         Height          =   375
         Left            =   3120
         Shape           =   1  'Quadrat
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3720
         TabIndex        =   3
         Top             =   2280
         Width           =   255
      End
   End
   Begin MSComDlg.CommonDialog CmD 
      Left            =   3480
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
CmD.ShowColor
Shape1.BackColor = CmD.Color
End Sub

Private Sub Command2_Click()
CmD.ShowColor
Shape2.BackColor = CmD.Color
End Sub

Private Sub Command3_Click()
CmD.ShowColor
Shape3.BackColor = CmD.Color
End Sub

Private Sub Command4_Click()
CmD.ShowColor
Shape4.BackColor = CmD.Color
End Sub

Private Sub Command5_Click()
LineColorCPU = Shape1.BackColor
LineColorRAM = Shape2.BackColor
LineColorGrid = Shape3.BackColor
frmMain.Graph.BackColor = Shape4.BackColor
GridScale = Text1.Text
Unload frmOption
frmMain.Show
End Sub

Private Sub Command6_Click()
Unload frmOption
End Sub

Private Sub Form_Load()
Shape1.BackColor = LineColorCPU
Shape2.BackColor = LineColorRAM
Shape3.BackColor = LineColorGrid
Shape4.BackColor = frmMain.Graph.BackColor
Text1.Text = GridScale
End Sub
