VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmSPLASHFORM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SPLASH SCREEN"
   ClientHeight    =   6690
   ClientLeft      =   2145
   ClientTop       =   2445
   ClientWidth     =   11910
   LinkTopic       =   "Form21"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "SPLASH SCREEN.frx":0000
   ScaleHeight     =   6690
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   10080
      Top             =   600
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   6240
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      MousePointer    =   2
      Max             =   105
      Scrolling       =   1
   End
   Begin VB.Image Image8 
      Height          =   2895
      Left            =   -120
      Picture         =   "SPLASH SCREEN.frx":5DBFF
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   600
   End
   Begin VB.Image Image7 
      Height          =   2895
      Left            =   11520
      Picture         =   "SPLASH SCREEN.frx":704EA
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   600
   End
   Begin VB.Image Image6 
      Height          =   1545
      Left            =   5160
      Picture         =   "SPLASH SCREEN.frx":82DD5
      Top             =   4080
      Width           =   1785
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   4680
      TabIndex        =   9
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   5880
      Width           =   2055
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "ON"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      TabIndex        =   6
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "PROJECT"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   5
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "a"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   4
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "THE COLLEGE MANAGEMENT SYSTEM"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   3480
      Width           =   8295
   End
   Begin VB.Image Image5 
      Height          =   1500
      Left            =   1320
      Picture         =   "SPLASH SCREEN.frx":84168
      Stretch         =   -1  'True
      Top             =   480
      Width           =   1995
   End
   Begin VB.Image Image4 
      Height          =   2505
      Left            =   9360
      Picture         =   "SPLASH SCREEN.frx":85D48
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2610
   End
   Begin VB.Image Image2 
      Height          =   2505
      Left            =   0
      Picture         =   "SPLASH SCREEN.frx":97BA8
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   2610
   End
   Begin VB.Image Image3 
      Height          =   2385
      Left            =   9240
      Picture         =   "SPLASH SCREEN.frx":A8BBC
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   2730
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "AS"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   2
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "RAJENDRA  NAGAR,PATNA-800016"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   1
      Top             =   1320
      Width           =   4695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ARCADE BUSINESS COLLEGE"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3240
      TabIndex        =   0
      Top             =   840
      Width           =   7335
   End
   Begin VB.Image Image1 
      Height          =   2505
      Left            =   0
      Picture         =   "SPLASH SCREEN.frx":B89FD
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2610
   End
End
Attribute VB_Name = "frmSPLASHFORM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
Timer1.Enabled = True
End Sub
Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1.Value + 5
Label8.Caption = "LOADING..."
Label9.Caption = ProgressBar1.Value & "%"
If (ProgressBar1.Value = ProgressBar1.Max) Then
Timer1.Enabled = False
Unload Me
frmlogin.Show
End If
End Sub
