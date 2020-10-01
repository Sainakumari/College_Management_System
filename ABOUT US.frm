VERSION 5.00
Begin VB.Form frmABOUT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ABOUT"
   ClientHeight    =   9195
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15150
   LinkTopic       =   "Form24"
   MaxButton       =   0   'False
   Picture         =   "ABOUT US.frx":0000
   ScaleHeight     =   9195
   ScaleWidth      =   15150
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H008080FF&
      Height          =   8745
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   13815
      Begin VB.Frame Frame2 
         BackColor       =   &H00FF8080&
         Height          =   7515
         Left            =   90
         TabIndex        =   2
         Top             =   270
         Width           =   13335
         Begin VB.Frame Frame4 
            BackColor       =   &H00004080&
            Height          =   2055
            Left            =   0
            TabIndex        =   13
            Top             =   0
            Width           =   13575
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "THE COLLEGE MANAGEMENT SYSTEM"
               BeginProperty Font 
                  Name            =   "Algerian"
                  Size            =   32.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Left            =   1320
               TabIndex        =   15
               Top             =   720
               Width           =   12015
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "THE WAY OF SUCCESS"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808080&
               Height          =   375
               Left            =   10200
               TabIndex        =   14
               Top             =   1440
               Width           =   3375
            End
            Begin VB.Image Image1 
               Height          =   1545
               Left            =   0
               Picture         =   "ABOUT US.frx":1E9608
               Stretch         =   -1  'True
               Top             =   240
               Width           =   1425
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00004080&
            Height          =   4815
            Left            =   840
            TabIndex        =   3
            Top             =   2160
            Width           =   12015
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "THE COLLEGE MANAGEMENT SYSTEM  is a system"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   18
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000B&
               Height          =   495
               Left            =   1200
               TabIndex        =   12
               Top             =   240
               Width           =   7935
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "used to keep record of college ,project will be done"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   18
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000B&
               Height          =   495
               Left            =   1200
               TabIndex        =   11
               Top             =   720
               Width           =   7815
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   " using VB 6.0 as front end ,and MS SQL SERVER  as"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   18
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000B&
               Height          =   495
               Left            =   1200
               TabIndex        =   10
               Top             =   1200
               Width           =   7935
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   " back end .It can used to college management. This"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   18
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000B&
               Height          =   495
               Left            =   1200
               TabIndex        =   9
               Top             =   1680
               Width           =   7815
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   " project is mainly useful for college.This system will"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   18
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000B&
               Height          =   495
               Left            =   1200
               TabIndex        =   8
               Top             =   2160
               Width           =   9135
            End
            Begin VB.Label Label8 
               BackStyle       =   0  'Transparent
               Caption         =   " help to manage all the activities in the colleges"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   18
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000B&
               Height          =   495
               Left            =   1200
               TabIndex        =   7
               Top             =   2640
               Width           =   7815
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   "using computers.Currently all the works are done"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   18
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000B&
               Height          =   495
               Left            =   1200
               TabIndex        =   6
               Top             =   3120
               Width           =   9135
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "inside a college campus can be managed easily and effectively."
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   18
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000B&
               Height          =   495
               Left            =   1200
               TabIndex        =   5
               Top             =   4080
               Width           =   9855
            End
            Begin VB.Label Label11 
               BackStyle       =   0  'Transparent
               Caption         =   " manaually,by computerizing all the activities"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   18
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000B&
               Height          =   495
               Left            =   1200
               TabIndex        =   4
               Top             =   3600
               Width           =   9135
            End
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "CLOSE"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5790
         Picture         =   "ABOUT US.frx":1E998F
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   7950
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmABOUT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Unload frmABOUT
End Sub
