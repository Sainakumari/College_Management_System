VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmFACULTYINFORMATION 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FACULTY "
   ClientHeight    =   10935
   ClientLeft      =   375
   ClientTop       =   210
   ClientWidth     =   20325
   BeginProperty Font 
      Name            =   "Cambria"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   Picture         =   "FACULITY INFORMATION DETAILS.frx":0000
   ScaleHeight     =   10935
   ScaleWidth      =   20325
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "s"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10095
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   19935
      Begin VB.CommandButton Command12 
         Caption         =   "REFRESH"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   18240
         Picture         =   "FACULITY INFORMATION DETAILS.frx":1E9608
         Style           =   1  'Graphical
         TabIndex        =   138
         Top             =   6096
         Width           =   1095
      End
      Begin VB.CommandButton Command5 
         Caption         =   "PRINT"
         Height          =   975
         Left            =   18240
         Picture         =   "FACULITY INFORMATION DETAILS.frx":231E92
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   7440
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "UPDATE"
         Height          =   975
         Left            =   18240
         Picture         =   "FACULITY INFORMATION DETAILS.frx":234A4E
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   4752
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "DELETE"
         Height          =   975
         Left            =   18240
         Picture         =   "FACULITY INFORMATION DETAILS.frx":23743E
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   3408
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "SAVE"
         Height          =   975
         Left            =   18240
         Picture         =   "FACULITY INFORMATION DETAILS.frx":249473
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "VIEW"
         Height          =   975
         Left            =   18240
         Picture         =   "FACULITY INFORMATION DETAILS.frx":278078
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FF8080&
         Caption         =   "  PERSONAL PROFILE"
         Height          =   9975
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   17655
         Begin VB.TextBox Text38 
            BackColor       =   &H80000016&
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   13440
            TabIndex        =   175
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H80000016&
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   5760
            TabIndex        =   163
            Top             =   360
            Width           =   1695
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H80000016&
            Caption         =   "VISITING"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2640
            TabIndex        =   133
            Top             =   360
            Width           =   1215
         End
         Begin VB.Frame Frame7 
            BackColor       =   &H00FF8080&
            Caption         =   "EXPERIENCE"
            Height          =   2415
            Left            =   120
            TabIndex        =   118
            Top             =   7440
            Width           =   17415
            Begin VB.VScrollBar VScroll10 
               Height          =   1335
               Left            =   14760
               TabIndex        =   142
               Top             =   960
               Width           =   135
            End
            Begin VB.VScrollBar VScroll8 
               Height          =   1335
               Left            =   10920
               TabIndex        =   141
               Top             =   960
               Width           =   135
            End
            Begin VB.VScrollBar VScroll7 
               Height          =   1335
               Left            =   7200
               TabIndex        =   140
               Top             =   960
               Width           =   135
            End
            Begin VB.VScrollBar VScroll5 
               Height          =   1335
               Left            =   3480
               TabIndex        =   139
               Top             =   960
               Width           =   135
            End
            Begin VB.ListBox LIST8 
               BackColor       =   &H00FF8080&
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1320
               ItemData        =   "FACULITY INFORMATION DETAILS.frx":2C033A
               Left            =   11400
               List            =   "FACULITY INFORMATION DETAILS.frx":2C033C
               TabIndex        =   132
               Top             =   960
               Width           =   3495
            End
            Begin VB.ListBox List7 
               BackColor       =   &H00FF8080&
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1320
               ItemData        =   "FACULITY INFORMATION DETAILS.frx":2C033E
               Left            =   7560
               List            =   "FACULITY INFORMATION DETAILS.frx":2C0340
               TabIndex        =   131
               Top             =   960
               Width           =   3495
            End
            Begin VB.ListBox List6 
               BackColor       =   &H00FF8080&
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1320
               ItemData        =   "FACULITY INFORMATION DETAILS.frx":2C0342
               Left            =   3840
               List            =   "FACULITY INFORMATION DETAILS.frx":2C0344
               TabIndex        =   130
               Top             =   960
               Width           =   3495
            End
            Begin VB.ListBox List5 
               BackColor       =   &H00FF8080&
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1320
               ItemData        =   "FACULITY INFORMATION DETAILS.frx":2C0346
               Left            =   120
               List            =   "FACULITY INFORMATION DETAILS.frx":2C0348
               TabIndex        =   129
               Top             =   960
               Width           =   3495
            End
            Begin VB.TextBox Text37 
               BackColor       =   &H80000016&
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   11400
               TabIndex        =   128
               Top             =   600
               Width           =   3495
            End
            Begin VB.TextBox Text36 
               BackColor       =   &H80000016&
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   7560
               TabIndex        =   127
               Top             =   600
               Width           =   3495
            End
            Begin VB.TextBox Text35 
               BackColor       =   &H80000016&
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   3840
               TabIndex        =   126
               Top             =   600
               Width           =   3495
            End
            Begin VB.TextBox Text34 
               BackColor       =   &H80000016&
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   120
               TabIndex        =   125
               Top             =   600
               Width           =   3495
            End
            Begin VB.CommandButton Command11 
               Caption         =   "REMOVE"
               Height          =   975
               Left            =   15600
               Picture         =   "FACULITY INFORMATION DETAILS.frx":2C034A
               Style           =   1  'Graphical
               TabIndex        =   124
               Top             =   1320
               Width           =   1095
            End
            Begin VB.CommandButton Command10 
               Caption         =   "ADD"
               Height          =   975
               Left            =   15600
               Picture         =   "FACULITY INFORMATION DETAILS.frx":308BD4
               Style           =   1  'Graphical
               TabIndex        =   123
               Top             =   240
               Width           =   1095
            End
            Begin VB.Line Line10 
               BorderColor     =   &H8000000B&
               X1              =   15000
               X2              =   15000
               Y1              =   120
               Y2              =   2640
            End
            Begin VB.Line Line9 
               BorderColor     =   &H8000000B&
               X1              =   11160
               X2              =   11160
               Y1              =   120
               Y2              =   2640
            End
            Begin VB.Line Line8 
               BorderColor     =   &H8000000B&
               X1              =   7440
               X2              =   7440
               Y1              =   120
               Y2              =   2640
            End
            Begin VB.Line Line7 
               BorderColor     =   &H8000000B&
               X1              =   3720
               X2              =   3720
               Y1              =   120
               Y2              =   2640
            End
            Begin VB.Line Line6 
               BorderColor     =   &H8000000B&
               X1              =   0
               X2              =   15000
               Y1              =   480
               Y2              =   480
            End
            Begin VB.Label Label65 
               BackStyle       =   0  'Transparent
               Caption         =   "DATE(FROM-TO)"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   11880
               TabIndex        =   122
               Top             =   240
               Width           =   1935
            End
            Begin VB.Label Label64 
               BackStyle       =   0  'Transparent
               Caption         =   "YEAR"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   8520
               TabIndex        =   121
               Top             =   240
               Width           =   1935
            End
            Begin VB.Label Label63 
               BackStyle       =   0  'Transparent
               Caption         =   "DESIGNATION/PROFILE"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   4320
               TabIndex        =   120
               Top             =   240
               Width           =   1935
            End
            Begin VB.Label Label62 
               BackStyle       =   0  'Transparent
               Caption         =   "ORGANISATION"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   720
               TabIndex        =   119
               Top             =   240
               Width           =   1935
            End
         End
         Begin VB.TextBox Text33 
            BackColor       =   &H80000016&
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   13440
            TabIndex        =   117
            Top             =   1719
            Width           =   1695
         End
         Begin VB.CommandButton Command7 
            Caption         =   "REMOVE"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   15720
            Picture         =   "FACULITY INFORMATION DETAILS.frx":340C0C
            Style           =   1  'Graphical
            TabIndex        =   106
            Top             =   3480
            Width           =   1095
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00FF8080&
            Caption         =   "FACULTY'S PHOTO"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1935
            Left            =   15240
            TabIndex        =   105
            Top             =   360
            Width           =   2295
            Begin VB.Image Image1 
               BorderStyle     =   1  'Fixed Single
               Height          =   1575
               Left            =   120
               Stretch         =   -1  'True
               Top             =   240
               Width           =   2055
            End
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   285
            Left            =   1920
            TabIndex        =   104
            Top             =   1229
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   503
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd/MMM/yyyy"
            Format          =   89522179
            CurrentDate     =   43091
         End
         Begin MSComDlg.CommonDialog cd2 
            Left            =   17040
            Top             =   3840
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            Filter          =   "*.JPG"
         End
         Begin VB.CommandButton Command6 
            Caption         =   "BROWSE.."
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   15720
            Picture         =   "FACULITY INFORMATION DETAILS.frx":389496
            Style           =   1  'Graphical
            TabIndex        =   82
            Top             =   2400
            Width           =   1095
         End
         Begin VB.TextBox Text28 
            BackColor       =   &H80000016&
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   13440
            TabIndex        =   81
            Top             =   813
            Width           =   1695
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00FF8080&
            Caption         =   "QUALIFICATION"
            Height          =   2655
            Left            =   120
            TabIndex        =   74
            Top             =   4680
            Width           =   17415
            Begin VB.VScrollBar VScroll4 
               Height          =   1455
               Left            =   14640
               TabIndex        =   137
               Top             =   960
               Width           =   135
            End
            Begin VB.VScrollBar VScroll3 
               Height          =   1455
               Left            =   10920
               TabIndex        =   136
               Top             =   960
               Width           =   135
            End
            Begin VB.VScrollBar VScroll2 
               Height          =   1455
               Left            =   7200
               TabIndex        =   135
               Top             =   960
               Width           =   135
            End
            Begin VB.VScrollBar VScroll1 
               Height          =   1455
               Left            =   3480
               TabIndex        =   134
               Top             =   960
               Width           =   135
            End
            Begin VB.CommandButton Command9 
               Caption         =   "REMOVE"
               Height          =   975
               Left            =   15600
               Picture         =   "FACULITY INFORMATION DETAILS.frx":3BE81A
               Style           =   1  'Graphical
               TabIndex        =   116
               Top             =   1560
               Width           =   1095
            End
            Begin VB.CommandButton Command8 
               Caption         =   "ADD"
               Height          =   975
               Left            =   15600
               Picture         =   "FACULITY INFORMATION DETAILS.frx":4070A4
               Style           =   1  'Graphical
               TabIndex        =   115
               Top             =   240
               Width           =   1095
            End
            Begin VB.TextBox Text32 
               BackColor       =   &H80000016&
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   11280
               TabIndex        =   114
               Top             =   600
               Width           =   3495
            End
            Begin VB.TextBox Text31 
               BackColor       =   &H80000016&
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   7560
               TabIndex        =   113
               Top             =   600
               Width           =   3495
            End
            Begin VB.TextBox Text30 
               BackColor       =   &H80000016&
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   3840
               TabIndex        =   112
               Top             =   600
               Width           =   3495
            End
            Begin VB.TextBox Text29 
               BackColor       =   &H80000016&
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   120
               TabIndex        =   111
               Top             =   600
               Width           =   3495
            End
            Begin VB.ListBox List4 
               BackColor       =   &H00FF8080&
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1500
               ItemData        =   "FACULITY INFORMATION DETAILS.frx":43F0DC
               Left            =   11280
               List            =   "FACULITY INFORMATION DETAILS.frx":43F0DE
               TabIndex        =   110
               Top             =   960
               Width           =   3495
            End
            Begin VB.ListBox List3 
               BackColor       =   &H00FF8080&
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1500
               ItemData        =   "FACULITY INFORMATION DETAILS.frx":43F0E0
               Left            =   7560
               List            =   "FACULITY INFORMATION DETAILS.frx":43F0E2
               TabIndex        =   109
               Top             =   960
               Width           =   3495
            End
            Begin VB.ListBox List2 
               BackColor       =   &H00FF8080&
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1500
               ItemData        =   "FACULITY INFORMATION DETAILS.frx":43F0E4
               Left            =   3840
               List            =   "FACULITY INFORMATION DETAILS.frx":43F0E6
               TabIndex        =   108
               Top             =   960
               Width           =   3495
            End
            Begin VB.ListBox List1 
               BackColor       =   &H00FF8080&
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1500
               ItemData        =   "FACULITY INFORMATION DETAILS.frx":43F0E8
               Left            =   120
               List            =   "FACULITY INFORMATION DETAILS.frx":43F0EA
               TabIndex        =   107
               Top             =   960
               Width           =   3495
            End
            Begin VB.Line Line5 
               BorderColor     =   &H8000000B&
               X1              =   15000
               X2              =   15000
               Y1              =   120
               Y2              =   3480
            End
            Begin VB.Line Line1 
               BorderColor     =   &H8000000B&
               X1              =   3720
               X2              =   3720
               Y1              =   120
               Y2              =   2640
            End
            Begin VB.Line Line4 
               BorderColor     =   &H8000000B&
               X1              =   0
               X2              =   15000
               Y1              =   480
               Y2              =   480
            End
            Begin VB.Line Line3 
               BorderColor     =   &H8000000B&
               X1              =   11160
               X2              =   11160
               Y1              =   120
               Y2              =   2640
            End
            Begin VB.Line Line2 
               BorderColor     =   &H8000000B&
               X1              =   7440
               X2              =   7440
               Y1              =   120
               Y2              =   2640
            End
            Begin VB.Label Label61 
               BackStyle       =   0  'Transparent
               Caption         =   "PERCENTAGE"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   12360
               TabIndex        =   103
               Top             =   240
               Width           =   1935
            End
            Begin VB.Label Label60 
               BackStyle       =   0  'Transparent
               Caption         =   "YEAR OF PASSING"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   8520
               TabIndex        =   102
               Top             =   240
               Width           =   1935
            End
            Begin VB.Label Label59 
               BackStyle       =   0  'Transparent
               Caption         =   "COLLEGE/UNIVERSITY"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   4680
               TabIndex        =   101
               Top             =   240
               Width           =   1935
            End
            Begin VB.Label Label56 
               BackStyle       =   0  'Transparent
               Caption         =   "QUALIFICATION"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   960
               TabIndex        =   100
               Top             =   240
               Width           =   1935
            End
         End
         Begin VB.ComboBox Combo8 
            BackColor       =   &H80000016&
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            ItemData        =   "FACULITY INFORMATION DETAILS.frx":43F0EC
            Left            =   13440
            List            =   "FACULITY INFORMATION DETAILS.frx":43F0F6
            TabIndex        =   73
            Text            =   "SELECT"
            Top             =   2160
            Width           =   1695
         End
         Begin VB.TextBox Text27 
            BackColor       =   &H80000016&
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   9600
            TabIndex        =   72
            Top             =   2160
            Width           =   1695
         End
         Begin VB.TextBox Text26 
            BackColor       =   &H80000016&
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   5760
            TabIndex        =   71
            Top             =   2160
            Width           =   1695
         End
         Begin VB.TextBox Text25 
            BackColor       =   &H80000016&
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1920
            TabIndex        =   70
            Top             =   2160
            Width           =   1695
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   255
            Left            =   9600
            TabIndex        =   66
            Top             =   360
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   450
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd/MMM/yyyy"
            Format          =   100401155
            CurrentDate     =   43083
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000016&
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1920
            TabIndex        =   53
            Top             =   787
            Width           =   1695
         End
         Begin VB.TextBox Text3 
            BackColor       =   &H80000016&
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   5760
            TabIndex        =   52
            Top             =   783
            Width           =   1695
         End
         Begin VB.TextBox Text4 
            BackColor       =   &H80000016&
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   9600
            TabIndex        =   51
            Top             =   720
            Width           =   1695
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00FF8080&
            Caption         =   "PARMANENT ADDRESS"
            Height          =   2055
            Left            =   120
            TabIndex        =   30
            Top             =   2520
            Width           =   7455
            Begin VB.TextBox Text5 
               BackColor       =   &H80000016&
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1920
               TabIndex        =   40
               Top             =   240
               Width           =   1695
            End
            Begin VB.TextBox Text6 
               BackColor       =   &H80000016&
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1920
               TabIndex        =   39
               Top             =   600
               Width           =   1695
            End
            Begin VB.TextBox Text7 
               BackColor       =   &H80000016&
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1920
               TabIndex        =   38
               Top             =   960
               Width           =   1695
            End
            Begin VB.TextBox Text8 
               BackColor       =   &H80000016&
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1920
               TabIndex        =   37
               Top             =   1320
               Width           =   1695
            End
            Begin VB.TextBox Text9 
               BackColor       =   &H80000016&
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1920
               TabIndex        =   36
               Top             =   1680
               Width           =   1695
            End
            Begin VB.TextBox Text10 
               BackColor       =   &H80000016&
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   5640
               TabIndex        =   35
               Top             =   240
               Width           =   1695
            End
            Begin VB.TextBox Text11 
               BackColor       =   &H80000016&
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   5640
               TabIndex        =   34
               Top             =   600
               Width           =   1695
            End
            Begin VB.TextBox Text12 
               BackColor       =   &H80000016&
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   5640
               TabIndex        =   33
               Top             =   960
               Width           =   1695
            End
            Begin VB.TextBox Text13 
               BackColor       =   &H80000016&
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   5640
               TabIndex        =   32
               Top             =   1320
               Width           =   1695
            End
            Begin VB.TextBox Text14 
               BackColor       =   &H80000016&
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   5640
               TabIndex        =   31
               Top             =   1680
               Width           =   1695
            End
            Begin VB.Label Label84 
               BackStyle       =   0  'Transparent
               Caption         =   "*"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   135
               Left            =   4560
               TabIndex        =   160
               Top             =   1560
               Width           =   135
            End
            Begin VB.Label Label83 
               BackStyle       =   0  'Transparent
               Caption         =   "*"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   135
               Left            =   4680
               TabIndex        =   159
               Top             =   1200
               Width           =   135
            End
            Begin VB.Label Label82 
               BackStyle       =   0  'Transparent
               Caption         =   "*"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   135
               Left            =   4800
               TabIndex        =   158
               Top             =   480
               Width           =   135
            End
            Begin VB.Label Label81 
               BackStyle       =   0  'Transparent
               Caption         =   "*"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   135
               Left            =   4560
               TabIndex        =   157
               Top             =   120
               Width           =   135
            End
            Begin VB.Label Label80 
               BackStyle       =   0  'Transparent
               Caption         =   "*"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   135
               Left            =   960
               TabIndex        =   156
               Top             =   1560
               Width           =   135
            End
            Begin VB.Label Label79 
               BackStyle       =   0  'Transparent
               Caption         =   "*"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   135
               Left            =   1200
               TabIndex        =   155
               Top             =   1200
               Width           =   135
            End
            Begin VB.Label Label78 
               BackStyle       =   0  'Transparent
               Caption         =   "*"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   135
               Left            =   600
               TabIndex        =   154
               Top             =   840
               Width           =   135
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "HOUSE NUMBER :"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   50
               Top             =   240
               Width           =   1335
            End
            Begin VB.Label Label11 
               BackStyle       =   0  'Transparent
               Caption         =   "VILLAGE :"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   49
               Top             =   600
               Width           =   855
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "CITY :"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   48
               Top             =   960
               Width           =   495
            End
            Begin VB.Label Label13 
               BackStyle       =   0  'Transparent
               Caption         =   "POST OFFICE :"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   47
               Top             =   1320
               Width           =   1095
            End
            Begin VB.Label Label14 
               BackStyle       =   0  'Transparent
               Caption         =   "DISTRICT :"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   46
               Top             =   1680
               Width           =   855
            End
            Begin VB.Label Label15 
               BackStyle       =   0  'Transparent
               Caption         =   "STREET :"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3840
               TabIndex        =   45
               Top             =   240
               Width           =   735
            End
            Begin VB.Label Label16 
               BackStyle       =   0  'Transparent
               Caption         =   "LANDMARK :"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3840
               TabIndex        =   44
               Top             =   600
               Width           =   975
            End
            Begin VB.Label Label17 
               BackStyle       =   0  'Transparent
               Caption         =   "POLICE STATION :"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3840
               TabIndex        =   43
               Top             =   960
               Width           =   1455
            End
            Begin VB.Label Label18 
               BackStyle       =   0  'Transparent
               Caption         =   "PIN CODE :"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3840
               TabIndex        =   42
               Top             =   1320
               Width           =   855
            End
            Begin VB.Label Label19 
               BackStyle       =   0  'Transparent
               Caption         =   "STATE :"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3840
               TabIndex        =   41
               Top             =   1680
               Width           =   735
            End
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00FF8080&
            Caption         =   "LOCAL ADDRESS"
            Height          =   2055
            Left            =   7680
            TabIndex        =   9
            Top             =   2520
            Width           =   7455
            Begin VB.TextBox Text15 
               BackColor       =   &H80000016&
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1920
               TabIndex        =   19
               Top             =   240
               Width           =   1695
            End
            Begin VB.TextBox Text16 
               BackColor       =   &H80000016&
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1920
               TabIndex        =   18
               Top             =   600
               Width           =   1695
            End
            Begin VB.TextBox Text17 
               BackColor       =   &H80000016&
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1920
               TabIndex        =   17
               Top             =   960
               Width           =   1695
            End
            Begin VB.TextBox Text18 
               BackColor       =   &H80000016&
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1920
               TabIndex        =   16
               Top             =   1320
               Width           =   1695
            End
            Begin VB.TextBox Text19 
               BackColor       =   &H80000016&
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1920
               TabIndex        =   15
               Top             =   1680
               Width           =   1695
            End
            Begin VB.TextBox Text20 
               BackColor       =   &H80000016&
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   5640
               TabIndex        =   14
               Top             =   240
               Width           =   1695
            End
            Begin VB.TextBox Text21 
               BackColor       =   &H80000016&
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   5640
               TabIndex        =   13
               Top             =   600
               Width           =   1695
            End
            Begin VB.TextBox Text22 
               BackColor       =   &H80000016&
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   5640
               TabIndex        =   12
               Top             =   960
               Width           =   1695
            End
            Begin VB.TextBox Text23 
               BackColor       =   &H80000016&
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   5640
               TabIndex        =   11
               Top             =   1320
               Width           =   1695
            End
            Begin VB.TextBox Text24 
               BackColor       =   &H80000016&
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   5640
               TabIndex        =   10
               Top             =   1680
               Width           =   1695
            End
            Begin VB.Label Label85 
               BackStyle       =   0  'Transparent
               Caption         =   "*"
               ForeColor       =   &H000000FF&
               Height          =   135
               Left            =   600
               TabIndex        =   173
               Top             =   840
               Width           =   135
            End
            Begin VB.Label Label86 
               BackStyle       =   0  'Transparent
               Caption         =   "*"
               ForeColor       =   &H000000FF&
               Height          =   135
               Left            =   1200
               TabIndex        =   172
               Top             =   1200
               Width           =   135
            End
            Begin VB.Label Label87 
               BackStyle       =   0  'Transparent
               Caption         =   "*"
               ForeColor       =   &H000000FF&
               Height          =   135
               Left            =   960
               TabIndex        =   171
               Top             =   1560
               Width           =   135
            End
            Begin VB.Label Label89 
               BackStyle       =   0  'Transparent
               Caption         =   "*"
               ForeColor       =   &H000000FF&
               Height          =   135
               Left            =   4680
               TabIndex        =   170
               Top             =   1200
               Width           =   135
            End
            Begin VB.Label Label90 
               BackStyle       =   0  'Transparent
               Caption         =   "*"
               ForeColor       =   &H000000FF&
               Height          =   135
               Left            =   4440
               TabIndex        =   169
               Top             =   1560
               Width           =   135
            End
            Begin VB.Label Label88 
               BackStyle       =   0  'Transparent
               Caption         =   "*"
               ForeColor       =   &H000000FF&
               Height          =   135
               Left            =   4560
               TabIndex        =   161
               Top             =   120
               Width           =   135
            End
            Begin VB.Label Label20 
               BackStyle       =   0  'Transparent
               Caption         =   "HOUSE NUMBER :"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   29
               Top             =   240
               Width           =   1335
            End
            Begin VB.Label Label21 
               BackStyle       =   0  'Transparent
               Caption         =   "VILLAGE :"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   28
               Top             =   600
               Width           =   855
            End
            Begin VB.Label Label22 
               BackStyle       =   0  'Transparent
               Caption         =   "CITY :"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   27
               Top             =   960
               Width           =   495
            End
            Begin VB.Label Label23 
               BackStyle       =   0  'Transparent
               Caption         =   "POST OFFICE :"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   26
               Top             =   1320
               Width           =   1095
            End
            Begin VB.Label Label24 
               BackStyle       =   0  'Transparent
               Caption         =   "DISTRICT :"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   25
               Top             =   1680
               Width           =   855
            End
            Begin VB.Label Label25 
               BackStyle       =   0  'Transparent
               Caption         =   "STREET :"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3840
               TabIndex        =   24
               Top             =   240
               Width           =   735
            End
            Begin VB.Label Label26 
               BackStyle       =   0  'Transparent
               Caption         =   "LANDMARK :"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3840
               TabIndex        =   23
               Top             =   600
               Width           =   975
            End
            Begin VB.Label Label27 
               BackStyle       =   0  'Transparent
               Caption         =   "POLICE STATION :"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3840
               TabIndex        =   22
               Top             =   960
               Width           =   1455
            End
            Begin VB.Label Label28 
               BackStyle       =   0  'Transparent
               Caption         =   "PIN CODE :"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3840
               TabIndex        =   21
               Top             =   1320
               Width           =   855
            End
            Begin VB.Label Label29 
               BackStyle       =   0  'Transparent
               Caption         =   "STATE :"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3840
               TabIndex        =   20
               Top             =   1680
               Width           =   615
            End
         End
         Begin VB.ComboBox Combo1 
            BackColor       =   &H80000016&
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            ItemData        =   "FACULITY INFORMATION DETAILS.frx":43F109
            Left            =   5760
            List            =   "FACULITY INFORMATION DETAILS.frx":43F113
            TabIndex        =   8
            Text            =   "SELECT"
            Top             =   1221
            Width           =   1695
         End
         Begin VB.ComboBox Combo2 
            BackColor       =   &H80000016&
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            ItemData        =   "FACULITY INFORMATION DETAILS.frx":43F125
            Left            =   9600
            List            =   "FACULITY INFORMATION DETAILS.frx":43F138
            TabIndex        =   7
            Text            =   "SELECT"
            Top             =   1250
            Width           =   1695
         End
         Begin VB.ComboBox Combo3 
            BackColor       =   &H80000016&
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            ItemData        =   "FACULITY INFORMATION DETAILS.frx":43F160
            Left            =   13440
            List            =   "FACULITY INFORMATION DETAILS.frx":43F170
            TabIndex        =   6
            Text            =   "SELECT"
            Top             =   1251
            Width           =   1695
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H80000016&
            Caption         =   "REGULAR"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1440
            TabIndex        =   5
            Top             =   360
            Width           =   1095
         End
         Begin VB.ComboBox Combo4 
            BackColor       =   &H80000016&
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            ItemData        =   "FACULITY INFORMATION DETAILS.frx":43F197
            Left            =   1920
            List            =   "FACULITY INFORMATION DETAILS.frx":43F199
            TabIndex        =   4
            Text            =   "SELECT"
            Top             =   1686
            Width           =   1695
         End
         Begin VB.ComboBox Combo5 
            BackColor       =   &H80000016&
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            ItemData        =   "FACULITY INFORMATION DETAILS.frx":43F19B
            Left            =   5760
            List            =   "FACULITY INFORMATION DETAILS.frx":43F1AB
            TabIndex        =   3
            Text            =   "SELECT"
            Top             =   1689
            Width           =   1695
         End
         Begin VB.ComboBox Combo6 
            BackColor       =   &H80000016&
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            ItemData        =   "FACULITY INFORMATION DETAILS.frx":43F1C6
            Left            =   9600
            List            =   "FACULITY INFORMATION DETAILS.frx":43F1C8
            TabIndex        =   2
            Text            =   "SELECT"
            Top             =   1705
            Width           =   1695
         End
         Begin VB.Label Label93 
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   135
            Left            =   12480
            TabIndex        =   176
            Top             =   240
            Width           =   135
         End
         Begin VB.Label Label92 
            BackStyle       =   0  'Transparent
            Caption         =   "SALARY :"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   11640
            TabIndex        =   174
            Top             =   445
            Width           =   855
         End
         Begin VB.Label Label66 
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   135
            Left            =   960
            TabIndex        =   168
            Top             =   240
            Width           =   135
         End
         Begin VB.Label Label32 
            BackStyle       =   0  'Transparent
            Caption         =   "SUBJECT :"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7680
            TabIndex        =   167
            Top             =   1710
            Width           =   735
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "CATEGORY :"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7680
            TabIndex        =   166
            Top             =   1260
            Width           =   975
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "SURNAME :"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7680
            TabIndex        =   165
            Top             =   810
            Width           =   975
         End
         Begin VB.Label Label38 
            BackStyle       =   0  'Transparent
            Caption         =   "CONTACT NO :"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7680
            TabIndex        =   164
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label Label91 
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   135
            Left            =   1200
            TabIndex        =   162
            Top             =   600
            Width           =   135
         End
         Begin VB.Label Label77 
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   135
            Left            =   12720
            TabIndex        =   153
            Top             =   1625
            Width           =   135
         End
         Begin VB.Label Label76 
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   135
            Left            =   12960
            TabIndex        =   152
            Top             =   770
            Width           =   135
         End
         Begin VB.Label Label75 
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   135
            Left            =   8880
            TabIndex        =   151
            Top             =   2040
            Width           =   135
         End
         Begin VB.Label Label74 
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   135
            Left            =   8520
            TabIndex        =   150
            Top             =   1560
            Width           =   135
         End
         Begin VB.Label Label73 
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   135
            Left            =   9120
            TabIndex        =   149
            Top             =   240
            Width           =   135
         End
         Begin VB.Label Label72 
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   135
            Left            =   5040
            TabIndex        =   148
            Top             =   240
            Width           =   135
         End
         Begin VB.Label Label71 
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   135
            Left            =   4800
            TabIndex        =   147
            Top             =   1200
            Width           =   135
         End
         Begin VB.Label Label70 
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   135
            Left            =   4680
            TabIndex        =   146
            Top             =   1680
            Width           =   135
         End
         Begin VB.Label Label69 
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   135
            Left            =   5160
            TabIndex        =   145
            Top             =   2160
            Width           =   135
         End
         Begin VB.Label Label68 
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   135
            Left            =   1320
            TabIndex        =   144
            Top             =   1560
            Width           =   135
         End
         Begin VB.Label Label67 
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   135
            Left            =   1440
            TabIndex        =   143
            Top             =   1080
            Width           =   135
         End
         Begin VB.Shape Shape2 
            BorderColor     =   &H8000000B&
            Height          =   2295
            Left            =   15240
            Top             =   2280
            Width           =   2295
         End
         Begin VB.Label Label40 
            BackStyle       =   0  'Transparent
            Caption         =   "FATHER'S NAME :"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   11640
            TabIndex        =   80
            Top             =   873
            Width           =   1335
         End
         Begin VB.Label Label39 
            BackStyle       =   0  'Transparent
            Caption         =   "NATIONALITY :"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   11640
            TabIndex        =   69
            Top             =   2160
            Width           =   1335
         End
         Begin VB.Label Label37 
            BackStyle       =   0  'Transparent
            Caption         =   "PAN CARD NO :"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3960
            TabIndex        =   68
            Top             =   2280
            Width           =   1215
         End
         Begin VB.Label Label36 
            BackStyle       =   0  'Transparent
            Caption         =   "AADHAR NO :"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   67
            Top             =   2160
            Width           =   1095
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "FACULTY ID :"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3960
            TabIndex        =   64
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "FIRST NAME :"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   63
            Top             =   810
            Width           =   1095
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "MIDDLE NAME :"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3960
            TabIndex        =   62
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "DATE OF BIRTH :"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   61
            Top             =   1260
            Width           =   1335
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "GENDER :"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3960
            TabIndex        =   60
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "RELIGION :"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   11640
            TabIndex        =   59
            Top             =   1301
            Width           =   975
         End
         Begin VB.Label Label30 
            BackStyle       =   0  'Transparent
            Caption         =   "DEPARTMENT :"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   58
            Top             =   1710
            Width           =   1215
         End
         Begin VB.Label Label31 
            BackStyle       =   0  'Transparent
            Caption         =   "COURSE :"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3960
            TabIndex        =   57
            Top             =   1800
            Width           =   735
         End
         Begin VB.Label Label33 
            BackStyle       =   0  'Transparent
            Caption         =   "STATUS :"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   56
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label34 
            BackStyle       =   0  'Transparent
            Caption         =   "DATE OF JOINING :"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7680
            TabIndex        =   55
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label35 
            BackStyle       =   0  'Transparent
            Caption         =   "EXPERIENCE :"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   11640
            TabIndex        =   54
            Top             =   1729
            Width           =   1095
         End
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H8000000B&
         Height          =   8175
         Left            =   17880
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Label Label58 
      Caption         =   "Label58"
      Height          =   495
      Left            =   9480
      TabIndex        =   99
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label57 
      Caption         =   "Label57"
      Height          =   495
      Left            =   9480
      TabIndex        =   98
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label55 
      Caption         =   "Label55"
      Height          =   495
      Left            =   9480
      TabIndex        =   97
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label54 
      Caption         =   "Label54"
      Height          =   495
      Left            =   9480
      TabIndex        =   96
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label53 
      Caption         =   "Label53"
      Height          =   495
      Left            =   9480
      TabIndex        =   95
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label52 
      Caption         =   "Label52"
      Height          =   495
      Left            =   9480
      TabIndex        =   94
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label51 
      Caption         =   "Label51"
      Height          =   495
      Left            =   9480
      TabIndex        =   93
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label50 
      Caption         =   "Label50"
      Height          =   495
      Left            =   9480
      TabIndex        =   92
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label49 
      Caption         =   "Label49"
      Height          =   495
      Left            =   9480
      TabIndex        =   91
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label48 
      Caption         =   "Label48"
      Height          =   495
      Left            =   9480
      TabIndex        =   90
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label47 
      Caption         =   "Label47"
      Height          =   495
      Left            =   9480
      TabIndex        =   89
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label46 
      Caption         =   "Label46"
      Height          =   495
      Left            =   9480
      TabIndex        =   88
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label45 
      Caption         =   "Label45"
      Height          =   495
      Left            =   9480
      TabIndex        =   87
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label44 
      Caption         =   "Label44"
      Height          =   495
      Left            =   9480
      TabIndex        =   86
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label43 
      Caption         =   "Label43"
      Height          =   495
      Left            =   9480
      TabIndex        =   85
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label42 
      Caption         =   "Label42"
      Height          =   495
      Left            =   9480
      TabIndex        =   84
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label41 
      Caption         =   "Label41"
      Height          =   495
      Left            =   9480
      TabIndex        =   83
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "FACULTY "
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000016&
      Height          =   615
      Left            =   4440
      TabIndex        =   65
      Top             =   0
      Width           =   7695
   End
End
Attribute VB_Name = "frmFACULTYINFORMATION"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command6_Click()
cd2.Filter = "picture file |*.jpg"
cd2.ShowOpen
If cd2.FileName <> " " Then
Image1.Picture = LoadPicture(cd2.FileName)
End If
End Sub

