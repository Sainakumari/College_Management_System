VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmFACULTY 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " FACULTY "
   ClientHeight    =   10935
   ClientLeft      =   375
   ClientTop       =   210
   ClientWidth     =   20325
   BeginProperty Font 
      Name            =   "Cambria"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   Picture         =   "FACULTY DETAILS.frx":0000
   ScaleHeight     =   10935
   ScaleWidth      =   20325
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10335
      Left            =   240
      TabIndex        =   63
      Top             =   600
      Width           =   19935
      Begin VB.CommandButton Command13 
         Caption         =   "DATA GRID"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   18360
         Picture         =   "FACULTY DETAILS.frx":1E9608
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   8910
         Width           =   1095
      End
      Begin VB.CommandButton refresh 
         Caption         =   "REFRESH"
         Height          =   975
         Left            =   18240
         Picture         =   "FACULTY DETAILS.frx":1E9E57
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   6096
         Width           =   1095
      End
      Begin VB.CommandButton print 
         Caption         =   "PRINT"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   18300
         Picture         =   "FACULTY DETAILS.frx":1EA4FB
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   7440
         Width           =   1095
      End
      Begin VB.CommandButton update 
         Caption         =   "UPDATE"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   18240
         Picture         =   "FACULTY DETAILS.frx":1EAD7E
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   4752
         Width           =   1095
      End
      Begin VB.CommandButton delete 
         Caption         =   "DELETE"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   18240
         Picture         =   "FACULTY DETAILS.frx":1EB50E
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   3408
         Width           =   1095
      End
      Begin VB.CommandButton save 
         Caption         =   "SAVE"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   18240
         Picture         =   "FACULTY DETAILS.frx":1EBA7A
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton view 
         Caption         =   "VIEW"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   18240
         Picture         =   "FACULTY DETAILS.frx":1EC254
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FF8080&
         Caption         =   "  PERSONAL PROFILE"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   10155
         Left            =   120
         TabIndex        =   64
         Top             =   120
         Width           =   17655
         Begin VB.ComboBox STATUS 
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
            ItemData        =   "FACULTY DETAILS.frx":1EC60B
            Left            =   1920
            List            =   "FACULTY DETAILS.frx":1EC615
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   360
            Width           =   1695
         End
         Begin VB.OptionButton Option4 
            Caption         =   "YEAR"
            Height          =   225
            Left            =   13200
            TabIndex        =   16
            Top             =   1740
            Width           =   1185
         End
         Begin VB.OptionButton Option3 
            Caption         =   "MONTH"
            Height          =   225
            Left            =   11640
            TabIndex        =   15
            Top             =   1740
            Width           =   1185
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00FF8080&
            Caption         =   "IF LOCAL ADDRESS  IS SAME AS FOR THE PERMANENT ADDRESS"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   150
            TabIndex        =   41
            Top             =   4650
            Width           =   6345
         End
         Begin VB.TextBox doj1 
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
            Left            =   9600
            TabIndex        =   2
            Top             =   300
            Width           =   1695
         End
         Begin VB.TextBox dob1 
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
            Left            =   1920
            TabIndex        =   8
            Top             =   1200
            Width           =   1695
         End
         Begin VB.TextBox salary 
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
            TabIndex        =   3
            Top             =   360
            Width           =   1695
         End
         Begin VB.Frame Frame7 
            BackColor       =   &H00FF8080&
            Caption         =   "EXPERIENCE"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2415
            Left            =   120
            TabIndex        =   132
            Top             =   7650
            Width           =   17415
            Begin VB.VScrollBar VScroll10 
               Height          =   1335
               Left            =   14760
               TabIndex        =   148
               Top             =   960
               Width           =   135
            End
            Begin VB.VScrollBar VScroll8 
               Height          =   1335
               Left            =   10920
               TabIndex        =   147
               Top             =   960
               Width           =   135
            End
            Begin VB.VScrollBar VScroll7 
               Height          =   1335
               Left            =   7200
               TabIndex        =   146
               Top             =   960
               Width           =   135
            End
            Begin VB.VScrollBar VScroll5 
               Height          =   1335
               Left            =   3480
               TabIndex        =   145
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
               ItemData        =   "FACULTY DETAILS.frx":1EC62C
               Left            =   11400
               List            =   "FACULTY DETAILS.frx":1EC62E
               TabIndex        =   140
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
               ItemData        =   "FACULTY DETAILS.frx":1EC630
               Left            =   7560
               List            =   "FACULTY DETAILS.frx":1EC632
               TabIndex        =   139
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
               ItemData        =   "FACULTY DETAILS.frx":1EC634
               Left            =   3840
               List            =   "FACULTY DETAILS.frx":1EC636
               TabIndex        =   138
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
               ItemData        =   "FACULTY DETAILS.frx":1EC638
               Left            =   120
               List            =   "FACULTY DETAILS.frx":1EC63A
               TabIndex        =   137
               Top             =   960
               Width           =   3495
            End
            Begin VB.TextBox dateofwork 
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
               TabIndex        =   49
               Top             =   600
               Width           =   3495
            End
            Begin VB.TextBox year 
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
               TabIndex        =   48
               Top             =   600
               Width           =   3495
            End
            Begin VB.TextBox designation 
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
               TabIndex        =   47
               Top             =   600
               Width           =   3495
            End
            Begin VB.TextBox organisation 
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
               TabIndex        =   46
               Top             =   600
               Width           =   3495
            End
            Begin VB.CommandButton remove2 
               Caption         =   "REMOVE"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   975
               Left            =   15600
               Picture         =   "FACULTY DETAILS.frx":1EC63C
               Style           =   1  'Graphical
               TabIndex        =   55
               Top             =   1320
               Width           =   1095
            End
            Begin VB.CommandButton add2 
               Caption         =   "ADD"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   975
               Left            =   15600
               Picture         =   "FACULTY DETAILS.frx":234EC6
               Style           =   1  'Graphical
               TabIndex        =   54
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
               Height          =   255
               Left            =   11880
               TabIndex        =   136
               Top             =   240
               Width           =   1935
            End
            Begin VB.Label Label64 
               BackStyle       =   0  'Transparent
               Caption         =   "YEAR"
               Height          =   255
               Left            =   8520
               TabIndex        =   135
               Top             =   240
               Width           =   1935
            End
            Begin VB.Label Label63 
               BackStyle       =   0  'Transparent
               Caption         =   "DESIGNATION/PROFILE"
               Height          =   255
               Left            =   4320
               TabIndex        =   134
               Top             =   240
               Width           =   1935
            End
            Begin VB.Label Label62 
               BackStyle       =   0  'Transparent
               Caption         =   "ORGANISATION"
               Height          =   255
               Left            =   720
               TabIndex        =   133
               Top             =   240
               Width           =   1935
            End
         End
         Begin VB.TextBox exp 
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
            Left            =   9540
            TabIndex        =   14
            Top             =   1719
            Width           =   1695
         End
         Begin VB.CommandButton Command7 
            Caption         =   "REMOVE"
            Height          =   975
            Left            =   15720
            Picture         =   "FACULTY DETAILS.frx":235780
            Style           =   1  'Graphical
            TabIndex        =   51
            Top             =   3480
            Width           =   1095
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00FF8080&
            Caption         =   "FACULTY'S PHOTO"
            Height          =   1935
            Left            =   15240
            TabIndex        =   127
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
         Begin MSComCtl2.DTPicker dob 
            Height          =   285
            Left            =   1920
            TabIndex        =   126
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
            Format          =   105054211
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
            Height          =   975
            Left            =   15720
            Picture         =   "FACULTY DETAILS.frx":27E00A
            Style           =   1  'Graphical
            TabIndex        =   50
            Top             =   2400
            Width           =   1095
         End
         Begin VB.TextBox fathername 
            BackColor       =   &H80000016&
            Height          =   270
            Left            =   13440
            TabIndex        =   7
            Top             =   813
            Width           =   1695
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00FF8080&
            Caption         =   "QUALIFICATION"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2655
            Left            =   150
            TabIndex        =   103
            Top             =   5070
            Width           =   17415
            Begin VB.VScrollBar VScroll4 
               Height          =   1455
               Left            =   14640
               TabIndex        =   144
               Top             =   960
               Width           =   135
            End
            Begin VB.VScrollBar VScroll3 
               Height          =   1455
               Left            =   10920
               TabIndex        =   143
               Top             =   960
               Width           =   135
            End
            Begin VB.VScrollBar VScroll2 
               Height          =   1455
               Left            =   7200
               TabIndex        =   142
               Top             =   960
               Width           =   135
            End
            Begin VB.VScrollBar VScroll1 
               Height          =   1455
               Left            =   3480
               TabIndex        =   141
               Top             =   960
               Width           =   135
            End
            Begin VB.CommandButton remove1 
               Caption         =   "REMOVE"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   975
               Left            =   15600
               Picture         =   "FACULTY DETAILS.frx":27E8B7
               Style           =   1  'Graphical
               TabIndex        =   53
               Top             =   1560
               Width           =   1095
            End
            Begin VB.CommandButton add1 
               Caption         =   "ADD"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   975
               Left            =   15600
               Picture         =   "FACULTY DETAILS.frx":2C7141
               Style           =   1  'Graphical
               TabIndex        =   52
               Top             =   240
               Width           =   1095
            End
            Begin VB.TextBox percentage 
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
               TabIndex        =   45
               Top             =   600
               Width           =   3495
            End
            Begin VB.TextBox yearofpass 
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
               Left            =   7590
               TabIndex        =   44
               Top             =   600
               Width           =   3495
            End
            Begin VB.TextBox college 
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
               TabIndex        =   43
               Top             =   600
               Width           =   3495
            End
            Begin VB.TextBox quali 
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
               TabIndex        =   42
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
               ItemData        =   "FACULTY DETAILS.frx":2C79FB
               Left            =   11280
               List            =   "FACULTY DETAILS.frx":2C79FD
               TabIndex        =   131
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
               ItemData        =   "FACULTY DETAILS.frx":2C79FF
               Left            =   7560
               List            =   "FACULTY DETAILS.frx":2C7A01
               TabIndex        =   130
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
               ItemData        =   "FACULTY DETAILS.frx":2C7A03
               Left            =   3840
               List            =   "FACULTY DETAILS.frx":2C7A05
               TabIndex        =   129
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
               ItemData        =   "FACULTY DETAILS.frx":2C7A07
               Left            =   120
               List            =   "FACULTY DETAILS.frx":2C7A09
               TabIndex        =   128
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
               Height          =   255
               Left            =   12360
               TabIndex        =   125
               Top             =   240
               Width           =   1935
            End
            Begin VB.Label Label60 
               BackStyle       =   0  'Transparent
               Caption         =   "YEAR OF PASSING"
               Height          =   255
               Left            =   8520
               TabIndex        =   124
               Top             =   240
               Width           =   1935
            End
            Begin VB.Label Label59 
               BackStyle       =   0  'Transparent
               Caption         =   "COLLEGE/UNIVERSITY"
               Height          =   255
               Left            =   4680
               TabIndex        =   123
               Top             =   240
               Width           =   1935
            End
            Begin VB.Label Label56 
               BackStyle       =   0  'Transparent
               Caption         =   "QUALIFICATION"
               Height          =   255
               Left            =   960
               TabIndex        =   122
               Top             =   240
               Width           =   1935
            End
         End
         Begin VB.ComboBox nationality 
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
            ItemData        =   "FACULTY DETAILS.frx":2C7A0B
            Left            =   13440
            List            =   "FACULTY DETAILS.frx":2C7A15
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   2160
            Width           =   1695
         End
         Begin VB.TextBox contactno 
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
            TabIndex        =   19
            Top             =   2160
            Width           =   1695
         End
         Begin VB.TextBox pano 
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
            TabIndex        =   18
            Top             =   2160
            Width           =   1695
         End
         Begin VB.TextBox aadharno 
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
            TabIndex        =   17
            Top             =   2160
            Width           =   1695
         End
         Begin MSComCtl2.DTPicker doj 
            Height          =   255
            Left            =   9600
            TabIndex        =   99
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
            Format          =   104660995
            CurrentDate     =   43083
         End
         Begin VB.TextBox firstname 
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
            TabIndex        =   4
            Top             =   787
            Width           =   1695
         End
         Begin VB.TextBox middlename 
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
            Left            =   5820
            TabIndex        =   5
            Top             =   720
            Width           =   1695
         End
         Begin VB.TextBox surname 
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
            TabIndex        =   6
            Top             =   720
            Width           =   1695
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00FF8080&
            Caption         =   "PARMANENT ADDRESS"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2055
            Left            =   120
            TabIndex        =   76
            Top             =   2520
            Width           =   7455
            Begin VB.TextBox hno 
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
               TabIndex        =   21
               Top             =   240
               Width           =   1695
            End
            Begin VB.TextBox village 
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
               Left            =   1950
               TabIndex        =   23
               Top             =   600
               Width           =   1695
            End
            Begin VB.TextBox city 
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
               TabIndex        =   25
               Top             =   960
               Width           =   1695
            End
            Begin VB.TextBox postoffice 
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
               TabIndex        =   27
               Top             =   1320
               Width           =   1695
            End
            Begin VB.TextBox district 
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
               TabIndex        =   29
               Top             =   1680
               Width           =   1695
            End
            Begin VB.TextBox street 
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
               TabIndex        =   22
               Top             =   240
               Width           =   1695
            End
            Begin VB.TextBox landmarkk 
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
               TabIndex        =   24
               Top             =   600
               Width           =   1695
            End
            Begin VB.TextBox policestation 
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
               TabIndex        =   26
               Top             =   960
               Width           =   1695
            End
            Begin VB.TextBox pincode 
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
               TabIndex        =   28
               Top             =   1320
               Width           =   1695
            End
            Begin VB.TextBox state 
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
               TabIndex        =   30
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
               TabIndex        =   164
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
               TabIndex        =   163
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
               TabIndex        =   162
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
               TabIndex        =   161
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
               TabIndex        =   160
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
               TabIndex        =   159
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
               TabIndex        =   158
               Top             =   840
               Width           =   135
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "HOUSE NUMBER :"
               Height          =   255
               Left            =   120
               TabIndex        =   86
               Top             =   240
               Width           =   1335
            End
            Begin VB.Label Label11 
               BackStyle       =   0  'Transparent
               Caption         =   "VILLAGE :"
               Height          =   255
               Left            =   120
               TabIndex        =   85
               Top             =   600
               Width           =   855
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "CITY :"
               Height          =   255
               Left            =   120
               TabIndex        =   84
               Top             =   960
               Width           =   495
            End
            Begin VB.Label Label13 
               BackStyle       =   0  'Transparent
               Caption         =   "POST OFFICE :"
               Height          =   255
               Left            =   120
               TabIndex        =   83
               Top             =   1320
               Width           =   1095
            End
            Begin VB.Label Label14 
               BackStyle       =   0  'Transparent
               Caption         =   "DISTRICT :"
               Height          =   255
               Left            =   120
               TabIndex        =   82
               Top             =   1680
               Width           =   855
            End
            Begin VB.Label Label15 
               BackStyle       =   0  'Transparent
               Caption         =   "STREET :"
               Height          =   255
               Left            =   3840
               TabIndex        =   81
               Top             =   240
               Width           =   735
            End
            Begin VB.Label Label16 
               BackStyle       =   0  'Transparent
               Caption         =   "LANDMARK :"
               Height          =   255
               Left            =   3840
               TabIndex        =   80
               Top             =   600
               Width           =   975
            End
            Begin VB.Label Label17 
               BackStyle       =   0  'Transparent
               Caption         =   "POLICE STATION :"
               Height          =   255
               Left            =   3840
               TabIndex        =   79
               Top             =   960
               Width           =   1455
            End
            Begin VB.Label Label18 
               BackStyle       =   0  'Transparent
               Caption         =   "PIN CODE :"
               Height          =   255
               Left            =   3840
               TabIndex        =   78
               Top             =   1320
               Width           =   855
            End
            Begin VB.Label Label19 
               BackStyle       =   0  'Transparent
               Caption         =   "STATE :"
               Height          =   255
               Left            =   3840
               TabIndex        =   77
               Top             =   1680
               Width           =   735
            End
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00FF8080&
            Caption         =   "LOCAL ADDRESS"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2055
            Left            =   7680
            TabIndex        =   65
            Top             =   2520
            Width           =   7455
            Begin VB.TextBox hno2 
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
               TabIndex        =   31
               Top             =   240
               Width           =   1695
            End
            Begin VB.TextBox village2 
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
               TabIndex        =   33
               Top             =   600
               Width           =   1695
            End
            Begin VB.TextBox city2 
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
               TabIndex        =   35
               Top             =   960
               Width           =   1695
            End
            Begin VB.TextBox postoffice2 
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
            Begin VB.TextBox district2 
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
               Top             =   1680
               Width           =   1695
            End
            Begin VB.TextBox street2 
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
               Top             =   240
               Width           =   1695
            End
            Begin VB.TextBox landmark 
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
            Begin VB.TextBox policestation2 
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
               TabIndex        =   36
               Top             =   960
               Width           =   1695
            End
            Begin VB.TextBox pincode2 
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
               TabIndex        =   38
               Top             =   1320
               Width           =   1695
            End
            Begin VB.TextBox state2 
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
               TabIndex        =   40
               Top             =   1710
               Width           =   1695
            End
            Begin VB.Label Label85 
               BackStyle       =   0  'Transparent
               Caption         =   "*"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   135
               Left            =   600
               TabIndex        =   175
               Top             =   840
               Width           =   135
            End
            Begin VB.Label Label86 
               BackStyle       =   0  'Transparent
               Caption         =   "*"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   135
               Left            =   1200
               TabIndex        =   174
               Top             =   1200
               Width           =   135
            End
            Begin VB.Label Label87 
               BackStyle       =   0  'Transparent
               Caption         =   "*"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   135
               Left            =   960
               TabIndex        =   173
               Top             =   1560
               Width           =   135
            End
            Begin VB.Label Label89 
               BackStyle       =   0  'Transparent
               Caption         =   "*"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   135
               Left            =   4680
               TabIndex        =   172
               Top             =   1200
               Width           =   135
            End
            Begin VB.Label Label90 
               BackStyle       =   0  'Transparent
               Caption         =   "*"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   135
               Left            =   4440
               TabIndex        =   171
               Top             =   1560
               Width           =   135
            End
            Begin VB.Label Label88 
               BackStyle       =   0  'Transparent
               Caption         =   "*"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   135
               Left            =   4560
               TabIndex        =   165
               Top             =   120
               Width           =   135
            End
            Begin VB.Label Label20 
               BackStyle       =   0  'Transparent
               Caption         =   "HOUSE NUMBER :"
               Height          =   255
               Left            =   120
               TabIndex        =   75
               Top             =   240
               Width           =   1335
            End
            Begin VB.Label Label21 
               BackStyle       =   0  'Transparent
               Caption         =   "VILLAGE :"
               Height          =   255
               Left            =   120
               TabIndex        =   74
               Top             =   600
               Width           =   855
            End
            Begin VB.Label Label22 
               BackStyle       =   0  'Transparent
               Caption         =   "CITY :"
               Height          =   255
               Left            =   120
               TabIndex        =   73
               Top             =   960
               Width           =   495
            End
            Begin VB.Label Label23 
               BackStyle       =   0  'Transparent
               Caption         =   "POST OFFICE :"
               Height          =   255
               Left            =   120
               TabIndex        =   72
               Top             =   1320
               Width           =   1095
            End
            Begin VB.Label Label24 
               BackStyle       =   0  'Transparent
               Caption         =   "DISTRICT :"
               Height          =   255
               Left            =   120
               TabIndex        =   71
               Top             =   1680
               Width           =   855
            End
            Begin VB.Label Label25 
               BackStyle       =   0  'Transparent
               Caption         =   "STREET :"
               Height          =   255
               Left            =   3840
               TabIndex        =   70
               Top             =   240
               Width           =   735
            End
            Begin VB.Label Label26 
               BackStyle       =   0  'Transparent
               Caption         =   "LANDMARK :"
               Height          =   255
               Left            =   3840
               TabIndex        =   69
               Top             =   600
               Width           =   975
            End
            Begin VB.Label Label27 
               BackStyle       =   0  'Transparent
               Caption         =   "POLICE STATION :"
               Height          =   255
               Left            =   3840
               TabIndex        =   68
               Top             =   960
               Width           =   1455
            End
            Begin VB.Label Label28 
               BackStyle       =   0  'Transparent
               Caption         =   "PIN CODE :"
               Height          =   255
               Left            =   3840
               TabIndex        =   67
               Top             =   1320
               Width           =   855
            End
            Begin VB.Label Label29 
               BackStyle       =   0  'Transparent
               Caption         =   "STATE :"
               Height          =   255
               Left            =   3840
               TabIndex        =   66
               Top             =   1680
               Width           =   615
            End
         End
         Begin VB.ComboBox gender 
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
            ItemData        =   "FACULTY DETAILS.frx":2C7A28
            Left            =   5760
            List            =   "FACULTY DETAILS.frx":2C7A32
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   1221
            Width           =   1695
         End
         Begin VB.ComboBox category 
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
            ItemData        =   "FACULTY DETAILS.frx":2C7A44
            Left            =   9600
            List            =   "FACULTY DETAILS.frx":2C7A57
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   1250
            Width           =   1695
         End
         Begin VB.ComboBox religion 
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
            ItemData        =   "FACULTY DETAILS.frx":2C7A7F
            Left            =   13440
            List            =   "FACULTY DETAILS.frx":2C7A8F
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   1251
            Width           =   1695
         End
         Begin VB.ComboBox department 
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
            ItemData        =   "FACULTY DETAILS.frx":2C7AB6
            Left            =   1920
            List            =   "FACULTY DETAILS.frx":2C7AB8
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   1686
            Width           =   1695
         End
         Begin VB.ComboBox course 
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
            ItemData        =   "FACULTY DETAILS.frx":2C7ABA
            Left            =   5760
            List            =   "FACULTY DETAILS.frx":2C7ACA
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   1689
            Width           =   1695
         End
         Begin VB.Label facid 
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
            Left            =   5760
            TabIndex        =   1
            Top             =   240
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
            TabIndex        =   177
            Top             =   240
            Width           =   135
         End
         Begin VB.Label Label92 
            BackStyle       =   0  'Transparent
            Caption         =   "SALARY :"
            Height          =   255
            Left            =   11640
            TabIndex        =   176
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
            TabIndex        =   170
            Top             =   240
            Width           =   135
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "CATEGORY :"
            Height          =   255
            Left            =   7680
            TabIndex        =   169
            Top             =   1260
            Width           =   975
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "SURNAME :"
            Height          =   255
            Left            =   7680
            TabIndex        =   168
            Top             =   810
            Width           =   975
         End
         Begin VB.Label Label38 
            BackStyle       =   0  'Transparent
            Caption         =   "CONTACT NO :"
            Height          =   255
            Left            =   7680
            TabIndex        =   167
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
            TabIndex        =   166
            Top             =   600
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
            TabIndex        =   157
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
            TabIndex        =   156
            Top             =   2040
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
            TabIndex        =   155
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
            TabIndex        =   154
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
            TabIndex        =   153
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
            TabIndex        =   152
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
            TabIndex        =   151
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
            TabIndex        =   150
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
            TabIndex        =   149
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
            Height          =   255
            Left            =   11640
            TabIndex        =   104
            Top             =   873
            Width           =   1335
         End
         Begin VB.Label Label39 
            BackStyle       =   0  'Transparent
            Caption         =   "NATIONALITY :"
            Height          =   255
            Left            =   11640
            TabIndex        =   102
            Top             =   2160
            Width           =   1335
         End
         Begin VB.Label Label37 
            BackStyle       =   0  'Transparent
            Caption         =   "PAN CARD NO :"
            Height          =   255
            Left            =   3960
            TabIndex        =   101
            Top             =   2280
            Width           =   1215
         End
         Begin VB.Label Label36 
            BackStyle       =   0  'Transparent
            Caption         =   "AADHAR NO :"
            Height          =   255
            Left            =   120
            TabIndex        =   100
            Top             =   2160
            Width           =   1095
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "FACULTY ID :"
            Height          =   255
            Left            =   3960
            TabIndex        =   97
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "FIRST NAME :"
            Height          =   255
            Left            =   120
            TabIndex        =   96
            Top             =   810
            Width           =   1095
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "MIDDLE NAME :"
            Height          =   255
            Left            =   3960
            TabIndex        =   95
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "DATE OF BIRTH :"
            Height          =   255
            Left            =   120
            TabIndex        =   94
            Top             =   1260
            Width           =   1335
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "GENDER :"
            Height          =   255
            Left            =   3960
            TabIndex        =   93
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "RELIGION :"
            Height          =   255
            Left            =   11640
            TabIndex        =   92
            Top             =   1301
            Width           =   975
         End
         Begin VB.Label Label30 
            BackStyle       =   0  'Transparent
            Caption         =   "DEPARTMENT :"
            Height          =   255
            Left            =   120
            TabIndex        =   91
            Top             =   1710
            Width           =   1215
         End
         Begin VB.Label Label31 
            BackStyle       =   0  'Transparent
            Caption         =   "COURSE :"
            Height          =   255
            Left            =   3960
            TabIndex        =   90
            Top             =   1800
            Width           =   735
         End
         Begin VB.Label Label33 
            BackStyle       =   0  'Transparent
            Caption         =   "STATUS :"
            Height          =   255
            Left            =   240
            TabIndex        =   89
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label34 
            BackStyle       =   0  'Transparent
            Caption         =   "DATE OF JOINING :"
            Height          =   255
            Left            =   7680
            TabIndex        =   88
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label35 
            BackStyle       =   0  'Transparent
            Caption         =   "EXPERIENCE :"
            Height          =   255
            Left            =   7740
            TabIndex        =   87
            Top             =   1725
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
   Begin VB.Label path 
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   120
      TabIndex        =   178
      Top             =   120
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label58 
      Caption         =   "Label58"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      TabIndex        =   121
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label57 
      Caption         =   "Label57"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      TabIndex        =   120
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label55 
      Caption         =   "Label55"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      TabIndex        =   119
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label54 
      Caption         =   "Label54"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      TabIndex        =   118
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label53 
      Caption         =   "Label53"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      TabIndex        =   117
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label52 
      Caption         =   "Label52"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      TabIndex        =   116
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label51 
      Caption         =   "Label51"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      TabIndex        =   115
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label50 
      Caption         =   "Label50"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      TabIndex        =   114
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label49 
      Caption         =   "Label49"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      TabIndex        =   113
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label48 
      Caption         =   "Label48"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      TabIndex        =   112
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label47 
      Caption         =   "Label47"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      TabIndex        =   111
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label46 
      Caption         =   "Label46"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      TabIndex        =   110
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label45 
      Caption         =   "Label45"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      TabIndex        =   109
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label44 
      Caption         =   "Label44"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      TabIndex        =   108
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label43 
      Caption         =   "Label43"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      TabIndex        =   107
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label42 
      Caption         =   "Label42"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      TabIndex        =   106
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label41 
      Caption         =   "Label41"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      TabIndex        =   105
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
      Height          =   525
      Left            =   7200
      TabIndex        =   98
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "frmFACULTY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub clearscreen()
salary.Text = ""
firstname.Text = ""
surname.Text = ""
middlename.Text = ""
fathername.Text = ""
exp.Text = ""
aadharno.Text = ""
contactno.Text = ""
pano.Text = ""
hno.Text = ""
hno2.Text = ""
village.Text = ""
village2.Text = ""
city.Text = ""
city2.Text = ""
street.Text = ""
street2.Text = ""
landmarkk.Text = ""
landmark.Text = ""
pincode.Text = ""
policestation.Text = ""
postoffice.Text = ""
pincode2.Text = ""
policestation2.Text = ""
postoffice2.Text = ""
List1.clear
List2.clear
List3.clear
List4.clear
List5.clear
List6.clear
List7.clear
LIST8.clear
state.Text = ""
state2.Text = ""
district.Text = ""
district2.Text = ""
End Sub
Private Sub aadharno_LostFocus()
If (Len(aadharno) < 12) Then
A = MsgBox("AADHAR NUMBER MUST BE OF 12 DIGIT", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
aadharno.SetFocus
End If
End If
End Sub
Private Sub Check1_Click()
If Check1.Value = 1 Then
    hno2.Text = hno.Text
    street2.Text = street.Text
    village2.Text = village.Text
    landmark.Text = landmarkk.Text
    city2.Text = city.Text
    policestation2.Text = policestation.Text
    postoffice2.Text = postoffice.Text
    pincode2.Text = pincode.Text
    district2.Text = district.Text
    state2.Text = state.Text
End If
End Sub

Private Sub Command6_Click()
cd2.Filter = "picture file |*.jpg"
cd2.ShowOpen
If cd2.FileName <> " " Then
    Image1.Picture = LoadPicture(cd2.FileName)
End If
path.Caption = cd2.FileName
End Sub
Private Sub add1_Click()
If (quali.Text = "" Or college.Text = "" Or yearofpass.Text = "" Or percentage.Text = "") Then
MsgBox "please enter valid data"
Else
List1.AddItem quali.Text
quali.Text = ""
List2.AddItem college.Text
college.Text = ""
List3.AddItem yearofpass.Text
yearofpass.Text = ""
List4.AddItem percentage.Text
percentage.Text = ""
End If
End Sub

Private Sub add2_Click()
If (organisation.Text = "" Or designation.Text = "" Or year.Text = "" Or dateofwork.Text = "") Then
MsgBox "please enter valid data"
Else
List5.AddItem organisation.Text
organisation.Text = ""
List6.AddItem designation.Text
designation.Text = ""
List7.AddItem year.Text
year.Text = ""
LIST8.AddItem dateofwork.Text
dateofwork.Text = ""
End If
End Sub

Private Sub Command7_Click()
Image1.Picture = unloadpicture(path.Caption)
End Sub

Private Sub contactno_LostFocus()
If (Len(contactno) < 10) Then
A = MsgBox("CONTACT NUMBER MUST BE OF 10 DIGIT", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
contactno.SetFocus
End If
End If
End Sub

'Private Sub Command6_Click()
'cd2.Filter = "picture file |*.jpg"
'cd2.ShowOpen
'If cd2.FileName <> " " Then
'Image1.Picture = LoadPicture(cd2.FileName)
'End If
'End Sub

Private Sub delete_Click()
If (facid.Caption = "" Or salary.Text = "" Or firstname.Text = "" Or surname.Text = "" Or middlename.Text = "" Or fathername.Text = "" Or exp.Text = "" Or nationality.Text = "" Or pano.Text = "" Or aadharno.Text = "" Or contactno.Text = "" Or hno.Text = "" Or street.Text = "" Or village.Text = "" Or city.Text = "" Or landmarkk.Text = "" Or district.Text = "" Or state.Text = "" Or pincode.Text = "" Or policestation.Text = "" Or postoffice.Text = "" Or hno2.Text = "" Or city2.Text = "" Or village2.Text = "" Or street2.Text = "" Or landmark.Text = "" Or district2.Text = "" Or postoffice2.Text = "" Or policestation2.Text = "" Or district2.Text = "" Or state2.Text = "" Or pincode2.Text = "") Then
frmINFO.SHOW 1
If (str = "OK") Then
status.SetFocus
End If
Else
Module1.conn
frmDELETE.SHOW 1
If (str = "YES") Then
sql = "delete from FacultyExp_Master where fac_id='" + facid.Caption + "'"
Set r = c.Execute(sql)
sql = "delete from FacultyQuali_Master where fac_id='" + facid.Caption + "'"
Set r = c.Execute(sql)
sql = "delete from fac_salinfo_master where facid='" + facid.Caption + "'"
Set r = c.Execute(sql)
sql = "delete from FacultyDetail_Master where fac_id='" + facid.Caption + "'"
Set r = c.Execute(sql)
b = MsgBox("RECORD DELETED", vbInformation + vbOKOnly, "INFORMATION")
If (b = vbOK) Then
Set Image1.Picture = Nothing
Image1.Picture = LoadPicture()
Call clearscreen
status.SetFocus
End If
ElseIf (str = "NO") Then
status.SetFocus
End If
End If
End Sub

Private Sub dob_LostFocus()
Dim age As Long
age = DateDiff("yyyy", dob.Value, Date)
If (age <= 25) Then
MsgBox ("WE NOT ALLOW A FACULTY UNDER AGE 25")
dob.SetFocus
If (age >= 50) Then
MsgBox ("WE NOT ALLOW A FACULTY OVER AGE 50")
dob.SetFocus
End If
End If
End Sub

Private Sub doj_keypress(KeyAscii As Integer)
If (KeyAscii = 13) Then
salary.SetFocus
End If
End Sub

Private Sub facid_keypress(KeyAscii As Integer)
If (KeyAscii = 13) Then
doj.SetFocus
End If
If (KeyAscii >= 65 And KeyAscii <= 87 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32 Or KeyAscii >= 48 And KeyAscii <= 57) Then
facid.Locked = False
Else
facid.Locked = True
End If
End Sub
'Private Sub exp_LostFocus()
'If (Option3.Value = True) Then
'exp.Text = (exp.Text) & (" ") & "month"
'Else
'If (Option4.Value = True) Then
'exp.Text = (exp.Text) & (" ") & "year"
'End If
'End If
'End Sub

Private Sub doj_LostFocus()
Dim d As Date
d = doj.Value
If (d > Date) Then
A = MsgBox("INVALID DATE", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOKOnly) Then
doj.SetFocus
End If
End If
End Sub


Private Sub firstname_keypress(KeyAscii As Integer)
If (KeyAscii = 13) Then
middlename.SetFocus
End If
If (kekyascii >= 65 And KeyAscii <= 87 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32) Then
firstname.Locked = False
Else
firstname.Locked = True
End If
End Sub

Private Sub Form_ACTIVATE()
status.SetFocus
exp.MaxLength = 7
delete.Enabled = False
update.Enabled = False
year.MaxLength = 4
End Sub

Private Sub Form_Load()
doj1.Visible = False
doj.Visible = True
dob1.Visible = False
'fac.Visible = False
'facid.Visible = True
Module1.conn
sql = "select count(fac_id) from FacultyDetail_Master"
Set r = c.Execute(sql)
'Dim j As Integer
'If IsNull(r.Fields(0)) Then
'studadm.Text = 1
'Else
facid.Caption = "FAC" & (r.Fields(0) + 1)
'End If
j = r.Fields(0)
If (j < 1) Then
'MsgBox "record not found"
Else
sql = "select fac_id from FacultyDetail_Master order by rownum desc"
Set r = c.Execute(sql)
Dim i As String
i = r.Fields(0)
A = Right(i, 1)
k = CInt(A)
facid.Caption = "FAC" & (k + 1)
End If
'facid.MaxLength = 5
exp.MaxLength = 7
salary.MaxLength = 10
firstname.MaxLength = 20
middlename.MaxLength = 20
surname.MaxLength = 20
fathername.MaxLength = 30
exp.MaxLength = 20
aadharno.MaxLength = 12
pano.MaxLength = 10
hno.MaxLength = 5
street.MaxLength = 20
landmarkk.MaxLength = 20
city.MaxLength = 20
contactno.MaxLength = 10
policestation.MaxLength = 20
postoffice.MaxLength = 20
pincode.MaxLength = 6
district.MaxLength = 15
state.MaxLength = 15
hno2.MaxLength = 5
street2.MaxLength = 20
village2.MaxLength = 20
landmark.MaxLength = 20
city2.MaxLength = 20
policestation2.MaxLength = 20
postoffice2.MaxLength = 20
pincode2.MaxLength = 6
district2.MaxLength = 15
state2.MaxLength = 15
quali.MaxLength = 20
college.MaxLength = 20
yearofpass.MaxLength = 4
percentage.MaxLength = 3
organisation.MaxLength = 20
designation.MaxLength = 20
year.MaxLength = 9
dateofwork.MaxLength = 15
Module1.conn
sql = "select code from CourseDetail_Master"
Set r = c.Execute(sql)
While (r.EOF = False)
course.AddItem r.Fields("code")
r.MoveNext
Wend
'sql = "select sub_id from SubjectDetail_Master"
'Set r = c.Execute(sql)
'While (r.EOF = False)
'subject.AddItem r.Fields("sub_id")
'r.MoveNext
'Wend
sql = "select dcode from DepartmentDetail_Master"
Set r = c.Execute(sql)
While (r.EOF = False)
department.AddItem r.Fields("dcode")
r.MoveNext
Wend
End Sub

Private Sub Label32_Click()

End Sub

Private Sub List1_Click()
quali.Text = List1.List(List1.ListIndex)
End Sub

Private Sub List2_Click()
college.Text = List2.List(List2.ListIndex)
End Sub

Private Sub List3_Click()
yearofpass.Text = List3.List(List3.ListIndex)
End Sub
Private Sub List4_Click()
percentage.Text = List4.List(List4.ListIndex)
End Sub
Private Sub List5_Click()
organisation.Text = List5.List(List5.ListIndex)
End Sub
Private Sub List6_Click()
designation.Text = List6.List(List6.ListIndex)
End Sub
Private Sub List7_Click()
year.Text = List7.List(List7.ListIndex)
End Sub
Private Sub List8_Click()
dateofwork.Text = LIST8.List(LIST8.ListIndex)
End Sub
Private Sub middlename_keypress(KeyAscii As Integer)
If (KeyAscii = 13) Then
surname.SetFocus
If (KeyAscii >= 65 And KeyAscii <= 87 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32) Then
middlename.Locked = False
Else
middlename.Locked = True
End If
End If
End Sub
Private Sub Option1_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
Option2.SetFocus
End If
End Sub
Private Sub Option2_keypress(KeyAscii As Integer)
If (KeyAscii = 13) Then
facid.SetFocus
End If
End Sub

Private Sub Option3_gotfocus()
If (exp.Text = "") Then
exp.SetFocus
Else
A = Left(exp.Text, 2)
exp.Text = ""
exp.Text = A & " " & "month"
LTrim (exp.Text)
End If
End Sub

Private Sub Option4_gotfocus()
If exp.Text = "" Then
exp.SetFocus
Else
 A = Left(exp.Text, 2)
 exp.Text = ""
 exp.Text = A & "year"
End If
End Sub

Private Sub organisation_click()
If (exp.Text = " ") Then
A = MsgBox("YOUR EXPERIANCE IS NULL", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
exp.SetFocus
End If
End If
End Sub



Private Sub percentage_keypress(KeyAscii As Integer)
If (KeyAscii = 13) Then
organisation.SetFocus
End If
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
percentage.Locked = False
Else
percentage.Locked = True
End If
End Sub

Private Sub refresh_Click()
quali.Text = ""
college.Text = ""
yearofpass.Text = ""
percentage.Text = ""
organisation.Text = ""
designation.Text = ""
dateofwork.Text = ""
year.Text = ""
save.Enabled = True
delete.Enabled = False
update.Enabled = False
Call clearscreen
Module1.conn
sql = "select fac_id from FacultyDetail_Master order by rownum desc"
Set r = c.Execute(sql)
Dim i As String
If r.EOF = True Then
status.SetFocus
Else
i = r.Fields(0)
A = Right(i, 1)
k = CInt(A)
facid.Caption = "FAC" & (k + 1)
'fac.Visible = False
dob1.Visible = False
doj1.Visible = False
doj.Visible = True
dob.Visible = True
facid.Visible = True
Set Image1.Picture = Nothing
Image1.Picture = LoadPicture()
End If
End Sub

Private Sub remove1_Click()
Item = List1.ListIndex
If (Item >= 0) Then
List1.RemoveItem Item
End If
Item = List2.ListIndex
If (Item >= 0) Then
List2.RemoveItem Item
End If
Item = List3.ListIndex
If (Item >= 0) Then
List3.RemoveItem Item
End If
Item = List4.ListIndex
If (Item >= 0) Then
List4.RemoveItem Item
End If
End Sub

Private Sub remove2_Click()
Item = List5.ListIndex
If (Item >= 0) Then
List5.RemoveItem Item
End If
Item = List6.ListIndex
If (Item >= 0) Then
List6.RemoveItem Item
End If
Item = List7.ListIndex
If (Item >= 0) Then
List7.RemoveItem Item
End If
Item = LIST8.ListIndex
If (Item >= 0) Then
LIST8.RemoveItem Item
End If
End Sub

Private Sub salary_keypress(KeyAscii As Integer)
If (KeyAscii = 13) Then
firstname.SetFocus
End If
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Or KeyAscii = 39 Or KeyAscii = 46) Then
salary.Locked = False
Else
salary.Locked = True
End If
End Sub
Private Sub save_Click()
'On Error GoTo label
If (facid.Caption = "" Or salary.Text = "" Or gender.Text = "" Or religion.Text = "" Or category.Text = "" Or department.Text = "" Or course.Text = "" Or firstname.Text = "" Or nationality.Text = "" Or contactno.Text = "" Or district.Text = "" Or state.Text = "" Or pincode.Text = "" Or policestation.Text = "" Or postoffice.Text = "" Or district2.Text = "" Or postoffice2.Text = "" Or policestation2.Text = "" Or district2.Text = "" Or state2.Text = "" Or pincode2.Text = "" Or List1.ListCount = 0 Or List2.ListCount = 0 Or List3.ListCount = 0 Or List4.ListCount = 0 Or List5.ListCount = 0 Or List6.ListCount = 0 Or LIST8.ListCount = 0 Or LIST8.ListCount = 0 Or status.Text = "" Or path.Caption = "") Then
frmINFO.SHOW 1
If str = "OK" Then
firstname.SetFocus
End If
Else
Module1.conn
frmSAVE.SHOW 1
If str = "YES" Then
sql = "insert into FacultyDetail_Master values('" + status.Text + "','" + facid.Caption + "','" + Format(doj.Value, "dd/MMM/yyyy") + "','" + firstname.Text + "','" + middlename.Text + "','" + surname.Text + "','" + fathername.Text + "','" + Format(dob.Value, "dd/MMM/yyyy") + "','" + gender.Text + "','" + category.Text + "','" + religion.Text + "','" + exp.Text + "','" + aadharno.Text + "','" + pano.Text + "'," + contactno.Text + ",'" + nationality.Text + "','" + hno.Text + "','" + street.Text + "','" + village.Text + "','" + landmarkk.Text + "','" + city.Text + "','" + policestation.Text + "','" + postoffice.Text + "'," + pincode.Text + ",'" + district.Text + "','" + state.Text + "','" + hno2.Text + "','" + street2.Text + "','" + village2.Text + "','" + landmark.Text + "','" + city.Text + "','" + policestation2.Text + "','" + postoffice2.Text + "'," + pincode2.Text + ",'" + district2.Text + "','" + state2.Text + "','" + department.Text + "','" + salary.Text + "','" + course.Text + "','" + path.Caption + "')"
Set r = c.Execute(sql)
i = 0
Counter = 0
While (i <= List1.ListCount - 1 And Counter <= Counter)
sql = "insert into facultyQuali_Master values('" & List1.List(i) & "','" & List2.List(i) & "'," & List3.List(i) & "," + List4.List(i) + ",'" & facid.Caption & "'," & Counter & ")"
Set r = c.Execute(sql)
MsgBox sql
i = i + 1
Counter = Counter + 1
Wend
i = 0
Counter = 0
While (i <= List5.ListCount - 1 And Counter <= Counter)
sql = "insert into facultyexp_Master values('" + List5.List(i) + "','" & List6.List(i) & "'," & List7.List(i) & ",'" + LIST8.List(i) + "','" & facid.Caption & "'," & Counter & ")"
Set r = c.Execute(sql)
i = i + 1
Counter = Counter + 1
Wend
A = MsgBox("RECORD SAVE", vbQuestion + vbOKCancel, "SAVE")
If (A = vbOK) Then
A = MsgBox("RECORD SAVE", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
Set Image1.Picture = Nothing
Image1.Picture = LoadPicture()
Call clearscreen
End If
ElseIf (str = "NO") Then
firstname.SetFocus
End If
End If
'Exit Sub
'label:
'MsgBox "data alredy exist"
'Call clearscreen
A = facid.Caption
k = Right(A, 1)
j = k + 1
facid.Caption = "FAC" & j
End If
End Sub

Private Sub surname_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
fathername.SetFocus
End If
If (KeyAscii >= 65 And KeyAscii <= 87 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32) Then
surname.Locked = False
Else
surname.Locked = True
End If
End Sub
Private Sub fathername_keypress(KeyAscii As Integer)
If (KeyAscii = 13) Then
dob.SetFocus
End If
If (KeyAscii >= 65 And KeyAscii <= 87 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32) Then
fathername.Locked = False
Else
fathername.Locked = True
End If
End Sub
Private Sub dob_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
gender.SetFocus
End If
End Sub
Private Sub gender_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
category.SetFocus
End If
End Sub
Private Sub category_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
religion.SetFocus
End If
End Sub
Private Sub religion_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
department.SetFocus
End If
End Sub
Private Sub department_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
course.SetFocus
End If
End Sub
Private Sub course_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
subject.SetFocus
End If
End Sub
Private Sub subject_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
exp.SetFocus
End If
End Sub
Private Sub exp_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
aadharno.SetFocus
End If
If (KeyAscii = 8 Or KeyAscii >= 48 And KeyAscii <= 57) Then
exp.Locked = False
Else
exp.Locked = True
End If
End Sub
Private Sub aadharno_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
pano.SetFocus
End If
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
aadharno.Locked = False
Else
aadharno.Locked = True
End If
End Sub
Private Sub pano_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
contactno.SetFocus
End If
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Or KeyAscii >= 65 And KeyAscii <= 87 Or KeyAscii >= 97 And KeyAscii <= 122) Then
pano.Locked = False
Else
pano.Locked = True
End If
End Sub
Private Sub contactno_keypress(KeyAscii As Integer)
If (KeyAscii = 13) Then
nationality.SetFocus
End If
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
contactno.Locked = False
Else
contactno.Locked = True
End If
End Sub
Private Sub nationality_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
hno.SetFocus
End If
End Sub
Private Sub hno_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
street.SetFocus
End If
End Sub
Private Sub street_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
village.SetFocus
End If
End Sub

Private Sub update_Click()
If (facid.Caption = "" Or salary.Text = "" Or firstname.Text = "" Or exp.Text = "" Or nationality.Text = "" Or contactno.Text = "" Or district.Text = "" Or state.Text = "" Or pincode.Text = "" Or policestation.Text = "" Or postoffice.Text = "" Or district2.Text = "" Or postoffice2.Text = "" Or policestation2.Text = "" Or district2.Text = "" Or state2.Text = "" Or pincode2.Text = "") Then
frmINFO.SHOW 1
If (str = "OK") Then
status.SetFocus
End If
Else
Module1.conn
frmUPDATE.SHOW 1
If str = "YES" Then
sql = "update FacultyDetail_Master set  fac_fnm='" + firstname.Text + "', fac_mnm='" + middlename.Text + "', fac_snm='" + surname.Text + "', fac_fathnm='" + fathername.Text + "', fac_dob='" + Format(dob1.Text, "dd-MMM-yyyy") + "', fac_gen='" + gender.Text + "', fac_cat='" + category.Text + "', fac_reli='" + religion.Text + "', course='" + course.Text + "', fac_exp='" + exp.Text + "', fac_aadhar='" + aadharno.Text + "', fac_pan='" + pano.Text + "', fac_con=" + contactno.Text + ", fac_nation='" + nationality.Text + "', fper_hno='" + hno.Text + "', fper_street='" + street.Text + "', fper_vill='" + village.Text + "', fper_landmark='" + landmarkk.Text + "', fper_city='" + city.Text + "', fper_ps='" + policestation.Text + "', fper_po='" + postoffice.Text + "', fper_pc='" + pincode.Text + "', fper_dist='" + district.Text + "', fper_state='" + state.Text + "', floc_hno='" + hno2.Text + "', floc_street='" + street2.Text + "', floc_vill='" + village2.Text + "'," _
& "floc_landmark='" + landmark.Text + "', floc_city='" + city2.Text + "', floc_ps='" + policestation2.Text + "', floc_po='" + postoffice2.Text + "', floc_pc=" + pincode2.Text + ", floc_dist='" + district2.Text + "', floc_state='" + state2.Text + "', dcode='" + department.Text + "',salary=" + salary.Text + ",path='" + path.Caption + "' where fac_id='" + facid.Caption + "'"
MsgBox sql
Set r = c.Execute(sql)
'MsgBox sql
'Set r = c.Execute(sql)
i = List1.ListIndex
k = List2.ListIndex
m = List3.ListIndex
n = List4.ListIndex
If List1.ListIndex = -1 Then
quali.Text = ""
Else
List1.RemoveItem List1.ListIndex
List1.AddItem quali.Text
j = i
While (j <= List1.ListCount - 1)
sql = "update facultyQuali_Master set fac_quali='" & List1.List(j) & "' where sno=" & i & ""
Set r = c.Execute(sql)
j = j + 1
Wend
End If
If List2.ListIndex = -1 Then
college.Text = ""
Else
List2.RemoveItem List2.ListIndex
List2.AddItem college.Text
l = k
While (l <= List2.ListCount - 1)
sql = "update facultyQuali_Master set fac_uni=" & List2.List(k) & " where sno=" & k & ""
Set r = c.Execute(sql)
l = l + 1
Wend
End If
If List3.ListIndex = -1 Then
yearofpass.Text = ""
Else
List3.RemoveItem List3.ListIndex
List3.AddItem yearofpass.Text
P = m
While (P <= List3.ListCount - 1)
sql = "update facultyQuali_Master set fac_pyear=" & List3.List(P) & " where sno=" & m & ""
Set r = c.Execute(sql)
P = P + 1
Wend
End If
If List4.ListIndex = -1 Then
percentage.Text = ""
Else
List4.RemoveItem List4.ListIndex
List4.AddItem percentage.Text
t = n
While (t <= List4.ListCount - 1)
sql = "update facultyQuali_Master set fac_per='" & List4.List(t) & "' where sno=" & n & ""
Set r = c.Execute(sql)
t = t + 1
Wend
End If
'List1.AddItem quali.Text
'List2.AddItem college.Text
'List3.AddItem yearofpass.Text
'List4.AddItem percentage.Text
i = List5.ListIndex
k = List6.ListIndex
m = List7.ListIndex
n = LIST8.ListIndex
If List5.ListIndex = -1 Then
organisation.Text = ""
Else
List5.RemoveItem List5.ListIndex
List5.AddItem organisation.Text
j = i
While (j <= List5.ListCount - 1)
sql = "update facultyexp_Master set fac_org='" & List5.List(j) & "' where sno=" & i & ""
Set r = c.Execute(sql)
j = j + 1
Wend
End If
If List6.ListIndex = -1 Then
designation.Text = ""
Else
List6.RemoveItem List6.ListIndex
List6.AddItem designation.Text
l = k
While (l <= List6.ListCount - 1)
sql = "update facultyexp_Master set fac_pro=" & List6.List(k) & " where sno=" & k & ""
Set r = c.Execute(sql)
l = l + 1
Wend
End If
If List7.ListIndex = -1 Then
year.Text = ""
Else
List7.RemoveItem List7.ListIndex
List7.AddItem year.Text
P = m
While (P <= List7.ListCount - 1)
sql = "update facultyexp_Master set fac_year=" & List7.List(P) & " where sno=" & m & ""
Set r = c.Execute(sql)
P = P + 1
Wend
End If
If LIST8.ListIndex = -1 Then
dateofwork.Text = ""
Else
LIST8.RemoveItem LIST8.ListIndex
LIST8.AddItem dateofwork.Text
t = n
While (t <= LIST8.ListCount - 1)
sql = "update facultyexp_Master set fac_dur='" & LIST8.List(t) & "' where sno=" & n & ""
Set r = c.Execute(sql)
t = t + 1
Wend
End If
'List5.AddItem organisation.Text
'List6.AddItem designation.Text
'List7.AddItem year.Text
'LIST8.AddItem dateofwork.Text
A = MsgBox("RECORD UPDATE", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
Call clearscreen
Else
If (A = vbNo) Then
status.SetFocus
End If
End If
End If
End If
End Sub

Private Sub view_Click()
save.Enabled = False
delete.Enabled = True
update.Enabled = True
dob.Visible = False
dob1.Visible = True
doj1.Visible = True
'facid.Visible = False
doj.Visible = False
'fac.Visible = True
Module1.conn
Value = UCase(InputBox("ENTER FACULTY ID", "FACULTY ID"))
facid.Caption = Value
sql = " select *from FacultyDetail_Master where fac_id ='" + facid.Caption + "'"
Set r = c.Execute(sql)
If (r.EOF = True) Then
MsgBox "DATA NOT FOUND"
Else
status.Text = r.Fields(0)
doj1.Text = r.Fields(2)
firstname.Text = r.Fields(3)
If (IsNull(r.Fields(4))) Then
middlename.Text = ""
Else
middlename.Text = r.Fields(4)
End If
If (IsNull(r.Fields(5))) Then
surname.Text = ""
Else
surname.Text = r.Fields(5)
End If
If (IsNull(r.Fields(6))) Then
fathername.Text = ""
Else
fathername.Text = r.Fields(6)
End If
dob1.Text = r.Fields(7)
gender.Text = r.Fields(8)
category.Text = r.Fields(9)
religion.Text = r.Fields(10)
If (IsNull(r.Fields(11))) Then
exp.Text = ""
Else
exp.Text = r.Fields(11)
End If
If (IsNull(r.Fields(12))) Then
aadharno.Text = ""
Else
aadharno.Text = r.Fields(12)
End If
If (IsNull(r.Fields(13))) Then
pano.Text = ""
Else
pano.Text = r.Fields(13)
End If
contactno.Text = r.Fields(14)
nationality.Text = r.Fields(15)
If (IsNull(r.Fields(16))) Then
hno.Text = ""
Else
hno.Text = r.Fields(16)
End If
If (IsNull(r.Fields(17))) Then
street.Text = ""
Else
street.Text = r.Fields(17)
End If
If (IsNull(r.Fields(18))) Then
village.Text = ""
Else
village.Text = r.Fields(18)
End If
If (IsNull(r.Fields(19))) Then
landmarkk.Text = ""
Else
landmarkk.Text = r.Fields(19)
End If
If (IsNull(r.Fields(20))) Then
city.Text = ""
Else
city.Text = r.Fields(20)
End If
policestation.Text = r.Fields(21)
postoffice.Text = r.Fields(22)
pincode.Text = r.Fields(23)
district.Text = r.Fields(24)
state.Text = r.Fields(25)
If (IsNull(r.Fields(26))) Then
hno2.Text = ""
Else
hno2.Text = r.Fields(26)
End If
If (IsNull(r.Fields(27))) Then
street2.Text = ""
Else
street2.Text = r.Fields(27)
End If
If (IsNull(r.Fields(28))) Then
village2.Text = ""
Else
village2.Text = r.Fields(28)
End If
If (IsNull(r.Fields(29))) Then
landmark.Text = ""
Else
landmark.Text = r.Fields(29)
End If
If (IsNull(r.Fields(30))) Then
city2.Text = ""
Else
city2.Text = r.Fields(30)
End If
policestation2.Text = r.Fields(31)
postoffice2.Text = r.Fields(32)
pincode2.Text = r.Fields(33)
district2.Text = r.Fields(34)
state2.Text = r.Fields(35)
department.Text = r.Fields(36)
salary.Text = r.Fields(37)
course.Text = r.Fields(38)
path.Caption = r.Fields(39)
Image1.Picture = LoadPicture(path.Caption)
sql = "select *from FacultyQuali_Master where fac_id='" + UCase(Value) + "'"
Set r = c.Execute(sql)
While (r.EOF = False)
List1.AddItem r.Fields("fac_quali")
List2.AddItem r.Fields("fac_uni")
List3.AddItem r.Fields("fac_pyear")
List4.AddItem r.Fields("fac_per")
r.MoveNext
Wend
sql = "select *from FacultyExp_Master where fac_id='" + UCase(Value) + "'"
Set r = c.Execute(sql)
While (r.EOF = False)
List5.AddItem r.Fields("fac_org")
List6.AddItem r.Fields("fac_pro")
List7.AddItem r.Fields("fac_year")
LIST8.AddItem r.Fields("fac_dur")
r.MoveNext
Wend
End If
End Sub

Private Sub village_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
landmarkk.SetFocus
End If
If (KeyAscii = 8 Or KeyAscii >= 65 And KeyAscii <= 87 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 32) Then
village.Locked = False
Else
village.Locked = True
End If
End Sub
Private Sub landmarkk_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
city.SetFocus
End If
If (KeyAscii = 8 Or KeyAscii >= 65 And KeyAscii <= 87 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 32) Then
landmarkk.Locked = False
Else
landmarkk.Locked = True
End If
End Sub
Private Sub city_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
policestation.SetFocus
End If
If (KeyAscii = 8 Or KeyAscii >= 65 And KeyAscii <= 87 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 32) Then
city.Locked = False
Else
city.Locked = True
End If
End Sub
Private Sub policestation_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
postoffice.SetFocus
End If
If (KeyAscii = 8 Or KeyAscii >= 65 And KeyAscii <= 87 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 32) Then
policestation.Locked = False
Else
village.Locked = True
End If
End Sub
Private Sub postoffice_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
pincode.SetFocus
End If
If (KeyAscii >= 65 And KeyAscii <= 87 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32) Then
postoffice.Locked = False
Else
postoffice.Locked = True
End If
End Sub
Private Sub pincode_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
district.SetFocus
End If
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
pincode.Locked = False
Else
pincode.Locked = True
End If
End Sub
Private Sub district_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
state.SetFocus
End If
If (KeyAscii >= 65 And KeyAscii <= 87 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32) Then
district.Locked = False
Else
district.Locked = True
End If
End Sub
Private Sub state_keypress(KeyAscii As Integer)
If (KeyAscii = 13) Then
hno2.SetFocus
End If
If (KeyAscii >= 65 And KeyAscii <= 87 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32) Then
state.Locked = False
Else
state.Locked = True
End If
End Sub
Private Sub hno2_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
street2.SetFocus
End If
End Sub
Private Sub street2_keypress(KeyAscii As Integer)
If (KeyAscii = 13) Then
village2.SetFocus
End If
End Sub
Private Sub village2_keypress(KeyAscii As Integer)
If (KeyAscii = 13) Then
landmark.SetFocus
End If
If (KeyAscii >= 65 And KeyAscii <= 87 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32) Then
village2.Locked = False
Else
village2.Locked = True
End If
End Sub
Private Sub landmark_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
city2.SetFocus
End If
If (KeyAscii >= 65 And KeyAscii <= 87 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32) Then
landmark.Locked = False
Else
landmark.Locked = True
End If
End Sub
Private Sub city2_keypress(KeyAscii As Integer)
If (KeyAscii = 13) Then
policestation2.SetFocus
End If
If (KeyAscii >= 65 And KeyAscii <= 87 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32) Then
city2.Locked = False
Else
city2.Locked = True
End If
End Sub
Private Sub policestation2_keypress(KeyAscii As Integer)
If (KeyAscii = 13) Then
postoffice2.SetFocus
End If
If (KeyAscii >= 65 And KeyAscii <= 87 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32) Then
policestation2.Locked = False
Else
policestation2.Locked = True
End If
End Sub
Private Sub postoffice2_keypress(KeyAscii As Integer)
If (KeyAscii = 13) Then
pincode2.SetFocus
End If
If (KeyAscii >= 65 And KeyAscii <= 87 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32) Then
postoffice2.Locked = False
Else
postoffice2.Locked = True
End If
End Sub
Private Sub pincode2_keypress(KeyAscii As Integer)
If (KeyAscii = 13) Then
district2.SetFocus
End If
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
pincode2.Locked = False
Else
pincode2.Locked = True
End If
End Sub
Private Sub district2_keypress(KeyAscii As Integer)
If (KeyAscii = 13) Then
state2.SetFocus
End If
If (KeyAscii >= 65 And KeyAscii <= 87 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32) Then
district2.Locked = False
Else
district2.Locked = True
End If
End Sub
Private Sub state2_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
quali.SetFocus
End If
If (KeyAscii >= 65 And KeyAscii <= 87 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32) Then
state2.Locked = False
Else
state2.Locked = True
End If
End Sub
Private Sub quali_keypress(KeyAscii As Integer)
'If (keyascii = 13) Then
'End If
'college.SetFocus
If (KeyAscii >= 65 And KeyAscii <= 87 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32) Then
quali.Locked = False
Else
quali.Locked = True
End If
End Sub

Private Sub college_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
yearofpass.SetFocus
End If
End Sub
Private Sub year_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
dateofwork.SetFocus
End If
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
year.Locked = False
Else
year.Locked = True
End If
End Sub

Private Sub yearofpass_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
percentage.SetFocus
End If
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
yearofpass.Locked = False
Else
yearofpass.Locked = True
End If
End Sub
