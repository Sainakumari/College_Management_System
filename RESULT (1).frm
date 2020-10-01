VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRESULT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RESULT"
   ClientHeight    =   9930
   ClientLeft      =   375
   ClientTop       =   705
   ClientWidth     =   19755
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   Picture         =   "RESULT.frx":0000
   ScaleHeight     =   9930
   ScaleWidth      =   19755
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Height          =   8895
      Left            =   720
      TabIndex        =   1
      Top             =   600
      Width           =   18375
      Begin VB.CommandButton Command7 
         Caption         =   "SHOW DATA"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   16170
         Picture         =   "RESULT.frx":1E9608
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   7830
         Width           =   1095
      End
      Begin VB.CommandButton refresh 
         Caption         =   "REFRESH"
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
         Left            =   16080
         Picture         =   "RESULT.frx":1E9E57
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   5310
         Width           =   1095
      End
      Begin VB.CommandButton print 
         BackColor       =   &H80000016&
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
         Left            =   16080
         Picture         =   "RESULT.frx":1EA4FB
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   6480
         Width           =   1095
      End
      Begin VB.CommandButton update 
         BackColor       =   &H80000016&
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
         Left            =   16080
         Picture         =   "RESULT.frx":1EAD7E
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   4110
         Width           =   1095
      End
      Begin VB.CommandButton delete 
         BackColor       =   &H80000016&
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
         Left            =   16080
         Picture         =   "RESULT.frx":1EB50E
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   2910
         Width           =   1095
      End
      Begin VB.CommandButton save 
         BackColor       =   &H80000016&
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
         Left            =   16080
         Picture         =   "RESULT.frx":1EBA7A
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   510
         Width           =   1095
      End
      Begin VB.CommandButton view 
         BackColor       =   &H80000016&
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
         Left            =   16080
         Picture         =   "RESULT.frx":1EC254
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   1710
         Width           =   1095
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FF8080&
         Height          =   8655
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   15255
         Begin VB.TextBox dt1 
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2550
            TabIndex        =   69
            Top             =   8160
            Width           =   2055
         End
         Begin VB.TextBox year 
            BackColor       =   &H80000016&
            Height          =   315
            Left            =   2910
            TabIndex        =   68
            Top             =   1740
            Width           =   2085
         End
         Begin VB.ComboBox ADMNO 
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
            Left            =   11910
            Style           =   2  'Dropdown List
            TabIndex        =   59
            Top             =   300
            Width           =   2055
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
            Left            =   7200
            Style           =   2  'Dropdown List
            TabIndex        =   53
            Top             =   300
            Width           =   2055
         End
         Begin MSComCtl2.DTPicker dt 
            Height          =   375
            Left            =   2520
            TabIndex        =   46
            Top             =   8160
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   661
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
            CalendarBackColor=   16744576
            CustomFormat    =   "dd/MMM/yyyy"
            Format          =   109314051
            CurrentDate     =   43087
         End
         Begin VB.TextBox clgcode 
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
            Height          =   345
            Left            =   2880
            TabIndex        =   12
            Top             =   1080
            Width           =   2055
         End
         Begin VB.TextBox resultno 
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
            Height          =   345
            Left            =   2940
            TabIndex        =   11
            Top             =   240
            Width           =   2055
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00FF8080&
            Height          =   5535
            Left            =   120
            TabIndex        =   10
            Top             =   2520
            Width           =   14055
            Begin VB.TextBox sub5 
               Alignment       =   2  'Center
               BackColor       =   &H00FF8080&
               Height          =   495
               Left            =   630
               TabIndex        =   66
               Top             =   4200
               Width           =   1215
            End
            Begin VB.TextBox sub4 
               Alignment       =   2  'Center
               BackColor       =   &H00FF8080&
               Height          =   495
               Left            =   660
               TabIndex        =   65
               Top             =   3510
               Width           =   1215
            End
            Begin VB.TextBox sub3 
               Alignment       =   2  'Center
               BackColor       =   &H00FF8080&
               Height          =   495
               Left            =   600
               TabIndex        =   64
               Top             =   2790
               Width           =   1215
            End
            Begin VB.TextBox sub2 
               Alignment       =   2  'Center
               BackColor       =   &H00FF8080&
               Height          =   495
               Left            =   600
               TabIndex        =   63
               Top             =   2040
               Width           =   1215
            End
            Begin VB.TextBox sub1 
               Alignment       =   2  'Center
               BackColor       =   &H00FF8080&
               Height          =   495
               Left            =   570
               TabIndex        =   62
               Top             =   1260
               Width           =   1215
            End
            Begin VB.TextBox total 
               Alignment       =   2  'Center
               BackColor       =   &H00FF8080&
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   8040
               TabIndex        =   44
               Top             =   4920
               Width           =   1455
            End
            Begin VB.TextBox mo5 
               Alignment       =   2  'Center
               BackColor       =   &H00FF8080&
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   8040
               TabIndex        =   43
               Top             =   4200
               Width           =   1455
            End
            Begin VB.TextBox mo4 
               Alignment       =   2  'Center
               BackColor       =   &H00FF8080&
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   8040
               TabIndex        =   42
               Top             =   3480
               Width           =   1455
            End
            Begin VB.TextBox mo3 
               Alignment       =   2  'Center
               BackColor       =   &H00FF8080&
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   8040
               TabIndex        =   41
               Top             =   2760
               Width           =   1455
            End
            Begin VB.TextBox mo2 
               Alignment       =   2  'Center
               BackColor       =   &H00FF8080&
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   8040
               TabIndex        =   40
               Top             =   2040
               Width           =   1455
            End
            Begin VB.TextBox mo1 
               Alignment       =   2  'Center
               BackColor       =   &H00FF8080&
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   8040
               TabIndex        =   39
               Top             =   1200
               Width           =   1455
            End
            Begin VB.TextBox c5 
               Alignment       =   2  'Center
               BackColor       =   &H00FF8080&
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   6240
               TabIndex        =   38
               Top             =   4200
               Width           =   1215
            End
            Begin VB.TextBox c4 
               Alignment       =   2  'Center
               BackColor       =   &H00FF8080&
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   6240
               TabIndex        =   37
               Top             =   3480
               Width           =   1215
            End
            Begin VB.TextBox c3 
               Alignment       =   2  'Center
               BackColor       =   &H00FF8080&
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   6240
               TabIndex        =   36
               Top             =   2760
               Width           =   1215
            End
            Begin VB.TextBox c2 
               Alignment       =   2  'Center
               BackColor       =   &H00FF8080&
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   6240
               TabIndex        =   35
               Top             =   2040
               Width           =   1215
            End
            Begin VB.TextBox c1 
               Alignment       =   2  'Center
               BackColor       =   &H00FF8080&
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   6240
               TabIndex        =   34
               Top             =   1320
               Width           =   1215
            End
            Begin VB.TextBox aq5 
               Alignment       =   2  'Center
               BackColor       =   &H00FF8080&
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   4320
               TabIndex        =   33
               Top             =   4200
               Width           =   1215
            End
            Begin VB.TextBox aq4 
               Alignment       =   2  'Center
               BackColor       =   &H00FF8080&
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   4320
               TabIndex        =   32
               Top             =   3480
               Width           =   1215
            End
            Begin VB.TextBox aq3 
               Alignment       =   2  'Center
               BackColor       =   &H00FF8080&
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   4350
               TabIndex        =   31
               Top             =   2760
               Width           =   1215
            End
            Begin VB.TextBox aq2 
               Alignment       =   2  'Center
               BackColor       =   &H00FF8080&
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   4320
               TabIndex        =   30
               Top             =   2040
               Width           =   1215
            End
            Begin VB.TextBox qu5 
               Alignment       =   2  'Center
               BackColor       =   &H00FF8080&
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   2520
               TabIndex        =   29
               Top             =   4200
               Width           =   1215
            End
            Begin VB.TextBox qu4 
               Alignment       =   2  'Center
               BackColor       =   &H00FF8080&
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   2520
               TabIndex        =   28
               Top             =   3480
               Width           =   1215
            End
            Begin VB.TextBox qu3 
               Alignment       =   2  'Center
               BackColor       =   &H00FF8080&
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   2520
               TabIndex        =   27
               Top             =   2760
               Width           =   1215
            End
            Begin VB.TextBox qu2 
               Alignment       =   2  'Center
               BackColor       =   &H00FF8080&
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   2520
               TabIndex        =   26
               Top             =   2040
               Width           =   1215
            End
            Begin VB.TextBox aq1 
               Alignment       =   2  'Center
               BackColor       =   &H00FF8080&
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   4320
               TabIndex        =   25
               Top             =   1320
               Width           =   1215
            End
            Begin VB.TextBox qu1 
               Alignment       =   2  'Center
               BackColor       =   &H00FF8080&
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   2520
               TabIndex        =   24
               Top             =   1320
               Width           =   1215
            End
            Begin VB.TextBox per 
               Alignment       =   2  'Center
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
               Height          =   615
               Left            =   12120
               TabIndex        =   22
               Top             =   2880
               Width           =   1575
            End
            Begin VB.TextBox tom 
               Alignment       =   2  'Center
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
               Height          =   615
               Left            =   12120
               TabIndex        =   21
               Top             =   2040
               Width           =   1575
            End
            Begin VB.TextBox fullmrks 
               Alignment       =   2  'Center
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
               Height          =   615
               Left            =   12120
               TabIndex        =   20
               Top             =   1200
               Width           =   1575
            End
            Begin VB.Line Line8 
               BorderColor     =   &H8000000B&
               X1              =   0
               X2              =   9720
               Y1              =   4800
               Y2              =   4800
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "TOTAL"
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
               Left            =   6960
               TabIndex        =   57
               Top             =   5040
               Width           =   855
            End
            Begin VB.Label Label17 
               BackStyle       =   0  'Transparent
               Caption         =   "MARKS OBTAINED"
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
               Left            =   8160
               TabIndex        =   56
               Top             =   720
               Width           =   1455
            End
            Begin VB.Label remarks 
               Alignment       =   2  'Center
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1095
               Left            =   10680
               TabIndex        =   47
               Top             =   4080
               Width           =   2895
            End
            Begin VB.Label Label27 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "REMARKS"
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
               Left            =   11160
               TabIndex        =   23
               Top             =   3720
               Width           =   1455
            End
            Begin VB.Label Label26 
               BackStyle       =   0  'Transparent
               Caption         =   "TOTAL OBTAINED MARKS"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Left            =   9840
               TabIndex        =   19
               Top             =   2160
               Width           =   2010
            End
            Begin VB.Label Label25 
               BackStyle       =   0  'Transparent
               Caption         =   "PERCENTGE"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Left            =   10800
               TabIndex        =   18
               Top             =   3000
               Width           =   945
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "FULL MARKS"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Left            =   10800
               TabIndex        =   17
               Top             =   1320
               Width           =   945
            End
            Begin VB.Line Line5 
               BorderColor     =   &H8000000B&
               X1              =   13920
               X2              =   13920
               Y1              =   600
               Y2              =   5880
            End
            Begin VB.Line Line17 
               BorderColor     =   &H8000000B&
               X1              =   9720
               X2              =   13920
               Y1              =   3960
               Y2              =   3960
            End
            Begin VB.Line Line16 
               BorderColor     =   &H8000000B&
               X1              =   9720
               X2              =   13920
               Y1              =   3600
               Y2              =   3600
            End
            Begin VB.Line Line15 
               BorderColor     =   &H8000000B&
               X1              =   9720
               X2              =   13920
               Y1              =   2760
               Y2              =   2760
            End
            Begin VB.Line Line14 
               BorderColor     =   &H8000000B&
               X1              =   9720
               X2              =   13920
               Y1              =   1920
               Y2              =   1920
            End
            Begin VB.Line Line13 
               BorderColor     =   &H8000000B&
               X1              =   11880
               X2              =   11880
               Y1              =   1080
               Y2              =   3600
            End
            Begin VB.Line Line12 
               BorderColor     =   &H8000000B&
               X1              =   11880
               X2              =   13920
               Y1              =   1080
               Y2              =   1080
            End
            Begin VB.Line Line11 
               BorderColor     =   &H8000000B&
               X1              =   9720
               X2              =   9720
               Y1              =   600
               Y2              =   6120
            End
            Begin VB.Line Line7 
               BorderColor     =   &H8000000B&
               X1              =   0
               X2              =   11880
               Y1              =   1080
               Y2              =   1080
            End
            Begin VB.Line Line6 
               BorderColor     =   &H8000000B&
               X1              =   0
               X2              =   13920
               Y1              =   600
               Y2              =   600
            End
            Begin VB.Line Line4 
               BorderColor     =   &H8000000B&
               X1              =   7800
               X2              =   7800
               Y1              =   600
               Y2              =   4800
            End
            Begin VB.Line Line3 
               BorderColor     =   &H8000000B&
               X1              =   5880
               X2              =   5880
               Y1              =   600
               Y2              =   4800
            End
            Begin VB.Line Line2 
               BorderColor     =   &H8000000B&
               X1              =   3960
               X2              =   3960
               Y1              =   600
               Y2              =   4800
            End
            Begin VB.Line Line1 
               BorderColor     =   &H8000000B&
               X1              =   2280
               X2              =   2280
               Y1              =   600
               Y2              =   4800
            End
            Begin VB.Label Label13 
               BackStyle       =   0  'Transparent
               Caption         =   "CORRECT"
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
               Left            =   6480
               TabIndex        =   16
               Top             =   720
               Width           =   855
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "ATTEMPTED QUESTION"
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
               Left            =   4080
               TabIndex        =   15
               Top             =   720
               Width           =   1815
            End
            Begin VB.Label Label11 
               BackStyle       =   0  'Transparent
               Caption         =   "TOTAL QUESTION"
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
               Left            =   2520
               TabIndex        =   14
               Top             =   720
               Width           =   1575
            End
            Begin VB.Label Label9 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "SUBJECT NAME"
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
               Left            =   360
               TabIndex        =   13
               Top             =   720
               Width           =   1575
            End
         End
         Begin VB.Label Label15 
            BackColor       =   &H00FF8080&
            Caption         =   "YEAR:"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1110
            TabIndex        =   67
            Top             =   1770
            Width           =   1575
         End
         Begin VB.Label fathernm 
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
            Height          =   345
            Left            =   11940
            TabIndex        =   61
            Top             =   990
            Width           =   2085
         End
         Begin VB.Label studnm 
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
            Height          =   375
            Left            =   7260
            TabIndex        =   60
            Top             =   1020
            Width           =   1995
         End
         Begin VB.Label Label31 
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   135
            Left            =   2190
            TabIndex        =   55
            Top             =   270
            Width           =   135
         End
         Begin VB.Label Label29 
            BackStyle       =   0  'Transparent
            Caption         =   "DATE :"
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
            TabIndex        =   45
            Top             =   8160
            Width           =   855
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   " COURSE :"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5940
            TabIndex        =   9
            Top             =   330
            Width           =   1575
         End
         Begin VB.Label Label7 
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
            Height          =   375
            Left            =   10080
            TabIndex        =   8
            Top             =   1020
            Width           =   1575
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "ADMISSION NUMBER:"
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
            Left            =   9960
            TabIndex        =   6
            Top             =   300
            Width           =   1935
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "COLLEGE CODE :"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   990
            TabIndex        =   5
            Top             =   1170
            Width           =   1575
         End
         Begin VB.Label Label3 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "STUDENT NAME :"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5550
            TabIndex        =   4
            Top             =   1050
            Width           =   1575
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "RESULT NO :"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1110
            TabIndex        =   3
            Top             =   315
            Width           =   1095
         End
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H8000000B&
         Height          =   7455
         Left            =   15720
         Top             =   270
         Width           =   1935
      End
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   495
      Left            =   8880
      TabIndex        =   7
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "RESULT"
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
      Left            =   6240
      TabIndex        =   0
      Top             =   0
      Width           =   6255
   End
End
Attribute VB_Name = "frmRESULT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub c1_GotFocus()
If (c1.Text = "0") Then
c1.Text = ""
End If
End Sub
Private Sub c2_GotFocus()
If (c2.Text = "0") Then
c2.Text = ""
End If
End Sub
Private Sub c3_GotFocus()
If (c3.Text = "0") Then
c3.Text = ""
End If
End Sub
Private Sub c4_GotFocus()
If (c4.Text = "0") Then
c4.Text = ""
End If
End Sub
Private Sub c5_GotFocus()
If (c5.Text = "0") Then
c5.Text = ""
End If
End Sub
Private Sub aq1_GotFocus()
If (aq1.Text = "0") Then
aq1.Text = ""
End If
End Sub
Private Sub aq1_LostFocus()
If (aq1.Text = "") Then
aq1.Text = "0"
End If
If (CInt(aq1.Text) > CInt(qu1.Text)) Then
aq1.SetFocus
End If
End Sub
Private Sub aq2_GotFocus()
If (aq2.Text = "0") Then
aq2.Text = ""
End If
End Sub
Private Sub aq2_LostFocus()
If (aq2.Text = "") Then
aq2.Text = "0"
End If
If (CInt(aq2.Text) > CInt(qu2.Text)) Then
aq2.SetFocus
End If
End Sub
Private Sub aq3_GotFocus()
If (aq3.Text = "0") Then
aq3.Text = ""
End If
End Sub
Private Sub aq3_LostFocus()
If (aq3.Text = "") Then
aq3.Text = "0"
End If
If (CInt(aq3.Text) > CInt(qu3.Text)) Then
aq3.SetFocus
End If
End Sub
Private Sub aq4_GotFocus()
If (aq4.Text = "0") Then
aq4.Text = ""
End If
End Sub
Private Sub aq4_LostFocus()
If (aq4.Text = "") Then
aq4.Text = "0"
End If
If (CInt(aq4.Text) > CInt(qu4.Text)) Then
aq4.SetFocus
End If
End Sub
Private Sub aq5_GotFocus()
If (aq5.Text = "0") Then
aq5.Text = ""
End If
End Sub
Private Sub aq5_LostFocus()
If (aq5.Text = "") Then
aq5.Text = "0"
End If
If (CInt(aq5.Text) > CInt(qu5.Text)) Then
aq5.SetFocus
End If
End Sub
Private Sub admno_LostFocus()
Module1.conn
sql = "select stud_fname,stud_fthnm,stud_year from studentdetail_master where stud_adm='" + admno.Text + "'"
Set r = c.Execute(sql)
If (r.EOF = True) Then
resultno.SetFocus
Else
studnm.Caption = r.Fields("stud_fname")
fathernm.Caption = r.Fields("stud_fthnm")
Year.Text = r.Fields("stud_year")
End If
End Sub
Private Sub c1_LostFocus()
If (c1.Text = "") Then
c1.Text = "0"
End If
mo1.Text = c1 * 2
End Sub
Private Sub c2_LostFocus()
If (c2.Text = "") Then
c2.Text = "0"
End If
mo2.Text = c2 * 2
End Sub
Private Sub c3_LostFocus()
If (c3.Text = "") Then
c3.Text = "0"
End If
mo3.Text = c3 * 2
End Sub
Private Sub c4_LostFocus()
If (c4.Text = "") Then
c4.Text = "0"
End If
mo4.Text = c4 * 2
End Sub
Private Sub c5_LostFocus()
If (c5.Text = "") Then
c5.Text = "0"
End If
mo5.Text = c5 * 2
End Sub

Private Sub Command7_Click()
frmRESULTDATA.Show
Unload Me
End Sub

Private Sub delete_Click()
If (sub1.Text = "" Or sub2.Text = "" Or sub3.Text = "" Or qu1.Text = "" Or qu2.Text = "" Or qu3.Text = "" Or aq1.Text = "" Or aq2.Text = "" Or aq3.Text = "" Or c1.Text = "" Or c2.Text = "" Or c3.Text = "" Or mo1.Text = "" Or mo2.Text = "" Or mo3.Text = "" Or total.Text = "" Or tom.Text = "" Or per.Text = "") Then
A = MsgBox("ALL FIELDS ARE MANDATORY", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
resultno.SetFocus
End If
Else
Module1.conn
A = MsgBox("WANT TO DELETE RECORD", vbInformation + vbYesNo, "INFORMATION")
If (A = vbYes) Then
sql = "delete from student_result where course='" + course.Text + "' and admno='" + admno.Text + "'"
Set r = c.Execute(sql)
A = MsgBox("RECORD DELETED", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
Call ref
resultno.SetFocus
End If
Else
If (A = vbNo) Then
resultno.SetFocus
End If
End If
End If
End Sub
Private Sub Form_Load()
dt1.Visible = False
resultno.Locked = True
per.Text = "0"
tom.Text = "0"
total.Text = "0"
mo1.Locked = True
mo2.Locked = True
mo3.Locked = True
mo4.Locked = True
mo5.Locked = True
c1.Text = "0"
c2.Text = "0"
c3.Text = "0"
c4.Text = "0"
c5.Text = "0"
qu1.Text = "0"
qu2.Text = "0"
qu3.Text = "0"
qu4.Text = "0"
qu5.Text = "0"
aq1.Text = "0"
aq2.Text = "0"
aq3.Text = "0"
aq4.Text = "0"
aq5.Text = "0"
total.Locked = True
tom.Locked = True
per.Locked = True
Update.Enabled = False
mo1.Text = "0"
mo2.Text = "0"
mo3.Text = "0"
mo4.Text = "0"
mo5.Text = "0"
Delete.Enabled = False
Update.Enabled = False
clgcode.Locked = True
fullmrks.Locked = True
clgcode.Text = "204"
fullmrks.Text = "150"
Module1.conn
sql = "select code from CourseDetail_Master"
Set r = c.Execute(sql)
While (r.EOF = False)
course.AddItem r.Fields("code")
r.MoveNext
Wend
sql = "select count(resno) from student_result"
Set r = c.Execute(sql)
i = r.Fields(0)
If (i < 1) Then
resultno.Text = "R00" & (1)
Else
sql = "select max(resno) from student_result"
Set r = c.Execute(sql)
A = r.Fields(0)
b = Right(A, 1)
'MsgBox b
b = b + 1
resultno.Text = "R00" & b
End If
End Sub
Private Sub admno_gotfocus()
Module1.conn
sql = "select stud_adm from studentdetail_master where course = '" + course.Text + "'"
Set r = c.Execute(sql)
While (r.EOF = False)
admno.AddItem r.Fields("stud_adm")
r.MoveNext
Wend
End Sub
Private Sub course_lostfocus()
admno.clear
End Sub
Private Sub ref()
Module1.conn
sql = "select count(resno) from student_result"
Set r = c.Execute(sql)
i = r.Fields(0)
If (i < 1) Then
resultno.Text = "R00" & (1)
Else
sql = "select max(resno) from student_result"
Set r = c.Execute(sql)
A = r.Fields(0)
b = Right(A, 1)
'MsgBox b
b = b + 1
resultno.Text = "R00" & b
End If
Year.Text = ""
studnm.Caption = ""
fathernm.Caption = ""
sub1.Text = ""
sub2.Text = ""
sub3.Text = ""
sub4.Text = ""
sub5.Text = ""
qu1.Text = ""
qu2.Text = ""
qu3.Text = ""
qu4.Text = ""
qu5.Text = ""
c1.Text = ""
c2.Text = ""
c3.Text = ""
c4.Text = ""
c5.Text = ""
aq1.Text = ""
aq2.Text = ""
aq3.Text = ""
aq4.Text = ""
aq5.Text = ""
mo1.Text = ""
mo2.Text = ""
mo3.Text = ""
mo4.Text = ""
mo5.Text = ""
tom.Text = ""
per.Text = ""
remarks.Caption = ""
total.Text = ""
End Sub
Private Sub per_lostfocus()
If (per.Text >= 0 And per.Text <= 50) Then
remarks.Caption = "BELOW AVERAGE"
ElseIf (per.Text >= 51 And per.Text <= 60) Then
remarks.Caption = "AVERAGE"
ElseIf (per.Text >= 61 And per.Text <= 70) Then
remarks.Caption = "SATISFACTORY"
ElseIf (per.Text >= 71 And per.Text <= 80) Then
remarks.Caption = "GOOD"
ElseIf (per.Text >= 81 And per.Text <= 90) Then
remarks.Caption = "VERY GOOD"
ElseIf (per.Text >= 91 And per.Text <= 100) Then
remarks.Caption = "EXCELLENT"
End If
End Sub


Private Sub qu1_GotFocus()
If (qu1.Text = "0") Then
qu1.Text = ""
End If
End Sub

Private Sub qu1_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
qu1.Locked = False
Else
qu1.Locked = True
End If
End Sub
Private Sub qu2_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
qu2.Locked = False
Else
qu2.Locked = True
End If
End Sub
Private Sub qu3_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
qu3.Locked = False
Else
qu3.Locked = True
End If
End Sub
Private Sub qu4_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
qu4.Locked = False
Else
qu4.Locked = True
End If
End Sub
Private Sub qu5_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
qu5.Locked = False
Else
qu5.Locked = True
End If
End Sub
Private Sub qu1_LostFocus()
If (qu1.Text = "") Then
qu1.Text = "0"
End If
If (qu1.Text > 25) Then
A = MsgBox("MAXIMUM NUMBER OF QUESTION SHOULD BE 25", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
qu1.SetFocus
End If
End If
End Sub
Private Sub qu2_GotFocus()
If (qu2.Text = "0") Then
qu2.Text = ""
End If
End Sub

Private Sub qu2_LostFocus()
If (qu2.Text = "") Then
qu2.Text = "0"
End If
If (qu2.Text > 25) Then
A = MsgBox("MAXIMUM NUMBER OF QUESTION SHOULD BE 25", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
qu2.SetFocus
End If
End If
End Sub
Private Sub qu3_GotFocus()
If (qu3.Text = "0") Then
qu3.Text = ""
End If
End Sub
Private Sub qu3_LostFocus()
If (qu3.Text = "") Then
qu3.Text = "0"
End If
If (qu3.Text > 25) Then
A = MsgBox("MAXIMUM NUMBER OF QUESTION SHOULD BE 25", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
qu3.SetFocus
End If
End If
End Sub
Private Sub qu4_GotFocus()
If (qu4.Text = "0") Then
qu4.Text = ""
End If
End Sub
Private Sub qu4_LostFocus()
If (qu4.Text = "") Then
qu4.Text = "0"
End If
If (qu4.Text > 25) Then
A = MsgBox("MAXIMUM NUMBER OF QUESTION SHOULD BE 25", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
qu4.SetFocus
End If
End If
End Sub
Private Sub qu5_GotFocus()
If (qu5.Text = "0") Then
qu5.Text = ""
End If
End Sub
Private Sub qu5_LostFocus()
If (qu5.Text = "") Then
qu5.Text = "0"
End If
If (qu5.Text > 25) Then
A = MsgBox("MAXIMUM NUMBER OF QUESTION SHOULD BE 25", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
qu5.SetFocus
End If
End If
End Sub

Private Sub refresh_Click()
dt.Visible = True
dt1.Visible = False
Call ref
Delete.Enabled = False
Update.Enabled = False
Save.Enabled = True
c1.Text = "0"
c2.Text = "0"
c3.Text = "0"
c4.Text = "0"
c5.Text = "0"
qu1.Text = "0"
qu2.Text = "0"
qu3.Text = "0"
qu4.Text = "0"
qu5.Text = "0"
aq1.Text = "0"
aq2.Text = "0"
aq3.Text = "0"
aq4.Text = "0"
aq5.Text = "0"
mo1.Text = "0"
mo2.Text = "0"
mo3.Text = "0"
mo4.Text = "0"
mo5.Text = "0"
per.Text = "0"
tom.Text = "0"
total.Text = "0"
End Sub

Private Sub save_Click()
On Error GoTo label
If (sub1.Text = "" Or sub2.Text = "" Or sub3.Text = "" Or qu1.Text = "" Or qu2.Text = "" Or qu3.Text = "" Or aq1.Text = "" Or aq2.Text = "" Or aq3.Text = "" Or c1.Text = "" Or c2.Text = "" Or c3.Text = "" Or mo1.Text = "" Or mo2.Text = "" Or mo3.Text = "" Or total.Text = "" Or tom.Text = "" Or per.Text = "") Then
A = MsgBox("ALL FIELDS ARE MANDATORY", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
resultno.SetFocus
End If
Else
Module1.conn
A = MsgBox("WANT TO SAVE RECORD", vbQuestion + vbOKCancel, "SAVE")
If (A = vbOK) Then
sql = "insert into student_result values('" + resultno.Text + "','" + course.Text + "','" + admno.Text + "'," + clgcode.Text + ",'" + sub1.Text + "','" + sub2.Text + "','" + sub3.Text + "','" + sub4.Text + "','" + sub5.Text + "','" + qu1.Text + "','" + qu2.Text + "','" + qu3.Text + "','" + qu4.Text + "','" + qu5.Text + "','" + aq1.Text + "','" + aq2.Text + "','" + aq3.Text + "','" + aq4.Text + "','" + aq5.Text + "','" + c1.Text + "','" + c2.Text + "','" + c3.Text + "','" + c4.Text + "','" + c5.Text + "','" + mo1.Text + "','" + mo2.Text + "','" + mo3.Text + "','" + mo4.Text + "','" + mo5.Text + "'," + total.Text + "," + per.Text + ",'" + remarks.Caption + "','" + Format(dt.Value, "dd/MMM/yyyy") + "','" + studnm.Caption + "','" + Year.Text + "','" + tom.Text + "','" + studnm.Caption + "')"
Set r = c.Execute(sql)
A = MsgBox("RECORD SAVE", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
Call ref
A = Right(resultno.Text, 1)
A = A + 1
resultno.Text = "R00" & A
resultno.SetFocus
End If
Else
If (A = vbCancel) Then
resultno.SetFocus
Exit Sub
label:
A = MsgBox("DATA ALREADY EXIST", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
resultno.SetFocus
End If
End If
End If
End If
End Sub
Private Sub tom_GotFocus()
tom.Text = total.Text
End Sub
Private Sub tom_LostFocus()
per.Text = (CInt(total.Text) / 150) * 100
End Sub

Private Sub total_GotFocus()
If (mo4.Text = "" And mo5.Text = "") Then
A = CInt(mo1.Text)
b = CInt(mo2.Text)
f = CInt(mo3.Text)
total.Text = A + b + f
Else
If (mo4.Text = "") Then
A = CInt(mo1.Text)
b = CInt(mo2.Text)
f = CInt(mo3.Text)
e = CInt(mo5.Text)
total.Text = A + b + f + e
Else
If (mo5.Text = "") Then
A = CInt(mo1.Text)
b = CInt(mo2.Text)
f = CInt(mo3.Text)
d = CInt(mo4.Text)
total.Text = A + b + f + d
Else
If (mo4.Text <> "" And mo5.Text <> "") Then
A = CInt(mo1.Text)
b = CInt(mo2.Text)
f = CInt(mo3.Text)
d = CInt(mo4.Text)
e = CInt(mo5.Text)
total.Text = A + b + f + d + e
End If
End If
End If
End If
End Sub

Private Sub total_LostFocus()
If (total.Text > 150) Then
A = MsgBox("WRONG EVALUATION", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
total.Text = "0"
total.SetFocus
End If
End If
End Sub

Private Sub update_Click()
If (sub1.Text = "" Or sub2.Text = "" Or sub3.Text = "" Or qu1.Text = "" Or qu2.Text = "" Or qu3.Text = "" Or aq1.Text = "" Or aq2.Text = "" Or aq3.Text = "" Or c1.Text = "" Or c2.Text = "" Or c3.Text = "" Or mo1.Text = "" Or mo2.Text = "" Or mo3.Text = "" Or total.Text = "" Or tom.Text = "" Or per.Text = "") Then
A = MsgBox("ALL FIELDS ARE MANDATORY", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
resultno.SetFocus
End If
Else
Module1.conn
A = MsgBox("WANT TO UPDATE RECORD", vbInformation + vbYesNo, "SAVE")
If (A = vbYes) Then
sql = "update student_result set sub1='" + sub1.Text + "',sub2='" + sub2.Text + "',sub3='" + sub3.Text + "',sub4='" + sub4.Text + "',sub5='" + sub5.Text + "',tq1='" + qu1.Text + "',tq2='" + qu2.Text + "',tq3='" + qu3.Text + "',tq4='" + qu4.Text + "',tq5='" + qu5.Text + "',aq1='" + aq1.Text + "',aq2='" + aq2.Text + "',aq3='" + aq3.Text + "',aq4='" + aq4.Text + "',aq5='" + aq5.Text + "',corr1='" + c1.Text + "',corr2='" + c2.Text + "',corr3='" + c3.Text + "',corr4='" + c4.Text + "',corr5='" + c5.Text + "',mo1='" + mo1.Text + "',mo2='" + mo2.Text + "',mo3='" + mo3.Text + "',mo4='" + mo4.Text + "',mo5='" + mo5.Text + "'"
Set r = c.Execute(sql)
A = MsgBox("RECORD UPDATED", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
Call ref
resultno.SetFocus
End If
Else
If (A = vbNo) Then
resultno.SetFocus
End If
End If
End If
End Sub

Private Sub view_Click()
dt1.Visible = True
dt.Visible = False
On Error GoTo label
Delete.Enabled = True
Update.Enabled = True
Save.Enabled = False
If (admno.Text = "" Or course.Text = "") Then
A = MsgBox("ALL FIEDLS ARE MANDATORY", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
resultno.SetFocus
End If
Else
Module1.conn
sql = "select *from student_result where course='" + course.Text + "' and admno='" + admno.Text + "'"
Set r = c.Execute(sql)
resultno.Text = r.Fields(0)
course.Text = r.Fields(1)
admno.Text = r.Fields(2)
sub1.Text = r.Fields(4)
sub2.Text = r.Fields(5)
sub3.Text = r.Fields(6)
If (IsNull(r.Fields(7))) Then
sub4.Text = ""
Else
sub4.Text = r.Fields(7)
End If
If (IsNull(r.Fields(8))) Then
sub5.Text = ""
Else
sub5.Text = r.Fields(8)
End If
qu1.Text = r.Fields(9)
qu2.Text = r.Fields(10)
qu3.Text = r.Fields(11)
qu4.Text = r.Fields(12)
qu5.Text = r.Fields(13)
aq1.Text = r.Fields(14)
aq2.Text = r.Fields(15)
aq3.Text = r.Fields(16)
aq4.Text = r.Fields(17)
aq5.Text = r.Fields(18)
c1.Text = r.Fields(19)
c2.Text = r.Fields(20)
c3.Text = r.Fields(21)
c4.Text = r.Fields(22)
c5.Text = r.Fields(23)
mo1.Text = r.Fields(24)
mo2.Text = r.Fields(25)
mo3.Text = r.Fields(26)
mo4.Text = r.Fields(27)
mo5.Text = r.Fields(28)
total.Text = r.Fields(29)
per.Text = r.Fields(30)
remarks.Caption = r.Fields(31)
dt1.Text = r.Fields(32)
Year.Text = r.Fields(34)
tom.Text = r.Fields(35)
studnm.Caption = r.Fields(33)
Exit Sub
label:
A = MsgBox("DATA NOT FOUND", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
resultno.SetFocus
End If
End If
End Sub
Private Sub aq1_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
aq1.Locked = False
Else
aq1.Locked = True
End If
End Sub
Private Sub aq2_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
aq2.Locked = False
Else
aq2.Locked = True
End If
End Sub
Private Sub aq3_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
aq3.Locked = False
Else
aq3.Locked = True
End If
End Sub
Private Sub aq4_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
aq4.Locked = False
Else
aq4.Locked = True
End If
End Sub
Private Sub aq5_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
aq5.Locked = False
Else
aq5.Locked = True
End If
End Sub
Private Sub c1_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
c1.Locked = False
Else
c1.Locked = True
End If
End Sub
Private Sub c2_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
c2.Locked = False
Else
c2.Locked = True
End If
End Sub
Private Sub c3_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
c3.Locked = False
Else
c3.Locked = True
End If
End Sub
Private Sub c4_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
c4.Locked = False
Else
c4.Locked = True
End If
End Sub
Private Sub c5_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
c5.Locked = False
Else
c5.Locked = True
End If
End Sub
Private Sub dt_lostfocus()
A = Date
If (dt.Value > A) Then
dt.SetFocus
End If
End Sub
