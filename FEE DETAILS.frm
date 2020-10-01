VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmFEE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FEE"
   ClientHeight    =   10380
   ClientLeft      =   1200
   ClientTop       =   750
   ClientWidth     =   18975
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FEE DETAILS.frx":0000
   ScaleHeight     =   10380
   ScaleWidth      =   18975
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Height          =   9735
      Left            =   810
      TabIndex        =   1
      Top             =   630
      Width           =   17295
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
         Left            =   15840
         Picture         =   "FEE DETAILS.frx":1E9608
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   6480
         Width           =   975
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
         Left            =   15840
         Picture         =   "FEE DETAILS.frx":1E9E8B
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   5280
         Width           =   975
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
         Left            =   15840
         Picture         =   "FEE DETAILS.frx":1EA52F
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   4080
         Width           =   975
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
         Left            =   15840
         Picture         =   "FEE DETAILS.frx":1EACBF
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   2880
         Width           =   975
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
         Left            =   15840
         Picture         =   "FEE DETAILS.frx":1EB22B
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   1680
         Width           =   975
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
         Left            =   15840
         Picture         =   "FEE DETAILS.frx":1EB5E2
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   480
         Width           =   975
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FF8080&
         Height          =   9495
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   15135
         Begin MSAdodcLib.Adodc Adodc1 
            Height          =   330
            Left            =   10860
            Top             =   3630
            Visible         =   0   'False
            Width           =   1995
            _ExtentX        =   3519
            _ExtentY        =   582
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   8
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   1
            Appearance      =   1
            BackColor       =   16744576
            ForeColor       =   -2147483640
            Orientation     =   0
            Enabled         =   -1
            Connect         =   ""
            OLEDBString     =   ""
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   ""
            Caption         =   "Adodc1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin VB.Frame Frame8 
            BackColor       =   &H00FF8080&
            Height          =   5415
            Left            =   7320
            TabIndex        =   33
            Top             =   3840
            Width           =   7695
         End
         Begin VB.Frame Frame7 
            BackColor       =   &H00FF8080&
            Caption         =   "PAID AMOUNT"
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
            Left            =   120
            TabIndex        =   30
            Top             =   8280
            Width           =   6855
            Begin VB.TextBox paid 
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
               Left            =   1320
               TabIndex        =   32
               Top             =   360
               Width           =   1605
            End
            Begin VB.Label Label35 
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
               TabIndex        =   48
               Top             =   240
               Width           =   135
            End
            Begin VB.Label Label21 
               BackStyle       =   0  'Transparent
               Caption         =   "PAID FEE :"
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
               Left            =   120
               TabIndex        =   31
               Top             =   360
               Width           =   1095
            End
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00FF8080&
            Caption         =   "PAYMENT DETAILS"
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
            TabIndex        =   21
            Top             =   6000
            Width           =   6855
            Begin VB.TextBox checkdt 
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
               Height          =   315
               Left            =   4950
               TabIndex        =   29
               Top             =   960
               Width           =   1575
            End
            Begin VB.TextBox checkno 
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
               Height          =   315
               Left            =   1800
               TabIndex        =   28
               Top             =   960
               Width           =   1575
            End
            Begin VB.TextBox bnknm 
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
               Height          =   315
               Left            =   1800
               TabIndex        =   27
               Top             =   1560
               Width           =   4335
            End
            Begin VB.ComboBox paymode 
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
               ItemData        =   "FEE DETAILS.frx":1EBDBC
               Left            =   1740
               List            =   "FEE DETAILS.frx":1EBDC6
               Style           =   2  'Dropdown List
               TabIndex        =   26
               Top             =   330
               Width           =   4335
            End
            Begin VB.Label Label34 
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
               TabIndex        =   47
               Top             =   1440
               Width           =   135
            End
            Begin VB.Label Label33 
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
               TabIndex        =   46
               Top             =   840
               Width           =   135
            End
            Begin VB.Label Label32 
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
               TabIndex        =   45
               Top             =   840
               Width           =   135
            End
            Begin VB.Label Label31 
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
               Left            =   1560
               TabIndex        =   44
               Top             =   240
               Width           =   135
            End
            Begin VB.Label Label20 
               BackStyle       =   0  'Transparent
               Caption         =   "BANK NAME :"
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
               TabIndex        =   25
               Top             =   1560
               Width           =   1095
            End
            Begin VB.Label Label19 
               BackStyle       =   0  'Transparent
               Caption         =   "CHECK DATE :"
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
               Left            =   3600
               TabIndex        =   24
               Top             =   960
               Width           =   1215
            End
            Begin VB.Label Label18 
               BackStyle       =   0  'Transparent
               Caption         =   "CHECK NO :"
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
               Left            =   480
               TabIndex        =   23
               Top             =   960
               Width           =   1095
            End
            Begin VB.Label Label17 
               BackStyle       =   0  'Transparent
               Caption         =   "PAYMENT MODE :"
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
               TabIndex        =   22
               Top             =   360
               Width           =   1455
            End
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00FF8080&
            Caption         =   "FEE DETAILS"
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
            TabIndex        =   14
            Top             =   3840
            Width           =   6855
            Begin VB.TextBox total 
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
               Left            =   1650
               TabIndex        =   52
               Top             =   1530
               Width           =   1575
            End
            Begin VB.TextBox disc 
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
               Left            =   5040
               TabIndex        =   50
               Top             =   960
               Width           =   1575
            End
            Begin VB.ComboBox feenm 
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
               ItemData        =   "FEE DETAILS.frx":1EBDD7
               Left            =   1680
               List            =   "FEE DETAILS.frx":1EBDE4
               Style           =   2  'Dropdown List
               TabIndex        =   20
               Top             =   360
               Width           =   1575
            End
            Begin VB.TextBox ltfee 
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
               Left            =   1620
               TabIndex        =   19
               Top             =   930
               Width           =   1575
            End
            Begin VB.TextBox amnt 
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
               Left            =   5040
               TabIndex        =   18
               Top             =   360
               Width           =   1575
            End
            Begin VB.Label Label37 
               BackStyle       =   0  'Transparent
               Caption         =   "TOTAL :"
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
               Left            =   180
               TabIndex        =   51
               Top             =   1530
               Width           =   1215
            End
            Begin VB.Label Label13 
               BackStyle       =   0  'Transparent
               Caption         =   "DISCOUNT :"
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
               Left            =   3480
               TabIndex        =   49
               Top             =   960
               Width           =   1215
            End
            Begin VB.Label Label30 
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
               Left            =   4200
               TabIndex        =   43
               Top             =   240
               Width           =   135
            End
            Begin VB.Label Label29 
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
               Left            =   720
               TabIndex        =   42
               Top             =   240
               Width           =   135
            End
            Begin VB.Label Label16 
               BackStyle       =   0  'Transparent
               Caption         =   "LATE FEE :"
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
               Left            =   120
               TabIndex        =   17
               Top             =   960
               Width           =   1455
            End
            Begin VB.Label Label15 
               BackStyle       =   0  'Transparent
               Caption         =   "NAME :"
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
               Left            =   120
               TabIndex        =   16
               Top             =   360
               Width           =   1455
            End
            Begin VB.Label Label14 
               BackStyle       =   0  'Transparent
               Caption         =   "AMOUNT :"
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
               Left            =   3480
               TabIndex        =   15
               Top             =   360
               Width           =   1455
            End
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00FF8080&
            Caption         =   "STUDENT DETAILS"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2175
            Left            =   120
            TabIndex        =   9
            Top             =   1410
            Width           =   14895
            Begin VB.ComboBox sem 
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
               Left            =   10620
               Style           =   2  'Dropdown List
               TabIndex        =   59
               Top             =   810
               Width           =   2415
            End
            Begin VB.ComboBox admno 
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
               Left            =   6900
               Style           =   2  'Dropdown List
               TabIndex        =   58
               Top             =   360
               Width           =   2415
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
               Left            =   2220
               Style           =   2  'Dropdown List
               TabIndex        =   57
               Top             =   360
               Width           =   2415
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "BATCH NO :"
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
               Left            =   5310
               TabIndex        =   62
               Top             =   893
               Width           =   1695
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "YEAR :"
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
               Left            =   330
               TabIndex        =   61
               Top             =   893
               Width           =   1455
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   "SEMESTER :"
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
               Left            =   9630
               TabIndex        =   60
               Top             =   893
               Width           =   1095
            End
            Begin VB.Label fthnm 
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
               Height          =   315
               Left            =   6870
               TabIndex        =   56
               Top             =   1470
               Width           =   2415
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
               Height          =   315
               Left            =   2160
               TabIndex        =   55
               Top             =   1530
               Width           =   2415
            End
            Begin VB.Label year 
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
               Height          =   315
               Left            =   2220
               TabIndex        =   54
               Top             =   840
               Width           =   2415
            End
            Begin VB.Label batno 
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
               Height          =   315
               Left            =   6870
               TabIndex        =   53
               Top             =   870
               Width           =   2415
            End
            Begin VB.Label Label11 
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
               Left            =   5280
               TabIndex        =   13
               Top             =   1530
               Width           =   1335
            End
            Begin VB.Label Label6 
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
               Left            =   390
               TabIndex        =   12
               Top             =   1590
               Width           =   1455
            End
            Begin VB.Label Label4 
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
               Height          =   375
               Left            =   360
               TabIndex        =   11
               Top             =   420
               Width           =   855
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "ADMISSION NO:"
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
               Left            =   5280
               TabIndex        =   10
               Top             =   420
               Width           =   1695
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00FF8080&
            Caption         =   "RECIEPT DETAILS"
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
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   14895
            Begin VB.TextBox dt1 
               Height          =   345
               Left            =   9660
               TabIndex        =   63
               Top             =   330
               Width           =   2145
            End
            Begin VB.TextBox rcptsno 
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
               Left            =   2184
               TabIndex        =   7
               Top             =   360
               Width           =   2535
            End
            Begin MSComCtl2.DTPicker dt 
               Height          =   375
               Left            =   7140
               TabIndex        =   8
               Top             =   300
               Width           =   2295
               _ExtentX        =   4048
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
               CustomFormat    =   "dd/MMM/yyyy"
               Format          =   109445123
               CurrentDate     =   43113
            End
            Begin VB.Label Label26 
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
               Left            =   6600
               TabIndex        =   41
               Top             =   240
               Width           =   135
            End
            Begin VB.Label Label25 
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
               Left            =   1620
               TabIndex        =   40
               Top             =   240
               Width           =   135
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "."
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
               Left            =   11040
               TabIndex        =   6
               Top             =   360
               Width           =   735
            End
            Begin VB.Label Label8 
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
               Height          =   375
               Left            =   6120
               TabIndex        =   5
               Top             =   360
               Width           =   735
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "RECIEPT SERIAL NO :"
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
               Left            =   150
               TabIndex        =   4
               Top             =   360
               Width           =   1695
            End
         End
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H8000000B&
         Height          =   7395
         Left            =   15600
         Top             =   300
         Width           =   1575
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "FEE "
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
      Left            =   8640
      TabIndex        =   0
      Top             =   0
      Width           =   1935
   End
End
Attribute VB_Name = "frmFEE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ref()
studnm.Caption = ""
amnt.Text = CInt(0)
ltfee.Text = CInt(0)
disc.Text = CInt(0)
batno.Caption = ""
Year.Caption = ""
fthnm.Caption = ""
total.Text = CInt(0)
checkno.Text = ""
'pamt.Text = CInt(0)
'dues.Text = CInt(0)
paid.Text = CInt(0)
End Sub
Private Sub admno_gotfocus()
Module1.conn
sql = "select stud_adm from studentdetail_master where course='" + course.Text + "'"
Set r = c.Execute(sql)
While (r.EOF = False)
admno.AddItem r.Fields("stud_adm")
r.MoveNext
Wend
End Sub

Private Sub admno_LostFocus()
Module1.conn
sql = "select stud_fname,stud_fthnm,stud_batch,stud_year from studentdetail_master where stud_adm='" + admno.Text + "'"
Set r = c.Execute(sql)
If (r.EOF = True) Then
admno.SetFocus
Else
studnm.Caption = r.Fields("stud_fname")
fthnm.Caption = r.Fields("stud_fthnm")
Year.Caption = r.Fields("stud_year")
batno.Caption = r.Fields("stud_batch")
End If
'sql = "select dues from studentfeedetail_master where adm_no = '" + ADMNO.Text + "' "
'Set r = c.Execute(sql)
'dues.Text = r.Fields(0)
End Sub
Private Sub amnt_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
amnt.Locked = False
Else
amnt.Locked = True
End If
End Sub

Private Sub delete_Click()
If (Year.Caption = "" Or studnm.Caption = "" Or course.Text = "" Or admno.Text = "" Or feenm.Text = "" Or amnt.Text = "" Or ltfee.Text = "" Or disc.Text = "" Or total.Text = "" Or paymode.Text = "") Then
A = MsgBox("ALL FIELDS ARE MANDATORY", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
rcptsno.SetFocus
End If
Else
Module1.conn
A = MsgBox("WANT TO DELETE RECORD", vbQuestion + vbYesNo, "DELETE")
If (A = vbYes) Then
sql = "delete from studentfeedetail_master where admno_no='" + admno.Text + "'"
Set r = c.Execute(sql)
A = MsgBox("RECORD DELETED", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
Call ref
rcptsno.SetFocus
Else
If (A = vbNo) Then
rcptsno.SetFocus
End If
End If
End If
End If
End Sub

Private Sub dt_lostfocus()
Dim d As Date
d = dt.Value
diff = DateDiff("yyyy", dt.Value, Date)
If (diff > 3) Then
A = MsgBox("INVALID DATE", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOKOnly) Then
dt.SetFocus
End If
End If
'Dim age As Long
'age = DateDiff("yyyy", dob.Value, Date)
'If (age <= 18 Or age >= 60) Then
'MsgBox ("YOU ARE NOT ELLIGIBLE FOR THIS COURSE")
'dob.SetFocus
'End If
End Sub

Private Sub feenm_GotFocus()
If (course.Text = "") Then
A = MsgBox("PLEASE CHOOSE COURSE FIRST", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
course.SetFocus
End If
End If
End Sub

Private Sub feenm_LostFocus()
Module1.conn
If (feenm.Text = "REGISTRATION FEE") Then
sql = "select reg_fee from coursedetail_master where code = '" + course.Text + "'"
Set r = c.Execute(sql)
amnt.Text = r.Fields(0)
ElseIf (feenm.Text = "ADMISSION FEE") Then
sql = "select adm_fee from coursedetail_master where code = '" + course.Text + "'"
Set r = c.Execute(sql)
amnt.Text = r.Fields(0)
ElseIf (feenm.Text = "SEMESTER FEE") Then
sql = "select sem_fee from coursedetail_master where code = '" + course.Text + "'"
Set r = c.Execute(sql)
amnt.Text = r.Fields(0)
End If
End Sub

Private Sub ltfee_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
ltfee.Locked = False
Else
ltfee.Locked = True
End If
End Sub
Private Sub course_Click()
admno.clear
End Sub
Private Sub disc_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
disc.Locked = False
Else
disc.Locked = True
End If
End Sub
Private Sub Form_Load()
rcptsno.Locked = True
ltfee.Text = CInt(0)
disc.Text = CInt(0)
'dues.Text = CInt(0)
amnt.Text = CInt(0)
paid.Text = CInt(0)
total.Text = CInt(0)
'pamt.Text = CInt(0)
'checkdt.Text = Format(Now, "dddd, mmmm dd, yyyy")
Module1.conn
sql = "select count(reciept_no) from studentfeedetail_master"
Set r = c.Execute(sql)
i = r.Fields(0)
If (i < 1) Then
rcptsno.Text = "RCP01"
Else
sql = "select max(reciept_no) from studentfeedetail_master"
Set r = c.Execute(sql)
i = r.Fields(0)
A = Right(i, 1)
A = A + 1
rcptsno.Text = "RCP0" & A
'MsgBox a
End If
sem.AddItem "1 semester"
sem.AddItem "2 semester"
sem.AddItem "3 semester"
sem.AddItem "4 semester"
sem.AddItem "5 semester"
sem.AddItem "6 semester"
Module1.conn
sql = "select code from CourseDetail_Master"
Set r = c.Execute(sql)
While (r.EOF = False)
course.AddItem r.Fields("code")
r.MoveNext
Wend
End Sub

Private Sub paid_Click()
paid.Text = ""
End Sub

Private Sub paid_GotFocus()
If paid.Text > total.Text Then
A = MsgBox("PAID AMMOUNT IS LARGER THAN TOTAL AMMOUNT", vbInformation + vbOKOnly, "INFORMATON")
If (A = vbOK) Then
paid.Text = ""
'pamt.Text = ""
paid.SetFocus
End If
End If
End Sub

Private Sub paid_LostFocus()
A = CInt(amnt.Text)
b = CInt(paid.Text)
If (A > b) Or (A < b) Then
f = MsgBox("PLEASE ENTER APPROPRIATE AMMOUNT", vbInformation + vbOKOnly, "INFORMATION")
If (f = vbOK) Then
paid.Text = ""
paid.SetFocus
End If
End If
End Sub

Private Sub paymode_LostFocus()
If (paymode.Text = "CASH") Then
checkno.Enabled = False
checkdt.Enabled = False
bnknm.Enabled = False
paid.SetFocus
Else
checkno.Enabled = True
checkdt.Enabled = True
bnknm.Enabled = True
checkno.SetFocus
End If
End Sub

Private Sub refresh_Click()
Module1.conn
sql = "select count(reciept_no) from studentfeedetail_master"
Set r = c.Execute(sql)
i = r.Fields(0)
If (i < 1) Then
rcptsno.Text = "RCP01"
Else
sql = "select max(reciept_no) from studentfeedetail_master"
Set r = c.Execute(sql)
i = r.Fields(0)
A = Right(i, 1)
A = A + 1
rcptsno.Text = "RCP0" & A
'MsgBox a
End If
Call ref
End Sub

Private Sub save_Click()
If (Year.Caption = "" Or studnm.Caption = "" Or course.Text = "" Or admno.Text = "" Or feenm.Text = "" Or amnt.Text = "" Or ltfee.Text = "" Or disc.Text = "" Or total.Text = "" Or paymode.Text = "") Then
A = MsgBox("ALL FIELDS ARE MANDATORY", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
rcptsno.SetFocus
End If
Else
Module1.conn
A = MsgBox("WANT TO SAVE RECORD", vbQuestion + vbYesNo, "SAVE")
If (A = vbYes) Then
sql = "insert into studentfeedetail_master values('" + rcptsno.Text + "','" + Format(dt.Value, "dd/MMM/yyyy") + "','" + studnm.Caption + "','" + course.Text + "','" + sem.Text + "','" + feenm.Text + "'," + amnt.Text + ",'" + ltfee.Text + "','" + disc.Text + "'," + total.Text + ",'" + paymode.Text + "','" + checkno.Text + "','" + checkdt.Text + "','" + bnknm.Text + "','" + admno.Text + "')"
Set r = c.Execute(sql)
A = MsgBox("RECORD SAVED", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
i = rcptsno.Text
A = Right(i, 1)
A = A + 1
rcptsno.Text = "RCP0" & A
Call ref
rcptsno.SetFocus
Else
If (A = vbNo) Then
rcptsno.SetFocus
End If
End If
End If
End If
End Sub
Private Sub total_GotFocus()
A = CInt(amnt.Text)
b = CInt(ltfee.Text)
d = CInt(disc.Text)
'e = CInt(dues.Text)

total.Text = A + b + d + e

End Sub

Private Sub update_Click()
If (Year.Caption = "" Or studnm.Caption = "" Or course.Text = "" Or admno.Text = "" Or feenm.Text = "" Or amnt.Text = "" Or ltfee.Text = "" Or disc.Text = "" Or total.Text = "" Or paymode.Text = "") Then
A = MsgBox("ALL FIELDS ARE MANDATORY", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
rcptsno.SetFocus
End If
Else
Module1.conn
A = MsgBox("")
End Sub

Private Sub view_Click()
If (admno.Text = "" And course.Text = "" And sem.Text = "") Then
A = MsgBox("PLEASE CHOOSE COURSE AND ADMISSION NUMBER AND SEMESTER", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
rcptsno.SetFocus
End If
Else
Module1.conn
sql = "select *from studentfeedetail_master where course = '" + course.Text + "' and  adm_no='" + admno.Text + "' and sem = '" + sem.Text + "' "
Set r = c.Execute(sql)
rcptsno.Text = r.Fields("reciept_no")
dt1.Text = r.Fields("d_ate")
'course.Text = r.Fields("course")
'sem.Text = r.Fields("sem")
feenm.Text = r.Fields("fee_nm")
amnt.Text = r.Fields("amount")
ltfee.Text = r.Fields("latefee")
disc.Text = r.Fields("discount")
total.Text = r.Fields("total")
paymode.Text = r.Fields("paymode")
'pamt.Text = r.Fields("dues")
If (IsNull(r.Fields("checkno"))) Then
checkno.Text = ""
Else
checkno.Text = r.Fields("checkno")
End If
If IsNull(r.Fields("checkdate")) Then
checkdt.Text = ""
Else
checkdt.Text = r.Fields("checkdate")
End If
If IsNull(r.Fields("bnknm")) Then
bnknm.Text = ""
Else
bnknm.Text = r.Fields("bnknm")
End If
'dues.Text = r.Fields("dues")
End If
End Sub

