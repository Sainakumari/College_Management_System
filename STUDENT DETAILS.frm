VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmSTUDENT 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   10305
   ClientLeft      =   540
   ClientTop       =   705
   ClientWidth     =   20115
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   Picture         =   "STUDENT DETAILS.frx":0000
   ScaleHeight     =   10305
   ScaleWidth      =   20115
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   405
      Left            =   4680
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   714
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
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=MSDAORA.1;User ID=cms;password=saina;Persist Security Info=False"
      OLEDBString     =   "Provider=MSDAORA.1;User ID=cms;password=saina;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select stud_adm,stud_reg,stud_fname,stud_lname,code,stud_roll,stud_year,stud_batch,stud_contact from studentdetail_master"
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9675
      Left            =   270
      TabIndex        =   60
      Top             =   600
      Width           =   19695
      Begin VB.CommandButton Command1 
         Caption         =   "DATA GRID"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   18180
         Picture         =   "STUDENT DETAILS.frx":1E9608
         Style           =   1  'Graphical
         TabIndex        =   151
         Top             =   8490
         Width           =   1035
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
         Height          =   1095
         Left            =   18120
         Picture         =   "STUDENT DETAILS.frx":1E9E57
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   5712
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
         Height          =   1095
         Left            =   18120
         Picture         =   "STUDENT DETAILS.frx":1EA4FB
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   6960
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
         Height          =   1095
         Left            =   18120
         Picture         =   "STUDENT DETAILS.frx":1EAD7E
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   4464
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
         Height          =   1095
         Left            =   18150
         Picture         =   "STUDENT DETAILS.frx":1EB50E
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   3216
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
         Height          =   1095
         Left            =   18120
         Picture         =   "STUDENT DETAILS.frx":1EBA7A
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton view 
         Caption         =   "SEARCH"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   18120
         Picture         =   "STUDENT DETAILS.frx":1EC254
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   1920
         UseMaskColor    =   -1  'True
         Width           =   1095
      End
      Begin VB.Frame frmSTUDENT 
         BackColor       =   &H00FF8080&
         Caption         =   "PERSONAL PROFILE"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   9435
         Left            =   120
         TabIndex        =   61
         Top             =   120
         Width           =   17655
         Begin VB.TextBox batch 
            BackColor       =   &H80000016&
            Height          =   345
            Left            =   5520
            TabIndex        =   157
            Top             =   780
            Width           =   1695
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H80000016&
            Caption         =   "IF LOCAL ADDRESS SAME AS FOR THAT OF PERMANENT ADDRESS"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   39
            Top             =   5190
            Width           =   5355
         End
         Begin VB.TextBox stadm 
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
            Left            =   2100
            TabIndex        =   0
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox stdob 
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
            Height          =   360
            Left            =   2100
            TabIndex        =   11
            Top             =   1680
            Width           =   1695
         End
         Begin VB.TextBox bloodgrp 
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
            Left            =   13380
            TabIndex        =   10
            Top             =   1296
            Width           =   1695
         End
         Begin VB.TextBox session 
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
            Left            =   2100
            TabIndex        =   4
            Top             =   750
            Width           =   1695
         End
         Begin VB.TextBox aadhar 
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
            Left            =   9630
            TabIndex        =   17
            Top             =   2130
            Width           =   1695
         End
         Begin VB.Frame Frame7 
            BackColor       =   &H00FF8080&
            Caption         =   "STUDENT'S PHOTO"
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
            TabIndex        =   104
            Top             =   480
            Width           =   2295
            Begin VB.Shape Shape2 
               Height          =   15
               Left            =   0
               Top             =   1920
               Width           =   2295
            End
            Begin VB.Image Image1 
               BorderStyle     =   1  'Fixed Single
               Height          =   1575
               Left            =   120
               Stretch         =   -1  'True
               Top             =   240
               Width           =   2055
            End
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
            Left            =   16410
            Picture         =   "STUDENT DETAILS.frx":1EC60B
            Style           =   1  'Graphical
            TabIndex        =   103
            Top             =   2550
            Width           =   945
         End
         Begin MSComCtl2.DTPicker dob 
            Height          =   360
            Left            =   2100
            TabIndex        =   102
            Top             =   1680
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   635
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
            Format          =   100335619
            CurrentDate     =   43091
         End
         Begin MSComDlg.CommonDialog cd1 
            Left            =   16200
            Top             =   3630
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            Filter          =   "*.jpg"
         End
         Begin VB.CommandButton Command6 
            Caption         =   "BROWSE..."
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
            Left            =   15330
            Picture         =   "STUDENT DETAILS.frx":234E95
            Style           =   1  'Graphical
            TabIndex        =   100
            Top             =   2550
            Width           =   975
         End
         Begin VB.ComboBox year 
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
            ItemData        =   "STUDENT DETAILS.frx":235742
            Left            =   9630
            List            =   "STUDENT DETAILS.frx":235744
            TabIndex        =   2
            Top             =   330
            Width           =   1695
         End
         Begin MSComCtl2.DTPicker admdate 
            Height          =   375
            Left            =   2100
            TabIndex        =   52
            Top             =   240
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
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
            CalendarBackColor=   -2147483626
            CalendarForeColor=   -2147483626
            CalendarTitleBackColor=   -2147483626
            CalendarTitleForeColor=   16744576
            CalendarTrailingForeColor=   -2147483626
            CustomFormat    =   "dd/MMM/yyyy"
            Format          =   100466691
            CurrentDate     =   43083
         End
         Begin VB.TextBox studadm 
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
            Left            =   13380
            TabIndex        =   3
            Top             =   330
            Width           =   1695
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
            Height          =   375
            Left            =   2100
            TabIndex        =   7
            Top             =   1200
            Width           =   1695
         End
         Begin VB.TextBox emailid 
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
            Left            =   5490
            TabIndex        =   16
            Top             =   2220
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
            Height          =   375
            Left            =   5490
            TabIndex        =   8
            Top             =   1259
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
            Height          =   375
            Left            =   2100
            TabIndex        =   15
            Top             =   2160
            Width           =   1695
         End
         Begin VB.TextBox regno 
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
            Left            =   13380
            TabIndex        =   6
            Top             =   813
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
            Height          =   375
            Left            =   9630
            TabIndex        =   9
            Top             =   1260
            Width           =   1695
         End
         Begin VB.TextBox rollnum 
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
            Left            =   9630
            TabIndex        =   5
            Top             =   780
            Width           =   1695
         End
         Begin VB.Frame Frame6 
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
            Height          =   2535
            Left            =   120
            TabIndex        =   84
            Top             =   2610
            Width           =   7455
            Begin VB.TextBox hno1 
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
            Begin VB.TextBox village1 
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
               Top             =   720
               Width           =   1695
            End
            Begin VB.TextBox city1 
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
               TabIndex        =   23
               Top             =   1200
               Width           =   1695
            End
            Begin VB.TextBox postoffice1 
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
               TabIndex        =   25
               Top             =   1680
               Width           =   1695
            End
            Begin VB.TextBox street1 
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
               TabIndex        =   20
               Top             =   240
               Width           =   1695
            End
            Begin VB.TextBox landmark1 
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
               Top             =   720
               Width           =   1695
            End
            Begin VB.TextBox policestation1 
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
               Top             =   1200
               Width           =   1695
            End
            Begin VB.TextBox pincode1 
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
               Top             =   1680
               Width           =   1695
            End
            Begin VB.TextBox district1 
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
               Top             =   2160
               Width           =   1695
            End
            Begin VB.TextBox state1 
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
               Top             =   2160
               Width           =   1695
            End
            Begin VB.Label Label25 
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
               TabIndex        =   149
               Top             =   1680
               Width           =   855
            End
            Begin VB.Label Label24 
               BackStyle       =   0  'Transparent
               Caption         =   "POLICE  STATION :"
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
               TabIndex        =   148
               Top             =   1200
               Width           =   1455
            End
            Begin VB.Label Label23 
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
               TabIndex        =   147
               Top             =   720
               Width           =   1095
            End
            Begin VB.Label Label22 
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
               TabIndex        =   146
               Top             =   240
               Width           =   855
            End
            Begin VB.Label Label81 
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
               TabIndex        =   145
               Top             =   2160
               Width           =   735
            End
            Begin VB.Label Label70 
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
               Left            =   4560
               TabIndex        =   135
               Top             =   2040
               Width           =   135
            End
            Begin VB.Label Label69 
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
               Left            =   4680
               TabIndex        =   134
               Top             =   120
               Width           =   135
            End
            Begin VB.Label Label68 
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
               Left            =   1080
               TabIndex        =   133
               Top             =   2040
               Width           =   135
            End
            Begin VB.Label Label67 
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
               Left            =   1200
               TabIndex        =   132
               Top             =   1560
               Width           =   135
            End
            Begin VB.Label Label66 
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
               Left            =   4680
               TabIndex        =   131
               Top             =   1560
               Width           =   135
            End
            Begin VB.Label Label64 
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
               Left            =   720
               TabIndex        =   130
               Top             =   1080
               Width           =   135
            End
            Begin VB.Label Label18 
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
               TabIndex        =   89
               Top             =   240
               Width           =   1335
            End
            Begin VB.Label Label19 
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
               TabIndex        =   88
               Top             =   720
               Width           =   855
            End
            Begin VB.Label Label20 
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
               TabIndex        =   87
               Top             =   1200
               Width           =   615
            End
            Begin VB.Label Label21 
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
               TabIndex        =   86
               Top             =   1680
               Width           =   1095
            End
            Begin VB.Label Label39 
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
               TabIndex        =   85
               Top             =   2160
               Width           =   975
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00FF8080&
            Caption         =   "LOCAL ADDRESS "
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2505
            Left            =   7650
            TabIndex        =   73
            Top             =   2610
            Width           =   7425
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
               Left            =   1950
               TabIndex        =   29
               Top             =   270
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
               TabIndex        =   31
               Top             =   720
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
               TabIndex        =   33
               Top             =   1200
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
               TabIndex        =   35
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
               TabIndex        =   30
               Top             =   240
               Width           =   1695
            End
            Begin VB.TextBox landmark2 
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
               Top             =   720
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
               TabIndex        =   34
               Top             =   1200
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
               TabIndex        =   36
               Top             =   1680
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
               TabIndex        =   37
               Top             =   2160
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
               TabIndex        =   38
               Top             =   2160
               Width           =   1695
            End
            Begin VB.Label Label77 
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
               Left            =   4560
               TabIndex        =   141
               Top             =   2040
               Width           =   135
            End
            Begin VB.Label Label76 
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
               Left            =   4680
               TabIndex        =   140
               Top             =   1560
               Width           =   135
            End
            Begin VB.Label Label75 
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
               Left            =   4560
               TabIndex        =   139
               Top             =   120
               Width           =   135
            End
            Begin VB.Label Label74 
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
               Left            =   960
               TabIndex        =   138
               Top             =   2040
               Width           =   135
            End
            Begin VB.Label Label73 
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
               Left            =   1200
               TabIndex        =   137
               Top             =   1560
               Width           =   135
            End
            Begin VB.Label Label72 
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
               Left            =   720
               TabIndex        =   136
               Top             =   1080
               Width           =   135
            End
            Begin VB.Label Label27 
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
               TabIndex        =   83
               Top             =   270
               Width           =   1335
            End
            Begin VB.Label Label28 
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
               TabIndex        =   82
               Top             =   720
               Width           =   855
            End
            Begin VB.Label Label29 
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
               TabIndex        =   81
               Top             =   1200
               Width           =   615
            End
            Begin VB.Label Label30 
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
               TabIndex        =   80
               Top             =   1680
               Width           =   1095
            End
            Begin VB.Label Label31 
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
               TabIndex        =   79
               Top             =   240
               Width           =   735
            End
            Begin VB.Label Label32 
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
               TabIndex        =   78
               Top             =   720
               Width           =   1095
            End
            Begin VB.Label Label33 
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
               TabIndex        =   77
               Top             =   1200
               Width           =   1455
            End
            Begin VB.Label Label34 
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
               TabIndex        =   76
               Top             =   1680
               Width           =   855
            End
            Begin VB.Label Label41 
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
               TabIndex        =   75
               Top             =   2160
               Width           =   855
            End
            Begin VB.Label Label42 
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
               TabIndex        =   74
               Top             =   2160
               Width           =   735
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
            ItemData        =   "STUDENT DETAILS.frx":235746
            Left            =   5490
            List            =   "STUDENT DETAILS.frx":235750
            TabIndex        =   12
            Text            =   "SELECT"
            Top             =   1776
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
            ItemData        =   "STUDENT DETAILS.frx":235762
            Left            =   9630
            List            =   "STUDENT DETAILS.frx":235775
            TabIndex        =   13
            Text            =   "SELECT"
            Top             =   1740
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
            ItemData        =   "STUDENT DETAILS.frx":23579C
            Left            =   13380
            List            =   "STUDENT DETAILS.frx":2357A9
            TabIndex        =   14
            Text            =   "SELECT"
            Top             =   1779
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
            ItemData        =   "STUDENT DETAILS.frx":2357C9
            Left            =   5490
            List            =   "STUDENT DETAILS.frx":2357CB
            TabIndex        =   1
            Top             =   300
            Width           =   1695
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00FF8080&
            Caption         =   "PARENT'S INFORMATION"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   180
            TabIndex        =   67
            Top             =   5580
            Width           =   14925
            Begin VB.TextBox fatheremail 
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
               Left            =   11880
               TabIndex        =   43
               Top             =   330
               Width           =   1695
            End
            Begin VB.TextBox fathername 
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
               Left            =   1770
               TabIndex        =   40
               Top             =   360
               Width           =   1695
            End
            Begin VB.TextBox mothername 
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
               Left            =   1770
               TabIndex        =   44
               Top             =   780
               Width           =   1695
            End
            Begin VB.TextBox fatheroccp 
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
               Left            =   5160
               TabIndex        =   41
               Top             =   330
               Width           =   1695
            End
            Begin VB.TextBox motherocc 
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
               Left            =   5190
               TabIndex        =   45
               Top             =   810
               Width           =   1695
            End
            Begin VB.TextBox fathercontact 
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
               Left            =   8730
               TabIndex        =   42
               Top             =   330
               Width           =   1695
            End
            Begin VB.Label Label80 
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
               Left            =   1770
               TabIndex        =   144
               Top             =   1350
               Width           =   135
            End
            Begin VB.Label Label79 
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
               Left            =   1440
               TabIndex        =   143
               Top             =   720
               Width           =   135
            End
            Begin VB.Label Label78 
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
               Left            =   1440
               TabIndex        =   142
               Top             =   240
               Width           =   135
            End
            Begin VB.Label Label45 
               BackStyle       =   0  'Transparent
               Caption         =   "EMAIL-ID :"
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
               Left            =   10800
               TabIndex        =   101
               Top             =   360
               Width           =   1095
            End
            Begin VB.Label Label43 
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
               Left            =   120
               TabIndex        =   72
               Top             =   360
               Width           =   1335
            End
            Begin VB.Label Label44 
               BackStyle       =   0  'Transparent
               Caption         =   "MOTHER'S NAME :"
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
               TabIndex        =   71
               Top             =   840
               Width           =   1335
            End
            Begin VB.Label Label46 
               BackStyle       =   0  'Transparent
               Caption         =   "OCCUPATION :"
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
               Left            =   3720
               TabIndex        =   70
               Top             =   360
               Width           =   1215
            End
            Begin VB.Label Label47 
               BackStyle       =   0  'Transparent
               Caption         =   "OCCUPATION :"
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
               Left            =   3690
               TabIndex        =   69
               Top             =   810
               Width           =   1215
            End
            Begin VB.Label Label65 
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
               Left            =   7440
               TabIndex        =   68
               Top             =   360
               Width           =   1215
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
            ItemData        =   "STUDENT DETAILS.frx":2357CD
            Left            =   13290
            List            =   "STUDENT DETAILS.frx":2357D7
            TabIndex        =   18
            Text            =   "SELECT"
            Top             =   2250
            Width           =   1695
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00FF8080&
            Caption         =   "ACEDEMIC DETAILS"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2535
            Left            =   90
            TabIndex        =   62
            Top             =   6840
            Width           =   17415
            Begin VB.VScrollBar VScroll4 
               Height          =   1455
               Left            =   3600
               TabIndex        =   112
               Top             =   840
               Width           =   135
            End
            Begin VB.VScrollBar VScroll3 
               Height          =   1455
               Left            =   7440
               TabIndex        =   111
               Top             =   840
               Width           =   135
            End
            Begin VB.VScrollBar VScroll2 
               Height          =   1455
               Left            =   11400
               TabIndex        =   110
               Top             =   840
               Width           =   135
            End
            Begin VB.VScrollBar VScroll1 
               Height          =   1455
               Left            =   15240
               TabIndex        =   109
               Top             =   840
               Width           =   135
            End
            Begin VB.CommandButton remove 
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
               Left            =   15840
               Picture         =   "STUDENT DETAILS.frx":2357EA
               Style           =   1  'Graphical
               TabIndex        =   51
               Top             =   1440
               Width           =   1215
            End
            Begin VB.CommandButton add 
               Caption         =   "ADD"
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
               Left            =   15840
               Picture         =   "STUDENT DETAILS.frx":27E074
               Style           =   1  'Graphical
               TabIndex        =   50
               Top             =   360
               Width           =   1215
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
               ItemData        =   "STUDENT DETAILS.frx":2B60AC
               Left            =   11760
               List            =   "STUDENT DETAILS.frx":2B60AE
               TabIndex        =   108
               Top             =   840
               Width           =   3615
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
               ItemData        =   "STUDENT DETAILS.frx":2B60B0
               Left            =   7920
               List            =   "STUDENT DETAILS.frx":2B60B2
               TabIndex        =   107
               Top             =   840
               Width           =   3615
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
               ItemData        =   "STUDENT DETAILS.frx":2B60B4
               Left            =   3960
               List            =   "STUDENT DETAILS.frx":2B60B6
               TabIndex        =   106
               Top             =   840
               Width           =   3615
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
               ItemData        =   "STUDENT DETAILS.frx":2B60B8
               Left            =   120
               List            =   "STUDENT DETAILS.frx":2B60BA
               TabIndex        =   105
               Top             =   840
               Width           =   3615
            End
            Begin VB.TextBox board 
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
               Left            =   11760
               TabIndex        =   49
               Top             =   600
               Width           =   3615
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
               Height          =   285
               Left            =   7920
               TabIndex        =   48
               Top             =   600
               Width           =   3615
            End
            Begin VB.TextBox passyear 
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
               Left            =   3960
               TabIndex        =   47
               Top             =   600
               Width           =   3615
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
               Height          =   285
               Left            =   120
               TabIndex        =   46
               Top             =   600
               Width           =   3615
            End
            Begin VB.Line Line5 
               BorderColor     =   &H8000000B&
               X1              =   15480
               X2              =   15480
               Y1              =   120
               Y2              =   2520
            End
            Begin VB.Line Line4 
               BorderColor     =   &H8000000B&
               X1              =   11640
               X2              =   11640
               Y1              =   120
               Y2              =   2520
            End
            Begin VB.Line Line3 
               BorderColor     =   &H8000000B&
               X1              =   7800
               X2              =   7800
               Y1              =   120
               Y2              =   2520
            End
            Begin VB.Line Line2 
               BorderColor     =   &H8000000B&
               X1              =   3840
               X2              =   3840
               Y1              =   120
               Y2              =   2520
            End
            Begin VB.Line Line1 
               BorderColor     =   &H8000000B&
               X1              =   0
               X2              =   15480
               Y1              =   480
               Y2              =   480
            End
            Begin VB.Label Label38 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "YEAR OF PASSING"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   4560
               TabIndex        =   66
               Top             =   240
               Width           =   2175
            End
            Begin VB.Label Label50 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "BOARD/UNIVERSITY"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   12480
               TabIndex        =   65
               Top             =   240
               Width           =   2175
            End
            Begin VB.Label Label26 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "QULIFICATION"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   720
               TabIndex        =   64
               Top             =   240
               Width           =   2175
            End
            Begin VB.Label Label49 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "PERCENTAGE"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   8520
               TabIndex        =   63
               Top             =   240
               Width           =   2175
            End
         End
         Begin VB.Label Label35 
            BackStyle       =   0  'Transparent
            Caption         =   "EMAIL ID :"
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
            Left            =   4200
            TabIndex        =   156
            Top             =   2280
            Width           =   855
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "BATCH :"
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
            Left            =   4200
            TabIndex        =   155
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label14 
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
            Left            =   4200
            TabIndex        =   154
            Top             =   300
            Width           =   1155
         End
         Begin VB.Label Label11 
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
            Left            =   4200
            TabIndex        =   153
            Top             =   1800
            Width           =   855
         End
         Begin VB.Label Label7 
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
            Left            =   4200
            TabIndex        =   152
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label adm 
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
            Left            =   13380
            TabIndex        =   150
            Top             =   330
            Width           =   1695
         End
         Begin VB.Label Label62 
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
            Left            =   13170
            TabIndex        =   129
            Top             =   840
            Width           =   135
         End
         Begin VB.Label Label61 
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
            Left            =   13110
            TabIndex        =   128
            Top             =   360
            Width           =   135
         End
         Begin VB.Label Label60 
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
            Left            =   8970
            TabIndex        =   127
            Top             =   780
            Width           =   135
         End
         Begin VB.Label Label59 
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
            Left            =   8400
            TabIndex        =   126
            Top             =   300
            Width           =   135
         End
         Begin VB.Label Label58 
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
            Left            =   4920
            TabIndex        =   125
            Top             =   1680
            Width           =   135
         End
         Begin VB.Label Label57 
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
            Left            =   4920
            TabIndex        =   124
            Top             =   690
            Width           =   135
         End
         Begin VB.Label Label56 
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
            Left            =   4950
            TabIndex        =   123
            Top             =   240
            Width           =   135
         End
         Begin VB.Label Label55 
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
            Left            =   1320
            TabIndex        =   122
            Top             =   2040
            Width           =   135
         End
         Begin VB.Label Label54 
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
            Left            =   1560
            TabIndex        =   121
            Top             =   1560
            Width           =   135
         End
         Begin VB.Label Label53 
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
            Left            =   1320
            TabIndex        =   120
            Top             =   1200
            Width           =   135
         End
         Begin VB.Label Label52 
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
            Left            =   1050
            TabIndex        =   119
            Top             =   690
            Width           =   135
         End
         Begin VB.Label Label51 
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
            Left            =   1710
            TabIndex        =   118
            Top             =   270
            Width           =   135
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "BLOOD GROUP :"
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
            TabIndex        =   117
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label48 
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
            TabIndex        =   116
            Top             =   2280
            Width           =   1215
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "ROLL NUMBER :"
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
            Left            =   7770
            TabIndex        =   115
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label13 
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
            TabIndex        =   114
            Top             =   1800
            Width           =   855
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "SESSION :"
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
            Left            =   270
            TabIndex        =   113
            Top             =   750
            Width           =   855
         End
         Begin VB.Shape Shape3 
            BorderColor     =   &H8000000B&
            Height          =   1875
            Left            =   15240
            Top             =   2400
            Width           =   2295
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "ADMISSION NO :"
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
            TabIndex        =   99
            Top             =   390
            Width           =   1335
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "ADMISSION DATE :"
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
            Left            =   270
            TabIndex        =   98
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "REGISTRATION NO :"
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
            TabIndex        =   97
            Top             =   900
            Width           =   1575
         End
         Begin VB.Label Label6 
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
            Left            =   240
            TabIndex        =   96
            Top             =   1260
            Width           =   1095
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "SUR NAME :"
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
            Left            =   7800
            TabIndex        =   95
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "DATE OF BIRTH : "
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
            TabIndex        =   94
            Top             =   1710
            Width           =   1335
         End
         Begin VB.Label Label12 
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
            Left            =   7800
            TabIndex        =   93
            Top             =   1800
            Width           =   975
         End
         Begin VB.Label Label15 
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
            Height          =   255
            Left            =   7800
            TabIndex        =   92
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label36 
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
            Left            =   240
            TabIndex        =   91
            Top             =   2160
            Width           =   1095
         End
         Begin VB.Label Label37 
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
            Left            =   7800
            TabIndex        =   90
            Top             =   2280
            Width           =   1095
         End
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H8000000B&
         BorderWidth     =   2
         Height          =   7575
         Left            =   17880
         Top             =   600
         Width           =   1575
      End
   End
   Begin VB.Label Label40 
      Caption         =   "Label40"
      Height          =   255
      Left            =   13200
      TabIndex        =   158
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "STUDENT INFORMATION"
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
      Height          =   735
      Left            =   6840
      TabIndex        =   59
      Top             =   0
      Width           =   6375
   End
End
Attribute VB_Name = "frmSTUDENT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub aadhar_LostFocus()
If (Len(aadhar) < 12) Then
A = MsgBox("AADHAR NUMBER MUST BE OF 12 DIGIT", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
aadhar.Text = ""
'aadhar.SetFocus
End If
End If
End Sub

Private Sub batch_lostfocus()
On Error GoTo label
batch.Locked = True
sql = "select count(stud_adm) from studentdetail_master where course='" + COURSE.Text + "' and stud_year='" + year.Text + "'"
Set r = c.Execute(sql)
Dim i As Integer
i = r.Fields(0)
If (i < 1) Then
batch.Text = 1
rollnum.Text = 1
Else
sql = "select strength from batchdetail_master where course='" + COURSE.Text + "' and year='" + year.Text + "' and batno='" + batch.Text + "'"
Set r = c.Execute(sql)
stren = r.Fields(0)
sql = "select max(stud_roll) from studentdetail_master where course='" + COURSE.Text + "' and stud_year='" + year.Text + "' and stud_batch=" + batch.Text + ""
Set r = c.Execute(sql)
roll = r.Fields(0)
If (roll = stren) Then
batch.Text = batch.Text + 1
rollnum.Text = 1
Else
sql = "select max(stud_roll) from studentdetail_master where course='" + COURSE.Text + "'and stud_year='" + year.Text + "' and stud_batch=" + batch.Text + ""
Set r = c.Execute(sql)
roll = r.Fields(0)
roll = roll + 1
rollnum.Text = roll
Exit Sub
label:
regno.SetFocus
End If
End If
End Sub
Private Sub Check1_Click()
If Check1.Value = 1 Then
hno2.Text = hno1.Text
street2.Text = street1.Text
village2.Text = village1.Text
landmark2.Text = landmark1.Text
city2.Text = city1.Text
policestation2.Text = policestation1.Text
postoffice2.Text = postoffice1.Text
pincode2.Text = pincode1.Text
district2.Text = district1.Text
state.Text = state1.Text
Else
hno2.Text = ""
street2.Text = ""
village2.Text = ""
landmark2.Text = ""
city2.Text = ""
policestation2.Text = ""
postoffice2.Text = ""
pincode2.Text = ""
district2.Text = ""
state.Text = ""
End If
End Sub
Private Sub Command6_Click()
cd1.Filter = "picture file |*.jpg"
cd1.ShowOpen
If cd1.FileName <> " " Then
Image1.Picture = LoadPicture(cd1.FileName)
End If
Label40.Caption = cd1.FileName
End Sub
Private Sub aadhar_keypress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
aadhar.Locked = False
Else
aadhar.Locked = True
End If
End Sub
Private Sub add_Click()
If (quali.Text = "" Or passyear.Text = "" Or percentage.Text = "" Or board.Text = "") Then
MsgBox "please enter valid data"
Else
List1.AddItem quali.Text
quali.Text = ""
List2.AddItem passyear.Text
passyear.Text = ""
List3.AddItem percentage.Text
percentage.Text = ""
List4.AddItem board.Text
board.Text = ""
End If
End Sub
Private Sub add_gotfocus()
If (quali.Text = "") Then
add.Enabled = False
If (passyear.Text = "") Then
add.Enabled = False
If (percentage.Text = "") Then
add.Enabled = False
If (board.Text = "") Then
add.Enabled = False
End If
End If
End If
End If
End Sub
Private Sub batch_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
batch.Locked = False
Else
batch.Locked = True
End If
End Sub
Private Sub bloodgrp_keypress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 87 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 43 Or KeyAscii = 45) Then
bloodgrp.Locked = False
Else
bloodgrp.Locked = True
End If
End Sub
Private Sub city1_keypress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 87 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32) Then
city1.Locked = False
Else
city1.Locked = True
End If
End Sub
Private Sub city2_keypress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 87 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32) Then
city2.Locked = False
Else
city2.Locked = True
End If
End Sub

Private Sub contactno_keypress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
contactno.Locked = False
Else
contactno.Locked = True
End If
End Sub
Private Sub contactno_LostFocus()
If (Len(contactno) < 10) Then
A = MsgBox("CONTACT NUMBER MUST BE OF 10 DIGIT", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
contactno.Text = ""
'contactno.SetFocus
End If
End If
End Sub
Private Sub fathercontact_LostFocus()
If (Len(fathercontact) < 10) Then
A = MsgBox(" CONTACT NUMBER MUST BE OF 10 DIGIT", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
fathercontact.Text = ""
fathercontact.SetFocus
End If
End If
End Sub

Private Sub passyear_LostFocus()
If Len(passyear.Text) < 4 Then
MsgBox "INVALID YEAR", vbInformation, "INFORMATION"
passyear.SetFocus
End If
End Sub

Private Sub studadm_Click()
If (COURSE.Text = "") Then
studadm.Text = ""
End If
End Sub
Private Sub year_lostfocus()
Module1.conn
If (COURSE.Text = COURSE.Text) Then
sql = "select count(stud_adm) from studentdetail_master where course='" + COURSE.Text + "'"
Set r = c.Execute(sql)
Dim i As Integer
i = r.Fields(0)
If (i < 1) Then
A = admdate.Value
b = Right(A, 2)
d = COURSE.Text & b & ("0")
studadm.Text = (d) & (1)
Else
sql = "select max(stud_adm) from studentdetail_master where course='" + COURSE.Text + "'"
Set r = c.Execute(sql)
Dim j As String
j = r.Fields(0)
MsgBox j
A = Right(j, 1)
A = A + 1
MsgBox A
d = Left(j, 6)
MsgBox d
studadm.Text = d & A
End If
End If
End Sub

Private Sub delete_Click()
If (adm.Caption = "" Or stadm.Text = "" Or session.Text = "" Or year.Text = "" Or batch.Text = "" Or rollnum.Text = "" Or firstname.Text = "" Or stdob.Text = "" Or gender.Text = "" Or category.Text = "" Or religion.Text = "" Or contactno.Text = "" Or nationality.Text = "" Or district1.Text = "" Or city1.Text = "" Or state1.Text = "" Or pincode1.Text = "" Or postoffice1.Text = "" Or city2.Text = "" Or postoffice2.Text = "" Or state.Text = "" Or pincode2.Text = "" Or district2.Text = "" Or fatheroccp.Text = "" Or fathercontact.Text = "" Or fathername.Text = "") Then
MsgBox "ALL FIELDS ARE MANDATORY"
Else
Module1.conn
A = MsgBox("WANT TO DELETE RECORD", vbQuestion + vbYesNo, "DELETION")
If (A = vbYes) Then
sql = "delete from StudQuali_Master where stud_adm=" + adm.Caption + ""
Set r = c.Execute(sql)
sql = "delete from StudentDetail_Master where stud_adm=" + adm.Caption + ""
Set r = c.Execute(sql)
MsgBox "record delete"
Else
If (A = vbNo) Then
MsgBox "RECORD NOT DELETE"
End If
End If
End If
End Sub

Private Sub district1_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 87 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32) Then
district1.Locked = False
Else
district1.Locked = True
End If
End Sub

Private Sub district2_keypress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 87 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32) Then
district2.Locked = False
Else
district2.Locked = True
End If
End Sub

Private Sub dob_LostFocus()
Dim age As Long
age = DateDiff("yyyy", dob.Value, Date)
If (age <= 18 Or age >= 60) Then
MsgBox ("YOU ARE NOT ELLIGIBLE FOR THIS COURSE")
dob.SetFocus
End If
End Sub

Private Sub fathercontact_keypress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
fathercontact.Locked = False
Else
fathercontact.Locked = True
End If
End Sub

Private Sub fathername_keypress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 87 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32) Then
fathername.Locked = False
Else
fathername.Locked = True
End If
End Sub

Private Sub fatheroccp_keypress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 87 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32) Then
fatheroccp.Locked = False
Else
fatheroccp.Locked = True
End If
End Sub

Private Sub firstname_keypress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 87 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32) Then
firstname.Locked = False
Else
firstname.Locked = True
End If
End Sub
Private Sub Form_Load()
studadm.Locked = True
stdob.Visible = False
stadm.Visible = False
studadm.MaxLength = 10
regno.MaxLength = 10
session.MaxLength = 9
'batch.MaxLength = 1
rollnum.MaxLength = 2
firstname.MaxLength = 20
middlename.MaxLength = 20
surname.MaxLength = 20
bloodgrp.MaxLength = 4
contactno.MaxLength = 10
hno1.MaxLength = 3
pincode1.MaxLength = 6
pincode2.MaxLength = 6
aadhar.MaxLength = 12
fathercontact.MaxLength = 10
emailid.MaxLength = 30
street1.MaxLength = 20
village1.MaxLength = 20
city1.MaxLength = 20
postoffice1.MaxLength = 20
district1.MaxLength = 20
state1.MaxLength = 20
policestation1.MaxLength = 20
landmark1.MaxLength = 20
hno2.MaxLength = 3
pincode2.MaxLength = 6
street2.MaxLength = 20
village2.MaxLength = 20
city2.MaxLength = 20
postoffice2.MaxLength = 20
district2.MaxLength = 20
state.MaxLength = 20
policestation2.MaxLength = 20
landmark2.MaxLength = 20
fathername.MaxLength = 20
fatheroccp.MaxLength = 20
mothername.MaxLength = 30
motherocc.MaxLength = 30
fatheremail.MaxLength = 30
passyear.MaxLength = 4
board.MaxLength = 50
quali.MaxLength = 50
percentage.MaxLength = 5
Module1.conn
sql = "select code from CourseDetail_Master"
Set r = c.Execute(sql)
While (r.EOF = False)
COURSE.AddItem r.Fields("code")
r.MoveNext
Wend
year.AddItem "I"
year.AddItem "II"
year.AddItem "III"
End Sub

Private Sub landmark1_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 87 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32) Then
landmark1.Locked = False
Else
landmark1.Locked = True
End If
End Sub

Private Sub landmark2_keypress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 87 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32) Then
landmark2.Locked = False
Else
landmark2.Locked = True
End If
End Sub

Private Sub middlename_keypress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 87 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32) Then
middlename.Locked = False
Else
middlename.Locked = True
End If
End Sub

Private Sub mothername_keypress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 87 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32) Then
mothername.Locked = False
Else
mothername.Locked = True
End If
End Sub

Private Sub motherocc_keypress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 87 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32) Then
motherocc.Locked = False
Else
motherocc.Locked = True
End If
End Sub
Private Sub pincode1_keypress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
pincode1.Locked = False
Else
pincode1.Locked = True
End If
End Sub

Private Sub pincode2_keypress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
pincode2.Locked = False
Else
pincode2.Locked = True
End If
End Sub

Private Sub policestation1_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 87 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32) Then
policestation1.Locked = False
Else
policestation1.Locked = True
End If
End Sub

Private Sub policestation2_keypress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 87 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32) Then
policestation2.Locked = False
Else
policestation2.Locked = True
End If
End Sub
Private Sub postoffice1_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 87 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32) Then
postoffice1.Locked = False
Else
postoffice1.Locked = True
End If
End Sub

Private Sub postoffice2_keypress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 87 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32) Then
postoffice2.Locked = False
Else
postoffice2.Locked = True
End If
End Sub

Private Sub refresh_Click()
dob.Visible = True
stadm.Visible = False
adm.Visible = False
studadm.Visible = True
admdate.Visible = True
stdob.Visible = False
regno.Text = ""
session.Text = ""
batch.Text = ""
rollnum.Text = ""
firstname.Text = ""
middlename.Text = ""
surname.Text = ""
bloodgrp.Text = ""
contactno.Text = ""
emailid.Text = ""
aadhar.Text = ""
hno1.Text = ""
street1.Text = ""
village1.Text = ""
landmark1.Text = ""
city1.Text = ""
policestation1.Text = ""
postoffice1.Text = ""
pincode1.Text = ""
district1.Text = ""
state1.Text = ""
hno2.Text = ""
street2.Text = ""
village2.Text = ""
landmark2.Text = ""
city2.Text = ""
policestation2.Text = ""
postoffice2.Text = ""
pincode2.Text = ""
district2.Text = ""
state.Text = ""
fathername.Text = ""
fatheroccp.Text = ""
fathercontact.Text = ""
fatheremail.Text = ""
mothername.Text = ""
motherocc.Text = ""
List1.clear
List2.clear
List3.clear
List4.clear
End Sub

Private Sub regno_keypress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
regno.Locked = False
Else
regno.Locked = True
End If
End Sub
Private Sub remove_Click()
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

Private Sub rollnum_click()
bat = batch.Text
Module1.conn
sql = "select max(batno) from batchdetail_master where course= '" + COURSE.Text + "' and year = '" + year.Text + "'"
Set r = c.Execute(sql)
batch = r.Fields(0)
If (bat > batch) Then
A = MsgBox("ALL BATCHES ARE FULL", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
batch.Text = ""
rollnum.Text = ""
session.Text = ""
studadm.Text = ""
admdate.SetFocus
End If
End If

End Sub
Private Sub rollnum_keypress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
rollnum.Locked = False
Else
rollnum.Locked = True
End If
End Sub
Private Sub save_Click()
'On Error GoTo label
If (studadm.Text = "" Or session.Text = "" Or year.Text = "" Or batch.Text = "" Or rollnum.Text = "" Or firstname.Text = "" Or gender.Text = "" Or category.Text = "" Or religion.Text = "" Or contactno.Text = "" Or nationality.Text = "" Or district1.Text = "" Or city1.Text = "" Or state1.Text = "" Or pincode1.Text = "" Or postoffice1.Text = "" Or city2.Text = "" Or postoffice2.Text = "" Or state.Text = "" Or pincode2.Text = "" Or district2.Text = "" Or fathercontact.Text = "" Or fathername.Text = "" Or fatheroccp.Text = "" Or List1.ListCount = 0 Or List2.ListCount = 0 Or List3.ListCount = 0 Or List4.ListCount = 0 Or Image1.Picture = Nothing) Then
MsgBox "ALL FIELDS ARE MANDATORY"
Else
Module1.conn
A = MsgBox("CLICK OK TO SAVE RECORD", vbQuestion + vbOKCancel, "SAVE")
If (A = vbOK) Then
sql = "insert into StudentDetail_Master values('" & Format(admdate.Value, "dd/MMM/yyyy") & "','" & regno.Text & "','" & session.Text & "','" & year.Text & "'," & batch.Text & "," & rollnum.Text & ",'" & firstname.Text & "','" & middlename.Text & "','" & surname.Text & "','" & bloodgrp.Text & "','" & Format(dob.Value, "dd/MMM/yyyy") & "','" & gender.Text & "','" & category.Text & "','" & religion.Text & "'," & contactno.Text & ",'" & emailid.Text & "','" & aadhar.Text & "','" + nationality.Text + "','" + hno1.Text + "','" + street1.Text + "','" + village1.Text + "','" + landmark1.Text + "','" + city1.Text + "','" + policestation1.Text + "','" + postoffice1.Text + "'," + pincode1.Text + ",'" + district1.Text + "','" + state1.Text + "','" + hno2.Text + "','" + street2.Text + "','" + village2.Text + "','" + landmark2.Text + "','" + city2.Text + "','" + policestation2.Text + "','" + postoffice2.Text + "'," + pincode2.Text + ",'" + district2.Text + "','" + state.Text + "'," _
& "'" + fathername.Text + "','" + fatheroccp.Text + "'," + fathercontact.Text + ",'" + fatheremail.Text + "','" + mothername.Text + "','" + motherocc.Text + "','" & studadm.Text & "','" & COURSE.Text & "','" + Label40.Caption + "')"
Set r = c.Execute(sql)
i = 0
Counter = 0
While (i <= List1.ListCount - 1 And Counter <= Counter)
sql = "insert into StudQuali_Master values('" + List1.List(i) + "'," & List2.List(i) & "," & List3.List(i) & ",'" + List4.List(i) + "'," & Counter & ",'" & studadm.Text & "')"
Set r = c.Execute(sql)
i = i + 1
Counter = Counter + 1
Wend
MsgBox "RECORD SAVE"
regno.Text = ""
session.Text = ""
batch.Text = ""
rollnum.Text = ""
firstname.Text = ""
middlename.Text = ""
surname.Text = ""
bloodgrp.Text = ""
contactno.Text = ""
emailid.Text = ""
aadhar.Text = ""
hno1.Text = ""
street1.Text = ""
village1.Text = ""
landmark1.Text = ""
city1.Text = ""
policestation1.Text = ""
postoffice1.Text = ""
pincode1.Text = ""
district1.Text = ""
state1.Text = ""
hno2.Text = ""
street2.Text = ""
village2.Text = ""
landmark2.Text = ""
city2.Text = ""
policestation2.Text = ""
postoffice2.Text = ""
pincode2.Text = ""
district2.Text = ""
state.Text = ""
fathername.Text = ""
fatheroccp.Text = ""
fathercontact.Text = ""
fatheremail.Text = ""
mothername.Text = ""
motherocc.Text = ""
Check1.Value = 0
List1.clear
List2.clear
List3.clear
List4.clear
Image1.refresh
sql = "select stud_adm"
'Exit Sub
'label:
'MsgBox "DATA ALREADY EXIST"
Else
If (A = vbCancel) Then
MsgBox "record not save"
End If
End If
End If
End Sub
'Private Sub Command6_Click()
'cd1.Filter = "picture file |*.jpg"
'cd1.ShowOpen
'If cd1.FileName <> " " Then
'Image1.Picture = LoadPicture(cd1.FileName)
'End If
'End Sub
Private Sub percentage_Change()
add.Enabled = Len(percentage.Text) > 0
End Sub
Private Sub percentage_lostfocus()
If (percentage.Text <= 55) Then
A = MsgBox("YOU ARE NOT ELIGIBLE FOR THIS COURSE", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
percentage.SetFocus
End If
End If
If (percentage.Text > 100) Then
A = MsgBox("PLEASE ENTER VALID PERCENTAGE", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
percentage.SetFocus
End If
End If
End Sub
Private Sub session_keypress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 45 Or KeyAscii = 8) Then
session.Locked = False
Else
session.Locked = True
End If
End Sub

Private Sub session_gotfocus()
A = Right(admdate.Value, 4)
b = A + 3
d = A & "-" & b
session.Text = d
End Sub
Private Sub admdate_lostfocus()
Dim d As Date
d = admdate.Value
If (d > Date) Then
A = MsgBox("INVALID DATE", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOKOnly) Then
admdate.SetFocus
End If
End If
End Sub
Private Sub session_lostfocus()
Module1.conn
sql = "select max(stud_batch) from studentdetail_master where course='" + COURSE.Text + "' and stud_year='" + year.Text + "'"
Set r = c.Execute(sql)
bat = r.Fields(0)
If (IsNull(bat)) Then
bat = 1
batch.Text = bat
Else
batch.Text = bat
End If
End Sub

Private Sub state_keypress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 87 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32) Then
state.Locked = False
Else
state.Locked = True
End If
End Sub

Private Sub state1_keypress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 87 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32) Then
state1.Locked = False
Else
state1.Locked = True
End If
End Sub
Private Sub street1_keypress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 87 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32) Then
street1.Locked = False
Else
street1.Locked = True
End If
End Sub

Private Sub street2_keypress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 87 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32) Then
street2.Locked = False
Else
street2.Locked = True
End If
End Sub
Private Sub surname_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 87 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32) Then
surname.Locked = False
Else
surname.Locked = True
End If
End Sub

Private Sub update_Click()
If (adm.Caption = "" Or stadm.Text = "" Or session.Text = "" Or year.Text = "" Or batch.Text = "" Or rollnum.Text = "" Or firstname.Text = "" Or stdob.Text = "" Or gender.Text = "" Or category.Text = "" Or religion.Text = "" Or contactno.Text = "" Or nationality.Text = "" Or district1.Text = "" Or city1.Text = "" Or state1.Text = "" Or pincode1.Text = "" Or postoffice1.Text = "" Or city2.Text = "" Or postoffice2.Text = "" Or state.Text = "" Or pincode2.Text = "" Or district2.Text = "" Or fatheroccp.Text = "" Or fathercontact.Text = "" Or fathername.Text = "") Then
MsgBox "ALL FIELDS ARE MANDATORY"
Else
Module1.conn
A = MsgBox("WANT TO UPDATE RECORD", vbQuestion + vbYesNo, "UPDATE")
If (A = vbYes) Then
sql = "update StudentDetail_Master set stud_admdate='" + Format(stadm.Text, "dd/MMM/yyyy") + "',stud_sess='" + session.Text + "',stud_fname='" + firstname.Text + "',stud_bgrp='" + bloodgrp.Text + "',stud_dob='" + Format(stdob.Text, "dd/MMM/yyyy") + "',stud_gen='" + gender.Text + "',stud_categ='" + category.Text + "',stud_relig='" + religion.Text + "',stud_contact=" + contactno.Text + ",stud_nation='" + nationality.Text + "',per_city='" + city1.Text + "',per_dist='" + district1.Text + "',per_po='" + postoffice1.Text + "',per_pincode=" + pincode1.Text + ",per_state='" + state1.Text + "', " _
& "loc_city='" + city2.Text + "',loc_po='" + postoffice2.Text + "',loc_pin=" + pincode2.Text + ",loc_state='" + state.Text + "',loc_dist='" + district2.Text + "',stud_fthnm='" + fathername.Text + "',stud_fthnmocc='" + fatheroccp.Text + "',stud_fthcontact=" + fathercontact.Text + ""
MsgBox sql
Set r = c.Execute(sql)
i = List1.ListIndex
k = List2.ListIndex
m = List3.ListIndex
n = List4.ListIndex
If (List1.ListIndex = -1) Then
quali.Text = ""
Else
List1.RemoveItem List1.ListIndex
List1.AddItem quali.Text
j = i
While (j <= List1.ListCount - 1)
sql = "update StudQuali_Master set quali='" & List1.List(j) & "' where sno=" & i & ""
Set r = c.Execute(sql)
j = j + 1
Wend
End If
If (List2.ListIndex = -1) Then
passyear.Text = ""
Else
List2.RemoveItem List2.ListIndex
List2.AddItem passyear.Text
l = k
While (l <= List2.ListCount - 1)
sql = "update StudQuali_Master set pass_year=" & List2.List(k) & " where sno=" & k & ""
Set r = c.Execute(sql)
l = l + 1
Wend
End If
If (List3.ListIndex = -1) Then
percentage.Text = ""
Else
List3.RemoveItem List3.ListIndex
List3.AddItem percentage.Text
P = m
While (P <= List3.ListCount - 1)
sql = "update StudQuali_Master set percentage=" & List3.List(P) & " where sno=" & m & ""
Set r = c.Execute(sql)
P = P + 1
Wend
End If
If (List4.ListIndex = -1) Then
board.Text = ""
Else
List4.RemoveItem List4.ListIndex
List4.AddItem board.Text
t = n
While (t <= List4.ListCount - 1)
sql = "update StudQuali_Master set board='" & List4.List(t) & "' where sno=" & n & ""
Set r = c.Execute(sql)
t = t + 1
Wend
End If
'List1.AddItem quali.Text
'List2.AddItem passyear.Text
'List3.AddItem percentage.Text
'List4.AddItem board.Text
'j = i
'l = k
'P = m
't = n
'While (j <= List1.ListCount - 1)
'sql = "update StudQuali_Master set quali='" & List1.List(j) & "' where sno=" & i & ""
'Set r = c.Execute(sql)
'j = j + 1
'Wend
'While (l <= List2.ListCount - 1)
'sql = "update StudQuali_Master set pass_year=" & List2.List(k) & " where sno=" & k & ""
'Set r = c.Execute(sql)
'l = l + 1
'Wend
'While (P <= List3.ListCount - 1)
'sql = "update StudQuali_Master set percentage=" & List3.List(P) & " where sno=" & m & ""
'Set r = c.Execute(sql)
'P = P + 1
'Wend
'While (t <= List4.ListCount - 1)
'sql = "update StudQuali_Master set board='" & List4.List(t) & "' where sno=" & n & ""
'Set r = c.Execute(sql)
't = t + 1
'Wend
MsgBox "RECORD UPDATE"
Else
If (A = vbNo) Then
MsgBox "RECORD NOT UPDATE"
End If
End If
End If
End Sub

Private Sub view_Click()
List1.clear
List2.clear
List3.clear
List4.clear
'On Error GoTo label
valu = InputBox("ENTER STUDENT ADMISSION NUMBER", "ADMISSION NUMBER")
If StrPtr(valu) = 0 Then
MsgBox "SEARCHING IS CANCELLED", vbInformation + vbOKOnly, "INFORMATION"
Else
adm = valu
studadm.Visible = False
dob.Visible = False
stdob.Visible = True
admdate.Visible = False
stadm.Visible = True
Module1.conn
sql = "select *from StudentDetail_Master where stud_adm='" + adm.Caption + "'"
Set r = c.Execute(sql)
stadm.Text = r.Fields(0)
If (IsNull(r.Fields(1))) Then
regno.Text = ""
Else
regno.Text = r.Fields(1)
End If
session.Text = r.Fields(2)
year.Text = r.Fields(3)
batch.Text = r.Fields(4)
rollnum.Text = r.Fields(5)
firstname.Text = r.Fields(6)
If (IsNull(r.Fields(7))) Then
middlename.Text = ""
Else
middlename.Text = r.Fields(7)
End If
If (IsNull(r.Fields(8))) Then
surname.Text = ""
Else
surname.Text = r.Fields(8)
End If
If (IsNull(r.Fields(9))) Then
bloodgrp.Text = ""
Else
bloodgrp.Text = r.Fields(9)
End If
stdob.Text = r.Fields(10)
gender.Text = r.Fields(11)
category.Text = r.Fields(12)
religion.Text = r.Fields(13)
contactno.Text = r.Fields(14)
If (IsNull(r.Fields(15))) Then
emailid.Text = ""
Else
emailid.Text = r.Fields(15)
End If
If (IsNull(r.Fields(16))) Then
aadhar.Text = ""
Else
aadhar.Text = r.Fields(16)
End If
nationality.Text = r.Fields(17)
If (IsNull(r.Fields(18))) Then
hno1.Text = ""
Else
hno1.Text = r.Fields(18)
End If
If (IsNull(r.Fields(19))) Then
street1.Text = ""
Else
street1.Text = r.Fields(19)
End If
If (IsNull(r.Fields(20))) Then
village1.Text = ""
Else
village1.Text = r.Fields(20)
End If
If (IsNull(r.Fields(21))) Then
landmark1.Text = ""
Else
landmark1.Text = r.Fields(21)
End If
city1.Text = r.Fields(22)
If (IsNull(r.Fields(23))) Then
policestation1.Text = ""
Else
policestation1.Text = r.Fields(23)
End If
postoffice1.Text = r.Fields(24)
pincode1.Text = r.Fields(25)
district1.Text = r.Fields(26)
state1.Text = r.Fields(27)
If (IsNull(r.Fields(28))) Then
hno2.Text = ""
Else
hno2.Text = r.Fields(28)
End If
If (IsNull(r.Fields(29))) Then
street2.Text = ""
Else
street2.Text = r.Fields(29)
End If
If (IsNull(r.Fields(30))) Then
village2.Text = ""
Else
village2.Text = r.Fields(30)
End If
If (IsNull(r.Fields(31))) Then
landmark2.Text = ""
Else
landmark2.Text = r.Fields(31)
End If
city2.Text = r.Fields(32)
If (IsNull(r.Fields(33))) Then
policestation2.Text = ""
Else
policestation2.Text = r.Fields(33)
End If
postoffice2.Text = r.Fields(34)
pincode2.Text = r.Fields(35)
district2.Text = r.Fields(36)
state.Text = r.Fields(37)
fathername.Text = r.Fields(38)
fatheroccp.Text = r.Fields(39)
fathercontact.Text = r.Fields(40)
If (IsNull(r.Fields(41))) Then
fatheremail.Text = ""
Else
fatheremail.Text = r.Fields(41)
End If
If (IsNull(r.Fields(42))) Then
mothername.Text = ""
Else
mothername.Text = r.Fields(42)
End If
If (IsNull(r.Fields(43))) Then
motherocc.Text = ""
Else
motherocc.Text = r.Fields(43)
End If
COURSE.Text = r.Fields(45)
Label40.Caption = r.Fields(46)
Image1.Picture = LoadPicture(Label40.Caption)
sql = "select *from StudQuali_Master where stud_adm='" + adm.Caption + "'"
Set r = c.Execute(sql)
While (r.EOF = False)
List1.AddItem r.Fields("Quali")
List2.AddItem r.Fields("Pass_Year")
List3.AddItem r.Fields("percentage")
List4.AddItem r.Fields("board")
r.MoveNext
Wend
'Exit Sub
'label:
'MsgBox "RECORD NOT FOUND"
adm.Visible = True
'studadm.Visible = True
End If
End Sub
Private Sub quali_change()
add.Enabled = Len(quali.Text) > 0
End Sub
Private Sub board_Change()
add.Enabled = Len(board.Text) > 0
End Sub
Private Sub passyear_Change()
add.Enabled = Len(passyear.Text) > 0
End Sub

Private Sub village1_keypress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 87 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32) Then
village1.Locked = False
Else
village1.Locked = True
End If
End Sub
Private Sub village2_keypress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 87 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32) Then
village2.Locked = False
Else
village2.Locked = True
End If
End Sub
Private Sub quali_keypress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 87 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32) Then
quali.Locked = False
Else
quali.Locked = True
End If
End Sub
Private Sub passyear_keypress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
passyear.Locked = False
Else
passyear.Locked = True
End If
End Sub
Private Sub percentage_keypress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Or KeyAscii = 46) Then
percentage.Locked = False
Else
percentage.Locked = True
End If
End Sub
Private Sub board_keypress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 87 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32) Then
board.Locked = False
Else
board.Locked = True
End If
End Sub
Private Sub List1_Click()
quali.Text = List1.List(List1.ListIndex)
End Sub
Private Sub List2_Click()
passyear.Text = List2.List(List2.ListIndex)
End Sub
Private Sub List3_Click()
percentage.Text = List3.List(List3.ListIndex)
End Sub
Private Sub List4_Click()
board.Text = List4.List(List4.ListIndex)
End Sub


