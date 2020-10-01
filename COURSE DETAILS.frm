VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmCOURSE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "COURSE"
   ClientHeight    =   10215
   ClientLeft      =   870
   ClientTop       =   870
   ClientWidth     =   18915
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   Picture         =   "COURSE DETAILS.frx":0000
   ScaleHeight     =   10215
   ScaleWidth      =   18915
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   4410
      Top             =   90
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   661
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
      RecordSource    =   "select *from SubjectDetail_Master "
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Cambria"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   345
      Left            =   990
      Top             =   120
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   609
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
      RecordSource    =   "select *from CourseDetail_Master"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Cambria"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FF8080&
      Height          =   4335
      Left            =   150
      TabIndex        =   19
      Top             =   4890
      Width           =   18255
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "COURSE DETAILS.frx":1E9608
         Height          =   3945
         Left            =   13380
         TabIndex        =   66
         Top             =   300
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   6959
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16393
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16393
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H80000002&
         Height          =   1155
         Left            =   3360
         TabIndex        =   59
         Top             =   2820
         Width           =   8025
         Begin VB.CommandButton sub_update 
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
            Left            =   4920
            Picture         =   "COURSE DETAILS.frx":1E961D
            Style           =   1  'Graphical
            TabIndex        =   64
            Top             =   150
            Width           =   1095
         End
         Begin VB.CommandButton sub_delete 
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
            Left            =   3420
            Picture         =   "COURSE DETAILS.frx":1E9DAD
            Style           =   1  'Graphical
            TabIndex        =   63
            Top             =   150
            Width           =   1095
         End
         Begin VB.CommandButton sub_view 
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
            Left            =   1935
            Picture         =   "COURSE DETAILS.frx":1EA319
            Style           =   1  'Graphical
            TabIndex        =   62
            Top             =   150
            Width           =   1095
         End
         Begin VB.CommandButton sub_refresh 
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
            Left            =   6405
            Picture         =   "COURSE DETAILS.frx":1EA6D0
            Style           =   1  'Graphical
            TabIndex        =   61
            Top             =   150
            Width           =   1095
         End
         Begin VB.OptionButton suboption 
            Caption         =   "FOR MODICATION"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   630
            TabIndex        =   60
            Top             =   450
            Width           =   1215
         End
      End
      Begin VB.CommandButton sub_save 
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
         Left            =   1590
         Picture         =   "COURSE DETAILS.frx":1EAD74
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   2910
         Width           =   1095
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FF8080&
         Caption         =   "SUBJECT "
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
         TabIndex        =   20
         Top             =   120
         Width           =   13125
         Begin VB.TextBox subid 
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
            Height          =   495
            Left            =   6480
            TabIndex        =   7
            Top             =   240
            Width           =   1935
         End
         Begin VB.TextBox fullmrks 
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
            Height          =   495
            Left            =   2208
            TabIndex        =   9
            Top             =   960
            Width           =   1935
         End
         Begin VB.TextBox passmrks 
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
            Height          =   495
            Left            =   2208
            TabIndex        =   12
            Top             =   1680
            Width           =   1935
         End
         Begin VB.ComboBox paper 
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
            ItemData        =   "COURSE DETAILS.frx":1EB54E
            Left            =   11040
            List            =   "COURSE DETAILS.frx":1EB558
            TabIndex        =   8
            Text            =   "SELECT"
            Top             =   240
            Width           =   1935
         End
         Begin VB.TextBox practical 
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
            Height          =   495
            Left            =   11040
            TabIndex        =   11
            Top             =   960
            Width           =   1935
         End
         Begin VB.TextBox theory 
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
            Height          =   495
            Left            =   6504
            TabIndex        =   10
            Top             =   960
            Width           =   1935
         End
         Begin VB.TextBox subnm 
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
            Height          =   495
            Left            =   2190
            TabIndex        =   6
            Top             =   300
            Width           =   1935
         End
         Begin VB.Label Label24 
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
            Left            =   10080
            TabIndex        =   36
            Top             =   960
            Width           =   135
         End
         Begin VB.Label Label23 
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
            Left            =   9960
            TabIndex        =   35
            Top             =   240
            Width           =   135
         End
         Begin VB.Label Label22 
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
            Left            =   5760
            TabIndex        =   34
            Top             =   960
            Width           =   135
         End
         Begin VB.Label Label21 
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
            Left            =   6000
            TabIndex        =   33
            Top             =   240
            Width           =   135
         End
         Begin VB.Label Label20 
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
            TabIndex        =   32
            Top             =   1680
            Width           =   135
         End
         Begin VB.Label Label19 
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
            TabIndex        =   31
            Top             =   960
            Width           =   135
         End
         Begin VB.Label Label18 
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
            TabIndex        =   30
            Top             =   240
            Width           =   135
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "THEORY :"
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
            Left            =   4896
            TabIndex        =   27
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "SUBJECT  ID :"
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
            Left            =   4740
            TabIndex        =   26
            Top             =   330
            Width           =   1215
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "SUBJECT NAME :"
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
            TabIndex        =   25
            Top             =   390
            Width           =   1215
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "PRACTICAL :"
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
            Left            =   9192
            TabIndex        =   24
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "FULL MARKS :"
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
            Left            =   240
            TabIndex        =   23
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "PAPER :"
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
            Left            =   9240
            TabIndex        =   22
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "PASS MARKS :"
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
            Left            =   240
            TabIndex        =   21
            Top             =   1800
            Width           =   1215
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Height          =   4335
      Left            =   180
      TabIndex        =   14
      Top             =   480
      Width           =   18210
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "COURSE DETAILS.frx":1EB571
         Height          =   4005
         Left            =   13500
         TabIndex        =   65
         Top             =   240
         Width           =   4635
         _ExtentX        =   8176
         _ExtentY        =   7064
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "CODE"
            Caption         =   "CODE"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16393
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "CNM"
            Caption         =   "NAME"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16393
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "DUR"
            Caption         =   "DURATION"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16393
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "FEE"
            Caption         =   "FEE"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16393
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "REG_FEE"
            Caption         =   "REG FEE"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16393
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "ADM_FEE"
            Caption         =   "ADM FEE"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16393
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "SEM_FEE"
            Caption         =   "SEM FEE"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16393
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1289.764
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   915.024
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H80000002&
         Height          =   1275
         Left            =   1200
         TabIndex        =   50
         Top             =   2940
         Width           =   11145
         Begin VB.CommandButton course_view 
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
            Left            =   2880
            Picture         =   "COURSE DETAILS.frx":1EB586
            Style           =   1  'Graphical
            TabIndex        =   58
            Top             =   180
            Width           =   1095
         End
         Begin VB.CommandButton course_save 
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
            Left            =   1380
            Picture         =   "COURSE DETAILS.frx":1EB93D
            Style           =   1  'Graphical
            TabIndex        =   57
            Top             =   180
            Width           =   1095
         End
         Begin VB.CommandButton course_delete 
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
            Left            =   4350
            Picture         =   "COURSE DETAILS.frx":1EC117
            Style           =   1  'Graphical
            TabIndex        =   56
            Top             =   180
            Width           =   1095
         End
         Begin VB.CommandButton course_update 
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
            Left            =   5850
            Picture         =   "COURSE DETAILS.frx":1EC683
            Style           =   1  'Graphical
            TabIndex        =   55
            Top             =   180
            Width           =   1095
         End
         Begin VB.CommandButton course_refresh 
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
            Left            =   7335
            Picture         =   "COURSE DETAILS.frx":1ECE13
            Style           =   1  'Graphical
            TabIndex        =   54
            Top             =   180
            Width           =   1095
         End
         Begin VB.CommandButton course_new 
            Caption         =   "NEW"
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
            Left            =   8700
            Picture         =   "COURSE DETAILS.frx":1ED4B7
            Style           =   1  'Graphical
            TabIndex        =   53
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton NEWCRSE 
            Caption         =   "ADD NEW COURSE"
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
            Left            =   90
            Picture         =   "COURSE DETAILS.frx":1EDC85
            Style           =   1  'Graphical
            TabIndex        =   52
            Top             =   210
            Width           =   1095
         End
         Begin VB.CommandButton RMVCRS 
            Caption         =   "REMOVE COURSE"
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
            Left            =   9960
            Picture         =   "COURSE DETAILS.frx":1EDE17
            Style           =   1  'Graphical
            TabIndex        =   51
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FF8080&
         Caption         =   "COURSE"
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
         Left            =   180
         TabIndex        =   15
         Top             =   210
         Width           =   13215
         Begin VB.OptionButton Option1 
            Caption         =   "YEAR"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   11100
            TabIndex        =   49
            Top             =   390
            Width           =   915
         End
         Begin VB.ComboBox duration 
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
            Left            =   9210
            TabIndex        =   48
            Top             =   360
            Width           =   1725
         End
         Begin VB.ComboBox coursenm 
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
            Height          =   300
            ItemData        =   "COURSE DETAILS.frx":1EE759
            Left            =   1950
            List            =   "COURSE DETAILS.frx":1EE75B
            TabIndex        =   0
            Text            =   "select"
            Top             =   330
            Width           =   2175
         End
         Begin VB.TextBox semfee 
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
            Height          =   300
            Left            =   1890
            TabIndex        =   5
            Top             =   1620
            Width           =   2175
         End
         Begin VB.TextBox admfee 
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
            Height          =   300
            Left            =   5730
            TabIndex        =   4
            Top             =   1080
            Width           =   2175
         End
         Begin VB.TextBox regfee 
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
            Height          =   300
            Left            =   1890
            TabIndex        =   3
            Top             =   1050
            Width           =   2175
         End
         Begin VB.TextBox code 
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
            Height          =   300
            Left            =   5700
            TabIndex        =   1
            Top             =   360
            Width           =   2175
         End
         Begin VB.TextBox fee 
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
            Height          =   300
            Left            =   5730
            TabIndex        =   2
            Top             =   1710
            Width           =   2175
         End
         Begin VB.Label Label14 
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
            Left            =   1590
            TabIndex        =   47
            Top             =   180
            Width           =   135
         End
         Begin VB.Label Label17 
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
            Left            =   8760
            TabIndex        =   46
            Top             =   360
            Width           =   135
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "DURATION :"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   8040
            TabIndex        =   45
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "COURSE NAME :"
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
            Left            =   300
            TabIndex        =   44
            Top             =   390
            Width           =   1455
         End
         Begin VB.Label Label25 
            BackStyle       =   0  'Transparent
            Caption         =   "REGISTRATION FEE :"
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
            TabIndex        =   43
            Top             =   1140
            Width           =   1575
         End
         Begin VB.Label Label28 
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
            TabIndex        =   42
            Top             =   930
            Width           =   135
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
            Left            =   1590
            TabIndex        =   41
            Top             =   1650
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
            Left            =   5430
            TabIndex        =   40
            Top             =   1020
            Width           =   135
         End
         Begin VB.Label Label27 
            BackStyle       =   0  'Transparent
            Caption         =   "ADMISSION FEE :"
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
            Left            =   4380
            TabIndex        =   39
            Top             =   1200
            Width           =   1335
         End
         Begin VB.Label Label26 
            BackStyle       =   0  'Transparent
            Caption         =   "SEMESTER FEE :"
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
            Left            =   630
            TabIndex        =   38
            Top             =   1770
            Width           =   1215
         End
         Begin VB.Label Label16 
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
            Left            =   5130
            TabIndex        =   37
            Top             =   300
            Width           =   135
         End
         Begin VB.Label Label15 
            BackColor       =   &H00FFFFFF&
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
            Left            =   5160
            TabIndex        =   29
            Top             =   1620
            Width           =   135
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Height          =   735
            Left            =   7080
            TabIndex        =   18
            Top             =   2520
            Width           =   2415
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   " FEE :"
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
            Left            =   4830
            TabIndex        =   17
            Top             =   1710
            Width           =   495
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "CODE :"
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
            Left            =   4710
            TabIndex        =   16
            Top             =   390
            Width           =   615
         End
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "COURSE"
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
      Left            =   7200
      TabIndex        =   13
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "frmCOURSE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub course_new_Click()
Frame2.Enabled = True
course_save.Enabled = True
course_view.Enabled = True
course_update.Enabled = True
course_delete.Enabled = True
course_refresh.Enabled = True
'courseprnt.Enabled = True
Frame3.Enabled = False
coursenm.Text = ""
fee.Text = ""
regfee.Text = ""
admfee.Text = ""
semfee.Text = ""
code.Text = ""
duration.Text = ""
suboption.Value = False
subnm.Text = ""
subid.Text = ""
paper.Text = ""
theory.Text = ""
practical.Text = ""
passmrks.Text = ""
End Sub

Private Sub coursenm_LostFocus()
code.Locked = True
End Sub

Private Sub courseprnt_Click()
If DataEnvironment4.rsprnt.state = 1 Then DataEnvironment4.rsprnt.Close
DataEnvironment4.prnt (code.Text)
'''DataReport5.Show
DataReport5.Sections("SECTION2").Controls("LABEL8").Caption = "HELLO FRNDS"
DataReport5.SHOW
End Sub

Private Sub Form_Activate()
fee.MaxLength = 10
coursenm.AddItem "BACHELOR IN COMPUTER APPLICATION"
coursenm.AddItem "BACHELOR OF SCIENCE IN INFORMATION TECHNOLOGY "
coursenm.AddItem "BACHELOR OF BUSINESS ADMINISTRATIVE"
coursenm.AddItem "BACHELOR OF BUSINESS MANAGEMENT "
fullmrks.Locked = True
Adodc2.Visible = False
End Sub

Private Sub NEWCRSE_Click()
coursenm.AddItem coursenm.Text
End Sub

Private Sub paper_lostfocus()
If (paper.Text = "SUBSIDARY") Then
practical.Locked = True
End If
End Sub
Private Sub course_refresh_Click()
code.Text = ""
coursenm.Text = ""
duration.Text = ""
fee.Text = ""
regfee.Text = ""
semfee.Text = ""
admfee.Text = ""
End Sub
Private Sub course_Delete_Click()
If code.Text = "" Or coursenm.Text = "" Or fee.Text = "" Or duration.Text = "" Then
MsgBox "ALL FIELDS ARE MANDATORY"
coursenm.SetFocus
Else
Module1.conn
A = MsgBox("ARE YOU SURE TO DELETE RECORD", vbQuestion + vbOKCancel, "DELETE")
If (A = vbOK) Then
sql = "delete from SubjectDetail_Master where code = '" + code.Text + "'"
Set r = c.Execute(sql)
Adodc2.refresh
sql = "delete from CourseDetail_Master where code='" + code.Text + "'"
Set r = c.Execute(sql)
Adodc1.refresh
MsgBox "RECORD DELETE"
code.Text = ""
coursenm.Text = ""
duration.Text = ""
fee.Text = ""
admfee.Text = ""
regfee.Text = ""
semfee.Text = ""
coursenm.SetFocus
Else
If (A = vbCancel) Then
MsgBox "RECORD NOT DELETED"
End If
End If
End If
End Sub
Private Sub course_save_Click()
If code.Text = "" Or coursenm.Text = "" Or fee.Text = "" Or duration.Text = "" Then
frmINFO.SHOW 1
If str = "OK" Then
coursenm.SetFocus
End If
Else
On Error GoTo label
Module1.conn
frmSAVE.SHOW 1
If str = "YES" Then
sql = "insert into CourseDetail_Master values('" + code.Text + "','" + coursenm.Text + "','" + duration.Text + "'," + fee.Text + "," + regfee.Text + "," + admfee.Text + "," + semfee.Text + ")"
Set r = c.Execute(sql)
frmRECORDSAVE.SHOW 1
If str = "OK" Then
Adodc1.refresh
coursenm.SetFocus
course_save.Enabled = False
course_view.Enabled = False
course_update.Enabled = False
course_delete.Enabled = False
course_refresh.Enabled = False
'courseprnt.Enabled = False
Frame2.Enabled = False
Frame4.Enabled = True
Frame3.Enabled = True
End If
MsgBox " PLEASE ENTER SUBJECT DETAIL REGARDING COURSE", vbInformation + vbOKOnly, "INFORMATION"
If A = vbOK Then
subnm.SetFocus
suboption.Value = False
sub_save.Enabled = True
fullmrks.Locked = True
End If
Else
If str = "NO" Then
code.Text = ""
coursenm.Text = ""
fee.Text = ""
duration.Text = ""
coursenm.SetFocus
End If
Exit Sub
label:
If (Err.Number = -2147217873) Then
frmDATAEXIST.SHOW 1
If str = "OK" Then
coursenm.SetFocus
End If
End If
End If
End If
End Sub

Private Sub fee_keypress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Or KeyAscii = 46) Then
fee.Locked = False
Else
fee.Locked = True
End If
If (KeyAscii = 13) Then
duration.SetFocus
If (KeyAscii = 9) Then
duration.SetFocus
End If
End If
End Sub
Private Sub code_lostfocus()
code.Text = UCase(code.Text)
End Sub

Private Sub coursenm_click()
If (coursenm.ListIndex = 0) Then
code.SetFocus
code.Text = "BCA"
End If
If (coursenm.ListIndex = 1) Then
code.SetFocus
code.Text = "BSCIT"
End If
If (coursenm.ListIndex = 2) Then
code.SetFocus
code.Text = "BBA"
End If
If (coursenm.ListIndex = 3) Then
code.SetFocus
code.Text = "BBM"
End If
End Sub

Private Sub course_update_Click()
On Error GoTo clear
Module1.conn
A = MsgBox("CLICK OK TO UPDATE RECORD", vbInformation + vbOKCancel, "INFORMATION")
If (A = vbOK) Then
sql = "update CourseDetail_Master set cnm='" + coursenm.Text + "' ,dur='" + duration.Text + "' ,fee=" + fee.Text + ""
Set r = c.Execute(sql)
Adodc1.refresh
MsgBox "RECORD UPDATE"
code.Text = ""
coursenm.Text = ""
fee.Text = ""
duration.Text = ""
Else
If (A = vbCancel) Then
MsgBox "RECORD NOT UPDATE"
Exit Sub
clear:
MsgBox "ALL FIELDS ARE MANDATORY"
End If
End If
End Sub

Private Sub course_view_Click()
Frame3.Enabled = True
sub_save.Enabled = False
course_save.Enabled = False
If code.Text = "" Then
MsgBox "ALL FIELDS ARE MANDATORY"
coursenm.SetFocus
Else
On Error GoTo label
Module1.conn
sql = "select *from CourseDetail_Master where code='" + code.Text + "'"
Set r = c.Execute(sql)
coursenm.Text = r.Fields(1)
fee.Text = r.Fields(3)
duration.Text = r.Fields(2)
regfee.Text = r.Fields(4)
admfee.Text = r.Fields(5)
semfee.Text = r.Fields(6)
Exit Sub
label:
If (r.EOF = True) Then
A = MsgBox("RECORD NOT FOUND", vbInformation + vbOKOnly, "INFORMATION")
code.Text = ""
coursenm.SetFocus
End If
End If
End Sub
Private Sub coursenm_keypress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 90 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 32 Or KeyAscii = 8) Then
coursenm.Locked = False
Else
coursenm.Locked = True
End If
If (KeyAscii = 13) Then
fee.SetFocus
If (KeyAscii = 9) Then
fee.SetFocus
End If
End If
End Sub

Private Sub Form_Load()
fullmrks.Text = 100
subnm.MaxLength = 20
'coursenm.MaxLength = 60
fee.MaxLength = 10

practical.MaxLength = 2
subid.MaxLength = 6
subnm.MaxLength = 20
theory.MaxLength = 3
practical.MaxLength = 2
fullmrks.MaxLength = 3
code.MaxLength = 6
passmrks.MaxLength = 2
theory.Locked = True
practical.Locked = True
passmrks.Locked = True
Adodc1.Visible = False
Frame4.Enabled = False
Frame3.Enabled = False
duration.AddItem "1"
duration.AddItem "2"
duration.AddItem "3"

End Sub

Private Sub semfee_LostFocus()
fee.Text = Val(semfee.Text) * 6
End Sub

Private Sub sub_delete_Click()
If (subid.Text = "" Or paper.Text = "" Or subnm.Text = "" Or fullmrks.Text = "" Or practical.Text = "" Or theory.Text = "" Or passmrks.Text = "") Then
MsgBox "ALL FIELDS ARE MANDATORY"
subnm.SetFocus
Else
Module1.conn
A = MsgBox("WANT TO DELETE RECORD", vbQuestion + vbOKCancel, "QUESTION")
If (A = vbOK) Then
sql = "delete from SubjectDetail_Master where sub_id='" + subid + "'"
Set r = c.Execute(sql)
Adodc2.refresh
MsgBox "RECORD DELETE"
subid = ""
subnm = ""
practical = ""
theory = ""
fullmrks = ""
passmrks = ""
paper.Text = ""
subnm.SetFocus
Else
If (A = vbCancel) Then
MsgBox "RECORD NOT DELETED"
End If
End If
End If
End Sub

Private Sub sub_save_Click()
If (subid.Text = "" Or paper.Text = "" Or subnm.Text = "" Or fullmrks.Text = "" Or practical.Text = "" Or theory.Text = "" Or passmrks.Text = "") Then
MsgBox "ALL FIELDS ARE MANDATORY"
subnm.SetFocus
Else
Module1.conn
A = MsgBox("CLICK OK TO SAVE RECORD", vbInformation + vbOKCancel, "INFORMATION")
If (A = vbOK) Then
sql = "insert into SubjectDetail_Master values('" + subid.Text + "','" + subnm.Text + "','" + paper.Text + "'," + fullmrks.Text + "," + theory.Text + "," + practical.Text + "," + passmrks.Text + ",'" + code.Text + "')"
Set r = c.Execute(sql)
Adodc2.refresh
MsgBox "RECORD SAVE"
b = MsgBox("ARE YOU WANT ADD ANOTHER RECORD", vbQuestion + vbYesNo, "QUESTION")
If (b = vbYes) Then
subid = ""
subnm = ""
practical = ""
theory = ""
passmrks = ""
subnm.SetFocus
If (b = vbNo) Then
subid = ""
subnm = ""
practical = ""
theory = ""
fullmrks = ""
passmrks = ""
course_save.Enabled = True
coursenm.SetFocus
Else
If (A = vbCancel) Then
MsgBox "RECORD NOT SAVE"
End If
End If
End If
End If
End If
End Sub

Private Sub sub_update_Click()
If (subid.Text = "" Or paper.Text = "" Or subnm.Text = "" Or fullmrks.Text = "" Or practical.Text = "" Or theory.Text = "" Or passmrks.Text = "") Then
MsgBox "ALL FIELDS ARE MANDATORY"
subnm.SetFocus
Else
Module1.conn
A = MsgBox("CLICK OK TO SAVE RECORD", vbInformation + vbOKCancel, "INFORMATION")
If (A = vbOK) Then
sql = "update SubjectDetail_Master set sub_nm='" + subnm.Text + "',sub_paper='" + paper.Text + "',sub_fullmarks=" + fullmrks.Text + ",sub_theorymarks=" + theory.Text + " ,sub_pracmarks=" + practical.Text + ",sub_passmarks=" + passmrks.Text + ""
Set r = c.Execute(sql)
MsgBox "RECORD UPDATE"
subid = ""
subnm = ""
practical = ""
theory = ""
fullmrks = ""
passmrks = ""
paper.Text = ""
subnm.SetFocus
Else
If (A = vbCancel) Then
MsgBox "RECORD NOT UPDATE"
End If
End If
End If
End Sub

Private Sub sub_view_Click()
If (subid.Text = "") Then
MsgBox "SUBJECT ID IS MANDATORY"
subid.SetFocus
Else
'On Error GoTo fnd
Module1.conn
sql = "select *from SubjectDetail_Master where sub_id='" + subid.Text + "'"
Set r = c.Execute(sql)
subnm.Text = r.Fields(1)
paper.Text = r.Fields(2)
fullmrks.Text = r.Fields(3)
theory.Text = r.Fields(4)
practical.Text = r.Fields(5)
passmrks.Text = r.Fields(6)
'Exit Sub
'fnd:
If (r.EOF = True) Then
MsgBox "RECORD NOT FOUNDD"
End If
End If
End Sub
'
'Private Sub subid_GotFocus()
'If (subnm = "") Then
'subid = ""
'End If
'End Sub

Private Sub subnm_LostFocus()
UCase (subnm.Text)
subid.Text = code.Text + Left(subnm.Text, 5)
End Sub

'Private Sub subid_keypress(keyascii As Integer)
'If (keyascii >= 48 And keyascii <= 57 Or keyascii = 32) Then
'subid.Locked = True
'Else
'subid.Locked = False
'End If
'End Sub
'Private Sub fullmrks_keypress(keyascii As Integer)
'If (keyascii >= 48 And keyascii <= 57 Or keyascii = 8) Then
'fullmrks.Locked = False
'Else
'fullmrks.Locked = True
'End If
'End Sub

Private Sub suboption_Click()
Frame2.Enabled = False
Frame3.Enabled = True
Frame4.Enabled = True
sub_save.Enabled = False
sub_view.Enabled = True
sub_delete.Enabled = True
sub_refresh.Enabled = True
'sub_prnt.Enabled = True
sub_update.Enabled = True
course_save.Enabled = False
course_view.Enabled = False
course_update.Enabled = False
course_delete.Enabled = False
course_refresh.Enabled = False
'courseprnt.Enabled = False
End Sub



Private Sub theory_lostFocus()
If (Val(theory.Text) > Val(fullmrks.Text)) Then
MsgBox "theory marks should be less than fullmarks"
theory.Text = ""
theory.SetFocus
End If
practical.Text = Val(fullmrks.Text) - Val(theory.Text)
passmrks.Text = Val(theory.Text) * 30 / 100
End Sub

Private Sub theory_keypress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
theory.Locked = False
Else
theory.Locked = True
End If
End Sub
Private Sub practical_keypress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
practical.Locked = False
Else
practical.Locked = True
End If
End Sub
Private Sub passmrks_keypress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 37 Or KeyAscii = 8) Then
passmrks.Locked = False
Else
passmrks.Locked = True
End If
End Sub
Private Sub subnm_keypress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 90 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32) Then
subnm.Locked = False
Else
subnm.Locked = True
End If
End Sub
'Private Sub code_keypress(keyascii As Integer)
'If (keyascii >= 65 And keyascii <= 90 Or keyascii >= 97 And keyascii <= 122 Or keyascii = 32 Or keyascii = 8) Then
'code.Locked = False
'Else
'code.Locked = True
'End If
'End Sub
'Private Sub fullmrks_lostfocus()
'If (fullmrks.Text < 100) Then
'MsgBox "PLEASE ENTER VALID VALUE"
'fullmrks.Text = ""
'End If
'End Sub
Private Sub subid_gotfocus()

End Sub

