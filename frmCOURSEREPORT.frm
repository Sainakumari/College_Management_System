VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmCOURSEREPORT 
   Caption         =   "COURSE REPORT"
   ClientHeight    =   3645
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11190
   LinkTopic       =   "Form2"
   ScaleHeight     =   3645
   ScaleWidth      =   11190
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Height          =   3675
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11235
      Begin VB.Frame Frame2 
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
         Height          =   3555
         Left            =   0
         TabIndex        =   1
         Top             =   240
         Width           =   10875
         Begin VB.OptionButton collective 
            Caption         =   "COLLECTIVE"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   7830
            TabIndex        =   6
            Top             =   390
            Width           =   2685
         End
         Begin VB.OptionButton selective 
            Caption         =   "SELECTIVE"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   7800
            TabIndex        =   5
            Top             =   1080
            Width           =   2685
         End
         Begin VB.CommandButton show 
            Caption         =   "SHOW"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   8280
            TabIndex        =   4
            Top             =   2040
            Width           =   1215
         End
         Begin VB.ComboBox COURSE 
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
            ItemData        =   "frmCOURSEREPORT.frx":0000
            Left            =   2760
            List            =   "frmCOURSEREPORT.frx":0010
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   480
            Width           =   2655
         End
         Begin MSAdodcLib.Adodc Adodc1 
            Height          =   330
            Left            =   120
            Top             =   1800
            Visible         =   0   'False
            Width           =   1200
            _ExtentX        =   2117
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
            RecordSource    =   "select *from coursedetail_master"
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
         Begin MSDataGridLib.DataGrid DataGrid1 
            Bindings        =   "frmCOURSEREPORT.frx":002B
            Height          =   975
            Left            =   1680
            TabIndex        =   3
            Top             =   1320
            Width           =   5535
            _ExtentX        =   9763
            _ExtentY        =   1720
            _Version        =   393216
            HeadLines       =   1
            RowHeight       =   14
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Cambria"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Cambria"
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
                  LCID            =   1033
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
                  LCID            =   1033
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
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "COURSE:"
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
            Left            =   630
            TabIndex        =   7
            Top             =   480
            Width           =   1575
         End
      End
   End
End
Attribute VB_Name = "frmCOURSEREPORT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub show_Click()
If (selective.Value = True) Then
    If DataEnvironment1.rscourseslc.state = 1 Then
    DataEnvironment1.rscourseslc.Close
    End If
    Module1.conn
   ' On Error GoTo Lbel
    sql = "select *from coursedetail_master where code='" + COURSE.Text + "' "
    Set r = c.Execute(sql)
    courseslct.Sections("SECTION2").Controls("label5").Caption = r.Fields(0)
    courseslct.Sections("SECTION2").Controls("label13").Caption = r.Fields(1)
    courseslct.Sections("SECTION2").Controls("label19").Caption = r.Fields(3)
     courseslct.Sections("SECTION2").Controls("label20").Caption = r.Fields(4)
      courseslct.Sections("SECTION2").Controls("label21").Caption = r.Fields(5)
       courseslct.Sections("SECTION2").Controls("label22").Caption = r.Fields(6)
        'courseslct.Sections("SECTION2").Controls("label23").Caption = r.Fields(7)
    courseslct.show
    Exit Sub
'Lbel:
'    A = MsgBox("RECORD NOT FOUND", vbInformation + vbOKOnly, "INFROMATION")
'    If (A = vbOK) Then
'    batch.SetFocus
'    End If
   ElseIf (collective.Value = True) Then
           coursecoll.show
           End If
End Sub
