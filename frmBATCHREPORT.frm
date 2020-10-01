VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmBATCHREPORT 
   Caption         =   "BATCH REPORT"
   ClientHeight    =   3660
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10875
   LinkTopic       =   "Form2"
   ScaleHeight     =   3660
   ScaleWidth      =   10875
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
            ItemData        =   "frmBATCHREPORT.frx":0000
            Left            =   2760
            List            =   "frmBATCHREPORT.frx":0010
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   480
            Width           =   2655
         End
         Begin MSAdodcLib.Adodc Adodc1 
            Height          =   330
            Left            =   120
            Top             =   2160
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
            RecordSource    =   "select *from batchdetail_master"
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
            Bindings        =   "frmBATCHREPORT.frx":002B
            Height          =   975
            Left            =   2040
            TabIndex        =   9
            Top             =   2400
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
         Begin VB.ComboBox batch 
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
            Left            =   2760
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   960
            Width           =   2655
         End
         Begin VB.ComboBox year 
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
            Left            =   2760
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   1560
            Width           =   2655
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
            TabIndex        =   3
            Top             =   1080
            Width           =   2685
         End
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
            TabIndex        =   2
            Top             =   390
            Width           =   2685
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
            TabIndex        =   10
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
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
            Height          =   465
            Left            =   360
            TabIndex        =   6
            Top             =   1680
            Width           =   2115
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00FF8080&
            Caption         =   "BATCH:"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   360
            TabIndex        =   5
            Top             =   1200
            Width           =   2115
         End
      End
   End
End
Attribute VB_Name = "frmBATCHREPORT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub batch_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Then
batch.Locked = False
Else
batch.Locked = True
End If
End Sub

Private Sub Form_Load()
batch.AddItem "1"
batch.AddItem "2"
batch.AddItem "3"
batch.AddItem "4"
year.AddItem "I"
year.AddItem "II"
year.AddItem "III"
End Sub

Private Sub show_Click()
If (selective.Value = True) Then
    If DataEnvironment1.rsprnt.state = 1 Then
    DataEnvironment1.rsprnt.Close
    End If
    Module1.conn
   ' On Error GoTo Lbel
    sql = "select *from batchdetail_master where course='" + COURSE.Text + "' and batno='" & batch.Text & "' and year='" + year.Text + "'"
    Set r = c.Execute(sql)
    batchslrpt.Sections("SECTION1").Controls("label14").Caption = r.Fields(0)
    batchslrpt.Sections("SECTION1").Controls("label15").Caption = r.Fields(1)
    batchslrpt.Sections("SECTION1").Controls("label16").Caption = r.Fields(2)
    batchslrpt.Sections("SECTION1").Controls("label17").Caption = r.Fields(3)
    batchslrpt.show
    Exit Sub
Lbel:
    A = MsgBox("RECORD NOT FOUND", vbInformation + vbOKOnly, "INFROMATION")
    If (A = vbOK) Then
    batch.SetFocus
    End If
   ElseIf (collective.Value = True) Then
    batch.Enabled = False
    year.Enabled = False
    COURSE.Enabled = False
           batchcollrpt.show
           End If
End Sub
