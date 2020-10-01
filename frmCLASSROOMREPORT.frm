VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmCLASSROOMREPORT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CLASS ROOM REPORT"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11010
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   11010
   StartUpPosition =   2  'CenterScreen
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
         Left            =   120
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
            TabIndex        =   7
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
            TabIndex        =   6
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
            TabIndex        =   5
            Top             =   2040
            Width           =   1215
         End
         Begin VB.ComboBox clsrm 
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
            ItemData        =   "frmCLASSROOMREPORT.frx":0000
            Left            =   2760
            List            =   "frmCLASSROOMREPORT.frx":0002
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   960
            Width           =   2655
         End
         Begin VB.ComboBox clstype 
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
            ItemData        =   "frmCLASSROOMREPORT.frx":0004
            Left            =   2760
            List            =   "frmCLASSROOMREPORT.frx":0006
            Style           =   2  'Dropdown List
            TabIndex        =   2
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
            RecordSource    =   "select *from classdetail_master"
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
            Bindings        =   "frmCLASSROOMREPORT.frx":0008
            Height          =   1575
            Left            =   1680
            TabIndex        =   3
            Top             =   1680
            Width           =   5535
            _ExtentX        =   9763
            _ExtentY        =   2778
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
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00FF8080&
            Caption         =   "CLASSROOM NO:"
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
            TabIndex        =   9
            Top             =   1200
            Width           =   2115
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "TYPE:"
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
            TabIndex        =   8
            Top             =   480
            Width           =   1575
         End
      End
   End
End
Attribute VB_Name = "frmCLASSROOMREPORT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub clstype_CLICK()
If (clstype.Text = "CLASS ROOM") Then
clsrm.clear
clsrm.AddItem "1"
clsrm.AddItem "2"
clsrm.AddItem "3"
clsrm.AddItem "4"
clsrm.AddItem "5"
End If
If (clstype.Text = "COMPUTER LAB") Then
clsrm.clear
clsrm.AddItem "X"
clsrm.AddItem "Y"
clsrm.AddItem "Z"
End If
If (clstype.Text = "PROJECT ROOM") Then
clsrm.clear
clsrm.AddItem "P1"
clsrm.AddItem "P2"
clsrm.AddItem "P3"
End If
End Sub

Private Sub Form_Load()
clstype.AddItem "CLASS ROOM"
clstype.AddItem "PROJECT ROOM"
clstype.AddItem "COMPUTER LAB"
End Sub

Private Sub show_Click()
If (selective.Value = True) Then
    If DataEnvironment1.rsclassslc.state = 1 Then
    DataEnvironment1.rsclassslc.Close
    End If
    Module1.conn
   ' On Error GoTo Lbel
    sql = "select *from classdetail_master where cls_rmno='" & clsrm.Text & "' and cls_type='" & clstype.Text & "' "
    Set r = c.Execute(sql)
    clsslc.Sections("SECTION2").Controls("label24").Caption = r.Fields(0)
    clsslc.Sections("SECTION2").Controls("label25").Caption = r.Fields(1)
    clsslc.Sections("SECTION2").Controls("label26").Caption = r.Fields(2)
    clsslc.show
    Exit Sub
'Lbel:
'    A = MsgBox("RECORD NOT FOUND", vbInformation + vbOKOnly, "INFROMATION")
'    If (A = vbOK) Then
'    batch.SetFocus
'    End If
   ElseIf (collective.Value = True) Then
           classcoll.show
           End If
End Sub

