VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmEXAMINATION 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EXAMINATION"
   ClientHeight    =   8520
   ClientLeft      =   1530
   ClientTop       =   1695
   ClientWidth     =   18735
   BeginProperty Font 
      Name            =   "Cambria"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   Picture         =   "EXAMINATION DETAILS.frx":0000
   ScaleHeight     =   8520
   ScaleWidth      =   18735
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   405
      Left            =   1890
      Top             =   180
      Visible         =   0   'False
      Width           =   2745
      _ExtentX        =   4842
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
      RecordSource    =   "select *from examdetail_master"
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
      Height          =   7095
      Left            =   1200
      TabIndex        =   5
      Top             =   720
      Width           =   16215
      Begin VB.Frame Frame2 
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
         Height          =   6735
         Left            =   210
         TabIndex        =   6
         Top             =   150
         Width           =   15855
         Begin MSDataGridLib.DataGrid DataGrid1 
            Bindings        =   "EXAMINATION DETAILS.frx":1E9608
            Height          =   2205
            Left            =   2400
            TabIndex        =   20
            Top             =   2640
            Width           =   10605
            _ExtentX        =   18706
            _ExtentY        =   3889
            _Version        =   393216
            HeadLines       =   1
            RowHeight       =   15
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Cambria"
               Size            =   9
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   5
            BeginProperty Column00 
               DataField       =   "EXAM_CODE"
               Caption         =   "EXAM CODE"
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
               DataField       =   "EXAM_NM"
               Caption         =   "EXAM NAME"
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
               DataField       =   "EXAM_DATE"
               Caption         =   "EXAM DATE"
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
               DataField       =   "EXAM_TIME"
               Caption         =   "EXAM TIME"
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
               DataField       =   "COURSE"
               Caption         =   "COURSE"
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
                  ColumnWidth     =   2429.858
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   2429.858
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   2429.858
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   1590.236
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   1275.024
               EndProperty
            EndProperty
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H80000002&
            Height          =   1365
            Left            =   2490
            TabIndex        =   14
            Top             =   5100
            Width           =   9615
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
               Left            =   2340
               Picture         =   "EXAMINATION DETAILS.frx":1E961D
               Style           =   1  'Graphical
               TabIndex        =   19
               Top             =   240
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
               Left            =   4290
               Picture         =   "EXAMINATION DETAILS.frx":1E99D4
               Style           =   1  'Graphical
               TabIndex        =   18
               Top             =   240
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
               Left            =   6180
               Picture         =   "EXAMINATION DETAILS.frx":1E9F40
               Style           =   1  'Graphical
               TabIndex        =   17
               Top             =   240
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
               Left            =   8130
               Picture         =   "EXAMINATION DETAILS.frx":1EA6D0
               Style           =   1  'Graphical
               TabIndex        =   16
               Top             =   240
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
               Left            =   420
               Picture         =   "EXAMINATION DETAILS.frx":1EAD74
               Style           =   1  'Graphical
               TabIndex        =   15
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.TextBox exmcode 
            Height          =   375
            Left            =   12450
            TabIndex        =   2
            Top             =   450
            Width           =   2505
         End
         Begin VB.TextBox dt1 
            Height          =   300
            Left            =   2730
            TabIndex        =   3
            Top             =   1680
            Width           =   2295
         End
         Begin VB.TextBox time 
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
            Left            =   8130
            TabIndex        =   4
            Top             =   1590
            Width           =   2295
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
            Left            =   2790
            TabIndex        =   0
            Text            =   "SELECT"
            Top             =   510
            Width           =   2295
         End
         Begin MSComCtl2.DTPicker dt 
            Height          =   300
            Left            =   2730
            TabIndex        =   13
            Top             =   1650
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   529
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
            CalendarBackColor=   -2147483638
            CustomFormat    =   "dd/MMM/yyyy"
            Format          =   107937795
            CurrentDate     =   43086
         End
         Begin VB.TextBox exmnm 
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
            Left            =   8160
            TabIndex        =   1
            Top             =   480
            Width           =   2295
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "TIME :"
            Height          =   375
            Left            =   5700
            TabIndex        =   12
            Top             =   1590
            Width           =   495
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "EXAM CODE :"
            Height          =   375
            Left            =   11310
            TabIndex        =   11
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "SUBJECT EXAM DATE :"
            Height          =   375
            Left            =   600
            TabIndex        =   10
            Top             =   1650
            Width           =   1695
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "COURSE CODE :"
            Height          =   375
            Left            =   750
            TabIndex        =   9
            Top             =   510
            Width           =   1215
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "EXAM NAME/TYPE :"
            Height          =   375
            Left            =   5700
            TabIndex        =   8
            Top             =   510
            Width           =   1575
         End
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "EXAMINATION"
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
      Left            =   5040
      TabIndex        =   7
      Top             =   120
      Width           =   8295
   End
End
Attribute VB_Name = "frmEXAMINATION"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub course_Click()
exmcode.Text = ""
End Sub

Private Sub delete_Click()
If (exmnm.Text = "" Or time.Text = "") Then
A = MsgBox("ALL FIELDS ARE MANDATORY", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
exmnm.SetFocus
End If
Else
Module1.conn
A = MsgBox("CLICK OK TO DELETE RECORD", vbQuestion + vbOKCancel, "DELETE")
If (A = vbOK) Then
sql = "DELETE FROM EXAMDETAIL_MASTER WHERE EXAM_CODE='" + exmcode.Text + "'"
Set r = c.Execute(sql)
A = MsgBox("RECORD DELETED", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
Adodc1.refresh
Call ref
dt1.Text = ""
exmnm.SetFocus
End If
Else
If (A = vbCancel) Then
exmnm.SetFocus
End If
End If
End If
End Sub

Private Sub exmnm_keypress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 32) Then
exmnm.Locked = True
Else
exmnm.Locked = False
End If
End Sub

Private Sub Form_Load()
dt1.Visible = False
update.Enabled = False
delete.Enabled = False
exmnm.MaxLength = 20
time.MaxLength = 5
exmcode.Text = ""
Module1.conn
sql = "select code from coursedetail_master"
Set r = c.Execute(sql)
While (r.EOF = False)
course.AddItem r.Fields("code")
r.MoveNext
Wend
End Sub

Private Sub refresh_Click()
update.Enabled = False
delete.Enabled = False
Call ref
End Sub

Private Sub save_Click()
On Error GoTo label
If (exmnm.Text = "" Or time.Text = "") Then
frmINFO.SHOW 1
If str = "OK" Then
exmnm.SetFocus
End If
Else
Module1.conn
frmSAVE.SHOW 1
If str = "YES" Then
sql = "INSERT INTO EXAMDETAIL_MASTER VALUES('" + exmcode.Text + "','" + exmnm.Text + "','" + Format(dt.Value, "dd/MMM/yyyy") + "','" + time.Text + "','" + course.Text + "')"
Set r = c.Execute(sql)
frmRECORDSAVE.SHOW 1
Adodc1.refresh
Call ref
exmnm.SetFocus
End If
ElseIf (str = "NO") Then
exmnm.SetFocus
End If
End If
Exit Sub
label:
frmDATAEXIST.SHOW 1
If str = "OK" Then
Call ref
exmnm.SetFocus
End If
End Sub
Private Sub ref()
save.Enabled = True
exmnm.Text = ""
time.Text = ""
End Sub

Private Sub time_keypress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Or KeyAscii = 58) Then
time.Locked = False
Else
time.Locked = True
End If
End Sub
Private Sub exmnm_lostfocus()
If (exmnm.Text = exmnm.Text) Then
A = Left(exmnm.Text, 5)
b = Right(dt.Value, 2)
i = 0
While (i <= course.ListCount - 1)
'exmcode.List(i) = (A) & course.List(i) & (b)
exmcode.Text = (A) & course.List(i) & (b)

i = i + 1
Wend
End If
End Sub

Private Sub update_Click()
If (exmnm.Text = "" Or time.Text = "") Then
frmINFO.SHOW 1
If str = "OK" Then
exmnm.SetFocus
End If
Else
Module1.conn
frmSAVE.SHOW 1
sql = "UPDATE EXAMDETAIL_MASTER WHERE EXAM_CODE='" + exmcode.Text + "'"
Set r = c.Execute(sql)
frmRECORDSAVE.SHOW 1
If str = "OK" Then
Adodc1.refresh
Call ref
dt1.Text = ""
exmnm.SetFocus
End If
ElseIf (str = "NO") Then
exmnm.SetFocus
End If
End If
End Sub

Private Sub view_Click()
On Error GoTo label
update.Enabled = True
delete.Enabled = True
save.Enabled = False
Module1.conn
frmEXAMCODE.SHOW 1
If str = "ok" Then
A = frmUSERINPUT.val
exmcode.Text = A
sql = "SELECT *FROM EXAMDETAIL_MASTER WHERE EXAM_CODE='" + exmcode.Text + "'"
Set r = c.Execute(sql)
exmcode.Text = r.Fields(0)
exmnm.Text = r.Fields(1)
dt1.Text = r.Fields(2)
time.Text = r.Fields(3)
course.Text = r.Fields(4)
Exit Sub
label:
frmDATAEXIST.SHOW 1
If str = "ok" Then
exmcode.SetFocus
End If
End Sub
