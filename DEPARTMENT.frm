VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmDEPARTMENT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DEPARTMENT"
   ClientHeight    =   8940
   ClientLeft      =   870
   ClientTop       =   1530
   ClientWidth     =   18885
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   Picture         =   "DEPARTMENT.frx":0000
   ScaleHeight     =   8940
   ScaleWidth      =   18885
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   405
      Left            =   420
      Top             =   1650
      Width           =   1200
      _ExtentX        =   2117
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
      RecordSource    =   "select *from DepartmentDetail_Master"
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
      Height          =   6675
      Left            =   1680
      TabIndex        =   1
      Top             =   630
      Width           =   15255
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "DEPARTMENT.frx":1E9608
         Height          =   2145
         Left            =   2880
         Negotiate       =   -1  'True
         TabIndex        =   17
         Top             =   2610
         Width           =   10245
         _ExtentX        =   18071
         _ExtentY        =   3784
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   14
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "DCODE"
            Caption         =   "DEPARTMENT CODE"
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
            DataField       =   "DNM"
            Caption         =   "DEPARTMENT NAME"
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
            DataField       =   "DTYPE"
            Caption         =   "DEPARTMENT TYPE"
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
            DataField       =   "HOD"
            Caption         =   "HEAD OF DEPARTMENT"
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
               ColumnWidth     =   2429.858
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H80000002&
         Height          =   1275
         Left            =   2040
         TabIndex        =   11
         Top             =   4980
         Width           =   10365
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
            Left            =   2490
            Picture         =   "DEPARTMENT.frx":1E961D
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   210
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
            Left            =   330
            Picture         =   "DEPARTMENT.frx":1E99D4
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   210
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
            Left            =   4560
            Picture         =   "DEPARTMENT.frx":1EA1AE
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   210
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
            Left            =   6660
            Picture         =   "DEPARTMENT.frx":1EA71A
            Style           =   1  'Graphical
            TabIndex        =   13
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
            Left            =   8760
            Picture         =   "DEPARTMENT.frx":1EAEAA
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   270
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FF8080&
         Height          =   2055
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   15015
         Begin VB.ComboBox dtype 
            BackColor       =   &H00E0E0E0&
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
            ItemData        =   "DEPARTMENT.frx":1EB54E
            Left            =   3870
            List            =   "DEPARTMENT.frx":1EB550
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   1320
            Width           =   2415
         End
         Begin VB.TextBox hod 
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
            Left            =   11160
            TabIndex        =   8
            Top             =   1320
            Width           =   2415
         End
         Begin VB.TextBox dnm 
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
            Left            =   3840
            TabIndex        =   7
            Top             =   240
            Width           =   2415
         End
         Begin VB.TextBox dcode 
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
            Left            =   11190
            TabIndex        =   6
            Top             =   300
            Width           =   2415
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "NAME"
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
            Left            =   1680
            TabIndex        =   10
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "HOD :"
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
            Left            =   9120
            TabIndex        =   5
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "CODE"
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
            Left            =   9120
            TabIndex        =   4
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "TYPE :"
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
            Left            =   1680
            TabIndex        =   3
            Top             =   1320
            Width           =   975
         End
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "DEPARTMENT"
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
      Left            =   5640
      TabIndex        =   0
      Top             =   0
      Width           =   6855
   End
End
Attribute VB_Name = "frmDEPARTMENT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub delete_Click()
If (dnm.Text = "" Or dtype.Text = "" Or dcode.Text = "" Or hod.Text = "") Then
frmINFO.SHOW 1
If (str = "OK") Then
dcode.SetFocus
End If
Else
Module1.conn
frmDELETE.SHOW 1
If (str = "YES") Then
sql = "delete from DepartmentDetail_Master where dcode='" + dcode.Text + "'"
Set r = c.Execute(sql)
Adodc1.refresh
frmRECORDELETE.SHOW 1
If (str = "OK") Then
dnm.Text = ""
hod.Text = ""
dcode.Text = ""
dcode.SetFocus
End If
ElseIf (str = "NO") Then
dcode.SetFocus
End If
End If
End Sub

Private Sub dtype_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
hod.SetFocus
End If
End Sub

Private Sub dnm_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 87 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32) Then
dnm.Locked = False
Else
dnm.Locked = True
End If
If (KeyAscii = 13) Then
dtype.SetFocus
End If
End Sub

Private Sub dnm_LostFocus()
dnm = UCase(dnm)
End Sub
Private Sub hod_LostFocus()
hod = UCase(hod)
End Sub

Private Sub hod_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 87 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32) Then
hod.Locked = False
Else
hod.Locked = True
End If
End Sub
Private Sub Form_Load()
delete.Enabled = False
update.Enabled = False
dcode.Locked = True
'dcode1.Visible = False
Module1.conn
sql = "select count(dcode) from DepartmentDetail_Master"
Set r = c.Execute(sql)
i = r.Fields(0)
If (i < 1) Then
dcode.Text = "D" & ("00") & (1)
Else
sql = "select dcode from DepartmentDetail_Master order by rownum desc"
Set r = c.Execute(sql)
A = r.Fields(0)
b = Right(A, 1)
b = b + 1
f = "D" & "00" & (b)
dcode.Text = f
End If
Adodc1.Visible = False
dtype.AddItem "NON-ACADEMIC"
dtype.AddItem "ACADEMIC"
dcode.MaxLength = 5
dnm.MaxLength = 30
hod.MaxLength = 20
End Sub

Private Sub refresh_Click()
Module1.conn
save.Enabled = True
delete.Enabled = False
update.Enabled = False
'dcode1.Visible = False
dnm.Text = ""
dcode.Text = ""
hod.Text = ""
sql = "select max(dcode) from departmentdetail_master"
Set r = c.Execute(sql)
i = r.Fields(0)
A = i
b = Right(A, 1)
b = b + 1
s = "D" & "00" & b
dcode.Text = s
End Sub

Private Sub save_Click()
On Error GoTo label
If (dnm.Text = "" Or dtype.Text = "" Or dcode.Text = "") Then
frmINFO.SHOW 1
If (str = "OK") Then
dcode.SetFocus
End If
Else
Module1.conn
frmSAVE.SHOW 1
If (str = "YES") Then
sql = "insert into DepartmentDetail_Master values('" + dcode.Text + "','" + dnm.Text + "','" + dtype.Text + "','" + hod.Text + "')"
Set r = c.Execute(sql)
'MsgBox sql
Adodc1.refresh
frmRECORDSAVE.SHOW 1
If (str = "OK") Then
A = dcode.Text
b = Right(A, 1)
b = b + 1
s = "D" & "00" & b
dcode.Text = s
hod.Text = ""
dnm.Text = ""
dcode.SetFocus
End If
ElseIf (str = "NO") Then
dnm.Text = ""
hod.Text = ""
dcode.Text = ""
dcode.SetFocus
Exit Sub
label:
A = MsgBox("DATA ALREADY EXIST", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
dnm.Text = ""
hod.Text = ""
dcode.Text = ""
dcode.SetFocus
End If
End If
End If
End Sub

Private Sub update_Click()
If (dnm.Text = "" Or dtype.Text = "" Or dcode.Text = "" Or hod.Text = "") Then
frmINFO.SHOW 1
If (str = "OK") Then
dcode.SetFocus
End If
Else
Module1.conn
frmUPDATE.SHOW 1
If (str = "YES") Then
sql = "update DepartmentDetail_Master set dnm='" + dnm.Text + "',dtype='" + dtype.Text + "',hod='" + hod.Text + "' where dcode='" + dcode.Text + "'"
Set r = c.Execute(sql)
Adodc1.refresh
frmRECORDUPDATE.SHOW 1
If (str = "OK") Then
dnm.Text = ""
hod.Text = ""
dcode.Text = ""
dcode.SetFocus
End If
ElseIf (str = "NO") Then
dcode.SetFocus
End If
End If
End If
End Sub

Private Sub view_Click()
save.Enabled = False
delete.Enabled = True
update.Enabled = True
A = InputBox("ENTER DEPARTMENT CODE", "VIEW")
dcode.Text = A
'dcode1.Text = a
'dcode1.Visible = True
'On Error GoTo label
If (dcode.Text = "") Then
A = MsgBox("PLEASE ENTER DEPARTMENT CODDE", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
dcode.SetFocus
End If
Else
Module1.conn
sql = "SELECT *FROM DepartmentDetail_Master where dcode='" + dcode.Text + "'"
Set r = c.Execute(sql)
'dcode1.Text = r.Fields(0)
dnm.Text = r.Fields(1)
dtype.Text = r.Fields(2)
If (IsNull(r.Fields(3))) Then
hod.Text = ""
Else
hod.Text = r.Fields(3)
End If
'Exit Sub
'label:
'A = MsgBox("RECORD NOT FOUND", vbInformation + vbOKOnly, "INFORMATION")
'If (A = vbOK) Then
'dcode.SetFocus
'End If
End If
End Sub

