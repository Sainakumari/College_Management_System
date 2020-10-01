VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmCLASSROOM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CLASS ROOM"
   ClientHeight    =   9075
   ClientLeft      =   1365
   ClientTop       =   1200
   ClientWidth     =   16470
   LinkTopic       =   "Form15"
   MaxButton       =   0   'False
   Picture         =   "CLASS ROOM DETAILS.frx":0000
   ScaleHeight     =   9075
   ScaleWidth      =   16470
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   390
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Height          =   4575
      Left            =   1440
      TabIndex        =   2
      Top             =   690
      Width           =   13695
      Begin VB.Frame Frame3 
         BackColor       =   &H80000002&
         Height          =   1485
         Left            =   2790
         TabIndex        =   11
         Top             =   2880
         Width           =   7485
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
            Left            =   4650
            Picture         =   "CLASS ROOM DETAILS.frx":1E9608
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   300
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
            Left            =   3150
            Picture         =   "CLASS ROOM DETAILS.frx":1E9D98
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   300
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
            Left            =   180
            Picture         =   "CLASS ROOM DETAILS.frx":1EA304
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   300
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
            Left            =   1665
            Picture         =   "CLASS ROOM DETAILS.frx":1EAADE
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   300
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
            Left            =   6135
            Picture         =   "CLASS ROOM DETAILS.frx":1EAE95
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   300
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FF8080&
         Height          =   2535
         Left            =   60
         TabIndex        =   3
         Top             =   150
         Width           =   13455
         Begin MSDataGridLib.DataGrid DataGrid1 
            Bindings        =   "CLASS ROOM DETAILS.frx":1EB539
            Height          =   2115
            Left            =   5700
            TabIndex        =   18
            Top             =   330
            Width           =   4935
            _ExtentX        =   8705
            _ExtentY        =   3731
            _Version        =   393216
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
            ColumnCount     =   3
            BeginProperty Column00 
               DataField       =   "CLS_RMNO"
               Caption         =   "CLS_RMNO"
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
               DataField       =   "CLS_TYPE"
               Caption         =   "CLS_TYPE"
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
               DataField       =   "SEAT_CAP"
               Caption         =   "SEAT_CAP"
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
                  ColumnWidth     =   945.071
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   2429.858
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   959.811
               EndProperty
            EndProperty
         End
         Begin VB.ComboBox clsrm 
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
            ItemData        =   "CLASS ROOM DETAILS.frx":1EB54E
            Left            =   2580
            List            =   "CLASS ROOM DETAILS.frx":1EB550
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   1080
            Width           =   2655
         End
         Begin VB.TextBox seat 
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
            Height          =   405
            Left            =   2580
            TabIndex        =   1
            Top             =   1710
            Width           =   2565
         End
         Begin VB.ComboBox clstype 
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
            ItemData        =   "CLASS ROOM DETAILS.frx":1EB552
            Left            =   2580
            List            =   "CLASS ROOM DETAILS.frx":1EB55F
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   450
            Width           =   2655
         End
         Begin VB.Label Label7 
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
            Left            =   2160
            TabIndex        =   10
            Top             =   1110
            Width           =   135
         End
         Begin VB.Label Label6 
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
            Left            =   2070
            TabIndex        =   9
            Top             =   1710
            Width           =   135
         End
         Begin VB.Label Label5 
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
            Left            =   1920
            TabIndex        =   8
            Top             =   420
            Width           =   135
         End
         Begin VB.Label Label4 
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "SEAT CAPACITY :"
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
            Left            =   840
            TabIndex        =   6
            Top             =   1800
            Width           =   1575
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "TYPES :"
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
            Left            =   1320
            TabIndex        =   5
            Top             =   510
            Width           =   735
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "CLASS ROOM NO :"
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
            Left            =   870
            TabIndex        =   4
            Top             =   1170
            Width           =   1455
         End
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "CLASS ROOM"
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
      TabIndex        =   7
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "frmCLASSROOM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ref()
seat.Text = ""
End Sub
Private Sub clsrm_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
clstype.SetFocus
End If
End Sub

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

Private Sub clstype_KeyPress(KeyAscii As Integer)
If (keyacii = 13) Then
clsrm.SetFocus
End If
End Sub

Private Sub delete_Click()
If (clsrm.Text = "" Or seat.Text = "" Or clstype.Text = "") Then
frmINFO.show 1
If str = "OK" Then
clsrm.SetFocus
End If
Else
Module1.conn
frmDELETE.show 1
If str = "YES" Then
sql = "delete from ClassDetail_Master where cls_rmno='" + clsrm.Text + "'"
Set r = c.Execute(sql)
frmRECORDELETE.show 1
If str = "OK" Then
Call ref
Adodc1.refresh
clsrm.SetFocus
End If
ElseIf str = "NO" Then
clsrm.SetFocus
End If
End If
End Sub

Private Sub Form_Load()
delete.Enabled = False
update.Enabled = False
seat.MaxLength = 3
'clsr.Visible = False
End Sub

Private Sub print_Click()
If DataEnvironment1.rsprint1.state = 1 Then DataEnvironment1.rsprint1.Close
DataEnvironment1.print1 clsrm
DataReport1.show
End Sub

Private Sub refresh_Click()
Call ref
save.Enabled = True
delete.Enabled = False
update.Enabled = False
'clsr.Visible = False
'clsrm.Visible = True
End Sub
Private Sub save_Click()
save.Enabled = False
update.Enabled = True
delete.Enabled = True
On Error GoTo label
If (clsrm.Text = "" Or seat.Text = "" Or clstype.Text = "") Then
frmINFO.show 1
If str = "OK" Then
clsrm.SetFocus
End If
Else
Module1.conn
frmSAVE.show 1
If str = "YES" Then
sql = "insert into ClassDetail_Master values('" + clsrm.Text + "','" + clstype.Text + "'," + seat.Text + ")"
Set r = c.Execute(sql)
frmRECORDSAVE.show 1
If str = "OK" Then
Adodc1.refresh
Call ref
clsrm.SetFocus
End If
frmANOTHERECORD.show 1
If str = "YES" Then
save.Enabled = True
delete.Enabled = False
update.Enabled = False
ElseIf str = "NO" Then
save.Enabled = False
delete.Enabled = True
update.Enabled = True
End If
Exit Sub
label:
frmDATAEXIST.show 1
If str = "OK" Then
Call ref
clsrm.SetFocus
End If
ElseIf str = "NO" Then
clsrm.SetFocus
End If
End If
'End If
End Sub
Private Sub seat_keypress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
seat.Locked = False
Else
seat.Locked = True
End If
End Sub
Private Sub update_Click()
If (clsrm.Text = "" Or seat.Text = "" Or clstype.Text = "") Then
frmINFO.show 1
If str = "OK" Then
clsrm.SetFocus
End If
Else
Module1.conn
frmUPDATE.show 1
If str = "YES" Then
sql = "update ClassDetail_Master set cls_type='" + clstype.Text + "',seat_cap=" + seat.Text + " where cls_rmno='" + clsrm.Text + "'"
Set r = c.Execute(sql)
Adodc1.refresh
frmRECORDUPDATE.show 1
If str = "OK" Then
Call ref
clsrm.SetFocus
End If
ElseIf str = "NO" Then
clsrm.SetFocus
End If
End If
End Sub
Private Sub view_Click()
save.Enabled = False
update.Enabled = True
delete.Enabled = True
On Error GoTo label
If (clsrm.Text = "") Then
frmINFO.Label2.Caption = "PLEASE ENTER CLASS NO"
frmINFO.show 1
If str = "OK" Then
clsrm.SetFocus
End If
Else
Module1.conn
sql = "select *from ClassDetail_Master where cls_rmno='" & clsrm.Text & "'"
Set r = c.Execute(sql)
seat.Text = r.Fields(2)
End If
Exit Sub
label:
frmNOTFOUND.show 1
If str = "OK" Then
clsrm.SetFocus
End If
End Sub
