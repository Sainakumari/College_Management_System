VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmBATCH 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BATCH"
   ClientHeight    =   8955
   ClientLeft      =   705
   ClientTop       =   1035
   ClientWidth     =   19005
   LinkTopic       =   "Form12"
   MaxButton       =   0   'False
   Picture         =   "BATCH DETAILS.frx":0000
   ScaleHeight     =   8955
   ScaleWidth      =   19005
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   435
      Left            =   240
      Top             =   2220
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   767
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
      RecordSource    =   "select *from BatchDetail_Master"
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
      Height          =   4215
      Left            =   1560
      TabIndex        =   5
      Top             =   630
      Width           =   16725
      Begin VB.Frame Frame3 
         BackColor       =   &H80000002&
         Height          =   1215
         Left            =   1080
         TabIndex        =   15
         Top             =   2760
         Width           =   9465
         Begin VB.CommandButton PRINT 
            Caption         =   "Command1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   8040
            TabIndex        =   21
            Top             =   240
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
            Left            =   5070
            Picture         =   "BATCH DETAILS.frx":1E9608
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   180
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
            Left            =   3600
            Picture         =   "BATCH DETAILS.frx":1E9D98
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   180
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
            Left            =   720
            Picture         =   "BATCH DETAILS.frx":1EA304
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   180
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
            Left            =   2160
            Picture         =   "BATCH DETAILS.frx":1EAADE
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   150
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
            Left            =   6555
            Picture         =   "BATCH DETAILS.frx":1EAE95
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   180
            Width           =   1095
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "BATCH DETAILS.frx":1EB539
         Height          =   4035
         Left            =   10800
         TabIndex        =   14
         Top             =   150
         Width           =   5865
         _ExtentX        =   10345
         _ExtentY        =   7117
         _Version        =   393216
         BackColor       =   -2147483626
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
         BeginProperty Column01 
            DataField       =   "BATNO"
            Caption         =   "BATCH"
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
            DataField       =   "YEAR"
            Caption         =   "YEAR"
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
            DataField       =   "STRENGTH"
            Caption         =   "STRENGTH"
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
               ColumnWidth     =   1214.929
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FF8080&
         Height          =   2175
         Left            =   90
         TabIndex        =   6
         Top             =   150
         Width           =   10605
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
            Left            =   2700
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   480
            Width           =   2775
         End
         Begin VB.TextBox strength 
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
            Left            =   7650
            TabIndex        =   0
            Top             =   1080
            Width           =   2775
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
            ItemData        =   "BATCH DETAILS.frx":1EB54E
            Left            =   7650
            List            =   "BATCH DETAILS.frx":1EB550
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   390
            Width           =   2775
         End
         Begin VB.ComboBox batch 
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
            ItemData        =   "BATCH DETAILS.frx":1EB552
            Left            =   2700
            List            =   "BATCH DETAILS.frx":1EB554
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   1050
            Width           =   2775
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            BackColor       =   &H00FF8080&
            Caption         =   "COURSE"
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
            Left            =   960
            TabIndex        =   13
            Top             =   540
            Width           =   1155
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
            Left            =   6690
            TabIndex        =   12
            Top             =   360
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
            Left            =   7080
            TabIndex        =   11
            Top             =   1050
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
            TabIndex        =   10
            Top             =   450
            Width           =   165
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
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
            Left            =   6090
            TabIndex        =   9
            Top             =   450
            Width           =   735
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   " STRENGTH :"
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
            Left            =   6150
            TabIndex        =   8
            Top             =   1110
            Width           =   1095
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "BATCH NUMBER :"
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
            Left            =   1020
            TabIndex        =   7
            Top             =   1050
            Width           =   1335
         End
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BATCH"
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
      Left            =   5640
      TabIndex        =   2
      Top             =   0
      Width           =   7095
   End
End
Attribute VB_Name = "frmBATCH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub save_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
BorderStyle = 1
appeareance = 1
End Sub

Private Sub save_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
BorderStyle = 1
End Sub

Private Sub save_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
BorderStyle = 0
appeareance = 0
End Sub

Private Sub view_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
delete.SetFocus
End If
End Sub
Private Sub print_Click()
If DataEnvironment2.rsprnt.state = 1 Then DataEnvironment2.rsprnt.Close
DataEnvironment2.prnt year, batch
DataReport2.show
End Sub

Private Sub batch_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
year.SetFocus
End If
End Sub

Private Sub course_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
strength.SetFocus
End If
End Sub

'Private Sub Form_Activate()
'batch.SetFocus
'End Sub

Private Sub refresh_Click()
save.Enabled = True
strength.Text = ""
update.Enabled = False
delete.Enabled = False
End Sub

Private Sub delete_Click()
If (year.Text = "" Or batch.Text = "" Or strength.Text = "" Or course.Text = "") Then
frmINFO.show 1
If str = "OK" Then
batch.SetFocus
End If
Else
Module1.conn
frmDELETE.show 1
If str = "YES" Then
sql = "delete from BatchDetail_Master where course = '" + course.Text + "' and year='" + year.Text + "' AND batno=" + batch.Text + ""
Set r = c.Execute(sql)
frmRECORDELETE.show 1
If str = "OK" Then
Adodc1.refresh
strength.Text = ""
batch.SetFocus
End If
ElseIf str = "NO" Then
End If
End If
End Sub
Private Sub Form_Load()
update.Enabled = False
delete.Enabled = False
Module1.conn
sql = "select code from CourseDetail_Master"
Set r = c.Execute(sql)
While (r.EOF = False)
course.AddItem r.Fields("code")
r.MoveNext
Wend
strength.MaxLength = 2
Adodc1.Visible = False
batch.AddItem 1
batch.AddItem 2
batch.AddItem 3
batch.AddItem 4
year.AddItem "I"
year.AddItem "II"
year.AddItem "III"
End Sub

Private Sub save_Click()
On Error GoTo label
If (course.Text = "" Or batch.Text = "" Or year.Text = "" Or strength.Text = "") Then
frmINFO.show 1
If (str = "OK") Then
batch.SetFocus
End If
Else
Module1.conn
frmSAVE.show 1
If (str = "YES") Then
sql = "insert into BatchDetail_Master values('" + course.Text + "','" + year.Text + "'," + batch.Text + "," + strength.Text + ")"
Set r = c.Execute(sql)
frmRECORDSAVE.show 1
If (str = "OK") Then
Adodc1.refresh
strength.Text = ""
End If
frmANOTHERECORD.show 1
If (str = "YES") Then
save.Enabled = True
delete.Enabled = False
update.Enabled = False
ElseIf (str = "NO") Then
save.Enabled = False
delete.Enabled = True
update.Enabled = True
End If
Exit Sub
label:
frmDATAEXIST.show 1
If (str = "OK") Then
strength.Text = ""
batch.SetFocus
End If
Else
If (str = "NO") Then
batch.SetFocus
End If
End If
End If
End Sub

Private Sub save_KeyPress(KeyAscii As Integer)
If (keyacii = 13) Then
view.SetFocus
End If
End Sub

Private Sub strength_keypress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
strength.Locked = False
Else
strength.Locked = True
End If
If (KeyAscii = 13) Then
save.SetFocus
End If
End Sub

Private Sub strength_LostFocus()
On Error GoTo label
If (val(strength.Text) > 70) Then
A = MsgBox("STRENGTH SHOULD BE LESS THEN 70 OR 70", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
strength.Text = ""
strength.SetFocus
End If
End If
Exit Sub
label:
A = MsgBox("PLEASE ENTER DATA", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
course.SetFocus
End If
End Sub

Private Sub update_Click()
If (course.Text = "" Or batch.Text = "" Or year.Text = "" Or strength.Text = "") Then
frmINFO.show 1
If str = "OK" Then
batch.SetFocus
End If
Else
Module1.conn
frmUPDATE.show 1
If str = "YES" Then
sql = "update BatchDetail_Master set strength=" + strength.Text + " where course='" + course.Text + "' and year='" + year.Text + "' and batno=" + batch.Text + ""
Set r = c.Execute(sql)
Adodc1.refresh
frmRECORDUPDATE.show 1
If str = "OK" Then
strength.Text = ""
batch.SetFocus
End If
ElseIf str = "NO" Then
batch.SetFocus
End If
End If
End Sub

Private Sub view_Click()
save.Enabled = False
delete.Enabled = True
update.Enabled = True
On Error GoTo label
If (batch.Text = "" Or year.Text = "" Or course.Text = "") Then
A = MsgBox(" PLEASE ENTER COURSE, YEAR AND BATCH DETAIL", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
batch.SetFocus
End If
Else
Module1.conn
sql = "select *from BatchDetail_Master where  course = '" + course.Text + "' and year='" + year.Text + "' and  batno=" + batch.Text + ""
Set r = c.Execute(sql)
strength.Text = r.Fields(3)
Exit Sub
label:
A = MsgBox("RECORD NOT FOUND", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
batch.SetFocus
End If
End If
End Sub
Private Sub view_lostfocus()
course.Visible = True
End Sub

Private Sub delete_KeyPress(KeyAscii As Integer)
If (keyacii = 13) Then
update.SetFocus
End If
End Sub
Private Sub year_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
course.SetFocus
End If
End Sub
Private Sub update_KeyPress(KeyAscii As Integer)
If (keyacii = 13) Then
refresh.SetFocus
End If
End Sub
