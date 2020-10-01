VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmCOLLEGE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "COLLEGE"
   ClientHeight    =   8820
   ClientLeft      =   1200
   ClientTop       =   1365
   ClientWidth     =   18990
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   Picture         =   "college details.frx":0000
   ScaleHeight     =   8820
   ScaleWidth      =   18990
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   300
      Top             =   150
      Width           =   1200
      _ExtentX        =   2117
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
      RecordSource    =   "select *from CollegeDetail_Master"
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
      Height          =   6405
      Left            =   1590
      TabIndex        =   1
      Top             =   600
      Width           =   15855
      Begin VB.Frame Frame3 
         BackColor       =   &H80000002&
         Height          =   3405
         Left            =   13770
         TabIndex        =   27
         Top             =   1500
         Width           =   1485
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
            Left            =   150
            Picture         =   "college details.frx":1E9608
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   630
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
            Left            =   210
            Picture         =   "college details.frx":1E9944
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   2040
            Width           =   1095
         End
         Begin VB.Line Line1 
            X1              =   30
            X2              =   1500
            Y1              =   1830
            Y2              =   1830
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FF8080&
         Height          =   6165
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   13095
         Begin MSDataGridLib.DataGrid DataGrid1 
            Bindings        =   "college details.frx":1E9CFB
            Height          =   885
            Left            =   2220
            TabIndex        =   26
            Top             =   4740
            Width           =   9615
            _ExtentX        =   16960
            _ExtentY        =   1561
            _Version        =   393216
            AllowUpdate     =   0   'False
            HeadLines       =   1
            RowHeight       =   14
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
         Begin VB.TextBox addr 
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
            Height          =   675
            Left            =   3210
            MultiLine       =   -1  'True
            TabIndex        =   25
            Top             =   2550
            Width           =   2385
         End
         Begin VB.TextBox affilated 
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
            Left            =   9930
            TabIndex        =   24
            Top             =   2520
            Width           =   2295
         End
         Begin VB.TextBox clgcode 
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
            Left            =   9960
            TabIndex        =   17
            Top             =   240
            Width           =   2295
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
            Height          =   495
            Left            =   9930
            TabIndex        =   16
            Top             =   3270
            Width           =   2295
         End
         Begin VB.TextBox website 
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
            Left            =   3240
            TabIndex        =   15
            Top             =   3360
            Width           =   2295
         End
         Begin VB.TextBox contact 
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
            Left            =   9960
            TabIndex        =   14
            Top             =   1680
            Width           =   2295
         End
         Begin VB.TextBox dirnm 
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
            Left            =   3240
            TabIndex        =   13
            Top             =   1800
            Width           =   2295
         End
         Begin VB.TextBox clgnm 
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
            Left            =   3240
            TabIndex        =   12
            Top             =   1020
            Width           =   6675
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
            Height          =   495
            Left            =   3240
            TabIndex        =   11
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            BackColor       =   &H00FF8080&
            Caption         =   "AFFILATION:"
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
            Left            =   7530
            TabIndex        =   23
            Top             =   2640
            Width           =   1425
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
            Left            =   9240
            TabIndex        =   22
            Top             =   1680
            Width           =   135
         End
         Begin VB.Label Label13 
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
            Left            =   9600
            TabIndex        =   21
            Top             =   240
            Width           =   135
         End
         Begin VB.Label Label12 
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
            TabIndex        =   20
            Top             =   2520
            Width           =   135
         End
         Begin VB.Label Label11 
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
            TabIndex        =   19
            Top             =   1080
            Width           =   135
         End
         Begin VB.Label Label10 
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
            Left            =   2760
            TabIndex        =   18
            Top             =   240
            Width           =   135
         End
         Begin VB.Label Label9 
            BackColor       =   &H80000016&
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
            Height          =   375
            Left            =   7830
            TabIndex        =   10
            Top             =   3360
            Width           =   1545
         End
         Begin VB.Label Label8 
            BackColor       =   &H80000016&
            BackStyle       =   0  'Transparent
            Caption         =   "WEBSITE :"
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
            Left            =   840
            TabIndex        =   9
            Top             =   3360
            Width           =   1935
         End
         Begin VB.Label Label7 
            BackColor       =   &H80000016&
            BackStyle       =   0  'Transparent
            Caption         =   "CONTACT NUMBER :"
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
            Left            =   7680
            TabIndex        =   8
            Top             =   1800
            Width           =   1575
         End
         Begin VB.Label Label6 
            BackColor       =   &H80000016&
            BackStyle       =   0  'Transparent
            Caption         =   "COLLEGE CODE :"
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
            Left            =   7680
            TabIndex        =   7
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label5 
            BackColor       =   &H80000016&
            BackStyle       =   0  'Transparent
            Caption         =   "ADDRESS :"
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
            Left            =   840
            TabIndex        =   6
            Top             =   2610
            Width           =   975
         End
         Begin VB.Label Label4 
            BackColor       =   &H80000016&
            BackStyle       =   0  'Transparent
            Caption         =   "DIRECTOR NAME :"
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
            Left            =   840
            TabIndex        =   5
            Top             =   1860
            Width           =   1935
         End
         Begin VB.Label Label3 
            BackColor       =   &H80000016&
            BackStyle       =   0  'Transparent
            Caption         =   "COLLEGE NAME :"
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
            Left            =   840
            TabIndex        =   4
            Top             =   1110
            Width           =   1335
         End
         Begin VB.Label Label2 
            BackColor       =   &H80000016&
            BackStyle       =   0  'Transparent
            Caption         =   " REGISTRATION NUMBER :"
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
            Left            =   840
            TabIndex        =   3
            Top             =   360
            Width           =   1935
         End
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "COLLEGE DETAILS"
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
      TabIndex        =   0
      Top             =   0
      Width           =   6615
   End
End
Attribute VB_Name = "frmCOLLEGE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
contact.MaxLength = 10
regno.MaxLength = 10
clgcode.MaxLength = 3
clgnm.Text = ""
dirnm.Text = ""
emailid.Text = ""
contact.Text = ""
clgcode.Text = ""
addr.Text = ""
website.Text = ""
regno.Text = ""
Module1.conn
sql = "select count(clg_code) from CollegeDetail_Master"
Set r = c.Execute(sql)
Dim i As Integer
i = r.Fields(0)
If (i < 1) Then
save.Caption = "SAVE"
Else
save.Caption = "UPDATE"
clgcode.Locked = True
regno.Locked = True
End If
End Sub
Private Sub print_Click()
DataReport3.SHOW
End Sub

Private Sub save_Click()
'On Error GoTo xx
If (clgcode.Text = "" Or regno.Text = "" Or dirnm.Text = "" Or addr.Text = "" Or contact.Text = "" Or website.Text = "" Or emailid.Text = "" Or affilated.Text = "" Or clgnm.Text = "") Then
frmINFO.SHOW 1
If str = "OK" Then
regno.SetFocus
End If
Else
Set c = New ADODB.Connection
c.Open "Provider=MSDAORA.1;User ID=cms/saina;Persist Security Info=true"
'Set r = New ADODB.Recordset
sql = "select count(clg_code) from CollegeDetail_Master"
Set r = c.Execute(sql)
Dim i As Integer
i = r.Fields(0)
If (i < 1) Then
frmSAVE.SHOW 1
If str = "YES" Then
sql = "insert into CollegeDetail_Master values(" + regno.Text + ",'" + clgnm.Text + "','" + dirnm.Text + "','" + addr.Text + "'," + clgcode.Text + "," + contact.Text + ",'" + website.Text + "','" + emailid.Text + "','" + affilated.Text + "')"
Set r = c.Execute(sql)
clgnm.Text = ""
clgcode.Text = ""
dirnm.Text = ""
website.Text = ""
contact.Text = ""
emailid.Text = ""
addr.Text = ""
affilated.Text = ""
regno.Text = ""
Adodc1.refresh
frmRECORDSAVE.SHOW 1
If str = "OK" Then
save.Caption = "UPDATE"
End If
ElseIf str = "NO" Then
regno.SetFocus
End If
Else
frmUPDATE.SHOW 1
If str = "YES" Then
Set r = New ADODB.Recordset
sql = "update CollegeDetail_Master set clg_nm='" + clgnm.Text + "',clg_dir='" + dirnm.Text + "',clg_addr='" + addr.Text + "',clg_num=" + contact.Text + ",clg_web='" + website.Text + "',clg_email='" + emailid.Text + "',clg_affl='" + affilated.Text + "'"
Set r = c.Execute(sql)
Adodc1.refresh
frmRECORDUPDATE.SHOW 1
If str = "OK" Then
clgnm.Text = ""
clgcode.Text = ""
dirnm.Text = ""
website.Text = ""
contact.Text = ""
emailid.Text = ""
addr.Text = ""
affilated.Text = ""
regno.Text = ""
'Exit Sub
'xx:
'MsgBox "RECORD NOT UPDATE"
'save.Caption = "update"
End If
End If
End If
End If
End Sub

Private Sub clgcode_keypress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
clgcode.Locked = False
Else
clgcode.Locked = True
End If
If KeyAscii = 13 Then
regno.SetFocus
End If
End Sub

Private Sub clgnm_keypress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 90 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32) Then
clgnm.Locked = False
Else
clgnm.Locked = True
End If
If KeyAscii = 13 Then
dirnm.SetFocus
End If
End Sub

Private Sub dirnm_keypress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 90 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32) Then
dirnm.Locked = False
Else
dirnm.Locked = True
End If
If (KeyAscii = 13) Then
contact.SetFocus
End If
End Sub
Private Sub clgnm_lostfocus()
clgnm.Text = UCase(clgnm.Text)
End Sub

Private Sub addr_LostFocus()
addr.Text = UCase(addr.Text)
End Sub
Private Sub dirnm_lostfocus()
dirnm.Text = UCase(dirnm.Text)
End Sub
Private Sub contact_keypress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
contact.Locked = False
Else
contact.Locked = True
End If
If (KeyAscii = 32) Then
contact.Locked = True
End If
If (KeyAscii = 13) Then
addr.SetFocus
End If
End Sub
Private Sub regno_keypress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
regno.Locked = False
Else
regno.Locked = True
End If
If (KeyAscii = 13) Then
clgnm.SetFocus
End If
End Sub
Private Sub website_keypress(KeyAscii As Integer)
If (KeyAscii = 13) Then
affilated.SetFocus
End If
End Sub

Private Sub affilated_keypress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 90 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32) Then
affilated.Locked = False
Else
affilated.Locked = True
End If
If (KeyAscii = 13) Then
affilated.SetFocus
End If
End Sub

Private Sub affilated_LostFocus()
affilated.Text = UCase(affilated.Text)
End Sub

Private Sub view_Click()
Module1.conn
sql = "select *from CollegeDetail_Master"
Set r = c.Execute(sql)
regno.Text = r.Fields(0)
clgnm.Text = r.Fields(1)
dirnm.Text = r.Fields(2)
addr.Text = r.Fields(3)
clgcode.Text = r.Fields(4)
contact.Text = r.Fields(5)
website.Text = r.Fields(6)
emailid.Text = r.Fields(7)
affilated.Text = r.Fields(8)
End Sub

