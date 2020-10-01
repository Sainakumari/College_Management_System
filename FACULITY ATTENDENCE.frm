VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmFACULTYATTENDENCE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FACULTY"
   ClientHeight    =   8730
   ClientLeft      =   3015
   ClientTop       =   1860
   ClientWidth     =   18300
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   Picture         =   "FACULITY ATTENDENCE.frx":0000
   ScaleHeight     =   8730
   ScaleWidth      =   18300
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   150
      Top             =   180
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
      RecordSource    =   "select *from facultyattend"
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
      Height          =   7845
      Left            =   1320
      TabIndex        =   1
      Top             =   600
      Width           =   15885
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
         Left            =   4770
         Picture         =   "FACULITY ATTENDENCE.frx":1E9608
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   6480
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
         Left            =   2310
         Picture         =   "FACULITY ATTENDENCE.frx":1E99BF
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   6480
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
         Left            =   7290
         Picture         =   "FACULITY ATTENDENCE.frx":1EA199
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   6480
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
         Left            =   9765
         Picture         =   "FACULITY ATTENDENCE.frx":1EA705
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   6480
         Width           =   1095
      End
      Begin VB.CommandButton refresh 
         Caption         =   "REGRESH"
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
         Left            =   12270
         Picture         =   "FACULITY ATTENDENCE.frx":1EAE95
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   6480
         Width           =   1095
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FF8080&
         Caption         =   "ATTENDENCE"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3675
         Left            =   150
         TabIndex        =   2
         Top             =   150
         Width           =   15495
         Begin VB.TextBox facid 
            BackColor       =   &H80000016&
            Height          =   465
            Left            =   8730
            TabIndex        =   27
            Top             =   300
            Width           =   2265
         End
         Begin VB.ComboBox month 
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
            ItemData        =   "FACULITY ATTENDENCE.frx":1EB539
            Left            =   8760
            List            =   "FACULITY ATTENDENCE.frx":1EB561
            TabIndex        =   24
            Text            =   "SELECT"
            Top             =   1080
            Width           =   2295
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00FF8080&
            Caption         =   "FACULTY'PHOTO"
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
            Left            =   11520
            TabIndex        =   14
            Top             =   240
            Width           =   2415
            Begin VB.Image Image1 
               BorderStyle     =   1  'Fixed Single
               Height          =   1575
               Left            =   240
               Stretch         =   -1  'True
               Top             =   240
               Width           =   1935
            End
         End
         Begin VB.TextBox year 
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
            Top             =   1080
            Width           =   2295
         End
         Begin VB.TextBox per 
            BackColor       =   &H80000016&
            Height          =   495
            Left            =   6510
            TabIndex        =   12
            Top             =   2970
            Width           =   2295
         End
         Begin VB.TextBox total 
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
            Top             =   1950
            Width           =   2295
         End
         Begin VB.TextBox present 
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
            Left            =   8760
            TabIndex        =   10
            Top             =   1800
            Width           =   2295
         End
         Begin VB.TextBox sno 
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
            TabIndex        =   9
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label Label12 
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
            Left            =   7800
            TabIndex        =   25
            Top             =   960
            Width           =   135
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "MONTH :"
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
            Left            =   7080
            TabIndex        =   23
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label10 
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
            Left            =   5910
            TabIndex        =   22
            Top             =   2970
            Width           =   135
         End
         Begin VB.Label Label6 
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
            Left            =   8160
            TabIndex        =   21
            Top             =   1800
            Width           =   135
         End
         Begin VB.Label Label3 
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
            Left            =   2160
            TabIndex        =   20
            Top             =   1920
            Width           =   135
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL NO OF DAY :"
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
            Left            =   720
            TabIndex        =   8
            Top             =   2040
            Width           =   1455
         End
         Begin VB.Label Label8 
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
            Left            =   780
            TabIndex        =   7
            Top             =   1200
            Width           =   615
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "FACULTY ID :"
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
            Left            =   7080
            TabIndex        =   6
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "PERCENTAGE :"
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
            TabIndex        =   5
            Top             =   3090
            Width           =   1215
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "PRESENT DAY :"
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
            Left            =   7020
            TabIndex        =   4
            Top             =   1920
            Width           =   1215
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "SERIAL NUMBER :"
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
            Left            =   720
            TabIndex        =   3
            Top             =   360
            Width           =   1335
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Bindings        =   "FACULITY ATTENDENCE.frx":1EB5C7
         Height          =   1935
         Left            =   690
         TabIndex        =   26
         Top             =   4050
         Width           =   14055
         _ExtentX        =   24791
         _ExtentY        =   3413
         _Version        =   393216
         BackColor       =   -2147483626
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Cambria"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H8000000B&
         Height          =   1455
         Left            =   630
         Top             =   6240
         Width           =   14175
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "FACULTY ATTENDENCE "
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
      Left            =   5400
      TabIndex        =   0
      Top             =   0
      Width           =   7335
   End
End
Attribute VB_Name = "frmFACULTYATTENDENCE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command6_Click()
cd4.Filter = "picture file |*.jpg"
cd4.ShowOpen
If cd4.FileName <> " " Then
Image1.Picture = LoadPicture(cd4.FileName)
End If
End Sub
Private Sub ref()
facid.Text = ""
Year.Text = ""
total.Text = ""
present.Text = ""
per.Text = ""
End Sub
Private Sub delete_Click()
If (sno.Text = "" Or facid.Text = "" Or Year.Text = "" Or Month.Text = "" Or total.Text = "" Or present.Text = "" Or per.Text = "") Then
A = MsgBox("ALL FIELDS ARE MANDATORY", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
sno.SetFocus
End If
Else
Module1.conn
A = MsgBox("CLICK OK TO DELETE RECORD", vbQuestion + vbOKCancel, "DELETE")
If (A = vbOK) Then
sql = "DELETE FROM facultyattend WHERE facid='" + facid.Text + "'"
Set r = c.Execute(sql)
A = MsgBox("RECORD DELETED", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
Adodc1.Refresh
Call ref
sno.SetFocus
End If
Else
If (A = vbCancel) Then
exmnm.SetFocus
End If
End If
End If
End Sub

Private Sub Form_Load()
sno.Locked = True
sno.MaxLength = 2
Module1.conn
sql = "select count(sno) from facultyattend order by rownum desc"
Set r = c.Execute(sql)
Dim i As Integer
i = r.Fields(0)
If (i < 1) Then
sno.Text = 1
End If
End Sub

Private Sub refresh_Click()
Call ref
End Sub

Private Sub save_Click()
If (sno.Text = "" Or facid.Text = "" Or Year.Text = "" Or Month.Text = "" Or total.Text = "" Or present.Text = "" Or per.Text = "") Then
A = MsgBox("ALL FIELDS ARE MANDATORY", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
sno.SetFocus
End If
Else
Module1.conn
A = MsgBox("CLICK OK TO SAVE RECORD", vbQuestion + vbOKCancel, "SAVE")
If (A = vbOK) Then
sql = "INSERT INTO FACULTYATTEND VALUES(" + sno.Text + ",'" + facid.Text + "'," + Year.Text + ",'" + Month.Text + "'," + total.Text + "," + present.Text + "," + per.Text + ")"
Set r = c.Execute(sql)
A = MsgBox("RECORD SAVED", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
n = sno.Text
n = n + 1
sno.Text = n
Adodc1.Refresh
Call ref
sno.SetFocus
End If
Else
If (A = vbCancel) Then
sno.SetFocus
End If
End If
End If
End Sub

Private Sub update_Click()
If (sno.Text = "" Or facid.Text = "" Or Year.Text = "" Or Month.Text = "" Or total.Text = "" Or present.Text = "" Or per.Text = "") Then
A = MsgBox("ALL FIELDS ARE MANDATORY", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
sno.SetFocus
End If
Else
Module1.conn
A = MsgBox("CLICK OK TO UPDATE RECORD", vbQuestion + vbOKCancel, "UPDATE")
If (A = vbOK) Then
sql = "UPDATE facultyattend WHERE facid='" + facid.Text + "'"
Set r = c.Execute(sql)
A = MsgBox("RECORD UPDATED", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
Adodc1.Refresh
Call ref
sno1.SetFocus
End If
Else
If (A = vbCancel) Then
sno1.SetFocus
End If
End If
End If
End Sub

Private Sub view_Click()
On Error GoTo label
Update.Enabled = True
Delete.Enabled = True
Save.Enabled = False
sql = "SELECT *FROM FACULTYATTEND WHERE facid='" + facid.Text + "'"
Set r = c.Execute(sql)
sno.Text = r.Fields(0)
facid.Text = r.Fields(1)
Year.Text = r.Fields(2)
Month.Text = r.Fields(3)
total.Text = r.Fields(4)
present.Text = r.Fields(5)
per.Text = r.Fields(6)
Exit Sub
label:
A = MsgBox("RECORD NOT FOUND", vbInformation + vbOKOnly)
If (A = vbOK) Then
sno.Text
End If
End Sub

