VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCREATEUSER 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CREATE USER"
   ClientHeight    =   8130
   ClientLeft      =   3015
   ClientTop       =   1365
   ClientWidth     =   15585
   LinkTopic       =   "Form20"
   MaxButton       =   0   'False
   Picture         =   "CREATE USER DETAILS.frx":0000
   ScaleHeight     =   8130
   ScaleWidth      =   15585
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Height          =   7095
      Left            =   840
      TabIndex        =   11
      Top             =   720
      Width           =   13695
      Begin VB.Frame Frame2 
         BackColor       =   &H00FF8080&
         Height          =   6855
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   13455
         Begin VB.CommandButton Command1 
            Caption         =   "DATA GRID"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   765
            Left            =   10710
            Picture         =   "CREATE USER DETAILS.frx":1E9608
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   5190
            Width           =   1035
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H80000002&
            Height          =   1425
            Left            =   1620
            TabIndex        =   33
            Top             =   4770
            Width           =   8745
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
               Height          =   1095
               Left            =   90
               Picture         =   "CREATE USER DETAILS.frx":1E9E57
               Style           =   1  'Graphical
               TabIndex        =   38
               Top             =   210
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
               Height          =   1095
               Left            =   1905
               Picture         =   "CREATE USER DETAILS.frx":1EA631
               Style           =   1  'Graphical
               TabIndex        =   37
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
               Height          =   1095
               Left            =   3735
               Picture         =   "CREATE USER DETAILS.frx":1EA9E8
               Style           =   1  'Graphical
               TabIndex        =   36
               Top             =   210
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
               Height          =   1095
               Left            =   7320
               MaskColor       =   &H0080C0FF&
               Picture         =   "CREATE USER DETAILS.frx":1EB178
               Style           =   1  'Graphical
               TabIndex        =   35
               Top             =   210
               UseMaskColor    =   -1  'True
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
               Height          =   1095
               Left            =   5340
               Picture         =   "CREATE USER DETAILS.frx":1EB81C
               Style           =   1  'Graphical
               TabIndex        =   34
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.TextBox dob1 
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
            Height          =   375
            Left            =   8610
            TabIndex        =   7
            Top             =   2940
            Width           =   2535
         End
         Begin VB.TextBox ans 
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
            Height          =   375
            Left            =   8640
            TabIndex        =   3
            Top             =   1458
            Width           =   3975
         End
         Begin VB.TextBox cnfrmpass 
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
            Height          =   375
            Left            =   8640
            TabIndex        =   9
            Top             =   3720
            Width           =   2535
         End
         Begin VB.ComboBox question 
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
            ItemData        =   "CREATE USER DETAILS.frx":1EBD88
            Left            =   8640
            List            =   "CREATE USER DETAILS.frx":1EBD9B
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   720
            Width           =   3975
         End
         Begin MSComCtl2.DTPicker dob 
            Height          =   375
            Left            =   8610
            TabIndex        =   21
            Top             =   2940
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   661
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
            CustomFormat    =   "dd/MMM/yyyy"
            Format          =   89063427
            CurrentDate     =   43109
         End
         Begin VB.ComboBox gender 
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
            ItemData        =   "CREATE USER DETAILS.frx":1EBE6F
            Left            =   8640
            List            =   "CREATE USER DETAILS.frx":1EBE79
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   2196
            Width           =   2535
         End
         Begin VB.ComboBox status 
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
            ItemData        =   "CREATE USER DETAILS.frx":1EBE8B
            Left            =   3240
            List            =   "CREATE USER DETAILS.frx":1EBE95
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   720
            Width           =   2535
         End
         Begin VB.TextBox pass 
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
            Height          =   375
            Left            =   3240
            TabIndex        =   8
            Top             =   3720
            Width           =   2535
         End
         Begin VB.TextBox userid 
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
            Height          =   375
            Left            =   3240
            TabIndex        =   6
            Top             =   2955
            Width           =   2535
         End
         Begin VB.TextBox lnm 
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
            Height          =   375
            Left            =   3240
            TabIndex        =   4
            Top             =   2175
            Width           =   2535
         End
         Begin VB.TextBox fnm 
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
            Height          =   375
            Left            =   3240
            TabIndex        =   2
            Top             =   1410
            Width           =   2535
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
            Left            =   7800
            TabIndex        =   32
            Top             =   600
            Width           =   135
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "ANSWER :"
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
            Left            =   6120
            TabIndex        =   31
            Top             =   1481
            Width           =   975
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "SECRET QUESTION :"
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
            Left            =   6120
            TabIndex        =   30
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label Label17 
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
            Left            =   7920
            TabIndex        =   29
            Top             =   3720
            Width           =   135
         End
         Begin VB.Label Label16 
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
            Left            =   7560
            TabIndex        =   28
            Top             =   2880
            Width           =   135
         End
         Begin VB.Label Label15 
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
            Left            =   6960
            TabIndex        =   27
            Top             =   2040
            Width           =   135
         End
         Begin VB.Label Label14 
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
            Left            =   1920
            TabIndex        =   26
            Top             =   3600
            Width           =   135
         End
         Begin VB.Label Label13 
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
            Left            =   1560
            TabIndex        =   25
            Top             =   2880
            Width           =   135
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
            Left            =   1800
            TabIndex        =   24
            Top             =   2040
            Width           =   135
         End
         Begin VB.Label Label11 
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
            Left            =   1800
            TabIndex        =   23
            Top             =   1320
            Width           =   135
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
            Left            =   1680
            TabIndex        =   22
            Top             =   600
            Width           =   135
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "STATUS :"
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
            Left            =   720
            TabIndex        =   20
            Top             =   720
            Width           =   975
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00FFC0C0&
            BorderWidth     =   8
            Height          =   6375
            Left            =   480
            Shape           =   4  'Rounded Rectangle
            Top             =   240
            Width           =   12375
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "DATE OF BIRTH :"
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
            Left            =   6120
            TabIndex        =   19
            Top             =   3003
            Width           =   1455
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "GENDER :"
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
            Left            =   6120
            TabIndex        =   18
            Top             =   2242
            Width           =   855
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "CONFIRM PASSWORD :"
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
            Left            =   6120
            TabIndex        =   17
            Top             =   3765
            Width           =   1815
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "PASSWORD :"
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
            Left            =   720
            TabIndex        =   16
            Top             =   3720
            Width           =   1215
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "USER ID :"
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
            Left            =   720
            TabIndex        =   15
            Top             =   2940
            Width           =   855
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "LAST NAME :"
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
            Left            =   720
            TabIndex        =   14
            Top             =   2160
            Width           =   1095
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "FIRST NAME :"
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
            Left            =   720
            TabIndex        =   13
            Top             =   1380
            Width           =   1095
         End
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "CREATE USER"
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
      Left            =   6120
      TabIndex        =   10
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "frmCREATEUSER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ref()
update.Enabled = False
delete.Enabled = False
dob1.Visible = False
dob.Visible = True
fnm.Text = ""
lnm.Text = ""
userid.Text = ""
pass.Text = ""
cnfrmpass.Text = ""
ans.Text = ""
End Sub

Private Sub Command1_Click()
frmUSERDATA.SHOW
Unload Me
End Sub

Private Sub delete_Click()
If (status.Text = "" Or question.Text = "" Or fnm.Text = "" Or lnm.Text = "" Or ans.Text = "" Or gender.Text = "" Or userid.Text = "" Or pass.Text = "" Or cnfrmpass.Text = "") Then
frmINFO.SHOW 1
If (str = "OK") Then
status.SetFocus
End If
Else
Module1.conn
frmDELETE.SHOW 1
If (str = "YES") Then
sql = "delete from Create_User where userid='" + userid.Text + "'"
Set r = c.Execute(sql)
frmRECORDELETE.SHOW 1
If (str = "OK") Then
Call ref
status.SetFocus
End If
ElseIf (str = "NO") Then
status.SetFocus
End If
End If
End Sub

Private Sub Form_Load()
dob1.Visible = False
update.Enabled = False
delete.Enabled = False
End Sub

Private Sub refresh_Click()
save.Enabled = True
Call ref
End Sub

Private Sub save_Click()
If (status.Text = "" Or question.Text = "" Or fnm.Text = "" Or lnm.Text = "" Or ans.Text = "" Or gender.Text = "" Or userid.Text = "" Or pass.Text = "" Or cnfrmpass.Text = "") Then
frmINFO.SHOW 1
If str = "OK" Then
status.SetFocus
End If
Else
Module1.conn
frmSAVE.SHOW 1
If str = "YES" Then
sql = "insert into Create_User values('" + status.Text + "','" + question.Text + "','" + fnm.Text + "','" + ans.Text + "','" + lnm.Text + "','" + gender.Text + "','" + userid.Text + "','" + Format(dob.Value, "dd/MMM/yyyy") + "','" + pass.Text + "','" + cnfrmpass.Text + "')"
Set r = c.Execute(sql)
frmRECORDSAVE.SHOW 1
If (str = OK) Then
Call ref
status.SetFocus
End If
Else
If (str = "NO") Then
status.SetFocus
End If
End If
End If
End Sub

Private Sub update_Click()
If (status.Text = "" Or question.Text = "" Or fnm.Text = "" Or lnm.Text = "" Or ans.Text = "" Or gender.Text = "" Or userid.Text = "" Or pass.Text = "" Or cnfrmpass.Text = "") Then
frmINFO.SHOW 1
If str = "OK" Then
status.SetFocus
End If
Else
Module1.conn
frmUPDATE.SHOW 1
If str = "YES" Then
sql = "update Create_User set status='" + status.Text + "',question='" + question.Text + "',fnm='" + fnm.Text + "',answer='" + ans.Text + "',lnm='" + lnm.Text + "',gender='" + gender.Text + "',dob='" + dob1.Text + "'"
Set r = c.Execute(sql)
frmRECORDUPDATE.SHOW 1
If str = "OK" Then
Call ref
status.SetFocus
End If
Else
If (str = "NO") Then
status.SetFocus
End If
End If
End If
End Sub

Private Sub view_Click()
update.Enabled = True
delete.Enabled = True
save.Enabled = False
dob1.Visible = True
dob.Visible = False
On Error GoTo label
Module1.conn
frmUSERINPUT.SHOW 1
If str = "ok" Then
a = frmUSERINPUT.val
userid.Text = a
sql = "select *from Create_User where userid='" + userid.Text + "'"
Set r = c.Execute(sql)
status.Text = r.Fields(0)
question.Text = r.Fields(1)
fnm.Text = r.Fields(2)
ans.Text = r.Fields(3)
lnm.Text = r.Fields(4)
gender.Text = r.Fields(5)
userid.Text = r.Fields(6)
dob1.Text = r.Fields(7)
pass.Text = r.Fields(8)
cnfrmpass.Text = r.Fields(9)
ElseIf str = "cancel" Then
status.SetFocus
Exit Sub
label:
frmNOTFOUND.SHOW 1
If str = "OK" Then
Call ref
status.SetFocus
End If
End If
End Sub
