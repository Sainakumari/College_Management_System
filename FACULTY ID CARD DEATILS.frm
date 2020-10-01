VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmFACULTYIDCARD 
   Caption         =   "FACULTY I.D CARD"
   ClientHeight    =   8820
   ClientLeft      =   2415
   ClientTop       =   1965
   ClientWidth     =   15570
   LinkTopic       =   "Form25"
   MaxButton       =   0   'False
   Picture         =   "FACULTY ID CARD DEATILS.frx":0000
   ScaleHeight     =   8820
   ScaleWidth      =   15570
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Height          =   7215
      Left            =   480
      TabIndex        =   0
      Top             =   840
      Width           =   14655
      Begin MSComDlg.CommonDialog cd1 
         Left            =   1200
         Top             =   5880
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton show 
         Caption         =   "SHOW DATA"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   10350
         Picture         =   "FACULTY ID CARD DEATILS.frx":1E9608
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   5850
         Width           =   1095
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FF8080&
         Height          =   2655
         Left            =   240
         TabIndex        =   11
         Top             =   2820
         Width           =   14295
         Begin VB.TextBox status 
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
            Left            =   9870
            TabIndex        =   22
            Top             =   930
            Width           =   2445
         End
         Begin VB.TextBox contact 
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
            Left            =   4830
            TabIndex        =   21
            Top             =   1590
            Width           =   2445
         End
         Begin VB.TextBox dept 
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
            Left            =   4830
            TabIndex        =   20
            Top             =   960
            Width           =   2445
         End
         Begin VB.TextBox nm 
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
            Left            =   4830
            TabIndex        =   19
            Top             =   300
            Width           =   2445
         End
         Begin VB.ComboBox facid 
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
            Left            =   9870
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   360
            Width           =   2445
         End
         Begin VB.Label Label6 
            BackColor       =   &H00FF8080&
            Caption         =   "STATUS:"
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
            Left            =   8430
            TabIndex        =   18
            Top             =   930
            Width           =   1185
         End
         Begin VB.Label Label4 
            BackColor       =   &H00FF8080&
            Caption         =   "CONTACT NO.:"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   3360
            TabIndex        =   17
            Top             =   1560
            Width           =   1185
         End
         Begin VB.Label Label5 
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
            Height          =   495
            Left            =   8460
            TabIndex        =   14
            Top             =   330
            Width           =   1215
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "DEPARTMENT :"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   3330
            TabIndex        =   13
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "NAME :"
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
            Left            =   3630
            TabIndex        =   12
            Top             =   360
            Width           =   765
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H8000000B&
            Height          =   2295
            Left            =   360
            Top             =   240
            Width           =   2535
         End
         Begin VB.Image Image1 
            BorderStyle     =   1  'Fixed Single
            Height          =   2055
            Left            =   480
            Stretch         =   -1  'True
            Top             =   360
            Width           =   2295
         End
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
         Left            =   4995
         Picture         =   "FACULTY ID CARD DEATILS.frx":1E9E57
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   5790
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
         Left            =   6900
         Picture         =   "FACULTY ID CARD DEATILS.frx":1EA20E
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   5790
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
         Left            =   8700
         Picture         =   "FACULTY ID CARD DEATILS.frx":1EA77A
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   5850
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
         Left            =   3120
         Picture         =   "FACULTY ID CARD DEATILS.frx":1EAE1E
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   5790
         Width           =   1095
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0C0&
         Height          =   2295
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   14295
         Begin VB.TextBox Text8 
            Alignment       =   2  'Center
            BackColor       =   &H80000016&
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   21.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2760
            Locked          =   -1  'True
            TabIndex        =   6
            Text            =   "ARCADE BUSINESS COLLEGE"
            Top             =   240
            Width           =   11175
         End
         Begin VB.TextBox Text9 
            Alignment       =   2  'Center
            BackColor       =   &H80000016&
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4080
            Locked          =   -1  'True
            TabIndex        =   5
            Text            =   "www.abcollege.org"
            Top             =   840
            Width           =   7455
         End
         Begin VB.TextBox Text10 
            Alignment       =   2  'Center
            BackColor       =   &H80000016&
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3360
            Locked          =   -1  'True
            TabIndex        =   4
            Text            =   "Rajendra Nagar,Patna-800016(Bihar)"
            Top             =   1200
            Width           =   9735
         End
         Begin VB.TextBox Text11 
            Alignment       =   2  'Center
            BackColor       =   &H80000016&
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5640
            Locked          =   -1  'True
            TabIndex        =   3
            Text            =   "[College Code:204]"
            Top             =   1680
            Width           =   4335
         End
         Begin VB.Image Image2 
            BorderStyle     =   1  'Fixed Single
            Height          =   2340
            Left            =   240
            Picture         =   "FACULTY ID CARD DEATILS.frx":1EB5F8
            Top             =   240
            Width           =   1980
         End
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H8000000B&
         Height          =   1455
         Left            =   3000
         Top             =   5580
         Width           =   8805
      End
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   135
      Left            =   840
      TabIndex        =   23
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "FACULTY I.D CARD"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   735
      Left            =   5040
      TabIndex        =   1
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "frmFACULTYIDCARD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub delete_Click()
If (facid.Text = "" Or nm.Text = "" Or dept.Text = "" Or status.Text = "" Or contact.Text = "") Then
A = MsgBox("ALL FIELDS ARE MANDATORY", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
facid.SetFocus
End If
Else
Module1.conn
A = MsgBox("CLICK OK TO DELETE RECORD", vbQuestion + vbOKCancel, "DELETE")
If (A = vbOK) Then
sql = "delete from facultyid_master where facid='" + facid.Text + "'"
Set r = c.Execute(sql)
A = MsgBox("RECORD DELETED", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
Call ref
facid.SetFocus
End If
Else
If (A = vbCancel) Then
facid.SetFocus
End If
End If
End If
End Sub

Private Sub facid_LostFocus()
'If (r.EOF = True) Then
'facid.SetFocus
'Else
Module1.conn
sql = "select stat,fac_fnm,fac_mnm,fac_snm,fac_con,dcode from facultydetail_master where fac_id = '" + facid.Text + "'"
Set r = c.Execute(sql)
If IsNull("fac_mnm") And IsNull(" fac_snm") Then
nm.Text = r.Fields("fac_fnm")
Else
nm.Text = r.Fields("fac_fnm") & " " & r.Fields("fac_mnm") & " " & r.Fields("fac_snm")
If (IsNull("fac_mnm")) Then
nm.Text = r.Fields("fac_fnm") & r.Fields("fac_snm")
Else
If (IsNull("fac_snm")) Then
nm.Text = r.Fields("fac_fnm") & r.Fields("fac_mnm")
End If
End If
End If
contact.Text = r.Fields("fac_con")
status.Text = r.Fields("stat")
dept.Text = r.Fields("dcode")
'End If
End Sub

Private Sub Form_Load()
Module1.conn
sql = "select fac_id from facultydetail_master"
Set r = c.Execute(sql)
While (r.EOF = False)
facid.AddItem r.Fields("fac_id")
r.MoveNext
Wend
nm.Locked = True
dept.Locked = True
contact.Locked = True
status.Locked = True
End Sub
Private Sub refresh_Click()
Call ref
End Sub
Private Sub ref()
nm.Text = ""
dept.Text = ""
contact.Text = ""
status.Text = ""
Set Image1.Picture = Nothing
Image1.Picture = LoadPicture()
End Sub

Private Sub save_Click()
On Error GoTo label
If (facid.Text = "" Or nm.Text = "" Or dept.Text = "" Or status.Text = "" Or contact.Text = "") Then
A = MsgBox("ALL FIELDS ARE MANDATORY", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
facid.SetFocus
End If
Else
Module1.conn
A = MsgBox("CLICK OK TO SAVE RECORD", vbQuestion + vbOKCancel, "SAVE")
If (A = vbOK) Then
sql = "insert into facultyid_master values('" + facid.Text + "')"
Set r = c.Execute(sql)
A = MsgBox("RECORD SAVE", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
Call ref
facid.SetFocus
End If
Exit Sub
label:
A = MsgBox("DATA ALREADY EXIST", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
Call ref
facid.SetFocus
End If
Else
If (A = vbCancel) Then
facid.SetFocus
End If
End If
End If
End Sub

Private Sub show_Click()
frmFACULTYIDCARDDATA.SHOW
Unload Me
End Sub

Private Sub view_Click()
If (facid.Text = "") Then
A = MsgBox("PLEASE ENTER FACULTY ID", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
facid.SetFocus
End If
Else
Module1.conn
sql = "select fac_fnm,dcode,fac_con,stat,path from facultydetail_Master where fac_id='" + facid.Text + "'"
Set r = c.Execute(sql)
nm.Text = r.Fields("fac_fnm")
dept.Text = r.Fields("dcode")
contact.Text = r.Fields("fac_con")
status.Text = r.Fields("stat")
Label7.Caption = r.Fields("path")
Image1.Picture = LoadPicture(Label7.Caption)
End If
End Sub
