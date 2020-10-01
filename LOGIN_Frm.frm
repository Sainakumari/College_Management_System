VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   Picture         =   "LOGIN_Frm.frx":0000
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton forgetpass 
      BackColor       =   &H80000016&
      Caption         =   "FORGOT PASSWORD"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   14760
      Picture         =   "LOGIN_Frm.frx":2A3042
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8160
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2895
      Left            =   10560
      TabIndex        =   1
      Top             =   5280
      Width           =   6015
      Begin VB.CommandButton login 
         Caption         =   "LOGIN"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   840
         Picture         =   "LOGIN_Frm.frx":2A5892
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2040
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FF8080&
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
         Height          =   375
         Left            =   4800
         TabIndex        =   5
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton cancel 
         Caption         =   "CANCEL"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3240
         Picture         =   "LOGIN_Frm.frx":2A7EF6
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2040
         Width           =   1455
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
         IMEMode         =   3  'DISABLE
         Left            =   2520
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1320
         Width           =   2055
      End
      Begin VB.TextBox id 
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
         Left            =   2520
         TabIndex        =   2
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000011&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "         PASSWORD"
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
         Left            =   270
         TabIndex        =   8
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label label1 
         BackColor       =   &H80000011&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "          LOGIN ID"
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
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   1695
      End
   End
   Begin VB.CommandButton exit 
      Appearance      =   0  'Flat
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   18600
      Picture         =   "LOGIN_Frm.frx":2F0780
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   210
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1530
      Left            =   150
      Picture         =   "LOGIN_Frm.frx":2F2FAD
      Top             =   510
      Width           =   1830
   End
   Begin VB.Image Image2 
      Height          =   750
      Left            =   9960
      Picture         =   "LOGIN_Frm.frx":2F3491
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   4
      Height          =   975
      Left            =   14640
      Top             =   8160
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "THE WAY OF SUCCESS"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   13800
      TabIndex        =   11
      Top             =   1800
      Width           =   3735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "THE  COLLEGE MANAGEMENT SYSTEM"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1215
      Left            =   2520
      TabIndex        =   10
      Top             =   720
      Width           =   16935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cancel_keypress(KeyAscii As Integer)
If KeyAscii = 13 Then
forgetpass.SetFocus
End If
End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then
pass.PasswordChar = ""
Else
pass.PasswordChar = "*"
End If
End Sub
Private Sub EXIT_Click()
Unload Me
End Sub
Private Sub forgetpass_Click()
Unload Me
frmFORGETPASSWORD.SHOW
End Sub


Private Sub id_keypress(KeyAscii As Integer)
If (KeyAscii = 13) Then
pass.SetFocus
End If
End Sub

Private Sub login_Click()
On Error GoTo label
If (id.Text = "" Or pass.Text = "") Then
A = MsgBox("PLEASE ENTER USERID OR PASSWORD", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
id.SetFocus
End If
Else
Module1.conn
sql = "select userid,pass from Create_User where userid='" + id.Text + "'"
Set r = c.Execute(sql)
USER = r.Fields("userid")
passs = r.Fields("pass")
If (id.Text <> USER Or pass.Text <> passs) Then
A = MsgBox("PLEASE ENTER CORRECT PASSWORD", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
pass.Text = ""
pass.SetFocus
End If
Else
frmWELCOMESCREEN.SHOW
End If
End If
'If (r.EOF = True Or r.BOF = True) Then
'a = MsgBox("PLEASE ENTER CORRECT USERID", vbInformation + vbOKOnly, "INFORMATION")
'End If
'If (a = vbOK) Then
'id.SetFocus
'End If
Exit Sub
label:
A = MsgBox("PLEASE ENTER CORRECT USERID", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
id.Text = ""
id.SetFocus
End If
End Sub

Private Sub login_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cancel.SetFocus
End If
End Sub

Private Sub pass_keypress(KeyAscii As Integer)
If (KeyAscii = 13) Then
login.SetFocus
End If
End Sub

