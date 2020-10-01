VERSION 5.00
Begin VB.Form frmFORGETPASSWORD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FORGET PASSWORD"
   ClientHeight    =   8835
   ClientLeft      =   2025
   ClientTop       =   1695
   ClientWidth     =   16125
   LinkTopic       =   "Form18"
   MaxButton       =   0   'False
   Picture         =   "FORGET PASSWORD.frx":0000
   ScaleHeight     =   8835
   ScaleWidth      =   16125
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Height          =   8415
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   15735
      Begin VB.Frame Frame2 
         BackColor       =   &H00FF8080&
         Height          =   8175
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   15495
         Begin VB.TextBox userid 
            BackColor       =   &H80000016&
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6120
            TabIndex        =   20
            Top             =   2340
            Width           =   3735
         End
         Begin VB.CommandButton back 
            Caption         =   "BACK TO LOGIN SCREEN"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6600
            TabIndex        =   17
            Top             =   7200
            Width           =   2295
         End
         Begin VB.CheckBox Check2 
            BackColor       =   &H00FF8080&
            Caption         =   "SHOW"
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
            Left            =   10080
            TabIndex        =   16
            Top             =   4920
            Width           =   855
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00FF8080&
            Caption         =   "SHOW"
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
            Left            =   10080
            TabIndex        =   15
            Top             =   4320
            Width           =   855
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
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   6120
            PasswordChar    =   "*"
            TabIndex        =   14
            Top             =   4920
            Width           =   3735
         End
         Begin VB.TextBox newpass 
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
            IMEMode         =   3  'DISABLE
            Left            =   6120
            PasswordChar    =   "*"
            TabIndex        =   13
            ToolTipText     =   "PASSWORD MUST CONTAIN AT LEAST 6 CHARACTERS,INCLUDING UPPER/LOWER AND NUMBER "
            Top             =   4320
            Width           =   3735
         End
         Begin VB.CommandButton submit 
            Caption         =   "SUBMIT"
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
            Left            =   6960
            Picture         =   "FORGET PASSWORD.frx":1E9608
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   5760
            Width           =   1575
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
            ItemData        =   "FORGET PASSWORD.frx":1EC57D
            Left            =   6120
            List            =   "FORGET PASSWORD.frx":1EC590
            TabIndex        =   8
            Text            =   "SELECT"
            Top             =   3000
            Width           =   3735
         End
         Begin VB.TextBox answer 
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
            Left            =   6120
            TabIndex        =   6
            Top             =   3600
            Width           =   3735
         End
         Begin VB.Label Label10 
            BackColor       =   &H00FF8080&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   11040
            TabIndex        =   19
            Top             =   4920
            Width           =   615
         End
         Begin VB.Label Label6 
            BackColor       =   &H00FF8080&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   11040
            TabIndex        =   18
            Top             =   4320
            Width           =   615
         End
         Begin VB.Label Label9 
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
            Height          =   315
            Left            =   3840
            TabIndex        =   12
            Top             =   4920
            Width           =   1815
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "NEW PASSWORD :"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4200
            TabIndex        =   11
            Top             =   4320
            Width           =   1455
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   $"FORGET PASSWORD.frx":1EC664
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
            Left            =   1320
            TabIndex        =   10
            Top             =   720
            Width           =   12855
         End
         Begin VB.Label Label4 
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
            Height          =   315
            Left            =   4320
            TabIndex        =   7
            Top             =   3000
            Width           =   1575
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   " ANSWER :"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4560
            TabIndex        =   5
            Top             =   3600
            Width           =   1095
         End
         Begin VB.Label Label2 
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
            Height          =   315
            Left            =   4500
            TabIndex        =   4
            Top             =   2400
            Width           =   1455
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FFFFFF&
            BorderColor     =   &H00FFC0C0&
            BorderWidth     =   8
            FillColor       =   &H00C000C0&
            Height          =   6735
            Left            =   3240
            Shape           =   2  'Oval
            Top             =   1080
            Width           =   8655
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "FORGET PASSWORD"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   5760
            TabIndex        =   3
            Top             =   240
            Width           =   3735
         End
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "FORGET PASSWORD"
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
      Left            =   5040
      TabIndex        =   1
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "frmFORGETPASSWORD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = 1 Then
newpass.PasswordChar = ""
Else
newpass.PasswordChar = "*"
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
cnfrmpass.PasswordChar = ""
Else
cnfrmpass.PasswordChar = "*"
End If
End Sub

Private Sub newpass_lostfocus()
l = Len(newpass.Text)
If (l < 8) Then
Label6.Caption = "poor"
Label6.BackColor = vbRed
'newpass.SetFocus
Else
If (l > 8 And l < 10) Then
Label6.Caption = "fair"
Label6.BackColor = vbYellow
Else
If (l > 10) Then
Label6.Caption = "strong"
Label6.BackColor = vbGreen
End If
End If
End If
End Sub
Private Sub cnfrmpass_lostfocus()
l = Len(cnfrmpass.Text)
If (l < 8) Then
Label10.Caption = "poor"
Label10.BackColor = vbRed
'cnfrmpass.SetFocus
Else
If (l > 8 And l < 10) Then
Label10.Caption = "fair"
Label10.BackColor = vbYellow
Else
If (l > 10) Then
Label10.Caption = "strong"
Label10.BackColor = vbGreen
End If
End If
End If
End Sub


Private Sub submit_Click()
If (userid.Text = "" Or question.Text = "" Or answer.Text = "" Or newpass.Text = "" Or cnfrmpass.Text = "") Then
a = MsgBox("ALL FIELDS ARE MANDATORY", vbInformation + vbOKOnly, "INFORMATION")
If (a = vbOK) Then
userid.SetFocus
End If
Else
Module1.conn
sql = "select userid,question,answer from Create_User where userid='" & userid.Text & "'"
Set r = c.Execute(sql)
User = r.Fields("userid")
MsgBox User
If (userid.Text <> User) Then
a = MsgBox("USERID IS NOT CORRECT", vbInformation + vbOKOnly, "INFORMATION")
If (a = vbOK) Then
userid.Text = ""
userid.SetFocus
End If
Else
ques = r.Fields(1)
If (question.Text <> ques) Then
a = MsgBox("QUESTION IS NOT CORRECT", vbInformation + vbOKOnly, "INFORMATION")
If (a = vbOK) Then
question.Text = ""
question.SetFocus
End If
Else
ans = r.Fields("answer")
If (answer.Text <> ans) Then
a = MsgBox("ANSWER IS NOT CORRECT", vbInformation + vbOKOnly, "INFORMATION")
If (a = vbOK) Then
answer.Text = ""
answer.SetFocus
End If
Else
sql = "update Create_User set pass='" + newpass.Text + "',cnfrm_pass='" + cnfrmpass.Text + "'"
Set r = c.Execute(sql)
a = MsgBox("PASSWORD IS SET", vbInformation + vbOKOnly, "INFORMATION")
If (a = vbOK) Then
userid.Text = ""
answer.Text = ""
newpass.Text = ""
cnfrmpass.Text = ""
userid.SetFocus
End If
End If
End If
End If
End If
If (newpass.Text <> cnfrmpass.Text) Then
MsgBox "new password and confirm password are not same"
cnfrmpass.Text = ""
cnfrmpass.SetFocus
Else
submit.SetFocus
End If
End Sub
