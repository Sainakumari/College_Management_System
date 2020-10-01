VERSION 5.00
Begin VB.Form frmRESET 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RESET PASSWORD"
   ClientHeight    =   8025
   ClientLeft      =   1695
   ClientTop       =   1365
   ClientWidth     =   16665
   LinkTopic       =   "Form19"
   MaxButton       =   0   'False
   Picture         =   "RE-SET PASSWORD DETAILS.frx":0000
   ScaleHeight     =   8025
   ScaleWidth      =   16665
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Height          =   6975
      Left            =   2640
      TabIndex        =   1
      Top             =   720
      Width           =   11415
      Begin VB.Frame Frame2 
         BackColor       =   &H00FF8080&
         Height          =   6495
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   11175
         Begin VB.TextBox currpass 
            BackColor       =   &H80000016&
            Height          =   375
            Left            =   4320
            TabIndex        =   17
            Top             =   2250
            Width           =   3255
         End
         Begin VB.OptionButton Option2 
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
            Left            =   7680
            TabIndex        =   16
            Top             =   3480
            Width           =   855
         End
         Begin VB.OptionButton Option1 
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
            Left            =   7680
            TabIndex        =   15
            Top             =   2880
            Width           =   855
         End
         Begin VB.CommandButton cancel 
            Caption         =   "CANCEL"
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
            Left            =   6240
            Picture         =   "RE-SET PASSWORD DETAILS.frx":1E9608
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   4440
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
            Left            =   4080
            Picture         =   "RE-SET PASSWORD DETAILS.frx":231E92
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   4440
            Width           =   1095
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
            Left            =   4320
            TabIndex        =   9
            Top             =   3480
            Width           =   3255
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
            Height          =   375
            Left            =   4320
            TabIndex        =   8
            ToolTipText     =   "PASSWORD MUST CONTAIN AT LEAST 6 CHARACTERS,INCLUDING UPPER/LOWER AND NUMBER "
            Top             =   2880
            Width           =   3255
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
            Left            =   4320
            TabIndex        =   7
            ToolTipText     =   "CURRENT PASSWORD IS REQ"
            Top             =   1680
            Width           =   3255
         End
         Begin VB.Label Label8 
            BackColor       =   &H00FF8080&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   8640
            TabIndex        =   14
            Top             =   3480
            Width           =   1095
         End
         Begin VB.Label Label7 
            BackColor       =   &H00FF8080&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   8640
            TabIndex        =   13
            Top             =   2880
            Width           =   1095
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "RE-SET PASSWORD"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3960
            TabIndex        =   10
            Top             =   840
            Width           =   2895
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00FFC0C0&
            BorderWidth     =   8
            Height          =   5535
            Left            =   960
            Shape           =   2  'Oval
            Top             =   480
            Width           =   9135
         End
         Begin VB.Label Label5 
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
            Left            =   2280
            TabIndex        =   6
            Top             =   3480
            Width           =   1815
         End
         Begin VB.Label Label4 
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
            Height          =   375
            Left            =   2280
            TabIndex        =   5
            Top             =   2880
            Width           =   1815
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "CURRENT PASSWORD :"
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
            Left            =   2280
            TabIndex        =   4
            Top             =   2280
            Width           =   1815
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
            Height          =   375
            Left            =   2280
            TabIndex        =   3
            Top             =   1680
            Width           =   1815
         End
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "RESET PASSWORD"
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
      TabIndex        =   0
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "frmRESET"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Sub currpass_lostfocus()
'On Error GoTo label
'Module1.conn
'sql = "select pass from Create_User where userid='" + id.Text + "'"
'Set r = c.Execute(sql)
'Password = r.Fields("pass")
'Exit Sub
'label:
'If (currpass.Text <> Password) Then
'a = MsgBox("CURRENT PASSWORD  IS NOT CORRECT", vbInformation + vbOKOnly, "INFORMATION")
'End If
'If (a = vbOK) Then
'id.SetFocus
'End If
'End Sub
Private Sub ref()
ID.Text = ""
currpass.Text = ""
newpass.Text = ""
cnfrmpass.Text = ""
End Sub

Private Sub Form_Activate()
ID.SetFocus
End Sub

Private Sub save_Click()
On Error GoTo label
If (ID.Text = "" Or currpass.Text = "" Or newpass.Text = "" Or cnfrmpass.Text = "") Then
A = MsgBox("ALL FIELDS ARE MANDATORY", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
ID.SetFocus
End If
Else
Module1.conn
sql = "select pass from Create_User where userid='" + ID.Text + "'"
Set r = c.Execute(sql)
Password = r.Fields("pass")
If (currpass.Text <> Password) Then
A = MsgBox("CURRENT PASSWORD  IS NOT CORRECT", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
ID.SetFocus
End If
Else
sql = "update Create_User set pass='" + newpass.Text + "',cnfrm_pass='" + cnfrmpass.Text + "'"
Set r = c.Execute(sql)
A = MsgBox("PASSWORD IS RESET", vbInformation + vbOKOnly, "information")
If (A = vbOK) Then
Call ref
ID.SetFocus
End If
End If
Exit Sub
label:
A = MsgBox("USERID IS NOT CORRECT", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
ID.SetFocus
End If
End If
End Sub

