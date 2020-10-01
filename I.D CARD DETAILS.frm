VERSION 5.00
Begin VB.Form frmSTUDENTIDCARD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "STUDENT IDENTITY CARD"
   ClientHeight    =   9630
   ClientLeft      =   2190
   ClientTop       =   870
   ClientWidth     =   16080
   BeginProperty Font 
      Name            =   "Cambria"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form17"
   MaxButton       =   0   'False
   Picture         =   "I.D CARD DETAILS.frx":0000
   ScaleHeight     =   9630
   ScaleWidth      =   16080
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame vv 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8535
      Left            =   150
      TabIndex        =   1
      Top             =   720
      Width           =   15855
      Begin VB.CommandButton Command1 
         Caption         =   "SHOW DATA"
         Height          =   765
         Left            =   12510
         TabIndex        =   30
         Top             =   7290
         Width           =   1965
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   150
         TabIndex        =   17
         Top             =   240
         Width           =   15615
         Begin VB.TextBox Text11 
            Alignment       =   2  'Center
            BackColor       =   &H00FF8080&
            Height          =   285
            Left            =   5670
            Locked          =   -1  'True
            TabIndex        =   21
            Text            =   "[College code : 204]"
            Top             =   1560
            Width           =   4335
         End
         Begin VB.TextBox Text10 
            Alignment       =   2  'Center
            BackColor       =   &H00FF8080&
            Height          =   375
            Left            =   3750
            Locked          =   -1  'True
            TabIndex        =   20
            Text            =   "rajendra nagar , patna-800016(BIHAR)"
            Top             =   1110
            Width           =   9735
         End
         Begin VB.TextBox Text9 
            Alignment       =   2  'Center
            BackColor       =   &H00FF8080&
            Height          =   285
            Left            =   4350
            Locked          =   -1  'True
            TabIndex        =   19
            Text            =   "www.abcollege.org"
            Top             =   780
            Width           =   7455
         End
         Begin VB.TextBox Text8 
            Alignment       =   2  'Center
            BackColor       =   &H00FF8080&
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
            Left            =   3750
            Locked          =   -1  'True
            TabIndex        =   18
            Text            =   "ARCADE BUSINESS COLLEGE"
            Top             =   240
            Width           =   10485
         End
         Begin VB.Image Image2 
            Height          =   2010
            Left            =   1200
            Picture         =   "I.D CARD DETAILS.frx":1E9608
            Stretch         =   -1  'True
            Top             =   210
            Width           =   1800
         End
      End
      Begin VB.CommandButton save 
         Caption         =   "SAVE"
         Height          =   975
         Left            =   2520
         Picture         =   "I.D CARD DETAILS.frx":1EAA89
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   7200
         Width           =   1095
      End
      Begin VB.CommandButton print 
         Caption         =   "PRINT"
         Height          =   975
         Left            =   10080
         Picture         =   "I.D CARD DETAILS.frx":21968E
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   7260
         Width           =   1095
      End
      Begin VB.CommandButton refresh 
         Caption         =   "REFRESH"
         Height          =   975
         Left            =   8160
         Picture         =   "I.D CARD DETAILS.frx":21C24A
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   7230
         Width           =   1095
      End
      Begin VB.CommandButton delete 
         Caption         =   "DELETE"
         Height          =   975
         Left            =   6270
         Picture         =   "I.D CARD DETAILS.frx":264AD4
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   7200
         Width           =   1095
      End
      Begin VB.CommandButton view 
         Caption         =   "VIEW"
         Height          =   975
         Left            =   4395
         Picture         =   "I.D CARD DETAILS.frx":276B09
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   7200
         Width           =   1095
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FF8080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4215
         Left            =   240
         TabIndex        =   2
         Top             =   2670
         Width           =   15615
         Begin VB.ComboBox course 
            BackColor       =   &H80000016&
            Height          =   345
            Left            =   5490
            TabIndex        =   32
            Top             =   450
            Width           =   3135
         End
         Begin VB.ComboBox adm 
            BackColor       =   &H80000016&
            Height          =   345
            Left            =   11760
            TabIndex        =   31
            Top             =   480
            Width           =   3105
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00FF8080&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2175
            Left            =   240
            TabIndex        =   3
            Top             =   360
            Width           =   2415
            Begin VB.Image Image1 
               BorderStyle     =   1  'Fixed Single
               Height          =   1815
               Left            =   120
               Top             =   240
               Width           =   2175
            End
         End
         Begin VB.Label addr 
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
            Height          =   495
            Index           =   6
            Left            =   5490
            TabIndex        =   29
            Top             =   2700
            Width           =   3135
         End
         Begin VB.Label session 
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
            Height          =   495
            Index           =   5
            Left            =   11760
            TabIndex        =   28
            Top             =   1110
            Width           =   3135
         End
         Begin VB.Label bloodgrp 
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
            Height          =   495
            Index           =   4
            Left            =   11760
            TabIndex        =   27
            Top             =   3360
            Width           =   3135
         End
         Begin VB.Label dob 
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
            Height          =   495
            Index           =   3
            Left            =   11760
            TabIndex        =   26
            Top             =   1860
            Width           =   3135
         End
         Begin VB.Label fathernm 
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
            Height          =   495
            Index           =   2
            Left            =   5490
            TabIndex        =   25
            Top             =   1900
            Width           =   3135
         End
         Begin VB.Label contact 
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
            Height          =   495
            Index           =   1
            Left            =   11760
            TabIndex        =   24
            Top             =   2610
            Width           =   3135
         End
         Begin VB.Label nm 
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
            Height          =   495
            Index           =   0
            Left            =   5490
            TabIndex        =   23
            Top             =   1100
            Width           =   3135
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "ADMISSION NO :"
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
            Left            =   9900
            TabIndex        =   22
            Top             =   510
            Width           =   1335
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "CONTACT NO. :"
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
            Left            =   9900
            TabIndex        =   11
            Top             =   2736
            Width           =   1695
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
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
            Height          =   495
            Left            =   3510
            TabIndex        =   10
            Top             =   2670
            Width           =   1335
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "BLOOD GROUP :"
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
            Left            =   9900
            TabIndex        =   9
            Top             =   3480
            Width           =   1695
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "SESSION :"
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
            Left            =   9900
            TabIndex        =   8
            Top             =   1252
            Width           =   1695
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "FATHER'S NAME :"
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
            Left            =   3510
            TabIndex        =   7
            Top             =   1950
            Width           =   1695
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "D.O.B :"
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
            Left            =   9900
            TabIndex        =   6
            Top             =   1994
            Width           =   1695
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "COURSE :"
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
            Left            =   3510
            TabIndex        =   5
            Top             =   510
            Width           =   1695
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
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
            Height          =   495
            Left            =   3510
            TabIndex        =   4
            Top             =   1230
            Width           =   1695
         End
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H8000000B&
         Height          =   1455
         Left            =   2160
         Top             =   6960
         Width           =   10155
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "STUDENT IDENTITY  CARD"
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
      Left            =   4800
      TabIndex        =   0
      Top             =   120
      Width           =   7215
   End
End
Attribute VB_Name = "frmSTUDENTIDCARD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmSTUDENTIDDATAVIEW.show
Unload Me
End Sub

Private Sub course_GotFocus()
adm.clear
End Sub

Private Sub COURSE_LostFocus()
Module1.conn
sql = "select stud_adm from studentdetail_master where course = '" + COURSE.Text + "'"
Set r = c.Execute(sql)
While (r.EOF = False)
adm.AddItem r.Fields("stud_adm")
r.MoveNext
Wend
End Sub

Private Sub delete_Click()
If (nm(0) = "" Or dob(3).Caption = "" Or session(5).Caption = "" Or COURSE.Text = "" Or adm.Text = "" Or contact(1).Caption = "" Or fathernm(2).Caption = "" Or addr(6).Caption = "") Then
A = MsgBox("ALL FIELDS ARE MANDATORY", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOKOnly) Then
adm.SetFocus
End If
Else
Module1.conn
A = MsgBox(UCase("are you sure to delete record"), vbQuestion + vbYesNo, UCase(delete))
If (A = vbYes) Then
sql = "delete from IdDetail where admno='" + adm.Text + "'"
Set r = c.Execute(sql)
A = MsgBox(UCase("record deleted"), vbInformation + vbOKOnly, UCase(delete))
If (A = vbOK) Then
Call allclear
adm.SetFocus
End If
Else
If (A = vbNo) Then
adm.SetFocus
End If
End If
End If
End Sub
Private Sub Form_ACTIVATE()
adm.SetFocus
'adm.MaxLength = 7
Module1.conn
sql = "select code from CourseDetail_Master"
Set r = c.Execute(sql)
While (r.EOF = False)
COURSE.AddItem r.Fields("code")
r.MoveNext
Wend
End Sub

Private Sub refresh_Click()
Call allclear
save.Enabled = True

End Sub
Private Sub allclear()
adm.Text = ""
nm(0).Caption = ""
contact(1).Caption = ""
fathernm(2).Caption = ""
dob(3).Caption = ""
session(5).Caption = ""
COURSE.Text = ""
bloodgrp(4).Caption = ""
addr(6).Caption = ""
End Sub
Private Sub mandtory()
If (nm(0) = "" Or dob(3).Caption = "" Or session(5).Caption = "" Or COURSE(8).Caption = "" Or rollnum(7).Caption = "" Or contact(1).Caption = "" Or fathernm(2).Caption = "" Or addr(6).Caption = "" Or bloodgrp(4).Caption = "") Then
A = MsgBox("ALL FIELDS ARE MANDATORY", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOKOnly) Then
adm.SetFocus
End If
End If
End Sub
Private Sub save_Click()
'On Error GoTo label
If (nm(0) = "" Or dob(3).Caption = "" Or session(5).Caption = "" Or COURSE.Text = "" Or contact(1).Caption = "" Or fathernm(2).Caption = "" Or addr(6).Caption = "") Then
A = MsgBox("ALL FIELDS ARE MANDATORY", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOKOnly) Then
adm.SetFocus
End If
'End If
Else
Module1.conn
sql = "insert into IdDetail values('" + adm.Text + "')"
MsgBox sql
Set r = c.Execute(sql)
A = MsgBox("WANT SAVE RECORD", vbQuestion + vbYesNo, "INFORMATON")
If (A = vbYes) Then
A = MsgBox("RECORD SAVE", vbInformation + vbOKOnly, "INFORMATION")
Call allclear
If (A = vbOKOnly) Then
Call allclear
End If
'Exit Sub
'label:
'a = MsgBox("RECORD ALREADY EXIST", vbInformation + vbOKOnly, "INFORMATION")
'If (a = vbOKOnly) Then
'adm.SetFocus
'End If
Else
If (A = vbNo) Then
A = MsgBox("RECORD NOT SAVE", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOKOnly) Then
adm.SetFocus
End If
End If
End If
End If
End Sub
Private Sub view_Click()
save.Enabled = True

'On Error GoTo label
If (adm.Text = "") Then
A = MsgBox("PLEASE ENTER ADMISSION NUMBER", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
adm.SetFocus
End If
Else
Module1.conn
sql = "select *from StudentDetail_Master where stud_adm='" + UCase(adm.Text) + "'"
Set r = c.Execute(sql)
If (IsNull(r.Fields(7)) Or IsNull(r.Fields(6))) Then
nm(0).Caption = r.Fields(6)
Else
nm(0).Caption = r.Fields(6) + " " + " " + r.Fields(7) + " " + r.Fields(8)
End If
session(5).Caption = r.Fields(2)
If (IsNull(r.Fields(9))) Then
bloodgrp(4).Caption = ""
Else
bloodgrp(4).Caption = r.Fields(9)
End If
dob(3).Caption = r.Fields(10)
contact(1).Caption = r.Fields(14)
fathernm(2).Caption = r.Fields(38)
'course. = r.Fields(43)
addr(6).Caption = r.Fields(29) & " " & r.Fields(30) & " " & r.Fields(31) & " " & r.Fields(34) & " " & r.Fields(35) & " " & r.Fields(36)
End If
'Exit Sub
'label:
'a = MsgBox("RECORD NOT FOUND", vbInformation + vbOKOnly, "INFORMATION")
'If (a = vbOK) Then
'adm.Text = ""
'adm.SetFocus
'End If
End Sub
