VERSION 5.00
Begin VB.Form frmFACULTYCLASSES 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FACULTY CLASSES"
   ClientHeight    =   8160
   ClientLeft      =   2535
   ClientTop       =   1515
   ClientWidth     =   17280
   LinkTopic       =   "Form23"
   MaxButton       =   0   'False
   Picture         =   "FACULTY CLASSES.frx":0000
   ScaleHeight     =   8160
   ScaleWidth      =   17280
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Height          =   7005
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   16815
      Begin VB.Frame Frame2 
         BackColor       =   &H00FF8080&
         Height          =   6615
         Left            =   210
         TabIndex        =   1
         Top             =   210
         Width           =   16455
         Begin VB.TextBox Text1 
            Height          =   225
            Left            =   60
            TabIndex        =   29
            Top             =   5910
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.TextBox department 
            Height          =   465
            Left            =   12960
            TabIndex        =   27
            Top             =   300
            Width           =   2235
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
            ItemData        =   "FACULTY CLASSES.frx":1E9608
            Left            =   1320
            List            =   "FACULTY CLASSES.frx":1E960A
            TabIndex        =   26
            Text            =   "SELECT"
            Top             =   330
            Width           =   2415
         End
         Begin VB.CommandButton showdata 
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
            Left            =   15030
            Picture         =   "FACULTY CLASSES.frx":1E960C
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   5370
            Width           =   1215
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
            Left            =   11385
            Picture         =   "FACULTY CLASSES.frx":1E9E5B
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   5310
            Width           =   1215
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
            Left            =   8940
            Picture         =   "FACULTY CLASSES.frx":1EA4FF
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   5310
            Width           =   1215
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
            Left            =   6480
            Picture         =   "FACULTY CLASSES.frx":1EAC8F
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   5310
            Width           =   1215
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
            Left            =   4035
            Picture         =   "FACULTY CLASSES.frx":1EB1FB
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   5310
            Width           =   1215
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
            Left            =   1620
            Picture         =   "FACULTY CLASSES.frx":1EB5B2
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   5355
            Width           =   1215
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00FF8080&
            Height          =   3615
            Left            =   120
            TabIndex        =   6
            Top             =   1200
            Width           =   16215
            Begin VB.ListBox duration 
               BackColor       =   &H00FF8080&
               Height          =   2010
               ItemData        =   "FACULTY CLASSES.frx":1EBD8C
               Left            =   11400
               List            =   "FACULTY CLASSES.frx":1EBD8E
               TabIndex        =   30
               Top             =   1290
               Width           =   3285
            End
            Begin VB.ListBox List1 
               BackColor       =   &H00FF8080&
               Height          =   2400
               ItemData        =   "FACULTY CLASSES.frx":1EBD90
               Left            =   3210
               List            =   "FACULTY CLASSES.frx":1EBD92
               TabIndex        =   28
               Top             =   1020
               Width           =   3105
            End
            Begin VB.ListBox selsub 
               BackColor       =   &H00FF8080&
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2400
               ItemData        =   "FACULTY CLASSES.frx":1EBD94
               Left            =   7800
               List            =   "FACULTY CLASSES.frx":1EBD96
               TabIndex        =   24
               Top             =   960
               Width           =   3135
            End
            Begin VB.TextBox dur 
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
               Left            =   11400
               TabIndex        =   22
               Top             =   840
               Width           =   3255
            End
            Begin VB.CommandButton move 
               Height          =   615
               Left            =   6480
               Picture         =   "FACULTY CLASSES.frx":1EBD98
               Style           =   1  'Graphical
               TabIndex        =   21
               Top             =   1680
               Width           =   975
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
               ForeColor       =   &H80000012&
               Height          =   300
               Left            =   420
               Style           =   2  'Dropdown List
               TabIndex        =   20
               Top             =   2520
               Width           =   2175
            End
            Begin VB.CommandButton remove 
               Caption         =   "REMOVE"
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
               Left            =   15120
               Picture         =   "FACULTY CLASSES.frx":1EE888
               Style           =   1  'Graphical
               TabIndex        =   12
               Top             =   2280
               Width           =   975
            End
            Begin VB.CommandButton add 
               Caption         =   "ADD"
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
               Left            =   15120
               Picture         =   "FACULTY CLASSES.frx":237112
               Style           =   1  'Graphical
               TabIndex        =   11
               Top             =   840
               Width           =   975
            End
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
               ItemData        =   "FACULTY CLASSES.frx":2379CC
               Left            =   360
               List            =   "FACULTY CLASSES.frx":2379CE
               Style           =   2  'Dropdown List
               TabIndex        =   10
               Top             =   840
               Width           =   2175
            End
            Begin VB.Line Line9 
               BorderColor     =   &H8000000B&
               X1              =   6960
               X2              =   6960
               Y1              =   120
               Y2              =   1680
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   "SELECTED SUBJECT"
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
               Left            =   8160
               TabIndex        =   23
               Top             =   240
               Width           =   1815
            End
            Begin VB.Line Line8 
               BorderColor     =   &H8000000B&
               X1              =   6960
               X2              =   6960
               Y1              =   2280
               Y2              =   3600
            End
            Begin VB.Line Line7 
               BorderColor     =   &H8000000B&
               X1              =   0
               X2              =   2880
               Y1              =   2400
               Y2              =   2400
            End
            Begin VB.Label Label8 
               BackStyle       =   0  'Transparent
               Caption         =   "YEAR"
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
               Left            =   1200
               TabIndex        =   19
               Top             =   2160
               Width           =   735
            End
            Begin VB.Line Line6 
               BorderColor     =   &H8000000B&
               X1              =   0
               X2              =   2880
               Y1              =   2040
               Y2              =   2040
            End
            Begin VB.Line Line5 
               BorderColor     =   &H8000000B&
               X1              =   0
               X2              =   2880
               Y1              =   1920
               Y2              =   1920
            End
            Begin VB.Line Line4 
               BorderColor     =   &H8000000B&
               X1              =   14880
               X2              =   14880
               Y1              =   120
               Y2              =   3840
            End
            Begin VB.Line Line3 
               BorderColor     =   &H8000000B&
               X1              =   11160
               X2              =   11160
               Y1              =   120
               Y2              =   3960
            End
            Begin VB.Line Line2 
               BorderColor     =   &H8000000B&
               X1              =   2880
               X2              =   2880
               Y1              =   120
               Y2              =   3840
            End
            Begin VB.Line Line1 
               BorderColor     =   &H8000000B&
               X1              =   0
               X2              =   14880
               Y1              =   720
               Y2              =   720
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "DURATION (TOTAL CLASS)"
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
               Left            =   12000
               TabIndex        =   9
               Top             =   240
               Width           =   2415
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "TOTAL SUBJECT"
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
               Left            =   4200
               TabIndex        =   8
               Top             =   240
               Width           =   1815
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
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
               Height          =   255
               Left            =   1080
               TabIndex        =   7
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.TextBox nm 
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
            Left            =   6990
            TabIndex        =   5
            Top             =   270
            Width           =   2175
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H8000000B&
            Height          =   1455
            Left            =   240
            Top             =   5040
            Width           =   14175
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
            Height          =   495
            Left            =   11532
            TabIndex        =   4
            Top             =   360
            Width           =   1290
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
            Height          =   495
            Left            =   6066
            TabIndex        =   3
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label1 
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
            Left            =   240
            TabIndex        =   2
            Top             =   360
            Width           =   1455
         End
      End
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "FACULTY CLASSES"
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
      Left            =   5880
      TabIndex        =   18
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "frmFACULTYCLASSES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub add_Click()
duration.AddItem dur.Text
dur.Text = ""
End Sub

Private Sub course_Click()
List1.clear
End Sub

Private Sub course_lostfocus()
sql = "select sub_id from subjectDetail_Master where code='" + course.Text + "'"
Set r = c.Execute(sql)
While (r.EOF = False)
List1.AddItem r.Fields("sub_id")
r.MoveNext
Wend
End Sub

Private Sub delete_Click()
If (facid.Text = "" Or course.Text = "" Or List1.ListCount = 0 Or selsub.ListCount = 0 Or duration.ListCount = 0 Or nm.Text = "" Or department.Text = "" Or Year.Text = "") Then
A = MsgBox("ALL FIELDS ARE MANDATORY", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
facid.SetFocus
End If
Else
Module1.conn
A = MsgBox("WANT TO DELETE RECORD", vbQuestion + vbYesNo, "DELETE")
If (A = vbYes) Then
sql = "delete from facultyclassesdetail_Master where facid='" + facid.Text + "'"
Set r = c.Execute(sql)
selsub.clear
A = MsgBox("RECORD DELETED", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
facid.SetFocus
End If
Else
If (A = vbNo) Then
facid.SetFocus
End If
End If
End If
End Sub

Private Sub facid_LostFocus()
Module1.conn
sql = "select fac_fnm ,dcode from facultydetail_master where fac_id='" + facid.Text + "'"
Set r = c.Execute(sql)
nm.Text = r.Fields("fac_fnm")
department.Text = r.Fields("dcode")
End Sub
Private Sub Form_Load()
'dur.Locked = True
Module1.conn
sql = "select fac_id from facultyDetail_Master"
Set r = c.Execute(sql)
While (r.EOF = False)
facid.AddItem r.Fields("fac_id")
r.MoveNext
Wend
sql = "select code from courseDetail_Master"
Set r = c.Execute(sql)
While (r.EOF = False)
course.AddItem r.Fields("code")
r.MoveNext
Wend
Year.AddItem "I"
Year.AddItem "II"
Year.AddItem "III"
End Sub

Private Sub move_Click()
If (List1.ListCount = 0) Then
course.SetFocus
Else
If (selsub.ListCount = 0) Then
Text1.Text = List1.List(List1.ListIndex)
selsub.AddItem Text1.Text
Else
i = 0
While (i <= List1.ListCount - 1)
Text1.Text = List1.List(List1.ListIndex)
If (List1.List(i) = Text1.Text) Then
A = MsgBox("SUBJECT ALREADY EXIST", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
List1.SetFocus
i = i + 1
End If
Else
selsub.AddItem Text1.Text
End If
Wend
End If
End If
End Sub
Private Sub ref()
List1.clear
selsub.clear
dur.Text = ""
duration.clear
nm.Text = ""
department.Text
End Sub
Private Sub refresh_Click()
Call ref
End Sub

Private Sub remove_Click()
Item = selsub.ListIndex
If (Item >= 0) Then
selsub.RemoveItem Item
End If
End Sub


'Private Sub selsub_LostFocus()
'If (selsub.ListCount = 0) Then
'course.SetFocus
'Else
'A = selsub.ListCount
'dur.Text = (A) & "hour"
'End If
'End Sub
Private Sub save_Click()
If (facid.Text = "" Or course.Text = "" Or List1.ListCount = 0 Or selsub.ListCount = 0 Or duration.ListCount = 0 Or nm.Text = "" Or department.Text = "" Or Year.Text = "") Then
A = MsgBox("ALL FIELDS ARE MANDATORY", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
facid.SetFocus
End If
Else
Module1.conn
A = MsgBox("CLICK OK TO SAVE RECORD", vbQuestion + vbOKCancel, "SAVE")
If (A = vbOK) Then
i = 0
Counter = 0
While (i <= List1.ListCount - 1 And Counter <= Counter)
sql = "insert into facultyclassesdetail_master values('" + facid.Text + "','" + nm.Text + "','" + department.Text + "','" + course.Text + "','" + Year.Text + "','" + selsub.List(i) + "','" + duration.List(i) + "'," & Counter & ")"
Set r = c.Execute(sql)
i = i + 1
Counter = Counter + 1
Wend
A = MsgBox("RECORD SAVE", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
List1.clear
selsub.clear
nm.Text = ""
department.Text = ""
facid.SetFocus
End If
Else
If (A = vbCancel) Then
facid.SetFocus
End If
End If
End If
End Sub
Private Sub duration_Click()
dur.Text = duration.List(duration.ListIndex)
End Sub
Private Sub update_Click()
If (facid.Text = "" Or course.Text = "" Or duration.ListCount = 0 Or nm.Text = "" Or department.Text = "" Or Year.Text = "") Then
A = MsgBox("ALL FIELDS ARE MANDATORY", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
facid.SetFocus
End If
Else
Module1.conn
A = MsgBox("CLICK OK TO UPDATE RECORD", vbQuestion + vbOKCancel, "UPDATE")
If (A = vbOK) Then
i = duration.ListIndex
If (duration.ListIndex = -1) Then
A = MsgBox("NO ITEM UPDATED", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
facid.SetFocus
End If
Else
duration.RemoveItem duration.ListIndex
duration.AddItem dur.Text
j = i
While (j <= duration.ListCount - 1)
sql = "update facultyclassesdetail_Master set dur='" & duration.List(j) & "' where sno=" & i & ""
Set r = c.Execute(sql)
j = j + 1
Wend
MsgBox "record update"
selsub.clear
dur.Text = ""
duration.clear
End If
End If
End If
End Sub

Private Sub view_Click()
If (facid.Text = "") Then
A = MsgBox("PLEASE ENTER FACULTY ID", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
facid.SetFocus
End If
Else
Module1.conn
sql = "select *from facultyclassesdetail_master where facid='" + facid.Text + "'"
Set r = c.Execute(sql)
nm.Text = r.Fields(1)
department.Text = r.Fields(2)
course.Text = r.Fields(3)
Year.Text = r.Fields(4)
While (r.EOF = False)
selsub.AddItem r.Fields(5)
duration.AddItem r.Fields(6)
r.MoveNext
Wend
End If
End Sub
