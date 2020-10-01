VERSION 5.00
Begin VB.Form frmTIMESLOT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TIME TABLE"
   ClientHeight    =   8790
   ClientLeft      =   870
   ClientTop       =   1365
   ClientWidth     =   18015
   LinkTopic       =   "Form13"
   MaxButton       =   0   'False
   Picture         =   "TIME SLOT DETAIL.frx":0000
   ScaleHeight     =   8790
   ScaleWidth      =   18015
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Height          =   7935
      Left            =   2400
      TabIndex        =   1
      Top             =   720
      Width           =   12975
      Begin VB.Frame Frame2 
         BackColor       =   &H00FF8080&
         Height          =   7695
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   12735
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
            Left            =   10680
            Picture         =   "TIME SLOT DETAIL.frx":1E9608
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   6240
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
            Left            =   7065
            Picture         =   "TIME SLOT DETAIL.frx":1E9E57
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   6180
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
            Left            =   2595
            Picture         =   "TIME SLOT DETAIL.frx":2326E1
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   6180
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
            Left            =   1110
            Picture         =   "TIME SLOT DETAIL.frx":27A9A3
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   6180
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
            Left            =   4080
            Picture         =   "TIME SLOT DETAIL.frx":2A95A8
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   6180
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
            Left            =   5580
            Picture         =   "TIME SLOT DETAIL.frx":2BB5DD
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   6180
            Width           =   1095
         End
         Begin VB.CommandButton print 
            BackColor       =   &H80000016&
            Caption         =   "PRINT"
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
            Left            =   8550
            Picture         =   "TIME SLOT DETAIL.frx":2BDFCD
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   6180
            Width           =   1095
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
            ItemData        =   "TIME SLOT DETAIL.frx":2C0B89
            Left            =   2160
            List            =   "TIME SLOT DETAIL.frx":2C0B8B
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   240
            Width           =   1815
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00FF8080&
            Caption         =   "TIME SLOT"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4695
            Left            =   120
            TabIndex        =   7
            Top             =   1080
            Width           =   12495
            Begin VB.ComboBox clsrm 
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
               Left            =   7590
               Style           =   2  'Dropdown List
               TabIndex        =   32
               Top             =   810
               Width           =   2265
            End
            Begin VB.ComboBox subject 
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
               Left            =   5070
               Style           =   2  'Dropdown List
               TabIndex        =   31
               Top             =   810
               Width           =   2265
            End
            Begin VB.ComboBox facid 
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
               Left            =   10110
               Style           =   2  'Dropdown List
               TabIndex        =   30
               Top             =   780
               Width           =   2265
            End
            Begin VB.CommandButton remove 
               Caption         =   "REMOVE"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1035
               Left            =   1170
               Picture         =   "TIME SLOT DETAIL.frx":2C0B8D
               Style           =   1  'Graphical
               TabIndex        =   29
               Top             =   2580
               Width           =   975
            End
            Begin VB.CommandButton add 
               Caption         =   "ADD"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1035
               Left            =   1140
               Picture         =   "TIME SLOT DETAIL.frx":309417
               Style           =   1  'Graphical
               TabIndex        =   28
               Top             =   1440
               Width           =   975
            End
            Begin VB.TextBox time 
               BackColor       =   &H80000016&
               Height          =   315
               Left            =   2520
               TabIndex        =   27
               Top             =   840
               Width           =   2265
            End
            Begin VB.ListBox List4 
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
               Height          =   3300
               ItemData        =   "TIME SLOT DETAIL.frx":34144F
               Left            =   10020
               List            =   "TIME SLOT DETAIL.frx":341451
               TabIndex        =   25
               Top             =   1080
               Width           =   2415
            End
            Begin VB.ListBox List3 
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
               Height          =   3300
               ItemData        =   "TIME SLOT DETAIL.frx":341453
               Left            =   7500
               List            =   "TIME SLOT DETAIL.frx":341455
               TabIndex        =   24
               Top             =   1140
               Width           =   2415
            End
            Begin VB.ListBox List2 
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
               Height          =   3300
               ItemData        =   "TIME SLOT DETAIL.frx":341457
               Left            =   4980
               List            =   "TIME SLOT DETAIL.frx":341459
               TabIndex        =   23
               Top             =   1140
               Width           =   2415
            End
            Begin VB.ListBox List1 
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
               Height          =   3300
               ItemData        =   "TIME SLOT DETAIL.frx":34145B
               Left            =   2460
               List            =   "TIME SLOT DETAIL.frx":34145D
               TabIndex        =   22
               Top             =   1140
               Width           =   2415
            End
            Begin VB.ComboBox day 
               BackColor       =   &H00FFC0C0&
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
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   13
               Top             =   720
               Width           =   2055
            End
            Begin VB.Line Line5 
               BorderColor     =   &H8000000B&
               X1              =   0
               X2              =   12480
               Y1              =   600
               Y2              =   600
            End
            Begin VB.Line Line4 
               BorderColor     =   &H8000000B&
               X1              =   9960
               X2              =   9960
               Y1              =   120
               Y2              =   6480
            End
            Begin VB.Line Line3 
               BorderColor     =   &H8000000B&
               X1              =   7440
               X2              =   7440
               Y1              =   120
               Y2              =   6480
            End
            Begin VB.Line Line2 
               BorderColor     =   &H8000000B&
               X1              =   4920
               X2              =   4920
               Y1              =   120
               Y2              =   6480
            End
            Begin VB.Line Line1 
               BorderColor     =   &H8000000B&
               X1              =   2400
               X2              =   2400
               Y1              =   120
               Y2              =   6480
            End
            Begin VB.Label Label8 
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
               Left            =   10440
               TabIndex        =   12
               Top             =   240
               Width           =   1335
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "CLASS ROOM NO :"
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
               Left            =   8040
               TabIndex        =   11
               Top             =   240
               Width           =   1335
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "SUBJECT :"
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
               Left            =   5640
               TabIndex        =   10
               Top             =   240
               Width           =   1335
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "TIME :"
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
               Left            =   3120
               TabIndex        =   9
               Top             =   240
               Width           =   1335
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   " DAY :"
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
               Top             =   240
               Width           =   1335
            End
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
            ItemData        =   "TIME SLOT DETAIL.frx":34145F
            Left            =   10560
            List            =   "TIME SLOT DETAIL.frx":341461
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   240
            Width           =   2055
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
            ItemData        =   "TIME SLOT DETAIL.frx":341463
            Left            =   6240
            List            =   "TIME SLOT DETAIL.frx":341465
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   240
            Width           =   2055
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H8000000B&
            Height          =   1575
            Left            =   990
            Top             =   5940
            Width           =   8775
         End
         Begin VB.Label Label9 
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
            Height          =   375
            Left            =   240
            TabIndex        =   14
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label3 
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
            Left            =   8760
            TabIndex        =   4
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label2 
            BackColor       =   &H00E0E0E0&
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
            Left            =   4560
            TabIndex        =   3
            Top             =   240
            Width           =   1575
         End
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TIME SLOT"
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
      Left            =   6210
      TabIndex        =   0
      Top             =   90
      Width           =   5025
   End
End
Attribute VB_Name = "frmTIMESLOT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub delete_Click()
If (List1.ListCount = 0 Or List2.ListCount = 0 Or List3.ListCount = 0 Or List4.ListCount = 0) Then
a = MsgBox("ALL FIELDS ARE MANDATORY", vbInformation + vbOKOnly, "INFORMATION")
If (a = vbOK) Then
course.SetFocus
End If
Else
Module1.conn
a = MsgBox("WANT TO DELETE RECORD", vbQuestion + vbYesNo, "DELETE")
If (a = vbYes) Then
sql = "delete from timeslot_Master where course='" + course.Text + "' and batno=" + batch.Text + " and year='" + year.Text + "' and day='" + day.Text + "'"
Set r = c.Execute(sql)
List1.clear
List2.clear
List3.clear
List4.clear
a = MsgBox("RECORD DELETED", vbInformation + vbOKOnly, "INFORMATION")
If (a = vbOK) Then
course.SetFocus
End If
Else
If (a = vbNo) Then
course.SetFocus
End If
End If
End If
End Sub

Private Sub refresh_Click()
List1.clear
List2.clear
List3.clear
List4.clear
time.Text = ""
subject.Text = ""
facid.Text = ""
clsrm.Text = ""
End Sub

Private Sub save_Click()
If (List1.ListCount = 0 Or List2.ListCount = 0 Or List3.ListCount = 0 Or List4.ListCount = 0) Then
a = MsgBox("ALL FIELDS ARE MANDATORY", vbInformation + vbOKOnly, "INFORMATION")
If (a = vbOK) Then
course.SetFocus
End If
Else
Module1.conn
i = 0
Counter = 0
While (i <= List1.ListCount - 1 And Counter <= Counter)
sql = "insert into timeslot_master values('" + course.Text + "','" + year.Text + "'," + batch.Text + ",'" + day.Text + "','" + List1.List(i) + "','" & List2.List(i) & "','" & List3.List(i) & "','" + List4.List(i) + "'," & Counter & ")"
Set r = c.Execute(sql)
i = i + 1
Counter = Counter + 1
Wend
MsgBox "RECORD SAVE"
End If
End Sub

Private Sub time_change()
add.Enabled = Len(time.Text) > 0
End Sub
Private Sub subject_Change()
add.Enabled = Len(subject.Text) > 0
End Sub
Private Sub clsrm_Change()
add.Enabled = Len(clsrm.Text) > 0
End Sub
Private Sub facid_Change()
add.Enabled = Len(facid.Text) > 0
End Sub
Private Sub add_Click()
If (time.Text = "" Or subject.Text = "" Or clsrm.Text = "" Or facid.Text = "") Then
MsgBox "please enter valid data"
Else
List1.AddItem time.Text
time.Text = ""
List2.AddItem subject.Text
'subject.Text = ""
List3.AddItem clsrm.Text
'clsrm.Text = ""
List4.AddItem facid.Text
'facid.Text = ""
End If
End Sub

Private Sub Form_Load()
year.AddItem "I"
year.AddItem "II"
year.AddItem "III"
Module1.conn
sql = "select code from CourseDetail_Master"
Set r = c.Execute(sql)
While (r.EOF = False)
course.AddItem r.Fields("code")
r.MoveNext
Wend
'sql = "select batno from batchdetail_master where course = '" + course.Text + "' and year = '" + year.Text + "'"
'Set r = c.Execute(sql)
'While (r.EOF = False)
'batch.AddItem r.Fields("batno")
'r.MoveNext
'Wend
sql = "select fac_id from facultyDetail_Master"
Set r = c.Execute(sql)
While (r.EOF = False)
facid.AddItem r.Fields("fac_id")
r.MoveNext
Wend
sql = "select sub_id from subjectDetail_Master"
Set r = c.Execute(sql)
While (r.EOF = False)
subject.AddItem r.Fields("sub_id")
r.MoveNext
Wend
sql = "select cls_rmno from classdetail_master"
Set r = c.Execute(sql)
While (r.EOF = Falsse)
clsrm.AddItem r.Fields("cls_rmno")
r.MoveNext
Wend
'batch.AddItem "1"
'batch.AddItem 2
'batch.AddItem 3
'batch.AddItem 4
'year.AddItem "I"
'year.AddItem "II"
'year.AddItem "III"
day.AddItem "MONDAY"
day.AddItem "TUESDAY"
day.AddItem "WEDNESDAY"
day.AddItem "THURSDAY"
day.AddItem "FRIDAY"
day.AddItem "SATURDAY"
day.AddItem "SUNDAY"
End Sub
Private Sub List1_Click()
time.Text = List1.List(List1.ListIndex)
End Sub
Private Sub List2_Click()
subject.Text = List2.List(List2.ListIndex)
End Sub
Private Sub List3_Click()
clsrm.Text = List3.List(List3.ListIndex)
End Sub
Private Sub List4_Click()
facid.Text = List4.List(List4.ListIndex)
End Sub
Private Sub remove_Click()
Item = List1.ListIndex
If (Item >= 0) Then
List1.RemoveItem Item
End If
Item = List2.ListIndex
If (Item >= 0) Then
List2.RemoveItem Item
End If
Item = List3.ListIndex
If (Item >= 0) Then
List3.RemoveItem Item
End If
Item = List4.ListIndex
If (Item >= 0) Then
List4.RemoveItem Item
End If
End Sub
Private Sub add_gotfocus()
If (time.Text = "") Then
add.Enabled = False
If (subject.Text = "") Then
add.Enabled = False
If (clsrm.Text = "") Then
add.Enabled = False
If (facid.Text = "") Then
add.Enabled = False
End If
End If
End If
End If
End Sub

Private Sub update_Click()
If (List1.ListCount = 0 Or List2.ListCount = 0 Or List3.ListCount = 0 Or List4.ListCount = 0) Then
a = MsgBox("ALL FIELDS ARE MANDATORY", vbInformation + vbOKOnly, "INFORMATION")
If (a = vbOK) Then
course.SetFocus
End If
Else
Module1.conn
a = MsgBox("WANT TO UPDATE RECORD", vbQuestion + vbYesNo, "UPDATE")
If (a = vbYes) Then
i = List1.ListIndex
k = List2.ListIndex
m = List3.ListIndex
n = List4.ListIndex
List1.RemoveItem List1.ListIndex
List2.RemoveItem List2.ListIndex
List3.RemoveItem List3.ListIndex
List4.RemoveItem List4.ListIndex
List1.AddItem time.Text
List2.AddItem subject.Text
List3.AddItem clsrm.Text
List4.AddItem facid.Text
j = i
l = k
P = m
t = n
While (j <= List1.ListCount - 1)
sql = "update timeslot_Master set time='" & List1.List(j) & "' where sno=" & i & ""
Set r = c.Execute(sql)
j = j + 1
Wend
While (l <= List2.ListCount - 1)
sql = "update timeslot_Master set subject='" & List2.List(k) & "' where sno=" & k & ""
Set r = c.Execute(sql)
l = l + 1
Wend
While (P <= List3.ListCount - 1)
sql = "update timeslot_Master set classno=" & List3.List(P) & " where sno=" & m & ""
Set r = c.Execute(sql)
P = P + 1
Wend
While (t <= List4.ListCount - 1)
sql = "update timeslot_Master set facid='" & List4.List(t) & "' where sno=" & n & ""
Set r = c.Execute(sql)
t = t + 1
Wend
MsgBox "RECORD UPDATE"
Else
If (a = vbNo) Then
MsgBox "RECORD NOT UPDATE"
End If
End If
End If
End Sub

Private Sub view_Click()
'On Error GoTo label
If course.Text = "" Or year.Text = "" Or batch.Text = "" Or day.Text = "" Then
a = MsgBox("ALL FIELDS ARE MANDATORY", vbInformation + vbOKOnly, "INFORMATION")
If (a = vbOK) Then
course.SetFocus
End If
Else
Module1.conn
sql = "select *from timeslot_Master where course='" + course.Text + "' and batno=" + batch.Text + " and year='" + year.Text + "' and day='" + day.Text + "'"
Set r = c.Execute(sql)
While (r.EOF = False)
List1.AddItem r.Fields("time")
List2.AddItem r.Fields("subject")
List3.AddItem r.Fields("classno")
List4.AddItem r.Fields("facid")
r.MoveNext
Wend
'Exit Sub
'label:
'a = MsgBox
End If
End Sub

Private Sub year_Click()
batch.clear

End Sub

Private Sub year_lostfocus()
sql = "select batno from batchdetail_master where course = '" + course.Text + "' and year = '" + year.Text + "'"
Set r = c.Execute(sql)
While (r.EOF = False)
batch.AddItem r.Fields("batno")
r.MoveNext
Wend
End Sub
