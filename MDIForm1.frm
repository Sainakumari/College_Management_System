VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H000000FF&
   Caption         =   "THE COLLEGE MANAGEMENT SYSTEM"
   ClientHeight    =   10650
   ClientLeft      =   510
   ClientTop       =   0
   ClientWidth     =   19470
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":0000
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00004080&
      Height          =   1695
      Left            =   0
      ScaleHeight     =   1635
      ScaleWidth      =   19410
      TabIndex        =   0
      Top             =   0
      Width           =   19470
      Begin VB.Timer Timer1 
         Interval        =   255
         Left            =   12840
         Top             =   600
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Castellar"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   14520
         TabIndex        =   5
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Castellar"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   17640
         TabIndex        =   4
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "THE WAY OF SUCCESS"
         BeginProperty Font 
            Name            =   "Algerian"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   8160
         TabIndex        =   3
         Top             =   1080
         Width           =   4305
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000040&
         BackStyle       =   0  'Transparent
         Caption         =   "THE COLLEGE MANAGEMENT SYSTEM"
         BeginProperty Font 
            Name            =   "Algerian"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1215
         Left            =   1920
         TabIndex        =   2
         Top             =   480
         Width           =   10635
      End
      Begin VB.Image Image1 
         Height          =   1425
         Left            =   480
         Picture         =   "MDIForm1.frx":4FBB9
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1545
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000040&
         BackStyle       =   0  'Transparent
         Height          =   1695
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   2055
      End
   End
   Begin VB.Menu COLLEGEINFO 
      Caption         =   "COLLEGE INFO"
      Begin VB.Menu COLLEGE 
         Caption         =   "COLLEGE"
         Shortcut        =   ^J
      End
      Begin VB.Menu BATCH 
         Caption         =   "BATCH"
         Shortcut        =   ^K
      End
      Begin VB.Menu CLASSROOM 
         Caption         =   "CLASS ROOM"
         Shortcut        =   ^L
      End
   End
   Begin VB.Menu COURSE 
      Caption         =   "COURSE"
      Begin VB.Menu COURSEDETAILS 
         Caption         =   "COURSE DETAILS"
         Shortcut        =   ^M
      End
      Begin VB.Menu SYLLABUS 
         Caption         =   "SYLLABUS"
         Shortcut        =   ^N
      End
   End
   Begin VB.Menu DEPARTMENT 
      Caption         =   "DEPARTMENT"
      Begin VB.Menu DEPARTMENTDETAILS 
         Caption         =   "DEPARTMENT DETAILS"
         Shortcut        =   ^I
      End
   End
   Begin VB.Menu STUDENT 
      Caption         =   "STUDENT"
      Begin VB.Menu STUDENTDETAILS 
         Caption         =   "STUDENT DETAILS"
         Shortcut        =   ^A
      End
      Begin VB.Menu IDCARD 
         Caption         =   "ID CARD"
         Shortcut        =   ^B
      End
      Begin VB.Menu STUDENTATTENDANCE 
         Caption         =   "STUDENT ATTENDANCE"
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu FACULTY 
      Caption         =   "FACULTY"
      Begin VB.Menu FACULTYINFO 
         Caption         =   "FACULTY INFO"
         Begin VB.Menu REGULARFACULTY 
            Caption         =   "REGULAR FACULTY "
            Shortcut        =   ^D
         End
         Begin VB.Menu VISITINGFACULTY 
            Caption         =   "VISITING FACULTY"
            Shortcut        =   ^E
         End
         Begin VB.Menu FACULTYCLASSES 
            Caption         =   "FACULTY CLASSES"
            Shortcut        =   ^F
         End
      End
      Begin VB.Menu FACULTYATTENDANCE 
         Caption         =   "FACULTY ATTENDANCE"
         Shortcut        =   ^G
      End
      Begin VB.Menu FACULTYIDCARD 
         Caption         =   "FACULTY ID CARD"
         Shortcut        =   ^H
      End
   End
   Begin VB.Menu EXAMINATION 
      Caption         =   "EXAMINATION"
      Begin VB.Menu EXAMINATIONDETAILS 
         Caption         =   "EXAMINATION DETAILS"
         Shortcut        =   ^O
      End
      Begin VB.Menu RESULT 
         Caption         =   "RESULT"
         Shortcut        =   ^P
      End
   End
   Begin VB.Menu FEE 
      Caption         =   "FEE"
   End
   Begin VB.Menu SETTING 
      Caption         =   "SETTING"
      Begin VB.Menu CREATEUSER 
         Caption         =   "CREATE USER"
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu RESETPASSWORD 
         Caption         =   "RE-SET PASSWORD"
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu LOGOUT 
         Caption         =   "LOGOUT"
      End
   End
   Begin VB.Menu ABOUT 
      Caption         =   "ABOUT"
   End
   Begin VB.Menu HELP 
      Caption         =   "HELP"
      Begin VB.Menu HELPMENU 
         Caption         =   "HELP MENU"
         Shortcut        =   +{INSERT}
      End
   End
   Begin VB.Menu REPORT 
      Caption         =   "REPORT"
      Begin VB.Menu DEPARTMENTREPORT 
         Caption         =   "DEPARTMENT REPORT"
      End
      Begin VB.Menu CLASSROOMREPORT 
         Caption         =   "CLASSROOM REPORT"
      End
      Begin VB.Menu BATCHREPORT 
         Caption         =   "BATCH REPORT"
      End
      Begin VB.Menu COLLEGEREPORT 
         Caption         =   "COLLEGE REPORT"
      End
      Begin VB.Menu COURSEREPORT 
         Caption         =   "COURSE REPORT"
      End
      Begin VB.Menu SUBJECTREPORT 
         Caption         =   "SUBJECT REPORT"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub ABOUTUS_Click()
frmABOUTUS.SHOW 1

End Sub


Private Sub BATCH_Click()
frmBATCH.SHOW
End Sub

Private Sub BATCHREPORT_Click()
frmBATCHREPORT.SHOW 1
End Sub

Private Sub CLASSROOM_Click()
frmCLASSROOM.SHOW 1
End Sub

Private Sub ABOUT_Click()
frmABOUT.SHOW 1
End Sub

Private Sub BATCHDETAIL_Click()
frmBATCH.SHOW 1
End Sub

Private Sub COLLEGEDETAILS_Click()
frmCOLLEGE.SHOW 1
End Sub





Private Sub CLASSROOMREPORT_Click()
frmCLASSROOMREPORT.SHOW 1
End Sub

Private Sub COLLEGE_Click()
frmCOLLEGE.SHOW 1
End Sub

Private Sub COLLEGEREPORT_Click()
frmCOLLEGEREPORT.SHOW 1
End Sub

Private Sub COURSEDETAILS_Click()
frmCOURSE.SHOW 1
End Sub

Private Sub COURSEREPORT_Click()
frmCOURSEREPORT.SHOW
End Sub

Private Sub CREATEUSER_Click()
frmCREATEUSER.SHOW 1

End Sub


Private Sub DEPARTMENTDETAILS_Click()
frmDEPARTMENT.SHOW 1
End Sub

Private Sub DEPARTMENTREPORT_Click()
frmDEPARTMENTREPORT.SHOW
End Sub

Private Sub EXAMINATIONDETAILS_Click()
frmEXAMINATION.SHOW 1
End Sub

Private Sub FACULTYATTENDANCE_Click()
frmFACULTYATTENDENCE.SHOW 1
End Sub

Private Sub FACULTYCLASSES_Click()
frmFACULTYCLASSES.SHOW 1

End Sub

Private Sub FACULTYIDCARD_Click()
frmFACULTYIDCARD.SHOW 1
End Sub

Private Sub FEE_Click()
frmFEE.SHOW 1
End Sub

Private Sub IDCARD_Click()
frmSTUDENTIDCARD.SHOW 1
End Sub

Private Sub REGULARFACULTY_Click()
frmFACULTY.SHOW 1
End Sub

Private Sub RESETPASSWORD_Click()
frmRESET.SHOW 1

End Sub

Private Sub RESULT_Click()
frmRESULT.SHOW 1

End Sub
Private Sub STUDENTATTENDANCE_Click()
frmSTUDENTATTENDENCE.SHOW 1
End Sub

Private Sub STUDENTDETAILS_Click()
frmSTUDENT.SHOW 1
End Sub



Private Sub SUBJECTREPORT_Click()
frmSUBJECTREPORT.SHOW 1
End Sub

Private Sub SYLLABUS_Click()
frmSYLLABUS.SHOW 1
End Sub

Private Sub Timer1_Timer()
Label3.Caption = Format(Date, " DD  mmm   YYYY")
Label5.Caption = time
End Sub

Private Sub VISITINGFACULTY_Click()
frmVISITINGFACULTY.SHOW 1
End Sub
