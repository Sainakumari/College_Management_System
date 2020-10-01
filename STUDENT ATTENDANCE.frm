VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmSTUDENTATTENDENCE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "STUDENT ATTENDANCE"
   ClientHeight    =   9060
   ClientLeft      =   870
   ClientTop       =   705
   ClientWidth     =   18825
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   Picture         =   "STUDENT ATTENDANCE.frx":0000
   ScaleHeight     =   9060
   ScaleWidth      =   18825
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   345
      Left            =   1770
      Top             =   120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   609
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
      RecordSource    =   "select *from studattend_master"
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
      Height          =   7935
      Left            =   1920
      TabIndex        =   0
      Top             =   600
      Width           =   15135
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "STUDENT ATTENDANCE.frx":1E9608
         Height          =   1515
         Left            =   1650
         TabIndex        =   31
         Top             =   4590
         Width           =   11475
         _ExtentX        =   20241
         _ExtentY        =   2672
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   14
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Cambria"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Cambria"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16393
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16393
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
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
         Left            =   9660
         Picture         =   "STUDENT ATTENDANCE.frx":1E961D
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   6480
         Width           =   1095
      End
      Begin VB.CommandButton print 
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
         Left            =   11610
         Picture         =   "STUDENT ATTENDANCE.frx":231EA7
         Style           =   1  'Graphical
         TabIndex        =   13
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
         Left            =   7725
         Picture         =   "STUDENT ATTENDANCE.frx":234A63
         Style           =   1  'Graphical
         TabIndex        =   12
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
         Left            =   5775
         Picture         =   "STUDENT ATTENDANCE.frx":237453
         Style           =   1  'Graphical
         TabIndex        =   11
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
         Left            =   1890
         Picture         =   "STUDENT ATTENDANCE.frx":2396F0
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   6510
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
         Left            =   3840
         Picture         =   "STUDENT ATTENDANCE.frx":2682F5
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   6480
         Width           =   1095
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FF8080&
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4215
         Left            =   90
         TabIndex        =   1
         Top             =   240
         Width           =   14895
         Begin VB.TextBox roll 
            BackColor       =   &H80000016&
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   8550
            TabIndex        =   30
            Top             =   1590
            Width           =   2355
         End
         Begin VB.TextBox batch 
            BackColor       =   &H80000016&
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2880
            TabIndex        =   28
            Top             =   1650
            Width           =   2355
         End
         Begin VB.ComboBox admno 
            BackColor       =   &H80000016&
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   8550
            TabIndex        =   23
            Top             =   300
            Width           =   2355
         End
         Begin VB.ComboBox month 
            BackColor       =   &H80000016&
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "STUDENT ATTENDANCE.frx":2B05B7
            Left            =   2850
            List            =   "STUDENT ATTENDANCE.frx":2B05DF
            TabIndex        =   22
            Top             =   2340
            Width           =   2415
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00FF8080&
            Caption         =   "STUDENT'S PHOTO"
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
            TabIndex        =   15
            Top             =   240
            Width           =   2655
            Begin VB.Image Image1 
               BorderStyle     =   1  'Fixed Single
               Height          =   1575
               Left            =   120
               Top             =   240
               Width           =   2415
            End
         End
         Begin VB.TextBox totalcls 
            BackColor       =   &H80000016&
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   8550
            TabIndex        =   8
            Top             =   2340
            Width           =   2355
         End
         Begin VB.TextBox prsntday 
            BackColor       =   &H80000016&
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2880
            TabIndex        =   7
            Top             =   3000
            Width           =   2355
         End
         Begin VB.TextBox sno 
            BackColor       =   &H80000016&
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2910
            TabIndex        =   6
            Top             =   240
            Width           =   2355
         End
         Begin VB.Label percentage 
            BackColor       =   &H80000016&
            Height          =   345
            Left            =   8520
            TabIndex        =   32
            Top             =   3090
            Width           =   2445
         End
         Begin VB.Label Label13 
            BackColor       =   &H00FF8080&
            Caption         =   "ROLL NO:"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   6840
            TabIndex        =   29
            Top             =   1620
            Width           =   1425
         End
         Begin VB.Label year 
            BackColor       =   &H80000016&
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   8550
            TabIndex        =   27
            Top             =   870
            Width           =   2355
         End
         Begin VB.Label course 
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2880
            TabIndex        =   24
            Top             =   990
            Width           =   2355
         End
         Begin VB.Label Label8 
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
            Height          =   255
            Left            =   900
            TabIndex        =   21
            Top             =   2310
            Width           =   1335
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "ADMISSION-NUMBER :"
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
            Left            =   6570
            TabIndex        =   20
            Top             =   360
            Width           =   1725
         End
         Begin VB.Label Label5 
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
            Left            =   900
            TabIndex        =   19
            Top             =   3180
            Width           =   1215
         End
         Begin VB.Label Label6 
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
            Left            =   6750
            TabIndex        =   18
            Top             =   3210
            Width           =   1215
         End
         Begin VB.Label Label7 
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
            Height          =   255
            Left            =   900
            TabIndex        =   17
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL NO OF CLASSES:"
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
            Left            =   6390
            TabIndex        =   5
            Top             =   2430
            Width           =   1815
         End
         Begin VB.Label Label9 
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
            Left            =   7110
            TabIndex        =   4
            Top             =   990
            Width           =   720
         End
         Begin VB.Label Label4 
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
            Height          =   255
            Left            =   480
            TabIndex        =   3
            Top             =   1620
            Width           =   1335
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
            TabIndex        =   2
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H8000000B&
         Height          =   1455
         Left            =   1800
         Top             =   6270
         Width           =   11295
      End
   End
   Begin VB.Label Label12 
      Caption         =   "Label12"
      Height          =   525
      Left            =   8820
      TabIndex        =   26
      Top             =   4290
      Width           =   1245
   End
   Begin VB.Label Label11 
      Caption         =   "Label11"
      Height          =   525
      Left            =   8820
      TabIndex        =   25
      Top             =   4290
      Width           =   1245
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "STUDENT ATTENDANCE"
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
      Left            =   5760
      TabIndex        =   14
      Top             =   0
      Width           =   7575
   End
End
Attribute VB_Name = "frmSTUDENTATTENDENCE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub admno_LostFocus()
Module1.conn
sql = "select *from studentdetail_master where stud_adm = '" + admno.Text + "'"
Set r = c.Execute(sql)
course.Caption = r.Fields(45)
year.Caption = r.Fields(3)
batch.Text = r.Fields(4)
roll.Text = r.Fields(5)
End Sub

Private Sub Form_Load()
sno.Locked = True
month.AddItem "JANUARY"
month.AddItem "FEBRUARY"
month.AddItem "MARCH"
month.AddItem "APRIL"
month.AddItem "MAY"
month.AddItem "JUNE"
month.AddItem "JULY"
month.AddItem "AUGUST"
month.AddItem "SEPETEMBER"
month.AddItem "OCTOBER"
month.AddItem "NOVEMBER"
month.AddItem "DECEMBER"
Module1.conn
sql = "select stud_adm from studentdetail_master"
Set r = c.Execute(sql)
While (r.EOF = False)
admno.AddItem r.Fields("stud_adm")
r.MoveNext
Wend
sql = "select count(sno )from studattend_master"
Set r = c.Execute(sql)
Dim i As Integer
i = r.Fields(0)
If (i < 1) Then
sno.Text = 1
Else
sql = "select max(sno) from studattend_master"
Set r = c.Execute(sql)
j = r.Fields(0)
sno.Text = j + 1
End If
End Sub

Private Sub percentage_click()
If (totalcls.Text = "" Or prsntday.Text = "") Then
A = MsgBox("PERCENTAGE CAN NOT CALCULATE WITHOUT TOTAL NUMBER OF CLASS OR PRESENT DAY", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
admno.SetFocus
End If
Else
tot = totalcls.Text
persent = prsntday.Text
Percent = persent / tot * 100
percentage.Caption = Percent
End If
End Sub

Private Sub save_Click()
If (sno.Text = "" Or admno.Text = "" Or course.Caption = "" Or year.Caption = "" Or batch.Text = "" Or roll.Text = "" Or month.Text = "" Or totalcls.Text = "" Or prsntday.Text = "" Or percentage.Caption = "") Then
A = MsgBox("ALL FIELDS ARE MANDATORY", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
admno.SetFocus
End If
Else
Module1.conn
A = MsgBox("CLICK OK TO SAVE RECORD", vbQuestion + vbOKCancel, "SAVE")
If (A = vbOK) Then
sql = "insert into studattend_master values(" + sno.Text + ",'" + admno.Text + "','" + course.Caption + "','" + year.Caption + "'," + batch.Text + "," + roll.Text + ",'" + month.Text + "'," + totalcls.Text + "," + prsntday.Text + "," + percentage.Caption + "   )"
Set r = c.Execute(sql)
A = MsgBox("RECORD SAVE", vbInformation + vbOKOnly, "INFORMATION")
If (A = vbOK) Then
sno.Text = sno.Text + 1
admno.Text = ""
course.Caption = ""
year.Caption = ""
batch.Text = ""
roll.Text = ""
month.Text = ""
totalcls.Text = ""
prsntday.Text = ""
percentage.Caption = ""
End If
Else
If (A = vbCancel) Then
admno.SetFocus
End If
End If
End If
End Sub

