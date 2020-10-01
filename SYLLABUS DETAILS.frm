VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmSYLLABUS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SYLLABUS"
   ClientHeight    =   9195
   ClientLeft      =   1035
   ClientTop       =   870
   ClientWidth     =   19605
   LinkTopic       =   "Form14"
   MaxButton       =   0   'False
   Picture         =   "SYLLABUS DETAILS.frx":0000
   ScaleHeight     =   9195
   ScaleWidth      =   19605
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Height          =   8055
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   16365
      Begin VB.Frame Frame2 
         BackColor       =   &H00FF8080&
         Height          =   7695
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   16095
         Begin VB.ComboBox subid 
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
            Left            =   12900
            TabIndex        =   12
            Top             =   270
            Width           =   2205
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
            ItemData        =   "SYLLABUS DETAILS.frx":1E9608
            Left            =   3144
            List            =   "SYLLABUS DETAILS.frx":1E9615
            TabIndex        =   11
            Text            =   "SELECT"
            Top             =   198
            Width           =   2415
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
            Left            =   8112
            TabIndex        =   9
            Text            =   "SELECT"
            Top             =   198
            Width           =   2415
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00FF8080&
            Height          =   6855
            Left            =   120
            TabIndex        =   4
            Top             =   840
            Width           =   15855
            Begin MSComDlg.CommonDialog cd1 
               Left            =   3450
               Top             =   6090
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
            Begin MSComctlLib.ImageList ImageList1 
               Left            =   6030
               Top             =   420
               _ExtentX        =   1005
               _ExtentY        =   1005
               BackColor       =   -2147483643
               ImageWidth      =   32
               ImageHeight     =   32
               MaskColor       =   12632256
               _Version        =   393216
               BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
                  NumListImages   =   5
                  BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "SYLLABUS DETAILS.frx":1E9625
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "SYLLABUS DETAILS.frx":1EBD49
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "SYLLABUS DETAILS.frx":1EE64B
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "SYLLABUS DETAILS.frx":1F0EB1
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "SYLLABUS DETAILS.frx":1F3A39
                     Key             =   ""
                  EndProperty
               EndProperty
            End
            Begin VB.HScrollBar HScroll1 
               Height          =   255
               Left            =   120
               TabIndex        =   7
               Top             =   5640
               Width           =   15375
            End
            Begin VB.VScrollBar VScroll1 
               Height          =   4935
               Left            =   15480
               TabIndex        =   6
               Top             =   960
               Width           =   255
            End
            Begin RichTextLib.RichTextBox rtb1 
               Height          =   4935
               Left            =   90
               TabIndex        =   5
               Top             =   990
               Width           =   15615
               _ExtentX        =   27543
               _ExtentY        =   8705
               _Version        =   393217
               BackColor       =   -2147483626
               HideSelection   =   0   'False
               TextRTF         =   $"SYLLABUS DETAILS.frx":1F3FF2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Cambria"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin MSComctlLib.Toolbar Toolbar1 
               Height          =   870
               Left            =   90
               TabIndex        =   13
               Top             =   180
               Width           =   15645
               _ExtentX        =   27596
               _ExtentY        =   1535
               ButtonWidth     =   1217
               ButtonHeight    =   1429
               Appearance      =   1
               ImageList       =   "ImageList1"
               _Version        =   393216
               BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
                  NumButtons      =   6
                  BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Caption         =   "NEW"
                     Object.ToolTipText     =   "create new file"
                     ImageIndex      =   1
                     BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                        NumButtonMenus  =   8
                        BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                        EndProperty
                        BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                        EndProperty
                        BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                        EndProperty
                        BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                        EndProperty
                        BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                        EndProperty
                        BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                        EndProperty
                        BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                        EndProperty
                        BeginProperty ButtonMenu8 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                        EndProperty
                     EndProperty
                  EndProperty
                  BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Caption         =   "OPEN"
                     Object.ToolTipText     =   "open file"
                     ImageIndex      =   2
                  EndProperty
                  BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Caption         =   "SAVE"
                     Object.ToolTipText     =   "save file"
                     ImageIndex      =   3
                  EndProperty
                  BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Caption         =   "FONT"
                     Object.ToolTipText     =   "change font"
                  EndProperty
                  BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Caption         =   "COLOR"
                     Object.ToolTipText     =   "change color"
                  EndProperty
                  BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Caption         =   "EXIT"
                     Object.ToolTipText     =   "exit"
                     ImageIndex      =   4
                  EndProperty
               EndProperty
            End
         End
         Begin VB.Label Label4 
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
            Height          =   450
            Left            =   1440
            TabIndex        =   10
            Top             =   195
            Width           =   1335
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "SUBJECT ID :"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   11520
            TabIndex        =   3
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "COURSE  :"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   6480
            TabIndex        =   2
            Top             =   195
            Width           =   1335
         End
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SYLLABUS"
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
      Left            =   7200
      TabIndex        =   8
      Top             =   0
      Width           =   4815
   End
End
Attribute VB_Name = "frmSYLLABUS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
Label5.Caption = Time
End Sub
Private Sub course_lostfocus()
If (course.Text = course.Text) Then
Module1.conn
sql = "select sub_id from subjectdetail_master where code='" + course.Text + "'"
Set r = c.Execute(sql)
While (r.EOF = False)
subid.AddItem r.Fields("sub_id")
r.MoveNext
Wend
End If
End Sub

Private Sub Form_Load()
rtb1.Locked = False
Module1.conn
sql = "select code from CourseDetail_Master"
Set r = c.Execute(sql)
While (r.EOF = False)
course.AddItem r.Fields("code")
r.MoveNext
Wend
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.Index = 1 Then
rtb1.Text = clear
Else
If Button.Index = 2 Then
cd1.ShowOpen
rtb1.LoadFile cd1.FileName
Else
If Button.Index = 3 Then
cd1.ShowSave
rtb1.SaveFile (cd1.FileName)
Else
If (Button.Index = 4) Then
cd1.ShowFont
rtb1.SelFontName = cd1.FontName
rtb1.SelBold = cd1.FontBold
rtb1.SelItalic = cd1.FontItalic
rtb1.SelFontSize = cd1.FontSize
Else
If Button.Index = 5 Then
cd1.ShowColor
Else
If (Button.Index = 6) Then
Unload Me
End If
End If
End If
End If
End If
End If
End Sub
