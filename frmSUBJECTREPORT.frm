VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3360
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11280
   LinkTopic       =   "Form2"
   ScaleHeight     =   3360
   ScaleWidth      =   11280
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H008080FF&
      Height          =   3315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11235
      Begin VB.Frame Frame2 
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
         Height          =   2925
         Left            =   180
         TabIndex        =   1
         Top             =   270
         Width           =   10875
         Begin VB.OptionButton collective 
            Caption         =   "COLLECTIVE"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   7830
            TabIndex        =   5
            Top             =   390
            Width           =   2685
         End
         Begin VB.OptionButton selective 
            Caption         =   "SELECTIVE"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   7800
            TabIndex        =   4
            Top             =   1080
            Width           =   2685
         End
         Begin VB.TextBox inpt 
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
            Height          =   615
            Left            =   2790
            TabIndex        =   3
            Top             =   480
            Width           =   2685
         End
         Begin VB.CommandButton show 
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
            Height          =   945
            Left            =   7740
            TabIndex        =   2
            Top             =   1890
            Width           =   2295
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00FF8080&
            Caption         =   "SUBJECT ID:"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   420
            TabIndex        =   6
            Top             =   630
            Width           =   2115
         End
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub inpt_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Then
inpt.Locked = True
Else
inpt.Locked = False
End If
End Sub

Private Sub show_Click()
If (selective.Value = True) Then
    If DataEnvironment1.rsSUBSLC.state = 1 Then
    DataEnvironment1.rsSUBSLC.Close
    End If
    Module1.conn
    On Error GoTo Lbel
    sql = "select *from subjectdetail_master where sub_id='" & inpt.Text & "'"
    Set r = c.Execute(sql)
    DataReport7.Sections("SECTION2").Controls("LabeL18").Caption = r.Fields(7)
    DataReport7.Sections("SECTION2").Controls("Label19").Caption = r.Fields(0)
    DataReport7.Sections("SECTION2").Controls("Label20").Caption = r.Fields(1)
    DataReport7.Sections("SECTION2").Controls("Label21").Caption = r.Fields(2)
    DataReport7.Sections("SECTION2").Controls("Label22").Caption = r.Fields(3)
    DataReport7.Sections("SECTION2").Controls("Label23").Caption = r.Fields(4)
    DataReport7.Sections("SECTION2").Controls("Label24").Caption = r.Fields(5)
    DataReport7.Sections("SECTION2").Controls("Label25").Caption = r.Fields(6)
    DataReport7.show
    Exit Sub
Lbel:
    A = MsgBox("RECORD NOT FOUND", vbInformation + vbOKOnly, "INFROMATION")
    If (A = vbOK) Then
    inpt.Text = ""
    inpt.SetFocus
   Else
       If (collective.Value = True) Then
           DataReport8.show
           End If
       End If
End If
End Sub

