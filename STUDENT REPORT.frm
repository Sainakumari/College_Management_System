VERSION 5.00
Begin VB.Form frmSTUDENTREPORT 
   Caption         =   "Form1"
   ClientHeight    =   3285
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11190
   LinkTopic       =   "Form1"
   ScaleHeight     =   3285
   ScaleWidth      =   11190
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H008080FF&
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11175
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
         Left            =   150
         TabIndex        =   1
         Top             =   210
         Width           =   10875
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
            Left            =   7920
            TabIndex        =   5
            Top             =   1740
            Width           =   2295
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
            Left            =   2310
            TabIndex        =   4
            Top             =   360
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
            Left            =   7950
            TabIndex        =   3
            Top             =   240
            Width           =   2685
         End
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
            Left            =   7920
            TabIndex        =   2
            Top             =   960
            Width           =   2685
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00FF8080&
            Caption         =   "ADMISSION NO:"
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
            Left            =   540
            TabIndex        =   6
            Top             =   570
            Width           =   1605
         End
      End
   End
End
Attribute VB_Name = "frmSTUDENTREPORT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub show_Click()
If (selective.Value = True) Then
 If DataEnvironment1.rsstdetail.State = 1 Then
    DataEnvironment1.rsstdetail.Close
    End If
    Module1.conn
    sql = "select stud_adm,stud_fname,stud_lname,stud_dob,stud_gen,code,per_city,per_po,per_ps,per_dist,per_state,stud_contact from StudentDetail_Master where stud_adm='" & UCase(inpt.Text) & "'"
    Set r = c.Execute(sql)
    MsgBox r.Fields(0)
    A = Right(inpt.Text, 3)
    MsgBox r.Fields(6)
    DataReport6.Sections("SECTION2").Controls("Label17").Caption = r.Fields(0)
    DataReport6.Sections("SECTION2").Controls("Label18").Caption = r.Fields(6) & "" & r.Fields(8)
    DataReport6.Sections("SECTION2").Controls("Label19").Caption = r.Fields(10)
    DataReport6.Sections("SECTION2").Controls("Label20").Caption = r.Fields(11)
    DataReport6.Sections("SECTION2").Controls("Label21").Caption = r.Fields(43)
    DataReport6.Sections("SECTION2").Controls("Label22").Caption = r.Fields(22) & "" & r.Fields(23) & "" & r.Fields(24) & "" & r.Fields(25) & "" & r.Fields(26) & "" & r.Fields(27)
    DataReport6.Sections("SECTION2").Controls("Label23").Caption = r.Fields(14)
    DataReport6.Show
    End If
End Sub
