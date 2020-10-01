VERSION 5.00
Begin VB.Form frmEXAMINATIONREPORT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EXAMINATION REPORT"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11370
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   11370
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Height          =   3315
      Left            =   60
      TabIndex        =   0
      Top             =   60
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
            Left            =   2940
            TabIndex        =   2
            Top             =   1830
            Width           =   2295
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00FF8080&
            Caption         =   "COURSE CODE:"
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
Attribute VB_Name = "frmEXAMINATIONREPORT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub show_Click()

End Sub
