VERSION 5.00
Begin VB.Form frmUPDATE 
   Caption         =   "UPDATE"
   ClientHeight    =   2115
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   3735
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   2295
      Left            =   0
      Picture         =   "frmUPDATE.frx":0000
      ScaleHeight     =   2235
      ScaleWidth      =   3735
      TabIndex        =   0
      Top             =   0
      Width           =   3795
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1110
         Picture         =   "frmUPDATE.frx":086B
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1590
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2220
         Picture         =   "frmUPDATE.frx":0DBE
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1620
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         Caption         =   "WANT TO UPDATE RECORd"
         BeginProperty Font 
            Name            =   "Algerian"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   870
         TabIndex        =   3
         Top             =   690
         Width           =   2505
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   300
         Picture         =   "frmUPDATE.frx":13A5
         Top             =   570
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmUPDATE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command3_Click()
str = "NO"
Unload Me
End Sub

Private Sub Command4_Click()
str = "YES"
Unload Me
End Sub

