VERSION 5.00
Begin VB.Form frmANOTHERECORD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SAVE"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3780
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   3780
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H8000000D&
      Height          =   2295
      Left            =   0
      Picture         =   "frmANOTHERECORD.frx":0000
      ScaleHeight     =   2235
      ScaleWidth      =   3735
      TabIndex        =   0
      Top             =   0
      Width           =   3795
      Begin VB.CommandButton Command1 
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
         Left            =   1140
         Picture         =   "frmANOTHERECORD.frx":086B
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1650
         Width           =   735
      End
      Begin VB.CommandButton Command2 
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
         Left            =   2280
         Picture         =   "frmANOTHERECORD.frx":0DBE
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1680
         Width           =   735
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   30
         Picture         =   "frmANOTHERECORD.frx":13A5
         Top             =   480
         Width           =   480
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         Caption         =   "WANT TO SAVE another RECORd"
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
         Left            =   600
         TabIndex        =   3
         Top             =   600
         Width           =   3075
      End
   End
End
Attribute VB_Name = "frmANOTHERECORD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
str = "YES"
Unload Me
End Sub
Private Sub Command2_Click()
str = "NO"
Unload Me
End Sub

