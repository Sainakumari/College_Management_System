VERSION 5.00
Begin VB.Form frmDATAEXIST 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INFORMATION"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3750
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   3750
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   2295
      Left            =   0
      Picture         =   "frmDATAEXIST.frx":0000
      ScaleHeight     =   2235
      ScaleWidth      =   3735
      TabIndex        =   0
      Top             =   0
      Width           =   3795
      Begin VB.CommandButton OK 
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   1440
         Picture         =   "frmDATAEXIST.frx":086B
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1590
         Width           =   825
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   180
         Picture         =   "frmDATAEXIST.frx":0B48
         Top             =   570
         Width           =   480
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         Caption         =   "DATA ALREADY EXIST"
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
         Left            =   840
         TabIndex        =   2
         Top             =   720
         Width           =   2445
      End
   End
End
Attribute VB_Name = "frmDATAEXIST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub OK_Click()
str = "OK"
Unload Me
End Sub
