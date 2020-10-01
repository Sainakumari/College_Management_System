VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmSYLLABUS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SYLLABUS"
   ClientHeight    =   9660
   ClientLeft      =   1035
   ClientTop       =   870
   ClientWidth     =   19605
   LinkTopic       =   "Form14"
   Picture         =   "SUBJECT DETAILS.frx":0000
   ScaleHeight     =   9660
   ScaleWidth      =   19605
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command9 
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
      Height          =   1095
      Left            =   17520
      Picture         =   "SUBJECT DETAILS.frx":1E9608
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2160
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
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
      Height          =   1095
      Left            =   17520
      Picture         =   "SUBJECT DETAILS.frx":2318CA
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
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
      Height          =   1095
      Left            =   17520
      Picture         =   "SUBJECT DETAILS.frx":2604CF
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3450
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
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
      Height          =   1095
      Left            =   17520
      Picture         =   "SUBJECT DETAILS.frx":272504
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4710
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
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
      Height          =   1095
      Left            =   17520
      Picture         =   "SUBJECT DETAILS.frx":274EF4
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7200
      Width           =   1095
   End
   Begin VB.CommandButton Command10 
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
      Height          =   1095
      Left            =   17520
      Picture         =   "SUBJECT DETAILS.frx":277AB0
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Height          =   8055
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   18495
      Begin VB.Frame Frame2 
         BackColor       =   &H00FF8080&
         Height          =   7695
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   16215
         Begin VB.ListBox List1 
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
            Height          =   420
            ItemData        =   "SUBJECT DETAILS.frx":2C033A
            Left            =   13080
            List            =   "SUBJECT DETAILS.frx":2C033C
            TabIndex        =   12
            Top             =   198
            Width           =   2415
         End
         Begin VB.ComboBox Combo2 
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
            Left            =   3144
            TabIndex        =   11
            Text            =   "SELECT"
            Top             =   198
            Width           =   2415
         End
         Begin VB.ComboBox Combo1 
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
            Begin MSComctlLib.ImageList ImageList1 
               Left            =   4920
               Top             =   360
               _ExtentX        =   1005
               _ExtentY        =   1005
               BackColor       =   -2147483643
               ImageWidth      =   32
               ImageHeight     =   32
               MaskColor       =   12632256
               _Version        =   393216
               BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
                  NumListImages   =   4
                  BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "SUBJECT DETAILS.frx":2C033E
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "SUBJECT DETAILS.frx":2C2A62
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "SUBJECT DETAILS.frx":2C5364
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "SUBJECT DETAILS.frx":2C7BCA
                     Key             =   ""
                  EndProperty
               EndProperty
            End
            Begin MSComctlLib.Toolbar Toolbar1 
               Height          =   870
               Left            =   120
               TabIndex        =   19
               Top             =   240
               Width           =   15375
               _ExtentX        =   27120
               _ExtentY        =   1535
               ButtonWidth     =   1032
               ButtonHeight    =   1429
               Appearance      =   1
               ImageList       =   "ImageList1"
               _Version        =   393216
               BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
                  NumButtons      =   4
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
                     Caption         =   "EXIT"
                     Object.ToolTipText     =   "exit"
                     ImageIndex      =   4
                  EndProperty
               EndProperty
            End
            Begin VB.HScrollBar HScroll1 
               Height          =   255
               Left            =   120
               TabIndex        =   7
               Top             =   4920
               Width           =   15375
            End
            Begin VB.VScrollBar VScroll1 
               Height          =   4935
               Left            =   15480
               TabIndex        =   6
               Top             =   240
               Width           =   255
            End
            Begin RichTextLib.RichTextBox rtb1 
               Height          =   4935
               Left            =   120
               TabIndex        =   5
               Top             =   240
               Width           =   15615
               _ExtentX        =   27543
               _ExtentY        =   8705
               _Version        =   393217
               BackColor       =   -2147483626
               Enabled         =   -1  'True
               TextRTF         =   $"SUBJECT DETAILS.frx":2CA752
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
      Begin VB.Shape Shape2 
         BorderColor     =   &H8000000B&
         Height          =   7575
         Left            =   16680
         Top             =   240
         Width           =   1575
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
      Left            =   5520
      TabIndex        =   8
      Top             =   0
      Width           =   7095
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

