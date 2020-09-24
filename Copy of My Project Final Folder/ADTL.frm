VERSION 5.00
Begin VB.Form Adt 
   BackColor       =   &H80000006&
   Caption         =   "Advanced Develpment Technologies"
   ClientHeight    =   5685
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   ScaleHeight     =   5685
   ScaleWidth      =   7410
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Advance 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   5295
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6975
      Begin VB.CommandButton cmdabout 
         Caption         =   "&About"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   2760
         Picture         =   "ADTL.frx":0000
         TabIndex        =   5
         Top             =   3480
         Width           =   1455
      End
      Begin VB.CommandButton cmdadmin 
         Caption         =   "&Admin Login"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   2760
         Picture         =   "ADTL.frx":62D0
         TabIndex        =   4
         Top             =   3000
         Width           =   1455
      End
      Begin VB.CommandButton cmdexit 
         Caption         =   "&Exit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   2760
         Picture         =   "ADTL.frx":C5A0
         TabIndex        =   3
         Top             =   3960
         Width           =   1455
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   1575
         Left            =   0
         Picture         =   "ADTL.frx":12870
         ScaleHeight     =   1575
         ScaleWidth      =   3855
         TabIndex        =   2
         Top             =   0
         Width           =   3855
      End
      Begin VB.CommandButton cmdmenu 
         BackColor       =   &H80000013&
         Caption         =   "&Menu"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   2760
         Picture         =   "ADTL.frx":18B40
         TabIndex        =   1
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "   WELCOME TO ADTL LTD."
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000013&
         Height          =   615
         Left            =   1080
         TabIndex        =   6
         Top             =   1680
         Width           =   5175
      End
   End
   Begin VB.Menu mfile 
      Caption         =   "File"
      Begin VB.Menu mchange 
         Caption         =   "Change User"
      End
      Begin VB.Menu mexit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mmenu 
      Caption         =   "Menu"
      Begin VB.Menu mdepartment 
         Caption         =   "Department"
      End
      Begin VB.Menu mprojects 
         Caption         =   "Projects"
      End
      Begin VB.Menu mcustomer 
         Caption         =   "Customers"
      End
      Begin VB.Menu mland 
         Caption         =   "Land Development"
      End
      Begin VB.Menu memployees 
         Caption         =   "Employees"
      End
   End
   Begin VB.Menu mhelp 
      Caption         =   "Help"
      Begin VB.Menu mabout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Adt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdabout_Click(Index As Integer)
frmAbout.Show
Adt.Hide


End Sub

Private Sub cmdexit_Click(Index As Integer)
Unload Me

End Sub
