VERSION 5.00
Begin VB.Form frmadt 
   BackColor       =   &H80000006&
   Caption         =   "Advanced Develpment Technologies Ltd."
   ClientHeight    =   6225
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   ScaleHeight     =   6225
   ScaleWidth      =   7680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Advance 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   5925
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   7040
      Begin VB.CommandButton Command2 
         Caption         =   "&Upcoming Project Report"
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
         Left            =   2280
         TabIndex        =   14
         Top             =   4440
         Width           =   2895
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Customers Login"
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
         Left            =   1560
         TabIndex        =   13
         Top             =   3840
         Width           =   1815
      End
      Begin VB.CommandButton cmdcreate 
         Caption         =   "&Created By"
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
         Index           =   1
         Left            =   3720
         Picture         =   "ADTLfinal.frx":0000
         TabIndex        =   12
         Top             =   3840
         Width           =   1935
      End
      Begin VB.CommandButton cmdcontact 
         Caption         =   "&Contact"
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
         Left            =   3720
         Picture         =   "ADTLfinal.frx":62D0
         TabIndex        =   11
         Top             =   3240
         Width           =   1935
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   4920
         Picture         =   "ADTLfinal.frx":C5A0
         ScaleHeight     =   1335
         ScaleWidth      =   1095
         TabIndex        =   10
         Top             =   0
         Width           =   1095
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   3840
         Picture         =   "ADTLfinal.frx":EBE3
         ScaleHeight     =   1335
         ScaleWidth      =   1095
         TabIndex        =   8
         Top             =   0
         Width           =   1095
         Begin VB.PictureBox Picture4 
            Height          =   255
            Left            =   1080
            ScaleHeight     =   255
            ScaleWidth      =   15
            TabIndex        =   9
            Top             =   360
            Width           =   15
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   6000
         Picture         =   "ADTLfinal.frx":1153D
         ScaleHeight     =   1335
         ScaleWidth      =   1215
         TabIndex        =   7
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdabout 
         Caption         =   "&About"
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
         Index           =   3
         Left            =   3720
         Picture         =   "ADTLfinal.frx":13E19
         TabIndex        =   5
         Top             =   2640
         Width           =   1935
      End
      Begin VB.CommandButton cmdadmin 
         Caption         =   "&Admin Login"
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
         Index           =   2
         Left            =   1560
         Picture         =   "ADTLfinal.frx":1A0E9
         TabIndex        =   4
         Top             =   3240
         Width           =   1815
      End
      Begin VB.CommandButton cmdexit 
         Caption         =   "&Exit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   2280
         Picture         =   "ADTLfinal.frx":203B9
         TabIndex        =   3
         Top             =   4920
         Width           =   2895
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   1575
         Left            =   0
         Picture         =   "ADTLfinal.frx":26689
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
         Left            =   1560
         Picture         =   "ADTLfinal.frx":2C959
         TabIndex        =   1
         Top             =   2640
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "      WELCOME TO ADTL LTD."
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
         Left            =   480
         TabIndex        =   6
         Top             =   1680
         Width           =   6135
      End
   End
   Begin VB.Menu mfile 
      Caption         =   "File"
      Begin VB.Menu mchange 
         Caption         =   "Admin Login"
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
      Begin VB.Menu mproj 
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
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mrepot 
      Caption         =   "Report"
      Begin VB.Menu mcustomers 
         Caption         =   "Upcoming Project Report"
      End
   End
   Begin VB.Menu mhelp 
      Caption         =   "Help"
      Begin VB.Menu mcontact 
         Caption         =   "Contact"
      End
      Begin VB.Menu mcreate 
         Caption         =   "Created By"
      End
      Begin VB.Menu mabout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmadt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdabout_Click(Index As Integer)
frmAbout.Show
frmadt.Hide

End Sub

Private Sub cmdadmin_Click(Index As Integer)
    frmLogin.Show
End Sub

Private Sub cmdcontact_Click(Index As Integer)
frmcontact.Show
frmadt.Hide

End Sub

Private Sub cmdcreate_Click(Index As Integer)
frmcreate.Show
frmadt.Hide

End Sub

Private Sub cmdexit_Click(Index As Integer)
    Unload frmadt
    End
    
 
       
End Sub

Private Sub cmdmenu_Click(Index As Integer)
frmmenu.Show
frmadt.Hide


End Sub

Private Sub Command1_Click()
frmcustom.Show
frmadt.Hide

End Sub



Private Sub Command2_Click()
Form8.Show
Unload Me


End Sub

Private Sub Command3_Click()
Form2.Show
Unload Me

End Sub

Private Sub mabout_Click()
    frmAbout.Show
    frmadt.Hide
    
End Sub

Private Sub mchange_Click()
    frmLogin.Show
    frmadt.Hide
    
    
End Sub

Private Sub mcontact_Click()
frmcontact.Show
frmadt.Hide

End Sub

Private Sub mcreate_Click()
frmcreate.Show
frmadt.Hide

End Sub

Private Sub mcustomer_Click()
frmcustom.Show
frmadt.Hide

End Sub

Private Sub mcustomers_Click()
Form8.Show
Unload Me


End Sub

Private Sub mdepartment_Click()
frmdepart.Show
frmadt.Hide

End Sub

Private Sub memployees_Click()
frmemp.Show
frmadt.Hide

End Sub

Private Sub mexit_Click()
End

End Sub

Private Sub mland_Click()
frmland.Show
frmadt.Hide

End Sub

Private Sub mproj_Click()
frmproject.Show
frmadt.Hide

End Sub

Private Sub mprojects_Click()

End Sub
