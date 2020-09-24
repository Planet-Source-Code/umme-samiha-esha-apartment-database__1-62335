VERSION 5.00
Begin VB.Form frmland 
   BackColor       =   &H80000008&
   Caption         =   "Lands Department Of ADTL"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7680
   LinkTopic       =   "Form3"
   ScaleHeight     =   6495
   ScaleWidth      =   7680
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   120
      Picture         =   "frmland.frx":0000
      ScaleHeight     =   1455
      ScaleWidth      =   3735
      TabIndex        =   1
      Top             =   120
      Width           =   3735
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   3840
         Picture         =   "frmland.frx":62D0
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
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   4920
         Picture         =   "frmland.frx":8C2A
         ScaleHeight     =   1335
         ScaleWidth      =   1095
         TabIndex        =   7
         Top             =   0
         Width           =   1095
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   6000
         Picture         =   "frmland.frx":B26D
         ScaleHeight     =   1335
         ScaleWidth      =   1215
         TabIndex        =   6
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdmenu16 
         Caption         =   "&Main Menu"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         TabIndex        =   5
         Top             =   3720
         Width           =   2535
      End
      Begin VB.CommandButton cmdnew 
         Caption         =   "&Upcoming Projects"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   2280
         TabIndex        =   4
         Top             =   3120
         Width           =   2535
      End
      Begin VB.CommandButton cmdowner 
         Caption         =   "&Land Owners List"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         TabIndex        =   3
         Top             =   2520
         Width           =   2535
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "   Land Developement"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   1680
         TabIndex        =   2
         Top             =   1680
         Width           =   3735
      End
   End
End
Attribute VB_Name = "frmland"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()


End Sub



Private Sub cmdmenu16_Click()
frmmenu.Show
frmland.Hide
End Sub

Private Sub cmdnew_Click(Index As Integer)
frmupcome.Show
frmland.Hide

End Sub

Private Sub cmdowner_Click()
frmowner.Show
frmland.Hide

End Sub

