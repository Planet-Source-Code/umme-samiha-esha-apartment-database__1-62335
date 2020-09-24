VERSION 5.00
Begin VB.Form frmadmin 
   BackColor       =   &H80000008&
   Caption         =   "Administration Of Advance Development Technologies Ltd."
   ClientHeight    =   6525
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   ScaleHeight     =   6525
   ScaleWidth      =   7680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Advance 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   6050
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   7340
      Begin VB.CommandButton Command1 
         Caption         =   "&Employee Informations"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2280
         TabIndex        =   10
         Top             =   4080
         Width           =   2775
      End
      Begin VB.CommandButton cmdadmin1 
         BackColor       =   &H80000013&
         Caption         =   "&Administration"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   2280
         Picture         =   "frmadmin.frx":0000
         TabIndex        =   8
         Top             =   2640
         Width           =   2775
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   1575
         Left            =   0
         Picture         =   "frmadmin.frx":62D0
         ScaleHeight     =   1575
         ScaleWidth      =   3855
         TabIndex        =   7
         Top             =   0
         Width           =   3855
      End
      Begin VB.CommandButton cmdmenu1 
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
         Height          =   615
         Index           =   1
         Left            =   2280
         Picture         =   "frmadmin.frx":C5A0
         TabIndex        =   6
         Top             =   4920
         Width           =   2775
      End
      Begin VB.CommandButton cmdprofile 
         Caption         =   "&Company Profiles"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   3
         Left            =   2280
         Picture         =   "frmadmin.frx":12870
         TabIndex        =   5
         Top             =   3360
         Width           =   2775
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   6240
         Picture         =   "frmadmin.frx":18B40
         ScaleHeight     =   1335
         ScaleWidth      =   1215
         TabIndex        =   4
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   4080
         Picture         =   "frmadmin.frx":1B41C
         ScaleHeight     =   1335
         ScaleWidth      =   1095
         TabIndex        =   2
         Top             =   0
         Width           =   1095
         Begin VB.PictureBox Picture4 
            Height          =   255
            Left            =   1080
            ScaleHeight     =   255
            ScaleWidth      =   15
            TabIndex        =   3
            Top             =   360
            Width           =   15
         End
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   5160
         Picture         =   "frmadmin.frx":1DD76
         ScaleHeight     =   1335
         ScaleWidth      =   1095
         TabIndex        =   1
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Company Profiles Of ADT"
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
         TabIndex        =   9
         Top             =   1800
         Width           =   5175
      End
   End
End
Attribute VB_Name = "frmadmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdmenu_Click(Index As Integer)


End Sub





Private Sub cmdadmin1_Click(Index As Integer)
frmadmins.Show
frmadmin.Hide
End Sub

Private Sub cmdmenu1_Click(Index As Integer)
frmadt.Show
frmadmin.Hide
End Sub

Private Sub cmdprofile_Click(Index As Integer)
frmprofile.Show
frmadmin.Hide

End Sub

Private Sub Command1_Click()
frmemp.Show
frmadmin.Hide

End Sub
