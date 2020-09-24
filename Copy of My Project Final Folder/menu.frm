VERSION 5.00
Begin VB.Form frmmenu 
   BackColor       =   &H80000007&
   Caption         =   "Menu List"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   7680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5775
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   4320
         Picture         =   "menu.frx":0000
         ScaleHeight     =   1335
         ScaleWidth      =   1095
         TabIndex        =   9
         Top             =   0
         Width           =   1095
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   5400
         Picture         =   "menu.frx":2643
         ScaleHeight     =   1335
         ScaleWidth      =   1215
         TabIndex        =   8
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdmenu19 
         Caption         =   "&Back To Menu"
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
         Index           =   5
         Left            =   2520
         TabIndex        =   7
         Top             =   4320
         Width           =   1815
      End
      Begin VB.CommandButton cmdprojects 
         Caption         =   "&Projects"
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
         Index           =   4
         Left            =   2520
         TabIndex        =   6
         Top             =   3360
         Width           =   1815
      End
      Begin VB.CommandButton cmdcustomer 
         Caption         =   "&Customer"
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
         Left            =   2520
         TabIndex        =   5
         Top             =   2400
         Width           =   1815
      End
      Begin VB.CommandButton cmdland 
         Caption         =   "& lands "
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
         Left            =   2520
         TabIndex        =   4
         Top             =   3840
         Width           =   1815
      End
      Begin VB.CommandButton cmddepartment 
         Caption         =   "&Department"
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
         Left            =   2520
         TabIndex        =   3
         Top             =   2880
         Width           =   1815
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1575
         Left            =   0
         Picture         =   "menu.frx":4F1F
         ScaleHeight     =   1575
         ScaleWidth      =   3735
         TabIndex        =   1
         Top             =   0
         Width           =   3735
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "       Menu  List"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000013&
         Height          =   615
         Left            =   2040
         TabIndex        =   2
         Top             =   1680
         Width           =   3015
      End
   End
End
Attribute VB_Name = "frmmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdcustomer_Click(Index As Integer)
frmcustom.Show
frmmenu.Hide

End Sub

Private Sub cmddepartment_Click(Index As Integer)
frmdepart.Show
frmmenu.Hide

End Sub

Private Sub cmdemployee_Click(Index As Integer)
frmemp.Show
frmmenu.Hide
End Sub

Private Sub cmdland_Click(Index As Integer)
frmland.Show
frmmenu.Hide


End Sub

Private Sub cmdmenu19_Click(Index As Integer)
frmadt.Show
frmmenu.Hide
End Sub

Private Sub cmdprojects_Click(Index As Integer)
frmproject.Show
frmmenu.Hide

End Sub
