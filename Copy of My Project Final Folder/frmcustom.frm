VERSION 5.00
Begin VB.Form frmcustom 
   BackColor       =   &H80000007&
   Caption         =   "Customers Informations"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7680
   LinkTopic       =   "Form3"
   ScaleHeight     =   6495
   ScaleWidth      =   7680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   7095
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   3720
         Picture         =   "frmcustom.frx":0000
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
         Left            =   4800
         Picture         =   "frmcustom.frx":295A
         ScaleHeight     =   1335
         ScaleWidth      =   1095
         TabIndex        =   6
         Top             =   0
         Width           =   1095
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   5880
         Picture         =   "frmcustom.frx":4F9D
         ScaleHeight     =   1335
         ScaleWidth      =   1215
         TabIndex        =   5
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdcnew 
         Caption         =   "&New Subscribers"
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
         Left            =   2400
         TabIndex        =   4
         Top             =   2640
         Width           =   2295
      End
      Begin VB.CommandButton cmdupcomings 
         Caption         =   "&Old Customers"
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
         Left            =   2400
         TabIndex        =   3
         Top             =   3120
         Width           =   2295
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   1575
         Left            =   0
         Picture         =   "frmcustom.frx":7879
         ScaleHeight     =   1575
         ScaleWidth      =   3735
         TabIndex        =   2
         Top             =   0
         Width           =   3735
      End
      Begin VB.CommandButton cmdmenu9 
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
         Height          =   375
         Index           =   0
         Left            =   2400
         TabIndex        =   1
         Top             =   3600
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "    Customers Info"
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
         TabIndex        =   7
         Top             =   1800
         Width           =   3135
      End
   End
End
Attribute VB_Name = "frmcustom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcnew_Click(Index As Integer)
frmcnew.Show
frmcustom.Hide

End Sub



Private Sub cmdmenu9_Click(Index As Integer)
frmmenu.Show
frmcustom.Hide
End Sub

Private Sub cmdupcomings_Click(Index As Integer)
frmcold.Show
frmcustom.Hide

End Sub
