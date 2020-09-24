VERSION 5.00
Begin VB.Form frmproject 
   BackColor       =   &H80000012&
   Caption         =   "Projects"
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
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6735
      Begin VB.CommandButton cmdmenu23 
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
         TabIndex        =   8
         Top             =   4080
         Width           =   2295
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   1575
         Left            =   0
         Picture         =   "frmproject.frx":0000
         ScaleHeight     =   1575
         ScaleWidth      =   3735
         TabIndex        =   6
         Top             =   0
         Width           =   3735
      End
      Begin VB.CommandButton cmdupcomings 
         Caption         =   "&Upcoming Project"
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
         TabIndex        =   5
         Top             =   3120
         Width           =   2295
      End
      Begin VB.CommandButton cmdongoing 
         Caption         =   "&Ongoing Project"
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
      Begin VB.CommandButton cmdcomplete 
         Caption         =   "&Completed Project"
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
         Left            =   2400
         TabIndex        =   3
         Top             =   3600
         Width           =   2295
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   5400
         Picture         =   "frmproject.frx":62D0
         ScaleHeight     =   1335
         ScaleWidth      =   1215
         TabIndex        =   2
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   4320
         Picture         =   "frmproject.frx":8BAC
         ScaleHeight     =   1335
         ScaleWidth      =   1095
         TabIndex        =   1
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "      Project List"
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
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmproject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdcomplete_Click(Index As Integer)
frmcomplete.Show
frmproject.Hide

End Sub



Private Sub cmdmenu23_Click(Index As Integer)
frmmenu.Show
frmproject.Hide
End Sub

Private Sub cmdongoing_Click(Index As Integer)
frmongoing.Show
frmproject.Hide

End Sub

Private Sub cmdupcomings_Click(Index As Integer)
frmupcoming.Show
frmproject.Hide

End Sub
