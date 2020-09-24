VERSION 5.00
Begin VB.Form frmemp 
   BackColor       =   &H80000007&
   Caption         =   "Employees Of ADTL"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7680
   LinkTopic       =   "Form2"
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
      Top             =   120
      Width           =   7095
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   5880
         Picture         =   "frmemp.frx":0000
         ScaleHeight     =   1335
         ScaleWidth      =   1215
         TabIndex        =   10
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   3720
         Picture         =   "frmemp.frx":28DC
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
         Picture         =   "frmemp.frx":5236
         ScaleHeight     =   1335
         ScaleWidth      =   1215
         TabIndex        =   7
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdmenu12 
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
         Index           =   4
         Left            =   2160
         TabIndex        =   6
         Top             =   4200
         Width           =   3135
      End
      Begin VB.CommandButton cmdcustomer 
         Caption         =   "&Customer Service Department"
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
         Index           =   2
         Left            =   2160
         TabIndex        =   5
         Top             =   3600
         Width           =   3135
      End
      Begin VB.CommandButton cmdfinance 
         Caption         =   "&Finance Department "
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
         Left            =   2160
         TabIndex        =   4
         Top             =   3000
         Width           =   3135
      End
      Begin VB.CommandButton cmdmarketing 
         Caption         =   "&Marketing Department"
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
         Index           =   0
         Left            =   2160
         TabIndex        =   3
         Top             =   2400
         Width           =   3135
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1575
         Left            =   0
         Picture         =   "frmemp.frx":7879
         ScaleHeight     =   1545
         ScaleWidth      =   3705
         TabIndex        =   1
         Top             =   0
         Width           =   3735
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "      Employee List"
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
         Left            =   2040
         TabIndex        =   2
         Top             =   1680
         Width           =   3255
      End
   End
End
Attribute VB_Name = "frmemp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)






End Sub



Private Sub cmdcustomer_Click(Index As Integer)
frmcustomer.Show
frmemp.Hide

End Sub

Private Sub cmdfinance_Click(Index As Integer)
frmfin.Show
frmemp.Hide


End Sub



Private Sub cmdmarketing_Click(Index As Integer)
frmmark.Show
frmemp.Hide

End Sub

Private Sub cmdmenu12_Click(Index As Integer)
frmadmin.Show
frmemp.Hide
End Sub
