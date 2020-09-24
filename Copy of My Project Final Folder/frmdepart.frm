VERSION 5.00
Begin VB.Form frmdepart 
   BackColor       =   &H80000012&
   Caption         =   "Departments Of ADTL"
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
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      Begin VB.CommandButton cmdadminss 
         Caption         =   "&Administration And Logistic"
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
         Left            =   1920
         TabIndex        =   11
         Top             =   4200
         Width           =   3495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1575
         Left            =   0
         Picture         =   "frmdepart.frx":0000
         ScaleHeight     =   1545
         ScaleWidth      =   3705
         TabIndex        =   9
         Top             =   0
         Width           =   3735
      End
      Begin VB.CommandButton cmdmarketing 
         Caption         =   "&Marketing And Customer Service"
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
         Left            =   1920
         TabIndex        =   8
         Top             =   2400
         Width           =   3495
      End
      Begin VB.CommandButton cmdfinance 
         Caption         =   "&Accounts And Finance"
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
         Left            =   1920
         TabIndex        =   7
         Top             =   3000
         Width           =   3495
      End
      Begin VB.CommandButton cmdengineering 
         Caption         =   "&Engineering And Design"
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
         Left            =   1920
         TabIndex        =   6
         Top             =   3600
         Width           =   3495
      End
      Begin VB.CommandButton cmdmenu11 
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
         Left            =   1920
         TabIndex        =   5
         Top             =   4800
         Width           =   3495
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   4920
         Picture         =   "frmdepart.frx":62D0
         ScaleHeight     =   1335
         ScaleWidth      =   1095
         TabIndex        =   4
         Top             =   0
         Width           =   1095
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   3840
         Picture         =   "frmdepart.frx":8913
         ScaleHeight     =   1335
         ScaleWidth      =   1215
         TabIndex        =   2
         Top             =   0
         Width           =   1215
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
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   6000
         Picture         =   "frmdepart.frx":B26D
         ScaleHeight     =   1335
         ScaleWidth      =   1215
         TabIndex        =   1
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "         Departments"
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
         Left            =   1800
         TabIndex        =   10
         Top             =   1680
         Width           =   3735
      End
   End
End
Attribute VB_Name = "frmdepart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdadminss_Click(Index As Integer)
frmadminist.Show
frmdepart.Hide

End Sub

Private Sub cmdengineering_Click(Index As Integer)
frmengineering.Show
frmdepart.Hide

End Sub

Private Sub cmdfinance_Click(Index As Integer)
frmfinance.Show
frmdepart.Hide

End Sub



Private Sub cmdmarketing_Click(Index As Integer)
frmmarket.Show
frmdepart.Hide

End Sub

Private Sub cmdmenu11_Click(Index As Integer)
frmmenu.Show
frmdepart.Hide
End Sub
