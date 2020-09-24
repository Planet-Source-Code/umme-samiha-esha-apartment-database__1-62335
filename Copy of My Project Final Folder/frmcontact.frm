VERSION 5.00
Begin VB.Form frmcontact 
   BackColor       =   &H80000012&
   Caption         =   "Contact Address of ADTL"
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
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   4920
         Picture         =   "frmcontact.frx":0000
         ScaleHeight     =   1335
         ScaleWidth      =   1095
         TabIndex        =   8
         Top             =   0
         Width           =   1095
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   6000
         Picture         =   "frmcontact.frx":2643
         ScaleHeight     =   1335
         ScaleWidth      =   1215
         TabIndex        =   7
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdmenu7 
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
         Index           =   5
         Left            =   2640
         TabIndex        =   6
         Top             =   5400
         Width           =   1815
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   1455
         Left            =   0
         Picture         =   "frmcontact.frx":4F1F
         ScaleHeight     =   1455
         ScaleWidth      =   3735
         TabIndex        =   5
         Top             =   0
         Width           =   3735
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   3960
         Picture         =   "frmcontact.frx":B1EF
         ScaleHeight     =   1335
         ScaleWidth      =   1095
         TabIndex        =   3
         Top             =   0
         Width           =   1095
         Begin VB.PictureBox Picture4 
            Height          =   255
            Left            =   1080
            ScaleHeight     =   255
            ScaleWidth      =   15
            TabIndex        =   4
            Top             =   360
            Width           =   15
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000013&
         Height          =   2895
         Left            =   240
         TabIndex        =   1
         Top             =   2400
         Width           =   6615
         Begin VB.Label Label1 
            BackColor       =   &H80000018&
            Caption         =   $"frmcontact.frx":DB49
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2415
            Index           =   0
            Left            =   360
            TabIndex        =   2
            Top             =   240
            Width           =   6015
         End
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "      Contact Info"
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
         Height          =   495
         Index           =   1
         Left            =   2160
         TabIndex        =   9
         Top             =   1680
         Width           =   3015
      End
   End
End
Attribute VB_Name = "frmcontact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdmenu7_Click(Index As Integer)
frmadt.Show
frmcontact.Hide
End Sub
