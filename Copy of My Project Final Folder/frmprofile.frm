VERSION 5.00
Begin VB.Form frmprofile 
   BackColor       =   &H80000008&
   Caption         =   "Company Profiles of ADT "
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
      Begin VB.Frame Frame2 
         BackColor       =   &H80000013&
         Height          =   2895
         Left            =   120
         TabIndex        =   8
         Top             =   2640
         Width           =   6855
         Begin VB.Label Label1 
            BackColor       =   &H80000018&
            Caption         =   $"frmprofile.frx":0000
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2655
            Left            =   120
            TabIndex        =   9
            Top             =   120
            Width           =   6615
         End
      End
      Begin VB.PictureBox Picture6 
         Height          =   1095
         Left            =   0
         Picture         =   "frmprofile.frx":02B0
         ScaleHeight     =   1035
         ScaleWidth      =   7035
         TabIndex        =   7
         Top             =   1440
         Width           =   7095
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   3960
         Picture         =   "frmprofile.frx":4011
         ScaleHeight     =   1335
         ScaleWidth      =   1095
         TabIndex        =   5
         Top             =   0
         Width           =   1095
         Begin VB.PictureBox Picture4 
            Height          =   255
            Left            =   1080
            ScaleHeight     =   255
            ScaleWidth      =   15
            TabIndex        =   6
            Top             =   360
            Width           =   15
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   1575
         Left            =   0
         Picture         =   "frmprofile.frx":696B
         ScaleHeight     =   1575
         ScaleWidth      =   3735
         TabIndex        =   4
         Top             =   0
         Width           =   3735
      End
      Begin VB.CommandButton cmdmenu22 
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
         TabIndex        =   3
         Top             =   5640
         Width           =   1815
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   6000
         Picture         =   "frmprofile.frx":CC3B
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
         Left            =   4920
         Picture         =   "frmprofile.frx":F517
         ScaleHeight     =   1335
         ScaleWidth      =   1095
         TabIndex        =   1
         Top             =   0
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmprofile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Text1_Change()

End Sub

Private Sub cmdmenu22_Click(Index As Integer)
frmadmin.Show
frmprofile.Hide
End Sub
