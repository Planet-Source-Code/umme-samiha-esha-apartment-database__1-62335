VERSION 5.00
Begin VB.Form frmadmins 
   BackColor       =   &H80000012&
   Caption         =   "Admin Department Of ADTL"
   ClientHeight    =   6990
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   ScaleHeight     =   6990
   ScaleWidth      =   7800
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Advance 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   7455
      Begin VB.Frame Frame1 
         BackColor       =   &H80000013&
         Height          =   3615
         Left            =   120
         TabIndex        =   9
         Top             =   2280
         Width           =   7215
         Begin VB.Label Label1 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Caption         =   $"frmadmins.frx":0000
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3135
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   6975
         End
      End
      Begin VB.PictureBox Picture7 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   0
         Picture         =   "frmadmins.frx":0501
         ScaleHeight     =   735
         ScaleWidth      =   7455
         TabIndex        =   8
         Top             =   1440
         Width           =   7455
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   5280
         Picture         =   "frmadmins.frx":451D
         ScaleHeight     =   1335
         ScaleWidth      =   1095
         TabIndex        =   6
         Top             =   0
         Width           =   1095
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   4200
         Picture         =   "frmadmins.frx":6B60
         ScaleHeight     =   1335
         ScaleWidth      =   1095
         TabIndex        =   4
         Top             =   0
         Width           =   1095
         Begin VB.PictureBox Picture4 
            Height          =   255
            Left            =   1080
            ScaleHeight     =   255
            ScaleWidth      =   15
            TabIndex        =   5
            Top             =   360
            Width           =   15
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   6360
         Picture         =   "frmadmins.frx":94BA
         ScaleHeight     =   1335
         ScaleWidth      =   1215
         TabIndex        =   3
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdmenu3 
         Caption         =   "&Back Menu"
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
         Picture         =   "frmadmins.frx":BD96
         TabIndex        =   2
         Top             =   6120
         Width           =   2775
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   1575
         Left            =   0
         Picture         =   "frmadmins.frx":12066
         ScaleHeight     =   1575
         ScaleWidth      =   3735
         TabIndex        =   1
         Top             =   0
         Width           =   3735
         Begin VB.PictureBox Picture6 
            Height          =   15
            Left            =   0
            ScaleHeight     =   15
            ScaleWidth      =   3855
            TabIndex        =   7
            Top             =   1560
            Width           =   3855
         End
      End
   End
End
Attribute VB_Name = "frmadmins"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdmenu3_Click(Index As Integer)
frmadmin.Show
frmadmins.Hide
End Sub

