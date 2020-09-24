VERSION 5.00
Begin VB.Form frmcreate 
   BackColor       =   &H80000012&
   Caption         =   "ADTL Project Created By  Umme Samiha  (03.01.04.108)"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   7680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Advance 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7335
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   3840
         Picture         =   "frmcreate.frx":0000
         ScaleHeight     =   1335
         ScaleWidth      =   1095
         TabIndex        =   10
         Top             =   0
         Width           =   1095
         Begin VB.PictureBox Picture4 
            Height          =   255
            Left            =   1080
            ScaleHeight     =   255
            ScaleWidth      =   15
            TabIndex        =   11
            Top             =   360
            Width           =   15
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000013&
         Height          =   3735
         Left            =   0
         TabIndex        =   6
         Top             =   1680
         Width           =   7335
         Begin VB.PictureBox Picture6 
            BackColor       =   &H80000013&
            BorderStyle     =   0  'None
            Height          =   3735
            Left            =   4200
            Picture         =   "frmcreate.frx":295A
            ScaleHeight     =   3735
            ScaleWidth      =   3135
            TabIndex        =   7
            Top             =   0
            Width           =   3135
            Begin VB.Label Label2 
               Caption         =   "Label2"
               Height          =   15
               Left            =   0
               TabIndex        =   8
               Top             =   1920
               Width           =   2295
            End
         End
         Begin VB.Label Label1 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Caption         =   $"frmcreate.frx":7F94
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2175
            Index           =   0
            Left            =   0
            TabIndex        =   9
            Top             =   360
            Width           =   4215
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   1575
         Left            =   0
         Picture         =   "frmcreate.frx":8206
         ScaleHeight     =   1575
         ScaleWidth      =   3855
         TabIndex        =   4
         Top             =   0
         Width           =   3855
      End
      Begin VB.CommandButton cmdmenu8 
         Caption         =   "&Main Menu"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   4800
         Picture         =   "frmcreate.frx":E4D6
         TabIndex        =   3
         Top             =   5520
         Width           =   1455
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   6000
         Picture         =   "frmcreate.frx":147A6
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
         Picture         =   "frmcreate.frx":17082
         ScaleHeight     =   1335
         ScaleWidth      =   1095
         TabIndex        =   1
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "     ABOUT UMME SAMIHA"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000013&
         Height          =   375
         Index           =   1
         Left            =   0
         TabIndex        =   5
         Top             =   5520
         Width           =   4095
      End
   End
End
Attribute VB_Name = "frmcreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdmenu8_Click(Index As Integer)
frmadt.Show
frmcreate.Hide
End Sub
