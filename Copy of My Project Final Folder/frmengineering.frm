VERSION 5.00
Begin VB.Form frmengineering 
   BackColor       =   &H80000012&
   Caption         =   "Engineering & Design Department of ADTL"
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
      Top             =   240
      Width           =   7215
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   5040
         Picture         =   "frmengineering.frx":0000
         ScaleHeight     =   1335
         ScaleWidth      =   1095
         TabIndex        =   11
         Top             =   0
         Width           =   1095
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   6120
         Picture         =   "frmengineering.frx":2643
         ScaleHeight     =   1335
         ScaleWidth      =   1335
         TabIndex        =   10
         Top             =   0
         Width           =   1335
      End
      Begin VB.CommandButton cmdmenu13 
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
         Height          =   375
         Index           =   5
         Left            =   2760
         TabIndex        =   9
         Top             =   5400
         Width           =   1815
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   1455
         Left            =   0
         Picture         =   "frmengineering.frx":4F1F
         ScaleHeight     =   1455
         ScaleWidth      =   3975
         TabIndex        =   8
         Top             =   0
         Width           =   3975
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   3960
         Picture         =   "frmengineering.frx":B1EF
         ScaleHeight     =   1335
         ScaleWidth      =   1095
         TabIndex        =   6
         Top             =   0
         Width           =   1095
         Begin VB.PictureBox Picture4 
            Height          =   255
            Left            =   1080
            ScaleHeight     =   255
            ScaleWidth      =   15
            TabIndex        =   7
            Top             =   360
            Width           =   15
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000013&
         Height          =   3015
         Left            =   120
         TabIndex        =   1
         Top             =   2280
         Width           =   6975
         Begin VB.PictureBox Picture6 
            BackColor       =   &H80000013&
            BorderStyle     =   0  'None
            Height          =   1335
            Index           =   1
            Left            =   5400
            Picture         =   "frmengineering.frx":DB49
            ScaleHeight     =   1335
            ScaleWidth      =   1335
            TabIndex        =   13
            Top             =   480
            Width           =   1335
            Begin VB.Label Label2 
               Caption         =   "Label2"
               Height          =   15
               Index           =   1
               Left            =   0
               TabIndex        =   14
               Top             =   1920
               Width           =   2295
            End
         End
         Begin VB.PictureBox Picture6 
            BackColor       =   &H80000013&
            BorderStyle     =   0  'None
            Height          =   1215
            Index           =   0
            Left            =   4320
            Picture         =   "frmengineering.frx":10B10
            ScaleHeight     =   1215
            ScaleWidth      =   1335
            TabIndex        =   2
            Top             =   480
            Width           =   1335
            Begin VB.Label Label2 
               Caption         =   "Label2"
               Height          =   15
               Index           =   0
               Left            =   0
               TabIndex        =   3
               Top             =   1920
               Width           =   2295
            End
         End
         Begin VB.Label Label3 
            BackColor       =   &H80000013&
            Caption         =   " Design Department "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4560
            TabIndex        =   5
            Top             =   1920
            Width           =   1935
         End
         Begin VB.Label Label1 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Caption         =   $"frmengineering.frx":132BB
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2055
            Index           =   0
            Left            =   120
            TabIndex        =   4
            Top             =   360
            Width           =   4095
         End
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "     Engineering And Design"
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
         Left            =   1560
         TabIndex        =   12
         Top             =   1560
         Width           =   4575
      End
   End
End
Attribute VB_Name = "frmengineering"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdmenu13_Click(Index As Integer)
frmdepart.Show
frmengineering.Hide
End Sub
