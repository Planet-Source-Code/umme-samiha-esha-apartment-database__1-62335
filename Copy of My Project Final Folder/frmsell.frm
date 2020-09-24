VERSION 5.00
Begin VB.Form frmsell 
   BackColor       =   &H80000007&
   Caption         =   "Selling Informations Of Lands"
   ClientHeight    =   6525
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7215
   LinkTopic       =   "Form3"
   ScaleHeight     =   6525
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      Begin VB.Frame Frame2 
         BackColor       =   &H80000013&
         Caption         =   "Selling Informations"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   3855
         Left            =   120
         TabIndex        =   6
         Top             =   1680
         Width           =   6735
         Begin VB.Frame Frame4 
            BackColor       =   &H80000013&
            Caption         =   "Employees"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   1335
            Left            =   240
            TabIndex        =   8
            Top             =   2160
            Width           =   6255
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H80000013&
            Caption         =   "Customers"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   1815
            Left            =   240
            TabIndex        =   7
            Top             =   360
            Width           =   6255
            Begin VB.TextBox txtname 
               Height          =   285
               Index           =   3
               Left            =   4560
               TabIndex        =   16
               Top             =   480
               Width           =   1455
            End
            Begin VB.TextBox txtname 
               Height          =   285
               Index           =   2
               Left            =   3240
               TabIndex        =   14
               Top             =   480
               Width           =   1215
            End
            Begin VB.TextBox txtname 
               Height          =   285
               Index           =   1
               Left            =   1680
               TabIndex        =   12
               Top             =   480
               Width           =   1455
            End
            Begin VB.TextBox txtname 
               Height          =   285
               Index           =   0
               Left            =   120
               TabIndex        =   9
               Top             =   480
               Width           =   1455
            End
            Begin VB.Label Label1 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " First Name"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   3
               Left            =   4560
               TabIndex        =   15
               Top             =   240
               Width           =   1455
            End
            Begin VB.Label Label1 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "   Gender"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   2
               Left            =   3240
               TabIndex        =   13
               Top             =   240
               Width           =   1215
            End
            Begin VB.Label Label1 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " Last Name"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   1
               Left            =   1680
               TabIndex        =   11
               Top             =   240
               Width           =   1455
            End
            Begin VB.Label Label1 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " First Name"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   10
               Top             =   240
               Width           =   1455
            End
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   1455
         Left            =   5880
         Picture         =   "frmsell.frx":0000
         ScaleHeight     =   1455
         ScaleWidth      =   1215
         TabIndex        =   5
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   4800
         Picture         =   "frmsell.frx":28DC
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
         Left            =   3720
         Picture         =   "frmsell.frx":4F1F
         ScaleHeight     =   1335
         ScaleWidth      =   1095
         TabIndex        =   2
         Top             =   0
         Width           =   1095
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
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   1575
         Left            =   0
         Picture         =   "frmsell.frx":7879
         ScaleHeight     =   1575
         ScaleWidth      =   3735
         TabIndex        =   1
         Top             =   0
         Width           =   3735
      End
   End
End
Attribute VB_Name = "frmsell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
