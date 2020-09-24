VERSION 5.00
Begin VB.Form Form8 
   BackColor       =   &H80000008&
   Caption         =   "Reports Of Upcoming Land Projects"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7425
   LinkTopic       =   "Form8"
   ScaleHeight     =   6390
   ScaleWidth      =   7425
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Advance 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   5925
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7040
      Begin VB.Frame Frame1 
         BackColor       =   &H80000013&
         Height          =   1815
         Left            =   1320
         TabIndex        =   7
         Top             =   2640
         Width           =   4455
         Begin VB.CommandButton Command2 
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
            Left            =   960
            TabIndex        =   9
            Top             =   1080
            Width           =   2415
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&Report"
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
            Left            =   960
            TabIndex        =   8
            Top             =   360
            Width           =   2415
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   1575
         Left            =   0
         Picture         =   "Form8.frx":0000
         ScaleHeight     =   1575
         ScaleWidth      =   3855
         TabIndex        =   5
         Top             =   0
         Width           =   3855
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   6000
         Picture         =   "Form8.frx":62D0
         ScaleHeight     =   1335
         ScaleWidth      =   1215
         TabIndex        =   4
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   3840
         Picture         =   "Form8.frx":8BAC
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
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   4920
         Picture         =   "Form8.frx":B506
         ScaleHeight     =   1335
         ScaleWidth      =   1095
         TabIndex        =   1
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "      Upcoming Project Reports"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000013&
         Height          =   615
         Left            =   480
         TabIndex        =   6
         Top             =   1800
         Width           =   6135
      End
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim res As Integer
    res = MsgBox(" Are You Sure....You Want To Go End And The See The Upcoming Projects Report ?" _
    , vbYesNo, "Delete Record?")
    If res = vbYes Then
    DataReport1.Show
    Unload Me
    Else
    
    
    Exit Sub
    End If
End Sub

Private Sub Command2_Click()
frmadt.Show
Unload Me

End Sub
