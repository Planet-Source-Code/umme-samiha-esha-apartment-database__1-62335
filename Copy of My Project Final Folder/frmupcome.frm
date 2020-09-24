VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmupcome 
   BackColor       =   &H80000007&
   Caption         =   "Upcoming Land Informations"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7680
   LinkTopic       =   "Form3"
   ScaleHeight     =   6495
   ScaleWidth      =   7680
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   120
      Picture         =   "frmupcome.frx":0000
      ScaleHeight     =   1575
      ScaleWidth      =   3735
      TabIndex        =   1
      Top             =   120
      Width           =   3735
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      Begin VB.CommandButton Command6 
         Caption         =   "&Search"
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
         Left            =   4320
         TabIndex        =   26
         Top             =   4800
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Delete"
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
         Left            =   3480
         TabIndex        =   20
         Top             =   4800
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Edit"
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
         Left            =   2640
         TabIndex        =   19
         Top             =   4800
         Width           =   855
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Cancel"
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
         Left            =   3480
         TabIndex        =   18
         Top             =   4800
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Update"
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
         Left            =   2640
         TabIndex        =   17
         Top             =   4800
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Add"
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
         Left            =   1800
         TabIndex        =   16
         Top             =   4800
         Width           =   855
      End
      Begin VB.CommandButton cmdmenu24 
         Caption         =   "&Back"
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
         Left            =   5160
         TabIndex        =   15
         Top             =   4800
         Width           =   975
      End
      Begin MSAdodcLib.Adodc adodc1 
         Height          =   375
         Left            =   600
         Top             =   4800
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   2
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   $"frmupcome.frx":62D0
         OLEDBString     =   $"frmupcome.frx":63C3
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "Upcomingland"
         Caption         =   "adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000013&
         Caption         =   "Upcoming Project Informations"
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
         Height          =   2895
         Index           =   0
         Left            =   480
         TabIndex        =   6
         Top             =   1800
         Width           =   5775
         Begin VB.Frame Frame2 
            BackColor       =   &H80000013&
            Caption         =   "Upcoming Project Informations"
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
            Height          =   2895
            Index           =   1
            Left            =   0
            TabIndex        =   8
            Top             =   0
            Width           =   5775
            Begin VB.Frame Frame2 
               BackColor       =   &H80000013&
               Caption         =   "Upcoming Project Informations"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000D&
               Height          =   2895
               Index           =   2
               Left            =   0
               TabIndex        =   9
               Top             =   0
               Width           =   5775
               Begin VB.TextBox Text5 
                  DataField       =   "Sproject"
                  DataSource      =   "adodc1"
                  Height          =   375
                  Left            =   1920
                  TabIndex        =   25
                  Top             =   2280
                  Width           =   3495
               End
               Begin VB.TextBox Text4 
                  DataField       =   "Saddress"
                  DataSource      =   "adodc1"
                  Height          =   375
                  Left            =   1920
                  TabIndex        =   24
                  Top             =   1800
                  Width           =   3495
               End
               Begin VB.TextBox Text3 
                  DataField       =   "Ssize"
                  DataSource      =   "adodc1"
                  Height          =   375
                  Left            =   1920
                  TabIndex        =   23
                  Top             =   1320
                  Width           =   3495
               End
               Begin VB.TextBox Text2 
                  DataField       =   "Slocation"
                  DataSource      =   "adodc1"
                  Height          =   375
                  Left            =   1920
                  TabIndex        =   22
                  Top             =   840
                  Width           =   3495
               End
               Begin VB.TextBox Text1 
                  DataField       =   "Sarea"
                  DataSource      =   "adodc1"
                  Height          =   375
                  Left            =   1920
                  TabIndex        =   21
                  Top             =   360
                  Width           =   3495
               End
               Begin VB.Label Label1 
                  Caption         =   "   Project      #"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000002&
                  Height          =   375
                  Index           =   5
                  Left            =   120
                  TabIndex        =   14
                  Top             =   2280
                  Width           =   1575
               End
               Begin VB.Label Label1 
                  Caption         =   "   Address     #"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000002&
                  Height          =   375
                  Index           =   4
                  Left            =   120
                  TabIndex        =   13
                  Top             =   1800
                  Width           =   1575
               End
               Begin VB.Label Label1 
                  Caption         =   "   Size            #"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000002&
                  Height          =   375
                  Index           =   3
                  Left            =   120
                  TabIndex        =   12
                  Top             =   1320
                  Width           =   1575
               End
               Begin VB.Label Label1 
                  Caption         =   "   Location     #"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000002&
                  Height          =   375
                  Index           =   2
                  Left            =   120
                  TabIndex        =   11
                  Top             =   840
                  Width           =   1575
               End
               Begin VB.Label Label1 
                  Caption         =   "   Area           #"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000002&
                  Height          =   375
                  Index           =   1
                  Left            =   120
                  TabIndex        =   10
                  Top             =   360
                  Width           =   1575
               End
            End
         End
         Begin VB.Label Label1 
            Caption         =   "Name              #"
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
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   7
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   1455
         Left            =   5760
         Picture         =   "frmupcome.frx":64B6
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
         Picture         =   "frmupcome.frx":8D92
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
         Picture         =   "frmupcome.frx":B3D5
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
   End
End
Attribute VB_Name = "frmupcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdmenu24_Click()
frmland.Show
frmupcome.Hide
End Sub


Private Sub Command1_Click()
Command2.Visible = False
Command5.Visible = False
Command3.Visible = True
Command4.Visible = True
Adodc1.Recordset.MoveLast
Adodc1.Recordset.AddNew
Command3.Enabled = True
Command4.Enabled = True
End Sub

Private Sub Command2_Click()
Dim x As Integer
    x = MsgBox(" Do You Want to Edit the current record?" _
    , vbYesNo, "Edit Record?")
    If x = vbYes Then
    Text1 = ""
    Text2 = ""
    Text3 = ""
    Text4 = ""
    Text5 = ""
    Command1.Visible = False
    Command3.Visible = True
    Command4.Visible = True
    Command2.Visible = False
    Command5.Visible = False
    
   Adodc1.Recordset.Update
    x = MsgBox("You Are Now Ready To Edit Your Informations...Go For Edit", vbOKOnly, "ATTENTION !")
    If x = vbOKOnly Then
    End
    End If
    Else
   Exit Sub
    End If
    Command1.Visible = True
    
End Sub

Private Sub Command3_Click()
Adodc1.Recordset.Update
Command3.Visible = False
Command4.Visible = False
Command2.Visible = True
Command5.Visible = True
Command1.Visible = True
End Sub

Private Sub Command4_Click()

Adodc1.Refresh

Adodc1.Recordset.Cancel

Command4.Visible = False
Command2.Visible = True
Command3.Visible = False
Command5.Visible = True
Command1.Visible = True

End Sub

Private Sub Command5_Click()
Dim res As Integer
    res = MsgBox(" Do You Want to delete the current record?" _
    , vbYesNo, "Delete Record?")
    If res = vbYes Then
    Adodc1.Recordset.Delete
    Adodc1.Recordset.UpdateBatch
    res = MsgBox("Your Informations Been Deleted Check Out Updates On Your Next Click", vbOKOnly, "CONFIRMATION MESSAGE !")
    If res = vbOKOnly Then
    End
    Adodc1.Recordset.MoveFirst
    End If
    Else
    Exit Sub
    End If
End Sub

Private Sub Command6_Click()
Form1.Show
Unload Me

End Sub

Private Sub Form_Activate()
Adodc1.Recordset.MoveFirst
Command3.Visible = False
Command4.Visible = False
Command2.Visible = True
Command1.Visible = True
Command5.Visible = True

End Sub

Private Sub Form_Load()
Command3.Visible = False
Command4.Value = False
Command2.Visible = True
Command5.Visible = True
Command1.Visible = True

End Sub

