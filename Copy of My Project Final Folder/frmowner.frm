VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmowner 
   BackColor       =   &H80000007&
   Caption         =   "Land Owners List"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7680
   LinkTopic       =   "Form3"
   ScaleHeight     =   6495
   ScaleWidth      =   7680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      Caption         =   "&Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   39
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton cmdmenu21 
      Caption         =   "&Back"
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
      Left            =   5760
      TabIndex        =   21
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      Begin VB.CommandButton Command5 
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         TabIndex        =   26
         Top             =   5640
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   25
         Top             =   5640
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         TabIndex        =   24
         Top             =   5640
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Update"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   23
         Top             =   5640
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   22
         Top             =   5640
         Width           =   1095
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   120
         Top             =   5640
         Width           =   1215
         _ExtentX        =   2143
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
         Connect         =   $"frmowner.frx":0000
         OLEDBString     =   $"frmowner.frx":00F3
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "Ownerslist"
         Caption         =   "Adodc1"
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
         Caption         =   "Owners Information"
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
         Height          =   4095
         Left            =   120
         TabIndex        =   6
         Top             =   1680
         Width           =   6735
         Begin VB.Frame Frame4 
            BackColor       =   &H80000013&
            Caption         =   "ADT Employee"
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
            Height          =   1095
            Left            =   120
            TabIndex        =   8
            Top             =   2640
            Width           =   6495
            Begin VB.TextBox Text12 
               DataField       =   "Sempdepart"
               DataSource      =   "Adodc1"
               Height          =   285
               Left            =   4320
               TabIndex        =   38
               Top             =   600
               Width           =   1695
            End
            Begin VB.TextBox Text11 
               DataField       =   "Sempdesign"
               DataSource      =   "Adodc1"
               Height          =   285
               Left            =   2400
               TabIndex        =   37
               Top             =   600
               Width           =   1695
            End
            Begin VB.TextBox Text10 
               DataField       =   "Sempname"
               DataSource      =   "Adodc1"
               Height          =   285
               Left            =   120
               TabIndex        =   36
               Top             =   600
               Width           =   1935
            End
            Begin VB.Label Label1 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "     Department"
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
               Index           =   11
               Left            =   4320
               TabIndex        =   20
               Top             =   360
               Width           =   1695
            End
            Begin VB.Label Label1 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "    Designation"
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
               Index           =   10
               Left            =   2400
               TabIndex        =   19
               Top             =   360
               Width           =   1695
            End
            Begin VB.Label Label1 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " Employee  Name"
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
               Index           =   9
               Left            =   120
               TabIndex        =   18
               Top             =   360
               Width           =   1935
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H80000013&
            Caption         =   "Owners Info"
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
            Height          =   2175
            Left            =   120
            TabIndex        =   7
            Top             =   360
            Width           =   6495
            Begin VB.TextBox Text9 
               DataField       =   "SPdate"
               DataSource      =   "Adodc1"
               Height          =   285
               Left            =   4200
               TabIndex        =   35
               Top             =   1680
               Width           =   2055
            End
            Begin VB.TextBox Text8 
               DataField       =   "Sland"
               DataSource      =   "Adodc1"
               Height          =   285
               Left            =   120
               TabIndex        =   34
               Top             =   1680
               Width           =   3615
            End
            Begin VB.TextBox Text7 
               DataField       =   "Scell"
               DataSource      =   "Adodc1"
               Height          =   285
               Left            =   4800
               TabIndex        =   33
               Top             =   1080
               Width           =   1455
            End
            Begin VB.TextBox Text6 
               DataField       =   "Sphone"
               DataSource      =   "Adodc1"
               Height          =   285
               Left            =   3360
               TabIndex        =   32
               Top             =   1080
               Width           =   1215
            End
            Begin VB.TextBox Text5 
               DataField       =   "SAddress"
               DataSource      =   "Adodc1"
               Height          =   285
               Left            =   120
               TabIndex        =   31
               Top             =   1080
               Width           =   3015
            End
            Begin VB.TextBox Text4 
               DataField       =   "SDob"
               DataSource      =   "Adodc1"
               Height          =   285
               Left            =   4800
               TabIndex        =   30
               Top             =   480
               Width           =   1455
            End
            Begin VB.TextBox Text3 
               DataField       =   "SGender"
               DataSource      =   "Adodc1"
               Height          =   285
               Left            =   3360
               TabIndex        =   29
               Top             =   480
               Width           =   1215
            End
            Begin VB.TextBox Text2 
               DataField       =   "SlastName"
               DataSource      =   "Adodc1"
               Height          =   285
               Left            =   1680
               TabIndex        =   28
               Top             =   480
               Width           =   1455
            End
            Begin VB.TextBox Text1 
               DataField       =   "SfirstName"
               DataSource      =   "Adodc1"
               Height          =   285
               Left            =   120
               TabIndex        =   27
               Top             =   480
               Width           =   1455
            End
            Begin VB.Label Label1 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "   Purchased Date"
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
               Index           =   8
               Left            =   4200
               TabIndex        =   17
               Top             =   1440
               Width           =   2055
            End
            Begin VB.Label Label1 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "            Land Description"
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
               Index           =   7
               Left            =   120
               TabIndex        =   16
               Top             =   1440
               Width           =   3615
            End
            Begin VB.Label Label1 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "    Cell  #"
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
               Index           =   6
               Left            =   4800
               TabIndex        =   15
               Top             =   840
               Width           =   1455
            End
            Begin VB.Label Label1 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "   Phone  #"
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
               Index           =   5
               Left            =   3360
               TabIndex        =   14
               Top             =   840
               Width           =   1215
            End
            Begin VB.Label Label1 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "                 Address"
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
               Index           =   4
               Left            =   120
               TabIndex        =   13
               Top             =   840
               Width           =   3015
            End
            Begin VB.Label Label1 
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " Date Of Birth"
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
               Left            =   4800
               TabIndex        =   12
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
               Left            =   3360
               TabIndex        =   11
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
               TabIndex        =   10
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
               TabIndex        =   9
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
         Picture         =   "frmowner.frx":01E6
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
         Picture         =   "frmowner.frx":2AC2
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
         Picture         =   "frmowner.frx":5105
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
         Picture         =   "frmowner.frx":7A5F
         ScaleHeight     =   1575
         ScaleWidth      =   3735
         TabIndex        =   1
         Top             =   0
         Width           =   3735
      End
   End
End
Attribute VB_Name = "frmowner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdmenu21_Click()
frmland.Show
frmowner.Hide
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
    Text6 = ""
    Text7 = ""
    Text8 = ""
    Text9 = ""
    Text10 = ""
    Text11 = ""
    Text12 = ""
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
Form2.Show
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


