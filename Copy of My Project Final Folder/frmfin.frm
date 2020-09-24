VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmfin 
   BackColor       =   &H80000008&
   Caption         =   "Employees Of Finance Department"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7680
   LinkTopic       =   "Form3"
   ScaleHeight     =   6495
   ScaleWidth      =   7680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   7335
      Begin VB.CommandButton Command5 
         Caption         =   "&Delete"
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
         Left            =   4320
         TabIndex        =   22
         Top             =   4680
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Edit"
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
         Left            =   3360
         TabIndex        =   21
         Top             =   4680
         Width           =   975
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Cancel"
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
         Left            =   4320
         TabIndex        =   20
         Top             =   4680
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Update"
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
         Left            =   3360
         TabIndex        =   19
         Top             =   4680
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Add"
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
         Left            =   2400
         TabIndex        =   18
         Top             =   4680
         Width           =   975
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   4920
         Picture         =   "frmfin.frx":0000
         ScaleHeight     =   1335
         ScaleWidth      =   1215
         TabIndex        =   17
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   1455
         Left            =   6000
         Picture         =   "frmfin.frx":2643
         ScaleHeight     =   1455
         ScaleWidth      =   1575
         TabIndex        =   16
         Top             =   0
         Width           =   1575
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   3840
         Picture         =   "frmfin.frx":4F1F
         ScaleHeight     =   1335
         ScaleWidth      =   1335
         TabIndex        =   14
         Top             =   0
         Width           =   1335
         Begin VB.PictureBox Picture4 
            Height          =   255
            Left            =   1080
            ScaleHeight     =   255
            ScaleWidth      =   15
            TabIndex        =   15
            Top             =   360
            Width           =   15
         End
      End
      Begin VB.CommandButton cmdmenu14 
         Caption         =   "&Back "
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
         Left            =   5280
         TabIndex        =   13
         Top             =   4680
         Width           =   975
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   1200
         Top             =   4680
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
         Connect         =   $"frmfin.frx":7879
         OLEDBString     =   $"frmfin.frx":796C
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "Finance"
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
         Caption         =   "Employee Informations"
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
         Height          =   3015
         Left            =   600
         TabIndex        =   2
         Top             =   1560
         Width           =   6255
         Begin VB.TextBox txtcontact 
            DataField       =   "scontact"
            DataSource      =   "Adodc1"
            Height          =   375
            Left            =   2520
            TabIndex        =   12
            Top             =   2280
            Width           =   3015
         End
         Begin VB.TextBox txtdob 
            DataField       =   "sdob"
            DataSource      =   "Adodc1"
            Height          =   375
            Left            =   2520
            TabIndex        =   11
            Top             =   1800
            Width           =   3015
         End
         Begin VB.TextBox txtaddress 
            DataField       =   "saddress"
            DataSource      =   "Adodc1"
            Height          =   375
            Left            =   2520
            TabIndex        =   10
            Top             =   1320
            Width           =   3015
         End
         Begin VB.TextBox txtdesign 
            DataField       =   "sdesignation"
            DataSource      =   "Adodc1"
            Height          =   375
            Left            =   2520
            TabIndex        =   9
            Top             =   840
            Width           =   3015
         End
         Begin VB.TextBox txtname 
            DataField       =   "sname"
            DataSource      =   "Adodc1"
            Height          =   375
            Left            =   2520
            TabIndex        =   8
            Top             =   360
            Width           =   3015
         End
         Begin VB.Label Label5 
            Caption         =   "  Contact           #"
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
            Height          =   375
            Left            =   360
            TabIndex        =   7
            Top             =   2280
            Width           =   1935
         End
         Begin VB.Label Label4 
            Caption         =   "  Date Of Birth   #"
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
            Height          =   375
            Left            =   360
            TabIndex        =   6
            Top             =   1800
            Width           =   1935
         End
         Begin VB.Label Label3 
            Caption         =   "  Address          #"
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
            Height          =   375
            Left            =   360
            TabIndex        =   5
            Top             =   1320
            Width           =   1935
         End
         Begin VB.Label Label2 
            Caption         =   "  Designation     #"
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
            Height          =   375
            Left            =   360
            TabIndex        =   4
            Top             =   840
            Width           =   1935
         End
         Begin VB.Label Label1 
            Caption         =   "  Name               #"
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
            Height          =   375
            Left            =   360
            TabIndex        =   3
            Top             =   360
            Width           =   1935
         End
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1455
         Left            =   0
         Picture         =   "frmfin.frx":7A5F
         ScaleHeight     =   1455
         ScaleWidth      =   3735
         TabIndex        =   1
         Top             =   0
         Width           =   3735
      End
   End
End
Attribute VB_Name = "frmfin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdmenu14_Click()
frmemp.Show
frmfin.Hide
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
    txtname = ""
    txtdesign = ""
    txtaddress = ""
    txtdob = ""
    txtcontact = ""
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
