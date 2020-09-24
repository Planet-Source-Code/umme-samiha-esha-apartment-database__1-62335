VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmcustomer 
   BackColor       =   &H80000008&
   Caption         =   "Employees of Customer Service Department"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7680
   LinkTopic       =   "Form3"
   ScaleHeight     =   6495
   ScaleWidth      =   7680
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   120
      Picture         =   "frmcustomer.frx":0000
      ScaleHeight     =   1455
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
      Width           =   6855
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
         Left            =   3840
         TabIndex        =   16
         Top             =   4800
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
         Left            =   2880
         TabIndex        =   15
         Top             =   4800
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
         Left            =   1920
         TabIndex        =   14
         Top             =   4800
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
         Left            =   3840
         TabIndex        =   13
         Top             =   4800
         Visible         =   0   'False
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
         Left            =   2880
         TabIndex        =   12
         Top             =   4800
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   5160
         Picture         =   "frmcustomer.frx":62D0
         ScaleHeight     =   1335
         ScaleWidth      =   1215
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   4080
         Picture         =   "frmcustomer.frx":8913
         ScaleHeight     =   1335
         ScaleWidth      =   1215
         TabIndex        =   9
         Top             =   0
         Width           =   1215
         Begin VB.PictureBox Picture4 
            Height          =   255
            Left            =   1080
            ScaleHeight     =   255
            ScaleWidth      =   15
            TabIndex        =   10
            Top             =   360
            Width           =   15
         End
      End
      Begin VB.CommandButton cmdmenu10 
         Caption         =   "&Back Menu"
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
         Left            =   4800
         TabIndex        =   8
         Top             =   4800
         Width           =   1095
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   720
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
         Connect         =   $"frmcustomer.frx":B26D
         OLEDBString     =   $"frmcustomer.frx":B360
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "CustomerCare"
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
         Left            =   480
         TabIndex        =   2
         Top             =   1680
         Width           =   5895
         Begin VB.TextBox Text5 
            DataField       =   "scontact"
            DataSource      =   "Adodc1"
            Height          =   375
            Left            =   2040
            TabIndex        =   21
            Top             =   2280
            Width           =   3375
         End
         Begin VB.TextBox Text4 
            DataField       =   "sdob"
            DataSource      =   "Adodc1"
            Height          =   375
            Left            =   2040
            TabIndex        =   20
            Top             =   1800
            Width           =   3375
         End
         Begin VB.TextBox Text3 
            DataField       =   "saddress"
            DataSource      =   "Adodc1"
            Height          =   375
            Left            =   2040
            TabIndex        =   19
            Top             =   1320
            Width           =   3375
         End
         Begin VB.TextBox Text2 
            DataField       =   "sdesignation"
            DataSource      =   "Adodc1"
            Height          =   375
            Left            =   2040
            TabIndex        =   18
            Top             =   840
            Width           =   3375
         End
         Begin VB.TextBox Text1 
            DataField       =   "sname"
            DataSource      =   "Adodc1"
            Height          =   375
            Left            =   2040
            TabIndex        =   17
            Top             =   360
            Width           =   3375
         End
         Begin VB.Label Label1 
            Caption         =   "Contacts       #"
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
            Left            =   240
            TabIndex        =   7
            Top             =   2280
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Date Of Birth #"
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
            Left            =   240
            TabIndex        =   6
            Top             =   1800
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Address         #"
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
            Left            =   240
            TabIndex        =   5
            Top             =   1320
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Designation    #"
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
            Left            =   240
            TabIndex        =   4
            Top             =   840
            Width           =   1575
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
            ForeColor       =   &H80000002&
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   3
            Top             =   360
            Width           =   1575
         End
      End
   End
End
Attribute VB_Name = "frmcustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdmenu10_Click()
frmemp.Show
frmcustomer.Hide
frmmark.Hide
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

Adodc1.Recordset.CancelBatch


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
