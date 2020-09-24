VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmcomplete 
   BackColor       =   &H80000008&
   Caption         =   "Completed Projects"
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
      Height          =   6015
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7095
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1455
         Left            =   0
         Picture         =   "frmcomplete.frx":0000
         ScaleHeight     =   1455
         ScaleWidth      =   3735
         TabIndex        =   15
         Top             =   0
         Width           =   3735
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000013&
         Caption         =   "Ongoing Projects"
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
         Height          =   4335
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   6735
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
            Height          =   495
            Left            =   5040
            TabIndex        =   33
            Top             =   2520
            Width           =   1455
         End
         Begin VB.TextBox Text10 
            DataField       =   "saptoadd"
            DataSource      =   "adodc1"
            Height          =   285
            Left            =   1920
            TabIndex        =   32
            Top             =   3600
            Width           =   2895
         End
         Begin VB.TextBox Text9 
            DataField       =   "saptowner"
            DataSource      =   "adodc1"
            Height          =   285
            Left            =   1920
            TabIndex        =   31
            Top             =   3240
            Width           =   2895
         End
         Begin VB.TextBox Text8 
            DataField       =   "saptsinfo"
            DataSource      =   "adodc1"
            Height          =   285
            Left            =   1920
            TabIndex        =   30
            Top             =   2880
            Width           =   2895
         End
         Begin VB.TextBox Text7 
            DataField       =   "sapthandate"
            DataSource      =   "adodc1"
            Height          =   285
            Left            =   1920
            TabIndex        =   29
            Top             =   2520
            Width           =   2895
         End
         Begin VB.TextBox Text6 
            DataField       =   "saptaddress"
            DataSource      =   "adodc1"
            Height          =   285
            Left            =   1920
            TabIndex        =   28
            Top             =   2160
            Width           =   2895
         End
         Begin VB.TextBox Text5 
            DataField       =   "Saptcost"
            DataSource      =   "adodc1"
            Height          =   285
            Left            =   1920
            TabIndex        =   27
            Top             =   1800
            Width           =   2895
         End
         Begin VB.TextBox Text4 
            DataField       =   "Sapttype"
            DataSource      =   "adodc1"
            Height          =   285
            Left            =   1920
            TabIndex        =   26
            Top             =   1440
            Width           =   2895
         End
         Begin VB.TextBox Text3 
            DataField       =   "Saptsize"
            DataSource      =   "adodc1"
            Height          =   285
            Left            =   1920
            TabIndex        =   25
            Top             =   1080
            Width           =   2895
         End
         Begin VB.TextBox Text2 
            DataField       =   "Saptname"
            DataSource      =   "adodc1"
            Height          =   285
            Left            =   1920
            TabIndex        =   24
            Top             =   720
            Width           =   2895
         End
         Begin VB.TextBox Text1 
            DataField       =   "Sarea"
            DataSource      =   "adodc1"
            Height          =   285
            Left            =   1920
            TabIndex        =   23
            Top             =   360
            Width           =   2895
         End
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
            Height          =   495
            Left            =   5040
            TabIndex        =   22
            Top             =   1920
            Width           =   1455
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
            Height          =   495
            Left            =   5040
            TabIndex        =   21
            Top             =   1320
            Width           =   1455
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
            Height          =   495
            Left            =   5040
            TabIndex        =   20
            Top             =   1920
            Width           =   1455
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
            Height          =   495
            Left            =   5040
            TabIndex        =   19
            Top             =   1320
            Width           =   1455
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
            Height          =   495
            Left            =   5040
            TabIndex        =   18
            Top             =   720
            Width           =   1455
         End
         Begin VB.CommandButton cmdmenu6 
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
            Height          =   495
            Index           =   0
            Left            =   5040
            TabIndex        =   6
            Top             =   3120
            Width           =   1455
         End
         Begin MSAdodcLib.Adodc adodc1 
            Height          =   375
            Left            =   240
            Top             =   3960
            Width           =   4695
            _ExtentX        =   8281
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
            Connect         =   $"frmcomplete.frx":62D0
            OLEDBString     =   $"frmcomplete.frx":63C3
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   "completedprojects"
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
         Begin VB.Label Label1 
            Caption         =   "  Owners Address"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   17
            Top             =   3600
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "  Owners Name"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   16
            Top             =   3240
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   " Types"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   255
            Index           =   9
            Left            =   240
            TabIndex        =   14
            Top             =   1440
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "  Cost"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   255
            Index           =   8
            Left            =   240
            TabIndex        =   13
            Top             =   1800
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "  Address"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   255
            Index           =   7
            Left            =   240
            TabIndex        =   12
            Top             =   2160
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "  Handover Date"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   255
            Index           =   6
            Left            =   240
            TabIndex        =   11
            Top             =   2520
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "  Sell Information"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   10
            Top             =   2880
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   " Size"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   9
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   " Apartment Name"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   8
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   " Area"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   7
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   3840
         Picture         =   "frmcomplete.frx":64B6
         ScaleHeight     =   1335
         ScaleWidth      =   1095
         TabIndex        =   3
         Top             =   0
         Width           =   1095
         Begin VB.PictureBox Picture4 
            Height          =   255
            Left            =   1080
            ScaleHeight     =   255
            ScaleWidth      =   15
            TabIndex        =   4
            Top             =   360
            Width           =   15
         End
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   4920
         Picture         =   "frmcomplete.frx":8E10
         ScaleHeight     =   1335
         ScaleWidth      =   1095
         TabIndex        =   2
         Top             =   0
         Width           =   1095
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   6000
         Picture         =   "frmcomplete.frx":B453
         ScaleHeight     =   1335
         ScaleWidth      =   1215
         TabIndex        =   1
         Top             =   0
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmcomplete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdmenu6_Click(Index As Integer)
frmproject.Show
frmcomplete.Hide
End Sub

Private Sub Command6_Click()
Form5.Show
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
Command4.Visible = False
Command2.Visible = True
Command1.Visible = True
Command5.Visible = True

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

