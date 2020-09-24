VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmcnew 
   BackColor       =   &H80000008&
   Caption         =   "New Subscription"
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
      TabIndex        =   53
      Top             =   6000
      Width           =   1095
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
      Height          =   375
      Left            =   3600
      TabIndex        =   52
      Top             =   6000
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
      Left            =   2520
      TabIndex        =   51
      Top             =   6000
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
      Left            =   3600
      TabIndex        =   32
      Top             =   6000
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
      Left            =   2520
      TabIndex        =   31
      Top             =   6000
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
      Left            =   1440
      TabIndex        =   30
      Top             =   6000
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   240
      Top             =   6000
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
      Connect         =   $"frmcnew.frx":0000
      OLEDBString     =   $"frmcnew.frx":00F3
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "custonew"
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
   Begin VB.CommandButton cmdmenu4 
      Caption         =   "&Back Menu"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   29
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Frame Advance 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   7095
      Begin VB.Frame Frame1 
         BackColor       =   &H80000013&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   4215
         Left            =   120
         TabIndex        =   6
         Top             =   1560
         Width           =   6855
         Begin VB.Frame fraAccount 
            BackColor       =   &H80000013&
            Caption         =   "Apartment Information"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Index           =   1
            Left            =   120
            TabIndex        =   24
            Top             =   2880
            Width           =   6735
            Begin VB.TextBox Text18 
               DataField       =   "sapthand"
               DataSource      =   "Adodc1"
               Height          =   285
               Left            =   5520
               TabIndex        =   50
               Top             =   480
               Width           =   855
            End
            Begin VB.TextBox Text17 
               DataField       =   "sapttype"
               DataSource      =   "Adodc1"
               Height          =   285
               Left            =   3840
               TabIndex        =   49
               Top             =   480
               Width           =   1575
            End
            Begin VB.TextBox Text16 
               DataField       =   "saptadd"
               DataSource      =   "Adodc1"
               Height          =   285
               Left            =   1800
               TabIndex        =   48
               Top             =   480
               Width           =   1935
            End
            Begin VB.TextBox Text15 
               DataField       =   "saptname"
               DataSource      =   "Adodc1"
               Height          =   285
               Left            =   0
               TabIndex        =   47
               Top             =   480
               Width           =   1695
            End
            Begin VB.Label lblLate 
               Alignment       =   2  'Center
               BackColor       =   &H00C00000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Handover"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   5520
               TabIndex        =   28
               Top             =   240
               Width           =   855
            End
            Begin VB.Label lblInstall 
               Alignment       =   2  'Center
               BackColor       =   &H00C00000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Apartment Type"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   3840
               TabIndex        =   27
               Top             =   240
               Width           =   1575
            End
            Begin VB.Label lblAmountdue 
               Alignment       =   2  'Center
               BackColor       =   &H00C00000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " Apartment Addres"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   1800
               TabIndex        =   26
               Top             =   240
               Width           =   1935
            End
            Begin VB.Label lblPayduedate 
               Alignment       =   2  'Center
               BackColor       =   &H00C00000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Apartment Name"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   0
               TabIndex        =   25
               Top             =   240
               Width           =   1695
            End
         End
         Begin VB.Frame fraAccount 
            BackColor       =   &H80000013&
            Caption         =   "Account Information"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Index           =   0
            Left            =   120
            TabIndex        =   18
            Top             =   1920
            Width           =   6735
            Begin VB.TextBox Text14 
               DataField       =   "ssigndate"
               DataSource      =   "Adodc1"
               Height          =   285
               Left            =   5520
               TabIndex        =   46
               Top             =   480
               Width           =   855
            End
            Begin VB.TextBox Text13 
               DataField       =   "sinstallamnt"
               DataSource      =   "Adodc1"
               Height          =   285
               Left            =   3600
               TabIndex        =   45
               Top             =   480
               Width           =   1815
            End
            Begin VB.TextBox Text12 
               DataField       =   "sadue"
               DataSource      =   "Adodc1"
               Height          =   285
               Left            =   1800
               TabIndex        =   44
               Top             =   480
               Width           =   1695
            End
            Begin VB.TextBox Text11 
               DataField       =   "spduedate"
               DataSource      =   "Adodc1"
               Height          =   285
               Left            =   0
               TabIndex        =   43
               Top             =   480
               Width           =   1695
            End
            Begin VB.Label lblPayduedate 
               Alignment       =   2  'Center
               BackColor       =   &H00C00000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Payment due date"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   22
               Top             =   240
               Width           =   1695
            End
            Begin VB.Label lblAmountdue 
               Alignment       =   2  'Center
               BackColor       =   &H00C00000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Amount due"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   1800
               TabIndex        =   21
               Top             =   240
               Width           =   1695
            End
            Begin VB.Label lblInstall 
               Alignment       =   2  'Center
               BackColor       =   &H00C00000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Install Amount"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   3600
               TabIndex        =   20
               Top             =   240
               Width           =   1815
            End
            Begin VB.Label lblLate 
               Alignment       =   2  'Center
               BackColor       =   &H00C00000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Sign Date"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   5520
               TabIndex        =   19
               Top             =   240
               Width           =   915
            End
         End
         Begin VB.Frame fraId_Name 
            BackColor       =   &H80000013&
            Caption         =   "Customer Informations"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000006&
            Height          =   795
            Index           =   1
            Left            =   120
            TabIndex        =   14
            Top             =   0
            Width           =   6735
            Begin VB.TextBox Text4 
               DataField       =   "Ssex"
               DataSource      =   "Adodc1"
               Height          =   285
               Left            =   5040
               TabIndex        =   36
               Top             =   480
               Width           =   1215
            End
            Begin VB.TextBox Text3 
               DataField       =   "Sdob"
               DataSource      =   "Adodc1"
               Height          =   285
               Left            =   3480
               TabIndex        =   35
               Top             =   480
               Width           =   1455
            End
            Begin VB.TextBox Text2 
               DataField       =   "Slast"
               DataSource      =   "Adodc1"
               Height          =   285
               Left            =   1680
               TabIndex        =   34
               Top             =   480
               Width           =   1695
            End
            Begin VB.TextBox Text1 
               DataField       =   "Sfirst"
               DataSource      =   "Adodc1"
               Height          =   285
               Left            =   0
               TabIndex        =   33
               Top             =   480
               Width           =   1575
            End
            Begin VB.Label lblFirstName 
               BackColor       =   &H00C00000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Sex"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   2
               Left            =   5040
               TabIndex        =   23
               Top             =   240
               Width           =   1215
            End
            Begin VB.Label lblFirstName 
               BackColor       =   &H00C00000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Date Of Birth"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   3480
               TabIndex        =   17
               Top             =   240
               Width           =   1455
            End
            Begin VB.Label lblFirstName 
               BackColor       =   &H00C00000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Last  Name"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   1680
               TabIndex        =   16
               Top             =   195
               Width           =   1695
            End
            Begin VB.Label lblLastName 
               BackColor       =   &H00C00000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "First  Name"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   0
               TabIndex        =   15
               Top             =   195
               Width           =   1575
            End
         End
         Begin VB.Frame fraInfo 
            BackColor       =   &H80000013&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   120
            TabIndex        =   7
            Top             =   720
            Width           =   6735
            Begin VB.TextBox Text10 
               DataField       =   "scell"
               DataSource      =   "Adodc1"
               Height          =   285
               Left            =   5160
               TabIndex        =   42
               Top             =   840
               Width           =   1455
            End
            Begin VB.TextBox Text9 
               DataField       =   "sphone"
               DataSource      =   "Adodc1"
               Height          =   285
               Left            =   3240
               TabIndex        =   41
               Top             =   840
               Width           =   1215
            End
            Begin VB.TextBox Text8 
               DataField       =   "scountry"
               DataSource      =   "Adodc1"
               Height          =   285
               Left            =   960
               TabIndex        =   40
               Top             =   840
               Width           =   1335
            End
            Begin VB.TextBox Text7 
               DataField       =   "szip"
               DataSource      =   "Adodc1"
               Height          =   285
               Left            =   3600
               TabIndex        =   39
               Top             =   480
               Width           =   1575
            End
            Begin VB.TextBox Text6 
               DataField       =   "scity"
               DataSource      =   "Adodc1"
               Height          =   285
               Left            =   720
               TabIndex        =   38
               Top             =   480
               Width           =   2295
            End
            Begin VB.TextBox Text5 
               DataField       =   "sstreet"
               DataSource      =   "Adodc1"
               Height          =   285
               Left            =   720
               TabIndex        =   37
               Top             =   120
               Width           =   4695
            End
            Begin VB.Label lblStreet 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C00000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Street :"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   0
               TabIndex        =   13
               Top             =   120
               Width           =   690
            End
            Begin VB.Label lblDOB 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C00000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Country :"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   0
               TabIndex        =   12
               Top             =   840
               Width           =   930
            End
            Begin VB.Label lblPhone 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C00000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Phone # :"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   2400
               TabIndex        =   11
               Top             =   840
               Width           =   825
            End
            Begin VB.Label lblSS 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C00000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Cell # :"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   4560
               TabIndex        =   10
               Top             =   840
               Width           =   615
            End
            Begin VB.Label lblZip 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C00000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Zip : "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   3120
               TabIndex        =   9
               Top             =   480
               Width           =   405
            End
            Begin VB.Label lblCity 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C00000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "City :"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   0
               TabIndex        =   8
               Top             =   480
               Width           =   690
            End
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   1575
         Left            =   0
         Picture         =   "frmcnew.frx":01E6
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
         Picture         =   "frmcnew.frx":64B6
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
         Picture         =   "frmcnew.frx":8D92
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
         Picture         =   "frmcnew.frx":B6EC
         ScaleHeight     =   1335
         ScaleWidth      =   1095
         TabIndex        =   1
         Top             =   0
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmcnew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdmenu4_Click()
frmcustom.Show
frmcnew.Hide

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
    Text13 = ""
    Text14 = ""
    Text15 = ""
    Text16 = ""
    Text17 = ""
    Text18 = ""
    
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
Form6.Show
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





