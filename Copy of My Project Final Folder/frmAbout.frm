VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H80000007&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Advance Developement Technologies Ltd."
   ClientHeight    =   6555
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   7710
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4524.378
   ScaleMode       =   0  'User
   ScaleWidth      =   7240.09
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture6 
      Height          =   375
      Index           =   6
      Left            =   5760
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   315
      ScaleWidth      =   1515
      TabIndex        =   18
      Top             =   4440
      Width           =   1575
   End
   Begin VB.PictureBox Picture6 
      Height          =   375
      Index           =   5
      Left            =   5760
      Picture         =   "frmAbout.frx":1936
      ScaleHeight     =   315
      ScaleWidth      =   1515
      TabIndex        =   17
      Top             =   3960
      Width           =   1575
   End
   Begin VB.PictureBox Picture6 
      Height          =   375
      Index           =   4
      Left            =   5760
      Picture         =   "frmAbout.frx":31D7
      ScaleHeight     =   315
      ScaleWidth      =   1515
      TabIndex        =   16
      Top             =   3480
      Width           =   1575
   End
   Begin VB.PictureBox Picture6 
      Height          =   375
      Index           =   3
      Left            =   5760
      Picture         =   "frmAbout.frx":4B22
      ScaleHeight     =   315
      ScaleWidth      =   1515
      TabIndex        =   15
      Top             =   3000
      Width           =   1575
   End
   Begin VB.PictureBox Picture6 
      Height          =   375
      Index           =   2
      Left            =   5760
      Picture         =   "frmAbout.frx":6408
      ScaleHeight     =   315
      ScaleWidth      =   1515
      TabIndex        =   14
      Top             =   2520
      Width           =   1575
   End
   Begin VB.PictureBox Picture6 
      Height          =   375
      Index           =   1
      Left            =   5760
      Picture         =   "frmAbout.frx":7CF4
      ScaleHeight     =   315
      ScaleWidth      =   1515
      TabIndex        =   13
      Top             =   2040
      Width           =   1575
   End
   Begin VB.PictureBox Picture6 
      Height          =   375
      Index           =   0
      Left            =   5760
      Picture         =   "frmAbout.frx":9563
      ScaleHeight     =   315
      ScaleWidth      =   1515
      TabIndex        =   12
      Top             =   1560
      Width           =   1575
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   1335
      Index           =   0
      Left            =   360
      Picture         =   "frmAbout.frx":AE10
      ScaleHeight     =   1335
      ScaleWidth      =   1095
      TabIndex        =   11
      Top             =   3000
      Width           =   1095
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   360
      Picture         =   "frmAbout.frx":D6EC
      ScaleHeight     =   1215
      ScaleWidth      =   1095
      TabIndex        =   10
      Top             =   1800
      Width           =   1095
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   1215
      Index           =   1
      Left            =   360
      Picture         =   "frmAbout.frx":FD2F
      ScaleHeight     =   1215
      ScaleWidth      =   1095
      TabIndex        =   8
      Top             =   4320
      Width           =   1095
      Begin VB.PictureBox Picture1 
         Height          =   255
         Left            =   1080
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   9
         Top             =   360
         Width           =   15
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   1215
      Index           =   0
      Left            =   360
      Picture         =   "frmAbout.frx":137E5
      ScaleHeight     =   1215
      ScaleWidth      =   1095
      TabIndex        =   6
      Top             =   600
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
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   1800
      Picture         =   "frmAbout.frx":1613F
      ScaleHeight     =   1085.105
      ScaleMode       =   0  'User
      ScaleWidth      =   2654.82
      TabIndex        =   1
      Top             =   600
      Width           =   3810
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2040
      TabIndex        =   0
      Top             =   5160
      Width           =   1185
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&System Info..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3360
      TabIndex        =   2
      Top             =   5160
      Width           =   1605
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   4
      X1              =   1464.921
      X2              =   5296.253
      Y1              =   2070.654
      Y2              =   2070.654
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   3
      X1              =   1464.921
      X2              =   5296.253
      Y1              =   1573.697
      Y2              =   1573.697
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   2
      X1              =   0
      X2              =   5210.799
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   1464.921
      X2              =   5183.566
      Y1              =   3395.872
      Y2              =   3395.872
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H80000003&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Advance Developement Technologies Ltd."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   1680
      TabIndex        =   4
      Top             =   2400
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   1464.921
      X2              =   5183.566
      Y1              =   3313.046
      Y2              =   3313.046
   End
   Begin VB.Label lblVersion 
      BackColor       =   &H80000003&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Version 7.07"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   225
      Left            =   1680
      TabIndex        =   5
      Top             =   3120
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      BackColor       =   &H80000003&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "This application created by Samiha Esha. For any comments or suggestions send me an email at amazon707@gmail.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   825
      Left            =   1680
      TabIndex        =   3
      Top             =   3720
      Width           =   3870
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long


Private Sub cmdSysInfo_Click()
  Call StartSysInfo
End Sub

Private Sub cmdOK_Click()
  Unload Me
  frmadt.Show
End Sub

Private Sub Form_Load()
    Me.Caption = "Advance Developement Technologies Ltd. " & App.Title
    
    
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
    
   
    
          
End Sub

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' Try To Get System Info Program Path\Name From Registry...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' Try To Get System Info Program Path Only From Registry...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' Validate Existance Of Known 32 Bit File Version
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' Error - File Can Not Be Found...
        Else
            GoTo SysInfoErr
        End If
    ' Error - Registry Entry Can Not Be Found...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "System Information Is Unavailable At This Time", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Search Data Types...
    Case REG_SZ                                             ' String Registry Key Data Type
        KeyVal = tmpVal                                     ' Copy String Value
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
    End Select
    
    GetKeyValue = True                                      ' Return Success
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:      ' Cleanup After An Error Has Occured...
    KeyVal = ""                                             ' Set Return Val To Empty String
    GetKeyValue = False                                     ' Return Failure
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function


