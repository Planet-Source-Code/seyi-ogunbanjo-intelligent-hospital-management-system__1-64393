VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1920
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
   Icon            =   "IHMS_Login.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1134.399
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboUserName 
      Height          =   315
      ItemData        =   "IHMS_Login.frx":6852
      Left            =   1320
      List            =   "IHMS_Login.frx":685F
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   360
      Width           =   2295
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   495
      TabIndex        =   3
      Top             =   1260
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2100
      TabIndex        =   4
      Top             =   1260
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   2
      Text            =   "351"
      Top             =   765
      Width           =   2325
   End
   Begin VB.Data datUsers 
      Caption         =   "IHMS_Users"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   3540
   End
   Begin VB.Label lblLabels 
      Caption         =   "&User Name:"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   390
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   1
      Top             =   780
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
 'check for correct password
 LoginSucceeded = False
 With datUsers.Recordset
    .MoveFirst
    Do While Not .EOF And Not LoginSucceeded
        If (cboUserName = .Fields("user")) And Trim(txtPassword) = .Fields("pass") Then
            'Login authenticated
            MsgBox "Login successful. Welcome to the system.", vbInformation
            LoginSucceeded = True
            Call ConfigMenus
            Unload Me
        End If
        .MoveNext
    Loop
    If LoginSucceeded = False Then
        MsgBox "Login failed, try again!", , "Login"
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
    End If
 End With
End Sub

Private Sub Form_Load()
 datUsers.DatabaseName = App.Path & "\IHMS_97.mdb"
 datUsers.RecordSource = "IHMS_Users"
End Sub

