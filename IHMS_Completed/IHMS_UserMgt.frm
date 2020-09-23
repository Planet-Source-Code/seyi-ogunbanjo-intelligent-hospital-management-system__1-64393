VERSION 5.00
Begin VB.Form frmUserMgt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Users Management"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   4920
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   4320
      Width           =   3855
   End
   Begin VB.Data datUsers 
      Caption         =   "Users"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4680
      Visible         =   0   'False
      Width           =   2940
   End
   Begin VB.Frame Frame1 
      Caption         =   "List of Users"
      Height          =   3975
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   3855
      Begin VB.CommandButton cmdChangePassword 
         Caption         =   "Change Password for the selected User"
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   3240
         Width           =   3375
      End
      Begin VB.CommandButton cmdAddUser 
         Caption         =   "Add New User"
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   2640
         Width           =   1575
      End
      Begin VB.CommandButton cmdDeleteUser 
         Caption         =   "Delete This User"
         Height          =   495
         Left            =   2040
         TabIndex        =   2
         Top             =   2640
         Width           =   1575
      End
      Begin VB.ListBox lstUsers 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2010
         Left            =   240
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   3375
      End
   End
End
Attribute VB_Name = "frmUserMgt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAddUser_Click()
 Dim userName As String
 Dim passPhrase As String
 
 userName = InputBox("Enter a user name:")
 passPhrase = InputBox("Enter a password for the user:")
 If Trim(userName = "") Or Trim(passPhrase = "") Then
    MsgBox "User name AND password have to be of non-zero lenght. Try again", , "Error"
    Exit Sub
 End If
 lstUsers.AddItem UCase(userName)
 With datUsers.Recordset
    .AddNew
    .Fields("user") = userName
    .Fields("pass") = passPhrase
    .Update
 End With
 MsgBox "New User successfully added."
End Sub

Private Sub cmdChangePassword_Click()
 On Error GoTo errhnd
 Dim oldPass As String
 Dim newPass As String
 Dim confirmPass As String
 
 With datUsers.Recordset
    .MoveFirst
    Do While UCase(.Fields("user")) <> lstUsers.List(lstUsers.ListIndex)
        .MoveNext
    Loop
    oldPass = InputBox("Enter the old password:")
    newPass = InputBox("Enter the new password:")
    confirmPass = InputBox("Confirm new password:")
    If UCase(oldPass) = UCase(.Fields("pass")) Then
        If newPass = confirmPass Then
            .Edit
            .Fields("pass") = newPass
            .Update
            MsgBox "Password change successful!"
        Else
            MsgBox "Password change failed! Please retry", vbInformation
        End If
    Else
        MsgBox "Old password incorrect. Please retry"
    End If
 End With
 Exit Sub
errhnd:
 Select Case Err.Number
 Case 3021
    MsgBox "You need to select a user to update"
 Case Else
    MsgBox "Critical error"
 End Select

End Sub

Private Sub cmdClose_Click()
 Unload Me
End Sub

Private Sub cmdDeleteUser_Click()
 On Error GoTo errhnd
 If MsgBox("Sure you want to delete this user?", vbYesNo, "Confirm") = vbNo Then Exit Sub
 
 With datUsers.Recordset
    .MoveFirst
    Do While UCase(.Fields("user")) <> lstUsers.List(lstUsers.ListIndex)
        .MoveNext
    Loop
    If UCase(lstUsers.List(lstUsers.ListIndex)) = "DOCTOR" Then
        MsgBox "The DOCTOR account is an administrative account and CANNOT be deleted!"
        Exit Sub
    End If
    .Delete
    lstUsers.RemoveItem lstUsers.ListIndex
 End With
Exit Sub

errhnd:
 Select Case Err.Number
 Case 3021
    MsgBox "You need to select a user to delete"
 Case Else
    MsgBox "Critical error"
 End Select
End Sub

Private Sub Form_Load()

'On Error Resume Next
 datUsers.DatabaseName = App.Path & "\IHMS_97.mdb"
 datUsers.RecordSource = "IHMS_Users"
 datUsers.Refresh
 
 'populate lstUsers
 With datUsers.Recordset
    .MoveFirst
    Do While Not .EOF
        lstUsers.AddItem (UCase(.Fields("user")))
        .MoveNext
    Loop
 End With
 
End Sub

