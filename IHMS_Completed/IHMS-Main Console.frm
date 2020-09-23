VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Intelligent H.M.S - "
   ClientHeight    =   8985
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   13650
   Icon            =   "IHMS-Main Console.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   8730
      Width           =   13650
      _ExtentX        =   24077
      _ExtentY        =   450
      Style           =   1
      SimpleText      =   "Not Logged In"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList imgLstToolBar1 
      Left            =   600
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IHMS-Main Console.frx":6852
            Key             =   "loginout"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IHMS-Main Console.frx":D0B4
            Key             =   "newreg"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IHMS-Main Console.frx":13916
            Key             =   "openexisting"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IHMS-Main Console.frx":1A178
            Key             =   "admitdischarge"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IHMS-Main Console.frx":209DA
            Key             =   "diagnose"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IHMS-Main Console.frx":2318C
            Key             =   "about"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrMainToolbar 
      Align           =   1  'Align Top
      Height          =   1095
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   13650
      _ExtentX        =   24077
      _ExtentY        =   1931
      ButtonWidth     =   3016
      ButtonHeight    =   1879
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgLstToolBar1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Log In"
            Object.ToolTipText     =   "Log-In / Log-Out"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Registration (New)"
            Object.ToolTipText     =   "Register New Patient"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Search For Patient"
            Object.ToolTipText     =   "Open Existing Patient File"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Admit/Discharge"
            Object.ToolTipText     =   "Admit / Discharge this patient"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Diagnose"
            Object.ToolTipText     =   "Diagnose this patient"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "About IHMS"
            Object.ToolTipText     =   "About this application"
            ImageIndex      =   6
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuLogIn 
         Caption         =   "Log &In"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuNewPatient 
         Caption         =   "&New Patient Registration"
         Shortcut        =   ^N
         Visible         =   0   'False
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open Existing Patient File"
         Shortcut        =   {F3}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuLogOut 
         Caption         =   "Log &Out"
         Enabled         =   0   'False
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Enabled         =   0   'False
      Begin VB.Menu mnuToolsUsers 
         Caption         =   "&User Management Console"
      End
      Begin VB.Menu mnuToolsKnowledgeBase 
         Caption         =   "Update &Knowledge Base"
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuToolsOptions 
         Caption         =   "&Options"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_DblClick()
 If LoginSucceeded = True Then
    frmNewReg.Show
 Else
    frmLogin.Show 1
 End If
End Sub

Private Sub MDIForm_Load()
 Me.Caption = Me.Caption & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 If LoginSucceeded = False Then Exit Sub
 If MsgBox("Exiting IHMS will abort all tasks in progress. Data and files might be lost. " & vbCrLf & "Are you sure you want to exit?", vbCritical + vbYesNo) = vbYes Then
    Cancel = 0
    End
 Else
    Cancel = 1
 End If
End Sub

Private Sub mnuClose_Click()
 Unload frmOldPatient
End Sub

Private Sub mnuExit_Click()
 Unload Me
End Sub

Private Sub mnuHelpAbout_Click()
 frmAbout.Show 1
End Sub

Private Sub mnuLogIn_Click()
 'Check if database is available. if not, exit the app.
 If Dir(App.Path & "\IHMS_97.mdb") = "" Then
    MsgBox "Working database file not found." + vbCrLf + "Please make sure the access database file 'IHMS_97' available and try launching the application again." + vbCrLf + "The application will now exit...", , "ERROR"
    End
 End If
 frmLogin.Show 1
End Sub

Private Sub mnuLogOut_Click()
 If MsgBox("Logging out will abort all tasks in progress. Data and files might be lost. " & vbCrLf & "Are you sure you want to Log Out?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
 LoginSucceeded = False
 Call ConfigMenus
End Sub

Private Sub mnuNewPatient_Click()
 frmNewReg.Show
 mnuNewPatient.Enabled = False
 tbrMainToolbar.Buttons(1).Enabled = False
End Sub

Private Sub mnuOpen_Click()
 'f3:  Search for record by hospital number.
 patientNumberX = 0     'hosp number being sought for
 patientNumberX = Val(InputBox("Please enter the patient's HOSPITAL NUMBER:"))
 If patientNumberX = 0 Then Exit Sub  'User selects cancel
 Unload frmOldPatient
 frmWait.Show 1
 
End Sub

Private Sub mnuToolsKnowledgeBase_Click()
 MsgBox "This function is still under development.", vbInformation
End Sub

Private Sub mnuToolsOptions_Click()
 MsgBox "This function is still under development.", vbInformation
End Sub

Private Sub mnuToolsUsers_Click()
 frmUserMgt.Show 1
End Sub

Private Sub tbrMainToolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
 Select Case Button.Index
 Case 1
 'Log In/Log Out.
    If LoginSucceeded = False Then
    'The user is currently logged out
        'Check if database is available. if not, exit the app.
        If Dir(App.Path & "\IHMS_97.mdb") = "" Then
            MsgBox "Working database file not found." + vbCrLf + "Please make sure the access database file 'IHMS_97' available and try launching the application again." + vbCrLf + "The application will now exit...", , "ERROR"
            End
        End If
        frmLogin.Show 1
    Else
    'The user is currently logged in
        If MsgBox("Logging out will abort all tasks in progress. Data and files might be lost. " & vbCrLf & "Are you sure you want to Log Out?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        LoginSucceeded = False
        Call ConfigMenus
    End If

 Case 2
 'Register New Patient.
    frmNewReg.Show
    mnuNewPatient.Enabled = False
    tbrMainToolbar.Buttons(2).Enabled = False
 Case 3
 'Search DB for existing patient file.
    'f3:  Search for record by hospital number.
    patientNumberX = 0     'hosp number being sought for
    patientNumberX = Val(InputBox("Please enter the patient's HOSPITAL NUMBER:"))
    If patientNumberX = 0 Then Exit Sub  'User selects cancel
    Unload frmOldPatient
    frmWait.Show 1
 Case 4
 'Admit/Discharge this patient.
  'MsgBox "This function is still under development.", vbInformation
  If somePatient.AdmissionStatus = "IN" Then
        'Patient is currently admitted. therefore, the admit command is
        'not available, but the discharge command is.
         frmDischarge.Show 1
    Else
        'Patient is currently NOT admitted. therefore, the admit command is
        'available, but the discharge command is not.
        'This part is also executed if no existing hospital history record is found for the patient
        'i.e when somePatient.AdmissionStatus = ""
         frmAdmitExisting.Show 1
    End If
 Case 5
 'Diagnose this patient.
 frmDiagnosis.Show 1
 'MsgBox "This function is still under development.", vbInformation
 Case 6
 'About This App.
    frmAbout.Show 1
 End Select
End Sub

