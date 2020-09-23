VERSION 5.00
Begin VB.Form frmNewReg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registration - New Patient"
   ClientHeight    =   6045
   ClientLeft      =   1515
   ClientTop       =   2895
   ClientWidth     =   12240
   ControlBox      =   0   'False
   Icon            =   "IHMS_NewPatient.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   12240
   Begin VB.Data datLabInfo 
      Caption         =   "Lab Info"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5640
      Visible         =   0   'False
      Width           =   2580
   End
   Begin VB.Data datPerInfo 
      Caption         =   "Personal Info"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5640
      Visible         =   0   'False
      Width           =   2340
   End
   Begin VB.CommandButton cmdCancelReg 
      Caption         =   "&Cancel Registration"
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
      Left            =   9000
      TabIndex        =   41
      Top             =   5160
      Width           =   2895
   End
   Begin VB.CommandButton cmdRestartReg 
      Caption         =   "Res&tart Registration"
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
      Left            =   9000
      TabIndex        =   40
      Top             =   4680
      Width           =   2895
   End
   Begin VB.CommandButton cmdRegAndAdmit 
      Caption         =   "Register and &Admit Patient"
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
      Left            =   9000
      TabIndex        =   39
      Top             =   4200
      Width           =   2895
   End
   Begin VB.CommandButton cmdRegOnly 
      Caption         =   "&Register Patient"
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
      Left            =   9000
      TabIndex        =   38
      Top             =   3720
      Width           =   2895
   End
   Begin VB.Frame fraLabInfo 
      Caption         =   "&Laboratory Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   5280
      TabIndex        =   29
      Top             =   3600
      Width           =   3615
      Begin VB.ComboBox cboRHFactor 
         Height          =   315
         ItemData        =   "IHMS_NewPatient.frx":6852
         Left            =   1680
         List            =   "IHMS_NewPatient.frx":685C
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   1200
         Width           =   1695
      End
      Begin VB.ComboBox cboBloodGrp 
         Height          =   315
         ItemData        =   "IHMS_NewPatient.frx":6874
         Left            =   1680
         List            =   "IHMS_NewPatient.frx":6881
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox txtAllergy 
         Height          =   285
         Left            =   1680
         TabIndex        =   37
         Text            =   "n/a"
         ToolTipText     =   "DD/MM/YYYY"
         Top             =   1560
         Width           =   1695
      End
      Begin VB.TextBox txtLabRefNo 
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label lblLabAllergy 
         Alignment       =   1  'Right Justify
         Caption         =   "Allergy:"
         Height          =   255
         Left            =   480
         TabIndex        =   36
         ToolTipText     =   "DD/MM/YYYY"
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label lblLabBloodGroup 
         Alignment       =   1  'Right Justify
         Caption         =   "Blood Group:"
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblLabRefNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Lab. Ref. Number:"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblLabGenotype 
         Alignment       =   1  'Right Justify
         Caption         =   "Rh Factor:"
         Height          =   255
         Left            =   480
         TabIndex        =   34
         Top             =   1200
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "&Sponsor's Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   240
      TabIndex        =   24
      Top             =   3600
      Width           =   4935
      Begin VB.TextBox txtNameOfSponsor 
         Height          =   285
         Left            =   1680
         TabIndex        =   26
         Top             =   360
         Width           =   3015
      End
      Begin VB.TextBox txtAddOfSponsor 
         Height          =   885
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         Text            =   "IHMS_NewPatient.frx":688E
         Top             =   960
         Width           =   4575
      End
      Begin VB.Label Label6 
         Caption         =   "Address of Sponsor:"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Name of Sponsor:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame fraNextOfKin 
      Caption         =   "&Next of Kin's Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   6960
      TabIndex        =   17
      Top             =   1080
      Width           =   4935
      Begin VB.TextBox txtKinAddress 
         Height          =   885
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   23
         Text            =   "IHMS_NewPatient.frx":689C
         Top             =   1320
         Width           =   4575
      End
      Begin VB.TextBox txtKinName 
         Height          =   285
         Left            =   1680
         TabIndex        =   19
         Text            =   "mr ogunbe"
         Top             =   360
         Width           =   3015
      End
      Begin VB.TextBox txtRelationship 
         Height          =   285
         Left            =   2160
         TabIndex        =   21
         Text            =   "father"
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label lblKinSName 
         Caption         =   "Name of Next of Kin:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label11 
         Caption         =   "Address of Next of Kin:"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label13 
         Caption         =   "Relationship to Next of Kin:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   1935
      End
   End
   Begin VB.TextBox txtHospNo 
      BackColor       =   &H80000016&
      DataSource      =   "datPerInfo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   720
      Width           =   1695
   End
   Begin VB.Frame fraPersonal 
      Caption         =   "&Personal Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   6495
      Begin VB.TextBox txtDoB 
         Height          =   285
         Left            =   1320
         TabIndex        =   8
         Text            =   "30-02-1983"
         ToolTipText     =   "DD/MM/YYYY"
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txtStateOfOrigin 
         Height          =   285
         Left            =   4560
         TabIndex        =   14
         Text            =   "ogun"
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox txtHomeAdd 
         Height          =   765
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Text            =   "IHMS_NewPatient.frx":68B5
         Top             =   1560
         Width           =   3015
      End
      Begin VB.TextBox txtSName 
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Text            =   "ogunbe"
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txtFName 
         Height          =   285
         Left            =   4560
         TabIndex        =   6
         Text            =   "dolapo"
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txtOccupation 
         Height          =   285
         Left            =   4560
         TabIndex        =   16
         Text            =   "student"
         Top             =   1440
         Width           =   1695
      End
      Begin VB.ComboBox cboSex 
         Height          =   315
         ItemData        =   "IHMS_NewPatient.frx":68D0
         Left            =   4560
         List            =   "IHMS_NewPatient.frx":68DA
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label10 
         Caption         =   "(DD/MM/YYYY)"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         ToolTipText     =   "DD/MM/YYYY"
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "State Of Origin:"
         Height          =   255
         Left            =   3360
         TabIndex        =   13
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblHomeAdd 
         Caption         =   "Address:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label lblSName 
         Alignment       =   1  'Right Justify
         Caption         =   "Surname:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblFName 
         Alignment       =   1  'Right Justify
         Caption         =   "First Name:"
         Height          =   255
         Left            =   3360
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Date Of Birth:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         ToolTipText     =   "DD/MM/YYYY"
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Sex:"
         Height          =   255
         Left            =   4080
         TabIndex        =   9
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Occupation:"
         Height          =   255
         Left            =   3360
         TabIndex        =   15
         Top             =   1440
         Width           =   1095
      End
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "New Patient Registration"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   3533
      TabIndex        =   42
      Top             =   120
      Width           =   5175
   End
   Begin VB.Label lblHospNo 
      Caption         =   "&Hospital Number:"
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
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   1575
   End
End
Attribute VB_Name = "frmNewReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public thisNewPatient As CPatient  'Public so that it can be accessed by frmAdmission
Private nextPatientNum As Integer

Private Sub cmdCancelReg_Click()
 If MsgBox("Are you sure you want to abort the current patient registration?" & vbCrLf & "(NOTE: You will loose all information you have entered.)", vbCritical + vbYesNo) = vbYes Then
    Unload Me
 End If
End Sub

Private Sub cmdRegAndAdmit_Click()
 Set thisNewPatient = New CPatient      'Create a Patient Object
 
 On Error GoTo errhnd
 Dim strErrorMessage As String
 datPerInfo.Recordset.Fields("Date_Of_Birth") = Trim(txtDOB)
 
 'write the info into a "patient" object
 With thisNewPatient
    .HospNo = Val(txtHospNo)
    .SName = Trim(txtSName)
    .FName = Trim(txtFName)
    .DoB = Trim(txtDOB)
    .Sex = cboSex
    .HomeAdd = Trim(txtHomeAdd)
    .StateOfOrigin = Trim(txtStateOfOrigin)
    .Occupation = Trim(txtOccupation)
    'NEXT OF KIN'S INFO
    .NameNoK = Trim(txtKinName)
    .RelaNok = Trim(txtRelationship)
    .AddNok = Trim(txtKinAddress)
    'SPONSOR'S INFO
    .SponsorName = Trim(txtNameOfSponsor)
    .SponsorAdd = Trim(txtAddOfSponsor)
    'LABORATORY INFO
    '.LabRefNo = Val(txtLabRefNo)
    .BloodGrp = Trim(cboBloodGrp)
    .RHFactor = cboRHFactor
    .Allergy = Trim(txtAllergy)
 End With
 frmAdmission.Show 1
 Call ClearRegForm 'clear the form b4 hiding it, in readiness for next use
 Unload Me
 Exit Sub
 
errhnd:
 Select Case Err.Number
 Case 3421
    strErrorMessage = "Input error. Please verify the syntax of the date with DD/MM/YYYY"
    strErrorMessage = strErrorMessage + vbCrLf
    strErrorMessage = strErrorMessage + "Also verify that you have not entered 29th feb. for a non-leap year."
    MsgBox strErrorMessage, , "Error #: " + Str(Err.Number)
    Exit Sub
 Case Else
    MsgBox "Unhandled error. Record could not be written to the DB. Please verify or re-enter your data."
   Exit Sub
 End Select
End Sub

Private Sub cmdRegOnly_Click()
 On Error GoTo errhnd
 Dim strErrorMessage As String
 'Collect the input
 With datPerInfo.Recordset
    'PERSONAL INFO
    .Fields("Date_Of_Birth") = Trim(txtDOB)
    .Fields("Hosp_No") = Val(txtHospNo)
    .Fields("SName") = Trim(txtSName)
    .Fields("FName") = Trim(txtFName)
    .Fields("Sex") = cboSex
    .Fields("Home_Add") = Trim(txtHomeAdd)
    .Fields("State_of_Origin") = Trim(txtStateOfOrigin)
    .Fields("Occupation") = Trim(txtOccupation)
    'NEXT OF KIN'S INFO
    .Fields("Name_of_NoK") = Trim(txtKinName)
    .Fields("Relationship_to_NoK") = Trim(txtRelationship)
    .Fields("Add_of_NoK") = Trim(txtKinAddress)
    'SPONSOR'S INFO
    .Fields("Name_of_Sponsor") = Trim(txtNameOfSponsor)
    .Fields("Add_of_Sponsor") = Trim(txtAddOfSponsor)
 End With
 With datLabInfo.Recordset
    'LABORATORY INFO
    '.Fields("Lab_Ref_No") = Val(txtLabRefNo)
    .Fields("Hosp_No") = Val(txtHospNo)
    .Fields("Blood_Group") = cboBloodGrp
    .Fields("RhFactor") = cboRHFactor
    .Fields("Allergy") = Trim(txtAllergy)
 End With
 
 'update d records into the db
 datPerInfo.Recordset.Update
 datLabInfo.Recordset.Update
 
 MsgBox "New Patient Registered Successfully.", vbInformation, "Success"
 Call ClearRegForm
 Unload Me
 Exit Sub
 
errhnd:
 Select Case Err.Number
 Case 3421
    strErrorMessage = "Input error. Please verify the syntax of the date with DD/MM/YYYY"
    strErrorMessage = strErrorMessage + vbCrLf
    strErrorMessage = strErrorMessage + "Also verify that you have not entered 29th feb. for a non-leap year."
    MsgBox strErrorMessage, , "Error #: " + Str(Err.Number)
    Exit Sub
 Case Else
    MsgBox "Unhandled error. Record could not be written to the DB. Please verify or re-enter your data."
   Exit Sub
 End Select
End Sub

Private Sub cmdRestartReg_Click()
 If MsgBox("Are you sure you want to clear the form and start over?", vbQuestion + vbYesNo) = vbYes Then
    'clear form
    Call ClearRegForm
 End If
End Sub

Private Sub Form_Load()
'On Error Resume Next

 datPerInfo.DatabaseName = App.Path & "\IHMS_97.mdb"
 datPerInfo.RecordSource = "Patient_Personal_Info"
 datPerInfo.Refresh
  
 datLabInfo.DatabaseName = App.Path + "\IHMS_97.mdb"
 datLabInfo.RecordSource = "Patient_Lab_Info"
 datLabInfo.Refresh
 
 'Automatically generate the next hosp. no.
 'Note that this has to be done before adding a new record, otherwise, you will get 0 (zero)
 datPerInfo.Recordset.MoveLast
 nextPatientNum = datPerInfo.Recordset.Fields("Hosp_No") + 1
 
 'Insert a new record at end of the DB.
 datPerInfo.Recordset.AddNew
 datLabInfo.Recordset.AddNew
 txtHospNo = nextPatientNum
 txtLabRefNo = datLabInfo.Recordset.Fields("Lab_Ref_No")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 If UnloadMode = 4 Then Exit Sub 'The MDI child form is closing because the MDI form is closing.
 If UnloadMode = 1 Then Exit Sub 'The Unload statement is invoked from code.
 If MsgBox("Are you sure you want to abort the current patient registration?" & vbCrLf & "(NOTE: You will loose all information you have entered.)", vbCritical + vbYesNo) = vbYes Then
    Cancel = 0
 Else
    Cancel = 1
 End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
 frmMain.mnuNewPatient.Enabled = True   'enable new patient menu item
 frmMain.tbrMainToolbar.Buttons(2).Enabled = True  'enable new patient toolbar button
End Sub

Private Sub txtSName_GotFocus()
 'Display the next hosp. number
End Sub
