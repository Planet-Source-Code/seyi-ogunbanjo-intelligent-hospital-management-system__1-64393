VERSION 5.00
Begin VB.Form frmOldPatient 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View Existing Patient"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10935
   ControlBox      =   0   'False
   Icon            =   "IHMS_ExistingPatient.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   10935
   Begin VB.Data datPerInfo 
      Caption         =   "Personal Info"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   2400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4560
      Visible         =   0   'False
      Width           =   2340
   End
   Begin VB.Data datLabInfo 
      Caption         =   "Lab Info"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   4800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4560
      Visible         =   0   'False
      Width           =   2580
   End
   Begin VB.CommandButton cmdCloseFile 
      Cancel          =   -1  'True
      Caption         =   "&Close (Unload) Patient File"
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
      Left            =   5640
      TabIndex        =   29
      Top             =   3960
      Width           =   4815
   End
   Begin VB.CommandButton cmdDiagnose 
      Caption         =   "&Diagnose Patient's Illness"
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
      Left            =   5640
      TabIndex        =   27
      Top             =   3480
      Width           =   4815
   End
   Begin VB.CommandButton cmdDischarge 
      Caption         =   "Dis&charge Patient"
      Enabled         =   0   'False
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
      Left            =   240
      TabIndex        =   28
      Top             =   3960
      Width           =   4815
   End
   Begin VB.CommandButton cmdAdmit 
      Caption         =   "&Admit Patient into care"
      Enabled         =   0   'False
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
      Left            =   240
      TabIndex        =   26
      Top             =   3480
      Width           =   4815
   End
   Begin VB.Frame fraLabInfo 
      Caption         =   "&Laboratory Information"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   6840
      TabIndex        =   17
      Top             =   1080
      Width           =   3615
      Begin VB.TextBox txtRHFactor 
         BackColor       =   &H80000016&
         Height          =   285
         Left            =   1680
         TabIndex        =   23
         ToolTipText     =   "DD/MM/YYYY"
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox txtBloodGroup 
         BackColor       =   &H80000016&
         Height          =   285
         Left            =   1680
         TabIndex        =   21
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox txtLabRefNo 
         BackColor       =   &H80000016&
         Height          =   285
         Left            =   1680
         TabIndex        =   19
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtAllergy 
         BackColor       =   &H80000016&
         Height          =   285
         Left            =   1680
         TabIndex        =   25
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label lblLabAllergy 
         Alignment       =   1  'Right Justify
         Caption         =   "RH Factor:"
         Height          =   255
         Left            =   480
         TabIndex        =   22
         ToolTipText     =   "DD/MM/YYYY"
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label lblLabBloodGroup 
         Alignment       =   1  'Right Justify
         Caption         =   "Blood Group:"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lblLabRefNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Lab. Ref. Number:"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lblLabGenotype 
         Alignment       =   1  'Right Justify
         Caption         =   "Allergy:"
         Height          =   255
         Left            =   480
         TabIndex        =   24
         Top             =   1680
         Width           =   1095
      End
   End
   Begin VB.TextBox txtHospNo 
      BackColor       =   &H80000016&
      Enabled         =   0   'False
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
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   6495
      Begin VB.TextBox txtSex 
         BackColor       =   &H80000016&
         DataField       =   "First_Name"
         DataSource      =   "datPerInfo"
         Height          =   285
         Left            =   4560
         TabIndex        =   10
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtStateOfOrigin 
         BackColor       =   &H80000016&
         DataSource      =   "datPerInfo"
         Height          =   285
         Left            =   4560
         TabIndex        =   14
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox txtHomeAdd 
         BackColor       =   &H80000016&
         DataSource      =   "datContactInfo"
         Height          =   765
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   1200
         Width           =   3015
      End
      Begin VB.TextBox txtSName 
         BackColor       =   &H80000016&
         DataField       =   "Surname"
         DataSource      =   "datPerInfo"
         Height          =   285
         Left            =   1560
         TabIndex        =   4
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtFName 
         BackColor       =   &H80000016&
         DataField       =   "First_Name"
         DataSource      =   "datPerInfo"
         Height          =   285
         Left            =   4560
         TabIndex        =   6
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtDOB 
         BackColor       =   &H80000016&
         DataField       =   "Date_of_Birth"
         DataSource      =   "datPerInfo"
         Height          =   285
         Left            =   1560
         TabIndex        =   8
         ToolTipText     =   "DD/MM/YYYY"
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtOccupation 
         BackColor       =   &H80000016&
         DataSource      =   "datPerInfo"
         Height          =   285
         Left            =   4560
         TabIndex        =   16
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "State Of Origin:"
         Height          =   255
         Left            =   3360
         TabIndex        =   13
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lblHomeAdd 
         Caption         =   "Address:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lblSName 
         Alignment       =   1  'Right Justify
         Caption         =   "Surname:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblFName 
         Alignment       =   1  'Right Justify
         Caption         =   "First Name:"
         Height          =   255
         Left            =   3360
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Date Of Birth:"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         ToolTipText     =   "DD/MM/YYYY"
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Sex:"
         Height          =   255
         Left            =   4080
         TabIndex        =   9
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Occupation:"
         Height          =   255
         Left            =   3360
         TabIndex        =   15
         Top             =   1320
         Width           =   1095
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "PS: You need to ADMIT a patient before he/she can be DIAGNOSED."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   2475
      TabIndex        =   31
      Top             =   4680
      Width           =   5985
   End
   Begin VB.Label lblHeading 
      Alignment       =   2  'Center
      Caption         =   "Patient#"
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
      Left            =   3000
      TabIndex        =   30
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
Attribute VB_Name = "frmOldPatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdmit_Click()
 frmAdmitExisting.Show 1
End Sub

Private Sub cmdCloseFile_Click()
 If MsgBox("Unload and close patient file?", vbYesNo + vbQuestion) = vbYes Then Unload Me
End Sub

Private Sub cmdDiagnose_Click()
 frmDiagnosis.Show 1
End Sub

Private Sub cmdDischarge_Click()
 frmDischarge.Show 1
End Sub

Private Sub cmdViewContact_Click()
 MsgBox "This function is still under development.", vbInformation
End Sub

Private Sub Form_Activate()
 If somePatient.AdmissionStatus = "IN" Then
    frmMain.tbrMainToolbar.Buttons(5).Enabled = True 'diagnose toolbar button
 Else
    frmMain.tbrMainToolbar.Buttons(5).Enabled = False 'diagnose toolbar button
 End If
 frmMain.tbrMainToolbar.Buttons(4).Enabled = True 'admit/discharge button
End Sub

Private Sub Form_Deactivate()
 frmMain.tbrMainToolbar.Buttons(4).Enabled = False 'admit/discharge button
 frmMain.tbrMainToolbar.Buttons(5).Enabled = False 'diagnose button
End Sub

Private Sub Form_Load()
    'PERSONAL INFO
    txtHospNo = somePatient.HospNo
    txtSName = somePatient.SName
    txtFName = somePatient.FName
    'txtDOB = somePatient.DoB
    txtSex = somePatient.Sex
    txtHomeAdd = somePatient.HomeAdd
    txtStateOfOrigin = somePatient.StateOfOrigin
    txtOccupation = somePatient.Occupation
    
    'LAB INFO
    txtLabRefNo = somePatient.LabRefNo
    txtBloodGroup = somePatient.BloodGrp
    txtRHFactor = somePatient.RHFactor
    txtAllergy = somePatient.Allergy
    
    lblHeading.Caption = lblHeading.Caption + Str(somePatient.HospNo)
    If somePatient.AdmissionStatus = "IN" Then
        'Patient is currently admitted. therefore, the admit command is
        'not available, but the discharge command is.
        cmdDischarge.Enabled = True
        cmdDiagnose.Enabled = True
        frmMain.tbrMainToolbar.Buttons(5).Enabled = True 'diagnose toolbar button
    Else
        'Patient is currently NOT admitted. therefore, the admit command is
        'available, but the discharge command is not.
        'This part is also executed if no existing hospital history record is found for the patient
        'i.e when somePatient.AdmissionStatus = ""
        cmdAdmit.Enabled = True
        cmdDiagnose.Enabled = False
        frmMain.tbrMainToolbar.Buttons(5).Enabled = False 'diagnose toolbar button
    End If
    frmMain.tbrMainToolbar.Buttons(4).Enabled = True 'admit/discharge button
End Sub

Private Sub Form_Unload(Cancel As Integer)
 frmMain.mnuClose.Enabled = False
 frmMain.tbrMainToolbar.Buttons(4).Enabled = False 'admit/discharge button
 frmMain.tbrMainToolbar.Buttons(5).Enabled = False 'diagnose button
End Sub
