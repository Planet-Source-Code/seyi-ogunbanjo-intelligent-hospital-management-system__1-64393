VERSION 5.00
Begin VB.Form frmDischarge 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Discharge - Existing Patient"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6915
   ControlBox      =   0   'False
   Icon            =   "IHMS_DischargePatient.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   6915
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data datHospHist 
      Caption         =   "Hospital History"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton cmdAbortDischarge 
      Cancel          =   -1  'True
      Caption         =   "Abort &Task"
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
      Left            =   3600
      TabIndex        =   12
      Top             =   4440
      Width           =   2895
   End
   Begin VB.CommandButton cmdConfirmDischarge 
      Caption         =   "&Discharge Patient"
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
      Left            =   360
      TabIndex        =   11
      Top             =   4440
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      Caption         =   "&Admission Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   360
      TabIndex        =   0
      Top             =   1080
      Width           =   6135
      Begin VB.ComboBox cboDischargeStatus 
         DataField       =   "Sex"
         DataSource      =   "datPerInfo"
         Height          =   315
         ItemData        =   "IHMS_DischargePatient.frx":6852
         Left            =   2640
         List            =   "IHMS_DischargePatient.frx":6862
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2760
         Width           =   3255
      End
      Begin VB.TextBox txtDateDischarged 
         DataField       =   "Date_of_Birth"
         DataSource      =   "datPerInfo"
         Height          =   285
         Left            =   240
         TabIndex        =   9
         ToolTipText     =   "DD/MM/YYYY"
         Top             =   2760
         Width           =   2055
      End
      Begin VB.TextBox txtDoctorsDiag 
         BackColor       =   &H80000016&
         DataSource      =   "datContactInfo"
         Enabled         =   0   'False
         Height          =   1725
         Left            =   2640
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   600
         Width           =   3255
      End
      Begin VB.TextBox txtAdmissionDate 
         BackColor       =   &H80000016&
         DataField       =   "Surname"
         DataSource      =   "datPerInfo"
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox txtCaseRefNo 
         BackColor       =   &H80000016&
         DataField       =   "First_Name"
         DataSource      =   "datPerInfo"
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         TabIndex        =   4
         Top             =   1320
         Width           =   2055
      End
      Begin VB.TextBox txtDoctorInCharge 
         BackColor       =   &H80000016&
         DataField       =   "Date_of_Birth"
         DataSource      =   "datPerInfo"
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         TabIndex        =   6
         ToolTipText     =   "DD/MM/YYYY"
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label Label4 
         Caption         =   "Status upon discharge:"
         Height          =   255
         Left            =   2640
         TabIndex        =   14
         ToolTipText     =   "DD/MM/YYYY"
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Date of Discharge:"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         ToolTipText     =   "DD/MM/YYYY"
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "Doctor's Diagnosis"
         Height          =   255
         Left            =   2640
         TabIndex        =   7
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Date of Admission:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Case Reference #:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Doctor-in-Charge:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         ToolTipText     =   "DD/MM/YYYY"
         Top             =   1800
         Width           =   1335
      End
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Discharge Form"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   1410
      TabIndex        =   16
      Top             =   600
      Width           =   4095
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
      Left            =   1830
      TabIndex        =   15
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "frmDischarge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAbortDischarge_Click()
 If MsgBox("Abort the discharge of this patient?", vbCritical + vbYesNo) = vbYes Then
    Unload Me
 End If
End Sub

Private Sub cmdConfirmDischarge_Click()
 With datHospHist.Recordset
    .Edit
    .Fields("Admission_Status") = "OUT"
    .Fields("Date_of_Discharge") = txtDateDischarged
    .Fields("Status_Upon_Discharge") = cboDischargeStatus
    .Update
 End With
 MsgBox "Patient DISCHARGED has been effected in the database.", vbInformation, "Success"
 Unload frmOldPatient
 Unload Me
End Sub

Private Sub Form_Load()
 Dim flgFound As Boolean
 lblHeading.Caption = lblHeading.Caption + Str(somePatient.HospNo)
 datHospHist.DatabaseName = App.Path & "\IHMS_97.mdb"
 datHospHist.RecordSource = "Patient_Hospital_History"
 datHospHist.Refresh
 
 'SEARCH THE PATIENT_HOSPITAL_HISTORY TABLE (for the patient's MOST RECENT record:
 datHospHist.Recordset.MoveLast
 With datHospHist.Recordset
    Do
        If .Fields("Hosp_No") = somePatient.HospNo Then
            flgFound = True
        Else
            .MovePrevious
        End If
    Loop Until (.BOF) Or (flgFound)
 End With
 
 'Display information that's already been collected.
 'If flgFound = False Then 'Although highly unlikely
 With somePatient
    txtAdmissionDate = .AdmissionDate
    txtCaseRefNo = datHospHist.Recordset.Fields("Case_Ref_No")
    txtDoctorInCharge = .DocName
    txtDoctorsDiag = .Diagnosis
 End With
End Sub

