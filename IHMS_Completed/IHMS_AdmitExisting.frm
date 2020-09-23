VERSION 5.00
Begin VB.Form frmAdmitExisting 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Admission - Existing Patient"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6915
   ControlBox      =   0   'False
   Icon            =   "IHMS_AdmitExisting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
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
   Begin VB.CommandButton cmdAbortAdmission 
      Cancel          =   -1  'True
      Caption         =   "Abort Admission Process"
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
      TabIndex        =   10
      Top             =   3960
      Width           =   2895
   End
   Begin VB.CommandButton cmdConfirmAdmission 
      Caption         =   "&Confirm Admission"
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
      TabIndex        =   9
      Top             =   3960
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
      Height          =   2655
      Left            =   360
      TabIndex        =   0
      Top             =   1200
      Width           =   6135
      Begin VB.TextBox txtDoctorsComments 
         DataSource      =   "datContactInfo"
         Height          =   1725
         Left            =   2640
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   600
         Width           =   3255
      End
      Begin VB.TextBox txtDateOfAdmission 
         DataField       =   "Surname"
         DataSource      =   "datPerInfo"
         Height          =   285
         Left            =   240
         TabIndex        =   4
         ToolTipText     =   "DD/MM/YYYY"
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox txtCaseRefNo 
         BackColor       =   &H80000016&
         DataField       =   "First_Name"
         DataSource      =   "datPerInfo"
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
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox txtDoctorInCharge 
         Height          =   285
         Left            =   240
         TabIndex        =   6
         ToolTipText     =   "DD/MM/YYYY"
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label Label11 
         Caption         =   "Doctor's Comments"
         Height          =   255
         Left            =   2640
         TabIndex        =   7
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Admission Date (dd/mm/yyyy):"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         ToolTipText     =   "DD/MM/YYYY"
         Top             =   1200
         Width           =   2535
      End
      Begin VB.Label Label9 
         Caption         =   "Case Reference #:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   480
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
      TabIndex        =   12
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Admission Form"
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
      TabIndex        =   11
      Top             =   600
      Width           =   4095
   End
End
Attribute VB_Name = "frmAdmitExisting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAbortAdmission_Click()
 If MsgBox("Abort the admission of this patient?", vbCritical + vbYesNo) = vbYes Then
    Unload Me
 End If
End Sub

Private Sub cmdConfirmAdmission_Click()
'On Error GoTo errHnd
 
 With Me.datHospHist.Recordset
    '.AddNew
    .Fields("Hosp_No") = somePatient.HospNo
    .Fields("Admission_Status") = "IN"
    .Fields("Date_of_Admission") = txtDateOfAdmission
    .Fields("Name_of_Doctor") = txtDoctorInCharge
    .Fields("Doctors_Diagnosis") = txtDoctorsComments
    .Update
 End With
 
 MsgBox "Patient has been ADMITTED into care.", vbInformation, "Success"
 Unload frmOldPatient
 Unload Me
 Exit Sub

errhnd:
 Debug.Print Err.Number; "   "; Err.Description
 MsgBox "An error has occured.", vbInformation, "Unhandled error!"
 Resume Next
End Sub

Private Sub Form_Load()
 lblHeading.Caption = lblHeading.Caption + Str(somePatient.HospNo)
 datHospHist.DatabaseName = App.Path & "\IHMS_97.mdb"
 datHospHist.RecordSource = "Patient_Hospital_History"
 datHospHist.Refresh
 datHospHist.Recordset.AddNew
 
 'Display information that's already been collected.
 txtCaseRefNo = datHospHist.Recordset.Fields("Case_Ref_No")
End Sub
