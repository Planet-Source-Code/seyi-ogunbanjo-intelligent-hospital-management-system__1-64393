VERSION 5.00
Begin VB.Form frmDiagnosis 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Diagnose - Existing Patient"
   ClientHeight    =   9090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7260
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   7260
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Save Results and Close"
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
      Left            =   360
      TabIndex        =   31
      Top             =   8400
      Width           =   3255
   End
   Begin VB.CommandButton cmdClose 
      Appearance      =   0  'Flat
      Caption         =   "Close (Without Saving)"
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
      Left            =   3720
      TabIndex        =   30
      Top             =   8400
      Width           =   3135
   End
   Begin VB.Frame fraDiagResults 
      Caption         =   "&Result of Diagnosis:"
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
      Height          =   2535
      Left            =   360
      TabIndex        =   25
      Top             =   5760
      Width           =   6495
      Begin VB.TextBox txtDisease 
         BackColor       =   &H80000016&
         DataField       =   "Date_of_Birth"
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
         Height          =   1125
         Index           =   1
         Left            =   360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   29
         ToolTipText     =   "DD/MM/YYYY"
         Top             =   1200
         Width           =   5775
      End
      Begin VB.TextBox txtDisease 
         BackColor       =   &H80000016&
         DataField       =   "Date_of_Birth"
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
         Index           =   0
         Left            =   540
         TabIndex        =   26
         ToolTipText     =   "DD/MM/YYYY"
         Top             =   480
         Width           =   5415
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "However, he/she could also be treated for the following diseases/illnesses:"
         Height          =   195
         Left            =   600
         TabIndex        =   28
         ToolTipText     =   "DD/MM/YYYY"
         Top             =   960
         Width           =   5295
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "The most likely disease this patient that this patient is suffering from is:"
         Height          =   195
         Left            =   795
         TabIndex        =   27
         ToolTipText     =   "DD/MM/YYYY"
         Top             =   240
         Width           =   4905
      End
   End
   Begin VB.CommandButton cmdDiagnose 
      Caption         =   "&Diagnose Patient"
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
      TabIndex        =   0
      Top             =   5160
      Width           =   6495
   End
   Begin VB.Data datSymptoms 
      Caption         =   "Symptoms"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5520
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data datDiseases 
      Caption         =   "Diseases"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5520
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Frame fraDiagInfo 
      Caption         =   "&Diagnosis Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   360
      TabIndex        =   8
      Top             =   2280
      Width           =   6495
      Begin VB.ComboBox cboSymptom 
         DataField       =   "Sex"
         DataSource      =   "datPerInfo"
         Height          =   315
         Index           =   4
         ItemData        =   "IHMS_Diagnosis.frx":0000
         Left            =   1200
         List            =   "IHMS_Diagnosis.frx":0007
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   2280
         Width           =   5055
      End
      Begin VB.ComboBox cboSymptom 
         DataField       =   "Sex"
         DataSource      =   "datPerInfo"
         Height          =   315
         Index           =   3
         ItemData        =   "IHMS_Diagnosis.frx":0010
         Left            =   1200
         List            =   "IHMS_Diagnosis.frx":0017
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   1920
         Width           =   5055
      End
      Begin VB.ComboBox cboSymptom 
         DataField       =   "Sex"
         DataSource      =   "datPerInfo"
         Height          =   315
         Index           =   2
         ItemData        =   "IHMS_Diagnosis.frx":0020
         Left            =   1200
         List            =   "IHMS_Diagnosis.frx":0027
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   1560
         Width           =   5055
      End
      Begin VB.ComboBox cboSymptom 
         DataField       =   "Sex"
         DataSource      =   "datPerInfo"
         Height          =   315
         Index           =   1
         ItemData        =   "IHMS_Diagnosis.frx":0030
         Left            =   1200
         List            =   "IHMS_Diagnosis.frx":0037
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   1200
         Width           =   5055
      End
      Begin VB.ComboBox cboSymptom 
         DataField       =   "Sex"
         DataSource      =   "datPerInfo"
         Height          =   315
         Index           =   0
         ItemData        =   "IHMS_Diagnosis.frx":0040
         Left            =   1200
         List            =   "IHMS_Diagnosis.frx":0047
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   840
         Width           =   5055
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "the top, and select N/A for the rest, if not applicable."
         Height          =   195
         Left            =   360
         TabIndex        =   24
         ToolTipText     =   "DD/MM/YYYY"
         Top             =   480
         Width           =   3705
      End
      Begin VB.Label lblSymptom 
         AutoSize        =   -1  'True
         Caption         =   "Symptom 5"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   17
         ToolTipText     =   "DD/MM/YYYY"
         Top             =   2280
         Width           =   780
      End
      Begin VB.Label lblSymptom 
         AutoSize        =   -1  'True
         Caption         =   "Symptom 4"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   16
         ToolTipText     =   "DD/MM/YYYY"
         Top             =   1920
         Width           =   780
      End
      Begin VB.Label lblSymptom 
         AutoSize        =   -1  'True
         Caption         =   "Symptom 3"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   15
         ToolTipText     =   "DD/MM/YYYY"
         Top             =   1560
         Width           =   780
      End
      Begin VB.Label lblSymptom 
         AutoSize        =   -1  'True
         Caption         =   "Symptom 2"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   14
         ToolTipText     =   "DD/MM/YYYY"
         Top             =   1200
         Width           =   780
      End
      Begin VB.Label lblSymptom 
         AutoSize        =   -1  'True
         Caption         =   "Symptom 1"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   13
         ToolTipText     =   "DD/MM/YYYY"
         Top             =   840
         Width           =   780
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "symptoms from the combo boxes, starting from"
         Height          =   195
         Left            =   2880
         TabIndex        =   12
         ToolTipText     =   "DD/MM/YYYY"
         Top             =   240
         Width           =   3240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "AT LEAST ONE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1440
         TabIndex        =   11
         ToolTipText     =   "DD/MM/YYYY"
         Top             =   240
         Width           =   1350
      End
      Begin VB.Label Label4 
         Caption         =   "Please select "
         Height          =   255
         Left            =   360
         TabIndex        =   10
         ToolTipText     =   "DD/MM/YYYY"
         Top             =   240
         Width           =   1095
      End
   End
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
   Begin VB.Frame fraPatientInfo 
      Caption         =   "&Patient Information"
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
      Height          =   1215
      Left            =   360
      TabIndex        =   3
      Top             =   960
      Width           =   6495
      Begin VB.TextBox txtCaseRefNo 
         BackColor       =   &H80000016&
         DataField       =   "First_Name"
         DataSource      =   "datPerInfo"
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox txtDoctorInCharge 
         BackColor       =   &H80000016&
         DataField       =   "Date_of_Birth"
         DataSource      =   "datPerInfo"
         Height          =   285
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   5
         ToolTipText     =   "DD/MM/YYYY"
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txtHospNo 
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label12 
         Caption         =   "Case Reference #:"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Doctor-in-Charge:"
         Height          =   255
         Left            =   4320
         TabIndex        =   7
         ToolTipText     =   "DD/MM/YYYY"
         Top             =   360
         Width           =   1335
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
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Diagnosis Assistant"
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
      TabIndex        =   2
      Top             =   480
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
      TabIndex        =   1
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "frmDiagnosis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
 If MsgBox("Close Diagnosis Assistant without saving?", vbCritical + vbYesNo) = vbYes Then
    Unload Me
 End If
End Sub

Private Sub cmdDiagnose_Click()
 Dim symptom(4) As String
 Dim diseaseID(4) As Integer
 Dim diseaseName(4) As String
 Dim diagnosisResults(4) As String
 
 Dim x As Integer   'Loop variable
 
 'Collect user input
 For x = 0 To 4
    symptom(x) = cboSymptom(x)
 Next x
 
 'Ensure the user chooses at least one symptom, otherwise, exit this procedure
 If symptom(0) = "N/A" Then symptom(0) = ""
 If symptom(0) = "" Then
    MsgBox "You must specify AT LEAST one sign/symptom to diagnose a patient, starting with Symptom 1.", vbInformation
    Exit Sub
 End If
 
 Call DiagnosePatient(symptom(0), diseaseID(0), diseaseName(0))
 txtDisease(0).Text = diseaseName(0)
 
 'Generate results for any other symptoms specified.
 For x = 1 To 4
    If symptom(x) <> "" Then
        Call DiagnosePatient(symptom(x), diseaseID(x), diseaseName(x))
        If diseaseName(x) <> diseaseName(0) Then diagnosisResults(x) = diseaseName(x)
    End If
 Next x
 
 'Sort the 4 resulting disease names in descending order, and eliminate duplicates
 'MsgBox diseaseName(0) + vbCrLf + diseaseName(1) + vbCrLf + diseaseName(2) + vbCrLf + diseaseName(3) + vbCrLf + diseaseName(4)
 
 Dim a As Integer, b As Integer
 Dim temp As String
 For a = 1 To 3
    For b = 1 To 4 - a
        If diagnosisResults(b) > diagnosisResults(b + 1) Then
            temp = diagnosisResults(b)
            diagnosisResults(b) = diagnosisResults(b + 1)
            diagnosisResults(b + 1) = temp
        ElseIf diagnosisResults(b) = diagnosisResults(b + 1) Then
            'eliminate duplicates
            diagnosisResults(b) = ""
        End If
    Next b
 Next a
 
 'Display the results.
 txtDisease(1) = ""
 For a = 1 To 4
    If diagnosisResults(a) <> "" Then txtDisease(1) = txtDisease(1) + diagnosisResults(a) + vbCrLf
 Next a
 
 fraDiagInfo.Enabled = False

 cmdDiagnose.Enabled = False
 cmdSave.Enabled = True
End Sub

Private Sub cmdSave_Click()
 Dim flgFound As Boolean
 'SEARCH THE PATIENT_HOSPITAL_HISTORY TABLE (for the patient's MOST RECENT record:
 With datHospHist.Recordset
    .MoveLast
    Do
        If .Fields("Hosp_No") = somePatient.HospNo Then
            flgFound = True
        Else
            .MovePrevious
        End If
    Loop Until (.BOF) Or (flgFound)
 End With
 
 With datHospHist.Recordset
    .Edit
    .Fields("Doctors_Diagnosis") = txtDisease(0)
    .Update
 End With
 MsgBox "Diagnosis Results have been recorded.", vbInformation, "Success"
 Unload Me
 Unload frmOldPatient
End Sub

Private Sub Form_Load()
 datSymptoms.DatabaseName = App.Path & "\IHMS_97.mdb"
 datSymptoms.RecordSource = "Symptoms"
 datSymptoms.Refresh

 datDiseases.DatabaseName = App.Path & "\IHMS_97.mdb"
 datDiseases.RecordSource = "Diseases"
 datDiseases.Refresh
 
 datHospHist.DatabaseName = App.Path & "\IHMS_97.mdb"
 datHospHist.RecordSource = "Patient_Hospital_History"
 datHospHist.Refresh

 lblHeading.Caption = lblHeading.Caption + Str(somePatient.HospNo)

 txtHospNo = somePatient.HospNo
 txtCaseRefNo = somePatient.CaseRefNo
 txtDoctorInCharge = somePatient.DocName
 
 'Populate the combo boxes...
 Dim y As Integer
 Dim x As Integer
 With datSymptoms.Recordset
    .MoveFirst
    y = 1
    Do While Not .EOF
        For x = 0 To 4
            cboSymptom(x).List(y) = .Fields("Symptom_Name")
        Next x
        y = y + 1
        .MoveNext
    Loop
 End With
 
End Sub

Public Sub DiagnosePatient(symptomX As String, d_ID As Integer, d_Name As String)
 'This procedure does the following:
 '1. Search for a symptom in the Symptom table, and "get" its Disease_ID
 '2. Match the disease_ID against a Disease_Name in the Diseases table
 '3. Return the Disease_Name found in the variable "d_Name"
 '     to the calling procedure.

 'Task 1: Locate the disease_ID in the Symptoms table
 Dim flgFoundSymptom As Boolean
 With datSymptoms.Recordset
    .MoveFirst
    flgFoundSymptom = False
    Do
        If symptomX = .Fields("Symptom_Name") Then
            flgFoundSymptom = True
        Else
            .MoveNext
        End If
    Loop Until flgFoundSymptom = True 'Or .EOF
    d_ID = .Fields("Disease_ID")
 End With
 
 'Task 2: Use the disease_ID to locate the Diseases_Name in the Diseases table
 With datDiseases.Recordset
    .MoveFirst
    Do
        If .Fields("Disease_ID") = d_ID Then
            'Task 3: Return the name of the disease
            d_Name = .Fields("Diseases_Name")
        Else
            .MoveNext
        End If
    Loop Until d_Name = .Fields("Diseases_Name") 'Or .EOF
 End With

End Sub

