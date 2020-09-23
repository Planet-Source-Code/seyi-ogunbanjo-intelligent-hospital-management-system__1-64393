VERSION 5.00
Begin VB.Form frmWait 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Searching..."
   ClientHeight    =   1560
   ClientLeft      =   5880
   ClientTop       =   3105
   ClientWidth     =   3495
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   3495
   Begin VB.Data datLabInfo 
      Caption         =   "Lab Info"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   1980
   End
   Begin VB.Data datPerInfo 
      Caption         =   "Personal Info"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   -240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.Data datHospHist 
      Caption         =   "Hospital History"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1320
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Please Wait..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   975
      TabIndex        =   0
      Top             =   1080
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   1200
      Picture         =   "IHMS_SearchProgress.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_GotFocus()
 Dim foundPersonalRec As Boolean
 Dim foundLabRec As Boolean
 Dim foundHospRec As Boolean
 Set somePatient = New CPatient
 
 'SEARCH THE PATIENT_PERSONAL_INFO TABLE:
 datPerInfo.Recordset.MoveFirst
 With datPerInfo.Recordset
    Do
        If .Fields("Hosp_No") = patientNumberX Then
            foundPersonalRec = True
        Else
            .MoveNext
        End If
    Loop Until (.EOF) Or (foundPersonalRec)
 End With
 
 'SEARCH THE PATIENT_LAB_INFO TABLE:
 datLabInfo.Recordset.MoveFirst
 With datLabInfo.Recordset
    Do
        If .Fields("Hosp_No") = patientNumberX Then
            foundLabRec = True
        Else
            .MoveNext
        End If
    Loop Until (.EOF) Or (foundLabRec)
 End With
 
 'SEARCH THE PATIENT_HOSPITAL_HISTORY TABLE (for the patient's MOST RECENT record:
 datHospHist.Recordset.MoveLast
 With datHospHist.Recordset
    Do
        If .Fields("Hosp_No") = patientNumberX Then
            foundHospRec = True
        Else
            .MovePrevious
        End If
    Loop Until (.BOF) Or (foundHospRec)
 End With
 
 If (foundPersonalRec = True) Then
    'The sought record was found: Display it
    Call GetPatientData(foundLabRec, foundHospRec)
    Unload Me
    frmOldPatient.Show
    frmMain.mnuClose.Enabled = True
 Else
    'The sought record doesn't exist: so inform the user.
    MsgBox "No record found for the patient number you entered. Sorry!", vbInformation, "Search Failed"
    Unload Me
 End If
End Sub

Private Sub Form_Load()
 Me.MousePointer = 11 'vbHourglass
 
 datPerInfo.DatabaseName = App.Path & "\IHMS_97.mdb"
 datPerInfo.RecordSource = "Patient_Personal_Info"
 datPerInfo.Refresh
  
 datLabInfo.DatabaseName = App.Path + "\IHMS_97.mdb"
 datLabInfo.RecordSource = "Patient_Lab_Info"
 datLabInfo.Refresh
 
 datHospHist.DatabaseName = App.Path & "\IHMS_97.mdb"
 datHospHist.RecordSource = "Patient_Hospital_History"
 datHospHist.Refresh
End Sub

Private Sub GetPatientData(flg1 As Boolean, flg2 As Boolean)
 With datPerInfo.Recordset
    'PERSONAL INFO
    'MsgBox Str(somePatient.HospNo)
    somePatient.HospNo = .Fields("Hosp_No")
    somePatient.SName = .Fields("SName")
    somePatient.FName = .Fields("FName")
    'somePatient.DoB = .Fields("Date_Of_Birth")
    somePatient.Sex = .Fields("Sex")
    somePatient.HomeAdd = .Fields("Home_Add")
    somePatient.StateOfOrigin = .Fields("State_of_Origin")
    somePatient.Occupation = .Fields("Occupation")
    
    somePatient.NameNoK = .Fields("Name_of_NoK")
    somePatient.RelaNok = .Fields("Relationship_to_NoK")
    somePatient.AddNok = .Fields("Add_of_NoK")
    
    somePatient.SponsorName = .Fields("Name_of_Sponsor")
    somePatient.SponsorAdd = .Fields("Add_of_Sponsor")
    
    somePatient.AdmissionStatus = ""
 End With
    
 With datLabInfo.Recordset
    'LAB INFO
    If (flg1 = False) Then Exit Sub 'although this is highly unlikely  :)
    somePatient.LabRefNo = .Fields("Lab_Ref_No")
    somePatient.BloodGrp = .Fields("Blood_Group")
    somePatient.RHFactor = .Fields("RhFactor")
    somePatient.Allergy = .Fields("Allergy")
 End With
 
 With datHospHist.Recordset
    'HOSPITAL HISTORY INFO
    If (flg2 = False) Then Exit Sub 'this is possible for patients
    'who have never been admitted: a hospital history record has never been created for them.
    somePatient.CaseRefNo = .Fields("Case_Ref_No")
    somePatient.AdmissionStatus = .Fields("Admission_Status")
    somePatient.AdmissionDate = .Fields("Date_of_Admission")
    somePatient.DocName = .Fields("Name_of_Doctor")
    somePatient.Diagnosis = .Fields("Doctors_Diagnosis")
    somePatient.DateDischarged = .Fields("Date_of_Discharge")
    somePatient.DischargeStatus = .Fields("Status_Upon_Discharge")
 End With
End Sub
