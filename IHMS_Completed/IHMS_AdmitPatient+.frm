VERSION 5.00
Begin VB.Form frmAdmission 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Admission - New Patient"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6915
   ControlBox      =   0   'False
   Icon            =   "IHMS_AdmitPatient.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7005
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
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6600
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton cmdAbortAdmission 
      Cancel          =   -1  'True
      Caption         =   "Abort Registration Process"
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
      TabIndex        =   26
      Top             =   6120
      Width           =   2895
   End
   Begin VB.CommandButton cmdConfirmAdmission 
      Caption         =   "A&dmit Patient"
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
      TabIndex        =   25
      Top             =   6120
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
      TabIndex        =   16
      Top             =   3240
      Width           =   6135
      Begin VB.TextBox txtDoctorsDiag 
         DataSource      =   "datContactInfo"
         Height          =   1725
         Left            =   2640
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   24
         Text            =   "IHMS_AdmitPatient.frx":6852
         Top             =   600
         Width           =   3255
      End
      Begin VB.TextBox txtDateOfAdmission 
         DataField       =   "Surname"
         DataSource      =   "datPerInfo"
         Height          =   285
         Left            =   240
         TabIndex        =   20
         Text            =   "1/1/2001"
         ToolTipText     =   "DD/MM/YYYY"
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox txtCaseRefNo 
         BackColor       =   &H80000000&
         DataField       =   "First_Name"
         DataSource      =   "datPerInfo"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox txtDoctorInCharge 
         Height          =   285
         Left            =   240
         TabIndex        =   22
         Text            =   "Dr. Mekufon"
         ToolTipText     =   "DD/MM/YYYY"
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label Label11 
         Caption         =   "Doctor's Diagnosis"
         Height          =   255
         Left            =   2640
         TabIndex        =   23
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Admission Date (dd/mm/yyyy):"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         ToolTipText     =   "DD/MM/YYYY"
         Top             =   1200
         Width           =   2535
      End
      Begin VB.Label Label9 
         Caption         =   "Case Reference #:"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Doctor-in-Charge:"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         ToolTipText     =   "DD/MM/YYYY"
         Top             =   1800
         Width           =   1335
      End
   End
   Begin VB.Frame fraPersonal 
      Caption         =   "&Patient Identification Information"
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
      Height          =   1935
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   6135
      Begin VB.TextBox txtSex 
         BackColor       =   &H80000000&
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
         Left            =   4200
         TabIndex        =   11
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox txtHospNo 
         BackColor       =   &H80000000&
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
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txtOccupation 
         BackColor       =   &H80000000&
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
         Left            =   4200
         TabIndex        =   15
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox txtDOB 
         BackColor       =   &H80000000&
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
         Left            =   1320
         TabIndex        =   9
         ToolTipText     =   "DD/MM/YYYY"
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox txtFName 
         BackColor       =   &H80000000&
         DataField       =   "First_Name"
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
         Left            =   4200
         TabIndex        =   7
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtSName 
         BackColor       =   &H80000000&
         DataField       =   "Surname"
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
         Left            =   960
         TabIndex        =   5
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox txtStateOfOrigin 
         BackColor       =   &H80000000&
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
         Left            =   1440
         TabIndex        =   13
         Top             =   1440
         Width           =   1575
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
         TabIndex        =   2
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Occupation:"
         Height          =   255
         Left            =   3240
         TabIndex        =   14
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Sex:"
         Height          =   255
         Left            =   3720
         TabIndex        =   10
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Date Of Birth:"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         ToolTipText     =   "DD/MM/YYYY"
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblFName 
         Caption         =   "First Name:"
         Height          =   255
         Left            =   3240
         TabIndex        =   6
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblSName 
         Caption         =   "Surname:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label19 
         Caption         =   "State Of Origin:"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1440
         Width           =   1095
      End
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
      Left            =   420
      TabIndex        =   27
      Top             =   600
      Width           =   6015
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "NEW PATIENT"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   420
      TabIndex        =   0
      Top             =   120
      Width           =   6015
   End
End
Attribute VB_Name = "frmAdmission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAbortAdmission_Click()
 If MsgBox("Are you sure you want to abort the current patient admission?" & vbCrLf & "(NOTE: You will loose all information you have entered.)", vbCritical + vbYesNo) = vbYes Then
    Unload Me
 End If
End Sub

Private Sub cmdConfirmAdmission_Click()
'On Error GoTo errHnd
    With frmNewReg.datPerInfo.Recordset
        'PERSONAL INFO
        .Fields("Hosp_No") = frmNewReg.thisNewPatient.HospNo
        .Fields("SName") = frmNewReg.thisNewPatient.SName
        .Fields("FName") = frmNewReg.thisNewPatient.FName
        .Fields("Sex") = frmNewReg.thisNewPatient.Sex
        .Fields("Home_Add") = frmNewReg.thisNewPatient.HomeAdd
        .Fields("State_of_Origin") = frmNewReg.thisNewPatient.StateOfOrigin
        .Fields("Occupation") = frmNewReg.thisNewPatient.Occupation
        'NEXT OF KIN'S INFO
        .Fields("Name_of_NoK") = frmNewReg.thisNewPatient.NameNoK
        .Fields("Relationship_to_NoK") = frmNewReg.thisNewPatient.RelaNok
        .Fields("Add_of_NoK") = frmNewReg.thisNewPatient.AddNok
        'SPONSOR'S INFO
        .Fields("Name_of_Sponsor") = frmNewReg.thisNewPatient.SponsorName
        .Fields("Add_of_Sponsor") = frmNewReg.thisNewPatient.SponsorAdd
    End With
    
    With frmNewReg.datLabInfo.Recordset
        'LABORATORY INFO
        .Fields("Hosp_No") = frmNewReg.thisNewPatient.HospNo
        .Fields("Blood_Group") = frmNewReg.thisNewPatient.BloodGrp
        .Fields("RhFactor") = frmNewReg.thisNewPatient.RHFactor
        .Fields("Allergy") = frmNewReg.thisNewPatient.Allergy
    End With
    
 With Me.datHospHist.Recordset
    .Fields("Hosp_No") = Val(txtHospNo)
    .Fields("Admission_Status") = "IN"
    .Fields("Date_of_Admission") = txtDateOfAdmission
    .Fields("Name_of_Doctor") = txtDoctorInCharge
    .Fields("Doctors_Diagnosis") = txtDoctorsDiag
 End With
 
 'update d records into the db
 frmNewReg.datPerInfo.Recordset.Update
 frmNewReg.datLabInfo.Recordset.Update
 Me.datHospHist.Recordset.Update
 
 MsgBox "New patient has been REGISTERED and ADMITTED into care.", vbInformation, "Success"

 Unload Me
 Exit Sub
errHnd:
 Debug.Print Err.Number; "   "; Err.Description
 MsgBox "An error has occured.", vbInformation, "Unhandled error!"
 Resume Next
End Sub

Private Sub Form_Load()
 datHospHist.DatabaseName = App.Path & "\IHMS_97.mdb"
 datHospHist.RecordSource = "Patient_Hospital_History"
 datHospHist.Refresh
 datHospHist.Recordset.AddNew
 
 'Display information that's already been collected.
 With frmNewReg.thisNewPatient
    txtHospNo = .HospNo
    txtSName = .SName
    txtFName = .FName
    txtDOB = .DoB
    txtSex = .Sex
    txtStateOfOrigin = .StateOfOrigin
    txtOccupation = .Occupation
    txtCaseRefNo = datHospHist.Recordset.Fields("Case_Ref_No")
 End With

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 If UnloadMode = 1 Then Exit Sub
 If MsgBox("Are you sure you want to abort the current patient admission?" & vbCrLf & "(NOTE: You will loose all information you have entered.)", vbCritical + vbYesNo) = vbYes Then
    Cancel = 0
 Else
    'User selected NO button, and does not wish to abort the admission.
    Cancel = 1
 End If
End Sub
