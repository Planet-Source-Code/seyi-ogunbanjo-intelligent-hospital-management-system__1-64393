VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_HospNo As Integer
Dim m_SName As String
Dim m_FName As String
Dim m_DoB As String
Dim m_Sex As String

Dim m_HomeAdd As String
Dim m_StateOfOrigin As String
Dim m_Occupation As String

Dim m_NameNoK As String
Dim m_RelaNok As String
Dim m_AddNok As String

Dim m_SponsorName As String
Dim m_SponsorAdd As String

Dim m_LabRefNo As Integer
Dim m_BloodGrp As String
Dim m_RHFactor As String
Dim m_Allergy As String

Dim m_AdmissionDate As String
Dim m_DischargeDate As String

Dim m_CaseRefNo As Integer
Dim m_AdmissionStatus As String
Dim m_DocName As String
Dim m_Diagnosis As String
Dim m_DateDischarged As String
Dim m_DischargeStatus As String

Public Property Get HospNo() As Integer
 HospNo = m_HospNo
End Property

Public Property Let HospNo(ByVal vHospNo As Integer)
 m_HospNo = vHospNo
End Property

Public Property Get SName() As String
 SName = m_SName
End Property

Public Property Let SName(ByVal vSName As String)
 m_SName = vSName
End Property

Public Property Get FName() As String
 FName = m_FName
End Property

Public Property Let FName(ByVal vFName As String)
 m_FName = vFName
End Property

Public Property Get DoB() As String
 DoB = m_DoB
End Property

Public Property Let DoB(ByVal vDoB As String)
 m_DoB = vDoB
End Property

Public Property Get Sex() As String
 Sex = m_Sex
End Property

Public Property Let Sex(ByVal vSex As String)
 m_Sex = vSex
End Property

Public Property Get HomeAdd() As String
 HomeAdd = m_HomeAdd
End Property

Public Property Let HomeAdd(ByVal vHomeAdd As String)
 m_HomeAdd = vHomeAdd
End Property

Public Property Get StateOfOrigin() As String
 StateOfOrigin = m_StateOfOrigin
End Property

Public Property Let StateOfOrigin(ByVal vStateOfOrigin As String)
 m_StateOfOrigin = vStateOfOrigin
End Property

Public Property Get Occupation() As String
 Occupation = m_Occupation
End Property

Public Property Let Occupation(ByVal vOccupation As String)
 m_Occupation = vOccupation
End Property

Public Property Get NameNoK() As String
 NameNoK = m_NameNoK
End Property

Public Property Let NameNoK(ByVal vNameNoK As String)
m_NameNoK = vNameNoK
End Property

Public Property Get RelaNok() As String
 RelaNok = m_RelaNok
End Property

Public Property Let RelaNok(ByVal vRelaNok As String)
 m_RelaNok = vRelaNok
End Property

Public Property Get AddNok() As String
 AddNok = m_AddNok
End Property

Public Property Let AddNok(ByVal vAddNok As String)
 m_AddNok = vAddNok
End Property

Public Property Get SponsorName() As String
 SponsorName = m_SponsorName
End Property

Public Property Let SponsorName(ByVal vSponsorName As String)
 m_SponsorName = vSponsorName
End Property

Public Property Get SponsorAdd() As String
 SponsorAdd = m_SponsorAdd
End Property

Public Property Let SponsorAdd(ByVal vSponsorAdd As String)
 m_SponsorAdd = vSponsorAdd
End Property

Public Property Get LabRefNo() As Integer
 LabRefNo = m_LabRefNo
End Property

Public Property Let LabRefNo(ByVal vLabRefNo As Integer)
 m_LabRefNo = vLabRefNo
End Property

Public Property Get BloodGrp() As String
 BloodGrp = m_BloodGrp
End Property

Public Property Let BloodGrp(ByVal vBloodGrp As String)
 m_BloodGrp = vBloodGrp
End Property

Public Property Get RHFactor() As String
 RHFactor = m_RHFactor
End Property

Public Property Let RHFactor(ByVal vRHFactor As String)
 m_RHFactor = vRHFactor
End Property

Public Property Get Allergy() As String
 Allergy = m_Allergy
End Property

Public Property Let Allergy(ByVal vAllergy As String)
 m_Allergy = vAllergy
End Property

Public Property Get AdmissionDate() As String
 AdmissionDate = m_AdmissionDate
End Property

Public Property Let AdmissionDate(ByVal vAdmissionDate As String)
 m_AdmissionDate = vAdmissionDate
End Property

Public Property Get DischargeDate() As String
 DischargeDate = m_DischargeDate
End Property

Public Property Let DischargeDate(ByVal vDischargeDate As String)
 m_DischargeDate = vDischargeDate
End Property

Public Property Get CaseRefNo() As Integer
 CaseRefNo = m_CaseRefNo
End Property

Public Property Let CaseRefNo(ByVal vCaseRefNo As Integer)
 m_CaseRefNo = vCaseRefNo
End Property

Public Property Get AdmissionStatus() As String
 AdmissionStatus = m_AdmissionStatus
End Property

Public Property Let AdmissionStatus(ByVal vAdmissionStatus As String)
 m_AdmissionStatus = vAdmissionStatus
End Property

Public Property Get DocName() As String
 DocName = m_DocName
End Property

Public Property Let DocName(ByVal vDocName As String)
 m_DocName = vDocName
End Property

Public Property Get Diagnosis() As String
 Diagnosis = m_Diagnosis
End Property

Public Property Let Diagnosis(ByVal vDiagnosis As String)
 m_Diagnosis = vDiagnosis
End Property

Public Property Get DateDischarged() As String
 DateDischarged = m_DateDischarged
End Property

Public Property Let DateDischarged(ByVal vDateDischarged As String)
 m_DateDischarged = vDateDischarged
End Property

Public Property Get DischargeStatus() As String
 DischargeStatus = m_DischargeStatus
End Property

Public Property Let DischargeStatus(ByVal vDischargeStatus As String)
 m_DischargeStatus = vDischargeStatus
End Property

