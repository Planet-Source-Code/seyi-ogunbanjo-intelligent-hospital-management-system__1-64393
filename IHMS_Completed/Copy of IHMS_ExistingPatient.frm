VERSION 5.00
Begin VB.Form frmOldPatient 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View Existing Patient"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10935
   Icon            =   "IHMS_ExistingPatient.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   10935
   Begin VB.Data datPerInfo 
      Caption         =   "Personal Info"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5040
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
      Height          =   345
      Left            =   4800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5040
      Visible         =   0   'False
      Width           =   2580
   End
   Begin VB.CommandButton cmdCloseFile 
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
      Left            =   240
      TabIndex        =   30
      Top             =   4560
      Width           =   10215
   End
   Begin VB.CommandButton cmdViewContact 
      Caption         =   "&Shred Patient's File from Database"
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
      Begin VB.TextBox txtAllergy 
         BackColor       =   &H80000000&
         Height          =   285
         Left            =   1680
         TabIndex        =   23
         ToolTipText     =   "DD/MM/YYYY"
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox txtBloodGroup 
         BackColor       =   &H80000000&
         Height          =   285
         Left            =   1680
         TabIndex        =   21
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox txtLabRefNo 
         BackColor       =   &H80000000&
         Height          =   285
         Left            =   1680
         TabIndex        =   19
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtGenotype 
         BackColor       =   &H80000000&
         Height          =   285
         Left            =   1680
         TabIndex        =   25
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label lblLabAllergy 
         Alignment       =   1  'Right Justify
         Caption         =   "Allergy:"
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
         Caption         =   "Genotype:"
         Height          =   255
         Left            =   480
         TabIndex        =   24
         Top             =   1680
         Width           =   1095
      End
   End
   Begin VB.TextBox txtHospNo 
      BackColor       =   &H80000000&
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
         BackColor       =   &H80000000&
         DataField       =   "First_Name"
         DataSource      =   "datPerInfo"
         Height          =   285
         Left            =   4560
         TabIndex        =   10
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtStateOfOrigin 
         BackColor       =   &H80000000&
         DataSource      =   "datPerInfo"
         Height          =   285
         Left            =   4560
         TabIndex        =   14
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox txtHomeAdd 
         BackColor       =   &H80000000&
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
         BackColor       =   &H80000000&
         DataField       =   "Surname"
         DataSource      =   "datPerInfo"
         Height          =   285
         Left            =   1560
         TabIndex        =   4
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtFName 
         BackColor       =   &H80000000&
         DataField       =   "First_Name"
         DataSource      =   "datPerInfo"
         Height          =   285
         Left            =   4560
         TabIndex        =   6
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtDOB 
         BackColor       =   &H80000000&
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
         BackColor       =   &H80000000&
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
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Patient# "
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
      TabIndex        =   31
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
 'MsgBox "THE PATIENT INFORMATION IN THE FORM IS CHECKED FOR ERRORS, AND THE USER IS NOTIFIED OF ANY ERRORS FOUND, LIKE OMMISION OF REQUIRED DATA OR WRONGLY ENTERED DATA. IF NO ERRORS ARE FOUND, THE PATIENT INFO IS WRITTEN TO THE DB AND A PATIENT ADMISSION FORM IS LOADED TO COMMENCE THE ADMISSION PROCESS."
 frmAdmission.Show 1
End Sub

Private Sub cmdCloseFile_Click()
 If MsgBox("Unload and close patient file?", vbYesNo + vbQuestion) = vbYes Then Unload Me
End Sub

Private Sub cmdDischarge_Click()
 frmDischarge.Show 1
 
End Sub

