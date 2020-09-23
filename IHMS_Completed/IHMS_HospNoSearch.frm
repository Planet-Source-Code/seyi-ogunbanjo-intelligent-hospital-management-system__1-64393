VERSION 5.00
Begin VB.Form frmSearchHospNo 
   Caption         =   "Form1"
   ClientHeight    =   6120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12195
   LinkTopic       =   "Form1"
   ScaleHeight     =   6120
   ScaleWidth      =   12195
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCloseFile 
      Caption         =   "&Close Patient File"
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
      TabIndex        =   42
      Top             =   5160
      Width           =   2895
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
      Left            =   9000
      TabIndex        =   41
      Top             =   4680
      Width           =   2895
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
      Left            =   9000
      TabIndex        =   40
      Top             =   4200
      Width           =   2895
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
      Height          =   2295
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   6495
      Begin VB.TextBox txtDoB 
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txtSex 
         Height          =   285
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "edo"
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtOccupation 
         Height          =   285
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "teacher"
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox txtFName 
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
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "harry"
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txtSName 
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
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "omoregie"
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txtHomeAdd 
         Height          =   765
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Text            =   "IHMS_H~1.frx":0000
         Top             =   1320
         Width           =   3015
      End
      Begin VB.TextBox txtStateOfOrigin 
         Height          =   285
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "edo"
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Occupation:"
         Height          =   255
         Left            =   3360
         TabIndex        =   16
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Sex:"
         Height          =   255
         Left            =   4080
         TabIndex        =   10
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Date Of Birth:"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         ToolTipText     =   "DD/MM/YYYY"
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblFName 
         Alignment       =   1  'Right Justify
         Caption         =   "First Name:"
         Height          =   255
         Left            =   3360
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblSName 
         Alignment       =   1  'Right Justify
         Caption         =   "Surname:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblHomeAdd 
         Caption         =   "Address:"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "State Of Origin:"
         Height          =   255
         Left            =   3360
         TabIndex        =   14
         Top             =   1080
         Width           =   1095
      End
   End
   Begin VB.TextBox txtHospNo 
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
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   720
      Width           =   1695
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
      TabIndex        =   18
      Top             =   1080
      Width           =   4935
      Begin VB.TextBox txtRelationship 
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "brother"
         Top             =   720
         Width           =   2535
      End
      Begin VB.TextBox txtKinName 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "imokwede igwe"
         Top             =   360
         Width           =   3015
      End
      Begin VB.TextBox txtKinAddress 
         Height          =   885
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   24
         Text            =   "IHMS_H~1.frx":0018
         Top             =   1320
         Width           =   4575
      End
      Begin VB.Label Label13 
         Caption         =   "Relationship to Next of Kin:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label11 
         Caption         =   "Address of Next of Kin:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label lblKinSName 
         Caption         =   "Name of Next of Kin:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   1575
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
      TabIndex        =   25
      Top             =   3600
      Width           =   4935
      Begin VB.TextBox txtAddOfSponsor 
         Height          =   885
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   29
         Text            =   "IHMS_H~1.frx":0030
         Top             =   960
         Width           =   4575
      End
      Begin VB.TextBox txtNameOfSponsor 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   27
         Text            =   "same as kin"
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label Label4 
         Caption         =   "Name of Sponsor:"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Address of Sponsor:"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   720
         Width           =   1695
      End
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
      TabIndex        =   30
      Top             =   3600
      Width           =   3615
      Begin VB.TextBox txtRHFactor 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox txtBloodGrp 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox txtLabRefNo 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtAllergy 
         Height          =   285
         Left            =   1680
         TabIndex        =   38
         Text            =   "n/a"
         ToolTipText     =   "DD/MM/YYYY"
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label lblLabGenotype 
         Alignment       =   1  'Right Justify
         Caption         =   "Rh Factor:"
         Height          =   255
         Left            =   480
         TabIndex        =   35
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblLabRefNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Lab. Ref. Number:"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblLabBloodGroup 
         Alignment       =   1  'Right Justify
         Caption         =   "Blood Group:"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblLabAllergy 
         Alignment       =   1  'Right Justify
         Caption         =   "Allergy:"
         Height          =   255
         Left            =   480
         TabIndex        =   37
         ToolTipText     =   "DD/MM/YYYY"
         Top             =   1560
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdAdmit 
      Caption         =   "&Admit Patient"
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
      Top             =   3720
      Width           =   2895
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
      TabIndex        =   1
      Top             =   720
      Width           =   1575
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
      Left            =   3540
      TabIndex        =   0
      Top             =   120
      Width           =   5175
   End
End
Attribute VB_Name = "frmSearchHospNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
 
 datPerInfo.DatabaseName = App.Path & "\IHMS_97.mdb"
 datPerInfo.RecordSource = "Patient_Personal_Info"
 datPerInfo.Refresh
  
 datLabInfo.DatabaseName = App.Path + "\IHMS_97.mdb"
 datLabInfo.RecordSource = "Patient_Lab_Info"
 datLabInfo.Refresh
 
 

End Sub
