Attribute VB_Name = "modGlobalVars"
Option Explicit
Public LoginSucceeded As Boolean    'Used to validate log-in procedure.
Public patientNumberX As Integer    'Stores hosp_no to be sought for in the db.
Public somePatient As CPatient  'Made public so it can be accessed by frmOldPatient.

Public Sub ConfigMenus()
 With frmMain
    If LoginSucceeded = True Then
        'menu items
        .mnuLogIn.Enabled = False
        .mnuSep1.Visible = True
        .mnuNewPatient.Visible = True
        .mnuOpen.Visible = True
        .mnuClose.Visible = True
        .mnuSep2.Visible = True
        .mnuLogOut.Enabled = True
        .mnuTools.Enabled = True
        .StatusBar1.SimpleText = "Logged in"
        
        'toolbar buttons
        .tbrMainToolbar.Buttons(1).Caption = "Log Out" 'login/logout button
        .tbrMainToolbar.Buttons(1).ToolTipText = "Log Out" 'login button
        .tbrMainToolbar.Buttons(2).Enabled = True  'new registration button
        .tbrMainToolbar.Buttons(3).Enabled = True  'open existing patient file
        .tbrMainToolbar.Buttons(4).Enabled = False 'admit/discharge button
        .tbrMainToolbar.Buttons(5).Enabled = False 'diagnose button
    Else
        'Close all open windows (all child forms)
        Unload frmNewReg
        Unload frmOldPatient

        'menu items
        .mnuLogIn.Enabled = True
        .mnuSep1.Visible = False
        .mnuNewPatient.Visible = False
        .mnuOpen.Visible = False
        .mnuClose.Visible = False
        .mnuSep2.Visible = False
        .mnuLogOut.Enabled = False
        .mnuTools.Enabled = False
        .StatusBar1.SimpleText = "Not Logged In"
        
        'toolbar buttons
        .tbrMainToolbar.Buttons(1).Caption = "Log In" 'loginout button
        .tbrMainToolbar.Buttons(1).ToolTipText = "Log In" 'login button
        .tbrMainToolbar.Buttons(2).Enabled = False 'new registration button
        .tbrMainToolbar.Buttons(3).Enabled = False 'open existing patient file
        .tbrMainToolbar.Buttons(4).Enabled = False 'admit/discharge button
        .tbrMainToolbar.Buttons(5).Enabled = False 'diagnose button
        
    End If
 End With
End Sub

Public Sub ClearRegForm()
 With frmNewReg
    '.txtHospNo.Text = ""
    .txtSName.Text = ""
    .txtFName.Text = ""
    .txtDOB.Text = ""
    .txtHomeAdd.Text = ""
    .txtStateOfOrigin.Text = ""
    .txtOccupation.Text = ""
    .txtNameOfSponsor.Text = ""
    .txtAddOfSponsor.Text = ""
    .txtKinName.Text = ""
    .txtRelationship.Text = ""
    .txtKinAddress.Text = ""
    .txtAllergy.Text = ""
    
'code snippet1
 End With
End Sub
