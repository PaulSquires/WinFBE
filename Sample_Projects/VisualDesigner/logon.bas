
#include once "WinFormsX\WinFormsX.bi"

#Include Once "frmPopup.inc"
#Include Once "frmLogon.inc"



''
''
Function frmPopup_Load( ByRef sender As wfxForm, ByRef e As EventArgs) As LRESULT
   ' Have to put this Load event handler here after the two #Includes because
   ' the frmPopup form references the frmLogon form. If this code was placed in
   ' the frmPopup.inc source file then the frmLogon TYPE would not yet be
   ' defined but it is referenced.
   
   frmPopup.lblType.Text     = "Access Type: " & frmLogon.lstAccess.SelectedItem.Text
   frmPopup.lblUserName.Text = "Username: "    & frmLogon.txtUserName.Text
   frmPopup.lblPassword.Text = "Password: "    & frmLogon.txtPassword.Text

   Function = 0
End Function


''
''  Show the main form for the application.
''
Application.Run(frmLogon)


