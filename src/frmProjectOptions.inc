'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2022 Paul Squires, PlanetSquires Software
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT any WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS for A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.

#include once "frmProjectOptions.bi"
#include once "frmBuildConfig.bi"


' ========================================================================================
' Display the abbreviated Font name and point size in the Label
' ========================================================================================
function SetDefaultFontLabel( byval hParent as HWND ) as long
   dim as CWSTR wszPropValue = gApp.ProjectDefaultFont
   dim as CWSTR wszFont = AfxStrParse(wszPropValue, 1, ",") & ", " & _ 
                          AfxStrParse(wszPropValue, 2, ",") & "pt"
   AfxSetWindowText( GetDlgItem(hParent, IDC_FRMPROJECTOPTIONS_LBLDEFAULTFONT), wszFont )
   function = 0
end function

' ========================================================================================
' Show the popup font selection dialog for the Project Default Control Font
' ========================================================================================
function ChooseProjectDefaultFont( byval hParent as HWND ) as long
   
   dim cf as CHOOSEFONT 
   dim lf as LOGFONT
   cf.lStructSize = sizeof(cf)
   cf.hwndOwner = hParent
   cf.lpLogFont = @lf
   cf.Flags = CF_SCREENFONTS or CF_EFFECTS or CF_INITTOLOGFONTSTRUCT

   lf = SetLogFontFromPropValue(gApp.ProjectDefaultFont)

   EnableWindow(hParent, false)
   if ChooseFont(@cf) then
      gApp.ProjectDefaultFont = SetPropValueFromLogFont(*cf.lpLogFont) 
      SetDefaultFontLabel( hParent )
   END IF
   EnableWindow(hParent, true)
   SetActiveWindow(hParent)

   function = 0
end function

' ========================================================================================
' Create a generic resource file and manifest for the Project
' ========================================================================================
function frmProjectOptions_CreateResourceManifest() as Long

   dim as CWSTR wszProjectPath      = AfxStrPathname("PATH", gApp.ProjectFilename)
   dim as CWSTR wszManifest         = wszProjectPath + "manifest.xml"
   dim as CWSTR wszResource         = wszProjectPath + "resource.rc"
   dim as CWSTR wszManifestTemplate = AfxGetExePathName + "Settings\WinFBE_manifest.xml"
   dim as CWSTR wszResourceTemplate = AfxGetExePathName + "Settings\WinFBE_resource.rc" 
   
   dim as CWSTR wszFilesMissing
   
   if AfxFileExists(wszManifest) = false then 
      if AfxFileExists(wszManifestTemplate) = false then 
         wszFilesMissing = wszFilesMissing & _
                           space(5) & "- " & AfxStrPathname( "NAMEX", wszManifestTemplate ) & vbcrlf
      else
         AfxCopyFile( wszManifestTemplate, wszManifest, true )
      end if
   end if
   
   if AfxFileExists(wszResource) = false then 
      if AfxFileExists(wszResourceTemplate) = false then 
         wszFilesMissing = wszFilesMissing & _
                           space(5) & "- " & AfxStrPathname( "NAMEX", wszResourceTemplate ) & vbcrlf
      else
         AfxCopyFile( wszResourceTemplate, wszResource, true )
      end if   
   end if
   
   if len(wszFilesMissing) then
      MessageBox( 0, _
                  "The following files(s) are missing from the WinFBE Settings folder. " + _
                  "Please reinstall the WinFBE application." & vbcrlf & vbcrlf & _ 
                  wszFilesMissing, _
                  "Error", _
                  MB_OK Or MB_ICONWARNING Or MB_DEFBUTTON1 Or MB_APPLMODAL )
      exit function
   end if
   
   
   ' Add the resource file to the project if it does not already exist in the project
   if gApp.GetDocumentPtrByFilename(wszResource) = 0 then
      dim pDoc as clsDocument ptr
      pDoc = frmMain_OpenFileSafely(HWND_FRMMAIN, _
                              False, _    ' bIsNewFile
                              False, _    ' bIsTemplate
                              false, _    ' bShowInTab
                              false, _    ' bIsInclude
                              wszResource, _  ' pwszName
                              0 )  ' pDocIn
      gApp.ProjectSetFileType(pDoc, FILETYPE_RESOURCE)
   END IF
   
   function = 0
end function
   

' ========================================================================================
' Save all options for the Project
' ========================================================================================
function frmProjectOptions_SaveProjectOptions( ByVal HWnd As HWnd ) As BOOLEAN    
   
   Dim wText As WString * MAX_PATH  
   
   ' Determine if the project path was specified
   wText = AfxGetWindowText(GetDlgItem(HWnd, IDC_FRMPROJECTOPTIONS_TXTPROJECTPATH))
   If Len(wText) = 0 Then
      MessageBox( HWnd, L(217,"Invalid project path specified."), L(201,"Error"), _
                        MB_OK Or MB_ICONINFORMATION Or MB_DEFBUTTON1 )
      return False
   End If

   if gApp.IsNewProjectFlag then 
      ' Need to close any open files or project at this point before loading
      ' and displaying the new project.
      if gApp.IsProjectActive then
         if OnCommand_ProjectClose(HWnd) = false then exit function
      else
         if OnCommand_FileClose(HWnd, EFC_CLOSEALL) = false then exit function
      end if
   end if

   ' Load the new or modified data into the Project variables
   gApp.IsProjectActive = true
   gApp.ProjectFilename = wText
   gApp.ProjectName = AfxStrPathname( "NAMEX", gApp.ProjectFilename )

   gApp.ProjectOther32     = AfxGetWindowText(GetDlgItem(HWnd, IDC_FRMPROJECTOPTIONS_TXTOPTIONS32))
   gApp.ProjectOther64     = AfxGetWindowText(GetDlgItem(HWnd, IDC_FRMPROJECTOPTIONS_TXTOPTIONS64))
   gApp.ProjectManifest    = Button_GetCheck(GetDlgItem(HWnd, IDC_FRMPROJECTOPTIONS_CHKMANIFEST))
   gApp.ProjectBuild       = frmBuildConfig_GetSelectedBuildGUID()
   gApp.ProjectNotes       = ""
   gApp.ProjectCommandLine = ""
   
 
   ' Create the resource file and manifest if applicable
   if gApp.ProjectManifest THEN frmProjectOptions_CreateResourceManifest()

   ' Save the actual Project data to disk.
   gConfig.ProjectSaveToFile()   

   if Button_GetCheck( GetDlgItem(Hwnd, IDC_FRMPROJECTOPTIONS_OPTNONE) ) THEN
      gApp.NewProjectTemplateType = IDC_FRMPROJECTOPTIONS_OPTNONE
   elseif Button_GetCheck( GetDlgItem(Hwnd, IDC_FRMPROJECTOPTIONS_OPTBLANK) ) THEN
      gApp.NewProjectTemplateType = IDC_FRMPROJECTOPTIONS_OPTBLANK
   elseif Button_GetCheck( GetDlgItem(Hwnd, IDC_FRMPROJECTOPTIONS_OPTVD) ) THEN
      gApp.NewProjectTemplateType = IDC_FRMPROJECTOPTIONS_OPTVD
   elseif Button_GetCheck( GetDlgItem(Hwnd, IDC_FRMPROJECTOPTIONS_OPTCONSOLE) ) THEN
      gApp.NewProjectTemplateType = IDC_FRMPROJECTOPTIONS_OPTCONSOLE
   elseif Button_GetCheck( GetDlgItem(Hwnd, IDC_FRMPROJECTOPTIONS_OPTDLL) ) THEN
      gApp.NewProjectTemplateType = IDC_FRMPROJECTOPTIONS_OPTDLL
   elseif Button_GetCheck( GetDlgItem(Hwnd, IDC_FRMPROJECTOPTIONS_OPTSTATIC) ) THEN
      gApp.NewProjectTemplateType = IDC_FRMPROJECTOPTIONS_OPTSTATIC
   end if

   function = true   
End Function



' ========================================================================================
' Process WM_COMMAND message for window/dialog: frmProjectOptions
' ========================================================================================
function frmProjectOptions_OnCommand( _
         ByVal HWnd As HWnd, _
         ByVal id As Long, _
         ByVal hwndCtl As HWnd, _
         ByVal codeNotify As UINT _
         ) As LRESULT

   Select Case id
      Case IDC_FRMPROJECTOPTIONS_CMDSELECT
         If codeNotify = BN_CLICKED Then
            dim as CWSTR wszProjectName = ""
            if gApp.IsNewProjectFlag = false then wszProjectName = gApp.ProjectName
            Dim pwszName As WString Ptr = AfxIFileSaveDialog(HWnd, wszProjectName, @wstr("wfbe"), IDC_FRMPROJECTOPTIONS_CMDSELECT)
            If pwszName Then
               AfxSetWindowText( GetDlgItem(HWnd, IDC_FRMPROJECTOPTIONS_TXTPROJECTPATH), pwszName )
               CoTaskMemFree pwszName
            End If
            Exit Function
         End If
         
      Case IDC_FRMPROJECTOPTIONS_CMDDEFAULTFONT
         If codeNotify = BN_CLICKED Then
            ChooseProjectDefaultFont( HWnd )
            Exit Function
         End If

      Case IDOK
         If codeNotify = BN_CLICKED Then
            If frmProjectOptions_SaveProjectOptions(HWnd) Then
               SendMessage HWnd, WM_CLOSE, 0, 0
            End If
            Exit Function
         End If
      
      Case IDCANCEL
         If codeNotify = BN_CLICKED Then
            gApp.IsNewProjectFlag = false
            SendMessage HWnd, WM_CLOSE, 0, 0
            exit function
         End If

   End Select

   Function = 0
End Function


' ========================================================================================
' Process WM_CREATE message for window/dialog: frmProjectOptions
' ========================================================================================
function frmProjectOptions_OnCreate( _
         ByVal HWnd As HWnd, _
         ByVal lpCreateStructPtr As LPCREATESTRUCT _
         ) As BOOLEAN

   ' This is a modal popup window so disable the parent window
   DisableAllModeless()

   '  Message cracker macro expects a True to be returned for a successful
   '  OnCreate handler even though returning -1 from a standard WM_CREATE
   '  call would stop creating the window. This is just one of those Windows
   '  inconsistencies.
   Return True
End Function


' ========================================================================================
' Process WM_CLOSE message for window/dialog: frmProjectOptions
' ========================================================================================
function frmProjectOptions_OnClose( byval HWnd As HWnd ) As LRESULT
   ' Enables parent window keeping parent's zorder
   EnableAllModeless()
   DestroyWindow HWnd
   Function = 0
end function


' ========================================================================================
' Process WM_DESTROY message for window/dialog: frmProjectOptions
' ========================================================================================
function frmProjectOptions_OnDestroy( byval HWnd As HWnd ) As LRESULT
   PostQuitMessage(0)
   Function = 0
End Function


' ========================================================================================
' frmProjectOptions Window procedure
' ========================================================================================
function frmProjectOptions_WndProc( _
         ByVal HWnd   As HWnd, _
         ByVal uMsg   As UINT, _
         ByVal wParam As WPARAM, _
         ByVal lParam As LPARAM _
         ) As LRESULT

   Select Case uMsg
      HANDLE_MSG (HWnd, WM_CREATE,   frmProjectOptions_OnCreate)
      HANDLE_MSG (HWnd, WM_CLOSE,    frmProjectOptions_OnClose)
      HANDLE_MSG (HWnd, WM_DESTROY,  frmProjectOptions_OnDestroy)
      HANDLE_MSG (HWnd, WM_COMMAND,  frmProjectOptions_OnCommand)
   End Select

   Function = DefWindowProc(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' frmProjectOptions_Show
' ========================================================================================
function frmProjectOptions_Show( _
         ByVal hWndParent As HWnd, _
         byval IsNewProject as boolean _
         ) As LRESULT

   Dim wszTitle As WString * MAX_PATH
   dim as long nLeft, nTop, nWidth, nHeight, nMidPoint
   
   '  Create the main window and child controls
   Dim pWindow As CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowOwnerPtr(hwndParent)->DPI

   Dim As HWnd hForm = _
   pWindow->Create( hWndParent, "", @frmProjectOptions_WndProc, 0, 0, 0, 0, _
        WS_POPUP Or WS_CAPTION Or WS_SYSMENU Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, _
        WS_EX_DLGMODALFRAME Or WS_EX_CONTROLPARENT Or WS_EX_LEFT)
   pWindow->SetClientSize(600, 342)
   pWindow->Center(pWindow->hWindow, hWndParent)
      
   dim as RECT rcClient
   GetClientRect( hForm, @rcClient )
   nMidPoint = pWindow->UnScaleX((rcClient.Right / 2))
   
   pWindow->AddControl("LABEL", , IDC_FRMPROJECTOPTIONS_LABEL1, L(240,"Project Path"), 26, 15, 235, 18, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT, _
        WS_EX_LEFT Or WS_EX_LTRREADING)                                              
   Dim As HWnd hProjectPath = _
   pWindow->AddControl("TEXTBOX", , IDC_FRMPROJECTOPTIONS_TXTPROJECTPATH, "", 26, 35, 508, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)     
   pWindow->AddControl("BUTTON", , IDC_FRMPROJECTOPTIONS_CMDSELECT, "...", 544, 33, 37, 24, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)

   nTop    = iIf(IsNewProject, 75, 95) 
   nWidth  = iIf(IsNewProject, 280, 555)
   nHeight = iIf(IsNewProject, 40, 20)
   pWindow->AddControl("LABEL", , IDC_FRMPROJECTOPTIONS_LABEL5, L(283,"Specify compiler options in addition to the selected build configuration") & ":", 26, nTop, nWidth, nHeight, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT, _
        WS_EX_LEFT Or WS_EX_LTRREADING)

   pWindow->AddControl("LABEL", , IDC_FRMPROJECTOPTIONS_LABEL2, L(241,"Other Options (32-bit compiler)"), 26, 121, nWidth, 18, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   nWidth = iIf(IsNewProject, 280, 555) 
   Dim As HWnd hProjectOther32 = _
   pWindow->AddControl("TEXTBOX", , IDC_FRMPROJECTOPTIONS_TXTOPTIONS32, "", 26, 141, nWidth, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   
   pWindow->AddControl("LABEL", , IDC_FRMPROJECTOPTIONS_LABEL4, L(242,"Other Options (64-bit compiler)"), 26, 170, nWidth, 18, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   nWidth = iIf(IsNewProject, 280, 555) 
   Dim As HWnd hProjectOther64 = _
   pWindow->AddControl("TEXTBOX", , IDC_FRMPROJECTOPTIONS_TXTOPTIONS64, "", 26, 190, nWidth, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)

   nWidth = iIf(IsNewProject, 280, 555) 
   dim as hwnd hManifest = _
   pWindow->AddControl("CHECKBOX", , IDC_FRMPROJECTOPTIONS_CHKMANIFEST, L(308,"Create resource file and manifest"), 26, 222, nWidth, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)

   nWidth = iIf(IsNewProject, 280, 555) 
   pWindow->AddControl("LABEL", , IDC_FRMPROJECTOPTIONS_LABEL8, L(448,"Default Control Font") & ":", 26, 254, 130, 18, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   dim as hwnd hLblDefaultFont = _
   pWindow->AddControl("LABEL", , IDC_FRMPROJECTOPTIONS_LBLDEFAULTFONT, "", 160, 254, 120, 18, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   SetDefaultFontLabel( hForm )
   dim as hwnd hCmdDefaultFont = _
   pWindow->AddControl("BUTTON", , IDC_FRMPROJECTOPTIONS_CMDDEFAULTFONT, wszMoreActions, 282, 254, 24, 18, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)

   If IsNewProject Then
      nMidPoint = nMidPoint + 40
      pWindow->AddControl("LABEL", , IDC_FRMPROJECTOPTIONS_LABEL6, "", nMidPoint, 75, 2, 200, _
           WS_CHILD Or WS_VISIBLE Or SS_LEFT Or SS_SUNKEN, _
           WS_EX_LEFT Or WS_EX_LTRREADING)
      nLeft  = nMidPoint + 20     
      nTop   = 75
      nWidth = 250
      pWindow->AddControl("LABEL", , IDC_FRMPROJECTOPTIONS_LABEL7, L(176,"Templates") & ":", nLeft, nTop, nWidth, 20, _
           WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT, _
           WS_EX_LEFT Or WS_EX_LTRREADING)
      nLeft  = nLeft + 20
      nTop   = nTop + 22
      nWidth = 260
      pWindow->AddControl("RADIOBUTTON", , IDC_FRMPROJECTOPTIONS_OPTNONE, L(388,"None"), nLeft, nTop, nWidth, 20, _
           WS_VISIBLE OR WS_TABSTOP OR BS_AUTORADIOBUTTON OR BS_LEFT OR BS_VCENTER or WS_GROUP, _
           WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
      nTop = nTop + 22
      pWindow->AddControl("RADIOBUTTON", , IDC_FRMPROJECTOPTIONS_OPTBLANK, L(389,"Blank Document"), nLeft, nTop, nWidth, 20, _
           WS_VISIBLE OR WS_TABSTOP OR BS_AUTORADIOBUTTON OR BS_LEFT OR BS_VCENTER, _
           WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
      nTop = nTop + 22
      pWindow->AddControl("RADIOBUTTON", , IDC_FRMPROJECTOPTIONS_OPTVD, L(390,"Visual Designer"), nLeft, nTop, nWidth, 20, _
           WS_VISIBLE OR WS_TABSTOP OR BS_AUTORADIOBUTTON OR BS_LEFT OR BS_VCENTER, _
           WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
      nTop = nTop + 22
      pWindow->AddControl("RADIOBUTTON", , IDC_FRMPROJECTOPTIONS_OPTCONSOLE, L(391,"Console"), nLeft, nTop, nWidth, 20, _
           WS_VISIBLE OR WS_TABSTOP OR BS_AUTORADIOBUTTON OR BS_LEFT OR BS_VCENTER, _
           WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
      nTop = nTop + 22
      pWindow->AddControl("RADIOBUTTON", , IDC_FRMPROJECTOPTIONS_OPTDLL, L(392,"Windows DLL"), nLeft, nTop, nWidth, 20, _
           WS_VISIBLE OR WS_TABSTOP OR BS_AUTORADIOBUTTON OR BS_LEFT OR BS_VCENTER, _
           WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
      nTop = nTop + 22
      pWindow->AddControl("RADIOBUTTON", , IDC_FRMPROJECTOPTIONS_OPTSTATIC, L(393,"Static Library"), nLeft, nTop, nWidth, 20, _
           WS_VISIBLE OR WS_TABSTOP OR BS_AUTORADIOBUTTON OR BS_LEFT OR BS_VCENTER, _
           WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
      CheckRadioButton( hForm, IDC_FRMPROJECTOPTIONS_OPTNONE, _
                               IDC_FRMPROJECTOPTIONS_OPTSTATIC, _
                               IDC_FRMPROJECTOPTIONS_OPTNONE )
   end if
   
   
   pWindow->AddControl("LABEL", , IDC_FRMPROJECTOPTIONS_LABEL3, "", 10, 282, 572, 2, _
        WS_CHILD Or WS_VISIBLE Or SS_LEFT Or SS_NOTIFY Or SS_SUNKEN, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("BUTTON", , IDOK, L(0,"&OK"), 423, 297, 74, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_DEFPUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("BUTTON", , IDCANCEL, L(1,"&Cancel"), 508, 297, 74, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)

   wszTitle = L(214,"Project Options")

   ' gApp.IsNewProjectFlag is only used within frmProjectOptions in order to
   ' keep the state when OK is pressed.
   gApp.IsNewProjectFlag = IsNewProject
   If IsNewProject Then
      ' Set the default values for a new project
      wszTitle = wszTitle & " - " & L(215, "New Project")
      Button_SetCheck( hManifest, true)
   Else
      wszTitle = wszTitle + " - " + gApp.ProjectName
      AfxSetWindowText( hProjectPath, gApp.ProjectFilename )
      AfxSetWindowText( hProjectOther32, gApp.ProjectOther32 )
      AfxSetWindowText( hProjectOther64, gApp.ProjectOther64 )
      Button_SetCheck( hManifest, gApp.ProjectManifest)
   End If
   AfxSetWindowText( hForm, wszTitle )

   SetFocus GetDlgItem(hForm, IDOK)
   
   ' Process Windows messages(modal)
   Function = pWindow->DoEvents(SW_SHOW)
   
   ' Delete the frmProjectOptions CWindow class manually allocated memory 
   Delete pWindow

End Function
