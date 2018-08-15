'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2018 Paul Squires, PlanetSquires Software
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


' ========================================================================================
' Create a generic resource file and manifest for the Project
' ========================================================================================
function CreateResourceManifest( byval idx as long ) as Long

   dim as CWSTR wszProjectPath = AfxStrPathname("PATH", gApp.Projects(idx).ProjectFilename)
   dim as CWSTR wszManifest = wszProjectPath + "manifest.xml"
   dim as CWSTR wszResource = wszProjectPath + "resource.rc"
   dim as CWSTR wszManifestTemplate = AfxGetExePathName + "Settings\WinFBE_manifest.xml"
   dim as CWSTR wszResourceTemplate = AfxGetExePathName + "Settings\WinFBE_resource.rc" 
   
   if AfxFileExists(wszManifest) = false THEN
      CopyFile( wszManifestTemplate, wszManifest, true )
   END IF
   
   if AfxFileExists(wszResource) = false THEN
      CopyFile( wszResourceTemplate, wszResource, true )
      ' Add the resource file to the project
      If AfxFileExists(wszResource) Then
         dim pDoc as clsDocument ptr
         pDoc = frmMain_OpenFileSafely(HWND_FRMMAIN, _
                                 False, _    ' bIsNewFile
                                 False, _    ' bIsTemplate
                                 FALSE, _    ' bShowInTab
                                 false, _    ' bIsInclude
                                 wszResource, _  ' pwszName
                                 0 )  ' pDocIn
         gApp.Projects(idx).ProjectSetFileType(pDoc, FILETYPE_RESOURCE)
         pDoc->ProjectIndex = idx
      end if
   END IF
   
   function = 0
end function
   

' ========================================================================================
' Save all options for the Project
' ========================================================================================
Function SaveProjectOptions( ByVal HWnd As HWnd ) As BOOLEAN    
   
   Dim wText As WString * MAX_PATH  
    
   Dim As HWnd hCtrl = GetDlgItem(HWnd, IDC_FRMPROJECTOPTIONS_TXTPROJECTPATH)
   
   dim as long idx = gApp.GetActiveProjectIndex
   
   gApp.Projects(idx).ProjectFilename = AfxGetWindowText(hCtrl)

   If Len(gApp.Projects(idx).ProjectFilename) = 0 Then
      MessageBoxW( HWnd, L(217,"Invalid project path specified."), L(201,"Error"), _
                        MB_OK Or MB_ICONINFORMATION Or MB_DEFBUTTON1 )
      SetFocus hCtrl
      Function = False: Exit Function
   End If

   gApp.Projects(idx).ProjectName = AfxStrPathname( "NAMEX", gApp.Projects(idx).ProjectFilename )

   hCtrl = GetDlgItem(HWnd, IDC_FRMPROJECTOPTIONS_TXTOPTIONS32)
   gApp.Projects(idx).ProjectOther32 = AfxGetWindowText(hCtrl)
   hCtrl = GetDlgItem(HWnd, IDC_FRMPROJECTOPTIONS_TXTOPTIONS64)
   gApp.Projects(idx).ProjectOther64 = AfxGetWindowText(hCtrl)
   hCtrl = GetDlgItem(HWnd, IDC_FRMPROJECTOPTIONS_CHKMANIFEST) 
   gApp.Projects(idx).ProjectManifest = Button_GetCheck(hCtrl)

   
   Return gConfig.ProjectSaveToFile()   

End Function



' ========================================================================================
' Process WM_CREATE message for window/dialog: frmProjectOptions
' ========================================================================================
Function frmProjectOptions_OnCreate(ByVal HWnd As HWnd, ByVal lpCreateStructPtr As LPCREATESTRUCT) As BOOLEAN

   ' Retrieve a reference to the CWindow class from the CREATESTRUCT structure
   'Dim pCreateStruct As CREATESTRUCT Ptr = lpCreateStructPtr
   'Dim pWindow As CWindow Ptr = Cast(CWindow Ptr, pCreateStruct->lpCreateParams)

   ' This is a modal popup window so disable the parent window
   EnableWindow GetParent(HWnd), False

   '  Message cracker macro expects a True to be returned for a successful
   '  OnCreate handler even though returning -1 from a standard WM_CREATE
   '  call would stop creating the window. This is just one of those Windows
   '  inconsistencies.
   Return True
End Function


' ========================================================================================
' Process WM_COMMAND message for window/dialog: frmProjectOptions
' ========================================================================================
Function frmProjectOptions_OnCommand(ByVal HWnd As HWnd, ByVal id As Long, ByVal hwndCtl As HWnd, ByVal codeNotify As UINT) As LRESULT

   dim as long idx = gApp.GetActiveProjectIndex
   
   Select Case id
      Case IDC_FRMPROJECTOPTIONS_CMDSELECT
         If codeNotify = BN_CLICKED Then
            Dim pwszName As WString Ptr = AfxIFileSaveDialog(HWnd, gApp.Projects(idx).ProjectName, WStr("wfbe"), IDC_FRMPROJECTOPTIONS_CMDSELECT)
            If pwszName Then
               SetWindowText( GetDlgItem(HWnd, IDC_FRMPROJECTOPTIONS_TXTPROJECTPATH), pwszName )
               CoTaskMemFree pwszName
            End If
            Exit Function
         End If
         
      Case IDOK
         If codeNotify = BN_CLICKED Then
            if gApp.IsNewProjectFlag then gProjectLoading = TRUE
            If SaveProjectOptions(HWnd) Then
               if gApp.IsNewProjectFlag then 
                  ' Update the most recently used project list 
                  UpdateMRUProjectList(gApp.Projects(idx).ProjectFilename)
                  gApp.IsNewProjectFlag = false
                  gApp.Projects(idx).InUse = FALSE   ' because OpenMRUProjectFile will assign one.
                  ' Open the new project so it displays in the Explorer treeview
                  OpenMRUProjectFile( HWND_FRMMAIN, 0, gApp.Projects(idx).ProjectFilename )
               end if
               ' Create the resource file and manifest if applicable
               if gApp.Projects(idx).ProjectManifest THEN
                  CreateResourceManifest(idx)
               END IF
               SendMessage( HWnd, WM_CLOSE, 0, 0 )
            End If
            gProjectLoading = FALSE
            Exit Function
         End If
      
      Case IDCANCEL
         If codeNotify = BN_CLICKED Then
            gApp.IsNewProjectFlag = false
            SendMessage( HWnd, WM_CLOSE, 0, 0 )
            exit function
         End If
   End Select

   Function = 0
End Function


' ========================================================================================
' Process WM_CLOSE message for window/dialog: frmProjectOptions
' ========================================================================================
Function frmProjectOptions_OnClose(HWnd As HWnd) As LRESULT
   ' Enables parent window keeping parent's zorder
   EnableWindow( GetParent(HWnd), True )
   DestroyWindow( HWnd )
   Function = 0
End Function


' ========================================================================================
' Process WM_DESTROY message for window/dialog: frmProjectOptions
' ========================================================================================
Function frmProjectOptions_OnDestroy(HWnd As HWnd) As LRESULT
   PostQuitMessage(0)
   Function = 0
End Function


' ========================================================================================
' frmProjectOptions Window procedure
' ========================================================================================
Function frmProjectOptions_WndProc( ByVal HWnd   As HWnd, _
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

   ' for messages that we don't deal with
   Function = DefWindowProc(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' frmProjectOptions_Show
' ========================================================================================
Function frmProjectOptions_Show( ByVal hWndParent As HWnd, _
                                 ByVal nCmdShow   As Long = 0, _
                                 byval IsNewProject as boolean = false _
                                 ) As Long

   Dim wszTitle As WString * MAX_PATH
   
   '  Create the main window and child controls
   Dim pWindow As CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowOwnerPtr(hwndParent)->DPI

   Dim As HWnd hForm = _
   pWindow->Create( hWndParent, "", @frmProjectOptions_WndProc, 0, 0, 0, 0, _
        WS_POPUP Or WS_CAPTION Or WS_SYSMENU Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, _
        WS_EX_DLGMODALFRAME Or WS_EX_CONTROLPARENT Or WS_EX_LEFT)
   pWindow->SetClientSize(600, 320)
   pWindow->Center(pWindow->hWindow, hWndParent)
      
   pWindow->AddControl("LABEL", , IDC_FRMPROJECTOPTIONS_LABEL1, L(240,"Project Path"), 26, 15, 235, 18, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   Dim As HWnd hProjectPath = _
   pWindow->AddControl("TEXTBOX", , IDC_FRMPROJECTOPTIONS_TXTPROJECTPATH, "", 26, 35, 346, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->AddControl("BUTTON", , IDC_FRMPROJECTOPTIONS_CMDSELECT, "...", 380, 33, 37, 24, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)

   pWindow->AddControl("LABEL", , IDC_FRMPROJECTOPTIONS_LABEL5, L(283,"Specify compiler options in addition to the selected build configuration"), 26, 95, 500, 18, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)

   pWindow->AddControl("LABEL", , IDC_FRMPROJECTOPTIONS_LABEL2, L(241,"Other Options (32-bit compiler)"), 26, 121, 235, 18, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   Dim As HWnd hProjectOther32 = _
   pWindow->AddControl("TEXTBOX", , IDC_FRMPROJECTOPTIONS_TXTOPTIONS32, "", 26, 141, 555, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   
   pWindow->AddControl("LABEL", , IDC_FRMPROJECTOPTIONS_LABEL4, L(242,"Other Options (64-bit compiler)"), 26, 170, 235, 18, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   Dim As HWnd hProjectOther64 = _
   pWindow->AddControl("TEXTBOX", , IDC_FRMPROJECTOPTIONS_TXTOPTIONS64, "", 26, 190, 555, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)

   dim as hwnd hManifest = _
   pWindow->AddControl("CHECKBOX", , IDC_FRMPROJECTOPTIONS_CHKMANIFEST, L(308,"Create resource file and manifest"), 26, 222, 400, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)

   pWindow->AddControl("LABEL", , IDC_FRMPROJECTOPTIONS_LABEL3, "", 10, 250, 572, 3, _
        WS_CHILD Or WS_VISIBLE Or SS_LEFT Or SS_NOTIFY Or SS_SUNKEN, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("BUTTON", , IDOK, L(0,"&OK"), 423, 265, 74, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_DEFPUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("BUTTON", , IDCANCEL, L(1,"&Cancel"), 508, 265, 74, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)

   wszTitle = L(214,"Project Options")

   
   dim as long idx = iif(IsNewProject, gApp.GetNewProjectIndex(), gApp.GetActiveProjectIndex() )
   gApp.ProjectOverrideIndex = idx
   
   ' For new projects or projects being opened from disk, the call to GetProjectIndex
   ' sets the InUse flag to TRUE. For a New project we simply want to get the details
   ' of the new project and then save the file to disk. A later call opens the project
   ' from disk and sets it up in the editor. Therefore, before exiting this dialog we
   ' will reset the InUse flag to FALSE because the open project code will assign a new one.
   gApp.IsNewProjectFlag = iif( IsNewProject, true, false )
   
   If gApp.IsNewProjectFlag Then
      ' Set the default values for a new project
      gApp.Projects(idx).InUse = true
      wszTitle = wszTitle & " - " & L(215, "New Project")
      gApp.Projects(idx).ProjectFilename = ""
      gApp.Projects(idx).ProjectName     = ""
      gApp.Projects(idx).ProjectBuild    = ""
      gApp.Projects(idx).ProjectOther32  = ""
      gApp.Projects(idx).ProjectOther64  = ""
      gApp.Projects(idx).ProjectManifest = true
      gApp.Projects(idx).ProjectNotes    = ""
   Else
      wszTitle = wszTitle + " - " + gApp.Projects(idx).ProjectName
   End If
   
   SetWindowText( hForm, wszTitle )
   SetWindowText( hProjectPath, gApp.Projects(idx).ProjectFilename )
   
   SetWindowTextW( hProjectOther32, gApp.Projects(idx).ProjectOther32 )
   SetWindowTextW( hProjectOther64, gApp.Projects(idx).ProjectOther64 )

   Button_SetCheck( hManifest, gApp.Projects(idx).ProjectManifest)

   SetFocus GetDlgItem(hForm, IDOK)
   
   ' Process Windows messages(modal)
   Function = pWindow->DoEvents(nCmdShow)
   
   ' Delete the frmProjectOptions CWindow class manually allocated memory 
   Delete pWindow

End Function










