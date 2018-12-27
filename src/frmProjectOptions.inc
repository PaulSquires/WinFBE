'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2019 Paul Squires, PlanetSquires Software
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
function CreateResourceManifest() as Long

   dim as CWSTR wszProjectPath = AfxStrPathname("PATH", gApp.ProjectFilename)
   dim as CWSTR wszManifest    = wszProjectPath + "manifest.xml"
   dim as CWSTR wszResource    = wszProjectPath + "resource.rc"
   dim as CWSTR wszManifestTemplate = AfxGetExePathName + "Settings\WinFBE_manifest.xml"
   dim as CWSTR wszResourceTemplate = AfxGetExePathName + "Settings\WinFBE_resource.rc" 
   
   if AfxFileExists(wszManifestTemplate) = false then exit function
   if AfxFileExists(wszResourceTemplate) = false then exit function
   if AfxFileExists(wszManifest) = false then CopyFile( wszManifestTemplate, wszManifest, true )
   if AfxFileExists(wszResource) = false then CopyFile( wszResourceTemplate, wszResource, true )

   ' Add the resource file to the project if it does not already exist in the project
   if gApp.GetDocumentPtrByFilename(wszResource) = 0 then
      dim pDoc as clsDocument ptr
      pDoc = frmMain_OpenFileSafely(HWND_FRMMAIN, _
                              False, _    ' bIsNewFile
                              False, _    ' bIsTemplate
                              FALSE, _    ' bShowInTab
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
Function SaveProjectOptions( ByVal HWnd As HWnd ) As BOOLEAN    
   
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
   TreeView_SetItemText(GetDlgItem(HWND_FRMEXPLORER, IDC_FRMEXPLORER_TREE), _
                          gApp.hExplorerRootNode, ucase(gApp.ProjectName))

   gApp.ProjectOther32     = AfxGetWindowText(GetDlgItem(HWnd, IDC_FRMPROJECTOPTIONS_TXTOPTIONS32))
   gApp.ProjectOther64     = AfxGetWindowText(GetDlgItem(HWnd, IDC_FRMPROJECTOPTIONS_TXTOPTIONS64))
   gApp.ProjectManifest    = Button_GetCheck(GetDlgItem(HWnd, IDC_FRMPROJECTOPTIONS_CHKMANIFEST))
   gApp.ProjectBuild       = GetSelectedBuildGUID()
   gApp.ProjectNotes       = ""
   gApp.ProjectCommandLine = ""
   

   ' Create the resource file and manifest if applicable
   if gApp.ProjectManifest THEN CreateResourceManifest()

   ' Save the actual Project data to disk.
   gConfig.ProjectSaveToFile()   

   if gApp.IsNewProjectFlag then 
      ' Need to remove resource file (if created) from the document list otherwise it
      ' will be loaded again and it will display twice in the Explorer.
      gApp.RemoveAllDocuments()
      
      ' Open the new project so it displays in the Explorer treeview. Pass the 
      ' variable wText because passing gApp.ProjectFilename will cause that variable
      ' to get reset because the parameter is byref and eventually that variable
      ' gets assigned a null.
      frmMain_OpenProjectSafely(HWND_FRMMAIN, wText) 

     ' ' Add a new blank file to the project and set it as the Main by default   
      dim pDoc as clsDocument ptr = OnCommand_FileNew(HWND_FRMMAIN)
      if pDoc then gApp.ProjectSetFileType(pDoc, FILETYPE_MAIN)
   end if
   
   function = true   
End Function



' ========================================================================================
' Process WM_COMMAND message for window/dialog: frmProjectOptions
' ========================================================================================
Function frmProjectOptions_OnCommand(ByVal HWnd As HWnd, ByVal id As Long, ByVal hwndCtl As HWnd, ByVal codeNotify As UINT) As LRESULT

   Select Case id
      Case IDC_FRMPROJECTOPTIONS_CMDSELECT
         If codeNotify = BN_CLICKED Then
            dim as CWSTR wszProjectName = ""
            if gApp.IsNewProjectFlag = false then wszProjectName = gApp.ProjectName
            Dim pwszName As WString Ptr = AfxIFileSaveDialog(HWnd, wszProjectName, WStr("wfbe"), IDC_FRMPROJECTOPTIONS_CMDSELECT)
            If pwszName Then
               AfxSetWindowText( GetDlgItem(HWnd, IDC_FRMPROJECTOPTIONS_TXTPROJECTPATH), pwszName )
               CoTaskMemFree pwszName
            End If
            Exit Function
         End If
         
      Case IDOK
         If codeNotify = BN_CLICKED Then
            If SaveProjectOptions(HWnd) Then
               SendMessage HWnd, WM_CLOSE, 0, 0
            End If
            Exit Function
         End If
      
      Case IDCANCEL
         If codeNotify = BN_CLICKED Then
            SendMessage HWnd, WM_CLOSE, 0, 0
            exit function
         End If
   End Select

   Function = 0
End Function


' ========================================================================================
' Process WM_CREATE message for window/dialog: frmProjectOptions
' ========================================================================================
Function frmProjectOptions_OnCreate(ByVal HWnd As HWnd, ByVal lpCreateStructPtr As LPCREATESTRUCT) As BOOLEAN

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
Function frmProjectOptions_OnClose(HWnd As HWnd) As LRESULT
   ' Enables parent window keeping parent's zorder
   EnableAllModeless()
   DestroyWindow HWnd
   Function = 0
End Function


' ========================================================================================
' Process WM_DESTROY message for window/dialog: frmProjectOptions
' ========================================================================================
Function frmProjectOptions_OnDestroy(HWnd As HWnd) As LRESULT
   gApp.IsProjectLoading = FALSE
   gApp.IsNewProjectFlag = false
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
   Function = pWindow->DoEvents(nCmdShow)
   
   ' Delete the frmProjectOptions CWindow class manually allocated memory 
   Delete pWindow

End Function
