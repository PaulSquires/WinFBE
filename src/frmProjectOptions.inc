'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2017 Paul Squires, PlanetSquires Software
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
' Save all options for the Project
' ========================================================================================
Function SaveProjectOptions( ByVal HWnd As HWnd ) As BOOLEAN    
   
   Dim wText As WString * MAX_PATH  
    
   Dim As HWnd hCtrl = GetDlgItem(HWnd, IDC_FRMPROJECTOPTIONS_TXTPROJECTPATH)
   gApp.ProjectFilename = AfxGetWindowText(hCtrl)

   If Len(gApp.ProjectFilename) = 0 Then
      MessageBoxW( HWnd, L(217,"Invalid project path specified."), L(201,"Error"), _
                        MB_OK Or MB_ICONINFORMATION Or MB_DEFBUTTON1 )
      SetFocus hCtrl
      Function = False: Exit Function
   End If

   gApp.ProjectName = AfxStrPathname( "NAMEX", gApp.ProjectFilename )

   hCtrl = GetDlgItem(HWnd, IDC_FRMPROJECTOPTIONS_OPTEXE)
   If Button_GetCheck(hCtrl) = BST_CHECKED Then gApp.ProjectType = 0
   hCtrl = GetDlgItem(HWnd, IDC_FRMPROJECTOPTIONS_OPTDLL)
   If Button_GetCheck(hCtrl) = BST_CHECKED Then gApp.ProjectType = 1
   hCtrl = GetDlgItem(HWnd, IDC_FRMPROJECTOPTIONS_OPTLIB)
   If Button_GetCheck(hCtrl) = BST_CHECKED Then gApp.ProjectType = 2
   
   hCtrl = GetDlgItem(HWnd, IDC_FRMPROJECTOPTIONS_OPTNOERROR)
   If Button_GetCheck(hCtrl) = BST_CHECKED Then gApp.ProjectErrorOption = 0
   hCtrl = GetDlgItem(HWnd, IDC_FRMPROJECTOPTIONS_OPTERROR)
   If Button_GetCheck(hCtrl) = BST_CHECKED Then gApp.ProjectErrorOption = 1
   hCtrl = GetDlgItem(HWnd, IDC_FRMPROJECTOPTIONS_OPTRESUME)
   If Button_GetCheck(hCtrl) = BST_CHECKED Then gApp.ProjectErrorOption = 2
   hCtrl = GetDlgItem(HWnd, IDC_FRMPROJECTOPTIONS_OPTERRORPLUS)
   If Button_GetCheck(hCtrl) = BST_CHECKED Then gApp.ProjectErrorOption = 3

   hCtrl = GetDlgItem(HWnd, IDC_FRMPROJECTOPTIONS_CHKDEBUG)
   gApp.ProjectDebug = Button_GetCheck( hCtrl )
   hCtrl = GetDlgItem(HWnd, IDC_FRMPROJECTOPTIONS_CHKTHREAD)
   gApp.ProjectThread = Button_GetCheck( hCtrl )
   hCtrl = GetDlgItem(HWnd, IDC_FRMPROJECTOPTIONS_CHKSHOWCONSOLE)
   gApp.ProjectShowConsole = Button_GetCheck( hCtrl )

   hCtrl = GetDlgItem(HWnd, IDC_FRMPROJECTOPTIONS_TXTOPTIONS32)
   gApp.ProjectOther32 = AfxGetWindowText(hCtrl)
   hCtrl = GetDlgItem(HWnd, IDC_FRMPROJECTOPTIONS_TXTOPTIONS64)
   gApp.ProjectOther64 = AfxGetWindowText(hCtrl)
   
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

   Select Case id
      Case IDC_FRMPROJECTOPTIONS_CMDSELECT
         If codeNotify = BN_CLICKED Then
            Dim pwszName As WString Ptr = AfxIFileSaveDialog(HWnd, gApp.ProjectName, WStr("wfbe"), IDC_FRMPROJECTOPTIONS_CMDSELECT)
            If pwszName Then
               SetWindowTextW( GetDlgItem(HWnd, IDC_FRMPROJECTOPTIONS_TXTPROJECTPATH), pwszName )
               CoTaskMemFree pwszName
            End If
            Exit Function
         End If
         
      Case IDOK
         If codeNotify = BN_CLICKED Then
            If SaveProjectOptions(HWnd) Then
               gApp.IsNewProjectFlag = False
               gApp.IsProjectActive = True
               ' Update the most recently used project list 
               UpdateMRUProjectList(gApp.ProjectFilename)
               SendMessageW HWnd, WM_CLOSE, 0, 0
            End If
            Exit Function
         End If
      
      Case IDCANCEL
         If codeNotify = BN_CLICKED Then
            SendMessageW HWnd, WM_CLOSE, 0, 0
            Exit Function
         End If
   End Select

   Function = 0
End Function


' ========================================================================================
' Process WM_CLOSE message for window/dialog: frmProjectOptions
' ========================================================================================
Function frmProjectOptions_OnClose(HWnd As HWnd) As LRESULT
   ' Enables parent window keeping parent's zorder
   EnableWindow GetParent(HWnd), True 
   DestroyWindow HWnd
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
   Function = DefWindowProcW(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' frmProjectOptions_Show
' ========================================================================================
Function frmProjectOptions_Show( ByVal hWndParent As HWnd, _
                                 ByVal nCmdShow   As Long = 0 _
                                 ) As Long

   Dim wszTitle As WString * MAX_PATH
   
   '  Create the main window and child controls
   Dim pWindow As CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowOwnerPtr(hwndParent)->DPI

   Dim As HWnd hForm = _
   pWindow->Create( hWndParent, "", @frmProjectOptions_WndProc, 0, 0, 0, 0, _
        WS_POPUP Or WS_CAPTION Or WS_SYSMENU Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, _
        WS_EX_DLGMODALFRAME Or WS_EX_CONTROLPARENT Or WS_EX_LEFT)
   pWindow->SetClientSize(720, 420)
   pWindow->Center(pWindow->hWindow, hWndParent)
      

   pWindow->AddControl("GROUPBOX", , IDC_FRMPROJECTOPTIONS_FRAME2, L(229,"Project Type"), 15, 15, 279, 125, _
        WS_CHILD Or WS_VISIBLE Or BS_TEXT Or BS_LEFT Or BS_NOTIFY Or BS_GROUPBOX, _
        WS_EX_TRANSPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("OPTION", , IDC_FRMPROJECTOPTIONS_OPTEXE, L(230,"Executable (EXE)"), 30, 34, 255, 21, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTORADIOBUTTON Or WS_GROUP Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("OPTION", , IDC_FRMPROJECTOPTIONS_OPTDLL, L(231,"Dynamic Link Library (-dll)"), 30, 55, 255, 21, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTORADIOBUTTON Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("OPTION", , IDC_FRMPROJECTOPTIONS_OPTLIB, L(232,"Static Library (-lib)"), 30, 76, 255, 21, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTORADIOBUTTON Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("GROUPBOX", , IDC_FRMPROJECTOPTIONS_FRAME1, L(233,"Error Checking"), 305, 15, 399, 125, _
        WS_CHILD Or WS_VISIBLE Or BS_TEXT Or BS_LEFT Or BS_NOTIFY Or BS_GROUPBOX, _
        WS_EX_TRANSPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("OPTION", , IDC_FRMPROJECTOPTIONS_OPTNOERROR, L(234,"No error checking"), 320, 37, 370, 21, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTORADIOBUTTON Or WS_GROUP Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("OPTION", , IDC_FRMPROJECTOPTIONS_OPTERROR, L(235,"Add error checking (-e)"), 320, 58, 370, 21, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTORADIOBUTTON Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("OPTION", , IDC_FRMPROJECTOPTIONS_OPTRESUME, L(236,"Error checking with RESUME support (-ex)"), 320, 79, 370, 21, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTORADIOBUTTON Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("OPTION", , IDC_FRMPROJECTOPTIONS_OPTERRORPLUS, L(237,"Same as -ex plus array bounds and null pointer (-exx)"), 320, 100, 370, 21, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTORADIOBUTTON Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   Dim As HWnd hProjectDebug = _
   pWindow->AddControl("CHECKBOX", , IDC_FRMPROJECTOPTIONS_CHKDEBUG, L(238,"Add debug info (-g)"), 30, 150, 278, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   Dim As HWnd hProjectThread = _
   pWindow->AddControl("CHECKBOX", , IDC_FRMPROJECTOPTIONS_CHKTHREAD, L(239,"Link with thread-safe runtime library (-mt)"), 30, 172, 278, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   Dim As HWnd hShowConsole = _
   pWindow->AddControl("CHECKBOX", , IDC_FRMPROJECTOPTIONS_CHKSHOWCONSOLE, L(245,"Show console for GUI applications"), 320, 150, 350, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("LABEL", , IDC_FRMPROJECTOPTIONS_LABEL1, L(240,"Project Path"), 30, 211, 235, 18, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   Dim As HWnd hProjectPath = _
   pWindow->AddControl("TEXTBOX", , IDC_FRMPROJECTOPTIONS_TXTPROJECTPATH, "", 31, 231, 346, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->AddControl("BUTTON", , IDC_FRMPROJECTOPTIONS_CMDSELECT, "...", 385, 229, 37, 24, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)

   pWindow->AddControl("LABEL", , IDC_FRMPROJECTOPTIONS_LABEL2, L(241,"Other Options (32-bit compiler)"), 31, 265, 235, 18, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   Dim As HWnd hProjectOther32 = _
   pWindow->AddControl("TEXTBOX", , IDC_FRMPROJECTOPTIONS_TXTOPTIONS32, "", 32, 285, 346, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   
   pWindow->AddControl("LABEL", , IDC_FRMPROJECTOPTIONS_LABEL4, L(242,"Other Options (64-bit compiler)"), 31, 314, 235, 18, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   Dim As HWnd hProjectOther64 = _
   pWindow->AddControl("TEXTBOX", , IDC_FRMPROJECTOPTIONS_TXTOPTIONS64, "", 32, 334, 346, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)

   pWindow->AddControl("LABEL", , IDC_FRMPROJECTOPTIONS_LABEL3, "", 10, 370, 700, 3, _
        WS_CHILD Or WS_VISIBLE Or SS_LEFT Or SS_NOTIFY Or SS_SUNKEN, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("BUTTON", , IDOK, L(0,"&OK"), 550, 380, 74, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_DEFPUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("BUTTON", , IDCANCEL, L(1,"&Cancel"), 635, 380, 74, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)

   wszTitle = L(214,"Project Options")

   If gApp.IsNewProjectFlag Then
      ' Set the default values for a new project
      wszTitle = wszTitle & " - " & L(215, "New Project")
      gApp.ProjectFilename = 0
      gApp.ProjectName         = 0
      gApp.ProjectType         = 0
      gApp.ProjectErrorOption  = 0
      gApp.ProjectDebug        = BST_UNCHECKED
      gApp.ProjectThread       = BST_UNCHECKED
      gApp.ProjectOther32      = 0
      gApp.ProjectOther64      = 0
      gApp.ProjectShowConsole  = 0
   Else
      wszTitle = wszTitle + " - " + gApp.ProjectName
   End If
   
   SetWindowTextW( hForm, wszTitle )
   SetWindowTextW( hProjectPath, gApp.ProjectFilename )
   
   Dim nOptions(3) As Long

   nOptions(0) = IDC_FRMPROJECTOPTIONS_OPTEXE
   nOptions(1) = IDC_FRMPROJECTOPTIONS_OPTDLL  
   nOptions(2) = IDC_FRMPROJECTOPTIONS_OPTLIB 
   CheckRadioButton( hForm, nOptions(0), nOptions(2), nOptions(gApp.ProjectType) )
 
   nOptions(0) = IDC_FRMPROJECTOPTIONS_OPTNOERROR  
   nOptions(1) = IDC_FRMPROJECTOPTIONS_OPTERROR
   nOptions(2) = IDC_FRMPROJECTOPTIONS_OPTRESUME
   nOptions(3) = IDC_FRMPROJECTOPTIONS_OPTERRORPLUS
   CheckRadioButton( hForm, nOptions(0), nOptions(3), nOptions(gApp.ProjectErrorOption) )

   Button_SetCheck( hProjectDebug, gApp.ProjectDebug )
   Button_SetCheck( hProjectThread, gApp.ProjectThread )
   Button_SetCheck( hShowConsole, gApp.ProjectShowConsole )
   
   SetWindowTextW( hProjectOther32, gApp.ProjectOther32 )
   SetWindowTextW( hProjectOther64, gApp.ProjectOther64 )

   SetFocus GetDlgItem(hForm, IDOK)
   
   ' Process Windows messages(modal)
   Function = pWindow->DoEvents(nCmdShow)
   
   ' Delete the frmProjectOptions CWindow class manually allocated memory 
   Delete pWindow

End Function










