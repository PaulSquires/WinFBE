
' ========================================================================================
' Save all options for the Project to a diskfile
' ========================================================================================
Function ProjectSaveToFile() As BOOLEAN    

   Dim wszText As WString * MAX_PATH
   Dim As Long nCount = gpApp->GetDocumentCount 
   Dim pDoc As clsDocument Ptr
   Dim As Long i

   Dim f As Long = Freefile
   If Open( *gpApp->ProjectDiskFilename For Output Encoding "utf16" As #f ) <> 0 Then Return False

   Print #f, "' WINFBE PROJECT FILE"
   Print #f, "ProjectType=";         gpApp->ProjectType
   Print #f, "ProjectErrorOption=";  gpApp->ProjectErrorOption
   Print #f, "ProjectDebug=";        gpApp->ProjectDebug
   Print #f, "ProjectThread=";       gpApp->ProjectThread
   Print #f, "ProjectOther=";        *gpApp->ProjectOther
   Print #f, "DefaultCompiler=";     gConfig.DefaultCompiler       
   Print #f, "DefaultCompileMode=";  gConfig.DefaultCompileMode    
   Print #f, "CommandLine=";         gConfig.CommandLine
   
   ' Save the properties for each file in the project
   For i = 0 To nCount - 1
      pDoc = gpApp->GetDocumentPtr(i)
      If pDoc Then 
         If pDoc->IsProjectFile Then
            wszText = *pDoc->DiskFilename
            ProcessToCurdrive(@wszText)
            Print #f, "File="; wszText
            Print #f, "FileType="; pDoc->ProjectFileType
            Print #f, "TabIndex="; gTTabCtl.GetTabIndexByDocumentPtr(pDoc)
            Print #f, "FileEnd=[-]"
         End If   
      End If
   Next

   Close #f
   
   Function = True   ' successul save
End Function


' ========================================================================================
' Load all options for the Project from a diskfile
' ========================================================================================
Function ProjectLoadFromFile() As BOOLEAN    

   Dim st    As WString * MAX_PATH
   Dim wKey  As WString * MAX_PATH
   Dim wData As WString * MAX_PATH
   Dim nData As Long  
   Dim i     As Long

   ' Info saved for each File and used to create the file once FileEnd is found
   Dim pDoc As clsDocument Ptr
   Dim wszFilename As WString * MAX_PATH
   Dim As Long nFileType, nTabIndex
   
   If Dir(*gpApp->ProjectDiskFilename) = "" Then Exit Function
   FF_Pathname(@WStr("NAMEX"), gpApp->ProjectDiskFilename, gpApp->ProjectName)
    
   Dim f As Long = Freefile
   If Open( *gpApp->ProjectDiskFilename For Input Encoding "utf16" As #f ) <> 0 Then Exit Function

   frmProjectManager_Show( HWND_FRMMAIN, SW_SHOW )
   
   ' Turn off redrawing of the listview in order to speed up loading
   Dim As HWnd hLV = GetDlgItem(HWND_FRMPROJECTMANAGER, IDC_FRMPROJECTMANAGER_LISTVIEW )
   SendMessage( hLV, WM_SETREDRAW, False, 0 )
   
   Do Until Eof(f)
      Line Input #f, st
      If Len(st) = 0 Then Continue Do
      If Left(st, 1) = "'" Then Continue Do
      If Left(st, 1) = "[" Then Continue Do
      
      i = Instr(st, "=")
      If i = 0 Then Continue Do
      
      wKey  = Left(st, i-1)
      wData = Mid(st, i+1)
      nData = Val(wData)
      
      ProcessFromCurdrive(@wData)

      Select Case wKey
         Case "ProjectType":         gpApp->ProjectType         = nData
         Case "ProjectErrorOption":  gpApp->ProjectErrorOption  = nData
         Case "ProjectDebug":        gpApp->ProjectDebug        = nData
         Case "ProjectThread":       gpApp->ProjectThread       = nData
         Case "ProjectOther":        gpApp->ProjectOther        = wData
         Case "DefaultCompiler":     gConfig.DefaultCompiler    = wData
         Case "DefaultCompileMode":  gConfig.DefaultCompileMode = wData
         Case "CommandLine":         gConfig.CommandLine        = wData
         
         Case "File":                wszFilename = wData
         Case "FileType":            nFileType = nData
         Case "TabIndex":            nTabIndex = Iif(nData <> -1, True, False)
         Case "FileEnd":
            If Dir(wszFilename) > "" Then
               pDoc = frmMain_OpenFileSafely(HWND_FRMMAIN, False, False, nTabIndex, @wszFilename, 0 )
               pDoc->ProjectFileType = nFileType
               gpApp->ProjectAddFile(pDoc)     ' Add the document to the active Project & Listview     
            End If
         
         
      End Select
   
   Loop                        
   
   gpApp->IsProjectActive = True
      
   Close #f

   Listview_SortColumn( hLV, gLVProjSortCol, gLVProjSortDir )
 
   SendMessage( hLV, WM_SETREDRAW, True, 0 )
   RedrawWindow(hLV, Null, Null, RDW_ERASE Or RDW_FRAME Or RDW_INVALIDATE Or RDW_ALLCHILDREN)
    
   Function = True   ' successul save
End Function


' ========================================================================================
' Save all options for the Project
' ========================================================================================
Function SaveProjectOptions( ByVal HWnd As HWnd ) As BOOLEAN    
   
   Dim wText As WString * MAX_PATH  
    
   Dim As HWnd hCtrl = GetDlgItem(HWnd, IDC_FRMPROJECTOPTIONS_TXTPROJECTPATH)
   FF_Control_GetTextW( hCtrl, gpApp->ProjectDiskFilename, MAX_PATH )

   If Len(*gpApp->ProjectDiskFilename) = 0 Then
      MessageBoxW( HWnd, L(217,"Invalid project path specified."), L(201,"Error"), _
                        MB_OK Or MB_ICONINFORMATION Or MB_DEFBUTTON1 )
      SetFocus hCtrl
      Function = False: Exit Function
   End If

   FF_Pathname(@WStr("NAMEX"), gpApp->ProjectDiskFilename, gpApp->ProjectName)

   hCtrl = GetDlgItem(HWnd, IDC_FRMPROJECTOPTIONS_OPTEXE)
   If Button_GetCheck(hCtrl) = BST_CHECKED Then gpApp->ProjectType = 0
   hCtrl = GetDlgItem(HWnd, IDC_FRMPROJECTOPTIONS_OPTDLL)
   If Button_GetCheck(hCtrl) = BST_CHECKED Then gpApp->ProjectType = 1
   hCtrl = GetDlgItem(HWnd, IDC_FRMPROJECTOPTIONS_OPTLIB)
   If Button_GetCheck(hCtrl) = BST_CHECKED Then gpApp->ProjectType = 2
   
   hCtrl = GetDlgItem(HWnd, IDC_FRMPROJECTOPTIONS_OPTNOERROR)
   If Button_GetCheck(hCtrl) = BST_CHECKED Then gpApp->ProjectErrorOption = 0
   hCtrl = GetDlgItem(HWnd, IDC_FRMPROJECTOPTIONS_OPTERROR)
   If Button_GetCheck(hCtrl) = BST_CHECKED Then gpApp->ProjectErrorOption = 1
   hCtrl = GetDlgItem(HWnd, IDC_FRMPROJECTOPTIONS_OPTRESUME)
   If Button_GetCheck(hCtrl) = BST_CHECKED Then gpApp->ProjectErrorOption = 2
   hCtrl = GetDlgItem(HWnd, IDC_FRMPROJECTOPTIONS_OPTERRORPLUS)
   If Button_GetCheck(hCtrl) = BST_CHECKED Then gpApp->ProjectErrorOption = 3

   hCtrl = GetDlgItem(HWnd, IDC_FRMPROJECTOPTIONS_CHKDEBUG)
   gpApp->ProjectDebug = Button_GetCheck( hCtrl )
   hCtrl = GetDlgItem(HWnd, IDC_FRMPROJECTOPTIONS_CHKTHREAD)
   gpApp->ProjectThread = Button_GetCheck( hCtrl )

   hCtrl = GetDlgItem(HWnd, IDC_FRMPROJECTOPTIONS_TXTOPTIONS)
   FF_Control_GetTextW( hCtrl, gpApp->ProjectOther, MAX_PATH )
   
   ' Ensure that the subfolders exist for the project
   Dim wzPath As WString * MAX_PATH
   FF_Pathname(@WStr("PATH"), gpApp->ProjectDiskFilename, @wzPath)

   SHCreateDirectory( 0, wzPath & "release32" )
   SHCreateDirectory( 0, wzPath & "release32\obj" )
   SHCreateDirectory( 0, wzPath & "release64" )
   SHCreateDirectory( 0, wzPath & "release64\obj" )

   Return ProjectSaveToFile()   

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
            Dim pwszName As WString Ptr = AfxIFileSaveDialog(HWnd, gpApp->ProjectName, WStr("wfbe"), IDC_FRMPROJECTOPTIONS_CMDSELECT)
            If pwszName Then
               SetWindowTextW( GetDlgItem(HWnd, IDC_FRMPROJECTOPTIONS_TXTPROJECTPATH), pwszName )
               CoTaskMemFree pwszName
            End If
            Exit Function
         End If
         
      Case IDOK
         If codeNotify = BN_CLICKED Then
            If SaveProjectOptions(HWnd) Then
               gpApp->IsNewProjectFlag = False
               gpApp->IsProjectActive = True
               ' Update the most recently used project list 
               UpdateMRUProjectList(gpApp->ProjectDiskFilename)
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
   pWindow->SetClientSize(720, 385)
   pWindow->Center(pWindow->hWindow, hWndParent)
      

   pWindow->AddControl("GROUPBOX", , IDC_FRMPROJECTOPTIONS_FRAME2, "Project Type", 15, 15, 279, 125, _
        WS_CHILD Or WS_VISIBLE Or BS_TEXT Or BS_LEFT Or BS_NOTIFY Or BS_GROUPBOX, _
        WS_EX_TRANSPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("OPTION", , IDC_FRMPROJECTOPTIONS_OPTEXE, "Executable (EXE)", 30, 34, 255, 21, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTORADIOBUTTON Or WS_GROUP Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("OPTION", , IDC_FRMPROJECTOPTIONS_OPTDLL, "Dynamic Link Library (-dll)", 30, 55, 255, 21, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTORADIOBUTTON Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("OPTION", , IDC_FRMPROJECTOPTIONS_OPTLIB, "Static Library (-lib)", 30, 76, 255, 21, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTORADIOBUTTON Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("GROUPBOX", , IDC_FRMPROJECTOPTIONS_FRAME1, "Error Checking", 305, 15, 399, 125, _
        WS_CHILD Or WS_VISIBLE Or BS_TEXT Or BS_LEFT Or BS_NOTIFY Or BS_GROUPBOX, _
        WS_EX_TRANSPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("OPTION", , IDC_FRMPROJECTOPTIONS_OPTNOERROR, "No error checking", 320, 37, 370, 21, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTORADIOBUTTON Or WS_GROUP Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("OPTION", , IDC_FRMPROJECTOPTIONS_OPTERROR, "Add error checking (-e)", 320, 58, 370, 21, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTORADIOBUTTON Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("OPTION", , IDC_FRMPROJECTOPTIONS_OPTRESUME, "Error checking with RESUME support (-ex)", 320, 79, 370, 21, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTORADIOBUTTON Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("OPTION", , IDC_FRMPROJECTOPTIONS_OPTERRORPLUS, "Same as -ex plus array bounds and null pointer (-exx)", 320, 100, 370, 21, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTORADIOBUTTON Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   Dim As HWnd hProjectDebug = _
   pWindow->AddControl("CHECKBOX", , IDC_FRMPROJECTOPTIONS_CHKDEBUG, "Add debug info (-g)", 30, 150, 278, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   Dim As HWnd hProjectThread = _
   pWindow->AddControl("CHECKBOX", , IDC_FRMPROJECTOPTIONS_CHKTHREAD, "Link with thread-safe runtime library (-mt)", 30, 172, 278, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("LABEL", , IDC_FRMPROJECTOPTIONS_LABEL1, "Project Path", 30, 211, 235, 18, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   Dim As HWnd hProjectPath = _
   pWindow->AddControl("TEXTBOX", , IDC_FRMPROJECTOPTIONS_TXTPROJECTPATH, "", 31, 231, 346, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->AddControl("BUTTON", , IDC_FRMPROJECTOPTIONS_CMDSELECT, "...", 385, 229, 37, 24, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("LABEL", , IDC_FRMPROJECTOPTIONS_LABEL2, "Other Options", 31, 265, 235, 18, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   Dim As HWnd hProjectOther = _
   pWindow->AddControl("TEXTBOX", , IDC_FRMPROJECTOPTIONS_TXTOPTIONS, "", 32, 285, 346, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->AddControl("LABEL", , IDC_FRMPROJECTOPTIONS_LABEL3, "", 10, 325, 700, 3, _
        WS_CHILD Or WS_VISIBLE Or SS_LEFT Or SS_NOTIFY Or SS_SUNKEN, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("BUTTON", , IDOK, L(0,"&OK"), 550, 337, 74, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_DEFPUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("BUTTON", , IDCANCEL, L(1,"&Cancel"), 635, 337, 74, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)

   wszTitle = L(214,"Project Options")

   If gpApp->IsNewProjectFlag Then
      ' Set the default values for a new project
      wszTitle = wszTitle & " - " & L(215, "New Project")
      gpApp->ProjectDiskFilename = 0
      gpApp->ProjectName         = 0
      gpApp->ProjectType         = 0
      gpApp->ProjectErrorOption  = 0
      gpApp->ProjectDebug        = BST_UNCHECKED
      gpApp->ProjectThread       = BST_UNCHECKED
      gpApp->ProjectOther        = 0
   Else
      wszTitle = wszTitle & " - " & *gpApp->ProjectName
   End If
   
   SetWindowTextW( hForm, wszTitle )
   SetWindowTextW( hProjectPath, gpApp->ProjectDiskFilename )
   
   Dim nOptions(3) As Long

   nOptions(0) = IDC_FRMPROJECTOPTIONS_OPTEXE
   nOptions(1) = IDC_FRMPROJECTOPTIONS_OPTDLL  
   nOptions(2) = IDC_FRMPROJECTOPTIONS_OPTLIB 
   CheckRadioButton( hForm, nOptions(0), nOptions(2), nOptions(gpApp->ProjectType) )
 
   nOptions(0) = IDC_FRMPROJECTOPTIONS_OPTNOERROR  
   nOptions(1) = IDC_FRMPROJECTOPTIONS_OPTERROR
   nOptions(2) = IDC_FRMPROJECTOPTIONS_OPTRESUME
   nOptions(3) = IDC_FRMPROJECTOPTIONS_OPTERRORPLUS
   CheckRadioButton( hForm, nOptions(0), nOptions(3), nOptions(gpApp->ProjectErrorOption) )

   Button_SetCheck( hProjectDebug, gpApp->ProjectDebug )
   Button_SetCheck( hProjectThread, gpApp->ProjectThread )
   
   SetWindowTextW( hProjectOther, gpApp->ProjectOther )

   SetFocus GetDlgItem(hForm, IDOK)
   
   ' Process Windows messages
   Function = pWindow->DoEvents(nCmdShow)
   
   ' Delete the frmProjectOptions CWindow class manually allocated memory 
   Delete pWindow

End Function
