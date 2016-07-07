
' ========================================================================================
' Position editor to document/function under the current caret position
' ========================================================================================
Function frmMain_GotoDefinition( ByVal pDoc As clsDocument Ptr ) As Long

   If pDoc = 0 Then Exit Function
   
   Dim As Long i, nCount
   Dim pList As FUNCTION_TYPE Ptr
   
   ' Determine the word at the current caret position.
   Dim As String sWord = Ucase(pDoc->GetWord)
   
   Dim pDocFind As clsDocument Ptr
   
   nCount = gpApp->GetDocumentCount()
   For i = 0 To nCount - 1
      pDocFind = gpApp->GetDocumentPtr(i)
      If pDocFind Then 
         ' Search all of the names for a match on the current select word
         pList = pDocFind->FnListPtr
         Do Until pList = 0
            If Ucase(pList->zFnName) = sWord Then
               ' Save our current position in case user invokes "Last Position" option,   
               gLastPosition.pDoc = pDoc
               gLastPosition.nFirstLine = SendMessageW( pDoc->hWindow, SCI_GETFIRSTVISIBLELINE, 0, 0) 
               gLastPosition.nPosition = SendMessageW( pDoc->hWindow, SCI_GETCURRENTPOS, 0, 0) 

               frmMain_OpenFileSafely( HWND_FRMMAIN, False, False, True, pDocFind->DiskFilename, pDocFind )
               ' Set the top line to display in the editor. I chose to start 3 lines before the
               ' function just to make it visually more appealing.
               SendMessageW( pDocFind->hWindow, SCI_SETFIRSTVISIBLELINE, pList->nLineNumber - 3, 0) 
               SendMessageW( pDocFind->hWindow, SCI_GOTOLINE, pList->nLineNumber, 0) 
               Exit Function
            End If
            pList = pList->pListNext
         Loop      
      End If   
   Next
   
   ' If we get this far then the word was not found.
   MessageBoxW( HWND_FRMMAIN, L(224,"Sub/Function definition not found."), @WStr(APPNAME), MB_ICONWARNING)
   
   Function = 0
End Function


' ========================================================================================
' Return to most previous position (invoked when "Last Position" is selected
' ========================================================================================
Function frmMain_GotoLastPosition() As Long

   Dim As Long i, nCount
   Dim pDoc As clsDocument Ptr
   
   ' Is the pDoc pointer still valid
   If gLastPosition.pDoc = 0 Then Exit Function
   
   nCount = gpApp->GetDocumentCount()
   For i = 0 To nCount - 1
      pDoc = gpApp->GetDocumentPtr(i)
      If pDoc = gLastPosition.pDoc Then 
         ' Position ourselves to the correct previous position.
         frmMain_OpenFileSafely( HWND_FRMMAIN, False, False, True, pDoc->DiskFilename, pDoc )
         SendMessageW( pDoc->hWindow, SCI_SETFIRSTVISIBLELINE, gLastPosition.nFirstLine, 0) 
         SendMessageW( pDoc->hWindow, SCI_GOTOPOS, gLastPosition.nPosition, 0) 
         Exit Function
      End If   
   Next
   
   Function = 0
End Function


' ========================================================================================
' Update display info related to currently active document
' ========================================================================================
Function frmMain_UpdateLineCol( ByVal HWnd As HWnd) As Long

   ' Update the statusbar with the current Line/Col position
   Dim pDoc As clsDocument Ptr = gTTabCtl.GetActiveDocumentPtr()

   Dim As HWnd hStatusbar = GetDlgItem(HWnd, IDC_FRMMAIN_STATUSBAR)
   Dim wText As WString * 100 
   
   wText = " " & gConfig.DefaultCompiler
   StatusBar_SetText(hStatusbar, 0, wText)

   wText = " " & gConfig.DefaultCompileMode
   StatusBar_SetText(hStatusbar, 1, wText)

   ' If there is not any file being edited
   If pDoc = 0 Then
      StatusBar_SetText(hStatusbar, 2, "")    ' line/col/sel
      StatusBar_SetText(hStatusbar, 3, "")    ' length/lines
      Dim wzText As WString * MAX_PATH
      wzText = APPNAME            
      If gpApp->IsProjectActive Then wzText = APPNAMESHORT & " - [" & *gpApp->ProjectName & "]"
      SetWindowTextW( HWnd, wzText )
   Else
      pDoc->DisplayStats
      SetFocus(pDoc->hWindow)     
   End If
   frmMain_ChangeToolbarButtonsState

  Function = 0
End Function


' ========================================================================================
' Process Scintilla Notifications
' ========================================================================================
Function Scintilla_OnNotify( ByVal HWnd As HWnd, _
                             ByVal pNSC As SCNOTIFICATION Ptr _
                             ) As Long

   If pNSC = 0 Then Exit Function
   
   Dim hEdit      As HWnd 
   Dim nLine      As Long 
   Dim nFoldLevel As Long 
   

   Select Case pNSC->hdr.code
                     
      Case SCN_UPDATEUI
         ' Show line and column
         frmMain_UpdateLineCol(HWnd)
         
      Case SCN_MODIFIED
         If ((pNSC->modificationType And SC_MOD_INSERTTEXT) = SC_MOD_INSERTTEXT) Then
            frmMain_UpdateLineCol(HWnd)
         End If

      Case SCN_MARGINCLICK
         ' Folder margin
         Dim pDoc As clsDocument Ptr = gTTabCtl.GetActiveDocumentPtr()
         If pDoc Then 
            hEdit      = pDoc->hWindow
            nLine      = SendMessage(hEdit, SCI_LINEFROMPOSITION, pNSC->position, 0)
            nFoldLevel = SendMessage(hEdit, SCI_GETFOLDLEVEL, nLine, 0)
   
            Select Case pNSC->margin 
               Case 1   ' left margin (bookmarks)
                  pDoc->ToggleBookmark(nLine)
               Case 2   ' fold margin
                  ' If is the head line...
                  If (nFoldLevel And SC_FOLDLEVELHEADERFLAG) <> 0 Then
                     SendMessageW(hEdit, SCI_TOGGLEFOLD, nLine, 0) 
                  End If   
            End Select
         End If
         
      CASE SCN_CHARADDED
         If pNSC->ch = 13 Then  ' ENTER KEY PRESSED
            If gConfig.AutoIndentation Then
               Dim pDoc As clsDocument Ptr = gTTabCtl.GetActiveDocumentPtr()
               If pDoc Then
                  hEdit = pDoc->hWindow
                  ' Current position
                  Dim curPos As Long = SendMessage(hEdit, SCI_GETCURRENTPOS, 0, 0)
                  ' Current line
                  nLine = SendMessage(hEdit, SCI_LINEFROMPOSITION, curPos, 0)
                  ' Line length of the previous line
                  Dim LineLen As Long = SendMessage(hEdit, SCI_LINELENGTH, nLine - 1, 0)
                  If LineLen < 1 Then Exit Function
                  ' Get the text of the previous line
                  Dim buffer As String = Space(LineLen)
                  SendMessage(hEdit, SCI_GETLINE, nLine - 1, Cast(LPARAM, Strptr(buffer)))
                  ' Get the tab width
                  Dim TabSize As Long = SendMessage(hEdit, SCI_GETTABWIDTH, 0, 0)
                  ' Calculate the number of spaces to fill on the left
                  Dim nSpaces As Long, i As Long
                  For i = 1 To Len(buffer)
                     If Mid(buffer, i, 1) <> " " Then
                        If Mid(buffer, i, 1) = Chr(9) Then
                           nSpaces = nSpaces + TabSize
                        Else
                           Exit For
                        End If
                     Else
                        nSpaces = nSpaces + 1
                     End If
                  Next
                  Dim strFill As String = ""
                  ' Remove CRLF
                  buffer = FF_Remove(buffer, Chr(13))
                  buffer = FF_Remove(buffer, Chr(10))
                  ' Remove spaces and tabs and convert to uppercase
                  buffer = Trim(Ucase(buffer), Any Chr(32, 9))
                  If (Left(buffer, 3) = "IF " And Right(buffer, 5) = " THEN") Or _
                     Left(buffer, 4) = "ELSE" Or _
                     Left(buffer, 7) = "SELECT " Or _
                     Left(buffer, 5) = "CASE " Or _
                     Left(buffer, 4) = "FOR " Or _
                     Left(buffer, 3) = "DO " Or _
                     buffer = "DO" Or _
                     Left$(buffer, 6) = "WHILE " Or _
                     buffer = "WHILE" Then
                     ' // Indentation size
                     Dim IndentSize As Long = SendMessage(hEdit, SCI_GETINDENT, 0, 0)
                     ' // Add spaces to indent the line
                     strFill = Space(nSpaces + IndentSize)
                  Else
                     ' Add the same spaces on the left that the line above
                     strFill = Space(nSpaces)
                  End If
                  ' *** TODO: If using tabs, replace spaces with tabs
                  ' Indents the line
                  SendMessage(hEdit, SCI_ADDTEXT, Len(strFill), Cast(LPARAM, Strptr(strFill)))
               End If
            End If
         End If
         
   End Select

   Function = 0
End Function


' ========================================================================================
' Set focus to currently active Scintilla window
' ========================================================================================
Function frmMain_SetFocusToCurrentCodeWindow() As Long
   ' Post a message to the main form CUSTOM handler that will
   ' set focus to the currently active Scintilla code window. We
   ' use PostMessage to ensure that all all other windows 
   ' messages are finished processing.
   PostMessageW HWND_FRMMAIN, MSG_USER_SETFOCUS, 0, 0
   Function = 0
End Function


' ========================================================================================
' Position all child windows. Called manually and/or by WM_SIZE
' ========================================================================================
Function frmMain_PositionWindows( ByVal HWnd As HWnd ) As LRESULT
   
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(HWnd)
   
   Dim As HWnd hEdit, hTabCtl
   Dim As Rect rc
   Dim As Long nHeightRebar, nHeightStatusBar, nHeightTabControl
   
   ' Get the entire client area
   GetClientRect(HWnd, @rc)

   ' Get the heights of the rebar and statusbar
   nHeightRebar = AfxGetWindowHeight( GetDlgItem(HWnd, IDC_FRMMAIN_REBAR) )
   nHeightStatusBar = AfxGetWindowHeight( GetDlgItem(HWnd, IDC_FRMMAIN_STATUSBAR) )
   
   ' If items exist in the top tabcontrol then show the tab control and account for its height
   hTabCtl = GetDlgItem(HWnd, IDC_FRMMAIN_TOPTABCONTROL)
   If TabCtrl_GetItemCount(hTabCtl) = 0 Then
      ShowWindow(hTabCtl, SW_HIDE)
   Else
      nHeightTabControl = AfxGetWindowHeight(hTabCtl)
      SetWindowPos hTabCtl, 0, _
                     rc.Left, rc.Top + nHeightRebar, _
                     rc.Right, nHeightTabControl, _
                     SWP_SHOWWINDOW Or SWP_NOZORDER
   End If
   
   ' Reduce the height of the client area by the size of the rebar and statusbar.
   Dim pDoc As clsDocument Ptr = gTTabCtl.GetActiveDocumentPtr()
   If pDoc Then hEdit = pDoc->hWindow

   SetWindowPos hEdit, 0, _
                  rc.Left, rc.Top + nHeightRebar + nHeightTabControl, _
                  rc.Right, rc.Bottom - nHeightRebar - nHeightStatusBar - nHeightTabControl, _
                  SWP_SHOWWINDOW Or SWP_NOZORDER

   frmMain_SetFocusToCurrentCodeWindow

   Function = 0
End Function


' ========================================================================================
' Save project file to disk
' ========================================================================================
Function OnCommand_ProjectSave( ByVal HWnd As HWnd, ByVal bSaveAs As BOOLEAN = False) As LRESULT

   ' Do the actual saving to disk
   gpApp->SaveProject(bSaveAs)

   frmMain_SetFocusToCurrentCodeWindow

   Function = 0
End Function


' ========================================================================================
' Close currently active project
' ========================================================================================
Function OnCommand_ProjectClose( ByVal HWnd As HWnd ) As LRESULT

   OnCommand_ProjectSave( HWnd, False )
   
   ' Close all documents that may be open
   OnCommand_FileClose( HWnd, True )

   ' If any documents remain in the global collection then ensure that
   ' they are reset to not be part of a project (because the project is now closed)
   Dim As Long nCount = gpApp->GetDocumentCount 
   Dim pDoc As clsDocument Ptr
   Dim As Long i
   For i = 0 To nCount - 1
      pDoc = gpApp->GetDocumentPtr(i)
      If pDoc Then pDoc->IsProjectFile = False
   Next
   
   DestroyWindow( HWND_FRMFNLIST )
   DestroyWindow( HWND_FRMCOMPILERESULTS )
   DestroyWindow( HWND_FRMPROJECTMANAGER )
   gpApp->IsProjectActive = False

   frmMain_SetFocusToCurrentCodeWindow

   Function = 0
End Function


' ========================================================================================
' Create a new project
' ========================================================================================
Function OnCommand_ProjectNew( ByVal HWnd As HWnd ) As LRESULT

   ' If a project is already open then save/close it.
   If gpApp->IsProjectActive Then OnCommand_ProjectClose(HWnd)
   
   ' Ensure that the Project Manager window has been created (show it now as well)
   gpApp->IsNewProjectFlag = True
   frmProjectOptions_Show( HWnd, SW_SHOW )
   
   ' Ensure that the Project Manager window has been created (show it now as well)
   If gpApp->IsProjectActive Then    
      frmProjectManager_Show( HWnd, SW_SHOW )
   End If
   
   ' This will update the main window to show the project name in the window caption
   frmMain_SetFocusToCurrentCodeWindow
   
   Function = 0
End Function


' ========================================================================================
' Open a Project
' ========================================================================================
Function OnCommand_ProjectOpen( ByVal HWnd As HWnd ) As LRESULT

   ' Display the Open File Dialog
   Dim pwszName As WString Ptr = AfxIFileOpenDialogW(HWnd, IDM_PROJECTOPEN)
   If pwszName Then 
      ' If a project is already open then save/close it.
      If gpApp->IsProjectActive Then OnCommand_ProjectClose(HWnd)
      ' Open the project
      gpApp->ProjectDiskFilename = pwszName
      If gConfig.ProjectLoadFromFile() Then
         gpApp->IsNewProjectFlag = False
         gpApp->IsProjectActive = True
         ' Update the most recently used project list 
         UpdateMRUProjectList(pwszName)
      End If
      CoTaskMemFree(pwszName)
   End If

   frmMain_PositionWindows(HWnd)
   frmMain_SetFocusToCurrentCodeWindow

   Function = 0
End Function


' ========================================================================================
' Attempt to open specified file. If it exists then position to Tab if applicable
' ========================================================================================
Function frmMain_OpenFileSafely( ByVal HWnd        As HWnd, _
                                 ByVal bIsNewFile  As BOOLEAN, _ 
                                 ByVal bIsTemplate As BOOLEAN, _
                                 ByVal bShowInTab  As BOOLEAN, _
                                 ByVal pwszName    As WString Ptr, _
                                 ByVal pDocIn      As clsDocument Ptr _
                                 ) As clsDocument Ptr
   
   Dim As Long idx = -1
   Dim pDoc As clsDocument Ptr 
   
   If pDocIn = 0 Then   ' Create a new pDoc 
      pDoc = New clsDocument
      If bIsNewFile Then
         pDoc->CreateCodeWindow(HWnd, True)  ' Create the new Scintilla window
      Else
         pDoc->CreateCodeWindow(HWnd, False, bIsTemplate, pwszName)
      End If   
      gpApp->AddDocument(pDoc)            ' Add the new document to the global app
   Else
      pDoc = pDocIn     ' deal with the existing incoming pDoc
   End If   

   If bShowInTab Then
      ' If the document is already open and loaded then simply switch to
      ' that document in the top tabcontrol.
      idx = gTTabCtl.GetTabIndexFromFilename(pwszName)
      If idx = -1 Then idx = gTTabCtl.AddTab( pDoc )  ' Add the new document to the top tabcontrol
      gTTabCtl.SetFocusTab(idx)
   End If

   ' Position all of the controls into place
   frmMain_PositionWindows(HWnd)
   frmMain_SetFocusToCurrentCodeWindow

   Function = pDoc
End Function


' ========================================================================================
' Create a new editing window
' ========================================================================================
Function OnCommand_FileNew( ByVal HWnd As HWnd ) As LRESULT
   frmMain_OpenFileSafely(HWnd, True, False, True, 0, 0 )
   Function = 0
End Function


' ========================================================================================
' Open one or more files for editing
' ========================================================================================
Function OnCommand_FileOpen( ByVal HWnd As HWnd ) As LRESULT

   ' Display the Open File Dialog
   Dim pItems As IShellItemArray Ptr = AfxIFileOpenDialogMultiple(HWnd)
   If pItems = Null Then Exit Function
   Dim dwItemCount As Long, i As Long, idx As Long, pItem As IShellItem Ptr, pwszName As WString Ptr
   pItems->lpVtbl->GetCount(pItems, @dwItemCount)
   
   For i = 0 To dwItemCount - 1
      pItems->lpVtbl->GetItemAt(pItems, i, @pItem)
      If pItem Then
         pItem->lpVtbl->GetDisplayName(pItem, SIGDN_FILESYSPATH, @pwszName)
         If pwszName Then 
            frmMain_OpenFileSafely(HWnd, False, False, True, pwszName, 0 )
            CoTaskMemFree(pwszName)
            pwszName = Null
         End If
         pItem->lpVtbl->Release(pItem)
         pItem = Null
      End If
   Next
   pItems->lpVtbl->Release(pItems)
   
   frmMain_PositionWindows(HWnd)
   frmMain_SetFocusToCurrentCodeWindow

   Function = 0
End Function

      
' ========================================================================================
' Add one or more files to the active Project
' ========================================================================================
Function OnCommand_ProjectAddFiles( ByVal HWnd As HWnd ) As LRESULT

   ' Display the Open File Dialog
   Dim pItems As IShellItemArray Ptr = AfxIFileOpenDialogMultiple(HWnd)
   If pItems = Null Then Exit Function
   Dim dwItemCount As Long, i As Long, idx As Long, pItem As IShellItem Ptr, pwszName As WString Ptr
   pItems->lpVtbl->GetCount(pItems, @dwItemCount)
   
   For i = 0 To dwItemCount - 1
      pItems->lpVtbl->GetItemAt(pItems, i, @pItem)
      If pItem Then
         pItem->lpVtbl->GetDisplayName(pItem, SIGDN_FILESYSPATH, @pwszName)
         If pwszName Then 
            ' Check to see if the document is already loaded. If it is then check to see if it needs
            ' to have its flag set to make it part of the project. If not then simply add it.
            Dim pDoc As clsDocument Ptr 
            pDoc = gpApp->GetDocumentPtrByFilename(pwszName)
            If pDoc Then
               If pDoc->IsProjectFile = False Then
                  ' Add it to the Project Manager Listview
                  gpApp->ProjectAddFile(pDoc)     ' Add the document to the active Project & Listview     
               End If   
            Else
               ' Create the new Scintilla window and load the file
               pDoc = frmMain_OpenFileSafely(HWnd, False, False, False, pwszName, 0 )
               gpApp->ProjectAddFile(pDoc)     ' Add the document to the active Project & Listview     
            End If
            ' Ensure that the item displays based on the correct sorting
            Dim As HWnd hLV = GetDlgItem( HWND_FRMPROJECTMANAGER, IDC_FRMPROJECTMANAGER_LISTVIEW)
            Listview_SortColumn( hLV, gLVProjSortCol, gLVProjSortDir )
            CoTaskMemFree(pwszName)
            pwszName = Null
         End If
         pItem->lpVtbl->Release(pItem)
         pItem = Null
      End If
   Next
   pItems->lpVtbl->Release(pItems)
   
   frmMain_PositionWindows(HWnd)
   frmMain_SetFocusToCurrentCodeWindow

   Function = 0
End Function


' ========================================================================================
' Open #Include file that was selected from the right click popup menu
' ========================================================================================
Function OnCommand_OpenIncludeFile( ByVal HWnd As HWnd ) As LRESULT
   frmMain_OpenFileSafely(HWnd, False, False, True, gpApp->IncludeFilename, 0 )
   gpApp->IncludeFilename = ""
   Function = 0
End Function


' ========================================================================================
' Save current active file to disk
' ========================================================================================
Function OnCommand_FileSave( ByVal HWnd As HWnd, ByVal bSaveAs As BOOLEAN = False) As LRESULT

   Dim pDoc As clsDocument Ptr = gTTabCtl.GetActiveDocumentPtr()
   If pDoc = 0 Then Return 0

   ' Do the actual saving to disk
   If pDoc->SaveFile(bSaveAs) Then
      ' If a project is active then we may need to update the filename/path for
      ' this document if it is part of the project.
      If (gpApp->IsProjectActive = True) And (pDoc->IsProjectFile = True) Then
        gpApp->ProjectSetFileType( pDoc, pDoc->ProjectFileType )    ' this updates all fields in listview
      End If
      ' Apply document properties to this file because the file extension may have
      ' changed. For example from Untitled to *.bas
      pDoc->ApplyProperties
      ' Parse the code for sub/function listing and create list
      pDoc->CreateFunctionList
   EndIf   

   ' Ensure that the Tab displays the correct filename in case it was changed
   gTTabCtl.SetTabText(-1)
   frmMain_SetFocusToCurrentCodeWindow
   frmFnList_UpdateListBox()

   Function = 0
End Function


' ========================================================================================
' Save all open files to disk
' ========================================================================================
Function OnCommand_FileSaveAll( ByVal HWnd As HWnd ) As LRESULT

   Dim As Long nCount = TabCtrl_GetItemCount(gTTabCtl.hWindow) 
   If nCount = 0 Then Exit Function
   
   Dim pDoc As clsDocument Ptr 
   Dim i    As Long               
   Dim TB   As TC_ITEM 
   
   For i = 0 To nCount - 1
      ' Get the document pointer and then save file to disk
      TB.Mask = TCIF_PARAM
      TabCtrl_GetItem( gTTabCtl.hWindow, i, @TB )
      pDoc = Cast(clsDocument Ptr, TB.lParam)
      If SendMessageW( pDoc->hWindow, SCI_GETMODIFY, 0, 0 ) Then 
         ' If this is a new Untitled document that is also dirty then we
         ' need to set the focus to that tab because the SaveAs dialog will be shown.
         If pDoc->IsNewFlag Then
            gTTabCtl.SetFocusTab(i)
         End If
         pDoc->SaveFile()
         pDoc->DisplayStats()   ' update the tab filename
      End If   
   Next
   frmMain_SetFocusToCurrentCodeWindow
   frmFnList_UpdateListBox()

   Function = 0
End Function


' ========================================================================================
' Close current (or all) open file(s)
' ========================================================================================
Function OnCommand_FileClose( ByVal HWnd As HWnd, ByVal bCloseAll As BOOLEAN = False) As LRESULT

   Dim wText As WString * MAX_PATH  
   Dim r     As Long 
   
   Do
      Dim pDoc As clsDocument Ptr = gTTabCtl.GetActiveDocumentPtr()
      If pDoc = 0 Then Return 0
   
      If SendMessageW( pDoc->hWindow, SCI_GETMODIFY, 0, 0 ) Then
         FF_PATHNAME( @WStr("NAMEX"), pDoc->Diskfilename, @wText)
         r = MessageBoxW( HWnd, L(76,"Save current changes?") & WStr(" ") & wText, @WStr(APPNAME), _
                           MB_YESNOCANCEL Or MB_ICONQUESTION)
         If r = IDCANCEL Then Exit Function
         If r = IDYES Then 
            r = pDoc->SaveFile()
            If r = False Then Exit Function   ' save was cancelled
         End If
      End If
      
      ' Destroy the Scintilla window and remove from top tab control
      Dim As Long idx = TabCtrl_GetCurSel(gTTabCtl.hWindow)
      TabCtrl_DeleteItem(gTTabCtl.hWindow, idx)
      ' Set the active tab to the closest tab to the one just removed
      gTTabCtl.SetFocusTab(idx-1)
      
      ShowWindow pDoc->hWindow, SW_HIDE

      ' Only remove this document from the global collection if it is not
      ' part of any active Project.
      If pDoc->IsProjectFile = False Then gpApp->RemoveDocument(pDoc)
      
      frmMain_PositionWindows(HWnd)
   
      If bCloseAll = False Then Exit Do
      
      If bCloseAll Then
         If TabCtrl_GetItemCount(gTTabCtl.hWindow) = 0 Then Exit Do
      End If
   Loop
    
   Function = 0
End Function


' ========================================================================================
' Process WM_CREATE message for window/dialog: frmMain
' ========================================================================================
Function frmMain_OnCreate(ByVal HWnd As HWnd, ByVal lpCreateStructPtr As LPCREATESTRUCT) As BOOLEAN

   ' Enable drag and drop files
   DragAcceptFiles HWnd, CTRUE
   
   '  Message cracker macro expects a True to be returned for a successful
   '  OnCreate handler even though returning -1 from a standard WM_CREATE
   '  call would stop creating the window. This is just one of those Windows
   '  inconsistencies.
   Return True
End Function


' ========================================================================================
' Process WM_SIZE message for window/dialog: frmMain
' ========================================================================================
Function frmMain_OnSize(ByVal HWnd As HWnd, ByVal state As UINT, ByVal cx As Long, ByVal cy As Long) As LRESULT
   If state <> SIZE_MINIMIZED Then
      ' Update the size of the Rebar control
      SendMessageW GetDlgItem(HWnd, IDC_FRMMAIN_REBAR), WM_SIZE, cx, cy
      
      ' Resize the status bar
      Dim As HWnd hStatusBar = GetDlgItem(HWnd, IDC_FRMMAIN_STATUSBAR)
      SendMessageW hStatusBar, WM_SIZE, cx, cy
      InvalidateRect hStatusBar, Null, True
      
      ' Position all of the child windows
      frmMain_PositionWindows(HWnd)
   End If
   Function = 0
End Function


' ========================================================================================
' Process WM_COMMAND message for window/dialog: frmMain
' ========================================================================================
Function frmMain_OnCommand(ByVal HWnd As HWnd, ByVal id As Long, ByVal hwndCtl As HWnd, ByVal codeNotify As UINT) As LRESULT

   Dim As HWnd hEdit 
   
   Dim pDoc As clsDocument Ptr = gTTabCtl.GetActiveDocumentPtr()
   If pDoc Then 
      hEdit = pDoc->hWindow
      frmMain_ChangeToolbarButtonsState
   End If   

   Select Case id 
      
      ''  FILE MENU
      Case IDM_EXIT
         SendMessageW( HWnd, WM_CLOSE, 0, 0 )


      Case IDM_FILENEW:            Function = OnCommand_FileNew( HWnd )
      Case IDM_FILEOPEN:           Function = OnCommand_FileOpen( HWnd )
      Case IDM_FILESAVE:           Function = OnCommand_FileSave( HWnd )
      Case IDM_FILESAVEAS:         Function = OnCommand_FileSave( HWnd, True )
      Case IDM_FILESAVEALL:        Function = OnCommand_FileSaveAll( HWnd ) 
      Case IDM_FILECLOSE:          Function = OnCommand_FileClose( HWnd, False )
      Case IDM_FILECLOSEALL:       Function = OnCommand_FileClose( HWnd, True )
      
      Case IDC_MRUBASE To (IDC_MRUBASE + 10)
         Function = OpenMRUFile( HWnd, id )

      Case IDM_OPENINCLUDE:        Function = OnCommand_OpenIncludeFile( HWnd )
      
      Case IDM_COMMAND:            Shell(Environ("COMSPEC"))

      ''  EDIT MENU
      Case IDM_REDO:               Function = SendMessageW( hEdit, SCI_REDO, 0, 0)
      Case IDM_UNDO, WM_UNDO:      Function = SendMessageW( GetFocus(), Iif(GetFocus()=hEdit, SCI_UNDO, WM_UNDO), 0, 0)
      Case IDM_CUT, WM_CUT:        Function = SendMessageW( GetFocus(), Iif(GetFocus()=hEdit, SCI_CUT, WM_CUT), 0, 0)
      Case IDM_COPY, WM_COPY:      Function = SendMessageW( GetFocus(), Iif(GetFocus()=hEdit, SCI_COPY, WM_COPY), 0, 0)
      Case IDM_PASTE, WM_PASTE:    Function = SendMessageW( GetFocus(), Iif(GetFocus()=hEdit, SCI_PASTE, WM_PASTE), 0, 0)
      Case IDM_DELETELINE:         Function = SendMessageW(hEdit, SCI_LINECUT, 0, 0)
      Case IDM_INSERTFILE:         Function = pDoc->InsertFile()
      Case IDM_INDENTBLOCK:        Function = SendMessageW(hEdit, SCI_TAB, 0, 0)
      Case IDM_UNINDENTBLOCK:      Function = SendMessageW(hEdit, SCI_BACKTAB, 0, 0)
      Case IDM_COMMENTBLOCK:       Function = pDoc->BlockComment(True) 
      Case IDM_UNCOMMENTBLOCK:     Function = pDoc->BlockComment(False) 
      Case IDM_DUPLICATELINE:      Function = SendMessageW(hEdit, SCI_LINEDUPLICATE, 0, 0)
      Case IDM_MOVELINEUP:         Function = pDoc->MoveCurrentLines(False) 
      Case IDM_MOVELINEDOWN:       Function = pDoc->MoveCurrentLines(True)  
      Case IDM_TOUPPERCASE:        Function = pDoc->ChangeSelectionCase(1)
      Case IDM_TOLOWERCASE:        Function = pDoc->ChangeSelectionCase(2)
      Case IDM_TOMIXEDCASE:        Function = pDoc->ChangeSelectionCase(3)
      Case IDM_EOLTOCRLF:          Function = pDoc->ConvertEOL(SC_EOL_CRLF)
      Case IDM_EOLTOCR:            Function = pDoc->ConvertEOL(SC_EOL_CR)
      Case IDM_EOLTOLF:            Function = pDoc->ConvertEOL(SC_EOL_LF)
      'Case IDM_SPACESTOTABS:
      Case IDM_TABSTOSPACES:       Function = pDoc->TabsToSpaces()
      Case IDM_SELECTALL:          Function = SendMessageW(hEdit, SCI_SELECTALL, 0, 0)
      Case IDM_SELECTLINE:         Function = pDoc->SelectLine(-1)
                                   
      ''  SEARCH MENU              
      Case IDM_FIND:               Function = frmFind_Show(pDoc->hWindow)
      Case IDM_REPLACE:            Function = frmReplace_Show(pDoc->hWindow)

      Case IDM_FINDNEXT, IDM_FINDPREV
         gFind.txtFind = pDoc->GetSelText()
         Find_UpOrDown( id, 0, True, 0 )

      Case IDM_DEFINITION:         Function = frmMain_GotoDefinition( pDoc )
      Case IDM_LASTPOSITION:       Function = frmMain_GotoLastPosition()
      
      Case IDM_GOTO:               Function = frmGoto_Show(pDoc->hWindow)
      Case IDM_FUNCTIONLIST:       Function = frmFnList_Show( HWnd )    ' this will also toggle visibility 
      Case IDM_BOOKMARKTOGGLE:     Function = pDoc->ToggleBookmark(pDoc->GetCurrentLineNumber() )
      Case IDM_BOOKMARKNEXT:       Function = pDoc->NextBookmark()
      Case IDM_BOOKMARKPREV:       Function = pDoc->PrevBookmark()
      Case IDM_BOOKMARKCLEARALL:   Function = SendMessageW(hEdit, SCI_MARKERDELETEALL, -1, 0)
      
      ''  VIEW MENU
      Case IDM_FOLDTOGGLE:         Function = pDoc->FoldToggle( pDoc->GetCurrentLineNumber() )
      Case IDM_FOLDBELOW:          Function = pDoc->FoldToggleOnwards( pDoc->GetCurrentLineNumber() )
      Case IDM_FOLDALL:            Function = pDoc->FoldAll()
      Case IDM_UNFOLDALL:          Function = pDoc->UnFoldAll()
      Case IDM_ZOOMIN:             Function = SendMessageW(hEdit, SCI_ZOOMIN, 0, 0)
      Case IDM_ZOOMOUT:            Function = SendMessageW(hEdit, SCI_ZOOMOUT, 0, 0)

      Case IDM_RESTOREMAIN
         Dim rc As Rect
         SystemParametersInfo SPI_GETWORKAREA, 0, @rc, 0
         MoveWindow HWnd, 0, 0, rc.Right - rc.Left, rc.Bottom - rc.Top, CTRUE
         
      Case IDM_OPTIONS:            Function = frmOptions_Show( HWnd, SW_NORMAL )
      
      
      ''  PROJECT MENU
      Case IDM_PROJECTNEW:            Function = OnCommand_ProjectNew( HWnd )
      Case IDM_PROJECTSAVE:           Function = OnCommand_ProjectSave( HWnd, False )
      Case IDM_PROJECTSAVEAS:         Function = OnCommand_ProjectSave( HWnd, True )
      Case IDM_PROJECTFILESADD:       Function = OnCommand_ProjectAddFiles( HWnd )
      Case IDM_PROJECTOPEN:           Function = OnCommand_ProjectOpen( HWnd )

      Case IDC_MRUPROJECTBASE To (IDC_MRUPROJECTBASE + 10)
         Function = OpenMRUProjectFile( HWnd, id )

      Case IDM_PROJECTCLOSE:          Function = OnCommand_ProjectClose( HWnd )
      Case IDM_PROJECTMANAGER:        Function = frmProjectManager_Show( HWnd )    ' this will also toggle visibility 
      Case IDM_PROJECTOPTIONS:        Function = frmProjectOptions_Show( HWnd ) 
    
      Case IDM_ADDFILETOPROJECT:      
           Dim As HWnd hLV = GetDlgItem( HWND_FRMPROJECTMANAGER, IDC_FRMPROJECTMANAGER_LISTVIEW)
           Function = gpApp->ProjectAddFile( pDoc )
           Listview_SortColumn( hLV, gLVProjSortCol, gLVProjSortDir )

      Case IDM_REMOVEFILEFROMPROJECT: Function = gpApp->ProjectRemoveFile( pDoc )
      Case IDM_SETFILENORMAL:         Function = gpApp->ProjectSetFileType( pDoc, FILETYPE_NORMAL )
      Case IDM_SETFILEMODULE:         Function = gpApp->ProjectSetFileType( pDoc, FILETYPE_MODULE )
      Case IDM_SETFILEMAIN:           Function = gpApp->ProjectSetFileType( pDoc, FILETYPE_MAIN )
      Case IDM_SETFILERESOURCE:       Function = gpApp->ProjectSetFileType( pDoc, FILETYPE_RESOURCE )
         
      
      ''  COMPILE MENU          
      Case IDM_BUILDEXECUTE, IDM_COMPILE
         Function = code_Compile(id)

      Case IDM_RUNEXE
         Exec gConfig.LastRunFilename, gConfig.CommandLine
      
      Case IDM_COMPILERESULTS :       Function = frmCompileResults_Show( HWnd )    ' this will also toggle visibility 
      Case IDM_COMMANDLINE:           Function = frmCommandLine_Show( HWnd )

      Case IDM_USE32BIT, IDM_USE64BIT
         Function = frmMain_MenuSetCompiler(id)
      Case IDM_GUI, IDM_CONSOLE
         Function = frmMain_MenuSetCompileMode(id)
      
      
      Case IDM_HELP
      

      Case IDM_ABOUT
         MessageBoxW( HWnd, APPNAME & vbcrlf & "Version " & APPVERSION & _
                       vbcrlf & vbcrlf & _
                       "Paul Squires / PlanetSquires Software" & vbcrlf & vbcrlf & _
                       "Special thanks to:" & vbcrlf & _
                       "José Roca (Scintilla and CWindow)", _
                       @WStr("About"), _
                       MB_OK Or MB_ICONINFORMATION Or MB_DEFBUTTON1 Or MB_TASKMODAL )
 
         frmMain_SetFocusToCurrentCodeWindow()
        
   End Select

   Function = 0
End Function


' ========================================================================================
' Process WM_NOTIFY message for window/dialog: frmMain
' ========================================================================================
Function frmMain_OnNotify(ByVal HWnd As HWnd, ByVal id As Long, ByVal pNMHDR As NMHDR Ptr) As LRESULT

   Dim ptttdi As NMTTDISPINFOW Ptr    ' Tooltip notification message information
   Dim wszText As WString * MAX_PATH  
   Dim nCurSel As Long
   Dim pInfo As TCHITTESTINFO
   Dim hPopUpMenu As HMENU
   Dim pt As Point   
   Dim rc As Rect      
   
   If id = IDC_SCINTILLA Then
      ' Process Scintilla control notification messages
      Scintilla_OnNotify HWnd, Cast(SCNOTIFICATION Ptr, pNMHDR)
      Exit Function
   End If
   
   
   Select Case pNMHDR->code 
   
      Case NM_CLICK
         ' Process clicking on the "compiler" or "compilemode" panels in the statusbar
         If id = IDC_FRMMAIN_STATUSBAR Then   
            Dim lpnm As NMMOUSE Ptr
            lpnm = Cast(NMMOUSE Ptr, pNMHDR)
            If lpnm->dwItemSpec = 0 Then   ' compiler panel is panel 0
               Select Case Ucase(gConfig.DefaultCompiler)
                  Case "FBC 32BIT"  ' toggle from 32 bit to 64 bit
                     frmMain_MenuSetCompiler( IDM_USE64BIT )
                  Case "FBC 64BIT"  ' toggle from 64 bit to 32 bit
                     frmMain_MenuSetCompiler( IDM_USE32BIT )
               End Select
            End If
            If lpnm->dwItemSpec = 1 Then   ' compilemode panel is panel 1
               Select Case Ucase(gConfig.DefaultCompileMode)
                  Case "GUI"      ' toggle from GUI to CONSOLE
                     frmMain_MenuSetCompileMode( IDM_CONSOLE )
                  Case "CONSOLE"  ' toggle from CONSOLE to GUI
                     frmMain_MenuSetCompileMode( IDM_GUI )
               End Select
            End If
         End If

      Case NM_RCLICK
         If id = IDC_FRMMAIN_TOPTABCONTROL Then
            ' Get the tab number
            GetCursorPos( @pInfo.pt )
            ScreenToClient gTTabCtl.hWindow, @pInfo.pt
            nCurSel = TabCtrl_HitTest(gTTabCtl.hWindow, @pInfo)
            gTTabCtl.SetFocusTab(nCurSel)
            ' Create the popup menu
            hPopupMenu = CreateTopTabCtlContextMenu(nCurSel)
            GetCursorPos @pt
            TrackPopupMenu(hPopUpMenu, 0, pt.x, pt.y, 0, HWnd, ByVal Null)
            DestroyMenu hPopUpMenu
            Return True   ' prevent further processing that leads to WM_CONTEXTMENU
         End If

      Case TBN_DROPDOWN
         If id = IDC_FRMMAIN_TOOLBAR Then
            Dim ptbn As TBNOTIFY Ptr = Cast(TBNOTIFY Ptr, pNMHDR)
            Select Case ptbn->iItem
               Case IDM_FILENEW    ' Dropdown new file menu
                  SendMessageW(GetDlgItem(HWnd, IDC_FRMMAIN_REBAR), RB_GETRECT, 0, Cast(LPARAM, @rc))
                  MapWindowPoints(ptbn->hdr.hwndFrom, HWND_DESKTOP, Cast(Point Ptr, @rc), 2)
                  frmTemplates_Show(HWnd, rc.Left, rc.Bottom)
               Case IDM_FILEOPEN    ' Dropdown most recently used files
                  SendMessageW(GetDlgItem(HWnd, IDC_FRMMAIN_REBAR), RB_GETRECT, 0, Cast(LPARAM, @rc))
                  MapWindowPoints(ptbn->hdr.hwndFrom, HWND_DESKTOP, Cast(Point Ptr, @rc), 2)
                  Dim As Long nBtnWidth = Toolbar_GetButtonWidth(HWND_FRMMAIN_TOOLBAR)
                  hPopupMenu = CreateMRUpopup()
                  TrackPopupMenu hPopupMenu, 0, rc.Left + nBtnWidth, rc.bottom, 0, HWnd, 0
                  DestroyMenu hPopupMenu
               
            End Select
         End If

      Case TCN_SELCHANGING
         If id = IDC_FRMMAIN_TOPTABCONTROL Then
            nCurSel = TabCtrl_GetCurSel(pNMHDR->hwndFrom) 
            gTTabCtl.DisplayScintilla(nCurSel, False)
         End If

      Case TCN_SELCHANGE
         If id = IDC_FRMMAIN_TOPTABCONTROL Then
            nCurSel = TabCtrl_GetCurSel(pNMHDR->hwndFrom) 
            gTTabCtl.DisplayScintilla(nCurSel, True)
            frmMain_PositionWindows(HWnd)
         End If
         
      Case TTN_GETDISPINFO
         ptttdi = Cast(NMTTDISPINFOW Ptr, pNMHDR)
         ptttdi->hinst = Null
         wszText = ""
         Select Case ptttdi->hdr.hwndFrom
            Case SendMessageW(GetDlgItem(GetDlgItem(HWnd, IDC_FRMMAIN_REBAR), IDC_FRMMAIN_TOOLBAR), TB_GETTOOLTIPS, 0, 0)
               Select Case ptttdi->hdr.idFrom
                  Case IDM_FILENEW         : wszText = L(81,"New (Ctrl+N)")
                  Case IDM_FILEOPEN        : wszText = L(82,"Open (Ctrl+O)")
                  Case IDM_FILESAVE        : wszText = L(83,"Save (Ctrl+S)")
                  Case IDM_CUT             : wszText = L(84,"Cut (Ctrl+X)")
                  Case IDM_COPY            : wszText = L(85,"Copy (Ctrl+C)")
                  Case IDM_PASTE           : wszText = L(86,"Paste (Ctrl+V)")
                  Case IDM_FIND            : wszText = L(87,"Find (Ctrl+F)")
                  Case IDM_REPLACE         : wszText = L(88,"Replace (Ctrl+R)")
                  Case IDM_UNDO            : wszText = L(89,"Undo (Ctrl+Z)")
                  Case IDM_REDO            : wszText = L(90,"Redo (Ctrl+E)")
                  Case IDM_INDENTBLOCK     : wszText = L(91,"Indent (TAB)")
                  Case IDM_UNINDENTBLOCK   : wszText = L(92,"UnIndent (Shift+TAB)")
                  Case IDM_COMMENTBLOCK    : wszText = L(93,"Comment Block (Ctrl+B)")
                  Case IDM_UNCOMMENTBLOCK  : wszText = L(94,"UnComment Block (Ctrl+Shift+B)")
                  Case IDM_FUNCTIONLIST    : wszText = L(222,"Function List (F4)")
                  Case IDM_COMPILE         : wszText = L(96,"Compile (Ctrl+F5)")
                  Case IDM_BUILDEXECUTE    : wszText = L(97,"Compile and Execute (F5)")
                  Case IDM_RUNEXE          : wszText = L(98,"Run Executable (Shift+F5)")
                  Case IDM_HELP            : wszText = L(99,"Help (F1)")
               End Select
               If Len(wszText) Then ptttdi->lpszText = @wszText
         End Select
   End Select

   Function = 0
End Function


' ========================================================================================
' Process WM_ACTIVATEAPP message for window/dialog: frmMain
' ========================================================================================
Function frmMain_OnActivateApp(ByVal HWnd As HWnd, ByVal fActivate As BOOLEAN, ByVal dwThreadId As DWORD) As LRESULT

   ' If the application is gaining focus then determine if any of the loaded
   ' documents have been modified by an external application. If yes, then ask
   ' the user if wish to reload.
   If fActivate Then
      ' Search all loaded documents
      Dim As Long nTabCount = TabCtrl_GetItemCount(gTTabCtl.hWindow)           
      Dim As Long i
      Dim As Double nSerial
      Dim As String sText
      Dim pDoc As clsDocument Ptr
      
      For i = 0 To nTabCount - 1
         pDoc = gTTabCtl.GetDocumentPtr(i)
         If pDoc Then
            ' Bypass any 'new' untitled files.
            If pDoc->IsNewFlag Then Continue For
            
            ' Compare the disk file date time to the value currently
            ' stored in document class.
            nSerial = FileDateTime( *pDoc->DiskFilename )
            If nSerial <> pDoc->DateFileTime Then
               gTTabCtl.SetTabIndexByDocumentPtr(pDoc)
               If MessageBoxW( HWND_FRMMAIN, _
                               *pDoc->DiskFilename & vbCrLf & _ 
                               WStr("File was changed by another application. Reload it?"), _
                               WStr("File Changed"), MB_ICONQUESTION Or MB_YESNO) = IDYES Then
                  Dim As Long f = Freefile
                  If Open(*pDoc->DiskFilename For Binary As #f) = 0 Then
                     sText = Space(Lof(f))
                     Get #f, , sText
                     Close #f
                     pDoc->SetText( sText ) 
                     pDoc->DateFileTime = FileDateTime( *pDoc->DiskFilename )
                  End If
               End If                
            End If
         End If
      Next
      frmMain_SetFocusToCurrentCodeWindow()
   End If
   
   Function = 0
End Function


' ========================================================================================
' Process WM_CONTEXTMENU message for window/dialog: frmMain
' ========================================================================================
Function frmMain_OnContextMenu( ByVal HWnd As HWnd, _
                                ByVal hwndContext As HWnd, _
                                ByVal xPos As Long, _
                                ByVal yPos As Long ) As LRESULT

   Dim hPopupMenu as HMENU 
   Dim pt         As Point 
   
   SetFocus hWndContext
   pt.x = xPos
   pt.y = yPos
   ScreenToClient hWndContext, @pt
   hPopupMenu = CreateScintillaContextMenu()
   GetCursorPos @pt
   TrackPopupMenu hPopupMenu, 0, pt.x, pt.y, 0, HWnd, 0
   DestroyMenu hPopupMenu

   Function = 0
End Function


' ========================================================================================
' Process WM_DROPFILES message for window/dialog: frmMain
' ========================================================================================
Function frmMain_OnDropFiles( ByVal HWnd As HWnd, ByVal hDrop As HDROP ) As LRESULT

   ' Get the number of dropped files
   Dim As Long nCount = DragQueryFile(hDrop, &HFFFFFFFF, Null, 0)
   If nCount = 0 Then Exit Function
   
   Dim As Long i, nLen
   Dim wszPath As WString * MAX_PATH
   
   For i = 0 To nCount - 1
      nLen = DragQueryFile(hDrop, i, @wszPath, MAX_PATH)
      ' Make sure it's a file, not a folder
      Dim fd As WIN32_FIND_DATAW
      Dim hFind As HANDLE = FindFirstFileW(@wszPath, @fd)
      If hFind <> INVALID_HANDLE_VALUE Then
         FindClose hFind
         If (fd.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY Then
            gpApp->IncludeFilename = wszPath
            OnCommand_OpenIncludeFile(HWnd)
         End If
      End If
   Next
   
   DragFinish hDrop
   
   Function = 0
End Function


' ========================================================================================
' Process WM_CLOSE message for window/dialog: frmMain
' ========================================================================================
Function frmMain_OnClose(ByVal HWnd As HWnd) As LRESULT

   ' If a project is already open then save/close it.
   If gpApp->IsProjectActive Then OnCommand_ProjectClose(HWnd)

   ' Close any open files asking to save any that are dirty 
   OnCommand_FileClose(HWnd, True)

   ' If no open documents remain then call DestroyWindow to close the application
   If TabCtrl_GetItemCount(gTTabCtl.hWindow) = 0 Then 
      DestroyWindow(HWnd)
   End If   
   
   Function = 0
End Function


' ========================================================================================
' Process WM_DESTROY message for window/dialog: frmMain
' ========================================================================================
Function frmMain_OnDestroy(HWnd As HWnd) As LRESULT
   ImageList_Destroy Cast(HIMAGELIST, SendMessageW(GetDlgItem(HWnd, IDC_FRMMAIN_TOOLBAR), TB_SETIMAGELIST, 0, 0))
   ImageList_Destroy Cast(HIMAGELIST, SendMessageW(GetDlgItem(HWnd, IDC_FRMMAIN_TOOLBAR), TB_SETHOTIMAGELIST, 0, 0))
   ImageList_Destroy Cast(HIMAGELIST, SendMessageW(GetDlgItem(HWnd, IDC_FRMMAIN_TOOLBAR), TB_SETDISABLEDIMAGELIST, 0, 0))

   ' Output the config settings to disk file
   gConfig.SaveToFile

   ' Disable drag and drop files
   DragAcceptFiles HWnd, False
   
   PostQuitMessage(0)
   Function = 0
End Function


' ========================================================================================
' frmMain Window procedure
' ========================================================================================
Function frmMain_WndProc( ByVal HWnd   As HWnd, _
                          ByVal uMsg   As UINT, _
                          ByVal wParam As WPARAM, _
                          ByVal lParam As LPARAM _
                          ) As LRESULT

   Select Case uMsg
      HANDLE_MSG (HWnd, WM_CREATE,      frmMain_OnCreate)
      HANDLE_MSG (HWnd, WM_SIZE,        frmMain_OnSize)
      HANDLE_MSG (HWnd, WM_CLOSE,       frmMain_OnClose)
      HANDLE_MSG (HWnd, WM_DESTROY,     frmMain_OnDestroy)
      HANDLE_MSG (HWnd, WM_COMMAND,     frmMain_OnCommand)
      HANDLE_MSG (HWnd, WM_NOTIFY,      frmMain_OnNotify)
      HANDLE_MSG (HWnd, WM_ACTIVATEAPP, frmMain_OnActivateApp)
      HANDLE_MSG (HWnd, WM_CONTEXTMENU, frmMain_OnContextMenu)
      HANDLE_MSG (HWnd, WM_DROPFILES,   frmMain_OnDropFiles)

      Case WM_ERASEBKGND
         ' Code to prevent main application space flickering. Only brush the
         ' background if no tabs are open.
         If TabCtrl_GetItemCount(gTTabCtl.hWindow) > 0 Then 
            Return True
         End If
         
      Case WM_SYSCOMMAND
         If (wParam And &HFFF0) = SC_CLOSE Then
            SendMessageW HWnd, WM_CLOSE, 0, 0
         End If

      Case WM_SETFOCUS
         frmMain_SetFocusToCurrentCodeWindow

      Case WM_INITMENUPOPUP
         ' A top menu is about to be displayed. Determine what menu items to enable/disable. We
         ' do all of the topmenu items here rather than deal with individual popup submenus.
         frmMain_ChangeTopMenuStates()
         UpdateMRUMenu( Cast(HMENU, wParam) )
         UpdateMRUProjectMenu( Cast(HMENU, wParam) )

      ''  CUSTOM MESSAGE
      Case MSG_USER_SETFOCUS
         ' Set focus to current Scintilla window and update the document
         ' display such as Line#, Col#, Filename, etc.
         frmMain_UpdateLineCol(HWnd)

   End Select

   ' for messages that we don't deal with
   Function = DefWindowProcW(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' frmMain_Show
' ========================================================================================
Function frmMain_Show( ByVal hWndParent As HWnd, _
                       ByVal nCmdShow   As Long = 0 _
                       ) As Long

   '  Create the main window and child controls
   Dim pWindow As CWindow Ptr = New CWindow("WinFBE_Class")
   'pWindow->DPI = 96   ' eg. 144 or any other value (96 is default)
   
   ' Disable background erasing by only assigning the one style
   pWindow->ClassStyle = CS_DBLCLKS

   HWND_FRMMAIN = pWindow->Create(Null, APPNAME, @frmMain_WndProc )

   ' Set the small and large icon for the main window (must be set after main window is created)
   pWindow->BigIcon   =  LoadImage( pWindow->InstanceHandle, "IMAGE_MAINICON", IMAGE_ICON, 32, 32, LR_SHARED)
   pWindow->SmallIcon =  LoadImage( pWindow->InstanceHandle, "IMAGE_MAINICON", IMAGE_ICON, 16, 16, LR_SHARED)

   ' Create the main application topmenu
   HWND_FRMMAIN_TOPMENU = frmMain_BuildMenu(pWindow)
   SetMenu pWindow->hWindow, HWND_FRMMAIN_TOPMENU

   ' Create the main application toolbar
   frmMain_CreateToolbar(pWindow)
   frmMain_ChangeToolbarButtonsState

   ' Add a status bar
   Dim As HWnd hStatusbar = pWindow->AddControl("Statusbar", pWindow->hWindow, IDC_FRMMAIN_STATUSBAR)
   Dim As Long rgParts(4)
   rgParts(0) = pWindow->ScaleX(100)
   rgParts(1) = pWindow->ScaleX(200)
   rgParts(2) = pWindow->ScaleX(400)
   rgParts(3) = pWindow->ScaleX(600)
   rgParts(4) = -1
   If SendMessageW( hStatusBar, SB_SETPARTS, 5, Cast(LPARAM, @rgParts(0)) ) <> 0 Then
      SendMessageW( hStatusBar, SB_SIMPLE, False, 0 )
   End If

   ' Add the top tabcontrol (initially not visible)
   gTTabCtl.hWindow = pWindow->AddControl("TABCONTROL", pWindow->hWindow, _
       IDC_FRMMAIN_TOPTABCONTROL, "", 0, 0, 0, 24, _
       WS_CHILD Or WS_TABSTOP Or TCS_SINGLELINE Or TCS_RAGGEDRIGHT Or TCS_HOTTRACK Or TCS_TABS Or TCS_FOCUSNEVER, _
       WS_EX_LEFT Or WS_EX_LTRREADING)
     

   ' SET STARTUP POSITION
   ' If no valid window size exists then set to the default working area of the screen
   If (gConfig.StartupRight = 0) OrElse (gConfig.StartupBottom = 0) Then
      ' Retrieve the size of the working area
      Dim rc As Rect = pWindow->GetWorkArea                
      gConfig.StartupRight  = rc.Right
      gConfig.StartupBottom = rc.Bottom
   End If
   
   Dim WinPla As WINDOWPLACEMENT
   With WinPla
      .Length = Sizeof(WinPla)
      .rcNormalPosition.Left   = gConfig.StartupLeft
      .rcNormalPosition.Top    = gConfig.StartupTop
      .rcNormalPosition.Right  = gConfig.StartupRight
      .rcNormalPosition.Bottom = gConfig.StartupBottom
      .showCmd = Iif( gConfig.StartupMaximized, SW_MAXIMIZE, SW_SHOWNORMAL )
   End With
   SetWindowPlacement(pWindow->hWindow, @WinPla)
     
   ' Ensure the window is placed on screen should the user had changed 
   ' the logical ordering of a multiple display setup.
   AfxForceVisibleDisplay(pWindow->hWindow)
   
   UpdateWindow pWindow->hWindow
   
   ' Process windows events
   Dim uMsg As MSG
   
   ' Message loop
   Do While GetMessageW(@uMsg, Null, 0, 0)
      ' Processes accelerator keys for menu commands
      If (pWindow->AccelHandle = 0) Or (TranslateAcceleratorW(pWindow->hWindow, pWindow->AccelHandle, @uMsg)) = 0 Then
         ' Determines whether a message is intended for the specified
         ' dialog box and, if it is, processes the message.
         ' Ensure keystrokes like TAB are properly handled by the modeless dialogs
         If IsDialogMessageW( HWND_FRMFIND, @uMsg ) Then Continue Do
         If IsDialogMessageW( HWND_FRMREPLACE, @uMsg ) Then Continue Do
         If IsDialogMessageW(pWindow->hWindow, @uMsg) = 0 Then
            TranslateMessage @uMsg    ' Translates virtual-key messages into character messages.
            DispatchMessageW @uMsg    ' Dispatches a message to a window procedure.
         End If
      End If
   Loop
   Function = uMsg.wParam

   
   
   ' Delete the frmCompileResults CWindow class manually allocated memory 
   pWindow = AfxCWindowPtr(HWND_FRMCOMPILERESULTS):   Delete pWindow

   ' Delete the frmProjectManager CWindow class manually allocated memory 
   pWindow = AfxCWindowPtr(HWND_FRMPROJECTMANAGER):   Delete pWindow

   ' Delete the Main form CWindow class manually allocated memory 
   pWindow = AfxCWindowPtr(HWND_FRMMAIN):  Delete pWindow

End Function
