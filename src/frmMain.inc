
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
   
   nCount = gApp.GetDocumentCount()
   For i = 0 To nCount - 1
      pDocFind = gApp.GetDocumentPtr(i)
      If pDocFind Then 
         ' Search all of the names for a match on the current select word
         pList = pDocFind->FnListPtr
         Do Until pList = 0
            If Ucase(pList->zFnName) = sWord Then
               ' Save our current position in case user invokes "Last Position" option,   
               gLastPosition.pDoc = pDoc
               gLastPosition.nFirstLine = SendMessageW( pDoc->hWindow, SCI_GETFIRSTVISIBLELINE, 0, 0) 
               gLastPosition.nPosition = SendMessageW( pDoc->hWindow, SCI_GETCURRENTPOS, 0, 0) 
               frmMain_OpenFileSafely(HWND_FRMMAIN, _
                        False, _    ' bIsNewFile
                        False, _    ' bIsTemplate
                        True,  _    ' bShowInTab
                        false, _    ' bIsInclude
                        pDocFind->DiskFilename, _ ' pwszName
                        pDocFind )  ' pDocIn

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
   
   nCount = gApp.GetDocumentCount()
   For i = 0 To nCount - 1
      pDoc = gApp.GetDocumentPtr(i)
      If pDoc = gLastPosition.pDoc Then 
         ' Position ourselves to the correct previous position.
         frmMain_OpenFileSafely(HWND_FRMMAIN, _
                  False, _    ' bIsNewFile
                  False, _    ' bIsTemplate
                  True,  _    ' bShowInTab
                  false, _    ' bIsInclude
                  pDoc->DiskFilename, _ ' pwszName
                  pDoc )  ' pDocIn
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
   Static wText As WString * MAX_PATH
   
   ' Only update statusbar text if changed in order to prevent flicker
   StatusBar_GetText(hStatusbar, 0, @wText)
   If LTrim(wText) <> gConfig.DefaultCompiler Then
      StatusBar_SetText(hStatusbar, 0, " " & gConfig.DefaultCompiler)
   End If
   StatusBar_GetText(hStatusbar, 1, @wText)
   If LTrim(wText) <> gConfig.DefaultCompileMode Then
      StatusBar_SetText(hStatusbar, 1, " " & gConfig.DefaultCompileMode)
   End If

   ' If there is not any file being edited
   If pDoc = 0 Then
      StatusBar_SetText(hStatusbar, 2, "")    ' line/col/sel
      StatusBar_SetText(hStatusbar, 3, "")    ' file type (normal, main, resource, module)
      StatusBar_SetText(hStatusbar, 4, "")    ' file disktype (Unicode/Ansi)
      Dim wzText As WString * MAX_PATH
      wzText = APPNAME            
      If gApp.IsProjectActive Then wzText = APPNAMESHORT + " - [" + gApp.ProjectName + "]"
      SetWindowTextW( HWnd, wzText )
   Else
      pDoc->DisplayStats
      SetFocus(pDoc->hWindow)     
   End If
   frmMain_ChangeToolbarButtonsState

  Function = 0
End Function


' ========================================================================================
' Process any command line that was passed to the editor
' ========================================================================================
Function frmMain_ProcessCommandLine( ByVal HWnd As HWnd) As Long

   ' The incoming command line may contain a regular file to open or a project file.
   
   ' Command: A space-separated list of all command-line arguments is returned. When the 
   '          command line is parsed for arguments, everything between double quotes in 
   '          the parameter list will be considered as a single parameter, and is returned 
   '          with the double quotes.
   '          A value of zero (0) returns the name of the executable; and values of 
   '          one (1) and greater return each command-line argument.

   If Len(Command(1)) = 0 Then Exit Function
   
   Dim wszPath As WString * MAX_PATH
   Dim wszExt  As WString * MAX_PATH
   Dim wszArg  As WString * MAX_PATH
   
   Dim As Long i = 1
   Do
       wszArg = Command(i)
       If Len(wszArg) = 0 Then Exit Do
       
       ' Remove any double quotes from the argument.
       wszPath = AfxStrRemove( wszArg, WChr(34) )
       
       ' If no path exists for the file then add the current folder
       wszPath = AfxStrPathname( "PATH", wszArg )
       If Len(wszPath) = 0 Then wszArg = ExePath & "\" & wszArg 
       
       If Dir(wszArg) > "" Then
          ' We have a valid filename so determine what type it is.
          wszExt = AfxStrPathname( "EXTN", wszArg )
          wszExt = Ucase(wszExt)
          
          Select Case wszExt
             Case ".WFBE"    ' project file
                ' Pass the info to our generic project open function to handle everything.
                OpenMRUProjectFile( HWnd, 0, @wszArg )
             
             Case Else   ' .bas, .bi, .rc, etc...
               frmMain_OpenFileSafely(HWND, _
                                       False, _    ' bIsNewFile
                                       False, _    ' bIsTemplate
                                       True,  _    ' bShowInTab
                                       false, _    ' bIsInclude
                                       @wszArg, _  ' pwszName
                                       0 )          ' pDocIn
          End Select   
          
       End If
       
       i += 1
   Loop

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
                  buffer = AfxStrRemove(buffer, Chr(13))
                  buffer = AfxStrRemove(buffer, Chr(10))
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
   dim as long nHeightTabControl, nLeft
   
   ' Get the entire client area
   GetClientRect(HWnd, @rc)

   ' Get the heights of the rebar and statusbar
   Dim As Long nHeightRebar     = AfxGetWindowHeight( GetDlgItem(HWnd, IDC_FRMMAIN_REBAR) )
   dim as long nHeightStatusBar = AfxGetWindowHeight( GetDlgItem(HWnd, IDC_FRMMAIN_STATUSBAR) )
   Dim As Long nHeightOutput    = AfxGetWindowHeight( HWND_FRMOUTPUT)
   Dim As Long nWidthExplorer   = AfxGetWindowWidth( HWND_FRMEXPLORER )

   nLeft = rc.Left 
   
   ' Set the Explorer pane into place if applicable
   If IsWindowVisible(HWND_FRMEXPLORER) Then
      SetWindowPos( HWND_FRMEXPLORER, 0, _
                     nLeft, rc.Top + nHeightRebar, _
                     nWidthExplorer, _
                     rc.Bottom - nHeightStatusBar - nHeightRebar, _
                     SWP_NOZORDER Or SWP_SHOWWINDOW ) 
      PositionExplorerWindows(HWND_FRMEXPLORER)
      nLeft = nWidthExplorer
   Else
      nWidthExplorer = 0
   End If

   ' If items exist in the top tabcontrol then show the tab control and account for its height
   hTabCtl = GetDlgItem(HWnd, IDC_FRMMAIN_TOPTABCONTROL)
   If TabCtrl_GetItemCount(hTabCtl) = 0 Then
      ShowWindow(hTabCtl, SW_HIDE)
   Else
      nHeightTabControl = AfxGetWindowHeight(hTabCtl)
      SetWindowPos hTabCtl, 0, _
                     nLeft, rc.Top + nHeightRebar, _
                     rc.Right - nWidthExplorer, nHeightTabControl, _
                     SWP_SHOWWINDOW Or SWP_NOZORDER
   End If
   
   ' Set the Output pane into place if applicable
   If IsWindowVisible(HWND_FRMOUTPUT) Then
      SetWindowPos( HWND_FRMOUTPUT, 0, _
                     nLeft, rc.Bottom - nHeightStatusbar - nHeightOutput, _
                     rc.Right - nLeft, _
                     nHeightOutput, _
                     SWP_NOZORDER Or SWP_SHOWWINDOW ) 
   Else
      nHeightOutput = 0
   End If

   ' Reduce the height of the client area by the size of the rebar and statusbar.
   Dim pDoc As clsDocument Ptr = gTTabCtl.GetActiveDocumentPtr()
   If pDoc Then hEdit = pDoc->hWindow

   SetWindowPos hEdit, 0, _
                  nLeft, rc.Top + nHeightRebar + nHeightTabControl, _
                  rc.Right - nWidthExplorer, _
                  rc.Bottom - nHeightRebar - nHeightStatusBar - nHeightTabControl - nHeightOutput, _
                  SWP_SHOWWINDOW Or SWP_NOZORDER

   frmMain_SetFocusToCurrentCodeWindow

   Function = 0
End Function


' ========================================================================================
' Save project file to disk
' ========================================================================================
Function OnCommand_ProjectSave( ByVal HWnd As HWnd, ByVal bSaveAs As BOOLEAN = False) As LRESULT

   ' Do the actual saving to disk
   gApp.SaveProject(bSaveAs)

   frmMain_SetFocusToCurrentCodeWindow

   Function = 0
End Function


' ========================================================================================
' Close currently active project
' ========================================================================================
Function OnCommand_ProjectClose( ByVal HWnd As HWnd ) As LRESULT

   ' Save the project configuration to disk
   OnCommand_ProjectSave( HWnd, False )
   
   ' Close all documents that may be open (this closes any files that
   ' are open in Tabs. Non-open documents are handled after this).
   if OnCommand_FileClose( HWnd, True ) = 0 then return 0
          
   ' If any open tabs remain that contain a Project file then the
   ' close procedure must have been cancelled.
   Dim As Long i
   Dim pDoc As clsDocument Ptr 
   Dim As Long nCount = TabCtrl_GetItemCount(gTTabCtl.hWindow) 

   Dim TB As TC_ITEM 
   TB.Mask = TCIF_PARAM
   
   For i = 0 To nCount - 1
      TabCtrl_GetItem( gTTabCtl.hWindow, i, @TB )
      pDoc = Cast(clsDocument Ptr, TB.lParam)
      If pDoc->IsProjectFile then
         gApp.IsProjectActive = true
         Exit Function
      End If   
   Next
   
   ' So far, so good... no open tabs with a project file, so now destroy
   ' all other documents.
   gApp.RemoveAllDocuments() 

   ' Ensure no tabs remain
   TabCtrl_DeleteAllItems(gTTabCtl.hWindow)
   
   DestroyWindow( HWND_FRMFNLIST )

   ' Reset the Explorer window
   Dim As Hwnd hTree = GetDlgItem(HWND_FRMEXPLORER, IDC_FRMEXPLORER_TREE)
   TreeView_DeleteAllItems(hTree)
   gApp.hExplorerRootNode = 0
   
   gApp.IsProjectActive = False

   PositionExplorerWindows(HWND_FRMEXPLORER)
   
   frmMain_SetFocusToCurrentCodeWindow

   Function = true
End Function


' ========================================================================================
' Create a new project
' ========================================================================================
Function OnCommand_ProjectNew( ByVal HWnd As HWnd ) As LRESULT

   ' If a project is already open then save/close it.
   If gApp.IsProjectActive Then OnCommand_ProjectClose(HWnd)
   
   ' Ensure that the Project Manager window has been created (show it now as well)
   gApp.IsNewProjectFlag = True
   frmProjectOptions_Show( HWnd, SW_SHOW )
   PositionExplorerWindows(HWND_FRMEXPLORER)
   
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
      ' Pass the info to our generic project open function to handle everything.
      OpenMRUProjectFile( HWnd, 0, pwszName )
      CoTaskMemFree(pwszName)
   End If
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
                                 byval bIsInclude  as BOOLEAN, _
                                 ByVal pwszName    As WString Ptr, _
                                 ByVal pDocIn      As clsDocument Ptr _
                                 ) As clsDocument Ptr
   
   Dim As Long idx = -1
   Dim pDoc As clsDocument Ptr 
   
   ' Check to see if the document is already loaded. 
   pDoc = gApp.GetDocumentPtrByFilename(pwszName)
   if pDoc then pDocIn = pDoc
   
   If pDocIn = 0 Then   ' Create a new pDoc 
      pDoc = New clsDocument
      If bIsNewFile Then
         pDoc->CreateCodeWindow(HWnd, True)  ' Create the new Scintilla window
      Else
         pDoc->CreateCodeWindow(HWnd, false, bIsTemplate, pwszName)
      End If   
      gApp.AddDocument(pDoc)            ' Add the new document to the global app
   Else
      pDoc = pDocIn     ' deal with the existing pDoc
   End If   
   pwszName = @pDoc->DiskFilename

   ' Create the root node of the Explorer tree if it does not already exist.
   CreateRootNodeExplorerTreeview
   
   ' If a project is active then do not add any Include files to it and don't
   ' show it in the Explorer's project list.
   If gApp.IsProjectActive Then 
      if bIsInclude = false then
         AddFunctionsToExplorerTreeview( pDoc, False )
         gApp.ProjectAddFile(pDoc)   ' Add the document to the active Project   
      end if
   else
      ' no project is active so add the file to the FILES Explorer list.
      AddFunctionsToExplorerTreeview( pDoc, False )
   end if
         
   Dim As Hwnd hTree = GetDlgItem(HWND_FRMEXPLORER, IDC_FRMEXPLORER_TREE)
   Treeview_SortChildren( hTree, gApp.hExplorerRootNode, 0 )
   TreeView_Expand( hTree, gApp.hExplorerRootNode, TVE_EXPAND)

   If bShowInTab Then
      ' If the document is already open and loaded then simply switch to
      ' that document in the top tabcontrol.
      idx = gTTabCtl.GetTabIndexFromFilename(pwszName)
      If idx = -1 Then idx = gTTabCtl.AddTab( pDoc )  ' Add the new document to the top tabcontrol
      gTTabCtl.SetFocusTab(idx)
      Treeview_SelectItem( GetDlgItem(HWND_FRMEXPLORER, IDC_FRMEXPLORER_TREE), _
                           pDoc->hNodeExplorer )
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
   frmMain_OpenFileSafely(HWnd, _
                           True,  _   ' bIsNewFile
                           False, _   ' bIsTemplate
                           True,  _   ' bShowInTab
                           false, _   ' bIsInclude
                           0, _       ' pwszName
                           0 )        ' pDocIn
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
            frmMain_OpenFileSafely(HWnd, _
                                    False, _    ' bIsNewFile
                                    False, _    ' bIsTemplate
                                    True,  _    ' bShowInTab
                                    false, _    ' bIsInclude
                                    pwszName, _ ' pwszName
                                    0 )         ' pDocIn
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
   frmMain_OpenFileSafely(HWnd, _
                           False, _    ' bIsNewFile
                           False, _    ' bIsTemplate
                           True,  _    ' bShowInTab
                           true, _     ' bIsInclude
                           gApp.IncludeFilename, _ ' pwszName
                           0 )         ' pDocIn
   gApp.IncludeFilename = ""
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
      If (gApp.IsProjectActive = True) And (pDoc->IsProjectFile = True) Then
        gApp.ProjectSetFileType( pDoc, pDoc->ProjectFileType )    
      End If
      ' Apply document properties to this file because the file extension may have
      ' changed. For example from Untitled to *.bas
      pDoc->ApplyProperties
      pDoc->DisplayStats()   
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
         pDoc->DisplayStats()   
         gTTabCtl.SetTabText(i)
      End If   
   Next
   frmMain_SetFocusToCurrentCodeWindow
   frmFnList_UpdateListBox()

   Function = 0
End Function


' ========================================================================================
' Close current (or all) open file(s)
' ========================================================================================
Function OnCommand_FileClose( ByVal HWnd As HWnd, _
                              ByVal bCloseAll As BOOLEAN = False _
                              ) As LRESULT

   Dim wText As WString * MAX_PATH  
   Dim pDoc As clsDocument Ptr
   Dim As Long r, idx 
   
   Dim As Hwnd hTree = GetDlgItem(HWND_FRMEXPLORER, IDC_FRMEXPLORER_TREE)

   Do
      pDoc = gTTabCtl.GetActiveDocumentPtr()
      If pDoc = 0 Then Return true
   
      If SendMessageW( pDoc->hWindow, SCI_GETMODIFY, 0, 0 ) Then
         wText = AfxStrPathname( "NAMEX", pDoc->Diskfilename )
         r = MessageBoxW( HWnd, L(76,"Save current changes?") & WStr(" ") & wText, @WStr(APPNAME), _
                           MB_YESNOCANCEL Or MB_ICONQUESTION)
         If r = IDCANCEL Then Exit Function
         If r = IDYES Then 
            r = pDoc->SaveFile()
            If r = False Then Exit Function   ' save was cancelled
         End If
      End If
      
      ' Remove from top tab control
      idx = TabCtrl_GetCurSel(gTTabCtl.hWindow)
      TabCtrl_DeleteItem(gTTabCtl.hWindow, idx)
      ' Set the active tab to the closest tab to the one just removed
      gTTabCtl.SetFocusTab(idx-1)
      
      ShowWindow pDoc->hWindow, SW_HIDE

      ' Only remove this document from the global collection if it is not
      ' part of any active Project. Files that are part of a project are
      ' all closed (tabs/nodes) automatically in one shot rather than individually.
      if gApp.IsProjectActive = false then 
         ' Also remove it from the Explorer
         Treeview_DeleteItem( hTree, pDoc->hNodeExplorer)
         gApp.RemoveDocument(pDoc)
      end if   
      
      frmMain_PositionWindows(HWnd)
   
      If bCloseAll = False Then Exit Do
      
      If bCloseAll Then
         If TabCtrl_GetItemCount(gTTabCtl.hWindow) = 0 Then Exit Do
      End If
   Loop
    
   ' If no project is active and all files/tabs are closed then we
   ' simply reset our Explorer to its default state.
   if gApp.IsProjectActive = false then 
      If TabCtrl_GetItemCount(gTTabCtl.hWindow) = 0 Then 
         ' Reset the Explorer window
         TreeView_DeleteAllItems(hTree)
         gApp.hExplorerRootNode = 0
      end if
      PositionExplorerWindows(HWND_FRMEXPLORER)
   end if
   
   Function = true
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
      
      Case IDM_NEXTTAB:            Function = gTTabCtl.NextTab()
      Case IDM_PREVTAB:            Function = gTTabCtl.PrevTab()
      Case IDM_CLOSETAB:           Function = gTTabCtl.CloseTab()

      ''  FILE MENU
      Case IDM_FILENEW:            Function = OnCommand_FileNew( HWnd )
      Case IDM_FILEOPEN:           Function = OnCommand_FileOpen( HWnd )
      Case IDM_FILESAVE:           Function = OnCommand_FileSave( HWnd )
      Case IDM_FILESAVEAS:         Function = OnCommand_FileSave( HWnd, True )
      Case IDM_FILESAVEALL:        Function = OnCommand_FileSaveAll( HWnd ) 
      Case IDM_FILECLOSE:          Function = OnCommand_FileClose( HWnd, False )
      Case IDM_FILECLOSEALL:       Function = OnCommand_FileClose( HWnd, True )
      
      Case IDC_MRUBASE To (IDC_MRUBASE + 10)
         Function = OpenMRUFile( HWnd, id )

      Case IDM_MRUCLEAR, IDM_MRUPROJECTCLEAR
         Function = ClearMRUlist( id )
         
      Case IDM_OPENINCLUDE:        Function = OnCommand_OpenIncludeFile( HWnd )
      Case IDM_COMMAND:            Shell(Environ("COMSPEC"))
      Case IDM_EXIT:               SendMessageW( HWnd, WM_CLOSE, 0, 0 )

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
         Dim As String sFindText = pDoc->GetSelText()
         If Len(sFindText) Then gFind.txtFind = sFindText
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

      Case IDM_VIEWEXPLORER
         ShowWindow HWND_FRMEXPLORER, Iif(IsWindowVisible(HWND_FRMEXPLORER), SW_HIDE, SW_SHOW)
         frmMain_PositionWindows(HWND_FRMMAIN)
      Case IDM_VIEWOUTPUT
         ShowWindow HWND_FRMOUTPUT, Iif(IsWindowVisible(HWND_FRMOUTPUT), SW_HIDE, SW_SHOW)
         frmMain_PositionWindows(HWND_FRMMAIN)
      Case IDM_RESTOREMAIN
         Dim rc As Rect
         SystemParametersInfo SPI_GETWORKAREA, 0, @rc, 0
         MoveWindow HWnd, 0, 0, rc.Right - rc.Left, rc.Bottom - rc.Top, CTRUE
         
      Case IDM_OPTIONS:            Function = frmOptions_Show( HWnd, SW_NORMAL )
      
      
      ''  PROJECT MENU
      Case IDM_PROJECTNEW:            Function = OnCommand_ProjectNew( HWnd )
      Case IDM_PROJECTSAVE:           Function = OnCommand_ProjectSave( HWnd, False )
      Case IDM_PROJECTSAVEAS:         Function = OnCommand_ProjectSave( HWnd, True )
      Case IDM_PROJECTFILESADD:       Function = OnCommand_FileOpen( HWnd )
      Case IDM_PROJECTOPEN:           Function = OnCommand_ProjectOpen( HWnd )

      Case IDC_MRUPROJECTBASE To (IDC_MRUPROJECTBASE + 10)
         Function = OpenMRUProjectFile( HWnd, id )

      Case IDM_PROJECTCLOSE:          Function = OnCommand_ProjectClose( HWnd )
      Case IDM_PROJECTOPTIONS:        Function = frmProjectOptions_Show( HWnd ) 
    
      Case IDM_ADDFILETOPROJECT:      
         ' Add the currently displayed document to the project.
         frmMain_OpenFileSafely(HWnd, _
                                 False, _    ' bIsNewFile
                                 False, _    ' bIsTemplate
                                 True,  _    ' bShowInTab
                                 false, _    ' bIsInclude
                                 0, _        ' pwszName
                                 pDoc )      ' pDocIn

      Case IDM_REMOVEFILEFROMPROJECT: Function = gApp.ProjectRemoveFile( pDoc )
      Case IDM_SETFILENORMAL:         Function = gApp.ProjectSetFileType( pDoc, FILETYPE_NORMAL )
      Case IDM_SETFILEMODULE:         Function = gApp.ProjectSetFileType( pDoc, FILETYPE_MODULE )
      Case IDM_SETFILEMAIN:           Function = gApp.ProjectSetFileType( pDoc, FILETYPE_MAIN )
      Case IDM_SETFILERESOURCE:       Function = gApp.ProjectSetFileType( pDoc, FILETYPE_RESOURCE )
         
           
      ''  COMPILE MENU          
      Case IDM_BUILDEXECUTE, IDM_COMPILE
         Function = code_Compile(id)

      Case IDM_RUNEXE
         RunExe gConfig.LastRunFilename, gConfig.CommandLine
      
      Case IDM_COMMANDLINE:           Function = frmCommandLine_Show( HWnd )

      Case IDM_USE32BIT, IDM_USE64BIT
         Function = frmMain_MenuSetCompiler(id)
      Case IDM_GUI, IDM_CONSOLE
         Function = frmMain_MenuSetCompileMode(id)
      
      
      Case IDM_HELP
         ShowContextHelp() 
      

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
      ' Process Scintilla control notification messages. First, check to see if the notifications
      ' have been suppressed by a bulk modification activity that should only update the screen
      ' at the end of its process (eg. moving large numbers of selected lines up or down).
      If gApp.SuppressNotify Then Exit Function
      Scintilla_OnNotify HWnd, Cast(SCNOTIFICATION Ptr, pNMHDR)
      Exit Function
   End If
   
   
   Select Case pNMHDR->code 
   
      Case NM_CLICK
         If id = IDC_FRMMAIN_TOPTABCONTROL Then
            ' Get the position of the hit and determine if the close icon was clicked.
            GetCursorPos( @pInfo.pt )
            ScreenToClient gTTabCtl.hWindow, @pInfo.pt
            nCurSel = TabCtrl_HitTest(gTTabCtl.hWindow, @pInfo)
            If pInfo.flags = TCHT_ONITEMICON Then gTTabCtl.CloseTab()
         End If
         
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
            If lpnm->dwItemSpec = 2 Then   ' line/column panel is panel 2
               dim pDoc as clsDocument ptr
               pDoc = gTTAbCtl.GetActiveDocumentPtr()
               If pDoc Then Function = frmGoto_Show(pDoc->hWindow)
            end if
            If lpnm->dwItemSpec = 3 Then   ' ProjectFiletype is panel 3
               hPopupMenu = CreateStatusBarFileTypeContextMenu
               GetCursorPos @pt
               TrackPopupMenu(hPopUpMenu, 0, pt.x, pt.y, 0, HWnd, ByVal Null)
               DestroyMenu hPopUpMenu
            end if
            If lpnm->dwItemSpec = 4 Then   ' Ansi/Unicode is panel 4
               ' Toggle the ANSI/Unicode state
               gConfig.UnicodeEncoding = iif( gConfig.UnicodeEncoding, false, true )
               frmMain_UpdateLineCol( HWND_FRMMAIN )
            end if   
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
            dim pDoc as clsDocument ptr
            pDoc = gTTAbCtl.GetActiveDocumentPtr()
            If pDoc Then
               Treeview_SelectItem( GetDlgItem(HWND_FRMEXPLORER, IDC_FRMEXPLORER_TREE), _
                                    pDoc->hNodeExplorer )
            end if
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
      Dim As FILETIME ft
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
            ft = AfxGetFileLastWriteTime(pDoc->DiskFilename)

            If AfxFileTimeToVariantTime(ft) <> AfxFileTimeToVariantTime(pDoc->DateFileTime) Then
               gTTabCtl.SetTabIndexByDocumentPtr(pDoc)
               If MessageBoxW( HWND_FRMMAIN, _
                               pDoc->DiskFilename & vbCrLf & _ 
                               L(245, "File was changed by another application. Reload it?"), _
                               L(246, "File Changed"), MB_ICONQUESTION Or MB_YESNO) = IDYES Then
                  Dim As Long f = Freefile
                  If Open(pDoc->DiskFilename For Binary As #f) = 0 Then
                     sText = Space(Lof(f))
                     Get #f, , sText
                     Close #f
                     Dim nResult As Long = IS_TEXT_UNICODE_SIGNATURE
                     If IsTextUnicode(StrPtr(sText), 2, Cast(LPINT, @nResult) ) Then
                        sText = Mid(sText, 3)
                        sText = AfxACode( Cast(WSTRING Ptr, StrPtr(sText)) )
                     End If   
                     pDoc->SetText( sText ) 
                     pDoc->DateFileTime = AfxGetFileLastWriteTime(pDoc->DiskFilename)
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
            gApp.IncludeFilename = wszPath
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

   gConfig.ShowExplorer = IsWindowVisible(HWND_FRMEXPLORER)
   
   ' Prevent Treeview flickering during closing
   dim as hwnd hTV = GetDlgItem(HWND_FRMEXPLORER, IDC_FRMEXPLORER_TREE)
   LockWindowUpdate(hTV)
   
   ' If a project is already open then save/close it.
   If gApp.IsProjectActive Then 
      if OnCommand_ProjectClose(HWnd) = 0 then 
         LockWindowUpdate(0): return 0
      end if   
   else      
      ' Close any open files asking to save any that are dirty 
      if OnCommand_FileClose(HWnd, True) = 0 then 
         LockWindowUpdate(0): return 0
      end if   
   end if
   
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
   ImageList_Destroy Cast(HIMAGELIST, SendMessageW(GetDlgItem(HWnd, IDC_FRMMAIN_TOOLBAR), TB_GETIMAGELIST, 0, 0))
   ImageList_Destroy Cast(HIMAGELIST, SendMessageW(GetDlgItem(HWnd, IDC_FRMMAIN_TOOLBAR), TB_GETHOTIMAGELIST, 0, 0))
   ImageList_Destroy Cast(HIMAGELIST, SendMessageW(GetDlgItem(HWnd, IDC_FRMMAIN_TOOLBAR), TB_GETDISABLEDIMAGELIST, 0, 0))
   ImageList_Destroy TabCtrl_GetImageList(gTTabCtl.hWindow) 
   
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
            Return True      ' prevents painting the background
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

      ''  CUSTOM MESSAGES
      Case MSG_USER_SETFOCUS
         ' Set focus to current Scintilla window and update the document
         ' display such as Line#, Col#, Filename, etc.
         frmMain_UpdateLineCol(HWnd)

      Case MSG_USER_PROCESS_COMMANDLINE 
         ' process any command line arguments that may have been passed to the program.
         frmMain_ProcessCommandLine(HWnd)


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
   pWindow->BigIcon   =  LoadImage( pWindow->InstanceHandle, "IMAGE_AAA_MAINICON", IMAGE_ICON, 32, 32, LR_SHARED)
   pWindow->SmallIcon =  LoadImage( pWindow->InstanceHandle, "IMAGE_AAA_MAINICON", IMAGE_ICON, 16, 16, LR_SHARED)

   ' Create the main application topmenu
   HWND_FRMMAIN_TOPMENU = frmMain_BuildMenu(pWindow)
   SetMenu pWindow->hWindow, HWND_FRMMAIN_TOPMENU

   ' Create the main application toolbar
   frmMain_CreateToolbar(pWindow)
   frmMain_ChangeToolbarButtonsState

   ' Add a status bar
   Dim As HWnd hStatusbar = pWindow->AddControl("Statusbar", pWindow->hWindow, IDC_FRMMAIN_STATUSBAR)
   Dim As Long rgParts(5)
   rgParts(0) = pWindow->ScaleX(100)   ' 32/64bit
   rgParts(1) = pWindow->ScaleX(200)   ' GUI/Console
   rgParts(2) = pWindow->ScaleX(400)   ' Line/Col
   rgParts(3) = pWindow->ScaleX(500)   ' Normal/Module/Main/Resource
   rgParts(4) = pWindow->ScaleX(600)   ' Unicode/ANSI
   rgParts(5) = -1
   If SendMessageW( hStatusBar, SB_SETPARTS, 6, Cast(LPARAM, @rgParts(0)) ) <> 0 Then
      SendMessageW( hStatusBar, SB_SIMPLE, False, 0 )
   End If

   ' Add the top tabcontrol (initially not visible)
   gTTabCtl.hWindow = pWindow->AddControl("TABCONTROL", pWindow->hWindow, _
       IDC_FRMMAIN_TOPTABCONTROL, "", 0, 0, 0, 24, _
       WS_CHILD Or WS_TABSTOP Or TCS_SINGLELINE Or TCS_RAGGEDRIGHT Or TCS_HOTTRACK Or _
       TCS_TABS Or TCS_FOCUSNEVER Or TCS_FORCEICONLEFT, WS_EX_LEFT Or WS_EX_LTRREADING)
     
   ' Create the imagelist for the TopTabControl tabs
   Dim As HICON hIcon = LoadIcon(pWindow->InstanceHandle, "IMAGE_TABCLOSE")
   Dim As HANDLE hImageList = ImageList_Create(16, 16, ILC_MASK Or ILC_COLOR32, 1, 0)
   ImageList_AddIcon(hImageList, hIcon)
   DestroyIcon( hIcon )
   TabCtrl_SetImageList(gTTabCtl.hWindow, hImageList)
   
   ' Create the ImageLists that will hold the images for the Explorer and Recent treeviews
   Dim cx As Long = 24 * (pWindow->DPI \ 96)
   ghImageListNormal = ImageList_Create( cx, cx, ILC_COLOR32 Or ILC_MASK, 4, 1)
   gidxImageBlank  = AfxGdipAddIconFromRes(ghImageListNormal, pWindow->InstanceHandle, "IMAGE_BLANK")
   gidxImageClosed = AfxGdipAddIconFromRes(ghImageListNormal, pWindow->InstanceHandle, "IMAGE_CLOSED")
   gidxImageOpened = AfxGdipAddIconFromRes(ghImageListNormal, pWindow->InstanceHandle, "IMAGE_OPENED")
   gidxImageCode   = AfxGdipAddIconFromRes(ghImageListNormal, pWindow->InstanceHandle, "IMAGE_CODE")

   ' Create the various child windows
   frmExplorer_Show( HWND_FRMMAIN )
   frmRecent_Show( HWND_FRMEXPLORER )     ' child of frmExplorer
   frmOutput_Show( HWND_FRMMAIN )

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

   ' Post a message to process the application's command line as applicable.
   PostMessage( pWindow->hWindow, MSG_USER_PROCESS_COMMANDLINE , 0, 0 )

   
   ' Process windows events
   Dim uMsg As MSG
   
   ' Message loop
   Do While GetMessageW(@uMsg, Null, 0, 0)
      ' Processes accelerator keys for menu commands
      If (pWindow->AccelHandle = 0) Or (TranslateAcceleratorW(pWindow->hWindow, pWindow->AccelHandle, @uMsg)) = 0 Then

         ' Intercept all MouseMove messages in order to test if we are over the splitter area. This
         ' easier and more efficent than trying to deal with mouse messages at the child control/form level.
         Select case uMsg.message
            Case WM_LBUTTONDOWN:   Splitter_OnLButtonDown
            Case WM_LBUTTONUP:     Splitter_OnLButtonUp
            Case WM_MOUSEMOVE:     Splitter_OnMouseMove
         End Select      

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

   
   ' Destroy the Explorer and Recent treeview imagelist
   ImageList_Destroy(ghImageListNormal)

   ' Delete the allocated memory for the various child windows
   pWindow = AfxCWindowPtr(HWND_FRMEXPLORER):       Delete pWindow
   pWindow = AfxCWindowPtr(HWND_FRMRECENT):         Delete pWindow
   pWindow = AfxCWindowPtr(HWND_FRMOUTPUT):         Delete pWindow
   pWindow = AfxCWindowPtr(HWND_FRMMAIN):           Delete pWindow

End Function






