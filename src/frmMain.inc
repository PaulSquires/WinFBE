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
' Position editor to document/function under the current caret position
' ========================================================================================
Function frmMain_GotoDefinition( ByVal pDoc As clsDocument Ptr ) As Long

   If pDoc = 0 Then Exit Function
   
   dim pData as DB2_DATA ptr
    
   dim as hwnd hEdit = pDoc->hWndActiveScintilla

   ' Determine the word at the current caret position.
   Dim As String sWord = pDoc->GetWord
    
   ' Search the list of Function names
   pData = gdb2.dbFindFunction(sWord)
   if pData THEN
      ' Save our current position in case user invokes "Last Position" option,   
      gLastPosition.pDoc = pDoc
      gLastPosition.nFirstLine = SciExec( hEdit, SCI_GETFIRSTVISIBLELINE, 0, 0) 
      gLastPosition.nPosition = SciExec( hEdit, SCI_GETCURRENTPOS, 0, 0) 
      OpenSelectedDocument(0, sWord)
   else
      MessageBox( HWND_FRMMAIN, L(224,"Sub/Function definition not found."), @WStr(APPNAME), MB_ICONWARNING)
   END IF
   
   Function = 0
End Function


' ========================================================================================
' Return to most previous position (invoked when "Last Position" is selected
' ========================================================================================
Function frmMain_GotoLastPosition() As Long

   ' Is the pDoc pointer still valid
   If gLastPosition.pDoc = 0 Then Exit Function
   
   ' Position ourselves to the correct previous position.
   frmMain_OpenFileSafely(HWND_FRMMAIN, _
            False, _    ' bIsNewFile
            False, _    ' bIsTemplate
            True,  _    ' bShowInTab
            false, _    ' bIsInclude
            gLastPosition.pDoc->DiskFilename, _ ' pwszName
            gLastPosition.pDoc )  ' pDocIn
   SciExec( GetFocus, SCI_SETFIRSTVISIBLELINE, gLastPosition.nFirstLine, 0) 
   SciExec( GetFocus, SCI_GOTOPOS, gLastPosition.nPosition, 0) 
   
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
   if IsWindowVisible(HWND_FRMRECENT) THEN
      StatusBar_SetText(hStatusbar, 0, "")    ' Build configuration
   else
      dim as long nCurSel = ComboBox_GetCurSel(HWND_FRMMAIN_COMBOBUILDS)
      if nCurSel > -1 THEN
         StatusBar_GetText(hStatusbar, 0, @wText)
         If LTrim(wText) <> gConfig.Builds(nCurSel).wszDescription Then
            StatusBar_SetText(hStatusbar, 0, " " & gConfig.Builds(nCurSel).wszDescription)
         End If
      else
         StatusBar_SetText(hStatusbar, 0, "")    ' Build configuration
      END IF
   END IF

   ' If there is not any file being edited
   If pDoc = 0 Then
      StatusBar_SetText(hStatusbar, 1, "")    ' line/col/sel
      StatusBar_SetText(hStatusbar, 2, "")    ' file type (normal, main, resource, module)
      StatusBar_SetText(hStatusbar, 3, "")    ' file disktype (Unicode/Ansi)
      StatusBar_SetText(hStatusbar, 4, "")    ' compile messages
      Dim wzText As WString * MAX_PATH
      wzText = APPNAME            
      If gApp.IsProjectActive Then
         dim idx as long = gApp.GetActiveProjectIndex
         if idx then wzText = APPNAMESHORT + " - [" + gApp.Projects(idx).ProjectName + "]"
      end if
      SetWindowText( HWnd, wzText )
   Else
      pDoc->DisplayStats
   End If
   frmMain_ChangeToolbarButtonsState
   LoadBuildComboBox

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
   Dim wszArg  As WString * MAX_PATH
   dim DataToSend AS COPYDATASTRUCT   
   
   IF IsIconic(hwnd) THEN 
      Dim WinPla As WINDOWPLACEMENT
      With WinPla
         .Length = Sizeof(WinPla)
         .rcNormalPosition.Left   = gConfig.StartupLeft
         .rcNormalPosition.Top    = gConfig.StartupTop
         .rcNormalPosition.Right  = gConfig.StartupRight
         .rcNormalPosition.Bottom = gConfig.StartupBottom
         .showCmd = Iif( gConfig.StartupMaximized, SW_MAXIMIZE, SW_SHOWNORMAL )
      End With
      SetWindowPlacement(HWND, @WinPla)
   end if
   SetForegroundWindow(hwnd)
   
   Dim As Long i = 1
   Do
      wszArg = Command(i)
      If Len(wszArg) = 0 Then Exit Do
       
      ' Remove any double quotes from the argument.
      wszPath = AfxStrRemove( wszArg, WChr(34) )
       
      ' If no path exists for the file then add the current folder
      wszPath = AfxStrPathname( "PATH", wszArg )
      If Len(wszPath) = 0 Then wszArg = AfxGetExePathName & wszArg 
       
      If AfxFileExists(wszArg) Then
         DataToSend.lpData  = @wszArg
         DataToSend.cbdata  = (LEN(wszArg)*2) + 1
         SendMessage(hwnd, WM_COPYDATA, LEN(DataToSend), cast(lParam, @DataToSend))
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
   
   Dim pDoc As clsDocument Ptr 

   Dim as HWND hEdit
   Dim as long nLine, nFoldLevel

   Select Case pNSC->hdr.code
         
      Case SCN_UPDATEUI    
         if gApp.SuppressNotify then return true
         If pNSC->hdr.idFrom = IDC_SCINTILLA Then
            pDoc = gTTabCtl.GetActiveDocumentPtr()
            If pDoc Then 
               ' This code will position our substitute scrollbar thumbs only when the 
               ' caret position crosses a page boundry via keyboard input
               If pNSC->updated And SC_UPDATE_SELECTION And pNSC->updated > SC_UPDATE_SELECTION Then
                  pDoc->ScrInfo.nPos = SciExec(pDoc->hWindow(0), SCI_GETFIRSTVISIBLELINE, 0, 0)
                  SetScrollInfo( pDoc->hScrollBar, SB_CTL, @pDoc->ScrInfo, TRUE)
               End If

               Select Case pNSC->updated
               Case SC_UPDATE_V_SCROLL
                  ' If there's only a V scroll update w/no selection update, then mousewheel was used - adj thumb pos.
                  ' Scintilla eats the WM_VSCROLL msg from a mousewheel so we never get one to process. We do it here.
                  pDoc->ScrInfo.nPos = SciExec(pDoc->hWindow(0), SCI_GETFIRSTVISIBLELINE, 0, 0)
                  SetScrollInfo( pDoc->hScrollBar, SB_CTL, @pDoc->ScrInfo, TRUE)
               end select
            end if
            frmMain_UpdateLineCol(HWnd)
         end if
         return true
         
            
      Case SCN_MODIFIED
         if gApp.SuppressNotify then return true
         ' Show line and column. Only do on modification of text otherwise we will have
         ' a huge slowdown when notifications sent for UI updates.
         If (pNSC->modificationType And SC_MOD_INSERTTEXT) Or (pNSC->modificationType And SC_MOD_DELETETEXT) Then
            frmMain_UpdateLineCol(HWnd)
         end if
         return true
         
      Case SCN_MARGINCLICK
         ' Folder margin
         pDoc = gTTabCtl.GetActiveDocumentPtr()
         If pDoc Then 
            hEdit      = pDoc->hWndActiveScintilla
            nLine      = SciExec(hEdit, SCI_LINEFROMPOSITION, pNSC->position, 0)
            nFoldLevel = SciExec(hEdit, SCI_GETFOLDLEVEL, nLine, 0)
            Select Case pNSC->margin 
               Case 1   ' left margin (bookmarks)
                  pDoc->ToggleBookmark(nLine)
               Case 2   ' fold margin
                  ' If is the head line...
                  If (nFoldLevel And SC_FOLDLEVELHEADERFLAG) <> 0 Then
                     SciExec(hEdit, SCI_TOGGLEFOLD, nLine, 0) 
                  End If   
            End Select
         End If           
                     
               
      case SCN_AUTOCSELECTIONCHANGE
         ' A listbox line has changed.
         ' Iterate all windows to find the popup Autocomplete listbox (need to find the
         ' handle here because the selection change fires when the listbox is created so
         ' if we wait until after this notification then we will have missed our chance
         ' to accurately display the codetip for the first item in the list).
         gApp.hWndAutoCListBox = 0
         EnumWindows(cast(EnumWindowsProc, @EnumWindowsProcCB), 0)
         if IsWindow(gApp.pWindowAutoC.hWindow) = 0 then
            ' Create the secondary popup window that shows additional information
            ' related to the selected autocomplete item
            gApp.pWindowAutoC.Create( HWND_FRMMAIN, "", @frmAutoCtooltip_WndProc, 0, 0, 0, 0, _
                                      WS_POPUP OR WS_BORDER or WS_CLIPCHILDREN OR WS_CLIPSIBLINGS)
            ' Set a monospaced font for better readibility
            gApp.pWindowAutoC.SetFont("Courier New", gApp.pWindowAutoC.DefaultFontSize)
            gApp.hWndAutoCTooltip = gApp.pWindowAutoC.AddControl("LABEL")
         END IF
        
'         if IsWindow(gApp.pWindowAutoC.hWindow) then
'            dim as string sID = *cast(zstring ptr, pNSC->lpText)
'            dim as string sText, sCodeTip
'            dim as RECT rc     
'            dim pData as DB2_DATA ptr
'            pData = GetAutoCTooltip(sID)    
'            if pData then
'               Dim pWindow As CWindow Ptr = AfxCWindowPtr(hWnd)
'               If pWindow = 0 Then Exit Function
'               GetWindowRect(gApp.hWndAutoCListBox, @rc )
                
'               dim as string sCodeTip = iif(len(pData->ElementValue), FormatCodetip(pData->ElementValue), " <no parameters>")
                    
'               select case pData->id
'                  CASE DB2_TYPE
'                     sText = iif(pData->IsEnum, "ENUM ", "TYPE ") & pData->ElementName
'                  case DB2_FUNCTION
'                     sText = "Function " & sCodeTip
'                  case DB2_SUB         
'                     sText = "Sub " & sCodeTip
'                  case DB2_PROPERTY  
'                     sText = "Property " & pData->ElementName & vbcrlf & sCodeTip
'                  case DB2_VARIABLE
'                     sText = pData->ElementName & vbcrlf & _
'                             pData->ElementValue
'                  case DB2_STANDARDDATATYPE
'                     sText = pData->ElementName
'               END SELECT
'               sText = sText & vbcrlf & pData->Description
'               sText = AfxStrReplace(sText, "..", ".")
                
'               ' Calculate the size of the tip window that will display all of
'               ' the tip text.
'               dim as HDC hDC = GetDC(gApp.hWndAutoCTooltip)
'               dim as RECT rcTip = (0, 0, 300, 50)
'               DrawText(hDC, sText, -1, @rcTip, DT_CALCRECT)
'               ReleaseDC(gApp.hWndAutoCTooltip, hDC)
'               ' Make the right side just a littler larger just in case
'               rcTip.right = rcTip.right + gApp.pWindowAutoC.ScaleX(20)
'               ' Make the static label control large enough
'               SetWindowPos( gApp.hWndAutoCTooltip, 0, _
'                             0, 0, _            
'                             gApp.pWindowAutoC.ScaleX(rcTip.right-rcTip.left+20), _
'                             gApp.pWindowAutoC.ScaleY(rcTip.bottom-rcTip.top+20), _
'                             SWP_NOZORDER or SWP_NOMOVE or SWP_NOACTIVATE)
'               ' Expand/contract the size of the parent tip window's client area              
'               gApp.pWindowAutoC.SetClientSize(rcTip.right-rcTip.left, rcTip.bottom-rcTip.top)
'               AfxSetWindowText(gApp.hWndAutoCTooltip, sText) 
               
'               SetWindowPos( gApp.pWindowAutoC.hWindow, HWND_TOPMOST, _
'                             rc.right + pWindow->ScaleX(2), rc.top, _
'                             0, 0, _
'                             SWP_SHOWWINDOW or SWP_NOSIZE or SWP_NOACTIVATE)
'               AfxRedrawWindow(gApp.pWindowAutoC.hWindow)
'               AfxRedrawWindow(gApp.hWndAutoCTooltip)
'            else     
'               ShowWindow(gApp.pWindowAutoC.hWindow, SW_HIDE)
'            end if      
'         end if              
        
      case SCN_AUTOCCANCELLED
         ' Destroy the secondary popup information window
         pDoc = gTTabCtl.GetActiveDocumentPtr()
         If pDoc Then 
            pDoc->AutoCompleteType = AUTOCOMPLETE_NONE
            KillAutoCTooltip()
         end if
         
      case SCN_AUTOCSELECTION
         ' A selection occured from the AutoComplete listbox. We do the insertion
         ' ourselves because the startpos of the word is not the same as the start position
         ' used by the autocomplete listbox (because we are doing incremental searches).
         pDoc = gTTabCtl.GetActiveDocumentPtr()
         If pDoc Then 
            hEdit = pDoc->hWndActiveScintilla
            ' Get the position where the listbox was opened.
            'dim as long nStartPos = SciExec( hEdit, SCI_AUTOCPOSSTART, 0, 0 )
            dim as long nStartPos = pDoc->AutoCStartPos
            Dim as long nPos = SciExec( hEdit, SCI_GETCURRENTPOS, 0, 0)
            ' Get the match word that existed when the listbox was displayed
            dim as long nLenMatchWord = len(pDoc->sMatchWord)
            ' Set the word that was selected in the listbox
            dim as String sText = *cast(zstring ptr, pNSC->lpText)
            ' Remove and (Get)/(Set) that would have been added for Property
            sText = AfxStrParse(sText, 1, " ")
            SciExec( hEdit, SCI_SETSEL, nStartPos-nLenMatchWord, nPos)
            SciExec( hEdit, SCI_REPLACESEL, 0, Cast(LPARAM, strptr(sText)))
            nPos = SciExec( hEdit, SCI_GETCURRENTPOS, 0, 0)
            SciExec( hEdit, SCI_SETSEL, nPos, nPos)
            ' Now that we have inserted our own text, cancel the autoinsertion by
            ' the autocomplete listbox.
            SciExec( hEdit, SCI_AUTOCCANCEL, 0, 0)
            pDoc->AutoCompleteType = AUTOCOMPLETE_NONE
            KillAutoCTooltip   ' Destroy the secondary popup information window
                            
            if chr(pNSC->ch) = "=" then
               ' Add a space before and after Equals sign whenever the user presses
               ' the equals sign to terminate ab autocomplete popup list.
               PostMessage(HWND, MSG_USER_APPENDEQUALSSIGN, 0, 0)
            end if
            return true   
         end if

      case SCN_AUTOCCHARDELETED    
         ' Use deleted a character while autocompletion list was active. Display
         ' new contents of list because the underlying word being typed as changed.
         ShowAutocompleteList(SCN_AUTOCCHARDELETED)
                                    
      CASE SCN_CHARADDED
         pDoc = gTTabCtl.GetActiveDocumentPtr()
         If pDoc = 0 Then exit function
         hEdit = pDoc->hWndActiveScintilla
       
         select CASE chr(pNSC->ch)
         case ".", ">"    '  Show autocomplete list for TYPE variables
            ' dot "." or ">"  part of a pointer
            ' Need to PostMessage in order to give time for notification to complete. the "."
            ' is also used as a character that selects and closes the popup list. If no
            ' PostMessage then the "." will automatically select the first entry in the list
            ' and close.
            pDoc->AutoCompleteType = AUTOCOMPLETE_NONE
            If PostMessage(HWND_FRMMAIN, MSG_USER_SHOWAUTOCOMPLETE, pNSC->ch, 0) then exit function
          
         case ")"      ' Close and active code tip 
            SciExec( hEdit, SCI_CALLTIPCANCEL, 0, 0 )
            KillAutoCTooltip()
               
         case "("      ' Show code tip 
            pDoc->AutoCompleteType = AUTOCOMPLETE_NONE
            If PostMessage(HWND_FRMMAIN, MSG_USER_SHOWAUTOCOMPLETE, pNSC->ch, 0) then exit function

         case chr(13)  ' ENTER KEY PRESSED
            AttemptAutoInsert()
    
         case chr(32)  ' Space key pressed (Insert AutoComplete constructs if applicable)
            pDoc->AutoCompleteType = AUTOCOMPLETE_NONE
            if ShowAutocompleteList() then exit function

         case else   ' all other letters
            if ShowAutocompleteList() then exit function
             
         END SELECT

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
   PostMessage( HWND_FRMMAIN, MSG_USER_SETFOCUS, 0, 0 )
   Function = 0
End Function


' ========================================================================================
' Process WM_PAINT message for window/dialog: frmMain
' ========================================================================================
Function frmMain_OnPaint( ByVal HWnd As HWnd) As LRESULT
            
   Dim As PAINTSTRUCT ps
   Dim As HPEN hPen
   Dim As HDC hDc
   dim as RECT rc
   
   hDC = BeginPaint(hWnd, @ps)

   SaveDC hDC

   dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr
   if pDoc THEN
      ' Get location of vertical scrollbar and draw splitter button right above it.
      if IsWindowVisible(pDoc->hScrollbar) then
         GetWindowRect(pDoc->hScrollbar, @rc)
         rc.top = rc.top - SPLITSIZE: rc.bottom = rc.top + SPLITSIZE
         MapWindowPoints(0, HWND_FRMMAIN, cast(POINT ptr, @rc), 2)
         pDoc->rcSplitButton = rc
         if IsWindowVisible(pDoc->hWindow(1)) THEN 
            ' If the window is visible then we need to extend the splitter hit
            ' area to include all the way to left edge of the client area.
            if IsWindowVisible(HWND_FRMEXPLORER) THEN 
               rc.left = AfxGetWindowWidth( HWND_FRMEXPLORER )
               pDoc->rcSplitButton.left = rc.left
            end if   
         End If
         DrawFrameControl( hDC, @rc, DFC_BUTTON, DFCS_BUTTONPUSH)
      end if
   end if
   
   RestoreDC hDC, -1 
   
   EndPaint hWnd, @ps
   
   Function = 0
End Function



' ========================================================================================
' Position all child windows. Called manually and/or by WM_SIZE
' ========================================================================================
Function frmMain_PositionWindows( ByVal HWnd As HWnd ) As LRESULT
   
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

   ShowWindow( GetDlgItem(HWnd, IDC_FRMMAIN_REBAR), iif(gConfig.HideToolbar, SW_HIDE, SW_SHOW) )
   ShowWindow( GetDlgItem(HWnd, IDC_FRMMAIN_STATUSBAR), iif(gConfig.HideStatusbar, SW_HIDE, SW_SHOW) )
   if gConfig.HideToolbar THEN nHeightRebar = 0
   if gConfig.HideStatusbar THEN nHeightStatusBar = 0
      
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
   If pDoc Then 
      ' Position the Scintilla child edit windows
      dim as RECT rcClient
      rcClient.left = nLeft
      rcClient.top = rc.Top + nHeightRebar + nHeightTabControl
      rcClient.right = rc.Right - nWidthExplorer
      rcClient.bottom = rc.bottom - nHeightStatusBar - nHeightOutput
      
      dim as long iScrollbarWidth  = GetSystemMetrics(SM_CXVSCROLL)
      dim as long iScrollbarHeight = GetSystemMetrics(SM_CXHSCROLL)
      dim as long iWidth = rcClient.right - iScrollbarWidth

      dim as long lStyle = SWP_NOZORDER Or SWP_NOOWNERZORDER Or _
                           SWP_SHOWWINDOW Or SWP_NOACTIVATE Or SWP_NOCOPYBITS
      
      dim as long nSplitSize 
      
      ' Position top (optional) edit window
      if pDoc->SplitY - rcClient.top <= 0 THEN 
         ShowWindow(pDoc->hWindow(1), SW_HIDE)      
         nSplitSize = 0              
      Else
         SetWindowPos( pDoc->hWindow(1), 0, _
                       rcClient.left, rcClient.top, _
                       rcClient.right, pDoc->SplitY - rcClient.top, _ 
                       lStyle)
         nSplitSize = SPLITSIZE              
      end if
      
      dim as long nTop = Max(rcClient.top, pDoc->SplitY)

      ' Position main edit window
      SetWindowPos( pDoc->hWindow(0), 0, _
                    rcClient.left, nTop + nSplitSize, _
                    iWidth, rcClient.bottom - nTop - nSplitSize, _
                    lStyle)

      ' Position the main view vertical scroll bar
      SetWindowPos( pDoc->hScrollbar, 0, _ 
                    rcClient.left + iWidth, _
                    nTop + SPLITSIZE, _
                    iScrollbarWidth, _
                    rcClient.bottom - SPLITSIZE - nTop - iScrollbarHeight, _
                    lStyle)

      ' Update vertical scroll bar page value
      pDoc->ScrInfo.nPage = SciExec(pDoc->hWindow(0), SCI_LINESONSCREEN, 0, 0)
      pDoc->ScrInfo.nMax = SciExec(pDoc->hWindow(0), SCI_GETLINECOUNT, 0, 0) - 1    
      pDoc->ScrInfo.nPos = SciExec(pDoc->hWindow(0), SCI_GETFIRSTVISIBLELINE, 0, 0)
      SetScrollInfo( pDoc->hScrollbar, SB_CTL, @pDoc->ScrInfo, TRUE)

      AfxRedrawWindow(HWND_FRMMAIN)
   end if
   
   ' Ensure that the correct notes are shown
   frmOutput_ShowNotes
   frmMain_UpdateLineCol(HWND)
   frmMain_SetFocusToCurrentCodeWindow

   Function = 0
End Function


' ========================================================================================
' Save project file to disk
' ========================================================================================
Function OnCommand_ProjectSave( ByVal HWnd As HWnd, ByVal bSaveAs As BOOLEAN = False) As LRESULT

   ' Do the actual saving to disk
   dim idx as long = gApp.GetActiveProjectIndex() 
   gApp.Projects(idx).SaveProject(bSaveAs)

   frmMain_SetFocusToCurrentCodeWindow

   Function = 0
End Function


' ========================================================================================
' Close currently active project
' ========================================================================================
Function OnCommand_ProjectClose( ByVal HWnd As HWnd ) As LRESULT

   Dim wText As WString * MAX_PATH  
   dim idx as long = gApp.GetActiveProjectIndex()
   dim as long r, iTab
   dim pDoc as clsDocument ptr

   ' Save the project configuration to disk
   OnCommand_ProjectSave( HWnd, False )

   ' Cycle all files in the Project and ask to save any dirty ones. If a 
   ' save is cancelled then do not proceed to finalize closing the project.
   for i as long = 0 to gApp.Projects(idx).GetDocumentCount - 1
      pDoc = gApp.Projects(idx).GetDocumentPtr(i)
      if pDoc = 0 THEN continue for
      if pDoc->ProjectIndex <> idx THEN continue for
         
      If cbool(SciExec( pDoc->hWindow(0), SCI_GETMODIFY, 0, 0 )) or pDoc->UserModified Then
         OpenSelectedDocument(pDoc, "")
         wText = AfxStrPathname( "NAMEX", pDoc->Diskfilename )
         r = MessageBox( HWnd, L(76,"Save current changes?") & WStr(" ") & wText, @WStr(APPNAME), _
                           MB_YESNOCANCEL Or MB_ICONQUESTION)
         If r = IDCANCEL Then Exit Function
         If r = IDYES Then 
            r = pDoc->SaveFile()
            If r = False Then Exit Function   ' save was cancelled
         End If
      End If
      
      ' Remove from top tab control
      iTab = gTTabCtl.GetTabIndexByDocumentPtr( pDoc)
      if iTab > -1 THEN
         TabCtrl_DeleteItem(gTTabCtl.hWindow, iTab)
         gTTabCtl.SetFocusTab(iTab-1)
      END IF
      for ii as long = lbound(pDoc->hWindow) to ubound(pDoc->hWindow)
         ShowWindow pDoc->hWindow(ii), SW_HIDE
      NEXT
      
   NEXT
   
   ' Reset the gdb database but reload the FB keywords (need valid pDocs)
   gdb2.dbDeleteByProject(idx)

   ' If we get this far then we can safely remove all of the pDocs for
   ' this project from the global array
   gApp.Projects(idx).RemoveAllDocuments

   DestroyWindow( HWND_FRMFNLIST )

   ' Remove the Explorer node for this project
   Dim As Hwnd hTree = GetDlgItem(HWND_FRMEXPLORER, IDC_FRMEXPLORER_TREE)
   TreeView_DeleteItem(hTree, gApp.Projects(idx).hExplorerRootNode)
   gApp.Projects(idx).hExplorerRootNode = 0
   gApp.Projects(idx).ProjectNotes = ""
   gApp.Projects(idx).InUse = false
   
   dim hCtl as hwnd = GetDlgItem(HWND_FRMOUTPUT, IDC_FRMOUTPUT_TXTNOTES)
   AfxSetWindowText(hCtl, gApp.NonProjectNotes)

   ' Make sure the next (if applicable) project is set as active so it can then be closed.
   gApp.EnsureDefaultActiveProject

   gApp.IsWindowIncludes = false
   gdb2.dbDeleteWinAPI

   PositionExplorerWindows(HWND_FRMEXPLORER)
   frmMain_PositionWindows(HWND_FRMMAIN)
   
   frmMain_UpdateLineCol(HWND_FRMMAIN)
   frmMain_SetFocusToCurrentCodeWindow

   Function = true
End Function


' ========================================================================================
' Create a new project
' ========================================================================================
Function OnCommand_ProjectNew( ByVal HWnd As HWnd ) As LRESULT

   ' Ensure that the Project Manager window has been created (show it now as well)
   frmProjectOptions_Show( HWnd, SW_SHOW, true )
   if gApp.IsNewProjectFlag = false then return 0  ' New project was Cancelled
      
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
                                 ByVal pDocIn      As clsDocument Ptr, _
                                 byval bIsDesigner as Boolean _
                                 ) As clsDocument Ptr
   
   Dim As Long iTab = -1
   dim as long idx
   dim as HTREEITEM hRootNode
   Dim pDoc As clsDocument Ptr 
   Dim As Hwnd hTree = GetDlgItem(HWND_FRMEXPLORER, IDC_FRMEXPLORER_TREE)
   dim as boolean bAttachToProject = true
   
   ' Check to see if the document is already loaded and part of the
   ' project it is being added to.
   idx = gApp.GetActiveProjectIndex()

   ' Need to determine if the document is already part of an open project
   ' because if it is then we don't want to add the file to the current
   ' project thus making a dual link.
   for i as long = 1 to ubound(gApp.Projects)
      if gApp.Projects(i).InUse = true THEN
         pDoc = gApp.Projects(i).GetDocumentPtrByFilename(pwszName)
         if pDoc then
            bAttachToProject = false
            exit for
         end if   
      END IF
   NEXT
   
   if pDoc then pDocIn = pDoc
   
   if bAttachToProject then
      If pDocIn = 0 Then   ' Create a new pDoc 
         gFileLoading = true  ' allows parsing of newly loading file
         pDoc = New clsDocument
         if bIsDesigner THEN
            If bIsNewFile Then
               pDoc->CreateDesignerWindow(HWnd, True)  ' Create the new visual designer window
            Else
               pDoc->CreateDesignerWindow(HWnd, false, pwszName)
            End If   
         else
            If bIsNewFile Then
               pDoc->CreateCodeWindow(HWnd, True)  ' Create the new Scintilla window
            Else
               pDoc->CreateCodeWindow(HWnd, false, bIsTemplate, pwszName)
            End If   
         END IF
         gApp.Projects(idx).AddDocument(pDoc) 
         ParseDocument(idx, pDoc, "")
         gFileLoading = false
      Else
         pDoc = pDocIn     ' deal with the existing pDoc
      End If   
      pwszName = @pDoc->DiskFilename

      if idx = 0 THEN   ' Project0 "FILES"
         If gApp.Projects(0).hExplorerRootNode = 0 Then
            gApp.Projects(0).hExplorerRootNode = FF_TreeView_InsertItem(hTree, TVI_ROOT, "FILES" )
            Treeview_RemoveCheckBox(hTree, gApp.Projects(0).hExplorerRootNode )
            Treeview_SetBold(hTree, gApp.Projects(0).hExplorerRootNode, TRUE)
         End If
         hRootNode = gApp.Projects(0).hExplorerRootNode 
         ' no project is active so add the file to the FILES Explorer list.
         AddFunctionsToExplorerTreeview( pDoc, False )
      else
         hRootNode = gApp.Projects(idx).hExplorerRootNode
         ' If a project is active then do not add any Include files to it and don't
         ' show it in the Explorer's project list.
         if bIsInclude = false then
            AddFunctionsToExplorerTreeview( pDoc, False )
            ' Give this document a default project type depending on its file extension
            gApp.Projects(idx).ProjectSetFileType(pDoc, pDoc->ProjectFileType)
         end if
      END IF
                       
      ' Relate this pDoc to the correct project even if it is Project0
      pDoc->ProjectIndex = idx
      
      ' Set the default build configuration for this document. If no IsDefault option
      ' has been checked then we simply use whatever the current selected build is.
      if pDoc->DocumentBuild = "" THEN 
         pDoc->DocumentBuild = GetDefaultBuildGUID()
         if pDoc->DocumentBuild = "" THEN 
            pDoc->DocumentBuild = GetSelectedBuildGUID()
         end if
      end if      
   end if
   
   Treeview_SortChildren( hTree, hRootNode, 0 )
   TreeView_Expand( hTree, hRootNode, TVE_EXPAND)

   If bShowInTab Then
      ' If the document is already open and loaded then simply switch to
      ' that document in the top tabcontrol.
      iTab = gTTabCtl.GetTabIndexFromFilename(pwszName)
      If iTab = -1 Then iTab = gTTabCtl.AddTab( pDoc )  ' Add the new document to the top tabcontrol
      gTTabCtl.SetFocusTab(iTab)
      Treeview_SelectItem( GetDlgItem(HWND_FRMEXPLORER, IDC_FRMEXPLORER_TREE), _
                           pDoc->hNodeExplorer )
   End If

   ' Position all of the controls into place
   frmMain_PositionWindows(HWnd)
   frmMain_SetFocusToCurrentCodeWindow

   Function = pDoc
End Function


' ========================================================================================
' Create a new visual designer window
' ========================================================================================
Function OnCommand_DesignerNewForm( ByVal HWnd As HWnd ) As LRESULT
   frmMain_OpenFileSafely(HWnd, _
                           True,  _   ' bIsNewFile
                           False, _   ' bIsTemplate
                           True,  _   ' bShowInTab
                           false, _   ' bIsInclude
                           0, _       ' pwszName
                           0, _       ' pDocIn
                           true )     ' Create a visual designer   
   Function = 0
End Function


' ========================================================================================
' Align controls on the form
' ========================================================================================
function OnCommand_DesignerAlign( byval HWND as HWND, byval id as long) as LRESULT

   dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr
   if pDoc = 0 THEN exit function

   dim pCtrlActive as clsControl ptr = pDoc->Controls.GetActiveControl()
   dim pCtrl as clsControl ptr
   dim as RECT rcCtrl, rcActive

   dim as long nWidth, nHeight
   dim as long nWidthActive, nHeightActive
   
   if pCtrlActive = 0 THEN exit function
   rcActive = GetControlRECT(pCtrlActive)
   nWidthActive  = rcActive.Right  - rcActive.Left
   nHeightActive = rcActive.Bottom - rcActive.Top 
   
   for i as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
      pCtrl = pDoc->Controls.ItemAt(i)
      if pCtrl->IsSelected THEN
         ' Skip processing the active control because that is unnecessary.
         if pCtrl = pCtrlActive THEN continue for
         rcCtrl = GetControlRECT(pCtrl)
         nWidth  = rcCtrl.Right  - rcCtrl.Left
         nHeight = rcCtrl.Bottom - rcCtrl.Top 
         select CASE id
            case IDM_ALIGNLEFTS
               SetControlProperty(pCtrl, "LEFT", str(rcActive.left))
               ApplyControlProperties(pCtrl)
            case IDM_ALIGNTOPS
               SetControlProperty(pCtrl, "TOP", str(rcActive.top))
               ApplyControlProperties(pCtrl)
            case IDM_ALIGNRIGHTS
               SetControlProperty(pCtrl, "LEFT", str(rcctrl.left + (rcActive.right-rcCtrl.right)) )
               ApplyControlProperties(pCtrl)
            case IDM_ALIGNBOTTOMS 
               SetControlProperty(pCtrl, "TOP", str(rcCtrl.top + (rcActive.bottom-rcCtrl.bottom)) )
               ApplyControlProperties(pCtrl)
            case IDM_ALIGNCENTERS  
               SetControlProperty(pCtrl, "LEFT", _
                  str(rcActive.Left + (nWidthActive / 2) - ( nWidth / 2 )) )
               ApplyControlProperties(pCtrl)
            case IDM_ALIGNMIDDLES  
               SetControlProperty(pCtrl, "TOP", _
                  str(rcActive.Top + (nHeightActive / 2) - ( nHeight / 2 )) )
               ApplyControlProperties(pCtrl)
            case IDM_SAMEWIDTHS  
               SetControlProperty(pCtrl, "WIDTH", str(nWidthActive))
               ApplyControlProperties(pCtrl)
            case IDM_SAMEHEIGHTS
               SetControlProperty(pCtrl, "HEIGHT", str(nHeightActive))
               ApplyControlProperties(pCtrl)
            case IDM_SAMEBOTH
               pCtrl->SuspendLayout = true
               SetControlProperty(pCtrl, "WIDTH", str(nWidthActive))
               SetControlProperty(pCtrl, "HEIGHT", str(nHeightActive))
               ApplyControlProperties(pCtrl)
               pCtrl->SuspendLayout = false
         END SELECT
      END IF
   next

   ' Ensure the grab handles are redrawn
   AfxRedrawWindow(pDoc->hWndForm)

   function = 0
end function


' ========================================================================================
' Center controls on the form
' ========================================================================================
function OnCommand_DesignerCenter( byval HWND as HWND, byval id as long) as LRESULT

   dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr
   if pDoc = 0 THEN exit function

   dim pCtrl as clsControl ptr

   dim as RECT rcCtrl, rcForm
   dim as long nLeft, nTop, nMaxRight, nMaxBottom
   dim as long nHorizSpacing, nVertSpacing, nFormWidth, nFormHeight
   dim as long nMinLeft = 999999999
   dim as long nMinTop  = 999999999
   
   ' Calculate the total distance from the far left (or top) control to the far
   ' right (or bottom) control and subtract the widths of each control. This will
   ' give us the amount of 'white' space. Divide that white space by
   ' half to determine the amount to put on the left/top and right/bottom of the
   ' group of controls.
   for i as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
      pCtrl = pDoc->Controls.ItemAt(i)
      if pCtrl->ControlType = CTRL_FORM THEN continue for
      if pCtrl->IsSelected = false THEN continue for
      rcCtrl     = GetControlRECT(pCtrl)
      nMinLeft   = Min( nMinLeft, rcCtrl.left )
      nMaxRight  = Max( nMaxRight, rcCtrl.right )
      nMinTop    = Min( nMinTop, rcCtrl.top )
      nMaxBottom = Max( nMaxBottom, rcCtrl.bottom )
   next

   ' Get the client area of the design form
   GetClientRect(pDoc->hWndForm, @rcForm)
   nFormWidth  = (rcForm.right - rcForm.left)
   nFormHeight = (rcForm.bottom - rcForm.top)
   
   ' Determine the value by which the group of controls needs to be shifted by.  
   nHorizSpacing = nMinLeft - ( (nFormWidth - (nMaxRight - nMinLeft)) / 2) 
   nVertSpacing  = nMinTop - ( (nFormHeight - (nMaxBottom - nMinTop) ) / 2) 

   for i as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
      pCtrl = pDoc->Controls.ItemAt(i)
      if pCtrl->ControlType = CTRL_FORM THEN continue for
      if pCtrl->IsSelected = false THEN continue for
      select CASE id
         case IDM_CENTERHORIZ
            nLeft = val(GetControlProperty(pCtrl, "LEFT"))
            SetControlProperty(pCtrl, "LEFT", str(nLeft-nHorizSpacing))
            ApplyControlProperties(pCtrl)
         case IDM_CENTERVERT
            nTop  = val(GetControlProperty(pCtrl, "TOP"))
            SetControlProperty(pCtrl, "TOP", str(nTop-nVertSpacing))
            ApplyControlProperties(pCtrl)
         case IDM_CENTERBOTH   
            nLeft = val(GetControlProperty(pCtrl, "LEFT"))
            nTop  = val(GetControlProperty(pCtrl, "TOP"))
            pCtrl->SuspendLayout = true
            SetControlProperty(pCtrl, "LEFT", str(nLeft-nHorizSpacing))
            SetControlProperty(pCtrl, "TOP", str(nTop-nVertSpacing))
            ApplyControlProperties(pCtrl)
            pCtrl->SuspendLayout = false
      END SELECT
   next

   ' Ensure the grab handles are redrawn
   AfxRedrawWindow(pDoc->hWndForm)

   function = 0
end function

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
   Dim dwItemCount As Long, i As Long, pItem As IShellItem Ptr, pwszName As WString Ptr
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
      ' Apply document properties to this file because the file extension may have
      ' changed. For example from Untitled to *.bas
      dim as long idx = pDoc->ProjectIndex
      If gApp.IsProjectActive Then
         gApp.Projects(idx).ProjectSetFileType( pDoc, pDoc->ProjectFileType )    
      End If
      pDoc->ApplyProperties
      pDoc->DisplayStats()   
      ' Parse the code for sub/function listing and create list
      ParseDocument(idx, pDoc, "")
   EndIf   

   ' Ensure that the Tab displays the correct filename in case it was changed
   gTTabCtl.SetTabText(-1)
   frmMain_SetFocusToCurrentCodeWindow
   AddFunctionsToExplorerTreeview(pDoc, true)
   frmFnList_UpdateListBox()

   Function = 0
End Function

' ========================================================================================
' Save Sub/Function Declares to disk
' ========================================================================================
Function OnCommand_FileSaveDeclares( ByVal HWnd As HWnd ) As LRESULT

   dim wFilename as wstring * MAX_PATH
   static wText as wstring * MAX_PATH
   
   Dim pwszName As WString Ptr = AfxIFileSaveDialog(HWND_FRMMAIN, @wText, "bi", IDM_FILESAVEDECLARES)
   If pwszName Then
      wFilename = *pwszName
      wText = AfxStrPathname( "NAMEX", wFilename )
      CoTaskMemFree(pwszName)
   Else
      Return False
   End If
 
   CreateDeclaresFile(wFilename)
   
   ' If the declares file is loaded in the project then we need to reload
   ' it in order to ensure most current document data.
   dim pDoc as clsDocument ptr
   dim as long idx = gApp.GetActiveProjectIndex()
   pDoc = gApp.Projects(idx).GetDocumentPtrByFilename(wFilename)
   if pDoc THEN 
      dim sText as string 
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
         SciExec(pDoc->hWindow(0), SCI_SETSAVEPOINT, 0, 0)
      End If
   end if
  
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
      If cbool(SciExec( pDoc->hWindow(0), SCI_GETMODIFY, 0, 0 )) or pDoc->UserModified Then
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
   Dim As Long r, iTab
   
   Dim As Hwnd hTree = GetDlgItem(HWND_FRMEXPLORER, IDC_FRMEXPLORER_TREE)

   Do
      pDoc = gTTabCtl.GetActiveDocumentPtr()
      If pDoc = 0 Then Return true
   
      If cbool(SciExec( pDoc->hWindow(0), SCI_GETMODIFY, 0, 0 )) or pDoc->UserModified Then
         wText = AfxStrPathname( "NAMEX", pDoc->Diskfilename )
         r = MessageBox( HWnd, L(76,"Save current changes?") & WStr(" ") & wText, @WStr(APPNAME), _
                           MB_YESNOCANCEL Or MB_ICONQUESTION)
         If r = IDCANCEL Then Exit Function
         If r = IDYES Then 
            r = pDoc->SaveFile()
            If r = False Then Exit Function   ' save was cancelled
         End If
      End If
      
      ' Remove from top tab control
      iTab = TabCtrl_GetCurSel(gTTabCtl.hWindow)
      TabCtrl_DeleteItem(gTTabCtl.hWindow, iTab)
      ' Set the active tab to the closest tab to the one just removed
      iTab = TabCtrl_GetCurSel(gTTabCtl.hWindow)
      gTTabCtl.SetFocusTab(iTab)
      
      for i as long = lbound(pDoc->hWindow) to ubound(pDoc->hWindow)
         ShowWindow( pDoc->hWindow(i), SW_HIDE)
      NEXT
      ShowWindow( HWND_FRMVDTOOLBOX, SW_HIDE)

      ' Only remove this document from the global collection if it is not
      ' part of any active Project or if it is new/unsaved file. Files that are part of a project are
      ' all closed (tabs/nodes) automatically in one shot rather than individually.
      if (gApp.IsProjectActive = false) or (pDoc->IsNewFlag = true) then 
         ' Remove it from the Explorer
         Treeview_DeleteItem( hTree, pDoc->hNodeExplorer)
      end if   
      if pDoc->ProjectIndex = 0 then
         gApp.Projects(pDoc->ProjectIndex).RemoveDocument(pDoc)
      end if   
      
      frmMain_PositionWindows(HWnd)
   
      If bCloseAll = False Then Exit Do
      
      If bCloseAll Then
         If TabCtrl_GetItemCount(gTTabCtl.hWindow) = 0 Then Exit Do
      End If
   Loop
    
   PositionExplorerWindows(HWND_FRMEXPLORER)
   
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
      SendMessage GetDlgItem(HWnd, IDC_FRMMAIN_REBAR), WM_SIZE, cx, cy
      
      ' Resize the status bar
      Dim As HWnd hStatusBar = GetDlgItem(HWnd, IDC_FRMMAIN_STATUSBAR)
      SendMessage hStatusBar, WM_SIZE, cx, cy
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
      select case codeNotify
         case SCEN_SETFOCUS:  pDoc->hWndActiveScintilla = hwndCtl
         case SCEN_KILLFOCUS: pDoc->hWndActiveScintilla = 0
      end select      
      hEdit = pDoc->hWndActiveScintilla
      frmMain_ChangeToolbarButtonsState
   End If   

   
   Select Case id
             
      case IDC_FRMMAIN_COMBOBUILDS
         if codeNotify = CBN_SELCHANGE THEN
            dim as long nCurSel = ComboBox_GetCurSel(hwndCtl)
            dim as long idx = gApp.GetActiveProjectIndex()
            if nCurSel = -1 THEN exit function
            If gApp.IsProjectActive Then
               gApp.Projects(idx).ProjectBuild = gConfig.Builds(nCurSel).Id
            Else
               Dim pDocMain As clsDocument Ptr 
               pDocMain = gTTabCtl.GetActiveDocumentPtr()
               If pDocMain Then pDocMain->DocumentBuild = gConfig.Builds(nCurSel).Id
            end if
            frmMain_UpdateLineCol(HWND)
         END IF  
      
      Case IDM_NEXTTAB:            Function = gTTabCtl.NextTab()
      Case IDM_PREVTAB:            Function = gTTabCtl.PrevTab()
      Case IDM_CLOSETAB:           Function = gTTabCtl.CloseTab()

      ''  FILE MENU
      Case IDM_FILENEW:            Function = OnCommand_FileNew( HWnd )
      Case IDM_FILEOPEN:           Function = OnCommand_FileOpen( HWnd )
      Case IDM_FILESAVE:           Function = OnCommand_FileSave( HWnd )
      Case IDM_FILESAVEAS:         Function = OnCommand_FileSave( HWnd, True )
      Case IDM_FILESAVEDECLARES:   Function = OnCommand_FileSaveDeclares( HWnd )
      Case IDM_FILESAVEALL:        Function = OnCommand_FileSaveAll( HWnd ) 
      Case IDM_FILECLOSE:          Function = OnCommand_FileClose( HWnd, False )
      Case IDM_FILECLOSEALL:       Function = OnCommand_FileClose( HWnd, True )
      
      Case IDC_MRUBASE To (IDC_MRUBASE + 10)
         Function = OpenMRUFile( HWnd, id )

      Case IDM_MRUCLEAR, IDM_MRUPROJECTCLEAR
         Function = ClearMRUlist( id )
         
      Case IDM_OPENINCLUDE:        Function = OnCommand_OpenIncludeFile( HWnd )
      Case IDM_COMMAND:            Shell(Environ("COMSPEC"))
      Case IDM_EXIT:               SendMessage( HWnd, WM_CLOSE, 0, 0 )

      ''  EDIT MENU
      Case IDM_REDO:               Function = SciExec( hEdit, SCI_REDO, 0, 0)
      Case IDM_UNDO, WM_UNDO:      Function = SciExec( GetFocus(), Iif(GetFocus()=hEdit, SCI_UNDO, WM_UNDO), 0, 0)
      
      Case IDM_CUT, WM_CUT
         if pDoc->IsDesigner then
            SendMessage(HWND_FRMMAIN, WM_COMMAND, MAKELONG(IDM_COPY, 0), 0)
            SendMessage(HWND_FRMMAIN, WM_COMMAND, MAKELONG(IDM_DELETE, 0), 0)
         else
            Function = SciExec( GetFocus(), Iif(GetFocus()=hEdit, SCI_CUT, WM_CUT), 0, 0)
         end if   
      
      Case IDM_COPY, WM_COPY
         if pDoc->IsDesigner then
            dim as long NumControls = pDoc->Controls.SelectedControlsCount
            if NumControls = 0 THEN exit function
            redim gCopyControls(NumControls-1) as clsControl

            dim as long NextControl = 0
            dim pCtrl as clsControl ptr
            for i as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
               pCtrl = pDoc->Controls.ItemAt(i)
               if pCtrl->IsSelected THEN 
                  gCopyControls(NextControl) = *pCtrl
                  NextControl = NextControl + 1
               END IF
            NEXT
         else
            Function = SciExec( GetFocus(), Iif(GetFocus()=hEdit, SCI_COPY, WM_COPY), 0, 0)
         end if

            
      Case IDM_PASTE, WM_PASTE
         if pDoc->IsDesigner then
            dim pCtrl as clsControl ptr
            dim rc as RECT
            dim as long nLeft, nTop, nLeftActive, nTopActive
            
            ' Need to get the Left/Top properties for the active control in the copy group
            for i as long = lbound(gCopyControls) to ubound(gCopyControls)
               if gCopyControls(i).IsActive THEN
                  nLeftActive = val(GetControlProperty(@gCopyControls(i), "LEFT"))
                  nTopActive  = val(GetControlProperty(@gCopyControls(i), "TOP"))
                  exit for
               END IF
            next
               
            ' Create the control and copy the properties to it
            for i as long = lbound(gCopyControls) to ubound(gCopyControls)
               pCtrl = CreateToolboxControl(pDoc, gCopyControls(i).ControlType, rc)
               
               for ii as long = lbound(gCopyControls(i).Properties) to ubound(gCopyControls(i).Properties)
                  select case ucase(gCopyControls(i).Properties(ii).wszPropName)
                     case "NAME"  ' don't copy b/c we will then have duplicate control names
                        continue for
                     case "LEFT": nLeft = val(gCopyControls(i).Properties(ii).wszPropValue)
                     case "TOP":  nTop  = val(gCopyControls(i).Properties(ii).wszPropValue)   
                  END select
                  pCtrl->Properties(ii) = gCopyControls(i).Properties(ii)
               NEXT
               
               ' Set the left/top properties relative to the current insert position
               dim pt as point
               GetCursorPos(@pt)
               MapWindowPoints(0, pDoc->hWndForm, cast(point ptr, @pt), 1)
               SetControlProperty(pCtrl, "LEFT", str(pt.x + (nLeft-nLeftActive)))
               SetControlProperty(pCtrl, "TOP", str(pt.y + (nTop-nTopActive)))
               pCtrl->IsActive = false
               pCtrl->IsSelected = false
               
               pCtrl->SuspendLayout = true
               ApplyControlProperties( pCtrl )
               pCtrl->SuspendLayout = false
            NEXT
            ' Ensure the grab handles of form and controls are redrawn or hidden
            pDoc->Controls.SetActiveControl(pDoc->hWndForm)
            AfxRedrawWindow(pDoc->hWndFrame)
            AfxRedrawWindow(pDoc->hWndForm)
            frmMain_UpdateLineCol(HWND_FRMMAIN)
            DisplayPropertyList(pDoc)
         else
            Function = SciExec( GetFocus(), Iif(GetFocus()=hEdit, SCI_PASTE, WM_PASTE), 0, 0)
         end if   
      
      case IDM_DELETE
         ' Delete selected controls
         dim pCtrl as clsControl ptr
         do until pDoc->Controls.SelectedControlsCount = 0
            for i as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
               pCtrl = pDoc->Controls.ItemAt(i)
               if pCtrl->IsSelected THEN 
                  pDoc->Controls.Remove(pCtrl)
                  exit for
               END IF
            NEXT
         loop
         ' Ensure the grab handles of form and controls are redrawn or hidden
         pDoc->Controls.SetActiveControl(pDoc->hWndForm)
         AfxRedrawWindow(pDoc->hWndFrame)
         AfxRedrawWindow(pDoc->hWndForm)
         frmMain_UpdateLineCol(HWND_FRMMAIN)
         DisplayPropertyList(pDoc)
        
      Case IDM_DELETELINE:         Function = SciExec(hEdit, SCI_LINECUT, 0, 0)
      Case IDM_INSERTFILE:         Function = pDoc->InsertFile()
      Case IDM_INDENTBLOCK
         ' If a TAB was pressed then manually terminate any active autocomplete.
         if SciExec(hEdit, SCI_AUTOCACTIVE, 0, 0) then
            SciExec(hEdit, SCI_AUTOCCOMPLETE, 0, 0) 
         else
            ' Do a block indent
            gApp.SuppressNotify = true: Function = SciExec(hEdit, SCI_TAB, 0, 0)
            gApp.SuppressNotify = false
         end if
      Case IDM_UNINDENTBLOCK
         gApp.SuppressNotify = true: Function = SciExec(hEdit, SCI_BACKTAB, 0, 0)
         gApp.SuppressNotify = false
      Case IDM_COMMENTBLOCK:       Function = pDoc->BlockComment(True) 
      Case IDM_UNCOMMENTBLOCK:     Function = pDoc->BlockComment(False) 
      Case IDM_DUPLICATELINE:      Function = pDoc->LineDuplicate()
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
      Case IDM_SELECTALL
            if pDoc->IsDesigner THEN
               pDoc->Controls.SelectAllControls
               ' Ensure the grab handles of form and controls are redrawn or hidden
               AfxRedrawWindow(pDoc->hWndFrame)
               AfxRedrawWindow(pDoc->hWndForm)
            else
               Function = SciExec(hEdit, SCI_SELECTALL, 0, 0)
            END IF
      Case IDM_SELECTLINE:         Function = pDoc->SelectLine(-1)
                                   
      case IDM_ANSI
         if pDoc->FileEncoding <> FILE_ENCODING_ANSI THEN 
            pDoc->UserModified = true
            ' Convert buffer to ANSI and redisplay the text in the editor
            ConvertTextBuffer(pDoc, FILE_ENCODING_ANSI)
         end if   
         pDoc->FileEncoding = FILE_ENCODING_ANSI 
         frmMain_UpdateLineCol( HWND_FRMMAIN )
      
      case IDM_UTF8BOM
         if pDoc->FileEncoding <> FILE_ENCODING_UTF8_BOM THEN 
            pDoc->UserModified = true
            ConvertTextBuffer(pDoc, FILE_ENCODING_UTF8_BOM)
         end if   
         pDoc->FileEncoding = FILE_ENCODING_UTF8_BOM 
         frmMain_UpdateLineCol( HWND_FRMMAIN )
      
      case IDM_UTF16BOM
         if pDoc->FileEncoding <> FILE_ENCODING_UTF16_BOM THEN 
            pDoc->UserModified = true
            ConvertTextBuffer(pDoc, FILE_ENCODING_UTF16_BOM)
         end if   
         frmMain_UpdateLineCol( HWND_FRMMAIN )


      ''  SEARCH MENU              
      Case IDM_FIND:               Function = frmFindReplace_Show(hEdit, false)
      Case IDM_REPLACE:            Function = frmFindReplace_Show(hEdit, true)
      Case IDM_FINDINFILES:        Function = frmFindInFiles_Show(hEdit)

      Case IDM_FINDNEXT, IDM_FINDPREV
         if pDoc = 0 then exit function

         dim as long startPos, endPos, curPos, mainSel
         Dim As String sFindText 
         mainSel  = SciExec(hEdit, SCI_GETMAINSELECTION, 0, 0)
         startPos = SciExec(hEdit, SCI_GETSELECTIONSTART, mainSel, 0)
         endPos   = SciExec(hEdit, SCI_GETSELECTIONEND, mainSel, 0)
         sFindText = pDoc->GetTextRange(startPos, endPos)

         if (GetForeGroundWindow = HWND_FRMFINDREPLACE) or (Len(sFindText) = 0) then
            id = iif(id = IDM_FINDNEXT, IDC_FRMFINDREPLACE_BTNRIGHT, IDC_FRMFINDREPLACE_BTNLEFT)
            SendMessage( HWND_FRMFINDREPLACE, WM_COMMAND, MAKELONG(id, 0), Cast(LPARAM, GetDlgItem(HWND_FRMFINDREPLACE, id)) )
            exit function
         end if

         curPos = SciExec( hEdit, SCI_GETCURRENTPOS, 0, 0) 
         startPos = iif( id = IDM_FINDNEXT, curPos + 1, curPos - 1 )
         endPos = iif( id = IDM_FINDNEXT, SciExec( hEdit, SCI_GETTEXTLENGTH, 0, 0), 0)
         SciExec( hEdit, SCI_SETTARGETSTART, startPos, 0)
         SciExec( hEdit, SCI_SETTARGETEND, endPos, 0)
         ' Search for the text to find
         dim as long r = SciExec( hEdit, SCI_SEARCHINTARGET, Len(sFindText), Strptr(sFindText))
         If r <> -1 Then SciExec( hEdit, SCI_SETSEL, r, r + len(sFindText)) 


      Case IDM_DEFINITION:         Function = frmMain_GotoDefinition( pDoc )
      Case IDM_LASTPOSITION:       Function = frmMain_GotoLastPosition()
      
      Case IDM_GOTO:               Function = frmGoto_Show(hEdit)
      Case IDM_FUNCTIONLIST:       Function = frmFnList_Show( HWnd )    ' this will also toggle visibility 
      Case IDM_BOOKMARKTOGGLE:     Function = pDoc->ToggleBookmark(pDoc->GetCurrentLineNumber() )
      Case IDM_BOOKMARKNEXT:       Function = pDoc->NextBookmark()
      Case IDM_BOOKMARKPREV:       Function = pDoc->PrevBookmark()
      Case IDM_BOOKMARKCLEARALL:   Function = SciExec(hEdit, SCI_MARKERDELETEALL, -1, 0)
      
      ''  VIEW MENU
      Case IDM_FOLDTOGGLE:         Function = pDoc->FoldToggle( pDoc->GetCurrentLineNumber() )
      Case IDM_FOLDBELOW:          Function = pDoc->FoldToggleOnwards( pDoc->GetCurrentLineNumber() )
      Case IDM_FOLDALL:            Function = pDoc->FoldAll()
      Case IDM_UNFOLDALL:          Function = pDoc->UnFoldAll()
      Case IDM_ZOOMIN:             Function = SciExec(hEdit, SCI_ZOOMIN, 0, 0)
      Case IDM_ZOOMOUT:            Function = SciExec(hEdit, SCI_ZOOMOUT, 0, 0)

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
               
      
      ''  PROJECT MENU
      Case IDM_PROJECTNEW:            Function = OnCommand_ProjectNew( HWnd )
      Case IDM_PROJECTSAVE:           Function = OnCommand_ProjectSave( HWnd, False )
      Case IDM_PROJECTSAVEAS:         Function = OnCommand_ProjectSave( HWnd, True )
      Case IDM_PROJECTOPEN:           Function = OnCommand_ProjectOpen( HWnd )
      Case IDM_PROJECTFILESADD:       Function = OnCommand_FileOpen( HWnd )
      
      Case IDM_PROJECTFILESADDTONODE
         ' Normally files are added to whatever project currently has a displayed
         ' file open. If a Project root node was right clicked on and displayed the
         ' menu then we want to add to the project that the node refers to rather than
         ' the currently opened document's project. We do this by simulating a project
         ' load and manually specifiying the project index to use.
         dim as hwnd hTree = GetDlgItem(HWND_FRMEXPLORER, IDC_FRMEXPLORER_TREE)
         dim as HTREEITEM hNode = TreeView_GetSelection(hTree)
         for i as long = 1 to ubound(gApp.Projects)
            if gApp.Projects(i).hExplorerRootNode = hNode THEN
               gApp.ProjectOverrideIndex = i
               Function = OnCommand_FileOpen( HWnd )
               gApp.ProjectOverrideIndex = 0
               exit for
            END IF
         NEXT

      Case IDC_MRUPROJECTBASE To (IDC_MRUPROJECTBASE + 10)
         Function = OpenMRUProjectFile( HWnd, id )

      Case IDM_PROJECTCLOSE:   Function = OnCommand_ProjectClose( HWnd )
      Case IDM_PROJECTOPTIONS: Function = frmProjectOptions_Show( HWnd ) 
    
      Case IDM_ADDFILETOPROJECT to IDM_ADDFILETOPROJECT + 99
         ' Invoked from popup menu. If more than one project was open then the popup
         ' menu would have displayed a list of projects to choose from. We need to 
         ' normalize the project index here.
         gApp.ProjectOverrideIndex = (id - IDM_ADDFILETOPROJECT)
         ' Add the currently displayed document to the project.
         frmMain_OpenFileSafely(HWnd, _
                                 False, _    ' bIsNewFile
                                 False, _    ' bIsTemplate
                                 True,  _    ' bShowInTab
                                 false, _    ' bIsInclude
                                 0, _        ' pwszName
                                 pDoc )      ' pDocIn
         gApp.ProjectOverrideIndex = 0

      Case IDM_REMOVEFILEFROMPROJECT
         Dim As Hwnd hTree   = GetDlgItem(HWND_FRMEXPLORER, IDC_FRMEXPLORER_TREE)
         dim as long idx = pDoc->ProjectIndex
         dim as HTREEITEM hNode = pDoc->hNodeExplorer
         if OnCommand_FileClose( HWnd, False ) then
            ' File close succeeded. Remove this document from the project collection and remove treeview node 
            Treeview_DeleteItem( hTree, hNode)
            gApp.Projects(idx).RemoveDocument(pDoc)
         end if   
      
      Case IDM_SETFILENORMAL:         Function = gApp.Projects(pDoc->ProjectIndex).ProjectSetFileType( pDoc, FILETYPE_NORMAL )
      Case IDM_SETFILEMODULE:         Function = gApp.Projects(pDoc->ProjectIndex).ProjectSetFileType( pDoc, FILETYPE_MODULE )
      Case IDM_SETFILEMAIN:           Function = gApp.Projects(pDoc->ProjectIndex).ProjectSetFileType( pDoc, FILETYPE_MAIN )
      Case IDM_SETFILERESOURCE:       Function = gApp.Projects(pDoc->ProjectIndex).ProjectSetFileType( pDoc, FILETYPE_RESOURCE )
         
           
      ''  COMPILE MENU          
      Case IDM_BUILDEXECUTE, IDM_COMPILE, IDM_REBUILDALL, IDM_QUICKRUN
         Function = code_Compile(id)
      Case IDM_RUNEXE
         RunExe( GetRunExecutableFilename, gApp.Projects(pDoc->ProjectIndex).ProjectCommandLine )
      Case IDM_COMMANDLINE:           Function = frmCommandLine_Show( HWnd )


      ''  OPTIONS MENU
      Case IDM_OPTIONS:               Function = frmOptions_Show( HWnd, SW_NORMAL )
      case IDM_COMPILECONFIG:         function = frmCompileConfig_Show( HWND ) 
      case IDM_USERTOOLSDIALOG:       function = frmUserTools_Show( HWND ) 
         
      
      ''   DESIGNER MENU
      case IDM_NEWFORM:               function = OnCommand_DesignerNewForm(HWND)
      case IDM_ALIGNLEFTS, IDM_ALIGNCENTERS, IDM_ALIGNRIGHTS, _
           IDM_ALIGNTOPS, IDM_ALIGNMIDDLES, IDM_ALIGNBOTTOMS, _
           IDM_SAMEWIDTHS, IDM_SAMEHEIGHTS, IDM_SAMEBOTH            
         function = OnCommand_DesignerAlign(HWND, id)
      case IDM_CENTERHORIZ, IDM_CENTERVERT, IDM_CENTERBOTH
         function = OnCommand_DesignerCenter(HWND, id)
         
      case IDM_LOCKCONTROLS

      
      ''   TOOLS MENU
      Case IDM_USERTOOL To (IDM_USERTOOL + UBound(gConfig.Tools))
         ExecuteUserTool( id - IDM_USERTOOL )

      Case IDM_HELP
 ' TODO: Remove debug call in IDM_HELP
' gdb2.dbDebug
         ShowContextHelp() 
     

      Case IDM_ABOUT
         MessageBox( HWnd, APPNAME & vbcrlf & "Version " & APPVERSION & _
                       vbcrlf & vbcrlf & _
                       "Paul Squires / PlanetSquires Software" & vbcrlf & vbcrlf & _
                       "Special thanks to:" & vbcrlf & _
                       "José Roca (WinFBX Framework)", _
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
         ' Process clicking on the panels in the statusbar
         If id = IDC_FRMMAIN_STATUSBAR Then   
            Dim lpnm As NMMOUSE Ptr
            lpnm = Cast(NMMOUSE Ptr, pNMHDR)
            If lpnm->dwItemSpec = 0 Then   ' compiler panel is panel 0
               frmCompileConfig_Show( HWND ) 
            End If
            If lpnm->dwItemSpec = 1 Then   ' line/column panel 
               dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr
               if pDoc THEN Function = frmGoto_Show(pDoc->hWindow(0))
            end if
            If lpnm->dwItemSpec = 2 Then   ' ProjectFiletype 
               hPopupMenu = CreateStatusBarFileTypeContextMenu
               GetCursorPos @pt
               TrackPopupMenu(hPopUpMenu, 0, pt.x, pt.y, 0, HWnd, ByVal Null)
               DestroyMenu hPopUpMenu
            end if
            If lpnm->dwItemSpec = 3 Then   ' Ansi/Unicode 
               hPopupMenu = CreateStatusBarFileEncodingContextMenu
               GetCursorPos @pt
               TrackPopupMenu(hPopUpMenu, 0, pt.x, pt.y, 0, HWnd, ByVal Null)
               DestroyMenu hPopUpMenu
            end if   
            If lpnm->dwItemSpec = 4 Then   ' Compile results is 
               ' Toggle show/hide compile results
               ShowWindow HWND_FRMOUTPUT, Iif(IsWindowVisible(HWND_FRMOUTPUT), SW_HIDE, SW_SHOW)
               frmMain_PositionWindows(HWND_FRMMAIN)
            End If   
         End If

      Case TBN_DROPDOWN
         If id = IDC_FRMMAIN_TOOLBAR Then
            Dim ptbn As TBNOTIFY Ptr = Cast(TBNOTIFY Ptr, pNMHDR)
            Select Case ptbn->iItem
               Case IDM_FILENEW    ' Dropdown new file menu
                  SendMessage(GetDlgItem(HWnd, IDC_FRMMAIN_REBAR), RB_GETRECT, 0, Cast(LPARAM, @rc))
                  MapWindowPoints(ptbn->hdr.hwndFrom, HWND_DESKTOP, Cast(Point Ptr, @rc), 2)
                  frmTemplates_Show(HWnd, rc.Left, rc.Bottom)
               Case IDM_FILEOPEN    ' Dropdown most recently used files
                  SendMessage(GetDlgItem(HWnd, IDC_FRMMAIN_REBAR), RB_GETRECT, 0, Cast(LPARAM, @rc))
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
            pDoc = gTTabCtl.GetActiveDocumentPtr()
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
            Case SendMessage(GetDlgItem(GetDlgItem(HWnd, IDC_FRMMAIN_REBAR), IDC_FRMMAIN_TOOLBAR), TB_GETTOOLTIPS, 0, 0)
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
                  Case IDM_QUICKRUN        : wszText = L(284,"Quick Run (Ctrl+Shift+F5)")
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
      Dim As FILETIME ft
      Dim As Long i, idx, nDocumentCount
      Dim As Double nSerial
      Dim As String sText
      Dim pDoc As clsDocument Ptr
      
      for idx = 0 to gApp.GetProjectCount
         
         nDocumentCount = gApp.Projects(idx).GetDocumentCount
         For i = 0 To nDocumentCount - 1
            pDoc = gApp.Projects(idx).GetDocumentPtr(i)
            If pDoc Then
               ' Bypass any 'new' untitled files.
               If pDoc->IsNewFlag Then Continue For
               
               ' Has the external file been deleted or is now not available?
               if AfxFileExists(pDoc->DiskFilename) = false THEN
                  If MessageBox( HWND_FRMMAIN, _
                                  pDoc->DiskFilename & vbCrLf & _ 
                                  L(286, "This document has been deleted...?"), _
                                  L(267, "File Changed"), MB_ICONQUESTION Or MB_YESNO) = IDYES Then
                     ' Keep the file open and simply mark it as dirty so it can be prompted to be saved.
                     pDoc->UserModified = true                
                  else
                     ' No. Close the file.
                     OnCommand_FileClose(HWND)
                  end if
               
               else
                  ' Compare the disk file date time to the value currently
                  ' stored in document class.
                  ft = AfxGetFileLastWriteTime(pDoc->DiskFilename)

                  If AfxFileTimeToVariantTime(ft) <> AfxFileTimeToVariantTime(pDoc->DateFileTime) Then
                     OpenSelectedDocument( pDoc, "" )

                     If MessageBox( HWND_FRMMAIN, _
                                     pDoc->DiskFilename & vbCrLf & _ 
                                     L(266, "File was changed by another application. Reload it?"), _
                                     L(267, "File Changed"), MB_ICONQUESTION Or MB_YESNO) = IDYES Then
                        dim sText as string
                        if GetFileToString(pDoc->DiskFilename, sText, pDoc) = false then
                           pDoc->SetText( sText ) 
                           pDoc->DateFileTime = AfxGetFileLastWriteTime( pDoc->DiskFilename )
                        else
                           print "Error opening: "; pDoc->DiskFilename
                        End If
                     End If                
                  End If
               end if
            
            End If
         Next
      NEXT
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
   
   Dim pDoc As clsDocument Ptr = gTTabCtl.GetActiveDocumentPtr()
   If pDoc = 0 Then Return 0

   SetFocus hWndContext
   pt.x = xPos
   pt.y = yPos
   ScreenToClient hWndContext, @pt
   if pDoc->IsDesigner THEN
      ' Right-click popup menu for visual design form is 
      ' handled in the HandleDesignerRButtonDown function.
   else
      hPopupMenu = CreateScintillaContextMenu()
   END IF
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
   Dim wFileExt As WString * MAX_PATH
   
   For i = 0 To nCount - 1
      nLen = DragQueryFile(hDrop, i, @wszPath, MAX_PATH)
      ' Make sure it's a file, not a folder
      Dim fd As WIN32_FIND_DATAW
      Dim hFind As HANDLE = FindFirstFileW(@wszPath, @fd)
      If hFind <> INVALID_HANDLE_VALUE Then
         FindClose hFind
         If (fd.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY Then
            ' Determine what type of file is being dropped. If it is a project file .wfbe then
            ' use the OpenProject routines, otherwise open it as a regular file.
            wFileExt = AfxStrPathname( "EXTN", wszPath)
            wFileExt = Ucase(wFileExt)
            if wFileExt = ".WFBE" then
               OpenMRUProjectFile( HWnd, 0, @wszPath )
            Else   
               frmMain_OpenFileSafely(HWnd, _
                                       False, _    ' bIsNewFile
                                       False, _    ' bIsTemplate
                                       True,  _    ' bShowInTab
                                       false, _    ' bIsInclude
                                       @wszPath, _ ' pwszName
                                       0 )         ' pDocIn
            end if                           
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

   ' If configuration option to confirm closing editor is active then ask now.
   if gConfig.AskExit THEN
      if MessageBox( HWND_FRMMAIN, L(275,"Are you sure you want to exit?"), L(276,"Confirm"), _
                  MB_YESNOCANCEL or MB_ICONQUESTION Or MB_DEFBUTTON1 ) <> IDYES then
         return true
      end if            
   END IF
   
   ' Save whether the Explorer should be shown the next time the program is run.
   ' Also save the Explorer panel width. 
   gConfig.ShowExplorer = IsWindowVisible(HWND_FRMEXPLORER)
   gConfig.ShowExplorerWidth = AfxGetWindowWidth( HWND_FRMEXPLORER )
   
   ' Prevent Treeview flickering during closing
   dim as hwnd hTV = GetDlgItem(HWND_FRMEXPLORER, IDC_FRMEXPLORER_TREE)
   LockWindowUpdate(hTV)
   
   ' If a project(s) is already open then save/close it.
   do
      If gApp.IsProjectActive Then 
         if OnCommand_ProjectClose(HWnd) = 0 then 
            LockWindowUpdate(0): return 0
         end if 
      ELSE
         EXIT DO
      END IF
   LOOP   
   
   ' Close any open files asking to save any that are dirty 
   if OnCommand_FileClose(HWnd, True) = 0 then 
      LockWindowUpdate(0): return 0
   end if   
   
   DestroyWindow(HWnd)
   
   Function = 0
End Function


' ========================================================================================
' Process WM_DESTROY message for window/dialog: frmMain
' ========================================================================================
Function frmMain_OnDestroy(byval HWnd As HWnd) As LRESULT
   ImageList_Destroy Cast(HIMAGELIST, SendMessage(GetDlgItem(HWnd, IDC_FRMMAIN_TOOLBAR), TB_GETIMAGELIST, 0, 0))
   ImageList_Destroy Cast(HIMAGELIST, SendMessage(GetDlgItem(HWnd, IDC_FRMMAIN_TOOLBAR), TB_GETHOTIMAGELIST, 0, 0))
   ImageList_Destroy Cast(HIMAGELIST, SendMessage(GetDlgItem(HWnd, IDC_FRMMAIN_TOOLBAR), TB_GETDISABLEDIMAGELIST, 0, 0))
   ImageList_Destroy TabCtrl_GetImageList(gTTabCtl.hWindow) 
   
   ' Output the config settings to disk file
   gConfig.SaveToFile

   ' Disable drag and drop files
   DragAcceptFiles HWnd, False
   
   PostQuitMessage(0)
   Function = 0
End Function


' ========================================================================================
' Processes messages for the subclassed frmMain TabControl control.
' ========================================================================================
Function frmMain_TabCtl_SubclassProc ( _
                  ByVal HWnd   As HWnd, _                 ' // Control window handle
                  ByVal uMsg   As UINT, _                 ' // Type of message
                  ByVal wParam As WPARAM, _               ' // First message parameter
                  ByVal lParam As LPARAM, _               ' // Second message parameter
                  ByVal uIdSubclass As UINT_PTR, _        ' // The subclass ID
                  ByVal dwRefData As DWORD_PTR _          ' // Pointer to reference data
                  ) As LRESULT

   Dim pInfo As TCHITTESTINFO
   Dim As HWND hLV 
   Dim As POINT pt
   dim as long nCurSel
   dim as HMENU hPopupMenu 
   static as long oldX, oldY


   ' Convert our ENTER key presses into LBUTTONDBLCLK to process them similarly
   If (uMsg = WM_KEYUP) And (Loword(wParam) = VK_RETURN) Then uMsg = WM_LBUTTONDBLCLK
      
   Select Case uMsg

      Case WM_RBUTTONDOWN
         ' Get the tab number
         GetCursorPos( @pInfo.pt )
         ScreenToClient gTTabCtl.hWindow, @pInfo.pt
         nCurSel = TabCtrl_HitTest(gTTabCtl.hWindow, @pInfo)
         gTTabCtl.SetFocusTab(nCurSel)
         ' Create the popup menu
         hPopupMenu = CreateTopTabCtlContextMenu(nCurSel)
         GetCursorPos @pt
         TrackPopupMenu(hPopUpMenu, 0, pt.x, pt.y, 0, HWnd_FRMMAIN, ByVal Null)
         DestroyMenu hPopUpMenu
         Return True   ' prevent further processing that leads to WM_CONTEXTMENU

      Case WM_LBUTTONDOWN
         ' Get the position of the hit and determine if the close icon was clicked.
         GetCursorPos( @pInfo.pt )
         ScreenToClient gTTabCtl.hWindow, @pInfo.pt
         nCurSel = TabCtrl_HitTest(gTTabCtl.hWindow, @pInfo)
         gTTabCtl.SetFocusTab(nCurSel)
         ' Was the close icon pressed
         If pInfo.flags = TCHT_ONITEMICON Then 
            gTTabCtl.CloseTab()
         else
            gApp.bDragTabActive = true: oldX = pt.x: oldY = pt.y
            SetCursor( LoadCursor( null, IDC_HAND ))
         end if
            
      Case WM_LBUTTONUP
         If gApp.bDragTabActive Then 
            gApp.bDragTabActive = false
            dim as long idx, iCurTab, iDropTab, nCount
            dim As TC_HITTESTINFO HT
            Dim AS TCITEM TB_CURRENT, TB_TEMP 
            
            nCount = TabCtrl_GetItemCount(gTTabCtl.hWindow) 
            iCurTab = TabCtrl_GetCurSel(gTTabCtl.hWindow)
            GetCursorPos(@HT.pt): ScreenToClient( gTTabCtl.hWindow, @HT.pt)
            ' Get the tab that is under the cursor drop location
            iDropTab = SendMessage(gTTabCtl.hWindow, TCM_HITTEST, 0, cast(LPARAM, @HT)) 
            If (iCurTab = iDropTab) Or (iDropTab = -1) Then Exit function

            ' Save the current TCITEM so it can be set to the new position
            TB_CURRENT.Mask = TCIF_PARAM Or TCIF_TEXT or TCIF_IMAGE
            TabCtrl_GetItem( gTTabCtl.hWindow, iCurTab, @TB_CURRENT )

            ' Move all of the TCITEM's up or down dependng on the drop point in
            ' relation to the current point.
            TB_TEMP.Mask = TCIF_PARAM Or TCIF_TEXT or TCIF_IMAGE
            if iDropTab < iCurTab THEN
               for idx = iCurTab to iDropTab + 1 step -1
                  TabCtrl_GetItem( gTTabCtl.hWindow, idx-1, @TB_TEMP )
                  TabCtrl_SetItem( gTTabCtl.hWindow, idx, @TB_TEMP )
                  gTTabCtl.SetTabText(idx)
               NEXT
            elseif iDropTab > iCurTab THEN
               for idx = iCurTab to iDropTab - 1 
                  TabCtrl_GetItem( gTTabCtl.hWindow, idx+1, @TB_TEMP )
                  TabCtrl_SetItem( gTTabCtl.hWindow, idx, @TB_TEMP )
                  gTTabCtl.SetTabText(idx)
               NEXT
            END IF

            ' Finally, set to the current tab (ie. the drop tab)
            TabCtrl_SetItem( gTTabCtl.hWindow, iDropTab, @TB_CURRENT )
            gTTabCtl.SetTabText(iDropTab)
            gTTabCtl.SetFocusTab(iDropTab)

         end if
         gApp.bDragTabActive = false
         ' Reset the mouse pointer
         SetCursor( LoadCursor( null, IDC_ARROW ))

      Case WM_DESTROY
         ' REQUIRED: Remove control subclassing
         RemoveWindowSubclass HWnd, @frmOutput_Listview_SubclassProc, uIdSubclass

   End Select

   ' Default processing of Windows messages
   Function = DefSubclassProc(HWnd, uMsg, wParam, lParam)

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
      HANDLE_MSG (HWnd, WM_PAINT,       frmMain_OnPaint)
      HANDLE_MSG (HWnd, WM_SIZE,        frmMain_OnSize)
      HANDLE_MSG (HWnd, WM_CLOSE,       frmMain_OnClose)
      HANDLE_MSG (HWnd, WM_DESTROY,     frmMain_OnDestroy)
      HANDLE_MSG (HWnd, WM_COMMAND,     frmMain_OnCommand)
      HANDLE_MSG (HWnd, WM_NOTIFY,      frmMain_OnNotify)
      HANDLE_MSG (HWnd, WM_ACTIVATEAPP, frmMain_OnActivateApp)
      HANDLE_MSG (HWnd, WM_CONTEXTMENU, frmMain_OnContextMenu)
      HANDLE_MSG (HWnd, WM_DROPFILES,   frmMain_OnDropFiles)

      
      case MSG_USER_SHOWAUTOCOMPLETE
         return ShowAutocompleteList()

      case MSG_USER_APPENDEQUALSSIGN
         ' The = key was used to terminate a popup autocomplete. Take that
         ' character and format it with a space before and a space after.
         Dim pDoc As clsDocument Ptr = gTTabCtl.GetActiveDocumentPtr()
         if pDoc THEN
            dim as hwnd hEdit = pDoc->hWndActiveScintilla
            Dim As Long nPos = SciExec( hEdit, SCI_GETCURRENTPOS, 0, 0)
            dim as string sWord = " = "
            SciExec( hEdit, SCI_SETSEL, nPos-1, nPos)
            SciExec( hEdit, SCI_REPLACESEL, 0, Cast(LPARAM, Strptr(sWord)))
            nPos = SciExec( hEdit, SCI_GETCURRENTPOS, 0, 0)
            SciExec( hEdit, SCI_SETSEL, nPos, nPos)
         end if   
              
      case WM_MOUSEMOVE
         ' Determine if we have mouse over a edit split area
         Dim pDoc As clsDocument Ptr = gTTabCtl.GetActiveDocumentPtr()
         if pDoc THEN
            dim as POINT pt = (LoWord(lParam), HiWord(lParam))
            if pDoc->bSizing then
               if pt.y <> pDoc->ptPrev.y then
                  pDoc->SplitY = pDoc->SplitY + (pt.y - pDoc->ptPrev.y)
                  frmMain_PositionWindows(HWND)
                  pDoc->ptPrev.y = pt.y
               end if   
            else   
               if PtInRect(@pDoc->rcSplitButton, pt) THEN SetCursor(ghCursorNS)
            END if
         end if
         
      Case WM_LBUTTONDOWN
         ' Are we over a split edit area
         dim as POINT pt = (LoWord(lParam), hiword(lParam))
         Dim pDoc As clsDocument Ptr = gTTabCtl.GetActiveDocumentPtr()
         if pDoc THEN 
            SetCapture(hWnd)
            if PtInRect(@pDoc->rcSplitButton, pt) THEN 
               pDoc->bSizing = true
               pDoc->SplitY = pDoc->rcSplitButton.top  
               pDoc->ptPrev.y = pt.y
               SetCursor(ghCursorNS)
            END IF
         end if

      case WM_LBUTTONUP
         Dim pDoc As clsDocument Ptr = gTTabCtl.GetActiveDocumentPtr()
         if pDoc THEN pDoc->bSizing = false
         ReleaseCapture
         
      Case WM_LBUTTONDBLCLK 
         ' Are we over a split edit area (toggle off the split)
         dim as POINT pt = (LoWord(lParam), hiword(lParam))
         Dim pDoc As clsDocument Ptr = gTTabCtl.GetActiveDocumentPtr()
         if pDoc THEN 
            if PtInRect(@pDoc->rcSplitButton, pt) THEN 
               pDoc->bSizing = false
               pDoc->SplitY = 0
               frmMain_PositionWindows(HWND)
            END IF
         end if
          
      Case WM_NCACTIVATE
         ' Ensure that the caption bar for any DesignerForm retains its active 
         ' state. This looks better than having it lose and gain active colors. 
         If wParam Then
            ' The DesignerForm never receives the WM_NCACTIVATE message because it
            ' is a child form. We need to manually send it the message whenever the
            ' focus of the Designer form changes (i.e. whenever we switch back and
            ' forth away from the main application window).
            PostMessage HWND, MSG_USER_SETFOCUS, 0, 0
         End If
          
      Case WM_SYSCOMMAND
         If (wParam And &HFFF0) = SC_CLOSE Then
            SendMessage( HWND, WM_CLOSE, wParam, lParam )
            Exit Function
         End If
            
      case WM_SETCURSOR
         if (gApp.bDragTabActive = true) Then 
            ' Change mouse cursor because Tab in TabCtl being moved
            SetCursor( LoadCursor( null, IDC_HAND ))
            return true
         elseif (gCompiling = true) then
            SetCursor( LoadCursor(0, IDC_WAIT) )
         end if
            
      Case WM_ERASEBKGND
         ' Code to prevent main application space flickering. Only brush the
         ' background if no tabs are open.
         If TabCtrl_GetItemCount(gTTabCtl.hWindow) > 0 Then 
            Return True      ' prevents painting the background
         End If
         
      case WM_CAPTURECHANGED
         gApp.bDragTabActive = false
         
      Case WM_SETFOCUS
         frmMain_SetFocusToCurrentCodeWindow

      Case WM_INITMENUPOPUP
         ' A top menu is about to be displayed. Determine what menu items to enable/disable. We
         ' do all of the topmenu items here rather than deal with individual popup submenus.
         UpdateMRUMenu( Cast(HMENU, wParam) )
         UpdateMRUProjectMenu( Cast(HMENU, wParam) )
         frmMain_ChangeTopMenuStates()

      CASE WM_COPYDATA      ' used during processing of commandline
         dim pDataToGet AS COPYDATASTRUCT ptr
         Dim pwszArg As WString ptr 
         Dim wszExt As WString * MAX_PATH

         pDataToGet = cast(COPYDATASTRUCT ptr, lParam)
         pwszArg = pDataToGet->lpData

         ' We have a valid filename so determine what type it is.
         wszExt = AfxStrPathname( "EXTN", *pwszArg )
         wszExt = Ucase(wszExt)
          
         Select Case wszExt
            Case ".WFBE"    ' project file
               ' Pass the info to our generic project open function to handle everything.
               OpenMRUProjectFile( HWnd, 0, pwszArg )
             
            Case Else   ' .bas, .bi, .rc, etc...
               frmMain_OpenFileSafely( HWND, _
                                       False, _    ' bIsNewFile
                                       False, _    ' bIsTemplate
                                       True,  _    ' bShowInTab
                                       false, _    ' bIsInclude
                                       pwszArg, _  ' pwszName
                                       0 )         ' pDocIn
         End Select   
          

      ''  CUSTOM MESSAGES
      Case MSG_USER_SETFOCUS
         ' Set focus to current Scintilla window and update the document
         ' display such as Line#, Col#, Filename, etc.
         Dim pDoc As clsDocument Ptr = gTTabCtl.GetActiveDocumentPtr()
         if pDoc THEN
            frmMain_UpdateLineCol(HWnd)
            if pDoc->IsDesigner THEN
               SendMessage pDoc->hWndForm, WM_NCACTIVATE, true, 0
            END IF
            SetFocus(pDoc->hWndActiveScintilla)
         END IF

      Case MSG_USER_PROCESS_COMMANDLINE 
         ' process any command line arguments that may have been passed to the program.
         frmMain_ProcessCommandLine(HWnd)


      '-------------------------------------------------------------------------------
      '                        Scrollbar Handler
      '-------------------------------------------------------------------------------
      Case WM_VSCROLL
         Dim pDoc As clsDocument Ptr = gTTabCtl.GetActiveDocumentPtr()
         If pDoc = 0 Then Return 0

         dim as long OldPos = pDoc->ScrInfo.nPos

         Select Case Loword(wParam)
            Case SB_THUMBTRACK
               GetScrollInfo(pDoc->hScrollbar, SB_CTL, @pDoc->ScrInfo)
               pDoc->ScrInfo.nPos = pDoc->ScrInfo.nTrackPos
            Case SB_LINEDOWN:   pDoc->ScrInfo.nPos += 1
            Case SB_LINEUP:     pDoc->ScrInfo.nPos -= 1
            Case SB_PAGEDOWN:   pDoc->ScrInfo.nPos += pDoc->ScrInfo.nPage - 1  
            Case SB_PAGEUP:     pDoc->ScrInfo.nPos -= pDoc->ScrInfo.nPage - 1  
         End Select 
          
         ' If the current position hasn't changed, do nothing.
         If pDoc->ScrInfo.nPos = oldPos Then return true

         ' Don't exceed range boundries
         if pDoc->ScrInfo.nPos < pDoc->ScrInfo.nMin then
            pDoc->ScrInfo.nPos = pDoc->ScrInfo.nMin         
         END IF
         if pDoc->ScrInfo.nPos > pDoc->ScrInfo.nMax - pDoc->ScrInfo.nPage + 1 then
            pDoc->ScrInfo.nPos = pDoc->ScrInfo.nMax - pDoc->ScrInfo.nPage + 1        
         END IF

         dim as long lParm = pDoc->ScrInfo.nPos - oldPos        ' Amount/direction to V scroll
         SciExec(pDoc->hWindow(0), SCI_LINESCROLL, 0, lParm)
         SetScrollInfo(pDoc->hScrollbar, SB_CTL, @pDoc->ScrInfo, TRUE)

   End Select

   ' for messages that we don't deal with
   Function = DefWindowProc(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' frmMain_Show
' ========================================================================================
Function frmMain_Show( ByVal hWndParent As HWnd, _
                       ByVal nCmdShow   As Long = 0 _
                       ) As Long

   '  Create the main window and child controls
   Dim pWindow As CWindow Ptr = New CWindow("WinFBE_Class")
   ' Comment out the next line to let WinFBE use the current active system DPI setting.
   'pWindow->DPI = 144   ' eg. 144 or any other value (96 is default)
   
   ' Disable background erasing by only assigning the one style
   pWindow->ClassStyle = CS_DBLCLKS
   
   HWND_FRMMAIN = pWindow->Create(Null, APPNAME, @frmMain_WndProc )

   ' Set the small and large icon for the main window (must be set after main window is created)
   pWindow->BigIcon   =  LoadImage( pWindow->InstanceHandle, "IMAGE_AAA_MAINICON", IMAGE_ICON, 32, 32, LR_SHARED)
   pWindow->SmallIcon =  LoadImage( pWindow->InstanceHandle, "IMAGE_AAA_MAINICON", IMAGE_ICON, 16, 16, LR_SHARED)

   ' Load the good and bad compile icons
   Dim cx As Long 
   cx = 16 * (pWindow->DPI \ 96)
   ghIconGood = LoadImage( pWindow->InstanceHandle, "IMAGE_GOODCOMPILE", IMAGE_ICON, cx, cx, LR_DEFAULTCOLOR)
   ghIconBad  = LoadImage( pWindow->InstanceHandle, "IMAGE_BADCOMPILE", IMAGE_ICON, cx, cx, LR_DEFAULTCOLOR)

   ' Create the main application topmenu
   HWND_FRMMAIN_TOPMENU = frmMain_BuildMenu(pWindow)
   SetMenu pWindow->hWindow, HWND_FRMMAIN_TOPMENU

   ' Create the main application toolbar
   frmMain_CreateToolbar(pWindow)
   frmMain_ChangeToolbarButtonsState

   ' Load the build configurations into the toolbar
   LoadBuildComboBox()

   ' Add a status bar
   Dim As HWnd hStatusbar = pWindow->AddControl("Statusbar", pWindow->hWindow, IDC_FRMMAIN_STATUSBAR)
   Dim As Long rgParts(5)
   rgParts(0) = pWindow->ScaleX(200)   ' Build configuration
   rgParts(1) = pWindow->ScaleX(400)   ' Line/Col
   rgParts(2) = pWindow->ScaleX(500)   ' Normal/Module/Main/Resource
   rgParts(3) = pWindow->ScaleX(600)   ' Unicode/ANSI
   rgParts(4) = -1
   If SendMessage( hStatusBar, SB_SETPARTS, 5, Cast(LPARAM, @rgParts(0)) ) <> 0 Then
      SendMessage( hStatusBar, SB_SIMPLE, False, 0 )
   End If


   ' Add the top tabcontrol (initially not visible)
   gTTabCtl.hWindow = pWindow->AddControl("TABCONTROL", pWindow->hWindow, _
       IDC_FRMMAIN_TOPTABCONTROL, "", 0, 0, 0, 24, _
       WS_CHILD Or WS_TABSTOP Or TCS_SINGLELINE Or TCS_RAGGEDRIGHT Or TCS_HOTTRACK Or _
       TCS_TABS Or TCS_FOCUSNEVER Or TCS_FORCEICONLEFT, WS_EX_LEFT Or WS_EX_LTRREADING, , _
       Cast(WNDPROC, @frmMain_TabCtl_SubclassProc), IDC_FRMMAIN_TOPTABCONTROL, Cast(DWORD_PTR, @pWindow))
    
   ' Create the imagelist for the TopTabControl tabs
   cx = 16 * (pWindow->DPI \ 96)
   Dim As HICON hIcon = LoadIcon(pWindow->InstanceHandle, "IMAGE_TABCLOSE")
   Dim As HANDLE hImageList = ImageList_Create(cx, cx, ILC_MASK Or ILC_COLOR32, 1, 0)
   ImageList_AddIcon(hImageList, hIcon)
   DestroyIcon( hIcon )
   TabCtrl_SetImageList(gTTabCtl.hWindow, hImageList)
                                                        
   ' Create the ImageLists that will hold the images for the Explorer and Recent treeviews
   cx = 24 * (pWindow->DPI \ 96)
   ghImageListNormal = ImageList_Create( cx, cx, ILC_COLOR32 Or ILC_MASK, 4, 1)
   gidxImageBlank  = AfxGdipAddIconFromRes(ghImageListNormal, pWindow->InstanceHandle, "IMAGE_BLANK")
   gidxImageClosed = AfxGdipAddIconFromRes(ghImageListNormal, pWindow->InstanceHandle, "IMAGE_CLOSED")
   gidxImageOpened = AfxGdipAddIconFromRes(ghImageListNormal, pWindow->InstanceHandle, "IMAGE_OPENED")
   gidxImageCode   = AfxGdipAddIconFromRes(ghImageListNormal, pWindow->InstanceHandle, "IMAGE_CODE")

   ' Load the splitter cursors
   ghCursorNS = LoadImage( Null, MAKEINTRESOURCEW(OCR_SIZENS), IMAGE_CURSOR, 0, 0, LR_DEFAULTSIZE Or LR_SHARED)

   ' Load the graphisc PNG pointers for the scintilla popup autocomplete listbox.
   gApp.pImageAutocompleteBaseType = LoadPNGfromRes(pWindow->InstanceHandle, "IMAGE_BASETYPE")     
   gApp.pImageAutocompleteClass    = LoadPNGfromRes(pWindow->InstanceHandle, "IMAGE_CLASS")     
   gApp.pImageAutocompleteMethod   = LoadPNGfromRes(pWindow->InstanceHandle, "IMAGE_METHOD")     
   gApp.pImageAutocompleteProperty = LoadPNGfromRes(pWindow->InstanceHandle, "IMAGE_PROPERTY")  

   ' Create the various child windows
   frmExplorer_Show( HWND_FRMMAIN )
   frmRecent_Show( HWND_FRMEXPLORER )     ' child of frmExplorer
   frmOutput_Show( HWND_FRMMAIN )

   ' Create the UserTools accelerator table
   CreateUserTools_AcceleratorTable            
   
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


' TODO: Remove SetTimer (debug for showing GDI count)
'SetTimer(HWND_FRMMAIN, 1000, 500, 0)
   
   ' Process windows events
   Dim uMsg As MSG
   
   ' Message loop
   Do While GetMessage(@uMsg, Null, 0, 0)


'      if uMsg.message = WM_TIMER then
'         Dim As HWnd hStatusbar = GetDlgItem(HWnd_FRMMAIN, IDC_FRMMAIN_STATUSBAR)
'         static GUILeak as wstring * 100
'         dim hProcess AS HANDLE
'         dim dwProcessID as dword
'         getwindowthreadprocessid( HWnd_FRMMAIN, @dwProcessID)
'         hProcess = openprocess(PROCESS_QUERY_INFORMATION OR PROCESS_VM_READ, 0, dwProcessID)
'         GUILeak = "GDICount: " & str(getguiresources(hProcess, GR_GDIOBJECTS))
'         StatusBar_SetText(hStatusbar, 0, GUILeak)
'      end if


      ' Processes accelerator keys for menu commands
      If (pWindow->AccelHandle = 0) Orelse (TranslateAccelerator(pWindow->hWindow, pWindow->AccelHandle, @uMsg) = 0) then
         
         if (ghAccelUserTools = 0) orelse (TranslateAccelerator(pWindow->hWindow, ghAccelUserTools, @uMsg) = 0) Then 

            ' Intercept all MouseMove messages in order to test if we are over the splitter area. This
            ' easier and more efficent than trying to deal with mouse messages at the child control/form level.
            Select case uMsg.message
               Case WM_LBUTTONDOWN: Splitter_OnLButtonDown
               Case WM_LBUTTONUP:   Splitter_OnLButtonUp
               Case WM_MOUSEMOVE:   Splitter_OnMouseMove
            End Select      

            ' Close any open Find/Replace window with ESC key
            if (uMsg.message = WM_KEYDOWN) andAlso (uMsg.wParam = VK_ESCAPE) THEN
               if IsWindowVisible( HWND_FRMFINDREPLACE ) THEN SendMessage( HWND_FRMFINDREPLACE, WM_CLOSE, 0, 0 )
            END IF

            ' Handle any keypress that would move or resize control(s) on a Designer Form.
            if uMsg.message = WM_KEYDOWN THEN
               ' Ensure that an arrow key was pressed
               select CASE uMsg.wParam

                  case VK_LEFT, VK_RIGHT, VK_UP, VK_DOWN
                     dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr
                     if pDoc THEN
                        if pDoc->IsDesigner THEN
                           select case GetFocus
                              case GetDlgItem(HWND_FRMVDTOOLBOX, IDC_FRMVDTOOLBOX_LSTTOOLBOX), _
                                   GetDlgItem(HWND_FRMVDTOOLBOX, IDC_FRMVDTOOLBOX_LSTPROPERTIES), _
                                   GetDlgItem(HWND_FRMEXPLORER, IDC_FRMEXPLORER_TREE)
                              CASE ELSE
                                 if pDoc->Controls.SelectedControlsCount > 0 then
                                    ' Move control(s)
                                    if (GetAsyncKeyState(VK_CONTROL) and &H8000) then
                                       KeyboardMoveControls(pDoc, uMsg.wParam)
                                       continue do
                                    end if
                                    ' Resize control(s)   
                                    if (GetAsyncKeyState(VK_SHIFT) and &H8000) THEN
                                       KeyboardResizeControls(pDoc, uMsg.wParam)
                                       continue do
                                    end if   
                                    ' Move active selected control focus within group of controls
                                    KeyboardCycleActiveControls(pDoc, uMsg.wParam)
                                    continue do
                                 end if   
                           END SELECT
                        END IF
                     END IF
               END SELECT
            end if
            
            
            ' Determines whether a message is intended for the specified
            ' dialog box and, if it is, processes the message.
            ' Ensure keystrokes like TAB are properly handled by the modeless dialogs
            If IsDialogMessage( HWND_FRMFINDINFILES, @uMsg ) Then Continue Do
            If IsDialogMessage( HWND_FRMFINDREPLACE, @uMsg ) Then Continue Do
            
            ' Check for any QuickRun exes that can be deleted.
            gApp.CheckQuickRunExe
            
            If IsDialogMessage(pWindow->hWindow, @uMsg) = 0 Then
               TranslateMessage @uMsg    ' Translates virtual-key messages into character messages.
               DispatchMessage @uMsg    ' Dispatches a message to a window procedure.
            End If
         end if
         
      End If
   Loop
   Function = uMsg.wParam

   
   ' Destroy the Explorer and Recent treeview imagelist
   ImageList_Destroy(ghImageListNormal)

   If ghIconGood Then DestroyIcon(ghIconGood) 
   If ghIconBad  Then DestroyIcon(ghIconBad)
   if ghAccelUserTools THEN DestroyAcceleratorTable(ghAccelUserTools)

   if gApp.pImageAutocompleteBaseType then deallocate(gApp.pImageAutocompleteBaseType)
   if gApp.pImageAutocompleteClass    then deallocate(gApp.pImageAutocompleteClass)
   if gApp.pImageAutocompleteMethod   then deallocate(gApp.pImageAutocompleteMethod)
   if gApp.pImageAutocompleteProperty then deallocate(gApp.pImageAutocompleteProperty)

   ' Delete the allocated memory for the various child windows
   pWindow = AfxCWindowPtr(HWND_FRMEXPLORER):       Delete pWindow
   pWindow = AfxCWindowPtr(HWND_FRMRECENT):         Delete pWindow
   pWindow = AfxCWindowPtr(HWND_FRMOUTPUT):         Delete pWindow
   pWindow = AfxCWindowPtr(HWND_FRMMAIN):           Delete pWindow

End Function


