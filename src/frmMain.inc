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

#include once "frmMain.bi"
#include once "clsDocument.bi"
#include once "frmUserTools.bi"
#include once "frmSnippets.bi"
#include once "frmProjectOptions.bi"
#include once "frmBuildConfig.bi"
#include once "frmExplorer.bi"
#include once "modMRU.bi"
#include once "modAutoInsert.bi"
#include once "modToolbar.bi"
#include once "modMenus.bi"
#include once "modCompile.bi"


' Global array to hold cut/copy/paste controls
dim shared gCopyControls(any) as clsControl


' ========================================================================================
' Load the Files combobox
' ========================================================================================
public Function frmMain_RemoveFromComboFiles( byval pDoc as clsDocument ptr ) As Long
   if pDoc = 0 then exit function
   
   dim as HWND hComboFiles = GetDlgItem( HWND_FRMMAIN, IDC_FRMMAIN_COMBOFILES )
   dim as long nCount = ComboBox_GetCount( hComboFiles )

   for i as long = 0 to nCount - 1
      if ComboBox_GetItemData( hComboFiles, i ) = pDoc then
         ComboBox_DeleteString( hComboFiles, i )
         exit for
      end if
   next
   
   function = 0
end function


' ========================================================================================
' Add a file to the Files combobox
' ========================================================================================
public Function frmMain_AddToComboFiles( byval pDoc as clsDocument ptr ) As Long
   if pDoc = 0 then exit function
   
   dim as HWND hComboFiles = GetDlgItem( HWND_FRMMAIN, IDC_FRMMAIN_COMBOFILES )
   
   dim as long idx = ComboBox_AddString( hComboFiles, AfxStrPathname( "NAMEX", pDoc->DiskFilename ).sptr )
   ' Save the pDoc in the itemData in order to quickly load the file later when
   ' selected from the Files combobox.
   ComboBox_SetItemData( hComboFiles, idx, cast(LPARAM, pDoc) )
   
   function = 0
end function


' ========================================================================================
' Load the Files combobox
' ========================================================================================
public Function frmMain_LoadComboFiles() As Long
   
   dim as HWND hComboFiles = GetDlgItem( HWND_FRMMAIN, IDC_FRMMAIN_COMBOFILES )
   ComboBox_ResetContent( hComboFiles )
   
   dim pDoc as clsDocument ptr = gApp.pDocList
   do until pDoc = 0
      frmMain_AddToComboFiles( pDoc )
      pDoc = pDoc->pDocNext
   loop

   function = 0
end function


' ========================================================================================
' Load the Functions combobox
' ========================================================================================
public Function frmMain_LoadComboFunctions() As Long
   
   Dim As Hwnd hTree = GetDlgItem(HWND_FRMEXPLORER, IDC_FRMEXPLORER_TREE)
   dim as HWND hComboFunctions = GetDlgItem( HWND_FRMMAIN, IDC_FRMMAIN_COMBOFUNCTIONS )
   ComboBox_ResetContent( hComboFunctions )

   Dim pDoc As clsDocument Ptr = gTTabCtl.GetActiveDocumentPtr()
   if pDoc = 0 then exit function
   
   dim wszText as wstring * MAX_PATH
   
   ' Add a default entry for when not positioned within a function.
   wszText = L(218,"(Module level)")
   ComboBox_AddString( hComboFunctions, @wszText )

   ' Get all functions in the file
   Dim As HTREEITEM hChild = Treeview_GetChild( hTree, pDoc->hNodeExplorer )
   do until hChild = 0 
      Treeview_GetItemText(hTree, hChild, @wszText, MAX_PATH)
      ComboBox_AddString( hComboFunctions, @wszText )
      hChild = Treeview_GetNextSibling(hTree, hChild)
   loop

   function = 0
end function


' ========================================================================================
' Select the Files combobox item based on incoming filename
' ========================================================================================
public Function frmMain_SelectComboFiles() As Long
   Dim pDoc As clsDocument Ptr = gTTabCtl.GetActiveDocumentPtr()
   if pDoc = 0 then exit function
   
   dim as HWND hComboFiles = GetDlgItem( HWND_FRMMAIN, IDC_FRMMAIN_COMBOFILES )
   dim as CWSTR wszFilename = AfxStrPathname( "NAMEX", pDoc->DiskFilename )
   dim as long idx = ComboBox_FindStringExact( hComboFiles, -1, wszFilename.sptr )
   ComboBox_SetCurSel( hComboFiles, idx )
   
   function = 0
end function


' ========================================================================================
' Select the Files combobox item based on incoming filename
' ========================================================================================
private Function frmMain_ComboFunctionsFindString( byval sFunctionName as string, _
                                                   byval nGetSet as long _
                                                   ) As Long
   
   dim as HWND hComboFunctions = GetDlgItem( HWND_FRMMAIN, IDC_FRMMAIN_COMBOFUNCTIONS )
   dim as long nCount = ComboBox_GetCount( hComboFunctions )
   dim as long nResult = -1
   
   dim wszText as WString * MAX_PATH
   dim as string sLookFor = ucase( sFunctionName )
   
   select case nGetSet 
      case ClassProperty.Getter: sLookFor = sLookFor & " (GET)"
      case ClassProperty.Setter: sLookFor = sLookFor & " (SET)"
      case ClassProperty.ctor:   sLookFor = sLookFor & " (CTOR)"
      case ClassProperty.dtor:   sLookFor = sLookFor & " (DTOR)"
   end select
   
   for i as long = 0 to nCount - 1
      ComboBox_GetLBText( hComboFunctions, i, @wszText )
      if ucase( wszText ) = sLookFor then
         nResult = i
         exit for
      end if
   next

   function = nResult
end function


' ========================================================================================
' Select the Files combobox item based on incoming filename
' ========================================================================================
public Function frmMain_SelectComboFunctions() As Long
   Dim pDoc As clsDocument Ptr = gTTabCtl.GetActiveDocumentPtr()
   if pDoc = 0 then exit function
   
   dim as HWND hComboFunctions = GetDlgItem( HWND_FRMMAIN, IDC_FRMMAIN_COMBOFUNCTIONS )
   dim as string sFunctionName 
   dim as ClassProperty nGetSet
   
   pDoc->GetCurrentFunctionName( sFunctionName, nGetSet )
   
   ' If no Functions exist in the Functions combo then it obviously has not
   ' been loaded yet.
   if ComboBox_GetCount( hComboFunctions ) = 0 then frmMain_LoadComboFunctions
   dim as long idx = frmMain_ComboFunctionsFindString( sFunctionName, nGetSet )

   ' If the function name is not found then check to see if the file is dirty
   ' and do a reparse and second check of the function list. If still not found
   ' then we must be between functions. Set the index to 0 so that the "(Module level)" 
   ' entry is selected.
   if idx = -1 then 
      if pDoc->bNeedsParsing then
         ParseDocument( pDoc )
         frmMain_LoadComboFunctions
         idx = frmMain_ComboFunctionsFindString( sFunctionName, nGetSet )
      end if
   end if
   if idx = -1 then idx = 0
   ComboBox_SetCurSel( hComboFunctions, idx ) 
   
   function = 0
end function


' ========================================================================================
' Position editor to document/function under the current caret position
' ========================================================================================
private Function frmMain_GotoDefinition( ByVal pDoc As clsDocument Ptr ) As Long

   If pDoc = 0 Then Exit Function
   
   dim pData as DB2_DATA ptr
   dim as CWSTR wszFilename
   dim as string sFunctionName, sTrigger

   dim as hwnd hEdit = pDoc->hWndActiveScintilla
   
   ' Version 1.8.8 added ability to open an underlying #Include file if such
   ' a file exists on the current editing line.
   if IsCurrentLineIncludeFilename() then
      wszFilename = gApp.IncludeFilename 
   else

      ' Determine the word at the current caret position.
      sFunctionName = pDoc->GetWord
      
      ' Determine if the function name is part of a class. This would be identified
      ' by a preceeding "." or "->" symbol.  eg  pDoc->GetCurrentLineNumber
      dim as long curPos = SciExec( hEdit, SCI_GETCURRENTPOS, 0, 0)
      dim as long wordStartPos = SciExec( hEdit, SCI_WORDSTARTPOSITION, curPos, True)
      sTrigger = right( pDoc->GetWord( wordStartPos - 1 ), 1 )
      if sTrigger = ">" then sTrigger = "->"
      if ( sTrigger = "." ) orelse ( sTrigger = "->" ) then
         ' In the DereferenceLine function any "->" are converted to "." for simplicity.
         pData = DereferenceLine( pDoc, sTrigger, wordStartPos - 1 )
         if pData then 
            sFunctionName = pData->ElementName & "." & sFunctionName
            ' Set pData to null because it currently points at the TYPE definition and
            ' we the gdb2 searches below to actually find the sub/function.
            pData = 0
         end if
      end if   

      ' Search the list of Function names
      if pData = 0 then pData = gdb2.dbFindFunction( sFunctionName )
      if pData = 0 then pData = gdb2.dbFindSub( sFunctionName )
      if pData = 0 then pData = gdb2.dbFindProperty( sFunctionName )
      
      IF pData = 0 then
         MessageBox( HWND_FRMMAIN, L(224,"Sub/Function definition not found."), @WStr(APPNAME), MB_ICONWARNING)
         exit function
      end if
      wszFilename = pData->filename
   end if
   
   ' Save our current position in case user invokes "Last Position" option,   
   gLastPosition.pDoc       = pDoc
   gLastPosition.nFirstLine = SciExec( hEdit, SCI_GETFIRSTVISIBLELINE, 0, 0) 
   gLastPosition.nPosition  = SciExec( hEdit, SCI_GETCURRENTPOS, 0, 0) 
   OpenSelectedDocument( wszFilename, sFunctionName )

   Function = 0
End Function


' ========================================================================================
' Return to most previous position (invoked when "Last Position" is selected
' ========================================================================================
private Function frmMain_GotoLastPosition() As Long

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
' Position editor to Header/Source/Main/Resource document
' ========================================================================================
private Function frmMain_GotoFile( ByVal pDoc As clsDocument Ptr, _
                                   byval nMenuId as long _
                                   ) As Long

   If pDoc = 0 Then Exit Function
   
   dim pDocFile as clsDocument ptr
   dim as CWSTR wszFilename
   
   select case nMenuId
      case IDM_GOTOMAINFILE
         pDocFile = gApp.GetMainDocumentPtr
         if pDocFile then wszFilename = pDocFile->DiskFilename
         
      case IDM_GOTORESOURCEFILE
         pDocFile = gApp.GetResourceDocumentPtr
         if pDocFile then wszFilename = pDocFile->DiskFilename

      case IDM_GOTOHEADERFILE
         pDocFile = gApp.GetHeaderDocumentPtr( pDoc )
         if pDocFile then wszFilename = pDocFile->DiskFilename

      case IDM_GOTOSOURCEFILE
         pDocFile = gApp.GetSourceDocumentPtr( pDoc )
         if pDocFile then wszFilename = pDocFile->DiskFilename
   end select

   
   ' If the document when going to is the same as the one that we are 
   ' already in then no need to execute the movement.
   if pDocFile = pDoc then exit function
   
   ' Save our current position in case user invokes "Last Position" option,   
   dim as hwnd hEdit = pDoc->hWndActiveScintilla
   gLastPosition.pDoc       = pDoc
   gLastPosition.nFirstLine = SciExec( hEdit, SCI_GETFIRSTVISIBLELINE, 0, 0) 
   gLastPosition.nPosition  = SciExec( hEdit, SCI_GETCURRENTPOS, 0, 0) 

   OpenSelectedDocument( wszFilename )
   
   Function = 0
End Function


' ========================================================================================
' Update the main form statusbar. This is the only routine that updates
' the statusbar in the entire program.
' ========================================================================================
public Function frmMain_SetStatusbar() as long

   Dim As HWnd hStatusbar = GetDlgItem(HWND_FRMMAIN, IDC_FRMMAIN_STATUSBAR)

   ' Update the statusbar with the current Line/Col position
   Dim pDoc As clsDocument Ptr = gTTabCtl.GetActiveDocumentPtr()
   
   dim wszText as wstring * MAX_PATH

   ' PANEL (0) Build Configuration
   ' PANEL (1) Line/Col/Sel -or- VD coordinates
   ' PANEL (2) Filetype (Normal/Main/Resource/Module)
   ' PANEL (3) File Encoding (Unicode/Ansi) -or- "Parsing:"
   ' PANEL (4) Compiling Results -or- Filename being parsed
   dim wszPanel(4) As WString * MAX_PATH
   dim hIconPanel4 as HANDLE
   for i as long = lbound(wszPanel) to ubound(wszPanel)
      StatusBar_GetText( hStatusbar, i, wszPanel(i) )
   next
   
   ' Default Panel4 text to be whatever is stored in gApp. That string will always hold
   ' the latest file being parsed or the latest compile success/failure message. We will
   ' reset the value when files are finished loading but we always what to keep the latest
   ' compiled message.
   wszPanel(4) = gApp.wszPanelText
   hIconPanel4 = gApp.hIconPanel

   if (gApp.IsProjectLoading) orelse (gApp.IsFileLoading) then
      wszPanel(1) = ""
      wszPanel(2) = ""
      wszPanel(3) = L(66,"Parsing") & ": "
      wszPanel(4) = "(" & gApp.FileLoadingCount & ")  " & gApp.wszPanelText 
   
   elseif gApp.IsCompiling then
      ' The gApp info is set in modCompile
      
   elseif pDoc <> 0 then
      dim as HWND hEdit = pDoc->hWndActiveScintilla
      wszText = ""
      if (pDoc->IsDesigner) andalso (IsDesignerView(pDoc)) THEN 
         dim pCtrl as clsControl ptr = pDoc->Controls.GetActiveControl
         if pCtrl THEN
            wszText = "L:" & GetControlProperty(pCtrl, "LEFT") & ", T:" & GetControlProperty(pCtrl, "TOP") & "  ::  " & _
                      "W:" & GetControlProperty(pCtrl, "WIDTH") & " x H:" & GetControlProperty(pCtrl, "HEIGHT")
         END IF
      else
         Dim As Long curPos, nLine, nCol
         Dim As Long startPos, endPos, startLine, endLine, nLines 
         ' Retrieve the information and show it in the status bar
         curPos    = SciExec( hEdit, SCI_GETCURRENTPOS, 0, 0)
         nLine     = SciExec( hEdit, SCI_LINEFROMPOSITION, curPos, 0) 
         nCol      = SciExec( hEdit, SCI_GETCOLUMN, curPos, 0) 
         nLines    = SciExec( hEdit, SCI_GETLINECOUNT, 0, 0) 
         pDoc->GetSelectedLineRange(startLine, endLine, startPos, endPos )

         wszText = " Ln " & Str(nLine + 1) & ", Col " & Str(nCol + 1) 
         If endPos - startPos Then  ' there is selected text
            wszText = wszText + " (" + Str(endPos - startPos) + " selected)"
            ' Update whether Selection icon should be highlighted
            if IsWindowVisible(HWND_FRMFINDREPLACE) then frmFindReplace_ShowControls()
         else
            wszText = wszText + ", Lines " + Str(nLines)
         End If 
         pDoc->ScrInfo.nMax = SciExec(hEdit, SCI_GETLINECOUNT, 0, 0) - 1    
         pDoc->ScrInfo.nPos = SciExec(hEdit, SCI_GETFIRSTVISIBLELINE, 0, 0)
         SetScrollInfo(pDoc->hScrollbar, SB_CTL, @pDoc->ScrInfo, TRUE) 

         ' Only search for the current sub/function if the pDoc or linenumber has changed
         ' from the last time we checked. This ensures that we are not constantly checking lines
         ' every time a character is entered.
         static pDocLast as clsDocument ptr
         static nLineLast as long
         if ( pDocLast <> pDoc ) orelse ( nLineLast <> nLine ) then
            frmMain_SelectComboFunctions
         end if
         pDocLast = pDoc
         nLineLast = nLine
         
      end if
      wszPanel(1) = wszText
   
      
      wszText = ""
      Select Case pDoc->ProjectFileType
         Case FILETYPE_NORMAL:    wszText = ucase(L(210,"Normal")) 
         Case FILETYPE_MODULE:    wszText = ucase(L(211,"Module"))
         Case FILETYPE_MAIN:      wszText = ucase(L(212,"Main"))
         Case FILETYPE_RESOURCE:  wszText = ucase(L(213,"Resource"))
      End Select
      if (pDoc->IsDesigner) andalso (IsDesignerView(pDoc)) andalso (pDoc->ProjectFileType = FILETYPE_UNDEFINED) THEN 
         wszText = "FORM"
      end if
      wszPanel(2) = wszText
      
      
      wszText = ""
      select case pDoc->FileEncoding
         case FILE_ENCODING_UTF8_BOM
            wszText = "UTF-8 (BOM)"   
            SciExec( hEdit, SCI_SETCODEPAGE, SC_CP_UTF8, 0 )
         case FILE_ENCODING_UTF16_BOM
            wszText = "UTF-16 (BOM)"   
            SciExec( hEdit, SCI_SETCODEPAGE, SC_CP_UTF8, 0 )
         case else
            wszText = "ANSI"  
            SciExec( hEdit, SCI_SETCODEPAGE, 0, 0 )
      END select
      wszPanel(3) = wszText
      
      if SciExec( hEdit, SCI_GETREADONLY, 0, 0 ) then
         wszPanel(4) = L(222,"Read Only")
         hIconPanel4 = ghIconBad
      end if

   else      
      ' No document must be active and nothing else happening so blank out the statusbar.
      wszPanel(1) = ""
      wszPanel(2) = ""
      wszPanel(3) = "" 
      wszPanel(4) = "" 
      hIconPanel4 = 0
   END IF
   

   ' MAIN WINDOW CAPTION 
   wszText = iif( gApp.GetDocumentCount, APPNAMESHORT, APPNAME )
   If (gApp.IsProjectActive = true) or (gApp.IsProjectLoading = true) Then 
      wszText = wszText & " - [" & gApp.ProjectName & "]"
   end if
   if pDoc then 
      wszText = wszText & " - [" & pDoc->DiskFilename & "]"
      if SciExec( pDoc->hWndActiveScintilla, SCI_GETREADONLY, 0, 0 ) then
         wszText = wszText & " - [" & L(222,"Read Only") & "]"
      end if
   end if
   If wszText <> AfxGetWindowText( HWND_FRMMAIN ) Then AfxSetWindowText( HWND_FRMMAIN, wszText )


   ' TAB CONTROL FILENAME ( * modified flag )
   gTTabCtl.SetTabText(-1)


   ' BUILD CONFIGURATION
   wszText = ""
   if gApp.GetDocumentCount THEN
      dim as long nCurSel = ComboBox_GetCurSel(HWND_FRMMAIN_COMBOBUILDS)
      if nCurSel > -1 THEN wszText = " " & gConfig.Builds(nCurSel).wszDescription
   END IF
   wszPanel(0) = wszText
   

   ' UPDATE THE PANELS THAT HAVE CHANGED
   for i as long = lbound(wszPanel) to ubound(wszPanel)
      StatusBar_GetText( hStatusbar, i, @wszText )
      if wszText <> wszPanel(i) then StatusBar_SetText( hStatusbar, i, wszPanel(i) )
   next
   if StatusBar_GetIcon( hStatusBar, 4 ) <> hIconPanel4 then StatusBar_SetIcon( hStatusBar, 4, hIconPanel4 )

   frmMain_ChangeToolbarButtonsState
   
   Function = 0
End Function


' ========================================================================================
' Process any command line that was passed to the editor
' ========================================================================================
private Function frmMain_ProcessCommandLine( ByVal HWnd As HWnd ) As Long

   ' The incoming command line may contain a regular file to open or a project file.
   
   ' Command: A space-separated list of all command-line arguments is returned. When the 
   '          command line is parsed for arguments, everything between double quotes in 
   '          the parameter list will be considered as a single parameter, and is returned 
   '          with the double quotes.
   '          A value of zero (0) returns the name of the executable; and values of 
   '          one (1) and greater return each command-line argument.

   ' As of v1.7.4 start to use AfxCommand which is a unicode aware replacement for FB's
   ' built in COMMAND function (that is not unicode compliant).
   
   If Len(AfxCommand(1)) = 0 Then Exit Function
   
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
      wszArg = AfxCommand(i)
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
private Function Scintilla_OnNotify( ByVal HWnd As HWnd, _
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
            frmMain_SetStatusbar
         end if
         return true
         
      
      case SCN_MODIFYATTEMPTRO
         ' Attempting to modify a read-only document.
         MessageBeep(MB_ICONWARNING)
                  
                  
      Case SCN_MODIFIED
         if gApp.SuppressNotify then return true
         ' Show line and column. Only do on modification of text otherwise we will have
         ' a huge slowdown when notifications sent for UI updates.
         If (pNSC->modificationType And SC_MOD_INSERTTEXT) Or (pNSC->modificationType And SC_MOD_DELETETEXT) Then
            ' Set parsing flag to true. We use a parse flag rather than depending on the dirty
            ' flags because the file could stay dirty for a long time before it gets saved and 
            ' every call to the parser would happen even in cases where only directional 
            ' movement has occured with no text modifications at all.
            pDoc = gTTabCtl.GetActiveDocumentPtr()
            If pDoc Then 
               pDoc->bNeedsParsing = true
               ' 1.8.8 Tried changing the parse routine to only parse from the modified line
               ' to the end of the file. The problem with that is that the parser looks for starting
               ' Sub/Function and End Sub / End Function line. Same thing with TYPE/END TYPE and ENUMS.
               ' If we start parsing in the middle of one of those block structures then we will
               ' not have them recognized because only the end points will be found in the parse.
               ' For now, until a better algorithm is found, it is wise to simply reparse the 
               ' whole file.
               pDoc->ParseStartLine = 0  'pDoc->GetCurrentLineNumber
            end if
            frmMain_SetStatusbar
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
                     
               
      case SCN_AUTOCCANCELLED
         ' Destroy the secondary popup information window
         pDoc = gTTabCtl.GetActiveDocumentPtr()
         If pDoc Then 
            pDoc->AutoCompleteType = AUTOCOMPLETE_NONE
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
                            
            if chr(pNSC->ch) = "=" then
               ' Add a space before and after Equals sign whenever the user presses
               ' the equals sign to terminate ab autocomplete popup list.
               PostMessage(HWND, MSG_USER_APPENDEQUALSSIGN, 0, 0)
            end if
            return true   
         end if

      case SCN_AUTOCCHARDELETED    
         if gConfig.CodeTips then
            ' Use deleted a character while autocompletion list was active. Display
            ' new contents of list because the underlying word being typed as changed.
            ShowAutocompleteList(SCN_AUTOCCHARDELETED)
         end if
                                    
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
            pDoc->AutoCompleteType = AUTOCOMPLETE_NONE
            SciExec( hEdit, SCI_CALLTIPCANCEL, 0, 0 )
               
         case "("      ' Show code tip 
            pDoc->AutoCompleteType = AUTOCOMPLETE_NONE
            If PostMessage(HWND_FRMMAIN, MSG_USER_SHOWAUTOCOMPLETE, pNSC->ch, 0) then exit function

         case chr(13)  ' ENTER KEY PRESSED
            pDoc->AutoCompleteType = AUTOCOMPLETE_NONE
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
public Function frmMain_SetFocusToCurrentCodeWindow() As Long
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
private Function frmMain_OnPaint( ByVal HWnd As HWnd ) As LRESULT
            
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

      ' Paint under the Files/Functions comboboxes because the boxes do not
      ' cover the entire width fo the client area.
      Dim pWindow As CWindow Ptr = AfxCWindowPtr( HWND_FRMMAIN )
      if pWindow then
         dim as hwnd hEdit = iif( IsWindowVisible( pDoc->hWindow(1)), pDoc->hWindow(1), pDoc->hWindow(0) )
         dim as long nLeft, nTop
         AfxGetWindowLocation( hEdit, nLeft, nTop )
         GetClientRect( HWND_FRMMAIN, @rc )
         rc.left = nLeft
         rc.right = nLeft + AfxGetWindowWidth( hEdit ) + AfxGetWindowWidth( pDoc->hScrollbar )
         rc.bottom = nTop
         FillRect( hDC, @rc, GetSysColorBrush( COLOR_BTNFACE ) )
         rc.top = rc.bottom - pWindow->ScaleY(1)
         FillRect( hDC, @rc, GetSysColorBrush( COLOR_WINDOWFRAME ) )
      end if
         
   end if
   
   RestoreDC hDC, -1 
   
   EndPaint hWnd, @ps
   
   Function = 0
End Function



' ========================================================================================
' Position all child windows. Called manually and/or by WM_SIZE
' ========================================================================================
public Function frmMain_PositionWindows() As LRESULT
   
   Dim As HWnd hEdit, hTabCtl, hComboFiles, hComboFunctions
   dim as long nHeightTabControl, nLeft
   Dim As Rect rc
   
   Dim pWindow As CWindow Ptr = AfxCWindowPtr( HWND_FRMMAIN )
   if pWindow = 0 then exit function
   
   Dim pDoc As clsDocument Ptr = gTTabCtl.GetActiveDocumentPtr()
   
   ' Get the entire client area
   GetClientRect( HWND_FRMMAIN, @rc )
  
   hTabCtl = GetDlgItem( HWND_FRMMAIN, IDC_FRMMAIN_TOPTABCONTROL )
   hComboFiles = GetDlgItem( HWND_FRMMAIN, IDC_FRMMAIN_COMBOFILES )
   hComboFunctions = GetDlgItem( HWND_FRMMAIN, IDC_FRMMAIN_COMBOFUNCTIONS )

   Dim As Long nHeightRebar     = AfxGetWindowHeight( GetDlgItem(HWND_FRMMAIN, IDC_FRMMAIN_REBAR) )
   dim as long nHeightStatusBar = AfxGetWindowHeight( GetDlgItem(HWND_FRMMAIN, IDC_FRMMAIN_STATUSBAR) )
   Dim As Long nHeightOutput    = AfxGetWindowHeight( HWND_FRMOUTPUT)
   Dim As Long nWidthExplorer   = AfxGetWindowWidth( HWND_FRMEXPLORER )
   Dim As Long nHeightCombo     = AfxGetWindowHeight( hComboFiles )

   ShowWindow( GetDlgItem(HWND_FRMMAIN, IDC_FRMMAIN_REBAR), iif(gConfig.HideToolbar, SW_HIDE, SW_SHOW) )
   ShowWindow( GetDlgItem(HWND_FRMMAIN, IDC_FRMMAIN_STATUSBAR), iif(gConfig.HideStatusbar, SW_HIDE, SW_SHOW) )
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
      frmExplorer_PositionWindows
      nLeft = nWidthExplorer
   Else
      nWidthExplorer = 0
   End If


   ' If items exist in the top tabcontrol then show the tab control and account for its height
   If TabCtrl_GetItemCount(hTabCtl) = 0 Then
      ShowWindow(hTabCtl, SW_HIDE)
      nHeightTabControl = 0
      ShowWindow( hComboFiles, SW_HIDE )
      ShowWindow( hComboFunctions, SW_HIDE )
      nHeightCombo = 0
   Else
      nHeightTabControl = AfxGetWindowHeight(hTabCtl)
      SetWindowPos( hTabCtl, 0, _
                    nLeft, rc.Top + nHeightRebar, _
                    rc.Right - nWidthExplorer, nHeightTabControl, _
                    SWP_SHOWWINDOW Or SWP_NOZORDER )
      ShowWindow( hComboFiles, SW_SHOW )
      ShowWindow( hComboFunctions, SW_SHOW )
   End If
   
   
   ' If the visual designer is being shown then we will hide the files and functions combos
   If pDoc Then 
      if (pDoc->IsDesigner) andalso (IsDesignerView(pDoc)) THEN
         ShowWindow( hComboFiles, SW_HIDE )
         ShowWindow( hComboFunctions, SW_HIDE )
         nHeightCombo = 0
      else
         ' Position the Files and Functions comboboxes just after the top tab control
         dim as long nWidthCombo = (rc.Right - nWidthExplorer - pWindow->ScaleX(10)) * .5
          
         SetWindowPos( hComboFiles, 0, _
                       nleft, _
                       rc.Top + nHeightRebar + nHeightTabControl + pWindow->ScaleY(2), _
                       nWidthCombo, nHeightCombo, _
                       SWP_SHOWWINDOW or SWP_NOZORDER )

         SetWindowPos( hComboFunctions, 0, _
                       nleft + nWidthCombo + pWindow->ScaleX(10), _
                       rc.Top + nHeightRebar + nHeightTabControl + pWindow->ScaleY(2), _
                       nWidthCombo, nHeightCombo, _
                       SWP_SHOWWINDOW or SWP_NOZORDER )
      end if
   end if
   
   
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


   If pDoc Then 
      ' Position the Scintilla child edit windows
      dim as RECT rcClient
      ' Reduce the height of the client area by the size of the rebar and statusbar.
      rcClient.left = nLeft
      rcClient.top = rc.Top + nHeightRebar + nHeightTabControl + nHeightCombo 
      rcClient.right = rc.Right - nWidthExplorer
      rcClient.bottom = rc.bottom - nHeightStatusBar - nHeightOutput
      
      dim as long iScrollbarWidth  = GetSystemMetrics(SM_CXVSCROLL)
      dim as long iScrollbarHeight = GetSystemMetrics(SM_CXHSCROLL)
      dim as long iWidth = rcClient.right - iScrollbarWidth

      dim as long nDesignCodeTabHeight = 0
      if pDoc->IsDesigner then
         dim pWindow as CWindow ptr = AfxCWindowPtr(pDoc->hWndDesigner)
         nDesignCodeTabHeight = pWindow->ScaleY(24)
      end if
      
      dim as long lStyle = SWP_SHOWWINDOW or SWP_NOZORDER Or SWP_NOOWNERZORDER Or SWP_NOACTIVATE Or SWP_NOCOPYBITS
      
      if pDoc->IsDesigner THEN
         SetWindowPos( pDoc->hDesignTabCtrl, 0, _
                       rcClient.left, rcClient.bottom - nDesignCodeTabHeight, _
                       rcClient.right, rcClient.bottom - nDesignCodeTabHeight, _
                       lStyle )
      else   
         ShowWindow(pDoc->hDesignTabCtrl, SW_HIDE)
      end if
      
      if (pDoc->IsDesigner) andalso (IsDesignerView(pDoc)) THEN
         ' If the DesignView has been set to show the visual designer rather than the code
         ' window then ensure that DesignMain is shown.
         SetWindowPos( pDoc->hWndDesigner, 0, _
                       rcClient.left, rcClient.top, _
                       rcClient.right, rcClient.bottom - rcClient.top - nDesignCodeTabHeight, _
                       lStyle )

         ' Display the ToolBox/PropertyList
         frmVDToolbox_Show( HWND_FRMMAIN, SW_SHOW )
         DisplayPropertyList( pDoc )
         ShowWindow( HWND_FRMVDTOOLBOX, SW_SHOWNORMAL )

         ShowWindow( pDoc->hWindow(0), SW_HIDE )
         ShowWindow( pDoc->hWindow(1), SW_HIDE )
         ShowWindow( pDoc->hScrollbar, SW_HIDE )
      else
      
         ShowWindow( pDoc->hWndDesigner, SW_HIDE )
         ShowWindow( HWND_FRMVDTOOLBOX, SW_HIDE )

         rcClient.top = rcClient.top + pWindow->ScaleY(4)
         dim as long nSplitSize 
         
         ' Position top (optional) edit window
         if pDoc->SplitY - rcClient.top <= 0 THEN 
            ShowWindow( pDoc->hWindow(1), SW_HIDE )      
            nSplitSize = 0              
         Else
            SetWindowPos( pDoc->hWindow(1), 0, _
                          rcClient.left, rcClient.top, _
                          rcClient.right, pDoc->SplitY - rcClient.top, _ 
                          lStyle )
            nSplitSize = SPLITSIZE              
         end if
         
         dim as long nTop = Max( rcClient.top, pDoc->SplitY )

         ' Position main edit window
         SetWindowPos( pDoc->hWindow(0), 0, _
                       rcClient.left, nTop + nSplitSize, _
                       iWidth, rcClient.bottom - nTop - nSplitSize - nDesignCodeTabHeight, _
                       lStyle )

         ' Position the main view vertical scroll bar
         SetWindowPos( pDoc->hScrollbar, 0, _ 
                       rcClient.left + iWidth, _
                       nTop + SPLITSIZE, _
                       iScrollbarWidth, _
                       rcClient.bottom - SPLITSIZE - nTop  - nDesignCodeTabHeight, _ 
                       lStyle )

         ' Update vertical scroll bar page value
         pDoc->ScrInfo.nPage = SciExec( pDoc->hWindow(0), SCI_LINESONSCREEN, 0, 0 )
         pDoc->ScrInfo.nMax  = SciExec( pDoc->hWindow(0), SCI_GETLINECOUNT, 0, 0 ) - 1    
         pDoc->ScrInfo.nPos  = SciExec( pDoc->hWindow(0), SCI_GETFIRSTVISIBLELINE, 0, 0 )
         SetScrollInfo( pDoc->hScrollbar, SB_CTL, @pDoc->ScrInfo, TRUE )
         
         ' If the FiND/REPLACE window is open then ensure that it is positioned correctly
         ' especially important if the Main window is resized.
         frmFindReplace_PositionWindow
         
         frmMain_SetFocusToCurrentCodeWindow
      end if
      
      AfxRedrawWindow(HWND_FRMMAIN)

   end if
   
   
   ' Ensure that the correct notes are shown
   frmOutput_ShowNotes
   frmMain_SetStatusbar
   
   Function = 0
End Function



' ========================================================================================
' Attempt to open specified project. 
' ========================================================================================
public Function frmMain_OpenProjectSafely( ByVal HWnd As HWnd, _
                                           byref wszProjectFileName as const WString _
                                           ) as Boolean

   ' If a project is open then close any files that may be open in the editor. No need to
   ' do this if a new project is being created/opened because that function has already
   ' performed this operation. Doing so again will cause the Recent Files/Projects panel
   ' to flash on the screen.
   if gApp.IsNewProjectFlag = false then
      If gApp.IsProjectActive Then 
         if OnCommand_ProjectClose(hwnd) = false then exit function
      else   
         if OnCommand_FileClose(HWnd, EFC_CLOSEALL) = false then exit function
      end if
      ' Clear any previous info from the Output windows
      frmOutput_ResetAllControls()
   end if
         
   ' Open the project
   If gConfig.ProjectLoadFromFile(wszProjectFileName) Then
      ' Update the most recently used project list 
      UpdateMRUProjectList(wszProjectFileName)
   End If
   gApp.IsProjectActive = true
   
   ' Position all of the controls into place
   frmExplorer_PositionWindows
   frmMain_PositionWindows

   function = true
end function


' ========================================================================================
' Attempt to open specified file. If it exists then position to Tab if applicable
' ========================================================================================
public Function frmMain_OpenFileSafely( ByVal HWnd        As HWnd, _
                                        ByVal bIsNewFile  As BOOLEAN, _ 
                                        ByVal bIsTemplate As BOOLEAN, _
                                        ByVal bShowInTab  As BOOLEAN, _
                                        byval bIsInclude  as BOOLEAN, _
                                        ByVal pwszName    As WString Ptr, _
                                        ByVal pDocIn      As clsDocument Ptr, _
                                        byval bIsDesigner as Boolean = false, _
                                        byval nFileType as long = FILETYPE_UNDEFINED _
                                        ) As clsDocument Ptr
   
   Dim As Long iTab = -1
   Dim pDoc As clsDocument Ptr 
   Dim As Hwnd hTree = GetDlgItem( HWND_FRMEXPLORER, IDC_FRMEXPLORER_TREE )
   
   If pDocIn = 0 Then   
      ' Create a new pDoc 
      gApp.FileLoadingCount = gApp.FileLoadingCount + 1

      pDoc = New clsDocument

      gApp.AddDocument( pDoc ) 
      pDoc->IsDesigner = bIsDesigner 

      ' All pDoc's will have an associated Scintilla window for code.
      If bIsNewFile Then
         pDoc->CreateCodeWindow( HWnd, True )  ' Create the new Scintilla window
      Else
         pDoc->CreateCodeWindow( HWnd, false, bIsTemplate, pwszName )
      End If   

      pDoc->bNeedsParsing = true
      ParseDocument( pDoc )

      ' When a Project is being loaded, the incoming nFileType is from the project
      ' file. We pass it to the AddFunctionsToExplorerTreeview function in order to
      ' ensure that the pDoc is attached to the right node.
      pDoc->ProjectFileType = nFileType
   Else
      ' deal with the existing pDoc
      pDoc = pDocIn     
   End If   
   pwszName = @pDoc->DiskFilename

   
   ' Add the document to the Explorer treeview and ensure that the root and special
   ' nodes are created.
   AddFunctionsToExplorerTreeview( pDoc, False )


   ' Set the default build configuration for this document. If no IsDefault option
   ' has been checked then we simply use whatever the current selected build is.
   if pDoc->DocumentBuild = "" THEN 
      pDoc->DocumentBuild = frmBuildConfig_GetDefaultBuildGUID()
      if pDoc->DocumentBuild = "" THEN 
         pDoc->DocumentBuild = frmBuildConfig_GetSelectedBuildGUID()
      end if
   end if      
   
   ' Only sort the root if no project is active because we do not want the 
   ' ordering of our special subnodes to be reorganized.
   if ( gApp.IsProjectActive = false ) andalso (gApp.IsProjectLoading = false ) then
      Treeview_SortChildren( hTree, gApp.hExplorerRootNode, 0 )
   end if   
   TreeView_Expand( hTree, gApp.hExplorerRootNode, TVE_EXPAND)

   If bShowInTab Then
      ' If the document is already open and loaded then simply switch to
      ' that document in the top tabcontrol.
      iTab = gTTabCtl.GetTabIndexFromFilename(pwszName)
      If iTab = -1 Then iTab = gTTabCtl.AddTab( pDoc )  ' Add the new document to the top tabcontrol
      gTTabCtl.SetFocusTab(iTab)
      Treeview_SelectItem( GetDlgItem(HWND_FRMEXPLORER, IDC_FRMEXPLORER_TREE), _
                           pDoc->hNodeExplorer )
   End If

   ' If a Menu, ToolBar or StatusBar exists on the Form then ensure that it
   ' resizes to the new Form width.
   frmMenuEditor_CreateFakeMainMenu(pDoc)
   frmToolBarEditor_CreateFakeToolBar(pDoc)
   frmStatusBarEditor_CreateFakeStatusBar(pDoc)


   ' Position all of the controls into place
   frmMain_PositionWindows
   frmMain_SetFocusToCurrentCodeWindow

   Function = pDoc
End Function


' ========================================================================================
' Process WM_CREATE message for window/dialog: frmMain
' ========================================================================================
private Function frmMain_OnCreate( ByVal HWnd As HWnd, _
                                   ByVal lpCreateStructPtr As LPCREATESTRUCT _
                                   ) As BOOLEAN
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
private Function frmMain_OnSize( ByVal HWnd As HWnd, _
                                 ByVal state As UINT, _
                                 ByVal cx As Long, _
                                 ByVal cy As Long _
                                 ) As LRESULT
   If state <> SIZE_MINIMIZED Then
      ' Update the size of the Rebar control
      SendMessage GetDlgItem(HWnd, IDC_FRMMAIN_REBAR), WM_SIZE, cx, cy
      
      ' Resize the status bar
      Dim As HWnd hStatusBar = GetDlgItem(HWnd, IDC_FRMMAIN_STATUSBAR)
      SendMessage hStatusBar, WM_SIZE, cx, cy
      InvalidateRect hStatusBar, Null, True
      
      ' Position all of the child windows
      frmMain_PositionWindows
   End If
   Function = 0
End Function


' ========================================================================================
' Process WM_COMMAND message for window/dialog: frmMain
' ========================================================================================
private Function frmMain_OnCommand( ByVal HWnd As HWnd, _
                                    ByVal id As Long, _
                                    ByVal hwndCtl As HWnd, _
                                    ByVal codeNotify As UINT _
                                    ) As LRESULT
   Dim As HWnd hEdit 

   Dim pDoc As clsDocument Ptr = gTTabCtl.GetActiveDocumentPtr()
   If pDoc Then 
      select case codeNotify
         case SCEN_SETFOCUS
            pDoc->hWndActiveScintilla = hwndCtl
            frmMain_SelectComboFiles
            frmMain_SelectComboFunctions
         case SCEN_KILLFOCUS
            pDoc->hWndActiveScintilla = 0
      end select      
      hEdit = pDoc->hWndActiveScintilla
      frmMain_ChangeToolbarButtonsState
   End If   

   ' If the popup Help Viewer window has focus then exit function because we don't
   ' want accelerator keys to act on the code editor.
   if IsWindow( HWND_FRMHELPVIEWER ) then
      if GetForegroundWindow() = HWND_FRMHELPVIEWER then 
         if id = IDM_FIND then
            frmHelpViewer_OnCommand(HWND_FRMHELPVIEWER, IDC_FRMHELPVIEWER_FIND, 0, BN_CLICKED)
         end if   
         exit function
      end if
   end if
   
   
   Select Case id
      
      case IDC_FRMMAIN_COMBOFILES
         if codeNotify = CBN_CLOSEUP THEN
            frmMain_SetFocusToCurrentCodeWindow
         
         elseif codeNotify = CBN_SELCHANGE THEN
            ' Load the pDoc that is held in the itemData 
            dim as long nCurSel = ComboBox_GetCurSel( hwndCtl )
            if nCurSel = -1 then exit function
            dim pDocSelected as clsDocument ptr 
            pDocSelected = cast( clsDocument ptr, ComboBox_GetItemData(hwndCtl, nCurSel) ) 
            if pDocSelected <> pDoc then
               OpenSelectedDocument( pDocSelected->DiskFilename )
            end if   
            frmMain_SetFocusToCurrentCodeWindow
         end if

      
      case IDC_FRMMAIN_COMBOFUNCTIONS
         if codeNotify = CBN_CLOSEUP THEN
            frmMain_SetFocusToCurrentCodeWindow
         
         elseif codeNotify = CBN_SELCHANGE THEN
            ' Load the pDoc that is held in the itemData of the Files combobox
            dim pDocSelected as clsDocument ptr 
            dim wszFilename as wstring * MAX_PATH
            dim wszFunctionName as wstring * MAX_PATH
            dim as HWND hComboFiles = GetDlgItem( HWND_FRMMAIN, IDC_FRMMAIN_COMBOFILES )
            dim as long nCurSel = ComboBox_GetCurSel( hComboFiles )
            if nCurSel = -1 then exit function
            
            ' Get the File name
            pDocSelected = cast( clsDocument ptr, ComboBox_GetItemData(hComboFiles, nCurSel) ) 
            wszFilename = pDocSelected->DiskFilename
            ' Get the Function name
            ' If the pDoc is dirty then we will need to reparse to the document otherwise
            ' the line numbers will not be correct in the DB2 database.
            ' If we have selected "(Module level)" then simply reposition to start of file.
            nCurSel = ComboBox_GetCurSel( hwndCtl )
            if nCurSel = 0 then
               OpenSelectedDocument( wszFilename, "", 0 )
            else   
               If cbool(SciExec( pDoc->hWindow(0), SCI_GETMODIFY, 0, 0 )) or pDoc->UserModified Then
                  ParseDocument( pDoc )
               end if
               ComboBox_GetText( hwndCtl, wszFunctionName, MAX_PATH )
               OpenSelectedDocument( wszFilename, wszFunctionName )
            end if
            frmMain_SetFocusToCurrentCodeWindow
         end if
         
         
      case IDC_FRMMAIN_COMBOBUILDS
         if codeNotify = CBN_CLOSEUP THEN
            frmMain_SetFocusToCurrentCodeWindow
         elseif codeNotify = CBN_SELCHANGE THEN
            dim as long nCurSel = ComboBox_GetCurSel(hwndCtl)
            if nCurSel = -1 THEN exit function
            If gApp.IsProjectActive Then
               gApp.ProjectBuild = gConfig.Builds(nCurSel).Id
            Else
               Dim pDocMain As clsDocument Ptr 
               pDocMain = gTTabCtl.GetActiveDocumentPtr()
               If pDocMain Then pDocMain->DocumentBuild = gConfig.Builds(nCurSel).Id
            end if
            frmMain_SetFocusToCurrentCodeWindow()
         END IF  
      
      
      Case IDM_GOTONEXTTAB
         gTTabCtl.NextTab
         
      Case IDM_GOTOPREVTAB
         gTTabCtl.PrevTab
         
      Case IDM_CLOSETAB
         gTTabCtl.CloseTab

      ''  FILE MENU
      Case IDM_FILENEW
         OnCommand_FileNew( HWnd )
      
      Case IDM_FILEOPEN
         OnCommand_FileOpen( HWnd, true )
      
      Case IDM_FILEOPENTEMPLATES
         dim as RECT rc
         SendMessage(GetDlgItem(HWnd, IDC_FRMMAIN_REBAR), RB_GETRECT, 0, Cast(LPARAM, @rc))
         MapWindowPoints(hwnd, HWND_DESKTOP, Cast(Point Ptr, @rc), 2)
         frmTemplates_Show(HWnd, rc.Left, rc.Bottom)

      Case IDM_FILESAVE
         OnCommand_FileSave( HWnd, pDoc, False )
      
      Case IDM_FILESAVEAS
         OnCommand_FileSave( HWnd, pDoc, True )

      case IDM_FILEOPEN_EXPLORERTREEVIEW, _
           IDM_FILESAVE_EXPLORERTREEVIEW, IDM_FILESAVEAS_EXPLORERTREEVIEW, _
           IDM_FILECLOSE_EXPLORERTREEVIEW
         dim as hwnd hTV = GetDlgItem( HWND_FRMEXPLORER, IDC_FRMEXPLORER_TREE )
         dim as HTREEITEM hItem = TreeView_GetSelection( hTV )
         dim as long ItemLParam = FF_TreeView_GetlParam( hTV, hItem )
         ' Must cast the ItemLParam Long to Integer or FBC will fail 64 bit compile.
         pDoc = Cast(clsDocument ptr, cast(integer, ItemLParam) )
         if id = IDM_FILEOPEN_EXPLORERTREEVIEW then
            dim as CWSTR wszFilename, wszFunctionName
            frmExplorer_GetFileNameFunctionName( hItem, wszFilename, wszFunctionName )
            OpenSelectedDocument( wszFilename, wszFunctionName )
            TreeView_SelectItem( hTV, hItem )
         elseif id = IDM_FILESAVE_EXPLORERTREEVIEW then
            OnCommand_FileSave( HWnd, pDoc, False )
         elseif id = IDM_FILESAVEAS_EXPLORERTREEVIEW then
            OnCommand_FileSave( HWnd, pDoc, true )
         elseif IDM_FILECLOSE_EXPLORERTREEVIEW then
            dim as long nCurSel = gTTabCtl.GetTabIndexByDocumentPtr( pDoc )
            if nCurSel <> -1 then
               OnCommand_FileClose( HWnd, EFC_CLOSECURRENT, nCurSel )
            end if   
         end if   
         
      Case IDM_FILECLOSE
         OnCommand_FileClose( HWnd, EFC_CLOSECURRENT )
      
      Case IDM_FILESAVEALL:        OnCommand_FileSaveAll( HWnd ) 
      
      Case IDM_FILECLOSEALL:       OnCommand_FileClose( HWnd, EFC_CLOSEALL )
      Case IDM_FILECLOSEALLOTHERS: OnCommand_FileClose( HWnd, EFC_CLOSEALLOTHERS )
      Case IDM_CLOSEALLFORWARD:    OnCommand_FileClose( HWnd, EFC_CLOSEALLFORWARD )
      Case IDM_CLOSEALLBACKWARD:   OnCommand_FileClose( HWnd, EFC_CLOSEALLBACKWARD )
      
      Case IDC_MRUBASE To (IDC_MRUBASE + 10)
         OpenMRUFile( HWnd, id )
      
      Case IDM_MRUCLEAR, IDM_MRUPROJECTCLEAR
         ClearMRUlist( id )
      
      Case IDM_OPENINCLUDE
         OpenSelectedDocument( gApp.IncludeFilename )

      Case IDM_COMMAND
         Shell(Environ("COMSPEC"))
         
      Case IDM_EXIT
         SendMessage( HWnd, WM_CLOSE, 0, 0 )

      ''  EDIT MENU
      Case IDM_REDO:               SciExec( hEdit, SCI_REDO, 0, 0)
      Case IDM_UNDO, WM_UNDO:      SciExec( GetFocus(), Iif(GetFocus()=hEdit, SCI_UNDO, WM_UNDO), 0, 0)
      
      Case IDM_CUT, WM_CUT
         ' Check if the Compiler log file or Notes windows or Snippets have
         ' input focus. If they do then select that text rather than Scintilla.
         select case GetFocus()
            case GetDlgItem(HWND_FRMOUTPUT, IDC_FRMOUTPUT_TXTNOTES), _
                 GetDlgItem(HWND_FRMOUTPUT, IDC_FRMOUTPUT_TXTLOGFILE), _ 
                 GetDlgItem(HWND_FRMSNIPPETS, IDC_FRMSNIPPETS_TXTCODE)
               SendMessage( GetFocus, WM_CUT, 0, 0 )
            case else     
               if IsDesignerView(pDoc) then
                  SendMessage(HWND_FRMMAIN, WM_COMMAND, MAKELONG(IDM_COPY, 0), 0)
                  SendMessage(HWND_FRMMAIN, WM_COMMAND, MAKELONG(IDM_DELETE, 0), 0)
               else
                  SciExec( GetFocus(), Iif(GetFocus()=hEdit, SCI_CUT, WM_CUT), 0, 0)
               end if   
         end select
      
      Case IDM_COPY, WM_COPY
         ' Check if the Compiler log file or Notes windows or Snippets have
         ' input focus. If they do then select that text rather than Scintilla.
         select case GetFocus()
            case GetDlgItem(HWND_FRMOUTPUT, IDC_FRMOUTPUT_TXTNOTES), _
                 GetDlgItem(HWND_FRMOUTPUT, IDC_FRMOUTPUT_TXTLOGFILE), _ 
                 GetDlgItem(HWND_FRMSNIPPETS, IDC_FRMSNIPPETS_TXTCODE)
               SendMessage( GetFocus, WM_COPY, 0, 0 )
            case else     
               if IsDesignerView(pDoc) then
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
                  SciExec( GetFocus(), Iif(GetFocus()=hEdit, SCI_COPY, WM_COPY), 0, 0)
               end if
         end select

            
      Case IDM_PASTE, WM_PASTE
         ' Check if the Compiler log file or Notes windows or Snippets have
         ' input focus. If they do then select that text rather than Scintilla.
         select case GetFocus()
            case GetDlgItem(HWND_FRMOUTPUT, IDC_FRMOUTPUT_TXTNOTES), _
                 GetDlgItem(HWND_FRMOUTPUT, IDC_FRMOUTPUT_TXTLOGFILE), _ 
                 GetDlgItem(HWND_FRMSNIPPETS, IDC_FRMSNIPPETS_TXTCODE)
               SendMessage( GetFocus, WM_PASTE, 0, 0 )
            case else     
               if IsDesignerView(pDoc) then
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
                           case "GROUPNAME"
                              ' Copy the GroupName and then increment the Index property
                              pCtrl->Properties(ii) = gCopyControls(i).Properties(ii)
                              continue for
                           CASE "TABINDEX"
                              continue for    
                        END select
                        pCtrl->Properties(ii) = gCopyControls(i).Properties(ii)
                     NEXT
                     
                     ' Set the left/top properties relative to the current insert position
                     dim pt as point
                     GetCursorPos(@pt)
                     MapWindowPoints(0, pDoc->hWndForm, cast(point ptr, @pt), 1)
                     Dim pWindow As CWindow Ptr = AfxCWindowPtr(HWND)
                     if pWindow then
                        pt.x = pWindow->UnScaleX(pt.x)
                        pt.y = pWindow->UnScaleY(pt.y)
                     end if   
                     SetControlProperty(pCtrl, "LEFT", str(pt.x + (nLeft-nLeftActive)))
                     SetControlProperty(pCtrl, "TOP", str(pt.y + (nTop-nTopActive)))
                     pCtrl->IsActive = false
                     pCtrl->IsSelected = false
                     
                     pCtrl->SuspendLayout = true
                     ApplyControlProperties( pDoc, pCtrl )
                     pCtrl->SuspendLayout = false
                  NEXT
                  ' Ensure the grab handles of form and controls are redrawn or hidden
                  pDoc->Controls.SetActiveControl(pDoc->hWndForm)
                  pDoc->bRegenerateCode = true
                  pDoc->UserModified = true
                  AfxRedrawWindow(pDoc->hWndFrame)
                  AfxRedrawWindow(pDoc->hWndForm)
                  frmMain_SetStatusbar
                  DisplayPropertyList(pDoc)
               else
                  SciExec( GetFocus(), Iif(GetFocus()=hEdit, SCI_PASTE, WM_PASTE), 0, 0)
               end if   
         end select
         
      case IDM_DELETE
         dim pCtrl as clsControl ptr
         ' Need to test if any of the selected controls are locked. If any of them are then
         ' we need to abort the deleting process.
         for i as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
            pCtrl = pDoc->Controls.ItemAt(i)
            if pCtrl->IsSelected THEN 
               ' Test to ensure that the specific control Locked property is not set, or
               ' the global Locked setting for the entire form is not set.
               if IsControlLocked(pDoc, pCtrl) then 
                  ' Maybe put a messagebox here indicating that can not delete because locked?
                  exit function
               end if
            end if         
         next

         ' Delete selected controls
         do until pDoc->Controls.SelectedControlsCount = 0
            for i as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
               pCtrl = pDoc->Controls.ItemAt(i)
               if pCtrl->IsSelected THEN 
                  ' If this was a BUTTON control then we need to potentially update the form ACCEPTBUTTON 
                  ' or CANCELBUTTON properties that depend on this button name.
                  if pCtrl->ControlType = CTRL_BUTTON then
                     dim as CWSTR wszCtrlName = GetControlProperty(pCtrl, "NAME")
                     dim pCtrlForm as clsControl ptr = GetFormCtrlPtr(pDoc)
                     if pCtrlForm then
                        If GetControlProperty(pCtrlForm, "ACCEPTBUTTON") = wszCtrlName then
                           SetControlProperty(pCtrlForm, "ACCEPTBUTTON", "")
                        end if    
                        If GetControlProperty(pCtrlForm, "CANCELBUTTON") = wszCtrlName then
                           SetControlProperty(pCtrlForm, "CANCELBUTTON", "")
                        end if    
                     end if
                  end if
                  pDoc->Controls.Remove(pCtrl)
                  pDoc->bRegenerateCode = true
                  pDoc->UserModified = true
                  exit for
               END IF
            NEXT
         loop
         ' Ensure the grab handles of form and controls are redrawn or hidden
         pDoc->Controls.SetActiveControl(pDoc->hWndForm)
         AfxRedrawWindow(pDoc->hWndFrame)
         AfxRedrawWindow(pDoc->hWndForm)
         frmMain_SetStatusbar
         DisplayPropertyList(pDoc)
        
      Case IDM_DELETELINE:  SciExec(hEdit, SCI_LINECUT, 0, 0)
      Case IDM_INSERTFILE:  pDoc->InsertFile()
      
      Case IDM_INDENTBLOCK
         ' If a TAB was pressed then manually terminate any active autocomplete.
         if SciExec(hEdit, SCI_AUTOCACTIVE, 0, 0) then
            SciExec(hEdit, SCI_AUTOCCOMPLETE, 0, 0) 
         end if
         if GetFocus = GetDlgItem(HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_TXTFIND) then
            if gFind.bShowReplace then SetFocus GetDlgItem(HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_TXTREPLACE)
         elseif GetFocus = GetDlgItem(HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_TXTREPLACE) then
            if gFind.bShowReplace then SetFocus GetDlgItem(HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_TXTFIND)
         else
            ' Determine if a Snippet can be inserted. This takes precedence over doing a block indent.
            if frmSnippets_DoInsertSnippet( pDoc ) = false then
               ' No snippet to insert, do a block indent instead.
               gApp.SuppressNotify = true
               Function = SciExec(hEdit, SCI_TAB, 0, 0)
               gApp.SuppressNotify = false
            end if
            frmMain_SetStatusbar
         end if
      
      Case IDM_UNINDENTBLOCK
         if GetFocus = GetDlgItem(HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_TXTFIND) then
            if gFind.bShowReplace then SetFocus GetDlgItem(HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_TXTREPLACE)
         elseif GetFocus = GetDlgItem(HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_TXTREPLACE) then
            if gFind.bShowReplace then SetFocus GetDlgItem(HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_TXTFIND)
         else
            gApp.SuppressNotify = true
            Function = SciExec(hEdit, SCI_BACKTAB, 0, 0)
            gApp.SuppressNotify = false
            frmMain_SetStatusbar
         end if
         
      Case IDM_COMMENTBLOCK:       pDoc->BlockComment(True) 
      Case IDM_UNCOMMENTBLOCK:     pDoc->BlockComment(False) 
      Case IDM_DUPLICATELINE:      pDoc->LineDuplicate()
      Case IDM_MOVELINEUP:         pDoc->MoveCurrentLines(False) 
      Case IDM_MOVELINEDOWN:       pDoc->MoveCurrentLines(True)  
      Case IDM_TOUPPERCASE:        pDoc->ChangeSelectionCase(1)
      Case IDM_TOLOWERCASE:        pDoc->ChangeSelectionCase(2)
      Case IDM_TOMIXEDCASE:        pDoc->ChangeSelectionCase(3)
      Case IDM_EOLTOCRLF:          pDoc->ConvertEOL(SC_EOL_CRLF)
      Case IDM_EOLTOCR:            pDoc->ConvertEOL(SC_EOL_CR)
      Case IDM_EOLTOLF:            pDoc->ConvertEOL(SC_EOL_LF)
      Case IDM_TABSTOSPACES:       pDoc->TabsToSpaces()
      Case IDM_SELECTLINE:         pDoc->SelectLine(-1)

      Case IDM_SELECTALL
         ' Check if the Compiler log file or Notes windows or Snippets have
         ' input focus. If they do then select that text rather than Scintilla.
         select case GetFocus()
            case GetDlgItem(HWND_FRMOUTPUT, IDC_FRMOUTPUT_TXTNOTES), _
                 GetDlgItem(HWND_FRMOUTPUT, IDC_FRMOUTPUT_TXTLOGFILE), _ 
                 GetDlgItem(HWND_FRMSNIPPETS, IDC_FRMSNIPPETS_TXTCODE)
               Edit_SetSel( GetFocus, 0, -1 )
            case else     
               if IsDesignerView(pDoc) THEN
                  pDoc->Controls.SelectAllControls
                  ' Ensure the grab handles of form and controls are redrawn or hidden
                  AfxRedrawWindow(pDoc->hWndFrame)
                  AfxRedrawWindow(pDoc->hWndForm)
               else
                  SciExec(hEdit, SCI_SELECTALL, 0, 0)
               END IF
         end select
                                   
      case IDM_ANSI
         if pDoc->FileEncoding <> FILE_ENCODING_ANSI THEN 
            pDoc->UserModified = true
            ' Convert buffer to ANSI and redisplay the text in the editor
            ConvertTextBuffer(pDoc, FILE_ENCODING_ANSI)
         end if   
         pDoc->FileEncoding = FILE_ENCODING_ANSI 
         frmMain_SetStatusbar
      
      case IDM_UTF8BOM
         if pDoc->FileEncoding <> FILE_ENCODING_UTF8_BOM THEN 
            pDoc->UserModified = true
            ConvertTextBuffer(pDoc, FILE_ENCODING_UTF8_BOM)
         end if   
         pDoc->FileEncoding = FILE_ENCODING_UTF8_BOM 
         frmMain_SetStatusbar
      
      case IDM_UTF16BOM
         if pDoc->FileEncoding <> FILE_ENCODING_UTF16_BOM THEN 
            pDoc->UserModified = true
            ConvertTextBuffer(pDoc, FILE_ENCODING_UTF16_BOM)
         end if   
         frmMain_SetStatusbar


      ''  SEARCH MENU              
      Case IDM_FIND, IDM_REPLACE
         frmFindReplace_Show( HWND_FRMMAIN, iif(id = IDM_FIND, false, true) )
      
      Case IDM_FINDINFILES
         frmFindInFiles_Show( hEdit )

      Case IDM_FINDNEXT, IDM_FINDPREV
         Dim As String sFindText = pDoc->GetSelText()
         if (GetForeGroundWindow = HWND_FRMFINDREPLACE) or (Len(sFindText) = 0) then
            id = iif(id = IDM_FINDNEXT, IDC_FRMFINDREPLACE_BTNRIGHT, IDC_FRMFINDREPLACE_BTNLEFT)
            SendMessage( HWND_FRMFINDREPLACE, WM_COMMAND, MAKELONG(id, 0), Cast(LPARAM, GetDlgItem(HWND_FRMFINDREPLACE, id)) )
         end if
         
      Case IDM_FINDNEXTACCEL, IDM_FINDPREVACCEL   ' F3/Shift+F3
         if pDoc = 0 then exit function

         dim as long startPos, endPos, curPos, mainSel
         Dim As String sFindText = pDoc->GetSelText()
         
         ' Use a static string variable to keep track of the F3/Shift F3 information
         ' We use the static value if there is no currently selected word/phrase.
         static as string sFindTextStatic
         if len(sFindText) = 0 then sFindText = sFindTextStatic
         sFindTextStatic = sFindText
         
         if len(sFindText) then
            curPos = SciExec( hEdit, SCI_GETCURRENTPOS, 0, 0) 
            startPos = iif( id = IDM_FINDNEXTACCEL, curPos + 1, curPos - 1 )
            endPos = iif( id = IDM_FINDNEXTACCEL, SciExec( hEdit, SCI_GETTEXTLENGTH, 0, 0), 0)
            SciExec( hEdit, SCI_SETTARGETSTART, startPos, 0)
            SciExec( hEdit, SCI_SETTARGETEND, endPos, 0)
            ' Search for the text to find
            dim as long r = SciExec( hEdit, SCI_SEARCHINTARGET, Len(sFindText), Strptr(sFindText))
            If r <> -1 Then SciExec( hEdit, SCI_SETSEL, r, r + len(sFindText)) 
         end if

      Case IDM_GOTO
         frmGoto_Show( hEdit )

      Case IDM_DEFINITION
         frmMain_GotoDefinition( pDoc )
         
      Case IDM_LASTPOSITION
         frmMain_GotoLastPosition()
      
      case IDM_GOTONEXTFUNCTION
         pDoc->GotoNextFunction
         
      case IDM_GOTOPREVFUNCTION
         pDoc->GotoPrevFunction

      case IDM_GOTONEXTTAB
         gTTabCtl.NextTab
      
      case IDM_GOTOPREVTAB
         gTTabCtl.PrevTab
      
      case IDM_GOTOHEADERFILE, _
           IDM_GOTOSOURCEFILE, _ 
           IDM_GOTOMAINFILE, _ 
           IDM_GOTORESOURCEFILE 
         frmMain_GotoFile( pDoc, id )
         
      Case IDM_BOOKMARKTOGGLE
         pDoc->ToggleBookmark( pDoc->GetCurrentLineNumber )
      
      Case IDM_BOOKMARKNEXT
         pDoc->NextBookmark
      
      Case IDM_BOOKMARKPREV
         pDoc->PrevBookmark
      
      Case IDM_BOOKMARKCLEARALL
         SciExec(hEdit, SCI_MARKERDELETEALL, -1, 0)
      
      
      ''  VIEW MENU
      Case IDM_FOLDTOGGLE:         pDoc->FoldToggle( pDoc->GetCurrentLineNumber() )
      Case IDM_FOLDBELOW:          pDoc->FoldToggleOnwards( pDoc->GetCurrentLineNumber() )
      Case IDM_FOLDALL:            pDoc->FoldAll()
      Case IDM_UNFOLDALL:          pDoc->UnFoldAll()
      Case IDM_ZOOMIN:             SciExec(hEdit, SCI_ZOOMIN, 0, 0)
      Case IDM_ZOOMOUT:            SciExec(hEdit, SCI_ZOOMOUT, 0, 0)

      Case IDM_FUNCTIONLIST
         ' this will also toggle visibility 
         frmFunctionList_Show( HWnd )

      Case IDM_VIEWEXPLORER
         ShowWindow HWND_FRMEXPLORER, Iif(IsWindowVisible(HWND_FRMEXPLORER), SW_HIDE, SW_SHOW)
         frmMain_PositionWindows

      Case IDM_VIEWOUTPUT
         ShowWindow HWND_FRMOUTPUT, Iif(IsWindowVisible(HWND_FRMOUTPUT), SW_HIDE, SW_SHOW)
         frmMain_PositionWindows

      Case IDM_RESTOREMAIN
         Dim rc As Rect
         SystemParametersInfo SPI_GETWORKAREA, 0, @rc, 0
         MoveWindow HWnd, 0, 0, rc.Right - rc.Left, rc.Bottom - rc.Top, CTRUE
               
      
      ''  PROJECT MENU
      Case IDM_PROJECTNEW:            OnCommand_ProjectNew( HWnd )
      Case IDM_PROJECTSAVE:           OnCommand_ProjectSave( HWnd, False )
      Case IDM_PROJECTSAVEAS:         OnCommand_ProjectSave( HWnd, True )
      Case IDM_PROJECTOPEN:           OnCommand_ProjectOpen( HWnd )
      Case IDM_PROJECTFILESADD:       OnCommand_FileOpen( HWnd, false )
      Case IDM_PROJECTCLOSE:          OnCommand_ProjectClose( HWnd )
      Case IDM_PROJECTOPTIONS:        frmProjectOptions_Show( HWnd, false ) 

      
      
      Case IDM_SETFILENORMAL, IDM_SETFILEMODULE, IDM_SETFILEMAIN, _
           IDM_SETFILERESOURCE, IDM_SETFILEHEADER, _
           IDM_SETFILEMAIN_EXPLORERTREEVIEW, _
           IDM_SETFILERESOURCE_EXPLORERTREEVIEW, _
           IDM_SETFILEHEADER_EXPLORERTREEVIEW, _
           IDM_SETFILEMODULE_EXPLORERTREEVIEW, _
           IDM_SETFILENORMAL_EXPLORERTREEVIEW 
         
         dim as hwnd hTV = GetDlgItem(HWND_FRMEXPLORER, IDC_FRMEXPLORER_TREE)
         dim as long nFileType
         
         ' This code also handles right click messages from the Explorer Treeview. In
         ' that case the pDoc needs to be taken from the Treeview rather than the 
         ' currently active document in the editor.
         select case id
            case IDM_SETFILENORMAL_EXPLORERTREEVIEW, IDM_SETFILEMODULE_EXPLORERTREEVIEW, _
                 IDM_SETFILEMAIN_EXPLORERTREEVIEW, IDM_SETFILERESOURCE_EXPLORERTREEVIEW, _
                 IDM_SETFILEHEADER_EXPLORERTREEVIEW
               dim as HTREEITEM hItem = TreeView_GetSelection( hTV )
               dim as long ItemLParam = FF_TreeView_GetlParam( hTV, hItem )
               ' Must cast the ItemLParam Long to Integer or FBC will fail 64 bit compile.
               pDoc = Cast(clsDocument ptr, cast(integer, ItemLParam) )
         end select

         select case id
            case IDM_SETFILENORMAL, IDM_SETFILENORMAL_EXPLORERTREEVIEW 
               nFileType = FILETYPE_NORMAL
            case IDM_SETFILEMODULE, IDM_SETFILEMODULE_EXPLORERTREEVIEW
               nFileType = FILETYPE_MODULE
            case IDM_SETFILEMAIN, IDM_SETFILEMAIN_EXPLORERTREEVIEW
               nFileType = FILETYPE_MAIN
            case IDM_SETFILERESOURCE, IDM_SETFILERESOURCE_EXPLORERTREEVIEW
               nFileType = FILETYPE_RESOURCE
            case IDM_SETFILEHEADER, IDM_SETFILEHEADER_EXPLORERTREEVIEW
               nFileType = FILETYPE_HEADER
         end select

         ' This code is also used for treeview drag/drop operation. For a drag/drop
         ' the pDoc is not necessarily the current active document. Therefore, if
         ' dragging is active then we need to get the pDoc from the drag node lParam.
         if treedragger.dragging then pDoc = treedragger.pDoc
         if pDoc = 0 then exit function
         
         if pDoc->ProjectFileType <> nFileType then
            ' If attempting to set MAIN or RESOURCE then we need to move any existing
            ' pDoc that have that filetype over to the NORMAL branch.
            if (nFileType = FILETYPE_MAIN) or (nFileType = FILETYPE_RESOURCE) then
               dim pDoc2 as clsDocument ptr = iif(nFileType = FILETYPE_MAIN, gApp.GetMainDocumentPtr, gApp.GetResourceDocumentPtr)
               if pDoc2 then
                  Treeview_DeleteItem( hTV, pDoc2->hNodeExplorer )
                  pDoc2->hNodeExplorer = 0
                  pDoc2->ProjectFileType = FILETYPE_NORMAL
                  AddFunctionsToExplorerTreeview( pDoc2, false )
               end if   
            end if
            ' Set the new FileType for the currently selected document
            gApp.ProjectSetFileType( pDoc, nFileType )
            Treeview_DeleteItem( hTV, pDoc->hNodeExplorer )
            pDoc->hNodeExplorer = 0
            AddFunctionsToExplorerTreeview( pDoc, false )
            Treeview_SelectItem( hTV, pDoc->hNodeExplorer )
            frmMain_SetStatusbar
         end if
   

      Case IDC_MRUPROJECTBASE To (IDC_MRUPROJECTBASE + 10)
         OpenMRUProjectFile( HWnd, id )

      Case IDM_REMOVEFILEFROMPROJECT, IDM_REMOVEFILEFROMPROJECT_EXPLORERTREEVIEW
         ' These messages will already have the correct node selected in the Explorer
         ' treeview. If the document is open in a tab then we do need to set that
         ' tab as current and issue a EFC_CLOSECURRENT. 
         dim as hwnd hTV = GetDlgItem( HWND_FRMEXPLORER, IDC_FRMEXPLORER_TREE )
         dim as HTREEITEM hItem = TreeView_GetSelection( hTV )
         dim as long ItemLParam = FF_TreeView_GetlParam( hTV, hItem )
         ' Must cast the ItemLParam Long to Integer or FBC will fail 64 bit compile.
         pDoc = Cast(clsDocument ptr, cast(integer, ItemLParam) )
         if pDoc = 0 then exit function
         if gTTabCtl.SetTabIndexByDocumentPtr( pDoc ) <> -1 then
           if OnCommand_FileClose( HWnd, EFC_CLOSECURRENT ) = false then exit function
         end if  
         ' File close succeeded. Remove this document from the project collection and remove treeview node. 
         gdb2.dbDelete( pDoc->DiskFilename, 0 )
         Treeview_DeleteItem( hTV, hItem)
         frmMain_RemoveFromComboFiles( pDoc )
         gApp.RemoveDocument( pDoc )
         
           
      ''  COMPILE MENU          
      Case IDM_BUILDEXECUTE, IDM_COMPILE, IDM_REBUILDALL, IDM_QUICKRUN
         code_Compile( id )
      
      Case IDM_RUNEXE
         RunExe( GetRunExecutableFilename, gApp.ProjectCommandLine )
      
      Case IDM_COMMANDLINE
         frmCommandLine_Show( HWnd )


      ''  OPTIONS MENU
      Case IDM_OPTIONS
         frmOptions_Show( HWnd )
         ' Repaint the File/Functions and Build comboboxes because the Theme may have changed.
         AfxRedrawWindow( GetDlgItem( HWND_FRMMAIN, IDC_FRMMAIN_COMBOFILES ) )
         AfxRedrawWindow( GetDlgItem( HWND_FRMMAIN, IDC_FRMMAIN_COMBOFUNCTIONS ) )
         AfxRedrawWindow( GetDlgItem( HWND_FRMMAIN, IDC_FRMMAIN_COMBOBUILDS ) )
         
      case IDM_BUILDCONFIG
         frmBuildConfig_Show( HWND ) 
      
      case IDM_USERTOOLSDIALOG
         frmUserTools_Show( HWND ) 
         
      case IDM_USERSNIPPETS
         frmSnippets_Show( HWND ) 
      
      
      ''   DESIGNER MENU
      case IDM_NEWFORM
         OnCommand_DesignerNewForm(HWND)

      case IDM_VIEWTOOLBOX
         frmVDToolbox_Show(HWND)   ' toggle show/hide

      case IDM_TOGGLEVIEWCODE
         if pDoc then
            dim as long idx = TabCtrl_GetCurSel( pDoc->hDesignTabCtrl )
            if idx = 0 then
               ' Designer is active, switch to code
               TabCtrl_SetCurSel( pDoc->hDesignTabCtrl, 1 )  ' 1 = code tab, 0 = Designer
            else
               ' Code is active, switch to designer
               TabCtrl_SetCurSel( pDoc->hDesignTabCtrl, 0 )  ' 1 = code tab, 0 = Designer
            end if
            frmMain_PositionWindows
         end if
            
      case IDM_MENUEDITOR            
         frmMenuEditor_Show( HWND )
         if pDoc then frmMenuEditor_CreateFakeMainMenu( pDoc )

      case IDM_TOOLBAREDITOR            
         frmToolBarEditor_Show( HWND )
         if pDoc then frmToolBarEditor_CreateFakeToolBar( pDoc )

      case IDM_STATUSBAREDITOR            
         frmStatusBarEditor_Show( HWND )
         if pDoc then frmStatusBarEditor_CreateFakeStatusBar( pDoc )

      case IDM_IMAGEMANAGER
         frmImageManager_Show( HWND, 0 )
         
      case IDM_ALIGNLEFTS, IDM_ALIGNCENTERS, IDM_ALIGNRIGHTS, _
           IDM_ALIGNTOPS, IDM_ALIGNMIDDLES, IDM_ALIGNBOTTOMS, _
           IDM_SAMEWIDTHS, IDM_SAMEHEIGHTS, IDM_SAMEBOTH            
         OnCommand_DesignerAlign( HWND, id )

      case IDM_CENTERHORIZ, IDM_CENTERVERT, IDM_CENTERBOTH
         OnCommand_DesignerCenter( HWND, id )
         
      case IDM_HORIZEQUAL, IDM_HORIZINCREASE, IDM_HORIZDECREASE, IDM_HORIZREMOVE
         OnCommand_DesignerHorizSpacing( hwnd, id )

      case IDM_VERTEQUAL, IDM_VERTINCREASE, IDM_VERTDECREASE, IDM_VERTREMOVE
         OnCommand_DesignerVertSpacing( hwnd, id )
      
      case IDM_SNAPLINES
         if pDoc then 
            pDoc->bSnapLines = not pDoc->bSnapLines
            pDoc->UserModified = true
            frmMain_SetStatusbar
         end if

      case IDM_LOCKCONTROLS
         if pDoc then 
            pDoc->bLockControls = not pDoc->bLockControls
            pDoc->UserModified = true
            AfxRedrawWindow(pDoc->hWndFrame)
            AfxRedrawWindow(pDoc->hWndForm)
            frmMain_SetStatusbar
         end if
      
      ''   TOOLS MENU
      Case IDM_USERTOOL To (IDM_USERTOOL + UBound(gConfig.Tools))
         frmUserTools_ExecuteUserTool( id - IDM_USERTOOL )

      Case IDM_HELP, IDM_HELPTIPS, IDM_HELPWINAPI, IDM_HELPWINFBE, IDM_HELPWINFBX
         ShowContextHelp(id) 
  
      case IDM_CHECKFORUPDATES
         DisableAllModeless()
         DoCheckForUpdates( hwnd, false )
         EnableAllModeless()
         
      Case IDM_ABOUT
         DisableAllModeless()
         frmAbout_Show( hwnd )
         EnableAllModeless()
         frmMain_SetFocusToCurrentCodeWindow()
        
   End Select

   Function = 0
End Function


' ========================================================================================
' Process WM_NOTIFY message for window/dialog: frmMain
' ========================================================================================
private Function frmMain_OnNotify( ByVal HWnd As HWnd, _
                                   ByVal id As Long, _
                                   ByVal pNMHDR As NMHDR Ptr _
                                   ) As LRESULT

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
               frmBuildConfig_Show( HWND ) 
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
               frmMain_PositionWindows
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
         select case id 
            case IDC_FRMMAIN_TOPTABCONTROL 
               ' Hide the current tab
               nCurSel = TabCtrl_GetCurSel(pNMHDR->hwndFrom) 
               gTTabCtl.DisplayScintilla(nCurSel, False)
            case IDC_DESIGNTABCTRL 
               frmMain_PositionWindows
         END SELECT
         
      Case TCN_SELCHANGE
         dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr()
         select case id 
            case IDC_FRMMAIN_TOPTABCONTROL 
               ' Show the new tab
               nCurSel = TabCtrl_GetCurSel(pNMHDR->hwndFrom) 
               gTTabCtl.DisplayScintilla(nCurSel, True)
               frmMain_PositionWindows
               dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr()
               If pDoc Then
                  ' Prevent Treeview flickering during the selection movement. This
                  ' is necessary because for large treeview we will see noticable flicker
                  ' as the new selection scrolls into view.
                  dim as hwnd hTV = GetDlgItem(HWND_FRMEXPLORER, IDC_FRMEXPLORER_TREE)
                  LockWindowUpdate(hTV)
                  Treeview_SelectItem( hTV, pDoc->hNodeExplorer )
                  LockWindowUpdate(0)
                  
                  ' Select the correct Files combobox item based on the current document
                  frmMain_SelectComboFiles
                  
                  ' Load the Functions list combobox and set it to the current function
                  ' based on line position.
                  frmMain_LoadComboFunctions
                  frmMain_SelectComboFunctions
               
                  if pDoc->IsDesigner THEN
                     DisplayPropertyList( pDoc )
                  end if   
               end if

            case IDC_DESIGNTABCTRL 
               frmMain_PositionWindows
               PostMessage(hwnd, MSG_USER_GENERATECODE, 0, 0)
         END SELECT  

         
      Case TTN_GETDISPINFO
         ptttdi = Cast(NMTTDISPINFOW Ptr, pNMHDR)
         ptttdi->hinst = Null
         wszText = ""
         Select Case ptttdi->hdr.hwndFrom
            Case SendMessage(GetDlgItem(GetDlgItem(HWnd, IDC_FRMMAIN_REBAR), IDC_FRMMAIN_TOOLBAR), TB_GETTOOLTIPS, 0, 0)
               Select Case ptttdi->hdr.idFrom
                  Case IDM_FILENEW         : wszText = L(03,"New")     & " (Ctrl+N)"
                  Case IDM_FILEOPEN        : wszText = L(04,"Open")    & " (Ctrl+O)"
                  Case IDM_FILESAVE        : wszText = L(07,"Save")    & " (Ctrl+S)"
                  Case IDM_CUT             : wszText = L(17,"Cut")     & " (Ctrl+X)"
                  Case IDM_COPY            : wszText = L(18,"Copy")    & " (Ctrl+C)"
                  Case IDM_PASTE           : wszText = L(19,"Paste")   & " (Ctrl+V)"
                  Case IDM_FIND            : wszText = L(43,"Find")    & " (Ctrl+F)"
                  Case IDM_REPLACE         : wszText = L(46,"Replace") & " (Ctrl+R)"
                  Case IDM_UNDO            : wszText = L(15,"Undo")    & " (Ctrl+Z)"
                  Case IDM_REDO            : wszText = L(16,"Redo")    & " (Ctrl+E)"
                  Case IDM_INDENTBLOCK     : wszText = L(23,"Indent Block")      & " (TAB)"
                  Case IDM_UNINDENTBLOCK   : wszText = L(24,"UnIndent Block")    & " (Shift+TAB)"
                  Case IDM_COMMENTBLOCK    : wszText = L(25,"Comment Block")     & " (Ctrl+B)"
                  Case IDM_UNCOMMENTBLOCK  : wszText = L(26,"UnComment Block")   & " (Ctrl+Shift+B)"
                  Case IDM_FUNCTIONLIST    : wszText = L(223,"Function List")    & " (F4)"
                  Case IDM_COMPILE         : wszText = L(64,"Compile")           & " (Ctrl+F5)"
                  Case IDM_BUILDEXECUTE    : wszText = L(65,"Build and Execute") & " (F5)"
                  Case IDM_QUICKRUN        : wszText = L(285,"Quick Run")        & " (Ctrl+Shift+F5)"
                  Case IDM_RUNEXE          : wszText = L(67,"Run Executable")    & " (Shift+F5)"
                  Case IDM_HELP            : wszText = L(73,"Help")              & " (F1)"
               End Select
               If Len(wszText) Then ptttdi->lpszText = @wszText
         End Select
   End Select

   Function = 0
End Function


' ========================================================================================
' Process WM_ACTIVATEAPP message for window/dialog: frmMain
' ========================================================================================
private Function frmMain_OnActivateApp( ByVal HWnd As HWnd, _
                                        ByVal fActivate As BOOLEAN, _
                                        ByVal dwThreadId As DWORD _
                                        ) As LRESULT

   Dim pDoc As clsDocument Ptr
   
   ' If the application is losing focus then validate any currently editing property
   ' in the ToolBox PropertyList.
   If fActivate = false Then
      pDoc = gTTabCtl.GetActiveDocumentPtr()
      DisplayPropertyList(pDoc)
      exit function
   end if

      
   static bInActivateApp as Boolean
   if bInActivateApp then exit function
   
   ' If the application is gaining focus then determine if any of the loaded
   ' documents have been modified by an external application. If yes, then ask
   ' the user if wish to reload.
   If fActivate = true Then
      if gApp.GetDocumentCount = 0 then exit function
      
      ' Search all loaded documents
      Dim As FILETIME ft
      Dim As Long i, idx, nDocumentCount
      Dim As Double nSerial
      Dim As String sText
      
      bInActivateApp = true
      
      ' Load all pDoc into array and then process. For each pDoc if the user chooses
      ' not to reload or wants to bypass, then set the pDoc to zero until eventually
      ' all pDocs are processed.
      dim pDocs( gApp.GetDocumentCount - 1) as clsDocument ptr
      idx = 0
      pDoc = gApp.pDocList
      do until pDoc = 0
         if AfxFileExists(pDoc->DiskFilename) THEN
            pDoc->DeletedButKeep = false
         end if
         pDocs(idx) = pDoc
         idx = idx + 1
         pDoc = pDoc->pDocNext
      loop
      
      do
         idx = -1
         for i as long = 0 to ubound(pDocs)
            if pDocs(i) then 
               idx = i: exit for
            end if
         next
         if idx = -1 then exit do
         pDoc = pDocs(idx)
         
         ' Bypass any 'new' untitled files.
         If pDoc->IsNewFlag Then 
            pDocs(idx) = 0: continue do
         end if

         ' Has the external file been deleted or is now not available
         if AfxFileExists(pDoc->DiskFilename) = false THEN
            if pDoc->DeletedButKeep = false then
               ' Ensure that the file is open and displayed
               frmMain_OpenFileSafely(HWND_FRMMAIN, _
                                       False, _    ' bIsNewFile
                                       False, _    ' bIsTemplate
                                       True,  _    ' bShowInTab
                                       false, _    ' bIsInclude
                                       0, _        ' pwszName
                                       pDoc )      ' pDocIn
               If MessageBox( HWND_FRMMAIN, _
                               pDoc->DiskFilename & vbCrLf & _ 
                               L(286, "This document has been deleted...?"), _
                               L(267, "File Changed"), MB_ICONQUESTION Or MB_YESNO) = IDYES Then
                  ' Keep the file open and simply mark it as dirty so it can be prompted to be saved.
                  pDoc->UserModified = true                
               else
                  ' No. Close the file.
                  frmMain_OnCommand(HWND_FRMMAIN, IDM_REMOVEFILEFROMPROJECT, 0, 0)
                  pDocs(idx) = 0
                  continue do
               end if
            end if
            pDoc->DeletedButKeep = true

         else
                  
            ' Compare the disk file date time to the value currently
            ' stored in document class.
            ft = AfxGetFileLastWriteTime(pDoc->DiskFilename)
            If AfxFileTimeToVariantTime(ft) <> AfxFileTimeToVariantTime(pDoc->DateFileTime) Then
               OpenSelectedDocument( pDoc->DiskFilename, "" )
               If MessageBox( HWND_FRMMAIN, _
                               pDoc->DiskFilename & vbCrLf & _ 
                               L(266, "File was changed by another application. Reload it?"), _
                               L(267, "File Changed"), MB_ICONQUESTION Or MB_YESNO) = IDYES Then
                  dim sText as string
                  if GetFileToString(pDoc->DiskFilename, sText, pDoc) = false then
                     pDoc->SetText( sText ) 
                  else
                     print "Error opening: "; pDoc->DiskFilename
                  End If
               End If                
            End If
            pDoc->DateFileTime = AfxGetFileLastWriteTime( pDoc->DiskFilename )
         end if

         pDocs(idx) = 0
      loop      
             
      bInActivateApp = false
            
   end if
   Function = 0
End Function


' ========================================================================================
' Process WM_CONTEXTMENU message for window/dialog: frmMain
' ========================================================================================
private Function frmMain_OnContextMenu( ByVal HWnd As HWnd, _
                                        ByVal hwndContext As HWnd, _
                                        ByVal xPos As Long, _  
                                        ByVal yPos As Long _
                                        ) As LRESULT
   Dim hPopupMenu as HMENU 
   Dim pt         As Point 
   
   Dim pDoc As clsDocument Ptr = gTTabCtl.GetActiveDocumentPtr()
   If pDoc = 0 Then Return 0

   SetFocus hWndContext
   pt.x = xPos
   pt.y = yPos
   ScreenToClient hWndContext, @pt
   if IsDesignerView(pDoc) THEN
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
private Function frmMain_OnDropFiles( ByVal HWnd As HWnd, _
                                      ByVal hDrop As HDROP _
                                      ) As LRESULT

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
               frmMain_OpenProjectSafely(HWND_FRMMAIN, wszPath) 
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
private Function frmMain_OnClose( ByVal HWnd As HWnd ) As LRESULT

   ' If configuration option to confirm closing editor is active then ask now.
   if gConfig.AskExit THEN
      if MessageBox( HWND_FRMMAIN, L(275,"Are you sure you want to exit?"), L(276,"Confirm"), _
                  MB_YESNOCANCEL or MB_ICONQUESTION Or MB_DEFBUTTON1 ) <> IDYES then
         return true
      end if            
   END IF
   
   ' Set global shutdown flag that will bypass removing nodes from project
   ' treeview, etc. Those tasks simply slow down the exiting of the program.
   gApp.IsShutdown = true

   ' Save whether the Explorer should be shown the next time the program is run.
   ' Also save the Explorer panel width. 
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(HWND)
   gConfig.ShowExplorer = IsWindowVisible(HWND_FRMEXPLORER)
   gConfig.ShowExplorerWidth = pWindow->UnScaleX(AfxGetWindowWidth( HWND_FRMEXPLORER ))
   
   ' If a project(s) is already open then save/close it.
   If gApp.IsProjectActive Then 
      if OnCommand_ProjectClose(HWnd) = 0 then return 0
   end if
   
   ' Close any open files asking to save any that are dirty 
   if OnCommand_FileClose(HWnd, EFC_CLOSEALL) = 0 then return 0
   
   DestroyWindow(HWnd)
   
   Function = 0
End Function


' ========================================================================================
' Process WM_DESTROY message for window/dialog: frmMain
' ========================================================================================
private Function frmMain_OnDestroy( byval HWnd As HWnd ) As LRESULT
   ImageList_Destroy Cast(HIMAGELIST, SendMessage(GetDlgItem(HWnd, IDC_FRMMAIN_TOOLBAR), TB_GETIMAGELIST, 0, 0))
   ImageList_Destroy Cast(HIMAGELIST, SendMessage(GetDlgItem(HWnd, IDC_FRMMAIN_TOOLBAR), TB_GETHOTIMAGELIST, 0, 0))
   ImageList_Destroy Cast(HIMAGELIST, SendMessage(GetDlgItem(HWnd, IDC_FRMMAIN_TOOLBAR), TB_GETDISABLEDIMAGELIST, 0, 0))

   ImageList_Destroy TabCtrl_GetImageList(gTTabCtl.hWindow) 
   ImageList_Destroy TreeView_GetImageList( GetDlgItem(HWND_FRMEXPLORER, IDC_FRMEXPLORER_TREE), TVSIL_NORMAL )
   
   ' Output the config settings to disk file
   gConfig.SaveConfigFile

   ' Disable drag and drop files
   DragAcceptFiles HWnd, False
   
   PostQuitMessage(0)
   Function = 0
End Function


' ========================================================================================
' Processes messages for the subclassed frmMain TabControl control.
' ========================================================================================
private Function frmMain_TabCtl_SubclassProc ( _
                  ByVal HWnd   As HWnd, _                 ' // Control window handle
                  ByVal uMsg   As UINT, _                 ' // Type of message
                  ByVal wParam As WPARAM, _               ' // First message parameter
                  ByVal lParam As LPARAM, _               ' // Second message parameter
                  ByVal uIdSubclass As UINT_PTR, _        ' // The subclass ID
                  ByVal dwRefData As DWORD_PTR _          ' // Pointer to reference data
                  ) As LRESULT

   Dim pInfo As TCHITTESTINFO
   dim as long nCurSel


   ' Convert our ENTER key presses into LBUTTONDBLCLK to process them similarly
   If (uMsg = WM_KEYUP) And (Loword(wParam) = VK_RETURN) Then uMsg = WM_LBUTTONDBLCLK
      
   Select Case uMsg

      Case WM_LBUTTONDBLCLK 
         ' Process double clicking on a tab in the top tab control (which between code and visual designer)
         ' Get the tab number
         GetCursorPos( @pInfo.pt )
         ScreenToClient gTTabCtl.hWindow, @pInfo.pt
         nCurSel = TabCtrl_HitTest(gTTabCtl.hWindow, @pInfo)
         if nCurSel = -1 then exit function
         
         ' Get the pDoc for this tab and check to see if it is a visual designer. If it is, then
         ' initiate the toggle between the code editor and visual designer.
         dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr
         if pDoc then
            if pDoc->IsDesigner then
               nCurSel = TabCtrl_GetCurSel(pDoc->hDesignTabCtrl)
               nCurSel = iif(nCurSel = 0, 1, 0)
               TabCtrl_SetCurFocus(pDoc->hDesignTabCtrl, nCurSel)
               frmMain_PositionWindows
            END IF
         END IF
         
      Case WM_RBUTTONDOWN
         ' Get the tab number
         GetCursorPos( @pInfo.pt )
         ScreenToClient gTTabCtl.hWindow, @pInfo.pt
         nCurSel = TabCtrl_HitTest(gTTabCtl.hWindow, @pInfo)
         gTTabCtl.SetFocusTab(nCurSel)
         ' Create the popup menu
         Dim As POINT pt
         dim as HMENU hPopupMenu = CreateTopTabCtlContextMenu(nCurSel)
         GetCursorPos @pt
         TrackPopupMenu(hPopUpMenu, 0, pt.x, pt.y, 0, HWnd_FRMMAIN, ByVal Null)
         DestroyMenu hPopUpMenu
         Return True   ' prevent further processing that leads to WM_CONTEXTMENU

      Case WM_LBUTTONDOWN
         ' Get the position of the hit and determine if the CLOSE icon was clicked.
         GetCursorPos( @pInfo.pt )
         ScreenToClient( gTTabCtl.hWindow, @pInfo.pt )
         nCurSel = TabCtrl_HitTest(gTTabCtl.hWindow, @pInfo)
         ' Was the close icon pressed
         If pInfo.flags = TCHT_ONITEMICON Then 
            OnCommand_FileClose( HWND_FRMMAIN, EFC_CLOSECURRENT, nCurSel )
         else
            gTTabCtl.SetFocusTab(nCurSel)
            if (wParam and MK_SHIFT) then
               gApp.bDragTabActive = true
               SetCursor( LoadCursor(0, MAKEINTRESOURCE(OCR_SIZEALL)) )
            end if
         end if
         SetCapture(hWnd)
            
      Case WM_MOUSEMOVE
         if gApp.bDragTabActive then
            SetCursor( LoadCursor(0, MAKEINTRESOURCE(OCR_SIZEALL)) )

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

      Case WM_LBUTTONUP
         ' Reset the mouse pointer
         SetCursor( LoadCursor( null, IDC_ARROW ))
         ReleaseCapture
         gApp.bDragTabActive = false

         
      Case WM_DESTROY
         ' REQUIRED: Remove control subclassing
         RemoveWindowSubclass HWnd, @frmOutput_Listview_SubclassProc, uIdSubclass

   End Select

   ' Default processing of Windows messages
   Function = DefSubclassProc(HWnd, uMsg, wParam, lParam)

End Function



' ========================================================================================
' Process WM_MEASUREITEM message for window/dialog: frmMain
' ========================================================================================
private Function frmMain_OnMeasureItem( ByVal HWnd As HWnd, _
                                        ByVal lpmis As MEASUREITEMSTRUCT Ptr _
                                        ) As Long
   ' Set the height of the List box items. 
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(HWnd)
   lpmis->itemHeight = pWindow->ScaleY(FRMVDTOOLBOX_LISTBOX_LINEHEIGHT)
   Function = 0
End Function


' ========================================================================================
' Process WM_DRAWITEM message for window/dialog: frmMain
' ========================================================================================
private Function frmMain_OnDrawItem( ByVal HWnd As HWnd, _
                                     ByVal lpdis As Const DRAWITEMSTRUCT Ptr _
                                     ) As Long

   If lpdis->itemID = -1 Then Exit Function
   
   Dim As HBRUSH hBrush 
   dim as HFONT hFontNormal, hFontBold 
   Dim As RECT rc
   Dim wszText As WString * MAX_PATH
   
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(HWnd)
   if pWindow = 0 THEN exit function
      
   dim pTheme as TYPE_THEMES ptr = gConfig.GetThemePtr
   

   Select Case lpdis->itemAction
      Case ODA_DRAWENTIRE, ODA_SELECT
         
         SaveDC(lpdis->hDC)
         
         if (lpdis->CtlID = IDC_FRMMAIN_COMBOFILES) orelse _
            (lpdis->CtlID = IDC_FRMMAIN_COMBOFUNCTIONS) orelse _
            (lpdis->CtlID = IDC_FRMMAIN_COMBOBUILDS) then
            
            ' CLEAR BACKGROUND
            ' The following code uses the theme as definied for the Scintilla
            ' code editing component. We could use those colors so that a dark
            ' code editing window would also give us dark comboboxes. I am not
            ' convinced that is the best route to go. Maybe. Propbably later
            ' add an option to customize the color settings or toggle the 
            ' visibility of these comboxes altogether.
            'If (lpdis->itemState And ODS_SELECTED) Then     
            '   SetTextColor( lpdis->hDC, pTheme->colors(CLR_SELECTION).nFg )
            '   SetBkColor( lpdis->hDC, pTheme->colors(CLR_SELECTION).nBg )
            '   hBrush = CreateSolidBrush( pTheme->colors(CLR_SELECTION).nBg )
            'else
            '   SetTextColor( lpdis->hDC, pTheme->colors(CLR_TEXT).nFg )
            '   SetBkColor( lpdis->hDC, pTheme->colors(CLR_WINDOW).nFg )
            '   hBrush = CreateSolidBrush( pTheme->colors(CLR_WINDOW).nFg )
            'end if
            If (lpdis->itemState And ODS_SELECTED) Then     
               SetBkColor(lpdis->hDC, GetSysColor(COLOR_HIGHLIGHT))   
               SetTextColor(lpdis->hDC, GetSysColor(COLOR_HIGHLIGHTTEXT))
               hBrush = GetSysColorBrush(COLOR_HIGHLIGHT) 
            else
               SetBkColor(lpdis->hDC, GetSysColor(COLOR_WINDOW))   
               SetTextColor(lpdis->hDC, GetSysColor(COLOR_WINDOWTEXT))
               hBrush = GetSysColorBrush(COLOR_WINDOW) 
            end if
            SelectObject(lpdis->hDC, hBrush)      
            FillRect(lpdis->hDC, @lpdis->rcItem, hBrush)  

            rc = lpdis->rcItem
            rc.left = rc.left + pWindow->ScaleX( iif(lpdis->CtlID = IDC_FRMMAIN_COMBOBUILDS, 8, 16) )
            ComboBox_GetLBText( lpdis->hwndItem, lpdis->itemID, @wszText ) 
            DrawText( lpdis->hDC, _
                      wszText, _
                      -1, Cast(lpRect, @rc), _
                      DT_LEFT Or DT_SINGLELINE Or DT_VCENTER or DT_NOPREFIX)
                
         end if

         RestoreDC(lpdis->hDC, -1)
         If hBrush then DeleteObject( hBrush )
         
         Function = True : Exit Function
   
   End Select

   Function = 0
End Function


' ========================================================================================
' Process WM_MOUSEMOVE message for window/dialog: frmMain
' ========================================================================================
public Function frmMain_OnMouseMove( ByVal HWnd As HWnd, _
                                     ByVal x as long, _
                                     byval y as long, _
                                     byval keyflags as UINT _
                                     ) As Long

   Dim pWindow As CWindow Ptr = AfxCWindowPtr( HWND_FRMMAIN )
   if pWindow = 0 then exit function
   
   dim as RECT rcClient
   GetClientRect( hwnd, @rcClient )
   
   ' HITTEST (DOCUMENT SPLITTER)
   Dim pDoc As clsDocument Ptr = gTTabCtl.GetActiveDocumentPtr()
   if pDoc THEN 
      Dim As Rect rc
      Dim As Point pt = (x, y)
      if pDoc->bSizing then
         if pt.y <> pDoc->ptPrev.y then
            pDoc->SplitY = pDoc->SplitY + (pt.y - pDoc->ptPrev.y)
            ' Don't let the split go all the way to the bottom of the edit window
            if pDoc->SplitY + pWindow->ScaleY(40) >= rcClient.bottom THEN 
            else
               frmMain_PositionWindows
               pDoc->ptPrev.y = pt.y
            end if   
         end if   
      else   
         if PtInRect(@pDoc->rcSplitButton, pt) THEN 
            SetCursor LoadImage( Null, MAKEINTRESOURCEW(OCR_SIZENS), IMAGE_CURSOR, 0, 0, LR_DEFAULTSIZE Or LR_SHARED )
         end if   
      END if
   end if
         
   function = 0
end function


' ========================================================================================
' Process WM_LBUTTONDOWN message for window/dialog: frmMain
' ========================================================================================
private Function frmMain_OnLButtonDown( ByVal HWnd As HWnd, _
                                        byval fDoubleClick as Boolean, _
                                        ByVal x as long, _
                                        byval y as long, _
                                        byval keyflags as UINT _
                                        ) As Long
                                        
   ' HITTEST (DOCUMENT SPLITTER)
   Dim pDoc As clsDocument Ptr = gTTabCtl.GetActiveDocumentPtr()
   if pDoc THEN 
      Dim As Rect rc 
      Dim As Point pt = (x, y)
      if PtInRect(@pDoc->rcSplitButton, pt) THEN 
         pDoc->bSizing = true
         pDoc->SplitY = pDoc->rcSplitButton.top  
         pDoc->ptPrev.y = pt.y
         SetCursor LoadImage( Null, MAKEINTRESOURCEW(OCR_SIZENS), IMAGE_CURSOR, 0, 0, LR_DEFAULTSIZE Or LR_SHARED )
         SetCapture( HWND_FRMMAIN )
      END IF
   end if

   function = 0
end function

         
' ========================================================================================
' Process WM_LBUTTONUP message for window/dialog: frmMain
' ========================================================================================
private Function frmMain_OnLButtonUp( ByVal HWnd As HWnd, _
                                      ByVal x as long, _
                                      byval y as long, _
                                      byval keyflags as UINT _
                                      ) As Long

   ' HITTEST (DOCUMENT SPLITTER)
   Dim pDoc As clsDocument Ptr = gTTabCtl.GetActiveDocumentPtr()
   if pDoc THEN 
      if pDoc->bSizing then
         pDoc->bSizing = false
         ReleaseCapture
      end if
   end if   
   SetCursor( LoadCursor( null, IDC_ARROW ))
         
   function = 0
end function


' ========================================================================================
' Process WM_LBUTTONDOWN message for window/dialog: frmMain
' ========================================================================================
private Function frmMain_OnLButtonDblClk( ByVal HWnd As HWnd, _
                                          byval fDoubleClick as Boolean, _
                                          ByVal x as long, _
                                          byval y as long, _
                                          byval keyflags as UINT _
                                          ) As Long

   if gApp.bDragActive = false then
      ' Are we over a split edit area (toggle off the split)
      dim as POINT pt = (x, y)
      Dim pDoc As clsDocument Ptr = gTTabCtl.GetActiveDocumentPtr()
      if pDoc THEN 
         if PtInRect(@pDoc->rcSplitButton, pt) THEN 
            pDoc->bSizing = false
            pDoc->SplitY = 0
            ReleaseCapture
            frmMain_PositionWindows
         END IF
      end if
   end if
   
   function = 0
end function


' ========================================================================================
' frmMain Window procedure
' ========================================================================================
private Function frmMain_WndProc( ByVal HWnd   As HWnd, _
                                  ByVal uMsg   As UINT, _
                                  ByVal wParam As WPARAM, _
                                  ByVal lParam As LPARAM _
                                  ) As LRESULT

   Select Case uMsg
      HANDLE_MSG (HWnd, WM_CREATE,        frmMain_OnCreate)
      HANDLE_MSG (HWnd, WM_PAINT,         frmMain_OnPaint)
      HANDLE_MSG (HWnd, WM_SIZE,          frmMain_OnSize)
      HANDLE_MSG (HWnd, WM_CLOSE,         frmMain_OnClose)
      HANDLE_MSG (HWnd, WM_DESTROY,       frmMain_OnDestroy)
      HANDLE_MSG (HWnd, WM_COMMAND,       frmMain_OnCommand)
      HANDLE_MSG (HWnd, WM_NOTIFY,        frmMain_OnNotify)
      HANDLE_MSG (HWnd, WM_ACTIVATEAPP,   frmMain_OnActivateApp)
      HANDLE_MSG (HWnd, WM_CONTEXTMENU,   frmMain_OnContextMenu)
      HANDLE_MSG (HWnd, WM_DROPFILES,     frmMain_OnDropFiles)
      HANDLE_MSG (HWnd, WM_MEASUREITEM,   frmMain_OnMeasureItem)
      HANDLE_MSG (HWnd, WM_DRAWITEM,      frmMain_OnDrawItem)
      HANDLE_MSG (HWnd, WM_LBUTTONDOWN,   frmMain_OnLButtonDown)
      HANDLE_MSG (HWnd, WM_LBUTTONUP,     frmMain_OnLButtonUp)
      HANDLE_MSG (HWnd, WM_MOUSEMOVE,     frmMain_OnMouseMove)
      HANDLE_MSG (HWnd, WM_LBUTTONDBLCLK, frmMain_OnLButtonDblClk)

      case WM_MOVE
         ' Ensure that if the Find/Replace dialog is open that it moves with
         ' underlying main WinFBE window.     
         frmFindReplace_PositionWindow
      
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
         If gApp.bDragActive = true Then
            ' Either the horiz or vert splitter bar is being resized. We handle
            ' setting the mouse cursor in the splitter_mousemove function.
            return true
            
         elseif gApp.bDragTabActive = true Then 
            ' We handle the cursor in the TabControl subclass. Don't allow the
            ' main window to change our cursor. Return True.
            return true

         elseif (gApp.IsCompiling = true) or _
                (gApp.IsProjectLoading = true) or _
                (gApp.IsFileLoading = true) then
            SetCursor( LoadCursor(0, IDC_WAIT) )
            return true
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
               frmMain_OpenProjectSafely(HWND_FRMMAIN, *pwszArg) 
             
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
         frmMain_SetStatusbar
         Dim pDoc As clsDocument Ptr = gTTabCtl.GetActiveDocumentPtr()
         if pDoc THEN
            if pDoc->IsDesigner THEN
               SendMessage pDoc->hWndForm, WM_NCACTIVATE, true, 0
            END IF
            SetFocus(pDoc->hWndActiveScintilla)
         END IF

      Case MSG_USER_PROCESS_COMMANDLINE 
         ' process any command line arguments that may have been passed to the program.
         frmMain_ProcessCommandLine(HWnd)

      Case MSG_USER_PROCESS_UPDATECHECK
         ' perform an update check one per day if the option has been set by the user.
         if gConfig.CheckForUpdates then
            dim as long curJulian = JulianDateNow
            if gConfig.LastUpdateCheck <> curJulian then
               DoCheckForUpdates( hwnd, true )  ' no messages if up to date
            end if
            ' Save the config file so that other editor instances will not also do update checks again
            gConfig.LastUpdateCheck = curJulian
            gConfig.SaveConfigFile
         end if   

      case MSG_USER_GENERATECODE
         Dim pDoc As clsDocument Ptr 
         if lParam = 0 then
            pDoc = gTTabCtl.GetActiveDocumentPtr()
         else
            pDoc = cast(clsDocument Ptr, lParam)
         end if   
         if pDoc then
            if IsDesignerView(pDoc) = false then  ' clicked on the code window
               dim as hwnd hEdit = pDoc->hWindow(0)
               dim as long nFirstLine = SendMessage( hEdit, SCI_GETFIRSTVISIBLELINE, 0, 0)
               dim as long curPos = SendMessage( hEdit, SCI_GETCURRENTPOS, 0, 0)
               GenerateFormCode(pDoc)
               if curPos >=0 then SendMessage( hEdit, SCI_GOTOPOS, curPos, 0)
               if nFirstLine >=0 then SendMessage( hEdit, SCI_SETFIRSTVISIBLELINE, nFirstLine, 0)
            else
               GenerateFormCode(pDoc)
            END IF
         end if

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
public Function frmMain_Show( ByVal hWndParent As HWnd ) as LRESULT

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
   
   ' Load the tick and untick icons
   ghIconTick   = LoadImage( pWindow->InstanceHandle, "IMAGE_TICK", IMAGE_ICON, cx, cx, LR_DEFAULTCOLOR )
   ghIconNoTick = LoadImage( pWindow->InstanceHandle, "IMAGE_NOTICK", IMAGE_ICON, cx, cx, LR_DEFAULTCOLOR )

   ' Create the main application topmenu
   HWND_FRMMAIN_TOPMENU = frmMain_BuildMenu(pWindow)
   SetMenu pWindow->hWindow, HWND_FRMMAIN_TOPMENU

   ' Create the main application toolbar
   frmMain_CreateToolbar(pWindow)
   frmMain_ChangeToolbarButtonsState

   ' Load the build configurations into the toolbar
   frmBuildConfig_LoadBuildComboBox()

   
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
      WS_CHILD Or TCS_SINGLELINE Or TCS_RAGGEDRIGHT Or TCS_HOTTRACK Or _
      TCS_TABS Or TCS_FOCUSNEVER Or TCS_FORCEICONLEFT, WS_EX_LEFT Or WS_EX_LTRREADING, , _
      Cast(SUBCLASSPROC, @frmMain_TabCtl_SubclassProc), IDC_FRMMAIN_TOPTABCONTROL, Cast(DWORD_PTR, @pWindow))
    
   ' Create the imagelist for the TopTabControl tabs
   cx = 16 * (pWindow->DPI \ 96)
   Dim As HANDLE hImageList = ImageList_Create(cx, cx, ILC_MASK Or ILC_COLOR32, 1, 0)
   Dim As HICON hIcon = LoadIcon( pWindow->InstanceHandle, "IMAGE_TABCLOSE" )
   ImageList_AddIcon( hImageList, hIcon )
   TabCtrl_SetImageList( gTTabCtl.hWindow, hImageList )
                                                        
   
   ' Add the file list navigation combo box
   pWindow->AddControl("COMBOBOX", pWindow->hWindow, _
      IDC_FRMMAIN_COMBOFILES, "", 0, 0, 0, 20, _
      WS_CHILD OR WS_VISIBLE OR WS_VSCROLL OR CBS_OWNERDRAWFIXED or _
      CBS_DROPDOWNLIST OR CBS_HASSTRINGS OR CBS_SORT, WS_EX_CLIENTEDGE )
   
   ' Add the functions list navigation combo box
   pWindow->AddControl("COMBOBOX", pWindow->hWindow, _
      IDC_FRMMAIN_COMBOFUNCTIONS, "", 0, 0, 0, 20, _
      WS_CHILD OR WS_VISIBLE OR WS_VSCROLL OR CBS_OWNERDRAWFIXED or  _
      CBS_DROPDOWNLIST OR CBS_HASSTRINGS OR CBS_SORT, WS_EX_CLIENTEDGE )

   
   ' Load the graphisc PNG pointers for the scintilla popup autocomplete listbox.
   gApp.pImageAutocompleteBaseType = LoadPNGfromRes( pWindow->InstanceHandle, "IMAGE_BASETYPE" )      
   gApp.pImageAutocompleteClass    = LoadPNGfromRes( pWindow->InstanceHandle, "IMAGE_CLASS" )     
   gApp.pImageAutocompleteMethod   = LoadPNGfromRes( pWindow->InstanceHandle, "IMAGE_METHOD" )     
   gApp.pImageAutocompleteProperty = LoadPNGfromRes( pWindow->InstanceHandle, "IMAGE_PROPERTY" )  


   ' Create the various child windows
   frmExplorer_Show( HWND_FRMMAIN )
   frmRecent_Show( HWND_FRMEXPLORER )     ' child of frmExplorer
   frmOutput_Show( HWND_FRMMAIN )


   ' Create the UserTools accelerator table
   frmUserTools_CreateAcceleratorTable            
   

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
   PostMessage( pWindow->hWindow, MSG_USER_PROCESS_COMMANDLINE, 0, 0 )

   ' Post a message to do an update check (if applicable)
   PostMessage( pWindow->hWindow, MSG_USER_PROCESS_UPDATECHECK, 0, 0 )

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

            ' Close any open Find/Replace window with ESC key
            ' Close any open FunctionList window
            if (uMsg.message = WM_KEYDOWN) andAlso (uMsg.wParam = VK_ESCAPE) THEN
               if IsWindowVisible( HWND_FRMFINDREPLACE ) THEN SendMessage( HWND_FRMFINDREPLACE, WM_CLOSE, 0, 0 )
               if IsWindowVisible( HWND_FRMFUNCTIONLIST ) THEN SendMessage( HWND_FRMFUNCTIONLIST, WM_CLOSE, 0, 0 )
            END IF

            ' Handle any keypress that would move or resize control(s) on a Designer Form.
            if uMsg.message = WM_KEYDOWN THEN
               dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr
               if pDoc THEN
                  ' We don't want keystrokes going into the scintilla control if the
                  ' visual designer form is active.
                  if pDoc->IsDesigner THEN
                     if IsDesignerView(pDoc) then
                        ' Don't allow keys to the scintilla control
                        if GetDlgCtrlID(GetFocus) = IDC_SCINTILLA then 
                           if uMsg.wParam = VK_DELETE then
                              frmMain_OnCommand(HWND_FRMMAIN, IDM_DELETE, 0, 0)
                              continue do
                           end if   
                        end if
                     end if
                  end if
               end if         

               ' Ensure that an arrow key was pressed
               select CASE uMsg.wParam

                  case VK_SPACE
                     dim as HWND hEventList = GetDlgItem(HWND_FRMVDTOOLBOX, IDC_FRMVDTOOLBOX_LSTEVENTS)
                     if pDoc andalso GetFocus() = hEventList then
                        ' Toggle the selected Event listbox item
                        dim as long nCurSel = ListBox_GetCurSel(hEventList)
                        if nCurSel > -1 then 
                           dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr
                           dim pCtrl as clsControl ptr 
                           if pDoc then
                              pCtrl = pDoc->Controls.GetActiveControl
                              if pCtrl THEN
                                 pCtrl->Events(nCurSel).bIsSelected = not pCtrl->Events(nCurSel).bIsSelected
                                 AfxRedrawWindow(hEventList)
                                 pDoc->UserModified = true
                                 pDoc->bRegenerateCode = true
                                 frmMain_SetStatusbar
                              end if
                           end if
                        end if   
                     END IF

                  case VK_LEFT, VK_RIGHT, VK_UP, VK_DOWN
                     if pDoc THEN
                        if pDoc->IsDesigner THEN
                           select case GetFocus
                              ' Allow arrows to move selected item in a listbox or edit textbox
                              case GetDlgItem(HWND_FRMVDTOOLBOX, IDC_FRMVDTOOLBOX_LSTTOOLBOX), _
                                   GetDlgItem(HWND_FRMVDTOOLBOX, IDC_FRMVDTOOLBOX_LSTPROPERTIES), _
                                   GetDlgItem(HWND_FRMVDTOOLBOX, IDC_FRMVDTOOLBOX_LSTEVENTS), _
                                   GetDlgItem(HWND_FRMEXPLORER, IDC_FRMEXPLORER_TREE), _
                                   HWND_PROPLIST_EDIT
                              CASE ELSE
                                 ' Allow arrows to move/resize selected controls
                                 if IsDesignerView(pDoc) then
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
                                 end if
                           END SELECT
                        END IF
                     END IF
               END SELECT
            end if
             
            ' Check for any QuickRun exes that can be deleted.
            gApp.CheckQuickRunExe
             
            ' Determines whether a message is intended for the specified
            ' dialog box and, if it is, processes the message.
            ' Ensure keystrokes like TAB are properly handled by the modeless dialogs
            IF AfxCAxHostForwardMessage(GetFocus, @uMsg) = FALSE THEN
               If IsDialogMessage( HWND_FRMFINDINFILES, @uMsg ) Then Continue Do
               If IsDialogMessage( HWND_FRMFINDREPLACE, @uMsg ) Then Continue Do
               IF IsDialogMessageW(HWND_FRMHELPVIEWER, @uMsg) Then Continue Do

               If IsDialogMessage(pWindow->hWindow, @uMsg) = 0 Then
                  TranslateMessage @uMsg    ' Translates virtual-key messages into character messages.
                  DispatchMessage @uMsg     ' Dispatches a message to a window procedure.
               end if
            End If
         end if
         
      End If
   Loop
   Function = uMsg.wParam

   If ghIconGood       Then DestroyIcon( ghIconGood ) 
   If ghIconBad        Then DestroyIcon( ghIconBad )
   if ghAccelUserTools THEN DestroyAcceleratorTable( ghAccelUserTools )
   
   if gApp.pImageAutocompleteBaseType then deallocate( gApp.pImageAutocompleteBaseType )
   if gApp.pImageAutocompleteClass    then deallocate( gApp.pImageAutocompleteClass )
   if gApp.pImageAutocompleteMethod   then deallocate( gApp.pImageAutocompleteMethod )
   if gApp.pImageAutocompleteProperty then deallocate( gApp.pImageAutocompleteProperty )

   ' Delete the allocated memory for the various child windows
   pWindow = AfxCWindowPtr( HWND_FRMHELPVIEWER ): Delete pWindow
   pWindow = AfxCWindowPtr( HWND_FRMEXPLORER ):   Delete pWindow
   pWindow = AfxCWindowPtr( HWND_FRMRECENT ):     Delete pWindow
   pWindow = AfxCWindowPtr( HWND_FRMOUTPUT ):     Delete pWindow
   pWindow = AfxCWindowPtr( HWND_FRMMAIN ):       Delete pWindow

End Function
