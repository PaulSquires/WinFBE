'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2025 Paul Squires, PlanetSquires Software
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
#include once "frmPanel.bi"
#include once "modMRU.bi"
#include once "modAutoInsert.bi"
#include once "modMenus.bi"
#include once "modCompile.bi"
#include once "mod302Upgrade.bi"

  
' ========================================================================================
' Update the main form statusbar. This is the only routine that updates
' the statusbar in the entire program.
' ========================================================================================
function frmMain_SetStatusbar() as long

   ' Update the statusbar with the current Line/Col position
   dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr()
   
   dim wszText as wstring * MAX_PATH
   
   ' PANEL (0) Line/Col/Sel -or- VD coordinates
   ' PANEL (1) Compiling Results -or- Filename being parsed
   ' PANEL (2) Build Configuration
   ' PANEL (3) Filetype (Normal/Main/Resource/Module)
   ' PANEL (4) Spacing
   ' PANEL (5) File Encoding (Ansi, UTF-8, etc)
   ' PANEL (6) Line Endings

   ' blank out the current statusbar values  
   for i as long = lbound(gSBPanels) to ubound(gSBPanels)
      gSBPanels(i).wszText = ""
   next

   ' PANEL (1) Compiling Results -or- Filename being parsed
   gSBPanels(1).wszText = gApp.wszPanelText
   'hIconPanel4 = gApp.hIconPanel

   if (gApp.IsProjectLoading) orelse (gApp.IsFileLoading) then
      gSBPanels(1).wszText = L(66,"Parsing") & ": (" & gApp.FileLoadingCount & ")  " & gApp.wszPanelText
      
   elseif pDoc <> 0 then
      dim as HWND hEdit = pDoc->hWndActiveScintilla
      wszText = ""
      if (pDoc->IsDesigner) andalso (IsDesignerView(pDoc)) then 
         dim pCtrl as clsControl ptr = pDoc->Controls.GetActiveControl
         if pCtrl then
            wszText = "L:" & GetControlProperty(pCtrl, "LEFT") & ", T:" & GetControlProperty(pCtrl, "TOP") & "  ::  " & _
                      "W:" & GetControlProperty(pCtrl, "WIDTH") & " x H:" & GetControlProperty(pCtrl, "HEIGHT")
         end if
      else
         dim as long curPos, nLine, nCol
         dim as long startPos, endPos, startLine, endLine, nLines 
         ' Retrieve the information and show it in the status bar
         curPos    = SciExec( hEdit, SCI_GETCURRENTPOS, 0, 0)
         nLine     = SciExec( hEdit, SCI_LINEFROMPOSITION, curPos, 0) 
         nCol      = SciExec( hEdit, SCI_GETCOLUMN, curPos, 0) 
         pDoc->GetSelectedLineRange(startLine, endLine, startPos, endPos )

         wszText = " Ln " & str(nLine + 1) & ", Col " & str(nCol + 1) 
         if endPos - startPos then  ' there is selected text
            wszText = wszText + " (" + str(endPos - startPos) + " selected)"
         end if 

         ' Should we display the "shadow" under the tabcontrol
         frmTopTabs_ShowShadow()

         ' Update the vertical scrollbar
         frmEditorVScroll_calcVThumbRect(pDoc)
         AfxRedrawWindow( iif( hEdit = pDoc->hWindow(0), HWND_FRMEDITOR_VSCROLLBAR(0), HWND_FRMEDITOR_VSCROLLBAR(1)) )

      end if
      
      ' PANEL (0) Line/Col/Sel -or- VD coordinates
      gSBPanels(0).wszText = wszText

      ' PANEL (2) Build Configuration
      if gApp.GetDocumentCount then
         gSBPanels(2).wszText = frmBuildConfig_GetSelectedBuildDescription()
      end if
      
      ' PANEL (3) Filetype (Normal/Main/Resource/Module)
      wszText = ""
      select case pDoc->ProjectFileType
         case FILETYPE_HEADER:    wszText = L(175,"Header") 
         case FILETYPE_NORMAL:    wszText = L(210,"Normal") 
         case FILETYPE_MODULE:    wszText = L(211,"Module")
         case FILETYPE_MAIN:      wszText = L(212,"Main")
         case FILETYPE_RESOURCE:  wszText = L(213,"Resource")
      end select 
      gSBPanels(3).wszText = wszText
      
      ' PANEL (4) Spacing
      gSBPanels(4).wszText = "Spaces: " & gConfig.TabSize
      
      ' PANEL (5) File Encoding (Ansi, UTF-8, UTF-16)
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
      end select 
      gSBPanels(5).wszText = wszText

      ' PANEL (6) Line Endings
      select case SciExec( hEdit, SCI_GETEOLMODE, 0, 0 )
         case SC_EOL_CRLF: wszText = "CRLF"  '(0)
         case SC_EOL_CR:   wszText = "CR"    '(1)
         case SC_EOL_LF:   wszText = "LF"    '(2)
      end select 
      gSBPanels(6).wszText = wszText
      
      if SciExec( hEdit, SCI_GETREADONLY, 0, 0 ) then
         gSBPanels(1).wszText = L(222,"Read Only")
      end if

   end if
   

   ' MAIN WINDOW CAPTION 
   wszText = iif( gApp.GetDocumentCount, APPNAMESHORT, APPNAME )
   if (gApp.IsProjectActive = true) or (gApp.IsProjectLoading = true) then 
      wszText = wszText & " - [" & gApp.ProjectName & "]"
   end if
   if pDoc then 
      wszText = wszText & " - [" & pDoc->DiskFilename & "]"
      if SciExec( pDoc->hWndActiveScintilla, SCI_GETREADONLY, 0, 0 ) then
         wszText = wszText & " - [" & L(222,"Read Only") & "]"
      end if
   end if
   if wszText <> AfxGetWindowText( HWND_FRMMAIN ) then AfxSetWindowText( HWND_FRMMAIN, wszText )

   ' TAB CONTROL FILENAME ( * modified flag )
   gTTabCtl.SetTabText(-1)  ' this will only repaint if text has changed
   ' repaint the tab to ensure that the unsaved indicator changes
   AfxRedrawWindow( HWND_FRMMAIN_TOPTABS )

   ' Call function to calculate the size/position of the panels and also paint the statusbar
   frmStatusBar_PositionWindows()

   function = 0
end Function


' ========================================================================================
' Open any Session or Project on editor startup
' ========================================================================================
function frmMain_RestoreSessionOrProject() as long
   ' Only restore session if that option is active. The session file may contain
   ' a reference to a project if the last session was a project.
   ' Only restore the session if no files were already opened via the command line.
   if gTTabCtl.GetItemCount = 0 then
       if gConfig.RestoreSession then 
          if AfxFileExists( gConfig.wszLastActiveSession ) then
             gConfig.LoadSessionFile( gConfig.wszLastActiveSession )
          end if   
       end if
   end if
    
   function = 0 
end function


' ========================================================================================
' Process any command line that was passed to the editor
' ========================================================================================
function frmMain_ProcessCommandLine( byval HWnd as HWnd ) as long

   ' The incoming command line may contain a regular file to open or a project file.
   
   ' Command: A space-separated list of all command-line arguments is returned. When the 
   '          command line is parsed for arguments, everything between double quotes in 
   '          the parameter list will be considered as a single parameter, and is returned 
   '          with the double quotes.
   '          A value of zero (0) returns the name of the executable; and values of 
   '          one (1) and greater return each command-line argument.

   ' as of v1.7.4 start to use AfxCommand which is a unicode aware replacement for FB's
   ' built in COMMAND function (that is not unicode compliant).
   
   if len(AfxCommand(1)) = 0 then 
       ' no command line files/project specified bu the user may be wanting
       ' to restore a previous session/project.
       frmMain_RestoreSessionOrProject()
       exit function
   end if
   
   dim wszPath as wstring * MAX_PATH
   dim wszArg  as wstring * MAX_PATH
   dim DataToSend as COPYDATASTRUCT   
   
   if IsIconic(hwnd) then 
      dim WinPla as WINDOWPLACEMENT
      with WinPla
         .Length = sizeof(WinPla)
         .rcNormalPosition.Left   = gConfig.StartupLeft
         .rcNormalPosition.Top    = gConfig.StartupTop
         .rcNormalPosition.Right  = gConfig.StartupRight
         .rcNormalPosition.Bottom = gConfig.StartupBottom
         .showCmd = iif( gConfig.StartupMaximized, SW_MAXIMIZE, SW_SHOWNORMAL )
      end with
      SetWindowPlacement(HWND, @WinPla)
   end if
   SetForegroundWindow(hwnd)
   
   dim as long NumCommandLineFiles = 0 
   dim as long i = 1
   Do
      wszArg = AfxCommand(i)
      if len(wszArg) = 0 then exit do
       
      ' Remove any double quotes from the argument.
      wszPath = AfxStrRemove( wszArg, wchr(34) )
       
      ' if no path exists for the file then add the current folder
      wszPath = AfxStrPathname( "PATH", wszArg )
      if len(wszPath) = 0 then wszArg = AfxGetExePathName & wszArg 
       
      if AfxFileExists(wszArg) then
         NumCommandLineFiles =+ 1
         DataToSend.lpData  = @wszArg
         DataToSend.cbdata  = (len(wszArg)*2) + 1
         SendMessage(hwnd, WM_COPYDATA, len(DataToSend), cast(lParam, @DataToSend))
      end if
       
      i += 1
   Loop

   ' If no files or project was loaded from the command line then check if the user 
   ' had session/project restore option.
   if NumCommandLineFiles = 0 then
       frmMain_RestoreSessionOrProject()
   end if    

   function = 0
end function   



' ========================================================================================
' Determine if the incoming character is a brace character
' ========================================================================================
function frmMain_HighlightWord( _
            byval pDoc as clsDocument ptr, _
            byref text as string _
            ) as long

   if pDoc = 0 then exit function
   
   dim as any ptr pSci = pDoc->GetActiveScintillaPtr()
   if pSci = 0 then exit function

   '// Indicators 0-7 could be in use by a lexer
   '// Indicator 8 is used by Find/Replace
   '// Indicator 9 is used for Brace Highlighting
   '// so we'll use indicator 10 to highlight words.
   dim as long NUM = 10

   '// Remove all uses of our Occurrence indicator
   dim as long nLength = SciMsg( pSci, SCI_GETTEXTLENGTH, 0, 0)
   SciMsg( pSci, SCI_SETINDICATORCURRENT, 10, 0)
   SciMsg( pSci, SCI_INDICATORCLEARRANGE, 0, nLength)

   if len(ltrim(text)) = 0 then exit function

   '// Update indicator appearance
   SciMsg( pSci, SCI_INDICSETSTYLE, NUM, INDIC_STRAIGHTBOX )
   SciMsg( pSci, SCI_INDICSETFORE, NUM, ghEditor.ForeColorOccurrence ) 
   SciMsg( pSci, SCI_INDICSETALPHA, NUM, 80 )   

   '// Search the document
   SciMsg( pSci, SCI_TARGETWHOLEDOCUMENT, 0, 0)
   SciMsg( pSci, SCI_SETSEARCHFLAGS, SCFIND_WHOLEWORD, 0)

   dim as long numfound = 0
	dim as long startPos = 0
   dim as long r
   do 
      r = SciMsg( pSci, SCI_SEARCHINTARGET, len(text), cast(LPARAM, strptr(text)))
      if r = -1 then exit do
      
      numfound = numfound + 1
		
		SciMsg( pSci, SCI_SETINDICATORVALUE, NUM, 0 )
      SciMsg( pSci, SCI_INDICATORFILLRANGE, r, len(text))
      startPos = r + len(text)
      
      ' Adjust the searching positions
      SciMsg( pSci, SCI_SETTARGETSTART, startPos, 0)
      SciMsg( pSci, SCI_SETTARGETEND, nLength, 0)
   loop

   '// if only the current word was found then we don't want any highlighting
	'// because the effect would be that the current word highlights as we type.
	'// Remove all uses of our indicator
   if numfound <= 1 then
		nLength = SciMsg( pSci, SCI_GETTEXTLENGTH, 0, 0)
		SciMsg( pSci, SCI_SETINDICATORCURRENT, 10, 0)
		SciMsg( pSci, SCI_INDICATORCLEARRANGE, 0, nLength)
	end if
	
   function = 0
end function


' ========================================================================================
' Determine if the incoming character is a brace character
' ========================================================================================
function frmMain_IsBrace( byval c as integer ) as boolean

   select case chr(c)
      case "(", ")", "[", "]", "{", "}"
         return true
   end select        

   return false
end function


' ========================================================================================
' Handle Character Autocompletion if that editor option is active
' ========================================================================================
function frmMain_InsertMatchedChars( _
            byval pDoc as clsDocument ptr, _
            byval ch as long _
            ) as boolean
   
   if gConfig.CharacterAutoComplete = 0 then exit function
  
   if pDoc = 0 then exit function
   
   dim as any ptr pSci = pDoc->GetActiveScintillaPtr()
   if pSci = 0 then exit function
   
   dim as long caretPos = SciMsg( pSci, SCI_GETCURRENTPOS, 0, 0)
   dim as long docStart = caretPos = 1
   dim as long docend   = caretPos = SciMsg( pSci, SCI_GETTEXTLENGTH, 0, 0)

   ' Get the styling of the current line to determine if we are in a 
   ' multiline or single line comment block then abort the autoinsert.
   select case SciMsg( pSci, SCI_GETSTYLEAT, caretPos, 0)
      case SCE_B_MULTILINECOMMENT, SCE_B_COMMENT
         exit function
   end select 

   dim as long nCharPrev = iif( docStart, _
                       SciMsg( pSci, SCI_GETCHARAT, caretPos, 0), _
                       SciMsg( pSci, SCI_GETCHARAT, caretPos-2, 0) )
   
   dim as long nCharNext = SciMsg( pSci, SCI_GETCHARAT, caretPos, 0)

   dim as boolean isCharPrevBlank = iif( instr(chr(nCharPrev), any chr(32, 9, 10, 13)), true, false )

   dim as boolean isCharNextBlank = iif( instr(chr(nCharNext), any chr(32, 9, 10, 13)), true, false ) 
   if nCharNext = docend then isCharNextBlank = true

   dim as boolean isEnclosed = false
   if (nCharPrev = asc("(") and nCharNext = asc(")") ) or _
      (nCharPrev = asc("{") and nCharNext = asc("}") ) or _
      (nCharPrev = asc("[") and nCharNext = asc("]") ) then
      isEnclosed = true
   end if                      
                         
   dim as boolean isSpaceEnclosed = false
   if (nCharPrev = asc("(") and isCharNextBlank ) or _
      (isCharPrevBlank and nCharNext = asc(")") ) or _
      (nCharPrev = asc("{") and isCharNextBlank)  or _
      (isCharPrevBlank and nCharNext = asc("}") ) or _
      (nCharPrev = asc("[") and isCharNextBlank)  or _
      (isCharPrevBlank and nCharNext = asc("]") ) then
      isSpaceEnclosed = true
   end if   

   dim as boolean isCharOrString = false
   if (isCharPrevBlank and isCharNextBlank) or isEnclosed or isSpaceEnclosed then
      isCharOrString = true
   end if

   dim as boolean charNextIsCharOrString = iif( instr(chr(nCharNext), any chr(34,92)), true, false )

   dim as wstring * 10 wszText
     
   select case chr(ch)
      case "("
         if (charNextIsCharOrString) then return false
         wszText = ")"
         SciMsg( pSci, SCI_INSERTTEXT, caretPos, cast(LPARAM, @wszText) )
         
      case "{"
         if (charNextIsCharOrString) then return false
         wszText = "}"
         SciMsg( pSci, SCI_INSERTTEXT, caretPos, cast(LPARAM, @wszText) )
         
      case "["
         if (charNextIsCharOrString) then return false
         wszText = "]"
         SciMsg( pSci, SCI_INSERTTEXT, caretPos, cast(LPARAM, @wszText) )
         
      case chr(34)
         '// 0x22 = "
         if (nCharPrev = 34) and (nCharNext = 34) then
            SciMsg( pSci, SCI_DELETERANGE, caretPos, 1)
            SciMsg( pSci, SCI_GOTOPOS, caretPos, 0)
            return false
         end if

         if (isCharOrString) then
            wszText = chr(34)
            SciMsg( pSci, SCI_INSERTTEXT, caretPos, cast(LPARAM, @wszText) )
         end if    
      
   end select 

   function = true
end function


' ========================================================================================
' Process Scintilla Notifications
' ========================================================================================
function Scintilla_OnNotify( _
            byval HWnd as HWnd, _
            byval pNSC as SCNOTIFICATION ptr _
            ) as long

   if pNSC = 0 then exit function
   
   dim pDoc as clsDocument ptr 

   dim as HWND hEdit
   dim as long nLine, nFoldLevel

   select case pNSC->hdr.code
         
      case SCN_UPDATEUI    
         if gApp.SuppressNotify then return true
         pDoc = gApp.GetDocumentPtrByWindow( pNSC->hdr.hwndFrom )
         if pDoc then 
            if pDoc->IsValidScintillaID( pNSC->hdr.idFrom ) then
               frmMain_SetStatusbar
                               
               '// Has the caret changed position
               var caretPos = SciExec( pDoc->hWndActiveScintilla, SCI_GETCURRENTPOS, 0, 0)

               if pDoc->lastCaretPos <> caretPos then
                  pDoc->lastCaretPos = caretPos
                  
                  '// If the xOffset has changed then we need to update the horizontal scrollbar
                  dim as long GetXOffset = SciExec( pDoc->hWndActiveScintilla, SCI_GETXOFFSET, 0, 0 )
                  if pDoc->lastXOffsetPos <> GetXOffset then
                     pDoc->lastXOffsetPos = GetXOffset
                     AfxRedrawWindow( HWND_FRMEDITOR_HSCROLLBAR(0) )
                     AfxRedrawWindow( HWND_FRMEDITOR_HSCROLLBAR(1) )
                  end if
               
                  '// Update any occurrence highlights
                  if gConfig.OccurrenceHighlight then
                     static as string sCurWord, sPrevWord
                     sCurWord = pDoc->GetWord()
                     if sCurWord <> sPrevWord then
                        frmMain_HighlightWord( pDoc, sCurWord )
                        sPrevWord = sCurWord
                     end if
                  end if
                  
                  '// Update any brace highlighting
                  if gConfig.BraceHighlight then
                     var bracePos1 = -1
                     var bracePos2 = -1
   
                     '// Is there a brace to the left or right
                     if (caretPos > 0) and ( frmMain_IsBrace(SciExec(pDoc->hWndActiveScintilla, SCI_GETCHARAT, caretPos - 1, 0))) then
                        bracePos1 = (caretPos - 1)
                     elseif frmMain_IsBrace(SciExec(pDoc->hWndActiveScintilla, SCI_GETCHARAT, caretPos, 0)) then
                        bracePos1 = caretPos
                     end if
                     
                     if bracePos1 >= 0 then
                        '// Find the matching brace
                        bracePos2 = SciExec( pDoc->hWndActiveScintilla, SCI_BRACEMATCH, bracePos1, 0)
                        if bracePos2 = -1 then
                           SciExec( pDoc->hWndActiveScintilla, SCI_INDICSETFORE, 9, ghEditor.ForeColorBracebad )
                           SciExec( pDoc->hWndActiveScintilla, SCI_BRACEBADLIGHT, bracePos1, 0 )
                        else
                           SciExec( pDoc->hWndActiveScintilla, SCI_INDICSETFORE, 9, ghEditor.ForeColorBracegood )
                           SciExec( pDoc->hWndActiveScintilla, SCI_BRACEHIGHLIGHT, bracePos1, bracePos2 )
                        end if
                        SciExec( pDoc->hWndActiveScintilla, SCI_INDICSETOUTLINEALPHA, 9, 127 ) ' transparency of outline
                        SciExec( pDoc->hWndActiveScintilla, SCI_INDICSETALPHA, 9, 127 )        ' transparency of interior      
                     else
                        '// Turn off brace matching
                        SciExec( pDoc->hWndActiveScintilla, SCI_BRACEHIGHLIGHT, -1, -1 )
                     end if
                     
                  else
                     SciExec( pDoc->hWndActiveScintilla, SCI_BRACEHIGHLIGHT, -1, -1 )
                  end if         
               end if
               
            end if
         end if
         return true
         
      
      case SCN_MODIFYATTEMPTRO
         ' Attempting to modify a read-only document.
         MessageBeep(MB_ICONWARNING)
                  
                  
      case SCN_MODIFIED
         if gApp.SuppressNotify then return true

         ' Show line and column. Only do on modification of text otherwise we will have
         ' a huge slowdown when notifications sent for UI updates.
         
         if (pNSC->modificationType and SC_MOD_INSERTTEXT) orelse _
            (pNSC->modificationType and SC_MOD_DELETETEXT) then

            ' Set parsing flag to true. We use a parse flag rather than depending on the dirty
            ' flags because the file could stay dirty for a long time before it gets saved and 
            ' every call to the parser would happen even in cases where only directional 
            ' movement has occured with no text modifications at all.
            
            ' Get the pDoc from the Scintilla HWND because this is the most reliable
            ' method of knowing which scintilla window sent the notification (rather than
            ' using gTTabCtl.GetActiveDocumentPtr. Possible that a control is sending the message
            ' as it is loading and is not yet the active document ptr.
            pDoc = gApp.GetDocumentPtrByWindow( pNSC->hdr.hwndFrom )
            if pDoc then 
               if pDoc->LoadingFromFile = false then
                  pDoc->AutoSaveRequired = true
               end if   
               frmMain_SetStatusbar

					pDoc->bNeedsParsing = true
					if (pNSC->modificationType and SC_MOD_DELETETEXT) then
						' Do a check to see if we have backspaced to a point where there is a
						' period to try to popup an autocomplete or codetip.
						dim as long nPos, nChar
						hEdit = pDoc->hWndActiveScintilla
						nPos  = SciExec( hEdit, SCI_GETCURRENTPOS, 0, 0 ) - 1
						nChar = SciExec( hEdit, SCI_GETCHARAT, nPos, 0 )
						select case chr(nChar)
							case ".", ">"    '  Show autocomplete list for TYPE variables
								' dot "." or ">"  part of a pointer
								' Need to PostMessage in order to give time for notification to complete. The "."
								' is also used as a character that selects and closes the popup list. If no
								' PostMessage then the "." will automatically select the first entry in the list
								' and close.
                        pDoc->AutoCTriggerStartPos = nPos
								pDoc->AutoCompleteType = AUTOCOMPLETE_NONE
								PostMessage( HWND_FRMMAIN, MSG_USER_SHOWAUTOCOMPLETE, nChar, 0 )
						end select 
					end if
            end if
         end if
         return true
         
      case SCN_MARGINCLICK
         ' Folder margin
         pDoc = gTTabCtl.GetActiveDocumentPtr()
         if pDoc then 
            hEdit      = pDoc->hWndActiveScintilla
            nLine      = SciExec(hEdit, SCI_LINEFROMPOSITION, pNSC->position, 0)
            nFoldLevel = SciExec(hEdit, SCI_GETFOLDLEVEL, nLine, 0)
            select case pNSC->margin 
               case 1   ' left margin (bookmarks)
                  pDoc->ToggleBookmark(nLine)
                  LoadBookmarksFiles
                  AfxRedrawWindow( HWND_FRMBOOKMARKS )
               case 2   ' fold margin
                  ' if is the head line...
                  if (nFoldLevel and SC_FOLDLEVELHEADERFLAG) <> 0 then
                     pDoc->FoldToggle(nLine) 
                  end if   
            end select 
         end if           
                     
       
      case SCN_AUTOCCANCELLED
         ' Destroy the popup information window
         pDoc = gTTabCtl.GetActiveDocumentPtr()
         if pDoc then 
            ' If the last typed character was a BACKSPACE then do NOT reset the 
            ' autocomplete type otherwise the autocomplete popup will not continue
            ' to popup on subsequent characters being added.
            if pDoc->LastCharTyped = VK_BACK then
               ' backspace pressed... do nothing, unless we are backspacing past where
               ' the popup trigger occurred.
               dim as long curPos = SciExec(pDoc->hWndActiveScintilla, SCI_GETCURRENTPOS, 0, 0) 
               if curPos <= pDoc->AutoCTriggerStartPos then
                  pDoc->AutoCompleteType = AUTOCOMPLETE_NONE
               end if
            else
               pDoc->AutoCompleteType = AUTOCOMPLETE_NONE
            end if   
         end if

         
      case SCN_AUTOCSELECTION
         ' A selection occured from the AutoComplete listbox. We do the insertion
         ' ourselves because the startpos of the word is not the same as the start position
         ' used by the autocomplete listbox (because we are doing incremental searches).
         pDoc = gTTabCtl.GetActiveDocumentPtr()
         
         if pDoc then 
            
            hEdit = pDoc->hWndActiveScintilla
            
            ' Get the position where the listbox was opened.
            dim as long nStartPos = pDoc->AutoCStartPos 
            dim as long nPos = SciExec( hEdit, SCI_GETCURRENTPOS, 0, 0)
            
            ' Get the match word that existed when the listbox was displayed
            dim as long nLenMatchWord = len(pDoc->sMatchWord)
            
            ' Set the word that was selected in the listbox
            dim as string sText = *cast(zstring ptr, pNSC->lpText)
            
            if ucase(sText) <> ucase(pDoc->sMatchWord) then
               SciExec( hEdit, SCI_SETSEL, nStartPos-nLenMatchWord, nPos)
               SciExec( hEdit, SCI_REPLACESEL, 0, cast(LPARAM, strptr(sText)))
               nPos = SciExec( hEdit, SCI_GETCURRENTPOS, 0, 0)
            end if
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
            
            ' if a CodeTip was displayed prior to the autocomplete popup list then
            ' redisplay that CodeTip now.
            ShowCodetip( pDoc )
            
            return true   
         end if
 

      case SCN_AUTOCCHARDELETED    
         if gConfig.CodeTips then
            ' User deleted a character while autocompletion list was active. Display
            ' new contents of list because the underlying word being typed as changed.
            ShowAutocompleteList(SCN_AUTOCCHARDELETED)
         end if
                                 
      
      case SCN_CHARADDED
         if gApp.SuppressNotify then exit function
         pDoc = gTTabCtl.GetActiveDocumentPtr()
         if pDoc = 0 then exit function
         hEdit = pDoc->hWndActiveScintilla
       
			pDoc->bNeedsParsing = true

         ' Attempt to do a Character Autocompletion if that editor option is active.
         frmMain_InsertMatchedChars( pDoc, pNSC->ch )
         
         select case chr(pNSC->ch)
         case ".", ">"    '  Show autocomplete list for TYPE variables
            ' dot "." or ">"  part of a pointer
            ' Need to PostMessage in order to give time for notification to complete. The "."
            ' is also used as a character that selects and closes the popup list. if no
            ' PostMessage then the "." will automatically select the first entry in the list
            ' and close.
            pDoc->AutoCTriggerStartPos = SciExec(hEdit, SCI_GETCURRENTPOS, 0, 0)
            pDoc->AutoCompleteType = AUTOCOMPLETE_NONE
            if PostMessage(HWND_FRMMAIN, MSG_USER_SHOWAUTOCOMPLETE, pNSC->ch, 0) then exit function
          
         case ")"      ' Close and active code tip 
            pDoc->AutoCompleteType = AUTOCOMPLETE_NONE
            SciExec( hEdit, SCI_CALLTIPCANCEL, 0, 0 )
               
         case "(", ","      ' Show code tip 
            pDoc->AutoCompleteType = AUTOCOMPLETE_NONE
            if PostMessage(HWND_FRMMAIN, MSG_USER_SHOWAUTOCOMPLETE, pNSC->ch, 0) then exit function

         case chr(13)  ' ENTER KEY PRESSED
            pDoc->AutoCompleteType = AUTOCOMPLETE_NONE
            AttemptAutoInsert()
  
         case chr(32)  ' Space key pressed (Insert AutoComplete constructs if applicable)
            pDoc->AutoCompleteType = AUTOCOMPLETE_NONE
            if ShowAutocompleteList() then exit function

         case else   ' all other letters
            if ShowAutocompleteList() then exit function

         end select 
         
   end select 

   function = 0
end Function


' ========================================================================================
' Set focus to currently active Scintilla window
' ========================================================================================
function frmMain_SetFocusToCurrentCodeWindow() as long
   ' Post a message to the main form CUSTOM handler that will
   ' set focus to the currently active Scintilla code window. We
   ' use PostMessage to ensure that all all other windows 
   ' messages are finished processing.
   PostMessage( HWND_FRMMAIN, MSG_USER_SETFOCUS, 0, 0 )
   function = 0
end Function


' ========================================================================================
' Process WM_PAINT message for window/dialog: frmMain
' ========================================================================================
function frmMain_OnPaint( byval HWnd as HWnd ) as LRESULT
            
   dim pWindow as CWindow ptr = AfxCWindowPtr( HWND_FRMMAIN )
   if pWindow = 0 then exit function

   dim as PAINTSTRUCT ps
   dim as HPEN hPen
   dim as HDC hDc
   dim as RECT rc
   
   hDC = BeginPaint(hWnd, @ps)

   SaveDC hDC
   
   FillRect( hDC, @ps.rcPaint, ghBrushMainBackground )

   ' Draw any horizontal splitter between the two edit windows
   dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr
   if pDoc then
      if pDoc->bEditorIsSplit then
         if IsWindowVisible( pDoc->hWindow(1) ) then 
            GetWindowRect( pDoc->hWindow(1), @rc )
            rc.top = rc.bottom
            if IsWindowVisible( HWND_FRMEDITOR_HSCROLLBAR(1) ) then
               rc.top = rc.top + AfxGetWindowHeight( HWND_FRMEDITOR_HSCROLLBAR(1) )
            end if   
            rc.bottom = rc.top + pWindow->ScaleY(SPLITSIZE)
            rc.right = rc.right + pWindow->ScaleX(SCROLLBAR_WIDTH_EDITOR)
            MapWindowPoints( HWND_DESKTOP, HWND_FRMMAIN, cast(POINT ptr, @rc), 2 )
            pDoc->rcSplitButton = rc
            dim as HBRUSH hBackBrush = CreateSolidBrush( ghEditor.Divider )
            FillRect( hDC, @rc, hBackBrush )
            DeleteObject( hBackBrush )
         end if
      end if
   end if
   
   RestoreDC hDC, -1 
   
   EndPaint hWnd, @ps
   
   function = 0
end Function

' ========================================================================================
' Hide the vertical & horizontal editor scrollbars
' ========================================================================================
function frmMain_HideScrollBars() as LRESULT
      ShowWindow( HWND_FRMEDITOR_HSCROLLBAR(0), SW_HIDE )
      ShowWindow( HWND_FRMEDITOR_HSCROLLBAR(1), SW_HIDE )
      ShowWindow( HWND_FRMEDITOR_VSCROLLBAR(0), SW_HIDE )
      ShowWindow( HWND_FRMEDITOR_VSCROLLBAR(1), SW_HIDE )
   function = 0
end function

' ========================================================================================
' Position all child windows. Called manually and/or by WM_SIZE
' ========================================================================================
function frmMain_PositionWindows() as LRESULT
   dim as HWnd hEdit
   dim as long nHeightTabControl, nLeft, nTop
   dim as Rect rc

   dim pWindow as CWindow ptr = AfxCWindowPtr( HWND_FRMMAIN )
   if pWindow = 0 then exit function
   
   dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr()
   
   ' Get the entire client area
   GetClientRect( HWND_FRMMAIN, @rc )
  
   dim as long nHeightMenuBar   = AfxGetWindowHeight( HWND_FRMMAIN_MENUBAR )
   dim as long nHeightStatusBar = AfxGetWindowHeight( HWND_FRMMAIN_STATUSBAR )
   dim as long nHeightOutput    = AfxGetWindowHeight( HWND_FRMOUTPUT )
   dim as long nWidthPanel      = AfxGetWindowWidth( HWND_FRMPANEL )

   ' Set the top menu menubar into place 
   SetWindowPos( HWND_FRMMAIN_MENUBAR, 0, _
                 0, 0, rc.right - rc.left, nHeightMenuBar, _
                 SWP_NOZORDER or SWP_SHOWWINDOW ) 

   ' Set the statusbar into place 
   SetWindowPos( HWND_FRMMAIN_STATUSBAR, 0, _
                 0, rc.bottom - nHeightStatusBar, rc.right - rc.left, nHeightStatusBar, _
                 SWP_NOZORDER or SWP_SHOWWINDOW ) 
   
   nLeft = rc.Left 
   nTop = rc.top + nHeightMenuBar
   
   ' Set the Panel pane into place if applicable
   if IsWindowVisible(HWND_FRMPANEL) then
      SetWindowPos( HWND_FRMPANEL, 0, _
                    nLeft, nTop, _
                    nWidthPanel, _
                    rc.Bottom - nHeightStatusBar - nHeightMenuBar, _
                    SWP_NOZORDER or SWP_SHOWWINDOW ) 
      frmPanel_PositionWindows
      nLeft = nWidthPanel
   Else
      nWidthPanel = 0
   end if


   ' if items exist in the top tabcontrol then show the tab control and account for its height
   if gTTabCtl.GetItemCount = 0 then
      ShowWindow( HWND_FRMMAIN_TOPTABS, SW_HIDE )
      nHeightTabControl = 0
      ' If Find/Replace is open then close it
      if IsWindowVisible( HWND_FRMFINDREPLACE ) then
         DestroyWindow( HWND_FRMFINDREPLACE )
      end if   
   else
      nHeightTabControl = AfxGetWindowHeight(HWND_FRMMAIN_TOPTABS)
      SetWindowPos( HWND_FRMMAIN_TOPTABS, 0, _
                    nLeft, nTop, _
                    rc.Right - nWidthPanel, nHeightTabControl, _
                    SWP_SHOWWINDOW or SWP_NOZORDER )
      frmTopTabs_PositionWindows()
   end if
   

   ' Set the Output pane into place if applicable
   if IsWindowVisible(HWND_FRMOUTPUT) then
      SetWindowPos( HWND_FRMOUTPUT, 0, _
                    nLeft, rc.Bottom - nHeightStatusbar - nHeightOutput, _
                    rc.Right - nLeft, _
                    nHeightOutput, _
                    SWP_NOZORDER or SWP_SHOWWINDOW ) 
   Else
      nHeightOutput = 0
   end if

   if pDoc = 0 then frmMain_HideScrollBars() 
   
   if pDoc then 
      ' Position the Scintilla child edit windows
      dim pWindow as CWindow ptr = AfxCWindowPtr(HWND_FRMMAIN)
      dim as RECT rcDoc
      ' Reduce the height of the client area by the size of the statusbar and menubar.
      rcDoc.left = nLeft
      rcDoc.top = nTop + nHeightTabControl 
      rcDoc.right = rc.Right
      rcDoc.bottom = rc.bottom - nHeightStatusBar - nHeightOutput
      
      dim as long nDesignCodeTabHeight = 0
      if pDoc->IsDesigner then
         nDesignCodeTabHeight = pWindow->ScaleY(DESIGNTABS_HEIGHT)
      end if
      
      dim as long lStyle = SWP_SHOWWINDOW or SWP_NOZORDER or SWP_NOOWNERZORDER or _
                           SWP_NOACTIVATE or SWP_NOCOPYBITS
      
      if pDoc->IsDesigner then
         SetWindowPos( HWND_FRMMAIN_DESIGNTABS, 0, _
                       rcDoc.left, _
                       rcDoc.bottom - nDesignCodeTabHeight, _
                       rcDoc.right, _
                       nDesignCodeTabHeight, _
                       lStyle )
         AfxRedrawWindow( HWND_FRMMAIN_DESIGNTABS )
      else   
         ShowWindow( HWND_FRMMAIN_DESIGNTABS, SW_HIDE )
      end if
      
      dim as long nShowVScrollBars = iif(pDoc->DesignTabsCurSel = 0, SW_HIDE, SW_SHOW)
      ' Only hide vertical scrollbars because hiding the horizontal could
      ' result in endless loop during mouseover test in modMsgPump.
      ShowWindow( HWND_FRMEDITOR_VSCROLLBAR(0), nShowVScrollBars )
      ShowWindow( HWND_FRMEDITOR_VSCROLLBAR(1), nShowVScrollBars )
      
      if (pDoc->IsDesigner) andalso (IsDesignerView(pDoc)) then
         ' if the DesignView has been set to show the visual designer rather than the code
         ' window then ensure that DesignMain is shown.
         SetWindowPos( pDoc->hWndDesigner, 0, _
                       rcDoc.left, rcDoc.top, _
                       rcDoc.right - rcDoc.left, _
                       rcDoc.bottom - rcDoc.top - nDesignCodeTabHeight, _
                       lStyle )

         ' Display the ToolBox/PropertyList
         frmVDToolbox_Show( HWND_FRMMAIN, SW_SHOW )
         DisplayPropertyList( pDoc )
         ShowWindow( HWND_FRMVDTOOLBOX, SW_SHOWNORMAL )
         ShowWindow( pDoc->hWindow(0), SW_HIDE )
         ShowWindow( pDoc->hWindow(1), SW_HIDE )
      else
      
         ShowWindow( pDoc->hWndDesigner, SW_HIDE )
         ShowWindow( HWND_FRMVDTOOLBOX, SW_HIDE )

         dim as long nSplitSize = pWindow->ScaleY(SPLITSIZE)
         dim as long iVScrollbarWidth = AfxGetWindowWidth(HWND_FRMEDITOR_VSCROLLBAR(0)) 
         dim as long iHScrollbarHeight 

         ' TOP EDIT WINDOW (optional)
         if pDoc->bEditorIsSplit = false then 
            ShowWindow( pDoc->hWindow(1), SW_HIDE )      
            ShowWindow( HWND_FRMEDITOR_VSCROLLBAR(1), SW_HIDE )
            ShowWindow( HWND_FRMEDITOR_HSCROLLBAR(1), SW_HIDE )
            nSplitSize = 0              
         Else
            ' Position the TOP split editor window
            iHScrollbarHeight = AfxGetWindowHeight(HWND_FRMEDITOR_HSCROLLBAR(1)) 
            if IsWindowVisible(HWND_FRMEDITOR_HSCROLLBAR(1)) = 0 then iHScrollbarHeight = 0
            SetWindowPos( pDoc->hWindow(1), 0, _
                          rcDoc.left, rcDoc.top, _
                          rcDoc.right - rcDoc.left - iVScrollbarWidth, _
                          pDoc->SplitY - rcDoc.top - iHScrollbarHeight, _ 
                          lStyle )

            ' Position the TOP view horizontal scroll bar
            if IsWindowVisible(HWND_FRMEDITOR_HSCROLLBAR(1)) then
               SetWindowPos( HWND_FRMEDITOR_HSCROLLBAR(1), 0, _ 
                             rcDoc.left, pDoc->SplitY - iHScrollbarHeight, _
                             rcDoc.right - rcDoc.left - iVScrollbarWidth, iHScrollbarHeight, _
                             SWP_NOZORDER )
            end if
            
            ' Position the TOP view vertical scroll bar
            ' Determine if the VScrollBar should be displayed
            if pDoc->GetLineCount > pDoc->LinesPerPage(1) then
               SetWindowPos( HWND_FRMEDITOR_VSCROLLBAR(1), 0, _ 
                             rcDoc.right - iVScrollbarWidth, rcDoc.top, _
                             iVScrollbarWidth, pDoc->SplitY - rcDoc.top, _ 
                             lStyle )
            end if
         end if
         
         ' BOTTOM (MAIN) EDIT WINDOW
         dim as long nTop = Max( rcDoc.top, pDoc->SplitY + nSplitSize )
         iHScrollbarHeight = AfxGetWindowHeight(HWND_FRMEDITOR_HSCROLLBAR(0)) 
         if IsWindowVisible(HWND_FRMEDITOR_HSCROLLBAR(0)) = 0 then iHScrollbarHeight = 0

         SetWindowPos( pDoc->hWindow(0), 0, _
                       rcDoc.left, nTop, _
                       rcDoc.right - rcDoc.left - iVScrollbarWidth, _
                       rcDoc.bottom - nTop - nDesignCodeTabHeight - iHScrollbarHeight, _
                       lStyle )

         ' Position the MAIN view horizontal scroll bar
         if IsWindowVisible(HWND_FRMEDITOR_HSCROLLBAR(0)) then
            SetWindowPos( HWND_FRMEDITOR_HSCROLLBAR(0), 0, _ 
                          rcDoc.left, rcDoc.bottom - nDesignCodeTabHeight - iHScrollbarHeight, _
                          rcDoc.right - rcDoc.left - iVScrollbarWidth, iHScrollbarHeight, _
                          SWP_NOZORDER )
         end if

         ' Position the MAIN view vertical scroll bar
         ' Determine if the VScrollBar should be displayed
         if pDoc->GetLineCount > pDoc->LinesPerPage(0) then
            SetWindowPos( HWND_FRMEDITOR_VSCROLLBAR(0), 0, _ 
                          rcDoc.right - iVScrollbarWidth, nTop + SPLITSIZE, _
                          iVScrollbarWidth, _
                          rcDoc.bottom - nTop - nDesignCodeTabHeight, _ 
                          lStyle )
         else
            ShowWindow( HWND_FRMEDITOR_VSCROLLBAR(0), SW_HIDE )
         end if
            
         ' Update editor vertical scrollbars
         frmEditorVScroll_calcVThumbRect( pDoc )
         
         ' if the FIND/REPLACE window is open then ensure that it is positioned correctly
         ' especially important if the Main window is resized.
         frmFindReplace_PositionWindows
         
      end if
   end if
   
   ' Ensure that the correct notes are shown
   frmOutput_ShowNotes
   
   function = 0
end function


' ========================================================================================
' Attempt to open specified project. 
' ========================================================================================
function frmMain_OpenProjectSafely( _
            byval HWnd as HWnd, _
            byref wszProjectFileName as const wstring _
            ) as boolean

   ' if a Project is not active then we need to save the current non-project notes
   ' when this file is closed. It is possible that this file is being closed and
   ' a project is being opened.
   if gApp.IsProjectActive = false then
      gApp.NonProjectNotes = AfxGetWindowText(GetDlgItem(HWND_FRMOUTPUT, IDC_FRMOUTPUT_TXTNOTES))
      gConfig.SaveConfigFile
   end if

   ' if a project is open then close any files that may be open in the editor. No need to
   ' do this if a new project is being created/opened because that function has already
   ' performed this operation. Doing so again will cause the Recent Files/Projects panel
   ' to flash on the screen.
   if gApp.IsNewProjectFlag = false then
      if gApp.IsProjectActive then 
         if OnCommand_ProjectClose(hwnd) = false then exit function
      else   
         if OnCommand_FileClose(HWnd, EFC_CLOSEALL) = false then exit function
      end if
      ' Clear any previous info from the Output windows
      frmOutput_ResetAllControls()
   end if
         
   ' Open the project
   if gConfig.ProjectLoadFromFile(wszProjectFileName) then
      ' Update the most recently used project list 
      UpdateMRUProjectList(wszProjectFileName)
   end if
   gApp.IsProjectActive = true
   
   ' Position all of the controls into place
   frmPanel_PositionWindows
   frmMain_PositionWindows

   function = true
end function


' ========================================================================================
' Attempt to open specified file. if it exists then position to Tab if applicable
' ========================================================================================
function frmMain_OpenFileSafely( _
            byval HWnd        as HWnd, _
            byval bIsNewFile  as boolean, _ 
            byval bIsTemplate as boolean, _
            byval bShowInTab  as boolean, _
            byval bIsInclude  as boolean, _
            byref wszName     as wstring, _
            byval pDocIn      as clsDocument ptr, _
            byval bIsDesigner as boolean = false, _
            byval wszFileType as CWSTR = FILETYPE_UNDEFINED _
            ) as clsDocument ptr
     
   dim as long iTab = -1
   dim pDoc as clsDocument ptr 

   ' This function opens/creates various types of files depending on the situation. 
   ' - New documents
   ' - Open document from disk (after WinFBE has been loaded)
   ' - Display document that is already loaded and has valid Scintilla loaded control.
   ' - Display document that is already loaded but does not have a valid Scintilla loaded control.

   ' if the incoming pDocIn is null then we need to create a new pDoc and add it to the collection. 
   if pDocIn = 0 then   
      ' Create a new pDoc 
      pDoc = gApp.AddNewDocument() 
      pDoc->IsDesigner = bIsDesigner 
      if bIsNewFile then
         pDoc->CreateCodeWindow( HWnd, true )  ' Create the new Scintilla window
      else
         wszName = OnCommand_FileAutoSaveFileCheck( wszName )
         pDoc->CreateCodeWindow( HWnd, false, bIsTemplate, wszName )
         pDoc->bNeedsParsing = true
         pDoc->ParseDocument()
      end if
      pDoc->ProjectFileType = wszFileType
      if gApp.IsProjectLoading = false then 
         LoadExplorerFiles()
         LoadFunctionsFiles()
      end if
   else   
      pDoc = pDocIn   
   end if

   ' Set the default build configuration for this document. If no IsDefault option
   ' has been checked then we simply use whatever the current selected build is.
   if pDoc->DocumentBuild = "" then 
      pDoc->DocumentBuild = frmBuildConfig_GetDefaultBuildGUID()
      if pDoc->DocumentBuild = "" then 
         pDoc->DocumentBuild = frmBuildConfig_GetSelectedBuildGUID()
      end if
   end if      
   
   if bShowInTab then
      if gApp.IsProjectLoading = false then 
         ' if the document is already open and loaded then simply switch to
         ' that document in the top tabcontrol, otherwise load the Explorer 
         ' files list, create a new tab, and switch to it.
         iTab = gTTabCtl.GetTabIndexFromFilename( pDoc->DiskFilename )
         if iTab = -1 then 
            LoadExplorerFiles()
            iTab = gTTabCtl.AddTab( pDoc )  ' Add the new document to the top tabcontrol
         end if   
         gTTabCtl.SetFocusTab(iTab)
      end if
   end if

   ' if a Menu, ToolBar or StatusBar exists on the Form then ensure that it
   ' resizes to the new Form width.
   if pDoc->IsDesigner then
      frmMenuEditor_CreateFakeMainMenu(pDoc)
      frmToolBarEditor_CreateFakeToolBar(pDoc)
      frmStatusBarEditor_CreateFakeStatusBar(pDoc)
   end if

   if gApp.IsProjectLoading = false then 
      frmMain_PositionWindows
      frmMain_SetFocusToCurrentCodeWindow
   end if
   
   ' Post a custom message in order to check if this now opened file needs
   ' to be upgraded to new 3.0.2 form format.
   if pDoc->IsDesigner then
      postmessage( HWND_FRMMAIN, MSG_USER_UPGRADE302FORM, 0, cast(LPARAM, pDoc) )
   end if

   function = pDoc
end Function


' ========================================================================================
' Process WM_CREATE message for window/dialog: frmMain
' ========================================================================================
function frmMain_OnCreate( _
            byval HWnd as HWnd, _
            byval lpCreateStructPtr as LPCREATESTRUCT _
            ) as boolean
   ' Enable drag and drop files
   DragAcceptFiles HWnd, CTRUE
   
   '  Message cracker macro expects a true to be returned for a successful
   '  OnCreate handler even though returning -1 from a standard WM_CREATE
   '  call would stop creating the window. This is just one of those Windows
   '  inconsistencies.
   return true
end Function


' ========================================================================================
' Process WM_SIZE message for window/dialog: frmMain
' ========================================================================================
function frmMain_OnSize( _
            byval HWnd as HWnd, _
            byval state as UINT, _
            byval cx as long, _
            byval cy as long _
            ) as LRESULT
   if state <> SIZE_MINIMIZED then
      ' Position all of the child windows
      frmMain_PositionWindows
   end if
   function = 0
end Function


' ========================================================================================
' Process WM_NOTIFY message for window/dialog: frmMain
' ========================================================================================
function frmMain_OnNotify( _
            byval HWnd as HWnd, _
            byval id as long, _
            byval pNMHDR as NMHDR ptr _
            ) as LRESULT

   dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr()
   if pDoc then 
      if pDoc->IsValidScintillaID(id) then
         ' Process Scintilla control notification messages. First, check to see if the notifications
         ' have been suppressed by a bulk modification activity that should only update the screen
         ' at the end of its process (eg. moving large numbers of selected lines up or down).
         if gApp.SuppressNotify then exit function
         Scintilla_OnNotify HWnd, cast(SCNOTIFICATION ptr, pNMHDR)
         exit function
      end if
   end if
        

   select case pNMHDR->code 

      case TCN_SELCHANGING
         select case id 
            case IDC_DESIGNTABCTRL 
               frmMain_PositionWindows
         end select 
         
      case TCN_SELCHANGE
         select case id 
            case IDC_DESIGNTABCTRL 
               frmMain_PositionWindows
               PostMessage(hwnd, MSG_USER_GENERATECODE, 0, 0)
         end select   

   end select 

   function = 0
end Function


' ========================================================================================
' Process WM_ACTIVATEAPP message for window/dialog: frmMain
' ========================================================================================
function frmMain_OnActivateApp( _
            byval HWnd as HWnd, _
            byval fActivate as boolean, _
            byval dwThreadId as DWORD _
            ) as LRESULT

   if gApp.PreventActivateApp then exit function

   dim pDoc as clsDocument ptr
   
   ' if the application is losing focus then validate any currently editing property
   ' in the ToolBox PropertyList and destroy any popup topmenus.
   if fActivate = false then
      killAllPopupMenus()
      pDoc = gTTabCtl.GetActiveDocumentPtr()
      if pDoc then DisplayPropertyList(pDoc)
      exit function
   end if
      
   static bInActivateApp as boolean
   if bInActivateApp then exit function
   
   ' if the application is gaining focus then determine if any of the loaded
   ' documents have been modified by an external application. if yes, then ask
   ' the user if wish to reload.
   if fActivate = true then
      if gApp.GetDocumentCount = 0 then exit function
      
      ' Search all loaded documents
      dim as FILETIME ft
      dim as long i, idx, nDocumentCount
      dim as double nSerial
      dim as string sText
      
      bInActivateApp = true
      
      ' Load all pDoc into array and then process. For each pDoc if the user chooses
      ' not to reload or wants to bypass, then set the pDoc to zero until eventually
      ' all pDocs are processed.
      dim pDocs( gApp.GetDocumentCount - 1) as clsDocument ptr
      idx = 0
      pDoc = gApp.pDocList
      do until pDoc = 0
         if AfxFileExists(pDoc->DiskFilename) then
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
         if pDoc->IsNewFlag then 
            pDocs(idx) = 0: continue do
         end if

         ' Has the external file been deleted or is now not available
         if AfxFileExists(pDoc->DiskFilename) = false then
            if pDoc->DeletedButKeep = false then
                ' Ensure that the file is open and displayed
                frmMain_OpenFileSafely(HWND_FRMMAIN, _
                                        false, _    ' bIsNewFile
                                        false, _    ' bIsTemplate
                                        true,  _    ' bShowInTab
                                        false, _    ' bIsInclude
                                        "", _       ' wszName
                                        pDoc )      ' pDocIn
                if MessageBox( HWND_FRMMAIN, _
                                pDoc->DiskFilename & vbCrLf & _ 
                                L(286, "This document has been deleted or is no longer available. Do you wish to keep the file open in the editor?"), _
                                L(267, "File Changed"), MB_ICONQUESTION or MB_YESNO) = IDYES then
                    ' Keep the file open and simply mark it as dirty so it can be prompted to be saved.
                    pDoc->UserModified = true                
                else
                    ' No. Close the file.
                    if gApp.IsProjectActive then 
                        OnCommand_ProjectRemove( 0, pDoc )  ' the first param (id) is not used.
                    else
                        if OnCommand_FileClose( HWND_FRMMAIN, EFC_CLOSECURRENT ) = false then continue do
                    end if    
                    pDocs(idx) = 0
                    continue do
                end if
            end if
            pDoc->DeletedButKeep = true

         else
            ' Compare the disk file date time to the value currently
            ' stored in document class.
            ft = AfxGetFileLastWriteTime(pDoc->DiskFilename)
            if AfxFileTimeToVariantTime(ft) <> AfxFileTimeToVariantTime(pDoc->DateFileTime) then
                OpenSelectedDocument( pDoc->DiskFilename, "" )
                dim as long res = MessageBox( HWND_FRMMAIN, _
                                pDoc->DiskFilename & vbCrLf & _ 
                                L(266, "File was changed by another application. Reload it?"), _
                                L(267, "File Changed"), MB_ICONQUESTION or MB_YESNOCANCEL) 
                    
                if res = IDCANCEL then
                    bInActivateApp = true
                    exit function
                end if
                
                if res = IDYES then
                    dim as CWSTR wszFilename = pDoc->DiskFilename
                    dim as CWSTR wszFileType = pDoc->ProjectFileType
                    
                    ' Remove the current file and open the changed file.
                    if gApp.IsProjectActive then 
                        OnCommand_ProjectRemove( 0, pDoc )  ' the first param (id) is not used.
                    else
                        if OnCommand_FileClose( HWND_FRMMAIN, EFC_CLOSECURRENT ) = false then continue do
                    end if    
                    
                    ' Ensure that the file is open and displayed
                    dim as clsDocument ptr pDocLoad = _
                    frmMain_OpenFileSafely( HWND_FRMMAIN, _
                                            false, _    ' bIsNewFile
                                            false, _    ' bIsTemplate
                                            true,  _    ' bShowInTab
                                            false, _    ' bIsInclude
                                            wszFilename, _ ' wszName
                                            0 )         ' pDocIn
                    
                    pDocLoad->ProjectFileType = wszFileType
                end if                
            end if
            pDoc->DateFileTime = AfxGetFileLastWriteTime( pDoc->DiskFilename )
         end if

         pDocs(idx) = 0
      loop      
             
      bInActivateApp = false
            
   end if
   function = 0
end Function


' ========================================================================================
' Process WM_CONTEXTMENU message for window/dialog: frmMain
' ========================================================================================
function frmMain_OnContextMenu( _
            byval HWnd as HWnd, _
            byval hwndContext as HWnd, _
            byval xPos as long, _  
            byval yPos as long _
            ) as LRESULT
   dim hPopupMenu as HMENU 
   dim pt         as POINT 
   
   dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr()
   if pDoc = 0 then return 0

   SetFocus hWndContext
   pt.x = xPos
   pt.y = yPos
   ScreenToClient hWndContext, @pt
   if IsDesignerView(pDoc) then
      ' Right-click popup menu for visual design form is 
      ' handled in the HandleDesignerRButtonDown function.
   else
      hPopupMenu = CreateScintillaContextMenu()
   end if
   GetCursorPos @pt
   TrackPopupMenu hPopupMenu, 0, pt.x, pt.y, 0, HWnd, 0
   DestroyMenu hPopupMenu

   function = 0
end Function


' ========================================================================================
' Process WM_DROPFILES message for window/dialog: frmMain
' ========================================================================================
function frmMain_OnDropFiles( _
            byval HWnd as HWnd, _
            byval hDrop as HDROP _
            ) as LRESULT

   ' Get the number of dropped files
   dim as long nCount = DragQueryFile(hDrop, &HFFFFFFFF, Null, 0)
   if nCount = 0 then exit function
   
   dim as long i, nLen
   dim wszPath as wstring * MAX_PATH
   dim wFileExt as wstring * MAX_PATH
   
   For i = 0 To nCount - 1
      nLen = DragQueryFile(hDrop, i, @wszPath, MAX_PATH)
      ' Make sure it's a file, not a folder
      dim fd as WIN32_FIND_DATAW
      dim hFind as HANDLE = FindFirstFileW(@wszPath, @fd)
      if hFind <> INVALID_HANDLE_VALUE then
         FindClose hFind
         if (fd.dwFileAttributes and FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY then
            ' Determine what type of file is being dropped. if it is a project file .wfbe then
            ' use the OpenProject routines, otherwise open it as a regular file.
            wFileExt = AfxStrPathname( "EXTN", wszPath)
            wFileExt = Ucase(wFileExt)
            if wFileExt = ".WFBE" then
               frmMain_OpenProjectSafely(HWND_FRMMAIN, wszPath) 
            Else   
               ' Test to see if the file is already loaded in the editor. if it is, then
               ' bypass loading it again thereby creating multiple ghost instances.
               dim pDoc as clsDocument ptr
               dim pDocIn as clsDocument ptr
               pDoc = gApp.GetDocumentPtrByFilename(wszPath) 
               if pDoc then
                  if pDoc->GetActiveScintillaPtr = 0 then pDocIn = pDoc
               end if   
               if (pDoc = 0) orelse (pDocIn <> 0) then 
                  pDoc = frmMain_OpenFileSafely(HWnd, _
                                          false, _    ' bIsNewFile
                                          false, _    ' bIsTemplate
                                          true,  _    ' bShowInTab
                                          false, _    ' bIsInclude
                                          wszPath, _  ' wszName
                                          pDocIn, _   ' pDocIn
                                          IsFormFilename(wszPath) _
                                          )
                                          
                  ' Give this document a default project type depending on its file extension
                  if (pDoc->IsNewFlag = false) andalso (pDoc->ProjectFileType = FILETYPE_UNDEFINED) then
                     if ( gApp.IsProjectActive = true ) orelse ( gApp.IsProjectLoading = true ) then
                        if pDoc->IsDesigner then 
                           pDoc->ProjectFileType = FILETYPE_NORMAL 
                        else
                           gApp.ProjectSetFileType( pDoc, pDoc->ProjectFileType )
                        end if
                     end if   
                  end if

               end if
            end if                           
         end if
      end if
   Next
   
   DragFinish hDrop

   LoadExplorerFiles()
   LoadFunctionsFiles()

   function = 0
end Function


' ========================================================================================
' Process WM_CLOSE message for window/dialog: frmMain
' ========================================================================================
function frmMain_OnClose( byval HWnd as HWnd ) as LRESULT

   ' if configuration option to confirm closing editor is active then ask now.
   if gConfig.AskExit then
      if MessageBox( HWND_FRMMAIN, L(275,"Are you sure you want to exit?"), L(276,"Confirm"), _
                  MB_YESNOCANCEL or MB_ICONQUESTION or MB_DEFBUTTON1 ) <> IDYES then
         return true
      end if            
   end if
   
   ' Set global shutdown flag that will bypass removing nodes from project
   ' treeview, etc. Those tasks simply slow down the exiting of the program.
   gApp.IsShutdown = true

   ' Save whether the Panel should be shown the next time the program is run.
   ' Also save the panel width. 
   dim pWindow as CWindow ptr = AfxCWindowPtr(HWND)
   gConfig.ShowPanel = IsWindowVisible(HWND_FRMPANEL)
   gConfig.ShowPanelWidth = pWindow->UnScaleX(AfxGetWindowWidth( HWND_FRMPANEL ))
   
   ' if a project(s) is already open then save/close it.
   if gApp.IsProjectActive then 
      if OnCommand_ProjectClose(HWnd) = 0 then return 0
   end if
   
   ' Close any open files asking to save any that are dirty 
   if OnCommand_FileClose(HWnd, EFC_CLOSEALL) = 0 then return 0
   
   DestroyWindow(HWnd)
   
   function = 0
end Function


' ========================================================================================
' Process WM_DESTROY message for window/dialog: frmMain
' ========================================================================================
function frmMain_OnDestroy( byval HWnd as HWnd ) as LRESULT

   ' Kill any existing AutoSave timer
   OnCommand_FileAutoSaveKillTimer()

   ' Output the config settings to disk file
   gConfig.SaveConfigFile

   ' Disable drag and drop files
   DragAcceptFiles HWnd, False
   
   PostQuitMessage(0)
   function = 0
end Function


' ========================================================================================
' Process WM_MOUSEMOVE message for window/dialog: frmMain
' ========================================================================================
function frmMain_OnMouseMove( _
            byval HWnd as HWnd, _
            byval x as long, _
            byval y as long, _
            byval keyflags as UINT _
            ) as long

   dim pWindow as CWindow ptr = AfxCWindowPtr( HWND_FRMMAIN )
   if pWindow = 0 then exit function
   
   dim as RECT rcClient
   GetClientRect( hwnd, @rcClient )
   
   ' HITTEST (DOCUMENT SPLITTER)
   dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr()
   if pDoc then 
      dim as Rect rc
      dim as Point pt = (x, y)
      if pDoc->bSizing then
         if pt.y <> pDoc->ptPrev.y then
            pDoc->SplitY = pDoc->SplitY + (pt.y - pDoc->ptPrev.y)
            ' Don't let the split go all the way to the bottom of the edit window
            if pDoc->SplitY + pWindow->ScaleY(40) >= rcClient.bottom then 
            else
               frmMain_PositionWindows
               pDoc->ptPrev.y = pt.y
            end if   
         end if   
      else   
         if PtInRect(@pDoc->rcSplitButton, pt) then 
            SetCursor( ghCursorSizeNS )
         end if   
      end if
   end if
         
   function = 0
end function


' ========================================================================================
' Process WM_LBUTTONDOWN message for window/dialog: frmMain
' ========================================================================================
function frmMain_OnLButtonDown( _
            byval HWnd as HWnd, _
            byval fDoubleClick as boolean, _
            byval x as long, _
            byval y as long, _
            byval keyflags as UINT _
            ) as long
                                        
   ' if main window area clicked then close any active top menus
   killAllPopupMenus()
   
   ' HITTEST (DOCUMENT SPLITTER)
   dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr()
   if pDoc then 
      dim as Rect rc 
      dim as POINT pt = (x, y)
      if PtInRect(@pDoc->rcSplitButton, pt) then 
         pDoc->bSizing = true
         pDoc->SplitY = pDoc->rcSplitButton.top  
         pDoc->ptPrev.y = pt.y
         SetCursor( ghCursorSizeNS )
         SetCapture( HWND_FRMMAIN )
      end if
   end if

   function = 0
end function

         
' ========================================================================================
' Process WM_LBUTTONUP message for window/dialog: frmMain
' ========================================================================================
function frmMain_OnLButtonUp( _
            byval HWnd as HWnd, _
            byval x as long, _
            byval y as long, _
            byval keyflags as UINT _
            ) as long

   ' HITTEST (DOCUMENT SPLITTER)
   dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr()
   if pDoc then 
      if pDoc->bSizing then
         pDoc->bSizing = false
         ReleaseCapture
         ' if we try to split before the start of the code window then simply
         ' stop the splitting altogether.
         dim as RECT rc = AfxGetWindowRect( pDoc->hWindow(1) )
         MapWindowPoints( HWND_DESKTOP, HWND_FRMMAIN, cast(POINT ptr, @rc), 2 )
         if pDoc->SplitY <= rc.top then
            pDoc->SplitY = 0
            pDoc->bEditorIsSplit = false
            frmMain_PositionWindows()
         end if   
      end if
   end if   
   SetCursor( LoadCursor( null, IDC_ARROW ))
         
   function = 0
end function


' ========================================================================================
' Process WM_LBUTTONDBLCLK message for window/dialog: frmMain
' ========================================================================================
function frmMain_OnLButtonDblClk( _
            byval HWnd as HWnd, _
            byval fDoubleClick as boolean, _
            byval x as long, _
            byval y as long, _
            byval keyflags as UINT _
            ) as long

   if gApp.bDragActive = false then
      ' Are we over a split edit area (toggle off the split)
      dim as POINT pt = (x, y)
      dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr()
      if pDoc then 
         if PtInRect(@pDoc->rcSplitButton, pt) then 
            pDoc->bSizing = false
            pDoc->SplitY = 0
            pDoc->bEditorIsSplit = false
            ReleaseCapture
            frmMain_PositionWindows
         end if
      end if
   end if
   
   function = 0
end function


' ========================================================================================
' frmMain Window procedure
' ========================================================================================
function frmMain_WndProc( _
            byval HWnd   as HWnd, _
            byval uMsg   as UINT, _
            byval wParam as WPARAM, _
            byval lParam as LPARAM _
            ) as LRESULT

   select case uMsg
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
      HANDLE_MSG (HWnd, WM_LBUTTONDOWN,   frmMain_OnLButtonDown)
      HANDLE_MSG (HWnd, WM_LBUTTONUP,     frmMain_OnLButtonUp)
      HANDLE_MSG (HWnd, WM_MOUSEMOVE,     frmMain_OnMouseMove)
      HANDLE_MSG (HWnd, WM_LBUTTONDBLCLK, frmMain_OnLButtonDblClk)

      case WM_ERASEBKGND
         return true      ' prevents painting the background
         
      case WM_MOVE
         ' Ensure that if the Find/Replace dialog is open that it moves with
         ' underlying main WinFBE window.     
         frmFindReplace_PositionWindows
      
      case WM_NCACTIVATE
         ' Ensure that the caption bar for any DesignerForm retains its active 
         ' state. This looks better than having it lose and gain active colors. 
         if wParam then
            ' The DesignerForm never receives the WM_NCACTIVATE message because it
            ' is a child form. We need to manually send it the message whenever the
            ' focus of the Designer form changes (i.e. whenever we switch back and
            ' forth away from the main application window).
            PostMessage HWND, MSG_USER_SETFOCUS, 0, 0
         end if
         
         ' if wParam is false then the system wants to draw an inactive title bar. we
         ' will prevent this action if a popup menu from the menubar is active.
         if wParam = false then
            if gPrevent_WM_NCACTIVATE then return false
         end if
          
      case WM_SYSCOMMAND
         if (wParam and &HFFF0) = SC_CLOSE then
            SendMessage( HWND, WM_CLOSE, wParam, lParam )
            exit function
         end if
            
      case WM_SETCURSOR
         if gApp.bDragActive = true then
            ' Either the horiz or vert splitter bar is being resized. We handle
            ' setting the mouse cursor in the splitter_mousemove function.
            return true
            
         elseif gApp.bDragTabActive = true then 
            ' We handle the cursor in the TabControl subclass. Don't allow the
            ' main window to change our cursor. return true.
            return true

         elseif (gApp.IsCompiling = true) or _
                (gApp.IsProjectLoading = true) or _
                (gApp.IsFileLoading = true) then
            SetCursor( LoadCursor(0, IDC_WAIT) )
            return true
         end if
            
      case WM_CAPTURECHANGED
         gApp.bDragTabActive = false
         
      case WM_SETFOCUS
         frmMain_SetFocusToCurrentCodeWindow

      case WM_COPYDATA      ' used during processing of commandline
         dim pDataToGet as COPYDATASTRUCT ptr
         dim pwszArg as wstring ptr 
         dim wszExt as wstring * MAX_PATH

         pDataToGet = cast(COPYDATASTRUCT ptr, lParam)
         pwszArg = pDataToGet->lpData

         ' We have a valid filename so determine what type it is.
         wszExt = AfxStrPathname( "EXTN", *pwszArg )
         wszExt = ucase(wszExt)
          
         select case wszExt
            case ".WFBE"    ' project file
               ' Pass the info to our generic project open function to handle everything.
               frmMain_OpenProjectSafely(HWND_FRMMAIN, *pwszArg) 
             
            case else   ' .bas, .bi, .rc, etc...
               dim pDoc as clsDocument ptr
               pDoc = frmMain_OpenFileSafely( HWND, _
                                       false, _    ' bIsNewFile
                                       false, _    ' bIsTemplate
                                       true,  _    ' bShowInTab
                                       false, _    ' bIsInclude
                                       *pwszArg, _  ' pwszName
                                       0 )         ' pDocIn
               if pDoc then UpdateMRUList(pDoc->DiskFilename) 
         end select    
          

      ''  CUSTOM MESSAGES

      case MSG_USER_UPGRADE302FORM
         ' After a form file is loaded a PostMessage is made using this custom message
         ' in order to initiate an upgrade check on that file.
         if lParam then
            dim pDoc as clsDocument ptr = cast(clsDocument ptr, lParam)
            FormUpgrade302Format( pDoc )
         end if   
   
      case MSG_USER_TOPTABS_CHANGING
         ' Hide the current tab
         gTTabCtl.DisplayScintilla( gTTabCtl.CurSel, false )
         exit function
         
      case MSG_USER_TOPTABS_CHANGED
         ' Show the new tab
         gTTabCtl.DisplayScintilla( gTTabCtl.CurSel, true )
         frmMain_PositionWindows
         dim pDoc as clsDocument ptr = gTTabCtl.tabs(gTTabCtl.CurSel).pDoc
         if pDoc then
            AfxRedrawWindow( pdoc->hWindow(0) )  ' makes the code window display quicker especially when loading Project.
            if pDoc->IsDesigner andalso IsDesignerView(pDoc) then
               DisplayPropertyList( pDoc )
            end if   
         end if   
         ' Set the focus to the Scintilla window. This will call frmMain_SetStatusbar.
         frmMain_SetFocusToCurrentCodeWindow()
         exit function

      case MSG_USER_SETFOCUS
         ' Set focus to current Scintilla window and update the document
         ' display such as Line#, Col#, Filename, etc.
         frmMain_SetStatusbar
         dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr()
         if pDoc then
            if pDoc->IsDesigner then
               SendMessage pDoc->hWndForm, WM_NCACTIVATE, true, 0
            end if
            SetFocus( pDoc->hWndActiveScintilla )
         end if

      case MSG_USER_PROCESS_COMMANDLINE 
         ' process any command line arguments that may have been passed to the program.
         frmMain_ProcessCommandLine(HWnd)

      case MSG_USER_PROCESS_UPDATECHECK
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

      case MSG_USER_PROCESS_STARTUPUSERTOOLS
         ' process any User Tools that must display immediately after the main WinFBE
         ' user interface is created and displayed.
         
         ' Wait a bit to ensure that the WinFBE window is visible... This loop
         ' is mostly never necessary but better to check just to be consistent.
         dim as double nStartTime = timer
         do
            if IsWindowVisible( HWND_FRMMAIN ) then exit do
            ' if we've waited more than 5 seconds then break out to avoid 
            ' an infite loop.
            if (timer - nStartTime) > 5 then exit do
         loop

         ' Only invoke any User Tools that have action for startup
         for y as long = lbound(gConfig.Tools) to ubound(gConfig.Tools)
            if gConfig.Tools(y).Action = USERTOOL_ACTION_WINFBESTARTUP then
               frmUserTools_ExecuteUserTool(y)
            end if   
         NEXT
         
         
      case MSG_USER_GENERATECODE
         dim pDoc as clsDocument ptr 
         if lParam = 0 then
            pDoc = gTTabCtl.GetActiveDocumentPtr()
         else
            pDoc = cast(clsDocument ptr, lParam)
         end if   
         if pDoc then
            pDoc->bRegenerateCode = true
            if IsDesignerView(pDoc) = false then  ' clicked on the code window
               dim as hwnd hEdit = pDoc->hWindow(0)
               dim as long nFirstLine = SendMessage( hEdit, SCI_GETFIRSTVISIBLELINE, 0, 0)
               dim as long curPos = SendMessage( hEdit, SCI_GETCURRENTPOS, 0, 0)
               GenerateFormCode(pDoc)
               if curPos >=0 then SendMessage( hEdit, SCI_GOTOPOS, curPos, 0)
               if nFirstLine >=0 then SendMessage( hEdit, SCI_SETFIRSTVISIBLELINE, nFirstLine, 0)
            else
               GenerateFormCode(pDoc)
            end if
         end if

      case MSG_USER_SHOWAUTOCOMPLETE
         return ShowAutocompleteList()

      case MSG_USER_APPENDEQUALSSIGN
         ' The = key was used to terminate a popup autocomplete. Take that
         ' character and format it with a space before and a space after.
         dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr()
         if pDoc then
            dim as hwnd hEdit = pDoc->hWndActiveScintilla
            dim as long nPos = SciExec( hEdit, SCI_GETCURRENTPOS, 0, 0)
            dim as string sWord = " = "
            SciExec( hEdit, SCI_SETSEL, nPos-1, nPos)
            SciExec( hEdit, SCI_REPLACESEL, 0, cast(LPARAM, Strptr(sWord)))
            nPos = SciExec( hEdit, SCI_GETCURRENTPOS, 0, 0)
            SciExec( hEdit, SCI_SETSEL, nPos, nPos)
         end if   
              

      '-------------------------------------------------------------------------------
      '                        Scrollbar Handler
      '-------------------------------------------------------------------------------
      case WM_VSCROLL
         dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr()
         if pDoc = 0 then return 0

'         dim as long OldPos = pDoc->ScrInfo.nPos

'         select case loword(wParam)
'            case SB_THUMBTRACK
'               GetScrollInfo(pDoc->hScrollbar, SB_CTL, @pDoc->ScrInfo)
'               pDoc->ScrInfo.nPos = pDoc->ScrInfo.nTrackPos
'            case SB_LINEDOWN:   pDoc->ScrInfo.nPos += 1
'            case SB_LINEUP:     pDoc->ScrInfo.nPos -= 1
'            case SB_PAGEDOWN:   pDoc->ScrInfo.nPos += pDoc->ScrInfo.nPage - 1  
'            case SB_PAGEUP:     pDoc->ScrInfo.nPos -= pDoc->ScrInfo.nPage - 1  
'         end select  
          
'         ' if the current position hasn't changed, do nothing.
'         if pDoc->ScrInfo.nPos = oldPos then return true

'         ' Don't exceed range boundries
'         if pDoc->ScrInfo.nPos < pDoc->ScrInfo.nMin then
'            pDoc->ScrInfo.nPos = pDoc->ScrInfo.nMin         
'         end if
'         if pDoc->ScrInfo.nPos > pDoc->ScrInfo.nMax - pDoc->ScrInfo.nPage + 1 then
'            pDoc->ScrInfo.nPos = pDoc->ScrInfo.nMax - pDoc->ScrInfo.nPage + 1        
'         end if

'         dim as long lParm = pDoc->ScrInfo.nPos - oldPos        ' Amount/direction to V scroll
'         SciExec(pDoc->hWindow(0), SCI_LINESCROLL, 0, lParm)
'         SetScrollInfo(pDoc->hScrollbar, SB_CTL, @pDoc->ScrInfo, true)

   end select 

   ' for messages that we don't deal with
   function = DefWindowProc(HWnd, uMsg, wParam, lParam)

end Function


' ========================================================================================
' frmMain_Show
' ========================================================================================
function frmMain_Show( byval hWndParent as HWnd ) as LRESULT

   '  Create the main window and child controls
   dim pWindow as CWindow ptr = new CWindow("WinFBE_Class")
   ' Comment out the next line to let WinFBE use the current active system DPI setting.
   'pWindow->DPI = 144   ' eg. 144 or any other value (96 is default)
   
   HWND_FRMMAIN = pWindow->Create( null, APPNAME, @frmMain_WndProc )
   
   ' Load the currently selected theme (this needs a valid pWindow HWND_FRMMAIN to exists)
   LoadThemeFile()
   
   ' Set the small and large icon for the main window (must be set after main window is created)
   pWindow->BigIcon   = LoadImage( pWindow->InstanceHandle, "IMAGE_AAA_MAINICON", IMAGE_ICON, 32, 32, LR_SHARED)
   pWindow->SmallIcon = LoadImage( pWindow->InstanceHandle, "IMAGE_AAA_MAINICON", IMAGE_ICON, 16, 16, LR_SHARED)

   ' Load the good and bad compile "icons"
   ghIconGood = 1  
   ghIconBad  = 2 
   
   ' Load the tick and untick icons
   dim cx as long =  16 * (pWindow->DPI \ 96)
   ghIconTick   = LoadImage( pWindow->InstanceHandle, "IMAGE_TICK", IMAGE_ICON, cx, cx, LR_DEFAULTCOLOR )
   ghIconNoTick = LoadImage( pWindow->InstanceHandle, "IMAGE_NOTICK", IMAGE_ICON, cx, cx, LR_DEFAULTCOLOR )

   ' Load the North/South and East/West cursor images
   ghCursorSizeNS = LoadImage( Null, MAKEINTRESOURCEW(OCR_SIZENS), IMAGE_CURSOR, 0, 0, LR_DEFAULTSIZE or LR_SHARED )
   ghCursorSizeWE = LoadImage( Null, MAKEINTRESOURCEW(OCR_SIZEWE), IMAGE_CURSOR, 0, 0, LR_DEFAULTSIZE or LR_SHARED )

   ' Build the keyboard accelerators
   frmKeyboard_BuildAcceleratorTable

   ' Create the main application topmenu menubar
   frmMenuBar_Show( HWND_FRMMAIN )

   ' Create the main application top tabs control
   frmTopTabs_Show( HWND_FRMMAIN )

   ' Create the main application design tabs control
   frmDesignTabs_Show( HWND_FRMMAIN )

   ' Create the main application statusbar
   frmStatusBar_Show( HWND_FRMMAIN )
  
   ' Create the various child windows
   frmPanel_Show( HWND_FRMMAIN )
   frmPanelVScroll_Show( HWND_FRMMAIN )
   frmOutput_Show( HWND_FRMMAIN )
   frmEditorHScroll_Show( HWND_FRMMAIN )
   frmEditorVScroll_Show( HWND_FRMMAIN )
   
   ' Create the UserTools accelerator table
   frmUserTools_CreateAcceleratorTable            

   ' SET STARTUP POSITION
   ' if no valid window size exists then set to the default working area of the screen
   if (gConfig.StartupRight = 0) orelse (gConfig.StartupBottom = 0) then
      ' Retrieve the size of the working area
      dim rc as Rect = pWindow->GetWorkArea                
      gConfig.StartupRight  = rc.Right
      gConfig.StartupBottom = rc.Bottom
   end if
   
   dim WinPla as WINDOWPLACEMENT
   with WinPla
      .Length = sizeof(WinPla)
      .rcNormalPosition.Left   = gConfig.StartupLeft
      .rcNormalPosition.Top    = gConfig.StartupTop
      .rcNormalPosition.Right  = gConfig.StartupRight
      .rcNormalPosition.Bottom = gConfig.StartupBottom
      .showCmd = iif( gConfig.StartupMaximized, SW_MAXIMIZE, SW_SHOWNORMAL )
   end with
   SetWindowPlacement( pWindow->hWindow, @WinPla )
     
   ' Ensure the window is placed on screen should the user had changed 
   ' the logical ordering of a multiple display setup.
   AfxForceVisibleDisplay( pWindow->hWindow )
   
   UpdateWindow( pWindow->hWindow )

   ' Post a message to process the application's command line as applicable.
   PostMessage( pWindow->hWindow, MSG_USER_PROCESS_COMMANDLINE, 0, 0 )

   ' Post a message to process any User Tools that must start after WinFBE displays.
   PostMessage( pWindow->hWindow, MSG_USER_PROCESS_STARTUPUSERTOOLS, 0, 0 )

   ' Post a message to do an update check (if applicable)
   PostMessage( pWindow->hWindow, MSG_USER_PROCESS_UPDATECHECK, 0, 0 )

   ' Start AutoSave timer if that option is enabled
   OnCommand_FileAutoSaveStartTimer()

   ' Process windows events
   dim uMsg as MSG
   
   ' Message loop
   do while GetMessage(@uMsg, null, 0, 0)

      if handleMouseScrollBar(uMsg) then continue do
      if handleMouseShowScrollBar(uMsg) then continue do
      if handleMouseTopMenu(uMsg) then continue do
      if handleAltKeyMenuBar(uMsg) then continue do

      ' Processes accelerator keys for menu commands
      if (pWindow->AccelHandle = 0) orelse (TranslateAccelerator(pWindow->hWindow, pWindow->AccelHandle, @uMsg) = 0) then
          
         if (ghAccelUserTools = 0) orelse (TranslateAccelerator(pWindow->hWindow, ghAccelUserTools, @uMsg) = 0) then 

            ' Prevent any < asc(32) characters from making their way to Scintilla where they 
            ' get shown as an embedded control graphic. Allow backspace.
            if (uMsg.message = WM_CHAR) then
               if (uMsg.wParam < 32) andalso (uMsg.wParam <> 8) then continue do
               ' Save the pdoc->LastCharTyped to use in our autocomplete popups
               dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr()
               If pDoc = 0 then pdoc->LastCharTyped = uMsg.wParam
            end if
 
            ' Save the pdoc->LastCharTyped to use in our autocomplete popups
            if (uMsg.message = WM_KEYDOWN) then
               dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr()
               if pDoc then pdoc->LastCharTyped = uMsg.wParam
            end if
            
            ' if topmenus or menubar are active then process all keyboard input for them rather than 
            ' passing it off to the system to process.
            if handleKeysTopMenu(uMsg) then continue do

            ' Was ENTER key pressed while scrolling the Explorer listbox via keyboard
            if handleKeysExplorerListBox(uMsg) then continue do
            
            ' Close with ESC key: Find/Replace, FindInFiles, FunctionList window
            if handleEscKeyModeless(uMsg) then continue do

            ' Handle any keypress that would move or resize control(s) on a Designer Form.
            if handleKeysVisualDesigner(uMsg) then continue do
            
            ' Handle ENTER key for active FindReplace dialog
            if handleKeysFindReplace(uMsg) then continue do
            
            ' Check for any QuickRun exes that can be deleted.
            gApp.CheckQuickRunExe()
             
            ' Check if the WinFBE.ini file was changed by an external program.
            ' Reload the config file in case a user has automated a change to it since
            ' the application started (for example, changing the compiler path).
            gConfig.ReloadConfigFileTest()
            
            ' Determines whether a message is intended for the specified
            ' dialog box and, if it is, processes the message.
            ' Ensure keystrokes like TAB are properly handled by the modeless dialogs
            if AfxCAxHostForwardMessage(GetFocus, @uMsg) = false then
               if IsDialogMessage( HWND_FRMFINDREPLACE, @uMsg ) then Continue Do
               if IsDialogMessage( HWND_FRMHELPVIEWER, @uMsg ) then Continue Do

               if IsDialogMessage(pWindow->hWindow, @uMsg) = 0 then
                  TranslateMessage @uMsg    ' Translates virtual-key messages into character messages.
                  DispatchMessage @uMsg     ' Dispatches a message to a window procedure.
               end if
            end if

         end if  'accelerator user tools    
      end if  ' accelerators
   loop  ' message loop
   
   function = uMsg.wParam

   if ghIconTick           then DestroyIcon( ghIconTick )
   if ghIconNoTick         then DestroyIcon( ghIconNoTick )
   
   if ghAccelUserTools then DestroyAcceleratorTable( ghAccelUserTools )
   
   ' delete the allocated memory for the various child windows
   pWindow = AfxCWindowPtr( HWND_FRMHELPVIEWER ): delete pWindow
   pWindow = AfxCWindowPtr( HWND_FRMPANEL ): delete pWindow
   pWindow = AfxCWindowPtr( HWND_FRMPANEL_VSCROLLBAR ): delete pWindow
   pWindow = AfxCWindowPtr( HWND_FRMEDITOR_HSCROLLBAR(0) ): delete pWindow
   pWindow = AfxCWindowPtr( HWND_FRMEDITOR_HSCROLLBAR(1) ): delete pWindow
   pWindow = AfxCWindowPtr( HWND_FRMEDITOR_VSCROLLBAR(0) ): delete pWindow
   pWindow = AfxCWindowPtr( HWND_FRMEDITOR_VSCROLLBAR(1) ): delete pWindow
   pWindow = AfxCWindowPtr( HWND_FRMOUTPUT ): delete pWindow
   pWindow = AfxCWindowPtr( HWND_FRMMAIN ): delete pWindow

end function

