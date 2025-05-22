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

' ========================================================================================
' Move next/prev amongst the various compiler errors
' ========================================================================================
function OnCommand_SearchGotoCompileError( byval bMoveNext as boolean ) as long
   ' Determine the current selection in the listview and move prev/next as needed
   dim as HWND hLV = GetDlgItem( HWND_FRMOUTPUT, IDC_FRMOUTPUT_LVRESULTS )
   Dim As Long nCurSel = ListView_GetSelection(hLV)
   If nCurSel < 0 Then Return 0 

   ' Determine how many lines in the listview so we can "wrap" around
   dim as long nCount = ListView_GetItemCount(hLV)
  
   if bMoveNext then 
      nCurSel = nCurSel + 1
   else   
      nCurSel = nCurSel - 1
   end if
   if nCurSel > nCount - 1 then nCurSel = 0
   if nCurSel < 0 then nCurSel = nCount - 1 

   ListView_SelectItem( hLV, nCurSel )
   SetDocumentErrorPosition( hLV, gCompile.CompileID )

   function = 0
end function


' ========================================================================================
' Save the current editor position so we can return to it should the user invoke
' the Last Position (Shift+F6) option. 
' ========================================================================================
function OnCommand_SearchSaveLastPosition( byval pDoc as clsDocument ptr ) as long
   if pDoc = 0 then exit function
   dim as hwnd hEdit = pDoc->hWndActiveScintilla
   gLastPosition.pDoc       = pDoc
   gLastPosition.nFirstLine = SciExec( hEdit, SCI_GETFIRSTVISIBLELINE, 0, 0) 
   gLastPosition.nPosition  = SciExec( hEdit, SCI_GETCURRENTPOS, 0, 0) 
   function = 0
end function

' ========================================================================================
' Position editor to document/function/variable under the current caret position
' ========================================================================================
function OnCommand_SearchGotoDefinition( byval pDoc as clsDocument ptr ) as LRESULT

   If pDoc = 0 Then Exit Function
   
   dim pData as DB2_DATA ptr
   dim as CWSTR wszFilename
   dim as string sTypeName, sFunctionName, sLookFor, sTrigger
   dim as long nLineNum = 0
   
   dim as hwnd hEdit = pDoc->hWndActiveScintilla
   
   ' Version 1.8.8 added ability to open an underlying #Include file if such
   ' a file exists on the current editing line.
   if IsCurrentLineIncludeFilename() then
      wszFilename = gApp.IncludeFilename 
   else

      ' Determine the word at the current caret position if nothing is selected.
      sFunctionName = trim(pDoc->GetSelText)
      if len(sFunctionName) = 0 then sFunctionName = trim(pDoc->GetWord)

      if len(sFunctionName) then
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
               sTypeName = pData->VariableType
               sLookFor = sTypeName & "." & sFunctionName
               ' Set pData to null because it currently points at the TYPE definition and
               ' we the gdb2 searches below to actually find the sub/function.
               pData = 0
            end if
         end if   

         ' Search the list of TYPE Function names
         if pData = 0 then pData = gdb2.dbFindFunctionTYPE( sTypeName, sLookFor )
         ' Search the list of Function names
         if pData = 0 then pData = gdb2.dbFindFunction( sFunctionName )
         ' Search for variable definitions
         if pData = 0 then pData = gdb2.dbFindVariable( "", sFunctionName )
         ' Search for the TYPE definition
         if pData = 0 then pData = gdb2.dbFindType( sTypeName )
      
      end if
      
      IF pData = 0 then
         MessageBox( HWND_FRMMAIN, L(224,"Sub/Function definition not found."), @WStr(APPNAME), MB_ICONWARNING)
         exit function
      end if
      wszFilename = pData->filename
      nLineNum = pData->nLineStart
   
   end if
   
   ' Save our current position in case user invokes "Last Position" option,   
   OnCommand_SearchSaveLastPosition( pDoc )
   OpenSelectedDocument( wszFilename, sFunctionName, nLineNum )

   Function = 0
End Function


' ========================================================================================
' Return to most previous position (invoked when "Last Position" is selected
' ========================================================================================
function OnCommand_SearchGotoLastPosition() as LRESULT

   ' Is the pDoc pointer still valid
   If gLastPosition.pDoc = 0 Then Exit Function

   ' Position ourselves to the correct previous position.
   frmMain_OpenFileSafely(HWND_FRMMAIN, _
            False, _    ' bIsNewFile
            False, _    ' bIsTemplate
            True,  _    ' bShowInTab
            false, _    ' bIsInclude
            gLastPosition.pDoc->DiskFilename, _ ' wszFileName
            gLastPosition.pDoc )  ' pDocIn
   SciExec( GetFocus, SCI_SETFIRSTVISIBLELINE, gLastPosition.nFirstLine, 0) 
   SciExec( GetFocus, SCI_GOTOPOS, gLastPosition.nPosition, 0) 
   
   Function = 0
End Function


' ========================================================================================
' Position editor to Header/Source/Main/Resource document
' ========================================================================================
function OnCommand_SearchGotoFile( _
            ByVal id as long, _
            byval pDoc as clsDocument ptr _
            ) as LRESULT

   If pDoc = 0 Then Exit Function
   
   dim pDocFile as clsDocument ptr
   dim as CWSTR wszFilename

   dim as HWND hEdit = pDoc->hWndActiveScintilla
   
   select case id
      Case IDM_GOTO
         frmGoto_Show( hEdit )
         exit function
      
      case IDM_GOTONEXTFUNCTION
         pDoc->GotoNextFunction
         exit function
      
      case IDM_GOTOPREVFUNCTION
         pDoc->GotoPrevFunction
         exit function
         
      Case IDM_GOTONEXTTAB
         gTTabCtl.NextTab
         exit function
      
      Case IDM_GOTOPREVTAB
         gTTabCtl.PrevTab
         exit function
      
      Case IDM_CLOSETAB
         gTTabCtl.CloseTab
         exit function

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
   gLastPosition.pDoc       = pDoc
   gLastPosition.nFirstLine = SciExec( hEdit, SCI_GETFIRSTVISIBLELINE, 0, 0) 
   gLastPosition.nPosition  = SciExec( hEdit, SCI_GETCURRENTPOS, 0, 0) 

   OpenSelectedDocument( wszFilename )
   
   Function = 0
End Function


' ========================================================================================
' Search action Common actions
' ========================================================================================
function OnCommand_SearchBookmarks( _
            ByVal id as long, _
            byval pDoc as clsDocument ptr _
            ) as LRESULT
      
   dim as HWND hEdit = pDoc->hWndActiveScintilla

   select case id
      Case IDM_BOOKMARKNEXT:     pDoc->NextBookmark
      Case IDM_BOOKMARKPREV:     pDoc->PrevBookmark
      
      Case IDM_BOOKMARKCLEARALL
         SciExec( hEdit, SCI_MARKERDELETEALL, -1, 0 )
         LoadBookmarksFiles
         AfxRedrawWindow( HWND_FRMBOOKMARKS )

      Case IDM_BOOKMARKTOGGLE
         pDoc->ToggleBookmark( pDoc->GetCurrentLineNumber )
         LoadBookmarksFiles
         AfxRedrawWindow( HWND_FRMBOOKMARKS )

      case IDM_BOOKMARKCLEARALLDOCS
         dim pDoc as clsDocument ptr = gApp.pDocList
         do until pDoc = 0
            hEdit = pDoc->hWindow(0)
            SciExec( hEdit, SCI_MARKERDELETEALL, -1, 0 )
            pDoc = pDoc->pDocNext
         loop
         LoadBookmarksFiles
         AfxRedrawWindow( HWND_FRMBOOKMARKS )
   end select
   
   function = 0
end function
