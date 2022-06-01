'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2022 Paul Squires, PlanetSquires Software
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
' Edit action Cut
' ========================================================================================
function OnCommand_EditCut( _
            byval pDoc as clsDocument ptr, _
            ByVal hEdit as HWND _
            ) as LRESULT
   
   ' Check if the Compiler log file or Notes windows or Snippets have
   ' input focus. If they do then select that text rather than Scintilla.
   select case GetFocus()
      case GetDlgItem(HWND_FRMOUTPUT, IDC_FRMOUTPUT_TXTNOTES), _
           GetDlgItem(HWND_FRMOUTPUT, IDC_FRMOUTPUT_TXTLOGFILE), _ 
           GetDlgItem(HWND_FRMSNIPPETS, IDC_FRMSNIPPETS_TXTCODE), _
           GetDlgItem(HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_TXTFIND), _
           GetDlgItem(HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_TXTREPLACE), _
           HWND_PROPLIST_EDIT
         SendMessage( GetFocus, WM_CUT, 0, 0 )
      case else     
         if IsDesignerView(pDoc) then
            SendMessage(HWND_FRMMAIN, WM_COMMAND, MAKELONG(IDM_COPY, 0), 0)
            SendMessage(HWND_FRMMAIN, WM_COMMAND, MAKELONG(IDM_DELETE, 0), 0)
         else
            SciExec( GetFocus(), Iif(GetFocus()=hEdit, SCI_CUT, WM_CUT), 0, 0)
            frmMain_PositionWindows()    ' ensure custom scrollbars display
         end if   
   end select

   function = 0
end function
      

' ========================================================================================
' Edit action Copy
' ========================================================================================
function OnCommand_EditCopy( _
            byval pDoc as clsDocument ptr, _
            ByVal hEdit as HWND _
            ) as LRESULT
   
   ' Check if the Compiler log file or Notes windows or Snippets have
   ' input focus. If they do then select that text rather than Scintilla.
   select case GetFocus()
      case GetDlgItem(HWND_FRMOUTPUT, IDC_FRMOUTPUT_TXTNOTES), _
           GetDlgItem(HWND_FRMOUTPUT, IDC_FRMOUTPUT_TXTLOGFILE), _ 
           GetDlgItem(HWND_FRMSNIPPETS, IDC_FRMSNIPPETS_TXTCODE), _
           GetDlgItem(HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_TXTFIND), _
           GetDlgItem(HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_TXTREPLACE), _
           HWND_PROPLIST_EDIT
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

   function = 0
end function


' ========================================================================================
' Edit action Paste
' ========================================================================================
function OnCommand_EditPaste( _
            byval pDoc as clsDocument ptr, _
            ByVal hEdit as HWND _
            ) as LRESULT

   ' Check if the Compiler log file or Notes windows or Snippets have
   ' input focus. If they do then select that text rather than Scintilla.
   select case GetFocus()
      case GetDlgItem(HWND_FRMOUTPUT, IDC_FRMOUTPUT_TXTNOTES), _
           GetDlgItem(HWND_FRMOUTPUT, IDC_FRMOUTPUT_TXTLOGFILE), _ 
           GetDlgItem(HWND_FRMSNIPPETS, IDC_FRMSNIPPETS_TXTCODE), _
           GetDlgItem(HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_TXTFIND), _
           GetDlgItem(HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_TXTREPLACE), _
           HWND_PROPLIST_EDIT
         SendMessage( GetFocus, WM_PASTE, 0, 0 )
      case else     
         if IsDesignerView(pDoc) then
            dim pCtrl as clsControl ptr
            dim rc as RECT
            dim as long nLeft, nTop, nLeftActive = 99999, nTopActive = 99999
            dim as long nCurLeft, nCurTop
            
            ' Need to get the Left/Top properties for least top/left in the copy group
            for i as long = lbound(gCopyControls) to ubound(gCopyControls)
               nCurLeft = val(GetControlProperty(@gCopyControls(i), "LEFT"))
               nCurTop  = val(GetControlProperty(@gCopyControls(i), "TOP"))
               if nCurLeft < nLeftActive then nLeftActive = nCurLeft
               if nCurTop < nTopActive then nTopActive = nCurTop 
               exit for
            next
               
            ' Create the control and copy the properties to it
            for i as long = lbound(gCopyControls) to ubound(gCopyControls)
               pCtrl = CreateToolboxControl(pDoc, gCopyControls(i).ControlType, rc)
               gCopyControls(i).hWindow = pCtrl->hWindow
               
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
               ' Map the point to the client area coordinates
               MapWindowPoints(0, pDoc->hWndForm, cast(point ptr, @pt), 1)

               ' Ensure that the point is within the client area
               dim rc as RECT
               GetWindowRect(pDoc->hWndForm, @rc)
               MapWindowPoints(0, pDoc->hWndForm, cast(point ptr, @rc), 2)
               if PtInRect(@rc, pt) = 0 then
                  pt.x = 10: pt.y = 10
               end if

               Dim pWindow As CWindow Ptr = AfxCWindowPtr(HWND_FRMMAIN)
               if pWindow then
                  pt.x = pWindow->UnScaleX(pt.x)
                  pt.y = pWindow->UnScaleY(pt.y)
               end if
               SetControlProperty(pCtrl, "LEFT", str(pt.x + (nLeft-nLeftActive)))
               SetControlProperty(pCtrl, "TOP", str(pt.y + (nTop-nTopActive)))
               
               pCtrl->SuspendLayout = true
               ApplyControlProperties( pDoc, pCtrl )
               pCtrl->SuspendLayout = false
            NEXT

            ' Select all of the newly created controls
            for i as long = lbound(gCopyControls) to ubound(gCopyControls)
               pDoc->Controls.SelectControl( gCopyControls(i).hWindow )
            next 

            pDoc->bRegenerateCode = true
            pDoc->UserModified = true
            AfxRedrawWindow(pDoc->hWndFrame)
            AfxRedrawWindow(pDoc->hWndForm)
            frmMain_SetStatusbar
            DisplayPropertyList(pDoc)
         else
            SciExec( GetFocus(), Iif(GetFocus()=hEdit, SCI_PASTE, WM_PASTE), 0, 0)
            frmMain_PositionWindows()    ' ensure custom scrollbars display
            SetFocus( hEdit )    ' ensure that the original edit window has not lost focus
         end if   
   end select
   
   function = 0
end function


' ========================================================================================
' Edit action Indent Block
' ========================================================================================
function OnCommand_EditIndentBlock( _
            byval pDoc as clsDocument ptr, _
            ByVal hEdit as HWND _
            ) as LRESULT

   if (pDoc->IsDesigner) andalso (IsDesignerView(pDoc)) then exit function
   
   ' If a TAB was pressed then manually terminate any active autocomplete.
   if SciExec(hEdit, SCI_AUTOCACTIVE, 0, 0) then
      SciExec(hEdit, SCI_AUTOCCOMPLETE, 0, 0) 
   end if
   if GetFocus = GetDlgItem(HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_TXTFIND) then
      if gFind.bExpanded then SetFocus GetDlgItem(HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_TXTREPLACE)
   elseif GetFocus = GetDlgItem(HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_TXTREPLACE) then
      if gFind.bExpanded then SetFocus GetDlgItem(HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_TXTFIND)
   elseif AfxGetFormHandle(GetFocus) = HWND_FRMFINDINFILES then
      Dim As HWnd hCtrl = GetNextDlgTabItem( HWND_FRMFINDINFILES, GetFocus, false )
      SetFocus(hCtrl)
   else
      ' Determine if a Snippet can be inserted. This takes precedence over doing a block indent.
      if frmSnippets_DoInsertSnippet( pDoc ) = false then
         ' No snippet to insert, do a block indent instead.
         gApp.SuppressNotify = true
         SetWindowRedraw( hEdit, false )
         Function = SciExec(hEdit, SCI_TAB, 0, 0)
         SetWindowRedraw( hEdit, true )
         gApp.SuppressNotify = false
         AfxRedrawWindow( hEdit )
      end if
      frmMain_SetStatusbar
   end if

   function = 0
end function


' ========================================================================================
' Edit action UnIndent Block
' ========================================================================================
function OnCommand_EditUnIndentBlock( _
            byval pDoc as clsDocument ptr, _
            ByVal hEdit as HWND _
            ) as LRESULT

   if (pDoc->IsDesigner) andalso (IsDesignerView(pDoc)) then exit function
   
   if GetFocus = GetDlgItem(HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_TXTFIND) then
      if gFind.bExpanded then SetFocus GetDlgItem(HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_TXTREPLACE)
   elseif GetFocus = GetDlgItem(HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_TXTREPLACE) then
      if gFind.bExpanded then SetFocus GetDlgItem(HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_TXTFIND)
   elseif AfxGetFormHandle(GetFocus) = HWND_FRMFINDINFILES then
      Dim As HWnd hCtrl = GetNextDlgTabItem( HWND_FRMFINDINFILES, GetFocus, true )
      SetFocus(hCtrl)
   else
      gApp.SuppressNotify = true
      SetWindowRedraw( hEdit, false )
      Function = SciExec(hEdit, SCI_BACKTAB, 0, 0)
      SetWindowRedraw( hEdit, true )
      gApp.SuppressNotify = false
      AfxRedrawWindow( hEdit )
      frmMain_SetStatusbar
   end if

   function = 0
end function


' ========================================================================================
' Edit action Select All
' ========================================================================================
function OnCommand_EditSelectAll( _
            byval pDoc as clsDocument ptr, _
            ByVal hEdit as HWND _
            ) as LRESULT

   ' Check if the Compiler log file or Notes windows or Snippets have
   ' input focus. If they do then select that text rather than Scintilla.
   select case GetFocus()
      case GetDlgItem(HWND_FRMOUTPUT, IDC_FRMOUTPUT_TXTNOTES), _
           GetDlgItem(HWND_FRMOUTPUT, IDC_FRMOUTPUT_TXTLOGFILE), _ 
           GetDlgItem(HWND_FRMSNIPPETS, IDC_FRMSNIPPETS_TXTCODE), _
           GetDlgItem(HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_TXTFIND), _
           GetDlgItem(HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_TXTREPLACE), _
           HWND_PROPLIST_EDIT
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
                             
   function = 0
end function


' ========================================================================================
' Edit action File Encoding
' ========================================================================================
function OnCommand_EditEncoding( _
            ByVal id as long, _
            byval pDoc as clsDocument ptr _
            ) as LRESULT

   select case id      
      case IDM_ANSI
         if pDoc->FileEncoding <> FILE_ENCODING_ANSI THEN 
            pDoc->UserModified = true
            ' Convert buffer to ANSI and redisplay the text in the editor
            ConvertTextBuffer(pDoc, FILE_ENCODING_ANSI)
         end if   
         pDoc->FileEncoding = FILE_ENCODING_ANSI 
      
      case IDM_UTF8BOM
         if pDoc->FileEncoding <> FILE_ENCODING_UTF8_BOM THEN 
            pDoc->UserModified = true
            ConvertTextBuffer(pDoc, FILE_ENCODING_UTF8_BOM)
         end if   
         pDoc->FileEncoding = FILE_ENCODING_UTF8_BOM 
      
      case IDM_UTF16BOM
         if pDoc->FileEncoding <> FILE_ENCODING_UTF16_BOM THEN 
            pDoc->UserModified = true
            ConvertTextBuffer(pDoc, FILE_ENCODING_UTF16_BOM)
         end if   
   end select
   
   frmMain_SetStatusbar

   function = 0
end function


' ========================================================================================
' Edit action Find
' ========================================================================================
function OnCommand_EditFindDialog() as LRESULT
   frmFindReplace_Show( HWND_FRMMAIN, false )
   function = 0
end function

' ========================================================================================
' Edit action Replace
' ========================================================================================
function OnCommand_EditReplaceDialog() as LRESULT
   frmFindReplace_Show( HWND_FRMMAIN, true )
   function = 0
end function

' ========================================================================================
' Edit action Find in Files
' ========================================================================================
function OnCommand_EditFindInFiles( byval hEdit as HWND ) as LRESULT
   if hEdit = 0 then hEdit = HWND_FRMMAIN   ' possible no files are open
   frmFindInFiles_Show( hEdit )
   function = 0
end function
      
' ========================================================================================
' Edit action Find actions
' ========================================================================================
function OnCommand_EditFindActions( _
            ByVal id as long, _
            byval pDoc as clsDocument ptr _
            ) as LRESULT

   dim as HWND hEdit = pDoc->hWndActiveScintilla

   select case id
      case IDM_FINDPREV
         if (GetForeGroundWindow = HWND_FRMFINDREPLACE) then
            if gFind.foundCount = 0 then exit function
            dim as long startPos = SendMessage( hEdit, SCI_GETANCHOR, 0, 0) 
            if frmFindReplace_NextSelection(startPos, false, true) then pDoc->CenterCurrentLine
            AfxRedrawWindow( HWND_FRMFINDREPLACE )
         end if

      case IDM_FINDNEXT
         if (GetForeGroundWindow = HWND_FRMFINDREPLACE) then
            if gFind.foundCount = 0 then exit function
            dim as long startPos = SendMessage( hEdit, SCI_GETCURRENTPOS, 0, 0) 
            if frmFindReplace_NextSelection(startPos, true, true) then pDoc->CenterCurrentLine
            AfxRedrawWindow( HWND_FRMFINDREPLACE )
         end if

      Case IDM_FINDNEXTACCEL, IDM_FINDPREVACCEL   ' F3/Shift+F3
         dim as long startPos, endPos, curPos, mainSel, r
         Dim As String sFindText = pDoc->GetSelText
         
         if len(sFindText) = 0 then
            if len(gFind.txtFind) then 
               sFindText = gFind.txtFind
            end if   
         end if
         gFind.txtFind = sFindText
         
         if len(sFindText) then
            curPos = SciExec( hEdit, SCI_GETCURRENTPOS, 0, 0) 
            if id = IDM_FINDNEXTACCEL then
               startPos = curPos + 1
               endPos = SciExec( hEdit, SCI_GETTEXTLENGTH, 0, 0)
               SciExec( hEdit, SCI_SETTARGETSTART, startPos, 0)
               SciExec( hEdit, SCI_SETTARGETEND, endPos, 0)
               ' Search for the text to find
               r = SciExec( hEdit, SCI_SEARCHINTARGET, Len(sFindText), Strptr(sFindText))
               If r = -1 Then
                  ' No match found. Do search starting from the beginning of the file.
                  SciExec( hEdit, SCI_SETTARGETSTART, 0, 0)
                  r = SciExec( hEdit, SCI_SEARCHINTARGET, Len(sFindText), Strptr(sFindText))
               end if
               If r <> -1 Then 
                  SciExec( hEdit, SCI_SETSEL, r, r + len(sFindText)) 
                  pDoc->CenterCurrentLine
               end if
               
            elseif id = IDM_FINDPREVACCEL then
               startPos = curPos - 1
               endPos = 0
               SciExec( hEdit, SCI_SETTARGETSTART, startPos, 0)
               SciExec( hEdit, SCI_SETTARGETEND, endPos, 0)
               ' Search for the text to find
               r = SciExec( hEdit, SCI_SEARCHINTARGET, Len(sFindText), Strptr(sFindText))
               If r = -1 Then
                  ' No match found. Do search starting from the end of the file.
                  startPos = SciExec( hEdit, SCI_GETTEXTLENGTH, 0, 0)
                  SciExec( hEdit, SCI_SETTARGETSTART, startPos, 0)
                  r = SciExec( hEdit, SCI_SEARCHINTARGET, Len(sFindText), Strptr(sFindText))
               end if
               If r <> -1 Then 
                  SciExec( hEdit, SCI_SETSEL, r, r + len(sFindText)) 
                  pDoc->CenterCurrentLine
               end if
            end if
         end if
   
   end select
   
   function = 0
end function


' ========================================================================================
' Edit action Replace actions
' ========================================================================================
function OnCommand_EditReplaceActions( _
            ByVal id as long, _
            byval pDoc as clsDocument ptr _
            ) as LRESULT
         
   dim as HWND hEdit = pDoc->hWndActiveScintilla

   select case id
      Case IDM_REPLACENEXT, IDM_REPLACEPREV, IDM_REPLACEALL
         if (GetForeGroundWindow = HWND_FRMFINDREPLACE) then
            if gFind.foundCount = 0 then exit function
            if id = IDM_REPLACEPREV then 
               frmFindReplace_DoReplace( false, false )
            elseif id = IDM_REPLACENEXT then 
               frmFindReplace_DoReplace( false, true )
            elseif id = IDM_REPLACEALL then 
               frmFindReplace_DoReplace( true, true )
            end if   
            frmFindReplace_HighlightSearches( true )
            pDoc->CenterCurrentLine
            AfxRedrawWindow( HWND_FRMFINDREPLACE )
         end if
   end select
   
   function = 0
end function


' ========================================================================================
' Edit action Common actions
' ========================================================================================
function OnCommand_EditCommon( _
            ByVal id as long, _
            byval pDoc as clsDocument ptr _
            ) as LRESULT
      
   dim as HWND hEdit = pDoc->hWndActiveScintilla

   select case id
      Case IDM_REDO:                SciExec( hEdit, SCI_REDO, 0, 0)
      Case IDM_UNDO, WM_UNDO:       SciExec( hEdit, SCI_UNDO, 0, 0)
      Case WM_UNDO:                 SciExec( GetFocus(), WM_UNDO, 0, 0)
      Case IDM_DELETELINE:          SciExec(hEdit, SCI_LINECUT, 0, 0)
      Case IDM_INSERTFILE:          pDoc->InsertFile()
      Case IDM_COMMENTBLOCK:        pDoc->BlockComment(True) 
      Case IDM_UNCOMMENTBLOCK:      pDoc->BlockComment(False) 
      Case IDM_DUPLICATELINE:       pDoc->LineDuplicate()
      Case IDM_MOVELINEUP:          pDoc->MoveCurrentLines(False) 
      Case IDM_MOVELINEDOWN:        pDoc->MoveCurrentLines(True)  
      Case IDM_NEWLINEBELOWCURRENT: pDoc->NewLineBelowCurrent()  
      Case IDM_TOUPPERCASE:         pDoc->ChangeSelectionCase(1)
      Case IDM_TOLOWERCASE:         pDoc->ChangeSelectionCase(2)
      Case IDM_TOMIXEDCASE:         pDoc->ChangeSelectionCase(3)
      Case IDM_EOLTOCRLF:           pDoc->ConvertEOL(SC_EOL_CRLF)
      Case IDM_EOLTOCR:             pDoc->ConvertEOL(SC_EOL_CR)
      Case IDM_EOLTOLF:             pDoc->ConvertEOL(SC_EOL_LF)
      Case IDM_TABSTOSPACES:        pDoc->TabsToSpaces()
      Case IDM_SELECTLINE:          pDoc->SelectLine(-1)
   end select
   
   function = 0
end function



