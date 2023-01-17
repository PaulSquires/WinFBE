'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2023 Paul Squires, PlanetSquires Software
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

#include once "frmFindReplace.bi"

' ========================================================================================
' Determine if the incoming string is uppercase
' ========================================================================================
function isUpperCaseString( byval sText as string ) as boolean
   ' Default that the string is uppercase
   function = true
   for i as long = 0 to len(sText) - 1
      if isupper( sText[i] ) = 0 then return false
   next
end function

' ========================================================================================
' Determine if the incoming string is lowercase
' ========================================================================================
function isLowerCaseString( byval sText as string ) as boolean
   ' Default that the string is lowercase
   function = true
   for i as long = 0 to len(sText) - 1
      if islower( sText[i] ) = 0 then return false
   next
end function

' ========================================================================================
' frmFindReplace_PositionWindows
' ========================================================================================
function frmFindReplace_PositionWindows() as LRESULT

   ' make sure that the shadows under the top tabs are showing and if the main form
   ' has moved, that the shadows move with it. 
   frmTopTabs_ShowShadow

   if IsWindow( HWND_FRMFINDREPLACE ) = 0 then exit function

   dim pWindow as CWindow ptr = AfxCWindowPtr( HWND_FRMFINDREPLACE )
   if pWindow = 0 then exit function

   pWindow->SetClientSize( 420, iif(gFind.bExpanded, 66, 33) )

   ' find/replace is a popup window so we need to use screen coordinates. Place the 
   ' window just below the toptabs or the menubar if the toptabs are not showing.
   ' Place the window the width of 2 scrollbars from the right of the main window.
   dim as hwnd hCtrl = iif( IsWindowVisible(HWND_FRMMAIN_TOPTABS), _
                              HWND_FRMMAIN_TOPTABS, HWND_FRMMAIN_MENUBAR )  
   dim as RECT rc = AfxGetWindowRect( hCtrl )
   SetWindowPos( HWND_FRMFINDREPLACE_SHADOW, HWND_TOP, _
                  rc.right - AfxGetWindowWidth(HWND_FRMFINDREPLACE) - _
                      pWindow->ScaleX(SCROLLBAR_WIDTH_EDITOR*3) - pWindow->ScaleX(3), _
                  rc.bottom, _
                  AfxGetWindowWidth(HWND_FRMFINDREPLACE) + pWindow->ScaleY(6), _
                  AfxGetWindowHeight(HWND_FRMFINDREPLACE) + pWindow->ScaleY(3), _
                  SWP_SHOWWINDOW )
   SetWindowPos( HWND_FRMFINDREPLACE, HWND_TOP, _
                  rc.right - AfxGetWindowWidth(HWND_FRMFINDREPLACE) - pWindow->ScaleX(SCROLLBAR_WIDTH_EDITOR * 3), _
                  rc.bottom, 0, 0, SWP_SHOWWINDOW or SWP_NOSIZE )

   dim as long frmWidth   = AfxGetWindowWidth( HWND_FRMFINDREPLACE )
   dim as long frmHeight  = AfxGetWindowHeight( HWND_FRMFINDREPLACE )
   dim as long iconWidth  = pWindow->ScaleX(24)
   dim as long iconHeight = pWindow->ScaleY(24)
   dim as long textBoxHeight = pWindow->ScaleY(24)
   dim as long hmargin    = pWindow->ScaleX(2)
   dim as long vmargin    = pWindow->ScaleX(6)
   dim as long nLeft      = pWindow->ScaleY(2)    ' account for the LeftEdge border
   dim as long nTop       = 0

   ' Set the RECT values for each element that we draw in the form. We use these RECTs for
   ' hot hittests and for clicking hittests.
   ' Each panel (find panel and replace panel) total height of 66px

   ' calculate the expand/collapse button
   gFind.rcExpand.Left = nLeft
   gFind.rcExpand.top = 0
   gFind.rcExpand.right = gFind.rcExpand.Left + iconWidth
   gFind.rcExpand.bottom = gFind.rcExpand.top + frmHeight
   
   ' Display the Find textbox               
   nTop  = vmargin
   nLeft = gFind.rcExpand.Right + hmargin
   SetWindowPos( GetDlgItem(HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_TXTFIND), 0, _
                  nLeft, nTop, 0, 0, SWP_SHOWWINDOW or SWP_NOSIZE )

   ' Match Case RECT
   gFind.rcMatchCase.Left = nLeft + _
      AfxGetWindowWidth( GetDlgItem(HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_TXTFIND) )
   gFind.rcMatchCase.top = nTop
   gFind.rcMatchCase.right = gFind.rcMatchCase.Left + iconWidth
   gFind.rcMatchCase.bottom = gFind.rcMatchCase.top + iconWidth

   ' Whole Word RECT
   gFind.rcWholeWord.Left = gFind.rcMatchCase.right
   gFind.rcWholeWord.top = nTop
   gFind.rcWholeWord.right = gFind.rcWholeWord.Left + iconWidth
   gFind.rcWholeWord.bottom = gFind.rcWholeWord.top + iconWidth

   ' Results from search eg. "1 of 10"
   gFind.rcResults.Left = gFind.rcWholeWord.right + hmargin + hmargin
   gFind.rcResults.top = nTop
   gFind.rcResults.right = gFind.rcResults.Left + pWindow->ScaleX(70)
   gFind.rcResults.bottom = gFind.rcResults.top + pWindow->ScaleY(24)

   ' UpArrow RECT
   gFind.rcUpArrow.Left = gFind.rcResults.right + hmargin
   gFind.rcUpArrow.top = nTop
   gFind.rcUpArrow.right = gFind.rcUpArrow.Left + iconWidth
   gFind.rcUpArrow.bottom = gFind.rcUpArrow.top + iconWidth

   ' DownArrow RECT
   gFind.rcDownArrow.Left = gFind.rcUpArrow.right + hmargin
   gFind.rcDownArrow.top = nTop
   gFind.rcDownArrow.right = gFind.rcDownArrow.Left + iconWidth
   gFind.rcDownArrow.bottom = gFind.rcDownArrow.top + iconWidth

   ' Selection RECT
   gFind.rcSelection.Left = gFind.rcDownArrow.right + hmargin
   gFind.rcSelection.top = nTop
   gFind.rcSelection.right = gFind.rcSelection.Left + iconWidth
   gFind.rcSelection.bottom = gFind.rcSelection.top + iconWidth

   ' Close RECT
   gFind.rcClose.Left = gFind.rcSelection.right + hmargin
   gFind.rcClose.top = nTop
   gFind.rcClose.right = gFind.rcClose.Left + iconWidth
   gFind.rcClose.bottom = gFind.rcClose.top + iconWidth

   ' Display the Replace textbox               
   nTop  = vmargin + textBoxHeight + vmargin
   nLeft = gFind.rcExpand.Right + hmargin
   SetWindowPos( GetDlgItem(HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_TXTREPLACE), 0, _
                  nLeft, nTop, 0, 0, iif(gFind.bExpanded, SWP_SHOWWINDOW, SWP_HIDEWINDOW) or SWP_NOSIZE )

   ' Preserve Case RECT
   gFind.rcPreserve.Left = nleft + _
      AfxGetWindowWidth(GetDlgItem(HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_TXTREPLACE))
   gFind.rcPreserve.top = nTop
   gFind.rcPreserve.right = gFind.rcPreserve.Left + iconWidth
   gFind.rcPreserve.bottom = gFind.rcPreserve.top + iconWidth
 
   ' Replace RECT
   gFind.rcReplace.Left = gFind.rcPreserve.right + hmargin + hmargin
   gFind.rcReplace.top = nTop
   gFind.rcReplace.right = gFind.rcReplace.Left + iconWidth
   gFind.rcReplace.bottom = gFind.rcReplace.top + iconWidth

   ' ReplaceAll RECT
   gFind.rcReplaceAll.Left = gFind.rcReplace.right 
   gFind.rcReplaceAll.top = nTop
   gFind.rcReplaceAll.right = gFind.rcReplaceAll.Left + iconWidth
   gFind.rcReplaceAll.bottom = gFind.rcReplaceAll.top + iconWidth

   function = 0
end function


' ========================================================================================
' Find next/prev selection based on current document position
' ========================================================================================
function frmFindReplace_NextSelection( _
      byval StartPos as long, _
      byval bGetnext as boolean, _
      byval bRepositionCaret as boolean _
      ) as boolean

   dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr()
   if pDoc = 0 then exit function

   dim as long mainIdx, lenFind, endPos, nPos, iStart, iEnd
   
   dim as hwnd hEdit = pDoc->hWndActiveScintilla
   dim as hwnd hWndFind = GetDlgItem(HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_TXTFIND)
   dim txtFind as string = str(AfxGetWindowText(hWndFind))
   lenFind = len(txtFind)
   endPos = SciExec( hEdit, SCI_GETTEXTLENGTH, 0, 0)
   if endPos = 0 then exit function
    
   ' Build an array holding all of the selection positions
   redim iPositions(100) as long
   
   gFind.foundCount = 0 
   while nPos <= endPos 
      iStart = SciExec( hEdit, SCI_INDICATORSTART, 8, nPos) 
      iEnd   = SciExec( hEdit, SCI_INDICATOREND, 8, iStart)
      if (iStart = 0) andalso (iEnd = 0) then exit while    ' no indicators found

      if SciExec( hEdit, SCI_INDICATORVALUEAT, 8, iStart-1) then
         gFind.foundCount = gFind.foundCount + 1
         if gFind.foundCount > ubound(iPositions) then
            redim preserve iPositions(gFind.foundCount) as long 
         end if
         iPositions(gFind.foundCount) = iStart - lenFind 
      end if
      nPos = iEnd + 1
   Wend    
   
   ' Determine the next/prev selection that we should position to.
   if bGetnext then
      ' next selection
      mainIdx = 1
      for i as long = 1 to gFind.foundCount
         if iPositions(i) > startPos then
            mainIdx = i: exit for
         end if
      next
   else
      ' Previous selection 
      mainIdx = gFind.foundCount
      for i as long = 1 to gFind.foundCount
         if iPositions(i) < startPos then
            mainIdx = i
         end if
      next
   end if

   ' Highlight the selection
   if gFind.foundCount > 0 then
      ' make sure target line is unfolded
      if bRepositionCaret then
         dim as long currentLine = SciExec( hEdit, SCI_LINEFROMPOSITION, iPositions(mainIdx), 0)
         SciExec( hEdit, SCI_ENSUREVISIBLE, currentLine, 0) 
         SciExec( hEdit, SCI_SETSEL, iPositions(mainIdx), iPositions(mainIdx) + lenFind)
      end if
      gFind.wszResults = mainIdx & " of " & gFind.foundCount
      return true
   else
      SciExec( hEdit, SCI_SETSEL, startPos, startPos)
      gFind.wszResults = L(269, "No results")
      return false
   end if
end function


' ========================================================================================
' Replace current selection or all selections
' ========================================================================================
function frmFindReplace_DoReplace( _
      byval fReplaceAll as boolean = false, _
      byval fMovenext as boolean = true _
      ) as long

   dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr()
   if pDoc = 0 then exit function

   dim as hwnd   hWndFind    = GetDlgItem(HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_TXTFIND)
   dim as hwnd   hWndReplace = GetDlgItem(HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_TXTREPLACE)
   dim as string txtFind     = str(AfxGetWindowText(hWndFind))
   dim as string txtReplace  = str(AfxGetWindowText(hWndReplace))
   dim as long   lenFind     = len(txtFind)

   if lenFind = 0 then exit function
   
   dim as long nPos, startPos, endPos, findFlags, r
    
   dim as hwnd hEdit = pDoc->hWndActiveScintilla
   if fReplaceAll = true then 
      if gFind.nMatchCase then findFlags = (findFlags or SCFIND_MATCHCASE)
      if gFind.nWholeWord then findFlags = (findFlags or SCFIND_WHOLEWORD)
      SendMessage( hEdit, SCI_SETSEARCHFLAGS, findFlags, 0)
       
      if pDoc->HasMarkerHighlight then
         startPos = SciExec( hEdit, SCI_POSITIONFROMLINE, pDoc->FirstMarkerHighlight, 0)
         endPos = SciExec( hEdit, SCI_GETLINEENDPOSITION, pDoc->LastMarkerHighlight, 0)
         SciExec( hEdit, SCI_SETTARGETSTART, startPos, 0)
         SciExec( hEdit, SCI_SETTARGETEND, endPos, 0)
      else
         SciExec( hEdit, SCI_TARGETWHOLEDOCUMENT, 0, 0)
         startPos = SciExec( hEdit, SCI_GETTARGETSTART, 0, 0)
         endPos = SciExec( hEdit, SCI_GETTARGETEND, 0, 0)
      end if
      SciExec( hEdit, SCI_INDICATORCLEARRANGE, startPos, endPos)
      
      gApp.SuppressNotify = true
      SetWindowRedraw( hEdit, false )
      SendMessage( hEdit, SCI_BEGINUNDOACTION, 0, 0)
      do 
         r = SciExec( hEdit, SCI_SEARCHINTARGET, len(txtFind), strptr(txtFind))
         if r = -1 then exit do
         dim as zstring * MAX_PATH sSelText
         SciExec( hEdit, SCI_GETTARGETTEXT, 0, @sSelText )
         ' If Preserve Case is active then check the selection to see if the entire
         ' selection is all UPPERCASE or all lowercase.
         if gFind.nPreserve then
            dim as string sReplaceText
            if isUpperCaseString( sSelText ) then
               sReplaceText = ucase(txtReplace)
            elseif isLowerCaseString( sSelText ) then
               sReplaceText = lcase(txtReplace)
            else
               sReplaceText = txtReplace
            end if
            SciExec( hEdit, SCI_REPLACETARGET, len(sReplaceText), strptr(sReplaceText) )
            startPos = r + len(sReplaceText)
            endPos = endPos + len(sReplaceText) - len(sSelText)
         else            
            SciExec( hEdit, SCI_REPLACETARGET, len(txtReplace), strptr(txtReplace) )
            startPos = r + len(txtReplace)  
            endPos = endPos + len(txtReplace) - len(sSelText)
         end if
         ' Need to update the end searching positions because the replace action may
         ' have made the document larger or smaller.
         SciExec( hEdit, SCI_SETTARGETSTART, startPos, 0)
         SciExec( hEdit, SCI_SETTARGETEND, endPos, 0)
      loop
      SendMessage( hEdit, SCI_ENDUNDOACTION, 0, 0)
      SetWindowRedraw( hEdit, true )
      gApp.SuppressNotify = false
      AfxRedrawWindow( hEdit )

   else   
      ' Need to ensure that the current cursor/selection is actually the find
      ' phrase that we want to replace. if not, then move to the closest selection.
      if ucase(pDoc->GetSelText) <> ucase(txtFind) then
         startPos = SciExec( hEdit, SCI_GETCURRENTPOS, 0, 0) 
         frmFindReplace_NextSelection(startPos, fMoveNext, true)
         exit function
      end if
      SciExec( hEdit, SCI_TARGETFROMSELECTION, 0, 0)
      ' If Preserve Case is active then check the selection to see if the entire
      ' selection is all UPPERCASE or all lowercase.
      if gFind.nPreserve then
         dim as string sSelText = pDoc->GetSelText
         dim as string sReplaceText
         if isUpperCaseString( sSelText ) then
            sReplaceText = ucase(txtReplace)
         elseif isLowerCaseString( sSelText ) then
            sReplaceText = lcase(txtReplace)
         else
            sReplaceText = txtReplace
         end if
         SciExec( hEdit, SCI_REPLACETARGET, len(sReplaceText), strptr(sReplaceText) )
         startPos = SciExec( hEdit, SCI_GETCURRENTPOS, 0, 0) + len(sReplaceText)
      else            
         SciExec( hEdit, SCI_REPLACETARGET, len(txtReplace), strptr(txtReplace) )
         startPos = SciExec( hEdit, SCI_GETCURRENTPOS, 0, 0) + len(txtReplace)
      end if
      frmFindReplace_NextSelection(startPos, fMoveNext, true)
   end if
      
   ' Need to reparse because some of the function names may been replaced
   ' with new names.
   pDoc->bNeedsParsing = true
   pDoc->ParseDocument()

   function = 0
end function


' ========================================================================================
' Highlight the found selections
' ========================================================================================
function frmFindReplace_HighlightSearches( byval bRepositionCaret as boolean ) as long

   dim as long r, startPos, endPos, findFlags, nLength
   dim as hwnd hWndFind = GetDlgItem(HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_TXTFIND)
   
   dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr()
   if pDoc = 0 then exit function

   dim as hwnd hEdit = pDoc->hWndActiveScintilla
   SciExec( hEdit, SCI_SETADDITIONALCARETSVISIBLE, false, 0) 
   SciExec( hEdit, SCI_INDICSETSTYLE, 8, INDIC_STRAIGHTBOX)
   SciExec( hEdit, SCI_SETINDICATORCURRENT, 8, 0)
   SciExec( hEdit, SCI_INDICSETFORE, 8, ghEditor.ForeColorOccurrence ) 
   SciExec( hEdit, SCI_INDICSETALPHA, 8, 60 )   
   
   dim txtFind as string = str(AfxGetWindowText(hWndFind))
      
   if gFind.nMatchCase then findFlags = (findFlags or SCFIND_MATCHCASE)
   if gFind.nWholeWord then findFlags = (findFlags or SCFIND_WHOLEWORD)
   SciExec( hEdit, SCI_SETSEARCHFLAGS, findFlags, 0)
   
   ' Remove all existing selection indicators
   nLength = SendMessage( hEdit, SCI_GETTEXTLENGTH, 0, 0)
   SciExec( hEdit, SCI_INDICATORCLEARRANGE, 0, nLength)
   
   ' Are we searching a selection or the whole document
   if pDoc->HasMarkerHighlight then
      startPos = SciExec( hEdit, SCI_POSITIONFROMLINE, pDoc->FirstMarkerHighlight, 0)
      endPos = SciExec( hEdit, SCI_GETLINEENDPOSITION, pDoc->LastMarkerHighlight, 0)
      SciExec( hEdit, SCI_SETTARGETSTART, startPos, 0)
      SciExec( hEdit, SCI_SETTARGETEND, endPos, 0)
   else
      SciExec( hEdit, SCI_TARGETWHOLEDOCUMENT, 0, 0)
      startPos = SciExec( hEdit, SCI_GETTARGETSTART, 0, 0)
      endPos = SciExec( hEdit, SCI_GETTARGETEND, 0, 0)
   end if

   ' Search for the text to find
   if len(txtFind) then
      do 
         r = SciExec( hEdit, SCI_SEARCHINTARGET, len(txtFind), strptr(txtFind))
         if r = -1 then exit do
         
         SciExec( hEdit, SCI_SETINDICATORVALUE, 8, 0 )
         SciExec( hEdit, SCI_INDICATORFILLRANGE, r, len(txtFind))
         startPos = r + len(txtFind)
         
         ' Adjust the searching positions
         SciExec( hEdit, SCI_SETTARGETSTART, startPos, 0)
         SciExec( hEdit, SCI_SETTARGETEND, endPos, 0)
      loop
      startPos = SciExec( hEdit, SCI_GETANCHOR, 0, 0) 
      endPos = SciExec( hEdit, SCI_GETCURRENTPOS, 0, 0) 
      if endPos < startPos then startPos = endPos
      startPos = startPos - 1
   else
      startPos = SciExec( hEdit, SCI_GETANCHOR, 0, 0) 
   end if

   startPos = iif(startPos < 0, 0, startPos)
   frmFindReplace_NextSelection(startPos, true, bRepositionCaret)
   
   function = 0

end function


' ========================================================================================
' Process WM_COMMAND message for window/dialog: frmFindReplace
' ========================================================================================
function frmFindReplace_OnCommand( _
      byval HWnd as HWnd, _
      byval id as Long, _
      byval hwndCtl as HWnd, _
      byval codeNotify as UINT _
      ) as LRESULT

   dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr()
   if pDoc = 0 then exit function

   dim as hwnd hEdit = pDoc->hWndActiveScintilla

   select case id

      case IDC_FRMFINDREPLACE_TXTFIND
         if (codeNotify = EN_CHANGE) then
            gFind.txtFind = AfxGetWindowText(hwndCtl)
            frmFindReplace_HighlightSearches( true )
            SetFocus( hwndCtl )   ' in case search stole focus from textbox
            AfxRedrawWindow( hWnd )
         elseif (codeNotify = EN_SETFOCUS) then
            AfxRedrawWindow( hWnd )
         end if
             
      case IDC_FRMFINDREPLACE_TXTREPLACE
         if (codeNotify = EN_SETFOCUS) then
            AfxRedrawWindow( hWnd )
         end if

   end select
   
   function = 0
end function


' ========================================================================================
' Process WM_PAINT message for window/dialog: frmFindReplace
' ========================================================================================
function frmFindReplace_OnPaint( byval HWnd as HWND) as LRESULT
            
   dim as PAINTSTRUCT ps
   dim as HPEN hPen
   dim as HWND hCtl
   dim as HDC hDc
   dim as Rect rc
   
   dim as long wsStyle
   
   dim pWindow as CWindow ptr = AfxCWindowOwnerPtr(HWND)
   if pWindow = 0 then exit function

   hDC = BeginPaint( hWnd, @ps )

   SaveDC( hDC )

   dim as long nWidth = ps.rcPaint.right - ps.rcPaint.left
   dim as long nHeight = ps.rcPaint.bottom - ps.rcPaint.top

   dim memDC as HDC      ' Double buffering
   dim hbit as HBITMAP   ' Double buffering

   memDC = CreateCompatibleDC( hDC )
   hbit  = CreateCompatibleBitmap( hDC, nWidth, nHeight )
   if hbit then hbit = SelectObject( memDC, hbit )

   SelectObject( memDC, ghFindReplace.hPanelBrush )

   ' Fill in the entire back panel width across the top of the screen
   FillRect( memDC, @ps.rcPaint, ghFindReplace.hPanelBrush )

   ' Paint the left edge
   hPen = CreatePen( PS_SOLID, pWindow->ScaleY(2), ghFindReplace.leftedge ) 
   SelectPen( memDC, hPen )
   MoveToEx( memDC, 0, 0, null )
   LineTo( memDC, 0, ps.rcPaint.bottom )   
   DeletePen( hPen )

   ' Paint the Expand/Collapse button
   if isMouseOverRECT( HWnd, gFind.rcExpand ) then
      SetBkColor( memDC, ghFindReplace.IconBackColorHot )
      SetTextColor( memDC, ghFindReplace.ForeColorHot )
      FillRect( memDC, @gFind.rcExpand, ghFindReplace.hIconBrushHot )
   else
      SetBkColor( memDC, ghFindReplace.BackColor )
      SetTextColor( memDC, ghFindReplace.ForeColor )
      FillRect( memDC, @gFind.rcExpand, ghFindReplace.hBackBrush )
  end if
   SelectObject( memDC, ghMenuBar.hFontSymbol )
   dim as CWSTR wszText = iif( gFind.bExpanded, wszChevronDown, wszChevronRight )
   wsStyle = DT_NOPREFIX or DT_CENTER or DT_VCENTER or DT_SINGLELINE
   DrawText( memDC, wszText.sptr, -1, cast(lpRect, @gFind.rcExpand), wsStyle )

   ' Paint a rectangle around the textbox control that has focus.
   hPen = CreatePen( PS_SOLID, 1, BGR(0,0,255) )  ' blue
   SelectPen( memDC, hPen )
   hCtl = GetFocus()
   if hCtl = GetDlgItem(HWND, IDC_FRMFINDREPLACE_TXTFIND) then
      SelectObject( hDC, ghFindReplace.hBackBrushTextBox )
      GetWindowRect( GetDlgItem(HWND, IDC_FRMFINDREPLACE_TXTFIND), @rc )
      MapWindowPoints( HWND_DESKTOP, HWND, cast(LPPOINT, @rc), 2)
      Rectangle( memDC, rc.left - 1, rc.top - 1, rc.right + pWindow->ScaleX(49), rc.bottom + 1 )
   elseif hCtl = GetDlgItem(HWND, IDC_FRMFINDREPLACE_TXTREPLACE) then
      SelectObject( hDC, ghFindReplace.hBackBrushTextBox )
      GetWindowRect( GetDlgItem(HWND, IDC_FRMFINDREPLACE_TXTREPLACE), @rc )
      MapWindowPoints( HWND_DESKTOP, HWND, cast(LPPOINT, @rc), 2)
      Rectangle( memDC, rc.left - 1, rc.top - 1, rc.right + pWindow->ScaleX(25), rc.bottom + 1 )
  end if
   DeleteObject( hPen )
   
   ' Match Case
   hPen = CreatePen( PS_NULL, 1, 0 )  ' null/invisible pen
   SelectPen( memDC, hPen )
   if gFind.nMatchCase then 
      SetBkColor( memDC, ghFindReplace.IconSelectedBackColor )
      SetTextColor( memDC, ghFindReplace.ForeColor )
      SelectObject( memDC, ghFindReplace.hBackBrushIconSelected )
   elseif isMouseOverRECT( HWnd, gFind.rcMatchCase ) then
      SetBkColor( memDC, ghFindReplace.IconBackColorHot )
      SetTextColor( memDC, ghFindReplace.ForeColorHot )
      SelectObject( memDC, ghFindReplace.hIconBrushHot )
   else
      SetBkColor( memDC, ghFindReplace.TextBoxBackColor )
      SetTextColor( memDC, ghFindReplace.ForeColor )
      SelectObject( memDC, ghFindReplace.hBackBrushTextBox )
  end if
   FillRect( memDC, @gFind.rcMatchCase, ghFindReplace.hBackBrushTextBox )
   ' shrink the rectangle in order to fit within the textbox
   rc = gFind.rcMatchCase
   InflateRect( @rc, -(pWindow->ScaleX(2)), -(pWindow->ScaleY(2)) )
   RoundRect( memDC, rc.left, rc.top, rc.right, rc.bottom, 20, 20 )
   SelectObject(memDC, ghStatusBar.hFontStatusBar)
   wsStyle = DT_NOPREFIX or DT_CENTER or DT_VCENTER or DT_SINGLELINE
   DrawText( memDC, wszMatchCase, -1, cast(lpRect, @rc), wsStyle )
   
   ' Whole Word
   if gFind.nWholeWord then 
      SetBkColor( memDC, ghFindReplace.IconSelectedBackColor )
      SetTextColor( memDC, ghFindReplace.ForeColor )
      SelectObject( memDC, ghFindReplace.hBackBrushIconSelected )
   elseif isMouseOverRECT( HWnd, gFind.rcWholeWord ) then
      SetBkColor( memDC, ghFindReplace.IconBackColorHot )
      SetTextColor( memDC, ghFindReplace.ForeColorHot )
      SelectObject( memDC, ghFindReplace.hIconBrushHot )
   else
      SetBkColor( memDC, ghFindReplace.TextBoxBackColor )
      SetTextColor( memDC, ghFindReplace.ForeColor )
      SelectObject( memDC, ghFindReplace.hBackBrushTextBox )
  end if
   rc = gFind.rcWholeWord
   FillRect( memDC, @rc, ghFindReplace.hBackBrushTextBox )
   ' shrink the rectangle in order to fit within the textbox
   InflateRect( @rc, -(pWindow->ScaleX(2)), -(pWindow->ScaleY(2)) )
   RoundRect( memDC, rc.left, rc.top, rc.right, rc.bottom, 20, 20 )
   SelectObject(memDC, ghStatusBar.hFontStatusBar)
   wsStyle = DT_NOPREFIX or DT_CENTER or DT_VCENTER or DT_SINGLELINE
   DrawText( memDC, wszWholeWord, -1, cast(lpRect, @rc), wsStyle )

   ' Paint the Results text
   if gFind.foundCount = 0 then 
      SetTextColor( memDC, ghFindReplace.NotFoundForeColor )
   else
      SetTextColor( memDC, ghFindReplace.ForeColor )
  end if
   SetBkColor( memDC, ghFindReplace.BackColor )
   SelectObject(memDC, ghStatusBar.hFontStatusBar)
   wsStyle = DT_NOPREFIX or DT_LEFT or DT_VCENTER or DT_SINGLELINE
   DrawText( memDC, gFind.wszResults.sptr, -1, cast(lpRect, @gFind.rcResults), wsStyle )
   
   ' Paint the Up Arrow
   if isMouseOverRECT( HWnd, gFind.rcUpArrow ) then
      SetBkColor( memDC, ghFindReplace.IconBackColorHot )
      SetTextColor( memDC, ghFindReplace.ForeColorHot )
      SelectObject( memDC, ghFindReplace.hIconBrushHot )
   else
      SetBkColor( memDC, ghFindReplace.BackColor )
      SetTextColor( memDC, ghFindReplace.ForeColor )
      SelectObject( memDC, ghFindReplace.hBackBrush )
  end if
   rc = gFind.rcUpArrow
   RoundRect( memDC, rc.left, rc.top, rc.right, rc.bottom, 20, 20 )
   SelectObject( memDC, ghMenuBar.hFontSymbol )
   wsStyle = DT_NOPREFIX or DT_CENTER or DT_VCENTER or DT_SINGLELINE
   DrawText( memDC, wszUpArrow, -1, cast(lpRect, @rc), wsStyle )
   
   ' Paint the Down Arrow
   if isMouseOverRECT( HWnd, gFind.rcDownArrow ) then
      SetBkColor( memDC, ghFindReplace.IconBackColorHot )
      SetTextColor( memDC, ghFindReplace.ForeColorHot )
      SelectObject( memDC, ghFindReplace.hIconBrushHot )
   else
      SetBkColor( memDC, ghFindReplace.BackColor )
      SetTextColor( memDC, ghFindReplace.ForeColor )
      SelectObject( memDC, ghFindReplace.hBackBrush )
  end if
   rc = gFind.rcDownArrow
   RoundRect( memDC, rc.left, rc.top, rc.right, rc.bottom, 20, 20 )
   SelectObject( memDC, ghMenuBar.hFontSymbol )
   wsStyle = DT_NOPREFIX or DT_CENTER or DT_VCENTER or DT_SINGLELINE
   DrawText( memDC, wszDownArrow, -1, cast(lpRect, @rc), wsStyle )

   ' Paint the Selection icon
   if gFind.nSelection then 
      SetBkColor( memDC, ghFindReplace.IconSelectedBackColor )
      SetTextColor( memDC, ghFindReplace.ForeColor )
      SelectObject( memDC, ghFindReplace.hBackBrushIconSelected )
   elseif isMouseOverRECT( HWnd, gFind.rcSelection) then
      SetBkColor( memDC, ghFindReplace.IconBackColorHot )
      SetTextColor( memDC, ghFindReplace.ForeColorHot )
      SelectObject( memDC, ghFindReplace.hIconBrushHot )
   else
      SetBkColor( memDC, ghFindReplace.BackColor )
      SetTextColor( memDC, ghFindReplace.ForeColor )
      SelectObject( memDC, ghFindReplace.hBackBrush )
  end if
   rc = gFind.rcSelection
   RoundRect( memDC, rc.left, rc.top, rc.right, rc.bottom, 20, 20 )
   SelectObject( memDC, ghMenuBar.hFontSymbol )
   wsStyle = DT_NOPREFIX or DT_CENTER or DT_VCENTER or DT_SINGLELINE
   DrawText( memDC, wszSelection, -1, cast(lpRect, @rc), wsStyle )

   ' Paint the Close icon
   if isMouseOverRECT( HWnd, gFind.rcClose) then
      SetBkColor( memDC, ghFindReplace.IconBackColorHot )
      SetTextColor( memDC, ghFindReplace.ForeColorHot )
      SelectObject( memDC, ghFindReplace.hIconBrushHot )
   else
      SetBkColor( memDC, ghFindReplace.BackColor )
      SetTextColor( memDC, ghFindReplace.ForeColor )
      SelectObject( memDC, ghFindReplace.hBackBrush )
   end if
   rc = gFind.rcClose
   RoundRect( memDC, rc.left, rc.top, rc.right, rc.bottom, 20, 20 )
   SelectObject( memDC, ghMenuBar.hFontSymbol )
   wsStyle = DT_NOPREFIX or DT_CENTER or DT_VCENTER or DT_SINGLELINE
   DrawText( memDC, wszClose, -1, cast(lpRect, @rc), wsStyle )

   ' Preserve Case
   if gFind.nPreserve then 
      SetBkColor( memDC, ghFindReplace.IconSelectedBackColor )
      SetTextColor( memDC, ghFindReplace.ForeColor )
      SelectObject( memDC, ghFindReplace.hBackBrushIconSelected )
   elseif isMouseOverRECT( HWnd, gFind.rcPreserve ) then
      SetBkColor( memDC, ghFindReplace.IconBackColorHot )
      SetTextColor( memDC, ghFindReplace.ForeColorHot )
      SelectObject( memDC, ghFindReplace.hIconBrushHot )
   else
      SetBkColor( memDC, ghFindReplace.TextBoxBackColor )
      SetTextColor( memDC, ghFindReplace.ForeColor )
      SelectObject( memDC, ghFindReplace.hBackBrushTextBox )
  end if
   FillRect( memDC, @gFind.rcPreserve, ghFindReplace.hBackBrushTextBox )
   ' shrink the rectangle in order to fit within the textbox
   rc = gFind.rcPreserve
   InflateRect( @rc, -(pWindow->ScaleX(2)), -(pWindow->ScaleY(2)) )
   RoundRect( memDC, rc.left, rc.top, rc.right, rc.bottom, 20, 20 )
   SelectObject(memDC, ghStatusBar.hFontStatusBar)
   wsStyle = DT_NOPREFIX or DT_CENTER or DT_VCENTER or DT_SINGLELINE
   DrawText( memDC, wszPreserveCase, -1, cast(lpRect, @rc), wsStyle )
   
   ' Paint the Replace icon
   if isMouseOverRECT( HWnd, gFind.rcReplace) then
      SetBkColor( memDC, ghFindReplace.IconBackColorHot )
      SetTextColor( memDC, ghFindReplace.ForeColorHot )
      SelectObject( memDC, ghFindReplace.hIconBrushHot )
   else
      SetBkColor( memDC, ghFindReplace.BackColor )
      SetTextColor( memDC, ghFindReplace.ForeColor )
      SelectObject( memDC, ghFindReplace.hBackBrush )
  end if
   rc = gFind.rcReplace
   RoundRect( memDC, rc.left, rc.top, rc.right, rc.bottom, 20, 20 )
   SelectObject( memDC, ghMenuBar.hFontSymbol )
   wsStyle = DT_NOPREFIX or DT_CENTER or DT_VCENTER or DT_SINGLELINE
   DrawText( memDC, wszReplace, -1, cast(lpRect, @rc), wsStyle )

   ' Paint the ReplaceAll icon
   if isMouseOverRECT( HWnd, gFind.rcReplaceAll) then
      SetBkColor( memDC, ghFindReplace.IconBackColorHot )
      SetTextColor( memDC, ghFindReplace.ForeColorHot )
      SelectObject( memDC, ghFindReplace.hIconBrushHot )
   else
      SetBkColor( memDC, ghFindReplace.BackColor )
      SetTextColor( memDC, ghFindReplace.ForeColor )
      SelectObject( memDC, ghFindReplace.hBackBrush )
  end if
   rc = gFind.rcReplaceAll
   RoundRect( memDC, rc.left, rc.top, rc.right, rc.bottom, 20, 20 )
   SelectObject( memDC, ghMenuBar.hFontSymbol )
   wsStyle = DT_NOPREFIX or DT_CENTER or DT_VCENTER or DT_SINGLELINE
   DrawText( memDC, wszReplaceAll, -1, cast(lpRect, @rc), wsStyle )

   BitBlt( hDC, 0, 0, nWidth, nHeight, memDC, 0, 0, SRCCOPY )

   ' Cleanup
   if hPen  then DeleteObject( hPen )
   if hbit  then DeleteObject( SelectObject(memDC, hbit) )
   if memDC then DeleteDC( memDC )
   RestoreDC( hDC, -1 )

   EndPaint( hWnd, @ps )
   
   function = true
end function


' ========================================================================================
' Process WM_CLOSE message for window/dialog: frmFindReplace
' ========================================================================================
function frmFindReplace_OnClose( byval HWnd as HWnd ) as LRESULT
   DestroyWindow HWnd
   function = 0
end function


' ========================================================================================
' Process WM_DESTROY message for window/dialog: frmFindReplace
' ========================================================================================
function frmFindReplace_OnDestroy( byval HWnd as HWnd ) as LRESULT 
    
   ' Remove selected attributes (Attribute #8). We need to remove it from all
   ' open documents because the user may have tabbed between documents while
   ' the FindReplace dialog was active thereby causing selection highlights.
   gApp.RemoveAllSelectionAttributes
    
   ' Destroy the "shadow" form
   DestroyWindow( HWND_FRMFINDREPLACE_SHADOW )
   
   ' Remove any markers set in the document that highlights
   ' entire lines (used for Selected text searching).
   dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr()
   if pDoc then 
      pDoc->RemoveMarkerHighlight
      SetFocusScintilla
   end if
      
   HWND_FRMFINDREPLACE = 0

   ' delete the pWindow because it gets recreated everytime frmFindReplace is invoked
   dim pWindow as CWindow ptr = AfxCWindowPtr(HWND)
   if pWindow then delete pWindow
   
   function = 0
end function


' ========================================================================================
' frmFindReplaceShadow_WndProc Window procedure
' ========================================================================================
function frmFindReplaceShadow_WndProc( _
      byval HWnd   as HWnd, _
      byval uMsg   as UINT, _
      byval wParam as WPARAM, _
      byval lParam as LPARAM _
      ) as LRESULT
      
   select case uMsg
      case WM_DESTROY
         dim pWindow as CWindow ptr = AfxCWindowPtr(HWnd)
         if pWindow = 0 then delete pWindow
         HWND_FRMFINDREPLACE_SHADOW = 0
   end select      

   ' for messages that we don't deal with
   function = DefWindowProc(HWnd, uMsg, wParam, lParam)

end function


' ========================================================================================
' frmFindReplace Window procedure
' ========================================================================================
function frmFindReplace_WndProc( _
      byval HWnd   as HWnd, _
      byval uMsg   as UINT, _
      byval wParam as WPARAM, _
      byval lParam as LPARAM _
      ) as LRESULT
   
   static hTooltip as HWND
   
   select case uMsg
      HANDLE_MSG (HWnd, WM_CLOSE,    frmFindReplace_OnClose)
      HANDLE_MSG (HWnd, WM_DESTROY,  frmFindReplace_OnDestroy)
      HANDLE_MSG (HWnd, WM_COMMAND,  frmFindReplace_OnCommand)
      HANDLE_MSG (HWnd, WM_PAINT,    frmFindReplace_OnPaint)

   case WM_ACTIVATE
      AfxRedrawWindow( HWnd )
      
   case WM_ERASEBKGND
      return true

   case WM_MOUSEMOVE
      dim tme as TrackMouseEvent
      tme.cbSize = sizeof(TrackMouseEvent)
      tme.dwFlags = TME_HOVER or TME_LEAVE
      tme.hwndTrack = HWnd
      tme.dwHoverTime = 400    ' system default is 400ms
      TrackMouseEvent(@tme) 

      if IsWindow(hTooltip) = 0 then hTooltip = AfxAddTooltip( HWnd, "", false, false )
      AfxRedrawWindow(HWnd)


   case WM_MOUSELEAVE
      AfxDeleteTooltip( hTooltip, HWnd )
      hTooltip = 0
      AfxRedrawWindow( HWnd )


   case WM_MOUSEHOVER
      dim as CWSTR wszTooltip = ""
      if isMouseOverRECT( HWnd, gFind.rcExpand ) then
         wszTooltip = L(270, "Toggle Replace mode")
      elseif isMouseOverRECT( HWnd, gFind.rcMatchCase ) then
         wszTooltip = L(163,"Match Case")
      elseif isMouseOverRECT( HWnd, gFind.rcWholeWord ) then
         wszTooltip = L(162,"Match Whole Words")
      elseif isMouseOverRECT( HWnd, gFind.rcUpArrow ) then
         wszTooltip = L(45,"Find Previous") & " (Shift+F3)"
      elseif isMouseOverRECT( HWnd, gFind.rcDownArrow ) then
         wszTooltip = L(44,"Find next") & " (F3)"
      elseif isMouseOverRECT( HWnd, gFind.rcSelection ) then
         wszTooltip = L(148,"Selection")
      elseif isMouseOverRECT( HWnd, gFind.rcPreserve ) then
         wszTooltip = L(424,"Preserve Case")
      elseif isMouseOverRECT( HWnd, gFind.rcReplace ) then
         wszTooltip = L(173,"Replace") & " (Enter)"
      elseif isMouseOverRECT( HWnd, gFind.rcReplaceAll ) then
         wszTooltip = L(174,"Replace All") & " (Ctrl+Alt+Enter)"
      elseif isMouseOverRECT( HWnd, gFind.rcClose ) then
         wszTooltip = L(161,"Close") & " (Esc)"
     end if

      ' Display the tooltip
      if len(wszTooltip) then AfxSetTooltipText( hTooltip, HWnd, wszTooltip )
      AfxRedrawWindow( HWnd )


   case WM_RBUTTONDOWN

   case WM_LBUTTONDOWN

   case WM_LBUTTONUP
      dim as hwnd hEdit 
      dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr()
      if pDoc then hEdit = pDoc->hWndActiveScintilla
      
      if isMouseOverRECT( HWnd, gFind.rcExpand ) then
         gFind.bExpanded = not gFind.bExpanded
         frmFindReplace_Show( HWnd, gFind.bExpanded )

      elseif isMouseOverRECT( HWnd, gFind.rcMatchCase ) then
         gFind.nMatchCase = not gFind.nMatchCase
         frmFindReplace_HighlightSearches( true )
         AfxRedrawWindow(HWnd)

      elseif isMouseOverRECT( HWnd, gFind.rcWholeWord ) then
         gFind.nWholeWord = not gFind.nWholeWord
         frmFindReplace_HighlightSearches( true )
         AfxRedrawWindow(HWnd)

      elseif isMouseOverRECT( HWnd, gFind.rcUpArrow ) then
         PostMessage( HWND_FRMMAIN, WM_COMMAND, MAKEWPARAM(IDM_FINDPREV, 0), 0 )

      elseif isMouseOverRECT( HWnd, gFind.rcDownArrow ) then
         PostMessage( HWND_FRMMAIN, WM_COMMAND, MAKEWPARAM(IDM_FINDnext, 0), 0 )

      elseif isMouseOverRECT( HWnd, gFind.rcSelection ) then
         if pDoc = 0 then exit function   
         if (pDoc->IsMultilineSelection) or (pDoc->HasMarkerHighlight) then
            gFind.nSelection = not gFind.nSelection
         else
            gFind.nSelection = false
         end if
         if gFind.nSelection then
            pDoc->SetMarkerHighlight
         else   
            pDoc->RemoveMarkerHighlight
         end if
         frmFindReplace_HighlightSearches( true )
         AfxRedrawWindow(HWnd)

      elseif isMouseOverRECT( HWnd, gFind.rcPreserve ) then
         if pDoc = 0 then exit function   
         gFind.nPreserve = not gFind.nPreserve

      elseif isMouseOverRECT( HWnd, gFind.rcReplace ) then
         if pDoc = 0 then exit function   
         PostMessage( HWND_FRMMAIN, WM_COMMAND, MAKEWPARAM(IDM_REPLACEnext, 0), 0 )

      elseif isMouseOverRECT( HWnd, gFind.rcReplaceAll ) then
         if pDoc = 0 then exit function   
         PostMessage( HWND_FRMMAIN, WM_COMMAND, MAKEWPARAM(IDM_REPLACEALL, 0), 0 )

      elseif isMouseOverRECT( HWnd, gFind.rcClose ) then
         SendMessage( HWnd, WM_CLOSE, 0, 0 )
     end if
         
   End Select

   function = DefWindowProc(HWnd, uMsg, wParam, lParam)

end function


' ========================================================================================
' frmFindReplace_RichEdit_SubclassProc Window procedure
' ========================================================================================
Function frmFindReplace_RichEdit_SubclassProc ( _
      byval hWin   as HWnd, _                 ' // Control window handle
      byval uMsg   as UINT, _                 ' // Type of message
      byval _wParam as WPARAM, _               ' // First message parameter
      byval _lParam as LPARAM, _               ' // Second message parameter
      byval uIdSubclass as UINT_PTR, _        ' // The subclass ID
      byval dwRefData as DWORD_PTR _          ' // Pointer to reference data
      ) as LRESULT

   dim pWindow as CWindow ptr = AfxCWindowPtr(hWin)

   select case uMsg
      
      case WM_PRINTCLIENT, WM_PAINT
         dim as long textLength = AfxGetWindowTextLength(hWin)
         if (textLength = 0 ) then 
            ' Get the needed DC with DCX_INTERSECTUPDATE before the EDIT
            ' control's WM_PAINT handler calls BeginPaint/EndPaint, which
            ' validates the update rect and would otherwise lead to drawing
            ' nothing later because the region is empty. Also, grab it from
            ' the cache so we don't mess with the EDIT's DC.
            dim as HDC hdc = iif(uMsg = WM_PRINTCLIENT, cast(HDC, _wParam), _
                 GetDCEx(hWin, 0, DCX_INTERSECTUPDATE or DCX_CACHE or DCX_CLIPCHILDREN or DCX_CLIPSIBLINGS) )

            ' Call the EDIT control so that the caret is properly handled,
            ' no caret litter left on the screen after tabbing away.
            dim as LRESULT result = DefSubclassProc(hWin, uMsg, _wParam, _lParam)

            ' Get the font and margin so the cue banner text has a
            ' consistent appearance and placement with existing text.
            dim as HFONT font = ghStatusBar.hFontStatusBar 
            dim as Rect editRect: SendMessage( hWin, EM_GETRECT, 0, cast(LPARAM, @editRect) )

            ' Ideally we would call Edit_GetCueBannerText, but since that message
            ' returns nothing when ES_MULTILINE, use a window property instead.
            dim cueBannerText as wstring * 250
            if hWin = GetDlgItem( HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_TXTFIND ) then
               cueBannerText = L(158,"Find")
            elseif hWin = GetDlgItem( HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_TXTREPLACE ) then
               cueBannerText = L(173,"Replace")
            end if

            dim as HFONT previousFont = SelectFont(hdc, font)
            SetTextColor( hDC, ghFindReplace.CueBannerForeColor )
            SetBkMode(hdc, TRANSPARENT)
            dim as long wsStyle = DT_TOP or DT_LEFT or DT_NOPREFIX or DT_NOCLIP
            DrawText(hdc, cueBannerText, -1, @editRect, wsStyle)
            SelectFont(hdc, previousFont)

            RestoreDC( hDC, -1 )
            ReleaseDC(hWin, hdc)

            ' Return the EDIT's result (could probably safely just return zero here,
            ' but seems safer to relay whatever value came from the edit).
            return result
         end if
     
      case WM_CONTEXTMENU 
         ' Create the right click popup menu
         dim as CWSTR wszText = RichEdit_GetSelText( hWin )
         Dim hPopUpMenu As HMENU = CreatePopupMenu()
         if len(wszText) then 
            AppendMenu( hPopUpMenu, MF_ENABLED, IDM_CUT, wstr("Cut") )
            AppendMenu( hPopUpMenu, MF_ENABLED, IDM_COPY, wstr("Copy") )
         end if
         if RichEdit_CanPaste( hWin, 0 ) then
            if len(wszText) then 
               AppendMenu( hPopUpMenu, MF_SEPARATOR, 0, "" )
            end if   
            AppendMenu( hPopUpMenu, MF_ENABLED, IDM_PASTE, wstr("Paste") )
         end if   

         dim as long nResult
         nResult = TrackPopupMenu( hPopUpMenu, TPM_RETURNCMD or TPM_NONOTIFY, _
                            loword(_lParam), hiword(_lParam), 0, HWND_FRMFINDREPLACE, 0 ) 
         select case nResult
            case IDM_CUT:   SendMessage( hWin, WM_CUT, 0, 0 )
            case IDM_COPY:  SendMessage( hWin, WM_COPY, 0, 0 )
            case IDM_PASTE: SendMessage( hWin, WM_PASTE, 0, 0 )
         end select
         DestroyMenu hPopUpMenu
         return 0

      Case WM_DESTROY
         ' REQUIRED: Remove control subclassing
         RemoveWindowSubclass( hWin, @frmFindReplace_RichEdit_SubclassProc, uIdSubclass )
   End Select
    
   ' For messages that we don't deal with
   function = DefSubclassProc(hWin, uMsg, _wParam, _lParam)

end function


' ========================================================================================
' frmFindReplace_Show
' ========================================================================================
function frmFindReplace_Show( _
      byval hWndParent as HWnd, _
      byval fShowReplace as BOOLEAN _
      ) as LRESULT

   '  Create the main window and child controls
   dim pWindow as CWindow ptr = new CWindow
   pWindow->DPI = AfxCWindowOwnerPtr(hwndParent)->DPI

   dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr()
   if pDoc = 0 then exit function
      
   dim as HWND hCtl
   dim as String sFindText
   dim as HWND hCtlFind, hCtlReplace 
   
   gFind.bExpanded = fShowReplace
   
   if IsWindow( HWND_FRMFINDREPLACE ) = 0 then
      gFind.nMatchCase = false
      gFind.nWholeWord = false
      gFind.nSelection = false

      HWND_FRMFINDREPLACE = _
      pWindow->Create(hWndParent, "FINDREPLACE", @frmFindReplace_WndProc, 0, 0, 0, 0, _
           WS_POPUP or WS_CLIPSIBLINGS or WS_CLIPCHILDREN, _   
           WS_EX_LEFT or WS_EX_LTRREADING or WS_EX_RIGHTSCROLLBAR )
      pWindow->SetClientSize( 420, iif(gFind.bExpanded, 66, 33) )
      pWindow->ClassStyle = CS_DBLCLKS
      
      ' we use RichEdit controls because regular Edit controls can not easily have
      ' their text vertically centered.
      dim as RECT rc
      dim cf as CHARFORMATW 
      cf.cbSize = sizeof(cf)
      cf.dwMask = CFM_COLOR
      cf.crTextColor = ghFindReplace.TextBoxForeColor

      hCtlFind = pWindow->AddControl("RICHEDIT", , IDC_FRMFINDREPLACE_TXTFIND, "", _
                    0, 0, 166, 24, _
                    WS_CHILD or WS_VISIBLE or WS_TABSTOP or ES_LEFT or ES_AUTOHSCROLL, _
                    WS_EX_LEFT or WS_EX_LTRREADING or WS_EX_RIGHTSCROLLBAR, _
                    0, @frmFindReplace_RichEdit_SubclassProc, IDC_FRMFINDREPLACE_TXTFIND, null )

      AfxSetWindowFont( hCtlFind, ghStatusBar.hFontStatusBar )
      SendMessage( hCtlFind, EM_SETCHARFORMAT, SCF_ALL, cast(LPARAM, @cf) ) 
      SendMessage( hCtlFind, EM_SETBKGNDCOLOR , 0, cast(LPARAM, ghFindReplace.TextBoxBackColor) )
      GetClientRect( hCtlFind, @rc )
      rc.top = rc.top + pWindow->ScaleY(2): rc.left = rc.left + pWindow->ScaleX(4)
      SendMessage( hCtlFind, EM_SETRECT, 0, cast(LPARAM, @rc) )
      SendMessage( hCtlFind, EM_SETEVENTMASK, 0, cast(LPARAM, ENM_SELCHANGE or ENM_CHANGE) )
      
      hCtlReplace = pWindow->AddControl("RICHEDIT", , IDC_FRMFINDREPLACE_TXTREPLACE, wstr(gFind.txtReplace), _
                    0, 0, 190, 24, _
                    WS_CHILD or WS_VISIBLE or WS_TABSTOP or ES_LEFT or ES_AUTOHSCROLL, _
                    WS_EX_LEFT or WS_EX_LTRREADING or WS_EX_RIGHTSCROLLBAR, _
                    0, @frmFindReplace_RichEdit_SubclassProc, IDC_FRMFINDREPLACE_TXTREPLACE, null )
      AfxSetWindowFont( hCtlReplace, ghStatusBar.hFontStatusBar )
      SendMessage( hCtlReplace, EM_SETCHARFORMAT, SCF_ALL, cast(LPARAM, @cf) ) 
      SendMessage( hCtlReplace, EM_SETBKGNDCOLOR , 0, cast(LPARAM, ghFindReplace.TextBoxBackColor) )
      GetClientRect( hCtlReplace, @rc )
      rc.top = rc.top + pWindow->ScaleY(2): rc.left = rc.left + pWindow->ScaleX(4)
      SendMessage( hCtlReplace, EM_SETRECT, 0, cast(LPARAM, @rc) )
      SendMessage( hCtlReplace, EM_SETEVENTMASK, 0, cast(LPARAM, ENM_SELCHANGE or ENM_CHANGE) )
     
      ' create semi-transparent window slightly offset under our popup menu in order to simulate a shadow.
      if gApp.isWineActive = false then
         pWindow = New CWindow
         pWindow->DPI = AfxCWindowPtr(HWND_FRMMAIN)->DPI
         HWND_FRMFINDREPLACE_SHADOW = pWindow->Create( HWND_FRMMAIN, "", _
              @frmFindReplaceShadow_WndProc, 0, 0, 0, 0, _
              WS_POPUP or WS_CLIPSIBLINGS or WS_CLIPCHILDREN, WS_EX_LAYERED )
         pWindow->ClassStyle = CS_DBLCLKS
         pWindow->Brush = GetSysColorBrush(COLOR_WINDOWTEXT + 1)   ' black background
         SetLayeredWindowAttributes( HWND_FRMFINDREPLACE_SHADOW, GetSysColor(COLOR_WINDOWTEXT + 1) , 80, LWA_ALPHA ) 
      end if
   end if

   ' Fills the search box with the selected word.
   ' if there are carriage returns or/and line feeds, this mean that
   ' there is a block selected.

   frmFindReplace_PositionWindows
   sFindText = pDoc->GetSelText
   if pDoc->IsMultilineSelection = false then
      hCtl = GetDlgItem( HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_TXTFIND )
      AfxSetWindowText( hCtl, sFindText )
      gFind.txtFind = sFindtext
   else
      gFind.nSelection = true   
      pDoc->SetMarkerHighlight
   end if   

   if gFind.bExpanded then
      hCtl = GetDlgItem( HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_TXTREPLACE )
   else   
      hCtl = GetDlgItem( HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_TXTFIND )
   end if   

   frmFindReplace_HighlightSearches( true )
   if len(AfxGetWindowText(hCtl)) then Edit_SetSel(hCtl, 0, -1)

   SetFocus hCtl 
   AfxRedrawWindow(HWND_FRMFINDREPLACE)

   function = 0
end function

