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

#include once "frmFindReplace.bi"



' ========================================================================================
' Find next/prev selection based on current document position
' ========================================================================================
private function frmFindReplace_NextSelection( byval StartPos as long, _
                                               byval bGetNext as boolean _
                                               ) as long

   dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr()
   if pDoc = 0 THEN exit function

   dim as long mainIdx, lenFind, endPos, nPos, iStart, iEnd
   
   dim as hwnd hEdit = pDoc->hWndActiveScintilla
   dim as hwnd hWndResults = GetDlgItem(HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_LBLRESULTS)      
   dim as hwnd hWndFind    = GetDlgItem(HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_TXTFIND)
   dim txtFind as string = Str(AfxGetWindowText(hWndFind))
   lenFind = len(txtFind)
   endPos = SciExec( hEdit, SCI_GETTEXTLENGTH, 0, 0)
   if endPos = 0 THEN exit function
    
   ' Build an array holding all of the selection positions
   redim iPositions(100) as long
   
   gFind.foundCount = 0 
   while nPos <= endPos 
      iStart = SciExec( hEdit, SCI_INDICATORSTART, 8, nPos) 
      iEnd   = SciExec( hEdit, SCI_INDICATOREND, 8, iStart)
      if (iStart = 0) andalso (iEnd = 0) THEN exit while    ' no indicators found

      if SciExec( hEdit, SCI_INDICATORVALUEAT, 8, iStart-1) then
         gFind.foundCount = gFind.foundCount + 1
         if gFind.foundCount > ubound(iPositions) THEN
            redim preserve iPositions(gFind.foundCount) as long 
         END IF
         iPositions(gFind.foundCount) = iStart - lenFind 
      END IF
      nPos = iEnd + 1
   Wend    
   
   ' Determine the next/prev selection that we should position to.
   if bGetNext THEN
      ' Next selection
      mainIdx = 1
      for i as long = 1 to gFind.foundCount
         if iPositions(i) > startPos THEN
            mainIdx = i: exit for
         END IF
      NEXT
   else
      ' Previous selection 
      mainIdx = gFind.foundCount
      for i as long = 1 to gFind.foundCount
         if iPositions(i) < startPos THEN
            mainIdx = i
         END IF
      NEXT
   END IF

   ' Highlight the selection
   if gFind.foundCount > 0 THEN
      ' make sure target line is unfolded
      dim as long currentLine = SciExec( hEdit, SCI_LINEFROMPOSITION, iPositions(mainIdx), 0)
      SciExec( hEdit, SCI_ENSUREVISIBLE, currentLine, 0) 
      SciExec( hEdit, SCI_SETSEL, iPositions(mainIdx), iPositions(mainIdx) + lenFind)
      AfxSetWindowText(hWndResults, mainIdx & " of " & gFind.foundCount)
   else
      SciExec( hEdit, SCI_SETSEL, startPos, startPos)
      AfxSetWindowText(hWndResults, L(269, "No results"))
   END IF
   
   function = 0
end function


' ========================================================================================
' Replace current selection or all selections
' ========================================================================================
private function frmFindReplace_DoReplace( byval fReplaceAll as BOOLEAN = false ) as long

   dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr()
   if pDoc = 0 THEN exit function

   dim as hwnd   hWndFind    = GetDlgItem(HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_TXTFIND)
   dim as hwnd   hWndReplace = GetDlgItem(HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_TXTREPLACE)
   dim as string txtFind     = Str(AfxGetWindowText(hWndFind))
   dim as string txtReplace  = str(AfxGetWindowText(hWndReplace))
   dim as long   lenFind     = len(txtFind)

   if lenFind = 0 then exit function
   
   dim as long nPos, startPos, endPos, findFlags, r
    
   dim as hwnd hEdit = pDoc->hWndActiveScintilla
   if fReplaceAll = true then 
      if gFind.nMatchCase THEN findFlags = (findFlags Or SCFIND_MATCHCASE)
      if gFind.nWholeWord THEN findFlags = (findFlags Or SCFIND_WHOLEWORD)
      SendMessage( hEdit, SCI_SETSEARCHFLAGS, findFlags, 0)
       
      if pDoc->HasMarkerHighlight THEN
         startPos = SciExec( hEdit, SCI_POSITIONFROMLINE, pDoc->FirstMarkerHighlight, 0)
         endPos = SciExec( hEdit, SCI_GETLINEENDPOSITION, pDoc->LastMarkerHighlight, 0)
         SciExec( hEdit, SCI_SETTARGETSTART, startPos, 0)
         SciExec( hEdit, SCI_SETTARGETEND, endPos, 0)
      else
         SciExec( hEdit, SCI_TARGETWHOLEDOCUMENT, 0, 0)
         startPos = SciExec( hEdit, SCI_GETTARGETSTART, 0, 0)
         endPos = SciExec( hEdit, SCI_GETTARGETEND, 0, 0)
      END IF
      SciExec( hEdit, SCI_INDICATORCLEARRANGE, startPos, endPos)
      
      gApp.SuppressNotify = true
      do 
         r = SciExec( hEdit, SCI_SEARCHINTARGET, Len(txtFind), Strptr(txtFind))
         if r = -1 THEN exit do
         SciExec( hEdit, SCI_REPLACETARGET, len(txtReplace), Strptr(txtReplace) )
         startPos = r + len(txtReplace)  'len(txtFind)
         ' Adjust the searching positions
         SciExec( hEdit, SCI_SETTARGETSTART, startPos, 0)
         SciExec( hEdit, SCI_SETTARGETEND, endPos, 0)
      loop
      gApp.SuppressNotify = false

   else   
      ' Need to ensure that the current cursor/selection is actually the find
      ' phrase that we want to replace. If not, then move to the closest selection.
      if ucase(pDoc->GetSelText) <> ucase(txtFind) THEN
         startPos = SciExec( hEdit, SCI_GETCURRENTPOS, 0, 0) 
         frmFindReplace_NextSelection(startPos, true)
         exit function
      END IF
      SciExec( hEdit, SCI_TARGETFROMSELECTION, 0, 0)
      SciExec( hEdit, SCI_REPLACETARGET, len(txtReplace), Strptr(txtReplace) )
   end if
      
   function = 0
end function


' ========================================================================================
' Highlight the found selections
' ========================================================================================
public function frmFindReplace_HighlightSearches() as long

   dim as long r, startPos, endPos, findFlags, nLength
   dim as hwnd hWndFind = GetDlgItem(HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_TXTFIND)
   dim as hwnd hWndResults = GetDlgItem(HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_LBLRESULTS)      
   
   dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr()
   if pDoc = 0 THEN exit function

   dim as hwnd hEdit = pDoc->hWndActiveScintilla
   SciExec( hEdit, SCI_SETADDITIONALCARETSVISIBLE, false, 0) 
   SciExec( hEdit, SCI_INDICSETSTYLE, 8, INDIC_STRAIGHTBOX)
   SciExec( hEdit, SCI_SETINDICATORCURRENT, 8, 0)
   
   dim txtFind as string = Str(AfxGetWindowText(hWndFind))
      
   if gFind.nMatchCase THEN findFlags = (findFlags Or SCFIND_MATCHCASE)
   if gFind.nWholeWord THEN findFlags = (findFlags Or SCFIND_WHOLEWORD)
   SciExec( hEdit, SCI_SETSEARCHFLAGS, findFlags, 0)
   
   ' Remove all existing selection indicators
   nLength = SendMessage( hEdit, SCI_GETTEXTLENGTH, 0, 0)
   SciExec( hEdit, SCI_INDICATORCLEARRANGE, 0, nLength)
   
   ' Are we searching a selection or the whole document
   if pDoc->HasMarkerHighlight THEN
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
         r = SciExec( hEdit, SCI_SEARCHINTARGET, Len(txtFind), Strptr(txtFind))
         if r = -1 THEN exit do
         
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
   frmFindReplace_NextSelection(startPos, true)
   
   function = 0

end function


' ========================================================================================
' Show the correct size for the form and the applicable controls
' ========================================================================================
public function frmFindReplace_ShowControls() as long
               
   Dim pWindow As CWindow ptr = AfxCWindowPtr(HWND_FRMFINDREPLACE)
   if pWindow = 0 THEN exit function
      
   dim as long fShow 
   dim as hwnd hCtl = GetDlgItem( HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_BTNTOGGLE)
    
   if gFind.bShowReplace THEN
      pWindow->SetClientSize(435, 64)
      pWindow->SetWindowPos(hCtl, 0, 0, 0, 20, 63, SWP_NOMOVE or SWP_NOZORDER)
      AfxCImageCtxPtr(hCtl)->LoadImageFromResource (pWindow->InstanceHandle, wstr("IMAGE_OPENED"))
      AfxCImageCtxPtr(hCtl)->SetImageSize( pWindow->ScaleX(20), pWindow->ScaleY(20)) 
      fShow = SW_SHOW
   else   
      pWindow->SetClientSize(435, 32)
      pWindow->SetWindowPos(hCtl, 0, 0, 0, 20, 31, SWP_NOMOVE or SWP_NOZORDER)
      AfxCImageCtxPtr(hCtl)->LoadImageFromResource (pWindow->InstanceHandle, wstr("IMAGE_CLOSED"))
      AfxCImageCtxPtr(hCtl)->SetImageSize( pWindow->ScaleX(20), pWindow->ScaleY(20)) 
      fShow = SW_HIDE
   END IF
   AfxRedrawWindow(hCtl)
   ShowWindow( GetDlgItem(HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_TXTREPLACE), fShow )
   ShowWindow( GetDlgItem(HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_BTNREPLACE), fShow )
   ShowWindow( GetDlgItem(HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_BTNREPLACEALL), fShow )

   hCtl = GetDlgItem( HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_BTNSELECTION)
   dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr()
   if pDoc = 0 THEN exit function
   if (pDoc->IsMultilineSelection) or (pDoc->HasMarkerHighlight) THEN
      AfxCImageCtxPtr(hCtl)->LoadImageFromResource (pWindow->InstanceHandle, wstr("IMAGE_SELECTED"))
      if gFind.nSelection then pDoc->SetMarkerHighlight
   else
      AfxCImageCtxPtr(hCtl)->LoadImageFromResource (pWindow->InstanceHandle, wstr("IMAGE_SELECTEDD"))
      if gFind.nSelection = false then pDoc->RemoveMarkerHighlight
   end if
   AfxRedrawWindow(hCtl)

   function = 0
end function


' ========================================================================================
' Process WM_COMMAND message for window/dialog: frmFindReplace
' ========================================================================================
private Function frmFindReplace_OnCommand( ByVal HWnd As HWnd, _
                                           ByVal id As Long, _
                                           ByVal hwndCtl As HWnd, _
                                           ByVal codeNotify As UINT _
                                           ) As LRESULT

   dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr()
   if pDoc = 0 THEN exit function

   dim as hwnd hEdit = pDoc->hWndActiveScintilla

   Select Case id

      case IDC_FRMFINDREPLACE_TXTFIND
         If (codeNotify = EN_CHANGE) Then
            frmFindReplace_HighlightSearches
            AfxRedrawWindow( hWnd )
         elseif (codeNotify = EN_SETFOCUS) then
            AfxRedrawWindow( hWnd )
         end if
             
      case IDC_FRMFINDREPLACE_TXTREPLACE
         If (codeNotify = EN_SETFOCUS) then
            AfxRedrawWindow( hWnd )
         end if

   end select
   
   Function = 0
End Function


' ========================================================================================
' Process WM_NOTIFY message for window/dialog: frmFindReplace
' ========================================================================================
private Function frmFindReplace_OnNotify( ByVal HWnd As HWnd, _
                                          ByVal id As Long, _
                                          ByVal pNMHDR As NMHDR Ptr _
                                          ) As LRESULT
 
   Select Case pNMHDR->code 
      
      Case NM_CLICK
         
         dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr()
         if pDoc = 0 THEN exit function
         dim as hwnd hEdit = pDoc->hWndActiveScintilla

         select case id
            case IDC_FRMFINDREPLACE_BTNCLOSE
               SendMessage( HWND_FRMFINDREPLACE, WM_CLOSE, 0, 0 )
               exit function
         
            case IDC_FRMFINDREPLACE_BTNTOGGLE
               gFind.bShowReplace = not gFind.bShowReplace
               frmFindReplace_ShowControls
               dim as hwnd hCtl = GetDlgItem( HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_TXTFIND )
               if GetFocus() <> hCtl Then
                  Edit_SetSel(hCtl, 0, -1)
                  SetFocus hCtl 
               end if   
               exit function         

            case IDC_FRMFINDREPLACE_BTNMATCHCASE
               gFind.nMatchCase = not gFind.nMatchCase
               frmFindReplace_HighlightSearches
               AfxRedrawWindow(HWND_FRMFINDREPLACE)
               exit function

            case IDC_FRMFINDREPLACE_BTNMATCHWHOLEWORD
               gFind.nWholeWord = not gFind.nWholeWord
               frmFindReplace_HighlightSearches
               AfxRedrawWindow(HWND_FRMFINDREPLACE)
               exit function
               
            case IDC_FRMFINDREPLACE_BTNLEFT
               if gFind.foundCount then
                  dim as long startPos = SendMessage( hEdit, SCI_GETANCHOR, 0, 0) 
                  frmFindReplace_NextSelection(startPos, false)
               end if
               exit function

            case IDC_FRMFINDREPLACE_BTNRIGHT
               if gFind.foundCount then
                  dim as long startPos = SendMessage( hEdit, SCI_GETCURRENTPOS, 0, 0) 
                  frmFindReplace_NextSelection(startPos, true)
               end if
               exit function

            case IDC_FRMFINDREPLACE_BTNSELECTION
               if (pDoc->IsMultilineSelection) or (pDoc->HasMarkerHighlight) THEN
                  gFind.nSelection = not gFind.nSelection
               else
                  gFind.nSelection = false
               end if
               if gFind.nSelection then
                  pDoc->SetMarkerHighlight
               else   
                  pDoc->RemoveMarkerHighlight
               end if
               frmFindReplace_HighlightSearches
               AfxRedrawWindow(HWND_FRMFINDREPLACE)
               exit function

            case IDC_FRMFINDREPLACE_BTNREPLACE
               frmFindReplace_DoReplace(false)
               frmFindReplace_HighlightSearches
               exit function

            case IDC_FRMFINDREPLACE_BTNREPLACEALL
               frmFindReplace_DoReplace(true)
               frmFindReplace_HighlightSearches
               exit function
         end select
         
   End Select
 
   Function = 0
End Function


' ========================================================================================
' Process WM_PAINT message for window/dialog: frmFindReplace
' ========================================================================================
private Function frmFindReplace_OnPaint( ByVal HWnd As HWnd) As LRESULT
            
   Dim As PAINTSTRUCT ps
   Dim As HPEN hPen
   dim as HWND hCtl
   Dim As HDC hDc
   Dim As Rect rc
   
   Dim pWindow As CWindow Ptr = AfxCWindowOwnerPtr(HWND)
   if pWindow = 0 THEN exit function

   hDC = BeginPaint(hWnd, @ps)

   SaveDC hDC

   ' Paint the extended part of the Find textbox white so that the
   ' MatchCase and WholeWord icons look like part of the textbox.
   hCtl = GetDlgItem(HWND, IDC_FRMFINDREPLACE_TXTFIND)
   GetWindowRect(hCtl, @rc)
   MapWindowPoints( HWND_DESKTOP, HWND, CAST(LPPOINT, @rc), 2)
   rc.right = rc.right + pWindow->ScaleX(44)
   FillRect( hDC, @rc, GetSysColorBrush(COLOR_WINDOW) )
   
 
   ' Outline the dialog with a dull gray pen
   SelectObject( hDC, GetStockObject( NULL_BRUSH ) )
   hPen = CreatePen( PS_SOLID, 1, GetSysColor(COLOR_ACTIVEBORDER) )
   SelectPen( hDC, hPen )
   GetClientRect HWnd, @rc
   MoveToEx( hDC, rc.left, rc.top, null)
   LineTo( hDC, rc.left, rc.bottom-1 ): LineTo( hDC, rc.right-1, rc.bottom-1): LineTo( hDC, rc.right-1, rc.top) 
   DeletePen(hPen)
   
   hPen = CreatePen( PS_SOLID, 1, BGR(0,0,255))  ' blue
   SelectPen( hDC, hPen )
 
   ' Paint a rectangle around the textbox control that has focus.
   hCtl = GetFocus()
   if hCtl THEN
      GetWindowRect(hCtl, @rc)
      ' If this is the Find textbox then extend it to cover the two embedded icons.
      if hCtl = GetDlgItem(HWND, IDC_FRMFINDREPLACE_TXTFIND) THEN
         rc.right = rc.right + pWindow->ScaleX(44)
      END IF
      SelectObject( hDC, GetSysColorBrush(COLOR_WINDOW) )
      MapWindowPoints( HWND_DESKTOP, HWND, CAST(LPPOINT, @rc), 2)
      Rectangle(hDC, rc.left - 1, rc.top - 1, rc.right + 1, rc.bottom + 1)
   END IF
   
   ' Paint blue selections around all icons that have TRUE selected states
   if gFind.nMatchCase then
      GetWindowRect(GetDlgItem(HWND, IDC_FRMFINDREPLACE_BTNMATCHCASE), @rc)
      SelectObject( hDC, GetSysColorBrush(COLOR_WINDOW) )
      MapWindowPoints( HWND_DESKTOP, HWND, CAST(LPPOINT, @rc), 2)
      Rectangle(hDC, rc.left - 1, rc.top - 1, rc.right + 1, rc.bottom + 1)
   end if
   if gFind.nWholeWord then
      GetWindowRect(GetDlgItem(HWND, IDC_FRMFINDREPLACE_BTNMATCHWHOLEWORD), @rc)
      SelectObject( hDC, GetSysColorBrush(COLOR_WINDOW) )
      MapWindowPoints( HWND_DESKTOP, HWND, CAST(LPPOINT, @rc), 2)
      Rectangle(hDC, rc.left - 1, rc.top - 1, rc.right + 1, rc.bottom + 1)
   end if
   if gFind.nSelection then
      GetWindowRect(GetDlgItem(HWND, IDC_FRMFINDREPLACE_BTNSELECTION), @rc)
      SelectObject( hDC, GetSysColorBrush(COLOR_BTNFACE) )
      MapWindowPoints( HWND_DESKTOP, HWND, CAST(LPPOINT, @rc), 2)
      Rectangle(hDC, rc.left - 1, rc.top - 1, rc.right + 1, rc.bottom + 1)
   end if

   DeletePen(hPen)
   RestoreDC hDC, -1 
   EndPaint hWnd, @ps
   
   Function = 0
End Function


' ========================================================================================
' Process WM_CLOSE message for window/dialog: frmFindReplace
' ========================================================================================
private Function frmFindReplace_OnClose( byval HWnd As HWnd ) As LRESULT
   DestroyWindow HWnd
   Function = 0
End Function


' ========================================================================================
' Process WM_DESTROY message for window/dialog: frmFindReplace
' ========================================================================================
private Function frmFindReplace_OnDestroy( byval HWnd As HWnd ) As LRESULT 
    
   ' Remove selected attributes (Attribute #8). We need to remove it from all
   ' open documents because the user may have tabbed between documents while
   ' the FindReplace dialog was active thereby causing selection highlights.
   gApp.RemoveAllSelectionAttributes
    
   ' Remove any markers set in the document that highlights
   ' entire lines (used for Selected text searching).
   Dim pDoc As clsDocument Ptr = gTTabCtl.GetActiveDocumentPtr()
   if pDoc THEN pDoc->RemoveMarkerHighlight
      
   SetFocusScintilla
   HWND_FRMFINDREPLACE = 0

   DIM pImageCtx AS CImageCtx ptr
   pImageCtx = AfxCImageCtxPtr(hwnd, IDC_FRMFINDREPLACE_BTNCLOSE)
   if pImageCtx then delete pImageCtx
   pImageCtx = AfxCImageCtxPtr(hwnd, IDC_FRMFINDREPLACE_BTNTOGGLE)
   if pImageCtx then delete pImageCtx
   pImageCtx = AfxCImageCtxPtr(hwnd, IDC_FRMFINDREPLACE_BTNMATCHCASE)
   if pImageCtx then delete pImageCtx
   pImageCtx = AfxCImageCtxPtr(hwnd, IDC_FRMFINDREPLACE_BTNMATCHWHOLEWORD)
   if pImageCtx then delete pImageCtx
   pImageCtx = AfxCImageCtxPtr(hwnd, IDC_FRMFINDREPLACE_BTNREPLACE)
   if pImageCtx then delete pImageCtx
   pImageCtx = AfxCImageCtxPtr(hwnd, IDC_FRMFINDREPLACE_BTNREPLACEALL)
   if pImageCtx then delete pImageCtx
   pImageCtx = AfxCImageCtxPtr(hwnd, IDC_FRMFINDREPLACE_BTNLEFT)
   if pImageCtx then delete pImageCtx
   pImageCtx = AfxCImageCtxPtr(hwnd, IDC_FRMFINDREPLACE_BTNRIGHT)
   if pImageCtx then delete pImageCtx
   pImageCtx = AfxCImageCtxPtr(hwnd, IDC_FRMFINDREPLACE_BTNSELECTION)
   if pImageCtx then delete pImageCtx

   ' Delete the pWindow because it gets recreated everytime frmFindReplace is invoked
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(HWND)
   if pWindow then Delete pWindow
   
   Function = 0
End Function


' ========================================================================================
' frmFindReplace_PositionWindow
' ========================================================================================
public Function frmFindReplace_PositionWindow() As LRESULT

   if IsWindow( HWND_FRMFINDREPLACE ) = 0 then exit function

   Dim pWindow As CWindow Ptr = AfxCWindowOwnerPtr( HWND_FRMFINDREPLACE )
   if pWindow = 0 THEN exit function

   dim as HWND hComboFunctions = GetDlgItem( HWND_FRMMAIN, IDC_FRMMAIN_COMBOFUNCTIONS )
   dim as RECT rc
   GetWindowRect( hComboFunctions, @rc )
   SetWindowPos HWND_FRMFINDREPLACE, HWND_TOP, _
                  rc.right - AfxGetWindowWidth(HWND_FRMFINDREPLACE) - pWindow->ScaleX(50), _
                  rc.bottom + pWindow->ScaleY(2), _
                  0, 0, SWP_SHOWWINDOW or SWP_NOSIZE or SWP_NOACTIVATE
   function = 0
end function


' ========================================================================================
' frmFindReplace Window procedure
' ========================================================================================
private Function frmFindReplace_WndProc( ByVal HWnd   As HWnd, _
                                         ByVal uMsg   As UINT, _
                                         ByVal wParam As WPARAM, _
                                         ByVal lParam As LPARAM _
                                         ) As LRESULT
   Select Case uMsg
      HANDLE_MSG (HWnd, WM_CLOSE,    frmFindReplace_OnClose)
      HANDLE_MSG (HWnd, WM_DESTROY,  frmFindReplace_OnDestroy)
      HANDLE_MSG (HWnd, WM_COMMAND,  frmFindReplace_OnCommand)
      HANDLE_MSG (HWnd, WM_NOTIFY,   frmFindReplace_OnNotify)
      HANDLE_MSG (HWnd, WM_PAINT,    frmFindReplace_OnPaint)
   End Select

   Function = DefWindowProc(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' frmFindReplace_Show
' ========================================================================================
public Function frmFindReplace_Show( ByVal hWndParent As HWnd, _
                                     byval fShowReplace as BOOLEAN _
                                     ) As LRESULT

   '  Create the main window and child controls
   Dim pWindow As CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowOwnerPtr(hwndParent)->DPI

   dim as HWND hCtl

   gFind.bShowReplace = fShowReplace
   
   ' If the FindReplace window already exists then simply set focus to it and exit.
   If IsWindowVisible(HWND_FRMFINDREPLACE) Then
      frmFindReplace_ShowControls
      hCtl = GetDlgItem( HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_TXTFIND )
      if len(AfxGetWindowText(hCtl)) then Edit_SetSel(hCtl, 0, -1)
      SetFocus hCtl 
      AfxRedrawWindow(HWND_FRMFINDREPLACE)
      Exit Function
   End If
   
   dim as RECT rc
   dim as long frmWidth   = 435
   dim as long frmHeight  = 64
   dim as long iconWidth  = 16
   dim as long iconHeight = 16
   dim as long margin     = 4
   dim as long vmargin    = 6
   dim as long nLeft      = 0
   dim as long nTop       = 0
   dim as long mainSel
   dim as HWND hCtlFind, hCtlReplace 
   dim wszImage as wstring * MAX_PATH
   
   Dim pDoc As clsDocument Ptr = gTTabCtl.GetActiveDocumentPtr()
   if pDoc = 0 THEN exit function
      
   gFind.nMatchCase = false
   gFind.nWholeWord = false
   gFind.nSelection = false
   
   ' Fills the search box with the selected word.
   ' If there are carriage returns or/and line feeds, this mean that
   ' there is a block selected, instead of a word, so avoid it.
   Dim As String sFindtext = gFind.txtFind
   if pDoc->IsMultilineSelection = false THEN
      ' Default to any highlighted text
      sFindText = pDoc->GetSelText
   else
      gFind.nSelection = true   
   END IF
   
   HWND_FRMFINDREPLACE = _
   pWindow->Create(hWndParent, "FINDREPLACE", @frmFindReplace_WndProc, 0, 0, 0, 0, _
        WS_POPUP Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, _   
        WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->SetClientSize(frmWidth, frmHeight)

   DIM pImageCtx AS CImageCtx ptr

   ' Display the expand toggle icon for the Find/Replace 
   pImageCtx = new CImageCtx(pWindow, IDC_FRMFINDREPLACE_BTNTOGGLE, "", nLeft, nTop, 20, frmHeight)
   pImageCtx->SetBkColor(GetSysColor(COLOR_BTNFACE))
   pImageCtx->SetBkColorHot(GetSysColor(COLOR_3DLIGHT))
   pImageCtx->LoadImageFromResource (pWindow->InstanceHandle, wstr("IMAGE_CLOSED"))
   pImageCtx->SetImageSize( pWindow->ScaleX(20), pWindow->ScaleY(20)) 
   pImageCtx->SetImageAdjustment(GDIP_IMAGECTX_ACTUALSIZE) 
   AfxAddTooltip(pImageCtx->hWindow, L(270, "Toggle Replace mode"))
                     

   ' Display the Find textbox               
   nTop  = margin
   nLeft = nLeft + 20 + margin
   hCtlFind = pWindow->AddControl("TEXTBOX", , IDC_FRMFINDREPLACE_TXTFIND, WSTR(sFindText), nLeft, nTop, 176, 22, _
              WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL, _
              WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   SendMessage hCtlFind, EM_SETCUEBANNER, 1, CAST(LPARAM, @L(158,"Find"))
       
   ' Display the Match Case button 
   nTop  = margin + 3
   nLeft = nLeft + pWindow->UnScaleX(AfxGetWindowWidth(hCtlFind)) + 4
   pImageCtx = new CImageCtx(pWindow, IDC_FRMFINDREPLACE_BTNMATCHCASE, "", nLeft, nTop, iconHeight, iconHeight)
   pImageCtx->SetBkColor(GetSysColor(COLOR_WINDOW))
   pImageCtx->SetBkColorHot(GetSysColor(COLOR_3DLIGHT))
   wszImage = iif(pWindow->DPI > 96, wstr("IMAGE_CASE24"), wstr("IMAGE_CASE"))
   pImageCtx->LoadImageFromResource (pWindow->InstanceHandle, wszImage)
   AfxAddTooltip(pImageCtx->hWindow, L(163,"Match Case"))
       
   ' Display the Match Whole Word button 
   nTop  = margin + 3
   nLeft = nLeft + pWindow->UnScaleX(AfxGetWindowWidth(pImageCtx->hWindow)) + 4  
   pImageCtx = new CImageCtx(pWindow, IDC_FRMFINDREPLACE_BTNMATCHWHOLEWORD, "", nLeft, nTop, iconHeight, iconHeight)
   pImageCtx->SetBkColor(GetSysColor(COLOR_WINDOW))
   pImageCtx->SetBkColorHot(GetSysColor(COLOR_3DLIGHT))
   wszImage = iif(pWindow->DPI > 96, wstr("IMAGE_WHOLE24"), wstr("IMAGE_WHOLE"))
   pImageCtx->LoadImageFromResource (pWindow->InstanceHandle, wszImage)
   AfxAddTooltip(pImageCtx->hWindow, L(162,"Match Whole Words"))

   ' Display the Results label
   nTop  = margin 
   nLeft = nLeft + pWindow->UnScaleX(AfxGetWindowWidth(pImageCtx->hWindow)) + margin + 10
   hCtl = pWindow->AddControl("LABEL", , IDC_FRMFINDREPLACE_LBLRESULTS, L(269, "No results"), _
           nLeft, nTop, 80, 22, _
           WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_CENTER Or SS_CENTERIMAGE, _
           WS_EX_LEFT Or WS_EX_LTRREADING)
       
   ' Display the Left button 
   nTop  = margin + 4
   nLeft = nLeft + pWindow->UnScaleX(AfxGetWindowWidth(hCtl)) + margin  
   pImageCtx = new CImageCtx(pWindow, IDC_FRMFINDREPLACE_BTNLEFT, "", nLeft, nTop, iconHeight, iconHeight)
   pImageCtx->SetBkColor(GetSysColor(COLOR_BTNFACE)): pImageCtx->SetBkColorHot(GetSysColor(COLOR_3DLIGHT))
   pImageCtx->LoadImageFromResource (pWindow->InstanceHandle, wstr("IMAGE_LEFT"))
   AfxAddTooltip(pImageCtx->hWindow, L(45,"Find Previous") & WSTR(" (Shift+F3)"))

   ' Display the Right button 
   nTop  = margin + 4
   nLeft = nLeft + pWindow->UnScaleX(AfxGetWindowWidth(pImageCtx->hWindow)) + margin  
   pImageCtx = new CImageCtx(pWindow, IDC_FRMFINDREPLACE_BTNRIGHT, "", nLeft, nTop, iconHeight, iconHeight)
   pImageCtx->SetBkColor(GetSysColor(COLOR_BTNFACE)): pImageCtx->SetBkColorHot(GetSysColor(COLOR_3DLIGHT))
   pImageCtx->LoadImageFromResource (pWindow->InstanceHandle, wstr("IMAGE_RIGHT"))
   AfxAddTooltip(pImageCtx->hWindow, L(44,"Find Next") & WSTR(" (F3)"))

   ' Display the Selection button 
   nTop  = margin + 4
   nLeft = nLeft + pWindow->UnScaleX(AfxGetWindowWidth(pImageCtx->hWindow)) + margin + 4 
   pImageCtx = new CImageCtx(pWindow, IDC_FRMFINDREPLACE_BTNSELECTION, "", nLeft, nTop, iconHeight, iconHeight)
   pImageCtx->SetBkColor(GetSysColor(COLOR_BTNFACE)): pImageCtx->SetBkColorHot(GetSysColor(COLOR_3DLIGHT))
   pImageCtx->LoadImageFromResource (pWindow->InstanceHandle, wstr("IMAGE_SELECTED"))
   AfxAddTooltip(pImageCtx->hWindow, L(148,"Selection"))

   ' Display the Replace textbox               
   nTop  = margin + pWindow->UnScaleX(AfxGetWindowHeight(hCtlFind)) + vmargin
   nLeft = 20 + margin
   hCtlReplace = pWindow->AddControl("TEXTBOX", , IDC_FRMFINDREPLACE_TXTREPLACE, wstr(gFind.txtReplace), nLeft, nTop, 220, 22, _
              WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL, _
              WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   SendMessage hCtlReplace, EM_SETCUEBANNER, 1, CAST(LPARAM, @L(173,"Replace"))
        
   ' Display the Replace button 
   nTop  = margin + pWindow->UnScaleX(AfxGetWindowHeight(hCtlFind)) + vmargin + 2
   nLeft = nLeft + pWindow->UnScaleX(AfxGetWindowWidth(hCtlReplace)) + margin + 4 
   pImageCtx = new CImageCtx(pWindow, IDC_FRMFINDREPLACE_BTNREPLACE, "", nLeft, nTop, iconHeight, iconHeight)
   pImageCtx->SetBkColor(GetSysColor(COLOR_BTNFACE)): pImageCtx->SetBkColorHot(GetSysColor(COLOR_3DLIGHT))
   wszImage = iif(pWindow->DPI > 96, wstr("IMAGE_REPLACE24"), wstr("IMAGE_REPLACE"))
   pImageCtx->LoadImageFromResource (pWindow->InstanceHandle, wszImage)
   AfxAddTooltip(pImageCtx->hWindow, L(173,"Replace"))

   ' Display the Replace All button 
   nTop  = margin + pWindow->UnScaleX(AfxGetWindowHeight(hCtlFind)) + vmargin + 2
   nLeft = nLeft + pWindow->UnScaleX(AfxGetWindowWidth(pImageCtx->hWindow)) + margin + 4 
   pImageCtx = new CImageCtx(pWindow, IDC_FRMFINDREPLACE_BTNREPLACEALL, "", nLeft, nTop, iconHeight, iconHeight)
   pImageCtx->SetBkColor(GetSysColor(COLOR_BTNFACE)): pImageCtx->SetBkColorHot(GetSysColor(COLOR_3DLIGHT))
   wszImage = iif(pWindow->DPI > 96, wstr("IMAGE_REPLACEALL24"), wstr("IMAGE_REPLACEALL"))
   pImageCtx->LoadImageFromResource (pWindow->InstanceHandle, wszImage)
   AfxAddTooltip(pImageCtx->hWindow, L(174,"Replace All"))

   ' Diplay the Close icon
   nTop  = margin 
   nLeft = frmWidth - 20 - margin 
   pImageCtx = new CImageCtx(pWindow, IDC_FRMFINDREPLACE_BTNCLOSE, "", nLeft, nTop, 20, 20)
   pImageCtx->SetBkColor(GetSysColor(COLOR_BTNFACE))
   pImageCtx->SetBkColorHot(GetSysColor(COLOR_3DLIGHT))
   pImageCtx->LoadImageFromResource (pWindow->InstanceHandle, wstr("IMAGE_CLOSE"))
   AfxAddTooltip(pImageCtx->hWindow, L(161, "Close"))

   frmFindReplace_ShowControls

   frmFindReplace_PositionWindow
   
   hCtl = GetDlgItem( HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_TXTFIND )
   if len(AfxGetWindowText(hCtl)) then Edit_SetSel(hCtl, 0, -1)

   frmFindReplace_HighlightSearches
   SetFocus hCtl 

   Function = 0
End Function

