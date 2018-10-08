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
' Replace current selection or all selections
' ========================================================================================
function frmFindReplace_DoReplace( byval fReplaceAll as BOOLEAN = false ) as long
   dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr()
   if pDoc = 0 THEN exit function

   dim as hwnd   hWndFind    = GetDlgItem(HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_TXTFIND)
   dim as hwnd   hWndReplace = GetDlgItem(HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_TXTREPLACE)
   dim as hwnd   hWndCase    = GetDlgItem(HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_BTNMATCHCASE)      
   dim as hwnd   hWndWhole   = GetDlgItem(HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_BTNMATCHWHOLEWORD)      
   dim as string txtFind     = Str(AfxGetWindowText(hWndFind))
   dim as string txtReplace  = str(AfxGetWindowText(hWndReplace))
   dim as long   lenFind     = len(txtFind)

   dim as long nPos, startPos, endPos, findFlags, r
   
   dim as hwnd hEdit = pDoc->hWndActiveScintilla
   if fReplaceAll = true then 
      if frmHotImgBtn_GetSelected(hWndCase) THEN findFlags = (findFlags Or SCFIND_MATCHCASE)
      if frmHotImgBtn_GetSelected(hWndWhole) THEN findFlags = (findFlags Or SCFIND_WHOLEWORD)
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
         startPos = r + len(txtFind)
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
' Find next/prev selection based on current document position
' ========================================================================================
function frmFindReplace_NextSelection( byval StartPos as long, byval bGetNext as boolean ) as long
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
' Highlight the found selections
' ========================================================================================
function frmFindReplace_HighlightSearches() as long

   dim as long r, startPos, endPos, findFlags, nLength
   dim as hwnd hWndFind = GetDlgItem(HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_TXTFIND)
   dim as hwnd hWndResults = GetDlgItem(HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_LBLRESULTS)      
   dim as hwnd hWndCase = GetDlgItem(HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_BTNMATCHCASE)      
   dim as hwnd hWndWhole = GetDlgItem(HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_BTNMATCHWHOLEWORD)      
   
   dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr()
   if pDoc = 0 THEN exit function

   dim as hwnd hEdit = pDoc->hWndActiveScintilla
   SciExec( hEdit, SCI_SETADDITIONALCARETSVISIBLE, false, 0) 
   SciExec( hEdit, SCI_INDICSETSTYLE, 8, INDIC_STRAIGHTBOX)
   SciExec( hEdit, SCI_SETINDICATORCURRENT, 8, 0)
   
   dim txtFind as string = Str(AfxGetWindowText(hWndFind))
      
   if frmHotImgBtn_GetSelected(hWndCase) THEN findFlags = (findFlags Or SCFIND_MATCHCASE)
   if frmHotImgBtn_GetSelected(hWndWhole) THEN findFlags = (findFlags Or SCFIND_WHOLEWORD)
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
' Process WM_COMMAND message for window/dialog: frmFindReplace
' ========================================================================================
function frmFindReplace_SetButtonStates() as LONG
   dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr()
   if pDoc = 0 THEN exit function

   dim as hwnd hCtl
   dim as COLORREF clrBack = GetSysColor(COLOR_BTNFACE) 
   dim as COLORREF clrHot  = GetSysColor(COLOR_3DLIGHT) 
   
   if gFind.foundCount > 0 THEN

      hCtl = GetDlgItem(HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_BTNLEFT)      
      frmHotImgBtn_SetBackColors(hCtl, clrBack, clrHot)
      frmHotImgBtn_SetImages(hCtl, "IMAGE_LEFT", "IMAGE_LEFT", true)
      frmHotImgBtn_Enabled(hCtl, true )

      hCtl = GetDlgItem(HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_BTNRIGHT)      
      frmHotImgBtn_SetBackColors(hCtl, clrBack, clrHot)
      frmHotImgBtn_SetImages(hCtl, "IMAGE_RIGHT", "IMAGE_RIGHT", true)
      frmHotImgBtn_Enabled(hCtl, true )

      hCtl = GetDlgItem(HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_BTNREPLACE)      
      frmHotImgBtn_SetBackColors(hCtl, clrBack, clrHot)
      frmHotImgBtn_SetImages(hCtl, "IMAGE_REPLACE", "IMAGE_REPLACE", true)
      frmHotImgBtn_Enabled(hCtl, true )

      hCtl = GetDlgItem(HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_BTNREPLACEALL)      
      frmHotImgBtn_SetBackColors(hCtl, clrBack, clrHot)
      frmHotImgBtn_SetImages(hCtl, "IMAGE_REPLACEALL", "IMAGE_REPLACEALL", true)
      frmHotImgBtn_Enabled(hCtl, true )

   else

      hCtl = GetDlgItem(HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_BTNLEFT)      
      frmHotImgBtn_SetBackColors(hCtl, clrBack, clrBack)
      frmHotImgBtn_SetImages(hCtl, "IMAGE_LEFTD", "IMAGE_LEFTD", true)
      frmHotImgBtn_Enabled(hCtl, false )

      hCtl = GetDlgItem(HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_BTNRIGHT)      
      frmHotImgBtn_SetBackColors(hCtl, clrBack, clrBack)
      frmHotImgBtn_SetImages(hCtl, "IMAGE_RIGHTD", "IMAGE_RIGHTD", true)
      frmHotImgBtn_Enabled(hCtl, false )

      hCtl = GetDlgItem(HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_BTNREPLACE)      
      frmHotImgBtn_SetBackColors(hCtl, clrBack, clrBack)
      frmHotImgBtn_SetImages(hCtl, "IMAGE_REPLACED", "IMAGE_REPLACED", true)
      frmHotImgBtn_Enabled(hCtl, false )
      
      hCtl = GetDlgItem(HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_BTNREPLACEALL)      
      frmHotImgBtn_SetBackColors(hCtl, clrBack, clrBack)
      frmHotImgBtn_SetImages(hCtl, "IMAGE_REPLACEALLD", "IMAGE_REPLACEALLD", true)
      frmHotImgBtn_Enabled(hCtl, false )
   END IF

   hCtl = GetDlgItem(HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_BTNSELECTION)      
   if (pDoc->IsMultilineSelection) or (pDoc->HasMarkerHighlight) THEN
      frmHotImgBtn_SetBackColors(hCtl, clrBack, clrHot)
      frmHotImgBtn_SetImages(hCtl, "IMAGE_SELECTED", "IMAGE_SELECTED", true)
      frmHotImgBtn_Enabled(hCtl, true)
   else   
      frmHotImgBtn_SetBackColors(hCtl, clrBack, clrBack)
      frmHotImgBtn_SetImages(hCtl, "IMAGE_SELECTEDD", "IMAGE_SELECTEDD", true)
      frmHotImgBtn_Enabled(hCtl, false )
   end if
   
   function = 0
end function



' ========================================================================================
' Show the correct size for the form and the applicable controls
' ========================================================================================
function frmFindReplace_ShowControls() as long
               
   Dim pWindow As CWindow ptr = AfxCWindowPtr(HWND_FRMFINDREPLACE)
   if pWindow = 0 THEN exit function
      
   dim as long fShow 
   dim as hwnd hCtl = GetDlgItem( HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_BTNTOGGLE)
   
   if gReplaceOpen THEN
      frmHotImgBtn_SetImages(hCtl, "IMAGE_OPENED", "IMAGE_OPENED", true)
      pWindow->SetClientSize(435, 64)
      pWindow->SetWindowPos(hCtl, 0, 0, 0, 20, 63, SWP_NOMOVE or SWP_NOZORDER)
      fShow = SW_SHOW
   else   
      frmHotImgBtn_SetImages(hCtl, "IMAGE_CLOSED", "IMAGE_CLOSED", true)
      pWindow->SetClientSize(435, 32)
      pWindow->SetWindowPos(hCtl, 0, 0, 0, 20, 31, SWP_NOMOVE or SWP_NOZORDER)
      fShow = SW_HIDE
   END IF
   ShowWindow( GetDlgItem(HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_TXTREPLACE), fShow )
   ShowWindow( GetDlgItem(HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_BTNREPLACE), fShow )
   ShowWindow( GetDlgItem(HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_BTNREPLACEALL), fShow )
   AfxRedrawWindow( HWND_FRMFINDREPLACE )

   function = 0
end function


' ========================================================================================
' Process WM_COMMAND message for window/dialog: frmFindReplace
' ========================================================================================
Function frmFindReplace_OnCommand(ByVal HWnd As HWnd, ByVal id As Long, ByVal hwndCtl As HWnd, ByVal codeNotify As UINT) As LRESULT

   dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr()
   if pDoc = 0 THEN exit function

   dim as hwnd hEdit = pDoc->hWndActiveScintilla

   Select Case id

      case IDC_FRMFINDREPLACE_BTNTOGGLE
         If codeNotify = BN_CLICKED Then
            gReplaceOpen = iif(gReplaceOpen, false, true )
            frmFindReplace_ShowControls
            dim as hwnd hCtl = GetDlgItem( HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_TXTFIND )
            if GetFocus() <> hCtl Then
               Edit_SetSel(hCtl, 0, -1)
               SetFocus hCtl 
            end if   
         end if
                  
      case IDC_FRMFINDREPLACE_BTNCLOSE
         If codeNotify = BN_CLICKED Then
            SendMessage( HWND_FRMFINDREPLACE, WM_CLOSE, 0, 0 )
         end if

      case IDC_FRMFINDREPLACE_BTNMATCHCASE, IDC_FRMFINDREPLACE_BTNMATCHWHOLEWORD
         If codeNotify = BN_CLICKED Then
            if frmHotImgBtn_GetSelected(hwndctl) THEN
               frmHotImgBtn_SetSelected(hwndCtl, false)
            else   
               frmHotImgBtn_SetSelected(hwndCtl, true)
            END IF
            frmFindReplace_HighlightSearches
            frmFindReplace_SetButtonStates
         end if

      case IDC_FRMFINDREPLACE_TXTFIND
         If (codeNotify = EN_CHANGE) Then
            frmFindReplace_HighlightSearches
            frmFindReplace_SetButtonStates
            AfxRedrawWindow( hWnd )
         elseif (codeNotify = EN_SETFOCUS) then
            frmFindReplace_SetButtonStates
            AfxRedrawWindow( hWnd )
         end if
            
      case IDC_FRMFINDREPLACE_TXTREPLACE
         If (codeNotify = EN_SETFOCUS) then
            frmFindReplace_SetButtonStates
            AfxRedrawWindow( hWnd )
         end if

      case IDC_FRMFINDREPLACE_BTNLEFT
         If codeNotify = BN_CLICKED Then
            if frmHotImgBtn_IsEnabled(hwndCtl) = false then exit function
            dim as long startPos = SendMessage( hEdit, SCI_GETANCHOR, 0, 0) 
            frmFindReplace_NextSelection(startPos, false)
         end if

      case IDC_FRMFINDREPLACE_BTNRIGHT
         If codeNotify = BN_CLICKED Then
            if frmHotImgBtn_IsEnabled(hwndCtl) = false then exit function
            dim as long startPos = SendMessage( hEdit, SCI_GETCURRENTPOS, 0, 0) 
            frmFindReplace_NextSelection(startPos, true)
         end if

      case IDC_FRMFINDREPLACE_BTNSELECTION
         If codeNotify = BN_CLICKED Then
            if frmHotImgBtn_IsEnabled(hwndCtl) = false then exit function
            if frmHotImgBtn_GetSelected(hwndCtl) = false then
               frmHotImgBtn_SetSelected(hwndCtl, TRUE)
               pDoc->SetMarkerHighlight
            else   
               frmHotImgBtn_SetSelected(hwndCtl, false)
               pDoc->RemoveMarkerHighlight
            end if
            frmFindReplace_SetButtonStates
         end if

      case IDC_FRMFINDREPLACE_BTNREPLACE
         If codeNotify = BN_CLICKED Then
            if frmHotImgBtn_IsEnabled(hwndCtl) = false then exit function
            frmFindReplace_DoReplace(false)
            frmFindReplace_HighlightSearches
            frmFindReplace_SetButtonStates
         end if

      case IDC_FRMFINDREPLACE_BTNREPLACEALL
         If codeNotify = BN_CLICKED Then
            if frmHotImgBtn_IsEnabled(hwndCtl) = false then exit function
            frmFindReplace_DoReplace(true)
            frmFindReplace_HighlightSearches
            frmFindReplace_SetButtonStates
         end if
   end select
   
   Function = 0
End Function


' ========================================================================================
' Process WM_PAINT message for window/dialog: frmFindReplace
' ========================================================================================
Function frmFindReplace_OnPaint( ByVal HWnd As HWnd) As LRESULT
            
   Dim As PAINTSTRUCT ps
   Dim As HPEN hPen
   dim as HWND hCtl
   Dim As HDC hDc
   Dim As Rect rc
   
   Dim pWindow As CWindow Ptr = AfxCWindowOwnerPtr(HWND)
   if pWindow = 0 THEN exit function

   hDC = BeginPaint(hWnd, @ps)

   SaveDC hDC

   SelectObject( hDC, GetStockObject( NULL_BRUSH ) )
 
   ' Outline the dialog with a dull gray pen
   hPen = CreatePen( PS_SOLID, 1, GetSysColor(COLOR_ACTIVEBORDER) )
   SelectPen( hDC, hPen )
   GetClientRect HWnd, @rc
   MoveToEx( hDC, rc.left, rc.top, null)
   LineTo( hDC, rc.left, rc.bottom-1 ): LineTo( hDC, rc.right-1, rc.bottom-1): LineTo( hDC, rc.right-1, rc.top) 
   DeletePen(hPen)
   
   hPen = CreatePen( PS_SOLID, 1, BGR(0,0,255))  ' blue
   SelectPen( hDC, hPen )
 
   ' Paint a null brushed rectangle around the control that has focus.
   hCtl = GetFocus()
   if hCtl THEN
      GetWindowRect(hCtl, @rc)
      ' If this is the Find textbox then extend it to cover the two embedded icons.
      if hCtl = GetDlgItem(HWND, IDC_FRMFINDREPLACE_TXTFIND) THEN
         rc.right = rc.right + pWindow->ScaleX(44)
      END IF
      SelectObject( hDC, GetSysColorBrush(COLOR_3DFACE) )
      MapWindowPoints( HWND_DESKTOP, HWND, CAST(LPPOINT, @rc), 2)
      Rectangle(hDC, rc.left - 1, rc.top - 1, rc.right + 1, rc.bottom + 1)
   END IF
   
   DeletePen(hPen)
   RestoreDC hDC, -1 
   
   EndPaint hWnd, @ps
   
   Function = 0
End Function


' ========================================================================================
' Process WM_CLOSE message for window/dialog: frmFindReplace
' ========================================================================================
Function frmFindReplace_OnClose(HWnd As HWnd) As LRESULT
   DestroyWindow HWnd
   Function = 0
End Function


' ========================================================================================
' Process WM_DESTROY message for window/dialog: frmFindReplace
' ========================================================================================
Function frmFindReplace_OnDestroy(HWnd As HWnd) As LRESULT 
   dim as hwnd hCtl = GetDlgItem( hWnd, IDC_FRMFINDREPLACE_TXTFIND)
   gFind.txtFind = str(AfxGetWindowText(hCtl))

   hCtl = GetDlgItem( hWnd, IDC_FRMFINDREPLACE_TXTREPLACE)
   gFind.txtReplace = str(AfxGetWindowText(hCtl))

   hCtl = GetDlgItem( hWnd, IDC_FRMFINDREPLACE_BTNMATCHCASE)
   gFind.nMatchCase = frmHotImgBtn_GetSelected(hCtl)
   hCtl = GetDlgItem( hWnd, IDC_FRMFINDREPLACE_BTNMATCHWHOLEWORD)
   gFind.nWholeWord = frmHotImgBtn_GetSelected(hCtl)
   gFind.txtFind = ""
   
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
   
   ' Delete the pWindow because it gets recreated everytime frmFindReplace is invoked
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(HWND)
   if pWindow then Delete pWindow
   
   Function = 0
End Function


' ========================================================================================
' frmFindReplace Window procedure
' ========================================================================================
Function frmFindReplace_WndProc( ByVal HWnd   As HWnd, _
                                 ByVal uMsg   As UINT, _
                                 ByVal wParam As WPARAM, _
                                 ByVal lParam As LPARAM _
                                 ) As LRESULT

   Select Case uMsg
      HANDLE_MSG (HWnd, WM_CLOSE,    frmFindReplace_OnClose)
      HANDLE_MSG (HWnd, WM_DESTROY,  frmFindReplace_OnDestroy)
      HANDLE_MSG (HWnd, WM_COMMAND,  frmFindReplace_OnCommand)
      HANDLE_MSG (HWnd, WM_PAINT,    frmFindReplace_OnPaint)
   End Select

   ' for messages that we don't deal with
   Function = DefWindowProc(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' frmFindReplace_Show
' ========================================================================================
Function frmFindReplace_Show( ByVal hWndParent As HWnd, byval fShowReplace as BOOLEAN ) As Long

   dim as HWND hCtl

   gReplaceOpen = fShowReplace

  
      
   '  Create the main window and child controls
   Dim pWindow As CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowOwnerPtr(hwndParent)->DPI
   pWindow->ClassStyle = CS_DROPSHADOW

   dim as RECT rc
   dim as long frmWidth   = 435
   dim as long frmHeight  = 64
   dim as long ctlHeight  = 22
   dim as long iconHeight = 16  
   dim as long margin     = 4
   dim as long vmargin    = 6
   dim as long nLeft      = 0
   dim as long nTop       = 0
   dim as long mainSel
   dim as HWND hCtlFind, hCtlReplace 
   
   Dim pDoc As clsDocument Ptr = gTTabCtl.GetActiveDocumentPtr()
   if pDoc = 0 THEN exit function
      
   ' Fills the search box with the selected word.
   ' If there are carriage returns or/and line feeds, this mean that
   ' there is a block selected, instead of a word, so avoid it.
   Dim As String sFindtext = gFind.txtFind
   if pDoc->IsMultilineSelection = false THEN
      ' Default to any highlighted text
      sFindText = pDoc->GetSelText
   END IF
   
   ' If the FindReplace window already exists then simply set focus to it and exit.
   If IsWindowVisible(HWND_FRMFINDREPLACE) Then
      frmFindReplace_ShowControls
      hCtl = GetDlgItem( HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_TXTFIND )
      if len(sFindText) then AfxSetWindowText(hCtl,sFindText) 
      if len(AfxGetWindowText(hCtl)) then Edit_SetSel(hCtl, 0, -1)
      SetFocus hCtl 
      Exit Function
   End If
      
   HWND_FRMFINDREPLACE = _
   pWindow->Create(hWndParent, "FINDREPLACE", @frmFindReplace_WndProc, 0, 0, 0, 0, _
        WS_POPUP Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, _
        WS_EX_CONTROLPARENT or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->SetClientSize(frmWidth, frmHeight)

   ' Display the expand toggle icon for the Find/Replace 
   hCtl = frmHotImgBtn( HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_BTNTOGGLE, "IMAGE_CLOSED", "IMAGE_CLOSED",  _
                  L(270, "Toggle Replace mode"), iconHeight, iconHeight, GetSysColor(COLOR_BTNFACE), GetSysColor(COLOR_3DLIGHT), _
                  nLeft+1, 0, 20, frmHeight-1, false, True, true)
                     
   ' Display the Find textbox               
   nTop  = margin
   nLeft = nLeft + 20 + margin
   hCtlFind = pWindow->AddControl("TEXTBOX", , IDC_FRMFINDREPLACE_TXTFIND, WSTR(sFindText), nLeft, nTop, 176, ctlHeight, _
              WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL, _
              WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   SendMessage hCtlFind, EM_SETCUEBANNER, 1, CAST(LPARAM, @L(158,"Find"))
      
   ' Display the Match Case button 
   nTop  = margin
   nLeft = nLeft + pWindow->UnScaleX(AfxGetWindowWidth(hCtlFind))
   hCtl  = frmHotImgBtn( HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_BTNMATCHCASE, "IMAGE_CASED", "IMAGE_CASED",  _
                  L(163,"Match Case"), iconHeight , iconHeight , GetSysColor(COLOR_WINDOW), GetSysColor(COLOR_WINDOW), _
                  nLeft, nTop, ctlHeight, ctlHeight, true, true, true)
   frmHotImgBtn_SetSelected(hCtl, gFind.nMatchCase)
      
   ' Display the Match Whole Word button 
   nTop  = margin
   nLeft = nLeft + pWindow->UnScaleX(AfxGetWindowWidth(hCtl))  
   hCtl  = frmHotImgBtn( HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_BTNMATCHWHOLEWORD, "IMAGE_WHOLED", "IMAGE_WHOLED",  _
                  L(162,"Match Whole Words"), iconHeight, iconHeight, GetSysColor(COLOR_WINDOW), GetSysColor(COLOR_WINDOW), _
                  nLeft, nTop, ctlHeight, ctlHeight, true, true, true)
   frmHotImgBtn_SetSelected(hCtl, gFind.nWholeWord)

   ' Display the Results label
   nTop  = margin
   nLeft = nLeft + pWindow->UnScaleX(AfxGetWindowWidth(hCtl)) + margin  
   hCtl = pWindow->AddControl("LABEL", , IDC_FRMFINDREPLACE_LBLRESULTS, L(269, "No results"), _
        nLeft, nTop, 80, ctlHeight, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_CENTER Or SS_CENTERIMAGE, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
      
   ' Display the Left button 
   nTop  = margin
   nLeft = nLeft + pWindow->UnScaleX(AfxGetWindowWidth(hCtl)) + margin  
   hCtl  = frmHotImgBtn( HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_BTNLEFT, "IMAGE_LEFTD", "IMAGE_LEFTD",  _
                  L(45,"Find Previous") & WSTR(" (Shift+F3)"), iconHeight, iconHeight, GetSysColor(COLOR_BTNFACE), GetSysColor(COLOR_3DLIGHT), _
                  nLeft, nTop, ctlHeight, ctlHeight, true, true, true)

   ' Display the Right button 
   nTop  = margin
   nLeft = nLeft + pWindow->UnScaleX(AfxGetWindowWidth(hCtl)) + margin  
   hCtl  = frmHotImgBtn( HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_BTNRIGHT, "IMAGE_RIGHTD", "IMAGE_RIGHTD",  _
                  L(44,"Find Next") & WSTR(" (F3)"), iconHeight, iconHeight, GetSysColor(COLOR_BTNFACE), GetSysColor(COLOR_3DLIGHT), _
                  nLeft, nTop, ctlHeight, ctlHeight, true, true, true)

   ' Display the Selection button 
   nTop  = margin
   nLeft = nLeft + pWindow->UnScaleX(AfxGetWindowWidth(hCtl)) + margin  
   hCtl  = frmHotImgBtn( HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_BTNSELECTION, "IMAGE_SELECTEDD", "IMAGE_SELECTEDD",  _
                  L(148,"Selection"), iconHeight, iconHeight, GetSysColor(COLOR_BTNFACE), GetSysColor(COLOR_3DLIGHT), _
                  nLeft, nTop, ctlHeight, ctlHeight, true, true, true)


   ' Display the Replace textbox               
   nTop  = margin + pWindow->UnScaleX(AfxGetWindowHeight(hCtlFind)) + vmargin
   nLeft = 20 + margin
   hCtlReplace = pWindow->AddControl("TEXTBOX", , IDC_FRMFINDREPLACE_TXTREPLACE, wstr(gFind.txtReplace), nLeft, nTop, 220, ctlHeight, _
              WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL, _
              WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   SendMessage hCtlReplace, EM_SETCUEBANNER, 1, CAST(LPARAM, @L(173,"Replace"))
      
   ' Display the Replace button 
   nTop  = margin + pWindow->UnScaleX(AfxGetWindowHeight(hCtlFind)) + vmargin
   nLeft = nLeft + pWindow->UnScaleX(AfxGetWindowWidth(hCtlReplace)) + margin  
   hCtl  = frmHotImgBtn( HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_BTNREPLACE, "IMAGE_REPLACED", "IMAGE_REPLACED",  _
                  L(173,"Replace"), iconHeight, iconHeight, GetSysColor(COLOR_BTNFACE), GetSysColor(COLOR_3DLIGHT), _
                  nLeft, nTop, ctlHeight, ctlHeight, true, true, true)

   ' Display the Replace All button 
   nTop  = margin + pWindow->UnScaleX(AfxGetWindowHeight(hCtlFind)) + vmargin
   nLeft = nLeft + pWindow->UnScaleX(AfxGetWindowWidth(hCtl)) + margin  
   hCtl  = frmHotImgBtn( HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_BTNREPLACEALL, "IMAGE_REPLACEALLD", "IMAGE_REPLACEALLD",  _
                  L(174,"Replace All"), iconHeight, iconHeight, GetSysColor(COLOR_BTNFACE), GetSysColor(COLOR_3DLIGHT), _
                  nLeft, nTop, ctlHeight, ctlHeight, true, true, true)

   ' Diplay the Close icon
   frmHotImgBtn( HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_BTNCLOSE, "IMAGE_CLOSE", "IMAGE_CLOSE",  _
                  L(161, "Close"), ctlHeight, ctlHeight, GetSysColor(COLOR_BTNFACE), GetSysColor(COLOR_3DLIGHT), _
                  frmWidth - ctlHeight - margin, margin, ctlHeight, ctlHeight, true, true, true)

   frmFindReplace_SetButtonStates
   frmFindReplace_ShowControls

   ' Position the window to the top right of the active edit window
   GetWindowRect( gTTabCtl.hWindow, @rc )
   SetWindowPos HWND_FRMFINDREPLACE, HWND_TOP, rc.right - pWindow->ScaleX(frmWidth + 50), rc.bottom, _
                  0, 0, SWP_SHOWWINDOW or SWP_NOSIZE or SWP_NOACTIVATE
   
   hCtl = GetDlgItem( HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_TXTFIND )
   if len(AfxGetWindowText(hCtl)) then Edit_SetSel(hCtl, 0, -1)
   frmFindReplace_HighlightSearches
   SetFocus hCtl 

   Function = 0
End Function

