'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2022 Paul Squires, PlanetSquires Software

'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT any WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS for A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.

'    VISUAL DESIGNER ROUTINES
'

#include once "modVDDesignForm.bi"


' ========================================================================================
' Determines if the current view should be Design or Code. This is done by checking
' the current selection of the design/code tabcontrol.
' ========================================================================================
function IsDesignerView( byval pDoc as clsDocument ptr ) as Boolean
   if pDoc = 0 THEN exit function

   dim as Boolean bDesignView

   ' If this is a Visual Designer document then the display depends on what tab is selected
   ' in the design|code tabcontrol.
   if pDoc->IsDesigner then 
      bDesignView = iif(pDoc->DesignTabsCurSel = 0, true, false)
   else
      ' Must be a code window
      bDesignView = false
   END IF

   function = bDesignView
end function


' ========================================================================================
' Reset any SnapLines data in order to ready it for future control movements.
' ========================================================================================
function ResetSnapLines( byval pDoc as clsDocument ptr ) as Long
   if pDoc->bSnapLines = false then exit function
   
   dim pt as POINT
   
   for i as long = SnapLinePosition.top to SnapLinePosition.right
      ' Hide any visible SnapLines
      ShowWindow( pDoc->hSnapLine(i), SW_HIDE )
      pDoc->bSnapActive(i) = false
      pDoc->ptCursorStart(i) = pt
   next
   
   function = 0
end function


' ========================================================================================
' Calculate any SnapLines and show them on screen.
' Returns True if SnapLines have been set and calling program should perform the snap.
' ========================================================================================
Function PerformSnapLines( _
            byval pDoc as clsDocument ptr, _
            byval pCtrlActive as clsControl ptr, _
            byval ptCursor as POINT, _
            byref xDelta as long, _
            byref yDelta as long _
            ) As long
                               
   if pDoc->bSnapLines = false then exit function
   
   ' Incoming ptCursor is in client coordinates
   
   dim pCtrl as clsControl ptr
   dim as RECT rc, rcClient, rcIntersect, rcActive
   
   dim as RECT rcTest1   ' the hit area for the control being tested
   dim as RECT rcTest2   ' the hit area for the active control 
    
   dim as long HitPixels = AfxScaleY(8)
   
   dim as Boolean bHit(3)
   dim as RECT rcSnap(3)
   dim as long nShowWindow
   dim as long xCursorDistance, yCursorDistance
   dim as long nSnap(3)
   dim as Boolean bDone
   
   ' Get the rectangle of the current active control being moved or sized 
   GetClientRect( pDoc->hWndForm, @rcClient )
   GetWindowRect( pDoc->pCtrlAction->hWindow, @rcActive )
   MapWindowPoints( 0, pDoc->hWndForm, cast(point ptr, @rcActive), 2 )
    
   ' NOTE: Based on testing, including the BOTTOM and RIGHT tests can cause
   ' jerkiness behaviour in the visual designer. For example, a control can
   ' sometimes get "stuck" alternating between a LEFT hit and a RIGHT hit
   ' causing the control to immediately snap LEFT and then when you try to
   ' move, it will snap RIGHT. Being able to dislodge the control from this
   ' behaviour is very hard.
   ' Therefore, until a better algorithm can be found, we will just do hit
   ' testing for TOP and LEFT.
   '
   for i as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
      pCtrl = pDoc->Controls.ItemAt(i)
      if pCtrl->ControlType = CTRL_FORM then continue for
      if pCtrl->IsSelected = true THEN continue for
      if pCtrl = pCtrlActive THEN continue for
       
      ' Get the control rectangle to test
      GetWindowRect( pCtrl->hWindow, @rc )
      MapWindowPoints( 0, pDoc->hWndForm, cast(point ptr, @rc), 2 )
      
      ' TOP TEST
      SetRect( @rcTest1, 0, rc.top - HitPixels, rcClient.right, rc.top + HitPixels )
      SetRect( @rcTest2, 0, rcActive.top, rcClient.right, rcActive.top + 1 )
      if IntersectRect( @rcIntersect, @rcTest1, @rcTest2 ) then
         bHit(SnapLinePosition.top) = true
         if IsRectEmpty( @rcSnap(SnapLinePosition.top) ) then
            SetRect( @rcSnap(SnapLinePosition.top), rcActive.left, rc.top - 1, rcActive.right, rc.top )
         end if
         if rc.left < rcSnap(SnapLinePosition.top).left then rcSnap(SnapLinePosition.top).left = rc.left
         if rc.right > rcSnap(SnapLinePosition.top).right then rcSnap(SnapLinePosition.top).right = rc.right
         nSnap(SnapLinePosition.top) = rc.top
      end if

'      ' BOTTOM TEST
'      if bHit(SnapLinePosition.top) = false then
'         SetRect( @rcTest1, 0, rc.bottom - HitPixels, rcClient.right, rc.bottom + HitPixels )
'         SetRect( @rcTest2, 0, rcActive.bottom, rcClient.right, rcActive.bottom + 1 )
'         if IntersectRect( @rcIntersect, @rcTest1, @rcTest2 ) then
'            if bHit(SnapLinePosition.top) = false then
'               bHit(SnapLinePosition.bottom) = true
'               if IsRectEmpty( @rcSnap(SnapLinePosition.bottom) ) then
'                  SetRect( @rcSnap(SnapLinePosition.bottom), rcActive.left, rc.bottom - 1, rcActive.right, rc.bottom )
'               end if
'               if rc.left < rcSnap(SnapLinePosition.bottom).left then rcSnap(SnapLinePosition.bottom).left = rc.left
'               if rc.right > rcSnap(SnapLinePosition.bottom).right then rcSnap(SnapLinePosition.bottom).right = rc.right
'               nSnap(SnapLinePosition.bottom) = rc.bottom
'            end if
'         end if
'      end if
      
      ' LEFT TEST
      SetRect( @rcTest1, rc.left - HitPixels, 0, rc.left + HitPixels, rcClient.bottom )
      SetRect( @rcTest2, rcActive.left, 0, rcActive.left + 1, rcClient.bottom )
      if IntersectRect( @rcIntersect, @rcTest1, @rcTest2 ) then
         bHit(SnapLinePosition.left) = true
         if IsRectEmpty( @rcSnap(SnapLinePosition.left) ) then
            SetRect( @rcSnap(SnapLinePosition.left), rc.left - 1, rcActive.top, rc.left, rcActive.bottom )
         end if
         if rc.top < rcSnap(SnapLinePosition.left).top then rcSnap(SnapLinePosition.left).top = rc.top
         if rc.bottom > rcSnap(SnapLinePosition.left).bottom then rcSnap(SnapLinePosition.left).bottom = rc.bottom
         nSnap(SnapLinePosition.left) = rc.left
      end if

'      ' RIGHT TEST
'      if bHit(SnapLinePosition.left) = false then
'         SetRect( @rcTest1, rc.right - HitPixels, 0, rc.right + HitPixels, rcClient.bottom )
'         SetRect( @rcTest2, rcActive.right, 0, rcActive.right + 1, rcClient.bottom )
'         if IntersectRect( @rcIntersect, @rcTest1, @rcTest2 ) then
'            if bHit(SnapLinePosition.left) = false then
'               bHit(SnapLinePosition.right) = true
'               if IsRectEmpty( @rcSnap(SnapLinePosition.right) ) then
'                  SetRect( @rcSnap(SnapLinePosition.right), rc.right - 1, rcActive.top, rc.right, rcActive.bottom )
'               end if
'               if rc.top < rcSnap(SnapLinePosition.right).top then rcSnap(SnapLinePosition.right).top = rc.top
'               if rc.bottom > rcSnap(SnapLinePosition.right).bottom then rcSnap(SnapLinePosition.right).bottom = rc.bottom
'               nSnap(SnapLinePosition.right) = rc.right
'            end if
'         end if
'      end if
   next


   ' Testing Sequence
   ' Top, Left, Bottom, Right
   for i as long = SnapLinePosition.top to SnapLinePosition.right

      if bHit(i) then
         ' If the snap has not performed then do it now and save the current screen cursor position.
         if pDoc->bSnapActive(i) = false then
            pDoc->bSnapActive(i) = true
            pDoc->ptCursorStart(i) = ptCursor
            select case i
               case SnapLinePosition.top
                  yDelta = nSnap(i) - rcActive.top 
               case SnapLinePosition.left
                  xDelta = nSnap(i) - rcActive.left 
               case SnapLinePosition.bottom      
                  yDelta = nSnap(i) - rcActive.bottom
               case SnapLinePosition.right
                  xDelta = nSnap(i) - rcActive.right 
            end select
            
         else
            ' A previous SnapLine action has occurred and the control is still snapped
            ' to the snapline. Need to test if the cursor has moved far enough outside
            ' the capture area to free the control from the snap.
            xCursorDistance = ( ptCursor.x - pDoc->ptCursorStart(i).x )
            yCursorDistance = ( ptCursor.y - pDoc->ptCursorStart(i).y )
            select case i
               case SnapLinePosition.top, SnapLinePosition.bottom
                  if abs(yCursorDistance) > HitPixels then
                     pDoc->bSnapActive(SnapLinePosition.top) = false
                     pDoc->bSnapActive(SnapLinePosition.bottom) = false
                     bHit(SnapLinePosition.left) = false  ' prevent further tests
                     bHit(SnapLinePosition.right) = false  ' prevent further tests
                     bHit(SnapLinePosition.top) = false  ' prevent further tests
                     bHit(SnapLinePosition.bottom) = false  ' prevent further tests
                     yDelta = yCursorDistance
                     bDone = true
                  else
                     ' We are still snapped and still within the snap area. Therefore, do not
                     ' allow the control to move by setting yDelta to zero.
                     yDelta = 0
                  end if
               case SnapLinePosition.left, SnapLinePosition.right     
                  if abs(xCursorDistance) > HitPixels then
                     pDoc->bSnapActive(SnapLinePosition.left) = false
                     pDoc->bSnapActive(SnapLinePosition.right) = false
                     bHit(SnapLinePosition.left) = false  ' prevent further tests
                     bHit(SnapLinePosition.right) = false  ' prevent further tests
                     bHit(SnapLinePosition.top) = false  ' prevent further tests
                     bHit(SnapLinePosition.bottom) = false  ' prevent further tests
                     xDelta = xCursorDistance
                     bDone = true
                  else
                     ' We are still snapped and still within the snap area. Therefore, do not
                     ' allow the control to move by setting xDelta to zero.
                     xDelta = 0
                  end if
            end select
            
         end if
      end if   
      
      nShowWindow = iif( bHit(i), SWP_SHOWWINDOW, SWP_HIDEWINDOW )
      if pDoc->bSnapActive(i) then nShowWindow = SWP_SHOWWINDOW
      
      SetWindowPos( pDoc->hSnapLine(i), HWND_TOP, _
                    rcSnap(i).left, rcSnap(i).top, _
                    rcSnap(i).right - rcSnap(i).left, _
                    rcSnap(i).bottom - rcSnap(i).top, _
                    nShowWindow )
      if bDone then exit for
   next              
   
   function = 0
end function


' ========================================================================================
' Create and display popup control menu.
' ========================================================================================
function DisplayControlPopupMenu( _
            byval hwnd as hwnd, _
            byval xpos as long, _
            byval ypos as long _
            ) As Long
   static as HMENU hPopupMenu
   
   If hPopupMenu Then DestroyMenu(hPopupMenu)
    
   hPopupMenu = CreatePopupMenu()
      AppendMenu hPopUpMenu, MF_ENABLED, IDM_CUT, L(17,"Cu&t") 
      AppendMenu hPopUpMenu, MF_ENABLED, IDM_COPY, L(18,"&Copy")
      AppendMenu hPopUpMenu, MF_ENABLED, IDM_PASTE, L(19,"&Paste")
      AppendMenu hPopUpMenu, MF_SEPARATOR, 0, ""
      AppendMenu hPopUpMenu, MF_ENABLED, IDM_DELETE, L(326,"&Delete")
   TrackPopupMenu(hPopupMenu, TPM_LEFTALIGN Or TPM_LEFTBUTTON, xpos, ypos, 0, hWnd, 0)

   function = 0
end function


' ========================================================================================
' Create and display popup form menu.
' ========================================================================================
Function DisplayFormPopupMenu( _
            byval hwnd as hwnd, _
            byval xpos as long, _
            byval ypos as long _
            ) As Long
   static as HMENU hPopupMenu
   
   If hPopupMenu Then DestroyMenu(hPopupMenu)
    
   hPopupMenu = CreatePopupMenu()
      AppendMenu hPopUpMenu, MF_ENABLED, IDM_PASTE, L(19,"Paste")
      AppendMenu hPopUpMenu, MF_SEPARATOR, 0, ""
      AppendMenu hPopUpMenu, MF_ENABLED, IDM_MENUEDITOR, L(312,"Menu Editor") & "..."
      AppendMenu hPopUpMenu, MF_ENABLED, IDM_TOOLBAREDITOR, L(313,"Toolbar Editor") & "..."
      AppendMenu hPopUpMenu, MF_ENABLED, IDM_STATUSBAREDITOR, L(314,"Statusbar Editor") & "..."
   TrackPopupMenu(hPopupMenu, TPM_LEFTALIGN Or TPM_LEFTBUTTON, xpos, ypos, 0, hWnd, 0)

   function = 0
end function


' ========================================================================================
' Change the mouse cursor depending on selected Toolbox control.
' ========================================================================================
function SetMouseCursor() As Long
   Dim As HWnd hList1 = GetDlgItem(HWND_FRMVDTOOLBOX, IDC_FRMVDTOOLBOX_LSTTOOLBOX)
   dim as long nCurSel = ListBox_GetCurSel(hList1)
   
   if nCurSel = -1 THEN exit function
   
   ' The index into the global gToolbox array is stored in the line's data area.
   dim as long idx = ListBox_GetItemData(hList1, nCurSel)
   if idx > 0 THEN
      SetCursor LoadCursor( GetModuleHandle(NULL), *gToolBox(idx).wszCursor)
   else
      SetCursor LoadCursor( 0, ByVal IDC_ARROW )
   END IF
   function = 0
End Function


' ========================================================================================
' Change the mouse cursor if over a valid grab handle
' ========================================================================================
function SetGrabHandleMouseCursor( _
            byval pDoc as clsDocument ptr, _
            byval x as long, _
            byval y as long, _
            byref pCtrlAction as clsControl Ptr _
            ) as LRESULT
   
   dim pCtrl as clsControl ptr
   dim rcCtrl as RECT
   dim pt as point: pt.x = x: pt.y = y   ' The point in is client coordinates
   
   for i as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
      pCtrl = pDoc->Controls.ItemAt(i)
      for ii as long = GRAB_TOPLEFT to GRAB_LEFT
         if PtInRect(@pCtrl->rcHandles(ii), pt) then
            ' Test to ensure that the specific control Locked property is not set, or
            ' the global Locked setting for the entire form is not set.
            if (pDoc->bLockControls = true) or (GetControlProperty(pCtrl, "Locked") = "True") then 
               SetCursor( LoadCursor(Null, ByVal IDC_NO) ): return GRAB_NOHIT
            end if
            select case ii
               Case GRAB_TOP, GRAB_BOTTOM:          SetCursor( LoadCursor(Null, ByVal IDC_SIZENS) )
               case GRAB_LEFT, GRAB_RIGHT:          SetCursor( LoadCursor(Null, ByVal IDC_SIZEWE) )
               case GRAB_TOPLEFT, GRAB_BOTTOMRIGHT: SetCursor( LoadCursor(Null, ByVal IDC_SIZENWSE) )
               case GRAB_TOPRIGHT, GRAB_BOTTOMLEFT: SetCursor( LoadCursor(Null, ByVal IDC_SIZENESW) )
            end select
            pCtrlAction = pCtrl
            return ii   ' return hit code GRAB_TOPLEFT to GRAB_LEFT         
         end if 
      NEXT
   next

   for i as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
      pCtrl = pDoc->Controls.ItemAt(i)
      if pCtrl->ControlType = CTRL_FORM THEN continue for
      GetWindowRect(pCtrl->hWindow, @rcCtrl)
      MapWindowPoints(0, pDoc->hWndForm, cast(point ptr, @rcCtrl), 2)
      if PtInRect(@rcCtrl, pt) THEN
         ' Test to ensure that the specific control Locked property is not set, or
         ' the global Locked setting for the entire form is not set.
         if (pDoc->bLockControls = true) or (GetControlProperty(pCtrl, "Locked") = "True") then 
            ' We will continue to allow the Arrow cursor for Locked controls in order to
            ' indicate that the control can be selected. If the user attempts to move the 
            ' control then the cursor will change to IDC_NO
            return GRAB_NOHIT
         else
            ' Cursor is over a valid control
            SetCursor( LoadCursor(Null, ByVal IDC_SIZEALL) ): return GRAB_NOHIT
         end if   
      END IF  
   next
   
   function = GRAB_NOHIT
      
end function


' ========================================================================================
' Calculate the size of form/controls grab handle rectangles 
' ========================================================================================
function CalculateGrabHandles( byval pDoc as clsDocument ptr) as long
   if pDoc = 0 THEN exit function
      
   dim pCtrl as clsControl ptr 

   dim as long cx = AfxScaleX(6)
   dim as long cy = AfxScaleY(6)
   dim as long HMid, VMid    ' horiz and vert middles
   
   for i as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
      pCtrl = pDoc->Controls.ItemAt(i)

      dim rc as RECT
      GetWindowRect(pCtrl->hWindow, @rc)
      MapWindowPoints(0, GetParent(pCtrl->hWindow), cast(point ptr, @rc), 2)
      
      ' Calculate the grab handle rectangles
      HMid = (rc.right - rc.left) / 2
      VMid = (rc.bottom - rc.top) / 2
      '
      '      1     2     3
      '
      '      8           4
      '
      '      7     6     5
      '
      ' Only calculate the grab handle rectangles if the control is selected
      if pCtrl->IsSelected THEN
         SetRect(@pCtrl->rcHandles(GRAB_TOPLEFT), rc.left-cx, rc.top-cy, rc.left, rc.top)
         SetRect(@pCtrl->rcHandles(GRAB_TOP), rc.left+HMid-(cx/2), rc.top-cy, rc.left+HMid+(cx/2), rc.top)
         SetRect(@pCtrl->rcHandles(GRAB_TOPRIGHT), rc.right, rc.top-cy, rc.right+cx, rc.top)
         SetRect(@pCtrl->rcHandles(GRAB_RIGHT), rc.right, rc.top+VMid-(cy/2), rc.right+cx, rc.top+VMid+(cy/2))
         SetRect(@pCtrl->rcHandles(GRAB_BOTTOMRIGHT), rc.right, rc.bottom, rc.right+cx, rc.bottom+cy)
         SetRect(@pCtrl->rcHandles(GRAB_BOTTOM), rc.left+HMid-(cx/2), rc.bottom, rc.left+HMid+(cx/2), rc.bottom+cy)
         SetRect(@pCtrl->rcHandles(GRAB_BOTTOMLEFT), rc.left-cx, rc.bottom, rc.left, rc.bottom+cy)
         SetRect(@pCtrl->rcHandles(GRAB_LEFT), rc.left-cx, rc.top+VMid-(cy/2), rc.left, rc.top+VMid+(cy/2))
      else
         for i as long = GRAB_TOPLEFT to GRAB_LEFT
            SetRectEmpty(@pCtrl->rcHandles(i))
         NEXT
      end if
      
      ' If this is a form then we don't want to display some of the grab handles
      ' so simply set them to be empty rectangles.
      if pCtrl->ControlType = CTRL_FORM THEN
         SetRectEmpty(@pCtrl->rcHandles(GRAB_TOPLEFT))
         SetRectEmpty(@pCtrl->rcHandles(GRAB_TOP))
         SetRectEmpty(@pCtrl->rcHandles(GRAB_TOPRIGHT))
         SetRectEmpty(@pCtrl->rcHandles(GRAB_BOTTOMLEFT))
         SetRectEmpty(@pCtrl->rcHandles(GRAB_LEFT))
      end if
         
   NEXT

   function = 0
end function

   
' ========================================================================================
' Draw the actual grab handles (this is called from WM_PAINT)
' ========================================================================================
function DrawGrabHandles( _
            byval hDC as HDC, _
            byval pDoc as clsDocument ptr, _ 
            byval bFormOnly as Boolean _
            ) as long

   if pDoc = 0 THEN exit function
      
   dim pCtrl as clsControl ptr 

   CalculateGrabHandles(pDoc)
   
   SaveDC hDC

   dim as LOGBRUSH LogBrush
   LogBrush.lbColor = BGR(0,0,0)
   LogBrush.lbStyle = PS_SOLID
   dim as HPEN hDottedPen = ExtCreatePen( PS_COSMETIC or PS_ALTERNATE, 1, @LogBrush, 0, NULL )
   dim as HPEN hSolidPen = CreatePen(PS_SOLID, 1, BGR(0,0,0))
   dim as HBRUSH hWhiteBrush = CreateSolidBrush(BGR(255,255,255))
   dim as HBRUSH hBlackBrush = CreateSolidBrush(BGR(0,0,0))
   dim as HBRUSH hRedBrush   = CreateSolidBrush(BGR(255,0,0))
   
   for i as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
      pCtrl = pDoc->Controls.ItemAt(i)

      if bFormOnly THEN
         if pCtrl->ControlType <> CTRL_FORM THEN continue for
      else
         if pCtrl->ControlType = CTRL_FORM THEN continue for
      end if

      dim rc as RECT
      GetWindowRect(pCtrl->hWindow, @rc)
      MapWindowPoints(0, GetParent(pCtrl->hWindow), cast(point ptr, @rc), 2)
      
      ' Draw the actual grab handles
      if pCtrl->IsSelected THEN
         ' Draw the dotted rectangle around the control
         dim as long nOffset = AfxScaleX(2)
         SelectObject( hDC, hDottedPen )
         SelectObject( hDC, GetStockObject( NULL_BRUSH ) )
         Rectangle(hDC, rc.left-nOffset, rc.top-nOffset, rc.right+nOffset, rc.bottom+nOffset)

         SelectObject( hDC, hSolidPen )
         
         ' If Form level locking is set or the Control is locked then paint in red.
         if IsControlLocked(pDoc, pCtrl) then
            SelectObject( hDC, hRedBrush )
         else   
            SelectObject( hDC, iif(pCtrl->IsActive, hWhiteBrush, hBlackBrush) )
         end if
         
         for ii as long = GRAB_TOPLEFT to GRAB_LEFT
            if pCtrl->IsActive THEN
               RoundRect(hDC, pCtrl->rcHandles(ii).left, pCtrl->rcHandles(ii).top, _
                              pCtrl->rcHandles(ii).right, pCtrl->rcHandles(ii).bottom, 2, 2 )
            else
               ' Make the non-active control grab handles a little smaller so as not to 
               ' visually overpower the active control's white handles.
               Rectangle(hDC, pCtrl->rcHandles(ii).left, pCtrl->rcHandles(ii).top, _
                              pCtrl->rcHandles(ii).right-2, pCtrl->rcHandles(ii).bottom-2 )
            END IF
         NEXT
      END IF
   next                           

   RestoreDC hDC, -1 

   DeleteObject(hDottedPen)
   DeleteObject(hSolidPen)
   DeleteObject(hWhiteBrush)
   DeleteObject(hBlackBrush)
   DeleteObject(hRedBrush)
   
   function = 0   
end function


' ========================================================================================
' Handle WM_LBUTTONDBLCLK messages for the Form (displays the code editor)
' ========================================================================================
function HandleDesignerLButtonDoubleClick( ByVal HWnd As HWnd ) as LRESULT
   ' When a Control is double clicked, it would have already set the default Event
   ' in the PropertyList. Therefore, all we need to do is call the doubleclick 
   ' handler for the Event listbox. This will create (if necessary) the Event 
   ' handler code and then position the code editor to that Event.
   dim as HWND hEvents = GetDlgItem( HWND_FRMVDTOOLBOX, IDC_FRMVDTOOLBOX_LSTEVENTS )
   SendMessage( HWND_FRMVDTOOLBOX, WM_COMMAND, MAKELONG(IDC_FRMVDTOOLBOX_LSTEVENTS, LBN_DBLCLK), cast(LPARAM,hEvents) )
   function = 0
end function


' ========================================================================================
' Handle WM_LBUTTONDOWN messages for the Form and Frame windows
' ========================================================================================
function HandleDesignerLButtonDown( ByVal HWnd As HWnd ) as LRESULT

   dim as POINT pt
   dim as RECT rc
   

   dim pDoc as clsDocument ptr = gApp.GetDocumentPtrByWindow(hwnd)
   if pDoc = 0 THEN exit function

   dim pCtrl as clsControl ptr

   GetCursorPos(@pt)
   MapWindowPoints(0, HWND, @pt, 1)

   ' Ensure that the cursor stays within the client area
   GetClientRect(hwnd, @rc)
   MapWindowPoints(hwnd, 0, cast(point ptr, @rc), 2)
   ClipCursor(@rc)
   
   SetCapture(hwnd)


   ' #1: Determine if a grab handle has been clicked on
   pDoc->GrabHit = SetGrabHandleMouseCursor(pDoc, pt.x, pt.y, pCtrl)
   if pDoc->GrabHit <> GRAB_NOHIT THEN
      ' One of the sizing handles was clicked on.
      ' Do not allow resizing of Timer controls.   
      if pCtrl->ControlType = CTRL_TIMER then
         ClipCursor(0)
         ReleaseCapture
         exit function
      end if
      pDoc->bSizing = true
      pDoc->pCtrlAction = pCtrl
      GetWindowRect(pDoc->pCtrlAction->hWindow, @pDoc->rcSize)
      pDoc->Controls.SetActiveControl(pDoc->pCtrlAction->hWindow)
   else
      ' #2: Determine what control/form was clicked on
      dim as hwnd hWndCtrl = RealChildWindowFromPoint(pDoc->hWndForm, pt)
      pCtrl = pDoc->Controls.GetCtrlPtr(hWndCtrl)
      if pCtrl then
         ' If Ctrl is held down then toggle adding/removing the control 
         if (GetAsyncKeyState(VK_CONTROL) and &H8000) THEN
            ' If we are individually selecting/deselecting controls then ensure that
            ' the form control itself is not part of the group otherwise any cut/copy
            ' operation will fail with a GPF.
            ' The form itself can not be part of a selected group
            if pCtrl->ControlType <> CTRL_FORM then
               dim pCtrlForm as clsControl ptr = GetFormCtrlPtr(pDoc)
               if pCtrlForm then pCtrlForm->IsSelected = false
               pCtrl->IsSelected = not(pCtrl->IsSelected)
            END IF
         else
            ' If the control being clicked on is already selected then it will become
            ' the active control. If not already selected then deselect all other controls
            ' in the selection group.
            if pCtrl->IsSelected = false THEN pDoc->Controls.DeselectAllControls
            pCtrl->IsSelected = true
            ' If the Form is clicked on the start the lasso process.
            if pCtrl->ControlType = CTRL_FORM THEN
               gLasso.Create(pDoc->hWndForm)
               gLasso.SetStartPoint(pt.x, pt.y)
               gLasso.SetEndPoint(pt.x, pt.y)
            else
               pDoc->pCtrlAction = pCtrl
               pDoc->bMoving = true
            END IF
         END IF
         pDoc->Controls.SetActiveControl( iif(pCtrl->IsSelected, hWndCtrl, 0) )
         frmMain_SetStatusbar
      END IF
   END IF

   
   ' Save the current mouse position
   pDoc->ptPrev.x = pt.x
   pDoc->ptPrev.y = pt.y
   
   ' Ensure the grab handles of form and controls are redrawn or hidden
   AfxRedrawWindow(pDoc->hWndFrame)
   AfxRedrawWindow(pDoc->hWndForm)
   DisplayPropertyList(pDoc)

   function = 0
end function


' ========================================================================================
' Handle WM_LBUTTONUP messages for the Form and Frame windows
' ========================================================================================
function HandleDesignerLButtonUp( ByVal HWnd As HWnd ) as LRESULT

   ClipCursor(0)
   ReleaseCapture

   dim pDoc as clsDocument ptr = gApp.GetDocumentPtrByWindow(hwnd)
   if pDoc = 0 THEN exit function
   
   dim as Rect rcIntersect, rcLasso, rcCtrl
   
   ' Reset any previous SnapLines data in order to ready it for any
   ' future movement of controls on the form.
   ResetSnapLines( pDoc )

   
   ' Hide any previous lasso (and select controls)
   if gLasso.IsActive THEN 
      dim pCtrl as clsControl ptr
      
      rcLasso = gLasso.GetLassoRect()
      gLasso.Destroy

      ' If the Toolbox Pointer/Arrow is selected then attempt to select the controls
      ' that intersect with the lasso, otherwise draw and create the new Toolbox control.
      if GetActiveToolboxControlType = CTRL_POINTER THEN
         dim as hwnd hCtrlSel
         pDoc->Controls.DeselectAllControls
         MapWindowPoints(pDoc->hWndForm, 0, cast(point ptr, @rcLasso), 2)
         for i as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
            pCtrl = pDoc->Controls.ItemAt(i)
            if pCtrl->ControlType <> CTRL_FORM THEN
               GetWindowRect( pCtrl->hWindow, @rcCtrl)
               If IntersectRect( @rcIntersect, @rcCtrl, @rcLasso ) Then
                  hCtrlSel = pCtrl->hWindow
                  pCtrl->IsSelected = true
               end if
            end if
         next
         if hCtrlSel = 0 THEN hCtrlSel = pDoc->hWndForm
         pDoc->Controls.SelectControl(hCtrlSel)
         pDoc->Controls.SetActiveControl(hCtrlSel)
         DisplayPropertyList(pDoc)
      else
         ' Create the selected Toolbox control.
         ' Need to modify rcLasso dimensions in case of HighDPI
         Dim pWindow As CWindow Ptr = AfxCWindowPtr(pDoc->hWndForm)
         SetRect( @rcCtrl, pWindow->UnScaleX(rcLasso.Left), pWindow->UnScaleY(rcLasso.Top), _
                         pWindow->UnScaleX(rcLasso.Right), pWindow->UnScaleY(rcLasso.Bottom) )
         ' Check for a minimum default size when creating. Check needed because a user could
         ' select a control in the Toolbox and simply click on the Form rather than drawing
         ' the control on the form.
         dim as long nControlType = GetActiveToolboxControlType()
         rcCtrl = CheckMinimumControlSize( nControlType, rcCtrl )
         pCtrl = CreateToolboxControl( pDoc, nControlType, rcCtrl )
         pDoc->UserModified = true
         pDoc->bRegenerateCode = true
         pDoc->Controls.SelectControl(pCtrl->hWindow)
         pDoc->Controls.SetActiveControl(pCtrl->hWindow)
         DisplayPropertyList(pDoc)
      end if   
   end if      
    
   pDoc->GrabHit  = GRAB_NOHIT
   pDoc->bSizing  = false
   pDoc->bMoving  = false
   pDoc->pCtrlAction = 0
   
   SetActiveToolboxControl(CTRL_POINTER)
   SetMouseCursor
   
   ' Ensure the grab handles of form and controls are redrawn or hidden
   AfxRedrawWindow(pDoc->hWndFrame)
   AfxRedrawWindow(pDoc->hWndForm)
   frmMain_SetStatusbar
   
   function = 0
end function


' ========================================================================================
' Handle WM_RBUTTONDOWN messages for the Form and Frame windows
' ========================================================================================
function HandleDesignerRButtonDown( ByVal HWnd As HWnd ) as LRESULT

   dim as POINT pt
   
   dim pDoc as clsDocument ptr = gApp.GetDocumentPtrByWindow(hwnd)
   if pDoc = 0 THEN exit function

   dim pCtrl as clsControl ptr

   ' Call LButtonDown to select control
   HandleDesignerLButtonDown(HWnd) 
   ClipCursor(0)
   ReleaseCapture

   GetCursorPos(@pt)
   pCtrl = pDoc->Controls.GetActiveControl
   if pCtrl THEN
      if pCtrl->ControlType = CTRL_FORM THEN
         DisplayFormPopupMenu(HWND_FRMMAIN, pt.x, pt.y)
      else   
         DisplayControlPopupMenu(HWND_FRMMAIN, pt.x, pt.y)
      END IF
   END IF

   ' Call LButtonUp to reset selections
   HandleDesignerLButtonUp(hwnd)

   function = 0
end function


' ========================================================================================
' Handle MOUSEMOVE messages for the Form and Frame windows
' ========================================================================================
function HandleDesignerMouseMove( ByVal HWnd As HWnd ) as LRESULT

   dim as POINT pt
   dim as long xDelta, yDelta
   
   dim pDoc as clsDocument ptr = gApp.GetDocumentPtrByWindow(hwnd)
   if pDoc = 0 THEN exit function
   
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(pDoc->hWndForm)
   dim pCtrl as clsControl ptr

   GetCursorPos(@pt)
   MapWindowPoints(0, HWND, @pt, 1)
   xDelta = pt.x - pDoc->ptPrev.x
   yDelta = pt.y - pDoc->ptPrev.y
   
   if (xDelta = 0) andalso (yDelta = 0) THEN exit function

   ' Test to ensure that the specific control Locked property is not set, or
   ' the global Locked setting for the entire form is not set.
   if (pDoc->bMoving) andalso IsControlLocked(pDoc, pDoc->pCtrlAction) then 
      SetCursor( LoadCursor(Null, ByVal IDC_NO) )
      exit function
   end if

   ' If the mouse cursor is not resizing or moving a control then determine the cursor
   ' based on what control is selected in the toolbox.
   if (pDoc->bSizing = false) andalso (pDoc->bMoving = false) then
      SetMouseCursor
   end if
   
   
   if pDoc->bSizing THEN

      Select Case pDoc->GrabHit
         Case GRAB_BOTTOMRIGHT 
            pDoc->rcSize.right  = pDoc->rcSize.right  + xDelta
            pDoc->rcSize.bottom = pDoc->rcSize.bottom + yDelta
         Case GRAB_RIGHT 
            pDoc->rcSize.right  = pDoc->rcSize.right + xDelta
         Case GRAB_BOTTOM
            pDoc->rcSize.bottom = pDoc->rcSize.bottom + yDelta
         case GRAB_BOTTOMLEFT
            pDoc->rcSize.left   = pDoc->rcSize.left   + xDelta
            pDoc->rcSize.bottom = pDoc->rcSize.bottom + yDelta
         Case GRAB_TOPLEFT 
            pDoc->rcSize.left   = pDoc->rcSize.left + xDelta
            pDoc->rcSize.top    = pDoc->rcSize.top  + yDelta
         Case GRAB_TOPRIGHT
            pDoc->rcSize.right  = pDoc->rcSize.right + xDelta
            pDoc->rcSize.top    = pDoc->rcSize.top   + yDelta
         Case GRAB_LEFT 
            pDoc->rcSize.left   = pDoc->rcSize.left + xDelta
         Case GRAB_TOP
            pDoc->rcSize.top    = pDoc->rcSize.top + yDelta
      End Select

      ' NOTE:
      ' Set our control to a default minimum value If Zero.
      ' We want to do this so we don't loose visibility of our control and we can still 
      ' see the handles when selected.

      ' Check for a minimum width and height
      If pDoc->rcSize.right - pDoc->rcSize.left <= AfxScaleX(8) Then pDoc->rcSize.right = pDoc->rcSize.left + AfxScaleX(8)
      If pDoc->rcSize.bottom - pDoc->rcSize.top <= AfxScaleY(8) Then pDoc->rcSize.bottom = pDoc->rcSize.top + AfxScaleY(8)

      ' Resize all selected the form/control 
      ' Convert pDoc->rcSize from Window to Client coordinates
      dim as rect rc = pDoc->rcSize
      MapWindowPoints(0, HWND, cast(point ptr, @rc), 2)
      ' Ensure that the rect is unscaled
      SetRect(@rc, pWindow->UnScaleX(rc.Left), pWindow->UnScaleY(rc.Top), _
                   pWindow->UnScaleX(rc.Right), pWindow->UnScaleY(rc.Bottom))
      
      pDoc->pCtrlAction->SuspendLayout = true
      SetControlProperty( pDoc->pCtrlAction, "LEFT", str(rc.left) )
      SetControlProperty( pDoc->pCtrlAction, "TOP", str(rc.top) )
      SetControlProperty( pDoc->pCtrlAction, "WIDTH", str(rc.right - rc.left) )
      SetControlProperty( pDoc->pCtrlAction, "HEIGHT", str(rc.bottom - rc.top) )
      ApplyControlProperties( pDoc, pDoc->pCtrlAction )
      pDoc->pCtrlAction->SuspendLayout = false
      
      ' If a Menu, ToolBar or StatusBar exists on the Form then ensure that it
      ' resizes to the new Form width.
      if pDoc->pCtrlAction->ControlType = CTRL_FORM then
         frmMenuEditor_CreateFakeMainMenu(pDoc)
         if pDoc->hWndRebar then
            dim as long nTopMenu_Height, nRebar_Height 
            if pDoc->hWndFakeMenu then
               nTopMenu_Height = AfxGetWindowHeight( pDoc->hWndFakeMenu )  
            end if 
            GetClientRect( hwnd, @rc )
            dim as long cx, cy
            cx = rc.left - rc.right
            cy = AfxGetWindowHeight( pDoc->hWndRebar )
            SetWindowPos pDoc->hWndRebar, 0, 0, nTopMenu_Height, rc.right, cy, SWP_NOZORDER
            SendMessage pDoc->hWndRebar, WM_SIZE, cx, cy
         end if
         frmStatusBarEditor_CreateFakeStatusBar(pDoc)
      end if   

      ' Indicate that the file is now dirty and will need to be saved
      pDoc->UserModified = true
      pDoc->bRegenerateCode = true

      ' Ensure the grab handles are redrawn
      AfxRedrawWindow(hwnd)   ' HWND because could be form or frame
      frmMain_SetStatusbar
      DisplayPropertyList(pDoc)
      
   elseif pDoc->bMoving then
      
      ' Determine if snap to Snaplines
      PerformSnapLines( pDoc, pDoc->pCtrlAction, pt, xDelta, yDelta )
      
      ' Move the control to its new position
      for i as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
         pCtrl = pDoc->Controls.ItemAt(i)
         if pCtrl->IsSelected THEN
            if IsControlLocked(pDoc, pCtrl) then continue for
            GetWindowRect(pCtrl->hWindow, @pDoc->rcSize)
            MapWindowPoints(0, pDoc->hWndForm, cast(point ptr, @pDoc->rcSize), 2)
      
            pDoc->rcSize.left  = pDoc->rcSize.left + xDelta
            pDoc->rcSize.top   = pDoc->rcSize.top + yDelta
            
            ' Ensure that the rect is unscaled
            SetRect(@pDoc->rcSize, pWindow->UnScaleX(pDoc->rcSize.Left), pWindow->UnScaleY(pDoc->rcSize.Top), _
                                   pWindow->UnScaleX(pDoc->rcSize.Right), pWindow->UnScaleY(pDoc->rcSize.Bottom))
            pCtrl->SuspendLayout = true
            SetControlProperty(pCtrl, "LEFT", str(pDoc->rcSize.left))
            SetControlProperty(pCtrl, "TOP", str(pDoc->rcSize.top))
            ApplyControlProperties(pDoc, pCtrl)
            pCtrl->SuspendLayout = false
         END IF
      next
      
      ' Indicate that the file is now dirty and will need to be saved
      pDoc->UserModified = true
      pDoc->bRegenerateCode = true

      ' Ensure the grab handles are redrawn
      AfxRedrawWindow(pDoc->hWndForm)
      frmMain_SetStatusbar
      DisplayPropertyList(pDoc)
      
   elseif gLasso.IsActive then
      gLasso.SetEndPoint(pt.x, pt.y)
      gLasso.Show()
   else   
      pDoc->GrabHit = SetGrabHandleMouseCursor(pDoc, pt.x, pt.y, pCtrl)
   end if      

   ' Save the current mouse position
   pDoc->ptPrev.x = pt.x
   pDoc->ptPrev.y = pt.y

   function = 0
   
END FUNCTION


' ========================================================================================
' Process WM_PAINT message for Visual Designer Form
' ========================================================================================
Function DesignerForm_OnPaint( ByVal HWnd As HWnd) As LRESULT

   Dim pWindow As CWindow Ptr = AfxCWindowPtr(hWnd)
   If pWindow = 0 Then Exit Function
   
   Dim As PAINTSTRUCT ps
   Dim As HDC hDC
   dim as HPEN hPen = CreatePen(PS_SOLID, 1, BGR(0,0,255))
   
   hDC = BeginPaint(hWnd, @ps)

   SaveDC hDC

   dim pDoc as clsDocument ptr = gApp.GetDocumentPtrByWindow(hwnd)
   if pDoc THEN
      ' Draw the control's grab handles 
      DrawGrabHandles(hDC, pDoc, false)
   end if   

   RestoreDC hDC, -1 
   DeleteObject(hPen)
   EndPaint hWnd, @ps
   
   ' Repaint any Frame controls because they disappear when the form's background is painted.
   if pDoc THEN
      dim pCtrl as clsControl ptr
      for i as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
         pCtrl = pDoc->Controls.ItemAt(i)
         if (pCtrl <> 0) andalso (pCtrl->ControlType = CTRL_FRAME) then
            AfxRedrawWindow(pCtrl->hWindow)
         end if
      next
   end if
         
   Function = 0

End Function


' ========================================================================================
' Process WM_DESTROY message for Visual Designer Form
' ========================================================================================
function DesignerForm_OnDestroy( byval HWnd As HWnd ) As LRESULT
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(hWnd)
   if pWindow then Delete pWindow
   Function = 0
End Function


' ========================================================================================
' Process WM_NOTIFY message for Visual Designer Form
' ========================================================================================
function DesignerForm_OnNotify( _
            ByVal HWnd As HWnd, _
            ByVal id As Long, _
            ByVal pNMHDR As NMHDR Ptr _
            ) As LRESULT

   Select Case pNMHDR->code 
      Case NM_CLICK
         if id = IDC_FAKESTATUSBAR then
            ' Determine if one of the Panels was clicked on. If it was then send
            ' that value to the routine that will open the sttausbar editor 
            ' defaulting to that panel as being selected.
            Dim lpnm As NMMOUSE Ptr = Cast(NMMOUSE Ptr, pNMHDR)
            ' lpnm->dwItemSpec  ' index of clicked panel
            PostMessage( HWND_FRMMAIN, WM_COMMAND, _
                         MAKELONG(IDM_STATUSBAREDITOR, lpnm->dwItemSpec), 0 )
         end if
   end select
   
   function = 0
end function


' ========================================================================================
' Process WM_COMMAND message for Visual Designer Form
' ========================================================================================
function DesignerForm_OnCommand( _
            ByVal HWnd As HWnd, _
            ByVal id As Long, _
            ByVal hwndCtl As HWnd, _
            ByVal codeNotify As UINT _
            ) As LRESULT
   Select Case id
      Case IDC_LBLFAKEMAINMENU
         If codeNotify = BN_CLICKED Then
            PostMessage( HWND_FRMMAIN, WM_COMMAND, MAKELONG(IDM_MENUEDITOR, 0), 0 )
         end if
   end select
   
   function = 0
end function


' ========================================================================================
' Visual Designer Form Window procedure 
' ========================================================================================
function DesignerForm_WndProc( _
            ByVal HWnd   As HWnd, _
            ByVal uMsg   As UINT, _
            ByVal wParam As WPARAM, _
            ByVal lParam As LPARAM _
            ) As LRESULT

   Select Case uMsg
      HANDLE_MSG (HWnd, WM_PAINT,   DesignerForm_OnPaint)
      HANDLE_MSG (HWnd, WM_NOTIFY,  DesignerForm_OnNotify)
      HANDLE_MSG (HWnd, WM_COMMAND, DesignerForm_OnCommand)
      HANDLE_MSG (HWnd, WM_DESTROY, DesignerForm_OnDestroy)

      case WM_LBUTTONDBLCLK: HandleDesignerLButtonDoubleClick(HWND)
      case WM_RBUTTONDOWN:   HandleDesignerRButtonDown(HWND)
      case WM_LBUTTONDOWN:   HandleDesignerLButtonDown(HWND)
      case WM_LBUTTONUP:     HandleDesignerLButtonUp(HWND)
      case WM_MOUSEMOVE:     HandleDesignerMouseMove(HWND)

      Case WM_NCLBUTTONDOWN, WM_NCRBUTTONDOWN, WM_NCLBUTTONDBLCLK  
         ' Will only fire when the Caption bar, Min/Max/Close buttons are clicked.
         SetActiveWindow HWND_FRMMAIN
         dim pDoc as clsDocument ptr = gApp.GetDocumentPtrByWindow(hwnd)
         if pDoc THEN 
            pDoc->Controls.DeselectAllControls
            pDoc->Controls.SetActiveControl(pDoc->hWndForm)
            AfxRedrawWindow(pDoc->hWndFrame)
            AfxRedrawWindow(pDoc->hWndForm)
            DisplayPropertyList(pDoc)
         END IF
         Function = TRUE: Exit Function


      Case WM_NCHITTEST
         ' Catch certain critical mouseover points on the form so we can stop processing them.
         dim as LRESULT nHitTest = DefWindowProc(hWnd, uMsg, wParam, lParam)
         
         Select Case nHitTest
            ' Border edges of the window and captionbar
            Case HTLEFT, HTTOP, HTTOPLEFT, HTTOPRIGHT, _
                 HTBOTTOMLEFT, HTRIGHT, HTBOTTOM, HTBOTTOMRIGHT, _
                 HTCLOSE, HTMENU, HTMINBUTTON, HTMAXBUTTON 
                 Function = 0  ' Return zero so the mousepointer will not change
                 Exit Function                                    
         End Select      
                                      
         Function = nHitTest    ' Return the default code from the default window handler.
         Exit Function
                              

      case WM_CTLCOLOREDIT, WM_CTLCOLORBTN, WM_CTLCOLORLISTBOX, WM_CTLCOLORSTATIC
         ' wParam: HDC of the control. 
         ' lParam: Handle to the control.
         ' Need to determine the child control handle that is sending the request to be painted.
         dim pCtrl as clsControl ptr 
         dim pDoc as clsDocument ptr = gApp.GetDocumentPtrByWindow(hwnd)
         if pDoc THEN 
            pCtrl = pDoc->Controls.GetCtrlPtr(cast(hwnd, lParam) )
            if pCtrl then
               ' Colors for WinFormsX Label controls are set in WM_DRAWITEM
               if pCtrl->ControlType <> CTRL_LABEL then
                  dim pProp as clsProperty ptr
                  pProp = GetControlPropertyPtr(pCtrl, "FORECOLOR")
                  if pProp then 
                     SetTextColor(cast(HDC, wParam), GetRGBColorFromProperty(pProp->wszPropValue))
                  end if
                  pProp = GetControlPropertyPtr(pCtrl, "BACKCOLOR")
                  if pProp then 
                     dim as COLORREF clrBack = GetRGBColorFromProperty(pProp->wszPropValue)
                     SetBkColor(cast(HDC, wParam), clrBack)
                     if pCtrl->hBackBrush then DeleteBrush(pCtrl->hBackBrush)
                     pCtrl->hBackBrush = CreateSolidBrush(clrBack)
                     Return Cast(LRESULT, pCtrl->hBackBrush)
                  end if
               end if
            end if   
         end if
         
         ' We would have handled any FRAME control backcolor in the code above. The following code
         ' deals with color for fake menu and snap lines.
         if uMsg = WM_CTLCOLORSTATIC then
            if pDoc THEN 
               ' Need to determine if the incoming control handle relates to
               ' a snapline or is the fake topmenu label.
               dim as hwnd hCtrl = cast(HWND, lParam)
               if hCtrl = pDoc->hWndFakeMenu then
                  Return Cast( LRESULT, GetSysColorBrush(COLOR_MENU+1) )
               else
                  for i as long = 0 to 3
                     if hCtrl = pDoc->hSnapLine(i) then
                        ' SnapLines labels have purple brush
                        if pDoc->hBrushSnapLine = 0 then 
                           pDoc->hBrushSnapLine = CreateSolidBrush( BGR(128,0,128) )  ' purple
                        end if
                        Return Cast(LRESULT, pDoc->hBrushSnapLine)
                     end if
                  next
               end if
            end if
         end if            


      case WM_DRAWITEM
         dim lpdis As DRAWITEMSTRUCT Ptr = cast( DRAWITEMSTRUCT Ptr, lParam )
         if lpdis = 0 then exit function
         dim pDoc as clsDocument ptr = gApp.GetDocumentPtrByWindow(hwnd)
         if pDoc = 0 THEN exit function
            
         if lpdis->hwndItem = pDoc->hWndStatusBar then
            ' StatusBar panels are OwnerDraw
            Dim memDC as HDC      ' Double buffering
            Dim hbit As HBITMAP   ' Double buffering
            Dim As RECT rc
            Dim wszText As WString * MAX_PATH
            dim as long nImageWidth 
            dim as long nImageHeight

            rc = lpdis->rcItem
            
            dim as long nItem = lpdis->itemID
            If (nItem < 0) or (nItem > ubound(pDoc->PanelItems)) Then Exit Function

            ' The image/text output rectangle is positioned within the main rc rectangle
            ' depending on the alignment.
            dim as RECT rc1 = rc
            
            dim wszImageName as wstring * MAX_PATH
            wszImageName = pDoc->PanelItems(nItem).pProp.wszPropValue
         
            dim as HANDLE hIcon
            dim hImageListNormal As HIMAGELIST 
            if len(wszImageName) then
               Dim cx As Long = AfxScaleX(16) 
               dim pImageType as IMAGES_TYPE ptr = GetImagesTypePtr(wszImageName)
               hImageListNormal = ImageList_Create( cx, cx, ILC_COLOR32 Or ILC_MASK, 1, 1)
               dim as long ii = AfxGdipAddIconFromFile( hImageListNormal, pImageType->wszFileName )
               hIcon = ImageList_GetIcon( hImageListNormal, ii, ILD_NORMAL )
               if hIcon then
                  nImageWidth = AfxScaleX(16)
                  nImageHeight = AfxScaleY(16)
               end if   
            End If   
            
            wszText = " " & pDoc->PanelItems(nItem).wszText & " "
            dim as long nTextWidth = GetTextWidthPixels( lpdis->hwndItem, wszText )  

            dim as long nTotalWidth 
            if nImageWidth then nTotalWidth = nImageWidth + AfxScaleX(4)
            if nTextWidth  then nTotalWidth = nTotalWidth + nTextWidth 
            
            select case pDoc->PanelItems(nItem).wszAlignment
               case "StatusBarPanelAlignment.Left"
                  ' No need to do anything because rc1 is already the full size of rc.
               case "StatusBarPanelAlignment.Center"
                  dim as long nPad = MAX(0, (rc.right - rc.left - nTotalWidth) / 2)
                  rc1.left = rc1.left + nPad
               case "StatusBarPanelAlignment.Right"
                  rc1.left  = rc1.right - nTotalWidth
            end select
            
            memDC = CreateCompatibleDC( lpdis->hDC )
            hbit  = CreateCompatibleBitmap( lpdis->hDC, rc.right, rc.bottom )
            If hbit Then hbit = SelectObject( memDC, hbit )

            dim as HFONT _hFont = AfxGetWindowFont( lpdis->hwndItem )
            SelectObject( memDC, _hFont )

            ' Paint the entire background
            dim as HBRUSH hBackBrush
            dim as COLORREF rgbBackClr, rgbForeClr 

            ' In the visual designer we only display the panels in the "non hot" state.
            rgbBackClr = GetRGBColorFromProperty(pDoc->PanelItems(nItem).wszBackColor)
            rgbForeClr = GetRGBColorFromProperty(pDoc->PanelItems(nItem).wszForeColor)

            SetBkColor( memDC, rgbBackClr )   
            SetTextColor( memDC, rgbForeClr )
            hBackBrush = CreateSolidBrush(rgbBackClr)
            FillRect( memDC, @rc, hBackBrush )
            DeleteObject(hBackBrush)

            ' Output any defined icon for the panel
            if hIcon then
               ' Center the image vertically within rc1
               dim as long nPad = (rc1.bottom - rc1.top - nImageHeight) / 2
               DrawIconEx( memDC, rc1.left, rc1.top + nPad, _
                                  hIcon, _
                                  nImageWidth, nImageHeight, 0, null, DI_NORMAL ) 
               DeleteObject( hIcon )
               ImageList_Destroy( hImageListNormal )
               rc1.left = rc1.left + nImageWidth + AfxScaleX(4)
            end if
            
            ' Prepare and paint the text coloring
            dim as long lFormat = DT_LEFT Or DT_VCENTER or DT_SINGLELINE
            wszText = pDoc->PanelItems(nItem).wszText
            DrawText( memDC, wszText, -1, Cast(lpRect, @rc1), lFormat )

            BitBlt lpdis->hDC, 0, 0, rc.right, rc.bottom, memDC, 0, 0, SRCCOPY 

            ' Cleanup
            If hbit  Then DeleteObject SelectObject(memDC, hbit)
            If memDC Then DeleteDC memDC
         
         else
            ' Label controls are OwnerDraw
            dim pCtrl as clsControl ptr = pDoc->Controls.GetCtrlPtr( lpdis->hwndItem )
            if pCtrl then
               dim pProp as clsProperty ptr
               
               select case pCtrl->ControlType
                  case CTRL_LABEL
                     pProp = GetControlPropertyPtr( pCtrl, "FORECOLOR" )
                     if pProp then SetTextColor( lpdis->hDC, GetRGBColorFromProperty(pProp->wszPropValue) )
                     pProp = GetControlPropertyPtr(pCtrl, "BACKCOLOR")
                     if pProp then 
                        dim as COLORREF clrBack = GetRGBColorFromProperty(pProp->wszPropValue)
                        SetBkColor( lpdis->hDC, clrBack )
                        dim as RECT rc
                        GetClientRect( lpdis->hwndItem, @rc )
                        if pCtrl->hBackBrush then DeleteBrush( pCtrl->hBackBrush )
                        pCtrl->hBackBrush = CreateSolidBrush( clrBack )
                        FillRect( lpdis->hDC, @rc, pCtrl->hBackBrush )
                     end if
                     dim wszText as WString * MAX_PATH
                     
                     ' Windows draws Labels with SS_LEFT, CC_CENTER and SS_RIGHT styles by using
                     ' DrawText with the DT_WORDBREAK and DT_EXPANDTABS parameters. Other styles
                     ' must use DT_SINGLELINE and do not word wrap.
                     dim as long lWrapMode = DT_SINGLELINE
                     dim as long lFormat 
                     Select Case **GetControlProperty(pCtrl, "TEXTALIGN")
                        Case "LabelAlignment.BottomCenter": lFormat = DT_CENTER Or DT_BOTTOM
                        Case "LabelAlignment.BottomLeft":   lFormat = DT_LEFT   Or DT_BOTTOM
                        Case "LabelAlignment.BottomRight":  lFormat = DT_RIGHT  Or DT_BOTTOM
                        Case "LabelAlignment.MiddleCenter": lFormat = DT_CENTER Or DT_VCENTER
                        Case "LabelAlignment.MiddleLeft":   lFormat = DT_LEFT   Or DT_VCENTER
                        Case "LabelAlignment.MiddleRight":  lFormat = DT_RIGHT  Or DT_VCENTER
                        Case "LabelAlignment.TopCenter"
                           lFormat = DT_CENTER Or DT_TOP
                           lWrapMode = DT_WORDBREAK
                        Case "LabelAlignment.TopLeft"
                           lFormat = DT_LEFT Or DT_TOP
                           lWrapMode = DT_WORDBREAK
                        Case "LabelAlignment.TopRight"
                           lFormat = DT_RIGHT Or DT_TOP
                           lWrapMode = DT_WORDBREAK
                     End Select
                     lFormat = lFormat or lWrapMode or DT_EXPANDTABS

                     wszText = AfxGetWindowText( lpdis->hwndItem )
                     DrawText( lpdis->hDC, wszText, -1, Cast(lpRect, @lpdis->rcItem), lFormat )
               end select
            end if   
         end if
            
   End Select

   ' for messages that we don't deal with
   Function = DefWindowProc(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' Processes messages for the subclassed controls.
' ========================================================================================
function Control_SubclassProc( _
            BYVAL hwnd   AS HWND, _                 ' Control window handle
            BYVAL uMsg   AS UINT, _                 ' Type of message
            BYVAL wParam AS WPARAM, _               ' First message parameter
            BYVAL lParam AS LPARAM, _               ' Second message parameter
            BYVAL uIdSubclass AS UINT_PTR, _        ' The subclass ID
            BYVAL dwRefData AS DWORD_PTR _          ' Pointer to reference data
            ) AS LRESULT

   dim pDoc as clsDocument ptr = cast(clsDocument ptr, dwRefData)
   
   SELECT CASE uMsg

      CASE WM_GETDLGCODE
         ' All keyboard input
         FUNCTION = DLGC_WANTALLKEYS
         EXIT FUNCTION

      case WM_RBUTTONDOWN:   HandleDesignerRButtonDown(pDoc->hWndForm): return CTRUE
      case WM_LBUTTONDOWN:   HandleDesignerLButtonDown(pDoc->hWndForm): return CTRUE
      case WM_LBUTTONDBLCLK: HandleDesignerLButtonDoubleClick(pDoc->hWndForm): return CTRUE
      case WM_LBUTTONUP:     HandleDesignerLButtonUp(pDoc->hWndForm): return CTRUE
      case WM_MOUSEMOVE:     HandleDesignerMouseMove(pDoc->hWndForm)

      Case WM_MOUSEACTIVATE
         ' Defeat this message so that mouse clicks do not activate the control. However, we
         ' do need to set focus to the control in order to allow keyboard move/resize.
         SetFocus(hwnd) 
         Function = MA_NOACTIVATE: uMsg = WM_NULL
         Exit Function

      Case WM_SETCURSOR
         Function = CTRUE: uMsg = WM_NULL
         Exit Function
          
      Case WM_SETFOCUS
         ' Defeat the caret activation, for some
         ' reason MA_NOACTIVATE does not work for right clicks.
         Function = 0: uMsg = WM_NULL
         Exit Function

      CASE WM_DESTROY
         ' REQUIRED: Remove control subclassing
         RemoveWindowSubclass( hwnd, @Control_SubclassProc, uIdSubclass )

   END SELECT

   ' Default processing of Windows messages
   FUNCTION = DefSubclassProc(hwnd, uMsg, wParam, lParam)

END FUNCTION



