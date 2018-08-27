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

'    VISUAL DESIGNER ROUTINES
'


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
      dim as long idx = TabCtrl_GetCurSel(pDoc->hDesignTabCtrl)
      bDesignView = iif(idx = 0, true, false)
   else
      ' Must be a code window
      bDesignView = false
   END IF

   function = bDesignView
END FUNCTION

' ========================================================================================
' Reset the references to the closest controls to the selected control.
' ========================================================================================
Function ResetClosestControls( byval pDoc as clsDocument ptr ) as long
   pDoc->pCtrlCloseLeft   = 0
   pDoc->pCtrlCloseTop    = 0
   pDoc->pCtrlCloseRight  = 0
   pDoc->pCtrlCloseBottom = 0
   function = 0
end function

' ========================================================================================
' Set the values for the blue color snap line.
' ========================================================================================
Function SetBlueSnapLine( byval pDoc as clsDocument ptr, _
                          byval x1 as long, byval y1 as long, _
                          byval x2 as long, byval y2 as long _
                          ) as long
   if pDoc = 0 THEN exit function
   pDoc->ptBlueStart.x = x1
   pDoc->ptBlueStart.y = y1
   pDoc->ptBlueEnd.x   = x2
   pDoc->ptBlueEnd.y   = y2
   function = 0
end function

      
' ========================================================================================
' Get the closest controls pointer to the current active control.
' ========================================================================================
Function SetClosestControls( byval pDoc as clsDocument ptr, _
                             byval pCtrlActive as clsControl ptr _
                             ) As clsControl ptr
   dim pCtrl as clsControl ptr
   dim as clsControl ptr pCtrlLeft, pCtrlTop, pCtrlRight, pCtrlBottom
   dim as long nLeft, nTop, nRight = 9999999, nBottom = 9999999
   dim as RECT rc, rcClient, rcActive, rcDest, rcTop, rcLeft, rcRight, rcBottom
   
   ResetClosestControls(pDoc)
   
   GetClientRect(pDoc->hWndForm, @rcClient)
   GetWindowRect(pDoc->pCtrlAction->hWindow, @rcActive)
   MapWindowPoints(0, pDoc->hWndForm, cast(point ptr, @rcActive), 2)
   
   ' Define the rectangle search areas based on the active control
   SetRect(@rcLeft, 0, rcActive.top, rcActive.left, rcActive.bottom)
   SetRect(@rcTop, rcActive.left, 0, rcActive.right, rcActive.top)
   SetRect(@rcRight, rcActive.right, rcActive.top, rcClient.right, rcActive.bottom)
   SetRect(@rcBottom, rcActive.left, rcActive.bottom, rcClient.bottom, rcActive.right)
   
   for i as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
      pCtrl = pDoc->Controls.ItemAt(i)
      if pCtrl->ControlType = CTRL_FORM then continue for
      if pCtrl->IsSelected = true THEN continue for
      if pCtrl = pCtrlActive THEN continue for
      
      GetWindowRect(pCtrl->hWindow, @rc)
      MapWindowPoints(0, pDoc->hWndForm, cast(point ptr, @rc), 2)
      
      ' LEFT
      if IntersectRect(@rcDest, @rcLeft, @rc) then
         if (rc.right > nLeft) THEN
            nLeft = rc.right: pCtrlLeft = pCtrl
         END IF
      end if
      ' TOP
      if IntersectRect(@rcDest, @rcTop, @rc) then
         if (rc.bottom > nTop) THEN
            nTop = rc.bottom: pCtrlTop = pCtrl
         END IF
      end if
      ' RIGHT
      if IntersectRect(@rcDest, @rcRight, @rc) then
         if (rc.left < nRight) THEN
            nRight = rc.left: pCtrlRight = pCtrl
         END IF
      end if
      ' BOTTOM
      if IntersectRect(@rcDest, @rcBottom, @rc) then
         if (rc.top < nBottom) THEN
            nBottom = rc.top: pCtrlBottom = pCtrl
         END IF
      end if
   next
   
   pDoc->pCtrlCloseLeft   = pCtrlLeft
   pDoc->pCtrlCloseTop    = pCtrlTop
   pDoc->pCtrlCloseRight  = pCtrlRight
   pDoc->pCtrlCloseBottom = pCtrlBottom

   function = 0
end function


' ========================================================================================
' Create and display popup control menu.
' ========================================================================================
Function DisplayControlPopupMenu( byval hwnd as hwnd, byval xpos as long, byval ypos as long) As Long
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
Function DisplayFormPopupMenu( byval hwnd as hwnd, byval xpos as long, byval ypos as long) As Long
   static as HMENU hPopupMenu
   
   If hPopupMenu Then DestroyMenu(hPopupMenu)
    
   hPopupMenu = CreatePopupMenu()
      AppendMenu hPopUpMenu, MF_ENABLED, IDM_PASTE, L(19,"&Paste")
      AppendMenu hPopUpMenu, MF_SEPARATOR, 0, ""
      AppendMenu hPopUpMenu, MF_ENABLED, IDM_MENUEDITOR, L(312,"&Menu Editor...")
      AppendMenu hPopUpMenu, MF_ENABLED, IDM_TOOLBAREDITOR, L(313,"Tool&bar Editor...")
      AppendMenu hPopUpMenu, MF_ENABLED, IDM_STATUSBAREDITOR, L(314,"&Statusbar Editor...")
   TrackPopupMenu(hPopupMenu, TPM_LEFTALIGN Or TPM_LEFTBUTTON, xpos, ypos, 0, hWnd, 0)

   function = 0
end function

' ========================================================================================
' Change the mouse cursor depending on selected Toolbox control.
' ========================================================================================
Function SetMouseCursor() As Long
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
function SetGrabHandleMouseCursor( byval pDoc as clsDocument ptr, _
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
function DrawGrabHandles( byval hDC as HDC, _
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

   ResetClosestControls(pDoc)


   ' #1: Determine if a grab handle has been clicked on
   pDoc->GrabHit = SetGrabHandleMouseCursor(pDoc, pt.x, pt.y, pCtrl)
   if pDoc->GrabHit <> GRAB_NOHIT THEN
      ' One of the sizing handles was clicked on   
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
               SetClosestControls( pDoc, pCtrl)
            END IF
         END IF
         pDoc->Controls.SetActiveControl( iif(pCtrl->IsSelected, hWndCtrl, 0) )
         frmMain_UpdateLineCol(HWND_FRMMAIN)
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
   
   ResetClosestControls(pDoc)

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
      else
         ' Create the selected Toolbox control.
         ' Need to modify rcLasso dimensions in case of HighDPI
         Dim pWindow As CWindow Ptr = AfxCWindowPtr(pDoc->hWndForm)
         SetRect(@rcCtrl, pWindow->UnScaleX(rcLasso.Left), pWindow->UnScaleY(rcLasso.Top), _
                         pWindow->UnScaleX(rcLasso.Right), pWindow->UnScaleY(rcLasso.Bottom))
         pCtrl = CreateToolboxControl( pDoc, GetActiveToolboxControlType, rcCtrl )
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
   
   SetBlueSnapLine(pDoc, -1, -1, -1, -1)
   SetActiveToolboxControl(CTRL_POINTER)
   SetMouseCursor
   
   ' Ensure the grab handles of form and controls are redrawn or hidden
   AfxRedrawWindow(pDoc->hWndFrame)
   AfxRedrawWindow(pDoc->hWndForm)
   frmMain_UpdateLineCol(HWND_FRMMAIN)
   
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
   
   SetBlueSnapLine(pDoc,-1,-1,-1,-1)
   
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
      SetControlProperty(pDoc->pCtrlAction, "LEFT", str(rc.left))
      SetControlProperty(pDoc->pCtrlAction, "TOP", str(rc.top))
      SetControlProperty(pDoc->pCtrlAction, "WIDTH", str(rc.right - rc.left))
      SetControlProperty(pDoc->pCtrlAction, "HEIGHT", str(rc.bottom - rc.top))
      ApplyControlProperties(pDoc, pDoc->pCtrlAction)
      pDoc->pCtrlAction->SuspendLayout = false
      
      ' Indicate that the file is now dirty and will need to be saved
      pDoc->UserModified = true
      pDoc->bRegenerateCode = true

      ' Ensure the grab handles are redrawn
      AfxRedrawWindow(hwnd)   ' HWND because could be form or frame
      frmMain_UpdateLineCol(HWND_FRMMAIN)
      DisplayPropertyList(pDoc)
      
   elseif pDoc->bMoving then
      ' Iterate controls to determine if we can snap to another control
      dim as long nDirection
      dim as RECT rcActive, rc
      dim as Boolean bSnap
      
'      SetClosestControls( pDoc, pDoc->pCtrlAction)

'      GetWindowRect(pDoc->pCtrlAction->hWindow, @rcActive)
'      MapWindowPoints(0, pDoc->hWndForm, cast(point ptr, @rcActive), 2)

'      if yDelta < 0 THEN nDirection = DIRECTION_TOP
'      if yDelta > 0 THEN nDirection = DIRECTION_BOTTOM
       
'      SetBlueSnapLine(pDoc, -1, -1, -1, -1)

'      if nDirection = DIRECTION_TOP then

'         if pDoc->SnapUpWait <> 0 THEN 
'            pDoc->SnapUpWait = pDoc->SnapUpWait - abs(yDelta)
'            pDoc->ptPrev.x = pt.x: pDoc->ptPrev.y = pt.y
'            if pDoc->SnapUpWait > 0 then exit function
'            pDoc->SnapUpWait = 0
'         END IF

'         pCtrl = pDoc->pCtrlCloseTop
'         if cbool(pCtrl <> 0) andalso (bSnap = false) THEN
'            GetWindowRect(pCtrl->hWindow, @rc)
'            MapWindowPoints(0, pDoc->hWndForm, cast(point ptr, @rc), 2)
'            if rcActive.top <= rc.top + 5 THEN 
'               yDelta = (rc.top - rcActive.top): bSnap = true
'            end if   
'         END IF
'         pCtrl = pDoc->pCtrlCloseLeft
'         if cbool(pCtrl <> 0) andalso (bSnap = false) THEN
'            GetWindowRect(pCtrl->hWindow, @rc)
'            MapWindowPoints(0, pDoc->hWndForm, cast(point ptr, @rc), 2)
'            if (rcActive.top > rc.top) andalso (rcActive.top <= rc.top + 5) THEN 
'               yDelta = (rc.top - rcActive.top): bSnap = true
'               pDoc->SnapUpWait = abs(yDelta)
'               SetBlueSnapLine(pDoc, rc.left, rc.top, rcActive.left, rcActive.top)
'            end if
'         END IF
'      end if
          
'      if nDirection = DIRECTION_BOTTOM then
''            ? "BOTTOM"
'      end if
             

      
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
      frmMain_UpdateLineCol(HWND_FRMMAIN)
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
      ' Draw any blue snap lines
      SelectObject(hDC, hPen)
      MoveToEx(hDC, pDoc->ptBlueStart.x, pDoc->ptBlueStart.y, 0)
      LineTo(hDC, pDoc->ptBlueEnd.x, pDoc->ptBlueStart.y)
      
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
Function DesignerForm_OnDestroy(HWnd As HWnd) As LRESULT
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(hWnd)
   if pWindow then Delete pWindow
   Function = 0
End Function


' ========================================================================================
' Visual Designer Form Window procedure 
' ========================================================================================
Function DesignerForm_WndProc( ByVal HWnd   As HWnd, _
                               ByVal uMsg   As UINT, _
                               ByVal wParam As WPARAM, _
                               ByVal lParam As LPARAM _
                               ) As LRESULT

   Select Case uMsg
      HANDLE_MSG (HWnd, WM_PAINT,       DesignerForm_OnPaint)
      HANDLE_MSG (HWnd, WM_DESTROY,     DesignerForm_OnDestroy)

      case WM_RBUTTONDOWN: HandleDesignerRButtonDown(HWND)
      case WM_LBUTTONDOWN: HandleDesignerLButtonDown(HWND)
      case WM_LBUTTONUP:   HandleDesignerLButtonUp(HWND)
      case WM_MOUSEMOVE:   HandleDesignerMouseMove(HWND)

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
                              
      case WM_CTLCOLORSTATIC, WM_CTLCOLOREDIT, WM_CTLCOLORBTN, WM_CTLCOLORLISTBOX 
         ' wParam: HDC of the control. 
         ' lParam: Handle to the control.
         ' Need to determine the child control handle that is sending the request to be painted.
         dim pCtrl as clsControl ptr 
         dim pDoc as clsDocument ptr = gApp.GetDocumentPtrByWindow(hwnd)
         if pDoc THEN 
            pCtrl = pDoc->Controls.GetCtrlPtr(cast(hwnd, lParam) )
            if pCtrl then
               dim pProp as clsProperty ptr
               pProp = GetControlPropertyPtr(pCtrl, "FORECOLOR")
               if pProp then 
                  SetTextColor(cast(HDC, wParam), GetRGBColorFromProperty(pProp))
               end if
               pProp = GetControlPropertyPtr(pCtrl, "BACKCOLOR")
               if pProp then 
                  dim as COLORREF clrBack = GetRGBColorFromProperty(pProp)
                  SetBkColor(cast(HDC, wParam), clrBack)
                  if pCtrl->hBackBrush then DeleteBrush(pCtrl->hBackBrush)
                  pCtrl->hBackBrush = CreateSolidBrush(clrBack)
                  Return Cast(LRESULT, pCtrl->hBackBrush)
               end if
            end if   
         end if
   
   End Select

   ' for messages that we don't deal with
   Function = DefWindowProc(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' Processes messages for the subclassed controls.
' ========================================================================================
FUNCTION Control_SubclassProc( BYVAL hwnd   AS HWND, _                 ' Control window handle
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

      case WM_RBUTTONDOWN: HandleDesignerRButtonDown(pDoc->hWndForm): return CTRUE
      case WM_LBUTTONDOWN: HandleDesignerLButtonDown(pDoc->hWndForm): return CTRUE
      case WM_LBUTTONUP:   HandleDesignerLButtonUp(pDoc->hWndForm): return CTRUE
      case WM_MOUSEMOVE:   HandleDesignerMouseMove(pDoc->hWndForm)

      Case WM_MOUSEACTIVATE
         ' Defeat this message so that mouse clicks do not activate the control. However, we
         ' do need to set focus to the control in order to allow keyboard movem/resize.
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



