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
' Show Visual Designer ToolBox
' ========================================================================================
Function OnCommand_DesignerToolBox( ByVal HWnd As HWnd ) As LRESULT
   frmVDToolbox_Show( HWnd )   ' toggle show/hide
   function = 0
end function

' ========================================================================================
' Show Visual Designer Menu Editor
' ========================================================================================
Function OnCommand_DesignerMenuEditor( _
            ByVal HWnd As HWnd, _
            byval pDoc as clsDocument ptr _
            ) As LRESULT
   frmMenuEditor_Show( HWND )
   if pDoc then frmMenuEditor_CreateFakeMainMenu( pDoc )
   function = 0
end function

' ========================================================================================
' Show Visual Designer ToolBar Editor
' ========================================================================================
Function OnCommand_DesignerToolBarEditor( _
            ByVal HWnd As HWnd, _
            byval pDoc as clsDocument ptr _
            ) As LRESULT
   frmToolBarEditor_Show( HWND )
   if pDoc then frmToolBarEditor_CreateFakeToolBar( pDoc )
   function = 0
end function

' ========================================================================================
' Show Visual Designer StatusBar Editor
' ========================================================================================
Function OnCommand_DesignerStatusBarEditor( _
            ByVal HWnd As HWnd, _
            byval pDoc as clsDocument ptr, _
            ByVal codeNotify As UINT _
            ) As LRESULT
   ' If a Panel was clicked on then WPARAM would have been set to the
   ' ControlID and panel index. When message cracked, that value would show as the
   ' incoming codeNotify for this message.
   ' Do a check to ensure that the panel click index is within the
   ' acceptable ranges.
   if pDoc then 
      dim as long nIndex = cast(long, codeNotify)
      if (nIndex < lbound(pDoc->PanelItems)) or _
         (nIndex > ubound(pDoc->PanelItems)) then
         nIndex = 0
      end if   
      frmStatusBarEditor_Show( HWND, nIndex )
      frmStatusBarEditor_CreateFakeStatusBar( pDoc )
   end if
   function = 0
end function

' ========================================================================================
' Show Visual Designer Image Manager
' ========================================================================================
Function OnCommand_DesignerImageManager( _
            ByVal HWnd As HWnd, _
            byval pDoc as clsDocument ptr _
            ) As LRESULT
   frmImageManager_Show( HWND, 0 )
   function = 0
end function

' ========================================================================================
' Visual Designer Snap Lines 
' ========================================================================================
Function OnCommand_DesignerSnapLines( byval pDoc as clsDocument ptr ) As LRESULT
   if pDoc then 
      pDoc->bSnapLines = not pDoc->bSnapLines
      pDoc->UserModified = true
      frmMain_SetStatusbar
   end if
   function = 0
end function

' ========================================================================================
' Visual Designer Lock Controls
' ========================================================================================
Function OnCommand_DesignerLockControls( byval pDoc as clsDocument ptr ) As LRESULT
   if pDoc then 
      pDoc->bLockControls = not pDoc->bLockControls
      pDoc->UserModified = true
      AfxRedrawWindow(pDoc->hWndFrame)
      AfxRedrawWindow(pDoc->hWndForm)
      frmMain_SetStatusbar
   end if
   function = 0
end function


' ========================================================================================
' Create a new visual designer window
' ========================================================================================
Function OnCommand_DesignerNewForm( ByVal HWnd As HWnd ) As LRESULT
   frmMain_OpenFileSafely( HWnd, _
                           True,  _   ' bIsNewFile
                           False, _   ' bIsTemplate
                           True,  _   ' bShowInTab
                           false, _   ' bIsInclude
                           "", _      ' wszName
                           0, _       ' pDocIn
                           true )     ' Create a visual designer   
   Function = 0
End Function


' ========================================================================================
' Horizonal Spacing (controls on the form)
' ========================================================================================
function OnCommand_DesignerHorizSpacing( _
            byval HWND as HWND, _
            byval id as long _
            ) as LRESULT

   dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr
   if pDoc = 0 THEN exit function

   dim pCtrlActive as clsControl ptr = pDoc->Controls.GetActiveControl()
   if pCtrlActive = 0 THEN exit function

   Dim pWindow As CWindow Ptr = AfxCWindowPtr(hwnd)
   if pWindow = 0 then exit function

   dim pCtrl as clsControl ptr
   dim as RECT rcCtrl, rcActive

   dim as long NumSelected = pDoc->Controls.SelectedControlsCount 
   dim as long nTotalRectWidth, nLeft

   ' Load all selected controls into array for easy processing
   dim pCtrls( 1 to NumSelected ) as clsControl ptr
   
   dim as long nextCtrl = 1
   for i as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
      pCtrl = pDoc->Controls.ItemAt(i)
      if pCtrl->IsSelected THEN
         rcCtrl = GetControlRECT(pCtrl)
         nTotalRectWidth = nTotalRectWidth + (rcCtrl.right - rcCtrl.left)
         pCtrls(nextCtrl) = pCtrl
         nextCtrl = nextCtrl + 1
      END IF
   next

   ' Sort the array based on Left (small to large)(bubble sort)
   dim done as long
   do
      done = 0
      For i as long = 1 To NumSelected - 1
         If val(GetControlProperty(pCtrls(i), "LEFT")) > val(GetControlProperty(pCtrls(i+1), "LEFT")) Then
            Swap pCtrls(i), pCtrls(i+1)
            done = 1
         End If
      Next
    Loop Until done = 0
       
   dim as RECT rcCtrlLeft  = GetControlRECT(pCtrls(1))
   dim as RECT rcCtrlRight = GetControlRECT(pCtrls(NumSelected))
   dim as long nFreeSpace = rcCtrlRight.right - rcCtrlLeft.left - nTotalRectWidth
   dim as long nEqualSpace = nFreeSpace / (NumSelected - 1)

   If id = IDM_HORIZREMOVE Then nEqualSpace = 0

   ' Determine the position in the array of the active control
   dim as long nActiveControlIndex = 0
   for i as long = 1 to NumSelected 
      if pCtrls(i) = pCtrlActive then 
         nActiveControlIndex = i: exit for
      end if
   next 

   ' Adjust all of the controls to the left of the active control and then adjust
   ' all of the controls to the right of the active control.
   for i as long = nActiveControlIndex - 1 to 1 step -1
      pCtrl = pCtrls(i)
      if IsControlLocked(pDoc, pCtrl) then continue for
      
      ' Get coordinates of current (Left) and (Right) controls
      rcCtrlLeft  = GetControlRECT( pCtrls(i) )
      rcCtrlRight = GetControlRECT( pCtrls(i+1) )

      select case id 
         case IDM_HORIZEQUAL, IDM_HORIZREMOVE
            nLeft = rcCtrlLeft.left + (rcCtrlRight.left - rcCtrlLeft.right) - nEqualSpace
         case IDM_HORIZINCREASE
            nLeft = rcCtrlLeft.left - 8
         case IDM_HORIZDECREASE
            nLeft = rcCtrlLeft.left + 8
      end select

      SetControlProperty( pCtrl, "LEFT", str(nLeft) )
      ApplyControlProperties( pDoc, pCtrl )
   next

   for i as long = nActiveControlIndex + 1 to NumSelected
      pCtrl = pCtrls(i)
      if IsControlLocked(pDoc, pCtrl) then continue for

      ' Get coordinates of current (Left) and (Right) controls
      rcCtrlLeft  = GetControlRECT( pCtrls(i-1) )
      rcCtrlRight = GetControlRECT( pCtrls(i) )

      select case id 
         case IDM_HORIZEQUAL, IDM_HORIZREMOVE
            nLeft = rcCtrlRight.left + nEqualSpace - (rcCtrlRight.left - rcCtrlLeft.right)  
         case IDM_HORIZINCREASE
            nLeft = rcCtrlRight.left + 8
         case IDM_HORIZDECREASE
            nLeft = rcCtrlRight.left - 8
      end select
             
      SetControlProperty( pCtrl, "LEFT", str(nLeft) )
      ApplyControlProperties( pDoc, pCtrl )
   next
   
   ' Ensure the grab handles are redrawn
   AfxRedrawWindow(pDoc->hWndForm)

   function = 0
end function


' ========================================================================================
' Vertical Spacing (controls on the form)
' ========================================================================================
function OnCommand_DesignerVertSpacing( _
            byval HWND as HWND, _
            byval id as long _
            ) as LRESULT

   dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr
   if pDoc = 0 THEN exit function

   dim pCtrlActive as clsControl ptr = pDoc->Controls.GetActiveControl()
   if pCtrlActive = 0 THEN exit function

   Dim pWindow As CWindow Ptr = AfxCWindowPtr(hwnd)
   if pWindow = 0 then exit function

   dim pCtrl as clsControl ptr
   dim as RECT rcCtrl, rcActive

   dim as long NumSelected = pDoc->Controls.SelectedControlsCount 
   dim as long nTotalRectHeight, nTop

   ' Load all selected controls into array for easy processing
   dim pCtrls( 1 to NumSelected ) as clsControl ptr
   
   dim as long nextCtrl = 1
   for i as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
      pCtrl = pDoc->Controls.ItemAt(i)
      if pCtrl->IsSelected THEN
         rcCtrl = GetControlRECT(pCtrl)
         nTotalRectHeight = nTotalRectHeight + (rcCtrl.bottom - rcCtrl.top)
         pCtrls(nextCtrl) = pCtrl
         nextCtrl = nextCtrl + 1
      END IF
   next
   
   ' Sort the array based on top (small to large)(bubble sort)
   dim done as long
   do
      done = 0
      For i as long = 1 To NumSelected - 1
         If val(GetControlProperty(pCtrls(i), "TOP")) > val(GetControlProperty(pCtrls(i+1), "TOP")) Then
            Swap pCtrls(i), pCtrls(i+1)
            done = 1
         End If
      Next
    Loop Until done = 0
       
   dim as RECT rcCtrlTop    = GetControlRECT(pCtrls(1))
   dim as RECT rcCtrlBottom = GetControlRECT(pCtrls(NumSelected))
   dim as long nFreeSpace = rcCtrlBottom.bottom - rcCtrlTop.top - nTotalRectHeight
   dim as long nEqualSpace = nFreeSpace / (NumSelected - 1)

   If id = IDM_VERTREMOVE Then nEqualSpace = 0

   ' Determine the position in the array of the active control
   dim as long nActiveControlIndex = 0
   for i as long = 1 to NumSelected 
      if pCtrls(i) = pCtrlActive then 
         nActiveControlIndex = i: exit for
      end if
   next 

   ' Adjust all of the controls above the active control and then adjust
   ' all of the controls below the active control.
   for i as long = nActiveControlIndex - 1 to 1 step -1
      pCtrl = pCtrls(i)
      if IsControlLocked(pDoc, pCtrl) then continue for
      
      ' Get coordinates of current (Left) and (Right) controls
      rcCtrlTop    = GetControlRECT( pCtrls(i) )
      rcCtrlBottom = GetControlRECT( pCtrls(i+1) )

      select case id 
         case IDM_VERTEQUAL, IDM_VERTREMOVE
            nTop = rcCtrlTop.top + (rcCtrlBottom.top - rcCtrlTop.bottom) - nEqualSpace
         case IDM_VERTINCREASE
            nTop = rcCtrlTop.top - 8
         case IDM_VERTDECREASE
            nTop = rcCtrlTop.top + 8
      end select

      SetControlProperty( pCtrl, "TOP", str(nTop) )
      ApplyControlProperties( pDoc, pCtrl )
   next

   for i as long = nActiveControlIndex + 1 to NumSelected
      pCtrl = pCtrls(i)
      if IsControlLocked(pDoc, pCtrl) then continue for

      ' Get coordinates of current (Left) and (Right) controls
      rcCtrlTop    = GetControlRECT( pCtrls(i-1) )
      rcCtrlBottom = GetControlRECT( pCtrls(i) )

      select case id 
         case IDM_VERTEQUAL, IDM_VERTREMOVE
            nTop = rcCtrlBottom.top + nEqualSpace - (rcCtrlBottom.top - rcCtrlTop.bottom)  
         case IDM_VERTINCREASE
            nTop = rcCtrlBottom.top + 8
         case IDM_VERTDECREASE
            nTop = rcCtrlBottom.top - 8
      end select
             
      SetControlProperty( pCtrl, "TOP", str(nTop) )
      ApplyControlProperties( pDoc, pCtrl )
   next
   
   ' Ensure the grab handles are redrawn
   AfxRedrawWindow(pDoc->hWndForm)

   function = 0
end function


' ========================================================================================
' Align controls on the form
' ========================================================================================
function OnCommand_DesignerAlign( _
            byval HWND as HWND, _
            byval id as long _
            ) as LRESULT

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
         if IsControlLocked(pDoc, pCtrl) then continue for
         ' Skip processing the active control because that is unnecessary.
         if pCtrl = pCtrlActive THEN continue for
         rcCtrl = GetControlRECT(pCtrl)
         nWidth  = rcCtrl.Right  - rcCtrl.Left
         nHeight = rcCtrl.Bottom - rcCtrl.Top 
         select CASE id
            case IDM_ALIGNLEFTS
               SetControlProperty(pCtrl, "LEFT", str(rcActive.left))
               ApplyControlProperties(pDoc, pCtrl)
            case IDM_ALIGNTOPS
               SetControlProperty(pCtrl, "TOP", str(rcActive.top))
               ApplyControlProperties(pDoc, pCtrl)
            case IDM_ALIGNRIGHTS
               SetControlProperty(pCtrl, "LEFT", str(rcCtrl.left + (rcActive.right-rcCtrl.right)) )
               ApplyControlProperties(pDoc, pCtrl)
            case IDM_ALIGNBOTTOMS 
               SetControlProperty(pCtrl, "TOP", str(rcCtrl.top + (rcActive.bottom-rcCtrl.bottom)) )
               ApplyControlProperties(pDoc, pCtrl)
            case IDM_ALIGNCENTERS  
               SetControlProperty(pCtrl, "LEFT", _
                  str(rcActive.Left + (nWidthActive / 2) - ( nWidth / 2 )) )
               ApplyControlProperties(pDoc, pCtrl)
            case IDM_ALIGNMIDDLES  
               SetControlProperty(pCtrl, "TOP", _
                  str(rcActive.Top + (nHeightActive / 2) - ( nHeight / 2 )) )
               ApplyControlProperties(pDoc, pCtrl)
            case IDM_SAMEWIDTHS  
               SetControlProperty(pCtrl, "WIDTH", str(nWidthActive))
               ApplyControlProperties(pDoc, pCtrl)
            case IDM_SAMEHEIGHTS
               SetControlProperty(pCtrl, "HEIGHT", str(nHeightActive))
               ApplyControlProperties(pDoc, pCtrl)
            case IDM_SAMEBOTH
               pCtrl->SuspendLayout = true
               SetControlProperty(pCtrl, "WIDTH", str(nWidthActive))
               SetControlProperty(pCtrl, "HEIGHT", str(nHeightActive))
               ApplyControlProperties(pDoc, pCtrl)
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
function OnCommand_DesignerCenter( _
            byval HWND as HWND, _
            byval id as long _
            ) as LRESULT

   Dim pWindow As CWindow Ptr = AfxCWindowPtr(hWnd)
   If pWindow = 0 Then Exit Function

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
   nFormWidth  = (rcForm.right - rcForm.left) / pWindow->rxRatio
   nFormHeight = (rcForm.bottom - rcForm.top) / pWindow->ryRatio
   
   ' Determine the value by which the group of controls needs to be shifted by.  
   nHorizSpacing = nMinLeft - ( (nFormWidth - (nMaxRight - nMinLeft)) / 2) 
   nVertSpacing  = nMinTop - ( (nFormHeight - (nMaxBottom - nMinTop) ) / 2) 

   for i as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
      pCtrl = pDoc->Controls.ItemAt(i)
      if pCtrl->ControlType = CTRL_FORM THEN continue for
      if pCtrl->IsSelected = false THEN continue for
      if IsControlLocked(pDoc, pCtrl) then continue for
      select CASE id
         case IDM_CENTERHORIZ
            nLeft = val(GetControlProperty(pCtrl, "LEFT"))
            SetControlProperty(pCtrl, "LEFT", str(nLeft-nHorizSpacing))
            ApplyControlProperties(pDoc, pCtrl)
         case IDM_CENTERVERT
            nTop  = val(GetControlProperty(pCtrl, "TOP"))
            SetControlProperty(pCtrl, "TOP", str(nTop-nVertSpacing))
            ApplyControlProperties(pDoc, pCtrl)
         case IDM_CENTERBOTH   
            nLeft = val(GetControlProperty(pCtrl, "LEFT"))
            nTop  = val(GetControlProperty(pCtrl, "TOP"))
            pCtrl->SuspendLayout = true
            SetControlProperty(pCtrl, "LEFT", str(nLeft-nHorizSpacing))
            SetControlProperty(pCtrl, "TOP", str(nTop-nVertSpacing))
            ApplyControlProperties(pDoc, pCtrl)
            pCtrl->SuspendLayout = false
      END SELECT
   next

   ' Ensure the grab handles are redrawn
   AfxRedrawWindow(pDoc->hWndForm)

   function = 0
end function


' ========================================================================================
' Delete a control (DEL key was pressed while control(s) selected)
' ========================================================================================
function OnCommand_DesignerDeleteKey( byval pDoc as clsDocument ptr ) as LRESULT
         
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
        
   function = 0
end function


