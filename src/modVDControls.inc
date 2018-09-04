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
' Respond to keyboard input and move selected controls by 1 pixel 
' ========================================================================================
function KeyboardMoveControls( byval pDoc as clsDocument ptr, byval vKeycode as long ) as Long

   dim pCtrl as clsControl ptr
   dim as long nTemp
   
   for i as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
      pCtrl = pDoc->Controls.ItemAt(i)
      if pCtrl->ControlType <> CTRL_FORM THEN
         if pCtrl->IsSelected then
            select CASE vKeycode
               case VK_UP
                  nTemp = val(GetControlProperty(pCtrl, "TOP")) - 1
                  SetControlProperty(pCtrl, "TOP", str(nTemp))
                  pDoc->UserModified = true
               case VK_DOWN
                  nTemp = val(GetControlProperty(pCtrl, "TOP")) + 1
                  SetControlProperty(pCtrl, "TOP", str(nTemp))
                  pDoc->UserModified = true
               case VK_LEFT
                  nTemp = val(GetControlProperty(pCtrl, "LEFT")) - 1
                  SetControlProperty(pCtrl, "LEFT", str(nTemp))
                  pDoc->UserModified = true
               case VK_RIGHT         
                  nTemp = val(GetControlProperty(pCtrl, "LEFT")) + 1
                  SetControlProperty(pCtrl, "LEFT", str(nTemp))
                  pDoc->UserModified = true
            END SELECT
            pCtrl->SuspendLayout = true
            ApplyControlProperties( pDoc, pCtrl )
            pCtrl->SuspendLayout = false
         end if
      end if
   next

   AfxRedrawWindow(pDoc->hWndForm)
   frmMain_UpdateLineCol(HWND_FRMMAIN)
   DisplayPropertyList(pDoc)

   function = 0
end function


' ========================================================================================
' Respond to keyboard input and resize selected controls by 1 pixel 
' ========================================================================================
function KeyboardResizeControls( byval pDoc as clsDocument ptr, byval vKeycode as long ) as Long

   dim pCtrl as clsControl ptr
   dim as long nTemp
   
   for i as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
      pCtrl = pDoc->Controls.ItemAt(i)
      if pCtrl->ControlType <> CTRL_FORM THEN
         if pCtrl->IsSelected then
            select CASE vKeycode
               case VK_UP
                  nTemp = val(GetControlProperty(pCtrl, "HEIGHT")) - 1
                  SetControlProperty(pCtrl, "HEIGHT", str(nTemp))
                  pDoc->UserModified = true
               case VK_DOWN
                  nTemp = val(GetControlProperty(pCtrl, "HEIGHT")) + 1
                  SetControlProperty(pCtrl, "HEIGHT", str(nTemp))
                  pDoc->UserModified = true
               case VK_LEFT
                  nTemp = val(GetControlProperty(pCtrl, "WIDTH")) - 1
                  SetControlProperty(pCtrl, "WIDTH", str(nTemp))
                  pDoc->UserModified = true
               case VK_RIGHT         
                  nTemp = val(GetControlProperty(pCtrl, "WIDTH")) + 1
                  SetControlProperty(pCtrl, "WIDTH", str(nTemp))
                  pDoc->UserModified = true
            END SELECT
            pCtrl->SuspendLayout = true
            ApplyControlProperties( pDoc, pCtrl )
            pCtrl->SuspendLayout = false
         end if
      end if
   next

   AfxRedrawWindow(pDoc->hWndForm)
   frmMain_UpdateLineCol(HWND_FRMMAIN)
   DisplayPropertyList(pDoc)

   function = 0
end function


' ========================================================================================
' Respond to keyboard input to cycle amongst the selected controls in the group.
' ========================================================================================
function KeyboardCycleActiveControls( byval pDoc as clsDocument ptr, byval vKeycode as long ) as Long

   dim pCtrl as clsControl ptr
   dim as long idxActive 
   dim as long idxPrev = -1
   dim as long idxNext = -1
   dim as Boolean bSelectedGroup = iif( pDoc->Controls.SelectedControlsCount > 1, true, false)
   
   ' Get the current active control in the group
   dim pCtrlActive as clsControl ptr = pDoc->Controls.GetActiveControl
   
   ' Determine the array position that the active control is at.
   for i as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
      if pDoc->Controls.ItemAt(i) = pCtrlActive then
         idxActive = i: exit for
      end if
   next      
   
   ' The following will cycle amongst controls in a selected group. If there is
   ' no selected group then the focus will simply shift to the next/prev control.
   select CASE vKeycode
      case VK_UP, VK_LEFT    ' get the previous control in the group
         for i as long = (idxActive - 1) to pDoc->Controls.ItemFirst step -1
            pCtrl = pDoc->Controls.ItemAt(i)
            if pCtrl->ControlType = CTRL_FORM THEN continue for
            if bSelectedGroup then
               if pCtrl->IsSelected then
                  idxPrev = i: exit for
               end if
            else
               idxPrev = i: exit for
            end if         
         next
         if idxPrev = -1 THEN  ' prev still not found. wrap to end of array
            for i as long = pDoc->Controls.ItemLast to idxActive step -1
               pCtrl = pDoc->Controls.ItemAt(i)
               if pCtrl->ControlType = CTRL_FORM THEN continue for
               if bSelectedGroup then
                  if pCtrl->IsSelected then
                     idxPrev = i: exit for
                  end if   
               else
                  idxPrev = i: exit for
               end if   
            next
         end if
         if idxPrev <> -1 then
            if bSelectedGroup = false THEN pDoc->Controls.DeselectAllControls
            pCtrl = pDoc->Controls.ItemAt(idxPrev)
            pDoc->Controls.SetActiveControl(pCtrl->hWindow)
         end if
         
      case VK_DOWN, VK_RIGHT   ' get the next control in the gorup
         for i as long = (idxActive + 1) to pDoc->Controls.ItemLast
            pCtrl = pDoc->Controls.ItemAt(i)
            if pCtrl->ControlType = CTRL_FORM THEN continue for
            if bSelectedGroup then
               if pCtrl->IsSelected then
                  idxNext = i: exit for
               end if   
            else
               idxNext = i: exit for
            end if      
         next
         if idxNext = -1 THEN  ' next still not found. wrap to start of array
            for i as long = pDoc->Controls.ItemFirst to idxActive 
               pCtrl = pDoc->Controls.ItemAt(i)
               if pCtrl->ControlType = CTRL_FORM THEN continue for
               if bSelectedGroup then
                  if pCtrl->IsSelected then
                     idxNext = i: exit for
                  end if   
               else
                  idxNext = i: exit for
               end if   
            next
         end if
         if idxNext <> -1 then
            if bSelectedGroup = false THEN pDoc->Controls.DeselectAllControls
            pCtrl = pDoc->Controls.ItemAt(idxNext)
            pDoc->Controls.SetActiveControl(pCtrl->hWindow)
         END IF

   end select

   AfxRedrawWindow(pDoc->hWndForm)
   frmMain_UpdateLineCol(HWND_FRMMAIN)
   DisplayPropertyList(pDoc)

   function = 0
end function


' ========================================================================================
' Determine if the specified control name already exists on the incoming form.
' ========================================================================================
function IsControlNameExists( byval pDoc as clsDocument ptr, byref wszControlName as wstring ) as boolean

   dim pCtrl as clsControl ptr 
   for ii as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
      pCtrl = pDoc->Controls.ItemAt(ii)
      if pCtrl->ControlType <> CTRL_FORM THEN
         if ucase(GetControlProperty(pCtrl, "NAME")) = ucase(wszControlName) THEN
            return true
         END IF
      END IF
   NEXT
   return false
end function

' ========================================================================================
' Determines if the control is locked or full form is locked
' ========================================================================================
function IsControlLocked( byval pDoc as clsDocument ptr, _
                          byval pCtrl as clsControl ptr _
                          ) as boolean
   if pDoc->bLockControls then return true
   if GetControlProperty(pCtrl, "LOCKED") = "True" then return true
   return false
end function

' ========================================================================================
' Determine if the specified form name already exists.
' ========================================================================================
function IsFormNameExists( byref wszFormName as wstring ) as boolean

   dim pDoc as clsDocument ptr 
   dim pCtrl as clsControl ptr 
   dim as Boolean bExists = false
   dim as long idx = gApp.GetActiveProjectIndex
   
   For i As Long = 0 To gApp.Projects(idx).GetDocumentCount - 1
      pDoc = gApp.Projects(idx).GetDocumentPtr(i)
      if pDoc THEN
         for ii as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
            pCtrl = pDoc->Controls.ItemAt(ii)
            if pCtrl->ControlType = CTRL_FORM THEN
               if ucase(GetControlProperty(pCtrl, "NAME")) = ucase(wszFormName) THEN
                  bExists = true
               END IF
               exit for
            END IF
         NEXT
      END IF
      if bExists THEN exit for
   Next
   
   return bExists
end function

' ========================================================================================
' Generate a default name for the new form
' ========================================================================================
function GenerateDefaultFormName() as CWSTR

   ' Generate a default name for the form, then iterate the document collection
   ' to see if it already exists.
   dim as long NextFormNumber = 0
   dim as CWSTR wszDefaultName 
   
   do 
      NextFormNumber = NextFormNumber + 1
      wszDefaultName = "Form" & NextFormNumber
      if IsFormNameExists(wszDefaultName) = false then exit do
   loop 
   
   return wszDefaultName
end function


' ========================================================================================
' Generate a default name for the new control
' ========================================================================================
function GenerateDefaultControlName( byval pDoc as clsDocument ptr, _
                                     byval nControlType as long _
                                     ) as CWSTR

   ' Generate a default name for the control, then iterate the control collection
   ' to see if it already exists.
   dim as long NextControlNumber = 0
   dim as CWSTR wszDefaultName 
   dim as CWSTR wszName 
   dim as Boolean bOK 
   dim pCtrl as clsControl ptr 
   
   wszName = GetControlName(nControlType)

   do 
      NextControlNumber = NextControlNumber + 1
      wszDefaultName = wszName & NextControlNumber
      bOK = true
      for i as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
         pCtrl = pDoc->Controls.ItemAt(i)
         if ucase(GetControlProperty(pCtrl, "NAME")) = ucase(wszDefaultName) THEN
            bOK = false: exit for
         END IF
      NEXT
   loop until bOK = true
   
   return wszDefaultName
end function


' ========================================================================================
' Generate the next Index in an OptionButton group.
' ========================================================================================
'function GenerateOptionButtonIndex( byval pDoc as clsDocument ptr, _
'                                    byval pCtrl as clsControl ptr _
'                                    ) as CWSTR

'   dim as long NextIndexNumber = 1
'   dim as long nIndex
'   dim as CWSTR wszOptionGroup
'   dim pCtrl2 as clsControl ptr 
    
'   wszOptionGroup = ucase(GetControlProperty(pCtrl, "GROUPNAME"))

'   for i as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
'      pCtrl2 = pDoc->Controls.ItemAt(i)
'      if pCtrl2->ControlType <> CTRL_OPTION then continue for
'      if pCtrl2 = pCtrl then continue for
'      if ucase(GetControlProperty(pCtrl2, "GROUPNAME")) = wszOptionGroup THEN
'         nIndex = val(GetControlProperty(pCtrl2, "INDEX")) 
'         if nIndex >= NextIndexNumber then NextIndexNumber = nIndex + 1
'      end if
'   NEXT
    
'   return wstr(NextIndexNumber)
'end function


' ========================================================================================
' Generate a default TabIndex value for the new control
' ========================================================================================
function GenerateDefaultTabIndex( byval pDoc as clsDocument ptr ) as CWSTR

   dim as long HighTabIndex = -1
   dim as long nPropValue 
   dim pCtrl as clsControl ptr 
   
   for i as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
      pCtrl = pDoc->Controls.ItemAt(i)
      nPropValue = val(GetControlProperty(pCtrl, "TABINDEX"))
      if nPropValue > HighTabIndex then HighTabIndex = nPropValue
   NEXT
   
   return wstr(HighTabIndex+1)
end function


' ========================================================================================
' Create the specified control on the design form
' ========================================================================================
function CreateToolboxControl( byval pDoc as clsDocument ptr, _
                               byval ControlType as long, _
                               byref rcDraw as RECT _
                               ) as clsControl ptr

   Dim pWindow As CWindow Ptr = AfxCWindowPtr(pDoc->hWndForm)
   dim pCtrl as clsControl ptr = new clsControl
   dim as CWSTR wszDefaultControlName, wszDefaultGroupName 

   pDoc->Controls.DeselectAllControls
   pCtrl->ControlType = ControlType
   pCtrl->IsActive    = false
   
   ' Set default properties and events for the created control
   AttachDefaultControlProperties(pCtrl)
   AttachDefaultControlEvents(pCtrl)

   if ControlType = CTRL_FORM THEN
      dim pForm as CWindow ptr = New CWindow
      pForm->DPI = AfxCWindowPtr(pDoc->hWndFrame)->DPI
      wszDefaultControlName = GenerateDefaultFormName()
      pDoc->hWndForm = _
         pForm->Create( pDoc->hWndFrame, "Form1", @DesignerForm_WndProc, 0, 0, 0, 0, _
              WS_CHILD or WS_OVERLAPPED Or WS_OVERLAPPEDWINDOW or _
              WS_CLIPSIBLINGS or WS_CLIPCHILDREN or WS_VISIBLE, _
              WS_EX_CONTROLPARENT OR WS_EX_WINDOWEDGE)
      pForm->ClassStyle = CS_DBLCLKS 
      SetWindowLongPtr( pDoc->hWndForm, GWLP_ID, IDC_DESIGNFORM )
      ' We will set our own mouse pointer as needed
      SetClassLongPtr(pDoc->hWndForm, GCLP_HCURSOR, 0)   
      pCtrl->hWindow = pDoc->hWndForm
   else
      if pWindow = 0 THEN exit function
      wszDefaultControlName = GenerateDefaultControlName(pDoc, ControlType)
      select case ControlType 
         case CTRL_BUTTON
            ' Use the custom control CXpButton instead of the standard Windows button control.
            pCtrl->AfxButtonPtr = new CXpButton(pWindow, 100, wszDefaultControlName, _
                             rcDraw.left, rcDraw.top, rcDraw.right-rcDraw.left, rcDraw.bottom-rcDraw.top)
            pCtrl->hWindow = pCtrl->AfxButtonPtr->hWindow
            SetWindowSubclass(pCtrl->hWindow, CAST(SUBCLASSPROC, @Control_SubclassProc), ControlType, CAST(DWORD_PTR, pDoc))
         case else   
            dim as long dwExStyle = -1
            dim as long dwStyle = -1
            if ControlType = CTRL_LISTBOX then
               ' WinFBX defaults to listbox with integralheight. We want to make sure that
               ' our listboxes have NoIntegralHeight.
               dwStyle = WS_CHILD or WS_VISIBLE OR WS_HSCROLL OR WS_VSCROLL OR WS_BORDER OR _
                         LBS_HASSTRINGS OR LBS_NOTIFY or LBS_NOINTEGRALHEIGHT
            end if
            if ControlType = CTRL_FRAME then
               dwStyle = WS_CHILD or WS_VISIBLE or WS_CLIPSIBLINGS or WS_CLIPCHILDREN or BS_GROUPBOX
               dwExStyle = WS_EX_TRANSPARENT
            end if   
            pCtrl->hWindow = pWindow->AddControl( GetControlClassName(pCtrl), , 100, wszDefaultControlName, _
                             rcDraw.left, rcDraw.top, rcDraw.right-rcDraw.left, rcDraw.bottom-rcDraw.top, _
                             dwStyle, dwExStyle, , _
                             CAST(SUBCLASSPROC, @Control_SubclassProc), ControlType, CAST(DWORD_PTR, pDoc))
      end select
   END IF

   pDoc->Controls.Add(pCtrl)

   SetControlProperty(pCtrl, "NAME", wszDefaultControlName)
   SetControlProperty(pCtrl, "TEXT", wszDefaultControlName)
   SetControlProperty(pCtrl, "LEFT", str(rcDraw.left))
   SetControlProperty(pCtrl, "TOP", str(rcDraw.top))
   SetControlProperty(pCtrl, "WIDTH", str(rcDraw.right-rcDraw.left))
   SetControlProperty(pCtrl, "HEIGHT", str(rcDraw.bottom-rcDraw.top))
   
   ' If this form/control is being created from reading an existing source file then 
   ' then we would wait until the all of the control's properties have been loaded. 
   ' See clsDocument.ParseDesignerString
   if pDoc->LoadingFromFile = false then
      ' By default, generate the next highest TabIndex value for this new control. 
      SetControlProperty(pCtrl, "TABINDEX", GenerateDefaultTabIndex(pDoc))
      'if ControlType = CTRL_OPTION then
      '   SetControlProperty(pCtrl, "INDEX", GenerateOptionButtonIndex(pDoc, pCtrl))
      'end if   
      pCtrl->SuspendLayout = true
      ApplyControlProperties( pDoc, pCtrl )
      pCtrl->SuspendLayout = false
      pDoc->bRegenerateCode = true
      pDoc->Controls.SelectControl(pCtrl->hWindow)
      pDoc->Controls.SetActiveControl(pCtrl->hWindow)
      DisplayPropertyList(pDoc)
   end if
   
   function = pCtrl
END FUNCTION


