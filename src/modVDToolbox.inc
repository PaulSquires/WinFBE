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
' Display the last selected property for the control (or it's default property). This
' allows the user to switch between controls always accessing a common property like 'Text'
' rather than having to seach the PropertyList every time for the property.
' ========================================================================================
Function ShowDefaultPropertyForControl( byval pCtrl as clsControl ptr ) as Long

   dim as hwnd hList 
   dim as long idx
   dim as CWSTR wszDefaultProp, wszDefaultEvent
   
   select case pCtrl->ControlType
      CASE CTRL_FORM, CTRL_BUTTON, CTRL_LABEL, CTRL_CHECKBOX, _
           CTRL_OPTION, CTRL_FRAME
         wszDefaultProp  = "Text" 
         wszDefaultEvent = "Click" 
      case CTRL_TEXTBOX
         wszDefaultProp  = "Text" 
         wszDefaultEvent = "KeyPress" 
      case CTRL_LISTBOX
         wszDefaultProp  = "Name" 
         wszDefaultEvent = "Click" 
   
   ' TODO: Set default property/event for each control.
   '   CTRL_PICTURE
   '   CTRL_COMBOBOX
   '   CTRL_IMAGEBUTTON
   '   CTRL_HSCROLL
   '   CTRL_VSCROLL
   '   CTRL_TIMER
   '   CTRL_TABCONTROL
   '   CTRL_RICHEDIT
   '   CTRL_PROGRESSBAR
   '   CTRL_UPDOWN
   '   CTRL_LISTVIEW
   '   CTRL_TREEVIEW
   '   CTRL_SLIDER
   '   CTRL_DATETIMEPICKER
   '   CTRL_MONTHCALENDAR
   '   CTRL_WEBBROWSER
   '   CTRL_CUSTOM
   '   CTRL_OCX
   END SELECT  
   

   ' Try to match the last accessed property/event name for this control. This allows
   ' you to switch between common properties/events of controls more quickly. If there 
   ' is no match on the last accessed, then use the default value instead.
   wszDefaultProp  = iif(IsPropertyExists(pCtrl, gwszPreviousPropName), gwszPreviousPropName, wszDefaultProp) 
   wszDefaultEvent = iif(IsEventExists(pCtrl, gwszPreviousEventName), gwszPreviousEventName, wszDefaultEvent)

   
   ' Match the default names in the properties and events lists
   hList = GetDlgItem(HWND_FRMVDTOOLBOX, IDC_FRMVDTOOLBOX_LSTPROPERTIES)
   for i as long = 0 to ListBox_GetCount(hList) - 1
      idx = ListBox_GetItemData(hList, i)
      if ucase(pCtrl->Properties(idx).wszPropName) = ucase(wszDefaultProp) then
         ListBox_SetCurSel(hList, i)
         exit for
      end if
   NEXT

   hList = GetDlgItem(HWND_FRMVDTOOLBOX, IDC_FRMVDTOOLBOX_LSTEVENTS)
   for i as long = 0 to ListBox_GetCount(hList) - 1
      idx = ListBox_GetItemData(hList, i)
      if ucase(pCtrl->Events(idx).wszEventName) = ucase(wszDefaultEvent) then
         ListBox_SetCurSel(hList, i)
         exit for
      end if
   NEXT

   ' Finally, show the selected name details visually in the description labels.
   DisplayPropertyDetails
   DisplayEventDetails

   function = 0
end function
   

' ========================================================================================
' Display the properties/events for the current active control/form.
' ========================================================================================
Function DisplayPropertyList( byval pDoc as clsDocument ptr ) as Long
   if pDoc = 0 then exit function

   ' If the ToolBox has not been created yet then exit.
   if IsWindow(HWND_FRMVDTOOLBOX) = 0 then exit function
   
   dim pCtrl as clsControl ptr  
   dim pCtrlActive as clsControl ptr = pDoc->Controls.GetActiveControl
   
   dim as long idx, nCurSel

   ' Always clear the controls combobox and repopuplate it. Do this even
   ' if the active control has not changed. It is possible that that a 
   ' group of controls may have been deleted.
   dim as hwnd hCombo = GetDlgItem(HWND_FRMVDTOOLBOX, IDC_FRMVDTOOLBOX_COMBOCONTROLS)
   Combobox_ResetContent(hCombo)
   ' Iterate through all controls and add them to the combobox
   for i as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
      pCtrl = pDoc->Controls.ItemAt(i)
      if pCtrl then
         idx = Combobox_AddString(hCombo, GetControlProperty(pCtrl, "NAME"))
         Combobox_SetItemData(hCombo, idx, i)   ' store array index in combobox item
      END if
   NEXT
   for i as long = 0 to ComboBox_GetCount(hCombo) - 1
      pCtrl = pDoc->Controls.ItemAt(Combobox_GetItemData(hCombo, i))
      if pCtrl = pCtrlActive then 
         ComboBox_SetCurSel(hCombo, i)
         exit for
      end if
   next

   if pCtrlActive = 0 THEN exit function
   
   ' If the current control property being displayed has not changed then do not
   ' clear the listboxes. It only causes flicker. Instead, simply invalidate them
   ' to cause them to repaint and show most up to date property values (for example,
   ' when controls are being moved or resized).
   static pCtrlDisplayed as clsControl ptr
   if pCtrlDisplayed = pCtrlActive then
      AfxRedrawWindow(GetDlgItem(HWND_FRMVDTOOLBOX, IDC_FRMVDTOOLBOX_LSTPROPERTIES))
      exit function
   END IF
   pCtrlDisplayed = pCtrlActive
   
      
   ' clear the listbox (PROPERTIES)
   dim as hwnd hList = GetDlgItem(HWND_FRMVDTOOLBOX, IDC_FRMVDTOOLBOX_LSTPROPERTIES)
   ListBox_ResetContent(hList)

   dim as CWSTR wszPropName
   
   for i as long = lbound(pCtrlActive->Properties) to ubound(pCtrlActive->Properties)
      ' There are some properties that are only used internally within the designer
      ' and we don't want to display them in the designer itself.
      wszPropName = pCtrlActive->Properties(i).wszPropName
      select case ucase(wszPropName)
         case "NAME" ' ensure it sorts to first position
            wszPropName = "_" & wszPropName
         CASE "STYLE", "EXSTYLE"
            continue for
      END SELECT
      idx = Listbox_AddString(hList, wszPropName)
      ListBox_SetItemData(hList, idx, i)   ' store array index in listbox item
   NEXT

   ' clear the listbox (EVENTS)
   hList = GetDlgItem(HWND_FRMVDTOOLBOX, IDC_FRMVDTOOLBOX_LSTEVENTS)
   ListBox_ResetContent(hList)
   ' Iterate through all events and add them to the listbox
   for i as long = lbound(pCtrlActive->Events) to ubound(pCtrlActive->Events)
      dim as CWSTR wszBlank = ""
      idx = Listbox_AddString(hList, wszBlank)
      ListBox_SetItemData(hList, idx, i)   ' store array index in listbox item
   NEXT

   ' Set the Property and Events listboxes to their default matches for the specified control
   ' or to the most previously accessed property/event should that exist.
   ShowDefaultPropertyForControl(pCtrlActive)

   Function = 0
End Function


' ========================================================================================
' Hide the PropetyList controls for edit, combobox, etc.
' ========================================================================================
function HidePropertyListControls() as long 
   dim as hwnd hList = GetDlgItem(HWND_FRMVDTOOLBOX, IDC_FRMVDTOOLBOX_LSTPROPERTIES)
   ShowWindow(HWND_PROPLIST_EDIT, SW_HIDE)
   ShowWindow(HWND_PROPLIST_COMBO, SW_HIDE)
   ShowWindow(HWND_PROPLIST_COMBOLIST, SW_HIDE)
   AfxRedrawWindow(hList)
   function = 0
END FUNCTION


' ========================================================================================
' Property value was manually changed... initiate the value change.
' ========================================================================================
function InitiatePropertyValueChange() as Long

   dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr
   dim pCtrl as clsControl ptr = pDoc->Controls.GetActiveControl
   dim pProp as clsProperty ptr = GetActivePropertyPtr()
   dim pCtrlForm as clsControl ptr = GetFormCtrlPtr(pDoc)
   dim as CWSTR wszValue, wszOldValue
   dim txtFind(3) as string
   dim txtReplace(3) as string
   dim as long startPos, endPos, r
   
   if pProp THEN 
      wszOldValue = pProp->wszPropValue
      select case pProp->PropType
         CASE PropertyType.EditEnter, PropertyType.EditEnterNumericOnly
            if IsWindowVisible(HWND_PROPLIST_EDIT) = false then exit function
            wszValue = AfxGetWindowText(HWND_PROPLIST_EDIT)
         case PropertyType.ComboPicker, PropertyType.TrueFalse
            wszValue = AfxGetListBoxText(HWND_PROPLIST_COMBOLIST, ListBox_GetCurSel(HWND_PROPLIST_COMBOLIST) )
         case PropertyType.FontPicker
            ' Font changes are never handled here. They are dealt with when the popup Font dialog
            ' closes. Exit out of this function now otherwise the property value will be deleted.
            exit function
         case PropertyType.ColorPicker
            ' Color changes are handled when the popup color selector loses focus.
            exit function
      END SELECT
  
      if wszValue <> wszOldValue then
         select case ucase(pProp->wszPropName)
            CASE "NAME"
               wszValue = trim(wszValue)
               ' The name can only be alphanumeric
               static wszRetain as CWSTR = "abcdefghijklmnopqrstuvwxyz0123456789"
               wszValue = AfxStrRetainAnyI(wszValue, wszRetain)
               ' Do a check to ensure that the name does not already exist.
               if pCtrl->ControlType = CTRL_FORM then               
                  if IsFormNameExists(wszValue) then
                     ' Don't put a messagebox here because it will cause the propertylist to
                     ' lose focus and therefore trigger the propertyvalue change.
                     wszValue = wszOldValue
                  else
                     txtFind(1) = wszOldValue    ' "Form1"  ' must use SCFIND_WHOLEWORD for this one
                     txtFind(2) = wszOldValue & "."
                     txtFind(3) = " " & wszOldValue & "_"
                     txtReplace(1) = wszValue    ' "Form2"
                     txtReplace(2) = wszValue & "."
                     txtReplace(3) = " " & wszValue & "_"
                  end if
               else
                  if IsControlNameExists(pDoc, wszValue) then
                     ' Don't put a messagebox here because it will cause the propertylist to
                     ' lose focus and therefore trigger the propertyvalue change.
                     wszValue = wszOldValue
                  else
                     dim as CWSTR wszFormName 
                     if pCtrlForm then wszFormName = GetControlProperty(pCtrlForm, "NAME")
                     txtFind(1) = ""
                     txtFind(2) = wszFormName & "." & wszOldValue & "."
                     txtFind(3) = " " & wszFormName & "_" & wszOldValue & "_"
                     txtReplace(1) = ""
                     txtReplace(2) = wszFormName & "." & wszValue & "."
                     txtReplace(3) = " " & wszFormName & "_" & wszValue & "_"
                  end if
               END IF
               
               ' If the form or control name has changed then update all source code
               dim pDocParse as clsDocument ptr
               dim as long idx = gApp.GetActiveProjectIndex
               dim as hwnd hEdit 
               dim as Boolean bReparse 

               for i as long = 1 to 3
                  if len(txtReplace(i)) = 0 then continue for
                  
                  For ii As Long = 0 To gApp.Projects(idx).GetDocumentCount - 1
                     pDocParse = gApp.Projects(idx).GetDocumentPtr(ii)
                     if pDocParse THEN
                        bReparse = false
                        hEdit = pDocParse->hWndActiveScintilla
                        if i = 1 then 
                           SendMessage( hEdit, SCI_SETSEARCHFLAGS, SCFIND_WHOLEWORD, 0)
                        else   
                           SendMessage( hEdit, SCI_SETSEARCHFLAGS, 0, 0)
                        end if
                        SciExec( hEdit, SCI_TARGETWHOLEDOCUMENT, 0, 0)
                        startPos = SciExec( hEdit, SCI_GETTARGETSTART, 0, 0)
                        endPos = SciExec( hEdit, SCI_GETTARGETEND, 0, 0)
                        do 
                           r = SciExec( hEdit, SCI_SEARCHINTARGET, Len(txtFind(i)), Strptr(txtFind(i)))
                           if r = -1 THEN exit do
                           bReparse = true
                           SciExec( hEdit, SCI_REPLACETARGET, len(txtReplace(i)), Strptr(txtReplace(i)) )
                           startPos = r + len(txtFind(i))
                           ' Adjust the searching positions
                           SciExec( hEdit, SCI_SETTARGETSTART, startPos, 0)
                           SciExec( hEdit, SCI_SETTARGETEND, endPos, 0)
                        loop
                     end if
                     if bReparse then ParseDocument(idx, pDocParse, "")
                  next
               next   
         end select
         
' TODO: Corruption in output string occurs if I simply try to assign CWSTR as below.
         ' Need to assign to intermediary wstring first for some reason.
         'pProp->wszPropValue = wszValue       
         '***** WORKAROUND *****           
         dim wst as wstring * 1000 = **wszValue
         pProp->wszPropValue = wst
         '**********************           

         select case ucase(pProp->wszPropName)
            CASE "NAME"
               ' If this was a NAME property change then we need to potentially update the form ACCEPTBUTTON 
               ' or CANCELBUTTON properties that depend on this button name.
               ' Ensure that the controls combox is repainted to show any changes
               AfxRedrawWindow(GetDlgItem(HWND_FRMVDTOOLBOX, IDC_FRMVDTOOLBOX_COMBOCONTROLS ))
               pCtrlForm = GetFormCtrlPtr(pDoc)
               if pCtrlForm then
                  If GetControlProperty(pCtrlForm, "ACCEPTBUTTON") = wszOldValue then
                     SetControlProperty(pCtrlForm, "ACCEPTBUTTON", wszValue)
                  end if    
                  If GetControlProperty(pCtrlForm, "CANCELBUTTON") = wszOldValue then
                     SetControlProperty(pCtrlForm, "CANCELBUTTON", wszValue)
                  end if    
               end if
            case "ACCEPTBUTTON", "CANCELBUTTON"
               ' If (none) was selected then ensure that the property value is set to blank
               if pProp->wszPropValue = "(none)" then pProp->wszPropValue = ""
            case "TABINDEX"
               ' Now that a new TabIndex has been entered, ensure that there are no duplicates 
               ' amongst all of the controls. Bypass the active control but increase the TabIndex
               ' value of all controls with a higher TabIndex number should a duplicate exist.
               dim as long nTabIndex = val(wst)
               dim as long nCurTabIndex
               dim pCtrlLoop as clsControl ptr
               dim pCtrlFound as clsControl ptr = 0
               
               ' Text to see if TabIndex value already exists
               for i as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
                  pCtrlLoop = pDoc->Controls.ItemAt(i)
                  if pCtrl <> pCtrlLoop then
                     nCurTabIndex = val(GetControlProperty(pCtrlLoop, "TABINDEX"))
                     if nCurTabIndex = nTabIndex then
                        pCtrlFound = pCtrlLoop
                        exit for
                     end if   
                  END if
               NEXT
               ' If exists then increase TabIndex for 
               if pCtrlFound then
                  for i as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
                     pCtrlLoop = pDoc->Controls.ItemAt(i)
                     if pCtrl <> pCtrlLoop then
                        nCurTabIndex = val(GetControlProperty(pCtrlLoop, "TABINDEX"))
                        if nCurTabIndex >= nTabIndex then
                           SetControlProperty(pCtrlLoop, "TABINDEX", wstr(nCurTabIndex + 1))
                        end if   
                     END if
                  NEXT
               end if   
               
         end select

         ApplyControlProperties(pDoc, pCtrl)
         ' Indicate that the file is now dirty and will need to be saved
         pDoc->UserModified = true
         pDoc->bRegenerateCode = true
         frmMain_UpdateLineCol(HWND_FRMMAIN)
      end if
   END IF

   function = 0
END FUNCTION


' ========================================================================================
' Show the propertylist control related to the current selected propertylist row.
' ========================================================================================
function ShowPropertyListControl() as Long
   
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(HWND_FRMVDTOOLBOX)
   if pWindow = 0 THEN exit function

   dim pCtrl as clsControl ptr 
   dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr
   if pDoc then pCtrl = pDoc->Controls.GetActiveControl
   if pCtrl = 0 THEN exit function

   ' Get the rectangle related to the textbox edit area and then display the edit control
   dim as hwnd hList = GetDlgItem(HWND_FRMVDTOOLBOX, IDC_FRMVDTOOLBOX_LSTPROPERTIES)
   dim as long nCurSel = ListBox_GetCurSel(hList)
   dim as long nItemData = ListBox_GetItemData(hList, nCurSel)
   dim as RECT rc, rcLine, rcList
   dim as CWSTR wszPropValue
   
   if nCurSel = -1 then exit function
   
   ' ItemRect coordinates are client relative to the Listbox. 
   ListBox_GetItemRect(hList, nCurSel, @rcLine)
   rc = rcLine
   GetClientRect(hList, @rcList)
   
   select case pCtrl->Properties(nItemData).PropType
      CASE PropertyType.EditEnter, PropertyType.EditEnterNumericOnly
         rc.left = rc.left + gPropDivPos
         rc.top = rc.top + pWindow->ScaleY(2)
         rc.bottom = rc.bottom - pWindow->ScaleY(2)
         rc.left = rc.left + pWindow->ScaleX(1)
         SetWindowPos(HWND_PROPLIST_EDIT, HWND_TOP, rc.left, rc.top, _
                             rc.right-rc.left, rc.bottom-rc.top, SWP_SHOWWINDOW)
         AfxSetWindowText(HWND_PROPLIST_EDIT, pCtrl->Properties(nItemData).wszPropValue)
         ' If the user clicked in the right hand side of the propertylist line then also
         ' set the focus immediately to the edit box. Saves the user from having to 
         ' click there a second time.
         rcLine.left = rcLine.left + gPropDivPos
         dim as POINT pt: GetCursorPos(@pt)
         MapWindowPoints(0, hList, cast(POINT ptr, @pt), 1)                    
         If PtInRect(@rcLine, pt) then
            dim as long nPos = len(pCtrl->Properties(nItemData).wszPropValue)
            SendMessage(HWND_PROPLIST_EDIT, EM_SETSEL, nPos, nPos)
            SetFocus(HWND_PROPLIST_EDIT)
         END IF
         
      CASE PropertyType.ComboPicker, PropertyType.TrueFalse
         rc.left = rc.right - (rc.bottom-rc.top)
         SetWindowPos(HWND_PROPLIST_COMBO, HWND_TOP, rc.left, rc.top, _
                             rc.right-rc.left, rc.bottom-rc.top, SWP_SHOWWINDOW)
         ' Fill the listbox with the correct options
         LoadPropertyComboListbox(pDoc, pCtrl, @pCtrl->Properties(nItemData))
         ' Set the active line based on the current stored value for the property
         dim as long nCurSel = ListBox_FindStringExact(HWND_PROPLIST_COMBOLIST, -1, _
                                                       pCtrl->Properties(nItemData).wszPropValue)
         if nCurSel = -1 then
            pCtrl->Properties(nItemData).wszPropValue = pCtrl->Properties(nItemData).wszPropDefault
            nCurSel = ListBox_FindStringExact( HWND_PROPLIST_COMBOLIST, -1, _
                                               pCtrl->Properties(nItemData).wszPropValue)
         END IF
         ListBox_SetCurSel(HWND_PROPLIST_COMBOLIST, nCurSel)
         ' Calculate height of the Listbox
         rcLine.left = rcLine.left + gPropDivPos
         dim as long numItems = ListBox_GetCount(HWND_PROPLIST_COMBOLIST)
         dim as long nLineHeight = SendMessage(HWND_PROPLIST_COMBOLIST, LB_GETITEMHEIGHT, 0, 0)
         dim as long nListHeight = (numItems * nLineHeight) + 4 

         if rcLine.bottom + nListHeight > rcList.bottom then
            SetWindowPos(HWND_PROPLIST_COMBOLIST, HWND_TOP, rcLine.left, rcLine.top - nListHeight, _
                                rcLine.right-rcLine.left, nListHeight, 0)
         else   
            SetWindowPos(HWND_PROPLIST_COMBOLIST, HWND_TOP, rcLine.left, rcLine.bottom, _
                                rcLine.right-rcLine.left, nListHeight, 0)
         end if
         ' If the user clicked in the right hand side of the propertylist line then also
         ' display the dropdown listbox rather than having the user click on the combo box
         ' icon as well.
         dim as POINT pt: GetCursorPos(@pt)
         MapWindowPoints(0, hList, cast(POINT ptr, @pt), 1)                    
         If PtInRect(@rcLine, pt) then ShowWindow(HWND_PROPLIST_COMBOLIST, SW_SHOW)

      CASE PropertyType.FontPicker
         rc.left = rc.right - (rc.bottom-rc.top)
         SetWindowPos(HWND_PROPLIST_COMBO, HWND_TOP, rc.left, rc.top, _
                             rc.right-rc.left, rc.bottom-rc.top, SWP_SHOWWINDOW)

      CASE PropertyType.ColorPicker
         rc.left = rc.right - (rc.bottom-rc.top)
         SetWindowPos(HWND_PROPLIST_COMBO, HWND_TOP, rc.left, rc.top, _
                             rc.right-rc.left, rc.bottom-rc.top, SWP_SHOWWINDOW)
         ' If the user clicked in the right hand side of the propertylist line then also
         ' display the dropdown listbox rather than having the user click on the combo box
         ' icon as well.
         rcLine.left = rcLine.left + gPropDivPos
         dim as POINT pt: GetCursorPos(@pt)
         MapWindowPoints(0, hList, cast(POINT ptr, @pt), 1)                    
         If PtInRect(@rcLine, pt) then 
            MapWindowPoints(hList, 0, cast(POINT ptr, @rc), 2)
            frmVDColors_Show(hList, pCtrl->Properties(nItemData).wszPropValue)  ' initialize the color popup if not already done so
            dim as long nWidth = AfxGetWindowWidth(HWND_FRMVDCOLORS)
            SetWindowPos(HWND_FRMVDCOLORS, HWND_TOP, rc.right - nWidth, rc.bottom, 0, 0, SWP_NOSIZE or SWP_SHOWWINDOW)
         end if
   END SELECT

   function = 0
END FUNCTION


' ========================================================================================
' Process WM_COMMAND message for window/dialog: frmVDToolbox
' ========================================================================================
Function frmVDToolbox_OnCommand(ByVal HWnd As HWnd, ByVal id As Long, ByVal hwndCtl As HWnd, ByVal codeNotify As UINT) As LRESULT

   select case id
      CASE IDC_FRMVDTOOLBOX_COMBOCONTROLS
         if codeNotify = CBN_SELCHANGE then
            dim as long nCurSel = Combobox_GetCurSel(hwndCtl)
            if nCurSel = -1 then exit function
            dim as long idx = Combobox_GetItemData(hwndCtl, nCurSel)
            dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr
            dim pCtrl as clsControl ptr
            pCtrl = pDoc->Controls.ItemAt(idx)
            if pCtrl then 
               pDoc->Controls.DeselectAllControls
               pDoc->Controls.SetActiveControl(pCtrl->hWindow)
               frmMain_UpdateLineCol(HWND_FRMMAIN)
               AfxRedrawWindow(pDoc->hWndFrame)
               AfxRedrawWindow(pDoc->hWndForm)
               DisplayPropertyList(pDoc)
            end if
         END IF

      case IDC_FRMVDTOOLBOX_LSTPROPERTIES 
         if codeNotify = LBN_SELCHANGE then
            HidePropertyListControls
            ShowPropertyListControl
            DisplayPropertyDetails
         END IF

      case IDC_FRMVDTOOLBOX_LSTEVENTS
         if codeNotify = LBN_SELCHANGE then
            DisplayEventDetails
         END IF

   END SELECT

   Function = 0
End Function

' ========================================================================================
' Process WM_NOTIFY message for window/dialog: frmVDToolbox
' ========================================================================================
Function frmVDToolbox_OnNotify(ByVal HWnd As HWnd, ByVal id As Long, ByVal pNMHDR As NMHDR Ptr) As LRESULT

   SELECT CASE id
      CASE IDC_FRMVDTOOLBOX_TABCONTROL
         dim as long iPage = TabCtrl_GetCurSel(pNMHDR->hwndFrom)
         SELECT CASE pNMHDR->code
            CASE TCN_SELCHANGE
               ' Show the selected page controls
               if iPage = 0 then 
                  ShowWindow( GetDlgItem(HWND, IDC_FRMVDTOOLBOX_COMBOCONTROLS), SW_HIDE)
                  ShowWindow( GetDlgItem(HWND, IDC_FRMVDTOOLBOX_LSTTOOLBOX), SW_SHOW)
               end if
               if iPage = 1 then 
                  ShowWindow( GetDlgItem(HWND, IDC_FRMVDTOOLBOX_COMBOCONTROLS), SW_SHOW)
                  ShowWindow( GetDlgItem(HWND, IDC_FRMVDTOOLBOX_LSTPROPERTIES), SW_SHOW)
                  ShowWindow( GetDlgItem(HWnd, IDC_FRMVDTOOLBOX_LBLPROPNAME), SW_SHOW)
                  ShowWindow( GetDlgItem(HWnd, IDC_FRMVDTOOLBOX_LBLPROPDESCRIBE), SW_SHOW)
               END IF
               if iPage = 2 then 
                  ShowWindow( GetDlgItem(HWND, IDC_FRMVDTOOLBOX_COMBOCONTROLS), SW_SHOW)
                  ShowWindow( GetDlgItem(HWND, IDC_FRMVDTOOLBOX_LSTEVENTS), SW_SHOW)
                  ShowWindow( GetDlgItem(HWnd, IDC_FRMVDTOOLBOX_LBLPROPNAME), SW_SHOW)
                  ShowWindow( GetDlgItem(HWnd, IDC_FRMVDTOOLBOX_LBLPROPDESCRIBE), SW_SHOW)
               END IF
               ' Ensure that we reset the ToolBox Tool to the Pointer
               SetActiveToolboxControl(CTRL_POINTER)
               DisplayPropertyDetails
               DisplayEventDetails
                                
            CASE TCN_SELCHANGING
               ' Hide the current page
               ShowWindow( GetDlgItem(HWND, IDC_FRMVDTOOLBOX_COMBOCONTROLS), SW_HIDE)
               if iPage = 0 then ShowWindow( GetDlgItem(HWND, IDC_FRMVDTOOLBOX_LSTTOOLBOX), SW_HIDE)
               if (iPage = 1) or (iPage = 2) then
                  ShowWindow( GetDlgItem(HWnd, IDC_FRMVDTOOLBOX_LBLPROPNAME), SW_HIDE)
                  ShowWindow( GetDlgItem(HWnd, IDC_FRMVDTOOLBOX_LBLPROPDESCRIBE), SW_HIDE)
               end if   
               if iPage = 1 then ShowWindow( GetDlgItem(HWND, IDC_FRMVDTOOLBOX_LSTPROPERTIES), SW_HIDE)
               if iPage = 2 then ShowWindow( GetDlgItem(HWND, IDC_FRMVDTOOLBOX_LSTEVENTS), SW_HIDE)
         END SELECT

   END SELECT

   function = 0
end function

' ========================================================================================
' Position all child windows. Called manually and/or by WM_SIZE
' ========================================================================================
Function frmVDToolbox_PositionWindows( ByVal HWnd As HWnd ) As LRESULT
   
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(HWnd)
   if pWindow = 0 THEN exit function
      
   ' Get the entire client area
   Dim As Rect rc
   GetClientRect(HWnd, @rc)
   
   Dim As HWnd hList1  = GetDlgItem(HWnd, IDC_FRMVDTOOLBOX_LSTTOOLBOX )
   Dim As HWnd hList2  = GetDlgItem(HWnd, IDC_FRMVDTOOLBOX_LSTPROPERTIES )
   Dim As HWnd hList3  = GetDlgItem(HWnd, IDC_FRMVDTOOLBOX_LSTEVENTS )
   Dim As HWnd hTabCtl = GetDlgItem(HWnd, IDC_FRMVDTOOLBOX_TABCONTROL )
   Dim As HWnd hCombo  = GetDlgItem(HWnd, IDC_FRMVDTOOLBOX_COMBOCONTROLS )
   Dim As HWnd hPropName = GetDlgItem(HWnd, IDC_FRMVDTOOLBOX_LBLPROPNAME)
   Dim As HWnd hPropDescribe = GetDlgItem(HWnd, IDC_FRMVDTOOLBOX_LBLPROPDESCRIBE )
   
   SetWindowPos( hTabCtl, 0, 0, 0, rc.Right-rc.Left, pWindow->ScaleY(24), SWP_SHOWWINDOW Or SWP_NOZORDER )

   ' TAB 1: ToolBox
   SetWindowPos( hList1, 0, 0, pWindow->ScaleY(24), rc.Right-rc.Left, rc.Bottom-rc.top-pWindow->ScaleY(24), SWP_NOZORDER )
   
   ' TAB 2 & 3: Properties & Events 
   '   Combolist of controls
   SetWindowPos( hCombo, 0, 0, pWindow->ScaleY(24), rc.Right-rc.Left, pWindow->ScaleY(20), SWP_NOZORDER )
   '   Bold label for Property/Event name
   SetWindowPos( hPropName, 0, pWindow->ScaleX(4), rc.Bottom - pWindow->ScaleY(90), _
                               rc.Right-rc.Left-pWindow->ScaleX(8), pWindow->ScaleY(20), SWP_NOZORDER )
   '   Multiline label for Property/Event description
   SetWindowPos( hPropDescribe, 0, pWindow->ScaleX(4), rc.Bottom - pWindow->ScaleY(70), _
                               rc.Right-rc.Left-pWindow->ScaleX(8), pWindow->ScaleY(70), SWP_NOZORDER )
   
   ' TAB 2: Properties 
   SetWindowPos( hList2, 0, 0, pWindow->ScaleY(48), rc.Right-rc.Left, rc.Bottom-rc.top-pWindow->ScaleY(138), SWP_NOZORDER )
   
   ' TAB 3: Events
   SetWindowPos( hList3, 0, 0, pWindow->ScaleY(48), rc.Right-rc.Left, rc.Bottom-rc.top-pWindow->ScaleY(138), SWP_NOZORDER )
   
   ShowPropertyListControl
   
   Function = 0
End Function


' ========================================================================================
' Process WM_SIZE message for window/dialog: frmVDToolbox
' ========================================================================================
Function frmVDToolbox_OnSize(ByVal HWnd As HWnd, ByVal state As UINT, ByVal cx As Long, ByVal cy As Long) As LRESULT
   If state <> SIZE_MINIMIZED Then 
      frmVDToolbox_PositionWindows(HWnd)
   End If   
   Function = 0
End Function


' ========================================================================================
' Process WM_CLOSE message for window/dialog: frmVDToolbox
' ========================================================================================
Function frmVDToolbox_OnClose( ByVal HWnd As HWnd ) As LRESULT
   ' Never close the window; simply hide it.
   ShowWindow( HWnd, SW_HIDE )
   Function = 0
End Function


' ========================================================================================
' Process WM_MEASUREITEM message for window/dialog: frmVDToolbox
' ========================================================================================
Function frmVDToolbox_OnMeasureItem( ByVal HWnd As HWnd, ByVal lpmis As MEASUREITEMSTRUCT Ptr ) As Long
   ' Set the height of the List box items. 
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(HWnd)
   lpmis->itemHeight = pWindow->ScaleY(LISTBOX_LINEHEIGHT)
   Function = 0
End Function


' ========================================================================================
' Process WM_DRAWITEM message for window/dialog: frmVDToolbox
' ========================================================================================
Function frmVDToolbox_OnDrawItem( ByVal HWnd As HWnd, ByVal lpdis As Const DRAWITEMSTRUCT Ptr ) As Long

   Dim As HBRUSH hBrush 
   dim as HANDLE hImage
   dim as HFONT hFont 
   Dim As RECT rc, rc2 
   Dim wzText As WString * MAX_PATH
   
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(HWnd)
   if pWindow = 0 THEN exit function
      
   If lpdis->itemID = -1 Then Exit Function
   
   Select Case lpdis->itemAction
      Case ODA_DRAWENTIRE, ODA_SELECT
         
         SaveDC(lpdis->hDC)
         
         if lpdis->CtlID = IDC_FRMVDTOOLBOX_COMBOCONTROLS then
            ' CLEAR BACKGROUND
            If (lpdis->itemState And ODS_SELECTED) Then     
               SetBkColor(lpdis->hDC, GetSysColor(COLOR_HIGHLIGHT))   
               SetTextColor(lpdis->hDC, GetSysColor(COLOR_HIGHLIGHTTEXT))
               hBrush = GetSysColorBrush(COLOR_HIGHLIGHT) 
            else
               SetBkColor(lpdis->hDC, GetSysColor(COLOR_WINDOW))   
               SetTextColor(lpdis->hDC, GetSysColor(COLOR_WINDOWTEXT))
               hBrush = GetSysColorBrush(COLOR_WINDOW) 
            end if
            SelectObject(lpdis->hDC, hBrush)      
            FillRect(lpdis->hDC, @lpdis->rcItem, hBrush)  

            dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr
            
            if (pDoc <> 0) andalso (lpdis->itemData < pDoc->Controls.Count) then
               dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr
               dim pCtrl as clsControl ptr
               pCtrl = pDoc->Controls.ItemAt(lpdis->itemData)
               if pCtrl then
                  SelectObject(lpdis->hDC, ghBoldFont)
                  rc = lpdis->rcItem
                  dim as CWSTR wszCtrlName = GetControlProperty(pCtrl, "NAME")
                  ' Get the size of this string because the Control type string must be drawn after this.
                  Dim nSize As SIZE
                  GetTextExtentPoint32( lpdis->hDC, wszCtrlName, Len(wszCtrlName), @nSIZE)
                  rc.left = rc.left + pWindow->ScaleX(4)
                  DrawText( lpdis->hDC, _
                            wszCtrlName, _
                            -1, Cast(lpRect, @rc), _
                            DT_LEFT Or DT_SINGLELINE Or DT_VCENTER or DT_NOPREFIX)
                
                  ' Determine the Control Type
                  SelectObject(lpdis->hDC, ghNormalFont)
                  wszCtrlName = GetToolBoxName(pCtrl->ControlType) & " (" & GetWinformsXClassName(pCtrl->ControlType) & ")"
                  ' Text is drawn immediately after the Control Name
                  rc.left = rc.left + nSize.cx + pWindow->ScaleX(4)
                  DrawText( lpdis->hDC, _
                            wszCtrlName, _
                            -1, Cast(lpRect, @rc), _
                            DT_LEFT Or DT_SINGLELINE Or DT_VCENTER or DT_NOPREFIX)
               end if
            end if
         end if

         
         if lpdis->CtlID = IDC_FRMVDTOOLBOX_LSTTOOLBOX then
            ' CLEAR BACKGROUND
            If (lpdis->itemState And ODS_SELECTED) Then     
               hBrush = GetSysColorBrush(COLOR_HIGHLIGHT) 
               SetBkColor(lpdis->hDC, GetSysColor(COLOR_HIGHLIGHT))   
               SetTextColor(lpdis->hDC, GetSysColor(COLOR_HIGHLIGHTTEXT))
            else
               hBrush = GetSysColorBrush(COLOR_WINDOW) 
               SetBkColor(lpdis->hDC, GetSysColor(COLOR_WINDOW))   
               SetTextColor(lpdis->hDC, GetSysColor(COLOR_WINDOWTEXT))
            end if
            SelectObject(lpdis->hDC, hBrush)      
            FillRect(lpdis->hDC, @lpdis->rcItem, hBrush)  
            ' The index into the gToolBox array is stored in the itemData of the line.
            
            ' DRAW IMAGE
            dim as HDC hdcMem = CreateCompatibleDC(lpdis->hDC)
            hImage = LoadImage(pWindow->InstanceHandle, gToolBox(lpdis->itemData).wszImage, IMAGE_BITMAP, _
                               pWindow->ScaleX(16), pWindow->ScaleX(16), LR_LOADTRANSPARENT)
            SelectObject(hdcMem, hImage)
            BitBlt( lpdis->hDC, _
                    lpdis->rcItem.left + pWindow->ScaleX(6), _
                    lpdis->rcItem.top + pWindow->ScaleY(2), _
                    pWindow->ScaleX(16), pWindow->ScaleY(16), _
                    hdcMem, 0, 0, SRCCOPY)
            DeleteDC(hdcMem)                   

            ' DRAW TEXT
            rc = lpdis->rcItem: rc.left = rc.left + pWindow->ScaleX(30)
            if (lpdis->itemData >= lbound(gToolBox)) andalso (lpdis->itemData <= ubound(gToolBox)) then
               DrawText( lpdis->hDC, gToolBox(lpdis->itemData).wszToolBoxName, -1, Cast(lpRect, @rc), _
                         DT_LEFT Or DT_SINGLELINE Or DT_VCENTER )
            end if          
         end if

         if lpdis->CtlID = IDC_FRMVDTOOLBOX_LSTPROPERTIES then
            dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr
            dim pCtrl as clsControl ptr 
            if pDoc then
               pCtrl = pDoc->Controls.GetActiveControl
               if pCtrl THEN
                  rc  = lpdis->rcItem
                  rc2 = lpdis->rcItem
                  rc.right = rc.left + gPropDivPos
                  rc2.left = rc.right
                  
                  ' CLEAR BACKGROUND
                  If (lpdis->itemState And ODS_SELECTED) Then     
                     SetBkColor(lpdis->hDC, GetSysColor(COLOR_HIGHLIGHT))   
                     SetTextColor(lpdis->hDC, GetSysColor(COLOR_HIGHLIGHTTEXT))
                     hBrush = GetSysColorBrush(COLOR_HIGHLIGHT) 
                  else
                     SetBkColor(lpdis->hDC, GetSysColor(COLOR_WINDOW))   
                     SetTextColor(lpdis->hDC, GetSysColor(COLOR_WINDOWTEXT))
                     hBrush = GetSysColorBrush(COLOR_WINDOW) 
                  end if

                  ' Draw the property name 
                  SelectObject(lpdis->hDC, hBrush)      
                  FillRect(lpdis->hDC, @rc, hBrush)  
                  rc.left = rc.left + pWindow->ScaleX(4)
                  DrawText( lpdis->hDC, _
                            pCtrl->Properties(lpdis->itemData).wszPropName, _
                            -1, Cast(lpRect, @rc), _
                            DT_LEFT Or DT_SINGLELINE Or DT_VCENTER )

                  ' Draw the current value
                  hBrush = GetSysColorBrush(COLOR_WINDOW) 
                  SelectObject(lpdis->hDC, hBrush)      
                  FillRect(lpdis->hDC, @rc2, hBrush)  
                  SetBkColor(lpdis->hDC, GetSysColor(COLOR_WINDOW))   
                  SetTextColor(lpdis->hDC, GetSysColor(COLOR_WINDOWTEXT))
                  rc2.left = rc2.left + pWindow->ScaleX(4)
                  
                  ' Some properties that are blank will show "(none)"
                  dim as CWSTR wszPropValue = pCtrl->Properties(lpdis->itemData).wszPropValue
                  if len(wszPropValue) = 0 then
                     select case ucase(pCtrl->Properties(lpdis->itemData).wszPropName)
                        CASE "ACCEPTBUTTON", "BACKGROUNDIMAGE", "CANCELBUTTON", "ICON", "IMAGE"
                           wszPropValue = "(none)"
                     END SELECT
                  end if   
                  if ucase(pCtrl->Properties(lpdis->itemData).wszPropName) = "FONT" then
                     dim as LOGFONT lf
                     GetObject(ghNormalFont, sizeof(lf), @lf)
                     dim as CWSTR wszFaceName = lf.lfFaceName
                     dim as long nHeight = lf.lfHeight
                     dim as long nCharSet = lf.lfCharSet
                     lf = SetLogFontFromPropValue(wszPropValue)  ' create font with bold, underline, strikeout
                     lf.lfFaceName = wszFaceName
                     lf.lfHeight   = nHeight
                     lf.lfCharSet  = nCharSet
                     hFont = CreateFontIndirect(@lf)
                     SelectObject(lpdis->hDC, hFont)
                     dim as CWSTR wszFont = AfxStrParse(wszPropValue, 1, ",") & ", " & _ 
                                            AfxStrParse(wszPropValue, 2, ",") & "pt"
                     wszPropValue = wszFont
                  end if
                  if pCtrl->Properties(lpdis->itemData).PropType = PropertyType.ColorPicker then
                     dim as CWSTR wszList, wszValue
                     wszList = AfxStrParse(wszPropValue, 1, "|")
                     wszValue = AfxStrParse(wszPropValue, 2, "|")
                     wszPropValue = wszValue
                     if wszList = "CUSTOM" then wszPropValue = "Custom Color"
                     
                     dim as HBRUSH hBrushColor, hBrushOld
                     dim as RECT rcColor = rc2
                     rcColor.left   = rcColor.left 
                     rcColor.right  = rcColor.left + pWindow->ScaleX(18)
                     rcColor.top    = rcColor.top + pWindow->ScaleY(1)
                     rcColor.bottom = rcColor.bottom - pWindow->ScaleY(4)
                     
                     dim as COLORREF rgbClr = GetRGBColorFromProperty(@pCtrl->Properties(lpdis->itemData))
                     hBrushColor = CreateSolidBrush(rgbClr)

                     ' DRAW COLOR RECT
                     hBrushOld = SelectObject(lpdis->hDC, hBrushColor)
                     Rectangle( lpdis->hDC, rcColor.Left, rcColor.Top, rcColor.Right, rcColor.Bottom)
                     SelectObject( lpdis->hDC, hBrushOld)
                     rc2.left = rcColor.right + pWindow->ScaleX(6)
                     DeleteObject(hBrushColor)
                  end if   

                  DrawText( lpdis->hDC, _
                            wszPropValue, _
                            -1, Cast(lpRect, @rc2), _
                            DT_LEFT Or DT_SINGLELINE Or DT_VCENTER )
                  SelectObject(lpdis->hDC, ghNormalFont)
                  if hFont then DeleteFont(hFont)
                  
 
                  ' Draw the border edges
                  SetBkMode(lpdis->hDC, TRANSPARENT)   
                  rc  = lpdis->rcItem: rc.right = rc.left + gPropDivPos
                  rc2 = lpdis->rcItem: rc2.left = rc.right
                  DrawEdge( lpdis->hDC, @rc, EDGE_SUNKEN, BF_BOTTOMRIGHT)
                  DrawEdge( lpdis->hDC, @rc2, EDGE_SUNKEN, BF_BOTTOMRIGHT)
                  
               end if
            END IF
         end if
         
         if lpdis->CtlID = IDC_FRMVDTOOLBOX_LSTEVENTS then
            dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr
            dim pCtrl as clsControl ptr 
            if pDoc then
               pCtrl = pDoc->Controls.GetActiveControl
               if pCtrl THEN
                  rc  = lpdis->rcItem
                  rc2 = lpdis->rcItem: rc2.Left = pWindow->ScaleX(20)
                  
                  ' CLEAR BACKGROUND
                  If (lpdis->itemState And ODS_SELECTED) Then     
                     SetBkColor(lpdis->hDC, GetSysColor(COLOR_HIGHLIGHT))   
                     SetTextColor(lpdis->hDC, GetSysColor(COLOR_HIGHLIGHTTEXT))
                     hBrush = GetSysColorBrush(COLOR_HIGHLIGHT) 
                  else
                     SetBkColor(lpdis->hDC, GetSysColor(COLOR_WINDOW))   
                     SetTextColor(lpdis->hDC, GetSysColor(COLOR_WINDOWTEXT))
                     hBrush = GetSysColorBrush(COLOR_WINDOW) 
                  end if

                  ' Draw the tick/untick image
                  ' The line height is LISTBOX_LINEHEIGHT so center the 16x16 icon vertically and horizontally
                  FillRect(lpdis->hDC, @rc, GetSysColorBrush(COLOR_WINDOW))  
                  DrawIconEx( lpdis->hDC, _
                              rc.Left + pWindow->ScaleX(2), _
                              rc.Top + pWindow->ScaleY(2), _
                              iIf( pCtrl->Events(lpdis->itemData).bIsSelected, ghIconTick, ghIconNoTick), _
                              pWindow->ScaleX(16), pWindow->ScaleY(16), 0, 0, DI_NORMAL) 
                  
                  ' Draw the event name 
                  SelectObject(lpdis->hDC, hBrush)      
                  FillRect(lpdis->hDC, @rc2, hBrush)  
                  DrawText( lpdis->hDC, _
                            pCtrl->Events(lpdis->itemData).wszEventName, _
                            -1, Cast(lpRect, @rc2), _
                            DT_LEFT Or DT_SINGLELINE Or DT_VCENTER )

                  ' Draw the border edges
                  SetBkMode(lpdis->hDC, TRANSPARENT)   
                  rc = lpdis->rcItem
                  DrawEdge( lpdis->hDC, @rc, EDGE_SUNKEN, BF_BOTTOMRIGHT)
               end if
            END IF

         end if
         RestoreDC(lpdis->hDC, -1)

         Function = True : Exit Function
   
   End Select

   Function = 0
End Function

   
' ========================================================================================
' Process WM_DESTROY message for window/dialog: frmVDToolbox
' ========================================================================================
Function frmVDToolbox_OnDestroy(HWnd As HWnd) As LRESULT
   HWND_FRMVDTOOLBOX = 0
   Function = 0
End Function


' ========================================================================================
' Processes messages for the subclassed ListBox window.
' ========================================================================================
Function frmVDToolbox_ListBox_SubclassProc ( _
                  ByVal HWnd   As HWnd, _                 ' // Control window handle
                  ByVal uMsg   As UINT, _                 ' // Type of message
                  ByVal wParam As WPARAM, _               ' // First message parameter
                  ByVal lParam As LPARAM, _               ' // Second message parameter
                  ByVal uIdSubclass As UINT_PTR, _        ' // The subclass ID
                  ByVal dwRefData As DWORD_PTR _          ' // Pointer to reference data
                  ) As LRESULT

   dim lpmis As MEASUREITEMSTRUCT Ptr 
   dim lpdis As DRAWITEMSTRUCT Ptr 
   
   Select Case uMsg

      case WM_ERASEBKGND
         ' Only erase the bottom portion of the listbox that extends from the last item
         ' to the bottom edge of the listbox. All other lines are already drawn. This helps
         ' reduce screen flicker.
         ' Determine the number of items in the listbox and multiply by the item height
         dim as RECT rc: GetClientRect(hwnd, @rc)
         rc.top = rc.top + (ListBox_GetCount(hwnd) * ListBox_GetItemHeight(hwnd, 0))
         if rc.top < rc.bottom then
            dim as HDC hDC = cast(HDC, wParam)
            FillRect(hDC, @rc, cast(HBRUSH, GetCurrentObject(hDC, OBJ_BRUSH)))
         end if
         return TRUE
         
      case WM_COMMAND
         if loword(wParam) = IDC_FRMVDTOOLBOX_COMBO then
            if hiword(wParam) = BN_CLICKED then
               dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr()
               dim pProp as clsProperty ptr = GetActivePropertyPtr()
               if pProp then
                  if pProp->PropType = PropertyType.FontPicker then
                     ChooseFontForProperty(pProp)
                     pDoc->UserModified = true
                     pDoc->bRegenerateCode = true
                     frmMain_UpdateLineCol(HWND_FRMMAIN)
                     PostMessage(HWND_FRMVDCOLORS, WM_ACTIVATE, WA_INACTIVE, 0)  ' to apply properties
                  elseif pProp->PropType = PropertyType.ColorPicker then
                     dim as hwnd hList = GetDlgItem(HWND_FRMVDTOOLBOX, IDC_FRMVDTOOLBOX_LSTPROPERTIES)
                     dim as long nCurSel = ListBox_GetCurSel(hList)
                     dim as RECT rcLine
                     ListBox_GetItemRect(hList, nCurSel, @rcLine)
                     MapWindowPoints(hList, 0, cast(POINT ptr, @rcLine), 2)
                     frmVDColors_Show(hList, pProp->wszPropValue)  ' initialize the color popup if not already done so
                     dim as long nWidth = AfxGetWindowWidth(HWND_FRMVDCOLORS)
                     SetWindowPos(HWND_FRMVDCOLORS, HWND_TOP, rcLine.right - nWidth, rcLine.bottom, 0, 0, SWP_NOSIZE or SWP_SHOWWINDOW)
                  else   
                     ' Show the ListBox portion of our "combo" control
                     ShowWindow(HWND_PROPLIST_COMBOLIST, SW_SHOW)
                  end if
               end if   
            end if
         END IF
         if loword(wParam) = IDC_FRMVDTOOLBOX_COMBOLIST then
            if hiword(wParam) = LBN_SELCHANGE then
               InitiatePropertyValueChange
               ShowWindow(HWND_PROPLIST_COMBOLIST, SW_HIDE)
            end if
         END IF
            
      Case WM_GETDLGCODE
         ' All keyboard input
         Function = DLGC_WANTALLKEYS
         Exit Function

      Case WM_KEYUP
         Select Case Loword(wParam)
            Case VK_RETURN
               ' Simulate the sending of a LBN_DBLCLK to the control.
               SendMessage( GetParent(HWnd), WM_COMMAND, MAKEWPARAM(IDC_FRMFNLIST_LISTBOX, LBN_DBLCLK), Cast(LPARAM,HWnd) )
               Exit Function
         End Select

      case WM_LBUTTONDOWN
         ' Hit test to see if the mouse is over the vertical splitter bar
         if hwnd = GetDlgItem(HWND_FRMVDTOOLBOX, IDC_FRMVDTOOLBOX_LSTPROPERTIES) THEN
            dim as long hPos = loword(lParam)
            dim as RECT rc: GetClientRect(hwnd, @rc)
            if (hPos >= gPropDivPos - 2) andalso (hPos <= gPropDivPos + 2) THEN
               SetCursor( LoadCursor(0, ByVal IDC_SIZEWE) )
               'InvertLine(hwnd, gPropDivPos, rc.top, gPropDivPos, rc.bottom)
               gPropDivTracking = true
               SetCapture(hwnd)
            else
               SetCursor( LoadCursor(0, ByVal IDC_ARROW) )
            END IF
         else
            SetCursor( LoadCursor(0, ByVal IDC_ARROW) )
         END IF
         
      case WM_LBUTTONUP
         if hwnd = GetDlgItem(HWND_FRMVDTOOLBOX, IDC_FRMVDTOOLBOX_LSTPROPERTIES) THEN
            if gPropDivTracking = true THEN
               gPropDivTracking = false
               ReleaseCapture
               AfxRedrawWindow(hwnd)
            end if
         end if
         
         if hwnd = GetDlgItem(HWND_FRMVDTOOLBOX, IDC_FRMVDTOOLBOX_LSTEVENTS) THEN
            ' Handle if the Events checkbox is clicked
            dim as long nCurSel = ListBox_GetCurSel(hwnd)
            if nCurSel = -1 then exit function
            dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr
            dim pCtrl as clsControl ptr 
            if pDoc then
               pCtrl = pDoc->Controls.GetActiveControl
               if pCtrl THEN
                  Dim pWindow As CWindow Ptr = AfxCWindowPtr(GetParent(HWnd))
                  if pWindow = 0 THEN exit function
                  dim as RECT rc: SendMessage(hwnd, LB_GETITEMRECT, nCurSel, cast(LPARAM, @rc))
                  ' The checkbox is the first 20x20 area (16x16 icon)
                  rc.Left   = rc.Left + pWindow->ScaleX(2)
                  rc.Top    = rc.Top  + pWindow->ScaleY(2)
                  rc.Right  = rc.Left + pWindow->ScaleX(16)
                  rc.Bottom = rc.Top  + pWindow->ScaleY(16)
                  dim as POINT pt = (loword(lParam), Hiword(lParam))
                  if PtInRect(@rc, pt) then
                     pCtrl->Events(nCurSel).bIsSelected = not pCtrl->Events(nCurSel).bIsSelected
                     InvalidateRect(hwnd, @rc, true): UpdateWindow(hwnd)
                     pDoc->UserModified = true
                     pDoc->bRegenerateCode = true
                     frmMain_UpdateLineCol(HWND_FRMMAIN)
                  END IF
               end if
            end if   
         end if
         
         
      case WM_MOUSEMOVE
         if hwnd = GetDlgItem(HWND_FRMVDTOOLBOX, IDC_FRMVDTOOLBOX_LSTPROPERTIES) THEN
            ' Hit test to see if the mouse is over the vertical splitter bar
            dim as long hPos = loword(lParam)
            dim as RECT rc: GetClientRect(hwnd, @rc)
            if (hPos >= gPropDivPos - 2) andalso (hPos <= gPropDivPos + 2) THEN
               SetCursor( LoadCursor(0, ByVal IDC_SIZEWE) )
               if gPropDivTracking = true THEN
                  gPropDivPos = loword(lParam)
                  AfxRedrawWindow(hwnd)
               end if
            else
               SetCursor( LoadCursor(0, ByVal IDC_ARROW) )
            END IF
         end if
         
      Case WM_DESTROY
         ' REQUIRED: Remove control subclassing
         RemoveWindowSubclass HWnd, @frmVDToolBox_ListBox_SubclassProc, uIdSubclass

   End Select

   ' Default processing of Windows messages
   Function = DefSubclassProc(HWnd, uMsg, wParam, lParam)

End Function

               
' ========================================================================================
' Processes messages for the subclassed TextEdit window.
' ========================================================================================
Function frmVDToolbox_TextEdit_SubclassProc ( _
                  ByVal HWnd   As HWnd, _                 ' // Control window handle
                  ByVal uMsg   As UINT, _                 ' // Type of message
                  ByVal wParam As WPARAM, _               ' // First message parameter
                  ByVal lParam As LPARAM, _               ' // Second message parameter
                  ByVal uIdSubclass As UINT_PTR, _        ' // The subclass ID
                  ByVal dwRefData As DWORD_PTR _          ' // Pointer to reference data
                  ) As LRESULT

   Select Case uMsg

      Case WM_GETDLGCODE
         ' All keyboard input
         Function = DLGC_WANTALLKEYS
         Exit Function

      case WM_KILLFOCUS
         InitiatePropertyValueChange
         HidePropertyListControls
         
      Case WM_CHAR   ' filter characters and also prevent the annoying beep!
         dim as clsProperty ptr pProp = GetActivePropertyPtr() 
         if pProp then 
            if pProp->PropType = PropertyType.EditEnterNumericOnly then
               select case wParam
                  CASE 48 to 57    ' 0 to 9 (only allow these characters)
                  case 8           ' allow backspace   
                  case else        ' Filter out all other characters
                     return 0
               END SELECT
            END IF
         END IF
         If wParam = 13 Then Return 0  ' ENTER
         If wParam = 27 Then Return 0  ' ESC
         

      Case WM_KEYUP
         Select Case Loword(wParam)
            Case VK_RETURN
               ' Hide the edit control. WM_KILLFOCUS will be fired.
               InitiatePropertyValueChange  ' must do here otherwise edit control is not visible and change property will fail
               HidePropertyListControls
            
            Case VK_ESCAPE
               ' Reset the edit value to original value
               dim as clsProperty ptr pProp = GetActivePropertyPtr() 
               if pProp then AfxSetWindowText(HWND_PROPLIST_EDIT, pProp->wszPropValue)
               HidePropertyListControls
         End Select

      Case WM_DESTROY
         ' REQUIRED: Remove control subclassing
         RemoveWindowSubclass HWnd, @frmVDToolBox_TextEdit_SubclassProc, uIdSubclass

   End Select

   ' Default processing of Windows messages
   Function = DefSubclassProc(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' frmVDToolbox Window procedure
' ========================================================================================
Function frmVDToolbox_SUBCLASSPROC( ByVal HWnd   As HWnd, _
                               ByVal uMsg   As UINT, _
                               ByVal wParam As WPARAM, _
                               ByVal lParam As LPARAM _
                               ) As LRESULT

   Select Case uMsg
      HANDLE_MSG (HWnd, WM_COMMAND,     frmVDToolbox_OnCommand)
      HANDLE_MSG (HWnd, WM_NOTIFY,      frmVDToolbox_OnNotify)
      HANDLE_MSG (HWnd, WM_SIZE,        frmVDToolbox_OnSize)
      HANDLE_MSG (HWnd, WM_CLOSE,       frmVDToolbox_OnClose)
      HANDLE_MSG (HWnd, WM_DESTROY,     frmVDToolbox_OnDestroy)
      HANDLE_MSG (HWnd, WM_MEASUREITEM, frmVDToolbox_OnMeasureItem)
      HANDLE_MSG (HWnd, WM_DRAWITEM,    frmVDToolbox_OnDrawItem)

      case WM_ACTIVATE
         if wParam = WA_INACTIVE then
            InitiatePropertyValueChange
            HidePropertyListControls
         end if 
         
   End Select

   ' for messages that we don't deal with
   Function = DefWindowProc(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' frmVDToolbox_Show
' ========================================================================================
Function frmVDToolbox_Show( ByVal hWndParent As HWnd, _
                            ByVal nCmdShow   As Long = 0 _
                            ) As Long

   ' If Toolbox already exists then toggle its visibility
   ' unless nCmdShow is explicitly set to show it.
   If IsWindow(HWND_FRMVDTOOLBOX) Then
      If nCmdShow <> SW_SHOW Then nCmdShow = Iif(IsWindowVisible(HWND_FRMVDTOOLBOX), SW_HIDE, SW_SHOW)
      ShowWindow HWND_FRMVDTOOLBOX, nCmdShow
      If nCmdShow = SW_SHOW Then
         SetFocus GetDlgItem(HWND_FRMFNLIST, IDC_FRMFNLIST_LISTBOX)
      End If
      Exit Function 
   Else 
      ' If the window does not exist yet then ensure that it becomes visible after creation.
      nCmdShow = SW_SHOW
   End If
      

   '  Create the main window and child controls
   Dim pWindow As CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowOwnerPtr(hwndParent)->DPI

   Dim rcWork  As Rect = pWindow->GetWorkArea
   Dim nHeight As Long = (rcWork.Bottom - rcWork.Top) * .70
   dim nLeft   as long = ((rcWork.Right - rcWork.Left) * .65) 
   Dim nTop    As Long = (rcWork.Bottom - rcWork.Top - nHeight) / 2
      
   HWND_FRMVDTOOLBOX = _
   pWindow->Create( hWndParent, L(352,"Toolbox"), @frmVDToolbox_SUBCLASSPROC, nLeft, nTop, 300, nHeight, _
        WS_POPUP Or WS_CAPTION or WS_SYSMENU or WS_THICKFRAME or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, _
        WS_EX_CLIENTEDGE)
   ' We will set our own mouse pointer as needed
   SetClassLongPtr(HWND_FRMVDTOOLBOX, GCLP_HCURSOR, 0)   
   
   ' Set the small and large icon for the main window (must be set after main window is created)
   pWindow->BigIcon   =  LoadImage( pWindow->InstanceHandle, "IMAGE_AAA_MAINICON", IMAGE_ICON, 32, 32, LR_SHARED)
   pWindow->SmallIcon =  LoadImage( pWindow->InstanceHandle, "IMAGE_AAA_MAINICON", IMAGE_ICON, 16, 16, LR_SHARED)

   Dim As HWnd hTabCtl = _ 
        pWindow->AddControl("TAB", , IDC_FRMVDTOOLBOX_TABCONTROL, "", 0, 0, 0, 0)
        
   TabCtrl_AddTab(hTabCtl, 0, L(351,"Tools"))
   TabCtrl_AddTab(hTabCtl, 0, L(350,"Properties"))
   TabCtrl_AddTab(hTabCtl, 0, L(353,"Events"))
   
   Dim As HWnd hCombo = _ 
        pWindow->AddControl("COMBOBOX", , IDC_FRMVDTOOLBOX_COMBOCONTROLS, "", 0, 0, 0, 0, _
        WS_CHILD or WS_TABSTOP Or CBS_DROPDOWNLIST or CBS_SORT or CBS_OWNERDRAWFIXED or CBS_HASSTRINGS Or WS_VSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_RIGHTSCROLLBAR)

   Dim As HWnd hList1 = _ 
        pWindow->AddControl("LISTBOX", , IDC_FRMVDTOOLBOX_LSTTOOLBOX, "", 0, 0, 0, 0, _
        WS_CHILD Or WS_VISIBLE or WS_TABSTOP Or LBS_NOINTEGRALHEIGHT Or WS_VSCROLL or  _
        LBS_OWNERDRAWFIXED Or LBS_HASSTRINGS Or LBS_NOTIFY, WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_RIGHTSCROLLBAR, , _
        Cast(SUBCLASSPROC, @frmVDToolbox_ListBox_SubclassProc), IDC_FRMVDTOOLBOX_LSTTOOLBOX, Cast(DWORD_PTR, @pWindow))

   Dim As HWnd hList2 = _ 
        pWindow->AddControl("LISTBOX", , IDC_FRMVDTOOLBOX_LSTPROPERTIES, "", 0, 0, 0, 0, _
        WS_CHILD Or WS_TABSTOP Or LBS_NOINTEGRALHEIGHT Or WS_VSCROLL or _
        LBS_OWNERDRAWFIXED Or LBS_HASSTRINGS Or LBS_SORT or LBS_NOTIFY, WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_RIGHTSCROLLBAR, , _
        Cast(SUBCLASSPROC, @frmVDToolbox_ListBox_SubclassProc), IDC_FRMVDTOOLBOX_LSTPROPERTIES, Cast(DWORD_PTR, @pWindow))

   Dim As HWnd hLabel = _ 
   pWindow->AddControl("LABEL", , IDC_FRMVDTOOLBOX_LBLPROPNAME, "", 0, 0, 0, 0, _
        WS_CHILD Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT, WS_EX_LEFT Or WS_EX_LTRREADING)
   AfxSetWindowFont(hLabel, ghBoldFont)     
        
   pWindow->AddControl("LABEL", , IDC_FRMVDTOOLBOX_LBLPROPDESCRIBE, "", 0, 0, 0, 0, _
        WS_CHILD Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT, WS_EX_LEFT Or WS_EX_LTRREADING)

   Dim As HWnd hList3 = _ 
        pWindow->AddControl("LISTBOX", , IDC_FRMVDTOOLBOX_LSTEVENTS, "", 0, 0, 0, 0, _
        WS_CHILD Or WS_TABSTOP Or LBS_NOINTEGRALHEIGHT Or WS_VSCROLL or _
        LBS_OWNERDRAWFIXED Or LBS_HASSTRINGS Or LBS_NOTIFY, WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_RIGHTSCROLLBAR, , _
        Cast(SUBCLASSPROC, @frmVDToolbox_ListBox_SubclassProc), IDC_FRMVDTOOLBOX_LSTEVENTS, Cast(DWORD_PTR, @pWindow))
   
   HWND_PROPLIST_EDIT = _
        pWindow->AddControl("TEXTBOX", hList2, IDC_FRMVDTOOLBOX_TEXTEDIT, "", 0, 0, 0, 0, _
            WS_CHILD Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL, _
            0, , _
            Cast(SUBCLASSPROC, @frmVDToolbox_TextEdit_SubclassProc), IDC_FRMVDTOOLBOX_TEXTEDIT, Cast(DWORD_PTR, @pWindow))

   DIM hBitmap AS HBITMAP
   HWND_PROPLIST_COMBO = _
         pWindow->AddControl("BITMAPBUTTON", hList2, IDC_FRMVDTOOLBOX_COMBO, "", 0, 0, 0, 0, _
              WS_CHILD Or WS_TABSTOP Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or BS_BITMAP Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER Or LR_DEFAULTCOLOR Or LR_LOADMAP3DCOLORS Or LR_LOADTRANSPARENT Or LR_SHARED, _
              WS_EX_LEFT Or WS_EX_LTRREADING)

   dim wszImage as wstring * 100
   wszImage = iif(pWindow->DPI > 96, "IMAGE_DOWN", "IMAGE_DOWN16")
         hBitmap = AfxGdipImageFromRes(pWindow->InstanceHandle, wszImage, 0, false, IMAGE_BITMAP, 0)
         SendMessage(HWND_PROPLIST_COMBO, BM_SETIMAGE, IMAGE_BITMAP, cast(LPARAM, hBitmap))
         IF hBitmap THEN DeleteObject(hBitmap)

   HWND_PROPLIST_COMBOLIST = _ 
        pWindow->AddControl("LISTBOX", hList2, IDC_FRMVDTOOLBOX_COMBOLIST, "", 0, 0, 0, 120, _
        WS_CHILD Or WS_BORDER or WS_VSCROLL or LBS_NOINTEGRALHEIGHT or LBS_HASSTRINGS Or LBS_NOTIFY, WS_EX_LEFT Or WS_EX_RIGHTSCROLLBAR)

            
   ' Set the default position for the vertical sizing bar
   gPropDivPos = pWindow->ScaleY(0.4 * 300)
   
   ' Add the Tools to the toolbox
   dim as long ndx
   for i as long = lbound(gToolBox) to ubound(gToolBox)
      ndx = ListBox_AddString(hList1, gToolBox(i).wszToolBoxName)
      ListBox_SetItemData(hList1, ndx, i)
   NEXT
   SetActiveToolboxControl(CTRL_POINTER)
   
   ' Initialize the popup Colors selection dialog
   frmVDColors_Show(hList2, "")
   
   frmVDToolbox_PositionWindows(HWND_FRMVDTOOLBOX)

   TabCtrl_SetCurSel(hTabCtl, 0)
   ShowWindow HWND_FRMVDTOOLBOX, nCmdShow
   ListBox_SetCurSel(hList1, 0)
   SetFocus hList1
   
   Function = 0
End Function



