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

#include once "modVDRoutines.bi"
#include once "frmStatusBarEditor.bi"


' ========================================================================================
' Retrieve the IMAGES_TYPE pointer for the IMAGE name.
' ========================================================================================
public Function GetImagesTypePtr( byref wszImageName as wstring ) As IMAGES_TYPE ptr
   dim pDoc as clsDocument ptr = gApp.pDocList
   do until pDoc = 0
      if (gApp.IsProjectActive = false) andalso _
         (gTTabCtl.GetActiveDocumentPtr <> 0) andalso _
         (pDoc <> gTTabCtl.GetActiveDocumentPtr) then 
         ' For non-projects only get images related to the active form.
      else
         for i as long = lbound(pDoc->AllImages) to ubound(pDoc->AllImages)
            if pDoc->AllImages(i).wszImageName = wszImageName then
               return @pDoc->AllImages(i)
            end if   
         next   
      end if
      pDoc = pDoc->pDocNext
   loop
   function = 0
end function


' ========================================================================================
' Check for a minimum default size when creating. Check needed because a user could
' select a control in the Toolbox and simply click on the Form rather than drawing
' the control on the form.
' ========================================================================================
function CheckMinimumControlSize( byval nControlType as long, byval rcIn as RECT ) as RECT
   dim as RECT rc = rcIn
   dim as long nWidth = rc.right - rc.left
   dim as long nHeight = rc.bottom - rc.top
   
   if (nWidth < 5) orelse (nHeight < 5) then
      'AFXMSG( "Control too small" )
      select case nControlType 
         case CTRL_LABEL:             nWidth = 97:  nHeight = 41  
         case CTRL_BUTTON:            nWidth = 188: nHeight = 58  
         case CTRL_TEXTBOX:           nWidth = 250: nHeight = 47  
         case CTRL_CHECKBOX:          nWidth = 197: nHeight = 45  
         case CTRL_OPTION:            nWidth = 228: nHeight = 45  
         case CTRL_FRAME:             nWidth = 500: nHeight = 250  
         case CTRL_PICTUREBOX:        nWidth = 250: nHeight = 125  
         case CTRL_COMBOBOX:          nWidth = 302: nHeight = 49  
         case CTRL_LISTBOX:           nWidth = 300: nHeight = 209  
         case CTRL_RICHEDIT:          nWidth = 250: nHeight = 240  
         case CTRL_MASKEDEDIT:        nWidth = 250: nHeight = 47  
         case CTRL_PROGRESSBAR:       nWidth = 250: nHeight = 58  
         case CTRL_LISTVIEW:          nWidth = 302: nHeight = 242  
         case CTRL_TREEVIEW:          nWidth = 302: nHeight = 242  
         case CTRL_MONTHCALENDAR:     nWidth = 489: nHeight = 399  
         case CTRL_DATETIMEPICKER:    nWidth = 500: nHeight = 47  
         case CTRL_TABCONTROL:        nWidth = 500: nHeight = 58  
         case CTRL_UPDOWN:            nWidth = 44:  nHeight = 47  
         case CTRL_TIMER:             nWidth = 32:  nHeight = 32  
         case CTRL_HSCROLL:           nWidth = 200: nHeight = 108  
         case CTRL_VSCROLL:           nWidth = 108: nHeight = 500   
         case CTRL_SLIDER:            nWidth = 260: nHeight = 114   
         case CTRL_WEBBROWSER:
         case CTRL_CUSTOM:
         case CTRL_OCX:
      end select
      
      rc.right = rc.left + AfxUnScaleX(nWidth)
      rc.bottom = rc.top + AfxUnScaleY(nHeight)
   end if
   
   return rc
end function


' ========================================================================================
' Retrieve the control type that is currently actively selected in the Toolbox
' ========================================================================================
public function GetActiveToolboxControlType() as Long

   Dim As HWnd hList1 = GetDlgItem(HWND_FRMVDTOOLBOX, IDC_FRMVDTOOLBOX_LSTTOOLBOX)
   dim as long nCurSel = ListBox_GetCurSel(hList1)
   
   if nCurSel = -1 THEN return CTRL_POINTER
   
   ' The index into the global gToolbox array is stored in the line's data area.
   dim as long idx = ListBox_GetItemData(hList1, nCurSel)
   function = gToolbox(idx).nToolType
end function


' ========================================================================================
' Retrieve a pointer to the currently selected Property in the PropertyList
' ========================================================================================
public function GetActivePropertyPtr() as clsProperty ptr

   ' This function allows for two use cases: (1) is to be able to choose colors
   ' in the PropertyList for various controls, and (2) is to be able to select
   ' colors for StatusBar Panels. If the StatusBar Editor is active then we
   ' must be looking for the pPropColor related to that panel, otherwise it
   ' must be the PropertyList that is being used.
   
   if IsWindowVisible(HWND_FRMSTATUSBAREDITOR) then
      dim as hwnd hList1 = GetDlgItem( HWND_FRMSTATUSBAREDITOR, IDC_FRMSTATUSBAREDITOR_LSTPANELS)
      dim as long nCurSel = ListBox_GetCurSel(hList1)
      if nCurSel = -1 then return 0
      return @gPanelItems(nCurSel).pPropColor
      
   else   
      Dim As HWnd hList = GetDlgItem(HWND_FRMVDTOOLBOX, IDC_FRMVDTOOLBOX_LSTPROPERTIES)
      dim as long nCurSel = ListBox_GetCurSel(hList)
      if nCurSel = -1 then return 0

      dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr
      dim pCtrl as clsControl ptr 
      if pDoc then pCtrl = pDoc->Controls.GetActiveControl
      if pCtrl = 0 THEN return 0

      dim as long idx = ListBox_GetItemData(hList, nCurSel)   ' property array index in listbox item
      
      return @pCtrl->Properties(idx)
   end if
   
end function


' ========================================================================================
' Retrieve a pointer to the currently selected Event in the PropertyList
' ========================================================================================
public function GetActiveEventPtr() as clsEvent ptr

   Dim As HWnd hList = GetDlgItem(HWND_FRMVDTOOLBOX, IDC_FRMVDTOOLBOX_LSTEVENTS)
   dim as long nCurSel = ListBox_GetCurSel(hList)
   if nCurSel = -1 then return 0

   dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr
   dim pCtrl as clsControl ptr 
   if pDoc then pCtrl = pDoc->Controls.GetActiveControl
   if pCtrl = 0 THEN return 0

   dim as long idx = ListBox_GetItemData(hList, nCurSel)   ' event array index in listbox item
   
   return @pCtrl->Events(idx)
end function


' ========================================================================================
' Get the pCtrl pointer for the Form
' ========================================================================================
public function GetFormCtrlPtr( byval pDoc as clsDocument ptr ) as clsControl ptr
   dim pCtrl as clsControl ptr
   for i as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
      pCtrl = pDoc->Controls.ItemAt(i)
      if pCtrl->ControlType = CTRL_FORM then return pCtrl
   NEXT
   return 0
end function


' ========================================================================================
' Set the Tools listbox to incoming control type.
' ========================================================================================
public function SetActiveToolboxControl( byval ControlType as long ) as Long

   Dim As HWnd hList1 = GetDlgItem(HWND_FRMVDTOOLBOX, IDC_FRMVDTOOLBOX_LSTTOOLBOX)
   dim as long NumItems = ListBox_GetCount(hList1)
   dim as long idx
   
   for i as long = 0 to NumItems
      idx = ListBox_GetItemData(hList1, i)
      if gToolbox(idx).nToolType = ControlType THEN 
         ListBox_SetCurSel(hList1, i)
         exit for
      end if   
   NEXT
   
   function = 0
end function


' ========================================================================================
' Get the class name based on the type of control
' ========================================================================================
public function GetControlClassName( byval pCtrl as clsControl ptr ) as CWSTR
   for i as long = lbound(gToolbox) to ubound(gToolbox)
      if gToolbox(i).nToolType = pCtrl->ControlType THEN 
         return gToolbox(i).wszClassName
      END IF
   NEXT
   return ""
end function

   
' ========================================================================================
' Get the control ToolBox name based on the type of control
' ========================================================================================
public function GetToolBoxName( byval nControlType as long ) as CWSTR
   if nControlType = CTRL_FORM then return "Form"
   for i as long = lbound(gToolbox) to ubound(gToolbox)
      if gToolbox(i).nToolType = nControlType THEN 
         return gToolbox(i).wszToolBoxName
      END IF
   NEXT
   return ""
end function


' ========================================================================================
' Get the control name based on the type of control
' ========================================================================================
public function GetControlName( byval nControlType as long ) as CWSTR
   if nControlType = CTRL_FORM then return "Form"
   for i as long = lbound(gToolbox) to ubound(gToolbox)
      if gToolbox(i).nToolType = nControlType THEN 
         return gToolbox(i).wszControlName
      END IF
   NEXT
   return ""
end function


' ========================================================================================
' Get the control type number based on the name of the control
' ========================================================================================
public function GetControlType( byval wszControlName as CWSTR ) as long
   wszControlName = ucase(wszControlName)
   if wszControlName = "FORM" then return CTRL_FORM 
   for i as long = lbound(gToolbox) to ubound(gToolbox)
      if ucase(gToolbox(i).wszToolBoxName) = wszControlName then
         return gToolbox(i).nToolType 
      END IF
   NEXT
   return 0
end function


' ========================================================================================
' Get the WinFormsX class name for the incoming control type
' ========================================================================================
public function GetWinformsXClassName( byval nControlType as long ) as CWSTR
   dim wszText as CWSTR
   select case nControlType 
      CASE CTRL_FORM:              wszText = "wfxForm"
      CASE CTRL_LABEL:             wszText = "wfxLabel"
      CASE CTRL_BUTTON:            wszText = "wfxButton"   
      CASE CTRL_TEXTBOX:           wszText = "wfxTextBox"  
      CASE CTRL_CHECKBOX:          wszText = "wfxCheckBox" 
      CASE CTRL_OPTION:            wszText = "wfxOptionButton"
      CASE CTRL_FRAME:             wszText = "wfxFrame"    
      CASE CTRL_PICTUREBOX:        wszText = "wfxPictureBox"  
      CASE CTRL_COMBOBOX:          wszText = "wfxComboBox" 
      CASE CTRL_LISTBOX:           wszText = "wfxListBox"  
      case CTRL_RICHEDIT:          wszText = "wfxRichEdit"
      case CTRL_MASKEDEDIT:        wszText = "wfxMaskedEdit"
      case CTRL_PROGRESSBAR:       wszText = "wfxProgressBar"
      case CTRL_LISTVIEW:          wszText = "wfxListView"
      case CTRL_TREEVIEW:          wszText = "wfxTreeView"
      case CTRL_MONTHCALENDAR:     wszText = "wfxMonthCalendar"
      case CTRL_DATETIMEPICKER:    wszText = "wfxDateTimePicker"
      case CTRL_TABCONTROL:        wszText = "wfxTabControl"
      case CTRL_UPDOWN:            wszText = "wfxUpDown"
      case CTRL_TIMER:             wszText = "wfxTimer"
      case CTRL_HSCROLL
      case CTRL_VSCROLL
      case CTRL_SLIDER
      case CTRL_WEBBROWSER
      case CTRL_CUSTOM
      case CTRL_OCX
   END SELECT

   return wszText
end function


' ========================================================================================
' Get the RECT of the specified control 
' ========================================================================================
public function GetControlRECT( byval pCtrl as clsControl ptr ) as RECT
   if pCtrl = 0 then exit function
   dim as long lb = lbound(pCtrl->Properties)
   dim as long ub = ubound(pCtrl->Properties)
   dim as long nLeft, nTop, nWidth, nHeight
   dim as RECT rc
   
   for i as long = lb to ub
      select CASE ucase(pCtrl->Properties(i).wszPropName)
         case "LEFT":   nLeft   = val(pCtrl->Properties(i).wszPropValue)
         case "TOP":    nTop    = val(pCtrl->Properties(i).wszPropValue)
         case "WIDTH":  nWidth  = val(pCtrl->Properties(i).wszPropValue)
         case "HEIGHT": nHeight = val(pCtrl->Properties(i).wszPropValue)
      END SELECT
   NEXT
   SetRect(@rc, nLeft, nTop, nLeft+nWidth, nTop+nHeight)
   function = rc
end function

