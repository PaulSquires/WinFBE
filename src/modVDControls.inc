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
' Retrieve the value for the specified control property
' ========================================================================================
function GetControlProperty( byval pCtrl as clsControl ptr, _
                             byval wszPropName as CWSTR _
                             ) as CWSTR
   if pCtrl = 0 THEN return ""
   dim as long lb = lbound(pCtrl->Properties)
   dim as long ub = ubound(pCtrl->Properties)
   dim as CWSTR wszPropName_ucase = ucase(wszPropName)
   
   ' All property values are returned as strings. 
   for i as long = lb to ub
      if ucase(pCtrl->Properties(i).wszPropName) = wszPropName_ucase THEN
         return pCtrl->Properties(i).wszPropValue 
      END IF
   NEXT

   return ""
end function

' ========================================================================================
' Set the value for the specified control property
' ========================================================================================
function SetControlProperty( byval pCtrl as clsControl ptr, _
                             byval wszPropName as CWSTR, _
                             byval wszPropValue as CWSTR _
                             ) as long
   dim as long lb = lbound(pCtrl->Properties)
   dim as long ub = ubound(pCtrl->Properties)
   dim as CWSTR wszPropName_ucase = ucase(wszPropName)
   
   for i as long = lb to ub
      if ucase(pCtrl->Properties(i).wszPropName) = wszPropName_ucase THEN
         pCtrl->Properties(i).wszPropValue = wszPropValue
         exit for
      END IF
   NEXT
   function = 0
end function


' ========================================================================================
' Get the RECT of the specified control 
' ========================================================================================
function GetControlRECT( byval pCtrl as clsControl ptr ) as RECT
   dim as long lb = lbound(pCtrl->Properties)
   dim as long ub = ubound(pCtrl->Properties)
   dim as long nLeft, nTop, nWidth, nHeight
   dim as RECT rc
   
   for i as long = lb to ub
      select CASE ucase(**pCtrl->Properties(i).wszPropName)
         case "LEFT":   nLeft   = val(pCtrl->Properties(i).wszPropValue)
         case "TOP":    nTop    = val(pCtrl->Properties(i).wszPropValue)
         case "WIDTH":  nWidth  = val(pCtrl->Properties(i).wszPropValue)
         case "HEIGHT": nHeight = val(pCtrl->Properties(i).wszPropValue)
      END SELECT
   NEXT
   SetRect(@rc, nLeft, nTop, nLeft+nWidth, nTop+nHeight)
   function = rc
end function


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
               case VK_DOWN
                  nTemp = val(GetControlProperty(pCtrl, "TOP")) + 1
                  SetControlProperty(pCtrl, "TOP", str(nTemp))
               case VK_LEFT
                  nTemp = val(GetControlProperty(pCtrl, "LEFT")) - 1
                  SetControlProperty(pCtrl, "LEFT", str(nTemp))
               case VK_RIGHT         
                  nTemp = val(GetControlProperty(pCtrl, "LEFT")) + 1
                  SetControlProperty(pCtrl, "LEFT", str(nTemp))
            END SELECT
            pCtrl->SuspendLayout = true
            ApplyControlProperties( pCtrl )
            pCtrl->SuspendLayout = false
         end if
      end if
   next

   AfxRedrawWindow(pDoc->hWndForm)
   frmMain_UpdateLineCol(HWND_FRMMAIN)

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
                  nTemp = val(GetControlProperty(pCtrl, "TOP")) - 1
                  SetControlProperty(pCtrl, "TOP", str(nTemp))
                  nTemp = val(GetControlProperty(pCtrl, "HEIGHT")) + 1
                  SetControlProperty(pCtrl, "HEIGHT", str(nTemp))
               case VK_DOWN
                  nTemp = val(GetControlProperty(pCtrl, "HEIGHT")) + 1
                  SetControlProperty(pCtrl, "HEIGHT", str(nTemp))
               case VK_LEFT
                  nTemp = val(GetControlProperty(pCtrl, "WIDTH")) - 1
                  SetControlProperty(pCtrl, "WIDTH", str(nTemp))
               case VK_RIGHT         
                  nTemp = val(GetControlProperty(pCtrl, "WIDTH")) + 1
                  SetControlProperty(pCtrl, "WIDTH", str(nTemp))
            END SELECT
            pCtrl->SuspendLayout = true
            ApplyControlProperties( pCtrl )
            pCtrl->SuspendLayout = false
         end if
      end if
   next

   AfxRedrawWindow(pDoc->hWndForm)
   frmMain_UpdateLineCol(HWND_FRMMAIN)

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
   
   ' Get the current active control in the group
   dim pCtrlActive as clsControl ptr = pDoc->Controls.GetActiveControl
   
   ' Determine the array position that the active control is at.
   for i as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
      if pDoc->Controls.ItemAt(i) = pCtrlActive then
         idxActive = i: exit for
      end if
   next      
   
   select CASE vKeycode
      case VK_UP, VK_LEFT    ' get the previous control in the group
         for i as long = (idxActive - 1) to pDoc->Controls.ItemFirst step -1
            pCtrl = pDoc->Controls.ItemAt(i)
            if pCtrl->ControlType = CTRL_FORM THEN continue for
            if pCtrl->IsSelected then
               idxPrev = i: exit for
            end if   
         next
         if idxPrev = -1 THEN  ' prev still not found. wrap to end of array
            for i as long = pDoc->Controls.ItemLast to idxActive step -1
               pCtrl = pDoc->Controls.ItemAt(i)
               if pCtrl->ControlType = CTRL_FORM THEN continue for
               if pCtrl->IsSelected then
                  idxPrev = i: exit for
               end if   
            next
         end if
         pCtrl = pDoc->Controls.ItemAt(idxPrev)
         pDoc->Controls.SetActiveControl(pCtrl->hWindow)
         
      case VK_DOWN, VK_RIGHT   ' get the next control in the gorup
         for i as long = (idxActive + 1) to pDoc->Controls.ItemLast
            pCtrl = pDoc->Controls.ItemAt(i)
            if pCtrl->ControlType = CTRL_FORM THEN continue for
            if pCtrl->IsSelected then
               idxNext = i: exit for
            end if   
         next
         if idxNext = -1 THEN  ' next still not found. wrap to start of array
            for i as long = pDoc->Controls.ItemFirst to idxActive 
               pCtrl = pDoc->Controls.ItemAt(i)
               if pCtrl->ControlType = CTRL_FORM THEN continue for
               if pCtrl->IsSelected then
                  idxNext = i: exit for
               end if   
            next
         end if
         pCtrl = pDoc->Controls.ItemAt(idxNext)
         pDoc->Controls.SetActiveControl(pCtrl->hWindow)

   end select

   AfxRedrawWindow(pDoc->hWndForm)
   frmMain_UpdateLineCol(HWND_FRMMAIN)

   function = 0
end function


' ========================================================================================
' Retrieve the value for the defined STYLE of the control
' ========================================================================================
function GetStyleValue( byval pCtrl as clsControl ptr, _
                        byref wszStyleType as CWSTR _
                        ) as long 
   dim as long dwStyle = 0
   dim as CWSTR wszStyle = GetControlProperty(pCtrl, wszStyleType)   ' STYLE, EXSTYLE
   
   dim as long nCount = AfxStrParseCount(wszStyle, "|")
   for i as long = 1 to nCount 
      select CASE **AfxStrParse(wszStyle, i, "|")
         case "WS_VISIBLE":       dwStyle = dwStyle or WS_VISIBLE
         case "WS_GROUP":         dwStyle = dwStyle or WS_GROUP
         case "WS_TABSTOP":       dwStyle = dwStyle or WS_TABSTOP
         Case "WS_BORDER":        dwStyle = dwStyle Or WS_BORDER
         Case "WS_HSCROLL":       dwStyle = dwStyle Or WS_HSCROLL
         Case "WS_VSCROLL":       dwStyle = dwStyle Or WS_VSCROLL

         Case "WS_EX_CLIENTEDGE":     dwStyle = dwStyle Or WS_EX_CLIENTEDGE
         Case "WS_EX_STATICEDGE":     dwStyle = dwStyle Or WS_EX_STATICEDGE
         Case "WS_EX_DLGMODALFRAME":  dwStyle = dwStyle Or WS_EX_DLGMODALFRAME
         Case "WS_EX_LEFT":           dwStyle = dwStyle Or WS_EX_LEFT
         Case "WS_EX_RIGHT":          dwStyle = dwStyle or WS_EX_RIGHT
         Case "WS_EX_LTRREADING":     dwStyle = dwStyle Or WS_EX_LTRREADING
         Case "WS_EX_RTLREADING":     dwStyle = dwStyle Or WS_EX_RTLREADING
         Case "WS_EX_LEFTSCROLLBAR":  dwStyle = dwStyle Or WS_EX_LEFTSCROLLBAR
         Case "WS_EX_RIGHTSCROLLBAR": dwStyle = dwStyle Or WS_EX_RIGHTSCROLLBAR 

         case "SS_LEFT":          dwStyle = dwStyle or SS_LEFT
         Case "SS_CENTER":        dwStyle = dwStyle Or SS_CENTER
         Case "SS_CENTERIMAGE":   dwStyle = dwStyle Or SS_CENTERIMAGE
         Case "SS_RIGHT":         dwStyle = dwStyle Or SS_RIGHT
         Case "SS_SIMPLE":        dwStyle = dwStyle Or SS_SIMPLE
         Case "SS_SUNKEN":        dwStyle = dwStyle Or SS_SUNKEN
         case "SS_NOTIFY":        dwStyle = dwStyle or SS_NOTIFY

         case "BS_AUTOCHECKBOX":  dwStyle = dwStyle Or BS_AUTOCHECKBOX
         case "BS_AUTORADIOBUTTON": dwStyle = dwStyle Or BS_AUTORADIOBUTTON
         Case "BS_MULTILINE":     dwStyle = dwStyle Or BS_MULTILINE
         Case "BS_PUSHBUTTON":    dwStyle = dwStyle Or BS_PUSHBUTTON
         Case "BS_DEFPUSHBUTTON": dwStyle = dwStyle Or BS_DEFPUSHBUTTON
         Case "BS_PUSHLIKE":      dwStyle = dwStyle Or BS_PUSHLIKE
         Case "BS_FLAT":          dwStyle = dwStyle Or BS_FLAT
         Case "BS_LEFTTEXT":      dwStyle = dwStyle Or BS_LEFTTEXT
         Case "BS_LEFT":          dwStyle = dwStyle Or BS_LEFT
         Case "BS_CENTER":        dwStyle = dwStyle Or BS_CENTER
         Case "BS_RIGHT":         dwStyle = dwStyle Or BS_RIGHT
         Case "BS_TOP":           dwStyle = dwStyle Or BS_TOP
         Case "BS_VCENTER":       dwStyle = dwStyle Or BS_VCENTER
         Case "BS_BOTTOM":        dwStyle = dwStyle Or BS_BOTTOM
         Case "BS_ICON":          dwStyle = dwStyle Or BS_ICON
         Case "BS_BITMAP":        dwStyle = dwStyle Or BS_BITMAP
         Case "BS_TEXT":          dwStyle = dwStyle Or BS_TEXT
         
         Case "ES_LEFT":          dwStyle = dwStyle Or ES_LEFT
         Case "ES_CENTER":        dwStyle = dwStyle Or ES_CENTER
         Case "ES_RIGHT":         dwStyle = dwStyle Or ES_RIGHT
         Case "ES_MULTILINE":     dwStyle = dwStyle Or ES_MULTILINE
         
         Case "CBS_DROPDOWNLIST":     dwStyle = dwStyle Or CBS_DROPDOWNLIST
         Case "CBS_NOINTEGRALHEIGHT": dwStyle = dwStyle Or CBS_NOINTEGRALHEIGHT

         Case "LBS_NOINTEGRALHEIGHT": dwStyle = dwStyle Or LBS_NOINTEGRALHEIGHT
         Case "LBS_STANDARD":         dwStyle = dwStyle Or LBS_STANDARD
         Case "LBS_HASSTRINGS":       dwStyle = dwStyle Or LBS_HASSTRINGS
         Case "LBS_NOTIFY":           dwStyle = dwStyle Or LBS_NOTIFY

         Case "PBS_SMOOTH":           dwStyle = dwStyle Or PBS_SMOOTH
         Case "PBS_VERTICAL":         dwStyle = dwStyle Or PBS_VERTICAL

         Case "UDS_ALIGNRIGHT":       dwStyle = dwStyle Or UDS_ALIGNRIGHT
         Case "UDS_ALIGNLEFT":        dwStyle = dwStyle Or UDS_ALIGNLEFT
         Case "UDS_HORZ":             dwStyle = dwStyle Or UDS_HORZ

         Case "TBS_HORZ":             dwStyle = dwStyle Or TBS_HORZ
         Case "TBS_VERT":             dwStyle = dwStyle Or TBS_VERT
         Case "TBS_TOP":              dwStyle = dwStyle Or TBS_TOP
         Case "TBS_BOTTOM":           dwStyle = dwStyle Or TBS_BOTTOM
         Case "TBS_RIGHT":            dwStyle = dwStyle Or TBS_RIGHT
         Case "TBS_LEFT":             dwStyle = dwStyle Or TBS_LEFT
         Case "TBS_BOTH":             dwStyle = dwStyle Or TBS_BOTH
         Case "TBS_AUTOTICKS":        dwStyle = dwStyle Or TBS_AUTOTICKS
         Case "TBS_NOTHUMB":          dwStyle = dwStyle Or TBS_NOTHUMB
          
         Case "TCS_SCROLLOPPOSITE":   dwStyle = dwStyle Or TCS_SCROLLOPPOSITE
         Case "TCS_BOTTOM":           dwStyle = dwStyle Or TCS_BOTTOM
         Case "TCS_RIGHT":            dwStyle = dwStyle Or TCS_RIGHT
         Case "TCS_FLATBUTTONS":      dwStyle = dwStyle Or TCS_FLATBUTTONS
         Case "TCS_FORCEICONLEFT":    dwStyle = dwStyle Or TCS_FORCEICONLEFT
         Case "TCS_VERTICAL":         dwStyle = dwStyle Or TCS_VERTICAL
         Case "TCS_FORCELABELLEFT":   dwStyle = dwStyle Or TCS_FORCELABELLEFT
         Case "TCS_TABS":             dwStyle = dwStyle Or TCS_TABS
         Case "TCS_BUTTONS":          dwStyle = dwStyle Or TCS_BUTTONS
         Case "TCS_SINGLELINE":       dwStyle = dwStyle Or TCS_SINGLELINE
         Case "TCS_MULTILINE":        dwStyle = dwStyle Or TCS_MULTILINE
         Case "TCS_RIGHTJUSTIFY":     dwStyle = dwStyle Or TCS_RIGHTJUSTIFY
         Case "TCS_FIXEDWIDTH":       dwStyle = dwStyle Or TCS_FIXEDWIDTH
         Case "TCS_RAGGEDRIGHT":      dwStyle = dwStyle Or TCS_RAGGEDRIGHT
         Case "TCS_MULTILINE":        dwStyle = dwStyle Or TCS_MULTILINE

         Case "DTS_UPDOWN":           dwStyle = dwStyle Or DTS_UPDOWN
         Case "DTS_SHOWNONE":         dwStyle = dwStyle Or DTS_SHOWNONE
         Case "DTS_SHORTDATEFORMAT":  dwStyle = dwStyle Or DTS_SHORTDATEFORMAT
         Case "DTS_LONGDATEFORMAT":   dwStyle = dwStyle Or DTS_LONGDATEFORMAT
         Case "DTS_LONGDATEFORMAT":   dwStyle = dwStyle Or DTS_LONGDATEFORMAT
         Case "DTS_SHORTDATECENTURYFORMAT": dwStyle = dwStyle Or DTS_SHORTDATECENTURYFORMAT 
         Case "DTS_SHORTDATECENTURYFORMAT": dwStyle = dwStyle Or DTS_SHORTDATECENTURYFORMAT 
         Case "DTS_SHORTDATECENTURYFORMAT": dwStyle = dwStyle Or DTS_SHORTDATECENTURYFORMAT 
         Case "DTS_TIMEFORMAT":       dwStyle = dwStyle Or DTS_TIMEFORMAT 
         Case "DTS_APPCANPARSE":      dwStyle = dwStyle Or DTS_APPCANPARSE
         Case "DTS_RIGHTALIGN":       dwStyle = dwStyle Or DTS_RIGHTALIGN

         Case "MCS_DAYSTATE":         dwStyle = dwStyle Or MCS_DAYSTATE
         Case "MCS_MULTISELECT":      dwStyle = dwStyle Or MCS_MULTISELECT
         Case "MCS_WEEKNUMBERS":      dwStyle = dwStyle Or MCS_WEEKNUMBERS
         Case "MCS_NOTODAYCIRCLE":    dwStyle = dwStyle Or MCS_NOTODAYCIRCLE
         Case "MCS_NOTODAY":          dwStyle = dwStyle Or MCS_NOTODAY

         Case "LVS_NOCOLUMNHEADER":   dwStyle = dwStyle Or LVS_NOCOLUMNHEADER    
         Case "LVS_REPORT":           dwStyle = dwStyle Or LVS_REPORT 
         Case "LVS_ICON":             dwStyle = dwStyle Or LVS_ICON             
         Case "LVS_SMALLICON":        dwStyle = dwStyle Or LVS_SMALLICON        
         Case "LVS_LIST":             dwStyle = dwStyle Or LVS_LIST             
         Case "LVS_SORTASCENDING":    dwStyle = dwStyle Or LVS_SORTASCENDING    
         Case "LVS_SORTDESCENDING":   dwStyle = dwStyle Or LVS_SORTDESCENDING   
         Case "LVS_NOLABELWRAP":      dwStyle = dwStyle Or LVS_NOLABELWRAP      
         Case "LVS_AUTOARRANGE":      dwStyle = dwStyle Or LVS_AUTOARRANGE      
         Case "LVS_ALIGNTOP":         dwStyle = dwStyle Or LVS_ALIGNTOP         
         Case "LVS_ALIGNLEFT":        dwStyle = dwStyle Or LVS_ALIGNLEFT        
          
         Case "TVS_LINESATROOT":      dwStyle = dwStyle Or TVS_LINESATROOT
         Case "TVS_HASBUTTONS":       dwStyle = dwStyle Or TVS_HASBUTTONS
         Case "TVS_HASLINES":         dwStyle = dwStyle Or TVS_HASLINES
         Case "TVS_CHECKBOXES":       dwStyle = dwStyle Or TVS_CHECKBOXES
         Case "TVS_NOSCROLL":         dwStyle = dwStyle Or TVS_NOSCROLL
         Case "TVS_NONEVENHEIGHT":    dwStyle = dwStyle Or TVS_NONEVENHEIGHT
         Case "TVS_NOHSCROLL":        dwStyle = dwStyle Or TVS_NOHSCROLL 
      END SELECT
   NEXT
   function = dwStyle
end function


' ========================================================================================
' Get the class name based on the type of control
' ========================================================================================
function GetControlClassName( byval pCtrl as clsControl ptr ) as CWSTR
   for i as long = lbound(gToolbox) to ubound(gToolbox)
      if gToolbox(i).nToolType = pCtrl->ControlType THEN 
         return gToolbox(i).wszClassName
      END IF
   NEXT
   return ""
end function
   

' ========================================================================================
' Generate a default name for the new form
' ========================================================================================
function GenerateDefaultFormName() as CWSTR

   ' Generate a default name for the form, then iterate the document collection
   ' to see if it already exists.
   dim pDoc as clsDocument ptr 
   dim pCtrl as clsControl ptr 
   dim as long NextFormNumber = 0
   dim as CWSTR wszDefaultName 
   dim as Boolean bOK 
   dim as long idx = gApp.GetActiveProjectIndex
   
   do 
      NextFormNumber = NextFormNumber + 1
      wszDefaultName = "Form" & NextFormNumber
      bOK = true
      For i As Long = 0 To gApp.Projects(idx).GetDocumentCount - 1
         pDoc = gApp.Projects(idx).GetDocumentPtr(i)
         if pDoc THEN
            for ii as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
               pCtrl = pDoc->Controls.ItemAt(ii)
               if pCtrl->ControlType = CTRL_FORM THEN
                  if ucase(GetControlProperty(pCtrl, "NAME")) = ucase(wszDefaultName) THEN
                     bOK = false
                  END IF
                  exit for
               END IF
            NEXT
         END IF
         if bOK = false THEN exit for
      Next
   loop until bOK = true
   
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
   
   for i as long = lbound(gToolbox) to ubound(gToolbox)
      if gToolbox(i).nToolType = nControlType THEN
         wszName = gToolbox(i).wszName
         exit for
      END IF
   NEXT

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
' Add a property to the incoming control and optionally a value
' ========================================================================================
function AddControlProperty( byval pCtrl as clsControl ptr, _
                             byref wszPropName as CWSTR, _
                             byref wszPropValue as CWSTR = "" _
                             ) as Long
   dim as long ub = ubound(pCtrl->Properties) + 1
   redim preserve pCtrl->Properties(ub) as clsProperty
   pCtrl->Properties(ub).wszPropName  = wszPropName
   pCtrl->Properties(ub).wszPropValue = wszPropValue
   function = 0
end function


' ========================================================================================
' Attach default properties for the incoming control
' ========================================================================================
function AttachDefaultControlProperties( byval pCtrl as clsControl ptr ) as Long

   ' Add the default same generic proprties that apply to all controls
   AddControlProperty(pCtrl, "Name")
   AddControlProperty(pCtrl, "Left")
   AddControlProperty(pCtrl, "Top")
   AddControlProperty(pCtrl, "Width")
   AddControlProperty(pCtrl, "Height")

   ' Add control specific properties
   select CASE pCtrl->ControlType
      case CTRL_LABEL
         AddControlProperty(pCtrl, "Text")
         AddControlProperty(pCtrl, "Style", "WS_VISIBLE|SS_LEFT|WS_GROUP|SS_NOTIFY")
         AddControlProperty(pCtrl, "ExStyle")
      case CTRL_BUTTON
         AddControlProperty(pCtrl, "Text")
         AddControlProperty(pCtrl, "Style", "WS_VISIBLE|WS_TABSTOP|BS_PUSHBUTTON|BS_CENTER|BS_VCENTER")
         AddControlProperty(pCtrl, "ExStyle")
      case CTRL_TEXTBOX
         AddControlProperty(pCtrl, "Text")
         AddControlProperty(pCtrl, "Style", "WS_VISIBLE|WS_TABSTOP|ES_LEFT|ES_AUTOHSCROLL")
         AddControlProperty(pCtrl, "ExStyle")
      case CTRL_CHECKBOX
         AddControlProperty(pCtrl, "Text")
         AddControlProperty(pCtrl, "Style", "WS_VISIBLE|WS_TABSTOP|BS_AUTOCHECKBOX|BS_LEFT|BS_VCENTER")
         AddControlProperty(pCtrl, "ExStyle")
      case CTRL_OPTION
         AddControlProperty(pCtrl, "Text")
         AddControlProperty(pCtrl, "Style", "WS_VISIBLE|WS_TABSTOP|BS_AUTORADIOBUTTON|BS_LEFT|BS_VCENTER")
         AddControlProperty(pCtrl, "ExStyle")
      case CTRL_FRAME
         AddControlProperty(pCtrl, "Text")
         AddControlProperty(pCtrl, "Style", "WS_VISIBLE|WS_CLIPSIBLINGS|WS_GROUP|BS_GROUPBOX")
         AddControlProperty(pCtrl, "ExStyle", "WS_EX_TRANSPARENT")
      case CTRL_PICTURE
      case CTRL_COMBOBOX
         AddControlProperty(pCtrl, "Style", "WS_VISIBLE|WS_BORDER|WS_TABSTOP|CBS_DROPDOWNLIST")
         AddControlProperty(pCtrl, "ExStyle")
      case CTRL_LISTBOX
         AddControlProperty(pCtrl, "Style", "WS_VISIBLE|WS_HSCROLL|WS_VSCROLL|WS_BORDER|WS_TABSTOP|LBS_STANDARD|LBS_HASSTRINGS|LBS_NOTIFY")
         AddControlProperty(pCtrl, "ExStyle")
      case CTRL_IMAGEBUTTON
      case CTRL_HSCROLL
      case CTRL_VSCROLL
      case CTRL_TIMER
      case CTRL_TABCONTROL
      case CTRL_RICHEDIT
      case CTRL_PROGRESSBAR
      case CTRL_UPDOWN
      case CTRL_LISTVIEW
      case CTRL_TREEVIEW
      case CTRL_SLIDER
      case CTRL_DATETIMEPICKER
      case CTRL_MONTHCALENDAR
      case CTRL_WEBBROWSER
      case CTRL_CUSTOM
      case CTRL_OCX
   END SELECT

   function = 0
end function


' ========================================================================================
' Apply properties to the incoming control
' ========================================================================================
function ApplyControlProperties( byval pCtrl as clsControl ptr ) as Long

   dim as long lb = lbound(pCtrl->Properties)
   dim as long ub = ubound(pCtrl->Properties)
   
   dim as long nPropValue, nTemp, nLeft, nTop, nWidth, nHeight
   dim as CWSTR wszPropValue
   
   for i as long = lb to ub
      ' Handle the generic properties that apply to all controls
      wszPropValue = pCtrl->Properties(i).wszPropValue
      nPropValue   = Val(pCtrl->Properties(i).wszPropValue)
      
      select CASE ucase(**pCtrl->Properties(i).wszPropName)
         case "NAME"
         CASE "LEFT"
            nLeft = nPropValue
            if pCtrl->SuspendLayout = false THEN
               nTemp = val(GetControlProperty(pCtrl, "TOP"))
               if pCtrl->ControlType = CTRL_FORM then
                  nLeft = 10: nTemp = 10
               end if
               SetWindowPos(pCtrl->hWindow, 0, nLeft, nTemp, 0, 0, SWP_NOZORDER or SWP_NOSIZE)
            end if   
         CASE "TOP"
            nTop = nPropValue
            if pCtrl->SuspendLayout = false THEN
               nTemp = val(GetControlProperty(pCtrl, "LEFT"))
               if pCtrl->ControlType = CTRL_FORM then
                  nTop = 10: nTemp = 10
               end if
               SetWindowPos(pCtrl->hWindow, 0, nTemp, nTop, 0, 0, SWP_NOZORDER or SWP_NOSIZE)
            end if   
         CASE "WIDTH"
            nWidth = nPropValue
            if pCtrl->SuspendLayout = false THEN
               nTemp = val(GetControlProperty(pCtrl, "HEIGHT"))
               SetWindowPos(pCtrl->hWindow, 0, 0, 0, nWidth, nTemp, SWP_NOZORDER or SWP_NOMOVE)
            end if   
         CASE "HEIGHT"      
            nHeight = nPropValue
            if pCtrl->SuspendLayout = false THEN
               nTemp = val(GetControlProperty(pCtrl, "WIDTH"))
               SetWindowPos(pCtrl->hWindow, 0, 0, 0, nTemp, nHeight, SWP_NOZORDER or SWP_NOMOVE)
            end if 
         case "TEXT"
            AfxSetWindowText(pCtrl->hWindow, wszPropValue)     
         CASE ELSE
      END SELECT
      
   NEXT

   ' Handle control specific properties
   for i as long = lb to ub
      select CASE pCtrl->ControlType
         case CTRL_LABEL
         case CTRL_BUTTON
         case CTRL_TEXTBOX      
            
         CASE ELSE
      END SELECT
   next
   
   ' If layout has been suspended then only move/resize the control after
   ' all of the relevant properties have now been set.
   if pCtrl->SuspendLayout THEN
      SetWindowPos(pCtrl->hWindow, 0, nLeft, nTop, nWidth, nHeight, SWP_NOZORDER)
   END IF
   
   function = 0

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
   dim as CWSTR wszDefaultControlName 
   

   pDoc->Controls.DeselectAllControls
   pCtrl->ControlType = ControlType
   pCtrl->IsActive    = false
   
   ' Set default properties for the created control
   AttachDefaultControlProperties( pCtrl )

   if ControlType = CTRL_FORM THEN
      dim pForm as CWindow ptr = New CWindow
      pForm->DPI = AfxCWindowPtr(pDoc->hWndFrame)->DPI
      wszDefaultControlName = GenerateDefaultFormName()
      pDoc->hWndForm = _
         pForm->Create( pDoc->hWndFrame, "Form1", @DesignerForm_WndProc, 0, 0, 0, 0, _
              WS_CHILD or WS_OVERLAPPED Or WS_OVERLAPPEDWINDOW or _
              WS_CLIPSIBLINGS Or WS_CLIPCHILDREN or WS_VISIBLE, _
              WS_EX_CONTROLPARENT OR WS_EX_WINDOWEDGE)
      pForm->ClassStyle = CS_DBLCLKS
      SetWindowLongPtr( pDoc->hWndForm, GWLP_ID, IDC_DESIGNFORM )
      ' We will set our own mouse pointer as needed
      SetClassLongPtr(pDoc->hWndForm, GCLP_HCURSOR, 0)   
      pCtrl->hWindow = pDoc->hWndForm
   else
      if pWindow = 0 THEN exit function
      wszDefaultControlName = GenerateDefaultControlName(pDoc, ControlType)
      pCtrl->hWindow = pWindow->AddControl( GetControlClassName(pCtrl), , 100, wszDefaultControlName, _
                       rcDraw.left, rcDraw.top, rcDraw.right-rcDraw.left, rcDraw.bottom-rcDraw.top, _
                       GetStyleValue(pCtrl, "STYLE"), GetStyleValue(pCtrl, "EXSTYLE"), , _
                       CAST(WNDPROC, @Control_SubclassProc), ControlType, CAST(DWORD_PTR, pDoc))
   END IF

   pDoc->Controls.Add(pCtrl)

   SetControlProperty(pCtrl, "NAME", wszDefaultControlName)
   SetControlProperty(pCtrl, "TEXT", wszDefaultControlName)
   SetControlProperty(pCtrl, "LEFT", str(rcDraw.left))
   SetControlProperty(pCtrl, "TOP", str(rcDraw.top))
   SetControlProperty(pCtrl, "WIDTH", str(rcDraw.right-rcDraw.left))
   SetControlProperty(pCtrl, "HEIGHT", str(rcDraw.bottom-rcDraw.top))
   
   ' Apply properties to the control
   pCtrl->SuspendLayout = true
   ApplyControlProperties( pCtrl )
   pCtrl->SuspendLayout = false

   pDoc->Controls.SelectControl(pCtrl->hWindow)
   pDoc->Controls.SetActiveControl(pCtrl->hWindow)

   function = pCtrl
END FUNCTION
