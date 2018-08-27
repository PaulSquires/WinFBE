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
' Display the Name and Description of the current selected Property.
' ========================================================================================
function DisplayPropertyDetails() as Long
   
   Dim As HWnd hList = GetDlgItem(HWND_FRMVDTOOLBOX, IDC_FRMVDTOOLBOX_LSTPROPERTIES)
   Dim As HWnd hPropName = GetDlgItem(HWND_FRMVDTOOLBOX, IDC_FRMVDTOOLBOX_LBLPROPNAME)
   Dim As HWnd hPropDescribe = GetDlgItem(HWND_FRMVDTOOLBOX, IDC_FRMVDTOOLBOX_LBLPROPDESCRIBE )

   dim pProp as clsProperty ptr = GetActivePropertyPtr
    
   ' If the listbox is not visible then do not set the labels. It could be that a different
   ' listbox (Events) is active and is using the labels to display other information.
   if IsWindowVisible(hList) = false then exit function
   
   dim as CWSTR wszPropName, wszPropDescribe
   if pProp then 
      wszPropName = pProp->wszPropName
      gwszPreviousPropName = pProp->wszPropName
   END IF
   
   select case ucase(wszPropName)
      CASE "NAME"
         wszPropDescribe = "Indicates the name used in code to identify the object."
      case "ACCEPTSRETURN"   
         wszPropDescribe = "Indicates if return characters are accepted as input for multiline edit controls."
      case "ACCEPTSTAB"   
         wszPropDescribe = "Indicates if tab characters are accepted as input for multiline edit controls."
      case "ACCEPTBUTTON"   
         wszPropDescribe = "If this is set, the button is 'clicked' whenever the user presses the 'ENTER' key."
      case "ALLOWDROP"   
         wszPropDescribe = "Indicates whether the control can accept data that the user drags onto it."
      case "BACKCOLOR"   
         wszPropDescribe = "The background color of the control. For Button controls the ThemeSupport property must be set to False."
      case "BACKCOLORDOWN"   
         wszPropDescribe = "The background color of the control when pressed. For Button controls the ThemeSupport property must be set to False."
      case "BACKCOLORHOT"   
         wszPropDescribe = "The background color of the control when the mouse is over the control. For Button controls the ThemeSupport property must be set to False."
      case "BACKGROUNDIMAGE"   
         wszPropDescribe = "The background image used for the control."
      case "BACKGROUNDIMAGELAYOUT"   
         wszPropDescribe = "The background image layout used for the control."
      case "BORDERSTYLE"   
         wszPropDescribe = "Indicates the appearance and behaviour of the border of the control."
      case "CANCELBUTTON"   
         wszPropDescribe = "If this is set, the button is 'clicked' whenever the user presses the 'ESC' key."
      case "CHARACTERCASING"   
         wszPropDescribe = "Indicates if all characters should be left alone or converted to uppercase or lowercase."
      case "CHECKALIGN"   
         wszPropDescribe = "Determines the location of the check box inside the control."
      case "CHECKED"   
         wszPropDescribe = "Indicates whether the control is checked."
      case "CHECKSTATE"   
         wszPropDescribe = "Indicates the state of the control."
      case "COLUMNWIDTH"   
         wszPropDescribe = "Indicates how wide each column should be in a multicolumn listbox."
      case "CONTROLBOX"   
         wszPropDescribe = "Determines whether a form has a Control/System menu box."
      case "ENABLED"   
         wszPropDescribe = "Indicates whether the control is enabled."
      case "FONT"   
         wszPropDescribe = "The font used to display text in the control."
      case "FORECOLOR"   
         wszPropDescribe = "The foreground color of the control which is used to display text."
      case "GROUPNAME"
         wszPropDescribe = "Controls with the same GroupName form part of the same OptionButton group."
      case "HIDESELECTION"   
         wszPropDescribe = "Indicates that the selection should be hidden when the edit control loses focus."
      case "HORIZONTALEXTENT"   
         wszPropDescribe = "The width in pixels by which a listbox can be scrolled horizontally. Only valid if HorizontalScrollBars is true."
      case "ICON"   
         wszPropDescribe = "The icon is displayed in the form's system menu box and when the form is minimized."
      case "IMAGE"   
         wszPropDescribe = "The image that will be displayed on the control."
      case "IMAGEALIGN"   
         wszPropDescribe = "The alignment of the image that will be displayed on the control."
'      case "INDEX"   
'         wszPropDescribe = "The position of the control within the OptionButton group."
      case "INTEGRALHEIGHT"   
         wszPropDescribe = "Indicates whether the list can contain only complete items."
      case "KEYPREVIEW"
         wszPropDescribe = "Determines whether the form will receive key events before the event is passed to the control that has focus."
      case "LOCKED"   
         wszPropDescribe = "Determines if the control can be moved or resized."
      case "MAXLENGTH"   
         wszPropDescribe = "Indicates the maximum number of characters that can be entered into the edit control."
      case "MAXIMIZEBOX"   
         wszPropDescribe = "Determines whether a form has a maximize box in the upper-right corner of its caption bar."
      case "MINIMIZEBOX"   
         wszPropDescribe = "Determines whether a form has a minimize box in the upper-right corner of its caption bar."
      case "MAXIMUMWIDTH"   
         wszPropDescribe = "The maximun width size the form can be resized to."
      case "MAXIMUMHEIGHT"   
         wszPropDescribe = "The maximun height size the form can be resized to."
      case "MINIMUMWIDTH"   
         wszPropDescribe = "The minimum width size the form can be resized to."
      case "MINIMUMHEIGHT"   
         wszPropDescribe = "The minimum height size the form can be resized to."
      case "MULTICOLUMN"   
         wszPropDescribe = "Indicates if values should be displayed in columns horizontally."
      case "MULTILINE"   
         wszPropDescribe = "Controls whether the text of the edit control can span more than one line."
      case "PASSWORDCHAR"   
         wszPropDescribe = "Indicates the character to display for password input for single-line edit control."
      case "READONLY"   
         wszPropDescribe = "Controls whether the text in the edit control can be changed or not."
      case "SCROLLALWAYSVISIBLE"   
         wszPropDescribe = "Indicates if the listbox should always have a scrollbar present regardless of how many items are in it."
      case "TEXTSCROLLBARS"   
         wszPropDescribe = "Indicates for multiline edit controls which scroll bars will be shown for the control."
      case "SELECTIONMODE"   
         wszPropDescribe = "Indicates if the listbox is to be simple-select, multiselect, or not selectable."
      case "SHOWINTASKBAR"   
         wszPropDescribe = "Determines whether the form appears in the Windows Taskbar."
      case "SORTED"   
         wszPropDescribe = "Controls whether the list is sorted."
      case "LEFT", "TOP"   
         wszPropDescribe = "The coordinates of the upper-left corner of the control relative to the upper-left corner of its container."
      case "WIDTH", "HEIGHT"   
         wszPropDescribe = "The size of the control in pixels."
      case "STARTPOSITION"   
         wszPropDescribe = "Determines the position of a form when it first appears."
      case "TABINDEX"   
         wszPropDescribe = "Determines the index in the TAB order that this control will occupy."
      case "TABSTOP"   
         wszPropDescribe = "Indicates whether the user can use the TAB key to give focus to the control."
      case "TAG"   
         wszPropDescribe = "User-defined text data associated with the control."
      case "TEXT"   
         wszPropDescribe = "The text associated with the control."
      case "TEXTALIGN"   
         wszPropDescribe = "The alignment of the text that will be displayed on the control."
      case "TEXTFORECOLOR"   
         wszPropDescribe = "The foreground color of the displayed control text. For Button controls the ThemeSupport property must be set to False."
      case "TEXTBACKCOLOR"   
         wszPropDescribe = "The background color of the displayed control text. For Button controls the ThemeSupport property must be set to False."
      case "TEXTFORECOLORDOWN"   
         wszPropDescribe = "The foreground color of the displayed control text when pressed. For Button controls the ThemeSupport property must be set to False."
      case "TEXTBACKCOLORDOWN"   
         wszPropDescribe = "The background color of the displayed control text when pressed. For Button controls the ThemeSupport property must be set to False."
      case "TEXTMARGIN"   
         wszPropDescribe = "The margin in pixels to apply to the text of a Button control."
      case "THEMESUPPORT"   
         wszPropDescribe = "If True, the Windows Theme is applied to the control. Set to False to specify a BackColor and/or ForeColor for the control."
      case "THREESTATE"
         wszPropDescribe = "Indicates whether the checkbox will allow three check states rather than two."
      case "TOGGLEMODE"
         wszPropDescribe = "Determines if the button allows dual states to toggle from ON (pressed) to OFF (unpressed)."
      case "USEMNEMONIC"   
         wszPropDescribe = "If True, the first character preceded by an ampersand will be used as the button's mnemonic key."
      case "USETABSTOPS"   
         wszPropDescribe = "Indicates if TAB character should be expanded into full spacing."
      CASE "WINDOWSTATE"
         wszPropDescribe = "Determines the initial visual state of the form."
      CASE "WORDWRAP"
         wszPropDescribe = "Indicates if lines are automatically word wrapped for multiline edit controls."
      CASE "VISIBLE"
         wszPropDescribe = "Determines whether the control is visible or hidden."
      case else
         wszPropDescribe = ""
   END SELECT
   
   AfxSetWindowText(hPropName, wszPropName)
   AfxSetWindowText(hPropDescribe, wszPropDescribe)
   
   function = 0
end function

            
' ========================================================================================
' Display the Name and Description of the current selected Event.
' ========================================================================================
function DisplayEventDetails() as Long
   
   Dim As HWnd hList = GetDlgItem(HWND_FRMVDTOOLBOX, IDC_FRMVDTOOLBOX_LSTEVENTS)
   Dim As HWnd hEventName = GetDlgItem(HWND_FRMVDTOOLBOX, IDC_FRMVDTOOLBOX_LBLPROPNAME)
   Dim As HWnd hEventDescribe = GetDlgItem(HWND_FRMVDTOOLBOX, IDC_FRMVDTOOLBOX_LBLPROPDESCRIBE )

   dim pEvent as clsEvent ptr = GetActiveEventPtr()

   ' If the listbox is not visible then do not set the labels. It could be that a different
   ' listbox (Properties) is active and is using the labels to display other information.
   if IsWindowVisible(hList) = false then exit function

   dim as CWSTR wszEventName, wszEventDescribe
   if pEvent then 
      wszEventName = pEvent->wszEventName
      gwszPreviousEventName = pEvent->wszEventName
   END IF
   
   select case ucase(wszEventName)
      case "ALLEVENTS"
         wszEventDescribe = "Allows the user to respond to all Windows messages for the control."
      case "LOAD"
         wszEventDescribe = "Occurs whenever the user loads the form. Form and control window handles are valid at this point."
      CASE "SHOWN"
         wszEventDescribe = "Occurs whenever the form is first shown."
      CASE "ACTIVATED"
         wszEventDescribe = "Occurs whenever the form is activated."
      CASE "DEACTIVATE"
         wszEventDescribe = "Occurs whenever the form is deactivated."
      CASE "DESTROY"
         wszEventDescribe = "Occurs immediately before the control is about to be destroyed and all resources associated with it released."
      CASE "GOTFOCUS"
         wszEventDescribe = "Occurs when the control receives focus."
      case "INITIALIZE"   
         wszEventDescribe = "Occurs before a form loads. Form and control window handles are invalid at this point."
      CASE "LOSTFOCUS"
         wszEventDescribe = "Occurs when the control loses focus."
      CASE "RESIZE"
         wszEventDescribe = "Occurs when a form is resized."
      CASE "FORMCLOSING"
         wszEventDescribe = "Occurs whenever the user closes the form, before the form has been closed."
      CASE "FORMCLOSED"
         wszEventDescribe = "Occurs whenever the user closes the form, after the form has been closed."
      CASE "MOVE"
         wszEventDescribe = "Occurs when a control is moved."
      CASE "CLICK"
         wszEventDescribe = "Occurs when a control is clicked."
      CASE "DROPFILES"
         wszEventDescribe = "Occurs whenever files are dropped on a control with the AllowDrop property enabled."
      CASE "MOUSEMOVE"
         wszEventDescribe = "Occurs when the mouse pointer is moved over the control."
      CASE "MOUSEDOWN"
         wszEventDescribe = "Occurs when the mouse pointer is over a control and a mouse button is pressed."
      CASE "MOUSEUP"
         wszEventDescribe = "Occurs when the mouse pointer is over a control and a mouse button is released."
      CASE "MOUSEDOUBLECLICK"
         wszEventDescribe = "Occurs when the control is double clicked by the mouse."
      CASE "MOUSEENTER"
         wszEventDescribe = "Occurs when the mouse pointer enters the control."
      CASE "MOUSEHOVER"
         wszEventDescribe = "Occurs when the mouse pointer rests on the control."
      CASE "MOUSELEAVE"
         wszEventDescribe = "Occurs when the mouse pointer leaves the control."
      case "TEXTCHANGED"   
         wszEventDescribe = "Occurs when the Text property is changed by either a programmatic modification or user interaction."
      CASE "KEYDOWN"
         wszEventDescribe = "Occurs when a key is first pressed."
      CASE "KEYPRESS"
         wszEventDescribe = "Occurs when a control has focus and the user presses and releases a key."
      CASE "KEYUP"
         wszEventDescribe = "Occurs when a key is released."
      case else
         wszEventDescribe = ""
   END SELECT
   
   AfxSetWindowText(hEventName, wszEventName)
   AfxSetWindowText(hEventDescribe, wszEventDescribe)
   
   function = 0
end function


' ========================================================================================
' Get the RGB (BGR) color value from the specified property value
' ========================================================================================
function GetRGBColorFromProperty( byval pProp as clsProperty ptr ) as COLORREF
   if pProp = 0 then exit function
   
   dim as CWSTR wszPropValue, wszList, wszValue
   wszPropValue = pProp->wszPropValue
   wszList = AfxStrParse(wszPropValue, 1, "|")
   wszValue = AfxStrParse(wszPropValue, 2, "|")

   dim as long nLookFor
   select case **wszList
      CASE "CUSTOM"
         return val(wszValue) 
      case "COLORS": nLookFor = COLOR_COLORS
      case "SYSTEM": nLookFor = COLOR_SYSTEM
   end select
   for i as long = lbound(gColors) to ubound(gColors)
      if gColors(i).ColorType = nLookFor then
         if gColors(i).wszColorName = wszValue then
            return gColors(i).ColorValue
         end if
      end if   
   NEXT
   function = 0
end function


' ========================================================================================
' Set values of a WinFormsX Font class (used for code generation)
' ========================================================================================
function SetFontClassFromPropValue( byref wszPropValue as wstring ) as CWSTR
   
   dim as CWSTR wszFont, wszStyles, wszCharSet
      
   if val(AfxStrParse(wszPropValue, 3, ",")) <> 400 then
      wszStyles = wszStyles & "FontStyles.Bold or "
   END IF
   if val(AfxStrParse(wszPropValue, 4, ",")) then
      wszStyles = wszStyles & "FontStyles.Italic or "
   END IF
   if val(AfxStrParse(wszPropValue, 5, ",")) then
      wszStyles = wszStyles & "FontStyles.Underline or "
   end if
   if val(AfxStrParse(wszPropValue, 6, ",")) then
      wszStyles = wszStyles & "FontStyles.StrikeOut or "
   end if
   wszStyles = rtrim(wszStyles, " or ")
   if len(wszStyles) = 0 then wszStyles = "FontStyles.Normal"

   select case val(AfxStrParse(wszPropValue, 7, ","))   ' charset
      case DEFAULT_CHARSET:     wszCharSet = "FontCharset.Default"
      case ANSI_CHARSET:        wszCharSet = "FontCharset.Ansi"        
      case ARABIC_CHARSET:      wszCharSet = "FontCharset.Arabic"      
      case BALTIC_CHARSET:      wszCharSet = "FontCharset.Baltic"      
      case CHINESEBIG5_CHARSET: wszCharSet = "FontCharset.ChineseBig5" 
      case EASTEUROPE_CHARSET:  wszCharSet = "FontCharset.EastEurope"  
      case GB2312_CHARSET:      wszCharSet = "FontCharset.GB2312"
      case GREEK_CHARSET:       wszCharSet = "FontCharset.Greek"      
      case HANGUL_CHARSET:      wszCharSet = "FontCharset.Hangul"      
      case HEBREW_CHARSET:      wszCharSet = "FontCharset.Hebrew"      
      case JOHAB_CHARSET:       wszCharSet = "FontCharset.Johab"       
      case MAC_CHARSET:         wszCharSet = "FontCharset.Mac"         
      case OEM_CHARSET:         wszCharSet = "FontCharset.OEM"         
      case RUSSIAN_CHARSET:     wszCharSet = "FontCharset.Russian"     
      case SHIFTJIS_CHARSET:    wszCharSet = "FontCharset.Shiftjis"    
      case SYMBOL_CHARSET:      wszCharSet = "FontCharset.Symbol"      
      case THAI_CHARSET:        wszCharSet = "FontCharset.Thai"        
      case TURKISH_CHARSET:     wszCharSet = "FontCharset.Turkish"     
      case VIETNAMESE_CHARSET:  wszCharSet = "FontCharset.Vietnamese"  
      case else: wszCharSet = "FontCharset.Default"
   end select
         
   wszFont = _
      chr(34) & AfxStrParse(wszPropValue, 1, ",") & chr(34) & "," & _
      AfxStrParse(wszPropValue, 2, ",") & "," & wszStyles & "," & wszCharSet

   function = wszFont

end function

' ========================================================================================
' Set values of a LOGFONT based on property value string
' ========================================================================================
function SetLogFontFromPropValue( byref wszPropValue as wstring ) as LOGFONT
   
   dim lf as LOGFONT
   lf.lfFaceName  = AfxStrParse(wszPropValue, 1, ",") 
   lf.lfHeight    = AfxGetFontHeight(val(AfxStrParse(wszPropValue, 2, ",")))
   lf.lfWeight    = val(AfxStrParse(wszPropValue, 3, ","))
   lf.lfItalic    = val(AfxStrParse(wszPropValue, 4, ","))
   lf.lfUnderline = val(AfxStrParse(wszPropValue, 5, ","))
   lf.lfStrikeOut = val(AfxStrParse(wszPropValue, 6, ","))
   lf.lfCharSet   = val(AfxStrParse(wszPropValue, 7, ","))
            
   function = lf
end function


' ========================================================================================
' Set values of a property value string based on LOGFONT 
' ========================================================================================
function SetPropValueFromLogFont( byref lf as LOGFONT ) as CWSTR
   
   dim as CWSTR wszPropValue = _
   lf.lfFaceName  & "," & _
   AfxGetFontPointSize(lf.lfHeight) & "," & _
   lf.lfWeight    & "," & _
   lf.lfItalic    & "," & _
   lf.lfUnderline & "," & _
   lf.lfStrikeOut & "," & _
   lf.lfCharSet   
            
   function = wszPropValue
end function


' ========================================================================================
' Show the popup font selection dialog for the current property
' ========================================================================================
function ChooseFontForProperty( byval pProp as clsProperty ptr ) as long
   if pProp = 0 then exit function
   
   dim cf as CHOOSEFONT 
   dim lf as LOGFONT
   cf.lStructSize = sizeof(cf)
   cf.hwndOwner = HWND_FRMVDTOOLBOX
   cf.lpLogFont = @lf
   cf.Flags = CF_SCREENFONTS or CF_EFFECTS or CF_INITTOLOGFONTSTRUCT

   lf = SetLogFontFromPropValue(pProp->wszPropValue)

   EnableWindow(HWND_FRMMAIN, false)
   if ChooseFont(@cf) then
      pProp->wszPropValue = SetPropValueFromLogFont(*cf.lpLogFont) 
      AfxRedrawWindow(GetDlgItem(HWND_FRMVDTOOLBOX, IDC_FRMVDTOOLBOX_LSTPROPERTIES))
   END IF
   EnableWindow(HWND_FRMMAIN, true)

   function = 0
end function

' ========================================================================================
' Create a default value for a FONT property of a control
' ========================================================================================
function CreateDefaultFontPropValue() as CWSTR
   
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(HWnd_FRMMAIN)
   if pWindow = 0 THEN exit function
   
   dim lf as LOGFONT
   GetObject(pWindow->Font, sizeof(lf), @lf)
   
   function = SetPropValueFromLogFont(lf)
end function


' ========================================================================================
' Determine if an Property exists for a Form
' ========================================================================================
function IsPropertyExists( byval pCtrl as clsControl ptr, _
                           byval wszPropName as CWSTR _
                           ) as boolean
   if pCtrl = 0 THEN return false
   dim as long lb = lbound(pCtrl->Properties)
   dim as long ub = ubound(pCtrl->Properties)
   dim as CWSTR wszPropName_ucase = ucase(wszPropName)
   
   for i as long = lb to ub
      if ucase(pCtrl->Properties(i).wszPropName) = wszPropName_ucase THEN
         return true
      END IF
   NEXT
   return false
end function


' ========================================================================================
' Determine if an Event exists for a Form
' ========================================================================================
function IsEventExists( byval pCtrl as clsControl ptr, _
                        byval wszEventName as CWSTR _
                        ) as boolean
   if pCtrl = 0 THEN return false
   dim as long lb = lbound(pCtrl->Events)
   dim as long ub = ubound(pCtrl->Events)
   dim as CWSTR wszEventName_ucase = ucase(wszEventName)
   
   for i as long = lb to ub
      if ucase(pCtrl->Events(i).wszEventName) = wszEventName_ucase THEN
         return true
      END IF
   NEXT
   return false
end function


' ========================================================================================
' Retrieve the Property pointer for the specified control property
' ========================================================================================
function GetControlPropertyPtr( byval pCtrl as clsControl ptr, _
                                byval wszPropName as CWSTR _
                                ) as clsProperty Ptr
   if pCtrl = 0 THEN return 0
   dim as long lb = lbound(pCtrl->Properties)
   dim as long ub = ubound(pCtrl->Properties)
   dim as CWSTR wszPropName_ucase = ucase(wszPropName)
   
   ' All property values are returned as strings. 
   for i as long = lb to ub
      if ucase(pCtrl->Properties(i).wszPropName) = wszPropName_ucase THEN
         return @pCtrl->Properties(i)
      END IF
   NEXT

   return 0
end function

' ========================================================================================
' Retrieve the value for the specified control property
' ========================================================================================
function GetControlProperty( byval pCtrl as clsControl ptr, _
                             byval wszPropName as CWSTR _
                             ) as CWSTR
   dim pProp as clsProperty ptr = GetControlPropertyPtr(pCtrl, wszPropName)
   if pProp then return pProp->wszPropValue 
   return 0
end function

' ========================================================================================
' Set the value for the specified control property
' ========================================================================================
function SetControlProperty( byval pCtrl as clsControl ptr, _
                             byval wszPropName as CWSTR, _
                             byval wszPropValue as CWSTR _
                             ) as long
   if pCtrl = 0 then exit function
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
' Add an event to the incoming control and optionally a value
' ========================================================================================
function AddControlEvent( byval pCtrl as clsControl ptr, _
                          byref wszEventName as CWSTR, _
                          byval bIsSelected as boolean = false _
                          ) as Long
   dim as long ub = ubound(pCtrl->Events) + 1
   redim preserve pCtrl->Events(ub) as clsEvent
   pCtrl->Events(ub).wszEventName = wszEventName
   pCtrl->Events(ub).bIsSelected = bIsSelected
   function = 0
end function


' ========================================================================================
' Set the value for the specified control event
' ========================================================================================
function SetControlEvent( byval pCtrl as clsControl ptr, _
                          byval wszEventName as CWSTR, _
                          byval bIsSelected as boolean _
                          ) as long
   if pCtrl = 0 then exit function
   dim as long lb = lbound(pCtrl->Events)
   dim as long ub = ubound(pCtrl->Events)
   dim as CWSTR wszEventName_ucase = ucase(wszEventName)
   
   for i as long = lb to ub
      if ucase(pCtrl->Events(i).wszEventName) = wszEventName_ucase THEN
         pCtrl->Events(i).bIsSelected = bIsSelected
         exit for
      END IF
   NEXT
   function = 0
end function

' ========================================================================================
' Attach default events for the incoming control
' ========================================================================================
function AttachDefaultControlEvents( byval pCtrl as clsControl ptr ) as Long

   ' Only add default events if the no events already exist for the control
   if ubound(pCtrl->Events) - lbound(pCtrl->Events) + 1 > 0 then exit function
  
   ' Add the default AllEvents event that applies to all controls
   AddControlEvent(pCtrl, "AllEvents")

   ' Add control specific properties
   select CASE pCtrl->ControlType
      case CTRL_FORM
         AddControlEvent(pCtrl, "Load")
         AddControlEvent(pCtrl, "Initialize")
         AddControlEvent(pCtrl, "Shown")
         AddControlEvent(pCtrl, "Activated")
         AddControlEvent(pCtrl, "Deactivate")
         AddControlEvent(pCtrl, "Resize")
         AddControlEvent(pCtrl, "FormClosing")
         AddControlEvent(pCtrl, "FormClosed")
         AddControlEvent(pCtrl, "Move")
         AddControlEvent(pCtrl, "Click")
         AddControlEvent(pCtrl, "DropFiles")
         AddControlEvent(pCtrl, "MouseMove")
         AddControlEvent(pCtrl, "MouseDown")
         AddControlEvent(pCtrl, "MouseUp")
         AddControlEvent(pCtrl, "MouseDoubleClick")
         AddControlEvent(pCtrl, "KeyDown")
         AddControlEvent(pCtrl, "KeyPress")
         AddControlEvent(pCtrl, "KeyUp")
      case CTRL_LABEL
         AddControlEvent(pCtrl, "Click")
         AddControlEvent(pCtrl, "Destroy")
         AddControlEvent(pCtrl, "DropFiles")
         AddControlEvent(pCtrl, "MouseMove")
         AddControlEvent(pCtrl, "MouseDown")
         AddControlEvent(pCtrl, "MouseUp")
         AddControlEvent(pCtrl, "MouseDoubleClick")
         AddControlEvent(pCtrl, "MouseEnter")
         AddControlEvent(pCtrl, "MouseHover")
         AddControlEvent(pCtrl, "MouseLeave")
      case CTRL_BUTTON
         AddControlEvent(pCtrl, "Click")
         AddControlEvent(pCtrl, "Destroy")
         AddControlEvent(pCtrl, "DropFiles")
         AddControlEvent(pCtrl, "GotFocus")
         AddControlEvent(pCtrl, "LostFocus")
         AddControlEvent(pCtrl, "MouseMove")
         AddControlEvent(pCtrl, "MouseDown")
         AddControlEvent(pCtrl, "MouseUp")
         AddControlEvent(pCtrl, "MouseDoubleClick")
         AddControlEvent(pCtrl, "MouseEnter")
         AddControlEvent(pCtrl, "MouseHover")
         AddControlEvent(pCtrl, "MouseLeave")
         AddControlEvent(pCtrl, "KeyDown")
         AddControlEvent(pCtrl, "KeyPress")
         AddControlEvent(pCtrl, "KeyUp")
      case CTRL_TEXTBOX
         AddControlEvent(pCtrl, "Click")
         AddControlEvent(pCtrl, "Destroy")
         AddControlEvent(pCtrl, "DropFiles")
         AddControlEvent(pCtrl, "GotFocus")
         AddControlEvent(pCtrl, "LostFocus")
         AddControlEvent(pCtrl, "MouseMove")
         AddControlEvent(pCtrl, "MouseDown")
         AddControlEvent(pCtrl, "MouseUp")
         AddControlEvent(pCtrl, "MouseDoubleClick")
         AddControlEvent(pCtrl, "MouseEnter")
         AddControlEvent(pCtrl, "MouseHover")
         AddControlEvent(pCtrl, "MouseLeave")
         AddControlEvent(pCtrl, "TextChanged")
         AddControlEvent(pCtrl, "KeyDown")
         AddControlEvent(pCtrl, "KeyPress")
         AddControlEvent(pCtrl, "KeyUp")
      case CTRL_CHECKBOX
         AddControlEvent(pCtrl, "Click")
         AddControlEvent(pCtrl, "Destroy")
         AddControlEvent(pCtrl, "DropFiles")
         AddControlEvent(pCtrl, "GotFocus")
         AddControlEvent(pCtrl, "LostFocus")
         AddControlEvent(pCtrl, "MouseMove")
         AddControlEvent(pCtrl, "MouseDown")
         AddControlEvent(pCtrl, "MouseUp")
         AddControlEvent(pCtrl, "MouseDoubleClick")
         AddControlEvent(pCtrl, "MouseEnter")
         AddControlEvent(pCtrl, "MouseHover")
         AddControlEvent(pCtrl, "MouseLeave")
         AddControlEvent(pCtrl, "KeyDown")
         AddControlEvent(pCtrl, "KeyPress")
         AddControlEvent(pCtrl, "KeyUp")
      case CTRL_LISTBOX
         AddControlEvent(pCtrl, "Click")
         AddControlEvent(pCtrl, "Destroy")
         AddControlEvent(pCtrl, "DropFiles")
         AddControlEvent(pCtrl, "GotFocus")
         AddControlEvent(pCtrl, "LostFocus")
         AddControlEvent(pCtrl, "MouseMove")
         AddControlEvent(pCtrl, "MouseDown")
         AddControlEvent(pCtrl, "MouseUp")
         AddControlEvent(pCtrl, "MouseDoubleClick")
         AddControlEvent(pCtrl, "MouseEnter")
         AddControlEvent(pCtrl, "MouseHover")
         AddControlEvent(pCtrl, "MouseLeave")
         AddControlEvent(pCtrl, "KeyDown")
         AddControlEvent(pCtrl, "KeyPress")
         AddControlEvent(pCtrl, "KeyUp")
      case CTRL_OPTION
         AddControlEvent(pCtrl, "Click")
         AddControlEvent(pCtrl, "Destroy")
         AddControlEvent(pCtrl, "DropFiles")
         AddControlEvent(pCtrl, "GotFocus")
         AddControlEvent(pCtrl, "LostFocus")
         AddControlEvent(pCtrl, "MouseMove")
         AddControlEvent(pCtrl, "MouseDown")
         AddControlEvent(pCtrl, "MouseUp")
         AddControlEvent(pCtrl, "MouseDoubleClick")
         AddControlEvent(pCtrl, "MouseEnter")
         AddControlEvent(pCtrl, "MouseHover")
         AddControlEvent(pCtrl, "MouseLeave")
         AddControlEvent(pCtrl, "KeyDown")
         AddControlEvent(pCtrl, "KeyPress")
         AddControlEvent(pCtrl, "KeyUp")
      case CTRL_FRAME
         AddControlEvent(pCtrl, "Click")
         AddControlEvent(pCtrl, "Destroy")
         AddControlEvent(pCtrl, "DropFiles")
         AddControlEvent(pCtrl, "MouseMove")
         AddControlEvent(pCtrl, "MouseDown")
         AddControlEvent(pCtrl, "MouseUp")
         AddControlEvent(pCtrl, "MouseDoubleClick")
         AddControlEvent(pCtrl, "MouseEnter")
         AddControlEvent(pCtrl, "MouseHover")
         AddControlEvent(pCtrl, "MouseLeave")
      case CTRL_PICTURE
      case CTRL_COMBOBOX
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
' Add a property to the incoming control and optionally a value
' ========================================================================================
function AddControlProperty( byval pCtrl as clsControl ptr, _
                             byref wszPropName as CWSTR, _
                             byref wszPropValue as CWSTR, _
                             byval nPropType as PropertyType _
                             ) as Long
   dim as long ub = ubound(pCtrl->Properties) + 1
   redim preserve pCtrl->Properties(ub) as clsProperty
   pCtrl->Properties(ub).wszPropName  = wszPropName
   pCtrl->Properties(ub).wszPropValue = wszPropValue
   pCtrl->Properties(ub).wszPropDefault = wszPropValue
   pCtrl->Properties(ub).PropType = nPropType
   function = 0
end function


' ========================================================================================
' Attach default properties for the incoming control
' ========================================================================================
function AttachDefaultControlProperties( byval pCtrl as clsControl ptr ) as Long

   ' Only add default properties if the no properties already exist for the control
   if ubound(pCtrl->Properties) - lbound(pCtrl->Properties) + 1 > 0 then exit function
  
   ' Add the default same generic proprties that apply to all controls
   AddControlProperty(pCtrl, "Name", "", PropertyType.EditEnter)
   AddControlProperty(pCtrl, "Left", "0", PropertyType.EditEnterNumericOnly)
   AddControlProperty(pCtrl, "Top", "0", PropertyType.EditEnterNumericOnly)
   AddControlProperty(pCtrl, "Width", "0", PropertyType.EditEnterNumericOnly)
   AddControlProperty(pCtrl, "Height", "0", PropertyType.EditEnterNumericOnly)

   ' Add control specific properties
   select CASE pCtrl->ControlType
      case CTRL_FORM
         AddControlProperty(pCtrl, "Text", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "WindowState", "FormWindowState.Normal", PropertyType.ComboPicker)
         AddControlProperty(pCtrl, "StartPosition", "FormStartPosition.Manual", PropertyType.ComboPicker)
         AddControlProperty(pCtrl, "BorderStyle", "FormBorderStyle.Sizable", PropertyType.ComboPicker)
         AddControlProperty(pCtrl, "MinimizeBox", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "MaximizeBox", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "ControlBox", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Enabled", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Visible", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "BackColor", "SYSTEM|Control", PropertyType.ColorPicker)
         AddControlProperty(pCtrl, "AcceptButton", "", PropertyType.ComboPicker)
         AddControlProperty(pCtrl, "AllowDrop", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "KeyPreview", "False", PropertyType.TrueFalse)
'         AddControlProperty(pCtrl, "BackgroundImage", "", PropertyType.ImagePicker)
'         AddControlProperty(pCtrl, "BackgroundImageLayout", "ImageLayout.None", PropertyType.ComboPicker)
         AddControlProperty(pCtrl, "CancelButton", "", PropertyType.ComboPicker)
'         AddControlProperty(pCtrl, "Icon", "", PropertyType.ImagePicker)
         AddControlProperty(pCtrl, "Locked", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "MaximumHeight", "0", PropertyType.EditEnterNumericOnly)
         AddControlProperty(pCtrl, "MaximumWidth", "0", PropertyType.EditEnterNumericOnly)
         AddControlProperty(pCtrl, "MinimumHeight", "0", PropertyType.EditEnterNumericOnly)
         AddControlProperty(pCtrl, "MinimumWidth", "0", PropertyType.EditEnterNumericOnly)
         AddControlProperty(pCtrl, "ShowInTaskBar", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Tag", "", PropertyType.EditEnter)

      case CTRL_LABEL
         AddControlProperty(pCtrl, "Text", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "AllowDrop", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "BackColor", "SYSTEM|Control", PropertyType.ColorPicker)
         AddControlProperty(pCtrl, "ForeColor", "SYSTEM|ControlText", PropertyType.ColorPicker)
         AddControlProperty(pCtrl, "BorderStyle", "ControlBorderStyle.None", PropertyType.ComboPicker)
         AddControlProperty(pCtrl, "TextAlign", "LabelAlignment.TopLeft", PropertyType.ComboPicker)
         AddControlProperty(pCtrl, "Font", CreateDefaultFontPropValue, PropertyType.FontPicker)
         AddControlProperty(pCtrl, "Enabled", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Tag", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "Locked", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "UseMnemonic", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Visible", "True", PropertyType.TrueFalse)
          
      case CTRL_BUTTON
         AddControlProperty(pCtrl, "BackColor", "SYSTEM|Control", PropertyType.ColorPicker)
         AddControlProperty(pCtrl, "BackColorDown", "SYSTEM|Control", PropertyType.ColorPicker)
         AddControlProperty(pCtrl, "BackColorHot", "SYSTEM|Control", PropertyType.ColorPicker)
         AddControlProperty(pCtrl, "AllowDrop", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Font", CreateDefaultFontPropValue, PropertyType.FontPicker)
         AddControlProperty(pCtrl, "TextForeColor", "SYSTEM|ControlText", PropertyType.ColorPicker)
         AddControlProperty(pCtrl, "TextBackColor", "SYSTEM|Control", PropertyType.ColorPicker)
         AddControlProperty(pCtrl, "TextForeColorDown", "SYSTEM|ControlText", PropertyType.ColorPicker)
         AddControlProperty(pCtrl, "TextBackColorDown", "SYSTEM|Control", PropertyType.ColorPicker)
'         AddControlProperty(pCtrl, "Image", "", PropertyType.ImagePicker)
'         AddControlProperty(pCtrl, "ImageAlign", "ImageLayout.None", PropertyType.ComboPicker)
         AddControlProperty(pCtrl, "Text", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "TextAlign", "ButtonAlignment.MiddleCenter", PropertyType.ComboPicker)
'         AddControlProperty(pCtrl, "TextMargin", "4", PropertyType.EditEnterNumericOnly)
         AddControlProperty(pCtrl, "UseMnemonic", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "ThemeSupport", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "ToggleMode", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Enabled", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "TabIndex", "0", PropertyType.EditEnterNumericOnly)
         AddControlProperty(pCtrl, "TabStop", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Locked", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Tag", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "Visible", "True", PropertyType.TrueFalse)

      case CTRL_TEXTBOX
         AddControlProperty(pCtrl, "AcceptsReturn", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "AcceptsTab", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "AllowDrop", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "BackColor", "SYSTEM|Window", PropertyType.ColorPicker)
         AddControlProperty(pCtrl, "BorderStyle", "ControlBorderStyle.Fixed3D", PropertyType.ComboPicker)
         AddControlProperty(pCtrl, "CharacterCasing", "CharacterCasing.Normal", PropertyType.ComboPicker)
         AddControlProperty(pCtrl, "Enabled", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Font", CreateDefaultFontPropValue, PropertyType.FontPicker)
         AddControlProperty(pCtrl, "ForeColor", "SYSTEM|WindowText", PropertyType.ColorPicker)
         AddControlProperty(pCtrl, "HideSelection", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Locked", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "MaxLength", "32767", PropertyType.EditEnterNumericOnly)
         AddControlProperty(pCtrl, "Multiline", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "PasswordChar", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "ReadOnly", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "TextScrollBars", "ScrollBars.None", PropertyType.ComboPicker)
         AddControlProperty(pCtrl, "TabIndex", "0", PropertyType.EditEnterNumericOnly)
         AddControlProperty(pCtrl, "TabStop", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Tag", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "Text", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "TextAlign", "TextAlignment.Left", PropertyType.ComboPicker)
         AddControlProperty(pCtrl, "Visible", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "WordWrap", "False", PropertyType.TrueFalse)

      case CTRL_CHECKBOX
         AddControlProperty(pCtrl, "AllowDrop", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "BackColor", "SYSTEM|Control", PropertyType.ColorPicker)
         'AddControlProperty(pCtrl, "CheckAlign", "ButtonAlignment.MiddleCenter", PropertyType.ComboPicker)
         AddControlProperty(pCtrl, "CheckState", "CheckBoxState.Unchecked", PropertyType.ComboPicker)
         AddControlProperty(pCtrl, "Enabled", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Font", CreateDefaultFontPropValue, PropertyType.FontPicker)
         'AddControlProperty(pCtrl, "ForeColor", "SYSTEM|ControlText", PropertyType.ColorPicker)
         AddControlProperty(pCtrl, "Locked", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "TabIndex", "0", PropertyType.EditEnterNumericOnly)
         AddControlProperty(pCtrl, "TabStop", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Tag", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "Text", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "TextAlign", "ButtonAlignment.MiddleLeft", PropertyType.ComboPicker)
         AddControlProperty(pCtrl, "ThreeState", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "UseMnemonic", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Visible", "True", PropertyType.TrueFalse)

      case CTRL_LISTBOX
         AddControlProperty(pCtrl, "AllowDrop", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "BackColor", "SYSTEM|Window", PropertyType.ColorPicker)
         AddControlProperty(pCtrl, "BorderStyle", "ControlBorderStyle.Fixed3D", PropertyType.ComboPicker)
         AddControlProperty(pCtrl, "ColumnWidth", "0", PropertyType.EditEnterNumericOnly)
         AddControlProperty(pCtrl, "Enabled", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Font", CreateDefaultFontPropValue, PropertyType.FontPicker)
         AddControlProperty(pCtrl, "ForeColor", "SYSTEM|WindowText", PropertyType.ColorPicker)
         AddControlProperty(pCtrl, "HorizontalExtent", "0", PropertyType.EditEnterNumericOnly)
         AddControlProperty(pCtrl, "IntegralHeight", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Locked", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "MultiColumn", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "ScrollAlwaysVisible", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "SelectionMode", "ListSelectionMode.One", PropertyType.ComboPicker)
         AddControlProperty(pCtrl, "Sorted", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "TabIndex", "0", PropertyType.EditEnterNumericOnly)
         AddControlProperty(pCtrl, "TabStop", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Tag", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "UseTabStops", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Visible", "True", PropertyType.TrueFalse)

      case CTRL_OPTION
         AddControlProperty(pCtrl, "AllowDrop", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "BackColor", "SYSTEM|Control", PropertyType.ColorPicker)
         'AddControlProperty(pCtrl, "CheckAlign", "ButtonAlignment.MiddleCenter", PropertyType.ComboPicker)
         AddControlProperty(pCtrl, "Checked", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Enabled", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Font", CreateDefaultFontPropValue, PropertyType.FontPicker)
         AddControlProperty(pCtrl, "GroupName", "OptionGroup1", PropertyType.EditEnter)
         'AddControlProperty(pCtrl, "Index", "0", PropertyType.EditEnterNumericOnly)
         AddControlProperty(pCtrl, "Locked", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "TabIndex", "0", PropertyType.EditEnterNumericOnly)
         AddControlProperty(pCtrl, "TabStop", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Tag", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "Text", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "TextAlign", "ButtonAlignment.MiddleLeft", PropertyType.ComboPicker)
         AddControlProperty(pCtrl, "UseMnemonic", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Visible", "True", PropertyType.TrueFalse)

      case CTRL_FRAME
         AddControlProperty(pCtrl, "Text", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "AllowDrop", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "BackColor", "SYSTEM|Control", PropertyType.ColorPicker)
'         AddControlProperty(pCtrl, "ForeColor", "SYSTEM|ControlText", PropertyType.ColorPicker)
'         AddControlProperty(pCtrl, "BorderStyle", "ControlBorderStyle.None", PropertyType.ComboPicker)
'         AddControlProperty(pCtrl, "TextAlign", "LabelAlignment.TopLeft", PropertyType.ComboPicker)
         AddControlProperty(pCtrl, "Font", CreateDefaultFontPropValue, PropertyType.FontPicker)
         AddControlProperty(pCtrl, "Enabled", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Tag", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "Locked", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "UseMnemonic", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Visible", "True", PropertyType.TrueFalse)

'      case CTRL_PICTURE
'      case CTRL_COMBOBOX
'         AddControlProperty(pCtrl, "TabIndex", "0")
'         AddControlProperty(pCtrl, "TabStop", "True", PropertyType.TrueFalse)
'         AddControlProperty(pCtrl, "Style", "WS_VISIBLE|WS_BORDER|WS_TABSTOP|CBS_DROPDOWNLIST")
'         AddControlProperty(pCtrl, "ExStyle", "WS_EX_CLIENTEDGE")
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
' Load the dropdown PropertyList listbox that shows chocies for the current property
' ========================================================================================
function LoadPropertyComboListbox( byval pDoc as clsDocument ptr, _
                                   byval pCtrl as clsControl ptr, _
                                   byval pProp as clsProperty ptr _
                                   ) as Long

   if pDoc = 0 then exit function
   if pCtrl = 0 then exit function
   if pProp = 0 then exit function
   
   ListBox_ResetContent(HWND_PROPLIST_COMBOLIST)
   
   if pProp->PropType = PropertyType.TrueFalse then
      ListBox_AddString(HWND_PROPLIST_COMBOLIST, @WSTR("True"))
      ListBox_AddString(HWND_PROPLIST_COMBOLIST, @WSTR("False"))
      exit function
   end if

   select case ucase(pProp->wszPropName)
      case "ACCEPTBUTTON", "CANCELBUTTON"
         ListBox_AddString(HWND_PROPLIST_COMBOLIST, @wstr("(none)"))
         for i as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
            pCtrl = pDoc->Controls.ItemAt(i)
            if pCtrl then
               if pCtrl->ControlType = CTRL_BUTTON then
                  ListBox_AddString(HWND_PROPLIST_COMBOLIST, GetControlProperty(pCtrl, "NAME"))
               end if
            END if
         NEXT
         
      CASE "BORDERSTYLE"
         select case pCtrl->ControlType
            case CTRL_FORM
               ListBox_AddString(HWND_PROPLIST_COMBOLIST, @wstr("FormBorderStyle.None"))
               ListBox_AddString(HWND_PROPLIST_COMBOLIST, @wstr("FormBorderStyle.Sizable"))
               ListBox_AddString(HWND_PROPLIST_COMBOLIST, @wstr("FormBorderStyle.Fixed3D"))
               ListBox_AddString(HWND_PROPLIST_COMBOLIST, @wstr("FormBorderStyle.FixedSingle"))
               ListBox_AddString(HWND_PROPLIST_COMBOLIST, @wstr("FormBorderStyle.FixedDialog"))
               ListBox_AddString(HWND_PROPLIST_COMBOLIST, @wstr("FormBorderStyle.FixedToolWindow"))
               ListBox_AddString(HWND_PROPLIST_COMBOLIST, @wstr("FormBorderStyle.SizableToolWindow"))
            case else
               ListBox_AddString(HWND_PROPLIST_COMBOLIST, @wstr("ControlBorderStyle.None"))
               ListBox_AddString(HWND_PROPLIST_COMBOLIST, @wstr("ControlBorderStyle.FixedSingle"))
               ListBox_AddString(HWND_PROPLIST_COMBOLIST, @wstr("ControlBorderStyle.Fixed3D"))
         end select
         
      case "CHARACTERCASING"
         ListBox_AddString(HWND_PROPLIST_COMBOLIST, @wstr("CharacterCase.Normal"))
         ListBox_AddString(HWND_PROPLIST_COMBOLIST, @wstr("CharacterCase.Upper"))
         ListBox_AddString(HWND_PROPLIST_COMBOLIST, @wstr("CharacterCase.Lower"))

      case "CHECKSTATE"
         ListBox_AddString(HWND_PROPLIST_COMBOLIST, @wstr("CheckBoxState.Checked"))
         ListBox_AddString(HWND_PROPLIST_COMBOLIST, @wstr("CheckBoxState.Unchecked"))
         ListBox_AddString(HWND_PROPLIST_COMBOLIST, @wstr("CheckBoxState.Indeterminate"))

      case "TEXTSCROLLBARS"
         ListBox_AddString(HWND_PROPLIST_COMBOLIST, @wstr("ScrollBars.None"))
         ListBox_AddString(HWND_PROPLIST_COMBOLIST, @wstr("ScrollBars.Horizontal"))
         ListBox_AddString(HWND_PROPLIST_COMBOLIST, @wstr("ScrollBars.Vertical"))
         ListBox_AddString(HWND_PROPLIST_COMBOLIST, @wstr("ScrollBars.Both"))

      CASE "STARTPOSITION"
         ListBox_AddString(HWND_PROPLIST_COMBOLIST, @wstr("FormStartPosition.CenterParent"))
         ListBox_AddString(HWND_PROPLIST_COMBOLIST, @wstr("FormStartPosition.CenterScreen"))
         ListBox_AddString(HWND_PROPLIST_COMBOLIST, @wstr("FormStartPosition.Manual"))
         ListBox_AddString(HWND_PROPLIST_COMBOLIST, @wstr("FormStartPosition.WindowsDefaultLocation"))

      CASE "SELECTIONMODE"
         ListBox_AddString(HWND_PROPLIST_COMBOLIST, @wstr("ListSelectionMode.None"))
         ListBox_AddString(HWND_PROPLIST_COMBOLIST, @wstr("ListSelectionMode.One"))
         ListBox_AddString(HWND_PROPLIST_COMBOLIST, @wstr("ListSelectionMode.MultiSimple"))
         ListBox_AddString(HWND_PROPLIST_COMBOLIST, @wstr("ListSelectionMode.MultiExtended"))

      case "WINDOWSTATE"
         ListBox_AddString(HWND_PROPLIST_COMBOLIST, @wstr("FormWindowState.Maximized"))
         ListBox_AddString(HWND_PROPLIST_COMBOLIST, @wstr("FormWindowState.Minimized"))
         ListBox_AddString(HWND_PROPLIST_COMBOLIST, @wstr("FormWindowState.Normal"))

      case "BACKGROUNDIMAGELAYOUT"
         ListBox_AddString(HWND_PROPLIST_COMBOLIST, @wstr("ImageLayout.None"))
         ListBox_AddString(HWND_PROPLIST_COMBOLIST, @wstr("ImageLayout.Tile"))
         ListBox_AddString(HWND_PROPLIST_COMBOLIST, @wstr("ImageLayout.Center"))
         ListBox_AddString(HWND_PROPLIST_COMBOLIST, @wstr("ImageLayout.Stretch"))
         ListBox_AddString(HWND_PROPLIST_COMBOLIST, @wstr("ImageLayout.Zoom"))

      case "TEXTALIGN"
         select case pCtrl->ControlType
            case CTRL_BUTTON, CTRL_CHECKBOX, CTRL_OPTION, CTRL_FRAME
               ListBox_AddString(HWND_PROPLIST_COMBOLIST, @wstr("ButtonAlignment.BottomCenter"))
               ListBox_AddString(HWND_PROPLIST_COMBOLIST, @wstr("ButtonAlignment.BottomLeft"))
               ListBox_AddString(HWND_PROPLIST_COMBOLIST, @wstr("ButtonAlignment.BottomRight"))
               ListBox_AddString(HWND_PROPLIST_COMBOLIST, @wstr("ButtonAlignment.MiddleCenter"))
               ListBox_AddString(HWND_PROPLIST_COMBOLIST, @wstr("ButtonAlignment.MiddleLeft"))
               ListBox_AddString(HWND_PROPLIST_COMBOLIST, @wstr("ButtonAlignment.MiddleRight"))
               ListBox_AddString(HWND_PROPLIST_COMBOLIST, @wstr("ButtonAlignment.TopCenter"))
               ListBox_AddString(HWND_PROPLIST_COMBOLIST, @wstr("ButtonAlignment.TopLeft"))
               ListBox_AddString(HWND_PROPLIST_COMBOLIST, @wstr("ButtonAlignment.TopRight"))
            case CTRL_LABEL
               ListBox_AddString(HWND_PROPLIST_COMBOLIST, @wstr("LabelAlignment.MiddleCenter"))
               ListBox_AddString(HWND_PROPLIST_COMBOLIST, @wstr("LabelAlignment.MiddleLeft"))
               ListBox_AddString(HWND_PROPLIST_COMBOLIST, @wstr("LabelAlignment.MiddleRight"))
               ListBox_AddString(HWND_PROPLIST_COMBOLIST, @wstr("LabelAlignment.TopCenter"))
               ListBox_AddString(HWND_PROPLIST_COMBOLIST, @wstr("LabelAlignment.TopLeft"))
               ListBox_AddString(HWND_PROPLIST_COMBOLIST, @wstr("LabelAlignment.TopRight"))
            case CTRL_TEXTBOX
               ListBox_AddString(HWND_PROPLIST_COMBOLIST, @wstr("TextAlignment.Left"))
               ListBox_AddString(HWND_PROPLIST_COMBOLIST, @wstr("TextAlignment.Right"))
               ListBox_AddString(HWND_PROPLIST_COMBOLIST, @wstr("TextAlignment.Center"))
         end select
   END SELECT

   function = 0
end function


