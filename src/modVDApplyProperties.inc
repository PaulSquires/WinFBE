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
' Apply properties to the form 
' ========================================================================================
function Form_ApplyProperties( byval pDoc as clsDocument ptr, _
                               byval pCtrl as clsControl ptr, _
                               byval pProp as clsProperty ptr _ 
                               ) as Long
      
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(pDoc->hWndForm)

   dim as long nPropValue
   dim as CWSTR wszPropName, wszPropValue
   
   wszPropName  = pProp->wszPropName
   wszPropValue = pProp->wszPropValue
   nPropValue   = Val(pProp->wszPropValue)
   
   select CASE ucase(wszPropName)

   case "BACKCOLOR"
      pWindow->Brush = CreateSolidBrush( GetRGBColorFromProperty(pProp) )
      AfxRedrawWindow(pCtrl->hWindow)
   
   case "CONTROLBOX"
      if wszPropValue = "True" then
         AfxAddWindowStyle(pCtrl->hWindow, WS_SYSMENU)
      else
         AfxRemoveWindowStyle(pCtrl->hWindow, WS_SYSMENU)
      end if
      SetWindowPos( pCtrl->hWindow, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_DRAWFRAME)
      
   case "MAXIMIZEBOX"
      if wszPropValue = "True" then
         AfxAddWindowStyle(pCtrl->hWindow, WS_MAXIMIZEBOX)
      else
         AfxRemoveWindowStyle(pCtrl->hWindow, WS_MAXIMIZEBOX)
      end if
      SetWindowPos( pCtrl->hWindow, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_DRAWFRAME)

   case "MINIMIZEBOX"
      if wszPropValue = "True" then
         AfxAddWindowStyle(pCtrl->hWindow, WS_MINIMIZEBOX)
      else
         AfxRemoveWindowStyle(pCtrl->hWindow, WS_MINIMIZEBOX)
      end if
      SetWindowPos( pCtrl->hWindow, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_DRAWFRAME)
      
   case "BORDERSTYLE"
      ' Remove all existing styles before adding the new style
      AfxRemoveWindowStyle(pCtrl->hWindow, WS_BORDER)
      AfxRemoveWindowStyle(pCtrl->hWindow, WS_THICKFRAME)
      AfxRemoveWindowExStyle(pCtrl->hWindow, WS_EX_TOOLWINDOW)
      AfxRemoveWindowExStyle(pCtrl->hWindow, WS_EX_CLIENTEDGE)
      AfxRemoveWindowExStyle(pCtrl->hWindow, WS_EX_WINDOWEDGE)
      AfxRemoveWindowExStyle(pCtrl->hWindow, WS_EX_DLGMODALFRAME)
      
      select case **wszPropValue
         case "FormBorderStyle.None"
            AfxRemoveWindowStyle(pCtrl->hWindow, WS_CAPTION)
         case "FormBorderStyle.SizableToolWindow"
            AfxAddWindowStyle(pCtrl->hWindow, WS_BORDER)
            AfxAddWindowStyle(pCtrl->hWindow, WS_THICKFRAME)
            AfxAddWindowExStyle(pCtrl->hWindow, WS_EX_TOOLWINDOW)
         case "FormBorderStyle.FixedToolWindow"
            AfxAddWindowStyle(pCtrl->hWindow, WS_BORDER)
            AfxAddWindowStyle(pCtrl->hWindow, WS_DLGFRAME)
            AfxAddWindowExStyle(pCtrl->hWindow, WS_EX_TOOLWINDOW)
         case "FormBorderStyle.Sizable"
            AfxAddWindowStyle(pCtrl->hWindow, WS_THICKFRAME)
            AfxAddWindowStyle(pCtrl->hWindow, WS_DLGFRAME)
            AfxAddWindowStyle(pCtrl->hWindow, WS_BORDER)
            AfxAddWindowExStyle(pCtrl->hWindow, WS_EX_WINDOWEDGE)
         case "FormBorderStyle.Fixed3D"
            AfxAddWindowStyle(pCtrl->hWindow, WS_DLGFRAME)
            AfxAddWindowStyle(pCtrl->hWindow, WS_BORDER)
            AfxAddWindowExStyle(pCtrl->hWindow, WS_EX_WINDOWEDGE)
            AfxAddWindowExStyle(pCtrl->hWindow, WS_EX_CLIENTEDGE)
         case "FormBorderStyle.FixedSingle"
            AfxAddWindowStyle(pCtrl->hWindow, WS_DLGFRAME)
            AfxAddWindowStyle(pCtrl->hWindow, WS_BORDER)
            AfxAddWindowExStyle(pCtrl->hWindow, WS_EX_WINDOWEDGE)
         case "FormBorderStyle.FixedDialog"
            AfxAddWindowStyle(pCtrl->hWindow, WS_DLGFRAME)
            AfxAddWindowStyle(pCtrl->hWindow, WS_BORDER)
            AfxAddWindowExStyle(pCtrl->hWindow, WS_EX_DLGMODALFRAME)
      end select     
      SetWindowPos( pCtrl->hWindow, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_DRAWFRAME)

   end select
   
   function = 0
end function


' ========================================================================================
' Apply properties to the button
' ========================================================================================
function Button_ApplyProperties( byval pDoc as clsDocument ptr, _
                                 byval pCtrl as clsControl ptr, _
                                 byval pProp as clsProperty ptr _ 
                                 ) as Long
   
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(pDoc->hWndForm)

   dim as long nPropValue
   dim as CWSTR wszPropName, wszPropValue
   
   wszPropName  = pProp->wszPropName
   wszPropValue = pProp->wszPropValue
   nPropValue   = Val(pProp->wszPropValue)
   
   select CASE ucase(wszPropName)
   
   case "THEMESUPPORT"
      if wszPropValue = "True" then
         pCtrl->AfxButtonPtr->EnableTheming
      else
         pCtrl->AfxButtonPtr->DisableTheming
      end if
      
   case "BACKCOLOR"
      pCtrl->AfxButtonPtr->SetButtonBkColor( GetRGBColorFromProperty(pProp), true )

   case "TEXTFORECOLOR"
      pCtrl->AfxButtonPtr->SetTextForeColor( GetRGBColorFromProperty(pProp), true )

   case "TEXTBACKCOLOR"
      pCtrl->AfxButtonPtr->SetTextBkColor( GetRGBColorFromProperty(pProp), true )

   case "TEXTMARGIN"
      pCtrl->AfxButtonPtr->SetTextMargin( nPropValue, true )

   case "TEXTALIGN"
      dim as long wsStyle
      select case **wszPropValue
         CASE "ButtonAlignment.BottomCenter": wsStyle = DT_CENTER OR DT_BOTTOM
         CASE "ButtonAlignment.BottomLeft":   wsStyle = DT_LEFT   OR DT_BOTTOM
         CASE "ButtonAlignment.BottomRight":  wsStyle = DT_RIGHT  OR DT_BOTTOM
         CASE "ButtonAlignment.MiddleCenter": wsStyle = DT_CENTER OR DT_VCENTER
         CASE "ButtonAlignment.MiddleLeft":   wsStyle = DT_LEFT   OR DT_VCENTER
         CASE "ButtonAlignment.MiddleRight":  wsStyle = DT_RIGHT  OR DT_VCENTER
         CASE "ButtonAlignment.TopCenter":    wsStyle = DT_CENTER OR DT_TOP
         CASE "ButtonAlignment.TopLeft":      wsStyle = DT_LEFT   OR DT_TOP
         CASE "ButtonAlignment.TopRight":     wsStyle = DT_RIGHT  OR DT_TOP
      END SELECT
      pCtrl->AfxButtonPtr->SetTextFormat(wsStyle OR DT_SINGLELINE, true)

   end select
   
   function = 0
end function


' ========================================================================================
' Apply properties to the label
' ========================================================================================
function Label_ApplyProperties( byval pDoc as clsDocument ptr, _
                                byval pCtrl as clsControl ptr, _
                                byval pProp as clsProperty ptr _ 
                                ) as Long
    
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(pDoc->hWndForm)

   dim as long nPropValue
   dim as CWSTR wszPropName, wszPropValue
   
   wszPropName  = pProp->wszPropName
   wszPropValue = pProp->wszPropValue
   nPropValue   = Val(pProp->wszPropValue)
   
   select CASE ucase(wszPropName)
      
   case "BACKCOLOR"
      dim as COLORREF clrBack = GetRGBColorFromProperty(pProp)
      if pCtrl->hBackBrush then DeleteBrush(pCtrl->hBackBrush)
      pCtrl->hBackBrush = CreateSolidBrush(clrBack)
      AfxRedrawWindow(pCtrl->hWindow)

   case "BORDERSTYLE"
      AfxRemoveWindowExStyle(pCtrl->hWindow, WS_EX_CLIENTEDGE)
      AfxRemoveWindowStyle(pCtrl->hWindow, WS_BORDER)
      dim as long wsStyle
      select case **wszPropValue
         case "ControlBorderStyle.None"
         case "ControlBorderStyle.Fixed3D"
            AfxAddWindowExStyle(pCtrl->hWindow, WS_EX_CLIENTEDGE)
         case "ControlBorderStyle.FixedSingle"
            AfxAddWindowStyle(pCtrl->hWindow, WS_BORDER)
      end select
      SetWindowPos( pCtrl->hWindow, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_DRAWFRAME)

   case "TEXTALIGN"
      AfxRemoveWindowStyle(pCtrl->hWindow, SS_CENTER)
      AfxRemoveWindowStyle(pCtrl->hWindow, SS_LEFT)
      AfxRemoveWindowStyle(pCtrl->hWindow, SS_RIGHT)
      AfxRemoveWindowStyle(pCtrl->hWindow, SS_CENTERIMAGE)
      dim as long wsStyle
      select case **wszPropValue
         CASE "LabelAlignment.MiddleCenter": wsStyle = SS_CENTER OR SS_CENTERIMAGE
         CASE "LabelAlignment.MiddleLeft":   wsStyle = SS_LEFT   OR SS_CENTERIMAGE
         CASE "LabelAlignment.MiddleRight":  wsStyle = SS_RIGHT  OR SS_CENTERIMAGE
         CASE "LabelAlignment.TopCenter":    wsStyle = SS_CENTER 
         CASE "LabelAlignment.TopLeft":      wsStyle = SS_LEFT   
         CASE "LabelAlignment.TopRight":     wsStyle = SS_RIGHT  
      END SELECT
      AfxAddWindowStyle(pCtrl->hWindow, wsStyle)
      SetWindowPos( pCtrl->hWindow, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_DRAWFRAME)

   end select
   
   function = 0
end function


' ========================================================================================
' Apply properties to the checkbox
' ========================================================================================
function CheckBox_ApplyProperties( byval pDoc as clsDocument ptr, _
                                   byval pCtrl as clsControl ptr, _
                                   byval pProp as clsProperty ptr _ 
                                   ) as Long
   
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(pDoc->hWndForm)

   dim as long nPropValue
   dim as CWSTR wszPropName, wszPropValue
   
   wszPropName  = pProp->wszPropName
   wszPropValue = pProp->wszPropValue
   nPropValue   = Val(pProp->wszPropValue)
   
   select CASE ucase(wszPropName)
      
   case "BACKCOLOR"
      dim as COLORREF clrBack = GetRGBColorFromProperty(pProp)
      if pCtrl->hBackBrush then DeleteBrush(pCtrl->hBackBrush)
      pCtrl->hBackBrush = CreateSolidBrush(clrBack)
      AfxRedrawWindow(pCtrl->hWindow)

   case "CHECKSTATE"
      dim as long wsState
      select case **wszPropValue
         Case "CheckBoxState.Checked": wsState = BST_CHECKED
         Case "CheckBoxState.UnChecked": wsState = BST_UNCHECKED
         Case "CheckBoxState.Indeterminate": wsState = BST_INDETERMINATE
      End Select   
      SendMessage( pCtrl->hWindow, BM_SETCHECK, wsState, 0)
      AfxRedrawWindow(pCtrl->hWindow)
      
   case "THREESTATE"
      AfxRemoveWindowStyle(pCtrl->hWindow, BS_AUTOCHECKBOX)
      AfxRemoveWindowStyle(pCtrl->hWindow, BS_AUTO3STATE)
      if wszPropValue = "True" then
         AfxAddWindowStyle(pCtrl->hWindow, BS_AUTO3STATE)
      else
         AfxAddWindowStyle(pCtrl->hWindow, BS_AUTOCHECKBOX)
      end if
         
   case "TEXTALIGN"
      AfxRemoveWindowStyle(pCtrl->hWindow, BS_CENTER)
      AfxRemoveWindowStyle(pCtrl->hWindow, BS_LEFT)
      AfxRemoveWindowStyle(pCtrl->hWindow, BS_RIGHT)
      AfxRemoveWindowStyle(pCtrl->hWindow, BS_TOP)
      AfxRemoveWindowStyle(pCtrl->hWindow, BS_BOTTOM)
      AfxRemoveWindowStyle(pCtrl->hWindow, BS_VCENTER)
      dim as long wsStyle
      select case **wszPropValue
         CASE "ButtonAlignment.BottomCenter": wsStyle = BS_CENTER OR BS_BOTTOM
         CASE "ButtonAlignment.BottomLeft":   wsStyle = BS_LEFT   OR BS_BOTTOM
         CASE "ButtonAlignment.BottomRight":  wsStyle = BS_RIGHT  OR BS_BOTTOM
         CASE "ButtonAlignment.MiddleCenter": wsStyle = BS_CENTER OR BS_VCENTER
         CASE "ButtonAlignment.MiddleLeft":   wsStyle = BS_LEFT   OR BS_VCENTER
         CASE "ButtonAlignment.MiddleRight":  wsStyle = BS_RIGHT  OR BS_VCENTER
         CASE "ButtonAlignment.TopCenter":    wsStyle = BS_CENTER OR BS_TOP
         CASE "ButtonAlignment.TopLeft":      wsStyle = BS_LEFT   OR BS_TOP
         CASE "ButtonAlignment.TopRight":     wsStyle = BS_RIGHT  OR BS_TOP
      END SELECT
      AfxAddWindowStyle(pCtrl->hWindow, wsStyle)
      SetWindowPos( pCtrl->hWindow, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_DRAWFRAME)

   end select
   
   function = 0
end function

                       
' ========================================================================================
' Apply properties to the optionbutton
' ========================================================================================
function OptionButton_ApplyProperties( byval pDoc as clsDocument ptr, _
                                       byval pCtrl as clsControl ptr, _
                                       byval pProp as clsProperty ptr _ 
                                       ) as Long
   
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(pDoc->hWndForm)

   dim as long nPropValue
   dim as CWSTR wszPropName, wszPropValue
   
   wszPropName  = pProp->wszPropName
   wszPropValue = pProp->wszPropValue
   nPropValue   = Val(pProp->wszPropValue)
   
   select CASE ucase(wszPropName)
      
   case "BACKCOLOR"
      dim as COLORREF clrBack = GetRGBColorFromProperty(pProp)
      if pCtrl->hBackBrush then DeleteBrush(pCtrl->hBackBrush)
      pCtrl->hBackBrush = CreateSolidBrush(clrBack)
      AfxRedrawWindow(pCtrl->hWindow)

   case "CHECKED"
      ' If Checked is True then cycle through all of the other OptionButtons in
      ' the same group and toggle their Checked property to False.
      if wszPropValue = "True" then
         dim pCtrl2 as clsControl ptr
         dim as CWSTR wszOptionGroup = ucase(GetControlProperty(pCtrl, "GROUPNAME")) 
         for i as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
            pCtrl2 = pDoc->Controls.ItemAt(i)
            if pCtrl2->ControlType <> CTRL_OPTION then continue for
            if pCtrl2 = pCtrl then continue for
            if ucase(GetControlProperty(pCtrl2, "GROUPNAME")) = wszOptionGroup THEN
               SetControlProperty(pCtrl2, "CHECKED", "False") 
               SendMessage( pCtrl2->hWindow, BM_SETCHECK, BST_UNCHECKED, 0)
               AfxRedrawWindow(pCtrl2->hWindow)
            end if
         NEXT
         SendMessage( pCtrl->hWindow, BM_SETCHECK, BST_CHECKED, 0)
         AfxRedrawWindow(pCtrl->hWindow)
      end if
      
   case "TEXTALIGN"
      AfxRemoveWindowStyle(pCtrl->hWindow, BS_CENTER)
      AfxRemoveWindowStyle(pCtrl->hWindow, BS_LEFT)
      AfxRemoveWindowStyle(pCtrl->hWindow, BS_RIGHT)
      AfxRemoveWindowStyle(pCtrl->hWindow, BS_TOP)
      AfxRemoveWindowStyle(pCtrl->hWindow, BS_BOTTOM)
      AfxRemoveWindowStyle(pCtrl->hWindow, BS_VCENTER)
      dim as long wsStyle
      select case **wszPropValue
         CASE "ButtonAlignment.BottomCenter": wsStyle = BS_CENTER OR BS_BOTTOM
         CASE "ButtonAlignment.BottomLeft":   wsStyle = BS_LEFT   OR BS_BOTTOM
         CASE "ButtonAlignment.BottomRight":  wsStyle = BS_RIGHT  OR BS_BOTTOM
         CASE "ButtonAlignment.MiddleCenter": wsStyle = BS_CENTER OR BS_VCENTER
         CASE "ButtonAlignment.MiddleLeft":   wsStyle = BS_LEFT   OR BS_VCENTER
         CASE "ButtonAlignment.MiddleRight":  wsStyle = BS_RIGHT  OR BS_VCENTER
         CASE "ButtonAlignment.TopCenter":    wsStyle = BS_CENTER OR BS_TOP
         CASE "ButtonAlignment.TopLeft":      wsStyle = BS_LEFT   OR BS_TOP
         CASE "ButtonAlignment.TopRight":     wsStyle = BS_RIGHT  OR BS_TOP
      END SELECT
      AfxAddWindowStyle(pCtrl->hWindow, wsStyle)
      SetWindowPos( pCtrl->hWindow, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_DRAWFRAME)

   end select
   
   function = 0
end function

' ========================================================================================
' Apply properties to the listbox
' ========================================================================================
function ListBox_ApplyProperties( byval pDoc as clsDocument ptr, _
                                  byval pCtrl as clsControl ptr, _
                                  byval pProp as clsProperty ptr _ 
                                  ) as Long
   
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(pDoc->hWndForm)
   if pWindow = 0 then exit function
   
   dim as long nPropValue
   dim as CWSTR wszPropName, wszPropValue
   
   wszPropName  = pProp->wszPropName
   wszPropValue = pProp->wszPropValue
   nPropValue   = Val(pProp->wszPropValue)
   
   select CASE ucase(wszPropName)
      
   case "BACKCOLOR"
      dim as COLORREF clrBack = GetRGBColorFromProperty(pProp)
      if pCtrl->hBackBrush then DeleteBrush(pCtrl->hBackBrush)
      pCtrl->hBackBrush = CreateSolidBrush(clrBack)
      AfxRedrawWindow(pCtrl->hWindow)

   case "BORDERSTYLE"
      AfxRemoveWindowExStyle(pCtrl->hWindow, WS_EX_CLIENTEDGE)
      AfxRemoveWindowStyle(pCtrl->hWindow, WS_BORDER)
      dim as long wsStyle                                     
      select case **wszPropValue
         case "ControlBorderStyle.None"
         case "ControlBorderStyle.Fixed3D"
            AfxAddWindowExStyle(pCtrl->hWindow, WS_EX_CLIENTEDGE)
         case "ControlBorderStyle.FixedSingle"
            AfxAddWindowStyle(pCtrl->hWindow, WS_BORDER)
      end select
      SetWindowPos( pCtrl->hWindow, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_DRAWFRAME)

   case "INTEGRALHEIGHT"
' Do not do anything with the IntegralHeight property. Let the user size at whatever
' size he wishes during runtime. If property is set to True then it will resize
' appropriately at runtime. It appears that even though you can dynamically set the
' NOINTEGRALHEIGHT flag, the control itself doesn't seem to respond to it correctly
' once the original listview has been cfreated.

'      ' Get the current dimensions of the control based on Properties
'      dim as RECT rc = GetControlRect(pCtrl)
'      if wszPropValue = "True" then
'         AfxRemoveWindowStyle(pCtrl->hWindow, LBS_NOINTEGRALHEIGHT)
'         pWindow->SetWindowPos( pCtrl->hWindow, 0, rc.left, rc.top, _
'                    rc.right-rc.left, rc.bottom-rc.top, SWP_DRAWFRAME)
'      else
'         AfxAddWindowStyle(pCtrl->hWindow, LBS_NOINTEGRALHEIGHT)
'      END IF
'      AfxRedrawWindow(pCtrl->hWindow)
'      ' Get the new height because the Integral setting can change it based on how many
'      ' row items can fit into the visible area.
'      rc.bottom = rc.top + pWindow->UnScaleY(AfxGetWindowHeight(pCtrl->hWindow)) 
'      SetControlProperty(pCtrl, "HEIGHT", str(rc.bottom-rc.top))
'      DisplayPropertyList(pDoc)

   end select
   
   function = 0
end function

' ========================================================================================
' Apply properties to the textbox
' ========================================================================================
function TextBox_ApplyProperties( byval pDoc as clsDocument ptr, _
                                  byval pCtrl as clsControl ptr, _
                                  byval pProp as clsProperty ptr _ 
                                  ) as Long
   
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(pDoc->hWndForm)

   dim as long nPropValue
   dim as CWSTR wszPropName, wszPropValue
   
   wszPropName  = pProp->wszPropName
   wszPropValue = pProp->wszPropValue
   nPropValue   = Val(pProp->wszPropValue)
   
   select CASE ucase(wszPropName)
      
   case "BACKCOLOR"
      dim as COLORREF clrBack = GetRGBColorFromProperty(pProp)
      if pCtrl->hBackBrush then DeleteBrush(pCtrl->hBackBrush)
      pCtrl->hBackBrush = CreateSolidBrush(clrBack)

   case "BORDERSTYLE"
      AfxRemoveWindowExStyle(pCtrl->hWindow, WS_EX_CLIENTEDGE)
      AfxRemoveWindowStyle(pCtrl->hWindow, WS_BORDER)
      select case **wszPropValue
         case "ControlBorderStyle.None"
         case "ControlBorderStyle.Fixed3D"
            AfxAddWindowExStyle(pCtrl->hWindow, WS_EX_CLIENTEDGE)
         case "ControlBorderStyle.FixedSingle"
            AfxAddWindowStyle(pCtrl->hWindow, WS_BORDER)
      end select
      SetWindowPos( pCtrl->hWindow, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_DRAWFRAME)

   case "CHARACTERCASING"
      AfxRemoveWindowStyle(pCtrl->hWindow, ES_LOWERCASE)
      AfxRemoveWindowStyle(pCtrl->hWindow, ES_UPPERCASE)
      select case **wszPropValue
         Case "CharacterCase.Normal"
         Case "CharacterCase.Upper"
            AfxAddWindowStyle(pCtrl->hWindow, ES_UPPERCASE)
         Case "CharacterCase.Lower"
            AfxAddWindowStyle(pCtrl->hWindow, ES_LOWERCASE)
      End Select
      AfxRedrawWindow(pCtrl->hWindow)

   CASE "MULTILINE"
      if wszPropValue = "True" then
         AfxAddWindowStyle(pCtrl->hWindow, ES_MULTILINE)
      else   
         AfxRemoveWindowStyle(pCtrl->hWindow, ES_MULTILINE)
      END IF
      AfxRedrawWindow(pCtrl->hWindow)

   case "PASSWORDCHAR"
      Dim As Long iChr = 0
      If Len(wszPropValue) Then iChr = wszPropValue[0]
      SendMessage(pCtrl->hWindow, EM_SETPASSWORDCHAR, iChr, 0)
      AfxRedrawWindow(pCtrl->hWindow)
      
   case "TEXTSCROLLBARS"
      AfxRemoveWindowStyle(pCtrl->hWindow, WS_HSCROLL)
      AfxRemoveWindowStyle(pCtrl->hWindow, WS_VSCROLL)
      dim as long wsStyle
      select case **wszPropValue
         case "ScrollBars.None"
         case "ScrollBars.Horizontal"
            wsStyle = WS_HSCROLL
         case "ScrollBars.Vertical"
            wsStyle = WS_VSCROLL
         case "ScrollBars.Both"
            wsStyle = (WS_HSCROLL or WS_VSCROLL)
      end select
      AfxAddWindowStyle(pCtrl->hWindow, wsStyle)
      SetWindowPos( pCtrl->hWindow, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_DRAWFRAME)

   case "TEXTALIGN"
      AfxRemoveWindowStyle(pCtrl->hWindow, ES_CENTER)
      AfxRemoveWindowStyle(pCtrl->hWindow, ES_LEFT)
      AfxRemoveWindowStyle(pCtrl->hWindow, ES_RIGHT)
      dim as long wsStyle
      select case **wszPropValue
         CASE "TextAlignment.Left":   wsStyle = ES_LEFT
         CASE "TextAlignment.Center": wsStyle = ES_CENTER 
         CASE "TextAlignment.Right":  wsStyle = ES_RIGHT  
      END SELECT
      AfxAddWindowStyle(pCtrl->hWindow, wsStyle)
      SetWindowPos( pCtrl->hWindow, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_DRAWFRAME)

   end select
   
   function = 0
end function


' ========================================================================================
' Apply properties to the frame
' ========================================================================================
function Frame_ApplyProperties( byval pDoc as clsDocument ptr, _
                                byval pCtrl as clsControl ptr, _
                                byval pProp as clsProperty ptr _ 
                                ) as Long
    
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(pDoc->hWndForm)

   dim as long nPropValue
   dim as CWSTR wszPropName, wszPropValue
   
   wszPropName  = pProp->wszPropName
   wszPropValue = pProp->wszPropValue
   nPropValue   = Val(pProp->wszPropValue)
   
   select CASE ucase(wszPropName)
      
   case "BACKCOLOR"
      dim as COLORREF clrBack = GetRGBColorFromProperty(pProp)
      if pCtrl->hBackBrush then DeleteBrush(pCtrl->hBackBrush)
      pCtrl->hBackBrush = CreateSolidBrush(clrBack)
      AfxRedrawWindow(pCtrl->hWindow)

'   case "BORDERSTYLE"
'      AfxRemoveWindowExStyle(pCtrl->hWindow, WS_EX_CLIENTEDGE)
'      AfxRemoveWindowStyle(pCtrl->hWindow, WS_BORDER)
'      dim as long wsStyle
'      select case **wszPropValue
'         case "ControlBorderStyle.None"
'         case "ControlBorderStyle.Fixed3D"
'            AfxAddWindowExStyle(pCtrl->hWindow, WS_EX_CLIENTEDGE)
'         case "ControlBorderStyle.FixedSingle"
'            AfxAddWindowStyle(pCtrl->hWindow, WS_BORDER)
'      end select
'      SetWindowPos( pCtrl->hWindow, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_DRAWFRAME)

'   case "TEXTALIGN"
'      AfxRemoveWindowStyle(pCtrl->hWindow, SS_CENTER)
'      AfxRemoveWindowStyle(pCtrl->hWindow, SS_LEFT)
'      AfxRemoveWindowStyle(pCtrl->hWindow, SS_RIGHT)
'      AfxRemoveWindowStyle(pCtrl->hWindow, SS_CENTERIMAGE)
'      dim as long wsStyle
'      select case **wszPropValue
'         CASE "LabelAlignment.MiddleCenter": wsStyle = SS_CENTER OR SS_CENTERIMAGE
'         CASE "LabelAlignment.MiddleLeft":   wsStyle = SS_LEFT   OR SS_CENTERIMAGE
'         CASE "LabelAlignment.MiddleRight":  wsStyle = SS_RIGHT  OR SS_CENTERIMAGE
'         CASE "LabelAlignment.TopCenter":    wsStyle = SS_CENTER 
'         CASE "LabelAlignment.TopLeft":      wsStyle = SS_LEFT   
'         CASE "LabelAlignment.TopRight":     wsStyle = SS_RIGHT  
'      END SELECT
'      AfxAddWindowStyle(pCtrl->hWindow, wsStyle)
'      SetWindowPos( pCtrl->hWindow, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_DRAWFRAME)

   end select
   
   function = 0
end function


' ========================================================================================
' Apply properties to the incoming control
' ========================================================================================
function ApplyControlProperties( byval pDoc as clsDocument ptr, _
                                 byval pCtrl as clsControl ptr _
                                 ) as long

   if pDoc = 0 then exit function
   if pCtrl = 0 then exit function
   
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(pDoc->hWndForm)
   dim pProp as clsProperty ptr
   
   dim as long lb = lbound(pCtrl->Properties)
   dim as long ub = ubound(pCtrl->Properties)
   
   dim as long nPropValue, nTemp, nLeft, nTop, nWidth, nHeight
   dim as CWSTR wszPropName, wszPropValue
   
   ' Loop through all properties and apply them
   for i as long = lb to ub
      pProp = @pCtrl->Properties(i)
      if pProp = 0 then continue for
      
      wszPropName  = pProp->wszPropName
      wszPropValue = pProp->wszPropValue
      nPropValue   = Val(pProp->wszPropValue)
      
      select CASE ucase(wszPropName)
         case "NAME"
            if pCtrl->ControlType = CTRL_LISTBOX then
               ListBox_ResetContent(pCtrl->hWindow)
               ListBox_AddString(pCtrl->hWindow, wszPropValue)
            END IF
         CASE "LEFT"
            nLeft = nPropValue
            if pCtrl->SuspendLayout = false THEN
               nTemp = val(GetControlProperty(pCtrl, "TOP"))
               if pCtrl->ControlType = CTRL_FORM then
                  nLeft = 10: nTemp = 10
               end if
               pWindow->SetWindowPos(pCtrl->hWindow, 0, nLeft, nTemp, 0, 0, SWP_NOZORDER or SWP_NOSIZE)
            end if   
         CASE "TOP"
            nTop = nPropValue
            if pCtrl->SuspendLayout = false THEN
               nTemp = val(GetControlProperty(pCtrl, "LEFT"))
               if pCtrl->ControlType = CTRL_FORM then
                  nTop = 10: nTemp = 10
               end if
               pWindow->SetWindowPos(pCtrl->hWindow, 0, nTemp, nTop, 0, 0, SWP_NOZORDER or SWP_NOSIZE)
            end if   
         CASE "WIDTH"
            nWidth = nPropValue
            if pCtrl->SuspendLayout = false THEN
               nTemp = val(GetControlProperty(pCtrl, "HEIGHT"))
               pWindow->SetWindowPos(pCtrl->hWindow, 0, 0, 0, nWidth, nTemp, SWP_NOZORDER or SWP_NOMOVE)
            end if   
         CASE "HEIGHT"      
            nHeight = nPropValue
            if pCtrl->SuspendLayout = false THEN
               nTemp = val(GetControlProperty(pCtrl, "WIDTH"))
               pWindow->SetWindowPos(pCtrl->hWindow, 0, 0, 0, nTemp, nHeight, SWP_NOZORDER or SWP_NOMOVE)
            end if 
         case "TEXT"
            AfxSetWindowText(pCtrl->hWindow, wszPropValue)     
         case "FONT"
            dim as HFONT hFontOld = AfxGetWindowFont(pCtrl->hWindow)
            dim as LOGFONT lf
            lf = SetLogFontFromPropValue(wszPropValue)
            dim as HFONT hFont = CreateFontIndirect(@lf)
            AfxSetWindowFont(pCtrl->hWindow, hFont, true)
            DeleteFont(hFontOld)
      END SELECT
      
   
   ' Handle control specific properties
   select CASE pCtrl->ControlType
      case CTRL_FORM:     Form_ApplyProperties(pDoc, pCtrl, pProp)
      case CTRL_BUTTON:   Button_ApplyProperties(pDoc, pCtrl, pProp)
      case CTRL_LABEL:    Label_ApplyProperties(pDoc, pCtrl, pProp)
      case CTRL_CHECKBOX: CheckBox_ApplyProperties(pDoc, pCtrl, pProp)      
      case CTRL_OPTION:   OptionButton_ApplyProperties(pDoc, pCtrl, pProp)      
      case CTRL_TEXTBOX:  TextBox_ApplyProperties(pDoc, pCtrl, pProp)      
      case CTRL_LISTBOX:  ListBox_ApplyProperties(pDoc, pCtrl, pProp)      
      case CTRL_FRAME:    Frame_ApplyProperties(pDoc, pCtrl, pProp)
      CASE ELSE
   END SELECT

   next

   ' If layout has been suspended then only move/resize the control after
   ' all of the relevant properties have now been set.
   if pCtrl->SuspendLayout THEN
      pWindow->SetWindowPos( pCtrl->hWindow, 0, nLeft, nTop, nWidth, nHeight, SWP_NOZORDER)
   END IF
   
   ' Repaint to ensure that the grab handles draw
   AfxRedrawWindow(pDoc->hWndFrame)
   AfxRedrawWindow(pDoc->hWndForm)
   
   function = 0

end function

