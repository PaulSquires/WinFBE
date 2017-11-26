'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2017 Paul Squires, PlanetSquires Software
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
   dim as long lb = lbound(pCtrl->Properties)
   dim as long ub = ubound(pCtrl->Properties)
   dim as CWSTR wszPropName_ucase = ucase(wszPropName)
   
   ' All property values are returned as strings. 
   for i as long = lb to ub
      if pCtrl->Properties(i).wszPropName = wszPropName_ucase THEN
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
      if pCtrl->Properties(i).wszPropName = wszPropName_ucase THEN
         pCtrl->Properties(i).wszPropValue = wszPropValue
         exit for
      END IF
   NEXT
   function = 0
end function


' ========================================================================================
' Generate a default name for the new control
' ========================================================================================
function GenerateDefaultControlName( byval pDoc as clsDocument ptr ) as CWSTR

   ' Generate a default name for the control, then iterate the control collection
   ' to see if it already exists.
   dim as long NextControlNumber = 0
   dim as CWSTR wszDefaultName 
   dim as Boolean bOK
   dim pCtrl as clsControl ptr 
   
   do until bOK = true
      NextControlNumber = NextControlNumber + 1
      wszDefaultName = GetActiveToolboxControlName & NextControlNumber
      for i as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
         pCtrl = pDoc->Controls.ItemAt(i)
         if ucase(GetControlProperty(pCtrl, "NAME")) <> ucase(wszDefaultName) THEN
            bOK = true: exit for
         END IF
      NEXT
   loop
   
   return wszDefaultName
end function


' ========================================================================================
' Add a property to the incoming control
' ========================================================================================
function AddControlProperty( byval pCtrl as clsControl ptr, _
                             byref wszPropDisplayName as CWSTR _
                             ) as Long
   dim as long ub = ubound(pCtrl->Properties) + 1
   redim preserve pCtrl->Properties(ub) as clsProperty
   pCtrl->Properties(ub).wszPropDisplayName = wszPropDisplayName
   pCtrl->Properties(ub).wszPropName        = ucase(wszPropDisplayName)
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
         
      CASE ELSE
   END SELECT

   function = 0
end function


' ========================================================================================
' Apply properties to the incoming control
' ========================================================================================
function ApplyControlProperties( byval pCtrl as clsControl ptr ) as Long

   dim as long lb = lbound(pCtrl->Properties)
   dim as long ub = ubound(pCtrl->Properties)
   
   dim as long nPropValue, nLeft, nTop, nWidth, nHeight
   dim as CWSTR wszPropValue
   
   for i as long = lb to ub
      ' Handle the generic properties that apply to all controls
      wszPropValue = pCtrl->Properties(i).wszPropValue
      nPropValue   = Val(pCtrl->Properties(i).wszPropValue)
      
      select CASE **pCtrl->Properties(i).wszPropName
         case "NAME"
         CASE "LEFT"
            nLeft = nPropValue
            if pCtrl->SuspendLayoutChange = false THEN
            end if   
         CASE "TOP"
            nTop = nPropValue
            if pCtrl->SuspendLayoutChange = false THEN
            end if   
         CASE "WIDTH"
            nWidth = nPropValue
            if pCtrl->SuspendLayoutChange = false THEN
            end if   
         CASE "HEIGHT"      
            nHeight = nPropValue
            if pCtrl->SuspendLayoutChange = false THEN
            end if   
         CASE ELSE
      END SELECT
      
   NEXT

   ' Handle control specific properties
   for i as long = lb to ub
      select CASE pCtrl->ControlType
         case CTRL_LABEL
            
         CASE ELSE
      END SELECT
   next
   
   ' If layout has been suspended then only move/resize the control after
   ' all of the relevant properties have now been set.
   if pCtrl->SuspendLayoutChange THEN
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
                               ) as Long

   Dim pWindow As CWindow Ptr = AfxCWindowPtr(pDoc->hWndForm)
   dim pCtrl as clsControl ptr = new clsControl
   dim as CWSTR wszDefaultControlName = GenerateDefaultControlName(pDoc)

   pDoc->Controls.DeselectAllControls

   pCtrl->ControlType = ControlType
   pCtrl->IsActive    = false
   
   pCtrl->hWindow = pWindow->AddControl( GetActiveToolboxControlClassName, , 100, wszDefaultControlName, _
                    rcDraw.left, rcDraw.top, rcDraw.right-rcDraw.left, rcDraw.bottom-rcDraw.top, , , , _
                    CAST(WNDPROC, @Control_SubclassProc), ControlType, CAST(DWORD_PTR, pDoc))
   pDoc->Controls.Add(pCtrl)

   ' Set default properties for the created control
   AttachDefaultControlProperties( pCtrl )
   
   SetControlProperty(pCtrl, "NAME", wszDefaultControlName)
   SetControlProperty(pCtrl, "LEFT", str(rcDraw.left))
   SetControlProperty(pCtrl, "TOP", str(rcDraw.top))
   SetControlProperty(pCtrl, "WIDTH", str(rcDraw.right-rcDraw.left))
   SetControlProperty(pCtrl, "HEIGHT", str(rcDraw.bottom-rcDraw.top))
   
   ' Apply properties to the control
   pCtrl->SuspendLayoutChange = true
   ApplyControlProperties( pCtrl )
   pCtrl->SuspendLayoutChange = false

   pDoc->Controls.SelectControl(pCtrl->hWindow)
   pDoc->Controls.SetActiveControl(pCtrl->hWindow)

   function = 0
END FUNCTION