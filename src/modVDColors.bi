'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2023 Paul Squires, PlanetSquires Software
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

#pragma once

#Define IDC_FRMVDCOLORS_TABCONTROL                  1000
#Define IDC_FRMVDCOLORS_LSTCUSTOM                   1001
#Define IDC_FRMVDCOLORS_LSTCOLORS                   1002
#Define IDC_FRMVDCOLORS_LSTSYSTEM                   1003

#define FRMVDCOLORS_LISTBOX_LINEHEIGHT  20 

enum
   COLOR_CUSTOM = 1
   COLOR_COLORS
   COLOR_SYSTEM
end enum
    
type clsColors
   private:
   public:
      wszColorName as CWSTR          
      ColorType    as long       ' COLOR_QUICK, COLOR_SYSTEM
      ColorValue   as COLORREF
      declare function SetColor( byref wszColorName as wstring, byval ColorType as long, byval ColorValue as COLORREF) as long
END TYPE
function clsColors.SetColor( byref wszColorName as wstring, byval ColorType as long, byval ColorValue as COLORREF) as Long
   this.wszColorName = wszColorName
   this.ColorType    = ColorType   ' COLOR_QUICK, COLOR_SYSTEM
   this.ColorValue   = ColorValue  ' COLORREF
   function = 0
end function
dim shared gColors(any) as clsColors

declare Function frmVDColors_Show( ByVal hWndParent As HWnd, byref wszPropValue as wstring ) as LRESULT
