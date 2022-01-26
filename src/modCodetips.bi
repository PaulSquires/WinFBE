'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2022 Paul Squires, PlanetSquires Software
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

' The type of autocomplete popup that is active. This is necessary
' because the autocomplete popup list is rebuilt every time a new
' character is entered.
enum
   AUTOCOMPLETE_NONE   = 0
   AUTOCOMPLETE_DIM_AS 
   AUTOCOMPLETE_TYPE
end enum   
                 
' Index for scintilla autocomplete raw PNG for SCI_REGISTERRGBAIMAGE
enum
   IMAGE_AUTOC_BASETYPE = 1        ' default variable types. Long, String, Single, etc
   IMAGE_AUTOC_CLASS    = 2        ' class/TYPEs
   IMAGE_AUTOC_METHOD   = 3        ' subs/functions
   IMAGE_AUTOC_PROPERTY = 4        ' variable within a TYPE that can be set directly
end enum

declare function DereferenceLine( byval pDoc as clsDocument ptr, byval sTrigger as String, byval nPosition as long ) as DB2_DATA ptr
