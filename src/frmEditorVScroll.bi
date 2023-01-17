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


type EDITOR_VSCROLL_TYPE
   numLines as long 
   linesPerPage as long
   thumbHeight as long
   thumbRatio as single
   rc as RECT
end type
dim shared gEditorVScroll(1) as EDITOR_VSCROLL_TYPE

declare function frmEditorVScroll_calcVThumbRect( byval pDoc as clsDocument ptr ) as boolean
