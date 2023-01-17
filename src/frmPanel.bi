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

declare function frmPanel_Show( byval hWndParent as HWnd ) as LRESULT
declare function frmPanelVScroll_Show( byval hWndParent as HWnd ) as LRESULT
declare function frmPanelVScroll_PositionWindows( byval nShowState as long ) as LRESULT

type PANEL_BUTTON_TYPE
   wszCaption as CWSTR
   hActionChild as HWND
   rc as RECT
end type
dim shared gPanelButton(any) as PANEL_BUTTON_TYPE

type PANEL_TYPE
   hActiveChild as HWND
   wszLabel as CWSTR
   rcLabel as RECT
   rcActionMenu as RECT
end type
dim shared gPanel as PANEL_TYPE

type PANEL_VSCROLL_TYPE
   hListBox as HWND
   listBoxHeight as long
   numItems as long 
   itemHeight as long
   itemsPerPage as long
   thumbHeight as long
   rc as RECT
end type
dim shared gPanelVScroll as PANEL_VSCROLL_TYPE

