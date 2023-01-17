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

#define IDC_DESIGNFRAME      100
#define IDC_DESIGNFORM       101
#define IDC_DESIGNTABCTRL    102

declare Function DesignerMain_WndProc( ByVal HWnd As HWnd, ByVal uMsg As UINT, ByVal wParam As WPARAM, ByVal lParam As LPARAM ) As LRESULT
