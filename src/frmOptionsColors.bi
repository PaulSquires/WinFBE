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

#pragma once

#Define IDC_FRMOPTIONSCOLORS_FRMCOLORS              1001
#Define IDC_FRMOPTIONSCOLORS_LSTTHEMES              1002
#Define IDC_FRMOPTIONSCOLORS_LBLTHEMES              1003
#Define IDC_FRMOPTIONSCOLORS_FRMFONT                1004
#Define IDC_FRMOPTIONSCOLORS_COMBOFONTNAME          1007
#Define IDC_FRMOPTIONSCOLORS_COMBOFONTCHARSET       1008
#Define IDC_FRMOPTIONSCOLORS_COMBOFONTSIZE          1009
#Define IDC_FRMOPTIONSCOLORS_LBLEXTRASPACE          1015
#Define IDC_FRMOPTIONSCOLORS_TXTEXTRASPACE          1016

declare Function frmOptionsColors_Show( ByVal hWndParent As HWnd ) as LRESULT
