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

#Define IDC_FRMRECENT_BTNNEWFILE                    1000
#Define IDC_FRMRECENT_BTNOPENFILE                   1001
#Define IDC_FRMRECENT_BTNNEWPROJECT                 1002
#Define IDC_FRMRECENT_BTNOPENPROJECT                1003
#Define IDC_FRMRECENT_TREEVIEW                      1004

declare Function frmRecent_LoadTreeview() As LRESULT
declare Function frmRecent_PositionWindows As LRESULT
declare Function frmRecent_Show( ByVal hWndParent As HWnd ) As LRESULT

