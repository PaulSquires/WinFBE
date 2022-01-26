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

#Define IDC_FRMIMAGEMANAGER_TOOLBAR                 1000
#Define IDC_FRMIMAGEMANAGER_LISTVIEW                1001 
#Define IDC_FRMIMAGEMANAGER_COMBO1                  1002
#Define IDC_FRMIMAGEMANAGER_IMAGECTX                1003
#Define IDC_FRMIMAGEMANAGER_FRAMEPREVIEW            1004
#Define IDC_FRMIMAGEMANAGER_LBLFILENAME             1005
#Define IDC_FRMIMAGEMANAGER_FORMATBITMAP            1006
#Define IDC_FRMIMAGEMANAGER_FORMATICON              1007
#Define IDC_FRMIMAGEMANAGER_FORMATCURSOR            1008
#Define IDC_FRMIMAGEMANAGER_FORMATRCDATA            1009
#Define IDC_FRMIMAGEMANAGER_CMDFILENAME             1010

declare Function frmImageManager_Show( ByVal hWndParent As HWnd, Byval pProp as clsProperty ptr = 0 ) As LRESULT
