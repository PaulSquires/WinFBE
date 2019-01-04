'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2019 Paul Squires, PlanetSquires Software
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

#Define IDC_FRMHELPVIEWER_WEBBROWSER_WINFBE         1000
#Define IDC_FRMHELPVIEWER_WEBBROWSER_WINFBX         1001
#Define IDC_FRMHELPVIEWER_TABCONTROL                1002
#Define IDC_FRMHELPVIEWER_TVWINFBE                  1003
#Define IDC_FRMHELPVIEWER_TVWINFBX                  1004
#Define IDC_FRMHELPVIEWER_BACK                      1005 
#Define IDC_FRMHELPVIEWER_FORWARD                   1006 
#Define IDC_FRMHELPVIEWER_PRINT                     1007 
#Define IDC_FRMHELPVIEWER_FIND                      1008 

declare Function frmHelpViewer_Show( ByVal hWndParent As HWnd, ByVal nCmdShow As Long = -1 ) As LRESULT
