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

#Define IDC_FRMVDTOOLBOX_LSTTOOLBOX                 1000
#Define IDC_FRMVDTOOLBOX_LSTPROPERTIES              1001
#Define IDC_FRMVDTOOLBOX_LSTEVENTS                  1002
#Define IDC_FRMVDTOOLBOX_TABCONTROL                 1003
#Define IDC_FRMVDTOOLBOX_COMBOCONTROLS              1004
#Define IDC_FRMVDTOOLBOX_TEXTEDIT                   1005
#Define IDC_FRMVDTOOLBOX_COMBO                      1006
#Define IDC_FRMVDTOOLBOX_COMBOLIST                  1007
#Define IDC_FRMVDTOOLBOX_LBLPROPNAME                1008
#Define IDC_FRMVDTOOLBOX_LBLPROPDESCRIBE            1009

#define FRMVDTOOLBOX_LISTBOX_LINEHEIGHT   20

declare function HidePropertyListControls() as long 
declare Function DisplayPropertyList( byval pDoc as clsDocument ptr ) as Long
declare Function frmVDToolbox_Show( ByVal hWndParent As HWnd, ByVal nCmdShow As Long = 0 ) As LRESULT
