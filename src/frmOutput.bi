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

#Define IDC_FRMOUTPUT_TABCONTROL                    1000
#Define IDC_FRMOUTPUT_LISTVIEW                      1001
#Define IDC_FRMOUTPUT_TXTLOGFILE                    1002
#Define IDC_FRMOUTPUT_BTNCLOSE                      1003
#Define IDC_FRMOUTPUT_LISTSEARCH                    1004
#Define IDC_FRMOUTPUT_LVTODO                        1005
#Define IDC_FRMOUTPUT_TXTNOTES                      1006

declare function frmOutput_ShowNotes() as long 
declare function frmOutput_UpdateToDoListview() as long 
declare Function frmOutput_ShowHideOutputControls( ByVal HWnd As HWnd ) As LRESULT
declare Function frmOutput_PositionWindows As LRESULT
declare Function frmOutput_Show( ByVal hWndParent As HWnd ) As LRESULT

