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

#Define IDC_FRMOUTPUT_TABS                          1000
#Define IDC_FRMOUTPUT_LVRESULTS                     1001
#Define IDC_FRMOUTPUT_TXTLOGFILE                    1002
#Define IDC_FRMOUTPUT_LVSEARCH                      1003
#Define IDC_FRMOUTPUT_LVTODO                        1004
#Define IDC_FRMOUTPUT_TXTNOTES                      1005
#Define IDC_FRMOUTPUT_BTNCLOSE                      1006

#define OUTPUT_TABS_HEIGHT  40


type OUTPUT_TABS
   wszText as CWSTR
   rcTab as RECT
   rcText as RECT     ' diff rect b/c line drawn under Text for CurSel
   isHot as boolean
end type

dim shared gOutputTabs(4) as OUTPUT_TABS
dim shared gOutputTabsCurSel as long = 0  ' default to first tab
dim shared gOutputCloseRect as RECT

declare function frmOutput_ShowNotes() as long 
declare function frmOutput_UpdateToDoListview() as long 
declare function frmOutput_UpdateSearchListview( byref wszResultFile as wstring ) as long 
declare Function frmOutput_ShowHideOutputControls( ByVal HWnd As HWnd ) As LRESULT
declare Function frmOutput_PositionWindows As LRESULT
declare Function frmOutput_Show( ByVal hWndParent As HWnd ) As LRESULT

