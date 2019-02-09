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

#Define IDC_FRMFINDREPLACE_BTNCLOSE                 IDCANCEL
#Define IDC_FRMFINDREPLACE_BTNTOGGLE                1001
#Define IDC_FRMFINDREPLACE_TXTFIND                  1002
#Define IDC_FRMFINDREPLACE_TXTREPLACE               1003
#Define IDC_FRMFINDREPLACE_BTNMATCHCASE             1004
#Define IDC_FRMFINDREPLACE_BTNMATCHWHOLEWORD        1005
#Define IDC_FRMFINDREPLACE_BTNREPLACE               1006
#Define IDC_FRMFINDREPLACE_BTNREPLACEALL            1007
#Define IDC_FRMFINDREPLACE_LBLRESULTS               1008
#Define IDC_FRMFINDREPLACE_BTNLEFT                  1009
#Define IDC_FRMFINDREPLACE_BTNRIGHT                 1010
#Define IDC_FRMFINDREPLACE_BTNSELECTION             1011

declare function frmFindReplace_HighlightSearches() as long
declare function frmFindReplace_ShowControls() as long
declare Function frmFindReplace_PositionWindow() As LRESULT
declare Function frmFindReplace_Show( ByVal hWndParent As HWnd, byval fShowReplace as BOOLEAN ) As LRESULT
