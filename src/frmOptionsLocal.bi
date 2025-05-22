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

#Define IDC_FRMOPTIONSLOCAL_FRAMELOCALIZATION       1001
#Define IDC_FRMOPTIONSLOCAL_CMDNEW                  1002
#Define IDC_FRMOPTIONSLOCAL_CMDEDIT                 1003
#Define IDC_FRMOPTIONSLOCAL_CMDDELETE               1004
#Define IDC_FRMOPTIONSLOCAL_FRAMEEDITAREA           1005
#Define IDC_FRMOPTIONSLOCAL_CMDLOCALIZATION         1006
#Define IDC_FRMOPTIONSLOCAL_LBLPHRASES              1007
#Define IDC_FRMOPTIONSLOCAL_LVWPHRASES              1008
#Define IDC_FRMOPTIONSLOCAL_LBLENGLISH              1009
#Define IDC_FRMOPTIONSLOCAL_TXTENGLISH              1010
#Define IDC_FRMOPTIONSLOCAL_LBLTRANSLATE            1011
#Define IDC_FRMOPTIONSLOCAL_TXTTRANSLATE            1012

declare function frmOptionsLocal_LocalEditCheck() as Long
declare Function frmOptionsLocal_Show( ByVal hWndParent As HWnd ) as LRESULT


