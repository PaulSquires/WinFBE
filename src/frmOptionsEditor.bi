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

#Define IDC_FRMOPTIONSEDITOR_LBLTABSIZE             1000
#Define IDC_FRMOPTIONSEDITOR_TXTTABSIZE             1001
#Define IDC_FRMOPTIONSEDITOR_LBLKEYWORDCASE         1002
#Define IDC_FRMOPTIONSEDITOR_COMBOCASE              1003
#Define IDC_FRMOPTIONSEDITOR_CHKSHOWLEFTMARGIN      1004
#Define IDC_FRMOPTIONSEDITOR_CHKSYNTAXHIGHLIGHTING  1005
#Define IDC_FRMOPTIONSEDITOR_CHKCURRENTLINE         1006
#Define IDC_FRMOPTIONSEDITOR_CHKLINENUMBERING       1007
#Define IDC_FRMOPTIONSEDITOR_CHKCONFINECARET        1008
#Define IDC_FRMOPTIONSEDITOR_CHKTABTOSPACES         1009
#Define IDC_FRMOPTIONSEDITOR_CHKSHOWFOLDMARGIN      1010
#Define IDC_FRMOPTIONSEDITOR_CHKINDENTGUIDES        1011
#Define IDC_FRMOPTIONSEDITOR_CHKSHOWRIGHTEDGE       1012
#Define IDC_FRMOPTIONSEDITOR_TXTRIGHTEDGE           1013
#Define IDC_FRMOPTIONSEDITOR_LBLRIGHTEDGE           1014
#Define IDC_FRMOPTIONSEDITOR_CHKPOSITIONMIDDLE      1015

declare Function frmOptionsEditor_Show( ByVal hWndParent As HWnd ) as LRESULT


