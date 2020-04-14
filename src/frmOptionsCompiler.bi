'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2020 Paul Squires, PlanetSquires Software
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

#Define IDC_FRMOPTIONSCOMPILER_CMDFBWIN32             1000
#Define IDC_FRMOPTIONSCOMPILER_LBLSWITCHES            1001
#Define IDC_FRMOPTIONSCOMPILER_LBLFBC32               1002
#Define IDC_FRMOPTIONSCOMPILER_TXTFBWIN32             1003
#Define IDC_FRMOPTIONSCOMPILER_TXTFBSWITCHES          1004
#Define IDC_FRMOPTIONSCOMPILER_LBLFBHELP              1005
#Define IDC_FRMOPTIONSCOMPILER_TXTFBHELPFILE          1006
#Define IDC_FRMOPTIONSCOMPILER_CMDFBHELPFILE          1007
#Define IDC_FRMOPTIONSCOMPILER_CMDFBWIN64             1008
#Define IDC_FRMOPTIONSCOMPILER_LBLFBC64               1009
#Define IDC_FRMOPTIONSCOMPILER_TXTFBWIN64             1010
#Define IDC_FRMOPTIONSCOMPILER_CMDAPIHELPPATH         1011
#Define IDC_FRMOPTIONSCOMPILER_LBLAPIHELP             1012
#Define IDC_FRMOPTIONSCOMPILER_TXTWIN32HELPPATH       1013
#Define IDC_FRMOPTIONSCOMPILER_CHKRUNVIACOMMANDWINDOW 1014
#Define IDC_FRMOPTIONSCOMPILER_CMDWINFBXHELPPATH      1015
#Define IDC_FRMOPTIONSCOMPILER_LBLWINFBXHELP          1016
#Define IDC_FRMOPTIONSCOMPILER_TXTWINFBXHELPPATH      1017
#Define IDC_FRMOPTIONSCOMPILER_CHKDISABLECOMPILEBEEP  1018
declare Function frmOptionsCompiler_Show( ByVal hWndParent As HWnd ) as LRESULT
