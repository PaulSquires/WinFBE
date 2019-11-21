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

#Define IDC_FRMPROJECTOPTIONS_LABEL1                1000
#Define IDC_FRMPROJECTOPTIONS_LABEL2                1001
#Define IDC_FRMPROJECTOPTIONS_LABEL3                1002
#Define IDC_FRMPROJECTOPTIONS_LABEL4                1003
#Define IDC_FRMPROJECTOPTIONS_LABEL5                1004
#Define IDC_FRMPROJECTOPTIONS_LABEL6                1005
#Define IDC_FRMPROJECTOPTIONS_LABEL7                1006
#Define IDC_FRMPROJECTOPTIONS_TXTPROJECTPATH        1007
#Define IDC_FRMPROJECTOPTIONS_CMDSELECT             1008
#Define IDC_FRMPROJECTOPTIONS_TXTOPTIONS32          1009
#Define IDC_FRMPROJECTOPTIONS_TXTOPTIONS64          1010
#Define IDC_FRMPROJECTOPTIONS_CHKMANIFEST           1011
#Define IDC_FRMPROJECTOPTIONS_OPTNONE               1012
#Define IDC_FRMPROJECTOPTIONS_OPTBLANK              1013
#Define IDC_FRMPROJECTOPTIONS_OPTVD                 1014
#Define IDC_FRMPROJECTOPTIONS_OPTCONSOLE            1015
#Define IDC_FRMPROJECTOPTIONS_OPTDLL                1016
#Define IDC_FRMPROJECTOPTIONS_OPTSTATIC             1017

declare Function frmProjectOptions_Show( ByVal hWndParent As HWnd, byval IsNewProject as boolean ) As LRESULT
