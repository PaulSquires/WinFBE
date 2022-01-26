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

#Define IDC_FRMFINDINFILES_LBLFINDWHAT              1000
#Define IDC_FRMFINDINFILES_COMBOFIND                1001
#Define IDC_FRMFINDINFILES_COMBOFILES               1002
#Define IDC_FRMFINDINFILES_COMBOFOLDER              1003
#Define IDC_FRMFINDINFILES_CHKWHOLEWORDS            1004
#Define IDC_FRMFINDINFILES_CHKMATCHCASE             1005
#Define IDC_FRMFINDINFILES_FRAMESCOPE               1006
#Define IDC_FRMFINDINFILES_OPTSCOPE1                1007
#Define IDC_FRMFINDINFILES_OPTSCOPE2                1008
#Define IDC_FRMFINDINFILES_OPTSCOPE3                1009
#Define IDC_FRMFINDINFILES_CHKSUBFOLDERS            1010
#Define IDC_FRMFINDINFILES_FRAMEOPTIONS             1011
#Define IDC_FRMFINDINFILES_CMDFILES                 1012
#Define IDC_FRMFINDINFILES_CMDFOLDERS               1013
#Define IDC_FRMFINDINFILES_LABEL1                   1014
#Define IDC_FRMFINDINFILES_LABEL2                   1015
#Define IDC_FRMFINDINFILES_CHKCURRENTDOC            1016
#Define IDC_FRMFINDINFILES_CHKALLOPENDOCS           1017
#Define IDC_FRMFINDINFILES_CHKPROJECT               1018
 
declare Function frmFindInFiles_Show( ByVal hWndParent As HWnd ) As LRESULT
