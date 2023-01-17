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


#DEFINE IDC_FRMBUILDCONFIG_LIST1                    1000
#DEFINE IDC_FRMBUILDCONFIG_LABEL1                   1001
#DEFINE IDC_FRMBUILDCONFIG_TXTDESCRIPTION           1002
#DEFINE IDC_FRMBUILDCONFIG_LABEL2                   1003
#DEFINE IDC_FRMBUILDCONFIG_TXTOPTIONS               1004
#DEFINE IDC_FRMBUILDCONFIG_CMDUP                    1005
#DEFINE IDC_FRMBUILDCONFIG_CMDDOWN                  1006
#DEFINE IDC_FRMBUILDCONFIG_CMDINSERT                1007
#DEFINE IDC_FRMBUILDCONFIG_CMDDELETE                1008
#DEFINE IDC_FRMBUILDCONFIG_OPT32                    1009
#DEFINE IDC_FRMBUILDCONFIG_OPT64                    1010
#Define IDC_FRMBUILDCONFIG_LSTOPTIONS               1011
#Define IDC_FRMBUILDCONFIG_CHKISDEFAULT             1012

#define FRMBUILDCONFIG_LISTBOX_LINEHEIGHT  20 

declare function frmBuildConfig_getActiveBuildIndex() as long
declare function frmBuildConfig_GetSelectedBuildDescription() as CWSTR
declare function frmBuildConfig_GetSelectedBuildGUID() as String
declare Function frmBuildConfig_Show( ByVal hWndParent As HWnd ) As LRESULT

