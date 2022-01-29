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


#DEFINE IDC_FRMCATEGORIES_LIST1                    1000
#DEFINE IDC_FRMCATEGORIES_TXTDESCRIPTION           1001
#DEFINE IDC_FRMCATEGORIES_CMDUP                    1002
#DEFINE IDC_FRMCATEGORIES_CMDDOWN                  1003
#DEFINE IDC_FRMCATEGORIES_CMDADD                   1004
#DEFINE IDC_FRMCATEGORIES_CMDEDIT                  1005
#DEFINE IDC_FRMCATEGORIES_CMDDELETE                1006
#DEFINE IDC_FRMCATEGORIES_CMDOK                    1007
#DEFINE IDC_FRMCATEGORIES_CMDCANCEL                1008

#define FRMCATEGORIES_LISTBOX_LINEHEIGHT  20 

declare Function frmCategories_Show( ByVal hWndParent As HWnd ) As LRESULT

