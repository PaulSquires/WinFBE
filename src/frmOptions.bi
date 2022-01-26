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

#Define IDC_FRMOPTIONS_LABEL1                       1000
#Define IDC_FRMOPTIONS_CMDCANCEL                    1001
#Define IDC_FRMOPTIONS_LBLCATEGORY                  1002
#Define IDC_FRMOPTIONS_CMDOK                        1003
#Define IDC_FRMOPTIONS_TVWCATEGORIES                1004

dim shared OptionsDialogLastOpened as long 

declare Function frmOptions_Show( ByVal hWndParent As HWnd ) as LRESULT

