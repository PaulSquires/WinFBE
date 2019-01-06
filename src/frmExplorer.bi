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


#Define IDC_FRMEXPLORER_LBLEXPLORER                 1000
#Define IDC_FRMEXPLORER_TREE                        1001
#Define IDC_FRMEXPLORER_BTNCLOSE                    1002

declare function GetExplorerSpecialSubNode( byval nFileType as long ) as HTREEITEM
declare Function AddFunctionsToExplorerTreeview( ByVal pDoc As clsDocument Ptr, ByVal fUpdateNodes As BOOLEAN ) As Long
declare Function frmExplorer_Show( ByVal hWndParent As HWnd ) As LRESULT
