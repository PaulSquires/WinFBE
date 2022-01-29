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

#Define IDC_FRMFUNCTIONLIST_LISTBOX           1000
#Define IDC_FRMFUNCTIONLIST_TXTSEARCH         1001
#define IDC_FRMFUNCTIONLIST_TREE              1002
#Define IDC_FRMFUNCTIONLIST_LBLFUNCTIONLIST   1003

declare function LoadFunctionListFiles() as long
declare function frmFunctionList_Show( ByVal hWndParent As HWnd ) as LRESULT
declare function frmFunctionList_ReparseFiles() as Long
declare function frmFunctionList_CreateSpecialNodes() as HTREEITEM
declare function frmFunctionList_GetSpecialNode( byval wszFileType as CWSTR ) as HTREEITEM
declare function frmFunctionList_GetFileNameFunctionName( byval hItem as HTREEITEM, byref wszFilename as CWSTR, byref wszFunctionName as CWSTR) as long
declare function frmFunctionList_AddParentNode( ByVal pDoc As clsDocument Ptr ) as HTREEITEM
declare function frmFunctionList_AddChildNodes( ByVal pDoc As clsDocument Ptr ) as long
declare function ShowFunctionList() as boolean



