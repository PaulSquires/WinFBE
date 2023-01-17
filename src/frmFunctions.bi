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


#define IDC_FRMFUNCTIONS_LISTBOX   1000

type FUNCTION_NODE_TYPE
   wszFunctionName as CWSTR
   wszPrototype as CWSTR      ' the sub/function parameters
   nLineNumber as long
end type

enum FunctionsDisplayState
   ViewAsTree = 0
   ViewAsList 
end enum

dim shared gFunctionsDisplay as FunctionsDisplayState = FunctionsDisplayState.ViewAsTree

declare function frmFunctions_Show( byval hWndParent as HWnd ) as LRESULT
declare function frmFunctions_ReparseFiles() as Long
declare function frmFunctions_SelectItemData( byval pDoc as clsDocument ptr ) as boolean
declare function LoadFunctionsFiles() as long
declare function frmFunctions_ViewAsTree() as long
declare function frmFunctions_ViewAsList() as long
declare function QuickSortpDocs( pDocs() As clsDocument ptr, lo as long, hi as long ) as long

