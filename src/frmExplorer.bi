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


#define IDC_FRMEXPLORER_LBLEXPLORER   1000
#define IDC_FRMEXPLORER_LISTBOX       1001

declare function frmExplorer_Show( byval hWndParent as HWnd ) as LRESULT
declare function LoadExplorerFiles() as long
declare function frmExplorer_SelectItemData( byval pDoc as clsDocument ptr ) as boolean
declare function frmExplorerVScroll_Show( byval hWndParent as HWnd ) as LRESULT
declare function frmExplorerVScroll_PositionWindows( byval nShowState as long ) as LRESULT

' NOTE: These node types are different values than the FileType defines from
' the cldDocument.bi file so we could not reuse those equates. These nodetype
' equates defined the order in which the files will be displayed in the 
' explorer listbox.
 #define NODETYPE_MAIN              0    ' when no project open then this is just all files (UNDEFINED)
 #define NODETYPE_RESOURCE          1
 #define NODETYPE_HEADER            2
 #define NODETYPE_MODULE            3
 #define NODETYPE_NORMAL            4
dim shared gExplorerNodeShow( NODETYPE_MAIN to NODETYPE_NORMAL ) as boolean

type EXPLORER_VSCROLL_TYPE
   listBoxHeight as long
   numItems as long 
   itemHeight as long
   itemsPerPage as long
   thumbHeight as long
   rc as RECT
end type
dim shared gExplorerVScroll as EXPLORER_VSCROLL_TYPE

