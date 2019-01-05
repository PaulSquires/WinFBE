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

declare Function frmMain_BuildMenu( ByVal pWindow As CWindow Ptr ) As HMENU
declare Function frmMain_ChangeTopMenuStates() As Long
declare Function CreateStatusBarFileTypeContextMenu() As HMENU
declare Function CreateStatusBarFileEncodingContextMenu() As HMENU
declare function CreateTopTabCtlContextMenu( ByVal idx As Long ) As HMENU
declare Function CreateExplorerContextMenu( ByVal pDoc As clsDocument Ptr ) As HMENU
declare Function CreateScintillaContextMenu() As HMENU

