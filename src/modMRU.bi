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

#define IDC_MRUBASE          5000  ' Windows id of MRU items 1 to 10 (located under File menu)
#define IDC_MRUPROJECTBASE   6000  ' Windows id of MRUPROJECT items 1 to 10 (located under Project menu)


declare Function OpenMRUFile( ByVal HWnd As HWnd, ByVal wID As Long ) As Long
declare Function ClearMRUlist( ByVal wID As Long ) As Long
declare Function CreateMRUpopup() As HMENU
declare Function UpdateMRUMenu( ByVal hMenu As HMENU ) As Long
declare Function UpdateMRUList( Byref wzFilename As WString ) As Long
declare Function OpenMRUProjectFile( ByVal HWnd As HWnd, ByVal wID As Long) As Long
declare Function UpdateMRUProjectMenu( ByVal hMenu As HMENU ) As Long
declare Function UpdateMRUProjectList( Byval wszFilename As CWSTR ) As Long


