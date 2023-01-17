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

#Define IDC_LBLFAKEMAINMENU                         1000
#Define IDC_LBLSNAPLINE                             1001
#Define IDC_FAKESTATUSBAR                           1002
#Define IDC_FAKEREBAR                               1003
#Define IDC_FAKETOOLBAR                             1004

declare function KeyboardMoveControls( byval pDoc as clsDocument ptr, byval vKeycode as long ) as Long
declare function KeyboardResizeControls( byval pDoc as clsDocument ptr, byval vKeycode as long ) as Long
declare function KeyboardCycleActiveControls( byval pDoc as clsDocument ptr, byval vKeycode as long ) as Long
declare function IsControlNameExists( byval pDoc as clsDocument ptr, byref wszControlName as wstring ) as boolean
Declare Function IsControlLocked( byval pDoc as clsDocument ptr, byval pCtrl as clsControl ptr ) as boolean
Declare Function IsFormNameExists( byref wszFormName as wstring ) as boolean
declare function CreateToolboxControl( byval pDoc as clsDocument ptr, byval ControlType as long, byref rcDraw as RECT ) as clsControl ptr
declare function ReCreateToolboxControl( byval pDoc as clsDocument ptr, byval pCtrl as clsControl ptr ) as long





