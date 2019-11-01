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

#Define IDC_FRMMAIN_TOPTABCONTROL                   1000
#Define IDC_FRMMAIN_TOOLBAR                         1001
#Define IDC_FRMMAIN_REBAR                           1002
#Define IDC_FRMMAIN_STATUSBAR                       1003
#Define IDC_FRMMAIN_COMPILETIMER                    1004
#Define IDC_FRMMAIN_COMBOBUILDS                     1005
#Define IDC_FRMMAIN_COMBOFILES                      1006
#Define IDC_FRMMAIN_COMBOFUNCTIONS                  1007
#Define IDC_FRMMAIN_COMBOTOOLCHAINS                 1008

'' Last position in document. Used when "Last Position" menu option is selected.
Type LASTPOSITION_TYPE
   pDoc       As clsDocument_ Ptr
   nFirstLine As Long     ' first visible line on screen
   nPosition  As Long     ' Position in Scintilla document where caret is positioned
End Type
Dim Shared gLastPosition As LASTPOSITION_TYPE

dim shared as long SPLITSIZE 
SPLITSIZE = AfxScaleY(6)       ' Width/Height of the scrollbar split buttons for split editing windows

declare Function frmMain_AddToComboFiles( byval pDoc as clsDocument ptr ) As Long
declare Function frmMain_RemoveFromComboFiles( byval pDoc as clsDocument ptr ) As Long
declare Function frmMain_SelectComboFunctions() As Long
declare Function frmMain_SelectComboFiles() As Long
declare Function frmMain_LoadComboFunctions() As Long
declare Function frmMain_LoadComboFiles() As Long
declare Function frmMain_SetStatusbar() as long
declare Function frmMain_SetFocusToCurrentCodeWindow() As Long
Declare Function frmMain_OpenFileSafely( ByVal HWnd As HWnd, ByVal bIsNewFile As BOOLEAN, ByVal bIsTemplate As BOOLEAN, ByVal bShowInTab As BOOLEAN, byval bIsInclude as BOOLEAN, ByVal pwszName As WString Ptr, ByVal pDocIn As clsDocument Ptr, byval bIsDesigner as Boolean = false, byval nFileType as long = FILETYPE_UNDEFINED ) As clsDocument Ptr
declare Function frmMain_OpenProjectSafely( ByVal HWnd As HWnd, byref wszProjectFileName as const WString ) as Boolean
declare Function frmMain_PositionWindows() As LRESULT
declare Function frmMain_Show( ByVal hWndParent As HWnd ) as LRESULT
