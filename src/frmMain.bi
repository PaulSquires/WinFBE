'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2025 Paul Squires, PlanetSquires Software
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
#Define IDC_FRMMAIN_COMPILETIMER                    1001


'' Last position in document. Used when "Last Position" menu option is selected.
Type LASTPOSITION_TYPE
   pDoc       As clsDocument_ Ptr
   nFirstLine As Long     ' first visible line on screen
   nPosition  As Long     ' Position in Scintilla document where caret is positioned
End Type
Dim Shared gLastPosition As LASTPOSITION_TYPE

declare Function frmMain_GotoFile( ByVal pDoc As clsDocument Ptr, byval nMenuId as long ) As Long
declare Function frmMain_GotoLastPosition() As Long
declare Function frmMain_GotoDefinition( ByVal pDoc As clsDocument Ptr ) As Long
declare Function frmMain_SetStatusbar() as long
declare Function frmMain_SetFocusToCurrentCodeWindow() As Long
Declare Function frmMain_OpenFileSafely( ByVal HWnd As HWnd, ByVal bIsNewFile As BOOLEAN, ByVal bIsTemplate As BOOLEAN, ByVal bShowInTab As BOOLEAN, byval bIsInclude as BOOLEAN, Byref wszName As WString, ByVal pDocIn As clsDocument Ptr, byval bIsDesigner as Boolean = false, byval wszFileType as CWSTR = FILETYPE_UNDEFINED ) As clsDocument Ptr
declare Function frmMain_OpenProjectSafely( ByVal HWnd As HWnd, byref wszProjectFileName as const WString ) as Boolean
declare Function frmMain_PositionWindows() As LRESULT
declare function frmMain_HighlightWord( byval pDoc as clsDocument ptr, byref text as string ) as long
declare Function frmMain_Show( ByVal hWndParent As HWnd ) as LRESULT
