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

enum eFileClose
   EFC_CLOSECURRENT
   EFC_CLOSEALL
   EFC_CLOSEALLFORWARD
   EFC_CLOSEALLOTHERS
   EFC_CLOSEALLBACKWARD 
end enum

Declare Function OnCommand_ProjectSave( ByVal HWnd As HWnd, ByVal bSaveAs As BOOLEAN = False ) As LRESULT
Declare Function OnCommand_ProjectClose( ByVal HWnd As HWnd ) As LRESULT
Declare Function OnCommand_ProjectNew( ByVal HWnd As HWnd ) As LRESULT
Declare Function OnCommand_ProjectOpen( ByVal HWnd As HWnd ) As LRESULT
Declare Function OnCommand_DesignerNewForm( ByVal HWnd As HWnd ) As LRESULT
Declare Function OnCommand_DesignerAlign( byval HWND as HWND, byval id as long ) as LRESULT
Declare Function OnCommand_DesignerCenter( byval HWND as HWND, byval id as long ) as LRESULT
Declare Function OnCommand_FileNew( ByVal HWnd As HWnd ) As clsDocument ptr
Declare Function OnCommand_FileOpen( ByVal HWnd As HWnd, byval bShowInTab as Boolean = true ) As LRESULT
Declare Function OnCommand_FileSave( ByVal HWnd As HWnd, ByVal bSaveAs As BOOLEAN = False ) As LRESULT
Declare Function OnCommand_FileSaveDeclares( ByVal HWnd As HWnd ) As LRESULT
Declare Function OnCommand_FileSaveAll( ByVal HWnd As HWnd ) As LRESULT
Declare Function OnCommand_FileClose( ByVal HWnd As HWnd, ByVal veFileClose As eFileClose = EFC_CLOSECURRENT ) As LRESULT

