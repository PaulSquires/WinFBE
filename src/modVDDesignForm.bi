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

'  Grab handles (clockwise starting at top left corner)
enum
   GRAB_NOHIT = 0
   GRAB_TOPLEFT 
   GRAB_TOP
   GRAB_TOPRIGHT
   GRAB_RIGHT
   GRAB_BOTTOMRIGHT
   GRAB_BOTTOM
   GRAB_BOTTOMLEFT
   GRAB_LEFT
end enum   

enum SnapLinePosition
   top = 0
   bottom
   left
   right
end enum


declare function IsDesignerView( byval pDoc as clsDocument ptr ) as Boolean
declare function DrawGrabHandles( byval hDC as HDC, byval pDoc as clsDocument ptr, byval bFormOnly as Boolean ) as long
declare function HandleDesignerLButtonDown( ByVal HWnd As HWnd ) as LRESULT
declare function HandleDesignerLButtonUp( ByVal HWnd As HWnd ) as LRESULT
declare function HandleDesignerRButtonDown( ByVal HWnd As HWnd ) as LRESULT
declare function HandleDesignerMouseMove( ByVal HWnd As HWnd ) as LRESULT
declare Function DesignerForm_WndProc( ByVal HWnd   As HWnd, _
                                       ByVal uMsg   As UINT, _
                                       ByVal wParam As WPARAM, _
                                       ByVal lParam As LPARAM _
                                       ) As LRESULT
declare FUNCTION Control_SubclassProc( BYVAL hwnd   AS HWND, _                 ' Control window handle
                                      BYVAL uMsg   AS UINT, _                 ' Type of message
                                      BYVAL wParam AS WPARAM, _               ' First message parameter
                                      BYVAL lParam AS LPARAM, _               ' Second message parameter
                                      BYVAL uIdSubclass AS UINT_PTR, _        ' The subclass ID
                                      BYVAL dwRefData AS DWORD_PTR _          ' Pointer to reference data
                                      ) AS LRESULT















