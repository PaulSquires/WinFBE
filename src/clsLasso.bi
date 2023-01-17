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


type clsLasso
   private:
      pWindow      as CWindow ptr 
      hWindow      as hwnd
      hWndParent   as hwnd
      ptStart      as POINT
      ptEnd        as POINT
      bLasso       as boolean
      
   public:
      declare destructor
      declare function IsActive() as boolean
      declare function GetLassoRect() as RECT
      declare function SetStartPoint( byval x as long, byval y as Long) as Long
      declare function SetEndPoint( byval x as long, byval y as Long) as Long
      declare function GetStartPoint() as POINT
      declare function GetEndPoint() as POINT
      declare function FillAlpha(byval hBmp as HBITMAP) as boolean 
      declare function Show() as Long
      declare function Create( byval hWndParent as HWND ) as boolean
      declare function Destroy() as boolean
end type

