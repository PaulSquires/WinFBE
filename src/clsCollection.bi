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

#include once "clsControl.bi"


type clsCollection
   private:
      dim _arrControls(any) as clsControl ptr
      
   public:
      declare property Count() as long
      declare property ItemFirst() as long
      declare property ItemLast() as long
      declare function ItemAt( byval nIndex as long ) as clsControl ptr 
      declare function DeselectAllControls() as long
      declare function SelectAllControls() as long
      declare function SelectControl( byval hWndCtrl as hwnd) as long
      declare function SelectedControlsCount() as long
      declare function SetActiveControl( byval hWndCtrl as hwnd) as long
      declare function GetActiveControl() as clsControl ptr
      declare function GetCtrlPtr( byval hWndCtrl as hwnd) as clsControl ptr
      declare function Add( byval pCtrl as clsControl ptr ) as long
      declare function Remove( byval pCtrl as clsControl ptr ) as long
      declare function Debug() as long
      declare constructor
      declare destructor
end type

