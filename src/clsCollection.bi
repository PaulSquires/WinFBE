'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2020 Paul Squires, PlanetSquires Software
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


Type clsCollection
   Private:
      Dim _arrControls(Any) As clsControl Ptr
      
   Public:
      Declare Property Count() As Long
      Declare Property ItemFirst() As Long
      Declare Property ItemLast() As Long
      Declare Function ItemAt( ByVal nIndex As Long ) As clsControl Ptr 
      declare function DeselectAllControls() as long
      declare function SelectAllControls() as long
      declare function SelectControl( byval hWndCtrl as hwnd) as long
      declare function SelectedControlsCount() as long
      declare function SetActiveControl( byval hWndCtrl as hwnd) as long
      declare function GetActiveControl() as clsControl ptr
      declare function GetCtrlPtr( byval hWndCtrl as hwnd) as clsControl ptr
      Declare Function Add( ByVal pCtrl As clsControl Ptr ) As Long
      declare function Remove( byval pCtrl as clsControl ptr ) as long
      declare function Debug() as long
      Declare Constructor
      Declare Destructor
End Type

