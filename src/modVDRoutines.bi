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

declare Function GetImagesTypePtr( byref wszImageName as wstring ) As IMAGES_TYPE ptr
declare function CheckMinimumControlSize( byval nCtrlType as long, byval rc as RECT ) as RECT
declare function GetActiveToolboxControlType() as Long
declare function GetActivePropertyPtr() as clsProperty ptr
declare function GetActiveEventPtr() as clsEvent ptr
declare function GetFormCtrlPtr( byval pDoc as clsDocument ptr ) as clsControl ptr
declare function SetActiveToolboxControl( byval ControlType as long ) as Long
declare function GetControlClassName( byval pCtrl as clsControl ptr ) as CWSTR
declare function GetToolBoxName( byval nControlType as long ) as CWSTR
declare function GetControlName( byval nControlType as long ) as CWSTR
declare function GetControlType( byval wszControlName as CWSTR ) as long
declare function GetWinformsXClassName( byval nControlType as long ) as CWSTR
declare function GetControlRECT( byval pCtrl as clsControl ptr ) as RECT



