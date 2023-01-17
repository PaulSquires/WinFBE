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

type TABORDER_TYPE
   pCtrl as clsControl ptr
   TabIndex   as Long        ' 999999 if TabStop=False or TabIndex property doesn't exist
END TYPE

declare function GenerateFormMetaData( byval pDoc as clsDocument ptr ) as long 
declare function GenerateFormCode( byval pDoc as clsDocument ptr ) as long
declare function GetFormName( byval pDoc as clsDocument ptr ) as CWSTR

