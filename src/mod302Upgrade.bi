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


' WinFBE Version 3.0.2+ requires that older version form file formats get upgraded
' to the new json format and also separated out from the source code file. There
' will now be two files: (1) *.frm for the json form definitions, and (2) *.inc/bas
' for the actual form code.

declare function FormUpgrade302Format( byval pDoc as clsDocument ptr ) as boolean

