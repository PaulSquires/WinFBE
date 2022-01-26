'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2022 Paul Squires, PlanetSquires Software
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

#define MODCBCOLOR_USERSELECTED    41

declare Function CreateCBColorList( ByVal HWnd As HWnd, _
                                    ByVal CtrlId As Long, _
                                    ByVal nLeft As Long, _
                                    ByVal nTop As Long, _
                                    ByVal nWidth As Long, _
                                    ByVal nHeight As Long _
                                    ) As HWnd
