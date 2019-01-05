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

' Directions when determining next closest control pointer
ENUM
   DIRECTION_TOP = 1
   DIRECTION_LEFT
   DIRECTION_RIGHT
   DIRECTION_BOTTOM
end enum   

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

