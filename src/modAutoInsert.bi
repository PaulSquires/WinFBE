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

enum
   BLOCK_STATEMENT_IF 
   BLOCK_STATEMENT_FOR
   BLOCK_STATEMENT_SELECT
   BLOCK_STATEMENT_WHILE
   BLOCK_STATEMENT_DO
   BLOCK_STATEMENT_FUNCTION
   BLOCK_STATEMENT_SUB
   BLOCK_STATEMENT_PROPERTY
   BLOCK_STATEMENT_OPERATOR
   BLOCK_STATEMENT_TYPE
   BLOCK_STATEMENT_WITH
   BLOCK_STATEMENT_ENUM
   BLOCK_STATEMENT_UNION
   BLOCK_STATEMENT_CONSTRUCTOR
   BLOCK_STATEMENT_DESTRUCTOR
end enum

declare function AttemptAutoInsert() as Long

