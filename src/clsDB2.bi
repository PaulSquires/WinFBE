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

#include once "clsParser.bi"

''
''  Application in-memory database
''
const DB2_VARIABLE         = 1
const DB2_FUNCTION         = 2    ' Standalone and TYPE functions
const DB2_SUB              = 3    ' Standalone and TYPE subs
const DB2_PROPERTY         = 4    ' GetProp/SetProp/Constr/Destr of TYPEs
const DB2_TYPE             = 5
const DB2_TODO             = 6
const DB2_STANDARDDATATYPE = 7    ' long, integer, string, etc...

type DB2_DATA
   deleted      as BOOLEAN = true  ' True/False
   fileName     as CWSTR           ' Filename of source file (needed for deleting).
   id           as LONG            ' See DB_* above for what type of record this is.
   nLineNum     as long            ' Location in the file where found
   ElementName  as string          ' Function name / Variable Name / Type Name
   ElementValue as string          ' Function Calltip / TYPE associated with ElementName variable
   TypeExtends  as String          ' The TYPE is extended from this TYPE
   IsPublic     as Boolean = true  ' Element is public in a TYPE (default) 
   IsTHIS       as Boolean         ' Dynamically set in DereferenceLine so caller can show/hide private elements
   IsWinApi     as Boolean         ' If data item is WinApi related
   IsEnum       as Boolean         ' If TYPE is treated as an ENUM
   GetSet       as ClassProperty   ' 0=sub/function, 1=propertyGet, 2=propertySet, 3=ctor, 4=dtor
END TYPE

TYPE clsDB2
   private:
      m_arrData(any) as DB2_DATA
      m_index as LONG
      
   public:  
      declare function dbAdd( byval parser as clsParser ptr, byval id as long) as DB2_DATA ptr
      declare function dbDelete( byref wszFilename as WString, byval nParseStartLine as long ) as long
      declare function dbDeleteAll() as boolean
      declare function dbDeleteWinAPI() as boolean
      declare function dbRewind() as LONG
      declare function dbGetNext() as DB2_DATA ptr
      declare function dbSeek( byval sLookFor as string, byval Action as long, byval sFilename as string = "" ) as DB2_DATA ptr
      declare function dbFindFunction( byref sFunctionName as string, byref sFilename as string = "" ) as DB2_DATA ptr
      declare function dbFindSub( byref sFunctionName as string, byref sFilename as string = "" ) as DB2_DATA ptr
      declare function dbFindProperty( byref sFunctionName as string, byref sFilename as string = "" ) as DB2_DATA ptr
      declare function dbFindVariable( byref sVariableName as string ) as DB2_DATA ptr
      declare function dbFindTYPE( byref sTypeName as string ) as DB2_DATA ptr
      declare function dbDebug() as long
      declare constructor
END TYPE

