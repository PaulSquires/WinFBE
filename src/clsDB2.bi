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

const DB2_FILETYPE_FB        = 100
const DB2_FILETYPE_WINAPI    = 101
const DB2_FILETYPE_WINFORMSX = 102
const DB2_FILETYPE_WINFBX    = 103
const DB2_FILETYPE_USERCODE  = 200

' Do not adjust sizes of this type definition because it is saved
' and reloaded from disk (codetip cache database).
type DB2_DATA
   deleted      as BOOLEAN = true      ' True/False
   nFileType    as long                ' See list of DB2_FILETYPE above
   fileName     as WSTRING * MAX_PATH  ' Filename of source file (needed for deleting).
   id           as LONG                ' See DB_* above for what type of record this is.
   nLineNum     as long                ' Location in the file where found
   ElementName  as ZSTRING * 75        ' Function name / Variable Name / Type Name
   ElementData  as Zstring * MAX_PATH  ' Generic text data related to ElementName (todo text, etc)
   CallTip      as Zstring * MAX_PATH  ' Function Calltip associated with ElementName variable
   VariableType as Zstring * 75        ' The type of variable this is. Could be a TYPE name.
   TypeExtends  as ZString * 75        ' The TYPE is extended from this TYPE
   IsPublic     as Boolean = true      ' Element is public in a TYPE (default) 
   IsTHIS       as Boolean             ' Dynamically set in DereferenceLine so caller can show/hide private elements
   IsEnum       as Boolean             ' If TYPE is treated as an ENUM
   GetSet       as ClassProperty       ' 0=sub/function, 1=propertyGet, 2=propertySet, 3=ctor, 4=dtor
END TYPE

TYPE clsDB2
   private:
      m_arrData(any) as DB2_DATA
      m_index as LONG
      
   public:  
      declare function dbGetFreeSlot() as Long
      declare function dbAddDirect( byval pData as DB2_DATA ptr ) as Long
      declare function dbAdd( byval parser as clsParser ptr, byval id as long) as DB2_DATA ptr
      declare function dbDelete( byref wszFilename as WString ) as long
      declare function dbDeleteAll() as boolean
      declare function dbDeleteByFileType( byval nFileType as long ) as boolean
      declare function dbRewind() as LONG
      declare function dbGetNext() as DB2_DATA ptr
      declare function dbSeek( byval sLookFor as string, byval Action as long, byval sFilename as string = "" ) as DB2_DATA ptr
      declare function dbFindFunction( byref sFunctionName as string, byref sFilename as string = "" ) as DB2_DATA ptr
      declare function dbFindSub( byref sFunctionName as string, byref sFilename as string = "" ) as DB2_DATA ptr
      declare function dbFindProperty( byref sFunctionName as string, byref sFilename as string = "" ) as DB2_DATA ptr
      declare function dbFindVariable( byref sVariableName as string ) as DB2_DATA ptr
      declare function dbFindTYPE( byref sTypeName as string ) as DB2_DATA ptr
      declare function dbWriteDB2( byref wszFilename as wstring ) as Long
      declare function dbReadDB2( byref wszFilename as wstring ) as Long
      declare function dbDebug() as long
      
      declare constructor
END TYPE

