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

#include once "clsParser.bi"

''
''  Application in-memory database
''
const DB2_VARIABLE         = 1
const DB2_FUNCTION         = 2    ' Standalone and type functions
const DB2_SUB              = 3    ' Standalone and type subs
const DB2_PROPERTY         = 4    ' GetProp/SetProp/Constr/Destr of TYPEs
const DB2_type             = 5
const DB2_TODO             = 6
const DB2_STANDARDDATAtype = 7    ' long, integer, string, etc...

const DB2_FILETYPE_FB        = 100
const DB2_FILETYPE_WINAPI    = 101
const DB2_FILETYPE_WINFORMSX = 102
const DB2_FILETYPE_WINFBX    = 103
const DB2_FILETYPE_USERCODE  = 200

' Do not adjust sizes of this type definition because it is saved
' and reloaded from disk (codetip cache database).
type DB2_DATA
   deleted      as boolean = true      ' True/False
   nFiletype    as long                ' See list of DB2_FILEtype above
   fileName     as wstring * MAX_PATH  ' Filename of source file (needed for deleting).
   id           as long                ' See DB_* above for what type of record this is.
   nLineNum     as long                ' Location in the file where found
   ElementName  as zstring * 75        ' Function name / Variable Name / type Name
   ElementData  as zstring * MAX_PATH  ' Generic text data related to ElementName (todo text, etc)
   CallTip      as zstring * MAX_PATH  ' Function Calltip associated with ElementName variable
   Variabletype as zstring * 75        ' The type of variable this is. Could be a type name.
   TypeExtends  as zstring * 75        ' The type is extended from this TYPE
   IsPublic     as boolean = true      ' Element is public in a type (default) 
   IsTHIS       as boolean             ' Dynamically set in DereferenceLine so caller can show/hide private elements
   IsEnum       as boolean             ' If type is treated as an ENUM
   GetSet       as ClassProperty       ' 0=sub/function, 1=propertyGet, 2=propertySet, 3=ctor, 4=dtor
end type

type clsDB2
   private:
      m_arrData(any) as DB2_DATA
      m_index as long
      
   public:  
      declare function dbGetFreeSlot() as long
      declare function dbAddDirect( byval pData as DB2_DATA ptr ) as long
      declare function dbAdd( byval parser as clsParser ptr, byval id as long) as DB2_DATA ptr
      declare function dbDelete( byref wszFilename as wstring ) as long
      declare function dbDeleteAll() as boolean
      declare function dbDeleteByFileType( byval nFiletype as long ) as boolean
      declare function dbRewind() as long
      declare function dbGetNext() as DB2_DATA ptr
      declare function dbSeek( byval sLookFor as string, byval Action as long, byval sFilename as string = "" ) as DB2_DATA ptr
      declare function dbFindFunction( byref sFunctionName as string, byref sFilename as string = "" ) as DB2_DATA ptr
      declare function dbFindSub( byref sFunctionName as string, byref sFilename as string = "" ) as DB2_DATA ptr
      declare function dbFindProperty( byref sFunctionName as string, byref sFilename as string = "" ) as DB2_DATA ptr
      declare function dbFindVariable( byref sVariableName as string ) as DB2_DATA ptr
      declare function dbFindTYPE( byref sTypeName as string ) as DB2_DATA ptr
      declare function dbWriteDB2( byref wszFilename as wstring ) as long
      declare function dbReadDB2( byref wszFilename as wstring ) as long
      declare function dbDebug() as long
      
      declare constructor
end type

