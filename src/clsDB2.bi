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

#pragma once

#include once "modParser.bi"

''
''  Application in-memory database
''
const DB2_VARIABLE         = 1
const DB2_FUNCTION         = 2    ' Standalone and type functions
const DB2_TYPE             = 3
const DB2_TODO             = 4
const DB2_STANDARDDATATYPE = 5    ' long, integer, string, etc...

const DB2_FILETYPE_FB        = 100
const DB2_FILETYPE_WINAPI    = 101
const DB2_FILETYPE_WINFORMSX = 102
const DB2_FILETYPE_WINFBX    = 103
const DB2_FILETYPE_USERCODE  = 200

' Do not adjust sizes of this type definition because it is saved
' and reloaded from disk (codetip cache database).
type DB2_DATA
   deleted       as boolean = true      ' True/False
   pDoc          as clsDocument ptr     ' Code Document
   nFiletype     as integer             ' See list of DB2_FILETYPE above
   fileName      as wstring * MAX_PATH  ' Filename of source file (needed for deleting).
   id            as integer             ' See DB_* above for what type of record this is.
   nLineStart    as integer             ' Location in the file where starts
   nLineEnd      as integer             ' Location in the file where ends
   ParentName    as zstring * 75        ' Function name / TYPE Name  (blank if global)
   ElementName   as zstring * 75        ' Function name / Variable Name / TYPE Name
   ElementData   as zstring * MAX_PATH  ' Generic text data related to ElementName (todo text, etc)
   CallTip       as zstring * MAX_PATH  ' Function Calltip associated with ElementName variable
   Variabletype  as zstring * 75        ' The type of variable this is. Could be a TYPE name.
   TypeExtends   as zstring * 75        ' The type is extended from this TYPE
   VariableScope as DIMSCOPE      ' Element is public in a type (default) 
   GetSet        as zstring * 75        ' (get) (set)
end type

type clsDB2
   private:
      m_arrData(any) as DB2_DATA
      m_index as integer
      
   public:  
      declare function dbGetFreeSlot() as integer
      declare function dbAdd( byval parser as ctxParser ptr, byval id as integer) as DB2_DATA ptr
      declare function dbDelete( byref wszFilename as wstring ) as integer
      declare function dbDeleteAll() as boolean
      declare function dbDeleteByDocumentPtr( byval pDoc as clsDocument ptr ) as boolean
      declare function dbDeleteByFileType( byval nFiletype as integer ) as boolean
      declare function dbRewind() as integer
      declare function dbGetNext() as DB2_DATA ptr
      declare function dbSeek( byval sParentName as string, byval sLookFor as string, byval Action as integer, byval sFilename as string = "" ) as DB2_DATA ptr
      declare function dbFindFunction( byref sFunctionName as string, byref sFilename as string = "" ) as DB2_DATA ptr
      declare function dbFindFunctionTYPE( byref sTypeName as string, byref sFunctionName as string ) as DB2_DATA ptr
      declare function dbFindVariable( byref sParentName as string, byref sVariableName as string ) as DB2_DATA ptr
      declare function dbFindTYPE( byref sTypeName as string ) as DB2_DATA ptr
      declare function dbWriteDB2( byref wszFilename as wstring ) as integer
      declare function dbReadDB2( byref wszFilename as wstring ) as integer
      declare function dbDebug() as integer
      
      declare constructor
end type

