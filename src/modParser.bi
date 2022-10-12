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

#include once "clsDocument.bi"

enum PARSINGACTION
   PARSE_NONE
   PARSE_SUB
   PARSE_FUNCTION
   PARSE_TYPE
   PARSE_DIM
   PARSE_TODO
end enum

enum DIMSCOPE
   SCOPEGLOBAL
   SCOPEFUNCTION
   SCOPETYPE
end enum
   
type ctxParser
   as clsDocument ptr  pDoc
   as zstring ptr      text
   as boolean          incomment
   as boolean          escaped
   as boolean          startofline
   as boolean          inprepro
   as boolean          EOL
   
   as integer          lineNum
   as integer          n
   as integer          s
   as integer          i
   as string           token
   as string           ltoken
   as string           fullLine
   as integer          nFileType    ' one of the DB2_FILETYPE_* codes
   
   as string           objectName 
   as string           objectAlias 
   as string           objectParams
   as string           objectReturnType
   as integer          objectStartLine
   as integer          objectEndLine
   as string           varName 
   as string           varType 
   as DIMSCOPE         varScope
   
   declare function Parse( byval pDoc as clsDocument ptr ) as boolean
   declare function ResetFunctionValues() as boolean
   declare function IsStandardDataType( byref sVarType as string ) as boolean
   declare function PeekChar( byval x as integer = 0 ) as integer
   declare function ReadChar() as integer
   declare function ReadToEOL() as boolean
   declare function ReadToSOL() as boolean
   declare function GetToken() as boolean
   declare function GetLine() as boolean
   declare function UnwindToken() as boolean
   declare function ParseSubFunction( byval action as PARSEACTION ) as boolean
   declare function ParseDIM( byval action as PARSEACTION, byval originFrom as DIMscope ) as boolean
   declare function ParseTYPE( byval action as PARSEACTION ) as boolean
   declare function ParseTODO( byval action as PARSEACTION ) as boolean
   declare function ReadQuoted( byval escapedonce as boolean = FALSE ) as boolean
end type
