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

' Enum used to distinguish bewteen a sub/function and Property Get/Set
enum ClassProperty
   None   = 0
   Getter = 1
   Setter = 2
   ctor   = 3   ' constructor
   dtor   = 4   ' destructor
end enum

'  Internal flags for the parser routines
enum
   ACTION_NONE
   ACTION_PARSEFUNCTION     
   ACTION_PARSESUB     
   ACTION_PARSEPROPERTY     
   ACTION_PARSETYPE
   ACTION_PARSEENUM
   ACTION_PARSECOMMENT
   ACTION_PARSEPARAMETERS   ' function parameters
   ACTION_PARSECONSTRUCTOR
   ACTION_PARSEDESTRUCTOR
end enum

type clsParser
   public:
      fileName        as CWSTR
      action          as Long          ' current active action
      lineNum         as long
      st              as String        ' full line, comments and double spaces removed
      st_ucase        as String        ' full line (UCASE), comments and double spaces removed
      funcName        as string        ' Name of function being parsed
      funcParams      as string        ' Parameters to a function identifed in sFuncName
      funcLineNum     as long          ' The line where the sub/function started. This is different than lineNum.
      typeName        as string        ' Name of TYPE being parsed
      typeElements    as string        ' Elements of the TYPE identified in sTypeName
      typeAlias       as string        ' Same as typeName unless ALIAS was detected
      varName         as string        ' Name of variable
      varType         as string        ' Type of variable identified in sVarName
      bIsAlias        as boolean       ' T/F if the stored TYPE name is an ALIAS for another TYPE.
      todoText        as string        ' text description associated with a TODO item
      bInTypePublic   as Boolean       ' PRIVATE/PUBLIC sections of a TYPE
      Description     as string        ' Text from '#Description: tag
      IsWinApi        as boolean       ' If windows.bi was found then database items added
      IsEnum          as Boolean       ' The TYPE is treated as an ENUM
      GetSet          as ClassProperty ' 0=sub/function, 1=propertyGet, 2=propertySet
      TypeExtends     as String        ' The TYPE is extended from this TYPE
      bParsingCodeGen as Boolean       ' flag to parser to properly mark codegen sub/function/properties
      declare function parseToDoItem(byval sText as string) as boolean
      declare function IsMultilineComment(byval sLine as String) as boolean
      declare function NormalizeLine() as boolean
      declare function InspectLine() as boolean
      declare function parseVariableDefinitions() as boolean
      declare function parseTYPE() as boolean
      declare function parseENUM() as boolean
      declare function IsStandardDataType( byref sVarType as string ) as Boolean
END TYPE   

