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


''
''  Application in-memory database
''

' VARIABLE NAMES
' ==============
' Every variable name is prefaced with its owner's name in order to differentiate
' it from other variables (Global variables have no owner name). Usually owners are class/TYPEs.
' The incoming parameter list for a Function is also processed as a variable.
' Therefore, the format is:  <OwnerName>.<FunctionName>.<VariableName>
'
' eg.  
'      TYPE MyType
'         PUBLIC:
'            st As String
'            Declare Function CalculateAmount( ByVal nDollars As Single ) As Single
'      END TYPE
'
'      DIM SHARED AS MyType g     '   ..g     ElementName = ..g   ElementValue = MyType
'
' eg. (Function as part of TYPE)
'      Function MyType.CalculateAmount( ByVal nDollars As Single ) As Single  
'         ' ElementName = MyType.CalculateAmount.nDollars   ElementValue = Single
'         Dim n As Long        ' ElementName = MyType.CalculateAmount.n   ElementValue = Long
'         Dim rc As RECT       ' ElementName = MyType.CalculateAmount.rc  ElementValue = RECT
'      End Function
'
' eg. (Standalone Function)
'      Function MyFunction( ByVal p As MyType Ptr ) As Long   ' ElementName = .MyFunction.p   ElementValue = MyType
'         Dim n As Long        ' ElementName = .MyFunction.n   ElementValue = Long
'         Dim m As MyType      ' ElementName = .MyFunction.m   ElementValue = MyType
'      End Function
'
'
' FUNCTION NAMES
' ==============
'    ElementName = Function name    (can also be fully qualified if from a TYPE.  eg.  MyType.CalculateAmount
'    ElementValue = The Calltip
'
' TYPE STRUCTURES
' ===============
'    ElementName  = The name of the CLASS/TYPE.
'    ElementValue = The true name of the CLASS/TYPE (b/c ElementName may be an ALIAS)
'
' ENUM STRUCTURES
' ===============
'    ElementName  = The name of the ENUM
'    ElementValue = The name of the ENUM
'
' TODO LABELS
' ===============
'    ElementValue = The text of the TODO as found in code.
'
' 

#include once "clsDB2.bi"
#include once "clsParser.bi"


''
''
constructor clsDB2
   redim m_arrData(500) as DB2_DATA
                   
   ' Add the standard builtin data types
   dim parser as clsParser
   with parser
     .lineNum    = -1
     .Filename   = ""
   end with

   dim as CWSTR wszMainStr = _
   "Boolean|Byte|UByte|Short|UShort|Integer|UInteger|Long|ULong|LongInt|ULongInt|Single|" & _
   "Double|String|WString|ZString|CWSTR"

   dim as long nCount = AfxStrParseCount(wszMainStr, "|")
   
   for i as long = 1 to nCount
      parser.typeName = AfxStrParse(wszMainStr, i, "|")
      this.dbAdd( @parser, DB2_STANDARDDATATYPE )  
   next
   this.dbRewind
end constructor    


''
''
function clsDB2.dbAdd( byval parser as clsParser ptr, _
                       byval id as long _
                       ) as DB2_DATA ptr
   
   dim as long nSlot = -1
   
   dim db as DB2_DATA
   
   with db
      .deleted      = false
      .id           = id                ' the type of database record that we are storing
      .nLineNum     = parser->lineNum
      .fileName     = parser->fileName
      
      select CASE id
         case DB2_STANDARDDATATYPE
            .ElementName  = parser->typeName
            .ElementValue = parser->typeName
            .IsWinApi     = parser->IsWinApi
            if .ElementName = "" then exit function
            
         case DB2_TYPE    ' this handles ENUM also
            ' If the type has already been added skip adding it again. 
            if this.dbFindTYPE( parser->typeName) then return 0
            .nLineNum     = parser->funcLineNum
            .ElementName  = parser->typeName
            .ElementValue = parser->typeAlias  ' same as typeName unless it was an ALIAS (always search using this name)
            .IsWinApi     = parser->IsWinApi
            .IsEnum       = parser->IsEnum
            .TypeExtends  = parser->TypeExtends
            if .ElementName = "" then exit function
         
         case DB2_TODO
            .ElementName  = ""
            .ElementValue = parser->todoText
            .nLineNum     = parser->lineNum + 1  ' display in listview
         
         case DB2_FUNCTION, DB2_SUB, DB2_PROPERTY   
            .nLineNum     = parser->funcLineNum
            .ElementName  = parser->funcName
            .ElementValue = parser->funcParams   ' Calltip
            .IsWinApi     = parser->IsWinApi
            .GetSet       = parser->GetSet
            if .ElementName = "" then exit function
         
         case DB2_VARIABLE
            ' Variables are always in the format:
            ' <OwnerName>.<FunctionName>.<VariableName>
            if instr(parser->funcName, ".") = 0 then
               .ElementName = parser->typeName & "." & parser->funcName & "." & parser->varName
            Else
               ' funcName is already mytype.MyFunctionName
               .ElementName = parser->funcName & "." & parser->varName
            end if
            .ElementValue = parser->varType 
            .IsPrivate = iif(parser->bInTypePublic, false, true)
            .IsWinApi     = parser->IsWinApi
            if .ElementName = "" then exit function

      END SELECT

   end with

   ' Find a slot to put the data in
   for i as long = lbound(m_arrData) to ubound(m_arrData)
      if m_arrData(i).deleted THEN
         nSlot = i: exit for
      END IF   
   NEXT
   
   ' If no slot found then resize the array to make room
   if nSlot = -1 THEN
      nSlot = ubound(m_arrData) + 1
      redim preserve m_arrData(nSlot+500) as DB2_DATA
   END IF
   
   m_arrData(nSlot) = db

   function = @m_arrData(nSlot)
end function                

''
''
function clsDB2.dbDelete( byref wszFilename as WString, _
                          byval nParseStartLine as long _
                           ) as long
   dim nCount as long 
   dim as CWSTR wszFile = ucase(wszFilename)
   for i as long = lbound(m_arrData) to ubound(m_arrData)
      if m_arrData(i).deleted = true then continue for
      if ucase(m_arrData(i).fileName) = wszFile THEN
         if m_arrData(i).nLineNum >= nParseStartLine then
            m_arrData(i).deleted = true
            nCount = nCount + 1
         end if
      END IF   
   NEXT
   function = nCount
end function

''
''
function clsDB2.dbDeleteAll() as boolean
   for i as long = lbound(m_arrData) to ubound(m_arrData)
      m_arrData(i).deleted = true
      function = true
   NEXT
end function

''
''
function clsDB2.dbDeleteWinAPI() as boolean
   for i as long = lbound(m_arrData) to ubound(m_arrData)
      if m_arrData(i).IsWinApi THEN
         m_arrData(i).deleted = true
         function = true
      END IF
   NEXT
end function

''
''
function clsDB2.dbRewind() as long
   ' Set index pointer to immediately before first array index 
   m_index = lbound(m_arrData) - 1
   function = 0
END FUNCTION

''
''
function clsDB2.dbGetNext() as DB2_DATA ptr
   ' Set index pointer to next array index that is not deleted
   dim as long ub = ubound(m_arrData)
   do
      m_index = m_index + 1
      if m_index > ub THEN return 0
      if m_arrData(m_index).deleted THEN 
         continue do
      else
         exit do   
      END IF
   loop
   function = @m_arrData(m_index)
END FUNCTION


''
''
function clsDB2.dbSeek( byval sLookFor as string, _
                        byval Action as long, _
                        byval sFilename as string = "" _
                        ) as DB2_DATA ptr
   ' GENERIC SEEK FUNCTION THAT OTHER FUNCTIONS CALL TO DO THE HARD WORK
   ' Find the array element that contains the function name beng searched for
   dim pData as DB2_DATA ptr
            
   ' Determine if this is a get or set property
   dim as long nGetSet, i

   i = instr( sLookFor, " (get)" )
   if i then
      nGetSet = ClassProperty.Getter
   else
      i = instr( sLookFor, " (set)" )
      if i then
         nGetSet = ClassProperty.Setter
      else      
         i = instr( sLookFor, " (ctor)" )
         if i then
            nGetSet = ClassProperty.ctor
         else      
            i = instr( sLookFor, " (dtor)" )
            if i then
               nGetSet = ClassProperty.dtor
            end if
         end if   
      end if
   end if
   if i then sLookFor = left( sLookFor, i - 1 )
      
   sLookFor = ucase(sLookFor)
   this.dbRewind()
   do 
      pData = this.dbGetNext()
      if pData = 0 THEN exit do
      if pData->id = Action THEN
         
         if ( sLookFor = ucase(pData->ElementName) ) andalso _
            ( nGetSet = pData->GetSet )  THEN
            
            if len(sFilename) then
               if ucase(pData->fileName) = ucase(sFilename) then 
                  return pData
               end if    
            else
               return pData
            end if      
         
         end if
      
      END IF
   loop
   
   function = 0
end function


''
''
function clsDB2.dbFindFunction( byref sFunctionName as string, byref sFilename as string = "") as DB2_DATA ptr
   return this.dbSeek( sFunctionName, DB2_FUNCTION, sFilename ) 
end function


''
''
function clsDB2.dbFindSub( byref sFunctionName as string, byref sFilename as string = "") as DB2_DATA ptr
   return this.dbSeek( sFunctionName, DB2_SUB, sFilename ) 
end function


''
''
function clsDB2.dbFindProperty( byref sFunctionName as string, byref sFilename as string = "") as DB2_DATA ptr
   return this.dbSeek( sFunctionName, DB2_PROPERTY, sFilename ) 
end function


''
''
function clsDB2.dbFindVariable( byref sVariableName as string ) as DB2_DATA ptr
   return this.dbSeek( sVariableName, DB2_VARIABLE )      
end function


''
''
function clsDB2.dbFindTYPE( byref sTypeName as string ) as DB2_DATA ptr
   dim pData as DB2_DATA ptr       
   pData = this.dbSeek(sTypeName, DB2_TYPE) 
   if pData = 0 then return 0
   ' If this Type is an ALIAS for another Type then get that real type
   if ucase(pData->ElementName) <> ucase(pData->ElementValue) then
      return this.dbSeek(pData->ElementValue, DB2_TYPE) 
   END IF   
   return pData
end function


''
''
function clsDB2.dbFilenameExists( byref sFilename as CWSTR ) as boolean
   ' Search to see if the incoming filename has already been parsed.
   dim pData as DB2_DATA ptr
   dim as CWSTR wszFilename = ucase(sFilename)
   this.dbRewind()
   do 
      pData = this.dbGetNext()
      if pData = 0 THEN exit do
      if wszFilename = ucase(pData->fileName) THEN return true
   loop
   return false
end function


''
''
function clsDB2.dbDebug() as long
   dim pStream as CTextStream
   pStream.Create("_debug.txt")

   dim pData as DB2_DATA ptr
   
   ' PRINT ALL TYPES
   pStream.WriteLine "CLASSES/TYPES"
   this.dbRewind()
   do 
      pData = this.dbGetNext
      if pData = 0 THEN exit do
      if pData->projectIndex = -1 THEN continue do
      if pData->deleted = true THEN continue do
      if pData->id <> DB2_TYPE then continue do
      pStream.WriteLine "   fileName: " & pData->fileName
      pStream.WriteLine "   ElementName: " & pData->ElementName
      pStream.WriteLine "   ElementValue: " & pData->ElementValue
      pStream.WriteLine "   lineNum: " & pData->nLineNum
      pStream.WriteLine ""
   loop
   pStream.WriteLine ""
   pStream.WriteLine ""
   pStream.WriteLine ""
      
   ' PRINT ALL FUNCTIONS
   pStream.WriteLine "FUNCTIONS"
   this.dbRewind()
   do 
      pData = this.dbGetNext
      if pData = 0 THEN exit do
      if pData->projectIndex = -1 THEN continue do
      if pData->projectIndex = -2 THEN continue do
      if pData->deleted = true THEN continue do
      if pData->id <> DB2_FUNCTION then continue do
      pStream.WriteLine "   fileName: " & pData->fileName
      pStream.WriteLine "   ElementName: " & pData->ElementName
      pStream.WriteLine "   ElementValue: " & pData->ElementValue
      pStream.WriteLine "   lineNum: " & pData->nLineNum
      pStream.WriteLine ""
   loop
   pStream.WriteLine ""
   pStream.WriteLine ""
   pStream.WriteLine ""
      
   ' PRINT ALL SUBS
   pStream.WriteLine "SUBS"
   this.dbRewind()
   do 
      pData = this.dbGetNext
      if pData = 0 THEN exit do
      if pData->projectIndex = -1 THEN continue do
      if pData->projectIndex = -2 THEN continue do
      if pData->deleted = true THEN continue do
      if pData->id <> DB2_SUB then continue do
      pStream.WriteLine "   fileName: " & pData->fileName
      pStream.WriteLine "   ElementName: " & pData->ElementName
      pStream.WriteLine "   ElementValue: " & pData->ElementValue
      pStream.WriteLine "   lineNum: " & pData->nLineNum
      pStream.WriteLine ""
   loop
   pStream.WriteLine ""
   pStream.WriteLine ""
   pStream.WriteLine ""

   ' PRINT ALL PROPERTIES
   pStream.WriteLine "PROPERTIES"
   this.dbRewind()
   do 
      pData = this.dbGetNext
      if pData = 0 THEN exit do
      if pData->projectIndex = -1 THEN continue do
      if pData->projectIndex = -2 THEN continue do
      if pData->deleted = true THEN continue do
      if pData->id <> DB2_PROPERTY then continue do
      pStream.WriteLine "   fileName: " & pData->fileName
      pStream.WriteLine "   ElementName: " & pData->ElementName
      pStream.WriteLine "   ElementValue: " & pData->ElementValue
      pStream.WriteLine "   lineNum: " & pData->nLineNum
      pStream.WriteLine ""
   loop
   pStream.WriteLine ""
   pStream.WriteLine ""
   pStream.WriteLine ""

   ' PRINT ALL VARIABLES
   pStream.WriteLine "VARIABLES"
   this.dbRewind()
   do 
      pData = this.dbGetNext
      if pData = 0 THEN exit do
      if pData->projectIndex = -1 THEN continue do
      if pData->projectIndex = -2 THEN continue do
      if pData->deleted = true THEN continue do
      if pData->id <> DB2_VARIABLE then continue do
      pStream.WriteLine "   fileName: " & pData->fileName
      pStream.WriteLine "   ElementName: " & pData->ElementName
      pStream.WriteLine "   ElementValue: " & pData->ElementValue
      pStream.WriteLine "   IsPrivate: " & pData->IsPrivate
      pStream.WriteLine "   lineNum: " & pData->nLineNum
      pStream.WriteLine ""
   loop
   pStream.WriteLine ""
   
'   ' PRINT ALL TODO
'   pStream.WriteLine "TODO DESCRIPTIONS"
'   this.dbRewind()
'   do 
'      pData = this.dbGetNext
'      if pData = 0 THEN exit do
'      if pData->projectIndex = -1 THEN continue do
'      if pData->deleted = true THEN continue do
'      if pData->id <> DB2_TODO then continue do
'      pStream.WriteLine "   fileName: " & pData->fileName
'      pStream.WriteLine "   ElementName: " & pData->ElementName
'      pStream.WriteLine "   ElementValue: " & pData->ElementValue
'      pStream.WriteLine "   lineNum: " & pData->nLineNum
'      pStream.WriteLine ""
'   loop
'   pStream.WriteLine ""

'   ' PRINT ALL WINAPI/WINFBX FUNCTION DEFINITIONS
'   pStream.WriteLine "WINAPI/WINFBX FUNCTIONS"
'   this.dbRewind()
'   do 
'      pData = this.dbGetNext
'      if pData = 0 THEN exit do
'      if pData->projectIndex <> -1 THEN continue do
'      if pData->deleted = true THEN continue do
'      if pData->id <> DB2_FUNCTION then continue do
'      pStream.WriteLine "   fileName: " & pData->fileName
'      pStream.WriteLine "   ElementName: " & pData->ElementName
'      pStream.WriteLine "   ElementValue: " & pData->ElementValue
'      pStream.WriteLine ""
'   loop
'   pStream.WriteLine ""
   
   pStream.Close
   function = 0
end function


dim shared gdb2 as clsDB2


