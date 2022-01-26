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


''
''  Application in-memory database
''

' VARIABLE NAMES
' ==============
' Every variable name is prefaced with its owner's name in order to differentiate
' it from other variables (Global variables have no owner name). Usually owners are class/TYPEs.
' The incoming parameter list for a Function is also processed as a variable.
'
' Therefore, the format is:  <OwnerName>.<FunctionName>.<VariableName>
'
' eg.  
'      TYPE MyType
'         PUBLIC:
'            st As String               ' ElementName = MyType..st  ElementValue = String
'            controls As clsCollection  ' ElementName = MyType..controls  ElementValue = clsCollection

'            ' The Declares in a TYPE would not be parsed. They are picked up elsewhere
'            ' in the source code when the Type body code is found.
'            Declare Function CalculateAmount( ByVal nDollars As Single ) As Single
'      END TYPE
'
' eg.
'      DIM SHARED AS MyType g     '   ..g     ElementName = ..g   ElementValue = MyType
'
' eg. (Function as part of TYPE)
'      Function MyType.CalculateAmount( ByVal nDollars As Single, myTypeVar as MYTYPE ) As Single  
'         ' The incoming nDollars would be discarded by the parser
'         ' MyTypeVar would be stored as:
'         '    ElementName = MyType.CalculateAmount.myTypeVar   ElementValue = MYTYPE
'         Dim n As long        ' This would be discarded by the parser
'         Dim rc As RECT       ' ElementName = MyType.CalculateAmount.rc  ElementValue = RECT
'      End Function
'
' eg. (Standalone Function)
'      Function MyFunction( ByVal p As MyType Ptr ) As long   ' ElementName = .MyFunction.p   ElementValue = MyType
'         Dim n As long        ' This would be discarded by the parser
'         Dim m As MyType      ' ElementName = .MyFunction.m   ElementValue = MyType
'      End Function
'
'
' FUNCTION NAMES
' ==============
'    ElementName = Function name  
'                  (can also be fully qualified if from a TYPE.  eg. MyType.CalculateAmount )
'    ElementValue = The Calltip
'
' TYPE STRUCTURES
' ===============
'    ElementName  = The name of the CLASS/TYPE.
'    ElementValue = The true name of the CLASS/TYPE (b/c ElementName may be an ALIAS)
'    TypeExtends  = The TYPE is extended from this TYPE
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
   "boolean|byte|ubyte|short|ushort|integer|uinteger|long|ulong|longint|ulongint|single|" & _
   "double|string|wstring|zstring|true|false|CWSTR"

   dim as long nCount = AfxStrParseCount(wszMainStr, "|")
   
   for i as long = 1 to nCount
      parser.typeName = AfxStrParse(wszMainStr, i, "|")
      this.dbAdd( @parser, DB2_STANDARDDATATYPE )  
   next
   this.dbRewind
end constructor    


''
''
function clsDB2.dbGetFreeSlot() as long
   
   dim as long nSlot = -1
   
   ' Find a slot to put the data in
   for i as long = lbound(m_arrData) to ubound(m_arrData)
      if m_arrData(i).deleted then
         nSlot = i: exit for
      end if   
   next
   
   ' If no slot found then resize the array to make room
   if nSlot = -1 then
      nSlot = ubound(m_arrData) + 1
      redim preserve m_arrData( nSlot + 10000 ) as DB2_DATA
   end if
   
   function = nSlot
end function
 

''
''
function clsDB2.dbAdd( _
            byval parser as clsParser ptr, _
            byval id as long _
            ) as DB2_DATA ptr
   
   dim db as DB2_DATA
   
   with db
      .deleted      = false
      .id           = id                ' the type of database record that we are storing
      .nLineNum     = parser->lineNum
      .fileName     = parser->fileName
      .nFileType    = parser->nFileType
      
      select CASE id
         case DB2_STANDARDDATATYPE
            .ElementName  = parser->typeName
            .VariableType = parser->typeName
            if .ElementName = "" then exit function
            
         case DB2_TYPE    ' this handles ENUM also
            ' If the type has already been added skip adding it again. 
            if this.dbFindTYPE( parser->typeName) then return 0
            .nLineNum     = parser->funcLineNum
            .ElementName  = parser->typeName
            .VariableType = parser->typeAlias  ' same as typeName unless it was an ALIAS (always search using this name)
            .IsEnum       = parser->IsEnum
            .TypeExtends  = parser->TypeExtends
            if .ElementName = "" then exit function
         
         case DB2_TODO
            .ElementName  = ""
            .ElementData  = parser->todoText
            .nLineNum     = parser->lineNum + 1  ' display in listview
         
         case DB2_FUNCTION, DB2_SUB, DB2_PROPERTY   
            .nLineNum     = parser->funcLineNum
            .ElementName  = parser->funcName
            .CallTip      = parser->funcParams   ' Calltip
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
            .CallTip      = parser->funcParams   ' Calltip
            .VariableType = parser->varType
            .IsPublic     = parser->bInTypePublic
            if .ElementName = "" then exit function

      END SELECT

   end with

   dim as long nSlot = this.dbGetFreeSlot()
   m_arrData(nSlot) = db

   function = @m_arrData(nSlot)
end function                


''
''
function clsDB2.dbAddDirect( byval pData as DB2_DATA ptr ) as long
   ' Thsi function is used when loading data directly from the external 
   ' cache file database. It loads the record directly into the memory array.
   if pData = 0 then exit function
   dim as long nSlot = this.dbGetFreeSlot()
   m_arrData(nSlot) = *pData
   function = 0
end function


''
''
function clsDB2.dbDelete( byref wszFilename as wstring ) as long
   dim nCount as long 
   dim as CWSTR wszFile = ucase(wszFilename)
   for i as long = lbound(m_arrData) to ubound(m_arrData)
      if m_arrData(i).deleted = true then continue for
      if ucase(m_arrData(i).fileName) = wszFile then
         m_arrData(i).deleted = true
         nCount = nCount + 1
      end if   
   next
   function = nCount
end function

''
''
function clsDB2.dbDeleteAll() as boolean
   for i as long = lbound(m_arrData) to ubound(m_arrData)
      m_arrData(i).deleted = true
      function = true
   next
end function

''
''
function clsDB2.dbDeleteByFileType( byval nFileType as long ) as boolean
   ' Delete database entry based on incoming DB2_FILETYPE_* value
   for i as long = lbound(m_arrData) to ubound(m_arrData)
      if m_arrData(i).nFileType = nFileType then
         m_arrData(i).deleted = true
      end if
      function = true
   next
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
function clsDB2.dbGetnext() as DB2_DATA ptr
   ' Set index pointer to next array index that is not deleted
   dim as long ub = ubound(m_arrData)
   do
      m_index = m_index + 1
      if m_index > ub then return 0
      if m_arrData(m_index).deleted then 
         continue do
      else
         exit do   
      end if
   loop
   function = @m_arrData(m_index)
END FUNCTION


''
''
function clsDB2.dbSeek( _
            byval sLookFor as string, _
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
      pData = this.dbGetnext()
      if pData = 0 then exit do
      if pData->id = Action then
         
         if ( sLookFor = ucase(pData->ElementName) ) andalso _
            ( nGetSet = pData->GetSet )  then
            
            if len(sFilename) then
               if ucase(pData->fileName) = ucase(sFilename) then 
                  return pData
               end if    
            else
               return pData
            end if      
         
         end if
      
      end if
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
   if ucase(pData->ElementName) <> ucase(pData->VariableType) then
      return this.dbSeek(pData->VariableType, DB2_TYPE) 
   end if   
   return pData
end function


''
''
function clsDB2.dbDebug() as long
   dim pStream as CTextStream
   pStream.Create("_debug.txt")

   dim as CWSTR wszTitle
   dim as long dataType
   
   dim pData as DB2_DATA ptr
   
   ' PRINT ALL TYPES
   for i as long = 1 to 5
      select case i
         case 1
            wszTitle = "CLASSES/TYPES"
            dataType = DB2_TYPE
         case 2
            wszTitle = "FUNCTIONS"
            dataType = DB2_FUNCTION
         case 3
            wszTitle = "SUBS"
            dataType = DB2_SUB
         case 4
            wszTitle = "PROPERTY"
            dataType = DB2_PROPERTY
         case 5
            wszTitle = "VARIABLES"
            dataType = DB2_VARIABLE
      end select

      pStream.WriteLine wszTitle
      this.dbRewind()
      do 
         pData = this.dbGetnext
         if pData = 0 then exit do
         if pData->id <> dataType then continue do
         'if len(pData->fileName) = 0 then continue do  ' bypass any predefined data
         pStream.WriteLine "   " & wszTitle
         pStream.WriteLine "   deleted: "           & pData->deleted
         pStream.WriteLine "   fileName: "          & pData->fileName
         
         dim as CWSTR wszFileType
         select case pData->nFileType
            case DB2_FILETYPE_FB:        wszFileType = "DB2_FILETYPE_FB"
            case DB2_FILETYPE_WINAPI:    wszFileType = "DB2_FILETYPE_WINAPI"
            case DB2_FILETYPE_WINFORMSX: wszFileType = "DB2_FILETYPE_WINFORMSX"
            case DB2_FILETYPE_WINFBX:    wszFileType = "DB2_FILETYPE_WINFBX"
            case DB2_FILETYPE_USERCODE:  wszFileType = "DB2_FILETYPE_USERCODE"
         end select
         pStream.WriteLine "   nFileType: "         & wszFileType
         pStream.WriteLine "   ElementName: "       & pData->ElementName
         pStream.WriteLine "   VariableType:"       & pData->VariableType
         pStream.WriteLine "   CallTip: "           & pData->CallTip
         pStream.WriteLine "   lineNum: "           & pData->nLineNum
         pStream.WriteLine "   Get/Set/ctor/dtor: " & pData->GetSet
         pStream.WriteLine "   IsPublic (TYPE elements): " & pData->IsPublic
         pStream.WriteLine ""
      loop
      pStream.WriteLine ""
      pStream.WriteLine ""
      pStream.WriteLine ""

   next
   pStream.WriteLine ""
   
   pStream.Close
   function = 0
end function


dim shared gdb2 as clsDB2


