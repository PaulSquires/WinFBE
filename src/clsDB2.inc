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

#include once "clsDB2.bi"
#include once "modParser.bi"


''
''
constructor clsDB2
   redim m_arrData(500) as DB2_DATA
                   
   ' Add the standard builtin data types
   dim parser as ctxParser
   parser.objectStartLine = -1
   parser.objectEndLine = -1

   dim as CWSTR wszMainStr = _
   "boolean|byte|ubyte|short|ushort|integer|uinteger|long|ulong|longint|ulongint|single|" & _
   "double|string|wstring|zstring|true|false|CWSTR"

   dim as integer nCount = AfxStrParseCount(wszMainStr, "|")
   
   for i as integer = 1 to nCount
      parser.varName = AfxStrParse(wszMainStr, i, "|")
      parser.varType = parser.varName
      this.dbAdd( @parser, DB2_STANDARDDATATYPE )  
   next
   this.dbRewind
end constructor    


''
''
function clsDB2.dbGetFreeSlot() as integer
   
   dim as integer nSlot = -1
   
   ' Find a slot to put the data in
   for i as integer = lbound(m_arrData) to ubound(m_arrData)
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
            byval parser as ctxParser ptr, _
            byval id as integer _
            ) as DB2_DATA ptr
   
   if parser = 0 then exit function
   
   dim db as DB2_DATA
   
   with db
      .deleted      = false
      .id           = id                ' the type of database record that we are storing
      .pDoc         = parser->pDoc
      if parser->pDoc then .fileName = parser->pDoc->DiskFilename
      .nFileType    = parser->nFileType
      
      select CASE id
         case DB2_STANDARDDATATYPE
            .ElementName  = parser->varName
            .VariableType = parser->varType
            if .ElementName = "" then exit function
            
         case DB2_TYPE    ' this handles ENUM also
            ' If the type has already been added skip adding it again. 
            if this.dbFindTYPE( parser->typeName) then return 0
            .nLineStart    = parser->objectStartLine
            .nLineEnd      = parser->objectEndLine
            .ElementName   = parser->typeName
            .VariableType  = parser->typeAlias  ' same as typeName unless it was an ALIAS (always search using this name)
            .TypeExtends   = parser->TypeExtends
            if .ElementName = "" then exit function

         case DB2_TODO
            .ElementName  = ""
            .ElementData  = parser->fullLine
            .nLineStart   = parser->objectStartLine + 1  ' display in listview
            .nLineEnd     = parser->objectEndLine + 1  ' display in listview
    
         case DB2_FUNCTION
            .nLineStart   = parser->objectStartLine
            .nLineEnd     = parser->objectEndLine
            .ElementName  = parser->functionName
            .ParentName   = parser->typeName          
            .CallTip      = parser->functionParams    ' Calltip
            .GetSet       = parser->GetSet
            if .ElementName = "" then exit function
        
         case DB2_VARIABLE
            .nLineStart     = parser->objectStartLine
            .nLineEnd       = parser->objectEndLine
            .ParentName     = parser->functionName
            .ElementName    = parser->varName
            .VariableType   = parser->varType
            .VariableScope  = parser->varScope
            if .ElementName = "" then exit function

      end select

   end with

   dim as integer nSlot = this.dbGetFreeSlot()
   m_arrData(nSlot) = db

   function = @m_arrData(nSlot)
end function                


''
''
function clsDB2.dbDelete( byref wszFilename as wstring ) as integer
   dim nCount as integer 
   dim as CWSTR wszFile = ucase(wszFilename)
   for i as integer = lbound(m_arrData) to ubound(m_arrData)
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
   for i as integer = lbound(m_arrData) to ubound(m_arrData)
      m_arrData(i).deleted = true
      function = true
   next
end function

''
''
function clsDB2.dbDeleteByDocumentPtr( byval pDoc as clsDocument ptr ) as boolean
   ' Delete database entry based on incoming clsDocument pointer
   for i as integer = lbound(m_arrData) to ubound(m_arrData)
      if m_arrData(i).pDoc = pDoc then
         m_arrData(i).deleted = true
      end if
      function = true
   next
end function

''
''
function clsDB2.dbDeleteByFileType( byval nFileType as integer ) as boolean
   ' Delete database entry based on incoming DB2_FILETYPE_* value
   for i as integer = lbound(m_arrData) to ubound(m_arrData)
      if m_arrData(i).nFileType = nFileType then
         m_arrData(i).deleted = true
      end if
      function = true
   next
end function

''
''
function clsDB2.dbRewind() as integer
   ' Set index pointer to immediately before first array index 
   m_index = lbound(m_arrData) - 1
   function = 0
END FUNCTION

''
''
function clsDB2.dbGetnext() as DB2_DATA ptr
   ' Set index pointer to next array index that is not deleted
   dim as integer ub = ubound(m_arrData)
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
            byval sParentName as string, _
            byval sLookFor as string, _
            byval Action as integer, _
            byval sFilename as string = "" _
            ) as DB2_DATA ptr

   ' GENERIC SEEK FUNCTION THAT OTHER FUNCTIONS CALL TO DO THE HARD WORK
   ' Find the array element that contains the function name beng searched for
   dim pData as DB2_DATA ptr
            
   sParentName = ucase(sParentName)
   sLookFor = ucase(sLookFor)
   
   this.dbRewind()
   do 
      pData = this.dbGetnext()
      if pData = 0 then exit do
      if pData->deleted = true then continue do
      if pData->id = Action then
    
         if ( sLookFor = ucase(pData->ElementName) ) andalso _
            ( sParentName = ucase(pData->ParentName) ) then

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
   return this.dbSeek( "", sFunctionName, DB2_FUNCTION, sFilename ) 
end function

''
''
function clsDB2.dbFindFunctionTYPE( byref sTypeName as string, byref sFunctionName as string) as DB2_DATA ptr
   return this.dbSeek( sTypeName, sFunctionName, DB2_FUNCTION ) 
end function

''
''
function clsDB2.dbFindVariable( byref sParentName as string, byref sVariableName as string ) as DB2_DATA ptr
   return this.dbSeek( sParentName, sVariableName, DB2_VARIABLE )      
end function


''
''
function clsDB2.dbFindTYPE( byref sTypeName as string ) as DB2_DATA ptr
   dim pData as DB2_DATA ptr       
   pData = this.dbSeek("", sTypeName, DB2_TYPE) 
   if pData = 0 then return 0
   ' If this Type is an ALIAS for another Type then get that real type
   if ucase(pData->ElementName) <> ucase(pData->VariableType) then
      return this.dbSeek("", pData->VariableType, DB2_TYPE) 
   end if   
   return pData
end function


''
''
function clsDB2.dbDebug() as integer
   dim pStream as CTextStream
   pStream.Create("_debug.txt")

   dim as CWSTR wszTitle
   dim as integer dataType
   
   dim pData as DB2_DATA ptr
   
   ' PRINT ALL TYPES
   for i as integer = 1 to 3
      select case i
         case 1
            wszTitle = "CLASSES/TYPES"
            dataType = DB2_TYPE
         case 2
            wszTitle = "FUNCTIONS"
            dataType = DB2_FUNCTION
         case 3
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
         
         dim as CWSTR wszFileType
         select case pData->nFileType
            case DB2_FILETYPE_FB:        wszFileType = "DB2_FILETYPE_FB"
            case DB2_FILETYPE_WINAPI:    wszFileType = "DB2_FILETYPE_WINAPI"
            case DB2_FILETYPE_WINFORMSX: wszFileType = "DB2_FILETYPE_WINFORMSX"
            case DB2_FILETYPE_WINFBX:    wszFileType = "DB2_FILETYPE_WINFBX"
            case DB2_FILETYPE_USERCODE:  wszFileType = "DB2_FILETYPE_USERCODE"
         end select

         pStream.WriteLine "   " & wszTitle
         pStream.WriteLine "   deleted: "           & pData->deleted
         pStream.WriteLine "   fileName: "          & pData->fileName

         pStream.WriteLine "   nFileType: "         & wszFileType
         pStream.WriteLine "   ParentName: "        & pData->ParentName
         pStream.WriteLine "   ElementName: "       & pData->ElementName
         pStream.WriteLine "   TypeExtends: "       & pData->TypeExtends
         pStream.WriteLine "   VariableType:"       & pData->VariableType
         pStream.WriteLine "   CallTip: "           & pData->CallTip
         pStream.WriteLine "   lineStart: "         & pData->nLineStart
         pStream.WriteLine "   lineEnd: "           & pData->nLineEnd
         pStream.WriteLine "   Get/Set/ctor/dtor: " & pData->GetSet
'         pStream.WriteLine "   Scope:             " & pData->varScope
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


