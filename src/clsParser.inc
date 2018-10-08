'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2018 Paul Squires, PlanetSquires Software
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


function clsParser.parseToDoItem( byval sText as string ) as boolean
   dim as long nPos = instr(ucase(sText), "' TODO:")
   if nPos = 0 THEN return false
   ' Everything leading up to the found "'" must be a space. Use pointer
   ' indexing for speed to bypass expensive ltrim, etc.
   ' convert to zero base and test up to char immediately before "'"
   for i as Long = 0 to nPos - 2    
      if sText[i] <> 32 THEN return false
   NEXT
   with this
      .OwnerNode4DB = 0
      .todoText = mid(sText, nPos + 7)
      .CurrentNode4DB = gdb2.dbAdd4DB( this, DB2_TODO)
   end with
   return true
end function

function clsParser.IsMultilineComment(byval sLine as String) as boolean   
   with this
      ' If we are already in a comment then test to see if we have now reached the end of it.
      if .action = ACTION_PARSECOMMENT THEN 
         if right(sLine, 2) = "'/" THEN .action = ACTION_NONE
      else   
         if StrNIsEqual(sLine, "/'", 2) = True THEN 
            .action = ACTION_PARSECOMMENT
         else
            ' Not in a multiline comment. Just a normal line.
            return false         
         END IF
      END IF
   end with
   return true
end function

function clsParser.NormalizeLine() as boolean   
   with this
      ' Remove double spaces and replace TABs with single space
      .st = trim(AfxStrShrink(.st, chr(32,9)))

      ' Also save our ucase version of the line
      .st_ucase = ucase(.st)
      
      ' Return TRUE if length of resulting line is zero.
      if len(.st) = 0 THEN return true
   end with
end function


function clsParser.InspectLine() as boolean
   ' Inspect the current line for any start/end points of an existing parse
   ' and take the appropriate action.
   ' Will return TRUE is the inspection changed the state of the current
   ' parser action.
   dim as long i, nCount
   dim as string sTemp
   
   with this
      ' CHECK FOR THE START OF A TYPE STRUCTURE.
      if gPreparsing = false then
         if len(.st_ucase) > 5 THEN 
            if (StrNIsEqual(.st_ucase, "TYPE ", 5) = True ) and (gConfig.CodeTips <> false) THEN
               ' If " AS " exists then this is not a TYPE structure but rather an ALIAS 
               ' for another TYPE.
               .OwnerNode4DB = 0
               .CurrentNode4DB = 0
               i = instr(.st_ucase, " AS ") 
               if i THEN 
                  .typeName  = trim(mid(.st, 6, (i-6)))
                  .typeAlias = trim(AfxStrRemainI(.st, " AS "))
                  .SubType = DB2_SUBTYPE4ALIAS
                  .funcLineNum = .lineNum
                  .CurrentNode4DB = gdb2.dbAdd4DB(this, DB2_TYPE)
                  .typeName = "": .typeAlias = ""
               else
                  ' We have encountered the start of a TYPE structure. Save the TYPE name.
                  ' Syntax:   TYPE typename [Extends base_typename] [Field = alignment]
                  .action = ACTION_PARSETYPE
                  .typeName  = AfxStrParse(st, 2, " ")
                  .typeAlias = .typeName
                  ' Handle case where this Type extends from another Type
                  i = instr(st_ucase, " EXTENDS") 
                  if i THEN 
                     sTemp = AfxStrParse(st, 4, " ")
                     if ucase(sTemp) <> "OBJECT" then
                        .TypeExtends = sTemp
                     end if   
                  end if
                  .bInTypePublic = true    ' assume that all elements in TYPE are public
                  .SubType = DB2_SUBTYPE4NORMAL
                  .funcLineNum = .lineNum
                  .CurrentNode4DB = gdb2.dbAdd4DB(this, DB2_TYPE)
                  .OwnerNode4DB = .CurrentNode4DB 
               END IF
               return true

            ' CHECK FOR THE END OF A TYPE STRUCTURE
            elseif (StrNIsEqual(st_ucase, "END TYPE", 8) = True) and (gConfig.CodeTips <> false) THEN
               if .OwnerNode4DB then
                  .OwnerNode4DB->nEndLineNum = .lineNum
                  .OwnerNode4DB=0
               endif
               .action = ACTION_NONE
               .typeName  = ""
               .typeAlias = ""
               return true
            end if   
         end if
      end if
      
      ' Check for any SUB/FUNCTION blocks of code. Bypass doing this test
      ' if the line length is less than the minimum (4) because in that case we
      ' know it can not be any of these so why even waste time testing them.
      
      ' TEST FOR START OF SUB/FUNCTION
      sTemp = ""
      if len(st_ucase) > 4 THEN 
         If StrNIsEqual(st_ucase, "STATIC ", 7) = True Then
            st = mid(st, 8)
         end if 
         If StrNIsEqual(st_ucase, "PUBLIC ", 7) = True Then
            st = mid(st, 8)
         end if 
         If StrNIsEqual(st_ucase, "PRIVATE ", 8) = True Then
            st = mid(st, 9)
         end if
         st_ucase = ucase(st)   
         If StrNIsEqual(st_ucase, "SUB ", 4) = True Then
            sTemp = Mid(st, 5): .action = ACTION_PARSESUB
         elseIf StrNIsEqual(st_ucase, "FUNCTION ", 9) = True  Then
            sTemp = Mid(st, 10): .action = ACTION_PARSEFUNCTION
         ElseIf StrNIsEqual(st_ucase, "PROPERTY ", 9) = True Then
            sTemp = Mid(st, 10): .action = ACTION_PARSEPROPERTY
         ElseIf StrNIsEqual(st_ucase, "CONSTRUCTOR ", 12) = True Then
            sTemp = Mid(st, 13): .action = ACTION_PARSECONSTRUCTOR
         ElseIf StrNIsEqual(st_ucase, "DESTRUCTOR ", 11) = True Then
            sTemp = Mid(st, 12): .action = ACTION_PARSEDESTRUCTOR
         End If
      end if   
      if len(sTemp) THEN 
         ' bypass any line with = embedded in it. Because the FUNCTION keyword has now
         ' been removed from the line, it is easy to check if the "=" sign is the next character.
         if sTemp[0] = 61 then return true   ' 61 is "="
         ' Save data values for later until the end of the block is encountered.
         .funcName = sTemp
         i = instr(sTemp, "(")
         if i THEN .funcName = rtrim(left(sTemp, i-1))
         if i then .funcParams = .funcName & mid(sTemp, i)
         .funcName = AfxStrRemoveI(.funcName, " OVERLOAD")
         .funcName = AfxStrRemoveI(.funcName, " CDECL")
         .funcName = AfxStrRemoveI(.funcName, " PASCAL")
         .funcName = AfxStrRemoveI(.funcName, " STDCALL")
         .funcLineNum = .lineNum
         
         ' Need to parse any function parameters to ensure that they get added
         ' as variables in the database. 
         ' Isolate the parameter list from the codetip
         sTemp = trim(AfxStrExtract ( .funcParams, "(", ")" ))

         ' Add the function to the database
         .OwnerNode4DB = 0
         select case .action
            CASE ACTION_PARSEFUNCTION
               .CurrentNode4DB = gdb2.dbAdd4DB(this, DB2_FUNCTION)
               .OwnerNode4DB = .CurrentNode4DB
            CASE ACTION_PARSESUB
               .CurrentNode4DB = gdb2.dbAdd4DB(this, DB2_SUB)
               .OwnerNode4DB = .CurrentNode4DB
            CASE ACTION_PARSECONSTRUCTOR
               .funcName &= "." & .funcName
               .CurrentNode4DB = gdb2.dbAdd4DB(this, DB2_SUB)
               .OwnerNode4DB = .CurrentNode4DB
            CASE ACTION_PARSEDESTRUCTOR
               .funcName &= ".~" & .funcName
               .CurrentNode4DB = gdb2.dbAdd4DB(this, DB2_SUB)
               .OwnerNode4DB = .CurrentNode4DB
            CASE ACTION_PARSEPROPERTY
               .CurrentNode4DB = gdb2.dbAdd4DB(this, DB2_PROPERTY)
               .OwnerNode4DB = .CurrentNode4DB  
         END SELECT
         
         ' Parse the function parameters
         if len(sTemp) THEN 
            .action = ACTION_PARSEPARAMETERS
            .st = sTemp
            'nCount = AfxStrParseCount(sTemp, ",")
            'for i = 1 to nCount
            '   .st = trim(AfxStrParse(sTemp, i, ","))
            '   .st_ucase = ucase(.st)
               .parseVariableDefinitions
            'NEXT
         end if
         
         return true
      END IF

      ' TEST FOR END OF SUB/FUNCTION BLOCK
      if len(st_ucase) >= 7 THEN 
         if (StrNIsEqual(st_ucase, "END SUB", 7) = True) orelse _
            (StrNIsEqual(st_ucase, "END FUNCTION", 12) = True) orelse _
            (StrNIsEqual(st_ucase, "END PROPERTY",12) = True) orelse _
            (StrNIsEqual(st_ucase, "END CONSTRUCTOR",15) = True) orelse _
            (StrNIsEqual(st_ucase, "END DESTRUCTOR",14) = True) THEN
            if .OwnerNode4DB then
               .OwnerNode4DB->nEndLineNum = .lineNum
               .OwnerNode4DB = 0
            endif
            .CurrentNode4DB = 0
            ' Reset our function variables
            .action = ACTION_NONE
            .funcName = ""
            .funcParams = ""
            return true
         end if
      end if
   end with
   return false
END FUNCTION



function clsParser.IsStandardDataType( byref sVarType as string ) as Boolean
   static as string sList 
   sList = " boolean byte double integer long longint short single" & _
           " string ubyte uinteger ushort wstring zstring "
   if instr(sList, " " & lcase(sVarType) & " " ) THEN return true
   return false 
end function

private function ParseVariableDefinitionsEx(byref parser as clsParser,byref sTemp as String) as boolean
   dim as long n, numParts
   dim as String sTemp_ucase = ucase(sTemp), sList       
   ' "AS Long var1, var2, var3"
   '  12345678
   with parser
      if StrNIsEqual(sTemp_ucase, "AS ", 3) = True THEN
         n = instr(4, sTemp, " ")
         .varType = mid(sTemp, 4, n-4) 
         ' Next is a comma separated list of elements DIM'd as the datatype
         sList = mid(sTemp, n)
         n = instr(sList, ",")
         if n then
            do
               n=n-1 
            loop while n>0 and mid(sList,n,1)=" " 
            n = instrrev(sList, " ",n)
            if n then sList = mid(sList, n+1)
         ELSE
            n = instrrev(sList, " ") 
            if n then sList = mid(sList, n+1) 
         end if
         numParts = AfxStrParseCount(sList, ",")
         for i as long = 1 to numParts
            .varName = trim(AfxStrParse(sList, i, ","), any " :")
            .CurrentNode4DB = gdb2.dbAdd4DB(parser, DB2_VARIABLE)
         NEXT
         return true
      elseIF instr(sTemp, ",") THEN
         'For example: "S1 AS Long , S2 AS BOOLEAN , ZZ AS WString * MAX_PATH"
         numParts = AfxStrParseCount(sTemp, ",")
         for i as long = 1 to numParts
            sList= trim(AfxStrParse(sTemp, i, ","), any " :")
            IF LEN(sList) THEN
               .varName = AfxStrParse(sList, 1, " ") 
               .varType = trim(AfxStrParse(sList, 3, " "), any " :") 
               .CurrentNode4DB = gdb2.dbAdd4DB(parser, DB2_VARIABLE)
            ENDIF
         NEXT
         return true
      else
         ' "var1 AS Long"
         .varName = AfxStrParse(sTemp, 1, " ") 
         .varType = trim(AfxStrParse(sTemp, 3, " "), any " :") 
         if .action = ACTION_PARSETYPE THEN 
            .CurrentNode4DB = gdb2.dbAdd4DB(parser, DB2_VARIABLE)
         else 
            .CurrentNode4DB = gdb2.dbAdd4DB(parser, DB2_VARIABLE)
            if .IsStandardDataType(.varType) = false THEN
               return true
            end if
         end if
      end if
   end with
   function = false
end function

function clsParser.parseVariableDefinitions() as boolean
   if gConfig.CodeTips = false then return false
   if gPreparsing then return false
   with this
      if len(.st) < 6 THEN return false  ' minimum line length needed "AS " (see below)
      if .action = ACTION_PARSEENUM THEN return false
      
      ' DIFFERENT WAYS TO DIMENSION VARIABLES
      ' dim shared va1 as long
      ' dim shared as long var1, var2, var3
      ' dim va1 as long
      ' dim as long var1, var2, var3
      ' As Long var1, var2, var3
      ' As Long Ptr var1, var2, var3
      ' As L v    ' minimum
      ' var1 as Long
      
      dim as String sTemp, sTemp_ucase, sList
      dim as long n, n2, numParts

      if StrNIsEqual(.st_ucase, "DIM SHARED ", 11) = True THEN
         sTemp = mid(.st, 12)
      elseif StrNIsEqual(.st_ucase, "DIM ", 4) = True THEN
         sTemp = mid(.st, 5)
      elseif StrNIsEqual(.st_ucase, "REDIM SHARED ", 13) = True THEN
         sTemp = mid(.st, 14)
      elseif StrNIsEqual(.st_ucase, "REDIM PRESERVE ", 15) = True THEN
         return false
      elseif StrNIsEqual(.st_ucase, "REDIM ", 6) = True THEN 
         if instr(.st_ucase, " AS ")=0 then return false
         sTemp = mid(.st, 7)
      elseif StrNIsEqual(.st_ucase, "FOR ", 4) = True THEN 
         if instr(.st_ucase, " AS ")=0 then return false
         sTemp = AfxStrExtract(5,.st, "=")
      elseif StrNIsEqual(.st_ucase, "AS ", 3) = True then
         ' This type of declare syntax can only be valid inside a class/TYPE
         ' the word following is the variable type
         if .action = ACTION_PARSETYPE THEN 
            sTemp = .st
         else
            return false
         end if      
      elseif instr(.st_ucase, " AS ") then
         ' This is the normal way variables are defined in a TYPE. Make sure
         ' to restrict this test to types only because otherwise we'll pick up
         ' statements like "For i As Long = 1 to 10 
         ' This is also the way function parameters appear.
         if .action = ACTION_PARSETYPE THEN 
            sTemp = .st
         elseif (.action = ACTION_PARSEPARAMETERS) then
            sTemp = .st
            sTemp = AfxStrRemoveI(sTemp, "byval ")
            sTemp = AfxStrRemoveI(sTemp, "byref ")
            sTemp = AfxStrRemoveI(sTemp, "const ")
         else   
            return false
         end if      
      end if 

      ' At this point the line (sTemp) has been reduced to one the two following
      ' types of declare lines, or zero length if no declare found.
      ' As Long var1, var2, var3
      ' As Long Ptr var1, var2, var3
      ' var1 as Long
      ' var1 as Long = 12345    ' need to remove initializers
      ' var1 as Long Pointer
      if len(sTemp) = 0 THEN return false

      ' Remove any text and array parenthesis.
      ' For example: "AS WString * MAX_PATH wszText=GETT("FFFF"), wszName ="TTTT""
      sTemp = AfxStrRemove(sTemp, """", """", true ) 
      sTemp = AfxStrRemoveWithMark(sTemp, "(", ")","",true,true) 

      ' Remove any initializers 
      n = instr(sTemp, "=")
      if n THEN
          sTemp = AfxStrRemoveWithMark(sTemp, "=", ",", ",", true, false) 
          n = instr(sTemp, "=")
          if n THEN 
              sTemp =rtrim(left(sTemp, n-1))
          end if
      end if
      
      ' Remove all Pointer references (we don't need them)
      sTemp = AfxStrRemoveI(sTemp, " ptr")
      sTemp = AfxStrRemoveI(sTemp, " pointer")      
      
      .bIsAlias = false
      numParts = AfxStrParseCount(sTemp, ContinueLineFlag)
      IF numParts>1 THEN
         function = false
         n2 = .lineNum
         for i as long = 1 to numParts
            sList= trim(AfxStrParse(sTemp, i, ContinueLineFlag))
            function = ParseVariableDefinitionsEx(this,sList)
            .lineNum =.lineNum+1     
         NEXT
         .lineNum = n2
      ELSE
         function = ParseVariableDefinitionsEx(this,sTemp)  
      end if
   end with
end function


function clsParser.parseTYPE() as boolean
   if gConfig.CodeTips = false then return false
   if gPreparsing then return false

   with this
      ' We are inside of a TYPE structure. Need to examine all lines that are marked
      ' as public in order to construct the "|" delimited list of TYPE elements (variables and functions).
      if .action <> ACTION_PARSETYPE THEN return false

      ' This function will parse for SUB/FUNCTIONS; Any variables will be identified through
      ' the subsequent parser calls to clsParser.parseVariableDefinitions
      
      dim as string sTemp
      
      if st_ucase[0] = 35 THEN return true   ' skip any embedded conditional "#" defines
      
      ' Determine if we are switching PRIVATE / PUBLIC accesses
      if len(.st_ucase) >= 7 THEN 
         if StrNIsEqual(.st_ucase, "PRIVATE:", 8) = True THEN 
            .bInTypePublic = false: return true
         elseif StrNIsEqual(.st_ucase, "PUBLIC:", 7) = True THEN 
            .bInTypePublic = true: return true
         elseif StrNIsEqual(.st_ucase, "PROTECTED:", 10) = True THEN 
            .bInTypePublic = true: return true
         END IF
      end if

      
      ' Handle any DECLARE's for TYPEs methods and properties. We only need to retrieve the
      ' function name here because the parameter list would have been picked up elsewhere
      ' during the parse of the action FUNCTION type / END FUNCTION code block.
      ' We don't save any CONSTRUCTOR or DESTRUCTOR declares.
      if len(st_ucase) > 12 THEN 
         if StrNIsEqual(st_ucase, "DECLARE FUNCTION ", 17) = True THEN
            sTemp = mid(st, 18)
         elseif StrNIsEqual(st_ucase, "DECLARE SUB ", 12) = True THEN
            sTemp = mid(st, 13)
         elseif StrNIsEqual(st_ucase, "DECLARE PROPERTY ", 17) = True THEN
            sTemp = mid(st, 18)
         elseif StrNIsEqual(st_ucase, "DECLARE CONSTRUCTOR ", 20) = True THEN
            sTemp = mid(st, 21)
         elseif StrNIsEqual(st_ucase, "DECLARE DESTRUCTOR ", 19) = True THEN
            sTemp = mid(st, 20)
         end if
      END IF

      if len(sTemp) then  ' we have found a sub/function definition
         return true
      END IF
   end with
   
   return false
end function


function clsParser.parseENUM() as boolean
   if gConfig.CodeTips = false then return false
   if gPreparsing then return false
   ' We are inside of an ENUM structure. 
   
   with this
      ' CHECK FOR THE START OF AN ENUM STRUCTURE.
      if len(.st_ucase) > 3 THEN 
         if .st_ucase = "ENUM" or StrNIsEqual(.st_ucase, "ENUM ", 5) = True THEN
            ' An ENUM can have a name associated with it and also the "explicit" keyword
            ' eg.  ENUM myEnum Explicit
            ' Therefore, the ENUM name is always Parse = 2
            ' We have encountered the start of an ENUM structure. Save the ENUM name.
            ' If no name then skip processing the ENUM.
            .action = ACTION_PARSEENUM
            .typeName  = AfxStrParse(.st, 2, " ")
            .typeAlias = .typeName
            .SubType = DB2_SUBTYPE4ENUM
            .bInTypePublic = true    ' assume that all elements in ENUM are public
            .funcLineNum = .lineNum
            .OwnerNode4DB = 0
            .CurrentNode4DB = 0 
            if len(.typeName) then
              .CurrentNode4DB = gdb2.dbAdd4DB(this, DB2_TYPE)   ' we save ENUM the same way as TYPEs
              .OwnerNode4DB = .CurrentNode4DB
            end if
            return true

         ' CHECK FOR THE END OF AN ENUM STRUCTURE
         elseif StrNIsEqual(st_ucase, "END ENUM", 8) = True THEN
            if .OwnerNode4DB then
               .OwnerNode4DB->nEndLineNum = .lineNum
               .OwnerNode4DB = 0
            endif
            .CurrentNode4DB = 0
            .action = ACTION_NONE
            .typeName  = ""
            .typeAlias = ""
            return true
         end if   
      end if
      
      if .action <> ACTION_PARSEENUM THEN return false

      ' Only continue to parse this ENUM if this is a named ENUM.
      if len(.typeName) = 0 then return false
      
      ' An ENUM can have elements on a single line or comma separated. Each element can then be
      ' further broken down as <element> = <value>

      dim as string sTemp
      dim as long nCount = AfxStrParseCount(.st, ",")
      
      for i as long = 1 to nCount
         ' Get the parse piece
         sTemp = trim(AfxStrParse(.st, i, ","))
         if len(sTemp) then
            ' Get the first parse because we don't need the = <value> piece
            .varName = trim(AfxStrParse(sTemp, 1, "=") )
            .varType = "Long"
            .CurrentNode4DB = gdb2.dbAdd4DB(this, DB2_VARIABLE)
         end if
      NEXT
   end with
   
   return false
end function


