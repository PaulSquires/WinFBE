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

#include once "clsParser.bi"

' ========================================================================================
' 
' ========================================================================================
private FUNCTION AfxStrRemoveWithMark( BYREF wszMainStr AS CONST WSTRING, _
                                       BYREF wszDelim1 AS CONST WSTRING, _
                                       BYREF wszDelim2 AS CONST WSTRING, _
                                       BYREF MarkKeys AS CONST WSTRING = "", _
                                       BYVAL fRemoveAll AS BOOLEAN = FALSE, _
                                       BYVAL IsInstrRev AS BOOLEAN = FALSE _
                                       ) AS CWSTR
   DIM nPos1 AS LONG =iif(IsInstrRev=true, INSTRRev(wszMainStr, wszDelim1), INSTR(wszMainStr, wszDelim1))
   IF nPos1 = 0 THEN RETURN wszMainStr
   DIM nPos2 AS LONG = INSTR(nPos1 + LEN(wszDelim1), wszMainStr, wszDelim2)
   IF nPos2 = 0 THEN RETURN wszMainStr
   nPos2 += LEN(wszDelim2)
   DIM nLen AS LONG = nPos2 - nPos1
   IF fRemoveAll = FALSE THEN 
      RETURN MID(wszMainStr, 1, nPos1 - 1)& MarkKeys & MID(wszMainStr, nPos2)
   else
      RETURN AfxStrRemoveWithMark(MID(wszMainStr, 1, nPos1 - 1) & MarkKeys & MID(wszMainStr, nPos2), wszDelim1, wszDelim2,MarkKeys, fRemoveAll,IsInstrRev)  
   end if
END FUNCTION


' ========================================================================================
' Get the details of a TODO item
' ========================================================================================
function clsParser.parseToDoItem( byval sText as string ) as boolean
   dim as long nPos = instr(ucase(sText), "' TODO:")
   if nPos = 0 THEN return false
   ' Everything leading up to the found "'" must be a space. Use pointer
   ' indexing for speed to bypass expensive ltrim, etc.
   ' convert to zero base and test up to char immediately before "'"
   for i as Long = 0 to nPos - 2    
      if sText[i] <> 32 THEN return false
   NEXT
   this.todoText = mid(sText, nPos + 7)
   gdb2.dbAdd( @this, DB2_TODO)
   return true
end function

   
' ========================================================================================
' Determine if we are inside a multiline comment block
' ========================================================================================
function clsParser.IsMultilineComment(byval sLine as String) as boolean
   ' If we are already in a comment then test to see if we have now reached the end of it.
   if action = ACTION_PARSECOMMENT THEN 
      if right(sLine, 2) = "'/" THEN 
         action = ACTION_NONE
      end if
   else   
      ' It is possible that the start and end comment markers may appear on 
      ' the same line. Do a test for this.
      if (left(sLine, 2) = "/'") andalso (right(sLine, 2) <> "'/") THEN 
         action = ACTION_PARSECOMMENT
      else
         ' Not in a multiline comment. Just a normal line.
         return false         
      END IF
   END IF
   return true
end function


' ========================================================================================
' Normalize the line by removing double spaces, tabs, etc
' ========================================================================================
function clsParser.NormalizeLine() as boolean
   
   ' Remove double spaces and replace TABs with single space
   this.st = AfxStrShrink(this.st, chr(32,9))

   ' Also save our ucase version of the line
   this.st_ucase = ucase(this.st)
   
   ' Return TRUE if length of resulting line is zero.
   if len(this.st) = 0 THEN return true
end function


' ========================================================================================
' Inspect the line for start/end points of the current parser action
' ========================================================================================
function clsParser.InspectLine() as boolean
   ' Inspect the current line for any start/end points of an existing parse
   ' and take the appropriate action.
   ' Will return TRUE is the inspection changed the state of the current
   ' parser action.
   dim as long i, nCount
   dim as string sTemp
   
   ' CHECK FOR THE START OF A TYPE STRUCTURE.
   if gConfig.CodeTips then
      if len(st_ucase) > 5 THEN 
         if (left(st_ucase, 5) = "TYPE ") THEN
            ' If " AS " exists then this is not a TYPE structure but rather an ALIAS 
            ' for another TYPE.
            i = instr(st_ucase, " AS ") 
            if i THEN 
               this.typeName  = trim(mid(st, 6, (i-6)))
               this.typeAlias = trim(AfxStrRemainI(st, " AS "))
               gdb2.dbAdd( @this, DB2_TYPE )
               this.typeName  = "": this.typeAlias = ""
               this.funcLineNum = this.lineNum
               return true
            else
               ' We have encountered the start of a TYPE structure. Save the TYPE name.
               ' Syntax:   TYPE typename [Extends base_typename] [Field = alignment]
               this.action = ACTION_PARSETYPE
               this.typeName  = AfxStrParse(st, 2, " ")
               this.typeAlias = this.typeName
               ' Handle case where this Type extends from another Type
               i = instr(st_ucase, " EXTENDS") 
               if i THEN 
                  sTemp = AfxStrParse(st, 4, " ")
                  if ucase(sTemp) <> "OBJECT" then
                     this.TypeExtends = sTemp
                  end if   
               end if
               this.bInTypePublic = true    ' assume that all elements in TYPE are public
               this.IsEnum = false
               this.funcLineNum = this.lineNum
               return true
            END IF

         ' CHECK FOR THE END OF A TYPE STRUCTURE
         elseif (left(st_ucase, 8) = "END TYPE") and (gConfig.CodeTips) THEN
            gdb2.dbAdd( @this, DB2_TYPE )
            this.action = ACTION_NONE
            this.typeName  = ""
            this.typeAlias = ""
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
      If Left(st_ucase, 7)  = "STATIC " Then
         st = mid(st, 8): st_ucase = ucase(st)
      end if   
      If Left(st_ucase, 4)  = "SUB " Then
         sTemp = Mid(st, 5): this.action = ACTION_PARSESUB
      elseIf Left(st_ucase, 9)  = "FUNCTION " Then
         ' bypass any line with = that is a "FUNCTION = " line.
         dim as string subText = ltrim( mid(st, 9) )
         if left( subText, 1 ) <> "=" then 
            sTemp = Mid(st, 10)
            this.action = ACTION_PARSEFUNCTION
         end if   
      ElseIf Left(st_ucase, 16) = "PUBLIC FUNCTION " Then
         sTemp = Mid(st, 17): this.action = ACTION_PARSEFUNCTION
      ElseIf Left(st_ucase, 17) = "PRIVATE FUNCTION " Then
         sTemp = Mid(st, 18): this.action = ACTION_PARSEFUNCTION
      ElseIf Left(st_ucase, 11) = "PUBLIC SUB " Then
         sTemp = Mid(st, 12): this.action = ACTION_PARSESUB
      ElseIf Left(st_ucase, 12) = "PRIVATE SUB " Then
         sTemp = Mid(st, 13): this.action = ACTION_PARSESUB
      ElseIf Left(st_ucase, 9) = "PROPERTY " Then
         ' bypass any line with = that is a "PROPERTY = " line.
         dim as string subText = ltrim( mid(st, 9) )
         if left( subText, 1 ) <> "=" then 
            sTemp = Mid(st, 10)
            this.action = ACTION_PARSEPROPERTY
         end if   
      ElseIf Left(st_ucase, 17) = "PRIVATE PROPERTY " Then
         sTemp = Mid(st, 18): this.action = ACTION_PARSEPROPERTY
      ElseIf Left(st_ucase, 12) = "CONSTRUCTOR " Then
         sTemp = Mid(st, 13): this.action = ACTION_PARSECONSTRUCTOR
      ElseIf Left(st_ucase, 11) = "DESTRUCTOR " Then
         sTemp = Mid(st, 12): this.action = ACTION_PARSEDESTRUCTOR
      End If
   end if   
   
   if len(sTemp) THEN 
      ' bypass any line with = embedded in it. Because the FUNCTION keyword has now
      ' been removed from the line, it is easy to check if the "=" sign is the next character.
      if sTemp[0] = 61 then return true   ' 61 is "="
      ' Save data values for later until the end of the block is encountered.
      dim as string sFuncParams, st
      this.funcName = sTemp
      i = instr(sTemp, "(")
      if i THEN this.funcName = rtrim(left(sTemp, i-1))
      if i then sFuncParams = mid(sTemp, i)
      
      ' If this is a PROPERTY then we need to make an educated assessment of whether
      ' it is a Get or Set property and set the flag appropriately.
      if this.action = ACTION_PARSEPROPERTY then
         ' if funcParams exist then this must be a Set property. Need to sanitize the
         ' property parameters first. Need to get the starting ( and the ending ) and
         ' evaluate the text between it. Not as easy as using Parse because there could
         ' be embedded array() parameters and the property could end in something like 
         ' AS LONG, etc.
         dim as long p1, p2
         p1 = instr( sFuncParams, "(" )
         p2 = InstrRev( sFuncParams, ")" )
         if ( p1 = 0 ) orelse (p2 = 0 ) then
            st = ""
         elseif p2 > p1 then
            st = mid( sFuncParams, p1, p2 - p1 )
         end if
         st = trim( st, any "( )" )
         
         if len( st ) then
            this.GetSet = ClassProperty.Setter
         else   
            this.GetSet = ClassProperty.Getter
         end if

      elseif this.action = ACTION_PARSECONSTRUCTOR then
         this.GetSet = ClassProperty.ctor

      elseif this.action = ACTION_PARSEDESTRUCTOR then
         this.GetSet = ClassProperty.dtor

      else      
         this.GetSet = ClassProperty.None
      end if

      if i then this.funcParams = this.funcName & sFuncParams
      this.funcName = AfxStrRemoveI(this.funcName, " OVERLOAD")
      this.funcName = AfxStrRemoveI(this.funcName, " CDECL")
      this.funcName = AfxStrRemoveI(this.funcName, " PASCAL")
      this.funcName = AfxStrRemoveI(this.funcName, " STDCALL")
      this.funcLineNum = this.lineNum
      
      return true
   END IF

   ' TEST FOR END OF SUB/FUNCTION BLOCK
   if len(st_ucase) >= 7 THEN 
      if (left(st_ucase, 7) = "END SUB") orelse _
         (left(st_ucase, 12) = "END FUNCTION") orelse _
         (left(st_ucase, 12) = "END PROPERTY") orelse _
         (left(st_ucase, 15) = "END CONSTRUCTOR") orelse _
         (left(st_ucase, 14) = "END DESTRUCTOR") then 

         ' Add the function to the database
         if this.bParsingCodeGen = false then
            select case this.action
               CASE ACTION_PARSEFUNCTION:    gdb2.dbAdd( @this, DB2_FUNCTION )
               CASE ACTION_PARSESUB:         gdb2.dbAdd( @this, DB2_SUB )
               CASE ACTION_PARSEPROPERTY:    gdb2.dbAdd( @this, DB2_PROPERTY )
               CASE ACTION_PARSECONSTRUCTOR: gdb2.dbAdd( @this, DB2_PROPERTY ) ' lump in with properties
               CASE ACTION_PARSEDESTRUCTOR:  gdb2.dbAdd( @this, DB2_PROPERTY ) ' lump in with properties
            END SELECT
         end if
         
         if gConfig.CodeTips then
            ' Need to parse any function parameters to ensure that they get added
            ' as variables in the database. 
            ' Isolate the parameter list from the codetip
            sTemp = trim(AfxStrExtract ( this.funcParams, "(", ")" ))

            ' Parse the function parameters
            if len(sTemp) THEN 
               this.action = ACTION_PARSEPARAMETERS
               nCount = AfxStrParseCount(sTemp, ",")
               dim as long tlineNum = this.lineNum
               this.lineNum = this.funcLineNum
               for i = 1 to nCount
                  this.st = trim(AfxStrParse(sTemp, i, ","))
                  this.st_ucase = ucase(this.st)
                  this.parseVariableDefinitions
               NEXT
               this.lineNum=tlineNum
            end if
         end if
         
         ' Reset our function variables
         this.action = ACTION_NONE
         this.funcName = ""
         this.funcParams = ""
         return true
      end if
   end if

   return false
END FUNCTION


' ========================================================================================
' Determine if the variable is one of the predefined standard data types
' ========================================================================================
function clsParser.IsStandardDataType( byref sVarType as string ) as Boolean
   static as string sList 
   sList = " boolean byte double integer long longint short single" & _
           " string ubyte uinteger ushort wstring zstring cwstr "
   if instr(sList, " " & lcase(sVarType) & " " ) THEN return true
   return false 
end function


' ========================================================================================
' Parse a DIM that has multiple variables
' ========================================================================================
function clsParser.parseVariableDefinitions() as boolean
   if gConfig.CodeTips = false then return false
   if len(this.st) < 6 THEN return false  ' minimum line length needed "AS " (see below)
   if this.action = ACTION_PARSEENUM THEN return false
   
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

   if left(this.st_ucase, 11) = "DIM SHARED " THEN
      sTemp = mid(this.st, 12)
   elseif left(this.st_ucase, 4) = "DIM " THEN
      sTemp = mid(this.st, 5)
   elseif left(this.st_ucase, 13) = "REDIM SHARED " THEN
      sTemp = mid(this.st, 14)
   elseif left(this.st_ucase, 15) = "REDIM PRESERVE " THEN
      return false
   elseif left(this.st_ucase, 6) = "REDIM " THEN 
      if instr(this.st_ucase, " AS ")=0 then return false
      sTemp = mid(this.st, 7)
   elseif left(this.st_ucase, 3) = "AS " then
      ' This type of declare syntax can only be valid inside a class/TYPE
      ' the word following is the variable type
      if this.action = ACTION_PARSETYPE THEN 
         sTemp = this.st
      else
         return false
      end if      
   elseif instr(this.st_ucase, " AS ") then
      ' This is the normal way variables are defined in a TYPE. Make sure
      ' to restrict this test to types only because otherwise we'll pick up
      ' statements like "For i As Long = 1 to 10 
      ' This is also the way function parameters appear.
      if this.action = ACTION_PARSETYPE THEN 
         sTemp = this.st
      elseif (this.action = ACTION_PARSEPARAMETERS) then
         sTemp = this.st
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
             
   sTemp_ucase = ucase(sTemp)
    
   this.bIsAlias = false
   
   ' "AS Long var1, var2, var3"
   '  12345678
   if left(sTemp_ucase, 3) = "AS " THEN
      n = instr(4, sTemp, " ")
      this.varType = mid(sTemp, 4, n-4) 
      ' Next is a comma separated list of elements DIM'd as the datatype
      sList = mid(sTemp, n)
      n = instr(sList, ",")
      if n then
         n = instrrev(sList, " ",n)
         if n then sList = mid(sList, n+1)
      ELSE
         n = instrrev(sList, " ") 
         if n then sList = mid(sList, n+1) 
      end if
      numParts = AfxStrParseCount(sList, ",")
      for i as long = 1 to numParts
         this.varName = trim(AfxStrParse(sList, i, ","), any " :")
         gdb2.dbAdd( @this, DB2_VARIABLE )
      NEXT
      return true
   
   elseIF instr(sTemp, ",") THEN
      'For example: "S1 AS Long , S2 AS BOOLEAN , ZZ AS WString * MAX_PATH"
      numParts = AfxStrParseCount(sTemp, ",")
      for i as long = 1 to numParts
         sList= trim(AfxStrParse(sTemp, i, ","), any " :")
         this.varName = AfxStrParse(sList, 1, " ") 
         this.varType = trim(AfxStrParse(sList, 3, " "), any " :") 
         gdb2.dbAdd( @this, DB2_VARIABLE )
      NEXT
      return true
   
   else
      ' "var1 AS Long"
      this.varName = AfxStrParse(sTemp, 1, " ") 
      this.varType = trim(AfxStrParse(sTemp, 3, " "), any " :") 
      if this.action = ACTION_PARSETYPE THEN 
         gdb2.dbAdd( @this, DB2_VARIABLE )
      else
         if IsStandardDataType(this.varType) = false THEN 
            gdb2.dbAdd( @this, DB2_VARIABLE )
            return true
         end if
      end if
   end if

   function = false
end function


' ========================================================================================
' Parse a TYPE/Class structure
' ========================================================================================
function clsParser.parseTYPE() as boolean
   if gConfig.CodeTips = false then return false

   ' We are inside of a TYPE structure. Need to examine all lines that are marked
   ' as public in order to construct the "|" delimited list of TYPE elements (variables and functions).
   if this.action <> ACTION_PARSETYPE THEN return false

   ' This function will parse for SUB/FUNCTIONS; Any variables will be identified through
   ' the subsequent parser calls to clsParser.parseVariableDefinitions
   
   dim as string sTemp
   
   if st_ucase[0] = 35 THEN return true   ' skip any embedded conditional "#" defines
   
   ' Determine if we are switching PRIVATE / PUBLIC accesses
   if len(this.st_ucase) >= 7 THEN 
      if left(this.st_ucase, 8) = "PRIVATE:" THEN 
         this.bInTypePublic = false: return true
      elseif left(this.st_ucase, 7) = "PUBLIC:" THEN 
         this.bInTypePublic = true: return true
      elseif left(this.st_ucase, 10) = "PROTECTED:" THEN 
         this.bInTypePublic = true: return true
      END IF
   end if

   
   ' Handle any DECLARE's for TYPEs methods and properties. We only need to retrieve the
   ' function name here because the parameter list would have been picked up elsewhere
   ' during the parse of the action FUNCTION type / END FUNCTION code block.
   ' We don't save any CONSTRUCTOR or DESTRUCTOR declares.
   if len(st_ucase) > 12 THEN 
      if left(st_ucase, 17) = "DECLARE FUNCTION " THEN
         sTemp = mid(st, 18)
      elseif left(st_ucase, 12) = "DECLARE SUB " THEN
         sTemp = mid(st, 13)
      elseif left(st_ucase, 17) = "DECLARE PROPERTY " THEN
         sTemp = mid(st, 18)
      end if
   END IF

   if len(sTemp) then  ' we have found a sub/function definition
      return true
   END IF

   return false
end function


' ========================================================================================
' Parse an ENUM structure
' ========================================================================================
function clsParser.parseENUM() as boolean
   if gConfig.CodeTips = false then return false
   ' We are inside of an ENUM structure. 
   
   ' CHECK FOR THE START OF AN ENUM STRUCTURE.
   if len(st_ucase) > 3 THEN 
      if left(st_ucase, 5) = "ENUM " or st_ucase = "ENUM" THEN
         ' An ENUM can have a name associated with it and also the "explicit" keyword
         ' eg.  ENUM myEnum Explicit
         ' Therefore, the ENUM name is always Parse = 2
         ' We have encountered the start of an ENUM structure. Save the ENUM name.
         ' If no name then skip processing the ENUM.
         action = ACTION_PARSEENUM
         this.typeName  = AfxStrParse(st, 2, " ")
         this.typeAlias = this.typeName
         this.IsEnum = true
         this.bInTypePublic = true    ' assume that all elements in ENUM are public
         this.funcLineNum = this.lineNum
         return true

      ' CHECK FOR THE END OF AN ENUM STRUCTURE
      elseif left(st_ucase, 8) = "END ENUM" THEN
         if len(this.typeName) then
            gdb2.dbAdd( @this, DB2_TYPE )   ' we save ENUM the same way as TYPEs
         end if
         this.action = ACTION_NONE
         this.typeName  = ""
         this.typeAlias = ""
         return true
      end if   
   end if
   
   if this.action <> ACTION_PARSEENUM THEN return false

   ' Only continue to parse this ENUM if this is a named ENUM.
   if len(this.typeName) = 0 then return false
   
   ' An ENUM can have elements on a single line or comma separated. Each element can then be
   ' further broken down as <element> = <value>

   dim as string sTemp
   dim as long nCount = AfxStrParseCount(st, ",")
   
   for i as long = 1 to nCount
      ' Get the parse piece
      sTemp = trim(AfxStrParse(st, i, ","))
      if len(sTemp) then
         ' Get the first parse because we don't need the = <value> piece
         this.varName = trim(AfxStrParse(sTemp, 1, "=") )
         this.varType = "Long"
         gdb2.dbAdd( @this, DB2_VARIABLE )
      end if
   NEXT
   
   return false  

end function


