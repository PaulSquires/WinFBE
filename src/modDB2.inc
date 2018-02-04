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
'         Dim rc As RECT       ' ElementName = MyType.CalculateAmount.m   ElementValue = RECT
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

''
''
constructor clsDB2
   redim m_arrData(500) as DB2_DATA
                   
   ' Add the standard builtin data types
   dim parser as clsParser
   with parser
     .lineNum    = -1
     .idxProject = -1
     .Filename   = ""
   end with
   parser.typeName = "Boolean"
   parser.Description = !"Data type that holds the value of True or False.\n"
   this.dbAdd(parser, DB2_STANDARDDATATYPE)  
               
   parser.typeName = "Byte"      
   parser.Description = !"8-bit signed whole-number data type.\nValue range of -128 to 127."
   this.dbAdd(parser, DB2_STANDARDDATATYPE)  

   parser.typeName = "UByte"
   parser.Description = !"8-bit signed whole-number data type.\nValue range of 0 to 255."
   this.dbAdd(parser, DB2_STANDARDDATATYPE)  
   
   parser.typeName = "Short"
   parser.Description = !"16-bit signed whole-number data type.\nValue range of -32768 to 32767."
   this.dbAdd(parser, DB2_STANDARDDATATYPE)  
   
   parser.typeName = "UShort"
   parser.Description = !"16-bit signed whole-number data type.\nValue range of 0 to 65535."
   this.dbAdd(parser, DB2_STANDARDDATATYPE)  

   parser.typeName = "Integer"
   parser.Description = !"32-bit or 64-bit signed whole-number data type depending on the platform.\n"
   this.dbAdd(parser, DB2_STANDARDDATATYPE)  
   
   parser.typeName = "UInteger"
   parser.Description = !"32-bit or 64-bit unsigned whole-number data type depending on the platform.\n"
   this.dbAdd(parser, DB2_STANDARDDATATYPE)  
   
   parser.typeName = "Long"
   parser.Description = !"32-bit signed whole-number data type.\nValue range of -2147483648 to 2147483647."
   this.dbAdd(parser, DB2_STANDARDDATATYPE)  

   parser.typeName = "ULong"
   parser.Description = !"32-bit unsigned whole-number data type.\nValue range of 0 to 4294967295."
   this.dbAdd(parser, DB2_STANDARDDATATYPE)  
   
   parser.typeName = "LongInt"
   parser.Description = !"A 64-bit signed whole-number data type.\nValues range of -9,223,372,036,854,775,808 to 9,223,372,036,854,775,807"
   this.dbAdd(parser, DB2_STANDARDDATATYPE)  

   parser.typeName = "ULongInt"
   parser.Description = !"A 64-bit unsigned whole-number data type.\nValues range of 0 to 18,446,744,073,709,551,615"
   this.dbAdd(parser, DB2_STANDARDDATATYPE)  

   parser.typeName = "Single"
   parser.Description = !"32-bit floating-point data type.\nContains at most 24 bits of precision, or about 6 decimal digits."
   this.dbAdd(parser, DB2_STANDARDDATATYPE)  

   parser.typeName = "Double"
   parser.Description = !"64-bit floating-point data type.\nContains at most 53 bits of precision, or about 15 decimal digits."
   this.dbAdd(parser, DB2_STANDARDDATATYPE)  

   parser.typeName = "String"
   parser.Description = "A String is an array of characters."
   this.dbAdd(parser, DB2_STANDARDDATATYPE)  
   
   parser.typeName = "WString"
   parser.Description = !"A fixed-size array of wide-chars that never overflows if the size is known at compile-time.\n"
   this.dbAdd(parser, DB2_STANDARDDATATYPE)  
   
   parser.typeName = "ZString"
   parser.Description = "A C-style fixed-size array of chars."
   this.dbAdd(parser, DB2_STANDARDDATATYPE)  

   this.dbRewind
end constructor    

''
''
function clsDB2.dbAdd( byref parser as clsParser, _
                       byref id as long _
                       ) as DB2_DATA ptr
   
   dim as long nSlot = -1
   
   dim db as DB2_DATA
   
   with db
      .deleted      = false
      .id           = id                ' the type of database record that we are storing
      .nLineNum     = parser.lineNum
      .projectIndex = parser.idxProject
      .fileName     = parser.fileName
      
      select CASE id
         case DB2_STANDARDDATATYPE
            .ElementName  = parser.typeName
            .ElementValue = parser.typeName
            .IsWinApi     = parser.IsWinApi
            .Description  = parser.Description
            if .ElementName = "" then exit function
            
         case DB2_TYPE    ' this handles ENUM also
            ' If the type has already been added skip adding it again. 
            if this.dbFindTYPE( parser.typeName) then return 0
            .nLineNum     = parser.funcLineNum
            .ElementName  = parser.typeName
            .ElementValue = parser.typeAlias     ' same as typeName unless it was an ALIAS (always search using this name)
            .IsWinApi     = parser.IsWinApi
            .IsEnum       = parser.IsEnum
            .TypeExtends  = parser.TypeExtends
            .Description  = ""
            if .ElementName = "" then exit function
         
         case DB2_TODO
            .ElementName  = ""
            .ElementValue = parser.todoText
            .Description  = ""
         
         case DB2_FUNCTION, DB2_SUB, DB2_PROPERTY   
            .nLineNum     = parser.funcLineNum
            .ElementName  = parser.funcName
            .ElementValue = parser.funcParams   ' Calltip
            .IsWinApi     = parser.IsWinApi
            .Description  = ""
            if .ElementName = "" then exit function
         
         case DB2_VARIABLE
            ' Variabls are always in the format:
            ' <OwnerName>.<FunctionName>.<VariableName>
            if instr(parser.funcName, ".") = 0 then
               .ElementName = parser.typeName & "." & parser.funcName & "." & parser.varName
            Else
               ' funcName is already mytype.MyFunctionName
               .ElementName = parser.funcName & "." & parser.varName
            end if
            .ElementValue = parser.varType 
            .IsPrivate = iif(parser.bInTypePublic, false, true)
            .IsWinApi     = parser.IsWinApi
            .Description  = ""
            if .ElementName = "" then exit function

         case DB2_INCLUDEFILE
            .ElementName  = ""
            .ElementValue = parser.varName
            .Description  = ""
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
function clsDB2.dbDelete( byref wszFilename as WString ) as long
   dim nCount as long 
   dim as CWSTR wszFile = ucase(wszFilename)
   for i as long = lbound(m_arrData) to ubound(m_arrData)
      if m_arrData(i).deleted = true then continue for
      if ucase(m_arrData(i).fileName) = wszFile THEN
         m_arrData(i).deleted = true
         nCount = nCount + 1
      END IF   
   NEXT
   function = nCount
end function

''
''
function clsDB2.dbDeleteByProject( byval idx as long ) as boolean
   for i as long = lbound(m_arrData) to ubound(m_arrData)
      if m_arrData(i).projectIndex = idx THEN
         m_arrData(i).deleted = true
         function = true
      END IF
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
function clsDB2.dbSeek( byval sLookFor as string, byval Action as long ) as DB2_DATA ptr
   ' GENERIC SEEK FUNCTION THAT OTHER FUNCTIONS CALL TO DO THE HARD WORK
   ' Find the array element that contains the function name beng searched for
   dim pData as DB2_DATA ptr
            
   sLookFor = ucase(sLookFor)
   this.dbRewind()
   do 
      pData = this.dbGetNext()
      if pData = 0 THEN exit do
      if pData->id = Action THEN
         if sLookFor = ucase(pData->ElementName) THEN
            return pData
         end if
      END IF
   loop
   
   function = 0
end function


''
''
function clsDB2.dbFindFunction( byref sFunctionName as string ) as DB2_DATA ptr
   return this.dbSeek(sFunctionName, DB2_FUNCTION) 
end function

''
''
function clsDB2.dbFindSub( byref sFunctionName as string ) as DB2_DATA ptr
   return this.dbSeek(sFunctionName, DB2_SUB) 
end function

''
''
function clsDB2.dbFindProperty( byref sFunctionName as string ) as DB2_DATA ptr
   return this.dbSeek(sFunctionName, DB2_PROPERTY) 
end function

''
''
function clsDB2.dbFindVariable( byref sVariableName as string ) as DB2_DATA ptr
   return this.dbSeek(sVariableName, DB2_VARIABLE)      
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
   dim as long f = freefile
   open "_debug.txt" for output as #f
   dim pData as DB2_DATA ptr
   
   ' PRINT ALL FILENAMES
   print #f, "INCLUDE FILENAMES"
   this.dbRewind()
   do 
      pData = this.dbGetNext
      if pData = 0 THEN exit do
      if pData->projectIndex = -1 THEN continue do
      if pData->deleted = true THEN continue do
      if pData->id <> DB2_INCLUDEFILE then continue do
      print #f, "   ElementValue: "; pData->ElementValue
   loop
   print #f, ""
   
   ' PRINT ALL TYPES
   print #f, "CLASSES/TYPES"
   this.dbRewind()
   do 
      pData = this.dbGetNext
      if pData = 0 THEN exit do
      if pData->projectIndex = -1 THEN continue do
      if pData->deleted = true THEN continue do
      if pData->id <> DB2_TYPE then continue do
      print #f, "   fileName: "; pData->fileName
      print #f, "   ElementName: "; pData->ElementName
      print #f, "   ElementValue: "; pData->ElementValue
      print #f, "   lineNum: "; pData->nLineNum
      print #f, ""
   loop
   print #f, ""
      
   ' PRINT ALL FUNCTIONS
   print #f, "FUNCTIONS"
   this.dbRewind()
   do 
      pData = this.dbGetNext
      if pData = 0 THEN exit do
      if pData->projectIndex = -1 THEN continue do
      if pData->deleted = true THEN continue do
      if pData->id <> DB2_FUNCTION then continue do
      print #f, "   fileName: "; pData->fileName
      print #f, "   ElementName: "; pData->ElementName
      print #f, "   ElementValue: "; pData->ElementValue
      print #f, "   lineNum: "; pData->nLineNum
      print #f, ""
   loop
   print #f, ""
      
   ' PRINT ALL SUBS
   print #f, "SUBS"
   this.dbRewind()
   do 
      pData = this.dbGetNext
      if pData = 0 THEN exit do
      if pData->projectIndex = -1 THEN continue do
      if pData->deleted = true THEN continue do
      if pData->id <> DB2_SUB then continue do
      print #f, "   fileName: "; pData->fileName
      print #f, "   ElementName: "; pData->ElementName
      print #f, "   ElementValue: "; pData->ElementValue
      print #f, "   lineNum: "; pData->nLineNum
      print #f, ""
   loop
   print #f, ""

   ' PRINT ALL PROPERTIES
   print #f, "PROPERTIES"
   this.dbRewind()
   do 
      pData = this.dbGetNext
      if pData = 0 THEN exit do
      if pData->projectIndex = -1 THEN continue do
      if pData->deleted = true THEN continue do
      if pData->id <> DB2_PROPERTY then continue do
      print #f, "   fileName: "; pData->fileName
      print #f, "   ElementName: "; pData->ElementName
      print #f, "   ElementValue: "; pData->ElementValue
      print #f, "   lineNum: "; pData->nLineNum
      print #f, ""
   loop
   print #f, ""

   ' PRINT ALL VARIABLES
   print #f, "VARIABLES"
   this.dbRewind()
   do 
      pData = this.dbGetNext
      if pData = 0 THEN exit do
      if pData->projectIndex = -1 THEN continue do
      if pData->deleted = true THEN continue do
      if pData->id <> DB2_VARIABLE then continue do
      print #f, "   fileName: "; pData->fileName
      print #f, "   ElementName: "; pData->ElementName
      print #f, "   ElementValue: "; pData->ElementValue
      print #f, "   IsPrivate: "; pData->IsPrivate
      print #f, "   lineNum: "; pData->nLineNum
      print #f, ""
   loop
   print #f, ""
   
   ' PRINT ALL TODO
   print #f, "TODO DESCRIPTIONS"
   this.dbRewind()
   do 
      pData = this.dbGetNext
      if pData = 0 THEN exit do
      if pData->projectIndex = -1 THEN continue do
      if pData->deleted = true THEN continue do
      if pData->id <> DB2_TODO then continue do
      print #f, "   fileName: "; pData->fileName
      print #f, "   ElementName: "; pData->ElementName
      print #f, "   ElementValue: "; pData->ElementValue
      print #f, "   lineNum: "; pData->nLineNum
      print #f, ""
   loop

   close #f
   function = 0
end function


dim shared gdb2 as clsDB2


