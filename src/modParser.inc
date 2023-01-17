'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2023 Paul Squires, PlanetSquires Software
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

#include once "modParser.bi"
#include once "modGenerateCode.bi"
#include once "modVDProperties.bi"
#include once "modVDRoutines.bi"

'':::::
function ctxParser.GenerateVisualDesignerVariables( byval pDoc as clsDocument ptr ) as boolean
   
   dim pData as DB2_DATA ptr
   dim pCtrl as clsControl ptr

   dim as CWSTR wszFormName = GetFormName(pDoc)
   dim as CWSTR wszTypeName = wszFormName & "type"

'type frmMainType extends wfxForm
'   List1 As wfxListBox
'end type
   this.objectStartLine = 0
   this.objectEndLine = 0
   this.typeName = wszTypeName
   this.typeAlias = wszTypeName
   this.typeExtends = "wfxForm"
   gdb2.dbAdd( @this, DB2_TYPE )

'dim shared frmMain as frmMainType
   this.varName = wszFormName
   this.varType = wszTypeName
   this.varScope = DIMSCOPE.SCOPEGLOBAL
   this.functionName = ""
   gdb2.dbAdd( @this, DB2_VARIABLE )

'List1 As wfxListBox
   for i as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
      pCtrl = pDoc->Controls.ItemAt(i)
      if pCtrl then
         if pCtrl->ControlType = CTRL_FORM then continue for
         this.varName = GetControlProperty(pCtrl, "NAME")
         this.varType = GetWinformsXClassName(pCtrl->ControlType) 
         this.varScope = DIMSCOPE.SCOPETYPE
         this.functionName = wszTypeName
         gdb2.dbAdd( @this, DB2_VARIABLE )
      end if
   next   

   return true
end function


'':::::
function ctxParser.ParseTODO( byval action as PARSEACTION ) as boolean

   this.objectStartLine = this.lineNum
   this.objectEndLine = this.lineNum

   ' Get the remainder of the line after the TODO: statement
   if this.GetLine() = true then 
      gdb2.dbAdd( @this, DB2_TODO )
   end if
      
   return true
end function


'':::::
function ctxParser.ParseDIM( _
         byval action as PARSEACTION, _
         byval originFrom as DIMSCOPE _
         ) as boolean

   ' DIMSCOPE tells us from where this function is being called. The origin is important
   ' because when saving the DIM data to the database we need to the parent scope of
   ' the DIM statement.
   
   dim as string text
   dim as boolean parenActive 
   dim as integer numStatements = 0   ' the number of "as" clauses

   ' Get the tokens for this DIM statement up until the EOL
   do until this.GetToken() = false
      if this.EOL then exit do
      select case this.ltoken 
         case "shared"
            ' change the scope of this variable to global and do not add it to the statement  
            originFrom = DIMSCOPE.SCOPEGLOBAL
         case "ptr", "pointer", "byref", "any", "preserve"
            ' simply skip these tokens
         case "("
            parenActive = true   
         case ")"
            parenActive = false
         case "="
            ' We have hit an initializer. No need to continue to parse.
            this.ReadToEOL()
            continue do   
         case "as"
            if parenActive = false then
               text = text & this.ltoken & " " 
               numStatements += 1
            end if   
         case else
            if parenActive = false then
               if this.token = "," then
                  text = rtrim(text) & ","
               else
                  text = text & this.token & " "
               end if   
            end if   
      end select
   loop
   
   ' remove any trailing spaces
   text = rtrim(text)
   
   ' If we are parsing a DIM from global space then reset any sub/function data.
   if originFrom = DIMSCOPE.SCOPEGLOBAL then
      this.ResetFunctionValues()
   end if

   ' Dim Shared name1 As DataType, name2 As DataType
   ' There could be multiple DIMS on the same line. These would have more than one " as "
   ' clause and the statements will be separated by a comma.
   dim statements( 1 to numStatements ) as string
   if numStatements = 1 then
      statements(1) = text
   else
      for ii as long = 1 to numStatements
         statements(ii) = AfxStrParse( text, ii, "," )
      next
   end if   

   for ii as long = 1 to ubound(statements) 
      text = statements(ii)
      
      ' Normalize the line to move "AS TYPE" to front of the list.
      ' va1 as long
      ' as long var1, var2, var3
      ' var1 as long
      ' p as long
      dim as long i = instr(text, " as ")
      if i > 1 then
         text = mid(text, i + 1) & " " & left(text, i - 1)
      end if
      
      ' as long val
      ' as long var1,var2,var3
      ' The first space deliminated token should be "as" followed by the data type at
      ' position two, followed by one or more list of variables.
      dim as string dataType = AfxStrParse( text, 2, " " )

      dim as boolean bAddToDatabase = true
      
      ' TYPE variables are always added to the database regardless of whether they include
      ' standard datatypes. This is essential in order to display all of the Type elements
      ' in the autocomplete popup.
      if originFrom <> DIMSCOPE.SCOPETYPE then
         if this.IsStandardDataType( dataType ) then
            bAddToDatabase = false
         end if
      end if
      
      if bAddToDatabase then
         dim as string varName, varNames 
         i = InStrRev( text, " " )
         if i then varNames = mid( text, i + 1 )
         dim as long numItems = AfxStrParseCount( varNames, "," )
         for i as long = 1 to numItems
            this.varName = AfxStrParse( varNames, i, "," )
            this.varType = dataType
            this.varScope = originFrom
            if originFrom = DIMSCOPE.SCOPETYPE then
               this.functionName = this.typeName
            end if   
            gdb2.dbAdd( @this, DB2_VARIABLE )
         next   
      end if
      
   next
   
   return true
end function


'':::::
function ctxParser.ParseFunctionParams() as boolean

   dim as string text = this.functionParams
   dim as integer numStatements = 0   ' the number of "as" clauses

   ' There could be multiple DIMS on the same line. These would have more than one " as "
   ' clause and the statements will be separated by a comma.
   dim as long numParts = AfxStrParseCount(this.functionParams, ",")
   for ii as long = 1 to numParts
      text = trim(AfxStrParse(this.functionParams, ii, ","))
      text = AfxStrRemoveI(text, "ptr")    
      text = AfxStrRemoveI(text, "pointer")    
      text = AfxStrRemoveI(text, "byval")    
      text = AfxStrRemoveI(text, "byref")    
      text = AfxStrRemoveI(text, "ptr")    
      
      ' Isolate the variable name and variable type
      dim as string varName, dataType
      dim as long f = instr( ucase(text), " AS " )
      if f then
         varName = trim(left(text, f-1))
         dataType = trim(mid(text, f+4))
         if len(varName) andalso len(dataType) then
            if this.IsStandardDataType( dataType ) = false then
               this.varName = varName
               this.varType = dataType
               this.varScope = DIMSCOPE.SCOPEFUNCTION
               ' We have the full DIM statement so now commit it to the database
               gdb2.dbAdd( @this, DB2_VARIABLE )
            end if
         end if   
      end if
   next
   
   return true
end function


'':::::
function ctxParser.ParseENUM( byval action as PARSEACTION ) as boolean

   ' We store ENUM data the same way as TYPE data
   this.ResetFunctionValues()
   this.objectStartLine = this.lineNum
   this.objectEndLine = this.lineNum

   ' Get next expected token (ENUM name)
   if this.EOL then return false
   
   if this.GetToken() then
      ' bypass any ENUM's that do not have names
      if this.EOL then return false
      this.typeName = this.token
   end if

   ' Finally, look for the END ENUM that will complete the parsing
   do until this.GetToken() = false
      ' tokens from the body of the ENUM should be processed here

      select case this.ltoken 
         case "explicit", "="
            this.ReadToEOL()
         case "'todo:"
            this.ParseTODO( PARSINGACTION.PARSE_TODO )
         case "end"   
         case else
            if len(this.token) then
               ' Everything else in this structure must be an actual ENUM 
               this.varName = this.token
               this.varType = "long"
               this.varScope = DIMSCOPE.SCOPETYPE
               this.functionName = this.typeName
               gdb2.dbAdd( @this, DB2_VARIABLE )
            end if   
      end select
      
      if( this.ltoken = "end" ) then
         if this.GetToken() then
            if this.EOL then return false
            if( this.ltoken = "enum" ) then
               this.objectEndLine = this.lineNum
               this.typeAlias = this.typeName
               gdb2.dbAdd( @this, DB2_TYPE )

               ' Also add a variable so that the enum can be accessed directly
               ' without having to assign it first to a variable.
               this.varName = this.typeName
               this.varType = this.typeName
               this.varScope = DIMSCOPE.SCOPEGLOBAL
               this.functionName = ""
               gdb2.dbAdd( @this, DB2_VARIABLE )

               exit do
            end if
         end if      
      end if
   loop

   return true
end function


'':::::
function ctxParser.ParseTYPE( byval action as PARSEACTION ) as boolean

   this.ResetFunctionValues()
   this.objectStartLine = this.lineNum
   this.objectEndLine = this.lineNum

   dim as boolean bReadingExtends = false
   dim as string extendsTYPE 
   
   ' Get next expected token (typeName)
   if this.EOL then return false
   
   if this.GetToken() then
      this.typeName = this.token
   end if

   ' Finally, look for the END TYPE that will complete the parsing
   do until this.GetToken() = false
      ' tokens from the body of the TYPE should be processed here
      
      if bReadingExtends then
         extendsTYPE = this.token
         bReadingExtends = false   
      end if
      
      select case this.ltoken 
         case "extends"
            bReadingExtends = true
         case "as"
            this.ReadToSOL()
            this.ParseDim( action, DIMSCOPE.SCOPETYPE )
         case "'todo:"
            this.ParseTODO( PARSINGACTION.PARSE_TODO )
      end select
      
      if( this.ltoken = "end" ) then
         if this.GetToken() then
            if this.EOL then return false
            if( this.ltoken = "type" ) then
               this.objectEndLine = this.lineNum
               this.typeAlias = this.typeName
               this.typeExtends = extendsTYPE
               gdb2.dbAdd( @this, DB2_TYPE )
               exit do
            end if
         end if      
      end if
   loop

   return true
end function


'':::::
function ctxParser.ResetFunctionValues() as boolean
   this.functionName = ""
   this.functionAlias = ""
   this.functionParams = ""
   this.functionReturnType = ""
   this.objectStartLine = -1
   this.objectEndLine = -1
   this.varName = ""
   this.varType = ""
   this.typeName = ""
   this.typeAlias = ""
   this.typeExtends = ""
   this.varScope = SCOPEGLOBAL
   this.GetSet = ""
   return true
end function   


'':::::
function ctxParser.ParseFunction( byval action as PARSEACTION ) as boolean

   ' This handles parsing FUNCTION, SUB, PROPERTY, CONSTRUCTOR, DESTRUCTOR
   
   dim as boolean readingAllParams = false
   dim as boolean readingSingleParam = false
   dim as boolean readingArrayParam = false
      
   this.ResetFunctionValues()

   ' Get next expected token (functionName)
   if this.EOL then return false
   if this.GetToken() then
      this.functionName = this.token
      ' If there is an embedded "." then the function belongs to a TYPE/CLASS
      dim as long i = instr(this.functionName, ".")
      if i then 
         this.typeName = left(this.functionName, i - 1)
         ' Add a "THIS" variable
         dim temp as ctxParser
         temp.pDoc = this.pDoc
         temp.functionName = this.functionName
         temp.varName = "this"
         temp.varType = this.typeName
         temp.varScope = DIMSCOPE.SCOPEFUNCTION
         temp.nFileType = this.nFileType
         temp.objectStartLine = this.lineNum
         temp.objectEndLine = this.lineNum
         gdb2.dbAdd( @temp, DB2_VARIABLE )
      end if   
   end if

   this.objectStartLine = this.lineNum
   this.objectEndLine = this.lineNum

   select case action
      case PARSINGACTION.PARSE_PROPERTY
         ' Default that this is a Set property. This could be changed later to Get
         ' if a Property statement is found within the body.
         this.GetSet = "(set)"
      case PARSINGACTION.PARSE_CONSTRUCTOR
         this.GetSet = "(ctor)"
      case PARSINGACTION.PARSE_DESTRUCTOR
         this.GetSet = "(dtor)"
   end select

   ' Get next token (left parenthesis but is optional for sub/function)
   do until this.GetToken() = false
      if this.EOL then 
         if (action = PARSINGACTION.PARSE_SUB) orelse _
            (action = PARSINGACTION.PARSE_CONSTRUCTOR) orelse _
            (action = PARSINGACTION.PARSE_DESTRUCTOR) then
            exit do  ' only for Sub b/c for Function we haven't encountered AS RETURN_TYPE yet
         else
            ' Functions/Properties   
            this.UnwindToken
         end if   
      end if

      select case this.ltoken 
         case "("    ' start of the list of parameters
            if readingSingleParam then
               ' This is start of an Array() parameter within the list so save it.
               this.functionParams = this.functionParams & "("
               readingArrayParam = true
               continue do
            end if
            readingAllParams = true
            readingSingleParam = true
         case ")"    ' end of the list of parameters
            if readingArrayParam then
               ' This is end of an Array() parameter within the list so save it.
               this.functionParams = this.functionParams & ")"
               readingArrayParam = false
               continue do
            end if
            readingAllParams = false
            readingSingleParam = false
            exit do
         case "="
            ' if reading a parameter then need to bypass any initializer
            if readingSingleParam then
               ' Need to continue until we reach a comma that separates the parameters
               ' or until we hit the closing parenthesis.
               readingSingleParam = false
            end if
         case else
            ' Get the parameter tokens until we get (right parenthesis)
            if readingAllParams andalso readingSingleParam then
               this.functionParams = this.functionParams & iif(this.token = ",", ",", " " & this.token)
               if this.token = "," then readingSingleParam = true
            else
               ' we are not reading params so maybe we have encountered the start of the RETURN_TYPE
               select case this.ltoken 
                  case "byref", "as"
                     this.UnwindToken
                     exit do
               end select      
            end if
      end select
   loop
   
   if (action = PARSINGACTION.PARSE_FUNCTION) orelse _
      (action = PARSINGACTION.PARSE_PROPERTY) then
      ' Look for the Function RETURN_TYPE
      do until this.GetToken() = false
         if this.EOL then exit do
         ' This should pickup any "byref" "as" RETURN_TYPE
         this.functionReturnType = this.functionReturnType & this.token & " "
      loop
   end if

   ' Parse any function parameters to add to DIM variables. We use a separate
   ' so that we don't interfere with the regular GetToken() parsing of the
   ' main code.
   if len(this.functionParams) then 
      this.ParseFunctionParams()
   end if

   ' Finally, look for the End Sub/Function that will complete the parsing
   do until this.GetToken() = false
      ' tokens from the body of the sub/function should be processed here
      
       select case this.ltoken 
         case "dim", "static"
            this.ParseDim( action, DIMSCOPE.SCOPEFUNCTION )
         case "'todo:"
            this.ParseTODO( PARSINGACTION.PARSE_TODO )
         case "property", "return"
            if action = PARSINGACTION.PARSE_PROPERTY then
               ' A Property statement found at this point would be a return statement
               ' indicating that this would be a (get) property.
               this.GetSet = "(get)"
            end if
      end select

      if( this.ltoken = "end" ) then
         if this.GetToken() then
            if this.EOL then return false
            if( this.ltoken = "sub" ) orelse _
               (this.ltoken = "function") orelse _
               (this.ltoken = "property") orelse _
               (this.ltoken = "constructor") orelse _
               (this.ltoken = "destructor") then
               this.objectEndLine = this.lineNum
               exit do
            end if
         end if      
      end if
   loop
      
   ' Construct the CallTip that we store in functionParams that dbAdd 
   ' will use when storing to the database.
   this.functionParams = this.functionName & "(" & this.functionParams & ") " & this.functionReturnType  
   gdb2.dbAdd( @this, DB2_FUNCTION )

   return true
end function


'':::::
function ctxParser.UnwindToken() as boolean
   this.s -= len(this.token)
   this.i = this.s
   return true
end function


'':::::
function ctxParser.IsStandardDataType( byref sVarType as string ) as boolean
   ' Determine if the variable is one of the predefined standard data types
   static as string sList
   sList = " boolean byte double integer long longint short single" & _
           " string ubyte uinteger ushort wstring zstring cwstr "
   if instr(sList, " " & lcase(sVarType) & " " ) then return true
   return false  
end function


'':::::
function ctxParser.PeekChar _
   ( _
      byval x as integer = 0 _
   ) as integer

   if( this.i + x < this.n ) then
      return cast( ubyte ptr, this.text )[this.i + x]
   end if

   return 0
   
end function 


'':::::
function ctxParser.ReadChar() as integer

   if( this.i < this.n ) then
      this.i += 1
      return cast( ubyte ptr, this.text )[this.i - 1]
   end if

   return 0
   
end function 


'':::::
function ctxParser.ReadToEOL() as boolean

   dim as integer c

   c = this.PeekChar()
   while( c )
      c = this.PeekChar()
      if( c = 10 or c = 13 ) then
         exit while
      elseif( c = asc("'") ) then
         c = this.ReadChar()
         if( this.incomment ) then
            if( this.PeekChar() = asc("/") ) then
               c = ReadChar()
               this.incomment = false
               exit while
            end if
         end if
      else
         c = this.ReadChar()
      end if
   wend

   return true

end function


'':::::
function ctxParser.ReadToSOL() as boolean
   ' Rewind to start of line. Used when a token is found part way through a 
   ' line but parsing it depends on tokens that would have occured earlier
   ' in the line.
   ' For example, a dim within a TYPE structure.
   '      firstName as string
   '
   dim as integer curPos = this.s
   for ii as integer = curPos to 0 step - 1
      if( (this.text[ii] = 10) or (this.text[ii] = 13) ) then
         exit for
      end if
      this.s = ii: this.i = ii
   next
   this.startofline = true
   
   return true

end function


'':::::
function ctxParser.ReadQuoted _
   ( _
      byval escapedonce as boolean = FALSE _
   ) as boolean

   dim as integer c

   c = this.PeekChar()
   if( c <> 34 ) then
      return false
   end if
   
   '' Expects Opening Quote
   c = ReadChar()

   if( escapedonce or this.escaped ) then

      while( this.i < this.n )
         c = this.ReadChar()
         if( c = asc("""") ) then
            c = this.PeekChar()
            if( c = asc("""") ) then
               c = this.ReadChar()
            else
               exit while
            end if
         elseif( c = asc("\") ) then
            c = this.PeekChar()
            if( c = 10 or c = 13 ) then
               exit while
            else
               c = this.ReadChar()
            end if
         end if
      wend 

   else
      
      while( this.i < this.n )
         c = this.ReadChar()
         if( c = asc("""") ) then
            if( this.PeekChar() = asc("""") ) then
               c = this.ReadChar()
            else
               exit while
            end if
         endif
      wend

   end if

   return true

end function


'':::::
function ctxParser.GetLine() as boolean
   this.ReadToEOL()
   this.fullLine = mid( *this.text, this.s + 1, this.i - this.s)
   this.s = this.i
   return true
end function


'':::::
function ctxParser.GetToken() as boolean

   while ( this.s < this.n )
   
   dim as integer c
   this.EOL = false

   '' newline ?
   c = this.PeekChar()
   if( c = 13 or c = 10 ) then
      c = this.ReadChar()
      if( c = 13 ) then
         c = this.PeekChar()
         if( c = 10 ) then
            c = this.ReadChar()
         end if
      end if
      
      this.EOL = true

      ' If line continuation was set then we can not set the
      ' statement terminated marker yet.
      if( this.ltoken = "_" ) then
         this.EOL = false
      end if

      this.lineNum += 1
      this.startofline = true
      this.inprepro = false
      this.s = this.i
      
      if this.EOL then 
         this.token = ""
         this.ltoken = ""
         return true
      else   
         continue while
      end if   
   end if

   if( this.incomment ) then
      this.ReadToEOL()
      this.s = this.i
      continue while
   end if

   '' tab | space ?
   c = this.PeekChar()
   if( c = 9 or c = 32 ) then
      c = this.ReadChar()
      while ( c )
         c = this.PeekChar()
         if( c = 9 or c = 32 )then
            c = this.ReadChar()
         else
            this.startofline = false
            exit while
         end if
      wend
      this.s = this.i
      continue while
   end if

   '' '
   c = this.PeekChar()
   if( c = asc("'") ) then
      c = this.ReadChar()

      ' We are in a Comment but we need to check if the comment starts
      ' with a TODO: statement. If yes, then we need to return that token
      c = this.PeekChar()
      if( c = asc("t") or c = asc("T") ) then
         c = this.ReadChar()
         while( c )
            c = this.PeekChar()
            select case c
               'case "O", "o", "D", "d", ":"
               case 79, 111, 68, 100, 58
                  c = this.ReadChar()
               case else
                  exit while
            end select
         wend
         this.token = mid( *this.text, this.s + 1, this.i - this.s)
         this.ltoken = lcase(this.token)
         this.s = this.i
         if this.ltoken = "'todo:" then
            return true
         end if
      end if
         
      ' Reach this point then this is just a normal comment line 
      ' to continue to the end of the line.
      this.ReadToEOL()
      this.s = this.i
      continue while
   end if

   '' #
   c = this.PeekChar()
   if( c = asc("#") and this.startofline ) then
      c = this.ReadChar()
      while( c )
         c = this.PeekChar()
         if( ( c = asc("_") ) or _
             ( c >= asc("A") and c <= asc("Z") ) or _
             ( c >= asc("a") and c <= asc("z") ) _
         ) then
            c = this.ReadChar()
         else
            exit while
         end if
      wend 
      this.inprepro = true
      this.s = this.i
      continue while
   end if

   '' /'
   c = this.PeekChar()
   if( c = asc("/") ) then
      if( this.PeekChar(1) = asc("'") ) then
         c = this.ReadChar()
         c = this.ReadChar()
         this.incomment = TRUE
         this.ReadToEOL()
         this.s = this.i
         continue while
      end if
   end if

   '' $"..." 
   c = this.PeekChar()
   if( c = asc("$") ) then
      if( this.PeekChar(1) = asc("""") ) then
         c = this.ReadChar() '' $
         if( this.ReadQuoted(true) ) then
            this.s = this.i
            continue while
         end if
      end if
   end if

   '' "
   c = this.PeekChar()
   if( c = asc("""") ) then
      if( this.ReadQuoted(false) ) then
         this.s = this.i
         continue while
      end if
   end if

   '' Special characters that we want to process
   c = this.PeekChar()
   if( c = asc("(") or _
       c = asc(")") or _
       c = asc("=") or _
       c = asc(",") _
      ) then

      c = this.ReadChar()

      this.token = chr(c)
      this.ltoken = this.token
      this.s = this.i

      return true
   end if

   '' Name | keyword
   c = this.PeekChar()
   if( c = asc("_") or _
      ( c >= asc("A") and c <= asc("Z") ) or _
       ( c >= asc("a") and c <= asc("z") ) _
      ) then

      c = this.ReadChar()
      while( c )
         c = this.PeekChar()
         if( ( c = asc("_") ) or _
             ( c >= asc("A") and c <= asc("Z") ) or _
             ( c >= asc("a") and c <= asc("z") ) or _
             ( c >= asc("0") and c <= asc("9") ) or _
             ( c = asc(".") ) _
            ) then

            c = this.ReadChar()
         else
            '_ReadSuffix( ctx )
            exit while
         end if
      wend

      this.token = mid( *this.text, this.s + 1, this.i - this.s)
      this.ltoken = lcase(this.token)

      ' Bypass REM statements and Line Continuation underscore character
      if (this.ltoken = "declare") orelse _
         (this.ltoken = "rem") orelse _
         (this.ltoken = "_") then
         this.ReadToEOL()
         this.s = this.i
         continue while
      end if

      this.s = this.i

      return true

   end if

   c = this.ReadChar()

   this.s = this.i
   wend
   
   return false

end function


'':::::
function ctxParser.Parse( byval pDoc as clsDocument ptr ) as boolean

   if pDoc = 0 then return false
   
   this.pDoc = pDoc
   
   this.text = Cast( ZString Ptr, SciExec(pDoc->hWindow(0), SCI_GETCHARACTERPOINTER, 0, 0) )
   if( this.text = 0 ) then return false

   this.incomment = false
   this.startofline = true
   this.lineNum = 0
   this.token = ""
   this.ltoken = ""
   this.fullline = ""
   this.nfileType = DB2_FILETYPE_USERCODE
   this.s = 0
   this.i = 0
   this.n = len(*this.text)

   do until this.GetToken() = false 

      select case this.ltoken
         case "dim", "redim"
            if this.ParseDIM( PARSINGACTION.PARSE_DIM, DIMSCOPE.SCOPEGLOBAL ) = false then
               '? this.lineNum, "Error parsing DIM statement. Unexpected token: "; this.token
            end if

         case "sub"
            if this.ParseFunction( PARSINGACTION.PARSE_SUB ) = false then
               '? this.lineNum, "Error parsing SUB statement. Unexpected token: "; this.token
            end if

         case "function"
            if this.ParseFunction( PARSINGACTION.PARSE_FUNCTION ) = false then
               '? this.lineNum, "Error parsing FUNCTION statement. Unexpected token: "; this.token
            end if

         case "property"
            if this.ParseFunction( PARSINGACTION.PARSE_PROPERTY ) = false then
               '? this.lineNum, "Error parsing PROPERTY statement. Unexpected token: "; this.token
            end if

         case "constructor"
            if this.ParseFunction( PARSINGACTION.PARSE_CONSTRUCTOR ) = false then
               '? this.lineNum, "Error parsing CONSTRUCTORE statement. Unexpected token: "; this.token
            end if

         case "destructor"
            if this.ParseFunction( PARSINGACTION.PARSE_DESTRUCTOR ) = false then
               '? this.lineNum, "Error parsing DESTRUCTOR statement. Unexpected token: "; this.token
            end if

         case "type"
            if this.ParseTYPE( PARSINGACTION.PARSE_TYPE ) = false then
               '? this.lineNum, "Error parsing TYPE statement. Unexpected token: "; this.token
            end if

         case "enum"
            if this.ParseENUM( PARSINGACTION.PARSE_ENUM ) = false then
               '? this.lineNum, "Error parsing ENUM statement. Unexpected token: "; this.token
            end if

         case "'todo:"
            if this.ParseTODO( PARSINGACTION.PARSE_TODO ) = false then
               '? this.lineNum, "Error parsing TODO statement. Unexpected token: "; this.token
            end if

      end select

   loop

   this.text = 0
   
   ' The main code source file has now been parsed. If this is a visual designer form then
   ' we also need to add all of the Form and Control TYPE's and variables.
   if pDoc->IsDesigner then
      this.GenerateVisualDesignerVariables( pDoc )
   end if

   return true

end function


   