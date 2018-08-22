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
''  Attempt to insert the second half of a syntax construct. for example,
''  The END FUNCTION, LOOP, NEXT, ...  This is activated when ENTER is pressed
''  in the code editor and gConfig. AutoComplete is TRUE.
''

enum
   BLOCK_STATEMENT_IF 
   BLOCK_STATEMENT_FOR
   BLOCK_STATEMENT_SELECT
   BLOCK_STATEMENT_WHILE
   BLOCK_STATEMENT_DO
   BLOCK_STATEMENT_FUNCTION
   BLOCK_STATEMENT_SUB
   BLOCK_STATEMENT_PROPERTY
   BLOCK_STATEMENT_TYPE
   BLOCK_STATEMENT_WITH
   BLOCK_STATEMENT_ENUM
   BLOCK_STATEMENT_UNION
   BLOCK_STATEMENT_CONSTRUCTOR
   BLOCK_STATEMENT_DESTRUCTOR
end enum

''
''
function CanCompleteBlockStatement( byval pDoc as clsDocument ptr, _
                                    byval idBlockType as long, _
                                    byval nStartLine as long _
                                    ) as boolean
   
   If gConfig.AutoComplete = false Then return false
   
   dim as long NumLines, nLenMatch
   dim as string sStartMatch, sEndMatch, sLine
   
   dim as hwnd hEdit = pDoc->hWndActiveScintilla

   ' Add a unique terminating delimiter to the phrase in order to make it
   ' easier to match lines later that only contain the matching end phrase
   ' and no following space character (ie. they exist alone on the line).
   ' The algorithm needs to search for a matching end phrase but if it first
   ' finds a matching start phrase for the block type 
   select case idBlockType
      case BLOCK_STATEMENT_IF:          sStartMatch = "IF ":     sEndMatch = "END IF|"
      case BLOCK_STATEMENT_FOR:         sStartMatch = "FOR ":    sEndMatch = "NEXT|"
      case BLOCK_STATEMENT_SELECT:      sStartMatch = "SELECT ": sEndMatch = "END SELECT|"
      case BLOCK_STATEMENT_WHILE:       sStartMatch = "WHILE ":  sEndMatch = "WEND|"
      case BLOCK_STATEMENT_DO:          sStartMatch = "DO ":     sEndMatch = "LOOP|"
      case BLOCK_STATEMENT_FUNCTION:    sStartMatch = "":        sEndMatch = "END FUNCTION|"
      case BLOCK_STATEMENT_SUB:         sStartMatch = "":        sEndMatch = "END SUB|"
      case BLOCK_STATEMENT_PROPERTY:    sStartMatch = "":        sEndMatch = "END PROPERTY|"
      case BLOCK_STATEMENT_TYPE:        sStartMatch = "TYPE ":   sEndMatch = "END TYPE|"
      case BLOCK_STATEMENT_WITH:        sStartMatch = "WITH ":   sEndMatch = "END WITH|"
      case BLOCK_STATEMENT_ENUM:        sStartMatch = "ENUM ":   sEndMatch = "END ENUM|"
      case BLOCK_STATEMENT_UNION:       sStartMatch = "UNION ":  sEndMatch = "END UNION|"
      case BLOCK_STATEMENT_CONSTRUCTOR: sStartMatch = "":        sEndMatch = "END CONSTRUCTOR|"
      case BLOCK_STATEMENT_DESTRUCTOR:  sStartMatch = "":        sEndMatch = "END DESTRUCTOR|"
   end select
   
   NumLines = SciExec( hEdit, SCI_GETLINECOUNT, 0, 0) 
   nStartLine = nStartLine + 1
   
   for i as long = nStartLine to NumLines - 1
      ' Does this line start with the matching phrase we are looking for?
      sLine = ucase(ltrim(pDoc->GetLine(i)))
      
      nLenMatch = len(sStartMatch)
      if nLenMatch then
         if sStartMatch = mid(sLine, 1, nLenMatch) then
            return true    ' a start block was found before a matching end was found.
         END IF
      end if
      
      nLenMatch = len(sEndMatch)
      if nLenMatch then
         if sEndMatch = mid(sLine, 1, nLenMatch-1) & "|" then
            return false    ' a matching end to the block statement already exists
         END IF
      end if
      
      ' Stop search if we find the start of another block statement structure and
      ' allow the insert.
      if (idBlockType <> BLOCK_STATEMENT_FUNCTION) and (idBlockType <> BLOCK_STATEMENT_SUB) then
         if left(sLine, 3)  = "IF "        then return true
         if left(sLine, 4)  = "FOR "       then return true
         if left(sLine, 3)  = "DO "        then return true
         if left(sLine, 7)  = "SELECT "    then return true
         if left(sLine, 6)  = "WHILE "     then return true
      end if
      if left(sLine, 9) = "FUNCTION " then
         if instr(10, sLine, "=") = 0 then return true
      end if
      if left(sLine, 4)  = "SUB "       then return true
      if left(sLine, 17) = "PRIVATE FUNCTION " then return true
      if left(sLine, 16) = "PUBLIC FUNCTION "  then return true
      if left(sLine, 12) = "PRIVATE SUB "  then return true
      if left(sLine, 11) = "PUBLIC SUB "   then return true
      if left(sLine, 9)  = "PROPERTY "     then return true
      if left(sLine, 6)  = "UNION "        then return true
      if left(sLine, 5)  = "ENUM "         then return true
   next
      
   ' Default that we can allow the insert
   function = true
end function


''
''
function AttemptAutoInsert() as Long
   ' Attempts to Autocomplete a FOR/DO/SELECT, etc block. This function also deals
   ' with AutoIndentation so need to handle both separation depending on the user
   ' chosen settings.
   
   Dim as HWND hEdit
   Dim as long nLine, curPos, LineLen, nFoldLevel, TabSize, nSpaces, IndentSize, i
   Dim as string strFill, strTemp, strCurLine, strPrevLine

   Dim pDoc As clsDocument Ptr 

   pDoc = gTTabCtl.GetActiveDocumentPtr()
   If pDoc = 0 Then exit function

   hEdit  = pDoc->hWndActiveScintilla

   ' Current and previous lines
   curPos      = SciExec(hEdit, SCI_GETCURRENTPOS, 0, 0)
   nLine       = pDoc->GetCurrentLineNumber
   strCurLine  = pDoc->GetLine(nLine)
   strPrevLine = pDoc->GetLine(nLine-1)
   if len(strPrevLine) = 0 then exit function

   ' Get the tab width and indent size
   TabSize     = SciExec(hEdit, SCI_GETTABWIDTH, 0, 0)
   IndentSize  = SciExec(hEdit, SCI_GETINDENT, 0, 0)
   
   ' Calculate the number of spaces to fill on the left
   For i = 1 To Len(strPrevLine)
      If Mid(strPrevLine, i, 1) <> " " Then
         If Mid(strPrevLine, i, 1) = Chr(9) Then
            nSpaces = nSpaces + TabSize
         Else
            Exit For
         End If
      Else
         nSpaces = nSpaces + 1
      End If
   Next
   strPrevLine = Trim(Ucase(strPrevLine), Any Chr(32, 9))
   
   If gConfig.AutoIndentation = false Then
      nSpaces = 0: IndentSize = 0: TabSize = 0
   END IF

   ''''''''''
   ' IF/THEN
   '   Before autoindenting an IF statement make sure that this
   '   is in fact a multiline IF statement.
   If (Left(strPrevLine, 3) = "IF " And Right(strPrevLine, 5) = " THEN") then
      ' Remove the current line because we will add it again below
      ' once we have wrapped it in an END IF. It will include any chunk of
      ' text that was after the THEN in a single line IF/THEN
      SciExec(hEdit, SCI_LINEDELETE, 0, 0)
      strFill = Space(nSpaces + IndentSize) & ltrim(strCurline, any chr(32,9)) & vbcrlf
      if CanCompleteBlockStatement(pDoc, BLOCK_STATEMENT_IF, nLine) then
         strFill = strFill & SPACE(nSpaces) & "end if" & vbcrlf
      end if          
      SciExec(hEdit, SCI_ADDTEXT, Len(strFill), Strptr(strFill))
      curPos = curPos + nSpaces + IndentSize
      SciExec(hEdit, SCI_SETSEL, curPos, curPos)
      exit function
   end if
    
   '''''''''''''
   ' FOR/NEXT
   If Left(strPrevLine, 4) = "FOR " then
      strFill = Space(nSpaces + IndentSize) 
      if CanCompleteBlockStatement(pDoc, BLOCK_STATEMENT_FOR, nLine) then
         strFill = strFill & vbcrlf & SPACE(nSpaces) & "next" & vbcrlf
      end if          
      SciExec(hEdit, SCI_ADDTEXT, Len(strFill), Strptr(strFill))
      curPos = curPos + nSpaces + IndentSize 
      SciExec(hEdit, SCI_SETSEL, curPos, curPos)
      exit function
   end if

   '''''''''''''
   ' SELECT CASE
   If Left(strPrevLine, 12) = "SELECT CASE " then
      strFill = Space(nSpaces + IndentSize)  
      if CanCompleteBlockStatement(pDoc, BLOCK_STATEMENT_SELECT, nLine) then
         strFill = strFill & "case " & vbcrlf & SPACE(nSpaces) & "end select" & vbcrlf
      end if          
      SciExec(hEdit, SCI_ADDTEXT, Len(strFill), Strptr(strFill))
      curPos = curPos + nSpaces + IndentSize + 5
      SciExec(hEdit, SCI_SETSEL, curPos, curPos)
      exit function
   end if

   '''''''''''''
   ' WHILE/WEND
   If (Left(strPrevLine, 6) = "WHILE ") or (strPrevLine = "WHILE") then
      strFill = Space(nSpaces + IndentSize) 
      if CanCompleteBlockStatement(pDoc, BLOCK_STATEMENT_WHILE, nLine) then
         strFill = strFill & vbcrlf & SPACE(nSpaces) & "wend" & vbcrlf
      end if          
      SciExec(hEdit, SCI_ADDTEXT, Len(strFill), Strptr(strFill))
      curPos = curPos + nSpaces + IndentSize 
      SciExec(hEdit, SCI_SETSEL, curPos, curPos)
      exit function
   end if

   '''''''''''''
   ' DO/LOOP
   If (Left(strPrevLine, 3) = "DO ") or (strPrevLine = "DO") then
      strFill = Space(nSpaces + IndentSize) 
      if CanCompleteBlockStatement(pDoc, BLOCK_STATEMENT_DO, nLine) then
         strFill = strFill & vbcrlf & SPACE(nSpaces) & "loop" & vbcrlf
      end if          
      SciExec(hEdit, SCI_ADDTEXT, Len(strFill), Strptr(strFill))
      curPos = curPos + nSpaces + IndentSize 
      SciExec(hEdit, SCI_SETSEL, curPos, curPos)
      exit function
   end if

   '''''''''''''
   ' FUNCTION/END FUNCTION
   If (Left(strPrevLine, 9)  = "FUNCTION ") or _
      (Left(strPrevLine, 17) = "PRIVATE FUNCTION ") or _
      (Left(strPrevLine, 16) = "PUBLIC FUNCTION ") then
      ' Determine if this is a FUNCTION = statement rather than a true function.
      strTemp = strPrevLine
      i = instr(strTemp, "(")
      if i then strTemp = left(strTemp, i-1)
      if instr(strTemp, "=") then exit function
      strFill = Space(nSpaces + IndentSize)
      if CanCompleteBlockStatement(pDoc, BLOCK_STATEMENT_FUNCTION, nLine) then
         strFill = strFill  & vbcrlf & SPACE(nSpaces) & "end function" & vbcrlf
      end if          
      SciExec(hEdit, SCI_ADDTEXT, Len(strFill), Strptr(strFill))
      curPos = curPos + nSpaces + IndentSize 
      SciExec(hEdit, SCI_SETSEL, curPos, curPos)
      exit function
   end if

   '''''''''''''
   ' SUB/END SUB
   If (Left(strPrevLine, 4)  = "SUB ") or _
      (Left(strPrevLine, 12) = "PRIVATE SUB ") or _
      (Left(strPrevLine, 11) = "PUBLIC SUB") then
      strFill = Space(nSpaces + IndentSize)
      if CanCompleteBlockStatement(pDoc, BLOCK_STATEMENT_SUB, nLine) then
         strFill = strFill  & vbcrlf & SPACE(nSpaces) & "end sub" & vbcrlf
      end if          
      SciExec(hEdit, SCI_ADDTEXT, Len(strFill), Strptr(strFill))
      curPos = curPos + nSpaces + IndentSize 
      SciExec(hEdit, SCI_SETSEL, curPos, curPos)
      exit function
   end if

   '''''''''''''
   ' PROPERTY/END PROPERTY
   If (Left(strPrevLine, 9)  = "PROPERTY ") then
      strFill = Space(nSpaces + IndentSize)
      if CanCompleteBlockStatement(pDoc, BLOCK_STATEMENT_SUB, nLine) then
         strFill = strFill  & vbcrlf & SPACE(nSpaces) & "end property" & vbcrlf
      end if          
      SciExec(hEdit, SCI_ADDTEXT, Len(strFill), Strptr(strFill))
      curPos = curPos + nSpaces + IndentSize 
      SciExec(hEdit, SCI_SETSEL, curPos, curPos)
      exit function
   end if

   '''''''''''''
   ' CONSTRUCTOR/END CONSTRUCTOR
   If (Left(strPrevLine, 12)  = "CONSTRUCTOR ") then
      strFill = Space(nSpaces + IndentSize) 
      if CanCompleteBlockStatement(pDoc, BLOCK_STATEMENT_CONSTRUCTOR, nLine) then
         strFill = strFill & vbcrlf & SPACE(nSpaces) & "end constructor" & vbcrlf
      end if          
      SciExec(hEdit, SCI_ADDTEXT, Len(strFill), Strptr(strFill))
      curPos = curPos + nSpaces + IndentSize 
      SciExec(hEdit, SCI_SETSEL, curPos, curPos)
      exit function
   end if

   '''''''''''''
   ' DESTRUCTOR/END DESTRUCTOR
   If (Left(strPrevLine, 11)  = "DESTRUCTOR ") then
      strFill = Space(nSpaces + IndentSize) 
      if CanCompleteBlockStatement(pDoc, BLOCK_STATEMENT_DESTRUCTOR, nLine) then
         strFill = strFill & vbcrlf & SPACE(nSpaces) & "end destructor" & vbcrlf
      end if          
      SciExec(hEdit, SCI_ADDTEXT, Len(strFill), Strptr(strFill))
      curPos = curPos + nSpaces + IndentSize 
      SciExec(hEdit, SCI_SETSEL, curPos, curPos)
      exit function
   end if

   '''''''''''''
   ' TYPE/END TYPE
   If (Left(strPrevLine, 5)  = "TYPE ") then
      ' Determine if this is a TYPE = statement rather than a true TYPE structure
      if instr(ucase(strPrevLine), " AS ") then exit function
      strFill = Space(nSpaces + IndentSize) 
      if CanCompleteBlockStatement(pDoc, BLOCK_STATEMENT_TYPE, nLine) then
         strFill = strFill & vbcrlf & SPACE(nSpaces) & "end type" & vbcrlf
      end if          
      SciExec(hEdit, SCI_ADDTEXT, Len(strFill), Strptr(strFill))
      curPos = curPos + nSpaces + IndentSize 
      SciExec(hEdit, SCI_SETSEL, curPos, curPos)
      exit function
   end if

   '''''''''''''
   ' WITH/END WITH
   If (Left(strPrevLine, 4)  = "WITH") then
      strFill = Space(nSpaces + IndentSize) 
      if CanCompleteBlockStatement(pDoc, BLOCK_STATEMENT_WITH, nLine) then
         strFill = strFill & vbcrlf & SPACE(nSpaces) & "end with" & vbcrlf
      end if          
      SciExec(hEdit, SCI_ADDTEXT, Len(strFill), Strptr(strFill))
      curPos = curPos + nSpaces + IndentSize 
      SciExec(hEdit, SCI_SETSEL, curPos, curPos)
      exit function
   end if

   '''''''''''''
   ' ENUM/END ENUM
   If (Left(strPrevLine, 4)  = "ENUM") then
      strFill = Space(nSpaces + IndentSize) 
      if CanCompleteBlockStatement(pDoc, BLOCK_STATEMENT_ENUM, nLine) then
         strFill = strFill & vbcrlf & SPACE(nSpaces) & "end enum" & vbcrlf
      end if          
      SciExec(hEdit, SCI_ADDTEXT, Len(strFill), Strptr(strFill))
      curPos = curPos + nSpaces + IndentSize 
      SciExec(hEdit, SCI_SETSEL, curPos, curPos)
      exit function
   end if

   '''''''''''''
   ' UNION/END UNION
   If (Left(strPrevLine, 6)  = "UNION ") then
      strFill = Space(nSpaces + IndentSize) 
      if CanCompleteBlockStatement(pDoc, BLOCK_STATEMENT_UNION, nLine) then
         strFill = strFill & vbcrlf & SPACE(nSpaces) & "end union" & vbcrlf
      end if          
      SciExec(hEdit, SCI_ADDTEXT, Len(strFill), Strptr(strFill))
      curPos = curPos + nSpaces + IndentSize 
      SciExec(hEdit, SCI_SETSEL, curPos, curPos)
      exit function
   end if

   '''''''''''''
   ' ELSE/CASE
   If Left(strPrevLine, 4) = "ELSE" Or _
      Left(strPrevLine, 5) = "CASE " then
      ' Add spaces to indent the line
      strFill = Space(nSpaces + IndentSize)
   Else
       ' Add the same spaces on the left that the line above
       strFill = Space(nSpaces)
   End If
   ' Indents the line
   SciExec(hEdit, SCI_ADDTEXT, Len(strFill), Strptr(strFill))

   function = 0
end function

