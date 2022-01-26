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


''  PARSER MODULE (used in conjunction with clsParser.inc)

#include once "modParser.bi"


' ========================================================================================
' Change everything between quotes to # symbols
' ========================================================================================
public function MaskStringCharacters( byval st as string) as string
   ' Iterate the line and change everything between quotes to # symbols. This
   ' ensures that we correctly deal with strings that have embedded single 
   ' quote characters.
   dim as long i 
   dim as Boolean bInString = false
   for i = 0 to len(st) - 1
      if st[i] = 34 THEN bInString = not(bInString)
      if bInString THEN  
         if st[i] <> 34 THEN st[i] = 35   ' # symbol
      END IF
   NEXT
   function = st
end function


' ========================================================================================
' Removes a single line comment.
' ========================================================================================
public function RemoveComments( byval st as string) as string
   function = AfxStrExtract( 1, st, "'")
end function  


' ========================================================================================
' Parse the incoming pDoc clsDocument.
' ========================================================================================
public Function ParseDocument( byval pDoc as clsDocument ptr ) As Long

   If pDoc = 0 Then exit function
   if pDoc->bNeedsParsing = false then exit function
   if gApp.IsProjectCacheLoaded = true then 
      pDoc->bNeedsParsing = false
      exit function
   end if

   dim as long nLineCount, nLineNum, nPos, nStartParseLine, numBufferArrayLines
   Dim As String st, st_ucase, sLine, sTempFileBuffer

   dim wszFilename as wstring * MAX_PATH
   wszFilename = pDoc->DiskFilename 

   ' The incoming document is already loaded into a clsDocument so we can simply 
   ' use that pDoc to retrieve the lines. We can use that approach for the actual
   ' editor lines but in order to parse any Form CodeGen lines then we need to 
   ' load them into an array.
   
   ' Dynamic string array to hold all lines loaded from the file. It
   ' is resized at reasonable capacity chunks in order to avoid
   ' too many memory reallocations if constantly redim'd. 
   ' These hold the ParsingCodeGen lines
   redim sLineArray( any ) as string 
 
   ' Array to hold the lines read in from a document that exists in the project
   ' but does not yet have the text loaded. Array access faster than Parse. 
   redim sTempFileBufferArray( any ) as string 
   
   ' Create a parse TYPE that handles the state information throughout parsing this document.
   dim parser as clsParser
   
   with parser
      .action    = ACTION_NONE
      .fileName  = wszFilename
      .nFileType = DB2_FILETYPE_USERCODE
      ' Flag whether we are parsing the codegen code. We do not want to save any
      ' of the Form Types sub/function/ctor/dtor because those will then display
      ' in the Explorer Treeview. The user would click on them expecting to be
      ' able to then see them in the editor. We can't allow that.
      .bParsingCodeGen = false
   end with

   ' Test to ensure that the document exists. If it does not then load the text now.
   if pDoc->GetActiveScintillaPtr = 0 then
      ' The document has not been loaded into a Scintilla window
      GetFileToString( pDoc->DiskFilename, sTempFileBuffer, pDoc )
      numBufferArrayLines = GetStringToArray( sTempFileBuffer, sTempFileBufferArray() )
   end if
   
   do
      if (gApp.IsProjectLoading) orelse (gApp.IsFileLoading) then
         gApp.wszPanelText = wszFilename
         gApp.hIconPanel = 0
         frmMain_SetStatusbar
      END IF

        
      if parser.bParsingCodeGen then
         ' We put the codegen code lines after the scintilla lines because this
         ' ensures that if nStartParseLine is > 0 then the codegen code will get
         ' reparsed when the main scintilla document is modified. Otherwise, 
         ' we'd have to gdb2.dbDelete all entries in the database and have to
         ' reparse the whole file when at times we only need to parse from the
         ' line currently being modified onwards to the end of the file.
         nLineCount = ubound( sLineArray )
         nStartParseLine = lbound( sLineArray )
      else
         if pDoc->GetActiveScintillaPtr = 0 then
            nLineCount = numBufferArrayLines 
            nStartParseLine = 1
         else
            nLineCount = pDoc->GetLineCount - 1  
            nStartParseLine = 0
         end if
      
         ' Before starting to parse this document we must remove any previously saved
         ' data that exists in the gdb2 database for this file. This will also delete
         ' any form codegen parsed structures because line numbers for those structures
         ' are after the main scintilla code.
         gdb2.dbDelete( wszFilename )
         
         ' Reset the count that tracks whether this file contains an Application.Run
         ' statement (multiple lines in the same code file may contain this function call).
         ' For visual designer projects, only one of these statements should exist in 
         ' the entire project.
         pDoc->AppRunCount = 0
      end if
      
      ' Analyze all lines in the file
      for nLineNum = nStartParseLine to nLineCount 
         parser.st       = ""
         parser.st_ucase = ""
         sLine           = ""   ' reset our full line string
         
         ' IMPORTANT: Get the full line as it could be across multiple lines via the line continuation underscore.
         do
            if parser.bParsingCodeGen THEN
               ' Parsing the codegen array
               st = sLineArray( nLineNum )
            else
               ' Parsing the lines in the scintilla control or temporary buffer
               if pDoc->GetActiveScintillaPtr = 0 then
                  st = ltrim(sTempFileBufferArray(nLineNum))
               else 
                  st = pDoc->GetLine( nLineNum )
               end if  
            END IF
            
            parser.lineNum = nLineNum 
                
            ' Check for a TODO entry. Do not normalize line until after this check is done because
            ' TODO's are found as part of a comment.
            if parser.parseToDoItem(st) THEN continue for
            
            ' Mask out all characters inside of a string. This enables us to correctly
            ' test for embedded comment characters and multiple line ":" separators.
            st = trim(MaskStringCharacters(st))   ' important to remove all pre/post spaces

            ' Check for multiline comments 
            if parser.IsMultilineComment(st) THEN continue for
                   
            ' Remove any single line comments
            st = trim(RemoveComments(st))
            
            if right(st, 2) = " _" THEN
               sLine = sLine & left(st, len(st) - 1) & " "
               nLineNum = nLineNum + 1
            else
               ' no line continuation character so we now have all the characters for the current line.   
               sLine = sLine & st 
               exit do
            end if
         loop
         
         parser.lineNum = nLineNum 
       
         ' If the line contains ":" characters then the line needs to be broken up
         ' into as many sublines to be processed. Don't use AfxStrParse because we
         ' need to manually check whether the substring was actually created because
         ' of a trailing ":". Using AfxStrParse and adding a trailing ":" will cause
         ' parse problems.
         for nPos = 1 to len(sLine)
            ' Get the subline but ensure that we add back the ":" because it 
            ' could be part of a Label or identifier like PRIVATE: or PUBLIC:
            dim AS LONG i = instr(nPos, sLine, ":")
            parser.st = trim(AfxStrExtract(nPos, sLine, ":"))
            if i THEN parser.st = parser.st & ":"
            nPos = iif(i, i + 1, len(sLine) + 1)
               
            ' Skip blank lines for speed
            if len(parser.st) = 0 THEN continue for
            parser.st_ucase = Ucase(parser.st)

            ' If normalized line is len = 0 then returns TRUE to continue to next line.
            ' This removes double spaces and also stores the UCASE version of the line.
            if parser.NormalizeLine() then continue for 
            
            ' Continue to parse a TYPE if applicable. If we found a declare line then
            ' this function will return True. Declare names in Types are picked up 
            ' elsewhere when the parser actually encounters the body code of the type/class.
            if parser.parseTYPE() then continue for

            ' Continue to parse an ENUM if applicable
            if parser.parseENUM() then continue for

            ' Test the line to see if we have encountered any start or end points
            ' for a parse. For example, starting to parse a TYPE or FUNCTION, or
            ' ending an existing parse of a TYPE or FUNCTION. If we have completed/ended
            ' a parse action then function returns TRUE and we can advance to the next line.
            if parser.InspectLine() then continue for
            
            ' Parse the line for any DIM'd variables. This call parses for variables 
            ' regardless whether they are global/local, inside a sub/function, or as
            ' part of a TYPE structure. The current parser state will dictate how the
            ' variable information is saved to the database.
            if parser.parseVariableDefinitions() then continue for
            
            ' If the line starts with Application.Run then set the flag in the document. We
            ' can only have one instance of Application.Run in visual designer projects
            ' otherwise the resulting exe will start phantom processes.
            if left(parser.st_ucase, 15) = "APPLICATION.RUN" then
               pDoc->AppRunCount = pDoc->AppRunCount + 1
            end if

         NEXT
      next

      ' Check to see if this is a visual design form. If it is, then we need
      ' to parse the codegen string.
      if parser.bParsingCodeGen = false then
         parser.bParsingCodeGen = true
         if pDoc->IsDesigner then
            ' Load the codegen lines into the sLineArray
            dim as CWSTR wst
            dim as long iLineStart = 1
            dim as long iLineEnd, nNextLine

            ' Convert the CodeGen string into an array that can be parsed.
            nNextLine = nLineCount
            do until iLineStart >= len( pDoc->wszFormCodeGen )
               iLineEnd = instr(iLineStart, pDoc->wszFormCodeGen, vbcrlf)
               if iLineEnd = 0 then iLineEnd = len( pDoc->wszFormCodeGen )  ' cr/lf not found
               wst = mid( pDoc->wszFormCodeGen.wstr, iLineStart, iLineEnd - iLineStart )
               iLineStart = iLineStart + len(wst) + len(vbcrlf)
               if nNextLine >= ubound(sLineArray) THEN
                  redim preserve sLineArray( nLineCount to nLineCount + 5000 )
               END IF
               sLineArray(nNextLine) = wst   
               nNextLine = nNextLine + 1
            loop
            continue do
         end if
      end if
   
      exit do
   loop 

   ' Document has now been fully parsed
   pDoc->bNeedsParsing = false

   ' Update the TODO listview in the frmOutput form
   frmOutput_UpdateToDoListview
   
   Function = 0
 
End Function

