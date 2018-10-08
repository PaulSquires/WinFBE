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
''
Function ScintillaGetLine( byval pDoc as clsDocument ptr, _
                           ByVal nLine As Long _
                           ) As String
   Dim nLen   As Long
   Dim buffer As String
   nLen = SciExec( pDoc->hWindow(0), SCI_LINELENGTH, nLine , 0)
   If nLen < 1 Then Exit Function
   buffer = Space(nLen)
   SciExec( pDoc->hWindow(0), SCI_GETLINE, nLine, Strptr(buffer))
   Function = Rtrim(buffer, Any Chr(13,10,0))
End Function


''
''
Function ParseDocument( byval idx as long, _     ' project index
                        byval pDoc as clsDocument ptr, _
                        byval wszFilename as CWSTR _    
                        ) As Long

   dim as long nLineCount, nLineNum, nPos
   Dim As String st, st_ucase, sLine
   redim as string sLineArray(5000)   ' dynamic array 
   
   dim as boolean bParseCodeGenString = false
   
ParseAgain:
   ' The incoming document may already be loaded into a clsDocument so we can simply 
   ' use that pDoc to retrieve the lines. If the pDoc = Null then we need to load
   ' the file from disk into a standard string array.
   if bParseCodeGenString then
      ' Load the codegen lines into the sLineArray
      if pDoc THEN
         wszFilename = pDoc->DiskFilename
         idx = pDoc->ProjectIndex
         dim as CWSTR wst
         dim as long iLineStart = 1
         dim as long iLineEnd
         nLineCount = 0
         do until iLineStart >= len(pDoc->wszFormCodeGen)
            iLineEnd = instr(iLineStart, pDoc->wszFormCodeGen, vbcrlf)
            if iLineEnd = 0 then iLineEnd = len(pDoc->wszFormCodeGen)  ' cr/lf not found
            wst = mid(pDoc->wszFormCodeGen.wstr, iLineStart, iLineEnd - iLineStart)
            iLineStart = iLineStart + len(wst) + len(vbcrlf)
            if nLineCount >= ubound(sLineArray) THEN
               redim preserve sLineArray(nLineCount + 5000)
            END IF
            sLineArray(nLineCount) = wst   
            nLineCount = nLineCount + 1
         loop
         pDoc = 0
      end if
   else
      if pDoc THEN
         nLineCount = SciExec(pDoc->hWindow(0), SCI_GETLINECOUNT, 0, 0)
         wszFilename = pDoc->DiskFilename
         idx = pDoc->ProjectIndex
         ' Before starting to parse this document we must remove any previously saved
         ' data that exists in the gdb2 database for this file.
         gdb2.dbDelete4File(idx,wszFilename)
      else
         ' These are often the FB/WinAPI codetip files and Include Files. Everything else in
         ' a project would have a valid pDoc.
         if AfxFileExists(wszFilename) = 0 THEN exit function

         dim pStream as CTextStream
         if pStream.Open(wszFilename) = S_OK then
            do until pStream.EOS
               if nLineCount >= ubound(sLineArray) THEN
                  redim preserve sLineArray(nLineCount + 5000)
               END IF
               sLineArray(nLineCount) = pStream.ReadLine 
               nLineCount = nLineCount + 1
            loop
         end if
         pStream.Close
         ' Before starting to parse this document we must remove any previously saved
         ' data that exists in the gdb2 database for this file.
         gdb2.dbDelete4File(idx,wszFilename)
      END IF                     
   end if
   
   if gFileLoading then
      SetCompileStatusBarMessage( "Parsing: " & wszFilename, 0 )
   END IF

   ' Create a parse TYPE that handles the state information throughout parsing this document.
   dim parser as clsParser
   
   gdb2.dbAdd4Proj(gApp.Projects(idx).ProjectFilename,idx)
   gdb2.dbAdd4File(idx,wszFilename) 
   with parser
      .action     = ACTION_NONE
      .idxProject = idx
      .fileName   = wszFilename
   end with


   ' Analyze all lines in the file
   for nLineNum = 0 to nLineCount - 1
      parser.st       = ""
      parser.st_ucase = ""
      sLine           = ""   ' reset our full line string
      
      parser.lineNum = nLineNum
      ' IMPORTANT: Get the full line as it could be across multiple lines via the line continuation underscore.
      do
         if pDoc THEN
            st = ScintillaGetLine(pDoc, nLineNum)
         else
            st = sLineArray(nLineNum)
         END IF
         
         'parser.lineNum = nLineNum 
             
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
            sLine = sLine & left(st, len(st) - 1) & " " & ContinueLineFlag
            nLineNum = nLineNum + 1
         else
            ' no line continuation character so we now have all the characters for the current line.   
            sLine = sLine & st 
            exit do
         end if
      loop
      
      'parser.lineNum = nLineNum 
    
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
         
         ' Continue to parse a TYPE if applicable
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

      NEXT
   next

   if bParseCodeGenString = false then
      ' Check to see if this is a visual design form. If it is, then we need
      ' to parse the codegen string.
      if pDoc->IsDesigner then
         bParseCodeGenString = true
         goto ParseAgain
      END IF
   END IF
   
   ' Update the TODO listview in the frmOutput form
   UpdateToDoListview
   
   SetCompileStatusBarMessage( "", 0 )
   Function = 0

End Function



