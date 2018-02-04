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
Function GetIncludeFilename( byref wszFilename as CWSTR, _    ' file where the #Include was found in (needed for pathing)
                             byref sLine as string _
                             ) as CWSTR
   ' With the incoming line parse it to determine the fully pathed
   ' include file that it represents. Incoming line is original case.
   ' #INCLUDE "filename" or #INCLUDE ONCE "filename"
   dim as long f1, f2
   dim as CWSTR sIncFilename, sPath, sCompilerPath, sTemp, sFilename, sTestPath
   
   f1 = instr(9, sLine, chr(34))                ' opening quote
   if f1 THEN f2 = instr(f1+1, sLine, chr(34))  ' ending quote
   sIncFilename = mid(sLine, f1+1, f2-f1-1)

   ' Search the compiler's include folder path
   ' Get the array index for the current selected build (project or document)
   dim as long idxBuild = ComboBox_GetCurSel(HWND_FRMMAIN_COMBOBUILDS)
   if idxBuild = -1 then idxBuild = 0
   if gConfig.Builds(idxBuild).Is32bit then
      sPath = ProcessFromCurdrive(gConfig.FBWINCompiler32)
      sCompilerPath = lcase(AfxStrPathname("PATH", sPath))
   end if
   if gConfig.Builds(idxBuild).Is64bit then
      sPath = ProcessFromCurdrive(gConfig.FBWINCompiler64)
      sCompilerPath = lcase(AfxStrPathname("PATH", sPath))
   End if
   sTemp = sCompilerPath & "inc\" & sIncFilename
   if AfxFileExists(sTemp) THEN 
      ' We don't want to parse all Windows include files. If the windows.bi file
      ' was found then later we will load our winapi include through gConfig.
      if ucase(sIncFilename) = "WINDOWS.BI" then return sTemp
      ' Never allow parsing of any windows api include files because we use our
      ' own version of winapi preparsed.
      sTestPath = sCompilerPath & "inc\win"
      if left(sTemp, len(sTestPath)) = sTestPath then 
         return ""
      else
         return sTemp
      end if      
   end if   

   ' Begin our search for the file relative to the file that this #Include directive was found in.
   sPath = AfxStrPathname("PATH", wszFilename)
   if AfxPathIsRelative(sIncFilename) then  ' If this is a relative filename then convert it back.
      sTemp = AfxPathCombine( sPath, sIncFilename)
      if AfxFileExists(sTemp) THEN return sTemp
   else
      sTemp = sPath & sIncFilename
      if AfxFileExists(sTemp) THEN return sTemp 
   end if
   
   ' search relative to the project file of the active project
   dim as long idx = gApp.GetActiveProjectIndex()
   if idx then
      sPath = AfxStrPathname("PATH", gApp.Projects(idx).ProjectFilename)
      if AfxPathIsRelative(sIncFilename) then  ' If this is a relative filename then convert it back.
         sTemp = AfxPathCombine( sPath, sIncFilename)
         if AfxFileExists(sTemp) THEN return sTemp
      else
         sTemp = sPath & sIncFilename
         if AfxFileExists(sTemp) THEN return sTemp 
      end if
   end if
      
   return ""
End Function

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
Function ParseOneDocument( byval idx as long, _     ' project index
                           byval pDoc as clsDocument ptr, _
                           byval wszFilename as CWSTR, _    
                           byval bAllowIncludes as Boolean _
                           ) As Long

   dim as long nLineCount, nLineNum, nPos
   Dim As String st, st_ucase, sLine
   redim as string sLineArray(5000)   ' dynamic array 
   
   ' The incoming document may already be loaded into a clsDocument so we can simply 
   ' use that pDoc to retrieve the lines. If the pDoc = Null then we need to load
   ' the file from disk into a standard string array.
   if pDoc THEN
      nLineCount = SciExec(pDoc->hWindow(0), SCI_GETLINECOUNT, 0, 0)
      wszFilename = pDoc->DiskFilename
      idx = pDoc->ProjectIndex
   else
      ' These are often the FB/WinAPI codetip files and Include Files. Everything else in
      ' a project would have a valid pDoc.
      if AfxFileExists(wszFilename) = 0 THEN exit function

      dim as long f = freefile
      open wszFilename for input as #f
      do until eof(f)
         if nLineCount >= ubound(sLineArray) THEN
            redim preserve sLineArray(nLineCount + 5000)
         END IF
         line input #f, sLineArray(nLineCount)
         nLineCount = nLineCount + 1
      loop
      close #f
   END IF                     

   if gFileLoading then
      SetCompileStatusBarMessage( "Parsing: " & wszFilename, 0 )
   END IF

   ' Create a parse TYPE that handles the state information throughout parsing this document.
   dim parser as clsParser
   
   with parser
      .action     = ACTION_NONE
      .idxProject = idx
      .fileName   = wszFilename
   end with


   ' Before starting to parse this document we must remove any previously saved
   ' data that exists in the gdb2 database for this file.
   gdb2.dbDelete(wszFilename)
   ' Add back a record to indicate that this file has been parsed
   parser.varName = wszFilename
   gdb2.dbAdd(parser, DB2_INCLUDEFILE)


   ' Analyze all lines in the file
   for nLineNum = 0 to nLineCount - 1
      parser.st       = ""
      parser.st_ucase = ""
      sLine           = ""   ' reset our full line string
      
      ' IMPORTANT: Get the full line as it could be across multiple lines via the line continuation underscore.
      do
         if pDoc THEN
            st = ScintillaGetLine(pDoc, nLineNum)
         else
            st = sLineArray(nLineNum)
         END IF
         
         parser.lineNum = nLineNum 
             
         ' Check for a TODO entry. Do not normalize line until after this check is done because
         ' TODO's are found as part of a comment.
         if parser.parseToDoItem(st) THEN continue for
         
         ' If we encounter a #Include then do a recursive parse of it. Must parse it here
         ' prior to the call that masks out all characters in a quoted string.
         if bAllowIncludes then
            if parser.IsIncludeFile(st) THEN continue for
         end if
           
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
      ' need to manually check whether the substring was actually crated because
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

   ' Update the TODO listview in the frmOutput form
   UpdateToDoListview
   
   SetCompileStatusBarMessage( "", 0 )
   Function = 0

End Function


''
''
Function ParseDocument( byval idx as long, _     ' project index
                        byval pDoc as clsDocument ptr, _
                        byval wszFilename as CWSTR, _    
                        byval bNoParseIncludes as Boolean = false _
                        ) As Long
   ' Parse the incoming document. During parsing of this main document we
   ' will collect any occurences of #INCLUDE files and then parse them afterwards.
   dim as Boolean bWindowsIncludes 
   
   erase gIncludeFiles
   ParseOneDocument(idx, pDoc, wszFilename, true)

   ' NoParseIncludes is set to True during normal editing of a file (rather than
   ' during file loading or saving when we do want to parse includes). Disabling 
   ' parsing during an edit prevents parsing when an autocomplete is open and the
   ' user is typing. Parsing includes during every keypress would be very slow.
   if bNoParseIncludes then exit function
   if gFileLoading = false THEN exit function   ' only parse includes during initial load
      
   dim as long i = lbound(gIncludeFiles)
   do until i > ubound(gIncludeFiles)
      ' ubound will continue to change as more include files are identified.
      if bWindowsIncludes = false then
         if lcase(AfxStrPathName("NAMEX", gIncludeFiles(i))) = "windows.bi" then
            ' Make sure that we don't add the windows.bi to the list.
            bWindowsIncludes = true
            i = i + 1
            continue do
         END IF
      end if
      ' TODO: Parsing shobjidl.bi causes a GPF for some reason.
      if lcase(AfxStrPathName("NAMEX", gIncludeFiles(i))) <> "shobjidl.bi" then
         ' If this is an initial loading then no need to reparse the same
         ' file multiple times.
         if gdb2.dbFilenameExists(gIncludeFiles(i)) = 0 THEN          
            ParseOneDocument(idx, 0, gIncludeFiles(i), true)
         end if
      end if
      i = i + 1
   loop
                                
   ' If the "windows.bi" file was included then parse the main windows files
   if gApp.IsWindowIncludes = false then     ' flag to only parse Win includes once
      if bWindowsIncludes then
         ' Load the WinAPI Codetips file
         gConfig.LoadCodetipsWinAPI( AfxGetExePathName & "Settings\codetips_winapi.ini" )
         gApp.IsWindowIncludes = true
      end if
   end if
   
   function = 0
   
end function

