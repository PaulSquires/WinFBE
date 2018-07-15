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


''  PREPARSE MODULE
''  Load metadata for previously parsed Include files. Only re-parse those files
''  that have changed since the last load. using preparsed files is much faster
''  than having to do a live parse everytime a file or project is loaded.
''  Files that are parsed include everything that is in the compiler's \inc
''  folder structure. This includes Windows system includes as well as any
''  user defined includes that are stored there.

' The functions in this module check a master database file that holds filenames
' and timestamps. If the actual file timestamp differs then the file is re-parsed
' and the results saved in the preparsed database (replacing any previous entries
' for that particular file).


''
''
function LoadPreparseTimestamps( byref sFilename as string ) as boolean
   if AfxFileExists(sFilename) = false then return false
   
   dim as string st
   dim as long f, ub, i
   
   dim as CWSTR wszIncludePath 
   #IfDef __FB_64BIT__
      wszIncludePath = ProcessFromCurdrive(gConfig.FBWINCompiler64)
   #Else
      wszIncludePath = ProcessFromCurdrive(gConfig.FBWINCompiler32)
   #EndIf
   wszIncludePath = ucase(AfxStrPathname("PATH", wszIncludePath)) & "INC\"

   f = freefile
   open sFilename for input as #f
   do until eof(f)
      line input #f, st
      i = instr(st, "|")
      if i = 0 then continue do
      ub = ubound(gPreparseTimestamps) + 1
      redim preserve gPreparseTimestamps(ub)
      gPreparseTimestamps(ub).wszFilename = wszIncludePath & left(st, i-1)
      gPreparseTimestamps(ub).tFiletime = val(mid(st, i+1))
   LOOP
   close #f
   
   return true 
END FUNCTION

''
''
function SavePreparseTimestamps( byref sFilename as string ) as boolean
   dim as long f = freefile
   
   dim as CWSTR wszIncludePath 
   #IfDef __FB_64BIT__
      wszIncludePath = ProcessFromCurdrive(gConfig.FBWINCompiler64)
   #Else
      wszIncludePath = ProcessFromCurdrive(gConfig.FBWINCompiler32)
   #EndIf
   wszIncludePath = ucase(AfxStrPathname("PATH", wszIncludePath)) & "INC\"

   open sFilename for output as #f
   for i as long = lbound(gPreparseTimestamps) to ubound(gPreparseTimestamps)
      print #f, MID(gPreparseTimestamps(i).wszFilename, LEN(wszIncludePath) + 1); "|"; _
                gPreparseTimestamps(i).tFiletime
   next
   close #f
   
   return true 
END FUNCTION

''
''
function LoadPreparseDatabase( byref sFilename as string ) as boolean
   if AfxFileExists(sFilename) = false then return false
   
   dim as string st
   dim as long f, i
   dim parser as clsParser
   
   dim as CWSTR wszIncludePath 
   #IfDef __FB_64BIT__
      wszIncludePath = ProcessFromCurdrive(gConfig.FBWINCompiler64)
   #Else
      wszIncludePath = ProcessFromCurdrive(gConfig.FBWINCompiler32)
   #EndIf
   wszIncludePath = ucase(AfxStrPathname("PATH", wszIncludePath)) & "INC\"

   f = freefile
   open sFilename for input as #f
   do until eof(f)
      line input #f, st
      i = instr(st, "|")
      if i = 0 then continue do
      
      with parser
         .idxProject = -2
         .fileName   = wszIncludePath & AfxStrParse(st, 1, "|")
         .funcName   = AfxStrParse(st, 2, "|")
         .funcParams = AfxStrParse(st, 3, "|")    ' Calltip
      end with
      
      ' Jose's Afx functions define Left/Right/Val in several files. We don't want
      ' these definitions to override FB's builtin definitions, so don't add them.
      select case ucase(parser.funcName)
         CASE "LEFT", "RIGHT", "VAL"
         case else   
            gdb2.dbAdd(parser, DB2_FUNCTION)
      END SELECT

   LOOP
   close #f
   
   return true 
END FUNCTION

''
''
function SavePreparseDatabase( byref sFilename as string ) as boolean
   dim as long f = freefile
   dim pData as DB2_DATA ptr
   
   dim as CWSTR wszIncludePath 
   #IfDef __FB_64BIT__
      wszIncludePath = ProcessFromCurdrive(gConfig.FBWINCompiler64)
   #Else
      wszIncludePath = ProcessFromCurdrive(gConfig.FBWINCompiler32)
   #EndIf
   wszIncludePath = ucase(AfxStrPathname("PATH", wszIncludePath)) & "INC\"

   open sFilename for output as #f
   gDB2.dbRewind()
   do 
      pData = gDB2.dbGetNext
      if pData = 0 THEN exit do
      if pData->projectIndex <> -2 THEN continue do   ' only look at preparse entries
      if pData->deleted = true THEN continue do

      select case pData->id
         case DB2_FUNCTION, DB2_SUB
            print #f, UCASE(MID(pData->fileName, LEN(wszIncludePath)+1)); "|"; _
                      pData->ElementName; "|"; pData->ElementValue
         case else
            continue do
      END SELECT
   loop
   close #f
   
   return true 
END FUNCTION


''
''
function PreparseFolder( byval wszIncludePath as CWSTR, byval wszSubFolder as CWSTR ) as Long
   dim wszUCaseFilename as CWSTR 
   dim wszPath as CWSTR = wszIncludePath & wszSubFolder & "\*.*"
   DIM pFinder AS CFindFile
   dim bFound  as Boolean
   dim LongFileTime as LONGLONG
   
   IF pFinder.FindFile(wszPath) = S_OK THEN
      DO
         IF pFinder.IsDots = FALSE THEN   ' // skip . and .. files
            ' If the file is a subfolder then recurse into that folder and process
            wszUCaseFilename = ucase(pFinder.FilePath)
            if pFinder.IsFolder then
               ' Filter out known folders that we are not interested in parsing.
               select case ucase(pFinder.FilenameX)
                  case "AL", "ALLEGRO", "ALLEGRO5", "CUNIT", "CURSES", "DOS", "GDK-PIXBUF"
                  case "GRX", "GSL", "GTK", "GTKGL", "WIN", "X11", "WINFORMSX"
                  CASE else
                     wszSubFolder = mid(pFinder.FilePath, len(wszIncludePath)+1)
                     PreparseFolder(wszIncludePath, wszSubFolder)
               END SELECT

            else
               ' Determine if the found file contents have changed since last save
               LongFileTime = AfxFileTimeToTime64(pFinder.LastWriteTime)
               bFound = false
               for i as long = lbound(gPreparseTimestamps) to ubound(gPreparseTimestamps)
                  if wszUCaseFilename = gPreparseTimestamps(i).wszFilename then
                     bFound = true
                     ' Compare filestamps to see if the file has changed
                     if gPreparseTimestamps(i).tFiletime <> LongFileTime then
                        ' Parse the found file   
                        gPreparseTimestamps(i).tFiletime = LongFileTime
                        gdb2.dbDelete(wszUCaseFilename)
                        ParseDocument(-2, 0, pFinder.FilePath)
                        gPreparsingChanges = true
                     end if
                     exit for
                  end if
               next

               ' If the file was not found then create an entry for it.
               if bFound = false then
                  dim as long ub = ubound(gPreparseTimestamps) + 1
                  redim preserve gPreparseTimestamps(ub)
                  gPreparseTimestamps(ub).wszFilename = wszUCaseFilename
                  gPreparseTimestamps(ub).tFiletime = LongFileTime
                  gdb2.dbDelete(wszUCaseFilename)
                  ParseDocument(-2, 0, pFinder.FilePath)
                  gPreparsingChanges = true
               end if
            end if
         END IF
         IF pFinder.FindNext = 0 THEN EXIT DO
      LOOP
   END IF
   pFinder.Close
   
   function = 0
END FUNCTION









