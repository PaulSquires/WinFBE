'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2017 Paul Squires, PlanetSquires Software
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
Function IsPreparsedFile( byref sFilename as string ) as boolean
   ' Create the list of preparsed WinAPI files that we will bypass during
   ' parsing. This depends on the compiler chosen and filepath.

   dim as string sPath, sTemp
   sFilename = AfxStrReplace( sFilename, "/", "\")
   
   Select Case gApp.Projects(gApp.GetActiveProjectIndex).GetProjectCompiler
      Case IDM_32BIT:  sPath = ProcessFromCurdrive(gConfig.FBWINCompiler32)
      Case IDM_64BIT:  sPath = ProcessFromCurdrive(gConfig.FBWINCompiler64)
   End Select
   sPath = AfxStrPathname("PATH", sPath) & "inc\"
   
   ' We do not want to parse the windows.bi file or any files in the \win
   ' include folder.
   sTemp = sPath & "windows.bi"
   if _stricmp(sTemp, sFilename) = 0 THEN return true
   
   sPath = sPath & "win\"
   sTemp = AfxStrPathName("PATH", sFilename)
   if _stricmp(sPath, sTemp) = 0 THEN return true
      
   function = false
end function


''
''
Function GetIncludeFilename( byref sFilename as string, _    ' file where the #Include was found in (needed for pathing)
                             byref sLine as const string _
                             ) as STRING
   ' With the incoming line parse it to determine the fully pathed
   ' include file that it represents. Incoming line is original case.
   ' #INCLUDE "filename" or #INCLUDE ONCE "filename"
   dim as long f1, f2
   dim as string sIncFilename, sPath, sTemp
   
   f1 = instr(9, sLine, chr(34))                ' opening quote
   if f1 THEN f2 = instr(f1+1, sLine, chr(34))  ' ending quote
   sIncFilename = mid(sLine, f1+1, f2-f1-1)

   ' Begin our search for the file relative to the file that this #Include directive was found in.
   sPath = AfxStrPathname("PATH", sFilename)
   if AfxPathIsRelative(sIncFilename) then  ' If this is a relative filename then convert it back.
      sTemp = AfxPathCombine( sPath, sIncFilename)
      if AfxFileExists(sTemp) THEN return sTemp
   else
      sTemp = sPath & sIncFilename
      if AfxFileExists(sTemp) THEN return sTemp 
   end if
   
   ' Search the compiler's include folder path
   Select Case gApp.Projects(gApp.GetActiveProjectIndex).GetProjectCompiler
      Case IDM_32BIT:  sPath = ProcessFromCurdrive(gConfig.FBWINCompiler32)
      Case IDM_64BIT:  sPath = ProcessFromCurdrive(gConfig.FBWINCompiler64)
   End Select
   sTemp = AfxStrPathname("PATH", sPath) & "inc\" & sIncFilename
   if AfxFileExists(sTemp) THEN 
      ' We have found an Include file that may be part of the already preparsed
      ' and loaded Include files. Do a check and bypass parsing the file if it
      ' is a preparsed file.
      if IsPreparsedFile(sTemp) THEN 
         return ""
      else
         return sTemp   
      END IF
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
   nLen = SciExec( pDoc->hWindow, SCI_LINELENGTH, nLine , 0)
   If nLen < 1 Then Exit Function
   buffer = Space(nLen)
   SciExec( pDoc->hWindow, SCI_GETLINE, nLine, Strptr(buffer))
   Function = Rtrim(buffer, Any Chr(13,10,0))
End Function


''
''
Function ParseDocument( byval idx as long, _     ' project index
                        byval pDoc as clsDocument ptr, _
                        byval sFilename as string _
                        ) As Long
   
   dim as long nLineCount, nLineNum, nFuncLineNumber
   Dim As String st, st_ucase, sLine, txtDescription, sFuncName, sFuncDeclare
   dim as string sTypeName, sTypeElements, sVarName, sVarType
   Dim As Long n, iParen, f, i
   Dim As BOOLEAN bIsSub, bIsFunction, bInType, bInTypePrivate, bInMultilineComment
   redim as string sLineArray(1000)   ' dynamic array 
   
   ' The incoming document may already be loaded into a clsDocument so we can simply 
   ' use that pDoc to retrieve the lines. If the pDoc = Null then we need to load
   ' the file from disk into a standard string array.
   if pDoc THEN
      nLineCount = SciExec(pDoc->hWindow, SCI_GETLINECOUNT, 0, 0)
      sFilename = pDoc->DiskFilename
   else
      if AfxFileExists(sFilename) = 0 THEN exit function
      f = freefile
      open sFilename for input as #f
      do until eof(f)
         if nLineCount >= ubound(sLineArray) THEN
            redim preserve sLineArray(nLineCount + 1000)
         END IF
         line input #f, sLineArray(nLineCount)
         nLineCount = nLineCount + 1
      loop
      close #f
   END IF                     
   
   ' Before starting to parse this document we must remove any previously saved
   ' data that exists in gdb database for this file.
   gdb.dbDelete(sFilename)
   gdb.dbAdd(sFilename, idx, DB_FILENAME_PARSE, 0, "", "")

   
   ' Analyze all lines in the file
   for nLineNum = 0 to nLineCount - 1
      bIsSub = false:  bIsFunction = false
      st = ""
      do
         ' Get the full line as it could be across multiple lines via the line continuation underscore.
         if pDoc THEN
            sLine = ScintillaGetLine(pDoc, nLineNum)
         else
            sLine = sLineArray(nLineNum)   
         END IF
      
         ' Check for a TODO entry
         if (left(ucase(ltrim(sLine)), 6) = "'TODO:") or _
            (left(ucase(ltrim(sLine)), 7) = "' TODO:") THEN
            txtDescription = AfxStrParse(sLine, 2, ":") 
            gdb.dbAdd(sFilename, idx, DB_TODO_ITEM, nLineNum + 1, txtDescription, "")
         END IF
         
         ' Deal with multiline comments
         if left(ltrim(sLine), 2) = "/'" THEN
            bInMultilineComment = true
         else
            ' Remove any comments (must be after the check for the TODO)
            if bInMultilineComment = false then
               n = Instr(sLine, "'")
               If n Then sLine = Left(sLine, n-1)
            end if
         END IF

         sLine = Trim(sLine)
         if right(sLine, 1) = "_" THEN
            st = st & left(sLine, len(sLine)-1) & " "
            nLineNum = nLineNum + 1
         else
            st = st & sLine
            exit do   
         END IF
      loop 

      If Len(st) = 0 Then Continue for
      st_ucase = Ucase(st)
      
      if bInMultilineComment then
         if right(rtrim(st), 2) = "'/" THEN
            bInMultilineComment = false
         END IF
         continue for
      end if   
         
      ' If we are currently processing a TYPE structure then continue
      ' to do so here rather than testing for other declares.
      if bInType THEN
         ' Is this the last line of the TYPE
         If Left(st_ucase, 8)  = "END TYPE" Then
            gdb.dbAdd(sFilename, idx, DB_TYPE_NAME, nLineNum, sTypeName, trim(sTypeElements,DELIM))
            bInTYPE = false: bInTypePrivate = false
            sTypeElements = ""
         else
            if len(st) = 0 THEN continue for
            if left(st_ucase, 8) = "PRIVATE:" THEN 
               bInTypePrivate = true: continue for
            elseif left(st_ucase, 7) = "PUBLIC:" THEN 
               bInTypePrivate = false: continue for
            elseif left(st_ucase, 1) = "#" THEN 
               continue for   ' conditional equate 
            end if
            if bInTypePrivate then continue for

            ' As Long var1, var2, var3
            ' As Long Ptr var1, var2, var3
            ' var1 as Long
            ' Declare Sub
            ' Declare Function
            ' Property
            ' Operator
            st = RemoveDuplicateSpaces(st)
            st_ucase = ucase(st)
            ' AS Long var1, var2, var3
            if left(st_ucase, 3) = "AS " THEN
               ' the word following is the element type
               n = instr(4, st_ucase, " ")
               if n then
                  ' Maybe the next word is "PTR" or "POINTER"
                  if mid(st_ucase, n + 1, 4) = "PTR " then
                     n = n + 5
                  elseif mid(st_ucase, n + 1, 8) = "POINTER " then
                     n = n + 9
                  END IF 
               END IF
               ' Next is a list of elements DIM'd as the datatype
               st = AfxStrReplace(mid(st, n), ", ", DELIM)
            end if
            ' var1 AS Long
            n = instr(st_ucase, " AS ") 
            if n THEN st = left(st, n-1)
            
            ' Process this line of the TYPE because it is Public
            sTypeElements = sTypeElements & DELIM & st
         end if
         continue for   
      END IF
      
      if left(st_ucase, 5) = "TYPE " THEN
         ' If " AS " exists then this is not a TYPE structure
         if instr(st_ucase, " AS ") THEN continue for
         sTypeName = AfxStrParse( ltrim(mid(st, 6)), 1, " ")
         sTypeElements = ""
         bInType = true: continue for
      end if   

      ' Search for UNICODE indicator
      if left(st_ucase, 7) = "#DEFINE" THEN
         if trim(mid(st_ucase, 8)) = "UNICODE" THEN
            gApp.IsUnicodeCodetips = true: continue for
         END IF
      END IF
         
      ' Is this a global shared variable
      if left(st_ucase, 11) = "DIM SHARED " THEN
         st = RemoveDuplicateSpaces(st)
         st = mid(st, 11): st_ucase = ucase(st)
         ' AS Long var1, var2, var3
         if left(st_ucase, 3) = "AS " THEN
            ' the word following is the element type
            n = instr(4, st_ucase, " ")
            if n then
               sVarType = mid(st, n-4+1)
               ' Maybe the next word is "PTR" or "POINTER"
               if mid(st_ucase, n + 1, 4) = "PTR " then
                  n = n + 5
               elseif mid(st_ucase, n + 1, 8) = "POINTER " then
                  n = n + 9
               END IF 
            END IF
            ' Next is a list of elements DIM'd as the datatype
            sVarName = AfxStrReplace(mid(st, n), ", ", DELIM)
         else
            ' var1 AS Long
            n = instr(st_ucase, " AS ") 
            if n = 0 THEN continue for
            sVarName = trim(left(st, n-1))
            sVarType = mid(st, n + 4)
            gdb.dbAdd(sFilename, idx, DB_DIM_SHARED, 0, sVarName, sVarType)
         end if
         continue for
      end if
      
      ' Is this an Include file that will need to be parsed? If yes, then
      ' get the filename and do a recursive call to ParseDocument.
      If Left(st_ucase, 9)  = "#INCLUDE " Then
         ' #INCLUDE "filename" or #INCLUDE ONCE "filename"
         dim as string sIncludeFilename = GetIncludeFilename(sFilename, st)
         
         ' If the include file was already parsed then skip it
         if AfxFileExists(sIncludeFilename) then
            if gdb.dbFilenameExists(sIncludeFilename) = false THEN
               ParseDocument( idx, 0, sIncludeFilename )
            end if
         END IF
      end if
      
      '' TEST FOR SUB/FUNCTION
      If Left(st_ucase, 9)  = "FUNCTION " Then
         st = Mid(st, 10): bIsFunction = true
      ElseIf Left(st_ucase, 16) = "PUBLIC FUNCTION " Then
         st = Mid(st, 17): bIsFunction = true
      ElseIf Left(st_ucase, 17) = "PRIVATE FUNCTION " Then
         st = Mid(st, 18): bIsFunction = true
      elseIf Left(st_ucase, 4)  = "SUB " Then
         st = Mid(st, 5): bIsSub = True
      ElseIf Left(st_ucase, 11) = "PUBLIC SUB " Then
         st = Mid(st, 12): bIsSub = True
      ElseIf Left(st_ucase, 12) = "PRIVATE SUB " Then
         st = Mid(st, 13): bIsSub = True
      End If
      st = ltrim(st)  ' remove any leading spaces in the new substring  
   
      ''  PROCESS ANY SUB/FUNCTION found
      if (bIsSub = true) orelse (bIsFunction = true) THEN 
         ' bypass any line with = embedded in it but need to check only up to the first
         ' open parenthesis because after that are the parameters and they could have equal
         ' signs for default parameter values.
         n = Instr( st, "=") 
         If n Then 
            iParen = Instr(st, "(") 
            If (iParen = 0) OrElse (n < iParen) Then
               Continue for
            End If
         End If

         ' Determine the name of this sub/function
         iParen = Instr(st, "(") 
         sFuncName = iIf(iParen, trim(Left(st, iParen - 1)), st)
         ' Remove any reference to "OVERLOAD" 
         n = instr(ucase(sFuncName), "OVERLOAD") 
         if n then sFuncName = rtrim(left(sFuncName, n-1))
         
         sFuncDeclare = iif(bIsFunction, "Function ", "Sub ") & st
         nFuncLineNumber = nLineNum    
         gdb.dbAdd(sFilename, idx, DB_FUNCTION_NAME, nLineNum, sFuncName, sFuncDeclare)
   
      end if
      
   next

   ' Update the TODO listview in the frmOutput form
   UpdateToDoListview
   
   Function = 0

End Function

