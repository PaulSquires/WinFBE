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
   Select Case Ucase(gConfig.DefaultCompiler)
      Case "FBC 32BIT":  sPath = ProcessFromCurdrive(gConfig.FBWINCompiler32)
      Case "FBC 64BIT":  sPath = ProcessFromCurdrive(gConfig.FBWINCompiler64)
   End Select
   sTemp = AfxStrPathname("PATH", sPath) & "inc\" & sIncFilename
   if AfxFileExists(sTemp) THEN 
      ' We have found an Include file that may be part of the already preparsed
      ' and loaded Include files. Do a check and bypass parsing the file if it
      ' is a preparsed file.
      '// TODO: Search list of preparsed files
      
      return ""    ' don't parse this file
   end if   
   
End Function

''
''
Function ScintillaGetLine( byval pDoc as clsDocument ptr, _
                           ByVal nLine As Long _
                           ) As String
   Dim nLen   As Long
   Dim buffer As String
   nLen = SendMessage( pDoc->hWindow, SCI_LINELENGTH, nLine , 0)
   If nLen < 1 Then Exit Function
   buffer = Space(nLen)
   SendMessage( pDoc->hWindow, SCI_GETLINE, nLine, Cast(LPARAM, Strptr(buffer)))
   Function = Rtrim(buffer, Any Chr(13,10,0))
End Function


''
''
Function ParseDocument( byval pDoc as clsDocument ptr, _
                        byval sFilename as string _
                        ) As Long
                        
   dim as long nLineCount, nLineNum, nFuncLineNumber
   Dim As String st, st_ucase, sLine, sFuncName, sFuncDeclare, sAllFunctionNames
   Dim As Long n, iParen, f
   Dim As BOOLEAN bFirstLine, bIsFunction
   redim as string sLineArray(1000)   ' dynamic array 
   
   ' The incoming document may already be loaded into a clsDocument so we can simply 
   ' use that pDoc to retrieve the lines. If the pDoc = Null then we need to load
   ' the file from disk into a standard string array.
   if pDoc THEN
      nLineCount = SendMessage(pDoc->hWindow, SCI_GETLINECOUNT, 0, 0)
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
   ' data that exists in the hash arrays. The most current data will replace it 
   ' once the parse completes.
   ' gFunctionDeclares: declares for all Subs/Functions
   gFunctionDeclares.DeleteAllData2(ucase(sFilename))

   ' gFunctionLists: filenames and subs/functions within it
   ' Key: Filename.
   ' This one is easy, simply delete the one hash entry for this filename
    gFunctionLists.DeleteItem(ucase(sFilename))
   
   
   
   for nLineNum = 0 to nLineCount - 1
      bFirstLine = false   
      bIsFunction = false
      st = ""
      do
         ' Get the full line as it could be across multiple lines via the line continuation underscore.
         if pDoc THEN
            sLine = ScintillaGetLine(pDoc, nLineNum)
         else
            sLine = sLineArray(nLineNum)   
         END IF
      
         ' Remove any comments
         n = Instr(sLine, "'")
         If n Then sLine = Left(sLine, n-1)
         
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

      ' Search for UNICODE indicator
      if left(st_ucase, 7) = "#DEFINE" THEN
         if trim(mid(st_ucase, 8)) = "UNICODE" THEN
            gApp.IsUnicodeCodetips = true
         END IF
      END IF
         
      ' Is this an Include file that will need to be parsed? If yes, then
      ' get the filename and do a recursive call to ParseDocument.
      If Left(st_ucase, 9)  = "#INCLUDE " Then
         ' #INCLUDE "filename" or #INCLUDE ONCE "filename"
         dim as string sIncludeFilename = GetIncludeFilename(sFilename, st)
         ParseDocument( 0, sIncludeFilename )
      end if
      
      If Left(st_ucase, 9)  = "FUNCTION " Then
         st = Mid(st, 10): bFirstLine = True: bIsFunction = true
      ElseIf Left(st_ucase, 16) = "PUBLIC FUNCTION " Then
         st = Mid(st, 17): bFirstLine = True: bIsFunction = true
      ElseIf Left(st_ucase, 17) = "PRIVATE FUNCTION " Then
         st = Mid(st, 18): bFirstLine = True: bIsFunction = true
      End If   
      
      If Left(st_ucase, 4)  = "SUB " Then
         st = Mid(st, 5): bFirstLine = True
      ElseIf Left(st_ucase, 11) = "PUBLIC SUB " Then
         st = Mid(st, 12): bFirstLine = True
      ElseIf Left(st_ucase, 12) = "PRIVATE SUB " Then
         st = Mid(st, 13): bFirstLine = True
      End If
      st = ltrim(st)  ' remove any leading spaces   
   
      if bFirstLine = false THEN continue for

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
      nFuncLineNumber = nLineNum    'b/c Scintilla is zero based
      gFunctionDeclares.AddItem(ucase(sFuncName), sFuncDeclare, ucase(sFilename), str(nFuncLineNumber))
            
      sAllFunctionNames = sAllFunctionNames & DELIM & sFuncName
   next
   
   ' We are now finished parsing this file. Add the found list of function
   ' names to the global hash. Data1 holds a DELIM delimited string of function names.
   gFunctionLists.AddItem(ucase(sFilename), trim(sAllFunctionNames,DELIM))

   Function = 0

End Function

