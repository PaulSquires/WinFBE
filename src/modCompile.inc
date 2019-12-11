''    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
''    Copyright (C) 2016-2019 Paul Squires, PlanetSquires Software
''
''    This program is free software: you can redistribute it and/or modify
''    it under the terms of the GNU General Public License as published by
''    the Free Software Foundation, either version 3 of the License, or
''    (at your option) any later version.
''
''    This program is distributed in the hope that it will be useful,
''    but WITHOUT any WARRANTY; without even the implied warranty of
''    MERCHANTABILITY or FITNESS for A PARTICULAR PURPOSE.  See the
''    GNU General Public License for more details.

''
''
''   modCompile.inc
''
''   Code here handles the entire compiling process. It prepares the necessary
''   command line data to pass to the applicable 32 or 64 version of the FBC
''   compiler. Compile time errors are parsed through analyis of the log file
''   and presented to the user via the Compiler Results listview and the Compiler
''   Lof File text box.
''
''

#include once "modCompile.bi"
#include once "modGenerateCode.bi"
#include once "frmUserTools.bi"
#include once "frmOutput.bi"


Dim Shared gCompile As COMPILE_TYPE


''
''  Set the statusbar text and icon for a good/bad compile
''
private function SetCompileStatusBarMessage( byref wszText as wstring, _
                                             byval hIconCompile as HICON _
                                             ) as LRESULT
   gApp.wszPanelText = wszText
   gApp.hIconPanel = hIconCompile
   frmMain_SetStatusbar
   function = 0
end function


''
''  Create Resource file (if necessary) for the compile and return
''  the full disk filename of the resource file.
''
private function CreateTempResourceFile() as boolean
   dim as CWSTR wszResourceFile = gCompile.ResourceFile
   
   ' Collect all of the IMAGE_NAME used by controls/properties in the file/project
   ' and generate a string to output to the resource file.
   dim pDoc as clsDocument ptr
   dim as CWSTR wszInternalText, wszExistingText
   
   pDoc = gApp.pDocList
   do until pDoc = 0
      for i as long = lbound(pDoc->AllImages) to ubound(pDoc->AllImages)
         ' IMAGE_ARROWLEFT    RCDATA "IMAGES\\LEFTARROW.PNG"
         wszInternalText = wszInternalText & _
                   pDoc->AllImages(i).wszImageName & wspace(6) & _
                   pDoc->AllImages(i).wszFormat    & wspace(2) & _
                   wchr(34) & pDoc->AllImages(i).wszFileName & wchr(34) & _
                   vbcrlf
      next   
      pDoc = pDoc->pDocNext
   loop

   ' If resource items were found then we need to generate the new resource file
   ' and return its name. The new file will be a temporary file that will need to
   ' added to the App.AddQuickRunEXE list. That list is checked in the main WinFBE
   ' message loop and deletes not only QuickRun EXEs but any other WinFBE generated
   ' temporary file that we want to get rid of.
   if len(wszInternalText) then

      ' Need to copy a default manifest file to the compiling folder as well because
      ' the newly generated resource file will depend on it.
      dim as CWSTR wszManifestTemplate = AfxGetExePathName + "Settings\WinFBE_manifest.xml"
      dim as CWSTR wszManifest = gCompile.MainFolder + "manifest.xml"
      if AfxFileExists(wszManifestTemplate) then
         if AfxFileExists(wszManifest) = false then
            AfxCopyFile( wszManifestTemplate, wszManifest, true )
         end if
      end if

      ' If a resource file already exists then use whatever text is in the file as
      ' the basis for the newly generated resource file. 
      if AfxFileExists(gCompile.ResourceFile) then
         ' Load the entire file into a string
         DIM dwCount AS DWORD, dwFileSize AS DWORD, dwHighSize AS DWORD, dwBytesRead AS DWORD
         DIM hFile AS HANDLE = CreateFileW(@gCompile.ResourceFile, GENERIC_READ, FILE_SHARE_READ, NULL, _
                                            OPEN_EXISTING, FILE_FLAG_SEQUENTIAL_SCAN, NULL)
         IF hFile = INVALID_HANDLE_VALUE THEN return false
         dwFileSize = GetFileSize(hFile, @dwHighSize)
         dim as string txtBuffer = String(dwFileSize, 0)
         ReadFile(hFile, strptr(txtBuffer), dwFileSize, @dwBytesRead, NULL)
         CloseHandle(hFile)

         ' Check for BOM signatures
         if left(txtBuffer, 3) = chr(&HEF, &HBB, &HBF) THEN
            ' UTF8 BOM encoded 
            wszExistingText = Utf8ToUnicode( mid(txtBuffer, 4) )   ' bypass the BOM
         elseif left(txtBuffer, 2) = chr(&HFF, &HFE) THEN
            ' UTF16 BOM (little endian) encoded
            wszExistingText = mid(txtBuffer, 3)   ' bypass the BOM
         else
            wszExistingText = wstr(txtBuffer)   ' ansi -> unicode
         END IF

      end if   
      
      ' Ensure that the reference to the manifest file exists
      dim as CWSTR wszLookFor = "1 24 " & wchr(34) & ".\manifest.xml" & wchr(34)
      if instr(wszExistingText, wszLookFor ) = 0 then
         wszInternalText = wszLookFor & vbcrlf & wszInternalText  
      end if
      
      ' Append the internally generated resource text to the new resource file.
      gCompile.TempResourceFile = GetTemporaryFilename(gCompile.MainFolder, "rc")
      dim pStream AS CTextStream   '(utf16)
      if pStream.Create(gCompile.TempResourceFile, true, true) = S_OK then 
         pStream.WriteLine wszExistingText
         pStream.WriteLine ""
         pStream.WriteLine wszInternalText
         pStream.WriteLine ""
         pStream.Close 
      end if
      
      return true
   end if
   
   return false
end function


''
''  Reset all Scintilla editing cursors
''
public function ResetScintillaCursors() as Long
   Dim pDoc As clsDocument Ptr
   Dim As Long nCount = TabCtrl_GetItemCount(gTTabCtl.hWindow) 
   Dim TB As TC_ITEM 

   For i as long = 0 To nCount - 1
      ' Get the document pointer and then save file to disk
      TB.Mask = TCIF_PARAM
      TabCtrl_GetItem( gTTabCtl.hWindow, i, @TB )
      pDoc = Cast(clsDocument Ptr, TB.lParam)
      if pDoc THEN 
         SciExec( pDoc->hWindow(0), SCI_SETCURSOR, SC_CURSORNORMAL, 0 )
         SciExec( pDoc->hWindow(1), SCI_SETCURSOR, SC_CURSORNORMAL, 0 )
      end if
   next
   function = 0
end function


'' July 2017, attempted to capture STDERR output from a running FB compiled
'' program in order to capture runtime errors. This code works okay on Windows 10
'' but does not PRINT on Windows 7. Also, dkl posted in FB forum that FB prints
'' to STDOUT rather than STDERR. This is too bad.

''
''
private function AlternateCompile( byref wszExe AS wstring, _
                                   byref wszCmdLine AS wstring, _
                                   byref sConsoleText AS string _
                                   ) as long

   Dim ShExecInfo As SHELLEXECUTEINFOW  

   ' Create a temporary batch that will contain the compile commands. We shell
   ' to and run this batch file in order to do the compile.
   ' Replace any embedded spaces with underscores
   dim as CWSTR wDQ = wchr(34)
   dim as CWSTR wszBatchFile = "_winfbe_compile.bat"
   dim as CWSTR wszLogFile   = "_winfbe_compile.log"
   
   ' Change to the output code directory
   'ChDir gCompile.MainFolder
   
   ' Must delete any previous instance of the compiler log file in case it was not deleted. Need
   ' to do this because the compiler appends to the log file rather than overwriting.
   if AfxFileExists(wszLogFile) then AfxDeleteFile(wszLogFile)
   
   dim as long f = Freefile
   Open wszBatchFile For Output as #f
   print #f, wDQ & wszEXE & wDQ & " " & wszCmdLine;
   Print #f, " >> "; wszLogFile; " 2>&1"   ' ensure stderr is redirected (appended) to log file as well
   close #f
   
   dim as CWSTR wszShellExe = Environ("COMSPEC")
   dim as CWSTR wszShellParam   = " /C " & wszBatchFile

   With ShExecInfo
      .cbSize       = Len(SHELLEXECUTEINFOW)
      .fMask        = SEE_MASK_NOCLOSEPROCESS
      .HWnd         = 0
      .lpVerb       = Null
      .lpFile       = wszShellExe.sptr
      .lpParameters = wszShellParam.sptr   
      .lpDirectory  = 0
      .nShow        = SW_HIDE
      .hInstApp     = 0 
   End With

   ' Shell to the compiler 
   ShellExecuteEx( @ShExecInfo )
    
   ' Give the process 33 ms (you can use a different value here depending on 
   ' how responsive you wish your app to be)
   Dim uMsg As MSG
   while WaitForSingleObject(ShExecInfo.hProcess, 33) = WAIT_TIMEOUT
   '  if after 33 ms the object's handle is not signaled..       
   '  we examine the message queue and if there are any waiting..
   '  Note:  see PeekMessage documentation for details on how to limit
   '         the types of messages to look for
      while(PeekMessage(@uMsg, HWND_FRMMAIN,  0, 0, PM_NOREMOVE))
        ' we process them..
        if(GetMessage(@uMsg, NULL, 0, 0) > 0) then
            TranslateMessage(@uMsg)
            DispatchMessage(@uMsg)
         end if
      wend
   wend
   
   ' Get the log file contents into a string in order to parse it for errors, etc
   If AfxFileExists(wszLogFile) Then
      f = Freefile
      Open wszLogFile For binary as #f
      sConsoleText = string(lof(f), 0)
      Get #f, , sConsoleText
      Close #f
   End If   

? "--sConsoleText--"
? sConsoleText

'   if AfxFileExists(wszBatchFile) then AfxDeleteFile(wszBatchFile)
'   if AfxFileExists(wszLogFile) then AfxDeleteFile(wszLogFile)

   function = 0
end function


''
''
private function RedirConsoleToFile( byref wszExe AS wstring, _
                                     byref wszCmdLine AS wstring, _
                                     byref sConsoleText AS string _
                                     ) as long

   ' It appears that CreateProcess will fail if szExe is wrapped in quotes or if
   ' it contains spaces in the filename. Use the older Batch file method instead.
   ' This means that the WinFBE installation and the location of the gcc tools should not
   ' have any embedded spaces in their folder paths. 
   if instr(wszExe, " ") then
      AlternateCompile( wszExe, wszCmdLine, sConsoleText )
      return 0
   end if   

   ' Continue on with the newer CreateProcess compile approach.
   dim SecurityAttribute AS SECURITY_ATTRIBUTES
   dim hChildStdOutRead  AS PHANDLE
   dim hChildStdOutWrite AS PHANDLE
   dim dwReadBytes       AS long
   dim sBuffer AS STRING * 4096
  
   ' Set the bInheritHandle flag so pipe handles are inherited.
   SecurityAttribute.nLength              = SIZEOF(SECURITY_ATTRIBUTES)
   SecurityAttribute.bInheritHandle       = TRUE
   SecurityAttribute.lpSecurityDescriptor = NULL

   ' Create a pipe for the child process's STDOUT.
   IF CreatePipe(@hChildStdOutRead, @hChildStdOutWrite, @SecurityAttribute, BYVAL 0) = FALSE THEN
      ? "error creating pipe"
   ELSE
      ' Ensure the read handle to the pipe for STDOUT is not inherited.
      IF SetHandleInformation(hChildStdOutRead, HANDLE_FLAG_INHERIT, BYVAL 0) = FALSE THEN
         ? "error SetHandleInformation"
      ELSE
         ' Create the child process and read from its standard output
         dim ProcessInfo AS PROCESS_INFORMATION
         dim StartupInf  AS STARTUPINFO

         StartupInf.cb         = SIZEOF(STARTUPINFO)
         StartupInf.hStdError  = hChildStdOutWrite   
         StartupInf.hStdOutput = hChildStdOutWrite  
         StartupInf.dwFlags    = STARTF_USESTDHANDLES

         dim wszFullExe as wstring * 32767
         
         if instr(wszExe, " ") then
            wszFullExe = wchr(34) & wszExe & wchr(34)
         else
            wszFullExe = wszExe 
         end if
         
' This call to CreateProcessW will succeed in all cases except for the one
' where there is spaces in wszFullExe because once I wrap the string in double
' quotes then the call to CreateProcessW fails.

         dim errCode as long 
         errCode = CreateProcessW( wszFullExe,  _    ' Create the child process.
                     wszCmdLine, _       ' Command line
                     BYVAL 0,     _  ' Process security attributes
                     BYVAL 0,     _  ' Primary thread security attributes
                     TRUE,        _  ' Handles are inherited
                     CREATE_NO_WINDOW, _  ' Creation flags  
                     BYVAL 0,     _  ' Use parent's environment
                     BYVAL 0,     _  ' Use parent's current directory
                     @StartupInf,  _  ' STARTUPINFO pointer
                     @ProcessInfo) 
         
         'if errCode = 0 then
         '   ? wszFullExe
         '   ? wszCmdLine
         '   ? "CreateProcess FAILED errCode: "; errCode, "GetLastError: "; GetLastError()
         '   'ERROR_INVALID_NAME
         '   '    123 (0x7B)
         '   '    The filename, directory name, or volume label syntax is incorrect.
         'end if


         if errCode <> 0 then            

            CloseHandle(hChildStdOutWrite) ' To avoid ReadFile hanging at the end.
            DO
               IF ReadFile(hChildStdOutRead, BYVAL VARPTR(sBuffer), 4096, @dwReadBytes, BYVAL 0) = FALSE THEN 
                  '? "ERROR ReadFile. hChildStdOutRead=";hChildStdOutRead
                  '? "GetLastError: "; GetLastError
                  EXIT DO
               end if  
               sConsoleText = sConsoleText & LEFT(sBuffer, dwReadBytes)
               '(at this stage we could SAVE the buffer that we've read from the child to a file, but 
               ' thats an extra ReadFile+WriteFile compared to the direct-to-hFile method)
            LOOP
            
         else
            '? "CreateChildProcess failed"      
            if instr( wszFullExe, " " ) then
               AfxMsg( "CreateChildProcess failed. Install WinFBE into folder without spaces." )
            end if    
            ' Close handles to the child process and its primary thread.
            ' Some applications might keep these handles to monitor the status 
            ' of the child process, for example.
            CloseHandle( ProcessInfo.hProcess )
            CloseHandle( ProcessInfo.hThread )
         END IF
         
      END IF
   END IF

   function = 0
END function


''
''
public Function RunEXE( ByRef wszFileExe As CWSTR, _
                        ByRef wszParam As CWSTR _
                        ) As Long
                         
   dim as CWSTR wszPath = AfxStrPathname("PATH", wszFileExe)
   dim as CWSTR wszBatchFile = AfxGetExePathName + AfxStrPathname("NAME", wszFileExe) + ".bat"
   dim as CWSTR wszQuickRunExe = wszFileExe
   
   ' If option is specified to launch command window prior to running the
   ' compiled program then we do so by executing a shell to a batch file.
   ' This allows being able to see any runtime errors that might be thrown.
   ' Otherwise, simple shell to the program.
   if gConfig.RunViaCommandWindow THEN
      dim pStream as CTextStream
      if pStream.Create(wszBatchFile) = S_OK then
         pStream.WriteLine "@echo off"
         pStream.WriteLine Left(**wszPath, 2)     ' ensure on correct drive
         pStream.WriteLine "cd " & chr(34) & wszPath & chr(34)   ' change to correct folder
         pStream.WriteLine chr(34) & wszFileExe & chr(34) & " " & wszParam
         pStream.WriteLine "pause"
         pStream.WriteLine "del " & chr(34) & wszBatchFile & chr(34) & " >> nul"
      end if
      pStream.Close
      wszFileExe = wszBatchFile
      wszParam = ""
   END IF
   

   Dim ShExecInfo As SHELLEXECUTEINFOW  

   ' Run the EXE
   With ShExecInfo
      .cbSize       = Len(SHELLEXECUTEINFOW)
      .fMask        = SEE_MASK_NOCLOSEPROCESS
      .HWnd         = 0
      .lpVerb       = Null
      .lpFile       = wszFileExe
      .lpParameters = wszParam   
      .lpDirectory  = 0
      .nShow        = SW_SHOW
      .hInstApp     = 0 
   End With
   ShellExecuteEx(@ShExecInfo)

   ' If this was a QuickRun then add the filename to the global tracking array
   ' that checks if the file exists and then deletes it.
   if gCompile.CompileID = IDM_QUICKRUN THEN
      ' Give time for the application to start to run (especially if being run via
      ' the RunViaCommandWindow batch file, otherwise the exe will be deleted even
      ' before it gets a chance to execute.
      sleep 500
      gApp.AddQuickRunEXE(wszQuickRunExe)
   END IF
   
   Function = 0
End Function



' ========================================================================================
' Set the cursor to the error position based on the selected line in frmCompileResults
' or the TODO listview. 
' ========================================================================================
public Function SetDocumentErrorPosition( ByVal hLV As HWnd, _
                                          Byval wID as long _
                                          ) As Long

   Dim wszErrorLine As WString * MAX_PATH
   Dim wszErrorFile As WString * MAX_PATH
   
   Dim As Long nCurSel = ListView_GetSelection(hLV)
   If nCurSel < 0 Then Return 0 
   
   Dim pDoc As clsDocument Ptr 

   ' Get the line number and filename of the selected item
   FF_ListView_GetItemText(hLV, nCurSel, 0, @wszErrorLine, MAX_PATH)
   FF_ListView_GetItemText(hLV, nCurSel, 1, @wszErrorFile, MAX_PATH)
   dim as long nLineNum = Val(wszErrorLine) - 1   ' because visually the editor is one based line numbers

   ' If we are doing a QuickRun then it must be on the currently loaded
   ' and active file in the editor, therefore simply use that pDoc.
   if wID = IDM_QUICKRUN THEN
      pDoc = gTTabCtl.GetActiveDocumentPtr()
      if pDoc then wszErrorFile = pDoc->DiskFilename 
   END IF
   
   pDoc = OpenSelectedDocument( wszErrorFile, "", nLineNum )
   
   ' Get the pDoc for this tab and check to see if it is a visual designer. If it is, then
   ' initiate the toggle between the code editor and visual designer.
   if pDoc then
      if pDoc->IsDesigner then
         ' Show the code editor.
         TabCtrl_SetCurSel( pDoc->hDesignTabCtrl, 1 )  ' 1 = code tab
         frmMain_PositionWindows
      end if
   END IF

   function = 0
End Function


''
''
private function SetLogFileTextbox() as long
   
   ' Construct the text that will be displayed:
   ' - FBC version, exe name, file size, compile time
   ' - raw log file output
   ' - compile command line used
   dim as CWSTR wszText
 
   if len(gCompile.wszOutputMsg) then 
      gCompile.wszOutputMsg = gCompile.wszOutputMsg + vbcrlf
   end if
   wszText = _
      gCompile.wszOutputMsg + _
      L(178,"Command Line") + ": " + vbcrlf + _
      gCompile.wszFullCommandLine + vbcrlf + vbcrlf + _
      gCompile.wszFullLogFile 
             
   AfxSetWindowText( GetDlgItem( HWND_FRMOUTPUT, IDC_FRMOUTPUT_TXTLOGFILE), wszText )

   ' Reset the type parameters so subsequent compiles will not append
   function = 0
end function


''
''
private function ParseLogForError( byref wsLogSt as CWSTR, _
                                   byval bAllowSuccessMessage as Boolean, _
                                   Byval wID as long, _
                                   byval fCompile64 as Boolean, _
                                   byval fCompilingObjFiles as Boolean _
                                   ) as Boolean
   
   ' Returns TRUE if error has been detected. This signals to the calling code
   ' that we can break out of performing any further compiles (eg. modules).
   Dim wszTemp As WString * MAX_PATH 

   Dim As Long Parenthesis_Start, Parenthesis_End, Error_Start       
   Dim As Long i, NumLines, NextLine, r, nCount, nListViewLine
   Dim As Long NumWarnings, NumErrors, NumDirectives, IsError, IsWarning
   Dim As Long nFirstErrorLine = -1
   Dim As HWnd hLV
   Dim As CWSTR wDQ, wst1, wst2, wst3, wst, wst_ucase, wszOutputMsg        
   dim as Boolean fGoodCompile

   wDQ = wchr(34) ' unicode double quotes

   hLV = GetDlgItem(HWND_FRMOUTPUT, IDC_FRMOUTPUT_LISTVIEW)
   ListView_DeleteAllItems( hLV )
   
   ' On some Windows systems with tight security policies, the CreateChildProcess used
   ' by the sub RedirConsoleToFile will fail. This seems to happen mostly when trying
   ' to compile using the FB 64 bit compiler on a 32-bit version of Windows. For this
   ' reason we will do a manual check here rather than relying on detecting the error
   ' through the log file string.
   if fCompile64 then   ' trying to compile a 64-bit program
      if AfxWindowsBitness() <> 64 then  ' Windows is not 64 bit
         wsLogSt = "This version of the FreeBASIC compiler is not compatible with the version " & _
                   "of Windows you're running. Check your computer's system information to see " & _
                   "whether you need a x86 (32-bit) or x64 (64-bit) version of the program, and " & _
                   "then contact the software publisher."  
         gCompile.wszFullLogFile = wsLogSt
      end if
   end if

   dim as Boolean bReadingCompileOutput = false
   
   ' Parse the log string.
   NumLines = AfxStrParseCount( wsLogSt, vbCrLf)
 
   For NextLine = 1 To NumLines                   
   
      wst = Trim(AfxStrParse(wsLogSt, NextLine, vbCrLf))                 
      wst_ucase = Ucase(wst)
      
      ' Deal with the situation where we might be trying to compile a 64-bit application
      ' from a 32-bit operating system.
      If Left(**wst_ucase, 16) = WSTR("THIS VERSION OF ") Then
         fGoodCompile = false
         FF_ListView_InsertItem( hLV, NumWarnings, 0, WStr("0") )
         FF_ListView_InsertItem( hLV, NumWarnings, 1, WStr("") ) 
         FF_ListView_InsertItem( hLV, NumWarnings, 2, WStr("compiling FAILED: Windows not 64-bit (refer to log file)") ) 
         wszOutputMsg = wsLogSt: exit For
         exit for
      End If   

      ' Save the FB version and copyright information. 
      If Left(**wst_ucase, 19) = WSTR("FREEBASIC COMPILER ") Then
         fGoodCompile = True
         wszOutputMsg = wszOutputMsg & wst & vbCrLf: Continue For
      End If        
      If Left(**wst_ucase, 13) = WSTR("COPYRIGHT (C)") Then
         fGoodCompile = True
         wszOutputMsg = wszOutputMsg & wst & vbCrLf: Continue For
      End If        
      
      ' Check to see if an error occurred in compiling a resource script.
      ' If there was a bad source name passed to the compiler (for 
      ' example, missing .rc file), then the output at the end of the
      ' log file is like the following:
      '
      ' Error!
      ' Could Not Open source file (p.RC)
      ' OBJ file Not made
      ' compiling resource FAILED: Error Code 1
      '
      ' Very sorry, but GoRC had a problem.
                                          
      ' Check to see if linking failed
      If (Left(**wst_ucase, 6) = WSTR("ERROR!")) or _
         (left(**wst_ucase, 34) = "VERY SORRY, BUT GORC HAD A PROBLEM") Then
         FF_ListView_InsertItem( hLV, nListViewLine, 0, WStr("0") )
         FF_ListView_InsertItem( hLV, nListViewLine, 1, WStr("") ) 
         FF_ListView_InsertItem( hLV, nListViewLine, 2, WStr("compiling resource FAILED: Error Code 1 (refer to log file)") ) 
         NumErrors = NumErrors + 1
      elseIf (Left(**wst_ucase, 5) = WSTR("ERROR")) Then
         FF_ListView_InsertItem( hLV, nListViewLine, 0, WStr("0") )
         FF_ListView_InsertItem( hLV, nListViewLine, 1, WStr("") ) 
         FF_ListView_InsertItem( hLV, nListViewLine, 2, WStr("compiling FAILED: Error Code 1 (refer to log file)") ) 
         NumErrors = NumErrors + 1
      elseIf (Left(**wst_ucase, 19) = WSTR("COMPILING C FAILED:")) Then
         FF_ListView_InsertItem( hLV, nListViewLine, 0, WStr("0") )
         FF_ListView_InsertItem( hLV, nListViewLine, 1, WStr("") ) 
         FF_ListView_InsertItem( hLV, nListViewLine, 2, WStr("compiling C FAILED: Error Code 1 (refer to log file)") ) 
         NumErrors = NumErrors + 1
      end if
      
      If Instr(**wst_ucase, WSTR("LINKING FAILED:")) Then 
         FF_ListView_InsertItem( hLV, nListViewLine, 0, WStr("0") ) 
         FF_ListView_InsertItem( hLV, nListViewLine, 1, WStr("") ) 
         FF_ListView_InsertItem( hLV, nListViewLine, 2, wst & " (refer to log file)" ) 
         NumErrors = NumErrors + 1
      End If
 
 
      ' Check for the "linking: " line because that line identifies the actual output filename being created
      If Left(**wst_ucase, 9) = WSTR("LINKING: ") Then 
         ' -o "WinFBE.exe" 
         i = Instr(wst, WSTR(" -o "))
         If i Then
            wszTemp = Mid(**wst, i+4)
            i = Instr(wszTemp, wDQ & WSTR(" "))
            If i Then gCompile.OutputFilename = AfxStrRemove(Left(wszTemp, i), wDQ)
         End If      
      End If
      
      ' Determine if we are reading any lines in the log file that appear after the "compiling: " line
      ' but before the "assembling: ". Anything there is data output by #print preprocessor statements.
      If Left(**wst_ucase, 11) = WSTR("COMPILING: ") Then 
         bReadingCompileOutput = true
         continue for
      elseIf Left(**wst_ucase, 13) = WSTR("COMPILING C: ") Then 
         bReadingCompileOutput = true
         continue for
      elseIf Left(**wst_ucase, 12) = WSTR("ASSEMBLING: ") Then 
         bReadingCompileOutput = false
      elseIf Left(**wst_ucase, 14) = WSTR("COMPILING RC: ") Then 
         bReadingCompileOutput = false
      elseIf Left(**wst_ucase, 9) = WSTR("LINKING: ") Then 
         bReadingCompileOutput = false
      end if

      if bReadingCompileOutput then

         if len(wst) = 0 then continue for

         ' Check for any compiler warnings
         Error_Start = 0
         IsWarning = Instr(**wst_ucase, WSTR(") WARNING "))
         If IsWarning Then Error_Start = IsWarning
         IsError   = Instr(**wst_ucase, WSTR(") ERROR ")) 
         If IsError Then Error_Start = IsError
         
         If (Error_Start > 0) Then 
            'sample warning message
            'c:\freebasic\test.bas(1394) warning 1(1): Passing scalar as pointer, at parameter 2 (hwndOldFocus) of ONSETFOCUS()
            'sample error message
            'c:\freebasic\test.bas(17) error 41: Variable not declared, kjljjada in 'kjljjada Error'

            ' Determine the line number. Error_Start variable holds the position of the closing
            ' parenthesis of the line number. We simply need to reverse search from there for the
            ' opening parenthesis.
            ' 2018-08-29: Code updated to handle case of embedded parenthesis in file name
            Parenthesis_End   = Error_Start
            Parenthesis_Start = InStrRev(**wst, WSTR("("), Parenthesis_End)
            
            If (Parenthesis_Start < Parenthesis_End) Andalso _
               (Parenthesis_End <= Error_Start) then
       
               wst1 = left(**wst, Parenthesis_Start-1)   ' filename
               wst2 = Mid( **wst, Parenthesis_Start + 1, Parenthesis_End - Parenthesis_Start - 1)  ' line# 
               wst3 = Mid( **wst, Error_Start + 1)  ' error message
     
               ' If this is a visual designer form file then we need to modify the line
               ' number based on the error offset number.
               dim pDoc as clsDocument ptr = gApp.GetDocumentPtrByFilename(wst1)
               if pDoc then
                  if pDoc->IsDesigner then
                     if len(pDoc->wszFormMetaData) = 0 then GenerateFormMetaData(pDoc)   ' recreate the metadata
                     if len(pDoc->wszFormCodegen) = 0 then GenerateFormCode(pDoc)   ' regenerate visual designer code
                     ' Offset is based on form metadata + code generation
                     pDoc->ErrorOffset = 0
                     pDoc->ErrorOffset = pDoc->ErrorOffset + AfxStrTally(pDoc->wszFormMetaData, vbcrlf) 
                     pDoc->ErrorOffset = pDoc->ErrorOffset + AfxStrTally(pDoc->wszFormCodegen, vbcrlf) 
                     wst2 = str( val(wst2) - pDoc->ErrorOffset)
                  end if
               END IF

               FF_ListView_InsertItem( hLV, nListViewLine, 0, wst2 )  ' line#
               FF_ListView_InsertItem( hLV, nListViewLine, 1, wst1 )  ' filename
               FF_ListView_InsertItem( hLV, nListViewLine, 2, wst3 )  ' error message
                
               If IsWarning Then 
                  NumWarnings = NumWarnings + 1
               ElseIf IsError Then
                  If nFirstErrorLine = -1 Then
                     nFirstErrorLine = NumWarnings + NumErrors
                  End If
                  NumErrors = NumErrors + 1
               End If            
            End If

         else
                     
            FF_ListView_InsertItem( hLV, nListViewLine, 0, wstr("") )  ' line#
            FF_ListView_InsertItem( hLV, nListViewLine, 1, wstr("") )  ' filename
            FF_ListView_InsertItem( hLV, nListViewLine, 2, wst )  ' message
            NumDirectives = NumDirectives + 1
         end if

      end if
 
      nListViewLine = NumWarnings + NumErrors + NumDirectives
   Next
 
   ' If there were no errors but the fGoodCompile flag was not set to True then the log file
   ' did not contain the required FB copyright notice. Something must have went wrong, like trying
   ' to call the 64 bit compiler using a 32 bit operating system. Show the message to the user via
   ' the Output window (log file).
   If fGoodCompile = False Then NumErrors = NumErrors + 1


   dim as CWSTR cwsCompileMsg

   ' If the EXE output file was never created then throw an error
   if fCompilingObjFiles = false then
      If (NumWarnings = 0) andalso (NumErrors = 0) then
         if FileLen(gCompile.OutputFilename) = 0 then 
            NumErrors = NumErrors + 1
         end if
      end if
   end if
   
   ' In all cases where warnings and/or errors exist, we will show the Compiler Results listview
   If (NumWarnings > 0) orelse (NumErrors > 0) orelse (NumDirectives > 0) then
   
      cwsCompileMsg = L(193, "Errors")   & " " & NumErrors   & "  " & _
                      L(192, "Warnings") & " " & NumWarnings & _
                      "  [" & _
                        AfxLocalDateStr( "yyyy-MM-dd" ) & "  " & _
                        AfxLocalTimeStr( "hh:mm:ss" ) & "]"
      gCompile.wszOutputMsg = L(229,"Failed Compile")  & " (" & cwsCompileMsg & ")" & vbcrlf
      SetCompileStatusBarMessage( cwsCompileMsg, ghIconBad )
      MessageBeep(MB_ICONWARNING)
      
      ' Position the Compiler Log to the first error/warning
      ListView_SelectItem( hLV, 0 )
      
      ' If fGoodCompile is False then something unusual occurred so better show the log file by default
      If fGoodCompile = False Then
         frmOutput_PositionWindows
      End If
       
      ' If the Search Results, TODO, Notes tab is active then reposition to 
      ' the first tab for error listview
      select case TabCtrl_GetCurSel( GetDlgItem(HWND_FRMOUTPUT, IDC_FRMOUTPUT_TABCONTROL) ) 
         case 2, 3, 4
         TabCtrl_SetCurSel( GetDlgItem(HWND_FRMOUTPUT, IDC_FRMOUTPUT_TABCONTROL), 0 ) 
         frmOutput_PositionWindows
      end select
      
      ShowWindow( HWND_FRMOUTPUT, SW_SHOW)
      frmMain_PositionWindows
      ' Set to error line position only after all windows have been shown and resized
      SetDocumentErrorPosition(hLV, wID)
   end if
   
   ' If there were no errors or warnings then close any previously open compiler results listviews
   If (NumWarnings = 0) andalso (NumErrors = 0) andalso (NumDirectives = 0) then
      if IsWindowVisible(HWND_FRMOUTPUT) THEN
         if TabCtrl_GetCurSel( GetDlgItem(HWND_FRMOUTPUT, IDC_FRMOUTPUT_TABCONTROL) ) = 0 THEN
            ShowWindow( HWND_FRMOUTPUT, SW_HIDE )
            frmMain_PositionWindows
         END IF   
      END IF 
   end if

   If NumErrors = 0 Then 
      ' 2018-08-29: Show popup compile message in cases where no errors but warnings may exist.
      ' Hide the Output window is already open but there is no longer any warnings
      ' or errors to show in the listview. Only close if the listview is active. Do not
      ' close if the user is looking at the logfile textbox.
      if bAllowSuccessMessage then  ' we are not compiling modules
         dim as long nFileSize = AfxGetFileSize(gCompile.OutputFilename)
         cwsCompileMsg = L(193, "Errors") & " " & NumErrors & "  " & _
                         L(192, "Warnings") & " " & NumWarnings & _
                         "  [ " & AfxStrPathname("NAMEX", gCompile.OutputFilename) & "  " & _
                         **AfxStrFormatKBSize(nFileSize) & " (" & _
                         nFileSize & " " & L(199,"bytes") & ") ]"
         SetCompileStatusBarMessage( cwsCompileMsg, ghIconGood )
         ' 2018-12-13: No longer offer the option to show the popup successful dialog.
         ' Simply sound a success
         MessageBeep(MB_OK)

         gCompile.wszOutputMsg = _
         L(228,"Successful Compile")  & " (" & _
               L(193,"Errors")        & " " & NumErrors & " " & _
               L(192,"Warnings")      & " " & NumWarnings & ")" & vbcrlf & vbcrlf & _
         L(194,"Primary Source:")     & " " & gCompile.MainFilename & vbCrLf & _
         L(195,"Target Compilation:") & " " & gCompile.OutputFilename & " (" & _
               **AfxStrFormatKBSize(nFileSize) & ", " & nFileSize & " " & L(199,"bytes") & ")" & vbcrlf & _
         L(196,"Compile Time:")       & " " & Format(gCompile.EndTime-gCompile.StartTime, "###0.0") & _
                                        " " & L(198,"seconds") & " (" & _
                                        AfxSystemTimeToDateStr( gCompile.SystemTime, "yyyy-MM-dd" ) & " " & _
                                        AfxSystemTimeToTimeStr( gCompile.SystemTime, "hh:mm:ss" ) & ")" & vbcrlf
      end if 
   End if

   ' Update the log file text box in the frmOutput window
   SetLogFileTextbox

   ' If Errors have occurred then return TRUE. We will allow processing
   ' to continue if only Warnings occurred.
   ' There were errors then do not allow running the EXE (warnings are okay)
   If NumErrors > 0 Then 
      gCompile.RunAfterCompile = False 
      return true
   end if
   
   function = false
end function


''
'' Main module that handles the entire compile process
''
public Function code_Compile( ByVal wID As Long ) As BOOLEAN

   Dim pDoc As clsDocument Ptr
   
   Dim wszTemp      As WString * MAX_PATH 
   Dim wszFileExe   As WString * MAX_PATH
   dim wszCompile   as CWSTR
   dim wszCommand   as CWSTR 
   dim wszParams    as CWSTR
   dim sConsoleText as string 
    
   Dim As CWSTR wExeCmd, wDQ, wsLogSt, wst
         
   Dim As BOOLEAN fCompile32, fCompile64
   
   dim as long idxBuild, f, i, n, nCount
   
   wDQ = wchr(34) ' unicode double quotes
                             
   ' Reload the config file in case a user has automated a change to it since
   ' the application started (for example, changing the compiler path).
   gConfig.LoadConfigFile()
   
   gCompile.wszFullCommandLine = ""
   gCompile.wszFullLogFile = ""
   gCompile.wszOutputMsg = ""
  
   ' Can only continue to compile if the primary source code file
   ' is not dirty and needs to be saved.
   if wID <> IDM_QUICKRUN THEN
      If gConfig.CompileAutosave Then
         ' If there any "New" files that have not been edited yet causing them to show 
         ' as dirty then we will mark them now so that the FileSaveAll command will
         ' ask to save it (otherwise a compile error will show).
         dim pDoc as clsDocument ptr = gApp.pDocList
         do until pDoc = 0
            if pDoc->IsNewFlag then pDoc->UserModified = true
            pDoc = pDoc->pDocNext
         loop
         If OnCommand_FileSaveAll(HWND_FRMMAIN) Then 
            Return True  ' user cancelled save
         end if   
         ' Reload the main form files combobox in case names have changed
         frmMain_LoadComboFiles
         frmMain_SelectComboFiles
      End If
   END IF
   
   ' Get the array index for the current selected build (project or document)
   idxBuild = ComboBox_GetCurSel(HWND_FRMMAIN_COMBOBUILDS)
   if idxBuild = -1 then idxBuild = 0
   
   ' Perform some pre-compile checks to see if we should continue.
   gConfig.FBWINCompiler32 = ProcessFromCurdrive(gConfig.FBWINCompiler32)
   gConfig.FBWINCompiler64 = ProcessFromCurdrive(gConfig.FBWINCompiler64)
   if gConfig.Builds(idxBuild).Is32bit then
      fCompile32 = True: fCompile64 = False
      gCompile.CompilerPath = gConfig.FBWINcompiler32
   end if
   if gConfig.Builds(idxBuild).Is64bit then
      gCompile.CompilerPath = gConfig.FBWINcompiler64
      fCompile32 = False: fCompile64 = True
   End if

   ' Convert relative path to absolute path if needed.
   if AfxPathIsRelative(gCompile.CompilerPath) then
      gCompile.CompilerPath = AfxPathCombine(AfxGetExePathName, gCompile.CompilerPath)
   END IF


   ' Check to see if the compiler exists
   If AfxFileExists(gCompile.CompilerPath) = 0 Then 
      wszTemp = L(202,"Invalid defined compiler path.")
      wszTemp = wszTemp & vbcrlf & "Path: " & gCompile.CompilerPath
      MessageBox( HWND_FRMMAIN, wszTemp, L(201,"Error"), _
                        MB_OK Or MB_ICONINFORMATION Or MB_DEFBUTTON1 )
      Function = False: Exit Function
   End If
   
   
   ' Set some compile flags depending on the type of compile requested.
   gCompile.CompileID = wID
   Select Case wID
      Case IDM_BUILDEXECUTE: gCompile.RunAfterCompile = True
      Case IDM_QUICKRUN:     gCompile.RunAfterCompile = True
      Case IDM_COMPILE:      gCompile.RunAfterCompile = False
   End Select
   
   
   Dim pDocMain As clsDocument Ptr 
   If gApp.IsProjectActive Then
      pDocMain = gApp.GetMainDocumentPtr()
      If pDocMain = 0 Then
         MessageBox( HWND_FRMMAIN, L(208,"No Main file specified for the project."), L(201,"Error"), _
                           MB_OK Or MB_ICONINFORMATION Or MB_DEFBUTTON1 )
         Function = False: Exit Function
      End If
   Else
      pDocMain = gTTabCtl.GetActiveDocumentPtr()
   End If
   If pDocMain = 0 Then return 0
   

   ' If this is a QuickRun then we need to copy the current source code to a temporary bas file
   ' and compile that bas file instead. The resulting bas file will be deleted after the compiling
   ' has completed and exe will be deleted when the exe is terminated. 
   if wID = IDM_QUICKRUN THEN
      if pDocMain->IsNewFlag then 
         gCompile.MainFolder = AfxGetExePathName
      else
         gCompile.MainFolder = AfxStrPathname("PATH", pDocMain->DiskFilename)
      end if
      gCompile.MainFilename = GetTemporaryFilename(gCompile.MainFolder, "bas")
      gCompile.MainName     = AfxStrPathname("NAME", gCompile.MainFilename)
      
      ' If this is a FORM document then we need to create the string of controls and properties that gets
      ' saved at the start of the file before any of the code is displayed.
      pDocMain->bRegenerateCode = true
      GenerateFormMetaData(pDocMain)
      GenerateFormCode(pDocMain)      ' Also, regenerate visual designer code

      ' Copy the contents of the active document to the temporary source file.
      ' Save text buffer to disk by directly accessing buffer rather
      ' saving it to an intermediary string variable first.
      Dim As ZString Ptr psz = Cast( ZString Ptr, SciExec(pDocMain->hWindow(0), SCI_GETCHARACTERPOINTER, 0, 0) )
      dim as long sciCodePage = SciExec(pDocMain->hWindow(0), SCI_GETCODEPAGE, 0, 0)   ' 0 or SC_CP_UTF8 
      
      dim pStream as CFileStream
      if pStream.Open(gCompile.MainFilename, STGM_CREATE or STGM_WRITE) = S_OK then
         dim as string st
         select CASE pDocMain->FileEncoding
            case FILE_ENCODING_ANSI
               if sciCodePage = 0 THEN    
                  pStream.Write psz, len(*psz)     ' no conversion necessary
               else
                  ' need to convert
                  st = Utf8ToAscii(*psz)
                  pStream.Write strptr(st), len(st) 
               end if    

            case FILE_ENCODING_UTF8_BOM
               ' Output the BOM first
               st = chr(&HEF, &HBB, &HBF)
               pStream.Write strptr(st), len(st)
               if sciCodePage = SC_CP_UTF8 THEN    
                  ' no conversion necessary
                  pStream.Write psz, len(*psz)     ' no conversion necessary
               else
                  ' need to convert
                  st = AnsiToUtf8(*psz)
                  pStream.Write strptr(st), len(st)
               end if    

            case FILE_ENCODING_UTF16_BOM
               ' Output the BOM first
               st = chr(&HFF, &HFE)
               pStream.Write strptr(st), len(st)
               if sciCodePage = SC_CP_UTF8 THEN    
                  pStream.Write pDocMain->wszFormMetaData.m_pBuffer, pDocMain->wszFormMetaData.m_BufferLen
                  pStream.Write pDocMain->wszFormCodeGen.m_pBuffer, pDocMain->wszFormCodeGen.m_BufferLen
                  ' convert utf8 to utf16
                  st = Utf8ToUnicode(*psz)   ' use a regular string b/c it is a byte buffer
                  pStream.Write strptr(st), len(st)
               else
                  pStream.Write pDocMain->wszFormMetaData.m_pBuffer, pDocMain->wszFormMetaData.m_BufferLen
                  pStream.Write pDocMain->wszFormCodeGen.m_pBuffer, pDocMain->wszFormCodeGen.m_BufferLen
                  ' need to convert ansi to unicode
                  wst = WStr(*psz)
                  pStream.Write wst.m_pBuffer, wst.m_BufferLen
               end if    
               
         END SELECT

      end if
      pStream.Close

   else
      gCompile.MainFilename = pDocMain->DiskFilename
      gCompile.MainName     = AfxStrPathname("NAME", pDocMain->DiskFilename)
      gCompile.MainFolder   = AfxStrPathname("PATH", pDocMain->DiskFilename)
   END IF
   
   gCompile.ResourceFile = WStr("")    ' default that there is no resource
   gCompile.CompileFlags = gConfig.Builds(idxBuild).wszOptions

   ' If the path to the WinFBX (Afx) library is valid then add that path
   ' as a -i switch to the compiler. This option is no longer presented in the
   ' Environment Options dialog but Jose Roca still uses it to assist in 
   ' more easily being able to maintain his library code.
   if AfxIsFolder(gConfig.WinFBXPath) THEN
      gCompile.CompileFlags = gCompile.CompileFlags + " -i " + chr(34) + gConfig.WinFBXPath + chr(34)
   end if
   
   If gApp.IsProjectActive Then
      Dim pDocResource As clsDocument Ptr = gApp.GetResourceDocumentPtr()
      if pDocResource then gCompile.ResourceFile = pDocResource->DiskFilename

      gCompile.ObjFolder = gCompile.MainFolder & wstr(".wfbe\")
      gCompile.ObjFolderShort = wstr(".\.wfbe\")

      If fCompile32 Then 
         gCompile.ObjID = wstr("32.o")
         gCompile.CompileFlags = gCompile.CompileFlags + " " + gApp.ProjectOther32           
      End If
      If fCompile64 Then 
         gCompile.ObjID = wstr("64.o")
         gCompile.CompileFlags = gCompile.CompileFlags + " " + gApp.ProjectOther64           
      End If   
      
      ' Make sure the folders exist
      SHCreateDirectory( 0, gCompile.ObjFolder )

   Else
      ' No active project
      gCompile.CompileFlags = gCompile.CompileFlags + " " + gConfig.CompilerSwitches   
   End If


   ' Search main source code for any user embedded compile directives. These will override
   ' anything that was set at the default or project level.
   redim directives(any) as COMPILE_DIRECTIVES
   pDocMain->CompileDirectives(directives())
   
   for i as long = lbound(directives) to ubound(directives)
      select case directives(i).DirectiveFlag 
         case IDM_CONSOLE
            n = instr(gCompile.CompileFlags, " -S GUI ")
            if n THEN gCompile.CompileFlags = AfxStrDelete(gCompile.CompileFlags, n, 8)
            gCompile.CompileFlags = gCompile.CompileFlags + " " + wstr(" -s console ")
         case IDM_GUI
            n = instr(gCompile.CompileFlags, " -S CONSOLE ")
            if n THEN gCompile.CompileFlags = AfxStrDelete(gCompile.CompileFlags, n, 12)
            gCompile.CompileFlags = gCompile.CompileFlags + " " + wstr(" -s gui ")
         case IDM_RESOURCE
            gCompile.ResourceFile = directives(i).DirectiveText
      END select
   next


   ' Need to determine the output filename in order to search for any existing
   ' running process.
   gCompile.OutputFilename = gCompile.MainFolder & gCompile.MainName & wstr(".exe")
   If Instr(" " & Ucase(gCompile.CompileFlags), wstr(" -DLL")) Then
      gCompile.OutputFilename = gCompile.MainFolder & gCompile.MainName & wstr(".dll")
   End If
   If Instr(" " & Ucase(gCompile.CompileFlags), wstr(" -DYLIB")) Then
      gCompile.OutputFilename = gCompile.MainFolder & gCompile.MainName & wstr(".dll")
   End If
   If Instr(" " & Ucase(gCompile.CompileFlags), wstr(" -LIB")) Then
      gCompile.OutputFilename = gCompile.MainFolder & wstr("lib") & gCompile.MainName & wstr(".a")
   End If


   ' Need to test if the resulting application to be compiled is actually running
   ' in memory. This would cause the compile to fail.
   If IsProcessRunning(@gCompile.OutputFilename) Then
      MessageBox( HWND_FRMMAIN, L(200,"Program running") & "...", L(201,"Error"), _
                        MB_OK Or MB_ICONINFORMATION Or MB_DEFBUTTON1 )
      Function = False: Exit Function
   End If
      

   gApp.IsCompiling = true    

   dim as HCURSOR hCurSave = GetCursor()
   SetCursor( LoadCursor(0, IDC_WAIT) )
   pDoc = gTTabCtl.GetActiveDocumentPtr()
   if pDoc THEN SciExec( pDoc->hWindow(0), SCI_SETCURSOR, SC_CURSORWAIT, 0 )


   ' Regenerate any visual designer code for Forms. Doing a pDoc->Save will regenerate
   ' the code for a document with pDoc->IsDesigner.
   pDoc = gApp.pDocList
   do until pDoc = 0
      if pDoc->IsDesigner then pDoc->SaveFile
      pDoc = pDoc->pDocNext
   loop
   
   
   ' Run any precompile tools
   for y as long = LBound(gConfig.Tools) To UBound(gConfig.Tools)
      if gConfig.Tools(y).Action = USERTOOL_ACTION_PRECOMPILE then
         frmUserTools_ExecuteUserTool(y)
      end if   
   NEXT

   ' Change to the output code directory
   ChDir gCompile.MainFolder
   
   
   ' Indicate in the StatusBar that we have started to compile
   SetCompileStatusBarMessage( L(203, "Compiling") & "...", 0 )
   
   gCompile.StartTime = Timer
   

   ' If this is a project then we need to compile all of the modules first
   ' in order to create the necessary *.o object files
   If gApp.IsProjectActive Then
      dim as long nFileCount   ' used to calculate when to fire an AfxDoEvents
      pDoc = gApp.pDocList
      do until pDoc = 0
         If pDoc->ProjectFileType = FILETYPE_MODULE Then
            ' Get the base name of the file for constructing the object filename
            wszTemp = AfxStrPathname("NAME", pDoc->DiskFilename)
            wszTemp = gCompile.ObjFolderShort & wszTemp & gCompile.ObjID
            
            ' Compare the source code file datetime to the object. If the source code
            ' date time is greater then we need to recompile it, otherwise we will simply
            ' link to the existing object file.
            Dim As FILETIME ft1 = AfxGetFileLastWriteTime(pDoc->DiskFilename) ' source file
            Dim As FILETIME ft2 = AfxGetFileLastWriteTime(wszTemp)  ' object file

            If (AfxFileTimeToVariantTime(ft1) > AfxFileTimeToVariantTime(ft2)) or _
               (wID = IDM_REBUILDALL)  Then

               nFileCount = nFileCount + 1
               if (nFileCount mod 5) = 0 then AfxDoEvents
               
               wst = L(203, "Compiling") + "... " + AfxStrPathname("NAMEX", pDoc->DiskFilename) 
               SetCompileStatusBarMessage( wst, 0 )
 
               wszCommand = gCompile.CompilerPath
               wszParams = " -c -b " + wDQ + pDoc->DiskFilename + wDQ + _
                           " -v -o " + wDQ + wszTemp + wDQ 
               RedirConsoleToFile( wszCommand, wszParams, sConsoleText )
               gCompile.wszFullCommandLine = wszCommand & " " & wszParams
               gCompile.wszFullLogFile = sConsoleText
               
               ' Need to check to see if any errors occurred while compiling this
               ' object file. If yes, then stop processing object files and display
               ' the error.
               if ParseLogForError(sConsoleText, false, wID, fCompile64, true) then   ' error detected
                  gApp.IsCompiling = false
                  SetCursor( hCurSave )
                  ResetScintillaCursors()
                  return true
               end if
               sConsoleText = ""
            End If
            
         End If
         pDoc = pDoc->pDocNext
      loop
   End If

   AfxDoEvents
   
   
   ' If this is a visual designer project then output specific equates that dictate
   ' which controls get compiled into the file source.
   dim as CWSTR wszVDEquates, wszCtrlEquate
   dim as boolean bIsVisual = false
   pDoc = gApp.pDocList
   do until pDoc = 0
      if pDoc->IsDesigner then 
         bIsVisual = true
         ' Search all of the controls and set the appropriate equate
         dim pCtrl as clsControl ptr
         for i as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
            pCtrl = pDoc->Controls.ItemAt(i)
            wszCtrlEquate = "#define CODEGEN_" & ucase(GetToolBoxName( pCtrl->ControlType ))
            if instr(wszVDEquates, wszCtrlEquate) = 0 then
               wszVDEquates = wszVDEquates & wszCtrlEquate & vbcrlf
               ' If adding a RichEdit control reference then we also need to add
               ' TextBox reference because a RichEdit extends a TextBox.
               if pCtrl->ControlType = CTRL_RICHEDIT then
                  wszCtrlEquate = "#define CODEGEN_TEXTBOX"
                  if instr(wszVDEquates, wszCtrlEquate) = 0 then
                     wszVDEquates = wszVDEquates & wszCtrlEquate & vbcrlf
                  end if
               end if
            end if
         NEXT

         if pDoc->MainMenuExists then
            wszCtrlEquate = "#define CODEGEN_MAINMENU"
            if instr(wszVDEquates, wszCtrlEquate) = 0 then
               wszVDEquates = wszVDEquates & wszCtrlEquate & vbcrlf
            end if
         end if
         if pDoc->ToolBarExists then
            wszCtrlEquate = "#define CODEGEN_TOOLBAR"
            if instr(wszVDEquates, wszCtrlEquate) = 0 then
               wszVDEquates = wszVDEquates & wszCtrlEquate & vbcrlf
            end if
         end if
         if pDoc->StatusBarExists then
            wszCtrlEquate = "#define CODEGEN_STATUSBAR"
            if instr(wszVDEquates, wszCtrlEquate) = 0 then
               wszVDEquates = wszVDEquates & wszCtrlEquate & vbcrlf
            end if
         end if
      end if
      pDoc = pDoc->pDocNext
   loop

   if bIsVisual then
      dim wszVDMain as CWSTR = gCompile.MainFolder & "WinFBE_VD_MAIN.bas"
      dim pStream as CTextStream
      if pStream.Create(wszVDMain, true, true) = S_OK then 
         dim as CWSTR wszText = _
            wszVDEquates & _
            "#Include Once " & wDQ & gCompile.MainFilename & wDQ
         pStream.WriteLine wszText
         pStream.Close
         gCompile.MainFilename = wszVDMain
      end if
   end if
   

   ' Compile the Main file. If this is a project then we also need to link
   ' in all of the *.o object files.
   ' Per fxm post https://www.freebasic.net/forum/viewtopic.php?f=8&p=266126#p266126
   ' Do not specify file extension with -m parameter, and also include the file
   ' a second time in the list of files but include the extension at that time.
   'dim as CWSTR wszMfile = wDQ + gCompile.MainFilename + wDQ
   'dim as CWSTR wszMfile = wDQ + AfxStrPathName("PATH", gCompile.MainFilename) + AfxStrPathName("NAME", gCompile.MainFilename) + wDQ
   'wszParams = "-m " + wszMfile + " " + wDQ + gCompile.MainFilename + wDQ 
   wszParams = "-m " + wDQ + gCompile.MainFilename + wDQ 


   ' Determine if we need to generate a temporary Resource file to hold any
   ' necessary for visual designer forms/projects where image references need to be outputted.
   if CreateTempResourceFile() then
      wszParams = wszParams + " " + wDQ + gCompile.TempResourceFile + wDQ
   else
      if len(gCompile.ResourceFile) then  
         wszParams = wszParams + " " + wDQ + gCompile.ResourceFile + wDQ
      end if   
   end if   

   
   ' Ensure verbose compiler message output
   wszParams = wszParams + " -v " + gCompile.CompileFlags 
   
   if instr(wszParams, " -x ") = 0 then
      wszParams = wszParams + " -x " + wDQ + gCompile.OutputFilename + wDQ
   end if
   
   If gApp.IsProjectActive Then
      pDoc = gApp.pDocList
      do until pDoc = 0
         If pDoc->ProjectFileType = FILETYPE_MODULE Then
            ' Get the base name of the file for constructing the object filename
            wszTemp = AfxStrPathname("NAME", pDoc->DiskFilename)
            wszParams = wszParams + " " + gCompile.ObjFolderShort + wszTemp + gCompile.ObjID
         End If
         pDoc = pDoc->pDocNext
      loop
   End If

   ' Attempt to delete any existing output file so it will not exist
   ' should the compile fail.
   if AfxFileExists( gCompile.OutputFilename ) then
      AfxDeleteFile( gCompile.OutputFilename )
   end if

   wst = L(203, "Compiling") + "... " + AfxStrPathname("NAMEX", gCompile.MainFilename) 
   SetCompileStatusBarMessage( wst, 0 )
   SetCursor( LoadCursor(0, IDC_WAIT) )


   sConsoleText = ""
   wszCommand = gCompile.CompilerPath
   RedirConsoleToFile( wszCommand, wszParams, sConsoleText )
   gCompile.wszFullCommandLine = wszCommand & " " & wszParams
   gCompile.wszFullLogFile = sConsoleText
               

   ' Attempt to fix the random problem whereby the spinning mouse wait cursor is not reset
   ' to an arrow. It seems to happen in random cases whereby an error is thrown and a new 
   ' document is loaded to position to the error line. Very hard to reproduce. We will manually
   ' reset all Scintilla cursors here for all open documents. Maybe this will help.
   SetCursor( hCurSave )
   ResetScintillaCursors()
   
   gCompile.EndTime = Timer
   gCompile.SystemTime = AfxSystemSystemTime 


   ' If this was a QuickRun then delete the temporary code file that was created.
   if wID = IDM_QUICKRUN THEN 
      AfxDeleteFile( gCompile.MainFilename )
   end if    
   
   ' Add any temp resource file to the list of temp files to later be deleted.
   gApp.AddQuickRunEXE(gCompile.TempResourceFile)
      
   if ParseLogForError(sConsoleText, true, wID, fCompile64, false) = true then 
      ' Error was found. Editor has now been set to the error position. Nothing
      ' more we can do now but exit out. 
      ' Make sure gApp.IsCompiling = false before returning
   else
      
      ' Run any postcompile tools
      for y as long = LBound(gConfig.Tools) To UBound(gConfig.Tools)
         if gConfig.Tools(y).Action = USERTOOL_ACTION_POSTCOMPILE then
            frmUserTools_ExecuteUserTool(y)
         end if   
      NEXT

      If gCompile.RunAfterCompile Then 
         RunEXE( gCompile.OutputFilename, gApp.ProjectCommandLine)
      End If
   
   end if
   
   gApp.IsCompiling = false

   Function = True   ' successful
End Function


