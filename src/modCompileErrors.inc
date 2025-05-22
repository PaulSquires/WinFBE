''    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
''    Copyright (C) 2016-2025 Paul Squires, PlanetSquires Software
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
''   modCompileErrors.inc
''

#include once "modCompile.bi"
#include once "modCompileErrors.bi"

''
''  Set the statusbar text and icon for a good/bad compile
''
function SetCompileStatusBarMessage( _
            byref wszText as wstring, _
            byval hIconCompile as long _
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
function CreateTempResourceFile() as boolean
   dim as CWSTR wszResourceFile = gCompile.ResourceFile
   
   ' Collect all of the IMAGE_NAME used by controls/properties in the file/project
   ' and generate a string to output to the resource file.
   dim pDoc as clsDocument ptr
   dim as CWSTR wszInternalText, wszExistingText, wszImage
   
   pDoc = gApp.pDocList
   do until pDoc = 0
      for i as long = lbound(pDoc->AllImages) to ubound(pDoc->AllImages)
         ' IMAGE_ARROWLEFT    RCDATA "IMAGES\\LEFTARROW.PNG"
         wszImage = pDoc->AllImages(i).wszImageName & wspace(6) & _
                    pDoc->AllImages(i).wszFormat    & wspace(2) & _
                    wchr(34) & pDoc->AllImages(i).wszFileName & wchr(34) & _
                    vbcrlf
         
         ' Only add this Image if code has not already been generated for it, otherwise
         ' there will be a resource compile error.
         if instr(wszInternalText, wszImage) = 0 then
            wszInternalText = wszInternalText & wszImage
         end if
            
         ' Check to ensure that the path to the image is valid otherwise the
         ' resource compiler will fail.
         if AfxFileExists( pDoc->AllImages(i).wszFileName ) = false then
            gCompile.bInvalidImagePath = true
         end if
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
            ' UTF8 BOM encoded  (convert from utf8)
            wszExistingText.Utf8 = mid(txtBuffer, 4)  ' bypass the BOM
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
function ResetScintillaCursors() as Long
   Dim As Long nCount = gTTabCtl.GetItemCount 

   For i as long = 0 To nCount - 1
      if gTTabCtl.IsSafeIndex(i) = false then continue for
      dim as clsDocument ptr pDoc = gTTabCtl.tabs(i).pDoc
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
function RedirConsoleToFile( _
            byref wszExe AS wstring, _
            byref wszCmdLine AS wstring, _
            byref sConsoleText AS string _
            ) as long

   ' From the MinGW "Getting Started" guide:
   '    MinGW may have problems with paths containing spaces, and if not, usually 
   '    other programs used with MinGW will experience problems with such paths. 
   '    Thus, we strongly recommend that you do not install MinGW in any location 
   '    with spaces in the path name reference; i.e. you should avoid installing 
   '    into any subdirectory of "Program Files" or "My Documents", or the like.
   if instr(wszExe, " ") then
      AfxMsg( "Compile failed (CreateChildProcess)." & vbcrlf & _
              "Please install WinFBE into a folder without spaces." )
      exit function
   end if

   ' Create the child process and read from its standard output
   dim ProcessInfo AS PROCESS_INFORMATION
   dim StartupInf  AS STARTUPINFO

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
   IF CreatePipe( @hChildStdOutRead, @hChildStdOutWrite, @SecurityAttribute, BYVAL 0 ) = FALSE THEN
      '? "error creating pipe"
   ELSE
      StartupInf.cb         = SIZEOF(STARTUPINFO)
      GetStartupInfo(@StartupInf )  
      StartupInf.hStdError  = hChildStdOutWrite   
      StartupInf.hStdOutput = hChildStdOutWrite  
      StartupInf.dwFlags    = STARTF_USESTDHANDLES or STARTF_USESHOWWINDOW
      StartupInf.wShowWindow = SW_HIDE
      
      ' This call to CreateProcessW will succeed in all cases except for the one
      ' where there are spaces in wszFullExe because once I wrap the string in double
      ' quotes then the call to CreateProcessW fails.

      dim errCode as long 
      errCode = CreateProcessW( _
                  wszExe,  _          ' Create the child process.
                  wszCmdLine, _       ' Command line
                  BYVAL 0,     _      ' Process security attributes
                  BYVAL 0,     _      ' Primary thread security attributes
                  TRUE,        _      ' Handles are inherited
                  byval 0,     _      ' Creation flags  
                  BYVAL 0,     _      ' Use parent's environment
                  BYVAL 0,     _      ' Use parent's current directory
                  @StartupInf,  _     ' STARTUPINFO pointer
                  @ProcessInfo) 
      
      if errCode <> 0 then            

         CloseHandle( hChildStdOutWrite ) ' To avoid ReadFile hanging at the end.
         DO
            IF ReadFile( hChildStdOutRead, BYVAL VARPTR(sBuffer), 4096, @dwReadBytes, BYVAL 0 ) = FALSE THEN 
               '? "ERROR ReadFile', "GetLastError: "; GetLastError
               EXIT DO
            end if  
            sConsoleText = sConsoleText & LEFT(sBuffer, dwReadBytes)
            '(at this stage we could SAVE the buffer that we've read from the child to a file, but 
            ' thats an extra ReadFile+WriteFile compared to the direct-to-hFile method)
         LOOP
            
      else
         '? "CreateChildProcess failed"      
         if instr( wszExe, " " ) then
            AfxMsg( "CreateChildProcess failed. Install WinFBE into folder without spaces." )
         end if    
         if hChildStdOutRead  then CloseHandle( hChildStdOutRead )
         if hChildStdOutWrite then CloseHandle( hChildStdOutWrite )
         exit function
      END IF
         
   END IF

   ' Close handles to the child process and its primary thread.
   ' Some applications might keep these handles to monitor the status 
   ' of the child process, for example.
   if ProcessInfo.hProcess then CloseHandle( ProcessInfo.hProcess )
   if ProcessInfo.hThread  then CloseHandle( ProcessInfo.hThread )
   if hChildStdOutRead     then CloseHandle( hChildStdOutRead )

   function = 0
END function


''
''
function RunEXE( _
            ByRef wszFileExe As CWSTR, _
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
         pStream.WriteLine Left(wszPath, 2)     ' ensure on correct drive
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
function SetDocumentErrorPosition( _
            ByVal hLV As HWnd, _
            Byval wID as long _
            ) As Long

   Dim wszErrorLine As WString * MAX_PATH
   Dim wszErrorFile As WString * MAX_PATH
   
   Dim As Long nCurSel = ListView_GetSelection(hLV)
   If nCurSel < 0 Then Return 0 
   
   Dim pDoc As clsDocument Ptr 

   ' Get the line number and filename of the selected item
   FF_ListView_GetItemText(hLV, nCurSel, 1, @wszErrorLine, MAX_PATH)
   FF_ListView_GetItemText(hLV, nCurSel, 2, @wszErrorFile, MAX_PATH)
   dim as long nLineNum = Val(wszErrorLine) - 1   ' because visually the editor is one based line numbers

   ' If we are doing a QuickRun then it must be on the currently loaded
   ' and active file in the editor, therefore simply use that pDoc.
   if wID = IDM_QUICKRUN THEN
      pDoc = gTTabCtl.GetActiveDocumentPtr()
      if pDoc then wszErrorFile = pDoc->DiskFilename 
   END IF
   
   pDoc = OpenSelectedDocument( wszErrorFile, "", nLineNum )
   
   function = 0
End Function


''
''
function SetLogFileTextbox() as long
   
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
function ParseLogForError( _
            byref wsLogSt as CWSTR, _
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

   hLV = GetDlgItem(HWND_FRMOUTPUT, IDC_FRMOUTPUT_LVRESULTS)
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
      If Left(wst_ucase, 16) = "THIS VERSION OF " Then
         fGoodCompile = false
         FF_ListView_InsertItem( hLV, NumWarnings, 0, "" )
         FF_ListView_InsertItem( hLV, NumWarnings, 1, "0" )
         FF_ListView_InsertItem( hLV, NumWarnings, 2, "" ) 
         FF_ListView_InsertItem( hLV, NumWarnings, 3, "compiling FAILED: Windows not 64-bit (refer to log file)" ) 
         wszOutputMsg = wsLogSt: exit For
         exit for
      End If   

      ' Save the FB version and copyright information. 
      If Left(wst_ucase, 19) = "FREEBASIC COMPILER " Then
         fGoodCompile = True
         wszOutputMsg = wszOutputMsg & wst & vbCrLf: Continue For
      End If        
      If Left(wst_ucase, 13) = "COPYRIGHT (C)" Then
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
      If (Left(wst_ucase, 6) = "ERROR!") or _
         (left(wst_ucase, 34) = "VERY SORRY, BUT GORC HAD A PROBLEM") Then
         FF_ListView_InsertItem( hLV, nListViewLine, 0, "" )
         FF_ListView_InsertItem( hLV, nListViewLine, 1, "0" )
         FF_ListView_InsertItem( hLV, nListViewLine, 2, "" ) 
         FF_ListView_InsertItem( hLV, nListViewLine, 3, "compiling resource FAILED: Error Code 1 (refer to log file)" ) 
         NumErrors = NumErrors + 1
      elseIf (Left(wst_ucase, 5) = "ERROR") Then
         FF_ListView_InsertItem( hLV, nListViewLine, 0, "" )
         FF_ListView_InsertItem( hLV, nListViewLine, 1, "0" )
         FF_ListView_InsertItem( hLV, nListViewLine, 2, "" ) 
         FF_ListView_InsertItem( hLV, nListViewLine, 3, "compiling FAILED: Error Code 1 (refer to log file)" ) 
         NumErrors = NumErrors + 1
      elseIf (Left(wst_ucase, 19) = "COMPILING C FAILED:") Then
         FF_ListView_InsertItem( hLV, nListViewLine, 0, "" )
         FF_ListView_InsertItem( hLV, nListViewLine, 1, "0" )
         FF_ListView_InsertItem( hLV, nListViewLine, 2, "" ) 
         FF_ListView_InsertItem( hLV, nListViewLine, 3, "compiling C FAILED: Error Code 1 (refer to log file)" ) 
         NumErrors = NumErrors + 1
      end if
      
      If Instr(wst_ucase, "LINKING FAILED:") Then 
         FF_ListView_InsertItem( hLV, nListViewLine, 0, "" )
         FF_ListView_InsertItem( hLV, nListViewLine, 1, "0" ) 
         FF_ListView_InsertItem( hLV, nListViewLine, 2, "" ) 
         FF_ListView_InsertItem( hLV, nListViewLine, 3, wst & " (refer to log file)" ) 
         NumErrors = NumErrors + 1
      End If
 
 
      ' Check for the "linking: " line because that line identifies the actual output filename being created
      If Left(wst_ucase, 9) = "LINKING: " Then 
         ' -o "WinFBE.exe" 
         i = Instr(wst, " -o ")
         If i Then
            wszTemp = Mid(wst, i+4)
            i = Instr(wszTemp, wDQ & " ")
            If i Then gCompile.OutputFilename = AfxStrRemove(Left(wszTemp, i), wDQ)
         End If      
      End If
      
      ' Determine if we are reading any lines in the log file that appear after the "compiling: " line
      ' but before the "assembling: ". Anything there is data output by #print preprocessor statements.
      If Left(wst_ucase, 11) = "COMPILING: " Then 
         bReadingCompileOutput = true
         continue for
      elseIf Left(wst_ucase, 13) = "COMPILING C: " Then 
         bReadingCompileOutput = true
         continue for
      elseIf Left(wst_ucase, 12) = "ASSEMBLING: " Then 
         bReadingCompileOutput = false
      elseIf Left(wst_ucase, 14) = "COMPILING RC: " Then 
         bReadingCompileOutput = false
      elseIf Left(wst_ucase, 9) = "LINKING: " Then 
         bReadingCompileOutput = false
      end if

      if bReadingCompileOutput then

         if len(wst) = 0 then continue for

         ' Check for any compiler warnings
         Error_Start = 0
         IsWarning = Instr(wst_ucase, ") WARNING ")
         If IsWarning Then Error_Start = IsWarning
         IsError   = Instr(wst_ucase, ") ERROR ") 
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
            Parenthesis_Start = InStrRev(wst, "(", Parenthesis_End)
            
            If (Parenthesis_Start < Parenthesis_End) Andalso _
               (Parenthesis_End <= Error_Start) then
       
               wst1 = left(wst, Parenthesis_Start-1)   ' filename
               wst2 = Mid( wst, Parenthesis_Start + 1, Parenthesis_End - Parenthesis_Start - 1)  ' line# 
               wst3 = Mid( wst, Error_Start + 1)  ' error message
     

               ' If the filename is "WinFBE_VD_MAIN.bas" then we need to load the corresponding 
               ' pDocMain instead and adjust for the error line offset.
               if ucase(AfxStrPathname( "NAMEX", wst1 )) = "WINFBE_VD_MAIN.BAS" then
                  wst1 = gCompile.pDocMain->DiskFilename
                  wst2 = str( val(wst2) - gCompile.pDocMainOffset)
               end if
              
               FF_ListView_InsertItem( hLV, nListViewLine, 0, "" )
               FF_ListView_InsertItem( hLV, nListViewLine, 1, wst2 )  ' line#
               FF_ListView_InsertItem( hLV, nListViewLine, 2, wst1 )  ' filename
               FF_ListView_InsertItem( hLV, nListViewLine, 3, wst3 )  ' error message
                
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
                     
            FF_ListView_InsertItem( hLV, nListViewLine, 0, "" )
            FF_ListView_InsertItem( hLV, nListViewLine, 1, "" )  ' line#
            FF_ListView_InsertItem( hLV, nListViewLine, 2, "" )  ' filename
            FF_ListView_InsertItem( hLV, nListViewLine, 3, wst )  ' message
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


   dim as CWSTR wszCompileMsg

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
   
      wszCompileMsg = L(193, "Errors")   & " " & NumErrors   & "  " & _
                      L(192, "Warnings") & " " & NumWarnings & _
                      "  [" & _
                        AfxLocalDateStr( "yyyy-MM-dd" ) & "  " & _
                        AfxLocalTimeStr( "hh:mm:ss" ) & "]"
      gCompile.wszOutputMsg = L(229,"Failed Compile")  & " (" & wszCompileMsg & ")" & vbcrlf
      SetCompileStatusBarMessage( wszCompileMsg, ghIconBad )
      MessageBeep(MB_ICONWARNING)
      
      ' Position the Compiler Log to the first error/warning
      ListView_SelectItem( hLV, 0 )
      
      ' If fGoodCompile is False then something unusual occurred so better show the log file by default
      If fGoodCompile = False Then
         frmOutput_PositionWindows
      End If
       
      ' If the Search Results, TODO, Notes tab is active then reposition to 
      ' the first tab for error listview
      select case gOutputTabsCurSel 
         case 2, 3, 4
         gOutputTabsCurSel = 0
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
         if gOutputTabsCurSel = 0 THEN
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
         wszCompileMsg = L(193, "Errors") & " " & NumErrors & "  " & _
                         L(192, "Warnings") & " " & NumWarnings & _
                         "  [ " & AfxStrPathname("NAMEX", gCompile.OutputFilename) & "  " & _
                         rtrim(AfxStrFormatKBSize(nFileSize)) & " (" & _
                         nFileSize & " " & L(199,"bytes") & ") ]"
         SetCompileStatusBarMessage( wszCompileMsg, ghIconGood )
         ' 2018-12-13: No longer offer the option to show the popup successful dialog.
         ' Simply sound a success
         if gConfig.DisableCompileBeep = false then
            MessageBeep(MB_OK)
         end if
         
         gCompile.wszOutputMsg = _   ' this gets output to the Output window
         L(228,"Successful Compile")  & " (" & _
               L(193,"Errors")        & " " & NumErrors & " " & _
               L(192,"Warnings")      & " " & NumWarnings & ")" & vbcrlf & vbcrlf & _
         L(194,"Primary Source:")     & " " & gCompile.MainFilename & vbCrLf & _
         L(195,"Target Compilation:") & " " & gCompile.OutputFilename & " (" & _
               rtrim(AfxStrFormatKBSize(nFileSize)) & ", " & nFileSize & " " & L(199,"bytes") & ")" & vbcrlf & _
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


