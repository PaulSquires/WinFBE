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
''  Handle information related to Compile process 
''
Type COMPILE_TYPE
   MainFilename    As WString * MAX_PATH   ' main source file (full path and file.ext)
   MainName        As WString * MAX_PATH   ' main source file (Name only, no extension)
   MainFolder      As WString * MAX_PATH   ' main source folder 
   ResourceFile    As WString * MAX_PATH   ' full path and file.ext to resource file (if applicable) 
   OutputFilename  As WString * MAX_PATH   ' resulting exe/dll/lib name 
   CompilerPath    As WString * MAX_PATH   ' full path and file.ext to fbc.exe
   ObjFolder       As WString * MAX_PATH   ' *.o for all modules (set depending on 32/64 bit) (full path)
   ObjFolderShort  As WString * MAX_PATH   ' ".\.wfbe\"
   ObjID           As WString * MAX_PATH   ' "32" or "64" appended to object name
   CompileMode     As WString * MAX_PATH 
   CompileFlags    As WString * MAX_PATH
   BatchFile       As WString * MAX_PATH
   LogFile         As WString * MAX_PATH  
   RunAfterCompile As BOOLEAN
   StartTime       As Double
   EndTime         As Double
End Type
Dim Shared gCompile As COMPILE_TYPE

''
''  Set the statusbar text and icon for a good/bad compile
'------------------------------------------------------------------------------------------------------------------------
function SetCompileStatusBarMessage( byref wszText as const wstring, byval hIconCompile as HICON ) as LRESULT
   dim as hwnd hStatusBar = GetDlgItem( HWND_FRMMAIN, IDC_FRMMAIN_STATUSBAR )
   StatusBar_SetIcon(hStatusBar, 5, hIconCompile)
   StatusBar_SetText(hStatusBar, 5, wszText)
   function = 0
end function


''
''  Main module that handles the entire compile process
''
'------------------------------------------------------------------------------------------------------------------------
Function code_Compile( ByVal wID As Long ) As BOOLEAN

   Dim ShExecInfo As SHELLEXECUTEINFOW  
    
   Dim As Long Parenthesis_Start, Parenthesis_End, Error_Start       
   Dim As Long f, i, NumLines, NextLine, idx, r, nCount
   Dim As Long NumWarnings, NumErrors, IsError, IsWarning
   Dim As Long nFirstErrorLine = -1
   Dim pDoc As clsDocument Ptr
   
   Dim wszTemp      As WString * MAX_PATH 
   Dim wszFileExe   As WString * MAX_PATH
   Dim wszParam     As WString * MAX_PATH  
   
   Dim As CWSTR wexecmd, wDQ, wsLogSt, wst, wst_ucase, wszOutputMsg
   Dim As CWSTR wst1, wst2, wst3        
   Dim As HWnd hLV
         
   Dim As BOOLEAN fCompile32, fCompile64, fGoodCompile
   
   idx = gApp.GetActiveProjectIndex()
   
   'sDQ = Chr(34)  ' double quotes
   wDQ = wchr(34) ' unicode double quotes
   
   ' Can only continue to compile if the primary source code file
   ' is not dirty and needs to be saved.
   If gConfig.CompileAutosave Then
      If OnCommand_FileSaveAll(HWND_FRMMAIN) Then Return True  ' user cancelled save
   End If
   
   ' Perform some pre-compile checks to see if we should continue.
   gConfig.FBWINCompiler32 = ProcessFromCurdrive(gConfig.FBWINCompiler32)
   gConfig.FBWINCompiler64 = ProcessFromCurdrive(gConfig.FBWINCompiler64)
   Select Case Ucase(gApp.Projects(idx).ProjectCompiler)
      Case "FBC 32BIT"
         fCompile32 = True: fCompile64 = False
         gCompile.CompilerPath = gConfig.FBWINcompiler32
      Case "FBC 64BIT"
         gCompile.CompilerPath = gConfig.FBWINcompiler64
         fCompile32 = False: fCompile64 = True
   End Select
   
   
   ' Check to see if the compiler exists
   If AfxFileExists(gCompile.CompilerPath) = 0 Then 
      MessageBox( HWND_FRMMAIN, L(202,"Invalid defined compiler path."), L(201,"Error"), _
                        MB_OK Or MB_ICONINFORMATION Or MB_DEFBUTTON1 )
      Function = False: Exit Function
   End If
   
   
   ' Set some compile flags depending on the type of compile requested.
   Select Case wID
      Case IDM_BUILDEXECUTE:  gCompile.RunAfterCompile = True
      Case IDM_COMPILE:       gCompile.RunAfterCompile = False
   End Select
   
   
   Dim pDocMain As clsDocument Ptr 
   If gApp.IsProjectActive Then
      pDocMain = gApp.Projects(idx).GetMainDocumentPtr()
      If pDocMain = 0 Then
         MessageBox( HWND_FRMMAIN, L(208,"No Main file specified for the project."), L(201,"Error"), _
                           MB_OK Or MB_ICONINFORMATION Or MB_DEFBUTTON1 )
         Function = False: Exit Function
      End If
   Else
      pDocMain = gTTabCtl.GetActiveDocumentPtr()
   End If
   If pDocMain = 0 Then Return 0
   
   gCompile.MainFilename = pDocMain->DiskFilename
   gCompile.MainName     = AfxStrPathname("NAME", pDocMain->DiskFilename)
   gCompile.MainFolder   = AfxStrPathname("PATH", pDocMain->DiskFilename)
   gCompile.ResourceFile = WStr("")    ' default that there is no resource
   
   If gApp.IsProjectActive Then
      Dim pDocResource As clsDocument Ptr = gApp.Projects(idx).GetResourceDocumentPtr()
      if pDocResource then gCompile.ResourceFile = pDocResource->DiskFilename

      gCompile.ObjFolder = gCompile.MainFolder & wstr(".wfbe\")
      gCompile.ObjFolderShort = wstr(".\.wfbe\")

      If fCompile32 Then 
         gCompile.ObjID = wstr("32.o")
         gCompile.CompileFlags = gApp.Projects(idx).ProjectOther32           
      End If
      If fCompile64 Then 
         gCompile.ObjID = wstr("64.o")
         gCompile.CompileFlags = gApp.Projects(idx).ProjectOther64           
      End If   
      
      ' Make sure the folders exist
      SHCreateDirectory( 0, gCompile.ObjFolder )
      
      Select Case gApp.Projects(idx).ProjectType
         Case 0:   ' EXE (no special compile switch)
            gCompile.OutputFilename = gCompile.MainFolder & gCompile.MainName & wstr(".exe") 
         Case 1
            gCompile.OutputFilename = gCompile.MainFolder & gCompile.MainName & wstr(".dll") 
            gCompile.CompileFlags = gCompile.CompileFlags & wstr(" -dll ")
         Case 2
            gCompile.OutputFilename = gCompile.MainFolder & wstr("lib") & gCompile.MainName & wstr(".a") 
            gCompile.CompileFlags = gCompile.CompileFlags & wstr(" -lib ")
      End Select

      Select Case gApp.Projects(idx).ProjectErrorOption
         Case 0:   ' No error checking (no special compile switch)
         Case 1:  gCompile.CompileFlags = gCompile.CompileFlags & wstr(" -e")           
         Case 2:  gCompile.CompileFlags = gCompile.CompileFlags & wstr(" -ex")           
         Case 3:  gCompile.CompileFlags = gCompile.CompileFlags & wstr(" -exx")           
      End Select

      If gApp.Projects(idx).ProjectDebug = BST_CHECKED Then
         gCompile.CompileFlags = gCompile.CompileFlags & wstr(" -g")           
      End If
      
      If gApp.Projects(idx).ProjectThread = BST_CHECKED Then
         gCompile.CompileFlags = gCompile.CompileFlags & wstr(" -mt")           
      End If

   Else
      ' No active project
      gCompile.CompileFlags = gConfig.CompilerSwitches   
      gCompile.OutputFilename = gCompile.MainFolder & gCompile.MainName & wstr(".exe")
      If Instr(Ucase(gCompile.CompileFlags), wstr(" -DLL")) Then
         gCompile.OutputFilename = gCompile.MainFolder & gCompile.MainName & wstr(".dll")
      End If
      If Instr(Ucase(gCompile.CompileFlags), wstr(" -DYLIB")) Then
         gCompile.OutputFilename = gCompile.MainFolder & gCompile.MainName & wstr(".dll")
      End If
      If Instr(Ucase(gCompile.CompileFlags), wstr(" -LIB")) Then
         gCompile.OutputFilename = gCompile.MainFolder & wstr("lib") & gCompile.MainName & wstr(".a")
      End If
   End If


   Select Case UCase(gApp.Projects(idx).ProjectCompileMode)   ' used on the Main source file
      Case "GUI"
         ' For non-projects simply suppress the console. For projects, the console
         ' will not show unless option is selected in Project Options.
         gCompile.CompileMode = wstr(" -s gui ")   ' default do not show console
         If gApp.Projects(idx).ProjectShowConsole = BST_CHECKED Then
            gCompile.CompileMode = wstr(" ")   ' console is not suppressed
         End If   
      Case "CONSOLE": gCompile.CompileMode = wstr(" -s console ")
   End Select   
  

   ' Need to test if the resulting application to be compiled is actually running
   ' in memory. This would cause the compile to fail.
   If IsProcessRunning(@gCompile.OutputFilename) Then
      MessageBox( HWND_FRMMAIN, L(200,"Program running..."), L(201,"Error"), _
                        MB_OK Or MB_ICONINFORMATION Or MB_DEFBUTTON1 )
      Function = False: Exit Function
   End If
      
   ' Create a temporary batch that will contain the compile commands. We shell
   ' to and run this batch file in order to do the compile.
   ' Replace any embedded spaces with underscores
   wszTemp = wstr("_") + AfxStrPathname("NAME", gApp.Projects(idx).ProjectName)
   wszTemp = AfxStrReplace(wszTemp, wstr(" "), wstr("_") )
   gCompile.BatchFile = wszTemp & wstr("_compile.bat") 
   gCompile.LogFile   = wszTemp & wstr("_compile.log")
   
   ' Change to the output code directory
   ChDir gCompile.MainFolder
   
   ' Must delete any previous instance of the compiler log file in case it was not deleted. Need
   ' to do this because the compiler appends to the log file rather than overwriting.
   Kill gCompile.LogFile
   
   f = Freefile
   Open gCompile.BatchFile For Output as #f


   ' If this is a project then we need to compile all of the modules first
   ' in order to create the necessary *.o object files
   If gApp.IsProjectActive Then
      nCount = gApp.Projects(idx).GetDocumentCount 
      For i = 0 To nCount - 1
         pDoc = gApp.Projects(idx).GetDocumentPtr(i)
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
               Print #f, wDQ; gCompile.CompilerPath; wDQ; _
                         " -c "; wDQ; pDoc->DiskFilename; wDQ; _
                         " -o "; wDQ; wszTemp; wDQ;
               Print #f, " >> "; gCompile.LogFile; " 2>&1"   ' ensure stderr is redirected (appended) to log file as well
            End If
            
         End If
      Next
   End If

   
   ' Compile the Main file. If this is a project then we also need to link
   ' in all of the *.o object files.
   Print #f, wDQ; gCompile.CompilerPath; wDQ; _
            " "; wDQ; gCompile.MainFilename; wDQ; 
   If Len(gCompile.ResourceFile) then
      print #f, " "; wDQ; gCompile.ResourceFile; wDQ; 
   end if   
   print #f, " -v "; gCompile.CompileFlags; gCompile.CompileMode;
        
   If gApp.IsProjectActive Then
      nCount = gApp.Projects(idx).GetDocumentCount 
      For i = 0 To nCount - 1
         pDoc = gApp.Projects(idx).GetDocumentPtr(i)
         If pDoc->ProjectFileType = FILETYPE_MODULE Then
            ' Get the base name of the file for constructing the object filename
            wszTemp = AfxStrPathname("NAME", pDoc->DiskFilename)
            Print #f, " "; gCompile.ObjFolderShort; wszTemp; gCompile.ObjID;
         End If
      Next
   End If

   Print #f, " >> "; gCompile.LogFile; " 2>&1"   ' ensure stderr is redirected (appended) to log file as well
   Close #f
         
   wszFileExe = Environ("COMSPEC")
   'wszParam   = " /U /C " & gCompile.BatchFile
   wszParam   = " /C " & gCompile.BatchFile

   With ShExecInfo
      .cbSize       = Len(SHELLEXECUTEINFOW)
      .fMask        = SEE_MASK_NOCLOSEPROCESS
      .HWnd         = 0
      .lpVerb       = Null
      .lpFile       = @wszFileExe
      .lpParameters = @wszParam   
      .lpDirectory  = 0
      .nShow        = SW_HIDE
      .hInstApp     = 0 
   End With
   
   
   ' Indicate in the StatusBar that we have started to compile
   SetCompileStatusBarMessage( "", 0 )

   dim as hwnd hStatusBar = GetDlgItem( HWND_FRMMAIN, IDC_FRMMAIN_STATUSBAR )
   dim as HWND hProgressBar = GetDlgItem(hStatusBar, IDC_FRMMAIN_PROGRESSBAR)
   SendMessage( hProgressBar, PBM_SETMARQUEE, 1, 16 )   ' every 16 milliseconds
   ShowWindow( hProgressBar, SW_SHOW)


   ' Shell to the command interpreter and run the batch file
   gCompile.StartTime = Timer
   gCompiling = true    
   
   dim as HCURSOR hCurSave = GetCursor()
   SetCursor( LoadCursor(0, IDC_WAIT) )

   pDoc = gTTabCtl.GetActiveDocumentPtr()
   if pDoc THEN SciExec( pDoc->hWindow, SCI_SETCURSOR, SC_CURSORWAIT, 0 )
      
   
   ' Shell to the compiler 
   ShellExecuteEx(@ShExecInfo)
    
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

   SetCursor( hCurSave )
   if pDoc THEN SciExec( pDoc->hWindow, SCI_SETCURSOR, SC_CURSORNORMAL, 0 )
   gCompile.EndTime = Timer
   gCompiling = false
   ShowWindow( hProgressBar, SW_HIDE)

   ' Get the log file contents into a string in order to parse it for errors, etc
   If AfxFileExists(gCompile.LogFile) Then
      dim st as string 
      wsLogSt = ""
      f = Freefile
      Open gCompile.LogFile For INPUT as #f
      while eof(f) = 0
         line input #f, st
         wsLogSt = wsLogSt + wstr(st) + vbcrlf
      WEND
      Close #f
   End If   
   
   ' Delete the temporary batch file and the compiler log file
   Kill gCompile.BatchFile
   Kill gCompile.LogFile

   ' Copy the log file text to the logfile textbox
   AfxSetWindowText( GetDlgItem( HWND_FRMOUTPUT, IDC_FRMOUTPUT_TXTLOGFILE), wsLogSt )
   
   hLV = GetDlgItem(HWND_FRMOUTPUT, IDC_FRMOUTPUT_LISTVIEW)
   ListView_DeleteAllItems( hLV )
   
   ' Parse the log string.
   NumLines = AfxStrParseCount( wsLogSt, vbCrLf)
 
   For NextLine = 1 To NumLines                   
   
      wst = Trim(AfxStrParse(wsLogSt, NextLine, vbCrlf))                 
      wst_ucase = Ucase(wst)
      
      ' Deal with the situation where we might be trying to compile a 64-bit application
      ' from a 32-bit operating system.
      If Left(**wst_ucase, 16) = WSTR("THIS VERSION OF ") Then
         ' Simply set the log file text and show the compiler log file window rather
         ' than the default compiler results window.
         SetCompileStatusBarMessage( "", 0 )
         TabCtrl_SetCurSel( GetDlgItem( HWND_FRMOUTPUT, IDC_FRMOUTPUT_TABCONTROL), 1 )
         AfxSetWindowText( GetDlgItem( HWND_FRMOUTPUT, IDC_FRMOUTPUT_TXTLOGFILE), wst )
         ShowWindow HWND_FRMOUTPUT, SW_SHOW
         frmMain_PositionWindows(HWND_FRMMAIN)
         Exit function
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
      
      ' Check to see if linking failed
      If (Left(**wst_ucase, 6) = WSTR("ERROR!")) Then
         FF_ListView_InsertItem( hLV, NumWarnings, 0, WStr("0") )
         FF_ListView_InsertItem( hLV, NumWarnings, 1, WStr("") ) 
         FF_ListView_InsertItem( hLV, NumWarnings, 2, WStr("compiling resource FAILED: Error Code 1") ) 
         NumErrors = NumErrors + 1
      End If
      
      If Instr(**wst_ucase, WSTR("LINKING FAILED:")) Then 
         FF_ListView_InsertItem( hLV, NumWarnings, 0, WStr("0") ) 
         FF_ListView_InsertItem( hLV, NumWarnings, 1, WStr("") ) 
         FF_ListView_InsertItem( hLV, NumWarnings, 2, wst ) 
         NumErrors = NumErrors + 1
      End If
 
 
      ' Check for the "linking: " line because that line identifies the actual output filename being created
      If Left(**wst_ucase, 9) = WSTR("LINKING: ") Then 
         ' -o "WinFBE.exe" 
         i = Instr(wst, WSTR(" -o "))
         If i Then
            wszTemp = Mid(wst, i+4)
            i = Instr(wszTemp, wDQ & WSTR(" "))
            If i Then gCompile.OutputFilename = AfxStrRemove(Left(wszTemp, i), wDQ)
         End If      
      End If
      
      ' Check for any compiler warnings
      IsWarning = Instr(wst_ucase, WSTR(") WARNING "))
      If IsWarning Then Error_Start = IsWarning
      IsError   = Instr(wst_ucase, WSTR(") ERROR ")) 
      If IsError Then Error_Start = IsError
      
      If (Error_Start > 0) Then 
         'sample warning message
         'c:\freebasic\test.bas(1394) warning 1(1): Passing scalar as pointer, at parameter 2 (hwndOldFocus) of ONSETFOCUS()
         'sample error message
         'c:\freebasic\test.bas(17) error 41: Variable not declared, kjljjada in 'kjljjada Error'
         Parenthesis_Start = Instr(1, wst, WSTR("("))
         Parenthesis_End   = Instr(Parenthesis_Start, wst, WSTR(")"))
         If (Parenthesis_Start < Parenthesis_End) And _
            (Parenthesis_End <= Error_Start)      And _
            (Error_Start > 0) Then
   
            wst1 = AfxStrParse(wst, 1, WSTR("("))   ' filename
            wst2 = Mid( wst, Parenthesis_Start + 1, Parenthesis_End - Parenthesis_Start - 1)  ' line# 
            wst3 = Mid( wst, Error_Start + 1)  ' error message
 
            FF_ListView_InsertItem( hLV, NumWarnings+NumErrors, 0, wst2 )  ' line#
            FF_ListView_InsertItem( hLV, NumWarnings+NumErrors, 1, wst1 )  ' filename
            FF_ListView_InsertItem( hLV, NumWarnings+NumErrors, 2, wst3 )  ' error message
            
            If IsWarning Then 
               NumWarnings = NumWarnings + 1
            ElseIf IsError Then
               If nFirstErrorLine = -1 Then
                  nFirstErrorLine = NumWarnings + NumErrors
               End If
               NumErrors = NumErrors + 1
            End If            
         End If
      End If
 
   Next
 
   ' If there were no errors but the fGoodCompile flag was not set to True then the log file
   ' did not contain the required FB copyright notice. Something must have went wrong, like trying
   ' to call the 64 bit compiler using a 32 bit operating system. Show the message to the user via
   ' the Output window (log file).
   If fGoodCompile = False Then NumErrors = NumErrors + 1

   dim as CWSTR cwsCompileMsg = L(193, "Errors:") & " " & NumErrors & "  " & _
                                L(192, "Warnings:") & " " & NumWarnings 
                                

   If (NumWarnings = 0) AndAlso (NumErrors = 0) Then 
      ' Hide the Output window is already open but there is no longer any warnings
      ' or errors to show in the listview. Only close if the listview is active. Do not
      ' close if the user is looking at the logfile textbox.
      if IsWindowVisible(HWND_FRMOUTPUT) THEN
         if TabCtrl_GetCurSel( GetDlgItem(HWND_FRMOUTPUT, IDC_FRMOUTPUT_TABCONTROL) ) = 0 THEN
            ShowWindow( HWND_FRMOUTPUT, SW_HIDE )
            frmMain_PositionWindows(HWND_FRMMAIN)
         END IF   
      END IF 
      cwsCompileMsg = cwsCompileMsg + "  [ " + gCompile.OutputFilename + " ]"
      SetCompileStatusBarMessage( cwsCompileMsg, ghIconGood )
      If gConfig.HideCompile = False Then
         wst = wszOutputMsg & vbCrLf & _
               L(194,"Primary Source:")     & " " & gCompile.MainFilename & vbCrLf & _
               L(195,"Target Compilation:") & " " & gCompile.OutputFilename & vbCrLf & _
               L(196,"Compile Time:")       & " " & Format(gCompile.EndTime-gCompile.StartTime, "###0.0") & " " & L(198,"seconds") & "." & vbCrLf & _
               L(197,"File Size:")          & " " & FileLen(gCompile.OutputFilename) & " " & L(199,"bytes") & "."
         MessageBoxW( HWND_FRMMAIN, wst, L(191,"Compiler Results"), MB_ICONINFORMATION Or MB_OK) 
      else
         ' Simply sound a success
         MessageBeep(MB_OK)
      End If
   Else 
      ' There were errors and/or warnings
      If NumErrors > 0 Then gCompile.RunAfterCompile = False 
      
      cwsCompileMsg = cwsCompileMsg + "  [" + Format(Now, "yyyy-mm-dd h:mm:ss AM/PM") & "]" 
      SetCompileStatusBarMessage( cwsCompileMsg, ghIconBad )
      MessageBeep(MB_ICONWARNING)
      
      ' Position the Compiler Log to the first error/warning
      ListView_SelectItem( hLV, 0 )
      
      ' If fGoodCompile is False then something unusual occurred so better show the log file by default
      If fGoodCompile = False Then
         PositionOutputWindows(HWND_FRMOUTPUT)
      End If 
      ' If the Search Results, TODO, Notes tab is active then reposition to 
      ' the first tab for error listview
      select case TabCtrl_GetCurSel( GetDlgItem(HWND_FRMOUTPUT, IDC_FRMOUTPUT_TABCONTROL) ) 
         case 2, 3, 4
         TabCtrl_SetCurSel( GetDlgItem(HWND_FRMOUTPUT, IDC_FRMOUTPUT_TABCONTROL), 0 ) 
         PositionOutputWindows(HWND_FRMOUTPUT)
      end select
      ShowWindow( HWND_FRMOUTPUT, SW_SHOW)
      frmMain_PositionWindows(HWND_FRMMAIN)
      ' Set to error line position only after all windows have been shown and resized
      SetDocumentErrorPosition(hLV)
   End If
   
   If gCompile.RunAfterCompile Then 
      RunEXE gCompile.OutputFilename, gApp.Projects(idx).ProjectCommandLine
   End If
   gConfig.LastRunFilename = gCompile.OutputFilename
      
   Function = True   ' successful
End Function


''
''
Function RunEXE( ByRef wszFileExe As CWSTR, _
                 ByRef wszParam As CWSTR _
                 ) As Long
                         
   Dim ShExecInfo As SHELLEXECUTEINFOW  

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

   Function = 0
End Function



' ========================================================================================
' Set the cursor to the error position based on the selected line in frmCompileResults
' ========================================================================================
Function SetDocumentErrorPosition( ByVal hLV As HWnd ) As Long

   Dim wszErrorLine As WString * MAX_PATH
   Dim wszErrorFile As WString * MAX_PATH
   
   Dim As Long idx
   Dim As Long nCurSel = ListView_GetSelection(hLV)
   If nCurSel < 0 Then Return 0 
   
   Dim pDoc As clsDocument Ptr 

   ' Get the line number and filename of the selected item
   FF_ListView_GetItemText(hLV, nCurSel, 0, @wszErrorLine, MAX_PATH)
   FF_ListView_GetItemText(hLV, nCurSel, 1, @wszErrorFile, MAX_PATH)

   If AfxFileExists(wszErrorFile) = 0 Then Exit Function
   
   ' Position the editor to the Error line
   idx = gTTabCtl.GetTabIndexFromFilename( @wszErrorFile )
   If idx >= 0 Then
      gTTabCtl.SetFocusTab(idx)
      pDoc = gTTabCtl.GetActiveDocumentPtr()
   Else
      ' File is not already loaded. Load it now.
      pDoc = frmMain_OpenFileSafely(HWND_FRMMAIN, _
                              False, _    ' bIsNewFile
                              False, _    ' bIsTemplate
                              true, _     ' bShowInTab
                              false, _    ' bIsInclude
                              @wszErrorFile, _  ' pwszName
                              0 )  ' pDocIn
      
      ' Position all of the controls into place
      frmMain_PositionWindows(HWND_FRMMAIN)
   End If 
   SciExec( pDoc->hWindow, SCI_GOTOLINE, Val(wszErrorLine)-1, 0 )

End Function













