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
''  Main module that handles the entire compile process
''
'------------------------------------------------------------------------------------------------------------------------
Function code_Compile( ByVal wID As Long ) As BOOLEAN

   Dim ShExecInfo As SHELLEXECUTEINFOW  
    
   Dim As Long Parenthesis_Start, Parenthesis_End, Error_Start       
   Dim As Long f, i, NumLines, NextLine, idx, r
   Dim As Long NumWarnings, NumErrors, IsError, IsWarning
   Dim As Long nFirstErrorLine = -1
   
   Dim wszTemp      As WString * MAX_PATH 
   Dim wszFileExe   As WString * MAX_PATH
   Dim wszParam     As WString * MAX_PATH  
   
   Dim As String execmd, sDQ, sLogSt, st, st_ucase, sTemp
   Dim As String sOutputMsg, sFileExe, sParam
   Dim As String st1, st2, st3        
   Dim As HWnd hLV
         
   Dim As BOOLEAN fCompile32, fCompile64, fGoodCompile
   
   sDQ = Chr(34)  ' double quotes

   ' Can only continue to compile if the primary source code file
   ' is not dirty and needs to be saved.
   If gConfig.CompileAutosave Then
      If OnCommand_FileSaveAll(HWND_FRMMAIN) Then Return True  ' user cancelled save
   End If
   

   ' Perform some pre-compile checks to see if we should continue.
   gConfig.FBWINCompiler32 = ProcessFromCurdrive(gConfig.FBWINCompiler32)
   gConfig.FBWINCompiler64 = ProcessFromCurdrive(gConfig.FBWINCompiler64)
   Select Case Ucase(gConfig.DefaultCompiler)
      Case "FBC 32BIT"
         fCompile32 = True: fCompile64 = False
         gCompile.CompilerPath = gConfig.FBWINcompiler32
      Case "FBC 64BIT"
         gCompile.CompilerPath = gConfig.FBWINcompiler64
         fCompile32 = False: fCompile64 = True
   End Select
   
   
   ' Check to see if the compiler exists
   If AfxFileExists(gCompile.CompilerPath) = 0 Then 
      MessageBoxW( HWND_FRMMAIN, L(202,"Invalid defined compiler path."), L(201,"Error"), _
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
      pDocMain = gApp.GetMainDocumentPtr()
      If pDocMain = 0 Then
         MessageBoxW( HWND_FRMMAIN, L(208,"No Main file specified for the project."), L(201,"Error"), _
                           MB_OK Or MB_ICONINFORMATION Or MB_DEFBUTTON1 )
         Function = False: Exit Function
      End If
   Else
      pDocMain = gTTabCtl.GetActiveDocumentPtr()
   End If
   If pDocMain = 0 Then Return 0
   
   gCompile.MainFilename = pDocMain->DiskFilename
   gCompile.MainName = AfxStrPathname("NAME", pDocMain->DiskFilename)
   gCompile.MainFolder = AfxStrPathname("PATH", pDocMain->DiskFilename)

   If gApp.IsProjectActive Then
      Dim pDocResource As clsDocument Ptr = gApp.GetResourceDocumentPtr()
      gCompile.ResourceFile = Iif( pDocResource, pDocResource->DiskFilename, WStr(""))

      gCompile.ObjFolder = gCompile.MainFolder & ".wfbe\"
      gCompile.ObjFolderShort = ".\.wfbe\"

      If fCompile32 Then 
         gCompile.ObjID = "32.o"
         gCompile.CompileFlags = gApp.ProjectOther32           
      End If
      If fCompile64 Then 
         gCompile.ObjID = "64.o"
         gCompile.CompileFlags = gApp.ProjectOther64           
      End If   
      
      ' Make sure the folders exist
      SHCreateDirectory( 0, gCompile.ObjFolder )
      
      Select Case gApp.ProjectType
         Case 0:   ' EXE (no special compile switch)
            gCompile.OutputFilename = gCompile.MainFolder & gCompile.MainName & ".exe" 
         Case 1
            gCompile.OutputFilename = gCompile.MainFolder & gCompile.MainName & ".dll" 
            gCompile.CompileFlags = gCompile.CompileFlags & " -dll "
         Case 2
            gCompile.OutputFilename = gCompile.MainFolder & "lib" & gCompile.MainName & ".a" 
            gCompile.CompileFlags = gCompile.CompileFlags & " -lib "
      End Select

      Select Case gApp.ProjectErrorOption
         Case 0:   ' No error checking (no special compile switch)
         Case 1:  gCompile.CompileFlags = gCompile.CompileFlags & " -e"           
         Case 2:  gCompile.CompileFlags = gCompile.CompileFlags & " -ex"           
         Case 3:  gCompile.CompileFlags = gCompile.CompileFlags & " -exx"           
      End Select

      If gApp.ProjectDebug = BST_CHECKED Then
         gCompile.CompileFlags = gCompile.CompileFlags & " -g"           
      End If
      
      If gApp.ProjectThread = BST_CHECKED Then
         gCompile.CompileFlags = gCompile.CompileFlags & " -mt"           
      End If

   Else
      ' No active project
      gCompile.CompileFlags = gConfig.CompilerSwitches   
      gCompile.OutputFilename = gCompile.MainFolder & gCompile.MainName & ".exe"
      If Instr(Ucase(gCompile.CompileFlags), " -DLL") Then
         gCompile.OutputFilename = gCompile.MainFolder & gCompile.MainName & ".dll"
      End If
      If Instr(Ucase(gCompile.CompileFlags), " -DYLIB") Then
         gCompile.OutputFilename = gCompile.MainFolder & gCompile.MainName & ".dll"
      End If
      If Instr(Ucase(gCompile.CompileFlags), " -LIB") Then
         gCompile.OutputFilename = gCompile.MainFolder & "lib" & gCompile.MainName & ".a"
      End If
   End If


   Select Case UCase(gConfig.DefaultCompileMode)   ' used on the Main source file
      Case "GUI"
         ' For non-projects simply suppress the console. For projects, the console
         ' will not show unless option is selected in Project Options.
         gCompile.CompileMode = " -s gui "   ' default do not show console
         If gApp.ProjectShowConsole = BST_CHECKED Then
            gCompile.CompileMode = " "   ' console is not suppressed
         End If   
      Case "CONSOLE": gCompile.CompileMode = " -s console "
   End Select   
  

   ' Need to test if the resulting application to be compiled is actually running
   ' in memory. This would cause the compile to fail.
   ' Test for the running exe if this project is an EXE project.
   If IsProcessRunning(@gCompile.OutputFilename) Then
      MessageBoxW( HWND_FRMMAIN, L(200,"Program running..."), L(201,"Error"), _
                        MB_OK Or MB_ICONINFORMATION Or MB_DEFBUTTON1 )
      Function = False: Exit Function
   End If
   
   ' Create a temporary batch that will contain the compile commands. We shell
   ' to and run this batch file in order to do the compile.
   ' Replace any embedded spaces with underscores
   wszTemp = AfxStrPathname("NAME", gApp.ProjectName)
   sTemp = "_" & Str(wszTemp)
   sTemp = AfxStrReplace(sTemp, " ", "_" )
   gCompile.BatchFile = sTemp & "_compile.bat" 
   gCompile.LogFile   = sTemp & "_compile.log"
   
   ' Change to the output code directory
   ChDir gCompile.MainFolder
   
   ' Must delete any previous instance of the compiler log file in case it was not deleted. Need
   ' to do this because the compiler appends to the log file rather than overwriting.
   Kill gCompile.LogFile
   
   f = Freefile
   Open gCompile.BatchFile For Output as #f
   
   Dim As Long nCount = gApp.GetDocumentCount 
   Dim pDoc As clsDocument Ptr

   ' If this is a project then we need to compile all of the modules first
   ' in order to create the necessary *.o object files
   If gApp.IsProjectActive Then

      For i = 0 To nCount - 1
         pDoc = gApp.GetDocumentPtr(i)
         If pDoc->IsProjectFile Then
            If pDoc->ProjectFileType = FILETYPE_MODULE Then
               ' Get the base name of the file for constructing the object filename
               wszTemp = AfxStrPathname("NAME", pDoc->DiskFilename)
               wszTemp = gCompile.ObjFolderShort & wszTemp & gCompile.ObjID
               
               ' Compare the source code file datetime to the object. If the source code
               ' date time is greater then we need to recompile it, otherwise we will simply
               ' link to the existing object file.
               Dim As FILETIME ft1 = AfxGetFileLastWriteTime(pDoc->DiskFilename) ' source file
               Dim As FILETIME ft2 = AfxGetFileLastWriteTime(wszTemp)  ' object file

               If AfxFileTimeToVariantTime(ft1) > AfxFileTimeToVariantTime(ft2) Then
                  Print #f, sDQ & gCompile.CompilerPath & sDQ & _
                            " -c " & sDQ & pDoc->DiskFilename & sDQ & _
                            " -o " & sDQ & wszTemp & sDQ;
                  Print #f, " >> " & gCompile.LogFile & " 2>&1"   ' ensure stderr is redirected (appended) to log file as well
               End If
               
            End If
         End If
      Next
   End If

   
   ' Compile the Main file. If this is a project then we also need to link
   ' in all of the *.o object files.
   Print #f, sDQ & gCompile.CompilerPath & sDQ & _
            " " & sDQ & gCompile.MainFilename & sDQ & _ 
            " " & Iif(Len(gCompile.ResourceFile), sDQ & gCompile.ResourceFile & sDQ, WStr("")) & _
            " -v " & gCompile.CompileFlags & gCompile.CompileMode;
        
   If gApp.IsProjectActive Then
      For i = 0 To nCount - 1
         pDoc = gApp.GetDocumentPtr(i)
         If pDoc->IsProjectFile Then
            If pDoc->ProjectFileType = FILETYPE_MODULE Then
               ' Get the base name of the file for constructing the object filename
               wszTemp = AfxStrPathname("NAME", pDoc->DiskFilename)
               Print #f, " " & gCompile.ObjFolderShort & wszTemp & gCompile.ObjID;
            End If
         End If
      Next
   End If

   Print #f, " >> " & gCompile.LogFile & " 2>&1"   ' ensure stderr is redirected (appended) to log file as well
   Close #f

         
   ' Shell to the command interpreter and run the batch file
   gCompile.StartTime = Timer

   wszFileExe = Environ("COMSPEC")
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
   ShellExecuteEx(@ShExecInfo)
   WaitForSingleObject(ShExecInfo.hProcess,INFINITE)

   gCompile.EndTime = Timer


   ' Get the log file contents into a string in order to parse it for errors, etc
   If AfxFileExists(gCompile.LogFile) Then
      f = Freefile
      Open gCompile.LogFile For Binary as #f 
      sLogSt = Space(LOF(f))
      Get #f, , sLogSt
      Close #f
   End If   
   
   ' Delete the temporary batch file and the compiler log file
   Kill gCompile.BatchFile
   Kill gCompile.LogFile

   
   ' Copy the log file text to the logfile textbox
   AfxSetWindowText( GetDlgItem( HWND_FRMOUTPUT, IDC_FRMOUTPUT_TXTLOGFILE), sLogSt )
   
   hLV = GetDlgItem(HWND_FRMOUTPUT, IDC_FRMOUTPUT_LISTVIEW)
   ListView_DeleteAllItems( hLV )
   
   ' Parse the log string.
   NumLines = AfxStrParseCount( sLogSt, vbCrLf)
 
   For NextLine = 1 To NumLines                   
   
      st = Trim(AfxStrParse(sLogSt, NextLine, vbCrlf))                 
      st_ucase = Ucase(st)
      
      ' Deal with the situation where we might be trying to compile a 64-bit application
      ' from a 32-bit operating system.
      If Left(st_ucase, 16) = "THIS VERSION OF " Then
         ' Simply set the log file text and show the compiler log file window rather
         ' than the default compiler results window.
         TabCtrl_SetCurSel( GetDlgItem( HWND_FRMOUTPUT, IDC_FRMOUTPUT_TABCONTROL), 1 )
         AfxSetWindowText( GetDlgItem( HWND_FRMOUTPUT, IDC_FRMOUTPUT_TXTLOGFILE), st )
         ShowWindow HWND_FRMOUTPUT, SW_SHOW
         frmMain_PositionWindows(HWND_FRMMAIN)
         Exit function
      End If   

      ' Save the FB version and copyright information. 
      If Left(st_ucase, 19) = "FREEBASIC COMPILER " Then
         fGoodCompile = True
         sOutputMsg = sOutputMsg & st & vbCrLf: Continue For
      End If        
      If Left(st_ucase, 13) = "COPYRIGHT (C)" Then
         fGoodCompile = True
         sOutputMsg = sOutputMsg & st & vbCrLf: Continue For
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
      If (Left(st_ucase, 6) = "ERROR!") Then
         FF_ListView_InsertItem( hLV, NumWarnings, 0, WStr("0") )
         FF_ListView_InsertItem( hLV, NumWarnings, 1, WStr("") ) 
         FF_ListView_InsertItem( hLV, NumWarnings, 2, WStr("compiling resource FAILED: Error Code 1") ) 
         NumErrors = NumErrors + 1
      End If
      
      If Instr(st_ucase, "LINKING FAILED:") Then 
         FF_ListView_InsertItem( hLV, NumWarnings, 0, WStr("0") ) 
         FF_ListView_InsertItem( hLV, NumWarnings, 1, WStr("") ) 
         FF_ListView_InsertItem( hLV, NumWarnings, 2, WStr(st) ) 
         NumErrors = NumErrors + 1
      End If
 
 
      ' Check for the "linking: " line because that line identifies the actual output filename being created
      If Left(st_ucase, 9) = "LINKING: " Then 
         ' -o "WinFBE.exe" 
         i = Instr(st, " -o ")
         If i Then
            sTemp = Mid(st, i+4)
            i = Instr(sTemp, sDQ & " ")
            If i Then gCompile.OutputFilename = AfxStrRemove(Left(sTemp, i), sDQ)
         End If      
      End If
      
      ' Check for any compiler warnings
      IsWarning = Instr(st_ucase, ") WARNING ")
      If IsWarning Then Error_Start = IsWarning
      IsError   = Instr(st_ucase, ") ERROR ") 
      If IsError Then Error_Start = IsError
      
      If (Error_Start > 0) Then 
         'sample warning message
         'c:\freebasic\test.bas(1394) warning 1(1): Passing scalar as pointer, at parameter 2 (hwndOldFocus) of ONSETFOCUS()
         'sample error message
         'c:\freebasic\test.bas(17) error 41: Variable not declared, kjljjada in 'kjljjada Error'
         Parenthesis_Start = Instr(1, st, "(")
         Parenthesis_End   = Instr(Parenthesis_Start, st, ")")
         If (Parenthesis_Start < Parenthesis_End) And _
            (Parenthesis_End <= Error_Start)       And _
            (Error_Start > 0) Then
   
            st1 = AfxStrParse(st, 1, "(")   ' filename
            st2 = Mid( st, Parenthesis_Start + 1, Parenthesis_End - Parenthesis_Start - 1)  ' line# 
            st3 = Mid(st, Error_Start+1)  ' error message
 
            FF_ListView_InsertItem( hLV, NumWarnings+NumErrors, 0, WStr(st2) )  ' line#
            FF_ListView_InsertItem( hLV, NumWarnings+NumErrors, 1, WStr(st1) )  ' filename
            FF_ListView_InsertItem( hLV, NumWarnings+NumErrors, 2, WStr(st3) )  ' error message
            
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

   If (NumWarnings = 0) AndAlso (NumErrors = 0) Then 
      If gConfig.HideCompile = False Then
         st = sOutputMsg & vbCrLf & _
               L(194,"Primary Source:")     & " " & gCompile.MainFilename & vbCrLf & _
               L(195,"Target Compilation:") & " " & gCompile.OutputFilename & vbCrLf & _
               L(196,"Compile Time:")       & " " & Format(gCompile.EndTime-gCompile.StartTime, "###0.0") & " " & L(198,"seconds") & "." & vbCrLf & _
               L(197,"File Size:")          & " " & FileLen(gCompile.OutputFilename) & " " & L(199,"bytes") & "."
         MessageBoxW( HWND_FRMMAIN, WStr(st), L(191,"Compiler Results"), MB_ICONINFORMATION Or MB_OK) 
      End If
   Else
      ' There were errors and/or warnings
      If NumErrors > 0 Then gCompile.RunAfterCompile = False 
      
      MessageBeep(MB_ICONWARNING)
      
      ' Position the Compiler Log to the first error/warning
      ListView_SelectItem( hLV, 0 )
      
      ' If fGoodCompile is False then something unusual occurred so better show the log file by default
      If fGoodCompile = False Then
         PositionOutputWindows(HWND_FRMOUTPUT)
      End If   
      ShowWindow HWND_FRMOUTPUT, SW_SHOW
      frmMain_PositionWindows(HWND_FRMMAIN)
      ' Set to error line position only after all windows have been shown and resized
      SetDocumentErrorPosition(hLV)
   End If
   
   If gCompile.RunAfterCompile Then 
      RunEXE gCompile.OutputFilename, gConfig.CommandLine
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

   If Dir(wszErrorFile) = "" Then Exit Function
   
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
   SendMessageW pDoc->hWindow, SCI_GOTOLINE, Val(wszErrorLine)-1, 0

End Function













