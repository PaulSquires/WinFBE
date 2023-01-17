''    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
''    Copyright (C) 2016-2022 Paul Squires, PlanetSquires Software
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
#include once "modCompileErrors.bi"
#include once "modGenerateCode.bi"
#include once "frmUserTools.bi"
#include once "frmOutput.bi"

Dim Shared gCompile As COMPILE_TYPE

''
'' Main module that handles the entire compile process
''
function code_Compile( ByVal wID As Long ) As BOOLEAN

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
                             
   gCompile.wszFullCommandLine = ""
   gCompile.wszFullLogFile = ""
   gCompile.wszOutputMsg = ""
   gCompile.bInvalidImagePath = false
   
   
   ' Indicate in the StatusBar that we have started to compile
   SetCompileStatusBarMessage( L(449, "Generating compiler files") & "...", 0 )
   

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
            SetCompileStatusBarMessage( "", 0 )
            Return True  ' user cancelled save
         end if   
         ' Reload the main form files combobox in case names have changed
         ' pDoc->SaveFile do this (OnCommand_FileSaveAll calls pDoc->SaveFile)
      End If
   END IF
   
   ' Get the array index for the current selected build (project or document)
   idxBuild = frmBuildConfig_getActiveBuildIndex()
   if idxBuild = -1 then idxBuild = 0
   
   ' Perform some pre-compile checks to see if we should continue.
   gConfig.FBWINCompiler32 = ProcessFromCurdriveApp(gConfig.FBWINCompiler32)
   gConfig.FBWINCompiler64 = ProcessFromCurdriveApp(gConfig.FBWINCompiler64)
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
      SetCompileStatusBarMessage( "", 0 )
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
         SetCompileStatusBarMessage( "", 0 )
         MessageBox( HWND_FRMMAIN, L(208,"No Main file specified for the project."), L(201,"Error"), _
                           MB_OK Or MB_ICONINFORMATION Or MB_DEFBUTTON1 )
         Function = False: Exit Function
      End If
   Else
      pDocMain = gTTabCtl.GetActiveDocumentPtr()
   End If
   If pDocMain = 0 Then 
      SetCompileStatusBarMessage( "", 0 )
      return 0
   end if
   
 
   ' Do a check to ensure that only one call to Application.Run exists, otherwise
   ' present a warning message before continuing.
   dim as CWSTR wszWarning = L(401,"Multiple calls to Application.Run exist. Check the source code of these files:") & ":"
   dim as long TotalAppRunCount = 0
   pDoc = gApp.pDocList
   do until pDoc = 0
      if pDoc->AppRunCount > 0 then
         wszWarning = wszWarning & vbcrlf &  "- " & AfxStrPathname( "NAMEX", pDoc->DiskFilename ) 
         TotalAppRunCount = TotalAppRunCount + pDoc->AppRunCount
      end if
      pDoc = pDoc->pDocNext
   loop
   if TotalAppRunCount > 1 then
      SetCompileStatusBarMessage( "", 0 )
      MessageBox( HWND_FRMMAIN, wszWarning, L(109,"Warning"), _
                        MB_OK Or MB_ICONWARNING Or MB_DEFBUTTON1 )
      Function = False: Exit Function
   end if

   
   ' If this is a QuickRun then we need to copy the current source code to a temporary bas file
   ' and compile that bas file instead. The resulting bas file will be deleted after the compiling
   ' has completed and exe will be deleted when the exe is terminated. 
   dim as CWSTR wszTempMainFilename
   if wID = IDM_QUICKRUN then
      if pDocMain->IsNewFlag then 
         gCompile.MainFolder = AfxGetExePathName
      else
         gCompile.MainFolder = AfxStrPathname("PATH", pDocMain->DiskFilename)
      end if
      wszTempMainFilename   = GetTemporaryFilename(gCompile.MainFolder, "bas")
      gCompile.MainFilename = wszTempMainFilename
      gCompile.MainName     = AfxStrPathname("NAME", gCompile.MainFilename)
      
      dim as CWSTR wszText = pDocMain->GetText()

      ' Regenerate any visual designer code for Forms.
      ' The codegen is stored in each pDoc's wszFormCodeGen
      dim as long initialCtrlID = 10000

      pDoc = gApp.pDocList
      do until pDoc = 0
         if pDoc->IsDesigner then 
            pDoc->initialCtrlID = initialCtrlID
            pDoc->bRegenerateCode = true
            GenerateFormCode( pDoc )
            initialCtrlID  = initialCtrlID  + 10000
         end if
         pDoc = pDoc->pDocNext
      loop

      dim pTextStream as CTextStream
      if pTextStream.OpenUnicode(gCompile.MainFilename, IOMODE_FORWRITING, true) = S_OK then
         pTextStream.Write( wszText )
      end if    
      pTextStream.Close

   else
      gCompile.MainFilename = pDocMain->DiskFilename
      gCompile.MainName     = AfxStrPathname("NAME", pDocMain->DiskFilename)
      gCompile.MainFolder   = AfxStrPathname("PATH", pDocMain->DiskFilename)
   end if
   
   gCompile.ResourceFile = ""    ' default that there is no resource
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

      gCompile.ObjFolder = gCompile.MainFolder & ".wfbe\"
      gCompile.ObjFolderShort = ".\.wfbe\"

      If fCompile32 Then 
         gCompile.ObjID = "32.o"
         gCompile.CompileFlags = gCompile.CompileFlags + " " + gApp.ProjectOther32           
      End If
      If fCompile64 Then 
         gCompile.ObjID = "64.o"
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
            gCompile.CompileFlags = gCompile.CompileFlags + " -s console "
         case IDM_GUI
            n = instr(gCompile.CompileFlags, " -S CONSOLE ")
            if n THEN gCompile.CompileFlags = AfxStrDelete(gCompile.CompileFlags, n, 12)
            gCompile.CompileFlags = gCompile.CompileFlags + " -s gui "
         case IDM_RESOURCE
            gCompile.ResourceFile = directives(i).DirectiveText
      END select
   next


   ' Need to determine the output filename in order to search for any existing
   ' running process.
   gCompile.OutputFilename = gCompile.MainFolder & gCompile.MainName & ".exe"
   If Instr(" " & Ucase(gCompile.CompileFlags), " -DLL") Then
      gCompile.OutputFilename = gCompile.MainFolder & gCompile.MainName & ".dll"
   End If
   If Instr(" " & Ucase(gCompile.CompileFlags), " -DYLIB") Then
      gCompile.OutputFilename = gCompile.MainFolder & gCompile.MainName & ".dll"
   End If
   If Instr(" " & Ucase(gCompile.CompileFlags), " -LIB") Then
      gCompile.OutputFilename = gCompile.MainFolder & "lib" & gCompile.MainName & ".a"
   End If


   ' Need to test if the resulting application to be compiled is actually running
   ' in memory. This would cause the compile to fail.
   If IsProcessRunning(@gCompile.OutputFilename) Then
      SetCompileStatusBarMessage( "", 0 )
      MessageBox( HWND_FRMMAIN, L(200,"Program running") & "...", L(201,"Error"), _
                        MB_OK Or MB_ICONINFORMATION Or MB_DEFBUTTON1 )
      Function = False: Exit Function
   End If
      

   gApp.IsCompiling = true    

   dim as HCURSOR hCurSave = GetCursor()
   SetCursor( LoadCursor(0, IDC_WAIT) )
   pDoc = gTTabCtl.GetActiveDocumentPtr()
   if pDoc THEN SciExec( pDoc->hWindow(0), SCI_SETCURSOR, SC_CURSORWAIT, 0 )


   ' Regenerate any visual designer code for Forms.
   ' The codegen is stored in each pDoc's wszFormCodeGen
   dim as long initialCtrlID = 10000

   pDoc = gApp.pDocList
   do until pDoc = 0
      if pDoc->IsDesigner then 
         pDoc->initialCtrlID = initialCtrlID
         pDoc->bRegenerateCode = true
         GenerateFormCode( pDoc )
         initialCtrlID  = initialCtrlID  + 10000
      end if
      pDoc = pDoc->pDocNext
   loop
   
   ' Do another Save All because a call to GenerateFormCode may have created
   ' Event code that previously did not exist.
   if wID <> IDM_QUICKRUN then
      OnCommand_FileSaveAll(HWND_FRMMAIN)
   end if

   ' Run any precompile tools
   for y as long = LBound(gConfig.Tools) To UBound(gConfig.Tools)
      if gConfig.Tools(y).Action = USERTOOL_ACTION_PRECOMPILE then
         frmUserTools_ExecuteUserTool(y)
      end if   
   next

   ' Change to the output code directory
   ChDir gCompile.MainFolder
   
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
               wszParams = " -c -b " + _
                           wDQ + pDoc->DiskFilename + wDQ + _
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

   
   ' If this is a visual designer project then output specific equates that dictate
   ' which controls get compiled into the file source.
   dim as CWSTR wszVDEquates, wszCtrlEquate
   dim as boolean bIsVisual = false
   
   ' In 2.0.8 and lower, the following two defines were put into the user code
   ' when a new main.bas file was created by New Project. From 2.0.9+ we move that
   ' code to the WinFBE_VD_MAIN.bas file and comment it out from the Main file.
   dim as CWSTR wszUnicodeEquate = "#Define UNICODE" & vbcrlf 
   dim as CWSTR wszDefineWin32Equate = "#Define _WIN32_WINNT &h0602" & vbcrlf
   
   wszVDEquates = wszUnicodeEquate & _
                  wszDefineWin32Equate & _
                  "#Include Once " & chr(34) & "windows.bi" & chr(34) & vbcrlf & _
                  "#Include Once " & chr(34) & "Afx\CLayout.inc" & chr(34) & vbcrlf & _
                  "#Include Once " & chr(34) & "Afx\CWindow.inc" & chr(34) & vbcrlf
      
   pDoc = iIf( gApp.IsProjectActive, gApp.pDocList, pDocMain )
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

      If gApp.IsProjectActive = false Then exit do
      pDoc = pDoc->pDocNext
   loop


   wszVDEquates = wszVDEquates & _
      "#Include once " & chr(34) & "WinFormsX\WinFormsX.bi" & chr(34) & vbcrlf & _
      "Using Afx" & vbcrlf 

   dim wszVDMain as CWSTR 
   if bIsVisual then
      wszVDMain = gCompile.MainFolder & "WINFBE_VD_MAIN.bas"
      dim pStream as CTextStream
      if pStream.Create(wszVDMain, true, true) = S_OK then 
         ' Get all of the text from the MAIN pDoc file
         if pDocMain->hWindow(0) = 0 then
            dim as CWSTR wszFileType = pDocMain->ProjectFileType
            pDocMain->CreateCodeWindow( HWND_FRMMAIN, false, false, pDocMain->DiskFilename )
            pDocMain->ProjectFileType = wszFileType
         end if
         
         dim as CWSTR wszMainFileText 
         wszMainFileText.Utf8 = pDocMain->GetText  ' Convert from utf8

         ' In 2.0.8 and lower, the following two defines were put into the user code
         ' when a new main.bas file was created by New Project. From 2.0.9+ we move that
         ' code to the WinFBE_VD_MAIN.bas file and comment it out from the Main file.
         wszMainFileText = AfxStrReplace( wszMainFileText, wszUnicodeEquate, "'" & wszUnicodeEquate )
         wszMainFileText = AfxStrReplace( wszMainFileText, wszDefineWin32Equate, "'" & wszDefineWin32Equate )

         dim as CWSTR wszAllTYPEs, wszAllEvents
         
         ' Output all of the visual designer generated code for each Form in the project
         pDoc = iIf( gApp.IsProjectActive, gApp.pDocList, pDocMain )
         do until pDoc = 0
				if pDoc->IsDesigner then 
					wszAllTYPEs = wszAllTYPEs & pDoc->wszFormCodeGen
					if pDoc <> pDocMain then
						wszAllEvents = wszAllEvents & _
											"#include once " & chr(34) & pDoc->DiskFilename & chr(34) & vbcrlf
					end if
				end if
            If gApp.IsProjectActive = false Then exit do
            pDoc = pDoc->pDocNext
         loop
         
         dim as CWSTR wszText = _
                        wszVDEquates    & vbcrlf & _
                        wszAllTYPEs     & vbcrlf       ' this also includes Declares for the Events

         ' Save the pDocMain in case WinFBE_VD_MAIN.bas throws an error and we have
         ' to load the original main file into the editor and reposition to the error line.
         gCompile.pDocMain = pDocMain
         ' Determine the number of vbcrlf's to this point because we need to store the line offset
         ' to where the Main code is to written.
         gCompile.pDocMainOffset = AfxStrTally( wszText, vbcrlf )
         
         ' Continue to build the output string and then write it disk.
         wszText = wszText & _
                        wszMainFileText & vbcrlf & _   '<-- ensure vbcrlf because Main file may not have crlf on last line
                        wszAllEvents

         
         pStream.WriteLine wszText
         pStream.Close
         gCompile.MainFilename = wszVDMain
      end if
   end if
   

   ' Compile the Main file. If this is a project then we also need to link
   ' in all of the *.o object files.
   ' Per fxm post https://www.freebasic.net/forum/viewtopic.php?f=8&p=266126#p266126
   ' and https://www.freebasic.net/forum/viewtopic.php?f=3&t=28723&p=274431#p274431
   ' Do not specify file extension with -m parameter.
   ' So the main module file must be called twice in the command line:
   ' - after compile option '-m', but without specified extension,
   ' - and also like any module to compile, but there with its specified extension.
   'dim as CWSTR wszMfile = wDQ + AfxStrPathName("PATH", gCompile.MainFilename) + AfxStrPathName("NAME", gCompile.MainFilename) + wDQ
   'dim as CWSTR wszMfile = AfxStrPathName("NAME", gCompile.MainFilename) 
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
   
   ' Present a warning that the compile will fail if one or more of the Images in the
   ' resource file have an invalid path.
   if gCompile.bInvalidImagePath then
      gApp.IsCompiling = false
      SetCursor( hCurSave )
      ResetScintillaCursors()
      SetCompileStatusBarMessage( "", 0 )
      dim as CWSTR wszWarning = L(404,"Image Manager contains invalid image paths. Resource file creation will fail.")
      MessageBox( HWND_FRMMAIN, wszWarning, L(109,"Warning"), _
                        MB_OK Or MB_ICONWARNING Or MB_DEFBUTTON1 )
      Function = False: Exit Function
   end if
   
   
   ' Ensure verbose compiler message output
   wszParams = wszParams + " -v " + gCompile.CompileFlags 
   
   f = instr(wszParams, " -x ") 
   if f = 0 then
      wszParams = wszParams + " -x " + wDQ + gCompile.OutputFilename + wDQ
   else
      ' The ParseLogForError() function below will identify the correct
      ' filename for gCompile.OutputFilename based on the man file being linked.
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
   ' Both the temporary disk file and (if applicable) the VD MAIN CODEGEN file.
   if wID = IDM_QUICKRUN THEN 
      AfxDeleteFile( gCompile.MainFilename )
      AfxDeleteFile( wszTempMainFilename )
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
         ' If gCompile.OutputFilename is a relative filename then convert it to full filename.
         gCompile.OutputFilename = AfxStrReplace( gCompile.OutputFilename, "/", "\" )
         if AfxPathIsRelative( gCompile.OutputFilename ) then 
            gCompile.OutputFilename = AfxPathCombine( gCompile.MainFolder, gCompile.OutputFilename )
         end if

         RunEXE( _
            gCompile.OutputFilename,    _
            iif( gApp.IsProjectActive, gApp.ProjectCommandLine, gApp.wszCommandLine ) _
            )
      End If
   
   end if
   
   gApp.IsCompiling = false

'   if AfxFileExists( wszVDMain ) then
'      AfxDeleteFile( wszVDMain )
'   end if

   Function = True   ' successful
End Function


