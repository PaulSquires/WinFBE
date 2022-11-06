'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2022 Paul Squires, PlanetSquires Software
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

#pragma once

''
''  COMPILE_TYPE
''  Handle information related to the currnet compile process 
''
Type COMPILE_TYPE
   MainFilename       As WString * MAX_PATH   ' main source file (full path and file.ext)
   MainName           As WString * MAX_PATH   ' main source file (Name only, no extension)
   MainFolder         As WString * MAX_PATH   ' main source folder 
   ResourceFile       As WString * MAX_PATH   ' full path and file.ext to resource file (if applicable) 
   TempResourceFile   As WString * MAX_PATH   ' full path and file.ext to temporary resource file (if applicable) 
   OutputFilename     As WString * MAX_PATH   ' resulting exe/dll/lib name 
   CompilerPath       As WString * MAX_PATH   ' full path and file.ext to fbc.exe
   ObjFolder          As WString * MAX_PATH   ' *.o for all modules (set depending on 32/64 bit) (full path)
   ObjFolderShort     As WString * MAX_PATH   ' ".\.wfbe\"
   ObjID              As WString * MAX_PATH   ' "32" or "64" appended to object name
   CompileFlags       As WString * 2048
   wszFullCommandLine as CWSTR                ' Command line sent to the FB compiler
   wszFullLogFile     as CWSTR                ' Full log file returned from the FB compiler
   wszOutputMsg       as CWSTR                ' Additional info during compile process (time/filesize)
   RunAfterCompile    As BOOLEAN
   SystemTime         AS SYSTEMTIME           ' System time when compile finished
   StartTime          As Double
   EndTime            As Double
   CompileID          as long                 ' Type of compile (wID). Needed in case frmOutput listview later clicked on.
   bInvalidImagePath  as Boolean              ' One or more images have invalid path will result in failed resource compile
   
   ' The following are used to diagnose an error thrown in WinFBE_VD_MAIN.bas. We will
   ' need to subtract the nMainStartLine from the error line number and then load pDocMain
   ' into the editor and position.
   pDocMain           as clsDocument ptr
   pDocMainOffset     as long                 ' Line in file where the pDocMain code is output
End Type

declare Function code_Compile( ByVal wID As Long ) As BOOLEAN



