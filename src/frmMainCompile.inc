'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2025 Paul Squires, PlanetSquires Software
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

#include once "frmMain.bi"
#include once "clsDocument.bi"

' ========================================================================================
' Compile common actions
' ========================================================================================
function OnCommand_CompileCommon( byval id as long ) as LRESULT

   select case id
      Case IDM_BUILDEXECUTE, IDM_COMPILE, IDM_REBUILDALL, IDM_QUICKRUN
         code_Compile( id )
      
      Case IDM_RUNEXE
         RunExe( _
             GetRunExecutableFilename, _
             iif( gApp.IsProjectActive, gApp.ProjectCommandLine, gApp.wszCommandLine ) _
             )
      
      Case IDM_COMMANDLINE
         frmCommandLine_Show( HWND_FRMMAIN )
   end select
   
   function = 0
end function

