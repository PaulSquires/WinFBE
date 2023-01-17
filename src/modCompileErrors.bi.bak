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

declare Function SetDocumentErrorPosition( ByVal hLV As HWnd, Byval wID as long ) As Long
declare function SetLogFileTextbox() as long
declare function ParseLogForError( _
            byref wsLogSt as CWSTR, _
            byval bAllowSuccessMessage as Boolean, _
            Byval wID as long, _
            byval fCompile64 as Boolean, _
            byval fCompilingObjFiles as Boolean _
            ) as Boolean
declare function ResetScintillaCursors() as Long
declare Function RunEXE( ByRef wszFileExe As CWSTR, ByRef wszParam As CWSTR ) As Long
declare function SetCompileStatusBarMessage(byref wszText as wstring, byval hIconCompile as long) as LRESULT
declare function RedirConsoleToFile(byref wszExe AS wstring, byref wszCmdLine AS wstring, byref sConsoleText AS string ) as long
declare function CreateTempResourceFile() as boolean

