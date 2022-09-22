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

enum eFileClose
   EFC_CLOSECURRENT
   EFC_CLOSEALL
   EFC_CLOSEALLFORWARD
   EFC_CLOSEALLOTHERS
   EFC_CLOSEALLBACKWARD 
end enum

Declare Function OnCommand_FileNew( ByVal HWnd As HWnd ) As clsDocument ptr
Declare Function OnCommand_FileOpen( ByVal HWnd As HWnd, byval bShowInTab as Boolean = true ) As LRESULT
declare function OnCommand_FileTemplates( ByVal HWnd As HWnd ) as LRESULT
Declare Function OnCommand_FileSave( ByVal HWnd As HWnd, byval pDoc as clsDocument ptr, ByVal bSaveAs As BOOLEAN = False, ByVal bSaveAll As BOOLEAN = False ) As LRESULT
Declare Function OnCommand_FileSaveDeclares( ByVal HWnd As HWnd ) As LRESULT
Declare Function OnCommand_FileSaveAll( ByVal HWnd As HWnd ) As LRESULT
Declare Function OnCommand_FileClose( ByVal HWnd As HWnd, ByVal veFileClose As eFileClose, byval nTabNum as long = -1 ) As LRESULT
declare function OnCommand_FileExplorerMessage( byval id as long, byval pDoc as clsDocument ptr ) as LRESULT
declare function OnCommand_FileAutoSave( byval HWnd As HWnd ) as LRESULT
declare function OnCommand_FileAutoSaveStartTimer() as LRESULT
declare function OnCommand_FileAutoSaveKillTimer() as LRESULT
declare function OnCommand_FileAutoSaveGenerateFilename( byval wszFilename as CWSTR ) as CWSTR
declare function OnCommand_FileAutoSaveFileCheck( byval wszFilename as CWSTR ) as CWSTR
declare function OnCommand_FileLoadSession( byval HWnd As HWnd ) as LRESULT
declare function OnCommand_FileSaveSession( byval HWnd As HWnd ) as LRESULT

declare function OnCommand_EditRedo( ByVal hEdit as HWND ) as LRESULT
declare function OnCommand_EditUndo( ByVal hEdit as HWND ) as LRESULT
declare function OnCommand_EditCut( byval pDoc as clsDocument ptr, ByVal hEdit as HWND ) as LRESULT
declare function OnCommand_EditCopy( byval pDoc as clsDocument ptr, ByVal hEdit as HWND ) as LRESULT
declare function OnCommand_EditPaste( byval pDoc as clsDocument ptr, ByVal hEdit as HWND ) as LRESULT
declare function OnCommand_EditFindDialog() as LRESULT
declare function OnCommand_EditReplaceDialog() as LRESULT
declare function OnCommand_EditFindInFiles( byval hEdit as HWND ) as LRESULT
declare function OnCommand_EditFindActions( ByVal id as long, byval pDoc as clsDocument ptr ) as LRESULT
declare function OnCommand_EditReplaceActions( ByVal id as long, byval pDoc as clsDocument ptr ) as LRESULT
declare function OnCommand_EditIndentBlock( byval pDoc as clsDocument ptr, ByVal hEdit as HWND ) as LRESULT
declare function OnCommand_EditUnIndentBlock( byval pDoc as clsDocument ptr, ByVal hEdit as HWND ) as LRESULT
declare function OnCommand_EditSelectAll( byval pDoc as clsDocument ptr, ByVal hEdit as HWND ) as LRESULT
declare function OnCommand_EditEncoding( ByVal id as long, byval pDoc as clsDocument ptr ) as LRESULT
declare function OnCommand_EditCommon( ByVal id as long, byval pDoc as clsDocument ptr ) as LRESULT

declare function OnCommand_SearchGotoDefinition( byval pDoc as clsDocument ptr ) as LRESULT
declare function OnCommand_SearchGotoLastPosition() as LRESULT
declare function OnCommand_SearchGotoCompileError( byval bMoveNext as boolean ) as long
declare function OnCommand_SearchGotoFile( ByVal id as long, byval pDoc as clsDocument ptr ) as LRESULT
declare function OnCommand_SearchBookmarks( ByVal id as long, byval pDoc as clsDocument ptr ) as LRESULT

declare function OnCommand_ViewFunctionList() as LRESULT
declare function OnCommand_ViewBookmarksList() as LRESULT
declare function OnCommand_ViewExplorer() as LRESULT
declare function OnCommand_ViewOutput() as LRESULT
declare function OnCommand_ViewFold( ByVal id as long, byval pDoc as clsDocument ptr ) as LRESULT
declare function OnCommand_ViewZoom( ByVal id as long, byval pDoc as clsDocument ptr ) as LRESULT
declare function OnCommand_ViewToDo() as LRESULT
declare function OnCommand_ViewNotes() as LRESULT
declare function OnCommand_ViewRestoreMain() as LRESULT

Declare Function OnCommand_ProjectSave( ByVal HWnd As HWnd, ByVal bSaveAs As BOOLEAN = False ) As LRESULT
Declare Function OnCommand_ProjectClose( ByVal HWnd As HWnd ) As LRESULT
Declare Function OnCommand_ProjectNew( ByVal HWnd As HWnd ) As LRESULT
Declare Function OnCommand_ProjectOpen( ByVal HWnd As HWnd ) As LRESULT
declare function OnCommand_ProjectFilesAdd( byval HWnd As HWnd ) as LRESULT
declare function OnCommand_ProjectSetFileType( byval id as long, byval pDoc as clsDocument ptr ) as LRESULT
declare function OnCommand_ProjectRemove( byval id as long, byval pDoc as clsDocument ptr ) as LRESULT

declare function OnCommand_CompileCommon( byval id as long ) as LRESULT

Declare Function OnCommand_DesignerNewForm( ByVal HWnd As HWnd ) As LRESULT
Declare Function OnCommand_DesignerAlign( byval HWND as HWND, byval id as long ) as LRESULT
Declare Function OnCommand_DesignerCenter( byval HWND as HWND, byval id as long ) as LRESULT
declare function OnCommand_DesignerHorizSpacing( byval HWND as HWND, byval id as long ) as LRESULT
declare function OnCommand_DesignerVertSpacing( byval HWND as HWND, byval id as long ) as LRESULT
declare function OnCommand_DesignerDeleteKey( byval pDoc as clsDocument ptr ) as LRESULT
declare function OnCommand_DesignerToolBox( byval HWnd as HWND ) as LRESULT
declare function OnCommand_DesignerMenuEditor( ByVal HWnd As HWnd, byval pDoc as clsDocument ptr ) As LRESULT
declare function OnCommand_DesignerToolBarEditor( ByVal HWnd As HWnd, byval pDoc as clsDocument ptr ) As LRESULT
declare function OnCommand_DesignerStatusBarEditor( ByVal HWnd As HWnd, byval pDoc as clsDocument ptr, ByVal codeNotify As UINT ) As LRESULT
declare function OnCommand_DesignerImageManager( ByVal HWnd As HWnd, byval pDoc as clsDocument ptr ) As LRESULT
declare function OnCommand_DesignerSnapLines( byval pDoc as clsDocument ptr ) As LRESULT
declare function OnCommand_DesignerLockControls( byval pDoc as clsDocument ptr ) As LRESULT

declare Function frmMain_OnCommand( ByVal HWnd As HWnd, _
                                    ByVal id As Long, _
                                    ByVal hwndCtl As HWnd, _
                                    ByVal codeNotify As UINT _
                                    ) As LRESULT
