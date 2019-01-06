'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2019 Paul Squires, PlanetSquires Software
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

declare function RemoveComments( byval st as string) as string
declare function MaskStringCharacters( byval st as string) as string
declare function JulianDateNow() as long
declare function ConvertWinFBEversion( byref wszVersion as wstring ) as long
declare function DisableAllModeless() as long
declare function EnableAllModeless() as long
declare FUNCTION GetTemporaryFilename( byref wszFolder as wstring, BYREF wszExtension AS wSTRING) AS string
declare FUNCTION ComboBox_ReplaceString (BYVAL hComboBox AS HWND, BYVAL index AS LONG, BYVAL pwszNewText AS WSTRING PTR, BYVAL pNewData AS LONG_PTR = 0) AS LONG
declare Function GetFontCharSetID(ByREF wzCharsetName As CWSTR ) As Long
declare function RemoveDuplicateSpaces( byref sText as const string) as string
declare function ConvertCase( byval sText as string) as string
declare FUNCTION Utf8ToAscii(byref strUtf8 AS STRING) AS STRING
declare FUNCTION AnsiToUtf8(BYREF sAnsi AS STRING) AS STRING
declare FUNCTION Utf8ToUnicode(BYREF ansiStr AS CONST STRING) AS STRING
declare FUNCTION UnicodeToUtf8(byval pswzUnicode as wstring ptr) AS STRING
declare function GetFileToString( byref wszFilename as const wstring, byref txtBuffer as string, byval pDoc as clsDocument ptr ) as boolean
declare function ConvertTextBuffer( byval pDoc as clsDocument ptr, byval FileEncoding as long ) as Long
declare function IsCurrentLineIncludeFilename() as Boolean
declare function OpenSelectedDocument( byref wszFilename as wstring, byref wszFunctionName as WSTRING = "", byval nLineNumber as long = -1 ) as clsDocument ptr
declare Function ProcessToCurdrive( Byval wzFilename As CWSTR ) As CWSTR
declare Function ProcessFromCurdrive( Byval wzFilename As CWSTR ) As CWSTR
Declare Function FF_TreeView_InsertItem( ByVal hWndControl As HWnd, ByVal hParent As HANDLE, ByRef TheText As WString, ByVal lParam As LPARAM = 0, ByVal iImage As Long = 0, ByVal iSelectedImage As Long = 0 ) As HANDLE
declare Function FF_TreeView_GetlParam( ByVal hWndControl As HWnd, ByVal hItem As HANDLE ) As Long
declare Function FF_TreeView_SetlParam (ByVal hWndControl as HWnd, ByVal hItem as HANDLE, ByVal lParam as Long) as Long
declare Function AfxIFileOpenDialogW( ByVal hwndOwner As HWnd, ByVal idButton As Long) As WString Ptr
declare Function AfxIFileOpenDialogMultiple( ByVal hwndOwner As HWnd, ByVal idButton As Long) As IShellItemArray Ptr
Declare Function AfxIFileSaveDialog( ByVal hwndOwner As HWnd, ByVal pwszFileName As WString Ptr, ByVal pwszDefExt As WString Ptr, ByVal id As Long = 0, ByVal sigdnName As SIGDN = SIGDN_FILESYSPATH ) As WString Ptr
declare Function FF_Toolbar_EnableButton (ByVal hToolBar As HWnd, ByVal idButton As Long) As BOOLEAN
declare Function FF_Toolbar_DisableButton (ByVal hToolBar As HWnd, ByVal idButton As Long) As BOOLEAN
Declare Function FF_ListView_InsertItem( ByVal hWndControl As HWnd, ByVal iRow As Long, ByVal iColumn As Long, ByVal pwszText As WString Ptr, ByVal lParam As LPARAM = 0 ) As BOOLEAN
Declare Function FF_ListView_GetItemText( ByVal hWndControl As HWnd, ByVal iRow As Long, ByVal iColumn As Long, ByVal pwszText As WString Ptr, ByVal nTextMax As Long ) As BOOLEAN
Declare Function FF_ListView_SetItemText( ByVal hWndControl As HWnd, ByVal iRow As Long, ByVal iColumn As Long, ByVal pwszText As WString Ptr, ByVal nTextMax As Long ) As Long
Declare Function FF_ListView_GetlParam( ByVal hWndControl As HWnd, ByVal iRow As Long ) As LPARAM
Declare Function FF_ListView_SetlParam( ByVal hWndControl As HWnd, ByVal iRow As Long, ByVal ilParam As LPARAM ) As long
Declare Function LoadLocalizationFile( Byref wszFileName As CWSTR, byval IsEnglish as boolean = false ) As BOOLEAN
Declare Function GetProcessImageName( ByVal pe32w As PROCESSENTRY32W Ptr, ByVal pwszExeName As WString Ptr ) As Long
Declare Function IsProcessRunning( ByVal pwszExeFileName As WString Ptr ) As BOOLEAN
Declare Function GetRunExecutableFilename() as CWSTR
Declare Function LoadPNGfromRes(BYVAL hInstance AS HINSTANCE, BYREF wszImageName AS WSTRING) as any ptr
Declare Function ParseDocument( byval pDoc as clsDocument ptr, byval wszFilename as CWSTR ) As Long
declare function DoCheckForUpdates( byval hWndParent as hwnd, byval bSilentCheck as Boolean = false ) as long












