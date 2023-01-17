'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2023 Paul Squires, PlanetSquires Software
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

#Define IDC_FRMHELPVIEWER_WEBBROWSER_WINFBE         1000
#Define IDC_FRMHELPVIEWER_WEBBROWSER_WINFBX         1001
#Define IDC_FRMHELPVIEWER_TABCONTROL                1002
#Define IDC_FRMHELPVIEWER_TVWINFBE                  1003
#Define IDC_FRMHELPVIEWER_TVWINFBX                  1004
#Define IDC_FRMHELPVIEWER_BACK                      1005 
#Define IDC_FRMHELPVIEWER_FORWARD                   1006 
#Define IDC_FRMHELPVIEWER_PRINT                     1007 
#Define IDC_FRMHELPVIEWER_FIND                      1008 

' Size = 32 bytes
TYPE HH_AKLINK 
   cbStruct     AS LONG         ' int       cbStruct;     // sizeof this structure
   fReserved    AS BOOLEAN      ' BOOL      fReserved;    // must be FALSE (really!)
   pszKeywords  AS WSTRING PTR  ' LPCTSTR   pszKeywords;  // semi-colon separated keywords
   pszUrl       AS WSTRING PTR  ' LPCTSTR   pszUrl;       // URL to jump to if no keywords found (may be NULL)
   pszMsgText   AS WSTRING PTR  ' LPCTSTR   pszMsgText;   // Message text to display in MessageBox if pszUrl is NULL and no keyword match
   pszMsgTitle  AS WSTRING PTR  ' LPCTSTR   pszMsgTitle;  // Message text to display in MessageBox if pszUrl is NULL and no keyword match
   pszWindow    AS WSTRING PTR  ' LPCTSTR   pszWindow;    // Window to display URL in
   fIndexOnFail AS BOOLEAN      ' BOOL      fIndexOnFail; // Displays index if keyword lookup fails.
END TYPE

#Define HH_DISPLAY_TOPIC   0000 
#Define HH_DISPLAY_TOC     0001
#Define HH_KEYWORD_LOOKUP  0013
#Define HH_HELP_CONTEXT    0015


'  Global holding all full path/name for HTML files linked to Help Treeview (index in lParam)
type HTMLHELPNODES
   wszFilename    as CWSTR
   wszLocationURL as CWSTR
   TreeviewNode   as HTREEITEM
   hTreeview      as HWND
end type
dim shared as HTMLHELPNODES gHTMLHelp(any)

Dim Shared As Any Ptr gpHelpLib

declare Function ShowContextHelp( byval id as long ) As Long
declare Function frmHelpViewer_Show( ByVal hWndParent As HWnd, byval idmHelpMessage as long ) As LRESULT
