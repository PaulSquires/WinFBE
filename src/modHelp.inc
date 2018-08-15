'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016 Paul Squires, PlanetSquires Software
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


' ========================================================================================
' Show context help or general contents if word does not exist
' ========================================================================================
Function ShowContextHelp( byval id as long ) As Long

   Dim HtmlHelpW As Function ( BYVAL hwndCaller AS HWnd, _
                               BYVAL pswzFile AS WSTRING Ptr, _
                               BYVAL uCommand AS UNIT, _
                               BYVAL dwData AS DWORD_PTR _
                               ) AS HWND

   dim as CWSTR wszHelpFilename, wszLabel
   
   ' Ensure that the CurDrive parameter is converted if applicable
   gConfig.CompilerHelpFile = ProcessFromCurdrive( gConfig.CompilerHelpFile ) 
   gConfig.Win32APIHelpfile = ProcessFromCurdrive( gConfig.Win32APIHelpfile ) 
   gConfig.WinFBXHelpfile   = ProcessFromCurdrive( gConfig.WinFBXHelpfile ) 

   select Case id
      case IDM_HELP
         wszLabel = "FreeBASIC "
         ' Convert relative path to absolute path if needed.
         if AfxPathIsRelative(gConfig.CompilerHelpFile) then
            wszHelpFilename = AfxPathCombine(AfxGetExePathName, gConfig.CompilerHelpFile)
         else
            wszHelpFilename = gConfig.CompilerHelpFile
         END IF
      case IDM_HELPWINAPI
         wszLabel = "WinAPI "
         ' Convert relative path to absolute path if needed.
         if AfxPathIsRelative(gConfig.Win32APIHelpfile) then
            wszHelpFilename = AfxPathCombine(AfxGetExePathName, gConfig.Win32APIHelpfile)
         else   
            wszHelpFilename = gConfig.Win32APIHelpfile 
         END IF
      case IDM_HELPWINFBX
         wszLabel = "WinFBX "
         ' Convert relative path to absolute path if needed.
         if AfxPathIsRelative(gConfig.WinFBXHelpfile) then
            wszHelpFilename = AfxPathCombine(AfxGetExePathName, gConfig.WinFBXHelpfile)
         else   
            wszHelpFilename = gConfig.WinFBXHelpfile
         END IF
   END SELECT


   If AfxFileExists(wszHelpFilename) = 0 Then
      MessageBoxW( HWND_FRMMAIN, wszLabel & L(244,"Help file not found."), L(201,"Error"), _
                    MB_OK Or MB_ICONWARNING Or MB_DEFBUTTON1 Or MB_APPLMODAL )
      Exit Function
   End If


   HtmlHelpW = DyLibSymbol( gpHelpLib, "HtmlHelpW" )

   If (gpHelpLib = 0) OrElse (HtmlHelpW = 0) Then
      MessageBox( HWND_FRMMAIN, L(243,"Error loading HtmlHelp."), L(201,"Error"), _
                    MB_OK Or MB_ICONWARNING Or MB_DEFBUTTON1 Or MB_APPLMODAL )
      Exit Function
   End If

   ' If we are currently in an active document then attempt to lookup the 
   ' word immediately under the caret.
   Dim pDoc as clsDocument Ptr = gTTabCtl.GetActiveDocumentPtr()
   
   If pDoc = 0 Then
      HtmlHelpW( 0, wszHelpFilename, HH_DISPLAY_TOC, NULL )
      Exit Function
   ENd If

   Dim wszKeyword as WString * MAX_PATH
   
   wszKeyword = WStr( pDoc->GetWord )
   
   Dim li As HH_AKLINK
   With li
      .cbStruct     = SizeOf(HH_AKLINK) 
      .fReserved    = FALSE 
      .pszKeywords  = @wszKeyword
      .pszUrl       = Null
      .pszMsgText   = Null
      .pszMsgTitle  = Null
      .pszWindow    = Null
      .fIndexOnFail = FALSE 
   End With
   
   ' Open the help and show the topic
   HtmlHelpW( 0, wszHelpFilename, HH_DISPLAY_TOC, Null )   '<-- needed?
   
   If HtmlHelpW( 0, wszHelpFilename, HH_KEYWORD_LOOKUP, Cast(DWORD_PTR, @li) ) = 0 Then
      ' Normal case search failed, try a ucase search
      wszKeyword     = UCase(wszKeyword)
      li.pszKeywords = @wszKeyword
      HtmlHelpW( 0, wszHelpFilename, HH_KEYWORD_LOOKUP, Cast(DWORD_PTR, @li) )
   End If

   Function = 0
   
End Function















