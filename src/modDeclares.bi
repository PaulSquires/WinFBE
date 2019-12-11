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


''  Menu message identifiers
Enum

   '' USER MESSAGES
   MSG_USER_SETFOCUS = WM_USER + 1
   MSG_USER_SHOWCOLORCOMBOBOXES
   MSG_USER_SETCOLORCUSTOM
   MSG_USER_GETCOLORCUSTOM
   MSG_USER_PROCESS_COMMANDLINE 
   MSG_USER_PROCESS_UPDATECHECK
   MSG_USER_SHOWAUTOCOMPLETE
   MSG_USER_APPENDEQUALSSIGN
   MSG_USER_GENERATECODE
   
   '' THEMES
   IDM_CREATE_THEME, IDM_IMPORT_THEME, IDM_DELETE_THEME
   
   '' FILE
   IDM_FILE, IDM_FILENEW, IDM_FILEOPEN, IDM_FILEOPENTEMPLATES, IDM_FILECLOSE
   IDM_FILESAVE, IDM_FILESAVEAS, IDM_FILESAVEALL, IDM_FILESAVEDECLARES
   IDM_FILEOPEN_EXPLORERTREEVIEW
   IDM_FILESAVE_EXPLORERTREEVIEW
   IDM_FILESAVEAS_EXPLORERTREEVIEW
   IDM_FILECLOSE_EXPLORERTREEVIEW
   IDM_FILECLOSEALL, IDM_FILECLOSEALLOTHERS, IDM_CLOSEALLFORWARD, IDM_CLOSEALLBACKWARD
   IDM_MRU, IDM_OPENINCLUDE, IDM_COMMAND, IDM_EXIT
   
   '' EDIT
   IDM_EDIT, IDM_UNDO, IDM_REDO
   IDM_CUT, IDM_COPY, IDM_PASTE, IDM_DELETELINE, IDM_DELETE, IDM_INSERTFILE
   IDM_ANSI, IDM_UTF8BOM, IDM_UTF16BOM
   IDM_INDENTBLOCK, IDM_UNINDENTBLOCK, IDM_COMMENTBLOCK, IDM_UNCOMMENTBLOCK
   IDM_DUPLICATELINE, IDM_MOVELINEUP, IDM_MOVELINEDOWN, IDM_TOUPPERCASE, IDM_TOLOWERCASE
   IDM_TOMIXEDCASE, IDM_EOLTOCRLF, IDM_EOLTOCR, IDM_EOLTOLF, IDM_SELECTALL, IDM_SELECTLINE
   IDM_SPACESTOTABS, IDM_TABSTOSPACES
   
   '' SEARCH
   IDM_SEARCH
   IDM_FIND, IDM_FINDNEXT, IDM_FINDPREV, IDM_FINDNEXTACCEL, IDM_FINDPREVACCEL 
   IDM_FINDINFILES, IDM_REPLACE, IDM_DEFINITION,IDM_LASTPOSITION
   IDM_GOTONEXTTAB, IDM_GOTOPREVTAB, IDM_CLOSETAB, IDM_GOTONEXTFUNCTION, IDM_GOTOPREVFUNCTION
   IDM_GOTOHEADERFILE, IDM_GOTOSOURCEFILE, IDM_GOTOMAINFILE, IDM_GOTORESOURCEFILE
   IDM_GOTO, IDM_BOOKMARKTOGGLE, IDM_BOOKMARKNEXT, IDM_BOOKMARKPREV, IDM_BOOKMARKCLEARALL
   
   '' VIEW
   IDM_VIEW
   IDM_FOLDTOGGLE, IDM_FOLDBELOW, IDM_FOLDALL, IDM_UNFOLDALL, IDM_ZOOMIN, IDM_ZOOMOUT
   IDM_FUNCTIONLIST, IDM_VIEWEXPLORER, IDM_VIEWOUTPUT, IDM_RESTOREMAIN
   
   '' PROJECT
   IDM_PROJECTNEW, IDM_PROJECTMANAGER, IDM_PROJECTOPEN, IDM_MRUPROJECT
   IDM_PROJECTCLOSE, IDM_PROJECTSAVE, IDM_PROJECTSAVEAS, IDM_PROJECTFILESADD, IDM_PROJECTOPTIONS  
   IDM_REMOVEFILEFROMPROJECT 
   IDM_REMOVEFILEFROMPROJECT_EXPLORERTREEVIEW
   
   '' COMPILE
   IDM_BUILDEXECUTE, IDM_COMPILE, IDM_REBUILDALL, IDM_RUNEXE, IDM_QUICKRUN, IDM_COMMANDLINE
   
   ''  OPTIONS
   IDM_OPTIONS, IDM_BUILDCONFIG, IDM_USERTOOLSDIALOG, IDM_USERSNIPPETS

   '' DESIGNER
   IDM_NEWFORM, IDM_VIEWTOOLBOX, IDM_TOGGLEVIEWCODE, IDM_MENUEDITOR
   IDM_TOOLBAREDITOR, IDM_STATUSBAREDITOR, IDM_IMAGEMANAGER
   IDM_ALIGNLEFTS, IDM_ALIGNCENTERS, IDM_ALIGNRIGHTS
   IDM_ALIGNTOPS, IDM_ALIGNMIDDLES, IDM_ALIGNBOTTOMS
   IDM_HORIZEQUAL, IDM_HORIZINCREASE, IDM_HORIZDECREASE, IDM_HORIZREMOVE
   IDM_VERTEQUAL, IDM_VERTINCREASE, IDM_VERTDECREASE, IDM_VERTREMOVE
   IDM_SAMEWIDTHS, IDM_SAMEHEIGHTS, IDM_SAMEBOTH 
   IDM_CENTERHORIZ, IDM_CENTERVERT, IDM_CENTERBOTH
   IDM_SNAPLINES, IDM_LOCKCONTROLS

   '' HELP
   IDM_HELP, IDM_HELPWINAPI, IDM_HELPWINFBE, IDM_HELPWINFBX
   IDM_HELPTIPS, IDM_CHECKFORUPDATES, IDM_ABOUT
   
   '' OTHER
   IDM_SETFILEMAIN
   IDM_SETFILERESOURCE
   IDM_SETFILEHEADER
   IDM_SETFILEMODULE
   IDM_SETFILENORMAL
   IDM_SETFILEMAIN_EXPLORERTREEVIEW
   IDM_SETFILERESOURCE_EXPLORERTREEVIEW
   IDM_SETFILEHEADER_EXPLORERTREEVIEW
   IDM_SETFILEMODULE_EXPLORERTREEVIEW
   IDM_SETFILENORMAL_EXPLORERTREEVIEW
   
   IDM_MRUCLEAR, IDM_MRUPROJECTCLEAR
   IDM_CONSOLE, IDM_GUI, IDM_RESOURCE   ' used for compiler directives in code
   IDM_ADDIMAGE, IDM_REMOVEIMAGE, IDM_FORMATIMAGE, IDM_ATTACHIMAGE, IDM_DETACHIMAGE
   IDM_32BIT, IDM_64BIT   ' mainly used for identifying compiler associated with a project
   IDM_USERTOOL   ' + n number of user tools
End Enum


'  Global window handle for the main form
Dim Shared As HWnd HWND_FRMMAIN, HWND_FRMMAIN_TOOLBAR, HWND_FRMEXPLORER, HWND_FRMRECENT, HWND_FRMOUTPUT
dim shared as hwnd HWND_FRMMAIN_COMBOBUILDS 
Dim Shared As HMENU HWND_FRMMAIN_TOPMENU   

'  Global window handles for some forms 
Dim Shared As HWnd HWND_FRMOPTIONS, HWND_FRMOPTIONSGENERAL, HWND_FRMOPTIONSEDITOR, HWND_FRMOPTIONSCOLORS
Dim Shared As HWnd HWND_FRMOPTIONSCOMPILER, HWND_FRMOPTIONSLOCAL, HWND_FRMOPTIONSKEYWORDS
Dim Shared As HWnd HWND_FRMFINDREPLACE, HWND_FRMFINDINFILES, HWND_FRMVDTOOLBOX, HWND_FRMVDCOLORS
Dim Shared As HWnd HWND_FRMFUNCTIONLIST, HWND_FRMBUILDCONFIG, HWND_FRMMENUEDITOR, HWND_FRMUSERTOOLS
Dim Shared As HWnd HWND_PROPLIST_EDIT, HWND_PROPLIST_COMBO, HWND_PROPLIST_COMBOLIST, HWND_FRMHELPVIEWER
dim shared as hwnd HWND_FRMIMAGES, HWND_FRMSNIPPETS, HWND_FRMSTATUSBAREDITOR, HWND_FRMTOOLBAREDITOR
dim shared as hwnd HWND_FRMVDTABCHILD

dim shared as HICON ghIconGood, ghIconBad, ghIconTick, ghIconNoTick

' Create a dynamic array that will hold all localization words/phrases while
' a language is being edited in frmOptionsLocal. Also create a global array
' that holds the english phrases. When a localization is loaded, any missing
' translations are replaced with the english version.
ReDim Shared gLangEnglish(Any) As WString * MAX_PATH
ReDim Shared gLocalPhrases(Any) As WString * MAX_PATH
common shared gLocalPhrasesEdit as boolean   ' a localization language is being edited. 


' Create a dynamic array that will hold all localization words/phrases. This
' array is resized and loaded using the LoadLocalizationFile function.
ReDim Shared LL(Any) As WString * MAX_PATH

' Define a macro that allows the user to specify the LL array subscript and
' also a descriptive label (that is ignored), and return the LL array value.
#Define L(e,s) LL(e)

#Define SetFocusScintilla  PostMessage( HWND_FRMMAIN, MSG_USER_SETFOCUS, 0, 0 )
#Define SciExec(h, m, w, l) SendMessage(h, m, w, CAST(LPARAM, l))


''
''  Save information related to Find/Replace and Find in Files operations
''
Type FINDREPLACE_TYPE
   foundCount          as long 
   txtFind             As String
   txtReplace          As String
   txtFindCombo(10)    As String
   txtFilesCombo(10)   As String
   txtFolderCombo(10)  As String
   txtLastFind         As String
   txtFiles            As String        ' *.*, *.bas, etc (FindInFolder)
   txtFolder           As String        ' start search from this folder (FindInFolder)
   nSearchSubFolders   As Long          ' search sub folders as well (FindInFolder)
   nWholeWord          As long          ' find/replace whole word search
   nMatchCase          As long          ' match case when searching
   nSelection          As long          ' search only selected text
   bShowReplace        as Boolean
End Type
dim Shared gFind As FINDREPLACE_TYPE
dim Shared gFindInFiles As FINDREPLACE_TYPE


' Tools/controls that can be drawn on a Form.
type TOOLBOX_TYPE
   nToolType       as long 
   wszToolBoxName  as CWSTR    ' eg. OptionButton
   wszControlName  as CWSTR    ' eg. Option
   wszImage        as CWSTR
   wszCursor       as CWSTR
   wszClassName    as CWSTR    ' eg. RADIOBUTTON
END TYPE
dim shared gToolBox(any) as TOOLBOX_TYPE






