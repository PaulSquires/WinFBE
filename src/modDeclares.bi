'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2017 Paul Squires, PlanetSquires Software
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


''  Control identifiers
#Define IDC_SCINTILLA 100

#Define IDC_FRMMAIN_TOPTABCONTROL                   1000
#Define IDC_FRMMAIN_TOOLBAR                         1001
#Define IDC_FRMMAIN_REBAR                           1002
#Define IDC_FRMMAIN_STATUSBAR                       1003
#Define IDC_FRMMAIN_PROGRESSBAR                     1004
#Define IDC_FRMMAIN_COMPILETIMER                    1005

#Define IDC_FRMOUTPUT_TABCONTROL                    1000
#Define IDC_FRMOUTPUT_LISTVIEW                      1001
#Define IDC_FRMOUTPUT_TXTLOGFILE                    1002
#Define IDC_FRMOUTPUT_BTNCLOSE                      1003
#Define IDC_FRMOUTPUT_LISTSEARCH                    1004
#Define IDC_FRMOUTPUT_LVTODO                        1005
#Define IDC_FRMOUTPUT_TXTNOTES                      1006

#Define IDC_FRMEXPLORER_LBLEXPLORER                 1000
#Define IDC_FRMEXPLORER_TREE                        1001
#Define IDC_FRMEXPLORER_BTNCLOSE                    1002

#Define IDC_FRMRECENT_BTNOPENFILE                   1000
#Define IDC_FRMRECENT_BTNOPENPROJECT                1001
#Define IDC_FRMRECENT_TREEVIEW                      1002

#Define IDC_FRMOPTIONS_LABEL1                       1000
#Define IDC_FRMOPTIONS_CMDCANCEL                    1001
#Define IDC_FRMOPTIONS_LBLCATEGORY                  1002
#Define IDC_FRMOPTIONS_CMDOK                        1003
#Define IDC_FRMOPTIONS_TVWCATEGORIES                1004

#Define IDC_FRMOPTIONSGENERAL_CHKMULTIPLEINSTANCES  1000
#Define IDC_FRMOPTIONSGENERAL_CHKCOMPILEAUTOSAVE    1001
#Define IDC_FRMOPTIONSGENERAL_CHKCLOSEFUNCLIST      1002
#Define IDC_FRMOPTIONSGENERAL_CHKCLOSEPROJMGR       1003
#Define IDC_FRMOPTIONSGENERAL_CHKHIDECOMPILE        1004

#Define IDC_FRMOPTIONSEDITOR_LBLTABSIZE             1000
#Define IDC_FRMOPTIONSEDITOR_TXTTABSIZE             1001
#Define IDC_FRMOPTIONSEDITOR_LBLKEYWORDCASE         1002
#Define IDC_FRMOPTIONSEDITOR_COMBOCASE              1003
#Define IDC_FRMOPTIONSEDITOR_CHKCODETIPS            1004
#Define IDC_FRMOPTIONSEDITOR_CHKAUTOCOMPLETE        1005
#Define IDC_FRMOPTIONSEDITOR_CHKSHOWLEFTMARGIN      1006
#Define IDC_FRMOPTIONSEDITOR_CHKSYNTAXHIGHLIGHTING  1007
#Define IDC_FRMOPTIONSEDITOR_CHKCURRENTLINE         1008
#Define IDC_FRMOPTIONSEDITOR_CHKLINENUMBERING       1009
#Define IDC_FRMOPTIONSEDITOR_CHKCONFINECARET        1010
#Define IDC_FRMOPTIONSEDITOR_CHKTABTOSPACES         1011
#Define IDC_FRMOPTIONSEDITOR_CHKAUTOINDENTATION     1012
#Define IDC_FRMOPTIONSEDITOR_CHKSHOWFOLDMARGIN      1013
#Define IDC_FRMOPTIONSEDITOR_CHKINDENTGUIDES        1014
#Define IDC_FRMOPTIONSEDITOR_CHKUNICODE             1015

#Define IDC_FRMOPTIONSCOLORS_LSTCOLORS              1000
#Define IDC_FRMOPTIONSCOLORS_FRMCOLORS              1001
#Define IDC_FRMOPTIONSCOLORS_LBLFOREGROUND          1002
#Define IDC_FRMOPTIONSCOLORS_LBLBACKGROUND          1003
#Define IDC_FRMOPTIONSCOLORS_FRMFONT                1004
#Define IDC_FRMOPTIONSCOLORS_CBCOLOR1               1005
#Define IDC_FRMOPTIONSCOLORS_CBCOLOR2               1006
#Define IDC_FRMOPTIONSCOLORS_COMBOFONTNAME          1007
#Define IDC_FRMOPTIONSCOLORS_COMBOFONTCHARSET       1008
#Define IDC_FRMOPTIONSCOLORS_COMBOFONTSIZE          1009

#Define IDC_FRMOPTIONSCOMPILER_CMDFBWIN32           1000
#Define IDC_FRMOPTIONSCOMPILER_LBLSWITCHES          1001
#Define IDC_FRMOPTIONSCOMPILER_LBLFBC32             1002
#Define IDC_FRMOPTIONSCOMPILER_TXTFBWIN32           1003
#Define IDC_FRMOPTIONSCOMPILER_TXTFBSWITCHES        1004
#Define IDC_FRMOPTIONSCOMPILER_LBLFBHELP            1005
#Define IDC_FRMOPTIONSCOMPILER_TXTFBHELPFILE        1006
#Define IDC_FRMOPTIONSCOMPILER_CMDFBHELPFILE        1007
#Define IDC_FRMOPTIONSCOMPILER_CMDFBWIN64           1008
#Define IDC_FRMOPTIONSCOMPILER_LBLFBC64             1009
#Define IDC_FRMOPTIONSCOMPILER_TXTFBWIN64           1010
#Define IDC_FRMOPTIONSCOMPILER_CMDAPIHELPPATH       1011
#Define IDC_FRMOPTIONSCOMPILER_LBLAPIHELP           1012
#Define IDC_FRMOPTIONSCOMPILER_TXTWIN32HELPPATH     1013

#Define IDC_FRMOPTIONSLOCAL_LBLLOCALIZATION         1000
#Define IDC_FRMOPTIONSLOCAL_CMDLOCALIZATION         1001
#Define IDC_FRMOPTIONSLOCAL_FRAMELOCALIZATION       1002

#Define IDC_FRMOPTIONSKEYWORDS_TXTKEYWORDS          1000

#Define IDC_FRMPROJECTOPTIONS_TXTPROJECTPATH        1000
#Define IDC_FRMPROJECTOPTIONS_CMDSELECT             1001
#Define IDC_FRMPROJECTOPTIONS_OPTEXE                1002
#Define IDC_FRMPROJECTOPTIONS_OPTDLL                1003
#Define IDC_FRMPROJECTOPTIONS_OPTLIB                1004
#Define IDC_FRMPROJECTOPTIONS_OPTNOERROR            1005
#Define IDC_FRMPROJECTOPTIONS_OPTERROR              1006
#Define IDC_FRMPROJECTOPTIONS_OPTRESUME             1007
#Define IDC_FRMPROJECTOPTIONS_OPTERRORPLUS          1008
#Define IDC_FRMPROJECTOPTIONS_FRAME1                1009
#Define IDC_FRMPROJECTOPTIONS_FRAME2                1010
#Define IDC_FRMPROJECTOPTIONS_CHKDEBUG              1011
#Define IDC_FRMPROJECTOPTIONS_CHKTHREAD             1012
#Define IDC_FRMPROJECTOPTIONS_LABEL1                1013
#Define IDC_FRMPROJECTOPTIONS_TXTOPTIONS32          1014
#Define IDC_FRMPROJECTOPTIONS_LABEL2                1015
#Define IDC_FRMPROJECTOPTIONS_LABEL3                1016
#Define IDC_FRMPROJECTOPTIONS_LABEL4                1017
#Define IDC_FRMPROJECTOPTIONS_TXTOPTIONS64          1018
#Define IDC_FRMPROJECTOPTIONS_CHKSHOWCONSOLE        1019

#Define IDC_FRMTEMPLATES_LISTBOX                    1000

#Define IDC_FRMFIND_LBLFINDWHAT                     1000
#Define IDC_FRMFIND_COMBOFIND                       1001
#Define IDC_FRMFIND_CHKWHOLEWORDS                   1004
#Define IDC_FRMFIND_CHKMATCHCASE                    1005
#Define IDC_FRMFIND_FRAMESCOPE                      1006
#Define IDC_FRMFIND_OPTSCOPE1                       1007
#Define IDC_FRMFIND_OPTSCOPE2                       1008
#Define IDC_FRMFIND_OPTSCOPE3                       1009
#Define IDC_FRMFIND_FRAMEOPTIONS                    1010

#Define IDC_FRMFINDINFILES_LBLFINDWHAT              1000
#Define IDC_FRMFINDINFILES_COMBOFIND                1001
#Define IDC_FRMFINDINFILES_COMBOFILES               1002
#Define IDC_FRMFINDINFILES_COMBOFOLDER              1003
#Define IDC_FRMFINDINFILES_CHKWHOLEWORDS            1004
#Define IDC_FRMFINDINFILES_CHKMATCHCASE             1005
#Define IDC_FRMFINDINFILES_FRAMESCOPE               1006
#Define IDC_FRMFINDINFILES_OPTSCOPE1                1007
#Define IDC_FRMFINDINFILES_OPTSCOPE2                1008
#Define IDC_FRMFINDINFILES_OPTSCOPE3                1009
#Define IDC_FRMFINDINFILES_CHKSUBFOLDERS            1010
#Define IDC_FRMFINDINFILES_FRAMEOPTIONS             1011
#Define IDC_FRMFINDINFILES_CMDFILES                 1012
#Define IDC_FRMFINDINFILES_CMDFOLDERS               1013
 
#Define IDC_FRMREPLACE_LBLFINDWHAT                  1000
#Define IDC_FRMREPLACE_COMBOFIND                    1001
#Define IDC_FRMREPLACE_CHKWHOLEWORDS                1002
#Define IDC_FRMREPLACE_CHKMATCHCASE                 1003
#Define IDC_FRMREPLACE_FRAMESCOPE                   1004
#Define IDC_FRMREPLACE_OPTSCOPE1                    1005
#Define IDC_FRMREPLACE_OPTSCOPE2                    1006
#Define IDC_FRMREPLACE_OPTSCOPE3                    1007
#Define IDC_FRMREPLACE_FRAMEOPTIONS                 1008
#Define IDC_FRMREPLACE_LBLREPLACEWITH               1009
#Define IDC_FRMREPLACE_COMBOREPLACE                 1010
#Define IDC_FRMREPLACE_CMDREPLACE                   1011
#Define IDC_FRMREPLACE_CMDREPLACEALL                1012
#Define IDC_FRMREPLACE_LBLSTATUS                    1013

#Define IDC_FRMFINDREPLACE_BTNCLOSE                 IDCANCEL
#Define IDC_FRMFINDREPLACE_BTNTOGGLE                1001
#Define IDC_FRMFINDREPLACE_TXTFIND                  1002
#Define IDC_FRMFINDREPLACE_TXTREPLACE               1003
#Define IDC_FRMFINDREPLACE_BTNMATCHCASE             1004
#Define IDC_FRMFINDREPLACE_BTNMATCHWHOLEWORD        1005
#Define IDC_FRMFINDREPLACE_BTNREPLACE               1006
#Define IDC_FRMFINDREPLACE_BTNREPLACEALL            1007
#Define IDC_FRMFINDREPLACE_LBLRESULTS               1008
#Define IDC_FRMFINDREPLACE_BTNLEFT                  1009
#Define IDC_FRMFINDREPLACE_BTNRIGHT                 1010
#Define IDC_FRMFINDREPLACE_BTNSELECTION             1011

#Define IDC_FRMGOTO_LBLLASTLINE                     1000
#Define IDC_FRMGOTO_LBLCURRENTLINE                  1001
#Define IDC_FRMGOTO_LBLGOTOLINE                     1002
#Define IDC_FRMGOTO_TXTLINE                         1003
#Define IDC_FRMGOTO_LBLLASTVALUE                    1004
#Define IDC_FRMGOTO_LBLCURRENTVALUE                 1005

#Define IDC_FRMCOMMANDLINE_LABEL1                   1000
#Define IDC_FRMCOMMANDLINE_TEXTBOX1                 1001

#Define IDC_FRMFNLIST_LISTBOX                       1000


Const DELIM = "|"                    ' character used as delimiter for function names in data1 of gFunctionLists hash
Const IDC_MRUBASE = 5000             ' Windows id of MRU items 1 to 10 (located under File menu)
Const IDC_MRUPROJECTBASE = 6000      ' Windows id of MRUPROJECT items 1 to 10 (located under Project menu)
CONST IDM_ADDFILETOPROJECT = 6100    ' 6100 to 6199 Popup menu will show list of open projects to choose from. 

Const FILETYPE_UNDEFINED = 0
Const FILETYPE_MAIN      = 1
Const FILETYPE_MODULE    = 2
Const FILETYPE_NORMAL    = 3
Const FILETYPE_RESOURCE  = 4
 
' File encodings
const FILE_ENCODING_ANSI      = 0
const FILE_ENCODING_UTF8_BOM  = 1
const FILE_ENCODING_UTF16_BOM = 2

   
''  Menu message identifiers
Enum
''  Custom messages
   MSG_USER_SETFOCUS = WM_USER + 1
   MSG_USER_PROCESS_COMMANDLINE 
   MSG_USER_TOGGLE_TVCHECKBOXES
   IDM_FILE, IDM_FILENEW 
   IDM_FILEOPEN, IDM_FILECLOSE, IDM_FILECLOSEALL, IDM_FILESAVE, IDM_FILESAVEAS
   IDM_FILESAVEALL, IDM_FILESAVEDECLARES
   IDM_MRU, IDM_OPENINCLUDE, IDM_COMMAND, IDM_EXIT
   IDM_EDIT
   IDM_UNDO, IDM_REDO, IDM_CUT, IDM_COPY, IDM_PASTE, IDM_DELETELINE, IDM_INSERTFILE
   IDM_ANSI, IDM_UTF8BOM, IDM_UTF16BOM
   IDM_INDENTBLOCK, IDM_UNINDENTBLOCK, IDM_COMMENTBLOCK, IDM_UNCOMMENTBLOCK
   IDM_DUPLICATELINE, IDM_MOVELINEUP, IDM_MOVELINEDOWN, IDM_TOUPPERCASE, IDM_TOLOWERCASE
   IDM_TOMIXEDCASE, IDM_EOLTOCRLF, IDM_EOLTOCR, IDM_EOLTOLF, IDM_SELECTALL, IDM_SELECTLINE
   IDM_SPACESTOTABS, IDM_TABSTOSPACES
   IDM_SEARCH
   IDM_FIND, IDM_FINDNEXT, IDM_FINDPREV, IDM_FINDINFILES, IDM_REPLACE, IDM_DEFINITION
   IDM_LASTPOSITION, IDM_FUNCTIONLIST
   IDM_GOTO, IDM_BOOKMARKTOGGLE, IDM_BOOKMARKNEXT, IDM_BOOKMARKPREV, IDM_BOOKMARKCLEARALL
   IDM_VIEW
   IDM_FOLDTOGGLE, IDM_FOLDBELOW, IDM_FOLDALL, IDM_UNFOLDALL, IDM_ZOOMIN, IDM_ZOOMOUT, IDM_RESTOREMAIN
   IDM_VIEWEXPLORER, IDM_VIEWOUTPUT
   IDM_OPTIONS
   IDM_PROJECTNEW, IDM_PROJECTMANAGER, IDM_PROJECTOPEN, IDM_MRUPROJECT
   IDM_PROJECTFILESADDTONODE, IDM_REMOVEFILEFROMPROJECT 
   IDM_PROJECTCLOSE, IDM_PROJECTSAVE, IDM_PROJECTSAVEAS, IDM_PROJECTFILESADD, IDM_PROJECTOPTIONS  
   IDM_BUILDEXECUTE, IDM_COMPILE, IDM_REBUILDALL, IDM_RUNEXE, IDM_COMMANDLINE
   IDM_USE32BIT, IDM_USE64BIT
   IDM_GUI, IDM_CONSOLE
   IDM_HELP, IDM_ABOUT
   IDM_SETFILENORMAL, IDM_SETFILEMODULE, IDM_SETFILEMAIN, IDM_SETFILERESOURCE
   IDM_MRUCLEAR, IDM_MRUPROJECTCLEAR, IDM_NEXTTAB, IDM_PREVTAB, IDM_CLOSETAB

End Enum



'  Global window handle for the main form
Dim Shared As HWnd HWND_FRMMAIN, HWND_FRMMAIN_TOOLBAR, HWND_FRMEXPLORER, HWND_FRMRECENT, HWND_FRMOUTPUT
Dim Shared As HMENU HWND_FRMMAIN_TOPMENU   

Dim Shared As HIMAGELIST ghImageListNormal
Dim Shared As Long gidxImageOpened, gidxImageClosed, gidxImageBlank, gidxImageCode
dim shared as BOOLEAN gProjectLoading  ' T/F to prevent screen flickering/updates during loading of many files.
dim shared as BOOLEAN gCompiling       ' T/F to show spinning mouse cursor.

'  Global window handles for some forms 
Dim Shared As HWnd HWND_FRMOPTIONS, HWND_FRMOPTIONSGENERAL, HWND_FRMOPTIONSEDITOR, HWND_FRMOPTIONSCOLORS
Dim Shared As HWnd HWND_FRMOPTIONSCOMPILER, HWND_FRMOPTIONSLOCAL, HWND_FRMOPTIONSKEYWORDS
Dim Shared As HWnd HWND_FRMFINDREPLACE, HWND_FRMFINDINFILES
Dim Shared As HWnd HWND_FRMFNLIST

'  Global handle to hhctrl.ocx for context sensitive help
Dim Shared As Any Ptr gpHelpLib

dim shared as HICON ghIconGood, ghIconBad
dim shared as BOOLEAN gReplaceOpen     ' replace dialog is open

' Create a temporary array to hold the selected color values
' for the different editor elements. When the form is saved then 
' the values from this temporary struture is saved to the 
' configuration table.
Type TYPE_COLORS
   nFg As COLORREF
   nBg As COLORREF
End Type
Dim Shared gTempColors(15) As TYPE_COLORS   


' Create a dynamic array that will hold all localization words/phrases. This
' array is resized and loaded using the LoadLocalizationFile function.
ReDim Shared LL(Any) As WString * MAX_PATH

' Define a macro that allows the user to specify the LL array subscript and
' also a descriptive label (that is ignored), and return the LL array value.
#Define L(e,s)  LL(e)

#Define SetFocusScintilla  PostMessage HWND_FRMMAIN, MSG_USER_SETFOCUS, 0, 0
#Define ResizeExplorerWindow  PostMessage HWND_FRMEXPLORER, MSG_USER_OPENEDITORS_RESIZE, 0, 0
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
   nWholeWord          As Long          ' checkmark for find/replace whole word search
   nMatchCase          As Long          ' match case when searching
   nSearchSubFolders   As Long          ' search sub folders as well (FindInFolder)
End Type
Dim Shared gFind As FINDREPLACE_TYPE
Dim Shared gFindInFiles As FINDREPLACE_TYPE


' Structure that holds all of the user embedded compiler directives
' in the source code. Currently, only the main source file is searched
' for the '#CONSOLE ON|OFF directive but others can be added as needed.
type COMPILE_DIRECTIVES
   ConsoleFlag as Boolean              ' True:CONSOLE ON, False:CONSOLE OFF
END TYPE


' Forward reference
Type clsDocument_ As clsDocument

'' Last position in document. Used when "Last Position" menu option is selected.
Type LASTPOSITION_TYPE
   pDoc       As clsDocument_ Ptr
   nFirstLine As Long     ' first visible line on screen
   nPosition  As Long     ' Position in Scintilla document where caret is positioned
End Type
Dim Shared gLastPosition As LASTPOSITION_TYPE


Type clsDocument
   Private:
      m_pSci           As Any Ptr      
      
   Public:
      hWindow          As HWnd
      IsNewFlag        As BOOLEAN
      ProjectIndex     as long      ' array index into gApp.Projects
      ProjectFileType  As Long = FILETYPE_UNDEFINED
      DiskFilename     As WString * MAX_PATH
      DateFileTime     As FILETIME  
      hNodeExplorer    As HTREEITEM
      FileEncoding     as long       
      UserModified     as boolean  ' occurs when user manually changes encoding state so that document will be saved in the new format

      Declare Function CreateCodeWindow( ByVal hWndParent As HWnd, ByVal IsNewFile As BOOLEAN, ByVal IsTemplate As BOOLEAN = False, ByVal pwszFile As WString Ptr = 0) As HWnd
      Declare Function FindReplace( ByVal strFindText As String, ByVal strReplaceText As String ) As Long
      Declare Function InsertFile() As BOOLEAN
      Declare Function SaveFile(ByVal bSaveAs As BOOLEAN = False) As BOOLEAN
      Declare Function ApplyProperties() As Long
      Declare Function GetTextRange( ByVal cpMin As Long, ByVal cpMax As Long) As String
      Declare Function ChangeSelectionCase( ByVal fCase As Long) As Long 
      Declare Function GetCurrentLineNumber() As Long
      Declare Function SelectLine( ByVal nLineNum As Long ) As Long
      Declare Function GetLine( ByVal nLine As Long) As String
      Declare Function GetSelText() As String
      Declare Function GetText() As String
      Declare Function SetText( ByRef sText As Const String ) As Long 
      Declare Function GetSelectedLineRange( ByRef startLine As Long, ByRef endLine As Long, ByRef startPos As Long, ByRef endPos As Long ) As Long 
      Declare Function BlockComment( ByVal flagBlock As BOOLEAN ) As Long
      Declare Function CurrentLineUp() As Long
      Declare Function CurrentLineDown() As Long
      Declare Function MoveCurrentLines( ByVal flagMoveDown As BOOLEAN ) As Long
      Declare Function ToggleBookmark( ByVal nLine As Long ) As Long
      Declare Function NextBookmark() As Long 
      Declare Function PrevBookmark() As Long 
      Declare Function FoldToggle( ByVal nLine As Long ) As Long
      Declare Function FoldAll() As Long
      Declare Function UnFoldAll() As Long
      Declare Function FoldToggleOnwards( ByVal nLine As Long) As Long
      Declare Function ConvertEOL( ByVal nMode As Long) As Long
      Declare Function DisplayStats() As Long
      Declare Function TabsToSpaces() As Long
      Declare Function GetWord( ByVal curPos As Long = -1 ) As String
      Declare Function GetBookmarks() As String
      Declare Function SetBookmarks( ByVal sBookmarks As String ) As Long
      declare Function InFunction() As BOOLEAN
      declare Function LineDuplicate() As Long
      declare function SetMarkerHighlight() As Long
      declare Function RemoveMarkerHighlight() As Long
      declare Function IsMultiLineSelection() As boolean
      declare Function HasMarkerHighlight() As BOOLEAN
      declare Function FirstMarkerHighlight() As long
      declare Function LastMarkerHighlight() As long
      declare function CompileDirectives( byval pDirectives as COMPILE_DIRECTIVES ptr) as Long
      Declare Constructor
      Declare Destructor
End Type

Type clsTopTabCtl
   Private:
      
   Public:
      hWindow   As HWnd
      Declare Function AddTab( ByVal pDoc As clsDocument Ptr ) As Long
      Declare Function GetTabIndexFromFilename( ByVal pwszName As WString Ptr ) As Long
      declare Function GetTabIndexByDocumentPtr( ByVal pDocIn As clsDocument Ptr ) As Long
      Declare Function SetTabIndexByDocumentPtr( ByVal pDocIn As clsDocument Ptr ) As Long
      Declare Function SetFocusTab( ByVal idx As Long ) As Long
      Declare Function GetActiveDocumentPtr() As clsDocument Ptr
      Declare Function GetDocumentPtr( ByVal idx As Long ) As clsDocument Ptr
      Declare Function DisplayScintilla( ByVal idx As Long, ByVal bShow As BOOLEAN ) As Long
      Declare Function SetTabText( ByVal idx As Long ) As Long
      Declare Function NextTab() As Long
      Declare Function PrevTab() As Long
      Declare Function CloseTab() As Long
      
End Type


Type clsConfig
   Private:
      _ConfigFilename As String 
      
   Public:
      FBKeywords           As String 
      bKeywordsDirty       As BOOLEAN = True       ' not saved to file
      LastRunFilename      As CWSTR                ' not saved to file
      CloseFuncList        As Long = True
      ShowExplorer         As Long = True
      ShowExplorerWidth    As Long = 250
      SyntaxHighlighting   As Long = True
      Codetips             As Long = True
      AutoComplete         As Long = True
      LeftMargin           As Long = True
      FoldMargin           As Long = false
      AutoIndentation      As Long = True
      ConfineCaret         As Long = false
      LineNumbering        As Long = True
      HighlightCurrentLine As Long = True
      IndentGuides         As Long = True
      TabIndentSpaces      As Long = True
      HideCompile          As Long = False
      MultipleInstances    As Long = True
      CompileAutosave      As Long = True
      UnicodeEncoding      As Long = False
      TabSize              As CWSTR = "3"
      LocalizationFile     As CWSTR = "english.lang"
      EditorFontname       As CWSTR = "Courier New"
      EditorFontCharSet    As CWSTR = "Default"
      EditorFontsize       As CWSTR = "10"
      KeywordCase          As Long = 2  ' "Original Case"
      StartupLeft          As Long = 0
      StartupTop           As Long = 0
      StartupRight         As Long = 0
      StartupBottom        As Long = 0
      StartupMaximized     As Long = False
      FBWINCompiler32      As CWSTR
      FBWINCompiler64      As CWSTR
      CompilerSwitches     As CWSTR
      CompilerHelpfile     As CWSTR
      Win32APIHelpfile     As CWSTR
      MRU(9)               As CWSTR
      MRUProject(9)        As CWSTR
      clrCaretFG           As Long = BGR(0,0,0)          ' black
      clrCaretBG           As Long = -1
      clrCommentsFG        As Long = BGR(0,128,0)        ' green
      clrCommentsBG        As Long = BGR(255,255,255)    ' white
      clrHighlightedFG     As Long = BGR(255,255,0)      ' yellow
      clrHighlightedBG     As Long = -1
      clrKeywordFG         As Long = BGR(0,0,255)        ' blue
      clrKeywordBG         As Long = BGR(255,255,255)    ' white
      clrFoldMarginFG      As Long = BGR(237,236,235)    ' pale gray
      clrFoldMarginBG      As Long = -1
      clrFoldSymbolFG      As Long = BGR(255,255,255)    ' white
      clrFoldSymbolBG      As Long = BGR(0,0,0)          ' black
      clrLineNumbersFG     As Long = BGR(0,0,0)          ' black
      clrLineNumbersBG     As Long = BGR(196,196,196)    ' light gray
      clrBookmarksFG       As Long = BGR(0,0,0)          ' black
      clrBookmarksBG       As Long = BGR(0,0,255)        ' blue
      clrOperatorsFG       As Long = BGR(196,0,0)        ' red
      clrOperatorsBG       As Long = BGR(255,255,255)    ' white
      clrIndentGuidesFG    As Long = BGR(255,255,255)    ' white
      clrIndentGuidesBG    As Long = BGR(0,0,0)          ' black
      clrPreprocessorFG    As Long = BGR(196,0,0)        ' red
      clrPreprocessorBG    As Long = BGR(255,255,255)    ' white
      clrSelectionFG       As Long = GetSysColor(COLOR_HIGHLIGHTTEXT) ' white
      clrSelectionBG       As Long = GetSysColor(COLOR_HIGHLIGHT)     ' light blue
      clrStringsFG         As Long = BGR(173,0,173)      ' Purple (little darker than Magenta)
      clrStringsBG         As Long = BGR(255,255,255)    ' white
      clrTextFG            As Long = BGR(0,0,0)          ' black
      clrTextBG            As Long = BGR(255,255,255)    ' white
      clrWinAPIFG          As Long = BGR(0,0,255)        ' blue
      clrWinAPIBG          As Long = BGR(255,255,255)    ' white
      clrWindowFG          As Long = BGR(255,255,255)    ' white
      clrWindowBG          As Long = -1 
      
      Declare Constructor()
      Declare Function LoadKeywords() As Long
      Declare Function SaveKeywords() As Long
      Declare Function SaveToFile() As Long
      Declare Function LoadFromFile() As Long
      Declare Function ProjectSaveToFile() As BOOLEAN    
      declare Function ProjectLoadFromFile( byref wzFile as WSTRING) As BOOLEAN    
      declare Function LoadCodetips( ByRef sFilename As Const String ) as boolean
      declare Function LoadCodetipsWinAPI( ByRef sFilename As Const String ) as boolean
End Type


type clsProject
   private:
      m_arrDocuments(Any) As clsDocument Ptr
   
   public:
      InUse                as boolean     ' this spot in the Projects array is in use
      ProjectType          As Long        ' dll, exe, lib               
      ProjectName          As CWSTR
      ProjectFilename      As CWSTR
      ProjectOther32       As CWSTR       ' compile flags 32 bit compiler
      ProjectOther64       As CWSTR       ' compile flags 64 bit compiler
      hExplorerRootNode    As HTREEITEM
      ProjectErrorOption   As Long
      ProjectDebug         As Long
      ProjectThread        As Long
      ProjectShowConsole   As Long
      ProjectNotes         as CWSTR       ' Save/Load from project file
      ProjectCompiler      as CWSTR = WSTR("FBC 32bit")
      ProjectCompileMode   as CWSTR = WSTR("CONSOLE")
      ProjectCommandLine   as CWSTR
      
      Declare Function AddDocument( ByVal pDoc As clsDocument Ptr ) As Long
      declare Function RemoveDocumentFromArray( ByVal idx As Long) As Long
      Declare Function RemoveDocument( ByVal pDoc As clsDocument Ptr ) As Long
      declare Function RemoveAllDocuments() As Long
      Declare Function GetDocumentCount() As Long
      Declare Function GetDocumentPtr( ByVal idx As Long ) As clsDocument Ptr
      Declare Function GetDocumentPtrByFilename( ByVal pswzName As WString Ptr ) As clsDocument Ptr
      Declare Function GetMainDocumentPtr() As clsDocument Ptr
      Declare Function GetResourceDocumentPtr() As clsDocument Ptr
      Declare Function SaveProject( ByVal bSaveAs As BOOLEAN = False ) As BOOLEAN
      Declare Function ProjectSetFileType( ByVal pDoc As clsDocument Ptr, ByVal nFileType As Long ) As LRESULT
      Declare Function Debug() As Long
END TYPE

Type clsApp
   Private: 
      
   Public:
      IsUnicodeCodetips       As BOOLEAN     ' UNICODE define exists. Use Unicode version of codetips
      SuppressNotify          As BOOLEAN     ' temporarily suppress Scintilla notifications
      hRecentFilesRootNode    As HTREEITEM
      hRecentProjectsRootNode As HTREEITEM
      bDragTabActive          as BOOLEAN     ' A tab in the top tabcontrol is being dragged
      bDragActive             As BOOLEAN     ' splitter drag is currently active 
      hWndPanel               As HWND        ' the panel being split left/right or up/down
      IncludeFilename         As CWSTR
      NonProjectNotes         as CWSTR       ' Save/load from config file
      IsNewProjectFlag        As BOOLEAN
      ProjectOverrideIndex    as long        ' Do action to specific project rather than the active project
      
      Projects(any) as clsProject 
      
      declare function GetProjectCount() as LONG
      declare Function GetActiveProjectIndex() As long
      declare Function SetActiveProject( byval hNode As HTREEITEM ) As long
      declare Function EnsureDefaultActiveProject(byval hNode as HTREEITEM = 0) As long
      declare Function RemoveAllSelectionAttributes() As long
      declare Function GetProjectIndexByFilename( byref sFilename as wstring ) As long
         
      Declare Function IsProjectActive() As boolean
      declare function GetNewProjectIndex() As Long
      
      Declare Constructor()
      Declare Destructor()
      
      
End Type


'  Global classes
Dim Shared gApp As clsApp
Dim Shared gConfig As clsConfig
Dim Shared gTTabCtl As clsTopTabCtl



