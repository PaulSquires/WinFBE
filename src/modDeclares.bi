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


''  Control identifiers
#Define IDC_SCINTILLA 100

#Define IDC_FRMMAIN_TOPTABCONTROL                   1000
#Define IDC_FRMMAIN_TOOLBAR                         1001
#Define IDC_FRMMAIN_REBAR                           1002
#Define IDC_FRMMAIN_STATUSBAR                       1003

#Define IDC_FRMOUTPUT_TABCONTROL                    1000
#Define IDC_FRMOUTPUT_LISTVIEW                      1001
#Define IDC_FRMOUTPUT_TXTLOGFILE                    1002
#Define IDC_FRMOUTPUT_BTNCLOSE                      1003

#Define IDC_FRMEXPLORER_LBLEXPLORER                 1000
#Define IDC_FRMEXPLORER_TREE                        1001
#Define IDC_FRMEXPLORER_BTNCLOSE                    1002

#Define IDC_FRMRECENT_BTNOPENFILE                   1000
#Define IDC_FRMRECENT_BTNOPENPROJECT                1001
#Define IDC_FRMRECENT_TREEFILES                     1002
#Define IDC_FRMRECENT_TREEPROJECTS                  1003

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
#Define IDC_FRMOPTIONSEDITOR_CHKSHOWLEFTMARGIN      1005
#Define IDC_FRMOPTIONSEDITOR_CHKSYNTAXHIGHLIGHTING  1006
#Define IDC_FRMOPTIONSEDITOR_CHKCURRENTLINE         1007
#Define IDC_FRMOPTIONSEDITOR_CHKLINENUMBERING       1008
#Define IDC_FRMOPTIONSEDITOR_CHKCONFINECARET        1009
#Define IDC_FRMOPTIONSEDITOR_CHKTABTOSPACES         1010
#Define IDC_FRMOPTIONSEDITOR_CHKAUTOINDENTATION     1011
#Define IDC_FRMOPTIONSEDITOR_CHKSHOWFOLDMARGIN      1012
#Define IDC_FRMOPTIONSEDITOR_CHKINDENTGUIDES        1013
#Define IDC_FRMOPTIONSEDITOR_CHKUNICODE             1014

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
#Define IDC_FRMFIND_FRAMEOPTIONS                    1009

#Define IDC_FRMREPLACE_LBLFINDWHAT                  1000
#Define IDC_FRMREPLACE_COMBOFIND                    1001
#Define IDC_FRMREPLACE_CHKWHOLEWORDS                1002
#Define IDC_FRMREPLACE_CHKMATCHCASE                 1003
#Define IDC_FRMREPLACE_FRAMESCOPE                   1004
#Define IDC_FRMREPLACE_OPTSCOPE1                    1005
#Define IDC_FRMREPLACE_OPTSCOPE2                    1006
#Define IDC_FRMREPLACE_FRAMEOPTIONS                 1007
#Define IDC_FRMREPLACE_LBLREPLACEWITH               1008
#Define IDC_FRMREPLACE_COMBOREPLACE                 1009
#Define IDC_FRMREPLACE_CMDREPLACE                   1010
#Define IDC_FRMREPLACE_CMDREPLACEALL                1011
#Define IDC_FRMREPLACE_LBLSTATUS                    1012

#Define IDC_FRMGOTO_LBLLASTLINE                     1000
#Define IDC_FRMGOTO_LBLCURRENTLINE                  1001
#Define IDC_FRMGOTO_LBLGOTOLINE                     1002
#Define IDC_FRMGOTO_TXTLINE                         1003
#Define IDC_FRMGOTO_LBLLASTVALUE                    1004
#Define IDC_FRMGOTO_LBLCURRENTVALUE                 1005

#Define IDC_FRMCOMMANDLINE_LABEL1                   1000
#Define IDC_FRMCOMMANDLINE_TEXTBOX1                 1001

#Define IDC_FRMFNLIST_LISTBOX                       1000


Const DELIM = Chr(1)                 ' character used as delimiter for Find/Replace text strings  
Const IDC_MRUBASE = 5000             ' Windows id of MRU items 1 to 10 (located under File menu)
Const IDC_MRUPROJECTBASE = 6000      ' Windows id of MRUPROJECT items 1 to 10 (located under Project menu)

Const FILETYPE_UNDEFINED = 0
Const FILETYPE_MAIN      = 1
Const FILETYPE_MODULE    = 2
Const FILETYPE_NORMAL    = 3
Const FILETYPE_RESOURCE  = 4
 
   
''  Menu message identifiers
Enum
''  Custom messages
   MSG_USER_SETFOCUS = WM_USER + 1
   MSG_USER_PROCESS_COMMANDLINE 
   IDM_FILE, IDM_FILENEW 
   IDM_FILEOPEN, IDM_FILECLOSE, IDM_FILECLOSEALL, IDM_FILESAVE, IDM_FILESAVEAS, IDM_FILESAVEALL
   IDM_MRU, IDM_OPENINCLUDE, IDM_COMMAND, IDM_EXIT
   IDM_EDIT
   IDM_UNDO, IDM_REDO, IDM_CUT, IDM_COPY, IDM_PASTE, IDM_DELETELINE, IDM_INSERTFILE
   IDM_INDENTBLOCK, IDM_UNINDENTBLOCK, IDM_COMMENTBLOCK, IDM_UNCOMMENTBLOCK
   IDM_DUPLICATELINE, IDM_MOVELINEUP, IDM_MOVELINEDOWN, IDM_TOUPPERCASE, IDM_TOLOWERCASE
   IDM_TOMIXEDCASE, IDM_EOLTOCRLF, IDM_EOLTOCR, IDM_EOLTOLF, IDM_SELECTALL, IDM_SELECTLINE
   IDM_SPACESTOTABS, IDM_TABSTOSPACES
   IDM_SEARCH
   IDM_FIND, IDM_FINDNEXT, IDM_FINDPREV, IDM_REPLACE, IDM_DEFINITION, IDM_LASTPOSITION, IDM_FUNCTIONLIST
   IDM_GOTO, IDM_BOOKMARKTOGGLE, IDM_BOOKMARKNEXT, IDM_BOOKMARKPREV, IDM_BOOKMARKCLEARALL
   IDM_VIEW
   IDM_FOLDTOGGLE, IDM_FOLDBELOW, IDM_FOLDALL, IDM_UNFOLDALL, IDM_ZOOMIN, IDM_ZOOMOUT, IDM_RESTOREMAIN
   IDM_VIEWEXPLORER, IDM_VIEWOUTPUT
   IDM_OPTIONS
   IDM_PROJECTNEW, IDM_PROJECTMANAGER, IDM_PROJECTOPEN, IDM_MRUPROJECT, IDM_PROJECTCLOSE 
   IDM_PROJECTSAVE, IDM_PROJECTSAVEAS, IDM_PROJECTFILESADD, IDM_PROJECTOPTIONS
   IDM_ADDFILETOPROJECT, IDM_REMOVEFILEFROMPROJECT
   IDM_BUILDEXECUTE, IDM_COMPILE, IDM_RUNEXE, IDM_COMMANDLINE
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

'  Global window handles for some forms 
Dim Shared As HWnd HWND_FRMOPTIONSGENERAL, HWND_FRMOPTIONSEDITOR, HWND_FRMOPTIONSCOLORS
Dim Shared As HWnd HWND_FRMOPTIONSCOMPILER, HWND_FRMOPTIONSLOCAL, HWND_FRMOPTIONSKEYWORDS
Dim Shared As HWnd HWND_FRMFIND, HWND_FRMREPLACE
Dim Shared As HWnd HWND_FRMFNLIST

'  Global handle to hhctrl.ocx for context sensitive help
Dim Shared As Any Ptr gpHelpLib

 
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

#Define SetFocusScintilla  PostMessageW HWND_FRMMAIN, MSG_USER_SETFOCUS, 0, 0
#Define ResizeExplorerWindow  PostMessageW HWND_FRMEXPLORER, MSG_USER_OPENEDITORS_RESIZE, 0, 0


''
''  Save information related to Find/Replace operations
''
Type FINDREPLACE_TYPE
   txtFind         As String
   txtReplace      As String
   txtFindCombo    As String
   txtReplaceCombo As String
   txtLastFind     As String
   nWholeWord      As Long          ' checkmark for find/replace whole word search
   nMatchCase      As Long          ' match case when searching
   nScopeFind      As Long          ' identifier of OptionButton that is checked
   nScopeReplace   As Long          ' identifier of OptionButton that is checked
End Type
Dim Shared gFind As FINDREPLACE_TYPE



' Forward reference
Type clsDocument_ As clsDocument

'' Last position in document. Used when "Last Position" menu option is selected.
Type LASTPOSITION_TYPE
   pDoc       As clsDocument_ Ptr
   nFirstLine As Long     ' first visible line on screen
   nPosition  As Long     ' Position in Scintilla document where caret is positioned
End Type
Dim Shared gLastPosition As LASTPOSITION_TYPE

' Linked list of Function names
Type FUNCTION_TYPE
   pListPrev   As FUNCTION_TYPE Ptr
   pListNext   As FUNCTION_TYPE Ptr
   bIsHeader   As BOOLEAN
   zFnName     As WString * MAX_PATH
   nLineNumber As Long
   pDoc        As clsDocument_ Ptr
End Type

Type clsDocument
   Private:
      m_pSci           As Any Ptr      
      
   Public:
      hWindow          As HWnd
      IsNewFlag        As BOOLEAN
      IsProjectFile    As BOOLEAN
      ProjectFileType  As Long = FILETYPE_UNDEFINED
      FnListPtr        As FUNCTION_TYPE Ptr
      DiskFilename     As WString * MAX_PATH
      DateFileTime     As FILETIME  
      hNodeExplorer    As HTREEITEM

      Declare Function CreateCodeWindow( ByVal hWndParent As HWnd, ByVal IsNewFile As BOOLEAN, ByVal IsTemplate As BOOLEAN = False, ByVal pwszFile As WString Ptr = 0 ) As HWnd
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
      Declare Function CreateFunctionList() As Long
      Declare Function DeallocateFunctionList() As Long
      Declare Function GetWord( ByVal curPos As Long = -1 ) As String
      Declare Function GetBookmarks() As String
      Declare Function SetBookmarks( ByVal sBookmarks As String ) As Long
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
      bKeywordsDirty       As BOOLEAN = True       ' not saved to file
      CommandLine          As CWSTR                ' not saved to file
      LastRunFilename      As CWSTR                ' not saved to file
      CloseFuncList        As Long = True
      ShowExplorer         As Long = True
      ShowExplorerWidth    As Long = 250
      SyntaxHighlighting   As Long = True
      Codetips             As Long = True
      LeftMargin           As Long = True
      FoldMargin           As Long = True
      AutoIndentation      As Long = True
      ConfineCaret         As Long = True
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
      DefaultCompiler      As CWSTR = "FBC 32bit"
      DefaultCompileMode   As CWSTR = "GUI"
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
      Declare Function SaveToFile() As Long
      Declare Function LoadFromFile() As Long
      Declare Function ProjectSaveToFile() As BOOLEAN    
      Declare Function ProjectLoadFromFile() As BOOLEAN    
End Type


Type clsApp
   Private: 
      
      m_arrDocuments(Any) As clsDocument Ptr
   
   Public:
      ProjectType          As Long                       
      IsProjectActive      As BOOLEAN
      IsNewProjectFlag     As BOOLEAN
      ProjectErrorOption   As Long
      ProjectDebug         As Long
      ProjectThread        As Long
      ProjectShowConsole   As Long
      SuppressNotify       As BOOLEAN     ' temporarily suppress Scintilla notifications
      ProjectName          As CWSTR
      ProjectFilename      As CWSTR
      ProjectOther32       As CWSTR       ' compile flags 32 bit compiler
      ProjectOther64       As CWSTR       ' compile flags 64 bit compiler
      IncludeFilename      As CWSTR
      hExplorerRootNode    As HTREEITEM
      bDragActive          As BOOLEAN     ' splitter drag is currently active
      hWndPanel            As HWND        ' the panel being split left/right or up/down

      Declare Function SaveProject( ByVal bSaveAs As BOOLEAN = False ) As BOOLEAN
      Declare Function ProjectAddFile( ByVal pDoc As clsDocument Ptr ) As LRESULT
      Declare Function ProjectRemoveFile( ByVal pDoc As clsDocument Ptr ) As LRESULT
      Declare Function ProjectSetFileType( ByVal pDoc As clsDocument Ptr, ByVal nFileType As Long ) As LRESULT

      Declare Function LoadKeywords() As Long
      Declare Function SaveKeywords() As Long

      Declare Function AddDocument( ByVal pDoc As clsDocument Ptr ) As Long
      declare Function RemoveDocumentFromArray( ByVal idx As Long) As Long
      Declare Function RemoveDocument( ByVal pDoc As clsDocument Ptr ) As Long
      declare Function RemoveAllDocuments() As Long
      Declare Function GetDocumentCount() As Long
      Declare Function GetDocumentPtr( ByVal idx As Long ) As clsDocument Ptr
      Declare Function GetDocumentPtrByFilename( ByVal pswzName As WString Ptr ) As clsDocument Ptr
      Declare Function GetMainDocumentPtr() As clsDocument Ptr
      Declare Function GetResourceDocumentPtr() As clsDocument Ptr

      Declare Constructor()
      Declare Destructor()
      
      FBKeywords As String 
      
      'Declare Function Debug() As Long
      
End Type


'  Global classes
Dim Shared gApp As clsApp
Dim Shared gConfig As clsConfig
Dim Shared gTTabCtl As clsTopTabCtl



Declare Function frmCommandLine_OnCreate(ByVal HWnd As HWnd, ByVal lpCreateStructPtr As LPCREATESTRUCT) As BOOLEAN
Declare Function frmCommandLine_OnCommand(ByVal HWnd As HWnd, ByVal id As Long, ByVal hwndCtl As HWnd, ByVal codeNotify As UINT) As LRESULT
Declare Function frmCommandLine_OnClose(HWnd As HWnd) As LRESULT
Declare Function frmCommandLine_OnDestroy(HWnd As HWnd) As LRESULT
Declare Function frmCommandLine_WndProc( ByVal HWnd As HWnd, ByVal uMsg As UINT, ByVal wParam As WPARAM, ByVal lParam As LPARAM ) As LRESULT
Declare Function frmCommandLine_Show( ByVal hWndParent As HWnd, ByVal nCmdShow As Long = 0 ) As Long
Declare Function SetDocumentErrorPosition( ByVal hLV As HWnd ) As Long
Declare Function Find_UpOrDown( ByVal flagUpDown As Long, ByVal findFlags As Long, ByVal flagSearchCurrentOnly As BOOLEAN, ByVal hWndDialog As HWnd ) As Long
Declare Function frmFind_AddToFindCombo( ByRef sText As Const String ) As Long
Declare Function frmFind_OnCreate(ByVal HWnd As HWnd, ByVal lpCreateStructPtr As LPCREATESTRUCT) As BOOLEAN
Declare Function frmFind_OnCommand(ByVal HWnd As HWnd, ByVal id As Long, ByVal hwndCtl As HWnd, ByVal codeNotify As UINT) As LRESULT
Declare Function frmFind_OnClose(HWnd As HWnd) As LRESULT
Declare Function frmFind_OnDestroy(HWnd As HWnd) As LRESULT
Declare Function frmFind_WndProc( ByVal HWnd As HWnd, ByVal uMsg As UINT, ByVal wParam As WPARAM, ByVal lParam As LPARAM ) As LRESULT
Declare Function frmFind_Show( ByVal hWndParent As HWnd ) As Long
Declare Function frmGoto_OnCreate(ByVal HWnd As HWnd, ByVal lpCreateStructPtr As LPCREATESTRUCT) As BOOLEAN
Declare Function frmGoto_OnCommand(ByVal HWnd As HWnd, ByVal id As Long, ByVal hwndCtl As HWnd, ByVal codeNotify As UINT) As LRESULT
Declare Function frmGoto_OnClose(HWnd As HWnd) As LRESULT
Declare Function frmGoto_OnDestroy(HWnd As HWnd) As LRESULT
Declare Function frmGoto_WndProc( ByVal HWnd As HWnd, ByVal uMsg As UINT, ByVal wParam As WPARAM, ByVal lParam As LPARAM ) As LRESULT
Declare Function frmGoto_Show( ByVal hWndParent As HWnd ) As Long
Declare Function frmMain_UpdateLineCol(ByVal HWnd As HWnd) As Long
Declare Function Scintilla_OnNotify( ByVal HWnd As HWnd, ByVal pNSC As SCNOTIFICATION Ptr ) As Long
Declare Function frmMain_SetFocusToCurrentCodeWindow() As Long
Declare Function frmMain_PositionWindows( ByVal HWnd As HWnd ) As LRESULT
Declare Function OnCommand_ProjectSave( ByVal HWnd As HWnd, ByVal bSaveAs As BOOLEAN = False) As LRESULT
Declare Function OnCommand_ProjectClose( ByVal HWnd As HWnd ) As LRESULT
Declare Function OnCommand_ProjectNew( ByVal HWnd As HWnd ) As LRESULT
Declare Function OnCommand_ProjectOpen( ByVal HWnd As HWnd ) As LRESULT
Declare Function frmMain_OpenFileSafely( ByVal HWnd As HWnd, ByVal bIsNewFile As BOOLEAN, ByVal bIsTemplate As BOOLEAN, ByVal bShowInTab As BOOLEAN, byval bIsInclude as BOOLEAN, ByVal pwszName As WString Ptr, ByVal pDocIn As clsDocument Ptr ) As clsDocument Ptr
Declare Function OnCommand_FileNew( ByVal HWnd As HWnd ) As LRESULT
Declare Function OnCommand_FileOpen( ByVal HWnd As HWnd ) As LRESULT
Declare Function OnCommand_ProjectAddFiles( ByVal HWnd As HWnd ) As LRESULT
Declare Function OnCommand_OpenIncludeFile( ByVal HWnd As HWnd ) As LRESULT
Declare Function OnCommand_FileSave( ByVal HWnd As HWnd, ByVal bSaveAs As BOOLEAN = False) As LRESULT
Declare Function OnCommand_FileSaveAll( ByVal HWnd As HWnd ) As LRESULT
Declare Function OnCommand_FileClose( ByVal HWnd As HWnd, ByVal bCloseAll As BOOLEAN = False) As LRESULT
Declare Function frmMain_OnCreate(ByVal HWnd As HWnd, ByVal lpCreateStructPtr As LPCREATESTRUCT) As BOOLEAN
Declare Function frmMain_OnSize(ByVal HWnd As HWnd, ByVal state As UINT, ByVal cx As Long, ByVal cy As Long) As LRESULT
Declare Function frmMain_OnCommand(ByVal HWnd As HWnd, ByVal id As Long, ByVal hwndCtl As HWnd, ByVal codeNotify As UINT) As LRESULT
Declare Function frmMain_OnNotify(ByVal HWnd As HWnd, ByVal id As Long, ByVal pNMHDR As NMHDR Ptr) As LRESULT
Declare Function frmMain_OnActivateApp(ByVal HWnd As HWnd, ByVal fActivate As BOOLEAN, ByVal dwThreadId As DWORD) As LRESULT
Declare Function frmMain_OnContextMenu( ByVal HWnd As HWnd, ByVal hwndContext As HWnd, ByVal xPos As Long, ByVal yPos As Long ) As LRESULT
Declare Function frmMain_OnDropFiles( ByVal HWnd As HWnd, ByVal hDrop As HDROP ) As LRESULT
Declare Function frmMain_OnClose(ByVal HWnd As HWnd) As LRESULT
Declare Function frmMain_OnDestroy(HWnd As HWnd) As LRESULT
Declare Function frmMain_WndProc( ByVal HWnd As HWnd, ByVal uMsg As UINT, ByVal wParam As WPARAM, ByVal lParam As LPARAM ) As LRESULT
Declare Function frmMain_Show( ByVal hWndParent As HWnd, ByVal nCmdShow As Long = 0 ) As Long
Declare Sub SaveEditorOptions()
Declare Function frmOptions_OnCreate(ByVal HWnd As HWnd, ByVal lpCreateStructPtr As LPCREATESTRUCT) As BOOLEAN
Declare Function frmOptions_OnCommand(ByVal HWnd As HWnd, ByVal id As Long, ByVal hwndCtl As HWnd, ByVal codeNotify As UINT) As LRESULT
Declare Function frmOptions_OnNotify(ByVal HWnd As HWnd, ByVal id As Long, ByVal pNMHDR As NMHDR Ptr) As LRESULT
Declare Function frmOptions_OnCtlColorStatic(HWnd As HWnd, hdc As HDC, hWndChild As HWnd, nType As Long) As HBRUSH
Declare Function frmOptions_OnClose(HWnd As HWnd) As LRESULT
Declare Function frmOptions_OnDestroy(HWnd As HWnd) As LRESULT
Declare Function frmOptions_WndProc( ByVal HWnd As HWnd, ByVal uMsg As UINT, ByVal wParam As WPARAM, ByVal lParam As LPARAM ) As LRESULT
Declare Function frmOptions_Show( ByVal hWndParent As HWnd, ByVal nCmdShow As Long = 0 ) As Long
Declare Function EnumFontName(lf As LOGFONTW, tm As TEXTMETRIC, ByVal FontType As Long, HWnd As HWnd) As Long
Declare Sub FillFontCombo( ByVal HWnd As HWnd)
Declare Function DrawFontCombo(ByVal HWnd As HWnd, ByVal wParam As WPARAM, ByVal lParam As LPARAM) As Long
Declare Sub FillFontSizeCombo( ByVal hCb As HWnd, ByVal strFontName As WString Ptr )
Declare Sub FillFontCharSets( ByVal hCtl As HWnd )
Declare Function SetTempColorSelection( ByVal HWnd As HWnd, ByVal nCtrlID As Long ) As Long
Declare Function ShowComboColors( ByVal HWnd As HWnd, ByVal nCurSel As Long ) As Long
Declare Function frmOptionsColors_OnCommand(ByVal HWnd As HWnd, ByVal id As Long, ByVal hwndCtl As HWnd, ByVal codeNotify As UINT) As LRESULT
Declare Function frmOptionsColors_WndProc( ByVal HWnd As HWnd, ByVal uMsg As UINT, ByVal wParam As WPARAM, ByVal lParam As LPARAM ) As LRESULT
Declare Function frmOptionsColors_Show( ByVal hWndParent As HWnd, ByVal nCmdShow As Long = 0 ) As Long
Declare Function frmOptionsCompiler_OnCommand(ByVal HWnd As HWnd, ByVal id As Long, ByVal hwndCtl As HWnd, ByVal codeNotify As UINT) As LRESULT
Declare Function frmOptionsCompiler_WndProc( ByVal HWnd As HWnd, ByVal uMsg As UINT, ByVal wParam As WPARAM, ByVal lParam As LPARAM ) As LRESULT
Declare Function frmOptionsCompiler_Show( ByVal hWndParent As HWnd, ByVal nCmdShow As Long = 0 ) As Long
Declare Function frmOptionsEditor_WndProc( ByVal HWnd As HWnd, ByVal uMsg As UINT, ByVal wParam As WPARAM, ByVal lParam As LPARAM ) As LRESULT
Declare Function frmOptionsEditor_Show( ByVal hWndParent As HWnd, ByVal nCmdShow As Long = 0 ) As Long
Declare Function frmOptionsKeywords_OnCommand(ByVal HWnd As HWnd, ByVal id As Long, ByVal hwndCtl As HWnd, ByVal codeNotify As UINT) As LRESULT
Declare Function frmOptionsKeywords_OnDestroy(HWnd As HWnd) As LRESULT
Declare Function frmOptionsKeywords_WndProc( ByVal HWnd As HWnd, ByVal uMsg As UINT, ByVal wParam As WPARAM, ByVal lParam As LPARAM ) As LRESULT
Declare Function frmOptionsKeywords_Show( ByVal hWndParent As HWnd, ByVal nCmdShow As Long = 0 ) As Long
Declare Function frmOptionsLocal_OnCommand(ByVal HWnd As HWnd, ByVal id As Long, ByVal hwndCtl As HWnd, ByVal codeNotify As UINT) As LRESULT
Declare Function frmOptionsLocal_WndProc( ByVal HWnd As HWnd, ByVal uMsg As UINT, ByVal wParam As WPARAM, ByVal lParam As LPARAM ) As LRESULT
Declare Function frmOptionsLocal_Show( ByVal hWndParent As HWnd, ByVal nCmdShow As Long = 0 ) As Long
declare Function PositionExplorerWindows( ByVal HWnd As HWnd ) As LRESULT
Declare Function SaveProjectOptions( ByVal HWnd As HWnd ) As BOOLEAN
Declare Function frmProjectOptions_OnCreate(ByVal HWnd As HWnd, ByVal lpCreateStructPtr As LPCREATESTRUCT) As BOOLEAN
Declare Function frmProjectOptions_OnCommand(ByVal HWnd As HWnd, ByVal id As Long, ByVal hwndCtl As HWnd, ByVal codeNotify As UINT) As LRESULT
Declare Function frmProjectOptions_OnClose(HWnd As HWnd) As LRESULT
Declare Function frmProjectOptions_OnDestroy(HWnd As HWnd) As LRESULT
Declare Function frmProjectOptions_WndProc( ByVal HWnd As HWnd, ByVal uMsg As UINT, ByVal wParam As WPARAM, ByVal lParam As LPARAM ) As LRESULT
Declare Function frmProjectOptions_Show( ByVal hWndParent As HWnd, ByVal nCmdShow As Long = 0 ) As Long
Declare Function frmReplace_AddToReplaceCombo( ByRef sText As Const String ) As Long
Declare Function frmReplace_OnCreate(ByVal HWnd As HWnd, ByVal lpCreateStructPtr As LPCREATESTRUCT) As BOOLEAN
Declare Function frmReplace_OnCommand(ByVal HWnd As HWnd, ByVal id As Long, ByVal hwndCtl As HWnd, ByVal codeNotify As UINT) As LRESULT
Declare Function frmReplace_OnClose(HWnd As HWnd) As LRESULT
Declare Function frmReplace_OnDestroy(HWnd As HWnd) As LRESULT
Declare Function frmReplace_WndProc( ByVal HWnd As HWnd, ByVal uMsg As UINT, ByVal wParam As WPARAM, ByVal lParam As LPARAM ) As LRESULT
Declare Function frmReplace_Show( ByVal hWndParent As HWnd ) As Long
Declare Function frmTemplates_OnCreate(ByVal HWnd As HWnd, ByVal lpCreateStructPtr As LPCREATESTRUCT) As BOOLEAN
Declare Function frmTemplates_OnCommand(ByVal HWnd As HWnd, ByVal id As Long, ByVal hwndCtl As HWnd, ByVal codeNotify As UINT) As LRESULT
Declare Function frmTemplates_OnClose(HWnd As HWnd) As LRESULT
Declare Function frmTemplates_OnDestroy(HWnd As HWnd) As LRESULT
Declare Function frmTemplates_WndProc (ByVal HWnd As HWnd, ByVal uMsg As UINT, ByVal wParam As WPARAM, ByVal lParam As LPARAM) As LRESULT
Declare Function frmTemplates_Show (ByVal hParent As HWnd, ByVal x As Long, ByVal y As Long) As Long
Declare Sub GetColorInfo( ByVal nIndex As Long, ByVal wColorName As WString Ptr, ByRef nColorValue As COLORREF )
Declare Function CBProc( ByVal HWnd As HWnd, ByVal wMsg As UInt, ByVal wParam As WPARAM, ByVal lParam As LPARAM ) As LRESULT
Declare Function CreateCBColorList( ByVal HWnd As HWnd, ByVal CtrlId As Long, ByVal nLeft As Long, ByVal nTop As Long, ByVal nWidth As Long, ByVal nHeight As Long ) As HWnd
Declare Function code_Compile( ByVal wID As Long ) As BOOLEAN
Declare Function GetMRUMenuHandle() As HMENU
Declare Function OpenMRUFile( ByVal HWnd As HWnd, ByVal wID As Long ) As Long
Declare Function UpdateMRUMenu( ByVal hMenu As HMENU ) As Long
Declare Function UpdateMRUList( ByVal wzFilename As WString Ptr ) As Long
Declare Function GetMRUProjectMenuHandle() As HMENU
Declare Function OpenMRUProjectFile( ByVal HWnd As HWnd, ByVal wID As Long, ByVal pwszFilename As WString Ptr = 0 ) As Long
Declare Function UpdateMRUProjectMenu( ByVal hMenu As HMENU ) As Long
Declare Function UpdateMRUProjectList( ByVal wzFilename As WString Ptr ) As Long
Declare Function GetFontCharSetID(ByREF wzCharsetName As CWSTR ) As Long
Declare Function ProcessToCurdrive( ByRef wzFilename As CWSTR ) As CWSTR
Declare Function ProcessFromCurdrive( ByRef wzFilename As CWSTR ) As CWSTR
Declare Function FF_TreeView_InsertItem( ByVal hWndControl As HWnd, ByVal hParent As HANDLE, ByRef TheText As WString, ByVal lParam As LPARAM = 0, ByVal iImage As Long = 0, ByVal iSelectedImage As Long = 0 ) As HANDLE
Declare Function FF_TreeView_GetlParam( ByVal hWndControl As HWnd, ByVal hItem As HANDLE ) As Long
declare Function AfxGetComboBoxText( ByVal hWndControl As HWnd, ByVal nIndex As Long ) As CWSTR
Declare Function AfxIFileOpenDialogW( ByVal hwndOwner As HWnd, ByVal idButton As Long) As WString Ptr
Declare Function AfxIFileOpenDialogMultiple( ByVal hwndOwner As HWnd, ByVal sigdnName As SIGDN = SIGDN_FILESYSPATH) As IShellItemArray Ptr
Declare Function AfxIFileSaveDialog( ByVal hwndOwner As HWnd, ByVal pwszFileName As WString Ptr, ByVal pwszDefExt As WString Ptr, ByVal id As Long = 0, ByVal sigdnName As SIGDN = SIGDN_FILESYSPATH ) As WString Ptr
Declare Function FF_Toolbar_EnableButton (ByVal hToolBar As HWnd, ByVal idButton As Long) As BOOLEAN
Declare Function FF_Toolbar_DisableButton (ByVal hToolBar As HWnd, ByVal idButton As Long) As BOOLEAN
Declare Function FF_ListView_InsertItem( ByVal hWndControl As HWnd, ByVal iRow As Long, ByVal iColumn As Long, ByVal pwszText As WString Ptr, ByVal lParam As LPARAM = 0 ) As BOOLEAN
Declare Function FF_ListView_GetItemText( ByVal hWndControl As HWnd, ByVal iRow As Long, ByVal iColumn As Long, ByVal pwszText As WString Ptr, ByVal nTextMax As Long ) As BOOLEAN
Declare Function FF_ListView_SetItemText( ByVal hWndControl As HWnd, ByVal iRow As Long, ByVal iColumn As Long, ByVal pwszText As WString Ptr, ByVal nTextMax As Long ) As Long
Declare Function FF_ListView_GetlParam( ByVal hWndControl As HWnd, ByVal iRow As Long ) As LPARAM
Declare Function LoadLocalizationFile( ByVal pwszName As WString Ptr ) As BOOLEAN
Declare Function GetProcessImageName( ByVal pe32w As PROCESSENTRY32W Ptr, ByVal pwszExeName As WString Ptr ) As Long
Declare Function IsProcessRunning( ByVal pwszExeFileName As WString Ptr ) As BOOLEAN
Declare Function frmMain_CreateToolbar( ByVal pWindow As CWindow Ptr ) As BOOLEAN
Declare Function frmMain_DisableToolbarButtons() As Long
Declare Function frmMain_ChangeToolbarButtonsState() As Long
Declare Function frmMain_BuildMenu ( ByVal pWindow As CWindow Ptr ) As HMENU
Declare Function frmMain_ChangeTopMenuStates() As Long
Declare Function frmMain_MenuSetCompiler( ByVal wID As Long ) As Long
Declare Function frmMain_MenuSetCompileMode( ByVal wID As Long ) As Long
Declare Function AddProjectFileTypesToMenu( ByVal hPopUpMenu As HMENU, ByVal pDoc As clsDocument Ptr, byval fNoSeparator as BOOLEAN = false ) As Long
declare Function CreateStatusBarFileTypeContextMenu() As HMENU
Declare Function CreateTopTabCtlContextMenu( ByVal idx As Long ) As HMENU
Declare Function CreateScintillaContextMenu() As HMENU
Declare Function frmFnList_UpdateListBox() As Long
Declare Function frmFnList_SetListBoxPosition() As Long
Declare Function CreateMRUpopup() As HMENU
Declare Function frmMain_GotoDefinition( ByVal pDoc As clsDocument Ptr ) As Long
Declare Function frmMain_GotoLastPosition() As Long
Declare Function ClearMRUlist( ByVal wID As Long ) As Long
declare Function RunEXE( ByRef wszFileExe As CWSTR, ByRef wszParam As CWSTR ) As Long
declare Function PositionOutputWindows( ByVal HWnd As HWnd ) As LRESULT
declare Function CreateRootNodeExplorerTreeview() As Long




