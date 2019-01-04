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


''  Control identifiers
#Define IDC_SCINTILLA 100
#Define IDC_SCROLLV   200

#Define IDC_FRMMAIN_TOPTABCONTROL                   1000
#Define IDC_FRMMAIN_TOOLBAR                         1001
#Define IDC_FRMMAIN_REBAR                           1002
#Define IDC_FRMMAIN_STATUSBAR                       1003
#Define IDC_FRMMAIN_COMPILETIMER                    1004
#Define IDC_FRMMAIN_COMBOBUILDS                     1005

#Define IDC_FRMVDCOLORS_TABCONTROL                  1000
#Define IDC_FRMVDCOLORS_LSTCUSTOM                   1001
#Define IDC_FRMVDCOLORS_LSTCOLORS                   1002
#Define IDC_FRMVDCOLORS_LSTSYSTEM                   1003

#Define IDC_LBLFAKEMAINMENU                         1000


Const DELIM = "|"                    ' character used as delimiter for function names in data1 of gFunctionLists hash
Const IDC_MRUBASE = 5000             ' Windows id of MRU items 1 to 10 (located under File menu)
Const IDC_MRUPROJECTBASE = 6000      ' Windows id of MRUPROJECT items 1 to 10 (located under Project menu)

dim shared as long SPLITSIZE 
SPLITSIZE = AfxScaleY(6)       ' Width/Height of the scrollbar split buttons for split editing windows

const LISTBOX_LINEHEIGHT = 20  ' used for ownerdrawn listboxes in the PropertyList and Build config

      
Const FILETYPE_UNDEFINED = 0
Const FILETYPE_MAIN      = 1
Const FILETYPE_MODULE    = 2
Const FILETYPE_NORMAL    = 3
Const FILETYPE_RESOURCE  = 4
 
' File encodings
const FILE_ENCODING_ANSI      = 0
const FILE_ENCODING_UTF8_BOM  = 1
const FILE_ENCODING_UTF16_BOM = 2

' User Tool actions (selected, pre/post compile)
const USERTOOL_ACTION_SELECTED    = 0   
const USERTOOL_ACTION_PRECOMPILE  = 1   
const USERTOOL_ACTION_POSTCOMPILE = 2

' The type of autocomplete popup that is active. This is necessary
' because the autocomplete popup list is rebuilt every time a new
' character is entered.
enum
   AUTOCOMPLETE_NONE   = 0
   AUTOCOMPLETE_DIM_AS 
   AUTOCOMPLETE_TYPE
end enum   
   
enum eFileClose
   EFC_CLOSECURRENT
   EFC_CLOSEALL
   EFC_CLOSEALLFORWARD
   EFC_CLOSEALLOTHERS
   EFC_CLOSEALLBACKWARD 
end enum
   
' Directions when determining next closest control pointer
ENUM
   DIRECTION_TOP = 1
   DIRECTION_LEFT
   DIRECTION_RIGHT
   DIRECTION_BOTTOM
end enum   

' Index for scintilla autocomplete raw PNG for SCI_REGISTERRGBAIMAGE
enum
   IMAGE_AUTOC_BASETYPE = 1        ' default variable types. Long, String, Single, etc
   IMAGE_AUTOC_CLASS    = 2        ' class/TYPEs
   IMAGE_AUTOC_METHOD   = 3        ' subs/functions
   IMAGE_AUTOC_PROPERTY = 4        ' variable within a TYPE that can be set directly
end enum

' Colors
enum
   ' Start the enum at 2 because when theme is saved to file the first parse is the
   ' theme id and theme description. The colors start at parse 2.
   CLR_CARET = 2          
   CLR_COMMENTS
   CLR_HIGHLIGHTED     
   CLR_KEYWORD         
   CLR_FOLDMARGIN
   CLR_FOLDSYMBOL      
   CLR_LINENUMBERS     
   CLR_OPERATORS       
   CLR_INDENTGUIDES    
   CLR_PREPROCESSOR    
   CLR_SELECTION       
   CLR_STRINGS         
   CLR_TEXT            
   CLR_WINDOW
end enum


'  Control types   
enum
   CTRL_FORM = 1
   CTRL_POINTER 
   CTRL_LABEL
   CTRL_BUTTON
   CTRL_TEXTBOX
   CTRL_CHECKBOX
   CTRL_OPTION
   CTRL_FRAME
   CTRL_PICTURE
   CTRL_COMBOBOX
   CTRL_LISTBOX
   CTRL_HSCROLL
   CTRL_VSCROLL
   CTRL_TIMER
   CTRL_TABCONTROL
   CTRL_RICHEDIT
   CTRL_PROGRESSBAR
   CTRL_UPDOWN
   CTRL_LISTVIEW
   CTRL_TREEVIEW
   CTRL_SLIDER
   CTRL_DATETIMEPICKER
   CTRL_MONTHCALENDAR
   CTRL_WEBBROWSER
   CTRL_CUSTOM
   CTRL_OCX
   CTRL_MASKEDEDIT
end enum
   
'  Grab handles (clockwise starting at top left corner)
enum
   GRAB_NOHIT = 0
   GRAB_TOPLEFT 
   GRAB_TOP
   GRAB_TOPRIGHT
   GRAB_RIGHT
   GRAB_BOTTOMRIGHT
   GRAB_BOTTOM
   GRAB_BOTTOMLEFT
   GRAB_LEFT
end enum   


''  Menu message identifiers
Enum
''  Custom messages
   MSG_USER_SETFOCUS = WM_USER + 1
   MSG_USER_SHOWCOLORCOMBOBOXES
   MSG_USER_SETCOLORCUSTOM
   MSG_USER_GETCOLORCUSTOM
   MSG_USER_PROCESS_COMMANDLINE 
   MSG_USER_PROCESS_UPDATECHECK
   MSG_USER_SHOWAUTOCOMPLETE
   MSG_USER_APPENDEQUALSSIGN
   MSG_USER_GENERATECODE
   IDM_CREATE_THEME, IDM_IMPORT_THEME, IDM_DELETE_THEME
   IDM_FILE, IDM_FILENEW 
   IDM_FILEOPEN, IDM_FILECLOSE, IDM_FILECLOSEALL, IDM_FILECLOSEALLOTHERS
   IDM_CLOSEALLFORWARD, IDM_CLOSEALLBACKWARD, IDM_FILESAVE, IDM_FILESAVEAS
   IDM_FILESAVEALL, IDM_FILESAVEDECLARES
   IDM_MRU, IDM_OPENINCLUDE, IDM_COMMAND, IDM_EXIT
   IDM_EDIT
   IDM_UNDO, IDM_REDO, IDM_CUT, IDM_COPY, IDM_PASTE, IDM_DELETELINE, IDM_DELETE, IDM_INSERTFILE
   IDM_ANSI, IDM_UTF8BOM, IDM_UTF16BOM
   IDM_INDENTBLOCK, IDM_UNINDENTBLOCK, IDM_COMMENTBLOCK, IDM_UNCOMMENTBLOCK
   IDM_DUPLICATELINE, IDM_MOVELINEUP, IDM_MOVELINEDOWN, IDM_TOUPPERCASE, IDM_TOLOWERCASE
   IDM_TOMIXEDCASE, IDM_EOLTOCRLF, IDM_EOLTOCR, IDM_EOLTOLF, IDM_SELECTALL, IDM_SELECTLINE
   IDM_SPACESTOTABS, IDM_TABSTOSPACES
   IDM_SEARCH
   IDM_FIND, IDM_FINDNEXT, IDM_FINDPREV, IDM_FINDNEXTACCEL, IDM_FINDPREVACCEL 
   IDM_FINDINFILES, IDM_REPLACE, IDM_DEFINITION,IDM_LASTPOSITION
   IDM_GOTONEXTTAB, IDM_GOTOPREVTAB, IDM_CLOSETAB, IDM_GOTONEXTFUNCTION, IDM_GOTOPREVFUNCTION
   IDM_GOTO, IDM_BOOKMARKTOGGLE, IDM_BOOKMARKNEXT, IDM_BOOKMARKPREV, IDM_BOOKMARKCLEARALL
   IDM_VIEW
   IDM_FOLDTOGGLE, IDM_FOLDBELOW, IDM_FOLDALL, IDM_UNFOLDALL, IDM_ZOOMIN, IDM_ZOOMOUT
   IDM_FUNCTIONLIST, IDM_VIEWEXPLORER, IDM_VIEWOUTPUT, IDM_RESTOREMAIN
   IDM_PROJECTNEW, IDM_PROJECTMANAGER, IDM_PROJECTOPEN, IDM_MRUPROJECT
   IDM_REMOVEFILEFROMPROJECT 
   IDM_PROJECTCLOSE, IDM_PROJECTSAVE, IDM_PROJECTSAVEAS, IDM_PROJECTFILESADD, IDM_PROJECTOPTIONS  
   IDM_BUILDEXECUTE, IDM_COMPILE, IDM_REBUILDALL, IDM_RUNEXE, IDM_QUICKRUN, IDM_COMMANDLINE
   IDM_NEWFORM, IDM_VIEWTOOLBOX, IDM_MENUEDITOR
   IDM_TOOLBAREDITOR, IDM_STATUSBAREDITOR, IDM_IMAGEMANAGER, IDM_ALIGNLEFTS
   IDM_ALIGNCENTERS, IDM_ALIGNRIGHTS, IDM_ALIGNTOPS, IDM_ALIGNMIDDLES
   IDM_ALIGNBOTTOMS, IDM_SAMEWIDTHS, IDM_SAMEHEIGHTS, IDM_SAMEBOTH 
   IDM_VSPACEDECREASE, IDM_VSPACEREMOVE, IDM_CENTERHORIZ
   IDM_CENTERVERT, IDM_CENTERBOTH, IDM_LOCKCONTROLS
   IDM_OPTIONS
   IDM_BUILDCONFIG, IDM_USERTOOLSDIALOG, IDM_USERSNIPPETS
   IDM_HELP, IDM_HELPWINAPI, IDM_HELPWINFBE, IDM_HELPWINFBX, IDM_CHECKFORUPDATES, IDM_ABOUT
   IDM_SETFILENORMAL, IDM_SETFILEMODULE, IDM_SETFILEMAIN, IDM_SETFILERESOURCE
   IDM_MRUCLEAR, IDM_MRUPROJECTCLEAR
   IDM_CONSOLE, IDM_GUI, IDM_RESOURCE   ' used for compiler directives in code
   IDM_ADDIMAGE, IDM_REMOVEIMAGE, IDM_FORMATIMAGE, IDM_ATTACHIMAGE, IDM_DETACHIMAGE
   IDM_32BIT, IDM_64BIT   ' mainly used for identifying compiler associated with a project
   IDM_USERTOOL   ' + n number of user tools
End Enum

Dim Shared As HIMAGELIST ghImageListNormal
Dim Shared As Long gidxImageOpened, gidxImageClosed, gidxImageBlank, gidxImageCode
dim shared as BOOLEAN gPreparsing      ' T/F we are preparsing one or more \inc files (idx = -2).
dim shared as BOOLEAN gPreparsingChanges  ' T/F preparsing had changes. Flag to save new databases.
dim shared as BOOLEAN gFileLoading     ' T/F only parse Includes during initial file load.
dim shared as BOOLEAN gCompiling       ' T/F to show spinning mouse cursor.

' Create a global bold font that is used in the PropertyList controls combobox and also for
' the label that describes the property name/description.
dim shared as HFONT ghNormalFont, ghBoldFont
dim shared as long gHelpViewerIndex

'  Global window handle for the main form
Dim Shared As HWnd HWND_FRMMAIN, HWND_FRMMAIN_TOOLBAR, HWND_FRMEXPLORER, HWND_FRMRECENT, HWND_FRMOUTPUT
Dim Shared As HMENU HWND_FRMMAIN_TOPMENU   
dim shared as hwnd HWND_FRMMAIN_COMBOBUILDS 

'  Global window handles for some forms 
Dim Shared As HWnd HWND_FRMOPTIONS, HWND_FRMOPTIONSGENERAL, HWND_FRMOPTIONSEDITOR, HWND_FRMOPTIONSCOLORS
Dim Shared As HWnd HWND_FRMOPTIONSCOMPILER, HWND_FRMOPTIONSLOCAL, HWND_FRMOPTIONSKEYWORDS
Dim Shared As HWnd HWND_FRMFINDREPLACE, HWND_FRMFINDINFILES, HWND_FRMVDTOOLBOX, HWND_FRMVDCOLORS
Dim Shared As HWnd HWND_FRMFUNCTIONLIST, HWND_FRMBUILDCONFIG, HWND_FRMUSERTOOLS, HWND_FRMMENUEDITOR
Dim Shared As HWnd HWND_PROPLIST_EDIT, HWND_PROPLIST_COMBO, HWND_PROPLIST_COMBOLIST, HWND_FRMHELPVIEWER
dim shared as hwnd HWND_FRMIMAGES, HWND_FRMSNIPPETS

'  Global handle to hhctrl.ocx for context sensitive help
Dim Shared As Any Ptr gpHelpLib

'  Global holding all full path/name for HTML files linked to Help Treeview (index in lParam)
type HTMLHELPNODES
   wszFilename as CWSTR
   wszLocationURL as CWSTR
   TreeviewNode as HTREEITEM
   hTreeview as HWND
end type
dim shared as HTMLHELPNODES gHTMLHelp(any)

dim shared as HICON ghIconGood, ghIconBad, ghIconTick, ghIconNoTick

dim shared as HACCEL ghAccelUserTools
dim shared as HCURSOR ghCursorNS

' Global string to track the last accessed property/event in the PropertyList. This allows the
' user to quickly sqitch between controls that share common properties like 'Text'.
dim shared as CWSTR gwszPreviousPropName, gwszPreviousEventName

' PropertyList divider globals
dim shared as long gPropDivPos
dim shared as boolean gPropDivTracking

' Create a dynamic array that will hold all localization words/phrases while
' a language is being edited in frmOptionsLocal. Also create a global array
' that holds the english phrases. When a localization is loaded, any missing
' translations are replaced with the english version.
ReDim Shared gLangEnglish(Any) As WString * MAX_PATH
ReDim Shared gLocalPhrases(Any) As WString * MAX_PATH
dim shared gLocalPhrasesEdit as boolean   ' a localization language is being edited. 


' Create a dynamic array that will hold all localization words/phrases. This
' array is resized and loaded using the LoadLocalizationFile function.
ReDim Shared LL(Any) As WString * MAX_PATH

' Define a macro that allows the user to specify the LL array subscript and
' also a descriptive label (that is ignored), and return the LL array value.
#Define L(e,s) LL(e)

#Define SetFocusScintilla  PostMessage HWND_FRMMAIN, MSG_USER_SETFOCUS, 0, 0
#Define ResizeExplorerWindow  PostMessage HWND_FRMEXPLORER, MSG_USER_OPENEDITORS_RESIZE, 0, 0
#Define SciExec(h, m, w, l) SendMessage(h, m, w, CAST(LPARAM, l))


type PreparseTimestamps
   wszFilename as CWSTR
   tFiletime AS LONGLONG
END TYPE
dim shared gPreparseTimestamps(any) as PreparseTimestamps


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
Dim Shared gFind As FINDREPLACE_TYPE
Dim Shared gFindInFiles As FINDREPLACE_TYPE


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


' Forward references
Type clsDocument_ As clsDocument


'' Last position in document. Used when "Last Position" menu option is selected.
Type LASTPOSITION_TYPE
   pDoc       As clsDocument_ Ptr
   nFirstLine As Long     ' first visible line on screen
   nPosition  As Long     ' Position in Scintilla document where caret is positioned
End Type
Dim Shared gLastPosition As LASTPOSITION_TYPE


' Structure that holds all Images found for an individual file.
type IMAGES_TYPE
   wszImageName as CWSTR              ' Based on filename. IMAGE_OPEN, IMAGE_CLOSE, etc
   wszFileName  as CWSTR
   wszFormat    as CWSTR              ' BITMAP, ICON, RCDATA, CURSOR
   pDoc         as clsDocument_ ptr   ' backpointer to pDoc in case search on wszImageName performed.
end type

Enum FontStyles
   Normal    = 0
   Bold      = 1
   Italic    = 2
   Strikeout = 4
   Underline = 8
End Enum

Enum FontCharset
   Default     = DEFAULT_CHARSET
   Ansi        = ANSI_CHARSET
   Arabic      = ARABIC_CHARSET
   Baltic      = BALTIC_CHARSET
   ChineseBig5 = CHINESEBIG5_CHARSET
   EastEurope  = EASTEUROPE_CHARSET
   GB2312      = GB2312_CHARSET
   Greek       = GREEK_CHARSET
   Hangul      = HANGUL_CHARSET
   Hebrew      = HEBREW_CHARSET
   Johab       = JOHAB_CHARSET
   Mac         = MAC_CHARSET
   OEM         = OEM_CHARSET
   Russian     = RUSSIAN_CHARSET
   Shiftjis    = SHIFTJIS_CHARSET
   Symbol      = SYMBOL_CHARSET
   Thai        = THAI_CHARSET
   Turkish     = TURKISH_CHARSET
   Vietnamese  = VIETNAMESE_CHARSET
End Enum

enum
   COLOR_CUSTOM = 1
   COLOR_COLORS
   COLOR_SYSTEM
end enum
    
type clsColors
   private:
   public:
      wszColorName as CWSTR          
      ColorType    as long       ' COLOR_QUICK, COLOR_SYSTEM
      ColorValue   as COLORREF
      declare function SetColor( byref wszColorName as wstring, byval ColorType as long, byval ColorValue as COLORREF) as long
END TYPE
function clsColors.SetColor( byref wszColorName as wstring, byval ColorType as long, byval ColorValue as COLORREF) as Long
   this.wszColorName = wszColorName
   this.ColorType    = ColorType   ' COLOR_QUICK, COLOR_SYSTEM
   this.ColorValue   = ColorValue  ' COLORREF
   function = 0
end function
dim shared gColors(any) as clsColors



' Create an array that holds all options in the Build compiler options listbox. The
' description contains a parenthesis enclosed action.
Dim shared gBuildOptions(...) as CWSTR => _
   {  "Language compatibility FreeBasic (-lang fb)", _
      "Language compatibility FreeBasic Lite (-lang fblite)", _
      "Language compatibility QuickBasic (-lang qb)", _
      "Subsystem to console (-s console)", _
      "Subsystem to GUI (-s gui)", _
      "Compiler assembler backend (-gen gas)", _
      "Compiler GCC backend (-gen gcc)", _
      "Compiler LLVM backend (-gen llvm)", _
      "Create DLL and import library (-dll)", _
      "Create static library (-lib)", _
      "Add error checking (-e)", _
      "Add error checking with RESUME support (-ex)", _
      "Same as -ex with array bounds and null pointer (-exx)", _
      "Add debug information (-g)", _
      "Compile only, do not link (-c)", _
      "Do not delete the object files (-C)", _
      "Emit preprocessed output, do not compile (-pp)" _
   }

' Create an array that holds all options in the replaceable parameters for snippets.
Dim shared gSnippetParameter(...) as CWSTR => _
   {  "CLIPBOARD", _
      "CURSOR_POSITION", _
      "FILENAME", _
      "FILENAME_SHORT", _
      "CURRENT_YEAR", _
      "CURRENT_YEAR_SHORT", _
      "CURRENT_MONTH", _
      "CURRENT_MONTH_NAME", _
      "CURRENT_MONTH_NAME_SHORT", _
      "CURRENT_DAY", _
      "CURRENT_DAY_NAME", _
      "CURRENT_DAY_NAME_SHORT", _
      "CURRENT_HOUR", _
      "CURRENT_MINUTE", _
      "CURRENT_SECOND" _
   }

