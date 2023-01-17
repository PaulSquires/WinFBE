'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2022 Paul Squires, PlanetSquires Software
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT any WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS for A PARTICULAR PURPOSE.  See the
'    GNU General public License for more details.

#pragma once

'  Scintilla Control identifiers
#define IDC_SCINTILLA              100

' File encodings
#define FILE_ENCODING_ANSI         0
#define FILE_ENCODING_UTF8_BOM     1
#define FILE_ENCODING_UTF16_BOM    2

#define FILETYPE_UNDEFINED        "0"
#define FILETYPE_MAIN             "1"
#define FILETYPE_MODULE           "2"
#define FILETYPE_NORMAL           "3"
#define FILETYPE_RESOURCE         "4"
#define FILETYPE_HEADER           "5"


#include once "clsMenuItem.bi"
#include once "clsToolBarItem.bi"
#include once "clsPanelItem.bi"
#include once "clsControl.bi"      ' Includes properties and events types
#include once "clsCollection.bi"

' Structure that holds all of the user embedded compiler directives
' in the source code. Currently, only the main source file is searched
' for the '#CONSOLE ON|OFF directive but others can be added as needed.
type COMPILE_DIRECTIVES
   DirectiveFlag as long              ' IDM_GUI, IDM_CONSOLE, IDM_RESOURCE
   DirectiveText as string            ' reource filename
end type

' Forward references
type clsDocument_ as clsDocument

' Structure that holds all Images found for an individual file.
type IMAGES_TYPE
   wszImageName as CWSTR              ' Based on filename. IMAGE_OPEN, IMAGE_CLOSE, etc
   wszFileName  as CWSTR
   wszFormat    as CWSTR              ' BITMAP, ICON, RCDATA, CURSOR
   pDoc         as clsDocument_ ptr   ' backpointer to pDoc in case search on wszImageName performed.
end type

' Enum used to distinguish bewteen a sub/function and Property Get/Set
enum ClassProperty
   None   = 0   ' sub/function
   Getter = 1
   Setter = 2
   ctor   = 3   ' constructor
   dtor   = 4   ' destructor
end enum

' type that holds data for all project files as it they are loaded from
' the project file.
type PROJECT_FILELOAD_DATA
   wszFilename    as CWSTR        ' full path and filename
   wszFiletype    as CWSTR        ' pDoc->ProjectFileType
   bLoadInTab     as boolean
   wszBookmarks   as CWSTR        ' pDoc->GetBookmarks()
   wszFoldPoints  as CWSTR        ' pDoc->GetFoldPoints()
   IsDesigner     as boolean
   IsDesignerView as long         
   nFirstLine     as long         ' first line of main view 
   nPosition      as long         ' current position of main view
   nFirstLine1    as long         ' first line of second view 
   nPosition1     as long         ' current position of second view
   nSplitPosition as long         ' pDoc->SplitY
   nFocusEdit     as long         ' View 0 or View 1
end type


type clsDocument
   private:
      ' 2 Scintilla direct pointers to accommodate split editing
      m_pSci(1)             as any ptr      
      m_hWndActiveScintilla as hwnd
      m_UserModified        as boolean  ' occurs when user manually changes state not captured by Scintilla changes
      
   public:
      pDocNext          as clsDocument_ ptr = 0  ' pointer to next document in linked list 
      IsDesigner        as boolean
      IsNewFlag         as boolean
      LoadingFromFile   as boolean
      
      docData           as PROJECT_FILELOAD_DATA    ' loaded from project files
      
      ' 2 Scintilla controls to accommodate split editing
      ' hWindow(0) is our MAIN control (bottom)
      ' hWindow(1) is our split control (top)
      hWindow(1)        as HWnd   ' Scintilla split edit windows 
      
      ' Visual designer related
      wszFormVersion    as CWSTR   
      MenuItems(any)    as clsMenuItem
      ToolBarItems(any) as clsToolBarItem
      wszToolBarSize    as CWSTR = wstr("SIZE_24")  ' SIZE_16, SIZE_24, SIZE_32, SIZE_48
      PanelItems(any)   as clsPanelItem
      Controls          as clsCollection
      AllImages(any)    as IMAGES_TYPE     ' All Images belonging to the Form
      GenerateMenu      as boolean = true  ' Indicates to generate code for the menu
      GenerateToolBar   as boolean = true  ' Indicates to generate code for the menu
      GenerateStatusBar as boolean = true  ' Indicates to generate code for the statusbar
      hWndDesigner      as HWnd            ' DesignMain window (switch to this window when in design mode (versus code mode)
      DesignTabsCurSel  as long 
      initialCtrlID     as long            ' The starting CtrlID to use for this form and all controls on it.
      hWndFrame         as hwnd            ' DesignFrame for visual designer windows
      hWndForm          as hwnd            ' DesignForm for visual designer windows
      hWndFakeMenu      as HWND            ' Fake top menu to display when using Menu Editor
      hFontFakeMenu     as HFONT           ' System font used for menus
      hWndStatusBar     as HWND            ' StatusBar for the form using StatusBar Editor
      hWndRebar         as HWND            ' Rebar for the form using ToolBar Editor
      hWndToolBar       as HWND            ' ToolBar for the form using ToolBar Editor
      GrabHit           as long            ' Which grab handle is currently active for sizing action
      ptPrev            as point           ' Used for sizing action
      bSizing           as boolean         ' Flag that sizing action is in progress
      bMoving           as boolean         ' Flag that moving action is in progress
      bRegenerateCode   as boolean         ' Flag to regenerate code when switching to the code tab
      bLockControls     as boolean         ' Global flag that locks the form and all controls from moving or resizing.
      rcSize            as RECT            ' Current size of form/control. Used during sizing action
      pCtrlAction       as clsControl ptr  ' The control that the size/move action is being performed on
      wszFormCodeGen    as CWSTR           ' Form code generated  
      wszFormMetaData   as CWSTR           ' Form metadata that defines the form
      AppRunCount       as long = 0        ' Only one should exist in the whole project so track if one or more exists in the code.
                  
      ' SnapLines       
      bSnapLines        as boolean = true  ' Enable/Disable SnapLines
      hBrushSnapLine    as HBRUSH
      hSnapLine(3)      as HWND      ' top, bottom, left, right (ENUM SnapLinePosition)
      bSnapActive(3)    as boolean
      ptCursorStart(3)  as POINT     ' Client coordinate of cursor at time of snap
      
      ' Code document related
      ProjectFiletype   as CWSTR = FILETYPE_UNDEFINED
      DiskFilename      as wstring * MAX_PATH
      DesignerFilename  as wstring * MAX_PATH
      AutoSaveFilename  as wstring * MAX_PATH    '#filename#
      AutoSaveRequired  as boolean
      DateFileTime      as FILETIME  
      bBookmarkExpanded as boolean = true     ' Bookmarks list expand/collapse state
      bFunctionsExpanded as boolean = true    ' Functions list expand/collapse state
      bHasFunctions     as boolean = false    ' FunctionList to determine if click will display the File
      FileEncoding      as long = FILE_ENCODING_ANSI       
      bNeedsParsing     as boolean     ' Document requires to be parsed due to changes.
      DeletedButKeep    as boolean     ' file no longer exists but keep open anyway
      DocumentBuild     as string      ' specific build configuration to use for this document
      sMatchWord        as string      ' for the incremental autocomplete search
      AutoCompletetype  as long        ' AUTOC_DIMAS, AUTOC_TYPE
      AutoCStartPos     as long
      AutoCTriggerStartPos as long
      lastCaretPos      as long        ' used for checking in SCN_UPDATEUI
      lastXOffsetPos    as long        ' used for checking in SCN_UPDATEUI (horizontal offset)
      LastCharTyped     as long        ' most recent entered character. Used to test for BACKSPACE resetting the autocomplete popup.

      ' Following used for split edit views
      rcSplitButton     as RECT        ' Split gripper vertical for Scintilla window
      SplitY            as long        ' Y coordinate of vertical splitter
      bEditorIsSplit    as boolean
      
      static NextFileNum as long
      
      declare property hWndActiveScintilla() as hwnd
      declare property hWndActiveScintilla(byval hWindow as hwnd)
      
      declare property UserModified() as boolean
      declare property UserModified( byval nModified as boolean )
      
      declare function ParseDocument() as boolean
      declare function MainMenuExists() as boolean
      declare function ToolBarExists() as boolean
      declare function StatusBarExists() as boolean
      declare function IsValidScintillaID( byval idScintilla as long ) as boolean
      declare function GetActiveScintillaPtr() as any ptr
      declare function CreateCodeWindow( byval hWndParent as HWnd, byval IsNewFile as boolean, byval IsTemplate as boolean = False, byref wszFileName as wstring = "") as HWnd
      declare function CreateDesignerWindow( byval hWndParent as HWnd) as HWnd   
      declare function FindReplace( byval strFindText as string, byval strReplaceText as string ) as long
      declare function InsertFile() as boolean
      declare function ParseFormMetaData( byval hWndParent as HWnd, byref sAllText as wstring, byval bLoadOnly as boolean = false ) as CWSTR
      declare function LoadFormJSONdata( byval hWndParent as HWnd, byref wszAllText as string, byval bLoadOnly as boolean = false ) as long
      declare function SaveFormJSONdata() as boolean
      declare function SaveFile(byval bSaveAs as boolean = False, byval bAutoSaveOnly as boolean = false) as boolean
      declare function ApplyProperties() as long
      declare function GetTextRange( byval cpMin as long, byval cpMax as long) as string
      declare function ChangeSelectionCase( byval fCase as long) as long 
      declare function GetCurrentLineNumber() as long
      declare function SelectLine( byval nLineNum as long ) as long
      declare function GetLine( byval nLine as long) as string
      declare function IsFunctionLine( byval lineNum as long ) as long
      declare function GotoNextFunction() as long
      declare function GotoPrevFunction() as long
      declare function GetLineCount() as long
      declare function GetSelText() as string
      declare function GetText() as string
      declare function SetText( byref sText as const string ) as long 
      declare function SetLine( byval nLineNum as long, byval sNewText as string) as long
      declare function AppendText( byref sText as const string ) as long 
      declare function CenterCurrentLine() as long 
      declare function GetSelectedLineRange( byref startLine as long, byref endLine as long, byref startPos as long, byref endPos as long ) as long 
      declare function BlockComment( byval flagBlock as boolean ) as long
      declare function CurrentLineUp() as long
      declare function CurrentLineDown() as long
      declare function MoveCurrentLines( byval flagMoveDown as boolean ) as long
      declare function NewLineBelowCurrent() as long
      declare function ToggleBookmark( byval nLine as long ) as long
      declare function NextBookmark() as long 
      declare function PrevBookmark() as long 
      declare function FoldToggle( byval nLine as long ) as long
      declare function FoldAll() as long
      declare function UnFoldAll() as long
      declare function FoldToggleOnwards( byval nLine as long) as long
      declare function ConvertEOL( byval nMode as long) as long
      declare function TabsToSpaces() as long
      declare function GetWord( byval curPos as long = -1 ) as string
      declare function GetBookmarks() as string
      declare function SetBookmarks( byval sBookmarks as string ) as long
      declare function GetFoldPoints() as string
      declare function SetFoldPoints( byval sFoldPoints as string ) as long
      declare function GetCurrentFunctionName( byref sFunctionName as string, byref nGetSet as ClassProperty ) as long
      declare function LineDuplicate() as long
      declare function SetMarkerHighlight() as long
      declare function RemoveMarkerHighlight() as long
      declare function IsMultiLineSelection() as boolean
      declare function HasMarkerHighlight() as boolean
      declare function FirstMarkerHighlight() as long
      declare function LastMarkerHighlight() as long
      declare function LinesPerPage( byval idxWindow as long ) as long
      declare function CompileDirectives( Directives() as COMPILE_DIRECTIVES) as long
      declare destructor
end type
dim clsDocument.NextFileNum as long = 0
