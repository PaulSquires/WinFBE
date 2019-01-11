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

'  Scintilla Control identifiers
#Define IDC_SCINTILLA              100
#Define IDC_SCROLLV                200

' File encodings
#define FILE_ENCODING_ANSI         0
#define FILE_ENCODING_UTF8_BOM     1
#define FILE_ENCODING_UTF16_BOM    2

#define FILETYPE_UNDEFINED         0
#define FILETYPE_MAIN              1
#define FILETYPE_MODULE            2
#define FILETYPE_NORMAL            3
#define FILETYPE_RESOURCE          4
#define FILETYPE_HEADER            5


#include once "clsMenuItem.bi"
#include once "clsControl.bi"      ' Includes properties and events types
#include once "clsCollection.bi"

' Structure that holds all of the user embedded compiler directives
' in the source code. Currently, only the main source file is searched
' for the '#CONSOLE ON|OFF directive but others can be added as needed.
type COMPILE_DIRECTIVES
   DirectiveFlag as long              ' IDM_GUI, IDM_CONSOLE, IDM_RESOURCE
   DirectiveText as String            ' reource filename
END TYPE

' Forward references
Type clsDocument_ As clsDocument

' Structure that holds all Images found for an individual file.
type IMAGES_TYPE
   wszImageName as CWSTR              ' Based on filename. IMAGE_OPEN, IMAGE_CLOSE, etc
   wszFileName  as CWSTR
   wszFormat    as CWSTR              ' BITMAP, ICON, RCDATA, CURSOR
   pDoc         as clsDocument_ ptr   ' backpointer to pDoc in case search on wszImageName performed.
end type

Type clsDocument
   Private:
      ' 2 Scintilla direct pointers to accommodate split editing
      m_pSci(1)             As Any Ptr      
      m_hWndActiveScintilla as hwnd
      
   Public:
      pDocNext         as clsDocument_ ptr = 0  ' pointer to next document in linked list 
      IsDesigner       As BOOLEAN
      IsNewFlag        As BOOLEAN
      LoadingFromFile  as Boolean
      
      ' 2 Scintilla controls to accommodate split editing
      ' hWindow(0) is our MAIN control (bottom)
      ' hWindow(1) is our split control (top)
      hWindow(1)       As HWnd   ' Scintilla split edit windows 
      
      ' Visual designer related
      MenuItems(any)   as clsMenuItem
      Controls         as clsCollection
      hWndDesigner     as HWnd      ' DesignMain window (switch to this window when in design mode (versus code mode)
      hDesignTabCtrl   as HWnd      ' TabCtrl to switch between Design/Code
      hWndFrame        as hwnd      ' DesignFrame for visual designer windows
      hWndForm         as hwnd      ' DesignForm for visual designer windows
      hWndFakeMenu     as HWND      ' Fake top menu to display when using Menu Editor
      hFontFakeMenu    as HFONT     ' System font used for menus
      ErrorOffset      as long      ' Number of lines to account for when error thrown for visual designer code files.
      GrabHit          as long      ' Which grab handle is currently active for sizing action
      ptPrev           as point     ' Used for sizing action
      bSizing          as Boolean   ' Flag that sizing action is in progress
      bMoving          as Boolean   ' Flag that moving action is in progress
      bRegenerateCode  as Boolean   ' Flag to regenerate code when switching to the code tab
      bLockControls    as Boolean   ' Global flag that locks the form and all controls from moving or resizing.
      rcSize           as RECT      ' Current size of form/control. Used during sizing action
      pCtrlAction      as clsControl ptr  ' The control that the size/move action is being performed on
      pCtrlCloseLeft   as clsControl ptr  ' closest control to the left of selected control
      pCtrlCloseTop    as clsControl ptr  ' closest control to the top of selected control
      pCtrlCloseRight  as clsControl ptr  ' closest control to the right of selected control
      pCtrlCloseBottom as clsControl ptr  ' closest control to the bottom of selected control
      ptBlueStart      as POINT     ' Start of blue line for snapping
      ptBlueEnd        as POINT     ' End of blue line for snapping
      SnapUpWait       as long      ' #pixels of movement to wait until snap operation ends
      wszFormCodeGen   as CWSTR     ' Form code generated  
      wszFormMetaData  as CWSTR     ' Form metadata that defines the form
      AllImages(any)   as IMAGES_TYPE ' All Images belonging to the Form
      
      ' Code document related
      ProjectFileType  As Long = FILETYPE_UNDEFINED
      DiskFilename     As WString * MAX_PATH
      DateFileTime     As FILETIME  
      hNodeExplorer    As HTREEITEM
      FileEncoding     as long       
      bNeedsParsing    as Boolean  ' Document requires to be parsed due to changes.
      UserModified     as boolean  ' occurs when user manually changes encoding state so that document will be saved in the new format
      DeletedButKeep   as boolean  ' file no longer exists but keep open anyway
      DocumentBuild    as string   ' specific build configuration to use for this document
      sMatchWord       as string   ' for the incremental autocomplete search
      AutoCompleteType as long     ' AUTOC_DIMAS, AUTOC_TYPE
      AutoCStartPos    as Long
      
      ' Following used for split edit views
      hScrollBar       as hwnd
      ScrInfo          As SCROLLINFO  ' Scrollbar parameters array
      rcSplitButton    as RECT        ' Split gripper vertical for Scintilla window
      SplitY           As long        ' Y coordinate of vertical splitter
      
      static NextFileNum as Long
      
      declare property hWndActiveScintilla() as hwnd
      declare property hWndActiveScintilla(byval hWindow as hwnd)
      
      declare function MainMenuExists() as Boolean
      declare function GetActiveScintillaPtr() as any ptr
      Declare Function CreateCodeWindow( ByVal hWndParent As HWnd, ByVal IsNewFile As BOOLEAN, ByVal IsTemplate As BOOLEAN = False, ByVal pwszFile As WString Ptr = 0) As HWnd
      declare Function CreateDesignerWindow( ByVal hWndParent As HWnd) As HWnd   
      Declare Function FindReplace( ByVal strFindText As String, ByVal strReplaceText As String ) As Long
      Declare Function InsertFile() As BOOLEAN
      declare function ParseFormMetaData( ByVal hWndParent As HWnd, byref sAllText as wstring ) as CWSTR
      Declare Function SaveFile(ByVal bSaveAs As BOOLEAN = False) As BOOLEAN
      Declare Function ApplyProperties() As Long
      Declare Function GetTextRange( ByVal cpMin As Long, ByVal cpMax As Long) As String
      Declare Function ChangeSelectionCase( ByVal fCase As Long) As Long 
      Declare Function GetCurrentLineNumber() As Long
      Declare Function SelectLine( ByVal nLineNum As Long ) As Long
      Declare Function GetLine( ByVal nLine As Long) As String
      declare function IsFunctionLine( byval lineNum as long ) as long
      declare function GotoNextFunction() as long
      declare function GotoPrevFunction() as long
      Declare Function GetLineCount() As long
      Declare Function GetSelText() As String
      Declare Function GetText() As String
      Declare Function SetText( ByRef sText As Const String ) As Long 
      declare Function SetLine( ByVal nLineNum As Long, byval sNewText as string) As long
      declare Function AppendText( ByRef sText As Const String ) As Long 
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
      Declare Function TabsToSpaces() As Long
      Declare Function GetWord( ByVal curPos As Long = -1 ) As String
      Declare Function GetBookmarks() As String
      Declare Function SetBookmarks( ByVal sBookmarks As String ) As Long
      declare Function GetCurrentFunctionName() As string
      declare Function LineDuplicate() As Long
      declare function SetMarkerHighlight() As Long
      declare Function RemoveMarkerHighlight() As Long
      declare Function IsMultiLineSelection() As boolean
      declare Function HasMarkerHighlight() As BOOLEAN
      declare Function FirstMarkerHighlight() As long
      declare Function LastMarkerHighlight() As long
      declare function CompileDirectives( Directives() as COMPILE_DIRECTIVES) as Long
      Declare Destructor
End Type
dim clsDocument.NextFileNum as long = 0
