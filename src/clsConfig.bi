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
   CTRL_PICTUREBOX
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
   
type TYPE_TOOLS
   wszDescription   as CWSTR
   wszCommand       as CWSTR
   wszParameters    as CWSTR
   wszKey           as CWSTR
   wszWorkingFolder as CWSTR
   IsCtrl           as long
   IsAlt            as long
   IsShift          as long
   IsPromptRun      as long
   IsMinimized      as long
   IsWaitFinish     as long
   IsDisplayMenu    as long
   Action           as long 
end type


type TYPE_SNIPPETS
   wszDescription as CWSTR
   wszTrigger     as CWSTR
   wszCode        as CWSTR
end type


type TYPE_BUILDS
   id             as string    ' GUID
   wszDescription as CWSTR
   IsDefault      as long      ' 0:False, 1:True
   Is32bit        as long      ' 0:False, 1:True
   Is64bit        as long      ' 0:False, 1:True
   wszOptions     as CWSTR     ' Compiler options (manual and selected from listbox)
end type

type TYPE_CATEGORIES
   idFileType     as string    ' GUID or special node value (FILETYPE_*)
   wszDescription as CWSTR
   hNodeExplorer  as HTREEITEM
   bShow          as boolean = true
end type

' NOTE: These node types are different values than the FileType defines from
' the clsDocument.bi file so we could not reuse those equates. These nodetype
' equates defined the order in which the files will be displayed in the 
' explorer listbox.
 #define CATINDEX_FILES             0
 #define CATINDEX_MAIN              1
 #define CATINDEX_RESOURCE          2
 #define CATINDEX_HEADER            3
 #define CATINDEX_MODULE            4
 #define CATINDEX_NORMAL            5

' Structure used to save codetip cache database information to disk. This
' data is checked when loading the codetip cache to see if any of the original
' codetip files had changed since the cache was created. If yes, then that
' codetip file needs to be reparsed.
type CODETIP_META_DATA
   nFiletype    as long           ' refer to DB2_FILETYPE_*  (filenames are not stored)
   DateFileTime as FILETIME       ' DateTime of original codetip file
   filler(1024) as ubyte          ' extra space for possible future expansion
end type


type clsConfig
   Private:
      _ConfigFilename            as CWSTR 
      _SnippetsFilename          as CWSTR
      _SnippetsDefaultFilename   as CWSTR
      _FBKeywordsFilename        as CWSTR 
      _WinApiKeywordsFilename    as CWSTR 
      _FBKeywordsDefaultFilename as CWSTR 
      _FBCodetipsFilename        as CWSTR
      _WinAPICodetipsFilename    as CWSTR 
      _WinFormsXCodetipsFilename as CWSTR 
      _WinFBXCodetipsFilename    as CWSTR
      _DateFileTime              as FILETIME
      
   public:
      WinFBEversion         as CWSTR
      Tools(any)            as TYPE_TOOLS
      ToolsTemp(any)        as TYPE_TOOLS  
      Builds(any)           as TYPE_BUILDS  
      BuildsTemp(any)       as TYPE_BUILDS  
      Cat(any)              as TYPE_CATEGORIES
      CatTemp(any)          as TYPE_CATEGORIES
      Snippets(any)         as TYPE_SNIPPETS
      SnippetsTemp(any)     as TYPE_SNIPPETS  
      rcSnippets            as rect                 ' Snippet window position (not saved to file)
      FBKeywords            as string
      WinApiKeywords        as string
      bKeywordsDirty        as boolean = true       ' not saved to file
      AskExit               as long = false         ' use long so true/False string not written to file
      CheckForUpdates       as long = true
      EnableProjectCache    as long = true          ' Fast project cache
      LastUpdateCheck       as long = 0             ' Julian date of last update check
      AutoSaveFiles         as long = true
      AutoSaveInterval      as long = 10            ' seconds between autosave checks
      idAutoSaveTimer       as long = 999           ' id of Autosave timer
      RestoreSession        as long = false
      wszLastActiveSession  as CWSTR
      CloseFuncList         as long = true
      ShowPanel             as long = true
      ShowPanelWidth        as long = 250
      SyntaxHighlighting    as long = true
      Codetips              as long = true
      AutoComplete          as long = true
      CharacterAutoComplete as long = false
      RightEdge             as long = false
      RightEdgePosition     as CWSTR = "80"
      LeftMargin            as long = true
      FoldMargin            as long = false
      AutoIndentation       as long = true
      ForNextVariable       as long = false
      ConfineCaret          as long = true
      LineNumbering         as long = true
      HighlightCurrentLine  as long = true
      IndentGuides          as long = false
      PositionMiddle        as long = false         ' position found text to middle of screen
      BraceHighlight        as long = false
      OccurrenceHighlight   as long = false
      TabIndentSpaces       as long = true
      MultipleInstances     as long = true
      CompileAutosave       as long = true
      UnicodeEncoding       as long = false
      TabSize               as CWSTR = "3"
      LocalizationFile      as CWSTR = "english.lang"
      EditorFontname        as CWSTR = "Consolas"
      EditorFontCharSet     as CWSTR = "Default"
      EditorFontsize        as CWSTR = "11"
      FontExtraSpace        as CWSTR = "10"
      ThemeFilename         as CWSTR = "winfbe_default_dark.theme"
      KeywordCase           as long = 3  ' "Original Case"
      StartupLeft           as long = 0
      StartupTop            as long = 0
      StartupRight          as long = 0
      StartupBottom         as long = 0
      StartupMaximized      as long = false
      ToolBoxLeft           as long = 0
      ToolBoxTop            as long = 0
      ToolBoxRight          as long = 0
      ToolBoxBottom         as long = 0
      FBWINCompiler32       as CWSTR
      FBWINCompiler64       as CWSTR
      CompilerBuild         as CWSTR     ' Build GUID
      CompilerSwitches      as CWSTR
      CompilerHelpfile      as CWSTR
      WinFBXHelpfile        as CWSTR
      WinFBXPath            as CWSTR
      RunViaCommandWindow   as long = false
      DisableCompileBeep    as long = false
      MRU(9)                as CWSTR
      MRUProject(9)         as CWSTR
      
      declare constructor()
      declare function SetCategoryDefaults() as long
      declare function LoadKeywords() as long
      declare function SaveKeywords() as long
      declare function WriteMRU() as long
      declare function WriteMRUProjects() as long
      declare function SaveConfigFile() as long
      declare function LoadConfigFile() as long
      declare function LoadSnippets() as long
      declare function SaveSnippets() as long
      declare function InitializeToolBox() as long
      declare function SaveSessionFile( byref wszSessionFile as wstring ) as boolean    
      declare function LoadSessionFile( byref wszSessionFile as wstring ) as boolean    
      declare function ProjectSaveToFile() as boolean    
      declare function ProjectLoadFromFile( byval wszFile as CWSTR ) as boolean    
      declare function LoadCodetipsFB() as boolean
      declare function LoadCodetipsWinAPI() as boolean
      declare function LoadCodetipsWinForms( byval wszFilename as CWSTR ) as boolean
      declare function LoadCodetipsWinFormsX() as boolean
      declare function LoadCodetipsWinFBX() as boolean
      declare function LoadCodetipsGeneric( byval wszFilename as CWSTR, byval nFiletype as long) as boolean
      declare function LoadCodetips() as long
      declare function ReloadConfigFileTest() as boolean    
end type
