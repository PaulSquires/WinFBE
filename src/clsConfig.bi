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
   IsAlt            as Long
   IsShift          as Long
   IsPromptRun      as Long
   IsMinimized      as Long
   IsWaitFinish     as Long
   IsDisplayMenu    as Long
   Action           as long 
END TYPE


type TYPE_SNIPPETS
   wszDescription as CWSTR
   wszTrigger     as CWSTR
   wszCode        as CWSTR
END TYPE


type TYPE_BUILDS
   id             as string    ' GUID
   wszDescription as CWSTR
   IsDefault      as Long      ' 0:False, 1:True
   Is32bit        as Long      ' 0:False, 1:True
   Is64bit        as Long      ' 0:False, 1:True
   wszOptions     as CWSTR     ' Compiler options (manual and selected from listbox)
END TYPE

' Used for Themes.
Type TYPE_COLORS
   nFg As COLORREF
   nBg As COLORREF
   ' bUseDefaultBg is currently not used. Do not use BOOLEAN because we don't 
   ' want the words true/false outputted to the ini file.
   bUseDefaultBg as Long = 0  
   bFontBold   as Long = 0
   bFontItalic as Long = 0
   bFontUnderline as long = 0
End Type

type TYPE_THEMES
   id             as string    ' GUID
   wszDescription as CWSTR
   colors(CLR_CARET to CLR_WINDOW) as TYPE_COLORS
END TYPE


Type clsConfig
   Private:
      _ConfigFilename            As CWSTR 
      _SnippetsFilename          as CWSTR
      _FBKeywordsFilename        as CWSTR 
      _FBCodetipsFilename        as CWSTR
      _WinAPICodetipsFilename    as CWSTR 
      _WinFormsXCodetipsFilename as CWSTR 
      _WinFBXCodetipsFilename    as CWSTR
      
   Public:
      WinFBEversion        as CWSTR
      SelectedTheme        as string          ' GUID of selected theme
      idxTheme             as long            ' need global b/c can't GetCurSel from CBN_EDITCHANGE
      Themes(any)          as TYPE_THEMES
      ThemesTemp(any)      as TYPE_THEMES
      Tools(any)           as TYPE_TOOLS
      ToolsTemp(any)       as TYPE_TOOLS  
      Builds(any)          as TYPE_BUILDS  
      BuildsTemp(any)      as TYPE_BUILDS  
      Snippets(any)        as TYPE_SNIPPETS
      SnippetsTemp(any)    as TYPE_SNIPPETS  
      rcSnippets           as rect                 ' Snippet window position (nt saved to file)
      FBKeywords           As String
      bKeywordsDirty       As BOOLEAN = True       ' not saved to file
      AskExit              As Long = false         ' use Long so True/False string not written to file
      CheckForUpdates      As Long = True
      LastUpdateCheck      as long = 0             ' Julian date of last update check
      HideToolbar          as long = false
      HideStatusbar        as long = false
      CloseFuncList        As Long = True
      ShowExplorer         As Long = True
      ShowExplorerWidth    As Long = 250
      SyntaxHighlighting   As Long = True
      Codetips             As Long = True
      AutoComplete         As Long = True
      RightEdge            As Long = TRUE
      RightEdgePosition    As CWSTR = "80"
      LeftMargin           As Long = True
      FoldMargin           As Long = false
      AutoIndentation      As Long = True
      ConfineCaret         As Long = true
      LineNumbering        As Long = True
      HighlightCurrentLine As Long = True
      IndentGuides         As Long = True
      TabIndentSpaces      As Long = True
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
      WinFBXHelpfile       As CWSTR
      WinFBXPath           as CWSTR
      RunViaCommandWindow  As Long = False
      MRU(9)               As CWSTR
      MRUProject(9)        As CWSTR
      
      Declare Constructor()
      declare function ImportTheme( byref st as wstring, byval bImportExternal as Boolean = false ) as Long
      declare Function GetThemePtr() as TYPE_THEMES ptr
      Declare Function LoadKeywords() As Long
      Declare Function SaveKeywords() As Long
      Declare Function WriteMRU() As Long
      Declare Function WriteMRUProjects() As Long
      Declare Function SaveConfigFile() As Long
      Declare Function LoadConfigFile() As Long
      declare Function LoadSnippets() as Long
      declare Function SaveSnippets() as Long
      declare Function InitializeToolBox() as Long
      Declare Function ProjectSaveToFile() As BOOLEAN    
      declare Function ProjectLoadFromFile( byval wszFile as CWSTR ) As BOOLEAN    
      declare Function LoadCodetipsFB() as boolean
      declare Function LoadCodetipsWinAPI() as boolean
      declare Function LoadCodetipsWinForms( byval wszFilename as CWSTR ) as boolean
      declare Function LoadCodetipsWinFormsX() as boolean
      declare Function LoadCodetipsWinFBX() as boolean
      declare Function LoadCodetipsGeneric( byval wszFilename as CWSTR, byval IsWinAPI as boolean ) as boolean
End Type
