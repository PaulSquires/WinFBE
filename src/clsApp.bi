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


Type clsApp
   Private: 
      m_arrQuickRun(Any) As WSTRING * MAX_PATH
      
   Public:
      pDocList                   As clsDocument Ptr   ' Single linked list of loaded files
      IsWindowIncludes           as Boolean           ' T/F that Windows includes have already been loaded
      SuppressNotify             As BOOLEAN           ' temporarily suppress Scintilla notifications
      hRecentFilesRootNode       As HTREEITEM
      hRecentProjectsRootNode    As HTREEITEM
      hExplorerRootNode          As HTREEITEM
      hExplorerNormalNode        As HTREEITEM
      hExplorerMainNode          As HTREEITEM
      hExplorerResourceNode      As HTREEITEM
      hExplorerHeaderNode        As HTREEITEM
      hExplorerModuleNode        As HTREEITEM
      bDragTabActive             as BOOLEAN           ' a tab in the top tabcontrol is being dragged
      bDragActive                As BOOLEAN           ' splitter drag is currently active 
      hWndPanel                  As HWND              ' the panel being split left/right or up/down
      IncludeFilename            As CWSTR
      NonProjectNotes            as CWSTR             ' Save/load from config file
      wszPanelText               as CWSTR             ' Current file loading or being compiled (for statusbar updating)
      hIconPanel                 as HANDLE            ' Success/failure of most previous compile (for Statusbar updating)
      FileLoadingCount           as long              ' Track count of files loading for statusbar display
      NewProjectTemplateType     as long              ' IDC of the new project type to create.
      IsNewProjectFlag           As BOOLEAN
      IsProjectLoading           as Boolean           ' Project loading. Disable some screen updating.
      IsFileLoading              as Boolean           ' File loading. Disable some screen updating.
      IsCompiling                as Boolean           ' File/Project currently being compiled (spinning mouse cursor).
      IsShutDown                 as boolean           ' App is currently closing
      wszCommandLine             as CWSTR             ' non-project commandline (not saved to file)
      
      pImageAutocompleteBaseType as any ptr           ' image used in scintilla popup autocomplete
      pImageAutocompleteClass    as any ptr           ' image used in scintilla popup autocomplete
      pImageAutocompleteMethod   as any ptr           ' image used in scintilla popup autocomplete
      pImageAutocompleteProperty as any ptr           ' image used in scintilla popup autocomplete
      hWndAutoCListBox           as hwnd              ' handle of popup autocomplete ListBox window
            
      IsProjectActive            As boolean
      ProjectBuild               As string            ' default build configuration for the project (GUID)
      ProjectName                As CWSTR
      ProjectFilename            As CWSTR
      ProjectOther32             As CWSTR             ' compile flags 32 bit compiler
      ProjectOther64             As CWSTR             ' compile flags 64 bit compiler
      ProjectNotes               as CWSTR             ' Save/Load from project file
      ProjectCommandLine         as CWSTR
      ProjectManifest            as long              ' T/F create a generic resource and manifest file

      ' Global string to track the last accessed property/event in the PropertyList. This allows the
      ' user to quickly sqitch between controls that share common properties like 'Text'.
      PreviousPropName           as CWSTR
      PreviousEventName          as CWSTR

      declare function AddQuickRunEXE( byref sFilename as wstring ) as Long
      declare function CheckQuickRunEXE() as Long
      declare Function RemoveAllSelectionAttributes() As long
      Declare Function AddDocument( ByVal pDoc As clsDocument Ptr ) As Long
      Declare Function RemoveDocument( ByVal pDoc As clsDocument Ptr ) As Long
      declare Function RemoveAllDocuments() As Long
      Declare Function GetDocumentCount() As Long
      declare Function GetDocumentPtrByWindow( byval hWindow as hwnd) As clsDocument ptr
      Declare Function GetDocumentPtrByFilename( ByVal pswzName As WString Ptr ) As clsDocument Ptr
      Declare Function GetMainDocumentPtr() As clsDocument Ptr
      Declare Function GetResourceDocumentPtr() As clsDocument Ptr
      declare function GetSourceDocumentPtr( byval pDocIn as clsDocument ptr ) As clsDocument Ptr
      declare function GetHeaderDocumentPtr( byval pDocIn as clsDocument ptr ) As clsDocument Ptr
      Declare Function SaveProject( ByVal bSaveAs As BOOLEAN = False ) As BOOLEAN
      Declare Function ProjectSetFileType( ByVal pDoc As clsDocument Ptr, ByVal nFileType As Long ) As LRESULT
      declare Function GetProjectCompiler() As long
      
End Type


