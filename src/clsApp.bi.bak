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

type clsApp
   private: 
      m_arrQuickRun(any) as wstring * MAX_PATH
      
   public:
      pDocList                   as clsDocument ptr   ' Single linked list of loaded files
      isWineActive               as boolean
      pfnCreateLexerfn           as CreateLexerFn
      IsWindowIncludes           as boolean           ' T/F that Windows includes have already been loaded
      PreventActivateApp         as boolean           ' temporarily suppress WM_ACTIVATEAPP (used in 3.02 form file upgrade)
      SuppressNotify             as boolean           ' temporarily suppress Scintilla notifications
      bDragTabActive             as boolean           ' a tab in the top tabcontrol is being dragged
      bDragActive                as boolean           ' splitter drag is currently active 
      hWndPanel                  as HWND              ' the panel being split left/right or up/down
      IncludeFilename            as CWSTR
      NonProjectNotes            as CWSTR             ' Save/load from config file
      wszPanelText               as CWSTR             ' Current file loading or being compiled (for statusbar updating)
      hIconPanel                 as long              ' Success/failure of most previous compile (for Statusbar updating)
      FileLoadingCount           as long              ' Track count of files loading for statusbar display
      NewProjectTemplatetype     as long              ' IDC of the new project type to create.
      IsNewProjectFlag           as boolean
      IsProjectLoading           as boolean           ' Project loading. Disable some screen updating.
      IsFileLoading              as boolean           ' File loading. Disable some screen updating.
      IsCompiling                as boolean           ' File/Project currently being compiled (spinning mouse cursor).
      IsShutDown                 as boolean           ' App is currently closing
      wszCommandLine             as CWSTR             ' non-project commandline (not saved to file)
      wszLastOpenFolder          as CWSTR             ' remembers the last opened folder for the Open Dialog
      
      hWndAutoCListBox           as hwnd              ' handle of popup autocomplete ListBox window
            
      IsProjectActive            as boolean
      ProjectBuild               as string            ' default build configuration for the project (GUID)
      ProjectName                as CWSTR
      ProjectFilename            as CWSTR
      ProjectOther32             as CWSTR             ' compile flags 32 bit compiler
      ProjectOther64             as CWSTR             ' compile flags 64 bit compiler
      ProjectNotes               as CWSTR             ' Save/Load from project file
      ProjectCommandLine         as CWSTR
      ProjectDefaultFont         as CWSTR = "Segoe UI,9,400,0,0,0,1"
      ProjectManifest            as long              ' T/F create a generic resource and manifest file

      ' Global string to track the last accessed property/event in the PropertyList. This allows the
      ' user to quickly sqitch between controls that share common properties like 'Text'.
      PreviousPropName           as CWSTR
      PreviousEventName          as CWSTR

      declare function AddQuickRunEXE( byref sFilename as wstring ) as long
      declare function CheckQuickRunEXE() as long
      declare function RemoveAllSelectionAttributes() as long
      Declare function AddNewDocument() as clsDocument ptr 
      Declare function RemoveDocument( byval pDoc as clsDocument ptr ) as long
      declare function RemoveAllDocuments() as long
      Declare function GetDocumentCount() as long
      declare function GetDocumentPtrByWindow( byval hWindow as hwnd) as clsDocument ptr
      Declare function GetDocumentPtrByFilename( Byref wszName as wstring ) as clsDocument ptr
      Declare function GetMainDocumentPtr() as clsDocument ptr
      Declare function GetResourceDocumentPtr() as clsDocument ptr
      declare function GetSourceDocumentPtr( byval pDocIn as clsDocument ptr ) as clsDocument ptr
      declare function GetHeaderDocumentPtr( byval pDocIn as clsDocument ptr ) as clsDocument ptr
      Declare function SaveProject( byval bSaveas as boolean = False ) as boolean
      Declare function ProjectSetFileType( byval pDoc as clsDocument ptr, byval wszFiletype as CWSTR ) as LRESULT
      declare function GetProjectCompiler() as long
      
end type

