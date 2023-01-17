' ========================================================================================
' WinFBE
' Windows FreeBASIC Editor (Windows 32/64 bit)
' Paul Squires (2016-2023)
' ========================================================================================

'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2023 Paul Squires, PlanetSquires Software
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

#define UNICODE
#define _WIN32_WINNT &h0602  

#include once "windows.bi"
#include once "vbcompat.bi"
#include once "win\shobjidl.bi"
#include once "win\TlHelp32.bi"
#include once "crt\string.bi"
#include once "win\Shlobj.bi"
#include once "Afx\CWindow.inc"
#include once "Afx\AfxFile.inc"
#include once "Afx\AfxStr.inc"
#include once "Afx\AfxRichEdit.inc"
#include once "Afx\AfxTime.inc"
#include once "Afx\AfxGdiplus.inc"
#include once "Afx\AfxMenu.inc" 
#include once "Afx\AfxCom.inc" 
#include once "Afx\CXpButton.inc"
#include once "Afx\CMaskedEdit.inc"
#include once "Afx\CImageCtx.inc"
#include once "Afx\CAxHost\CWebCtx.inc"
#include once "Afx\CWinHttpRequest.inc"

using Afx


#define APPNAME        wstr("WinFBE - FreeBASIC Editor")
#define APPNAMESHORT   wstr("WinFBE")

#define APPVERSION     wstr("3.0.8") 
#define PREVENT_UPDATE_CHECK FALSE     ' used for betas

#define APPCOPYRIGHT   wstr("Paul Squires, PlanetSquires Software, Copyright (C) 2016-2023") 
dim shared as CWSTR gwszDefaultToolchain = "FreeBASIC-1.09.0-winlibs-gcc-9.3.0"

#ifdef __FB_64BIT__
   #define APPBITS wstr(" (64-bit)")
#else
   #define APPBITS wstr(" (32-bit)")
#endif

#include once "modScintilla.bi"
#include once "modDeclares.bi"         
#include once "cJSON.bi"

#include once "clsLasso.bi"
#include once "clsDocument.bi"
#include once "clsTopTabCtl.bi"
#include once "clsDB2.bi"
#include once "clsConfig.bi"
#include once "clsApp.bi"

'  Global classes
dim shared gApp     as clsApp
dim shared gConfig  as clsConfig
dim shared gTTabCtl as clsTopTabCtl
dim shared gLasso   as clsLasso


#include once "clsDB2.inc"
#include once "clsConfig.inc"
#include once "modThemes.inc"
#include once "modRoutines.inc"
#include once "modParser.inc"
#include once "clsControl.inc"
#include once "clsCollection.inc"
#include once "clsDocument.inc"
#include once "clsApp.inc"
#include once "clsTopTabCtl.inc"
#include once "clsLasso.inc"
#include once "modVDDesignFrame.inc"
#include once "modVDRoutines.inc"
#include once "modVDProperties.inc"
#include once "modVDApplyProperties.inc"
#include once "modVDColors.inc"
#include once "modVDAnchors.inc"
#include once "modVDControls.inc"
#include once "modVDDesignForm.inc"
#include once "modVDDesignMain.inc"
#include once "modVDToolbox.inc"
#include once "modAutoInsert.inc"
#include once "modCompile.inc"
#include once "modCompileErrors.inc"
#include once "modMenus.inc"
#include once "modCodetips.inc"
#include once "modGenerateCode.inc"
#include once "modMenuDefinitions.inc"
#include once "modMRU.inc"
#include once "mod302Upgrade.inc"

#include once "frmVDTabChild.inc"
#include once "frmAbout.inc" 
#include once "frmPopupMenu.inc"
#include once "frmTopTabs.inc"
#include once "frmMenuBar.inc"
#include once "frmStatusBar.inc"
#include once "frmImageManager.inc" 
#include once "frmPanelVScroll.inc" 
#include once "frmEditorHScroll.inc" 
#include once "frmEditorVScroll.inc" 
#include once "frmPanel.inc" 
#include once "frmExplorer.inc" 
#include once "frmBookmarks.inc" 
#include once "frmFunctions.inc" 
#include once "frmKeyboardEdit.inc" 
#include once "frmKeyboard.inc" 
#include once "frmUserTools.inc" 
#include once "frmSnippets.inc"
#include once "frmCategories.inc" 
#include once "frmBuildConfig.inc" 
#include once "frmDesignTabs.inc"
#include once "frmOutput.inc" 
#include once "frmOptionsGeneral.inc"
#include once "frmOptionsEditor.inc"
#include once "frmOptionsEditor2.inc"
#include once "frmOptionsColors.inc"
#include once "frmOptionsCompiler.inc"
#include once "frmOptionsLocal.inc"
#include once "frmOptionsKeywords.inc"
#include once "frmOptionsKeywordsWinApi.inc"
#include once "frmOptions.inc"
#include once "frmTemplates.inc"
#include once "frmGoto.inc"
#include once "frmCommandLine.inc"
#include once "frmFindInFiles.inc"
#include once "frmFindReplace.inc"
#include once "frmProjectOptions.inc"
#include once "frmHelpViewer.inc"
#include once "frmMenuEditor.inc"
#include once "frmToolBarEditor.inc"
#include once "frmStatusBarEditor.inc"
#include once "frmMainOnCommand.inc"
#include once "modMsgPump.inc"
#include once "frmMainFile.inc"
#include once "frmMainEdit.inc"
#include once "frmMainSearch.inc"
#include once "frmMainView.inc"
#include once "frmMainProject.inc"
#include once "frmMainCompile.inc"
#include once "frmMainDesigner.inc"
#include once "frmMain.inc"

' ========================================================================================
' Check if running under Linux Wine
' ========================================================================================
function isWineActive() as boolean
   dim hLib as HMODULE = LoadLibraryW("NtDll.dll")
   if hLib = null then exit function
   dim pwine as function() as long
   pwine = cast(any ptr, GetProcAddress(hLib, "wine_get_version"))
   function = iif( pwine, true, false )
   FreeLibrary hLib
end function
' ========================================================================================


' ========================================================================================
' WinMain
' ========================================================================================
function WinMain( _
            byval hInstance     as HINSTANCE, _
            byval hPrevInstance as HINSTANCE, _
            byval szCmdLine     as zstring ptr, _
            byval nCmdShow      as long _
            ) as long

   gApp.isWineActive = isWineActive()

   ' Load configuration files 
   gConfig.LoadConfigFile()
   gConfig.LoadKeywords()

  
   ' Attempt to load the english localization file. This is necessary because
   ' any non-english localization file will have missing entries filled by the
   ' english version.
   dim as CWSTR wszLocalizationFile
   wszLocalizationFile = AfxGetExePathName + wstr("Languages\english.lang")
   if LoadLocalizationFile(wszLocalizationFile, true) = false Then
      MessageBox( 0, _
                  "English Localization file could not be loaded. Aborting application." + vbcrlf + _
                  wszLocalizationFile, _
                  "Error", _
                  MB_OK or MB_ICONWARNING or MB_DEFBUTTON1 or MB_APPLMODAL )
      return 1
   end if
   
   
   ' Load the selected localization file
   wszLocalizationFile = AfxGetExePathName + "Languages\" + gConfig.LocalizationFile
   if LoadLocalizationFile(wszLocalizationFile, false) = false then
      MessageBox( 0, _
                  "Localization file could not be loaded. Aborting application." + vbcrlf + _
                  wszLocalizationFile, _
                  "Error", _
                  MB_OK or MB_ICONWARNING or MB_DEFBUTTON1 or MB_APPLMODAL )
      Return 1
   end if

   ' Load default Explorer Categories should none exist. Need to do it here
   ' rather than from within Config because the localization file must be 
   ' loaded first.
   gConfig.SetCategoryDefaults()

   ' Check for previous instance 
   if gConfig.MultipleInstances = false Then
      dim as HWND hWindow = FindWindow("WinFBE_Class", 0)
      if hWindow then
         SetForegroundWindow(hWindow)
         frmMain_ProcessCommandLine(hWindow)
         return true
      end if
   end if
   

   ' Initialize the COM library
   CoInitialize(null)


   #IfDef __FB_64BIT__
      ' Load the Scintilla code editing dll
      dim as any ptr pLibLexilla = dylibload("Lexilla64.dll")
      dim as any ptr pLibScintilla = dylibload("Scintilla64.dll")
   #Else
      ' Load the Scintilla code editing dll
      dim as any ptr pLibLexilla = dylibload("Lexilla32.dll")
      dim as any ptr pLibScintilla = dylibload("Scintilla32.dll")
   #EndIf
   gApp.pfnCreateLexerfn = cast(CreateLexerFn , GetProcAddress(pLibLexilla, "CreateLexer"))

   ' Load the HTML help library for displaying FreeBASIC help *.chm file
   gpHelpLib = dylibload( "hhctrl.ocx" )

   ' Load codetip files 
   if gConfig.Codetips then gConfig.LoadCodetips

   
   ' Load any user code snippets and initialize the ToolBox
   gConfig.LoadSnippets
   gConfig.InitializeToolBox
   

   ' Show the main form
   function = frmMain_Show( 0 )


   ' Free the Scintilla, CaptureConsole and HTML help libraries
   dylibfree(pLibLexilla)
   dylibfree(pLibScintilla)
   dylibfree(gpHelpLib)
   

   ' Uninitialize the COM library
   CoUninitialize

end function


' ========================================================================================
' Main program entry point
' ========================================================================================
end WinMain( GetModuleHandle(null), null, command(), SW_NORMAL )

