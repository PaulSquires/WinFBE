' ========================================================================================
' WinFBE
' Windows FreeBASIC Editor (Windows 32/64 bit)
' Paul Squires (2016-2019)
' ========================================================================================

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

#Define UNICODE
#Define _WIN32_WINNT &h0602  

#Include Once "windows.bi"
#Include Once "vbcompat.bi"
#Include Once "win\shobjidl.bi"
#Include Once "win\TlHelp32.bi"
#Include Once "crt\string.bi"
#Include Once "win\Shlobj.bi"
#Include Once "Afx\CWindow.inc"
#Include Once "Afx\AfxFile.inc"
#Include Once "Afx\AfxStr.inc"
#Include Once "Afx\AfxTime.inc"
#Include Once "Afx\AfxGdiplus.inc"
#Include Once "Afx\AfxMenu.inc" 
#Include Once "Afx\AfxCom.inc" 
#Include Once "Afx\CXpButton.inc"
#Include Once "Afx\CMaskedEdit.inc"
#Include Once "Afx\CImageCtx.inc"
#Include Once "Afx\CAxHost\CWebCtx.inc"
#Include Once "Afx\CWinHttpRequest.inc"

Using Afx


#Define APPNAME        WStr("WinFBE - FreeBASIC Editor")
#Define APPNAMESHORT   WStr("WinFBE")
#Define APPVERSION     WStr("2.0.2") 
#Define APPCOPYRIGHT   WStr("Paul Squires, PlanetSquires Software, Copyright (C) 2016-2019") 


#ifdef __FB_64BIT__
   #Define APPBITS WStr(" (64-bit)")
#else
   #Define APPBITS WStr(" (32-bit)")
#endif

#Define FILES_MENU_POSITION        0
#Define PROJECT_MENU_POSITION      4
#Define TOOLS_MENU_POSITION        8
#Define MRUFILES_MENU_POSITION     11
#Define MRUPROJECTS_MENU_POSITION  3

#Include Once "modScintilla.bi"
#Include Once "modDeclares.bi"         

#include once "clsLasso.bi"
#include once "clsDocument.bi"
#include once "clsTopTabCtl.bi"
#include once "clsConfig.bi"
#include once "clsApp.bi"
#include once "clsDB2.bi"

'  Global classes
Dim Shared gApp     As clsApp
Dim Shared gConfig  As clsConfig
Dim Shared gTTabCtl As clsTopTabCtl
dim shared gLasso   as clsLasso


#Include Once "clsDB2.inc"
#Include Once "clsParser.inc"
#Include Once "clsConfig.inc"
#Include Once "modRoutines.inc"
#Include Once "modCBColor.inc"
#Include Once "clsControl.inc"
#Include Once "clsCollection.inc"
#Include Once "clsDocument.inc"
#Include Once "clsApp.inc"
#Include Once "clsTopTabCtl.inc"
#Include Once "clsLasso.inc"
#Include Once "modVDDesignFrame.inc"
#Include Once "modVDRoutines.inc"
#Include Once "modVDProperties.inc"
#Include Once "modVDApplyProperties.inc"
#Include Once "modVDColors.inc"
#Include Once "modVDControls.inc"
#Include Once "modVDDesignForm.inc"
#Include Once "modVDDesignMain.inc"
#Include Once "modVDToolbox.inc"
#Include Once "modAutoInsert.inc"
#Include Once "modCompile.inc"
#Include Once "modMenus.inc"
#Include Once "modToolbar.inc"
#Include Once "modMRU.inc"
#Include Once "modCodetips.inc"
#Include Once "modGenerateCode.inc"

#Include Once "frmVDTabChild.inc"
#Include Once "frmAbout.inc" 
#Include Once "frmImageManager.inc" 
#Include Once "frmRecent.inc" 
#Include Once "frmExplorer.inc" 
#Include Once "frmUserTools.inc" 
#Include Once "frmSnippets.inc"
#Include Once "frmBuildConfig.inc" 
#Include Once "frmOutput.inc" 
#Include Once "frmOptionsGeneral.inc"
#Include Once "frmOptionsEditor.inc"
#Include Once "frmOptionsColors.inc"
#Include Once "frmOptionsCompiler.inc"
#Include Once "frmOptionsLocal.inc"
#Include Once "frmOptionsKeywords.inc"
#Include Once "frmOptions.inc"
#Include Once "frmTemplates.inc"
#Include Once "frmFunctionList.inc"
#Include Once "frmGoto.inc"
#Include Once "frmCommandLine.inc"
#Include Once "frmFindInFiles.inc"
#Include Once "frmFindReplace.inc"
#Include Once "frmProjectOptions.inc"
#Include Once "frmHelpViewer.inc"
#Include Once "frmMenuEditor.inc"
#Include Once "frmToolBarEditor.inc"
#Include Once "frmStatusBarEditor.inc"
#Include Once "frmMainOnCommand.inc"
#Include Once "frmMain.inc"


' ========================================================================================
' WinMain
' ========================================================================================
Function WinMain( ByVal hInstance     As HINSTANCE, _
                  ByVal hPrevInstance As HINSTANCE, _
                  ByVal szCmdLine     As ZString Ptr, _
                  ByVal nCmdShow      As Long _
                  ) As Long

   ' Load configuration files 
   gConfig.LoadConfigFile()
   gConfig.LoadKeywords()
   

   ' Attempt to load the english localization file. This is necessary because
   ' any non-english localization file will have missing entries filled by the
   ' english version.
   dim as CWSTR wszLocalizationFile
   wszLocalizationFile = AfxGetExePathName + wstr("Languages\english.lang")
   If LoadLocalizationFile(wszLocalizationFile, true) = False Then
      MessageBox( 0, _
                  WStr("English Localization file could not be loaded. Aborting application.") + vbcrlf + _
                  wszLocalizationFile, _
                  WStr("Error"), _
                  MB_OK Or MB_ICONWARNING Or MB_DEFBUTTON1 Or MB_APPLMODAL )
      Return 1
   End If
   
   
   ' Load the selected localization file
   wszLocalizationFile = AfxGetExePathName + wstr("Languages\") + gConfig.LocalizationFile
   If LoadLocalizationFile(wszLocalizationFile, false) = False Then
      MessageBox( 0, _
                  WStr("Localization file could not be loaded. Aborting application.") + vbcrlf + _
                  wszLocalizationFile, _
                  WStr("Error"), _
                  MB_OK Or MB_ICONWARNING Or MB_DEFBUTTON1 Or MB_APPLMODAL )
      Return 1
   End If


   ' Check for previous instance 
   If gConfig.MultipleInstances = False Then
      dim as HWND hWindow = FindWindow("WinFBE_Class", 0)
      If hWindow Then
         SetForegroundWindow(hWindow)
         frmMain_ProcessCommandLine(hWindow)
         Return True
      End If
   End If
   

   ' Initialize the COM library
   CoInitialize(Null)


   ' Load the Scintilla code editing dll
   #IfDef __FB_64BIT__
      Dim As Any Ptr pLib = Dylibload("SciLexer64.dll")
   #Else
      Dim As Any Ptr pLib = Dylibload("SciLexer32.dll")
   #EndIf

   
   ' Load the HTML help library for displaying FreeBASIC help *.chm file
   gpHelpLib = DyLibLoad( "hhctrl.ocx" )

   
   ' Load preparsed codetip files for the compiler's \inc folders.
   if gConfig.Codetips then
      gConfig.LoadCodetipsFB()
      gConfig.LoadCodetipsWinAPI()
      gConfig.LoadCodetipsWinFormsX()
      gConfig.LoadCodetipsWinFBX()
   end if

   
   ' Load any user code snippets and initialize the ToolBox
   gConfig.LoadSnippets
   gConfig.InitializeToolBox
   

   ' Show the main form
   Function = frmMain_Show( 0 )


   ' Free the Scintilla and HTML help libraries
   Dylibfree(pLib)
   Dylibfree(gpHelpLib)
   

   ' Uninitialize the COM library
   CoUninitialize

End Function


' ========================================================================================
' Main program entry point
' ========================================================================================
End WinMain( GetModuleHandleW(Null), Null, Command(), SW_NORMAL )





