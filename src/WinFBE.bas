
' ========================================================================================
' WinFBE
' Windows FreeBASIC Editor (Windows 32/64 bit)
' Paul Squires (2016)
' ========================================================================================


#Define UNICODE
#Define _WIN32_WINNT &h0602  

#Include Once "windows.bi"
#Include Once "vbcompat.bi"
#Include Once "win\shobjidl.bi"
#Include Once "win\TlHelp32.bi"
#Include Once "win\Shlobj.bi"
#Include Once "Afx\CWindow.inc"
#Include Once "Afx\AfxStr.inc"
#Include Once "Afx\AfxTime.inc"
#Include Once "Afx\AfxGdiplus.inc"
#Include Once "Afx\AfxCtl.inc" 
#Include Once "Afx\AfxMenu.inc" 

Using Afx
' $FB_RESPATH = "WinFBE.rc"


#Define APPNAME       WStr("WinFBE - FreeBASIC Editor")
#Define APPNAMESHORT  WStr("WinFBE")
#Define APPVERSION    WStr("0.0.0")


#Include Once "windowsxx.bi"      ' needed because version that ships with FB is broken and incomplete. 
#Include Once "modScintilla.bi"
#Include Once "modDeclares.bi"
#Include Once "clsConfig.inc"
#Include Once "modRoutines.inc"
#Include Once "modCBColor.inc"
#Include Once "clsDocument.inc"
#Include Once "clsApp.inc"
#Include Once "clsTopTabCtl.inc"
#Include Once "modHelp.inc"
#Include Once "modTopMenu.inc"
#Include Once "modToolbar.inc"
#Include Once "modCompile.inc"
#Include Once "modMRU.inc"

#Include Once "frmOptionsGeneral.inc"
#Include Once "frmOptionsEditor.inc"
#Include Once "frmOptionsColors.inc"
#Include Once "frmOptionsCompiler.inc"
#Include Once "frmOptionsLocal.inc"
#Include Once "frmOptionsKeywords.inc"
#Include Once "frmOptions.inc"
#Include Once "frmTemplates.inc"
#Include Once "frmFnList.inc"
#Include Once "frmCompileResults.inc"
#Include Once "frmGoto.inc"
#Include Once "frmCommandLine.inc"
#Include Once "frmFind.inc"
#Include Once "frmReplace.inc"
#Include Once "frmProjectManager.inc"
#Include Once "frmProjectOptions.inc"
#Include Once "frmMain.inc"


' ========================================================================================
' WinMain
' ========================================================================================
Function WinMain( ByVal hInstance     As HINSTANCE, _
                  ByVal hPrevInstance As HINSTANCE, _
                  ByVal szCmdLine     As ZString Ptr, _
                  ByVal nCmdShow      As Long _
                  ) As Long

   ' Load configuration file 
   gConfig.LoadFromFile()

   ' Load the selected localization file
   If LoadLocalizationFile( Exepath & "\Languages\" & gConfig.LocalizationFile ) = False Then
      MessageBoxW( 0, WStr("Localization file could not be loaded. Aborting application.") & vbcrlf & _
                   Exepath & "\Languages\" & gConfig.LocalizationFile, _
                   WStr("Error"), MB_OK Or MB_ICONWARNING Or MB_DEFBUTTON1 Or MB_APPLMODAL )
      Return 1
   End If

   ' Check for previous instance 
   If gConfig.MultipleInstances = False Then
      If FindWindow("WinFBE_Class", 0) Then
         Return True
      End If
   End If
   
   ' Initialize the COM library
   CoInitialize(Null)

   ' Create a global App class to control all projects and documents
   gpApp = New clsApp

   ' Load the Scintilla code editing dll
   #IfDef __FB_64BIT__
      Dim As Any Ptr pLib = Dylibload("SciLexer64.dll")
   #Else
      Dim As Any Ptr pLib = Dylibload("SciLexer32.dll")
   #EndIf
   
   ' Load the HTML help library for displaying FreeBASIC help *.chm file
   gpHelpLib = DyLibLoad( "hhctrl.ocx" )
   
   ' Show the main form
   Function = frmMain_Show( 0, nCmdShow )

   ' Destroy the global App class which in turn frees all allocated projects and documents
   Delete gpApp
   
   ' Free the Scintilla library
   Dylibfree(pLib)
   
   ' Free the HTML help library
   Dylibfree(gpHelpLib)
   
   ' Uninitialize the COM library
   CoUninitialize

End Function


' ========================================================================================
' Main program entry point
' ========================================================================================
End WinMain( GetModuleHandleW(Null), Null, Command(), SW_NORMAL )













