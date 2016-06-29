
' ========================================================================================
' WinFBE
' Windows FreeBASIC Editor (Windows 32/64 bit)
' Paul Squires and José Roca (2016)
' ========================================================================================


#Define UNICODE
#Define _WIN32_WINNT &h0602  

#Include Once "windows.bi"
#Include Once "vbcompat.bi"
#Include Once "win\shobjidl.bi"
#Include Once "Afx\CWindow.inc"
#Include Once "Afx\AfxGdiplus.inc"
#Include Once "Afx\AfxCtl.inc" 
#Include Once "Afx\AfxMenu.inc" 

Using Afx.CWindowClass
' $FB_RESPATH = "WinFBE.rc"


#Define APPNAME       WStr("WinFBE - FreeBASIC Editor")
#Define APPNAMESHORT  WStr("WinFBE")
#Define APPVERSION    WStr("1.0.0")


'  Global window handle for the main form
Dim Shared As HWnd HWND_FRMMAIN, HWND_FRMMAIN_TOOLBAR
Dim Shared As HMENU HWND_FRMMAIN_TOPMENU   

'  Global window handles for some forms 
Dim Shared As HWnd HWND_FRMOPTIONSEDITOR, HWND_FRMOPTIONSCOLORS, HWND_FRMOPTIONSCOMPILER
Dim Shared As HWnd HWND_FRMOPTIONSLOCAL, HWND_FRMOPTIONSKEYWORDS
Dim Shared As HWnd HWND_FRMCOMPILERESULTS, HWND_FRMFIND, HWND_FRMREPLACE
Dim Shared As HWnd HWND_FRMPROJECTMANAGER


' Create a temporary array to hold the selected color values
' for the different editor elements. When the form is saved then 
' the values from this temporary struture is saved to the 
' configuration table.
Type TYPE_COLORS
   nFg As COLORREF
   nBg As COLORREF
End Type
Dim Shared gTempColors(15) As TYPE_COLORS

' Create a dynamic array that will hold all localization words/phrases. This
' array is resized and loaded using the LoadLocalizationFile function.
ReDim Shared LL(Any) As WString * MAX_PATH

' Define a macro that allows the user to specify the LL array subscript and
' also a descriptive label (that is ignored), and return the LL array value.
#Define L(e,s)  LL(e)


#Include Once "Modules\windowsxx.bi"
#Include Once "Modules\modScintilla.bi"
#Include Once "Modules\modDeclares.inc"
#Include Once "Modules\clsConfig.inc"
#Include Once "Modules\modRoutines.inc"
#Include Once "Modules\modCBColor.inc"
#Include Once "Modules\clsDocument.inc"
#Include Once "Modules\clsApp.inc"
#Include Once "Modules\clsTopTabCtl.inc"
#Include Once "Modules\modTopMenu.inc"
#Include Once "Modules\modToolbar.inc"
#Include Once "Modules\modCompile.inc"
#Include Once "Modules\modMRU.inc"

#Include Once "Forms\frmOptionsEditor.inc"
#Include Once "Forms\frmOptionsColors.inc"
#Include Once "Forms\frmOptionsCompiler.inc"
#Include Once "Forms\frmOptionsLocal.inc"
#Include Once "Forms\frmOptionsKeywords.inc"
#Include Once "Forms\frmOptions.inc"
#Include Once "Forms\frmTemplates.inc"
#Include Once "Forms\frmCompileResults.inc"
#Include Once "Forms\frmGoto.inc"
#Include Once "Forms\frmCommandLine.inc"
#Include Once "Forms\frmFind.inc"
#Include Once "Forms\frmReplace.inc"
#Include Once "Forms\frmProjectManager.inc"
#Include Once "Forms\frmMain.inc"


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
   If LoadLocalizationFile( Exepath & "\" & gConfig.LocalizationFile ) = False Then
      MessageBoxW( 0, WStr("Localization file could not be loaded. Aborting application.") & vbcrlf & _
                   Exepath & "\" & gConfig.LocalizationFile, _
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
   Dim As Any Ptr pLib = Dylibload("SciLexer32.dll")

   ' Show the main form
   Function = frmMain_Show( 0, nCmdShow )

   ' Destroy the global App class which in turn frees all allocated projects and documents
   Delete gpApp
   
   ' Free the Scintilla library
   Dylibfree(pLib)
   
   ' Uninitialize the COM library
   CoUninitialize

End Function


' ========================================================================================
' Main program entry point
' ========================================================================================
End WinMain( GetModuleHandleW(Null), Null, Command(), SW_NORMAL )


