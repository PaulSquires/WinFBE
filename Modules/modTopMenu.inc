
' ========================================================================================
' Build the frmMain topmenu
' ========================================================================================
Function frmMain_BuildMenu ( ByVal pWindow As CWindow Ptr ) As HMENU

   Dim As HMENU hMenu, hPopUpMenu, hPopUpSubMenu

   hMenu = CreateMenu
   hPopUpMenu = CreatePopupMenu
      AppendMenuW hMenu, MF_POPUP Or MF_ENABLED, Cast(UINT_PTR, hPopUpMenu), "&File"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_FILENEW, "&New" & Chr(9) & "Ctrl+N"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_FILEOPEN, "&Open..." & Chr(9) & "Ctrl+O"
         AppendMenuW hPopUpMenu, MF_SEPARATOR, 0, ""
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_FILECLOSE, "&Close" & Chr(9) & "Ctrl+Q"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_FILECLOSEALL, "Close All"
         AppendMenuW hPopUpMenu, MF_SEPARATOR, 0, ""
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_FILESAVE, "&Save" & Chr(9) & "Ctrl+S"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_FILESAVEAS, "Save &As" & Chr(9) & "F12"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_FILESAVEALL, "Sa&ve All" & Chr(9) & "Ctrl+Shift+S"
         AppendMenuW hPopUpMenu, MF_SEPARATOR, 0, ""
         hPopUpSubMenu = CreatePopupMenu
            AppendMenuW hPopUpMenu, MF_POPUP Or MF_ENABLED, Cast(UINT_PTR, hPopUpSubMenu), "&Recent Files"
              AppendMenuW hPopUpSubMenu, MF_ENABLED, IDM_MRU, "Most Recently Used"
         AppendMenuW hPopUpMenu, MF_SEPARATOR, 0, ""
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_COMMAND, "Co&mmand Prompt..."
         AppendMenuW hPopUpMenu, MF_SEPARATOR, 0, ""
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_EXIT, "E&xit" & Chr(9) & "Alt+F4"

   hPopUpMenu = CreatePopupMenu
      AppendMenuW hMenu, MF_POPUP Or MF_ENABLED, Cast(UINT_PTR, hPopUpMenu), "&Edit"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_UNDO, "&Undo" & Chr(9) & "Ctrl+Z"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_REDO, "&Redo" & Chr(9) & "Ctrl+E"
         AppendMenuW hPopUpMenu, MF_SEPARATOR, 0, ""
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_CUT, "Cu&t" & Chr(9) & "Ctrl+X"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_COPY, "&Copy" & Chr(9) & "Ctrl+C"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_PASTE, "&Paste" & Chr(9) & "Ctrl+V"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_DELETELINE, "&Delete Line" & Chr(9) & "Ctrl+Y"
         AppendMenuW hPopUpMenu, MF_SEPARATOR, 0, ""
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_INSERTFILE, "Insert &File..." & Chr(9) & "Ctrl+I"
         AppendMenuW hPopUpMenu, MF_SEPARATOR, 0, ""
         hPopUpSubMenu = CreatePopupMenu
            AppendMenuW hPopUpMenu, MF_POPUP Or MF_ENABLED, Cast(UINT_PTR, hPopUpSubMenu), "&Block Operations"
              AppendMenuW hPopUpSubMenu, MF_ENABLED, IDM_INDENTBLOCK, "&Indent Block" & Chr(9) & "TAB"
              AppendMenuW hPopUpSubMenu, MF_ENABLED, IDM_UNINDENTBLOCK, "U&nIndent Block" & Chr(9) & "Shift+TAB"
              AppendMenuW hPopUpSubMenu, MF_SEPARATOR, 0, ""
              AppendMenuW hPopUpSubMenu, MF_ENABLED, IDM_COMMENTBLOCK, "Co&mment Block" & Chr(9) & "Ctrl+B"
              AppendMenuW hPopUpSubMenu, MF_ENABLED, IDM_UNCOMMENTBLOCK, "UnComment &Block" & Chr(9) & "Ctrl+Shift+B"
         hPopUpSubMenu = CreatePopupMenu
            AppendMenuW hPopUpMenu, MF_POPUP Or MF_ENABLED, Cast(UINT_PTR, hPopUpSubMenu), "&Line Operations"
              AppendMenuW hPopUpSubMenu, MF_ENABLED, IDM_DUPLICATELINE, "&Duplicate Current Line" & Chr(9) & "Ctrl+D"
              AppendMenuW hPopUpSubMenu, MF_ENABLED, IDM_MOVELINEUP, "Move &Up Current Line" & Chr(9) & "Ctrl+Shift+Up"
              AppendMenuW hPopUpSubMenu, MF_ENABLED, IDM_MOVELINEDOWN, "Move &Down Current Line" & Chr(9) & "Ctrl+Shift+Down"
         hPopUpSubMenu = CreatePopupMenu
            AppendMenuW hPopUpMenu, MF_POPUP Or MF_ENABLED, Cast(UINT_PTR, hPopUpSubMenu), "&Case Conversions"
              AppendMenuW hPopUpSubMenu, MF_ENABLED, IDM_TOUPPERCASE, "Selection to Upp&er Case" & Chr(9) & "Ctrl+Alt+U"
              AppendMenuW hPopUpSubMenu, MF_ENABLED, IDM_TOLOWERCASE, "Selection to &Lower Case" & Chr(9) & "Ctrl+Alt+L"
              AppendMenuW hPopUpSubMenu, MF_ENABLED, IDM_TOMIXEDCASE, "Selection to Mi&xed Case" & Chr(9) & "Ctrl+Alt+M"
         hPopUpSubMenu = CreatePopupMenu
            AppendMenuW hPopUpMenu, MF_POPUP Or MF_ENABLED, Cast(UINT_PTR, hPopUpSubMenu), "&End of Line Character Conversions"
              AppendMenuW hPopUpSubMenu, MF_ENABLED, IDM_EOLTOCRLF, "Convert to CRLF"
              AppendMenuW hPopUpSubMenu, MF_ENABLED, IDM_EOLTOCR, "Convert to CR"
              AppendMenuW hPopUpSubMenu, MF_ENABLED, IDM_EOLTOLF, "Convert to LF"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_TABSTOSPACES, "Replace Tabs &with Spaces"
'         hPopUpSubMenu = CreatePopupMenu
'            AppendMenuW hPopUpMenu, MF_POPUP Or MF_ENABLED, Cast(UINT_PTR, hPopUpSubMenu), "Replace &Spaces or Tabs"
'              AppendMenuW hPopUpSubMenu, MF_ENABLED, IDM_SPACESTOTABS, "Replace Spaces with Tabs"
'              AppendMenuW hPopUpSubMenu, MF_ENABLED, IDM_TABSTOSPACES, "Replace Tabs with Spaces"
         AppendMenuW hPopUpMenu, MF_SEPARATOR, 0, ""
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_SELECTALL, "Select &All" & Chr(9) & "Ctrl+A"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_SELECTLINE, "&Select Line" & Chr(9) & "Ctrl+L"

   hPopUpMenu = CreatePopupMenu
      AppendMenuW hMenu, MF_POPUP Or MF_ENABLED, Cast(UINT_PTR, hPopUpMenu), "&Search"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_FIND, "&Find..." & Chr(9) & "Ctrl+F"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_FINDNEXT, "Find &Next" & Chr(9) & "F3"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_FINDPREV, "Find &Previous" & Chr(9) & "Shift+F3"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_REPLACE, "&Replace..." & Chr(9) & "Ctrl+R"
         AppendMenuW hPopUpMenu, MF_SEPARATOR, 0, ""
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_DEFINITION, "Sub/Function &Definition" & Chr(9) & "F6"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_LASTPOSITION, "&Last Position" & Chr(9) & "Shift+F6"
         AppendMenuW hPopUpMenu, MF_SEPARATOR, 0, ""
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_GOTO, "&Goto Line..." & Chr(9) & "Ctrl+G"
         AppendMenuW hPopUpMenu, MF_SEPARATOR, 0, ""
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_BOOKMARKTOGGLE, "&Toggle Bookmark" & Chr(9) & "Ctrl+F2"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_BOOKMARKNEXT, "Next &Bookmark" & Chr(9) & "F2"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_BOOKMARKPREV, "Pre&vious Bookmark" & Chr(9) & "Shift+F2"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_BOOKMARKCLEARALL, "&Clear Bookmarks" & Chr(9) & "Ctrl+Shift+F2"

   hPopUpMenu = CreatePopupMenu
      AppendMenuW hMenu, MF_POPUP Or MF_ENABLED, Cast(UINT_PTR, hPopUpMenu), "&View"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_FOLDTOGGLE, "Toggle &Current Fold Point" & Chr(9) & "F8"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_FOLDBELOW, "Toggle Current And All &Below" & Chr(9) & "Ctrl+F8"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_FOLDALL, "Fold &All" & Chr(9) & "Alt+F8"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_UNFOLDALL, "&Unfold All" & Chr(9) & "Shift+F8"
         AppendMenuW hPopUpMenu, MF_SEPARATOR, 0, ""
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_ZOOMIN, "Zoom &In" & Chr(9) & "Ctrl++"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_ZOOMOUT, "Zoom &Out" & Chr(9) & "Ctrl+-"
         AppendMenuW hPopUpMenu, MF_SEPARATOR, 0, ""
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_RESTOREMAIN, "&Restore Main Window Size"
         AppendMenuW hPopUpMenu, MF_SEPARATOR, 0, ""
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_OPTIONS, "En&vironment Options..."

   hPopUpMenu = CreatePopupMenu
      AppendMenuW hMenu, MF_POPUP Or MF_ENABLED, Cast(UINT_PTR, hPopUpMenu), "&Project"

   hPopUpMenu = CreatePopupMenu
      AppendMenuW hMenu, MF_POPUP Or MF_ENABLED, Cast(UINT_PTR, hPopUpMenu), "&Compile"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_BUILDEXECUTE, "Build And E&xecute" & Chr(9) & "F5"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_COMPILE, "&Compile" & Chr(9) & "Ctrl+F5"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_RUNEXE, "&Run Executable" & Chr(9) & "Shift+F5"
         AppendMenuW hPopUpMenu, MF_SEPARATOR, 0, ""
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_COMMANDLINE, "Command &Line..."
         AppendMenuW hPopUpMenu, MF_SEPARATOR, 0, ""
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_USE32BIT, "Use &32-Bit Compiler" & Chr(9) & "Ctrl+1"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_USE64BIT, "Use &64-Bit Compiler" & Chr(9) & "Ctrl+2"
         AppendMenuW hPopUpMenu, MF_SEPARATOR, 0, ""
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_GUI, "Compile &GUI application" & Chr(9) & "Ctrl+3"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_CONSOLE, "Compile C&onsole application" & Chr(9) & "Ctrl+4"

   hPopUpMenu = CreatePopupMenu
      AppendMenuW hMenu, MF_POPUP Or MF_ENABLED, Cast(UINT_PTR, hPopUpMenu), "&Help"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_HELP, "&Help" & Chr(9) & "F1"
         AppendMenuW hPopUpMenu, MF_SEPARATOR, 0, ""
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_ABOUT, "&About"

   Function = hMenu

   '  Create the accelerator table for the menu
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL, VK_N,  IDM_FILENEW )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL, VK_O,  IDM_FILEOPEN )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL, VK_Q,  IDM_FILECLOSE )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL, VK_S,  IDM_FILESAVE )
   pWindow->AddAccelerator( FVIRTKEY, VK_F12,  IDM_FILESAVEAS )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL Or FSHIFT, VK_S,  IDM_FILESAVEALL )
   pWindow->AddAccelerator( FVIRTKEY Or FALT,     VK_F4, IDM_EXIT )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL, VK_Z,  IDM_UNDO )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL, VK_E,  IDM_REDO )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL, VK_X,  IDM_CUT )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL, VK_C,  IDM_COPY )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL, VK_V,  IDM_PASTE )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL, VK_Y,  IDM_DELETELINE )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL, VK_I,  IDM_INSERTFILE )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL, VK_B,  IDM_COMMENTBLOCK )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL Or FSHIFT, VK_B,  IDM_UNCOMMENTBLOCK )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL, VK_D,  IDM_DUPLICATELINE )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL Or FSHIFT, VK_UP,  IDM_MOVELINEUP )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL Or FSHIFT, VK_DOWN,  IDM_MOVELINEDOWN )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL, VK_A,  IDM_SELECTALL )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL, VK_L,  IDM_SELECTLINE )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL, VK_F,  IDM_FIND )
   pWindow->AddAccelerator( FVIRTKEY, VK_F3,  IDM_FINDNEXT )
   pWindow->AddAccelerator( FVIRTKEY Or FSHIFT, VK_F3,  IDM_FINDPREV )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL, VK_R,  IDM_REPLACE )
   pWindow->AddAccelerator( FVIRTKEY, VK_F6,  IDM_DEFINITION )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL Or FSHIFT, VK_F6,  IDM_LASTPOSITION )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL, VK_G,  IDM_GOTO )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL, VK_F2,  IDM_BOOKMARKTOGGLE )
   pWindow->AddAccelerator( FVIRTKEY, VK_F2,  IDM_BOOKMARKNEXT )
   pWindow->AddAccelerator( FVIRTKEY Or FSHIFT, VK_F2,  IDM_BOOKMARKPREV )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL Or FSHIFT, VK_F2,  IDM_BOOKMARKCLEARALL )
   pWindow->AddAccelerator( FVIRTKEY, VK_F8,  IDM_FOLDTOGGLE )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL, VK_F8,  IDM_FOLDBELOW )
   pWindow->AddAccelerator( FVIRTKEY Or FALT, VK_F8,  IDM_FOLDALL )
   pWindow->AddAccelerator( FVIRTKEY Or FSHIFT, VK_F8,  IDM_UNFOLDALL )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL, VK_OEM_PLUS,  IDM_ZOOMIN )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL, VK_OEM_MINUS,  IDM_ZOOMOUT )
   pWindow->AddAccelerator( FVIRTKEY, VK_F5,  IDM_BUILDEXECUTE )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL, VK_F5,  IDM_COMPILE )
   pWindow->AddAccelerator( FVIRTKEY Or FSHIFT, VK_F5,  IDM_RUNEXE )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL, VK_1,  IDM_USE32BIT )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL, VK_2,  IDM_USE64BIT )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL, VK_3,  IDM_GUI )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL, VK_4,  IDM_CONSOLE )
   pWindow->AddAccelerator( FVIRTKEY, VK_F1, IDM_HELP )
   pWindow->CreateAcceleratorTable()

'   DIM hInst AS HINSTANCE = GetModuleHandle(NULL)
'   DIM hSubMenu AS HMENU = GetSubMenu(hMenu, 0)
'   AfxAddIconToMenuItem(hSubMenu, 0, TRUE, AfxGdipIconFromRes(hInst, "IMAGE_NEW_FILE"))
'   AfxAddIconToMenuItem(hSubMenu, 1, TRUE, AfxGdipIconFromRes(hInst, "IMAGE_OPEN_FILE"))
'   AfxAddIconToMenuItem(hSubMenu, 6, TRUE, AfxGdipIconFromRes(hInst, "IMAGE_SAVE_FILE"))

End Function


' ========================================================================================
' Enable/Disable top menu items depending on application state
' ========================================================================================
Function frmMain_ChangeTopMenuStates() As Long

   Dim hMenu As HMENU = GetMenu(HWND_FRMMAIN)
   Dim idMenu      As Long
   Dim startSelPos As Long
   Dim endSelPos   As Long
   Dim hEdit       As HWnd
   
   ' Start off by disabling everything in order to make things easier. (We only
   ' then need to enable things)
   For idMenu = IDM_FILE To IDM_ABOUT
      EnableMenuItem hMenu, idMenu, MF_BYCOMMAND Or MF_GRAYED
   Next
   
   ' Items that will always be enabled.
   EnableMenuItem hMenu, IDM_FILENEW, MF_BYCOMMAND Or MF_ENABLED
   EnableMenuItem hMenu, IDM_FILEOPEN, MF_BYCOMMAND Or MF_ENABLED
   EnableMenuItem hMenu, IDM_COMMAND, MF_BYCOMMAND Or MF_ENABLED
   EnableMenuItem hMenu, IDM_RESTOREMAIN, MF_BYCOMMAND Or MF_ENABLED
   EnableMenuItem hMenu, IDM_OPTIONS, MF_BYCOMMAND Or MF_ENABLED
   EnableMenuItem hMenu, IDM_USE32BIT, MF_BYCOMMAND Or MF_ENABLED
   EnableMenuItem hMenu, IDM_USE64BIT, MF_BYCOMMAND Or MF_ENABLED
   EnableMenuItem hMenu, IDM_GUI, MF_BYCOMMAND Or MF_ENABLED
   EnableMenuItem hMenu, IDM_CONSOLE, MF_BYCOMMAND Or MF_ENABLED
   EnableMenuItem hMenu, IDM_HELP, MF_BYCOMMAND Or MF_ENABLED
   EnableMenuItem hMenu, IDM_ABOUT, MF_BYCOMMAND Or MF_ENABLED

   ' Check the correct default compiler
   Select Case UCase(gConfig.DefaultCompiler)
      Case "FBC 32BIT"
         CheckMenuItem hMenu, IDM_USE32BIT, MF_BYCOMMAND Or MF_CHECKED
         CheckMenuItem hMenu, IDM_USE64BIT, MF_BYCOMMAND Or MF_UNCHECKED
      Case "FBC 64BIT"
         CheckMenuItem hMenu, IDM_USE32BIT, MF_BYCOMMAND Or MF_UNCHECKED
         CheckMenuItem hMenu, IDM_USE64BIT, MF_BYCOMMAND Or MF_CHECKED
   End Select
   
   ' Check the correct default compile mode
   Select Case Ucase(gConfig.DefaultCompileMode)
      Case "GUI"
         CheckMenuItem hMenu, IDM_GUI, MF_BYCOMMAND Or MF_CHECKED
         CheckMenuItem hMenu, IDM_CONSOLE, MF_BYCOMMAND Or MF_UNCHECKED
      Case "CONSOLE"
         CheckMenuItem hMenu, IDM_GUI, MF_BYCOMMAND Or MF_UNCHECKED
         CheckMenuItem hMenu, IDM_CONSOLE, MF_BYCOMMAND Or MF_CHECKED
   End Select
   
   ' Enable everything if a document is open
   If TabCtrl_GetItemCount(gTTabCtl.hWindow) > 0 Then
      For idMenu = IDM_FILE To IDM_ABOUT
         EnableMenuItem hMenu, idMenu, MF_BYCOMMAND Or MF_ENABLED
      Next

      ' Enable or disable specific edit menu items (based on the currently active document)
      Dim pDoc As clsDocument Ptr = gTTabCtl.GetActiveDocumentPtr()
      If pDoc Then 
         hEdit = pDoc->hWindow
         If SendMessageW(hEdit, SCI_CANPASTE, 0, 0) = 0 Then 
            EnableMenuItem( hMenu, IDM_PASTE, MF_GRAYED)
         End If
         
         startSelPos = SendMessageW(hEdit, SCI_GETSELECTIONSTART, 0, 0)
         endSelPos   = SendMessageW(hEdit, SCI_GETSELECTIONEND, 0, 0)
         If startSelPos = endSelPos Then
            EnableMenuItem hMenu, IDM_CUT, MF_GRAYED
            EnableMenuItem hMenu, IDM_COPY, MF_GRAYED
         End If
      
         If SendMessageW(hEdit, SCI_CANUNDO, 0, 0) = 0 Then 
            EnableMenuItem hMenu, IDM_UNDO, MF_GRAYED
         End If
      
         If SendMessageW(hEdit, SCI_CANREDO, 0, 0) = 0 Then 
            EnableMenuItem hMenu, IDM_REDO, MF_GRAYED
         End If

         If SendMessageW(hEdit, SCI_CANPASTE, 0, 0) = 0 Then 
            EnableMenuItem hMenu, IDM_PASTE, MF_GRAYED
         End If
      End If
   End If
   
   Function = 0
End Function


''
''
Function frmMain_MenuSetCompiler( ByVal wID As Long ) As Long

   Select Case wID
      Case IDM_USE32BIT
         gConfig.DefaultCompiler = "FBC 32bit"
      Case IDM_USE64BIT
         gConfig.DefaultCompiler = "FBC 64bit"
   End Select
   
   frmMain_ChangeTopMenuStates()
   frmMain_UpdateLineCol(HWND_FRMMAIN)   ' so that the statusbar updates
   
   Function = 0
End Function


''
''
Function frmMain_MenuSetCompileMode( ByVal wID As Long ) As Long

   Select Case wID
      Case IDM_GUI
         gConfig.DefaultCompileMode = "GUI"
      Case IDM_CONSOLE
         gConfig.DefaultCompileMode = "CONSOLE"
   End Select
   
   frmMain_ChangeTopMenuStates()
   frmMain_UpdateLineCol(HWND_FRMMAIN)   ' so that the statusbar updates
   
   Function = 0
End Function


''
''
Function CreateContextMenu() As HMENU

   Dim pDoc As clsDocument Ptr = gTTabCtl.GetActiveDocumentPtr()
   If pDoc = 0 Then Exit Function

   Dim hPopupMenu    As HMENU
   Dim wPath         As WString * MAX_PATH
   Dim wCompilerPath As WString * MAX_PATH
   Dim sFilename     As String 
   Dim hWindow       As HWnd   = pDoc->hWindow
   Dim nLine         As Long   = pDoc->GetCurrentLineNumber()
   Dim sLine         As String = LTrim(pDoc->GetLine(nLine))
   Dim i             As Long
   Dim pt            As Point
   
   ' Ensure that the window under the cursor is actually a Scintilla window
   ' and not another window like a tab in the top tabcontrol.
   GetCursorPos(@pt)
   ScreenToClient(pDoc->hWindow, @pt)

   If RealChildWindowFromPoint(HWND_FRMMAIN, pt) <> pDoc->hWindow Then Exit Function
   
   ' Determine the compiler path in case it needs to be prepended to the filename
   ProcessFromCurdrive(@gConfig.FBWINCompiler32)
   ProcessFromCurdrive(@gConfig.FBWINCompiler64)
   Select Case Ucase(gConfig.DefaultCompiler)
      Case "FBC 32BIT":   wCompilerPath = gConfig.FBWINcompiler32
      Case "FBC 64BIT":   wCompilerPath = gConfig.FBWINcompiler64
   End Select
   FF_Pathname(WStr("PATH"), @wCompilerPath, @wCompilerPath)
   wCompilerPath = wCompilerPath & "inc"


   hPopupMenu = CreatePopupMenu

   If Left(Ucase(sLine), 9) = "#INCLUDE " Then sFilename = Mid(sLine, 10)
   If Left(Ucase(sLine), 14) = "#INCLUDE ONCE " Then sFilename = Mid(sLine, 15)

   If Len(sFilename) Then
      ' remove any comments at the end of the line
      i = Instr(sFilename, "'")
      If i Then sFilename = Left(sFilename, i-1)
      sFilename = Trim(sFilename, Any Chr(32,34))  ' remove spaces and double quotes

      If Dir(sFilename) > "" Then gpApp->IncludeFilename = sFilename
      If Dir(Curdir & "\" & sFilename) > "" Then gpApp->IncludeFilename = Curdir & "\" & sFilename
      If Dir(Exepath & "\" & sFilename) > "" Then gpApp->IncludeFilename = Exepath & "\" & sFilename
      If Dir(Str(wCompilerPath) & "\" & sFilename) > "" Then gpApp->IncludeFilename = Str(wCompilerPath) & "\" & sFilename
      gpApp->IncludeFilename = FF_Replace(gpApp->IncludeFilename, "/", "\")
      
      If Dir(gpApp->IncludeFilename) > "" Then        
         AppendMenuW hPopupMenu, MF_ENABLED, IDM_OPENINCLUDE, "Open " & Chr(34) & gpApp->IncludeFilename & Chr(34)
         AppendMenuW hPopupMenu, MF_SEPARATOR, 0, ""
      End If
   End If

   AppendMenuW hPopupMenu, MF_ENABLED, IDM_UNDO, "&Undo"
   AppendMenuW hPopupMenu, MF_ENABLED, IDM_REDO, "Red&o"
   AppendMenuW hPopupMenu, MF_SEPARATOR, 0, ""
   AppendMenuW hPopupMenu, MF_ENABLED, IDM_CUT, "Cu&t"
   AppendMenuW hPopupMenu, MF_ENABLED, IDM_COPY, "&Copy"
   AppendMenuW hPopupMenu, MF_ENABLED, IDM_PASTE, "&Paste"
   AppendMenuW hPopupMenu, MF_SEPARATOR, 0, ""
   AppendMenuW hPopupMenu, MF_ENABLED, IDM_INDENTBLOCK, "Block &Indent" 
   AppendMenuW hPopupMenu, MF_ENABLED, IDM_UNINDENTBLOCK, "Block &Unindent"
   AppendMenuW hPopupMenu, MF_SEPARATOR, 0, ""
   AppendMenuW hPopupMenu, MF_ENABLED, IDM_COMMENTBLOCK, "&Block Comment" 
   AppendMenuW hPopupMenu, MF_ENABLED, IDM_UNCOMMENTBLOCK, "Bloc&k Uncomment"
   AppendMenuW hPopupMenu, MF_SEPARATOR, 0, ""
   AppendMenuW hPopupMenu, MF_ENABLED, IDM_SELECTALL, "Select &All"

   Function = hPopupMenu

End Function
    