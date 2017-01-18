'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2017 Paul Squires, PlanetSquires Software
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


' ========================================================================================
' Build the frmMain topmenu
' ========================================================================================
Function frmMain_BuildMenu ( ByVal pWindow As CWindow Ptr ) As HMENU

   Dim As HMENU hMenu, hPopUpMenu, hPopUpSubMenu

   hMenu = CreateMenu
   hPopUpMenu = CreatePopupMenu
      AppendMenuW hMenu, MF_POPUP Or MF_ENABLED, Cast(UINT_PTR, hPopUpMenu), L(2,"&File")
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_FILENEW, L(3,"&New") & Chr(9) & "Ctrl+N"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_FILEOPEN, L(4,"&Open...") & Chr(9) & "Ctrl+O"
         AppendMenuW hPopUpMenu, MF_SEPARATOR, 0, ""
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_FILECLOSE, L(5,"&Close") & Chr(9) & "Ctrl+Q"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_FILECLOSEALL, L(6,"C&lose All")
         AppendMenuW hPopUpMenu, MF_SEPARATOR, 0, ""
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_FILESAVE, L(7,"&Save") & Chr(9) & "Ctrl+S"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_FILESAVEAS, L(8,"Save &As") & Chr(9) & "F12"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_FILESAVEALL, L(9,"Sa&ve All") & Chr(9) & "Ctrl+Shift+S"
         AppendMenuW hPopUpMenu, MF_SEPARATOR, 0, ""
         hPopUpSubMenu = CreatePopupMenu
            AppendMenuW hPopUpMenu, MF_POPUP Or MF_ENABLED, Cast(UINT_PTR, hPopUpSubMenu), L(10,"&Recent Files")
              AppendMenuW hPopUpSubMenu, MF_ENABLED, IDM_MRU, L(11,"(Empty)")
         AppendMenuW hPopUpMenu, MF_SEPARATOR, 0, ""
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_COMMAND, L(12,"Co&mmand Prompt...")
         AppendMenuW hPopUpMenu, MF_SEPARATOR, 0, ""
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_EXIT, L(13,"E&xit") & Chr(9) & "Alt+F4"

   hPopUpMenu = CreatePopupMenu
      AppendMenuW hMenu, MF_POPUP Or MF_ENABLED, Cast(UINT_PTR, hPopUpMenu), L(14,"&Edit")
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_UNDO, L(15,"&Undo") & Chr(9) & "Ctrl+Z"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_REDO, L(16,"&Redo") & Chr(9) & "Ctrl+E"
         AppendMenuW hPopUpMenu, MF_SEPARATOR, 0, ""
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_CUT, L(17,"Cu&t") & Chr(9) & "Ctrl+X"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_COPY, L(18,"&Copy") & Chr(9) & "Ctrl+C"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_PASTE, L(19,"&Paste") & Chr(9) & "Ctrl+V"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_DELETELINE, L(20,"&Delete Line") & Chr(9) & "Ctrl+Y"
         AppendMenuW hPopUpMenu, MF_SEPARATOR, 0, ""
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_INSERTFILE, L(21,"Insert &File...") & Chr(9) & "Ctrl+I"
         AppendMenuW hPopUpMenu, MF_SEPARATOR, 0, ""
         hPopUpSubMenu = CreatePopupMenu
            AppendMenuW hPopUpMenu, MF_POPUP Or MF_ENABLED, Cast(UINT_PTR, hPopUpSubMenu), L(22,"&Block Operations")
              AppendMenuW hPopUpSubMenu, MF_ENABLED, IDM_INDENTBLOCK, L(23,"&Indent Block") & Chr(9) & "TAB"
              AppendMenuW hPopUpSubMenu, MF_ENABLED, IDM_UNINDENTBLOCK, L(24,"U&nIndent Block") & Chr(9) & "Shift+TAB"
              AppendMenuW hPopUpSubMenu, MF_SEPARATOR, 0, ""
              AppendMenuW hPopUpSubMenu, MF_ENABLED, IDM_COMMENTBLOCK, L(25,"Co&mment Block") & Chr(9) & "Ctrl+B"
              AppendMenuW hPopUpSubMenu, MF_ENABLED, IDM_UNCOMMENTBLOCK, L(26,"UnComment &Block") & Chr(9) & "Ctrl+Shift+B"
         hPopUpSubMenu = CreatePopupMenu
            AppendMenuW hPopUpMenu, MF_POPUP Or MF_ENABLED, Cast(UINT_PTR, hPopUpSubMenu), L(27,"&Line Operations")
              AppendMenuW hPopUpSubMenu, MF_ENABLED, IDM_DUPLICATELINE, L(28,"&Duplicate Current Line") & Chr(9) & "Ctrl+D"
              AppendMenuW hPopUpSubMenu, MF_ENABLED, IDM_MOVELINEUP, L(29,"Move &Up Current Line") & Chr(9) & "Ctrl+Shift+Up"
              AppendMenuW hPopUpSubMenu, MF_ENABLED, IDM_MOVELINEDOWN, L(30,"Move &Down Current Line") & Chr(9) & "Ctrl+Shift+Down"
         hPopUpSubMenu = CreatePopupMenu
            AppendMenuW hPopUpMenu, MF_POPUP Or MF_ENABLED, Cast(UINT_PTR, hPopUpSubMenu), L(31,"&Case Conversions")
              AppendMenuW hPopUpSubMenu, MF_ENABLED, IDM_TOUPPERCASE, L(32,"Selection to Upp&er Case") & Chr(9) & "Ctrl+Alt+U"
              AppendMenuW hPopUpSubMenu, MF_ENABLED, IDM_TOLOWERCASE, L(33,"Selection to &Lower Case") & Chr(9) & "Ctrl+Alt+L"
              AppendMenuW hPopUpSubMenu, MF_ENABLED, IDM_TOMIXEDCASE, L(34,"Selection to Mi&xed Case") & Chr(9) & "Ctrl+Alt+M"
         hPopUpSubMenu = CreatePopupMenu
            AppendMenuW hPopUpMenu, MF_POPUP Or MF_ENABLED, Cast(UINT_PTR, hPopUpSubMenu), L(35,"&End of Line Character Conversions")
              AppendMenuW hPopUpSubMenu, MF_ENABLED, IDM_EOLTOCRLF, L(36,"Convert to CRLF")
              AppendMenuW hPopUpSubMenu, MF_ENABLED, IDM_EOLTOCR, L(37,"Convert to CR")
              AppendMenuW hPopUpSubMenu, MF_ENABLED, IDM_EOLTOLF, L(38,"Convert to LF")
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_TABSTOSPACES, L(39,"Replace Tabs &with Spaces")
'         hPopUpSubMenu = CreatePopupMenu
'            AppendMenuW hPopUpMenu, MF_POPUP Or MF_ENABLED, Cast(UINT_PTR, hPopUpSubMenu), "Replace &Spaces or Tabs"
'              AppendMenuW hPopUpSubMenu, MF_ENABLED, IDM_SPACESTOTABS, "Replace Spaces with Tabs"
'              AppendMenuW hPopUpSubMenu, MF_ENABLED, IDM_TABSTOSPACES, "Replace Tabs with Spaces"
         AppendMenuW hPopUpMenu, MF_SEPARATOR, 0, ""
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_SELECTALL, L(40,"Select &All") & Chr(9) & "Ctrl+A"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_SELECTLINE, L(41,"&Select Line") & Chr(9) & "Ctrl+L"

   hPopUpMenu = CreatePopupMenu
      AppendMenuW hMenu, MF_POPUP Or MF_ENABLED, Cast(UINT_PTR, hPopUpMenu), L(42,"&Search")
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_FIND, L(43,"&Find...") & Chr(9) & "Ctrl+F"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_FINDNEXT, L(44,"Find &Next") & Chr(9) & "F3"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_FINDPREV, L(45,"Find &Previous") & Chr(9) & "Shift+F3"
'         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_FINDINFILES, L(257,"Find &In Files...") & Chr(9) & "Ctrl+Shift+F"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_REPLACE, L(46,"&Replace...") & Chr(9) & "Ctrl+R"
         AppendMenuW hPopUpMenu, MF_SEPARATOR, 0, ""
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_DEFINITION, L(47,"Sub/Function &Definition") & Chr(9) & "F6"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_LASTPOSITION, L(48,"&Last Position") & Chr(9) & "Shift+F6"
         AppendMenuW hPopUpMenu, MF_SEPARATOR, 0, ""
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_GOTO, L(49,"&Goto Line...") & Chr(9) & "Ctrl+G"
         AppendMenuW hPopUpMenu, MF_SEPARATOR, 0, ""
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_BOOKMARKTOGGLE, L(50,"&Toggle Bookmark") & Chr(9) & "Ctrl+F2"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_BOOKMARKNEXT, L(51,"Next &Bookmark") & Chr(9) & "F2"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_BOOKMARKPREV, L(52,"Pre&vious Bookmark") & Chr(9) & "Shift+F2"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_BOOKMARKCLEARALL, L(53,"&Clear Bookmarks") & Chr(9) & "Ctrl+Shift+F2"

   hPopUpMenu = CreatePopupMenu
      AppendMenuW hMenu, MF_POPUP Or MF_ENABLED, Cast(UINT_PTR, hPopUpMenu), L(54,"&View")
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_FOLDTOGGLE, L(55,"Toggle &Current Fold Point") & Chr(9) & "F8"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_FOLDBELOW, L(56,"Toggle Current And All &Below") & Chr(9) & "Ctrl+F8"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_FOLDALL, L(57,"Fold &All") & Chr(9) & "Alt+F8"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_UNFOLDALL, L(58,"&Unfold All") & Chr(9) & "Shift+F8"
         AppendMenuW hPopUpMenu, MF_SEPARATOR, 0, ""
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_ZOOMIN, L(59,"Zoom &In") & Chr(9) & "Ctrl++"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_ZOOMOUT, L(60,"Zoom &Out") & Chr(9) & "Ctrl+-"
         AppendMenuW hPopUpMenu, MF_SEPARATOR, 0, ""
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_FUNCTIONLIST, L(221,"View &Function List") & Chr(9) & "F4"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_VIEWEXPLORER, L(250,"View &Explorer Window") & Chr(9) & "F9"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_VIEWOUTPUT, L(251,"View Ou&tput Window") & Chr(9) & "Ctrl+F9"
         AppendMenuW hPopUpMenu, MF_SEPARATOR, 0, ""
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_RESTOREMAIN, L(61,"&Restore Main Window Size")
         AppendMenuW hPopUpMenu, MF_SEPARATOR, 0, ""
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_OPTIONS, L(62,"En&vironment Options...")

   hPopUpMenu = CreatePopupMenu
      AppendMenuW hMenu, MF_POPUP Or MF_ENABLED, Cast(UINT_PTR, hPopUpMenu), L(63,"&Project")
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_PROJECTNEW, L(180,"&New Project")
         AppendMenuW hPopUpMenu, MF_SEPARATOR, 0, ""
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_PROJECTOPEN, L(182,"&Open Project...") 
         hPopUpSubMenu = CreatePopupMenu
            AppendMenuW hPopUpMenu, MF_POPUP Or MF_ENABLED, Cast(UINT_PTR, hPopUpSubMenu), L(219,"&Recent Projects")
              AppendMenuW hPopUpSubMenu, MF_ENABLED, IDM_MRUPROJECT, L(11,"(Empty)")
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_PROJECTCLOSE, L(183,"&Close Project")
         AppendMenuW hPopUpMenu, MF_SEPARATOR, 0, ""
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_PROJECTSAVE, L(184,"&Save Project")
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_PROJECTSAVEAS, L(185,"Save Project &As...")
         AppendMenuW hPopUpMenu, MF_SEPARATOR, 0, ""
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_PROJECTFILESADD, L(207,"&Add Files to Project...")
         AppendMenuW hPopUpMenu, MF_SEPARATOR, 0, ""
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_PROJECTOPTIONS, L(186,"Project O&ptions...")

   hPopUpMenu = CreatePopupMenu
      AppendMenuW hMenu, MF_POPUP Or MF_ENABLED, Cast(UINT_PTR, hPopUpMenu), L(64,"&Compile")
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_BUILDEXECUTE, L(65,"Build And E&xecute") & Chr(9) & "F5"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_COMPILE, L(66,"&Compile") & Chr(9) & "Ctrl+F5"
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_RUNEXE, L(67,"&Run Executable") & Chr(9) & "Shift+F5"
         AppendMenuW hPopUpMenu, MF_SEPARATOR, 0, ""
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_COMMANDLINE, L(68,"Command &Line...")
         AppendMenuW hPopUpMenu, MF_SEPARATOR, 0, ""
         AppendMenuW hPopUpMenu, MF_ENABLED Or MFT_RADIOCHECK, IDM_USE32BIT, L(69,"Use &32-Bit Compiler") & Chr(9) & "Ctrl+1"
         AppendMenuW hPopUpMenu, MF_ENABLED Or MFT_RADIOCHECK, IDM_USE64BIT, L(70,"Use &64-Bit Compiler") & Chr(9) & "Ctrl+2"
         AppendMenuW hPopUpMenu, MF_SEPARATOR, 0, ""
         AppendMenuW hPopUpMenu, MF_ENABLED Or MFT_RADIOCHECK, IDM_GUI, L(71,"Compile &GUI application") & Chr(9) & "Ctrl+3"
         AppendMenuW hPopUpMenu, MF_ENABLED Or MFT_RADIOCHECK, IDM_CONSOLE, L(72,"Compile C&onsole application") & Chr(9) & "Ctrl+4"

   hPopUpMenu = CreatePopupMenu
      AppendMenuW hMenu, MF_POPUP Or MF_ENABLED, Cast(UINT_PTR, hPopUpMenu), L(73,"&Help")
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_HELP, L(73,"&Help") & Chr(9) & "F1"
         AppendMenuW hPopUpMenu, MF_SEPARATOR, 0, ""
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_ABOUT, L(74,"&About")

   Function = hMenu

   '  Create the accelerator table for the menu
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL, VK_TAB,  IDM_NEXTTAB )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL Or FSHIFT, VK_TAB,  IDM_PREVTAB )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL, VK_W,  IDM_CLOSETAB )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL, VK_N,  IDM_FILENEW )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL, VK_O,  IDM_FILEOPEN )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL, VK_Q,  IDM_FILECLOSE )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL, VK_S,  IDM_FILESAVE )
   pWindow->AddAccelerator( FVIRTKEY, VK_F12,  IDM_FILESAVEAS )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL Or FSHIFT, VK_S,  IDM_FILESAVEALL )
   pWindow->AddAccelerator( FVIRTKEY Or FALT,     VK_F4, IDM_EXIT )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL, VK_E,  IDM_REDO )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL, VK_Y,  IDM_DELETELINE )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL, VK_I,  IDM_INSERTFILE )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL, VK_B,  IDM_COMMENTBLOCK )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL Or FSHIFT, VK_B,  IDM_UNCOMMENTBLOCK )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL, VK_D,  IDM_DUPLICATELINE )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL Or FSHIFT, VK_UP,  IDM_MOVELINEUP )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL Or FSHIFT, VK_DOWN,  IDM_MOVELINEDOWN )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL Or FALT, VK_U,  IDM_TOUPPERCASE )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL Or FALT, VK_L,  IDM_TOLOWERCASE )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL Or FALT, VK_M,  IDM_TOMIXEDCASE )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL, VK_A,  IDM_SELECTALL )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL, VK_L,  IDM_SELECTLINE )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL, VK_F,  IDM_FIND )
   pWindow->AddAccelerator( FVIRTKEY, VK_F3,  IDM_FINDNEXT )
   pWindow->AddAccelerator( FVIRTKEY Or FSHIFT, VK_F3,  IDM_FINDPREV )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL Or FSHIFT, VK_F, IDM_FINDINFILES)
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL, VK_R,  IDM_REPLACE )
   pWindow->AddAccelerator( FVIRTKEY, VK_F6,  IDM_DEFINITION )
   pWindow->AddAccelerator( FVIRTKEY Or FSHIFT, VK_F6,  IDM_LASTPOSITION )
   pWindow->AddAccelerator( FVIRTKEY, VK_F4,  IDM_FUNCTIONLIST )
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
   pWindow->AddAccelerator( FVIRTKEY, VK_F9,  IDM_VIEWEXPLORER)
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL, VK_F9,  IDM_VIEWOUTPUT )
   pWindow->AddAccelerator( FVIRTKEY, VK_F5,  IDM_BUILDEXECUTE )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL, VK_F5,  IDM_COMPILE )
   pWindow->AddAccelerator( FVIRTKEY Or FSHIFT, VK_F5,  IDM_RUNEXE )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL, VK_1,  IDM_USE32BIT )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL, VK_2,  IDM_USE64BIT )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL, VK_3,  IDM_GUI )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL, VK_4,  IDM_CONSOLE )
   pWindow->AddAccelerator( FVIRTKEY, VK_F1, IDM_HELP )
   pWindow->CreateAcceleratorTable()

End Function


' ========================================================================================
' Enable/Disable top menu items depending on application state
' ========================================================================================
Function frmMain_ChangeTopMenuStates() As Long

   Dim hMenu As HMENU = GetMenu(HWND_FRMMAIN)
   Dim idMenu      As Long
   Dim startSelPos As Long
   Dim endSelPos   As Long
   Dim nSelected   As Long 
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
   EnableMenuItem hMenu, IDM_EXIT, MF_BYCOMMAND Or MF_ENABLED
   EnableMenuItem hMenu, IDM_VIEWEXPLORER, MF_BYCOMMAND Or MF_ENABLED
   EnableMenuItem hMenu, IDM_VIEWOUTPUT, MF_BYCOMMAND Or MF_ENABLED
   EnableMenuItem hMenu, IDM_RESTOREMAIN, MF_BYCOMMAND Or MF_ENABLED
   EnableMenuItem hMenu, IDM_OPTIONS, MF_BYCOMMAND Or MF_ENABLED
   EnableMenuItem hMenu, IDM_PROJECTNEW, MF_BYCOMMAND Or MF_ENABLED
   EnableMenuItem hMenu, IDM_PROJECTOPEN, MF_BYCOMMAND Or MF_ENABLED
   EnableMenuItem hMenu, IDM_USE32BIT, MF_BYCOMMAND Or MF_ENABLED
   EnableMenuItem hMenu, IDM_USE64BIT, MF_BYCOMMAND Or MF_ENABLED
   EnableMenuItem hMenu, IDM_GUI, MF_BYCOMMAND Or MF_ENABLED
   EnableMenuItem hMenu, IDM_CONSOLE, MF_BYCOMMAND Or MF_ENABLED
   EnableMenuItem hMenu, IDM_HELP, MF_BYCOMMAND Or MF_ENABLED
   EnableMenuItem hMenu, IDM_ABOUT, MF_BYCOMMAND Or MF_ENABLED


   ' Check the correct default compiler
   Select Case UCase(gConfig.DefaultCompiler)
      Case "FBC 32BIT": nSelected = IDM_USE32BIT
      Case "FBC 64BIT": nSelected = IDM_USE64BIT
   End Select
   CheckMenuRadioItem( hMenu, IDM_USE32BIT, IDM_USE64BIT, nSelected, MF_BYCOMMAND )
   
   ' Check the correct default compile mode
   Select Case Ucase(gConfig.DefaultCompileMode)
      Case "GUI":       nSelected = IDM_GUI
      Case "CONSOLE":   nSelected = IDM_CONSOLE
   End Select
   CheckMenuRadioItem( hMenu, IDM_GUI, IDM_CONSOLE, nSelected, MF_BYCOMMAND )
   
   ' Enable everything if a document is open
   Dim As Long NumTabs = TabCtrl_GetItemCount(gTTabCtl.hWindow)
   If NumTabs > 0 Then
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
   
   ' Check if a Project is open and active
   Dim As Long fEnabled = Iif( gApp.IsProjectActive, MF_ENABLED, MF_GRAYED )
   EnableMenuItem hMenu, IDM_PROJECTMANAGER, MF_BYCOMMAND Or fEnabled
   EnableMenuItem hMenu, IDM_PROJECTCLOSE, MF_BYCOMMAND Or fEnabled
   EnableMenuItem hMenu, IDM_PROJECTSAVE, MF_BYCOMMAND Or fEnabled
   EnableMenuItem hMenu, IDM_PROJECTSAVEAS, MF_BYCOMMAND Or fEnabled
   EnableMenuItem hMenu, IDM_PROJECTFILESADD, MF_BYCOMMAND Or fEnabled
   EnableMenuItem hMenu, IDM_PROJECTOPTIONS, MF_BYCOMMAND Or fEnabled

   If (gApp.IsProjectActive = True) OrElse cbool(NumTabs > 0) Then
      fEnabled = MF_ENABLED
   Else
      fEnabled = MF_GRAYED 
   End If
   EnableMenuItem hMenu, IDM_BUILDEXECUTE, MF_BYCOMMAND Or fEnabled
   EnableMenuItem hMenu, IDM_COMPILE, MF_BYCOMMAND Or fEnabled
   EnableMenuItem hMenu, IDM_COMMANDLINE, MF_BYCOMMAND Or fEnabled

   If AfxFileExists(gConfig.LastRunFilename) Then
      EnableMenuItem hMenu, IDM_RUNEXE, MF_BYCOMMAND Or MF_ENABLED
   Else
      EnableMenuItem hMenu, IDM_RUNEXE, MF_BYCOMMAND Or MF_GRAYED
   End If

   ' If at least one document exists then allow the Function List to be shown.
   EnableMenuItem hMenu, IDM_FUNCTIONLIST, MF_BYCOMMAND Or Iif( gApp.GetDocumentCount, MF_ENABLED, MF_GRAYED)

   Function = 0
End Function


''
''
Function frmMain_MenuSetCompiler( ByVal wID As Long ) As Long

   Select Case wID
      Case IDM_USE32BIT:  gConfig.DefaultCompiler = "FBC 32bit"
      Case IDM_USE64BIT:  gConfig.DefaultCompiler = "FBC 64bit"
   End Select
   
   frmMain_UpdateLineCol( HWND_FRMMAIN )

   Function = 0
End Function


''
''
Function frmMain_MenuSetCompileMode( ByVal wID As Long ) As Long

   Select Case wID
      Case IDM_GUI:      gConfig.DefaultCompileMode = "GUI"
      Case IDM_CONSOLE:  gConfig.DefaultCompileMode = "CONSOLE"
   End Select
   
   frmMain_UpdateLineCol( HWND_FRMMAIN )

   Function = 0
End Function


''
''
Function AddProjectFileTypesToMenu( ByVal hPopUpMenu As HMENU, _
                                    ByVal pDoc As clsDocument Ptr, _
                                    byval fNoSeparator as BOOLEAN = false _
                                    ) As Long

   if fNoSeparator = false then
      AppendMenuW hPopUpMenu, MF_SEPARATOR, 0, "" 
   end if
   AppendMenuW hPopUpMenu, MF_ENABLED Or MFT_RADIOCHECK, IDM_SETFILENORMAL,   L(204,"Normal file") 
   AppendMenuW hPopUpMenu, MF_ENABLED Or MFT_RADIOCHECK, IDM_SETFILEMODULE,   L(205,"Module file") 
   AppendMenuW hPopUpMenu, MF_ENABLED Or MFT_RADIOCHECK, IDM_SETFILEMAIN,     L(206,"Main file") 
   AppendMenuW hPopUpMenu, MF_ENABLED Or MFT_RADIOCHECK, IDM_SETFILERESOURCE, L(209,"Resource file") 

   Dim nFileType As Long
   Select Case pDoc->ProjectFileType
      Case FILETYPE_NORMAL:   nFileType = IDM_SETFILENORMAL
      Case FILETYPE_MODULE:   nFileType = IDM_SETFILEMODULE
      Case FILETYPE_MAIN:     nFileType = IDM_SETFILEMAIN
      Case FILETYPE_RESOURCE: nFileType = IDM_SETFILERESOURCE
   End Select
   CheckMenuRadioItem( hPopupMenu, IDM_SETFILENORMAL, IDM_SETFILERESOURCE, nFileType, MF_BYCOMMAND )

   Function = 0
End Function


''
''
Function CreateStatusBarFileTypeContextMenu() As HMENU
   Dim hPopUpMenu As HMENU
   Dim wszText As WString * MAX_PATH
   
   Dim pDoc As clsDocument Ptr = gTTabCtl.GetActiveDocumentPtr()
   If pDoc Then
      If gApp.IsProjectActive Then
         hPopUpMenu = CreatePopupMenu
         AddProjectFileTypesToMenu( hPopUpMenu, pDoc, true ) 
      End If
   End If

   Function = hPopupMenu
End Function


''
''
Function CreateTopTabCtlContextMenu( ByVal idx As Long ) As HMENU
   Dim hPopUpMenu As HMENU
   Dim wszText As WString * MAX_PATH
   
   hPopUpMenu = CreatePopupMenu
   AppendMenuW hPopUpMenu, MF_ENABLED, IDM_FILECLOSE,  L(5,"Close") 
   AppendMenuW hPopUpMenu, MF_ENABLED, IDM_FILESAVE,   L(7,"Save") 
   AppendMenuW hPopUpMenu, MF_ENABLED, IDM_FILESAVEAS, L(8,"Save As") 
   AppendMenuW hPopUpMenu, MF_SEPARATOR, 0, "" 
   AppendMenuW hPopUpMenu, MF_ENABLED, IDM_FILECLOSEALL, L(6,"Close All")

   Dim pDoc As clsDocument Ptr = gTTabCtl.GetDocumentPtr(idx)
   If pDoc Then
      If (gApp.IsProjectActive = True) Then
        AppendMenuW hPopUpMenu, MF_SEPARATOR, 0, ""
        If pDoc->IsProjectFile Then     
           wszText = L(190,"&Remove from project")
           AppendMenuW hPopUpMenu, MF_ENABLED, IDM_REMOVEFILEFROMPROJECT, wszText
           AddProjectFileTypesToMenu(hPopUpMenu, pDoc)
        Else
           wszText = L(189,"A&dd to project")
           AppendMenuW hPopUpMenu, MF_ENABLED, IDM_ADDFILETOPROJECT, wszText 
        End If
      End If
   End If

   Function = hPopupMenu
End Function


''
''
Function CreateExplorerContextMenu( ByVal pDoc As clsDocument Ptr ) As HMENU
   
   If pDoc = 0 Then Exit Function
   Dim wszText As WString * MAX_PATH
   
   Dim hPopUpMenu As HMENU = CreatePopupMenu()

   If (gApp.IsProjectActive = True) Then
      If pDoc->IsProjectFile Then     
         wszText = L(190,"&Remove from project") 
         AppendMenuW hPopUpMenu, MF_ENABLED, IDM_REMOVEFILEFROMPROJECT, wszText
         AddProjectFileTypesToMenu(hPopUpMenu, pDoc)
      End If
   Else
      AppendMenuW hPopUpMenu, MF_ENABLED, IDM_FILECLOSE,  L(5,"Close") 
      AppendMenuW hPopUpMenu, MF_ENABLED, IDM_FILESAVE,   L(7,"Save") 
      AppendMenuW hPopUpMenu, MF_ENABLED, IDM_FILESAVEAS, L(8,"Save As") 
   End If

   Function = hPopupMenu
 
End Function

''
''
Function CreateScintillaContextMenu() As HMENU

   Dim pDoc As clsDocument Ptr = gTTabCtl.GetActiveDocumentPtr()
   If pDoc = 0 Then Exit Function

   Dim hPopupMenu    As HMENU
   Dim wPath         As WString * MAX_PATH
   Dim wCompilerPath As WString * MAX_PATH
   Dim wzText        As WString * MAX_PATH
   Dim sFilename     As String 
   Dim hWindow       As HWnd   = pDoc->hWindow
   Dim nLine         As Long   = pDoc->GetCurrentLineNumber()
   Dim sLine         As String = LTrim(pDoc->GetLine(nLine))
   Dim i             As Long
   Dim pt            As Point
   
   ' Ensure that the window under the cursor is actually a Scintilla window
   ' and not another window like a tab in the top tabcontrol.
   GetCursorPos(@pt)
   ScreenToClient(HWND_FRMMAIN, @pt)

   If RealChildWindowFromPoint(HWND_FRMMAIN, pt) <> pDoc->hWindow Then Exit Function
   
   ' Determine the compiler path in case it needs to be prepended to the filename
   gConfig.FBWINCompiler32 = ProcessFromCurdrive(gConfig.FBWINCompiler32)
   gConfig.FBWINCompiler64 = ProcessFromCurdrive(gConfig.FBWINCompiler64)
   Select Case Ucase(gConfig.DefaultCompiler)
      Case "FBC 32BIT":   wCompilerPath = gConfig.FBWINcompiler32
      Case "FBC 64BIT":   wCompilerPath = gConfig.FBWINcompiler64
   End Select
   wCompilerPath = AfxStrPathname( "PATH", wCompilerPath )
   wCompilerPath = wCompilerPath & "inc"


   hPopupMenu = CreatePopupMenu

   If Left(Ucase(sLine), 9) = "#INCLUDE " Then sFilename = Mid(sLine, 10)
   If Left(Ucase(sLine), 14) = "#INCLUDE ONCE " Then sFilename = Mid(sLine, 15)

   If Len(sFilename) Then
      ' remove any comments at the end of the line
      i = Instr(sFilename, "'")
      If i Then sFilename = Left(sFilename, i-1)
      sFilename = Trim(sFilename, Any Chr(32,34))  ' remove spaces and double quotes
      wPath = AfxStrPathname( "PATH", pDoc->DiskFilename )

      If AfxFileExists(wPath & sFilename) Then 
         gApp.IncludeFilename = wPath & sFilename
      ElseIf AfxFileExists(sFilename) Then 
         gApp.IncludeFilename = sFilename
      ElseIf AfxFileExists(Curdir & "\" & sFilename) Then 
         gApp.IncludeFilename = Curdir & "\" & sFilename
      ElseIf AfxFileExists(Exepath & "\" & sFilename) Then 
         gApp.IncludeFilename = Exepath & "\" & sFilename
      ElseIf AfxFileExists(Str(wCompilerPath) & "\" & sFilename) Then 
         gApp.IncludeFilename = Str(wCompilerPath) & "\" & sFilename
      End If
      gApp.IncludeFilename = AfxStrReplace(gApp.IncludeFilename, "/", "\")
      
      If AfxFileExists(gApp.IncludeFilename) Then        
         AppendMenuW hPopupMenu, MF_ENABLED, IDM_OPENINCLUDE, L(75,"Open") + Chr(32, 34) + gApp.IncludeFilename + Chr(34)
         AppendMenuW hPopupMenu, MF_SEPARATOR, 0, ""
      End If
   End If

   AppendMenuW hPopupMenu, MF_ENABLED, IDM_UNDO, L(15,"&Undo")
   AppendMenuW hPopupMenu, MF_ENABLED, IDM_REDO, L(16,"Red&o")
   AppendMenuW hPopupMenu, MF_SEPARATOR, 0, ""
   AppendMenuW hPopupMenu, MF_ENABLED, IDM_CUT, L(17,"Cu&t")
   AppendMenuW hPopupMenu, MF_ENABLED, IDM_COPY, L(18,"&Copy")
   AppendMenuW hPopupMenu, MF_ENABLED, IDM_PASTE, L(19,"&Paste")
   AppendMenuW hPopupMenu, MF_SEPARATOR, 0, ""
   AppendMenuW hPopupMenu, MF_ENABLED, IDM_INDENTBLOCK, L(23,"Block &Indent") 
   AppendMenuW hPopupMenu, MF_ENABLED, IDM_UNINDENTBLOCK, L(24,"Block &Unindent")
   AppendMenuW hPopupMenu, MF_SEPARATOR, 0, ""
   AppendMenuW hPopupMenu, MF_ENABLED, IDM_COMMENTBLOCK, L(25,"&Block Comment") 
   AppendMenuW hPopupMenu, MF_ENABLED, IDM_UNCOMMENTBLOCK, L(26,"Bloc&k Uncomment")
   AppendMenuW hPopupMenu, MF_SEPARATOR, 0, ""
   AppendMenuW hPopupMenu, MF_ENABLED, IDM_SELECTALL, L(40,"Select &All")
   
   If (gApp.IsProjectActive = True) Then
      AppendMenuW hPopupMenu, MF_SEPARATOR, 0, ""
      If pDoc->IsProjectFile Then     
         wzText = L(190,"&Remove from project")
         AppendMenuW hPopupMenu, MF_ENABLED, IDM_REMOVEFILEFROMPROJECT, wzText
         AddProjectFileTypesToMenu(hPopUpMenu, pDoc)
      Else
         wzText = L(189,"A&dd to project") 
         AppendMenuW hPopupMenu, MF_ENABLED, IDM_ADDFILETOPROJECT, wzText 
      End If
   End If
   
   Function = hPopupMenu

End Function
    





















