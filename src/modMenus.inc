'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2018 Paul Squires, PlanetSquires Software
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
      AppendMenu hMenu, MF_POPUP Or MF_ENABLED, Cast(UINT_PTR, hPopUpMenu), L(2,"&File")
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_FILENEW, L(3,"&New") & Chr(9) & "Ctrl+N"
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_FILEOPEN, L(4,"&Open...") & Chr(9) & "Ctrl+O"
         AppendMenu hPopUpMenu, MF_SEPARATOR, 0, ""
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_FILECLOSE, L(5,"&Close") & Chr(9) & "Ctrl+Q"
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_FILECLOSEALL, L(6,"C&lose All")
         AppendMenu hPopUpMenu, MF_SEPARATOR, 0, ""
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_FILESAVE, L(7,"&Save") & Chr(9) & "Ctrl+S"
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_FILESAVEAS, L(8,"Save &As") & Chr(9) & "F12"
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_FILESAVEALL, L(9,"Sa&ve All") & Chr(9) & "Ctrl+Shift+S"
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_FILESAVEDECLARES, L(265,"Save &Declares File") & Chr(9) & "Ctrl+Shift+D"
         AppendMenu hPopUpMenu, MF_SEPARATOR, 0, ""
         hPopUpSubMenu = CreatePopupMenu
            AppendMenu hPopUpMenu, MF_POPUP Or MF_ENABLED, Cast(UINT_PTR, hPopUpSubMenu), L(10,"&Recent Files")
              AppendMenu hPopUpSubMenu, MF_ENABLED, IDM_MRU, L(11,"(Empty)")
         AppendMenu hPopUpMenu, MF_SEPARATOR, 0, ""
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_COMMAND, L(12,"Co&mmand Prompt...")
         AppendMenu hPopUpMenu, MF_SEPARATOR, 0, ""
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_EXIT, L(13,"E&xit") & Chr(9) & "Alt+F4"

   hPopUpMenu = CreatePopupMenu
      AppendMenu hMenu, MF_POPUP Or MF_ENABLED, Cast(UINT_PTR, hPopUpMenu), L(14,"&Edit")
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_UNDO, L(15,"&Undo") & Chr(9) & "Ctrl+Z"
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_REDO, L(16,"&Redo") & Chr(9) & "Ctrl+E"
         AppendMenu hPopUpMenu, MF_SEPARATOR, 0, ""
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_CUT, L(17,"Cu&t") & Chr(9) & "Ctrl+X"
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_COPY, L(18,"&Copy") & Chr(9) & "Ctrl+C"
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_PASTE, L(19,"&Paste") & Chr(9) & "Ctrl+V"
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_DELETELINE, L(20,"&Delete Line") & Chr(9) & "Ctrl+Y"
         AppendMenu hPopUpMenu, MF_SEPARATOR, 0, ""
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_INSERTFILE, L(21,"Insert &File...") & Chr(9) & "Ctrl+I"
         AppendMenu hPopUpMenu, MF_SEPARATOR, 0, ""
         hPopUpSubMenu = CreatePopupMenu
            AppendMenu hPopUpMenu, MF_POPUP Or MF_ENABLED, Cast(UINT_PTR, hPopUpSubMenu), L(271,"File Encoding")
              AppendMenu hPopUpSubMenu, MF_ENABLED or MFT_RADIOCHECK, IDM_ANSI, "ANSI"
              AppendMenu hPopUpSubMenu, MF_ENABLED or MFT_RADIOCHECK, IDM_UTF8BOM, "UTF-8 (BOM)"
              AppendMenu hPopUpSubMenu, MF_ENABLED or MFT_RADIOCHECK, IDM_UTF16BOM, "UTF-16 (BOM)"
         hPopUpSubMenu = CreatePopupMenu
            AppendMenu hPopUpMenu, MF_POPUP Or MF_ENABLED, Cast(UINT_PTR, hPopUpSubMenu), L(22,"&Block Operations")
              AppendMenu hPopUpSubMenu, MF_ENABLED, IDM_INDENTBLOCK, L(23,"&Indent Block") & Chr(9) & "TAB"
              AppendMenu hPopUpSubMenu, MF_ENABLED, IDM_UNINDENTBLOCK, L(24,"U&nIndent Block") & Chr(9) & "Shift+TAB"
              AppendMenu hPopUpSubMenu, MF_SEPARATOR, 0, ""
              AppendMenu hPopUpSubMenu, MF_ENABLED, IDM_COMMENTBLOCK, L(25,"Co&mment Block") & Chr(9) & "Ctrl+B"
              AppendMenu hPopUpSubMenu, MF_ENABLED, IDM_UNCOMMENTBLOCK, L(26,"UnComment &Block") & Chr(9) & "Ctrl+Shift+B"
         hPopUpSubMenu = CreatePopupMenu
            AppendMenu hPopUpMenu, MF_POPUP Or MF_ENABLED, Cast(UINT_PTR, hPopUpSubMenu), L(27,"&Line Operations")
              AppendMenu hPopUpSubMenu, MF_ENABLED, IDM_DUPLICATELINE, L(28,"&Duplicate Current Line") & Chr(9) & "Ctrl+D"
              AppendMenu hPopUpSubMenu, MF_ENABLED, IDM_MOVELINEUP, L(29,"Move &Up Current Line") & Chr(9) & "Ctrl+Shift+Up"
              AppendMenu hPopUpSubMenu, MF_ENABLED, IDM_MOVELINEDOWN, L(30,"Move &Down Current Line") & Chr(9) & "Ctrl+Shift+Down"
         hPopUpSubMenu = CreatePopupMenu
            AppendMenu hPopUpMenu, MF_POPUP Or MF_ENABLED, Cast(UINT_PTR, hPopUpSubMenu), L(31,"&Case Conversions")
              AppendMenu hPopUpSubMenu, MF_ENABLED, IDM_TOUPPERCASE, L(32,"Selection to Upp&er Case") & Chr(9) & "Ctrl+Alt+U"
              AppendMenu hPopUpSubMenu, MF_ENABLED, IDM_TOLOWERCASE, L(33,"Selection to &Lower Case") & Chr(9) & "Ctrl+Alt+L"
              AppendMenu hPopUpSubMenu, MF_ENABLED, IDM_TOMIXEDCASE, L(34,"Selection to Mi&xed Case") & Chr(9) & "Ctrl+Alt+M"
         hPopUpSubMenu = CreatePopupMenu
            AppendMenu hPopUpMenu, MF_POPUP Or MF_ENABLED, Cast(UINT_PTR, hPopUpSubMenu), L(35,"&End of Line Character Conversions")
              AppendMenu hPopUpSubMenu, MF_ENABLED, IDM_EOLTOCRLF, L(36,"Convert to CRLF")
              AppendMenu hPopUpSubMenu, MF_ENABLED, IDM_EOLTOCR, L(37,"Convert to CR")
              AppendMenu hPopUpSubMenu, MF_ENABLED, IDM_EOLTOLF, L(38,"Convert to LF")
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_TABSTOSPACES, L(39,"Replace Tabs &with Spaces")
         AppendMenu hPopUpMenu, MF_SEPARATOR, 0, ""
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_SELECTALL, L(40,"Select &All") & Chr(9) & "Ctrl+A"
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_SELECTLINE, L(41,"&Select Line") & Chr(9) & "Ctrl+L"

   hPopUpMenu = CreatePopupMenu
      AppendMenu hMenu, MF_POPUP Or MF_ENABLED, Cast(UINT_PTR, hPopUpMenu), L(42,"&Search")
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_FIND, L(43,"&Find...") & Chr(9) & "Ctrl+F"
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_FINDNEXT, L(44,"Find &Next") & Chr(9) & "F3"
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_FINDPREV, L(45,"Find &Previous") & Chr(9) & "Shift+F3"
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_FINDINFILES, L(257,"Find &In Files...") & Chr(9) & "Ctrl+Shift+F"
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_REPLACE, L(46,"&Replace...") & Chr(9) & "Ctrl+R"
         AppendMenu hPopUpMenu, MF_SEPARATOR, 0, ""
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_DEFINITION, L(47,"Sub/Function &Definition") & Chr(9) & "F6"
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_LASTPOSITION, L(48,"&Last Position") & Chr(9) & "Shift+F6"
         AppendMenu hPopUpMenu, MF_SEPARATOR, 0, ""
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_GOTO, L(49,"&Goto Line...") & Chr(9) & "Ctrl+G"
         AppendMenu hPopUpMenu, MF_SEPARATOR, 0, ""
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_BOOKMARKTOGGLE, L(50,"&Toggle Bookmark") & Chr(9) & "Ctrl+F2"
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_BOOKMARKNEXT, L(51,"Next &Bookmark") & Chr(9) & "F2"
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_BOOKMARKPREV, L(52,"Pre&vious Bookmark") & Chr(9) & "Shift+F2"
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_BOOKMARKCLEARALL, L(53,"&Clear Bookmarks") & Chr(9) & "Ctrl+Shift+F2"

   hPopUpMenu = CreatePopupMenu
      AppendMenu hMenu, MF_POPUP Or MF_ENABLED, Cast(UINT_PTR, hPopUpMenu), L(54,"&View")
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_FOLDTOGGLE, L(55,"Toggle &Current Fold Point") & Chr(9) & "F8"
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_FOLDBELOW, L(56,"Toggle Current And All &Below") & Chr(9) & "Ctrl+F8"
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_FOLDALL, L(57,"Fold &All") & Chr(9) & "Alt+F8"
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_UNFOLDALL, L(58,"&Unfold All") & Chr(9) & "Shift+F8"
         AppendMenu hPopUpMenu, MF_SEPARATOR, 0, ""
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_ZOOMIN, L(59,"Zoom &In") & Chr(9) & "Ctrl++"
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_ZOOMOUT, L(60,"Zoom &Out") & Chr(9) & "Ctrl+-"
         AppendMenu hPopUpMenu, MF_SEPARATOR, 0, ""
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_FUNCTIONLIST, L(221,"View &Function List") & Chr(9) & "F4"
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_VIEWEXPLORER, L(250,"View &Explorer Window") & Chr(9) & "F9"
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_VIEWOUTPUT, L(251,"View Ou&tput Window") & Chr(9) & "Ctrl+F9"
         AppendMenu hPopUpMenu, MF_SEPARATOR, 0, ""
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_RESTOREMAIN, L(61,"&Restore Main Window Size")

   hPopUpMenu = CreatePopupMenu
      AppendMenu hMenu, MF_POPUP Or MF_ENABLED, Cast(UINT_PTR, hPopUpMenu), L(63,"&Project")
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_PROJECTNEW, L(180,"&New Project")
         AppendMenu hPopUpMenu, MF_SEPARATOR, 0, ""
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_PROJECTOPEN, L(182,"&Open Project...") 
         hPopUpSubMenu = CreatePopupMenu
            AppendMenu hPopUpMenu, MF_POPUP Or MF_ENABLED, Cast(UINT_PTR, hPopUpSubMenu), L(219,"&Recent Projects")
              AppendMenu hPopUpSubMenu, MF_ENABLED, IDM_MRUPROJECT, L(11,"(Empty)")
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_PROJECTCLOSE, L(183,"&Close Project")
         AppendMenu hPopUpMenu, MF_SEPARATOR, 0, ""
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_PROJECTSAVE, L(184,"&Save Project")
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_PROJECTSAVEAS, L(185,"Save Project &As...")
         AppendMenu hPopUpMenu, MF_SEPARATOR, 0, ""
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_PROJECTFILESADD, L(207,"&Add Files to Project...")
         AppendMenu hPopUpMenu, MF_SEPARATOR, 0, ""
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_PROJECTOPTIONS, L(186,"Project O&ptions...")      

   hPopUpMenu = CreatePopupMenu
      AppendMenu hMenu, MF_POPUP Or MF_ENABLED, Cast(UINT_PTR, hPopUpMenu), L(64,"&Compile")
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_BUILDEXECUTE, L(65,"Build And E&xecute") & Chr(9) & "F5"
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_COMPILE, L(66,"&Compile") & Chr(9) & "Ctrl+F5"
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_REBUILDALL, L(268,"&Rebuild All") & Chr(9) & "Ctrl+Alt+F5"
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_QUICKRUN, L(285,"&Quick Run") & Chr(9) & "Ctrl+Shift+F5"
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_RUNEXE, L(67,"&Run Executable") & Chr(9) & "Shift+F5"
         AppendMenu hPopUpMenu, MF_SEPARATOR, 0, ""
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_COMMANDLINE, L(68,"Command &Line...")

   hPopUpMenu = CreatePopupMenu
      AppendMenu hMenu, MF_POPUP Or MF_ENABLED, Cast(UINT_PTR, hPopUpMenu), L(288,"&Options")
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_OPTIONS, L(62,"En&vironment Options...")
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_COMPILECONFIG, L(277,"&Build Configurations...") & Chr(9) & "F7"
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_USERTOOLSDIALOG, L(289,"&User Tools...") & Chr(9) & "Ctrl+F7"

   hPopUpMenu = CreatePopupMenu
      AppendMenu hMenu, MF_POPUP Or MF_ENABLED, Cast(UINT_PTR, hPopUpMenu), L(309,"&Designer")
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_NEWFORM, L(310,"&New Form")
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_VIEWTOOLBOX, L(311,"View &Toolbox")
         AppendMenu hPopUpMenu, MF_SEPARATOR, 0, ""
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_MENUEDITOR, L(312,"&Menu Editor...")
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_TOOLBAREDITOR, L(313,"Tool&bar Editor...")
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_STATUSBAREDITOR, L(314,"&Statusbar Editor...")
         AppendMenu hPopUpMenu, MF_SEPARATOR, 0, ""
         hPopUpSubMenu = CreatePopupMenu
            AppendMenu hPopUpMenu, MF_POPUP Or MF_ENABLED, Cast(UINT_PTR, hPopUpSubMenu), L(315,"&Align")
               AppendMenu hPopUpSubMenu, MF_ENABLED, IDM_ALIGNLEFTS, L(316,"Lefts")
               AppendMenu hPopUpSubMenu, MF_ENABLED, IDM_ALIGNCENTERS, L(317,"Centers")
               AppendMenu hPopUpSubMenu, MF_ENABLED, IDM_ALIGNRIGHTS, L(318,"Rights")
               AppendMenu hPopUpSubMenu, MF_SEPARATOR, 0, ""
               AppendMenu hPopUpSubMenu, MF_ENABLED, IDM_ALIGNTOPS, L(319,"Tops")
               AppendMenu hPopUpSubMenu, MF_ENABLED, IDM_ALIGNMIDDLES, L(320,"Middles")
               AppendMenu hPopUpSubMenu, MF_ENABLED, IDM_ALIGNBOTTOMS, L(321,"Bottoms")
         hPopUpSubMenu = CreatePopupMenu
            AppendMenu hPopUpMenu, MF_POPUP Or MF_ENABLED, Cast(UINT_PTR, hPopUpSubMenu), L(322,"Make Same &Size")
               AppendMenu hPopUpSubMenu, MF_ENABLED, IDM_SAMEWIDTHS, L(323,"Widths")
               AppendMenu hPopUpSubMenu, MF_ENABLED, IDM_SAMEHEIGHTS, L(324,"Heights")
               AppendMenu hPopUpSubMenu, MF_ENABLED, IDM_SAMEBOTH, L(325,"Both")
         AppendMenu hPopUpMenu, MF_SEPARATOR, 0, ""
         hPopUpSubMenu = CreatePopupMenu
            AppendMenu hPopUpMenu, MF_POPUP Or MF_ENABLED, Cast(UINT_PTR, hPopUpSubMenu), L(332,"&Center in Form")
               AppendMenu hPopUpSubMenu, MF_ENABLED, IDM_CENTERHORIZ, L(333,"&Horizontally")
               AppendMenu hPopUpSubMenu, MF_ENABLED, IDM_CENTERVERT, L(334,"&Vertically")
               AppendMenu hPopUpSubMenu, MF_ENABLED, IDM_CENTERBOTH, L(325,"&Both")
         AppendMenu hPopUpMenu, MF_SEPARATOR, 0, ""
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_LOCKCONTROLS, L(335,"&Lock Controls")

   hPopUpMenu = CreatePopupMenu
      AppendMenu hMenu, MF_POPUP Or MF_ENABLED, Cast(UINT_PTR, hPopUpMenu), L(305,"&Tools")

   hPopUpMenu = CreatePopupMenu
      AppendMenu hMenu, MF_POPUP Or MF_ENABLED, Cast(UINT_PTR, hPopUpMenu), L(73,"&Help")
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_HELP, L(73,"&Help") & Chr(9) & "F1"
         AppendMenu hPopUpMenu, MF_SEPARATOR, 0, ""
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_HELPWINAPI, "WinAPI Help"
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_HELPWINFBX, "WinFBX Help"
         AppendMenu hPopUpMenu, MF_SEPARATOR, 0, ""
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_ABOUT, L(74,"&About")

   Function = hMenu

   '  Create the accelerator table for the menu
   pWindow->AddAccelerator( FVIRTKEY, VK_TAB,  IDM_INDENTBLOCK )
   pWindow->AddAccelerator( FVIRTKEY or FSHIFT, VK_TAB,  IDM_UNINDENTBLOCK )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL, VK_TAB,  IDM_NEXTTAB )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL Or FSHIFT, VK_TAB,  IDM_PREVTAB )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL, VK_W,  IDM_CLOSETAB )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL, VK_N,  IDM_FILENEW )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL, VK_O,  IDM_FILEOPEN )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL, VK_Q,  IDM_FILECLOSE )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL, VK_S,  IDM_FILESAVE )
   pWindow->AddAccelerator( FVIRTKEY, VK_F12,  IDM_FILESAVEAS )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL Or FSHIFT, VK_S,  IDM_FILESAVEALL )
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL or FSHIFT, VK_D,  IDM_FILESAVEDECLARES )
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
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL or FALT, VK_F5,  IDM_REBUILDALL)
   pWindow->AddAccelerator( FVIRTKEY Or FCONTROL or FSHIFT, VK_F5,  IDM_QUICKRUN)
   pWindow->AddAccelerator( FVIRTKEY Or FSHIFT, VK_F5,  IDM_RUNEXE )
   pWindow->AddAccelerator( FVIRTKEY, VK_F7,  IDM_COMPILECONFIG )
   pWindow->AddAccelerator( FVIRTKEY or FCONTROL, VK_F7,  IDM_USERTOOLSDIALOG )
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
   dim wzText      as WSTRING * MAX_PATH
   
   ' Start off by disabling everything in order to make things easier. (We only
   ' then need to enable things)
   For idMenu = IDM_FILE To IDM_ABOUT
      EnableMenuItem hMenu, idMenu, MF_BYCOMMAND Or MF_GRAYED
   Next
   For idMenu = IDM_ADDFILETOPROJECT To IDM_PROJECTOPTIONS
      EnableMenuItem hMenu, idMenu, MF_BYCOMMAND Or MF_GRAYED
   Next
   For idMenu = IDM_BUILDEXECUTE To IDM_LOCKCONTROLS
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
   EnableMenuItem hMenu, IDM_FUNCTIONLIST, MF_BYCOMMAND Or MF_ENABLED
   EnableMenuItem hMenu, IDM_PROJECTNEW, MF_BYCOMMAND Or MF_ENABLED
   EnableMenuItem hMenu, IDM_PROJECTOPEN, MF_BYCOMMAND Or MF_ENABLED
   EnableMenuItem hMenu, IDM_COMPILECONFIG, MF_BYCOMMAND Or MF_ENABLED
   EnableMenuItem hMenu, IDM_USERTOOLSDIALOG, MF_BYCOMMAND Or MF_ENABLED
   EnableMenuItem hMenu, IDM_NEWFORM, MF_BYCOMMAND Or MF_ENABLED
   EnableMenuItem hMenu, IDM_HELP, MF_BYCOMMAND Or MF_ENABLED
   EnableMenuItem hMenu, IDM_HELPWINAPI, MF_BYCOMMAND Or MF_ENABLED
   EnableMenuItem hMenu, IDM_HELPWINFBX, MF_BYCOMMAND Or MF_ENABLED
   EnableMenuItem hMenu, IDM_ABOUT, MF_BYCOMMAND Or MF_ENABLED

   ' Enable everything if a document is open
   Dim As Long NumTabs = TabCtrl_GetItemCount(gTTabCtl.hWindow)
   If NumTabs > 0 Then
      For idMenu = IDM_FILE To IDM_ABOUT
         EnableMenuItem hMenu, idMenu, MF_BYCOMMAND Or MF_ENABLED
      Next

      ' Enable or disable specific edit menu items (based on the currently active document)
      Dim pDoc As clsDocument Ptr = gTTabCtl.GetActiveDocumentPtr()
      If pDoc Then 
         
         dim as hwnd hEdit = pDoc->hWndActiveScintilla

         ' Enable/Disable depending on if Visual Designer or Code Editor
         if (pDoc->IsDesigner = true) andalso (IsDesignerView(pDoc) = true) THEN
            ' Disable specific code editing menu options
            ' all of Edit/Search/View menu up to ZoomOut
            ' except for cut/copy/paste and FunctionList
            for idMenu = IDM_UNDO to IDM_FUNCTIONLIST
               EnableMenuItem( hMenu, idMenu, MF_GRAYED)
            NEXT
            if pDoc->Controls.SelectedControlsCount > 0 THEN
               EnableMenuItem hMenu, IDM_CUT, MF_BYCOMMAND Or MF_ENABLED
               EnableMenuItem hMenu, IDM_COPY, MF_BYCOMMAND Or MF_ENABLED
            END IF
            EnableMenuItem hMenu, IDM_PASTE, MF_BYCOMMAND Or MF_ENABLED
            EnableMenuItem hMenu, IDM_SELECTALL, MF_BYCOMMAND Or MF_ENABLED
            CheckMenuItem( hMenu, IDM_LOCKCONTROLS, MF_BYCOMMAND or iif(pDoc->bLockControls, MF_CHECKED, MF_UNCHECKED))
         else
            ' Code editor is active
            ' Disable visual designer menu options except for "New Form"
            for idMenu = IDM_VIEWTOOLBOX to IDM_LOCKCONTROLS
               EnableMenuItem( hMenu, idMenu, MF_GRAYED)
            NEXT

            If SendMessage(hEdit, SCI_CANPASTE, 0, 0) = 0 Then 
               EnableMenuItem( hMenu, IDM_PASTE, MF_GRAYED)
            End If
            
            startSelPos = SendMessage(hEdit, SCI_GETSELECTIONSTART, 0, 0)
            endSelPos   = SendMessage(hEdit, SCI_GETSELECTIONEND, 0, 0)
            If startSelPos = endSelPos Then
               EnableMenuItem hMenu, IDM_CUT, MF_GRAYED
               EnableMenuItem hMenu, IDM_COPY, MF_GRAYED
            End If
         
            If SendMessageW(hEdit, SCI_CANUNDO, 0, 0) = 0 Then 
               EnableMenuItem hMenu, IDM_UNDO, MF_GRAYED
            End If
                  
            If SendMessage(hEdit, SCI_CANREDO, 0, 0) = 0 Then 
               EnableMenuItem hMenu, IDM_REDO, MF_GRAYED
            End If

            If SendMessage(hEdit, SCI_CANPASTE, 0, 0) = 0 Then 
               EnableMenuItem hMenu, IDM_PASTE, MF_GRAYED
            End If
         
            ' Check the File Encoding
            select CASE pDoc->FileEncoding
               case FILE_ENCODING_ANSI: nSelected = IDM_ANSI
               case FILE_ENCODING_UTF8_BOM: nSelected = IDM_UTF8BOM
               case FILE_ENCODING_UTF16_BOM: nSelected = IDM_UTF16BOM
               CASE ELSE
            END SELECT
            CheckMenuRadioItem( hMenu, IDM_ANSI, IDM_UTF16BOM, nSelected, MF_BYCOMMAND )
            EnableMenuItem hMenu, IDM_ANSI, MF_BYCOMMAND Or MF_ENABLED
            EnableMenuItem hMenu, IDM_UTF8BOM, MF_BYCOMMAND Or MF_ENABLED
            EnableMenuItem hMenu, IDM_UTF16BOM, MF_BYCOMMAND Or MF_ENABLED
         end if
         
      End If
   End If
   
   ' Check if a Project is open and active
   Dim As Long fEnabled = Iif( gApp.IsProjectActive, MF_ENABLED, MF_GRAYED )
   EnableMenuItem hMenu, IDM_ADDFILETOPROJECT, MF_BYCOMMAND Or fEnabled
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

   If gApp.IsProjectActive then 
      ' QuickRun should be disabled for Projects
      EnableMenuItem hMenu, IDM_QUICKRUN, MF_BYCOMMAND Or MF_GRAYED
   end if

   If AfxFileExists(GetRunExecutableFilename) Then
      EnableMenuItem hMenu, IDM_RUNEXE, MF_BYCOMMAND Or MF_ENABLED
   Else
      EnableMenuItem hMenu, IDM_RUNEXE, MF_BYCOMMAND Or MF_GRAYED
   End If

   Function = 0
End Function


''
''
Function AddProjectFileTypesToMenu( ByVal hPopUpMenu As HMENU, _
                                    ByVal pDoc As clsDocument Ptr, _
                                    byval fNoSeparator as BOOLEAN = false _
                                    ) As Long

   if fNoSeparator = false then
      AppendMenu hPopUpMenu, MF_SEPARATOR, 0, "" 
   end if
   AppendMenu hPopUpMenu, MF_ENABLED Or MFT_RADIOCHECK, IDM_SETFILENORMAL,   L(204,"Normal file") 
   AppendMenu hPopUpMenu, MF_ENABLED Or MFT_RADIOCHECK, IDM_SETFILEMODULE,   L(205,"Module file") 
   AppendMenu hPopUpMenu, MF_ENABLED Or MFT_RADIOCHECK, IDM_SETFILEMAIN,     L(206,"Main file") 
   AppendMenu hPopUpMenu, MF_ENABLED Or MFT_RADIOCHECK, IDM_SETFILERESOURCE, L(209,"Resource file") 

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
Function CreateStatusBarFileEncodingContextMenu() As HMENU
   Dim hPopUpMenu As HMENU
   
   Dim pDoc As clsDocument Ptr = gTTabCtl.GetActiveDocumentPtr()
   If pDoc Then
      hPopUpMenu = CreatePopupMenu
      ' winFBE forms must be UTF-16 so only allow other encodings for non-forms
      if pDoc->IsDesigner = false then
         AppendMenu hPopUpMenu, MF_ENABLED or MFT_RADIOCHECK, IDM_ANSI, "ANSI"
         AppendMenu hPopUpMenu, MF_ENABLED or MFT_RADIOCHECK, IDM_UTF8BOM, "UTF-8 (BOM)"
      end if
      AppendMenu hPopUpMenu, MF_ENABLED or MFT_RADIOCHECK, IDM_UTF16BOM, "UTF-16 (BOM)"

      Dim nEncoding As Long
      Select Case pDoc->FileEncoding
         Case FILE_ENCODING_ANSI:      nEncoding = IDM_ANSI
         Case FILE_ENCODING_UTF8_BOM:  nEncoding = IDM_UTF8BOM
         Case FILE_ENCODING_UTF16_BOM: nEncoding = IDM_UTF16BOM
      End Select
      CheckMenuRadioItem( hPopupMenu, IDM_ANSI, IDM_UTF16BOM, nEncoding, MF_BYCOMMAND )
   End If

   Function = hPopupMenu
End Function


''
''
Function CreateTopTabCtlContextMenu( ByVal idx As Long ) As HMENU
   Dim hPopUpMenu As HMENU
   Dim hPopUpSubMenu As HMENU
   Dim wszText As WString * MAX_PATH
   Dim As Long CuriTab=TabCtrl_GetCurSel(gTTabCtl.hWindow)
   Dim As Long MaxiTab=TabCtrl_GetItemCount(gTTabCtl.hWindow)-1

   hPopUpMenu = CreatePopupMenu
   AppendMenu hPopUpMenu, MF_ENABLED, IDM_FILECLOSE,  L(5,"Close") 
   AppendMenu hPopUpMenu, MF_ENABLED, IDM_FILESAVE,   L(7,"Save") 
   AppendMenu hPopUpMenu, MF_ENABLED, IDM_FILESAVEAS, L(8,"Save As") 
   AppendMenu hPopUpMenu, MF_SEPARATOR, 0, ""
   AppendMenu hPopUpMenu, iif(MaxiTab>0,MF_ENABLED,MF_DISABLED), IDM_FILECLOSEALL, L(6,"Close All")
   AppendMenu hPopUpMenu, iif(MaxiTab>0,MF_ENABLED,MF_DISABLED), IDM_FILECLOSEALLOTHERS, L(354,"Close Others")
   AppendMenu hPopUpMenu, iif(CuriTab>0,MF_ENABLED,MF_DISABLED), IDM_CLOSEALLFORWARD, L(355,"Close Tabs on the Left")
   AppendMenu hPopUpMenu, iif(MaxiTab>CuriTab,MF_ENABLED,MF_DISABLED), IDM_CLOSEALLBACKWARD, L(356,"Close Tabs on the Right") 

   Dim pDoc As clsDocument Ptr = gTTabCtl.GetDocumentPtr(idx)
   If pDoc Then
      If gApp.IsProjectActive = True Then
         AppendMenu hPopUpMenu, MF_SEPARATOR, 0, ""
         wszText = L(190,"&Remove from project")
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_REMOVEFILEFROMPROJECT, wszText
         AddProjectFileTypesToMenu(hPopUpMenu, pDoc)
      End If
   End If

   Function = hPopupMenu
End Function


''
''
Function CreateExplorerContextMenu( ByVal pDoc As clsDocument Ptr ) As HMENU
   
   Dim wszText As WString * MAX_PATH
   Dim hPopUpMenu As HMENU = CreatePopupMenu()

   ' if pDoc is zero then most likely we have right clicked on a Project or FILES root node
   if pDoc = 0 THEN
      If gApp.IsProjectActive = True Then
         dim as hwnd hTree = GetDlgItem( hWnd_FRMEXPLORER, IDC_FRMEXPLORER_TREE)
         dim as HTREEITEM hNode = Treeview_GetSelection(hTree)
         for i as long = 1 to ubound(gApp.Projects)
            if gApp.Projects(i).hExplorerRootNode = hNode THEN
               wszText = L(207,"&Add Files to Project...")
               AppendMenu( hPopUpMenu, MF_ENABLED, IDM_PROJECTFILESADDTONODE, wszText )
               exit for
            END IF
         NEXT
      END IF
   else
      If gApp.IsProjectActive Then
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_FILESAVE,   L(7,"Save") 
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_FILESAVEAS, L(8,"Save As") 
         AppendMenu hPopUpMenu, MF_SEPARATOR, 0, "" 
         wszText = L(190,"&Remove from project") 
         AppendMenu( hPopUpMenu, MF_ENABLED, IDM_REMOVEFILEFROMPROJECT, wszText )
         AddProjectFileTypesToMenu(hPopUpMenu, pDoc)
      Else
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_FILECLOSE,  L(5,"Close") 
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_FILESAVE,   L(7,"Save") 
         AppendMenu hPopUpMenu, MF_ENABLED, IDM_FILESAVEAS, L(8,"Save As") 
      End If
   END IF

   Function = hPopupMenu
 
End Function

''
''
Function CreateScintillaContextMenu() As HMENU

   Dim pDoc As clsDocument Ptr = gTTabCtl.GetActiveDocumentPtr()
   If pDoc = 0 Then Exit Function

   Dim hPopupMenu    As HMENU
   Dim hPopupSubMenu As HMENU
   Dim wPath         As WString * MAX_PATH
   Dim wCompilerPath As WString * MAX_PATH
   Dim wzText        As WString * MAX_PATH
   Dim sFilename     As String 
   Dim nLine         As Long   = pDoc->GetCurrentLineNumber()
   Dim sLine         As String = LTrim(pDoc->GetLine(nLine))
   Dim i             As Long
   Dim pt            As Point
   
   dim as hwnd hEdit = pDoc->hWndActiveScintilla

   ' Ensure that the window under the cursor is actually a Scintilla window
   ' and not another window like a tab in the top tabcontrol.
   GetCursorPos(@pt)
   ScreenToClient(HWND_FRMMAIN, @pt)

   If RealChildWindowFromPoint(HWND_FRMMAIN, pt) <> hEdit Then Exit Function
   
   ' Determine the compiler path in case it needs to be prepended to the filename
   gConfig.FBWINCompiler32 = ProcessFromCurdrive(gConfig.FBWINCompiler32)
   gConfig.FBWINCompiler64 = ProcessFromCurdrive(gConfig.FBWINCompiler64)
   Select Case gApp.Projects(gApp.GetActiveProjectIndex).GetProjectCompiler
      Case IDM_32BIT:   wCompilerPath = gConfig.FBWINcompiler32
      Case IDM_64BIT:   wCompilerPath = gConfig.FBWINcompiler64
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
      ElseIf AfxFileExists(AfxGetCurDir & "\" & sFilename) Then 
         gApp.IncludeFilename = AfxGetCurDir & "\" & sFilename
      ElseIf AfxFileExists(AfxGetExePathName & sFilename) Then 
         gApp.IncludeFilename = AfxGetExePathName & sFilename
      ElseIf AfxFileExists(Str(wCompilerPath) & "\" & sFilename) Then 
         gApp.IncludeFilename = Str(wCompilerPath) & "\" & sFilename
      End If
      gApp.IncludeFilename = AfxStrReplace(gApp.IncludeFilename, "/", "\")
      
      If AfxFileExists(gApp.IncludeFilename) Then        
         AppendMenu hPopupMenu, MF_ENABLED, IDM_OPENINCLUDE, L(75,"Open") + Chr(32, 34) + gApp.IncludeFilename + Chr(34)
         AppendMenu hPopupMenu, MF_SEPARATOR, 0, ""
      End If
   End If

   AppendMenu hPopupMenu, MF_ENABLED, IDM_UNDO, L(15,"&Undo")
   AppendMenu hPopupMenu, MF_ENABLED, IDM_REDO, L(16,"Red&o")
   AppendMenu hPopupMenu, MF_SEPARATOR, 0, ""
   AppendMenu hPopupMenu, MF_ENABLED, IDM_CUT, L(17,"Cu&t")
   AppendMenu hPopupMenu, MF_ENABLED, IDM_COPY, L(18,"&Copy")
   AppendMenu hPopupMenu, MF_ENABLED, IDM_PASTE, L(19,"&Paste")
   AppendMenu hPopupMenu, MF_SEPARATOR, 0, ""
   AppendMenu hPopupMenu, MF_ENABLED, IDM_INDENTBLOCK, L(23,"Block &Indent") 
   AppendMenu hPopupMenu, MF_ENABLED, IDM_UNINDENTBLOCK, L(24,"Block &Unindent")
   AppendMenu hPopupMenu, MF_SEPARATOR, 0, ""
   AppendMenu hPopupMenu, MF_ENABLED, IDM_COMMENTBLOCK, L(25,"&Block Comment") 
   AppendMenu hPopupMenu, MF_ENABLED, IDM_UNCOMMENTBLOCK, L(26,"Bloc&k Uncomment")
   AppendMenu hPopupMenu, MF_SEPARATOR, 0, ""
   AppendMenu hPopupMenu, MF_ENABLED, IDM_SELECTALL, L(40,"Select &All")
   
   If (gApp.IsProjectActive = True) Then
      AppendMenu hPopupMenu, MF_SEPARATOR, 0, ""
      wzText = L(190,"&Remove from project")
      AppendMenu hPopupMenu, MF_ENABLED, IDM_REMOVEFILEFROMPROJECT, wzText
      AddProjectFileTypesToMenu(hPopUpMenu, pDoc)
   End If
   
   Function = hPopupMenu

End Function
    







