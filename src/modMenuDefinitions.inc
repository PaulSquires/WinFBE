'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2025 Paul Squires, PlanetSquires Software
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

#include once "modMRU.bi"
#include once "frmKeyboard.bi"

' ========================================================================================
' Get the menu accelerator text (if one exists) for this menu item
' ========================================================================================
function getMenuAccelText( byval idMenu as long ) As CWSTR
   dim as CWSTR wszAccel
   ' Search the gKeys array for a match
   for i as long = lbound(gKeys) to ubound(gKeys)
      if gKeys(i).idAction = idMenu then
         if len(gKeys(i).wszUserKeys) then
            wszAccel = gKeys(i).wszUserKeys
         else
            if gKeys(i).bDefaultDisabled = false then
               wszAccel = gKeys(i).wszDefaultKeys
            end if   
         end if
         exit for   
      end if
   next
   
   if len(wszAccel) then wszAccel = chr(9) & wszAccel
   return wszAccel
end function

' ========================================================================================
' Get the menu text based on the incoming control id
' ========================================================================================
Function getMenuText( byval idMenu as long ) As CWSTR
   dim as CWSTR wszText
   select case idMenu
      case 0: return "-"    ' this is a separator line
      case IDC_MENUBAR_FILE: wszText = L(2,"&File")
      case IDC_MENUBAR_EDIT: wszText =  L(14,"Edit")
      case IDC_MENUBAR_SEARCH: wszText = L(42,"Search")
      case IDC_MENUBAR_VIEW: wszText = L(54,"View")
      case IDC_MENUBAR_PROJECT: wszText = L(63,"Project")
      case IDC_MENUBAR_COMPILE: wszText = L(64,"Compile")
      case IDC_MENUBAR_DESIGNER: wszText = L(309,"Designer")
      case IDC_MENUBAR_HELP: wszText = L(73,"Help")
      case IDM_FILENEW: wszText = L(3,"New")
      case IDM_FILEOPEN: wszText = L(4,"Open") & "..." 
      case IDM_FILEOPENTEMPLATES: wszText = L(197,"Open Templates") & "..." 
      case IDM_MRU: return L(10,"Open Recent")
      case IDM_MRUBASE to IDM_MRUBASE + 9: return gConfig.MRU( idMenu - IDM_MRUBASE )
      case IDM_MRUCLEAR: return L(225, "Clear this list")   
      case IDM_MRUFILES: return L(11,"(Empty)")
      case IDM_FILECLOSE: wszText = L(5,"Close") 
      case IDM_FILECLOSEALL: wszText = L(6,"Close All") 
      case IDM_FILESAVE: wszText = L(7,"Save")
      case IDM_FILESAVEAS: wszText = L(8,"Save As") & "..."
      case IDM_FILESAVEAS: wszText = L(8,"Save As") & "..."
      case IDM_FILESAVEALL: wszText = L(9,"Save All")
      case IDM_AUTOSAVE: wszText = L(427,"Auto Save") 
      case IDM_SAVESESSION: wszText = L(425,"Save Session") & "..."
      case IDM_LOADSESSION: wszText = L(426,"Load Session") & "..."
      case IDM_KEYBOARDSHORTCUTS: wszText = L(220,"Keyboard Shortcuts") & "..."
      case IDM_USERTOOLS: return L(289,"User Tools")
      case IDM_USERTOOLSDIALOG: wszText = L(289,"User Tools") & "..."
      case IDM_USERTOOLSBASE to IDM_USERTOOLSBASE + ubound(gConfig.Tools)
         dim as CWSTR wszText = gConfig.Tools(idMenu - IDM_USERTOOLSBASE).wszDescription
         dim as CWSTR wszShortcut = createToolsMenuShortcut(idMenu - IDM_USERTOOLSBASE)
         if len(wszShortcut) then wszText = wszText & chr(9) & wszShortcut
         return wszText
      case IDM_USERTOOLSLIST: return L(11,"(Empty)")
      case IDM_OPTIONS: return L(288,"Preferences")
      case IDM_OPTIONSDIALOG: wszText = L(62,"Environment Options") & "..."
      case IDM_BUILDCONFIG: wszText = L(277,"Build Configurations") & "..."
      case IDM_USERSNIPPETS: wszText = L(88,"User Snippets") & "..." 
      case IDM_CATEGORIES: wszText = L(436,"Explorer Categories") & "..." 
      case IDM_EXIT: wszText = L(13,"Exit") 

      case IDM_UNDO: wszText = L(15,"Undo")
      case IDM_REDO: wszText = L(16,"Redo")
      case IDM_CUT: wszText = L(17,"Cut") 
      case IDM_COPY: wszText = L(18,"Copy")
      case IDM_PASTE: wszText = L(19,"Paste")
      case IDM_DELETELINE: wszText = L(20,"Delete Line")
      case IDM_FIND: wszText = L(43,"Find") & "..."
      case IDM_FINDINFILES: wszText = L(257,"Find In Files") & "..."
      case IDM_REPLACE: wszText = L(46,"Replace") & "..."
      case IDM_COMMENTBLOCK: wszText = L(25,"Comment Block")
      case IDM_UNCOMMENTBLOCK: wszText = L(26,"UnComment Block")
      case IDM_DUPLICATELINE: wszText = L(28,"Duplicate Line")
      case IDM_MOVELINEUP: wszText = L(29,"Move Line Up")
      case IDM_MOVELINEDOWN: wszText = L(30,"Move Line Down")
      case IDM_SELECTLINE: wszText = L(41,"Select Line")
      case IDM_SELECTALL: wszText = L(40,"Select All")

      case IDM_DEFINITION: wszText = L(47,"Sub/Function Definition")
      case IDM_LASTPOSITION: wszText = L(48,"Last Position")
      case IDM_GOTONEXTFUNCTION: wszText = L(84,"Next Function")
      case IDM_GOTOPREVFUNCTION: wszText = L(85,"Previous Function")
      case IDM_GOTOHEADERFILE: wszText = L(231,"Goto Header File")
      case IDM_GOTOSOURCEFILE: wszText = L(232,"Goto Code File")
      case IDM_GOTOMAINFILE: wszText = L(233,"Goto Main File")
      case IDM_GOTORESOURCEFILE: wszText = L(234,"Goto Resource File")
      case IDM_GOTO: wszText = L(49,"Goto Line") & "..."
      case IDM_BOOKMARKTOGGLE: wszText = L(50,"Toggle Bookmark")
      case IDM_BOOKMARKNEXT: wszText = L(51,"Next Bookmark")
      case IDM_BOOKMARKPREV: wszText = L(52,"Previous Bookmark")
      case IDM_BOOKMARKCLEARALL: wszText = L(53,"Clear Bookmarks")

      case IDM_FOLDTOGGLE: wszText = L(55,"Toggle Current Fold Point")
      case IDM_FOLDBELOW: wszText = L(56,"Toggle Current And All Below")
      case IDM_FOLDALL: wszText = L(57,"Fold All")
      case IDM_UNFOLDALL: wszText = L(58,"Unfold All")
      case IDM_ZOOMIN: wszText = L(59,"Zoom In")
      case IDM_ZOOMOUT: wszText = L(60,"Zoom Out")
      case IDM_BOOKMARKSLIST: wszText = L(443,"View Bookmarks List")
      case IDM_FUNCTIONLIST: wszText = L(221,"View Function List")
      case IDM_VIEWEXPLORER: wszText = L(250,"View Explorer Window")
      case IDM_VIEWOUTPUT: wszText = L(251,"View Output Window")
      case IDM_VIEWTODO: wszText = L(263,"TODO")
      case IDM_VIEWNOTES: wszText = L(264,"Notes")
      case IDM_RESTOREMAIN: wszText = L(61,"Restore Main Window Size")

      case IDM_PROJECTNEW: wszText = L(180,"New Project")
      case IDM_PROJECTOPEN: wszText = L(182,"Open Project") & "..." 
      case IDM_MRUPROJECT: wszText = L(219,"Recent Projects")
      case IDM_MRUPROJECTBASE to IDM_MRUPROJECTBASE + 9
         return gConfig.MRUPROJECT( idMenu - IDM_MRUPROJECTBASE )
      case IDM_MRUPROJECTCLEAR: return L(225, "Clear this list")   
      case IDM_MRUPROJECTFILES: return L(11,"(Empty)")
      case IDM_PROJECTCLOSE: wszText = L(183,"Close Project")
      case IDM_PROJECTSAVE: wszText = L(184,"Save Project")
      case IDM_PROJECTSAVEAS: wszText = L(185,"Save Project As") & "..."
      case IDM_PROJECTFILESADD: wszText = L(207,"Add Files to Project") & "..."
      case IDM_PROJECTOPTIONS: wszText = L(186,"Project Options") & "..."      

      case IDM_BUILDEXECUTE: wszText = L(65,"Build And Execute")
      case IDM_COMPILE: wszText = L(64,"Compile")
      case IDM_REBUILDALL: wszText = L(268,"Rebuild All")
      case IDM_QUICKRUN: wszText = L(285,"Quick Run")
      case IDM_RUNEXE: wszText = L(67,"Run Executable")
      case IDM_COMMANDLINE: wszText = L(68,"Command Line") & "..."

      case IDM_NEWFORM: wszText = L(310,"New Form")
      case IDM_VIEWTOOLBOX: wszText = L(311,"View Toolbox")
      case IDM_TOGGLEVIEWCODE: wszText = L(382,"Toggle View Form/Code")
      case IDM_ALIGN: wszText = L(315,"Align")
      case IDM_ALIGNLEFTS: wszText = L(316,"Lefts")
      case IDM_ALIGNCENTERS: wszText = L(317,"Centers")
      case IDM_ALIGNRIGHTS: wszText = L(318,"Rights")
      case IDM_ALIGNTOPS: wszText = L(319,"Tops")
      case IDM_ALIGNMIDDLES: wszText = L(320,"Middles")
      case IDM_ALIGNBOTTOMS: wszText = L(321,"Bottoms")
      case IDM_MAKESAME: wszText = L(322,"Make Same Size")
      case IDM_SAMEWIDTHS: wszText = L(323,"Widths")
      case IDM_SAMEHEIGHTS: wszText = L(324,"Heights")
      case IDM_SAMEBOTH: wszText = L(325,"Both")
      case IDM_HORIZSPACING: wszText = L(265,"Horizontal Spacing")
      case IDM_HORIZEQUAL: wszText = L(329,"Make Equal")
      case IDM_HORIZINCREASE: wszText = L(330,"Increase")
      case IDM_HORIZDECREASE: wszText = L(331,"Decrease")
      case IDM_HORIZREMOVE: wszText = L(349,"Remove")
      case IDM_VERTSPACING: wszText = L(284,"Vertical Spacing")
      case IDM_VERTEQUAL: wszText = L(329,"Make Equal")
      case IDM_VERTINCREASE: wszText = L(330,"Increase")
      case IDM_VERTDECREASE: wszText = L(331,"Decrease")
      case IDM_VERTREMOVE: wszText = L(349,"Remove")
      case IDM_CENTER: wszText = L(332,"Center in Form")
      case IDM_CENTERHORIZ: wszText = L(333,"Horizontally")
      case IDM_CENTERVERT: wszText = L(334,"Vertically")
      case IDM_CENTERBOTH: wszText = L(325,"Both")
      case IDM_MENUEDITOR: wszText = L(312,"Menu Editor") & "..."
      case IDM_TOOLBAREDITOR: wszText = L(313,"Toolbar Editor") & "..."
      case IDM_STATUSBAREDITOR: wszText = L(314,"Statusbar Editor") & "..."
      case IDM_IMAGEMANAGER: wszText = L(368,"Image Manager") & "..."
      case IDM_SNAPLINES: wszText = L(365,"Enable SnapLines")
      case IDM_LOCKCONTROLS: wszText = L(335,"Lock Controls")

      case IDM_HELP: wszText = "FreeBASIC " & L(73,"Help")
      case IDM_HELPWINFBE: wszText = "WinFBE Editor " & L(73,"Help")
      case IDM_HELPWINFBX: wszText = "WinFBX Framework " & L(73,"Help")
      case IDM_CHECKFORUPDATES: wszText = L(91,"Check for Updates")
      case IDM_ABOUT: wszText = L(74,"About")

      case else
         return ""
   end select
   
   wszText = wszText & getMenuAccelText( idMenu )
   return wszText
end function


' ========================================================================================
' Set the values for one element in the gTopMenu TYPE array
' ========================================================================================
function setTopMenuItem( byval nParentID as long, _
                         byval nID as long, _
                         byval nChildID as long, _
                         byval isDisabled as boolean, _
                         byval isSeparator as boolean _
                         ) as long
                         
   redim preserve gTopMenu(ubound(gTopMenu)+1) as TOPMENU_TYPE
   dim as long i = ubound(gTopMenu)

   gTopMenu(i).nParentID = nParentID
   gTopMenu(i).nID = nID
   gTopMenu(i).nChildID = nChildID
   gTopMenu(i).isDisabled = isDisabled
   gTopMenu(i).isSeparator = isSeparator
   
   function = 0

end function

' ========================================================================================
' Set the values for one MRU element in the gTopMenu TYPE array
' ========================================================================================
function setTopMenuMRUItem( byval nIndex as long, _
                            byval nParentID as long, _
                            byval nID as long, _
                            byval nChildID as long, _
                            byval isDisabled as boolean, _
                            byval isSeparator as boolean _
                            ) as long
   ' do sanity check to ensure that nIndex is within range
   if (nIndex >= lbound(gTopMenu)) andalso (nIndex <= ubound(gTopMenu)) then
      gTopMenu(nIndex).nParentID = nParentID
      gTopMenu(nIndex).nID = nID
      gTopMenu(nIndex).nChildID = nChildID
      gTopMenu(nIndex).isDisabled = isDisabled
      gTopMenu(nIndex).isSeparator = isSeparator
   end if
   
   function = 0
end function


' ========================================================================================
' Reurns a pointer to the gTopMenu TYPE array for the incoming nID
' Used mostly during the painting of the line
' ========================================================================================
public function getTopMenuPtr( byval nID as long ) as TOPMENU_TYPE ptr
   static dummy as TOPMENU_TYPE    ' returned in case no match in order to avoid null pointer
   for i as long = lbound(gTopMenu) to ubound(gTopMenu)
      if gTopMenu(i).nID = nID then
         return @gTopMenu(i)
      end if
   next
   return @dummy
end function



function buildTopMenuDefinitions() as long
   setTopMenuItem( IDC_MENUBAR_FILE, IDM_FILENEW, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_FILE, IDM_FILEOPEN, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_FILE, IDM_FILEOPENTEMPLATES, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_FILE, IDM_MRU, IDM_MRUFILES, false, false )

   updateMRUFilesItems()
   
   setTopMenuItem( IDC_MENUBAR_FILE, 0, 0, false, true )
   setTopMenuItem( IDC_MENUBAR_FILE, IDM_FILECLOSE, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_FILE, IDM_FILECLOSEALL, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_FILE, 0, 0, false, true )
   setTopMenuItem( IDC_MENUBAR_FILE, IDM_FILESAVE, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_FILE, IDM_FILESAVEAS, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_FILE, IDM_FILESAVEALL, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_FILE, IDM_AUTOSAVE, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_FILE, 0, 0, false, true )
   setTopMenuItem( IDC_MENUBAR_FILE, IDM_LOADSESSION, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_FILE, IDM_SAVESESSION, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_FILE, 0, 0, false, true )
   setTopMenuItem( IDC_MENUBAR_FILE, IDM_KEYBOARDSHORTCUTS, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_FILE, IDM_USERTOOLS, IDM_USERTOOLSLIST, false, false )

   updateUserToolsMenuItems()

   setTopMenuItem( IDC_MENUBAR_FILE, IDM_OPTIONS, IDM_OPTIONSDIALOG, false, false )
   setTopMenuItem( IDM_OPTIONS, IDM_OPTIONSDIALOG, 0, false, false )
   setTopMenuItem( IDM_OPTIONS, IDM_BUILDCONFIG, 0, false, false )
   setTopMenuItem( IDM_OPTIONS, IDM_USERSNIPPETS, 0, false, false )
   setTopMenuItem( IDM_OPTIONS, IDM_CATEGORIES, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_FILE, 0, 0, false, true )
   setTopMenuItem( IDC_MENUBAR_FILE, IDM_EXIT, 0, false, false )
   
   setTopMenuItem( IDC_MENUBAR_EDIT, IDM_UNDO, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_EDIT, IDM_REDO, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_EDIT, 0, 0, false, true )
   setTopMenuItem( IDC_MENUBAR_EDIT, IDM_CUT, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_EDIT, IDM_COPY, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_EDIT, IDM_PASTE, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_EDIT, IDM_DELETELINE, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_EDIT, 0, 0, false, true )
   setTopMenuItem( IDC_MENUBAR_EDIT, IDM_FIND, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_EDIT, IDM_FINDINFILES, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_EDIT, IDM_REPLACE, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_EDIT, 0, 0, false, true )
   setTopMenuItem( IDC_MENUBAR_EDIT, IDM_DUPLICATELINE, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_EDIT, IDM_MOVELINEUP, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_EDIT, IDM_MOVELINEDOWN, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_EDIT, 0, 0, false, true )
   setTopMenuItem( IDC_MENUBAR_EDIT, IDM_COMMENTBLOCK, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_EDIT, IDM_UNCOMMENTBLOCK, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_EDIT, 0, 0, false, true )
   setTopMenuItem( IDC_MENUBAR_EDIT, IDM_SELECTLINE, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_EDIT, IDM_SELECTALL, 0, false, false )
   
   setTopMenuItem( IDC_MENUBAR_SEARCH, IDM_DEFINITION, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_SEARCH, IDM_LASTPOSITION, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_SEARCH, 0, 0, false, true )
   setTopMenuItem( IDC_MENUBAR_SEARCH, IDM_GOTONEXTFUNCTION, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_SEARCH, IDM_GOTOPREVFUNCTION, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_SEARCH, 0, 0, false, true )
   setTopMenuItem( IDC_MENUBAR_SEARCH, IDM_GOTOHEADERFILE, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_SEARCH, IDM_GOTOSOURCEFILE, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_SEARCH, IDM_GOTOMAINFILE, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_SEARCH, IDM_GOTORESOURCEFILE, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_SEARCH, 0, 0, false, true )
   setTopMenuItem( IDC_MENUBAR_SEARCH, IDM_BOOKMARKTOGGLE, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_SEARCH, IDM_BOOKMARKNEXT, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_SEARCH, IDM_BOOKMARKPREV, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_SEARCH, IDM_BOOKMARKCLEARALL, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_SEARCH, 0, 0, false, true )
   setTopMenuItem( IDC_MENUBAR_SEARCH, IDM_GOTO, 0, false, false )

   setTopMenuItem( IDC_MENUBAR_VIEW, IDM_VIEWEXPLORER, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_VIEW, IDM_VIEWOUTPUT, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_VIEW, IDM_FUNCTIONLIST, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_VIEW, IDM_BOOKMARKSLIST, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_VIEW, 0, 0, false, true )
   setTopMenuItem( IDC_MENUBAR_VIEW, IDM_ZOOMIN, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_VIEW, IDM_ZOOMOUT, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_VIEW, 0, 0, false, true )
   setTopMenuItem( IDC_MENUBAR_VIEW, IDM_FOLDTOGGLE, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_VIEW, IDM_FOLDBELOW, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_VIEW, IDM_FOLDALL, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_VIEW, IDM_UNFOLDALL, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_VIEW, 0, 0, false, true )
   setTopMenuItem( IDC_MENUBAR_VIEW, IDM_RESTOREMAIN, 0, false, false )

   setTopMenuItem( IDC_MENUBAR_PROJECT, IDM_PROJECTNEW, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_PROJECT, IDM_PROJECTOPEN, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_PROJECT, IDM_MRUPROJECT, IDM_MRUPROJECTFILES, false, false )
   
   updateMRUProjectFilesItems()

   setTopMenuItem( IDC_MENUBAR_PROJECT, 0, 0, false, true )
   setTopMenuItem( IDC_MENUBAR_PROJECT, IDM_PROJECTCLOSE, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_PROJECT, 0, 0, false, true )
   setTopMenuItem( IDC_MENUBAR_PROJECT, IDM_PROJECTSAVEAS, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_PROJECT, IDM_PROJECTSAVE, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_PROJECT, 0, 0, false, true )
   setTopMenuItem( IDC_MENUBAR_PROJECT, IDM_PROJECTFILESADD, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_PROJECT, 0, 0, false, true )
   setTopMenuItem( IDC_MENUBAR_PROJECT, IDM_PROJECTOPTIONS, 0, false, false )

   setTopMenuItem( IDC_MENUBAR_COMPILE, IDM_BUILDEXECUTE, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_COMPILE, IDM_COMPILE, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_COMPILE, IDM_REBUILDALL, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_COMPILE, 0, 0, false, true )
   setTopMenuItem( IDC_MENUBAR_COMPILE, IDM_QUICKRUN, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_COMPILE, IDM_RUNEXE, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_COMPILE, 0, 0, false, true )
   setTopMenuItem( IDC_MENUBAR_COMPILE, IDM_COMMANDLINE, 0, false, false )

   setTopMenuItem( IDC_MENUBAR_DESIGNER, IDM_NEWFORM, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_DESIGNER, IDM_VIEWTOOLBOX, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_DESIGNER, 0, 0, false, true )
   setTopMenuItem( IDC_MENUBAR_DESIGNER, IDM_TOGGLEVIEWCODE, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_DESIGNER, 0, 0, false, true )
   setTopMenuItem( IDC_MENUBAR_DESIGNER, IDM_ALIGN, IDM_ALIGNLEFTS, false, false )
   setTopMenuItem( IDM_ALIGN, IDM_ALIGNLEFTS, 0, false, false )
   setTopMenuItem( IDM_ALIGN, IDM_ALIGNCENTERS, 0, false, false )
   setTopMenuItem( IDM_ALIGN, IDM_ALIGNRIGHTS, 0, false, false )
   setTopMenuItem( IDM_ALIGN, 0, 0, false, true )
   setTopMenuItem( IDM_ALIGN, IDM_ALIGNTOPS, 0, false, false )
   setTopMenuItem( IDM_ALIGN, IDM_ALIGNMIDDLES, 0, false, false )
   setTopMenuItem( IDM_ALIGN, IDM_ALIGNBOTTOMS, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_DESIGNER, IDM_MAKESAME, IDM_SAMEWIDTHS, false, false )
   setTopMenuItem( IDM_MAKESAME, IDM_SAMEWIDTHS, 0, false, false )
   setTopMenuItem( IDM_MAKESAME, IDM_SAMEHEIGHTS, 0, false, false )
   setTopMenuItem( IDM_MAKESAME, 0, 0, false, true )
   setTopMenuItem( IDM_MAKESAME, IDM_SAMEBOTH, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_DESIGNER, 0, 0, false, true )
   setTopMenuItem( IDC_MENUBAR_DESIGNER, IDM_HORIZSPACING, IDM_HORIZEQUAL, false, false )
   setTopMenuItem( IDM_HORIZSPACING, IDM_HORIZEQUAL, 0, false, false )
   setTopMenuItem( IDM_HORIZSPACING, 0, 0, false, true )
   setTopMenuItem( IDM_HORIZSPACING, IDM_HORIZINCREASE, 0, false, false )
   setTopMenuItem( IDM_HORIZSPACING, IDM_HORIZDECREASE, 0, false, false )
   setTopMenuItem( IDM_HORIZSPACING, 0, 0, false, true )
   setTopMenuItem( IDM_HORIZSPACING, IDM_HORIZREMOVE, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_DESIGNER, IDM_VERTSPACING, IDM_VERTEQUAL, false, false )
   setTopMenuItem( IDM_VERTSPACING, IDM_VERTEQUAL, 0, false, false )
   setTopMenuItem( IDM_VERTSPACING, 0, 0, false, true )
   setTopMenuItem( IDM_VERTSPACING, IDM_VERTINCREASE, 0, false, false )
   setTopMenuItem( IDM_VERTSPACING, IDM_VERTDECREASE, 0, false, false )
   setTopMenuItem( IDM_VERTSPACING, 0, 0, false, true )
   setTopMenuItem( IDM_VERTSPACING, IDM_VERTREMOVE, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_DESIGNER, 0, 0, false, true )
   setTopMenuItem( IDC_MENUBAR_DESIGNER, IDM_CENTER, IDM_CENTERHORIZ, false, false )
   setTopMenuItem( IDM_CENTER, IDM_CENTERHORIZ, 0, false, false )
   setTopMenuItem( IDM_CENTER, IDM_CENTERVERT, 0, false, false )
   setTopMenuItem( IDM_CENTER, 0, 0, false, true )
   setTopMenuItem( IDM_CENTER, IDM_CENTERBOTH, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_DESIGNER, 0, 0, false, true )
   setTopMenuItem( IDC_MENUBAR_DESIGNER, IDM_MENUEDITOR, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_DESIGNER, IDM_TOOLBAREDITOR, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_DESIGNER, IDM_STATUSBAREDITOR, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_DESIGNER, IDM_IMAGEMANAGER, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_DESIGNER, 0, 0, false, true )
   setTopMenuItem( IDC_MENUBAR_DESIGNER, IDM_SNAPLINES, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_DESIGNER, IDM_LOCKCONTROLS, 0, false, false )

   setTopMenuItem( IDC_MENUBAR_HELP, IDM_HELP, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_HELP, 0, 0, false, true )
   setTopMenuItem( IDC_MENUBAR_HELP, IDM_HELPWINFBE, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_HELP, IDM_HELPWINFBX, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_HELP, 0, 0, false, true )
   setTopMenuItem( IDC_MENUBAR_HELP, IDM_CHECKFORUPDATES, 0, false, false )
   setTopMenuItem( IDC_MENUBAR_HELP, 0, 0, false, true )
   setTopMenuItem( IDC_MENUBAR_HELP, IDM_ABOUT, 0, false, false )
   
   function = 0
end function
