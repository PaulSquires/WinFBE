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

#include once "clsConfig.bi"
#include once "frmKeyboard.bi"


' ========================================================================================
' Save the gKeys() keybinding array to disk file
' ========================================================================================
function frmKeyboard_SaveKeyBindings( byval wszFilename as CWSTR ) as long
   dim pStream AS CTextStream
   if pStream.Create( wszFilename ) = S_OK then 
      dim as CWSTR wszLine
      for i as long = lbound(gKeys) to ubound(gKeys)
         wszLine = gKeys(i).wszMsgString 
         
         if gKeys(i).bDefaultDisabled then
            wszLine = wszLine & "(DISABLED)"
         end if
            
         wszLine = wszLine & ":" & gKeys(i).wszUserKeys 
         pStream.WriteLine wszLine
      next
      pStream.Close
   end if
   function = 0
end function


' ========================================================================================
' Load a keybinding into the global array
' ========================================================================================
function frmKeyBoard_AddKeyBinding( _
            byval wszCategory as CWSTR, _
            byval idAction as long, _
            byval wszMsgString as CWSTR, _
            byval wszDescription as CWSTR, _
            byval wszDefaultKeys as CWSTR, _
            byval wszUserKeys as CWSTR, _
            byval bDisabled as boolean _
            ) as long
            
   ' Search the array to see if the id exists. If it does then update
   ' that binding, otherwise, add it to the array.
   dim as boolean bFound = false
   dim as long nFoundIdx = -1
    
   for i as long = lbound(gKeys) to ubound(gKeys)
      if gKeys(i).wszMsgString = wszMsgString then
         bFound = true
         nFoundIdx = i
         exit for
      end if
   next
   
   if bFound = false then
      nFoundIdx = ubound(gKeys) + 1
      redim preserve gKeys(nFoundIdx) as KEYBINDINGS_TYPE
   end if   
   
   with gKeys(nFoundIdx)
      if bFound = false then   ' don't update existing system entries
         .idAction = idAction
         .wszMsgString = wszMsgString
         .wszCategory = wszCategory
         .wszDescription = wszDescription
         .wszDefaultKeys = wszDefaultKeys
      else
         ' Existing entry found which means we are reading the existing keybindings
         ' file. Update the entry with any user defined keybinding or if the user has
         ' disabled the Default keyboard shortcut.
         .wszUserKeys = wszUserKeys
         .bDefaultDisabled = bDisabled
      end if
   end with
   
   function = 0
end function


' ========================================================================================
' ' Create the default keyboard bindings / shortcut mappings
' ========================================================================================
function frmKeyboard_CreateDefaultKeyBindings() as long
   '' FILE MENU
   frmKeyBoard_AddKeyBinding( "File", IDM_FILENEW, "IDM_FILENEW", "New file", "Ctrl+N", "", false )
   frmKeyBoard_AddKeyBinding( "File", IDM_FILEOPEN, "IDM_FILEOPEN", "Open one or more files", "Ctrl+O", "", false )
   frmKeyBoard_AddKeyBinding( "File", IDM_FILECLOSE, "IDM_FILECLOSE", "Close file", "Ctrl+W", "", false )
   frmKeyBoard_AddKeyBinding( "File", IDM_FILECLOSEALL, "IDM_FILECLOSEALL", "Close all files", "Ctrl+Shift+W", "", false )
   frmKeyBoard_AddKeyBinding( "File", IDM_FILESAVE, "IDM_FILESAVE", "Save file", "Ctrl+S", "", false )
   frmKeyBoard_AddKeyBinding( "File", IDM_FILESAVEAS, "IDM_FILESAVEAS", "Save file as a different name", "", "", false )
   frmKeyBoard_AddKeyBinding( "File", IDM_FILESAVEALL, "IDM_FILESAVEALL", "Save all files", "Ctrl+Shift+S", "", false )
   frmKeyBoard_AddKeyBinding( "File", IDM_AUTOSAVE, "IDM_AUTOSAVE", "Toggle Auto Save", "", "", false )
   frmKeyBoard_AddKeyBinding( "File", IDM_LOADSESSION, "IDM_LOADSESSION", "Load session", "", "", false )
   frmKeyBoard_AddKeyBinding( "File", IDM_SAVESESSION, "IDM_SAVESESSION", "Save session", "", "", false )
   frmKeyBoard_AddKeyBinding( "File", IDM_KEYBOARDSHORTCUTS, "IDM_KEYBOARDSHORTCUTS", "Keyboard shortcuts", "Ctrl+K", "", false )
   frmKeyBoard_AddKeyBinding( "File", IDM_OPTIONSDIALOG, "IDM_OPTIONSDIALOG", "Environment options", "Shift+F7", "", false )
   frmKeyBoard_AddKeyBinding( "File", IDM_BUILDCONFIG, "IDM_BUILDCONFIG", "Build configurations", "F7", "", false )
   frmKeyBoard_AddKeyBinding( "File", IDM_EXIT, "IDM_EXIT", "Exit application", "Alt+F4", "", false )

   '' EDIT MENU
   frmKeyBoard_AddKeyBinding( "Edit", IDM_UNDO, "IDM_UNDO", "Undo", "Ctrl+Z", "", false )
   frmKeyBoard_AddKeyBinding( "Edit", IDM_REDO, "IDM_REDO", "Redo", "Ctrl+Shift+Z", "", false )
   frmKeyBoard_AddKeyBinding( "Edit", IDM_CUT, "IDM_CUT", "Cut selection to the clipboard", "Ctrl+X", "", false )
   frmKeyBoard_AddKeyBinding( "Edit", IDM_COPY, "IDM_COPY", "Copy selection to the clipboard", "Ctrl+C", "", false )
   frmKeyBoard_AddKeyBinding( "Edit", IDM_PASTE, "IDM_PASTE", "Paste clipboard contents", "Ctrl+V", "", false )
   frmKeyBoard_AddKeyBinding( "Edit", IDM_DELETELINE, "IDM_DELETELINE", "Delete the current editor line", "Ctrl+Y", "", false )
   frmKeyBoard_AddKeyBinding( "Edit", IDM_FIND, "IDM_FIND", "Find", "Ctrl+F", "", false )
   frmKeyBoard_AddKeyBinding( "Edit", IDM_FINDINFILES, "IDM_FINDINFILES", "Find in files", "Ctrl+Shift+F", "", false )
   frmKeyBoard_AddKeyBinding( "Edit", IDM_REPLACE, "IDM_REPLACE", "Replace", "Ctrl+H", "", false )
   frmKeyBoard_AddKeyBinding( "Edit", IDM_DUPLICATELINE, "IDM_DUPLICATELINE", "Duplicate line", "Ctrl+D", "", false )
   frmKeyBoard_AddKeyBinding( "Edit", IDM_MOVELINEUP, "IDM_MOVELINEUP", "Move line up", "Alt+Up", "", false )
   frmKeyBoard_AddKeyBinding( "Edit", IDM_MOVELINEDOWN, "IDM_MOVELINEDOWN", "Move line down", "Alt+Down", "", false )
   frmKeyBoard_AddKeyBinding( "Edit", IDM_COMMENTBLOCK, "IDM_COMMENTBLOCK", "Comment block", "Ctrl+/", "", false )   'VK_OEM_2
   frmKeyBoard_AddKeyBinding( "Edit", IDM_UNCOMMENTBLOCK, "IDM_UNCOMMENTBLOCK", "UnComment block", "Ctrl+Shift+/", "", false )  ' VK_OEM_2
   frmKeyBoard_AddKeyBinding( "Edit", IDM_SELECTLINE, "IDM_SELECTLINE", "Select line", "Ctrl+L", "", false )
   frmKeyBoard_AddKeyBinding( "Edit", IDM_SELECTALL, "IDM_SELECTALL", "Select all", "Ctrl+A", "", false )
   ' The following are non-visual (no topmenu item)
   frmKeyBoard_AddKeyBinding( "Edit", IDM_FINDNEXTACCEL, "IDM_FINDNEXTACCEL", "Find next", "F3", "", false )
   frmKeyBoard_AddKeyBinding( "Edit", IDM_FINDPREVACCEL, "IDM_FINDPREVACCEL", "Find previous", "Shift+F3", "", false )
   frmKeyBoard_AddKeyBinding( "Edit", IDM_INDENTBLOCK, "IDM_INDENTBLOCK", "Indent block", "TAB", "", false )
   frmKeyBoard_AddKeyBinding( "Edit", IDM_UNINDENTBLOCK, "IDM_UNINDENTBLOCK", "UnIndent block", "Shift+TAB", "", false )
   frmKeyBoard_AddKeyBinding( "Edit", IDM_INSERTFILE, "IDM_INSERTFILE", "Insert file", "Ctrl+I", "", false )
   frmKeyBoard_AddKeyBinding( "Edit", IDM_TOUPPERCASE, "IDM_TOUPPERCASE", "Change selection to uppercase", "Ctrl+Alt+U", "", false )
   frmKeyBoard_AddKeyBinding( "Edit", IDM_TOLOWERCASE, "IDM_TOLOWERCASE", "Change selection to lowercase", "Ctrl+Alt+L", "", false )
   frmKeyBoard_AddKeyBinding( "Edit", IDM_TOMIXEDCASE, "IDM_TOMIXEDCASE", "Change selection to mixedcase", "Ctrl+Alt+X", "", false )
   frmKeyBoard_AddKeyBinding( "Edit", IDM_NEWLINEBELOWCURRENT, "IDM_NEWLINEBELOWCURRENT", "Insert new line below current line", "Ctrl+Enter", "", false ) 
   frmKeyBoard_AddKeyBinding( "Edit", IDM_SETFOCUSEDITOR, "IDM_SETFOCUSEDITOR", "Set keyboard focus to the editing window", "Ctrl+Tilde", "", false ) ' VK_OEM_3
   
   '' SEARCH MENU
   frmKeyBoard_AddKeyBinding( "Search", IDM_DEFINITION, "IDM_DEFINITION", "Goto the sub/function definition", "F12", "", false )
   frmKeyBoard_AddKeyBinding( "Search", IDM_LASTPOSITION, "IDM_LASTPOSITION", "Goto previous file position", "Shift+F12", "", false )
   frmKeyBoard_AddKeyBinding( "Search", IDM_GOTONEXTFUNCTION, "IDM_GOTONEXTFUNCTION", "Goto next sub/function", "Ctrl+PgDn", "", false )
   frmKeyBoard_AddKeyBinding( "Search", IDM_GOTOPREVFUNCTION, "IDM_GOTOPREVFUNCTION", "Goto previous sub/function", "Ctrl+PgUp", "", false )
   frmKeyBoard_AddKeyBinding( "Search", IDM_GOTOHEADERFILE, "IDM_GOTOHEADERFILE", "Goto header file", "Ctrl+Shift+H", "", false )
   frmKeyBoard_AddKeyBinding( "Search", IDM_GOTOSOURCEFILE, "IDM_GOTOSOURCEFILE", "Goto code file", "Ctrl+Shift+C", "", false )
   frmKeyBoard_AddKeyBinding( "Search", IDM_GOTOMAINFILE, "IDM_GOTOMAINFILE", "Goto main file", "Ctrl+Shift+M", "", false )
   frmKeyBoard_AddKeyBinding( "Search", IDM_GOTORESOURCEFILE, "IDM_GOTORESOURCEFILE", "Goto resource file", "Ctrl+Shift+R", "", false )
   frmKeyBoard_AddKeyBinding( "Search", IDM_BOOKMARKTOGGLE, "IDM_BOOKMARKTOGGLE", "Toggle bookmark", "Ctrl+F2", "", false )
   frmKeyBoard_AddKeyBinding( "Search", IDM_BOOKMARKNEXT, "IDM_BOOKMARKNEXT", "Goto the next bookmark", "F2", "", false )
   frmKeyBoard_AddKeyBinding( "Search", IDM_BOOKMARKPREV, "IDM_BOOKMARKPREV", "Goto the previous bookmark", "Shift+F2", "", false )
   frmKeyBoard_AddKeyBinding( "Search", IDM_BOOKMARKCLEARALL, "IDM_BOOKMARKCLEARALL", "Clear bookmarks", "Ctrl+Shift+F2", "", false )
   frmKeyBoard_AddKeyBinding( "Search", IDM_GOTO, "IDM_GOTO", "Goto line", "Ctrl+G", "", false )
   ' The following are non-visual (no topmenu item)
   frmKeyBoard_AddKeyBinding( "Search", IDM_GOTONEXTTAB, "IDM_GOTONEXTTAB", "Goto next open editor tab", "Ctrl+Tab", "", false )
   frmKeyBoard_AddKeyBinding( "Search", IDM_GOTOPREVTAB, "IDM_GOTOPREVTAB", "Goto previous open editor tab", "Ctrl+Shift+Tab", "", false )
   frmKeyBoard_AddKeyBinding( "Search", IDM_GOTONEXTCOMPILEERROR, "IDM_GOTONEXTCOMPILEERROR", "Goto next compile error", "Ctrl+\", "", false )
   frmKeyBoard_AddKeyBinding( "Search", IDM_GOTOPREVCOMPILEERROR, "IDM_GOTOPREVCOMPILEERROR", "Goto previous compile error", "Ctrl+Shift+\", "", false)

   '' VIEW MENU
   frmKeyBoard_AddKeyBinding( "View", IDM_VIEWEXPLORER, "IDM_VIEWEXPLORER", "View Explorer window", "F9", "", false )
   frmKeyBoard_AddKeyBinding( "View", IDM_VIEWOUTPUT, "IDM_VIEWOUTPUT", "View Output window", "Ctrl+F9", "", false )
   frmKeyBoard_AddKeyBinding( "View", IDM_FUNCTIONLIST, "IDM_FUNCTIONLIST", "View Function List", "F4", "", false )
   frmKeyBoard_AddKeyBinding( "View", IDM_BOOKMARKSLIST, "IDM_BOOKMARKSLIST", "View Bookmarks List", "Shift+F4", "", false )
   frmKeyBoard_AddKeyBinding( "View", IDM_ZOOMIN, "IDM_ZOOMIN", "Zoom in", "Ctrl+Plus", "", false )     ' VK_OEM_PLUS
   frmKeyBoard_AddKeyBinding( "View", IDM_ZOOMOUT, "IDM_ZOOMOUT", "Zoom out", "Ctrl+Minus", "", false )   ' VK_OEM_MINUS
   frmKeyBoard_AddKeyBinding( "View", IDM_FOLDTOGGLE, "IDM_FOLDTOGGLE", "Toggle current fold point", "F8", "", false )
   frmKeyBoard_AddKeyBinding( "View", IDM_FOLDBELOW, "IDM_FOLDBELOW", "Fold current line and all below", "Ctrl+F8", "", false )
   frmKeyBoard_AddKeyBinding( "View", IDM_FOLDALL, "IDM_FOLDALL", "Fold all lines", "Shift+F8", "", false )
   frmKeyBoard_AddKeyBinding( "View", IDM_UNFOLDALL, "IDM_UNFOLDALL", "Unfold all lines", "Ctrl+Shift+F8", "", false )
   ' The following are non-visual (no topmenu item)
   frmKeyBoard_AddKeyBinding( "View", IDM_VIEWNOTES, "IDM_VIEWNOTES", "Notes for this file or project", "", "", false )
   frmKeyBoard_AddKeyBinding( "View", IDM_VIEWTODO, "IDM_VIEWTODO", "List of TODO's for this file or project", "", "", false )
   frmKeyBoard_AddKeyBinding( "View", IDM_EXPLORER_EXPANDALL, "IDM_EXPLORER_EXPANDALL", "Expand all nodes in the Explorer window", "", "", false )
   frmKeyBoard_AddKeyBinding( "View", IDM_EXPLORER_COLLAPSEALL, "IDM_EXPLORER_COLLAPSEALL", "Collapse all nodes in the Explorer window", "", "", false )
   frmKeyBoard_AddKeyBinding( "View", IDM_FUNCTIONS_EXPANDALL, "IDM_FUNCTIONS_EXPANDALL", "Expand all nodes in the Function List window", "", "", false )
   frmKeyBoard_AddKeyBinding( "View", IDM_FUNCTIONS_COLLAPSEALL, "IDM_FUNCTIONS_COLLAPSEALL", "Collapse all nodes in the Function List window", "", "", false )
   frmKeyBoard_AddKeyBinding( "View", IDM_BOOKMARKS_EXPANDALL, "IDM_BOOKMARKS_EXPANDALL", "Expand all nodes in the Bookmarks window", "", "", false )
   frmKeyBoard_AddKeyBinding( "View", IDM_BOOKMARKS_COLLAPSEALL, "IDM_BOOKMARKS_COLLAPSEALL", "Collapse all nodes in the Bookmarks window", "", "", false )
   frmKeyBoard_AddKeyBinding( "View", IDM_CLOSEPANEL, "IDM_CLOSEPANEL", "Close left side panel", "", "", false )

   '' PROJECT MENU
   frmKeyBoard_AddKeyBinding( "Project", IDM_PROJECTNEW, "IDM_PROJECTNEW", "New project", "", "", false )
   frmKeyBoard_AddKeyBinding( "Project", IDM_PROJECTSAVE, "IDM_PROJECTSAVE", "Save project", "", "", false )
   frmKeyBoard_AddKeyBinding( "Project", IDM_PROJECTSAVEAS, "IDM_PROJECTSAVEAS", "Save project as a different name ", "", "", false )
   frmKeyBoard_AddKeyBinding( "Project", IDM_PROJECTOPEN, "IDM_PROJECTOPEN", "Open project", "F11", "", false )
   frmKeyBoard_AddKeyBinding( "Project", IDM_PROJECTFILESADD, "IDM_PROJECTFILESADD", "Add files to project", "Ctrl+F11", "", false )
   frmKeyBoard_AddKeyBinding( "Project", IDM_PROJECTCLOSE, "IDM_PROJECTCLOSE", "Close project", "", "", false )
   frmKeyBoard_AddKeyBinding( "Project", IDM_PROJECTOPTIONS, "IDM_PROJECTOPTIONS", "Project options", "", "", false )

   '' COMPILE MENU
   frmKeyBoard_AddKeyBinding( "Compile", IDM_BUILDEXECUTE, "IDM_BUILDEXECUTE", "Build and Execute", "F5", "", false )
   frmKeyBoard_AddKeyBinding( "Compile", IDM_COMPILE, "IDM_COMPILE", "Compile the project or the current active file", "Ctrl+F5", "", false )
   frmKeyBoard_AddKeyBinding( "Compile", IDM_REBUILDALL, "IDM_REBUILDALL", "Rebuild all", "Ctrl+Alt+F5", "", false )
   frmKeyBoard_AddKeyBinding( "Compile", IDM_QUICKRUN, "IDM_QUICKRUN", "Quick Run current active file", "Ctrl+Shift+F5", "", false )
   frmKeyBoard_AddKeyBinding( "Compile", IDM_RUNEXE, "IDM_RUNEXE", "Run executable", "Shift+F5", "", false )
   frmKeyBoard_AddKeyBinding( "Compile", IDM_COMMANDLINE, "IDM_COMMANDLINE", "Command line", "", "", false )
   
   '' HELP MENU
   frmKeyBoard_AddKeyBinding( "Help", IDM_HELP, "IDM_HELP", "FreeBasic Help (context sensitive help)", "F1", "", false )
   
   ' Load the keybindings settings file. This will update the gKeys array with any
   ' user defined key mappings.
   dim as CWSTR wszFilename = AfxGetExePathName & "settings\keybindings.ini" 
   
   dim pStream AS CTextStream
   
   if pStream.Open( wszFilename ) = S_OK then 
      do until pStream.EOS
         dim as CWSTR wst = pStream.ReadLine
         wst = trim(wst)
         if len(wst) = 0 then Continue Do
         
         ' Each keybinding entry has fields separated by a colon 
         if instr( wst, ":" ) = 0 then continue do
         
         dim as CWSTR wszMsgString = AfxStrParse(wst, 1, ":")  
         
         ' The MsgString could have an embedded (DISABLED) flag indicating that the
         ' default key combination should be disabled.
         dim as boolean bDisabled = false
         dim as long i
         i = instr(wszMsgString, "(DISABLED)")
         if i then
            bDisabled = true
            wszMsgString = left(wszMsgString, i-1 )
         end if   
         
         dim as CWSTR wszUserKeys = trim(AfxStrParse(wst, 2, ":"))
         frmKeyBoard_AddKeyBinding( "", 0, wszMsgString, "", "", wszUserKeys, bDisabled )
      Loop
      pStream.Close
   end if

   ' Lastly, save the keybindings settings file (we do this in case the file never existed
   ' in the first place so now we will have a default keybindings file).
   frmKeyboard_SaveKeyBindings( wszFilename )
   
   function = 0   
end function

' ========================================================================================
' Convert a virtual key string (Ctrl, Alt, Shift) to virtual key value
' ========================================================================================
function frmKeyboard_VirtKeyToValue( byval wszString as CWSTR ) as long
   select case ucase(wszString)
      case "CTRL":  return FCONTROL
      case "ALT":   return FALT
      case "SHIFT": return FSHIFT
   end select
   function = 0
end function


' ========================================================================================
' Converts a VK_OEM key to propery virtual key because of different keyboard layouts
' ========================================================================================
function frmKeyboard_OEMtoVirtKey( byval ch as ubyte ) as long
   ' declare function VkKeyScanExW(byval ch as WCHAR, byval dwhkl as HKL) as SHORT
   dim as HKL layout = GetKeyboardLayout(0)  '// Current layout
   dim as short result = VkKeyScanEx(ch, layout)
   dim as ubyte vkCode = lobyte(result)   
   dim as ubyte shiftState = hibyte(result)
   return vkCode
end function


' ========================================================================================
' Convert an accelerator key string (F1, TAB, "A", etc..) to virtual key value
' ========================================================================================
function frmKeyboard_AccelKeyToValue( byval wszString as CWSTR ) as long
   select case ucase(wszString)
      case "BACKSPACE": return VK_BACK
      case "TAB": return VK_TAB
      case "ENTER": return VK_RETURN
      case "ESCAPE": return VK_ESCAPE
      case "SPACEBAR": return VK_SPACE
      case "PGUP": return VK_PRIOR
      case "PGDN": return VK_NEXT
      case "END": return VK_END
      case "HOME": return VK_HOME
      case "LEFT": return VK_LEFT
      case "UP": return VK_UP
      case "RIGHT": return VK_RIGHT
      case "DOWN": return VK_DOWN
      case "INS": return VK_INSERT
      case "DEL": return VK_DELETE
      case "NUMPAD0": return VK_NUMPAD0
      case "NUMPAD1": return VK_NUMPAD1
      case "NUMPAD2": return VK_NUMPAD2
      case "NUMPAD3": return VK_NUMPAD3
      case "NUMPAD4": return VK_NUMPAD4
      case "NUMPAD5": return VK_NUMPAD5
      case "NUMPAD6": return VK_NUMPAD6
      case "NUMPAD7": return VK_NUMPAD7
      case "NUMPAD8": return VK_NUMPAD8
      case "NUMPAD9": return VK_NUMPAD9
      case "NUMPAD*": return VK_MULTIPLY
      case "NUMPAD+": return VK_ADD
      case "NUMPAD-": return VK_SUBTRACT
      case "NUMPAD.": return VK_DECIMAL
      case "NUMPAD/": return VK_DIVIDE
      case "F1": return VK_F1
      case "F2": return VK_F2
      case "F3": return VK_F3
      case "F4": return VK_F4
      case "F5": return VK_F5
      case "F6": return VK_F6
      case "F7": return VK_F7
      case "F8": return VK_F8
      case "F9": return VK_F9
      case "F10": return VK_F10
      case "F11": return VK_F11
      case "F12": return VK_F12
      case "+", "PLUS": return VK_OEM_PLUS           
      case ",", "COMMA": return VK_OEM_COMMA           
      case "-", "MINUS": return VK_OEM_MINUS          
      case ".", "PERIOD": return VK_OEM_PERIOD
      case ";", "SEMICOLON"       ' VK_OEM_1
         return frmKeyboard_OEMtoVirtKey(asc(";"))
      case ":", "COLON"           ' VK_OEM_1
         return frmKeyboard_OEMtoVirtKey(asc(":"))
      case "/", "FORWARDSLASH"    ' VK_OEM_2
         return frmKeyboard_OEMtoVirtKey(asc("/"))
      case "?", "QUESTIONMARK"    ' VK_OEM_2
         return frmKeyboard_OEMtoVirtKey(asc("?"))
      case "`", "BACKTICK"        ' VK_OEM_3
         return frmKeyboard_OEMtoVirtKey(asc("`"))
      case "~", "TILDE"           ' VK_OEM_3
         return frmKeyboard_OEMtoVirtKey(asc("~"))
      case "[", "OPENSQUAREBRACKET" ' VK_OEM_4
         return frmKeyboard_OEMtoVirtKey(asc("["))
      case "{", "OPENCURLYBRACE"    ' VK_OEM_4
         return frmKeyboard_OEMtoVirtKey(asc("{"))
      case "\", "BACKSLASH"         ' VK_OEM_5
         return frmKeyboard_OEMtoVirtKey(asc("\"))
      case "|", "PIPE"              ' VK_OEM_5
         return frmKeyboard_OEMtoVirtKey(asc("|"))
      case "]", "CLOSESQUAREBRACKET"  ' VK_OEM_6
         return frmKeyboard_OEMtoVirtKey(asc("]"))
      case "}", "CLOSECURLYBRACE"     ' VK_OEM_6
         return frmKeyboard_OEMtoVirtKey(asc("}"))
      case "'", "SINGLEQUOTE"         ' VK_OEM_7
         return frmKeyboard_OEMtoVirtKey(asc("'"))
      case chr(34), "DOUBLEQUOTE"     ' VK_OEM_7
         return frmKeyboard_OEMtoVirtKey(34)
      case "0" to "9": return asc(wszString)   
      case "A" to "Z": return asc(wszString)
   end select
   function = 0
end function


' ========================================================================================
' Build the main keyboard accelerator table
' ========================================================================================
Function frmKeyboard_BuildAcceleratorTable() As long
    Dim pWindow As CWindow Ptr = AfxCWindowPtr(HWND_FRMMAIN)
    if pWindow = 0 then exit function

   ' Create mapping array that equates the string representation of the
   ' IDM_* messages to their values. We use that array we saving and 
   ' restoring the keybindings from file. We can not use the literal IDM
   ' value because if we modify any of the IDM source code values then the
   ' mappings stored in the keybindings file will instantly be incorrect.
   
   ' Create the default keyboard bindings / shortcut mappings
   frmKeyboard_CreateDefaultKeyBindings()
   
   ' Destroy any existing accelerator table
   pWindow->DestroyAcceleratorTable
   
   ' The following ALWAYS get created and can not be altered because they are shortcuts
   ' to activate each of the top menus.
   pWindow->AddAccelerator( FVIRTKEY Or FALT,     VK_F,   IDC_MENUBAR_FILE )
   pWindow->AddAccelerator( FVIRTKEY Or FALT,     VK_E,   IDC_MENUBAR_EDIT )
   pWindow->AddAccelerator( FVIRTKEY Or FALT,     VK_S,   IDC_MENUBAR_SEARCH )
   pWindow->AddAccelerator( FVIRTKEY Or FALT,     VK_V,   IDC_MENUBAR_VIEW )
   pWindow->AddAccelerator( FVIRTKEY Or FALT,     VK_P,   IDC_MENUBAR_PROJECT )
   pWindow->AddAccelerator( FVIRTKEY Or FALT,     VK_C,   IDC_MENUBAR_COMPILE )
   pWindow->AddAccelerator( FVIRTKEY Or FALT,     VK_H,   IDC_MENUBAR_HELP )

   ' Create the accelerator table based on the data in the gKeys() array
   for i as long = lbound(gKeys) to ubound(gKeys)
      
      dim as CWSTR wszKeys = gKeys(i).wszDefaultKeys

      ' wszUserKeys will override the wszDefaultKeys entry
      if len(gKeys(i).wszUserKeys) then 
         wszKeys = gKeys(i).wszUserKeys
      else
         ' Only create the default keyboard short if it has not been disabled
         if gKeys(i).bDefaultDisabled then
            wszKeys = ""
         end if
      end if
      
      ' No need to proceed further if no User defined shortcut and Default keys disabled
      if len(wszKeys) = 0 then continue for
      
      ' convert the key combination into something that AddAcelerator can understand
      ' F5
      ' TAB
      ' Alt+F4
      ' Shift+F5
      ' Ctrl+F5
      ' Ctrl+Shift+F5
      ' etc...
      
      ' Parse for + signs
      dim as long nCount = AfxStrParseCount( wszKeys, "+" )
      dim as long nVirtValue = FVIRTKEY
      dim as long nAccelValue = 0
      
      for ii as long = nCount to 1 step -1
         ' The accelerator key will always be the last parse regardless of 
         ' number of items in the string
         dim as CWSTR wszString = AfxStrParse( wszKeys, ii, "+" )
         if ii = nCount then
            nAccelValue = frmKeyboard_AccelKeyToValue( wszString )
         else
            nVirtValue = nVirtValue or frmKeyboard_VirtKeyToValue( wszString )
         end if
      next
      if nAccelValue <> 0 then   ' value can be zero if no default or user key defined
         pWindow->AddAccelerator( nVirtValue, nAccelValue, gKeys(i).idAction )
      end if   
   next
   
   pWindow->CreateAcceleratorTable()

   function = 0
End Function

' ========================================================================================
' Check for keybinding conflict with other keys and display warning label message
' ========================================================================================
Function frmKeyboard_CheckForKeyConflict ( _
            byval wszKeys as CWSTR, _
            byval nSkipIndex as long _
            ) as long
            
   dim as CWSTR wszMessage
   dim as CWSTR wszCheck
   
   for i as long = lbound(gKeys) to ubound(gKeys)
      if i = nSkipIndex then continue for
      ' If user binding exists then check that b/c it would override the default binding
      wszCheck = iif( len(gKeys(i).wszUserKeys), gKeys(i).wszUserKeys, gKeys(i).wszDefaultKeys )
      if len(wszCheck) = 0 then continue for
      if wszKeys = wszCheck then
         wszMessage = "Conflict: " & gKeys(i).wszDescription & " (" & wszKeys & ")"
         exit for
      end if   
   next

   dim as HWND hCtrl = GetDlgItem( HWND_FRMKEYBOARD, IDC_FRMKEYBOARD_LBLCONFLICT )
   AfxSetWindowText( hCtrl, wszMessage )

   function = 0
end function


' ========================================================================================
' Process WM_CREATE message for window/dialog: frmKeyboard
' ========================================================================================
Function frmKeyboard_OnCreate( _
            ByVal HWnd As HWnd, _
            ByVal lpCreateStructPtr As LPCREATESTRUCT _
            ) As BOOLEAN

   ' This is a modal popup window so disable the parent window
   DisableAllModeless()

   '  Message cracker macro expects a True to be returned for a successful
   '  OnCreate handler even though returning -1 from a standard WM_CREATE
   '  call would stop creating the window. This is just one of those Windows
   '  inconsistencies.
   Return True
End Function


' ========================================================================================
' Process WM_NOTIFY message for window/dialog: frmMain
' ========================================================================================
function frmKeyboard_OnNotify( _
            byval HWnd as HWnd, _
            byval id as long, _
            byval pNMHDR as NMHDR ptr _
            ) as LRESULT

   select case pNMHDR->code 

      case LVN_ITEMCHANGED 
         if id = IDC_FRMKEYBOARD_LIST1 then
            Dim As Long nCurSel = ListView_GetSelection( HWND_FRMKEYBOARD_LISTVIEW )
            If nCurSel < 0 Then exit function
            dim as CWSTR wszKeys = iif( len(gKeys(nCurSel).wszUserKeys), _
                                          gKeys(nCurSel).wszUserKeys, gKeys(nCurSel).wszDefaultKeys )
            frmKeyboard_CheckForKeyConflict( wszKeys, nCurSel )
         end if   

   end select
   
   Function = 0

end function

' ========================================================================================
' Process WM_COMMAND message for window/dialog: frmKeyboard
' ========================================================================================
Function frmKeyboard_OnCommand( _
            ByVal HWnd As HWnd, _
            ByVal id As Long, _
            ByVal hwndCtl As HWnd, _
            ByVal codeNotify As UINT _
            ) As LRESULT

   Select Case id

      case IDC_FRMKEYBOARD_CMDMODIFY
         if codeNotify = BN_CLICKED THEN
            PostMessage( HWnd, MSG_USER_SHOW_KEYBOARDEDIT, 0, 0 )
         end if

      case IDC_FRMKEYBOARD_CMDCLEAR
         if codeNotify = BN_CLICKED THEN
            Dim As Long nCurSel = ListView_GetSelection( HWND_FRMKEYBOARD_LISTVIEW )
            If nCurSel < 0 Then exit function
            gKeys(nCurSel).wszUserKeys = ""
            dim wszText as wstring * MAX_PATH = gKeys(nCurSel).wszUserKeys
            FF_ListView_SetItemText( HWND_FRMKEYBOARD_LISTVIEW, nCurSel, 5, wszText, MAX_PATH ) 
            frmKeyboard_CheckForKeyConflict( "", nCurSel )
            ListView_SelectItem( HWND_FRMKEYBOARD_LISTVIEW, nCurSel )
            SetFocus( HWND_FRMKEYBOARD_LISTVIEW )
         end if

      Case IDCANCEL
         If codeNotify = BN_CLICKED Then
            SendMessage( HWnd, WM_CLOSE, 0, 0 )
            Exit Function
         End If
   End Select

   Function = 0
End Function


' ========================================================================================
' Process WM_CLOSE message for window/dialog: frmKeyboard
' ========================================================================================
Function frmKeyboard_OnClose( byval HWnd As HWnd ) As LRESULT
   dim as CWSTR wszFilename = AfxGetExePathName & "settings\keybindings.ini" 
   frmKeyboard_SaveKeyBindings( wszFilename )
   frmKeyboard_BuildAcceleratorTable

   ' Enables parent window keeping parent's zorder
   EnableAllModeless()
   DestroyWindow( HWnd )
   Function = 0
End Function


' ========================================================================================
' Process WM_DESTROY message for window/dialog: frmKeyboard
' ========================================================================================
Function frmKeyboard_OnDestroy( byval HWnd As HWnd) As LRESULT
   PostQuitMessage(0)
   Function = 0
End Function


' ========================================================================================
' Processes messages for the subclassed ListBox window.
' ========================================================================================
Function frmKeyboard_ListView_SubclassProc ( _
            ByVal HWnd   As HWnd, _                 ' // Control window handle
            ByVal uMsg   As UINT, _                 ' // Type of message
            ByVal wParam As WPARAM, _               ' // First message parameter
            ByVal lParam As LPARAM, _               ' // Second message parameter
            ByVal uIdSubclass As UINT_PTR, _        ' // The subclass ID
            ByVal dwRefData As DWORD_PTR _          ' // Pointer to reference data
            ) As LRESULT
   

   ' Convert our ENTER key presses into LBUTTONDBLCLK to process them similarly
   If (uMsg = WM_KEYUP) And (Loword(wParam) = VK_RETURN) Then uMsg = WM_LBUTTONDBLCLK
   
   Select Case uMsg

      Case WM_LBUTTONDBLCLK
         PostMessage( GetParent(HWnd), MSG_USER_SHOW_KEYBOARDEDIT, 0, 0 )
         Exit Function
         
      Case WM_CHAR   ' prevent the annoying beep
         If wParam = VK_RETURN Then Return 0

      Case WM_DESTROY
         ' REQUIRED: Remove control subclassing
         RemoveWindowSubclass( HWnd, @frmKeyboard_ListView_SubclassProc, uIdSubclass )

   End Select

   ' Default processing of Windows messages
   Function = DefSubclassProc( HWnd, uMsg, wParam, lParam )

End Function

' ========================================================================================
' frmKeyboard Window procedure
' ========================================================================================
Function frmKeyboard_WndProc( _
            ByVal HWnd   As HWnd, _
            ByVal uMsg   As UINT, _
            ByVal wParam As WPARAM, _
            ByVal lParam As LPARAM _
            ) As LRESULT

   Select Case uMsg
      HANDLE_MSG (HWnd, WM_CREATE,      frmKeyboard_OnCreate)
      HANDLE_MSG (HWnd, WM_CLOSE,       frmKeyboard_OnClose)
      HANDLE_MSG (HWnd, WM_DESTROY,     frmKeyboard_OnDestroy)
      HANDLE_MSG (HWnd, WM_COMMAND,     frmKeyboard_OnCommand)
      HANDLE_MSG (HWnd, WM_NOTIFY,      frmKeyboard_OnNotify)

      case MSG_USER_SHOW_KEYBOARDEDIT   
         Dim As Long nCurSel = ListView_GetSelection( HWND_FRMKEYBOARD_LISTVIEW )
         If nCurSel < 0 Then Return 0 
         dim wszText as wstring * MAX_PATH
         FF_ListView_GetItemText( HWND_FRMKEYBOARD_LISTVIEW, nCurSel, 0, @wszText, MAX_PATH) 
         dim as long idx = val( wszText )
         if (idx >= lbound(gKeys)) and (idx <= ubound(gKeys)) then
            gKeysEdit = gKeys(idx)
            frmKeyboardEdit_Show( Hwnd )
            gKeys(idx) = gKeysEdit

            ' Update the Default keys column because it could have changed (DISABLED)
            wszText = gKeysEdit.wszDefaultKeys
            if gKeys(idx).bDefaultDisabled then
               wszText = wszText & " (DISABLED)"
            end if   
            FF_ListView_SetItemText( HWND_FRMKEYBOARD_LISTVIEW, nCurSel, 4, wszText, MAX_PATH ) 
            
            ' Update the UserKeys column because it could have changed
            wszText = gKeysEdit.wszUserKeys
            FF_ListView_SetItemText( HWND_FRMKEYBOARD_LISTVIEW, nCurSel, 5, wszText, MAX_PATH ) 
            frmKeyboard_CheckForKeyConflict( wszText, nCurSel )
         end if
         ListView_SelectItem( HWND_FRMKEYBOARD_LISTVIEW, nCurSel )
         SetFocus( HWND_FRMKEYBOARD_LISTVIEW )

   End Select
   
   ' for messages that we don't deal with
   Function = DefWindowProc( HWnd, uMsg, wParam, lParam )

End Function


' ========================================================================================
' frmKeyboard_Show
' ========================================================================================
Function frmKeyboard_Show( ByVal hWndParent As HWnd ) As LRESULT

   dim hCtrl as HWnd

   '  Create the main window and child controls
   Dim pWindow As CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowOwnerPtr(hwndParent)->DPI

   HWND_FRMKEYBOARD = _
   pWindow->Create(hWndParent, L(220,"Keyboard Shortcuts"), _
        @frmKeyboard_WndProc, 0, 0, 0, 0, _
        WS_POPUP Or WS_CAPTION Or WS_SYSMENU Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, _
        WS_EX_DLGMODALFRAME Or WS_EX_CONTROLPARENT Or WS_EX_LEFT )
   pWindow->SetClientSize(622, 436)
   pWindow->Center(pWindow->hWindow, hWndParent)

   HWND_FRMKEYBOARD_LISTVIEW = _
      pWindow->AddControl("LISTVIEW", , IDC_FRMKEYBOARD_LIST1, "", 0, 0, 622, 364, _
      WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or LVS_REPORT Or LVS_SHOWSELALWAYS or LVS_SINGLESEL, _
      WS_EX_LEFT Or WS_EX_RIGHTSCROLLBAR, , _
      Cast(SUBCLASSPROC, @frmKeyboard_ListView_SubclassProc), IDC_FRMKEYBOARD_LIST1, Cast(DWORD_PTR, @pWindow))
                                     
   ' Make the listview header flat
   ListView_MakeHeaderFlat( HWND_FRMKEYBOARD_LISTVIEW )
          
   ' Add some extended styles
   Dim dwExStyle As DWORD
   dwExStyle = ListView_GetExtendedListViewStyle( HWND_FRMKEYBOARD_LISTVIEW )
   dwExStyle = dwExStyle Or LVS_EX_FULLROWSELECT Or LVS_EX_GRIDLINES Or LVS_EX_DOUBLEBUFFER Or LVS_EX_FLATSB
   ListView_SetExtendedListViewStyle( HWND_FRMKEYBOARD_LISTVIEW, dwExStyle )

   ' Configure the ListView
   dim as HWND hLV = HWND_FRMKEYBOARD_LISTVIEW
   ListView_AddColumn( hLV, 0, "idMenu", 0 )  ' hidden
   ListView_AddColumn( hLV, 1, "", pWindow->ScaleX(10) )  ' left padding
   ListView_AddColumn( hLV, 2, "Category", pWindow->ScaleX(90) )
   ListView_AddColumn( hLV, 3, "Description", pWindow->ScaleX(250) )
   ListView_AddColumn( hLV, 4, "Default Keys", pWindow->ScaleX(125) )
   ListView_AddColumn( hLV, 5, "User Keys", pWindow->ScaleX(125) )
   
   pWindow->AddControl("BUTTON", , IDC_FRMKEYBOARD_CMDMODIFY, L(433, "Modify"), 8, 388, 74, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("BUTTON", , IDC_FRMKEYBOARD_CMDCLEAR, L(434, "Clear"), 90, 388, 74, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("BUTTON", , IDCANCEL, L(161,"Close"), 536, 388, 74, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)

   hCtrl = pWindow->AddControl("LABEL", , IDC_FRMKEYBOARD_LBLCONFLICT, "", 168, 394, 350, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_CENTER Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   AfxSetWindowFont( hCtrl, ghStatusBar.hFontStatusBar, true )

   ' Load the listview with these temp keys
   for i as long = lbound(gKeys) to ubound(gKeys)
      FF_ListView_InsertItem( hLV, i, 0, str(i) )  ' array index (hidden column)
      FF_ListView_InsertItem( hLV, i, 1, "" ) 
      FF_ListView_InsertItem( hLV, i, 2, gKeys(i).wszCategory ) 
      FF_ListView_InsertItem( hLV, i, 3, gKeys(i).wszDescription ) 
      dim as CWSTR wszTemp = gKeys(i).wszDefaultKeys
      if gKeys(i).bDefaultDisabled then
         wszTemp = wszTemp & " (DISABLED)"
      end if
      FF_ListView_InsertItem( hLV, i, 4, wszTemp ) 
      FF_ListView_InsertItem( hLV, i, 5, gKeys(i).wszUserKeys ) 
   next

   AfxSetWindowFont( hLV, ghStatusBar.hFontStatusBar, true )

   ListView_SelectItem( hLV, 0 )
   SetFocus( hLV )    
   
   ShowWindow( HWND_FRMKEYBOARD, SW_SHOW )
   UpdateWindow HWND_FRMKEYBOARD

   ' Message loop (modal)
   DIM uMsg AS MSG
   WHILE GetMessage( @uMsg, NULL, 0, 0 )
      if (uMsg.message = WM_KEYDOWN) andalso (uMsg.wParam = VK_ESCAPE) then
         SendMessage( HWND_FRMKEYBOARD, WM_CLOSE, 0, 0 )
      end if
      IF IsDialogMessage( HWND_FRMKEYBOARD, @uMsg ) = 0 THEN
         TranslateMessage( @uMsg )
         DispatchMessage( @uMsg )
      end if
   WEND
   function = uMsg.wParam
   
   ' Delete the CWindow class manually allocated memory 
   Delete pWindow

End Function

