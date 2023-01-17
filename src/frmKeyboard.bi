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

#pragma once


#DEFINE IDC_FRMKEYBOARD_LIST1           1000
#DEFINE IDC_FRMKEYBOARD_CMDMODIFY       1001
#DEFINE IDC_FRMKEYBOARD_CMDCLEAR        1002
#DEFINE IDC_FRMKEYBOARD_LBLCONFLICT     1003

#define IDC_FRMKEYBOARDEDIT_CHKCTRL     1100
#define IDC_FRMKEYBOARDEDIT_CHKALT      1101
#define IDC_FRMKEYBOARDEDIT_CHKSHIFT    1102
#define IDC_FRMKEYBOARDEDIT_CHKDISABLE  1103
#define IDC_FRMKEYBOARDEDIT_LABEL1      1104
#define IDC_FRMKEYBOARDEDIT_LABEL2      1105
#define IDC_FRMKEYBOARDEDIT_LABEL3      1106
#define IDC_FRMKEYBOARDEDIT_COMBOACCEL  1107
#define IDC_FRMKEYBOARDEDIT_CHKDISABLED 1108

TYPE KEYBINDINGS_TYPE
   idAction as long         ' IDM_* message
   wszMsgString as CWSTR    ' "IDM_SAVE", "IDM_SAVEAS", etc 
   wszCategory as CWSTR
   wszDescription as CWSTR
   wszDefaultKeys as CWSTR
   wszUserKeys as CWSTR
   bDefaultDisabled as boolean = false
end type
dim shared gKeys(any) as KEYBINDINGS_TYPE
dim shared gKeysEdit as KEYBINDINGS_TYPE

declare Function frmKeyboard_Show( ByVal hWndParent As HWnd ) As LRESULT
declare function frmKeyboard_SaveKeyBindings( byval wszFilename as CWSTR ) as long
declare function frmKeyBoard_AddKeyBinding( _
            byval wszCategory as CWSTR, _
            byval idAction as long, _
            byval wszMsgString as CWSTR, _
            byval wszDescription as CWSTR, _
            byval wszDefaultKeys as CWSTR, _
            byval wszUserKeys as CWSTR, _
            byval bDisabled as boolean _
            ) as long
declare function frmKeyboard_CheckForKeyConflict ( byval wszKeys as CWSTR, byval nSkipIndex as long ) as long

