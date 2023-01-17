'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2022 Paul Squires, PlanetSquires Software
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

#DEFINE IDC_FRMUSERTOOLS_LSTTOOLS                   1000
#DEFINE IDC_FRMUSERTOOLS_CMDINSERT                  1001
#DEFINE IDC_FRMUSERTOOLS_CMDDELETE                  1002
#DEFINE IDC_FRMUSERTOOLS_CMDUP                      1003
#DEFINE IDC_FRMUSERTOOLS_CMDDOWN                    1004
#DEFINE IDC_FRMUSERTOOLS_TXTTOOLNAME                1005
#DEFINE IDC_FRMUSERTOOLS_TXTCOMMAND                 1006
#DEFINE IDC_FRMUSERTOOLS_TXTPARAMETERS              1007
#DEFINE IDC_FRMUSERTOOLS_TXTKEY                     1008
#DEFINE IDC_FRMUSERTOOLS_LABEL1                     1009
#DEFINE IDC_FRMUSERTOOLS_LABEL2                     1010
#DEFINE IDC_FRMUSERTOOLS_LABEL3                     1011
#DEFINE IDC_FRMUSERTOOLS_CMDBROWSEEXE               1012
#DEFINE IDC_FRMUSERTOOLS_LABEL4                     1013
#DEFINE IDC_FRMUSERTOOLS_TXTWORKINGFOLDER           1014
#DEFINE IDC_FRMUSERTOOLS_CMDBROWSEFOLDER            1015
#DEFINE IDC_FRMUSERTOOLS_LABEL5                     1016
#DEFINE IDC_FRMUSERTOOLS_LABEL6                     1017
#DEFINE IDC_FRMUSERTOOLS_CHKCTRL                    1018
#DEFINE IDC_FRMUSERTOOLS_CHKALT                     1019
#DEFINE IDC_FRMUSERTOOLS_CHKSHIFT                   1020
#DEFINE IDC_FRMUSERTOOLS_LABEL7                     1021
#DEFINE IDC_FRMUSERTOOLS_CHKPROMPT                  1022
#DEFINE IDC_FRMUSERTOOLS_CHKMINIMIZED               1023
#DEFINE IDC_FRMUSERTOOLS_CHKWAIT                    1024
#DEFINE IDC_FRMUSERTOOLS_COMBOACTION                1025
#DEFINE IDC_FRMUSERTOOLS_LABEL8                     1026
#DEFINE IDC_FRMUSERTOOLS_LABEL9                     1027
#Define IDC_FRMUSERTOOLS_CHKDISPLAYMENU             1028
#Define IDC_FRMUSERTOOLS_CMDOK                      1029

' User Tool actions (selected, pre/post compile, after WinFBE starts up)
const USERTOOL_ACTION_SELECTED      = 0   
const USERTOOL_ACTION_PRECOMPILE    = 1   
const USERTOOL_ACTION_POSTCOMPILE   = 2
const USERTOOL_ACTION_WINFBESTARTUP = 3

common shared as HACCEL ghAccelUserTools

declare Function frmUserTools_ExecuteUserTool( ByVal nToolNum As Long ) As Long            
declare Function frmUserTools_CreateAcceleratorTable() As Long
declare Function frmUserTools_Show( ByVal hWndParent As HWnd ) as LRESULT
declare Function updateUserToolsMenuItems() as long
declare function createToolsMenuShortcut( byval nCtrlID as long ) as CWSTR







