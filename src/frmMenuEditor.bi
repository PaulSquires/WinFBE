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

#DEFINE IDC_FRMMENUEDITOR_LABEL1                    1000
#DEFINE IDC_FRMMENUEDITOR_LABEL2                    1001
#DEFINE IDC_FRMMENUEDITOR_LABEL3                    1002
#DEFINE IDC_FRMMENUEDITOR_TXTCAPTION                1003
#DEFINE IDC_FRMMENUEDITOR_TXTNAME                   1004
#DEFINE IDC_FRMMENUEDITOR_CMBOSHORTCUT              1005
#DEFINE IDC_FRMMENUEDITOR_CHKCHECKED                1006
#DEFINE IDC_FRMMENUEDITOR_CHKGRAYED                 1007
#DEFINE IDC_FRMMENUEDITOR_CMDDELETE                 1008
#DEFINE IDC_FRMMENUEDITOR_CMDINSERT                 1009
#DEFINE IDC_FRMMENUEDITOR_CMDNEXT                   1010
#DEFINE IDC_FRMMENUEDITOR_CMDCANCEL                 1011
#DEFINE IDC_FRMMENUEDITOR_CMDOK                     1012
#DEFINE IDC_FRMMENUEDITOR_LABEL4                    1013
#DEFINE IDC_FRMMENUEDITOR_LSTDETAIL                 1014
#DEFINE IDC_FRMMENUEDITOR_CMDLEFT                   1015
#DEFINE IDC_FRMMENUEDITOR_CMDRIGHT                  1016
#DEFINE IDC_FRMMENUEDITOR_CMDUP                     1017
#DEFINE IDC_FRMMENUEDITOR_CMDDOWN                   1018
#Define IDC_FRMMENUEDITOR_CHKCTRL                   1019
#Define IDC_FRMMENUEDITOR_CHKALT                    1020
#Define IDC_FRMMENUEDITOR_CHKSHIFT                  1021
#Define IDC_FRMMENUEDITOR_LBLSHORTCUT               1022
#Define IDC_FRMMENUEDITOR_LBLSTATE                  1023
#Define IDC_FRMMENUEDITOR_CHKDISPLAYONFORM          1024

declare Function frmMenuEditor_CreateFakeMainMenu( ByVal pDoc as clsDocument ptr ) As Long
declare Function frmMenuEditor_Show( ByVal hWndParent As HWnd ) as LRESULT
