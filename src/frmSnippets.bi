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

#include once "clsDocument.bi"

#DEFINE IDC_FRMSNIPPETS_LIST1                       1000
#DEFINE IDC_FRMSNIPPETS_TXTDESCRIPTION              1001
#DEFINE IDC_FRMSNIPPETS_TXTTRIGGER                  1002
#DEFINE IDC_FRMSNIPPETS_TXTCODE                     1003
#DEFINE IDC_FRMSNIPPETS_LABEL1                      1004
#DEFINE IDC_FRMSNIPPETS_LABEL2                      1005
#DEFINE IDC_FRMSNIPPETS_LABEL3                      1006
#DEFINE IDC_FRMSNIPPETS_LABEL4                      1007
#DEFINE IDC_FRMSNIPPETS_CMDUP                       1008
#DEFINE IDC_FRMSNIPPETS_CMDDOWN                     1009
#DEFINE IDC_FRMSNIPPETS_CMDINSERT                   1010
#DEFINE IDC_FRMSNIPPETS_CMDDELETE                   1011
#DEFINE IDC_FRMSNIPPETS_CMDOK                       1012
#DEFINE IDC_FRMSNIPPETS_LBLPARAMETERS               1013
#DEFINE IDM_FRMSNIPPETS_PARAMETERBASE               2000

declare function frmSnippets_DoInsertSnippet( byval pDoc as clsDocument ptr ) as Boolean
declare Function frmSnippets_Show( ByVal hWndParent As HWnd ) as LRESULT


