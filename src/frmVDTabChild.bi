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

#pragma once


#DEFINE IDC_FRMVDTABCHILD_LIST1                    1000
#DEFINE IDC_FRMVDTABCHILD_TXTDESCRIPTION           1001
#DEFINE IDC_FRMVDTABCHILD_TXTIMAGE                 1002
#DEFINE IDC_FRMVDTABCHILD_CMDIMAGE                 1003
#DEFINE IDC_FRMVDTABCHILD_COMBOTABPAGES            1004
#DEFINE IDC_FRMVDTABCHILD_CHKISDEFAULT             1005
#DEFINE IDC_FRMVDTABCHILD_CMDADD                   1006
#DEFINE IDC_FRMVDTABCHILD_CMDINSERT                1007
#DEFINE IDC_FRMVDTABCHILD_CMDDELETE                1008
#DEFINE IDC_FRMVDTABCHILD_CMDUP                    1009
#DEFINE IDC_FRMVDTABCHILD_CMDDOWN                  1010
#DEFINE IDC_FRMVDTABCHILD_CMDOK                    1011
#DEFINE IDC_FRMVDTABCHILD_CMDCANCEL                1012
#DEFINE IDC_FRMVDTABCHILD_LABEL1                   1013
#DEFINE IDC_FRMVDTABCHILD_LABEL2                   1014
#DEFINE IDC_FRMVDTABCHILD_LABEL3                   1015

type TabPage
   wszText      as CWSTR
   wszImage     as CWSTR
   wszTabPage   as CWSTR 
   IsActiveTab  as long     ' must be Long rather than boolean
   wszReserved1 as CWSTR    ' for future use
   wszReserved2 as CWSTR 
   wszReserved3 as CWSTR 
   wszReserved4 as CWSTR 
end type

dim shared gTabPages(any) as TabPage

dim shared gTabRecordSep as CWSTR = "-[]-"
dim shared gTabFieldSep as CWSTR = "-||-"

declare function frmVDTabChild_LoadTabPagesArray( byref wszPropValue as wstring ) as Long
declare Function frmVDTabChild_Show( ByVal hWndParent As HWnd, byref wszPropValue as wstring ) As LRESULT

