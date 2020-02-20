'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2020 Paul Squires, PlanetSquires Software
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

'#include once "clsDocument.bi"

Type clsTopTabCtl
   Private:
      
   Public:
      hWindow   As HWnd
      Declare Function AddTab( ByVal pDoc As clsDocument Ptr ) As Long
      Declare Function GetTabIndexFromFilename( ByVal pwszName As WString Ptr ) As Long
      declare Function GetTabIndexByDocumentPtr( ByVal pDocIn As clsDocument Ptr ) As Long
      Declare Function SetTabIndexByDocumentPtr( ByVal pDocIn As clsDocument Ptr ) As Long
      Declare Function SetFocusTab( ByVal idx As Long ) As Long
      Declare Function GetActiveDocumentPtr() As clsDocument Ptr
      Declare Function GetDocumentPtr( ByVal idx As Long ) As clsDocument Ptr
      Declare Function DisplayScintilla( ByVal idx As Long, ByVal bShow As BOOLEAN ) As Long
      Declare Function SetTabText( ByVal idx As Long ) As Long
      Declare Function NextTab() As Long
      Declare Function PrevTab() As Long
      Declare Function CloseTab() As Long
      
End Type

