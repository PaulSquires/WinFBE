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
' Forward references
Type clsDocument_ As clsDocument
TYPE TOPTABS_TYPE
   pDoc as clsDocument_ ptr
   wszText as CWSTR
   rcTab as RECT          ' client coordinates 
   rcIcon as RECT         ' client coordinates 
   rcText as RECT         ' client coordinates 
   rcClose as RECT        ' client coordinates 
   isHot as boolean
end type


Type clsTopTabCtl
   Private:
      
   Public:
      hWindow          as HWnd
      ClientRightEdge  as long        ' the right edge (client right - action Panel)
      CurSel           as long = -1
      FirstDisplayTab  as long = 0   
      rcActionPanel    as RECT
      rcActionButton   as RECT
      rcPrevTabs       as RECT
      rcNextTabs       as RECT
      rcSplitEditor    as RECT
      tabs(any)        as TOPTABS_TYPE
      
      declare function IsSafeIndex( byval idx as long ) as boolean
      declare function GetItemCount() as long
      declare function RemoveElement( byval idx as long ) as long
      Declare Function AddTab( ByVal pDoc As clsDocument Ptr ) As Long
      Declare Function GetTabIndexFromFilename( Byref wszName As WString ) As Long
      declare Function GetTabIndexByDocumentPtr( ByVal pDocIn As clsDocument Ptr ) As Long
      Declare Function SetTabIndexByDocumentPtr( ByVal pDocIn As clsDocument Ptr ) As Long
      Declare Function SetFocusTab( ByVal idx As Long ) As Long
      Declare Function GetActiveDocumentPtr() As clsDocument Ptr
      Declare Function GetDocumentPtr( ByVal idx As Long ) As clsDocument Ptr
      Declare Function DisplayScintilla( ByVal idx As Long, ByVal bShow As BOOLEAN ) As Long
      Declare Function SetTabText( ByVal idx As Long ) As Long
      Declare Function NextTab() As Long
      Declare Function PrevTab() As Long
      Declare Function CloseTab( ByVal idx As Long = -1 ) As Long
      
End Type

