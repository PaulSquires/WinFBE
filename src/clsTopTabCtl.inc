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

#include once "clsTopTabCtl.bi"

''
''
function clsTopTabCtl.IsSafeIndex( byval idx as long ) as boolean
   ' test to see if the incoming idx is valid based on the boundaries
   ' of the global tab array
   if (idx >= lbound(this.tabs)) andalso (idx <= ubound(this.tabs)) then
      return true
   else   
      return false   
   end if
end function

''
''
function clsTopTabCtl.GetItemCount() as long
   ' get the number tabs in the global tab array
   return ubound(this.tabs) - lbound(this.tabs) + 1
end function

''
''
Function clsTopTabCtl.AddTab( ByVal pDoc As clsDocument Ptr ) As Long
   ' Return the index of the added tab. we attempt to add it immediately
   ' after the current active tab.
   Dim idx As Long
   redim preserve this.tabs(ubound(this.tabs) + 1) as TOPTABS_TYPE
   
   if (this.CurSel = -1) orelse (this.CurSel = this.GetItemCount - 1) then
      ' add the tab at the end of the array
      idx = ubound(this.tabs)
   else
      ' insert the tab after the current tab
      for i as long = ubound(this.tabs) to this.CurSel + 1 step -1
         this.tabs(i) = this.tabs(i - 1)
      next
      idx = this.CurSel + 1
   end if
   
   this.tabs(idx).pDoc = pDoc
   this.SetTabText(idx)

   Function = idx
End Function

''
''
function clsTopTabCtl.RemoveElement( byval idx as long ) as long
   ' remove an element from the this.tabs array
   if this.IsSafeIndex(idx) = false then exit function
   ' if this is the last element being removed then we need
   ' to erase the global array
   if this.GetItemCount() = 1 then
      erase this.tabs
   else
      for i as long = idx to ubound(this.tabs) - 1
         this.tabs(i) = this.tabs(i+1)
      next   
      redim preserve this.tabs(ubound(this.tabs)-1)
   end if
   function = 0
end function


''
''
Function clsTopTabCtl.GetTabIndexFromFilename( Byref wszName As WString ) As Long
   ' Get the index of tab that holds the incoming filename. This is used
   ' mostly for positioning the editor to the correct document when an
   ' error occurs during compiling. 
   ' Returns -1 if not found, otherwise zero based index of tab.
   
   Dim As Long nCount = this.GetItemCount()
   If nCount = 0 Then Return -1 
   If len(wszName) = 0 Then Return -1
   
   for i as long = 0 To nCount - 1
      if this.tabs(i).pDoc then
         If Ucase(this.tabs(i).pDoc->DiskFilename) = Ucase(wszName) Then Return i
      end if   
   Next

   Function = -1  ' if not found
End Function


''
''
Function clsTopTabCtl.GetTabIndexByDocumentPtr( ByVal pDocIn As clsDocument Ptr ) As Long

   Dim As Long nCount = this.GetItemCount()
   If nCount = 0 Then Return -1 
   If pDocIn = 0 Then Return -1

   for i as long = 0 To nCount - 1
      if this.tabs(i).pDoc then
         If this.tabs(i).pDoc = pDocIn Then Return i
      end if   
   next

   Function = -1  ' if not found
End Function


''
''
Function clsTopTabCtl.SetTabIndexByDocumentPtr( ByVal pDocIn As clsDocument Ptr ) As Long
   Dim As Long nCount = this.GetItemCount()
   If nCount = 0 Then Return -1 
   If pDocIn = 0 Then Return -1
   for i as long = 0 To nCount - 1
      If this.tabs(i).pDoc = pDocIn Then 
         function = this.SetFocusTab(i)
         Exit Function
      end if   
   next
   Function = -1  ' if not found
End Function


''
''
Function clsTopTabCtl.SetFocusTab( ByVal idx As Long ) As Long

   Dim As Long nCount = this.GetItemCount()
   If nCount = 0 Then Return -1 
   if idx < 0 then idx = nCount - 1
   if this.IsSafeIndex(idx) = false then exit function
   if this.CurSel <> idx then
      ' Send a user message to accomplish the same thing as TCN_SELCHANGING and TCN_SELCHANGE 
      SendMessage( HWND_FRMMAIN, MSG_USER_TOPTABS_CHANGING, 0, 0 )
      this.CurSel = idx
      SendMessage( HWND_FRMMAIN, MSG_USER_TOPTABS_CHANGED, 0, 0 )
      
      ' V3.0.0 Disable the highlighting of searches when a user switches tabs? Report
      ' from user that such action disrupts the user's train of thought because the spot
      ' is lost within the file. Other users may feel differently?
      if IsWindowVisible(HWND_FRMFINDREPLACE) then
         frmFindReplace_HighlightSearches( false )
      end if
      
      ' Highlight the selected tab file in the Explorer listbox
      frmExplorer_SelectItemData( this.tabs(idx).pDoc )
      ' Highlight the selected tab file in the Function List listbox
      frmFunctions_SelectItemData( this.tabs(idx).pDoc )
   end if

   function = this.CurSel
End Function


''
''
Function clsTopTabCtl.NextTab() As Long
   ' Invoked by Ctl+TAB
   ' Set the tab with focus (this sends the TCN_SELCHANGING and TCN_SELCHANGE 
   ' notification codes to its parent window). 
   Dim As Long nCount = this.GetItemCount
   Dim As Long idx = this.CurSel
   If nCount = 0 Then Exit Function
   idx += 1
   If idx > nCount - 1 Then idx = 0
   Function = this.SetFocusTab(idx)
End Function

''
''
Function clsTopTabCtl.PrevTab() As Long
   ' Invoked by Ctl+Shift+TAB
   ' Set the tab with focus (this sends the TCN_SELCHANGING and TCN_SELCHANGE 
   ' notification codes to its parent window). 
   Dim As Long nCount = this.GetItemCount 
   Dim As Long idx = this.CurSel
   If nCount = 0 Then Exit Function
   idx -= 1
   If idx < 0 Then idx = nCount - 1 
   Function = this.SetFocusTab(idx)
End Function

''
''
Function clsTopTabCtl.CloseTab( byval nTabIdx as long = -1) As Long
   ' Invoked by clicking "X" on tab
   OnCommand_FileClose( HWND_FRMMAIN, EFC_CLOSECURRENT, nTabIdx )
   Function = 0
End Function


''
''
Function clsTopTabCtl.GetActiveDocumentPtr() As clsDocument Ptr
   ' Return pointer to clsDocument class for the current active tab
   If this.GetItemCount() = 0 Then Return 0 
   if this.IsSafeIndex(this.CurSel) = false then exit function
   function = this.tabs(this.CurSel).pDoc   
End Function

''
''
Function clsTopTabCtl.GetDocumentPtr( ByVal idx As Long ) As clsDocument Ptr
   ' Return pointer to clsDocument class for the current active tab
   if this.IsSafeIndex(idx) = false then exit function
   function = this.tabs(idx).pDoc
End Function

''
''
Function clsTopTabCtl.DisplayScintilla( ByVal idx As Long, ByVal bShow As BOOLEAN ) As Long
   ' Show/Hide the Scintilla editing window (or visual designer window) for the incoming tab index
   Dim pDocShow As clsDocument Ptr

   if this.IsSafeIndex(idx) = false then exit function
   pDocShow = this.tabs(idx).pDoc
   If pDocShow = 0 Then exit function

   ' Hide all documents and their associated scrollbars
   Dim pDoc As clsDocument Ptr = gApp.pDocList
   
   do until pDoc = 0
      ' Scintilla windows and scrollbars
      ShowWindow(pDoc->hWindow(0), SW_HIDE)
      ShowWindow(pDoc->hWindow(1), SW_HIDE)
      if pDoc->IsDesigner THEN 
         ShowWindow(pDoc->hWndDesigner, SW_HIDE)
      END IF
      pDoc = pDoc->pDocNext
   loop
   
   if (pDocShow->IsDesigner) andalso (IsDesignerView(pDocShow)) THEN
      ShowWindow(pDocShow->hWndDesigner, Iif(bShow, SW_SHOWNORMAL, SW_HIDE))
      ShowWindow(HWND_FRMVDTOOLBOX, Iif(bShow, SW_SHOWNORMAL, SW_HIDE))
   else
      ' Show/Hide our current active Scintilla window and scrollbar
      ShowWindow(pDocShow->hWindow(0), Iif(bShow, SW_SHOWNORMAL, SW_HIDE))
      ShowWindow(pDocShow->hWindow(1), Iif(bShow, SW_SHOWNORMAL, SW_HIDE))
   end if
   
   Function = 0
End Function
   
''
''
Function clsTopTabCtl.SetTabText( ByVal idx As Long ) As Long
   ' Set the text for the incoming tab index. If the index
   ' is -1 then set the text of the current tab.
   Dim wszText As WString * MAX_PATH  
   Dim wszTemp As WString * MAX_PATH  
   Dim pDoc As clsDocument Ptr

   If idx = -1 Then idx = gTTabCtl.CurSel
   if this.IsSafeIndex(idx) = false then exit function

   pDoc = this.tabs(idx).pDoc
   If pDoc Then 
      ' Set the text that displays on the tab
      wszText = AfxStrPathname( "NAMEX", pDoc->DiskFilename )
            
      ' We only update the text if it has changed in order to prevent flicker.
      If wszText <> this.tabs(idx).wszText Then
         this.tabs(idx).wszText = wszText
         frmTopTabs_PositionWindows()
      End If
   End If
   
   function = idx
End Function

