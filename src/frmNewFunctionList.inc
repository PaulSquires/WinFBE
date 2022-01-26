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


''
''  frmNewFunctionList.inc
''   

#include once "frmFunctionList.bi"
#include once "clsDB2.bi"


' ========================================================================================
' Reparse any dirty files
' ========================================================================================
public function frmFunctionList_ReparseFiles() as Long
   '  Need to re-parse any dirty files. This is especially important for QuickRun scenarios
   '  where the file is never physically saved.
   Dim pDoc As clsDocument Ptr = gApp.pDocList
   do until pDoc = 0
      If cbool(SciExec( pDoc->hWindow(0), SCI_GETMODIFY, 0, 0 )) or pDoc->UserModified Then
         pDoc->bNeedsParsing = true
         ParseDocument(pDoc)
         frmFunctionList_AddChildNodes( pDoc )
      end if
      pDoc = pDoc->pDocNext
   loop   
   function = 0
end function


' ========================================================================================
' Create the special root nodes
' ========================================================================================
public function frmFunctionList_CreateSpecialNodes() as HTREEITEM

   dim as hwnd hTV = GetDlgItem( HWND_FRMFUNCTIONLIST, IDC_FRMFUNCTIONLIST_TREE )

   Treeview_DeleteAllItems(hTV)

   gApp.hExplorerRootNode = TreeView_AppendItem( hTV, TVI_ROOT, "" )
   ' Add the subnodes for the various file types
   gApp.hExplorerFilesNode = _
      TreeView_AppendItem( hTV, gApp.hExplorerRootNode, "Files", cast(LPARAM, 0) )
   gApp.hExplorerMainNode = _
      TreeView_AppendItem( hTV, gApp.hExplorerRootNode, "Main", cast(LPARAM, 0) )
   gApp.hExplorerResourceNode = _
      TreeView_AppendItem( hTV, gApp.hExplorerRootNode, "Resource", cast(LPARAM, 0) )
   gApp.hExplorerHeaderNode = _
      TreeView_AppendItem( hTV, gApp.hExplorerRootNode, "Header", cast(LPARAM, 0) )
   gApp.hExplorerModuleNode = _
      TreeView_AppendItem( hTV, gApp.hExplorerRootNode, "Module", cast(LPARAM, 0) )
   gApp.hExplorerNormalNode = _
      TreeView_AppendItem( hTV, gApp.hExplorerRootNode, "Normal", cast(LPARAM, 0) )

   ' reset array that determines which Explorer nodes to display
   for i as long = lbound(gExplorerNodeShow) to ubound(gExplorerNodeShow)
      gExplorerNodeShow(i) = true
   next

   function = 0
end function


' ========================================================================================
' Get the special subnode based on the filetype
' ========================================================================================
public function frmFunctionList_GetSpecialNode( byval nFileType as long ) as HTREEITEM
   
   dim hRootNode as HTREEITEM
   
   ' If project is not active then we retrieve the data from the Files node
   if gApp.IsProjectActive = false then hRootNode = gApp.hExplorerFilesNode

   ' Add the pDoc to the node depending on what file type it is.
   select case nFileType
      case FILETYPE_UNDEFINED
         if ( gApp.IsProjectActive = true ) orelse ( gApp.IsProjectLoading = true ) then
            hRootNode = gApp.hExplorerNormalNode
         end if   
      case FILETYPE_MAIN:      hRootNode = gApp.hExplorerMainNode
      case FILETYPE_MODULE:    hRootNode = gApp.hExplorerModuleNode
      case FILETYPE_NORMAL:    hRootNode = gApp.hExplorerNormalNode
      case FILETYPE_RESOURCE:  hRootNode = gApp.hExplorerResourceNode
      case FILETYPE_HEADER:    hRootNode = gApp.hExplorerHeaderNode
   end select

   function = hRootNode

end function


' ========================================================================================
' Retrieve the filename and function name of the specified Treeview node
' ========================================================================================
private function frmFunctionList_GetFileNameFunctionName( byval hItem as HTREEITEM, _
                                                          byref wszFilename as CWSTR, _
                                                          byref wszFunctionName as CWSTR _
                                                          ) as long

   Dim As Hwnd hTree = GetDlgItem(HWND_FRMFUNCTIONLIST, IDC_FRMFUNCTIONLIST_TREE)
   Dim pDoc As clsDocument Ptr
   dim as LPARAM lParam 
   dim zFilename as wstring * MAX_PATH
   dim zFunctionName as wstring * MAX_PATH
            
   lParam = TreeView_GetlParam(hTree, hItem)
   pDoc = Cast(clsDocument Ptr, lParam)
   if pDoc then
      ' This is a header node. Open the document based on the pDoc info
      ' and position to first line in the file.
      zFilename = pDoc->DiskFilename
      zFunctionName = ""
   else
      ' pDoc is null so the node text contains the name of the sub/function. We
      ' need to get the parent node to retrieve the pDoc in order to get the 
      ' correct filename.
      TreeView_GetItemText(hTree, hItem, @zFunctionName, MAX_PATH)
      hItem = TreeView_GetNextItem(hTree, hItem, TVGN_PARENT)   
      lParam = TreeView_GetlParam(hTree, hItem)
      pDoc = Cast(clsDocument Ptr, lParam)
      if pDoc then zFilename = pDoc->DiskFilename
   end if

   wszFilename = zFilename
   wszFunctionName = zFunctionName
   
   function = 0
end function


' ========================================================================================
' Add the incoming Document to a parent node in the treeview
' ========================================================================================
public Function frmFunctionList_AddParentNode( ByVal pDoc As clsDocument Ptr ) as HTREEITEM
 
   If pDoc = 0 Then Exit Function
   
   Dim As HTREEITEM hSpecialNode, hItem
   
   Dim As Hwnd hTree    = GetDlgItem(HWND_FRMFUNCTIONLIST, IDC_FRMFUNCTIONLIST_TREE)
   Dim As CWSTR wszFile = AfxStrPathname("NAMEX", pDoc->DiskFilename)

   ' Give this document a default project type depending on its file extension
   if (pDoc->IsNewFlag = false) andalso (pDoc->ProjectFileType = FILETYPE_UNDEFINED) then
      if ( gApp.IsProjectActive = true ) orelse ( gApp.IsProjectLoading = true ) then
         if pDoc->IsDesigner then 
            pDoc->ProjectFileType = FILETYPE_NORMAL 
         else
            gApp.ProjectSetFileType( pDoc, pDoc->ProjectFileType )
         end if
      end if   
   end if

   ' If the node already exists in the tree then delete it because we will add it again. This is
   ' necessary for situations where the filename changes and the file then moves to a different
   ' node branch.
   if pDoc->hNodeExplorer then
      TreeView_DeleteItem( hTree, pDoc->hNodeExplorer )
      pDoc->hNodeExplorer  = 0
   end if

   hItem = pDoc->hNodeExplorer 
   if hItem = 0 then
      hSpecialNode = frmFunctionList_GetSpecialNode( pDoc->ProjectFileType )
      hItem = TreeView_AppendItem( hTree, hSpecialNode, wszFile, cast(LPARAM, pDoc), 1, 1 )
   end if

   function = hItem

end function


' ========================================================================================
' Add the incoming Document to a parent node in the treeview
' ========================================================================================
public Function frmFunctionList_AddChildNodes( ByVal pDoc As clsDocument Ptr ) as long
 
   If pDoc = 0 Then Exit Function
   
   Dim As HTREEITEM hItem

   Dim As Hwnd hTree = GetDlgItem(HWND_FRMFUNCTIONLIST, IDC_FRMFUNCTIONLIST_TREE)

   ' Delete any existing children sub/function nodes and then update with new values
   hItem = TreeView_GetChild( hTree, pDoc->hNodeExplorer )
   Do Until hItem = 0
      TreeView_DeleteItem( hTree, hItem )
      hItem = TreeView_GetChild( hTree, pDoc->hNodeExplorer )
   Loop

   dim pData as DB2_DATA ptr
   dim as CWSTR wszFilename = ucase(pDoc->DiskFilename)
   dim as CWSTR wszNodeName
    
   gdb2.dbRewind()
   do 
      pData = gdb2.dbGetNext
      if pData = 0 THEN exit do
      if pData->deleted then continue do
      if pData->nFileType <> DB2_FILETYPE_USERCODE then continue do
      select case pData->id 
         case DB2_FUNCTION, DB2_SUB, DB2_PROPERTY
            if wszFilename = ucase(pData->fileName) THEN
               wszNodeName = pData->ElementName
               select case pData->GetSet 
                  case ClassProperty.Getter: wszNodeName = wszNodeName & " (get)"
                  case ClassProperty.Setter: wszNodeName = wszNodeName & " (set)"
                  case ClassProperty.ctor:   wszNodeName = wszNodeName & " (ctor)"
                  case ClassProperty.dtor:   wszNodeName = wszNodeName & " (dtor)"
               end select      
               hItem = TreeView_AppendItem( hTree, pDoc->hNodeExplorer, wszNodeName, 0, 2, 2 )
            END IF
      END SELECT
   loop      
       
   function = 0
end function

' ========================================================================================
' Set the Listbox selected item to the currently selected open document (if applicable)
' ========================================================================================
private Function frmFunctionList_SetListBoxPosition() As Long

   Dim nCurSel As Long = -1

   Dim pDocActive As clsDocument Ptr = gTTabCtl.GetActiveDocumentPtr() 
   Dim pDoc As clsDocument Ptr
   Dim As HTREEITEM hItem

   Dim As Hwnd hTree = GetDlgItem(HWND_FRMFUNCTIONLIST, IDC_FRMFUNCTIONLIST_TREE)
   Dim As HWnd hList  = GetDlgItem(HWND_FRMFUNCTIONLIST, IDC_FRMFUNCTIONLIST_LISTBOX)
   Dim As Long nCount = Listbox_GetCount(hList)
   dim as integer nParam  ' must be INTEGER otherwise the Cast below fails on 64-bit.
   dim wszListText as wstring * MAX_PATH
   
   ' Search for a match
   For i As Long = 0 To nCount - 1
      hItem = cast(HTREEITEM, ListBox_GetItemData(hList, i))
      nParam = TreeView_GetlParam(hTree, hItem)
      pDoc = Cast(clsDocument ptr, nParam)
      if pDoc = pDocActive THEN
         ' Attempt to match the function name in the child nodes
         dim as string sCurrentFunction
         dim as ClassProperty nGetSet
         pDoc->GetCurrentFunctionName( sCurrentFunction, nGetSet )
         For ii As Long = i To nCount - 1
            SendMessage( hList, LB_GETTEXT, ii, Cast(LPARAM, @wszListText))
            if ucase(wszListText) = ucase(sCurrentFunction) then
               nCurSel = ii: exit for
            end if   
         next
         ' Didn't find an exact function name match so exit out with just 
         ' the parent document match instead.
         if nCurSel = -1 then nCurSel = i
         Exit For
      END IF
   Next
   
   Listbox_SetCurSel(hList, nCurSel)
   
   Function = 0
End Function


' ========================================================================================
' Search child nodes to see if search text exists. Needed to allow output of parent node.
' ========================================================================================
private function frmFunctionList_SearchChildNodes( byval hParent as HTREEITEM, _
                                                   byval hChild as HTREEITEM, _
                                                   byref wszSearch as wstring _
                                                   ) as Boolean
   dim as Boolean bFound = true
   
   ' If a search term exists then only add the header if a valid found
   ' sub/function matches in one of the child nodes.
   if len(wszSearch) then
      Dim As Hwnd hTree = GetDlgItem(HWND_FRMFUNCTIONLIST, IDC_FRMFUNCTIONLIST_TREE)
      dim wszText as wstring * MAX_PATH
      hChild = Treeview_GetChild( hTree, hParent)
      bFound = false
      do until hChild = 0 
         Treeview_GetItemText(hTree, hChild, @wszText, MAX_PATH)
         if instr(ucase(wszText), ucase(wszSearch)) then 
            bFound = true: exit do
         end if
         hChild = Treeview_GetNextSibling(hTree, hChild)
      loop
   end if
   function = bFound
end function

' ========================================================================================
' Update the Listbox with the list of functions from all loaded documents
' ========================================================================================
private Function frmFunctionList_UpdateListBox() As Long

   ' Check for a valid HWND because this function will be called during a File Save
   ' and it is entirely possible that the FunctionList is not displayed. Screen
   ' flicker will otherwise result.
   if IsWindow(HWND_FRMFUNCTIONLIST) = 0 then exit function
   Dim As HWnd hList = GetDlgItem(HWND_FRMFUNCTIONLIST, IDC_FRMFUNCTIONLIST_LISTBOX)
   
   ' Clear the listbox and add the new items
   Listbox_ResetContent(hList)
   
   ' We get the list of files from the Project Explorer treeview rather
   ' than going through all pDoc documents
   Dim As Hwnd hTree = GetDlgItem(HWND_FRMFUNCTIONLIST, IDC_FRMFUNCTIONLIST_TREE)
   Dim As HTREEITEM hChild, hChildSearch, hParent
   dim as long p, pDoc
   dim as Boolean bFound
   dim wszText as wstring * MAX_PATH
   dim as CWSTR wszSearch = AfxGetWindowText(GetDlgItem(HWND_FRMFUNCTIONLIST, IDC_FRMFUNCTIONLIST_TXTSEARCH))
   
   SetWindowRedraw(hList, false)
   
   redim hRootNodes(any) as HTREEITEM
   
   if gApp.IsProjectActive then
      redim hRootNodes(4) as HTREEITEM
      hRootNodes(0) = gApp.hExplorerNormalNode        
      hRootNodes(1) = gApp.hExplorerMainNode          
      hRootNodes(2) = gApp.hExplorerResourceNode      
      hRootNodes(3) = gApp.hExplorerHeaderNode        
      hRootNodes(4) = gApp.hExplorerModuleNode        
   else   
      redim hRootNodes(0) as HTREEITEM
      hRootNodes(0) = gApp.hExplorerFilesNode
   end if
   
   for ii as long = lbound(hRootNodes) to ubound(hRootNodes)
      Treeview_SortChildren( hTree, hRootNodes(ii), 0 )
      hParent = Treeview_GetChild( hTree, hRootNodes(ii))
      do until hParent = 0 
         ' Filename
         Treeview_GetItemText(hTree, hParent, @wszText, MAX_PATH)
         pDoc = TreeView_GetlParam(hTree, hParent)
         wszText = ucase(wszText)

         bFound = frmFunctionList_SearchChildNodes(hParent, hChild, wszSearch) 
         if bFound then   
            p = Listbox_AddString( hList, strptr(wszText) )
            ' Treeview node item is stored in ListBox ItemData
            ListBox_SetItemData( hList, p, hParent )    
         end if
         
         ' Retrieve all functions in the file (child nodes)
         Treeview_SortChildren( hTree, hParent, 0 )
         hChild = Treeview_GetChild( hTree, hParent)
         do until hChild = 0 
            bFound = true
            Treeview_GetItemText(hTree, hChild, @wszText, MAX_PATH)
            if len(wszSearch) then
               bFound = instr(ucase(wszText), ucase(wszSearch))
            end if   
            if bFound then      
               p = Listbox_AddString( hList, strptr(wszText) )
               ListBox_SetItemData( hList, p, hChild )  ' Treeview node item is stored in ListBox ItemData
            end if
            hChild = Treeview_GetNextSibling(hTree, hChild)
         loop
         
         ' Get the next filename
         hParent = Treeview_GetNextSibling(hTree, hParent)
      loop 
   
   next
   
   SetWindowRedraw(hList, true)
   AfxRedrawWindow(hList)

   Function = 0
End Function


' ========================================================================================
' Select the listbox item that matches the incoming HTREE item. Do not open nodes to
' find a possible hidden document because the user may have purposely closed a node
' and we should not re-open it automatically.
' ========================================================================================
public function frmFunctionList_SelectItemData( byval hItem as HTREEITEM ) as boolean
   
   ' Select the Explorer listbox item where the ItemData holds the HTREE handle
   ' to the hidden FunctionList treeview. Search the listbox to match the hNodeExplorer
   ' value and the set that listbox item as current selection.
   dim as long idx = -1
   for i as long = 0 to ListBox_GetCount( HWND_FRMFUNCTIONLIST_LISTBOX ) - 1 
      if ListBox_GetItemData(HWND_FRMFUNCTIONLIST_LISTBOX, i) = hItem then
         idx = i: exit for
      end if   
   next 
   ListBox_SetCurSel( HWND_FRMFUNCTIONLIST_LISTBOX, idx )
   AfxRedrawWindow( HWND_FRMFUNCTIONLIST_LISTBOX )
   
   if idx >= 0 then function = true   
end function


' ========================================================================================
' Position all child windows. Called manually and/or by WM_SIZE
' ========================================================================================
private Function frmFunctionList_PositionWindows() As LRESULT

   Dim pWindow As CWindow Ptr = AfxCWindowPtr(HWND_FRMFUNCTIONLIST)
   If pWindow = 0 Then Exit Function
   
   ' Get the entire client area
   Dim As Rect rc
   GetClientRect(HWND_FRMFUNCTIONLIST, @rc)

   ' The Explorer label was painted in the top area so start our listbox immediately after that

   ' show/hide of the listbox is done separately in order to prevent blank client area from 
   ' painting before the lists are loaded. Allow a 2 pixel right border between listbox and
   ' form edge in order to allow for panel resizing. 
   
   dim as long nLeft = 0
   dim as long nTop = rc.top + pWindow->ScaleY(46)
   dim as long nWidth = rc.right  - rc.Left - pWindow->ScaleX(2)
   dim as long nHeight = rc.bottom - rc.top - nTop

   SetWindowPos( HWND_FRMFUNCTIONLIST_LISTBOX, 0, nLeft, nTop, nWidth, nHeight, SWP_NOZORDER )

   Function = 0
End Function


' ========================================================================================
' Process WM_SIZE message for window/dialog: frmFunctionList
' ========================================================================================
private Function frmFunctionList_OnSize( ByVal HWnd As HWnd, _
                                         ByVal state As UINT, _
                                         ByVal cx As Long, _
                                         ByVal cy As Long _
                                         ) As LRESULT
   If state <> SIZE_MINIMIZED Then
      ' Position all of the child windows
      frmFunctionList_PositionWindows
   End If
   Function = 0
End Function
       

' ========================================================================================
' Process WM_PAINT message for window/dialog: frmFunctionList
' ========================================================================================
private Function frmFunctionList_OnPaint( ByVal HWnd As HWnd ) As LRESULT
            
   Dim As PAINTSTRUCT ps
   Dim As HDC hDc
   
   hDC = BeginPaint(hWnd, @ps)
   FillRect( hDC, @ps.rcPaint, ghFunctionList.hPanelBrush )
   EndPaint hWnd, @ps
   
   Function = 0
End Function


' ========================================================================================
' Process WM_MOUSEMOVE message for window/dialog: frmFunctionList
' ========================================================================================
private Function frmFunctionList_OnMouseMove( ByVal HWnd As HWnd, _
                                              ByVal x as long, _
                                              byval y as long, _
                                              byval keyflags as UINT _
                                              ) As Long

   DIM pWindow AS CWindow PTR = AfxCWindowPtr(HWND_FRMFUNCTIONLIST)
   If pWindow = 0 Then Exit Function

   ' These are relative to <hWnd>
   dim as POINT pt = (x, y)
  
   ' PANELS LEFT/RIGHT SPLITTER
   Dim As Rect rc
   GetWindowRect( HWND_FRMFUNCTIONLIST, @rc )
   rc.Left = rc.Right - pWindow->ScaleX(3)
   GetCursorPos(@pt)
   If PtInRect( @rc, pt ) Then
      if WindowFromPoint(pt) = HWND_FRMFUNCTIONLIST then
         SetCursor LoadImage( Null, MAKEINTRESOURCEW(OCR_SIZEWE), IMAGE_CURSOR, 0, 0, LR_DEFAULTSIZE Or LR_SHARED )
      end if
   End If

   If gApp.bDragActive Then
      If gApp.hWndPanel = HWND_FRMFUNCTIONLIST Then
         ' Get the current rect of the frmFunctionList and compare right side to the current cursor position
         ' and then move the right side to equal the cursor position.
         Dim As Long nDiff = pt.x - rc.Right
         GetClientRect HWND_FRMFUNCTIONLIST, @rc
         Dim As Long nWidth = (rc.Right - rc.Left + nDiff)
         ' make sure width does not go below a specific size otherwise the user will not
         ' be able to grab the panel edge to resize
         nWidth = max( nWidth, pWindow->ScaleX(80) )
         SetWindowPos HWND_FRMFUNCTIONLIST, 0, 0, 0, nWidth, rc.Bottom - rc.Top, SWP_NOMOVE Or SWP_NOZORDER
         frmMain_PositionWindows
         AfxDoEvents   ' allow screen to repaint
         Exit Function
      End If
   End If
      
   function = 0
end function


' ========================================================================================
' Process WM_LBUTTONDOWN message for window/dialog: frmFunctionList
' ========================================================================================
private Function frmFunctionList_OnLButtonDown( ByVal HWnd As HWnd, _
                                                byval fDoubleClick as Boolean, _
                                                ByVal x as long, _
                                                byval y as long, _
                                                byval keyflags as UINT _
                                                ) As Long

   DIM pWindow AS CWindow PTR = AfxCWindowPtr(HWND_FRMFUNCTIONLIST)
   If pWindow = 0 Then Exit Function

   ' PANELS LEFT/RIGHT SPLITTER
   Dim As Rect rc
   Dim As Point pt 

   gApp.bDragActive = False 
   GetWindowRect HWND_FRMFUNCTIONLIST, @rc
   rc.Left = rc.Right - pWindow->ScaleX(3)
   GetCursorPos(@pt)
   If PtInRect( @rc, pt ) Then
      if WindowFromPoint(pt) = HWND_FRMFUNCTIONLIST then
         SetCursor LoadImage( Null, MAKEINTRESOURCEW(OCR_SIZEWE), IMAGE_CURSOR, 0, 0, LR_DEFAULTSIZE Or LR_SHARED )
         gApp.bDragActive = True 
         gApp.hWndPanel   = HWND_FRMFUNCTIONLIST
         SetCapture( HWND_FRMFUNCTIONLIST )
      end if
   End If

   function = 0
end function


' ========================================================================================
' Process WM_MEASUREITEM message for window/dialog: frmFunctionList
' ========================================================================================
private Function frmFunctionList_OnMeasureItem( ByVal HWnd As HWnd, _
                                                ByVal lpmis As MEASUREITEMSTRUCT Ptr _
                                                ) As Long
   ' Set the height of the list box items. 
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(HWnd)
   lpmis->itemHeight = pWindow->ScaleY(FUNCTIONLISTITEM_HEIGHT)
      
   Function = 0
End Function


' ========================================================================================
' Process WM_DRAWITEM message for window/dialog: frmFunctionList
' ========================================================================================
private Function frmFunctionList_OnDrawItem( ByVal HWnd As HWnd, _
                                             ByVal lpdis As Const DRAWITEMSTRUCT Ptr _
                                             ) As Long

   DIM pWindow AS CWindow PTR = AfxCWindowPtr(HWND_FRMFUNCTIONLIST)
   If pWindow = 0 Then Exit Function

   if lpdis = 0 then exit function
 
   if ( lpdis->itemAction = ODA_DRAWENTIRE ) orelse _
      ( lpdis->itemAction = ODA_SELECT ) orelse _
      ( lpdis->itemAction = ODA_FOCUS ) then

      if lpdis->itemAction = ODA_FOCUS then return true

      dim as RECT rc = lpdis->rcItem
      dim as long nWidth  = rc.right-rc.left
      dim as long nHeight = rc.bottom-rc.top 

      SaveDC(lpdis->hDC)

      Dim memDC as HDC      ' Double buffering
      Dim hbit As HBITMAP   ' Double buffering

      memDC = CreateCompatibleDC( lpdis->hDC )
      hbit  = CreateCompatibleBitmap( lpdis->hDC, nWidth, nHeight )
      If hbit Then hbit = SelectObject( memDC, hbit )

      SelectObject( memDC, ghMenuBar.hFontMenuBar )
         
      ' Default to using normal
      dim as HBRUSH hBrush = ghFunctionList.hBackBrush
      dim as COLORREF foreclr = ghFunctionList.ForeColor
      dim as COLORREF backclr = ghFunctionList.BackColor
         
      dim as boolean IsHot = false
      dim as boolean isNodeHeader = false
      dim as boolean isIconDown = false
         
      dim as POINT pt
      GetCursorPos( @pt )
      MapWindowPoints( lpdis->hwndItem, HWND_DESKTOP, cast( POINT ptr, @rc ), 2 )
      if PtInRect( @rc, pt ) then IsHot = true
      
      ' If mouse is over VScrollBar then reset hot
      if isMouseOverWindow( HWND_FRMEXPLORER_VSCROLLBAR ) then IsHot = false
         
      if ListBox_GetCurSel(lpdis->hwndItem) = lpdis->itemID then IsHot = true

      hBrush = iif( IsHot, ghFunctionList.hBackBrushHot, ghFunctionList.hBackBrush)
      backclr = iif( IsHot, ghFunctionList.BackColorHot, ghFunctionList.BackColor)
      foreclr = iif( IsHot, ghFunctionList.ForeColorHot, ghFunctionList.ForeColor)

      dim as CWSTR wszCaption = AfxGetListBoxText(lpdis->hwndItem, lpdis->ItemID)

      ' If this is a "node" header then use those colors
      if left(wszCaption, 1) = "%" then isNodeHeader = true
      if mid(wszCaption, 4, 1) = "1" then isIconDown = true
      
      if isNodeHeader then
         hBrush = iif( IsHot, ghFunctionList.hBackBrushHot, ghFunctionList.hBackBrushNode)
         backclr = iif( IsHot, ghFunctionList.BackColorHot, ghFunctionList.BackColorNode)
         foreclr = iif( IsHot, ghFunctionList.ForeColorNode, ghFunctionList.ForeColorNode)
      end if
         
      ' Paint the entire background
      ' Create our rect that works with the entire line
      SetRect( @rc, 0, 0, nWidth, nHeight )
      FillRect( memDC, @rc, hBrush )

      SetBkColor( memDC, backclr )   
      SetTextColor( memDC, foreclr )

      dim as RECT rcText = rc
      dim as RECT rcBitmap = rc

      dim as long wsStyle 
      
      ' indent the text based on its type
      if isNodeHeader then
         rcBitmap.right = rcBitmap.left + pWindow->ScaleX(20)
         SelectObject( memDC, ghMenuBar.hFontSymbol )
         wsStyle = DT_NOPREFIX or DT_CENTER or DT_TOP or DT_SINGLELINE  
         if isIconDown then 
            DrawText( memDC, wszChevronDown, -1, Cast(lpRect, @rcBitmap), wsStyle )
         else   
            DrawText( memDC, wszChevronRight, -1, Cast(lpRect, @rcBitmap), wsStyle )
         end if
         wszCaption = mid(wszCaption, 6)
         rcText.left = rcBitmap.right
         SelectObject( memDC, ghMenuBar.hFontMenuBar )
         wsStyle = DT_NOPREFIX or DT_LEFT or DT_VCENTER or DT_SINGLELINE  
         DrawText( memDC, wszCaption.sptr, -1, Cast(lpRect, @rcText), wsStyle )
      else
         ' This would be a regular file.
         rcBitmap.left = rcText.left + pWindow->ScaleX(20)
         rcBitmap.right = rcBitmap.left + pWindow->ScaleX(20) 

         SelectObject( memDC, ghMenuBar.hFontSymbol )
         wsStyle = DT_NOPREFIX or DT_CENTER or DT_TOP or DT_SINGLELINE  
         DrawText( memDC, wszDocumentIcon, -1, Cast(lpRect, @rcBitmap), wsStyle )

         rcText.left = rcBitmap.right
         SelectObject( memDC, ghMenuBar.hFontMenuBar )
         wsStyle = DT_NOPREFIX or DT_LEFT or DT_VCENTER or DT_SINGLELINE  
         DrawText( memDC, wszCaption.sptr, -1, Cast(lpRect, @rcText), wsStyle )
      end if
         
      BitBlt lpdis->hDC, lpdis->rcItem.left, lpdis->rcItem.top, _
             nWidth, nHeight, memDC, 0, 0, SRCCOPY 

      ' Cleanup
      If hbit  Then DeleteObject SelectObject(memDC, hbit)
      If memDC Then DeleteDC memDC
      RestoreDC(lpdis->hDC, -1)
   end if   

   Function = True
   
end function


' ========================================================================================
' Process WM_LBUTTONUP message for window/dialog: frmFunctionList
' ========================================================================================
private Function frmFunctionList_OnLButtonUp( ByVal HWnd As HWnd, _
                                              ByVal x as long, _
                                              byval y as long, _
                                              byval keyflags as UINT _
                                              ) As Long
  
   if gApp.bDragActive then
      gApp.bDragActive = False 
      gApp.hWndPanel = 0
      ReleaseCapture()
   end if
   
   function = 0
end function


' ========================================================================================
' Process WM_COMMAND message for window/dialog: frmFunctionList
' ========================================================================================
private Function frmFunctionList_OnCommand( ByVal HWnd As HWnd, _
                                            ByVal id As Long, _
                                            ByVal hwndCtl As HWnd, _
                                            ByVal codeNotify As UINT _
                                            ) As LRESULT
   select case codeNotify
      case LBN_SELCHANGE
         ' update the highlighting of the current line
         AfxRedrawWindow(hwndCtl)
         ' update the scrollbar position if necessary
         frmFunctionList_PositionWindows()
   end select      

   function = 0
end function


' ========================================================================================
' frmFunctionList Window procedure
' ========================================================================================
private Function frmFunctionList_WndProc( ByVal HWnd   As HWnd, _
                                          ByVal uMsg   As UINT, _
                                          ByVal wParam As WPARAM, _
                                          ByVal lParam As LPARAM _
                                          ) As LRESULT

   Select Case uMsg
      HANDLE_MSG (HWnd, WM_SIZE,        frmFunctionList_OnSize)
      HANDLE_MSG (HWnd, WM_PAINT,       frmFunctionList_OnPaint)
      HANDLE_MSG (HWnd, WM_COMMAND,     frmFunctionList_OnCommand)
      HANDLE_MSG (HWnd, WM_LBUTTONDOWN, frmFunctionList_OnLButtonDown)
      HANDLE_MSG (HWnd, WM_LBUTTONUP,   frmFunctionList_OnLButtonUp)
      HANDLE_MSG (HWnd, WM_MOUSEMOVE,   frmFunctionList_OnMouseMove)
      HANDLE_MSG (HWnd, WM_MEASUREITEM, frmFunctionList_OnMeasureItem)
      HANDLE_MSG (HWnd, WM_DRAWITEM,    frmFunctionList_OnDrawItem)
   
   case WM_ERASEBKGND
      return true

   case WM_CTLCOLORSTATIC
      ' Foreground and Background colors for the Explorer label
      SetTextColor Cast(HDC, wParam), ghFunctionList.forecolor
      SetBkColor Cast(HDC, wParam), ghFunctionList.backcolor
      return Cast( LRESULT, ghFunctionList.hPanelBrush)

   End Select

   ' for messages that we don't deal with
   Function = DefWindowProc( HWnd, uMsg, wParam, lParam )

End Function


' ========================================================================================
' TrackMouse. setup TrackMouseEvent for a specififc window.
' ========================================================================================
function frmFunctionList_TrackMouse( byval hWin as HWND ) as long
   Dim tme As TrackMouseEvent
   tme.cbSize = Sizeof(TrackMouseEvent)
   tme.dwFlags = TME_HOVER Or TME_LEAVE
   tme.hwndTrack = hWin
   TrackMouseEvent(@tme) 
   function = 0
end function         

' ========================================================================================
' frmFunctionListListBox_SubclassProc 
' ========================================================================================
Function frmFunctionListListBox_SubclassProc ( _
                  ByVal hWin   As HWnd, _                 ' // Control window handle
                  ByVal uMsg   As UINT, _                 ' // Type of message
                  ByVal _wParam As WPARAM, _               ' // First message parameter
                  ByVal _lParam As LPARAM, _               ' // Second message parameter
                  ByVal uIdSubclass As UINT_PTR, _        ' // The subclass ID
                  ByVal dwRefData As DWORD_PTR _          ' // Pointer to reference data
                  ) As LRESULT

   DIM pWindow AS CWindow PTR = AfxCWindowPtr(hWin)
   static as long accumDelta
   
   ' keep track of last index we were over so that we only issue a 
   ' repaint if the cursor has moved off of the line
   static as long nLastIdx = -1
   static as boolean isTracking = false
   
   Select Case uMsg
      case MSG_USER_LOAD_FUNCTIONLISTFILES
         LoadFunctionListFiles()

      Case WM_MOUSEWHEEL
         ' accumulate delta until scroll one line (up +120, down -120). 
         ' 120 is the Microsoft default delta
         dim as long zDelta = GET_WHEEL_DELTA_WPARAM( _wParam )
         dim as long nTopIndex = SendMessage( hWin, LB_GETTOPINDEX, 0, 0 ) 
         accumDelta = accumDelta + zDelta
         if accumDelta >= 120 then       ' scroll up 3 lines
            nTopIndex = nTopIndex - 3
            SendMessage( hWin, LB_SETTOPINDEX, nTopIndex, 0 ) 
            accumDelta = 0
            frmExplorerVScroll_PositionWindows( SW_SHOWNA )
         elseif accumDelta <= -120 then  ' scroll down 3 lines
            nTopIndex = nTopIndex + 3
            SendMessage( hWin, LB_SETTOPINDEX, nTopIndex, 0 ) 
            accumDelta = 0
            frmExplorerVScroll_PositionWindows( SW_SHOWNA )
         end if
         
      Case WM_MOUSEMOVE
         ' Track that we are over the control in order to catch the 
         ' eventual WM_MOUSELEAVE event
         if not isTracking then
            frmFunctionList_TrackMouse( hWin )
            isTracking = true
         end if
         
         ' get the item rect that the mouse is over and only invalidate
         ' that instead of the entire listbox
         dim as RECT rc
         dim as long idx = Listbox_ItemFromPoint( hWin, GET_X_LPARAM(_lParam), GET_Y_LPARAM(_lParam)) 
         if idx <> nLastIdx then
            ListBox_GetItemRect( hWin, idx, @rc )
            InvalidateRect( hWin, @rc, true )
            ListBox_GetItemRect( hWin, nLastIdx, @rc )
            InvalidateRect( hWin, @rc, true )
            nLastIdx = idx
         end if
                     
      case WM_MOUSELEAVE
         isTracking = false   
         nLastIdx = -1
         AfxRedrawWindow(hWin)   
         
      case WM_RBUTTONDOWN
         ' call WM_LBUTTONDOWN to set the correct node/tab
         SendMessage( hWin, WM_LBUTTONDOWN, _wParam, _lParam )
         ' Create the popup menu
         Dim As Hwnd hTree = GetDlgItem(HWND_FRMFUNCTIONLIST, IDC_FRMFUNCTIONLIST_TREE)
         dim as long idx = Listbox_ItemFromPoint( hWin, GET_X_LPARAM(_lParam), GET_Y_LPARAM(_lParam)) 
         dim as long p = ListBox_GetItemData( hWin, idx )
         dim as HTREEITEM hItem = cast( HTREEITEM, cast(integer, p) )
         ' we need to actualy select the treeview item in the hidden functionlist treeview because
         ' the IDM_REMOVEFILEFROMPROJECT_EXPLORERLISTBOX message will examine the current selected
         ' node to determine the necessary action to take.
         TreeView_SelectItem(hTree, hItem)
         dim as clsDocument ptr pDoc = cast(clsDocument ptr, TreeView_GetlParam(hTree, hItem))
         dim as HMENU hPopupMenu = CreateExplorerContextMenu(pDoc)
         Dim As POINT pt: GetCursorPos @pt
         TrackPopupMenu(hPopUpMenu, 0, pt.x, pt.y, 0, HWND_FRMMAIN, ByVal Null)
         DestroyMenu hPopUpMenu
         Return True   ' prevent further processing that leads to WM_CONTEXTMENU

      case WM_LBUTTONDOWN
         ' determine if we clicked on a regular file or a node header
         dim as RECT rc
         dim as long idx = Listbox_ItemFromPoint( hWin, GET_X_LPARAM(_lParam), GET_Y_LPARAM(_lParam)) 
         dim as CWSTR wszCaption = AfxGetListBoxText( hWin, idx )
         if left(wszCaption, 1) = "%" then
            ' Toggle the show/hide of files under this node
            dim as long idxArray = val(mid(wszCaption, 2, 1))
            gExplorerNodeShow(idxArray) = not gExplorerNodeShow(idxArray)
            ' allow listboc click event to fully process before loading new explorer files
            ' so that we can correctly select the current item.
            PostMessage( hWin, MSG_USER_LOAD_FUNCTIONLISTFILES, 0, 0 )
         end if
         
      case WM_LBUTTONDBLCLK 
         ' determine if we clicked on a regular file or a node header. If it is a regular
         ' file then load it.
         dim as RECT rc
         dim as long idx = Listbox_ItemFromPoint( hWin, GET_X_LPARAM(_lParam), GET_Y_LPARAM(_lParam)) 
         dim as CWSTR wszCaption = AfxGetListBoxText( hWin, idx )
         if left(wszCaption, 1) <> "%" then
            dim as CWSTR wszFilename
            dim as long p = ListBox_GetItemData( hWin, idx )
            dim as HTREEITEM hItem = cast( HTREEITEM, cast(integer, p) )
            frmFunctionList_GetFileNameFunctionName( hItem, wszFilename, "" )
            OpenSelectedDocument( wszFilename, "" )
         end if
         
      case WM_ERASEBKGND
         ' If the number of lines in the listbox maybe less than the number per page then 
         ' calculate from last item to bottom of listbox, otherwise calculate based on
         ' the mod of the lineheight to listbox height so we can color the partial line
         ' that won't be displayed at the bottom of the list.
         dim as RECT rc: GetClientRect( hWin, @rc )

         dim as RECT rcItem  
         SendMessage( hWin, LB_GETITEMRECT, 0, cast(LPARAM, @rcItem) )
         dim as long itemHeight = rcItem.bottom - rcItem.top
         dim as long NumItems = ListBox_GetCount(hWin)
         dim as long nTopIndex = SendMessage( hWin, LB_GETTOPINDEX, 0, 0 ) 
         dim as long visible_rows = 0
         dim as long ItemsPerPage = 0
         dim as long bottom_index = 0
                     
         if NumItems > 0 then
            ItemsPerPage = (rc.bottom - rc.top) / itemHeight
            bottom_index = (nTopIndex + ItemsPerPage) 
            if bottom_index >= NumItems then bottom_index = NumItems - 1
            visible_rows = (bottom_index - nTopIndex) + 1
            rc.top = visible_rows * itemHeight 
         end if

         if rc.top < rc.bottom then
            dim as HDC _hDC = cast(HDC, _wParam)
            FillRect( _hDC, @rc, ghFunctionList.hPanelBrush )
         end if

         ValidateRect( hWin, @rc )
         return true
      
      Case WM_DESTROY
         ' REQUIRED: Remove control subclassing
         RemoveWindowSubclass( hWin, @frmFunctionListListBox_SubclassProc, uIdSubclass )
   End Select
    
   ' For messages that we don't deal with
   Function = DefSubclassProc(hWin, uMsg, _wParam, _lParam)

End Function


' ========================================================================================
' frmFunctionList_Show
' ========================================================================================
public Function frmFunctionList_Show( ByVal hWndParent As HWnd ) As LRESULT

   '  Create the main window and child controls
   Dim pWindow As CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowPtr(hwndParent)->DPI

   dim as long nWidth = iif(gConfig.ShowExplorer, gConfig.ShowExplorerWidth, 250)
   
   HWND_FRMFUNCTIONLIST = pWindow->Create( hWndParent, "Function List", _
        @frmFunctionList_WndProc, 0, 0, nWidth, 0, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, _
        WS_EX_CONTROLPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)

   ' Disable background erasing by only assigning the one style
   pWindow->ClassStyle = CS_DBLCLKS
   pWindow->SetFont( "Segoe UI Semibold", 9)

   dim as CWSTR wszText = ucase(L(223,"Function List"))
   Dim As HWnd hBtn = _
   pWindow->AddControl("LABEL", , IDC_FRMFUNCTIONLIST_LBLFUNCTIONLIST, wszText, 10, 12, 75, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT, _
        WS_EX_LEFT Or WS_EX_LTRREADING)

   pWindow->Brush = ghFunctionList.hPanelBrush

   HWND_FRMFUNCTIONLIST_LISTBOX = _ 
        pWindow->AddControl("LISTBOX", , IDC_FRMFUNCTIONLIST_LISTBOX, "", 0, 0, 0, 0, _
        WS_CHILD Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or WS_TABSTOP Or LBS_NOINTEGRALHEIGHT or _   
        LBS_OWNERDRAWFIXED Or LBS_HASSTRINGS Or LBS_NOTIFY, WS_EX_LEFT Or WS_EX_RIGHTSCROLLBAR, , _
        Cast(SUBCLASSPROC, @frmFunctionListListBox_SubclassProc), IDC_FRMFUNCTIONLIST_LISTBOX, Cast(DWORD_PTR, @pWindow))

   ' Basically a hidden treeview that we use to manage our nodes and subnodes and to do the
   ' required sorting for us.
   Dim As HWnd hTree = _
        pWindow->AddControl("TREEVIEW", , IDC_FRMFUNCTIONLIST_TREE, "", 0, 0, 0, 0, _
        WS_CHILD Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN or TVS_HASBUTTONS, 0 )

   ' Create the special functionlist treeview nodes 
   frmFunctionList_CreateSpecialNodes()

   frmFunctionList_PositionWindows

   Function = 0
   
End Function


' ========================================================================================
' LoadFunctionListFiles
' This will clear the current list of files in the Explorer and repopulate it
' with the latest list of filenames that are stored in the hidden FunctionList treeview.
' Files are excluded from being added based on the state of the global array gExplorerNodeShow()
' ========================================================================================
public function LoadFunctionListFiles() as long
   dim as HWND hList = GetDlgItem(HWND_FRMFUNCTIONLIST, IDC_FRMFUNCTIONLIST_LISTBOX)
   Dim As Hwnd hTree = GetDlgItem(HWND_FRMFUNCTIONLIST, IDC_FRMFUNCTIONLIST_TREE)
    
   ' Hide the listbox while it is loading so that we don't get the unpainted
   ' white background from the empty listbox
   ShowWindow( hList, SW_HIDE )
   
   ' Save the topindex because we will restore it after filling the new contents 
   dim as long nTopIndex = SendMessage( hList, LB_GETTOPINDEX, 0, 0 ) 

   Dim As HTREEITEM hCurrentSelection
   Dim pDoc As clsDocument Ptr = gTTabCtl.GetActiveDocumentPtr()
   If pDoc Then hCurrentSelection = pDoc->hNodeExplorer 
   
   dim as long idxRestore = -1

   ' Clear all content from the listbox
   ListBox_ResetContent(hList)
   
   ' Add the list of filenames (allow the treeview to sort the data for us)
   Dim As HTREEITEM hParent
   dim as long p
   dim wszText as wstring * MAX_PATH

   redim hRootNodes(any) as HTREEITEM
   
'#define NODETYPE_MAIN              0
'#define NODETYPE_RESOURCE          1
'#define NODETYPE_HEADER            2
'#define NODETYPE_MODULE            3
'#define NODETYPE_NORMAL            4

   if gApp.IsProjectActive then
      redim hRootNodes(4) as HTREEITEM
      hRootNodes(0) = gApp.hExplorerMainNode          
      hRootNodes(1) = gApp.hExplorerResourceNode      
      hRootNodes(2) = gApp.hExplorerHeaderNode        
      hRootNodes(3) = gApp.hExplorerModuleNode        
      hRootNodes(4) = gApp.hExplorerNormalNode        
   else   
      redim hRootNodes(0) as HTREEITEM
      hRootNodes(0) = gApp.hExplorerFilesNode
   end if
   
   for ii as long = lbound(hRootNodes) to ubound(hRootNodes)
      ' The node header description starts with a % in order for the
      ' drawing routine to know that it is a node header and not a
      ' regular file.
      ' e.g.
      '       %0%1%Main
      '       %1%1%Resource
      '       %2%0%Header
      '       %3%0%Module
      '       %4%1%Normal
      ' The value after the % is the array index into gExplorerNodeShow
      ' that allows the mouse routines to toggle on/off displaying of
      ' the files. The value after the second % indicates to the drawing 
      ' routine that files exist ("1") under the node, or ("0") that 
      ' no files exist. This wil then dictate what icon to show next
      ' to the line.
      
      wszText = "%" & ii & "%"
      if gExplorerNodeShow(ii) = false then 
         wszText = wszText & "0" & "%"
      else   
         wszText = wszText & "1" & "%"
      end if

      select case ii
         case NODETYPE_MAIN
            if gApp.IsProjectActive then
               wszText = wszText & L(212, "Main")
            else   
               wszText = wszText & L(2, "File")
            end if   
         case NODETYPE_RESOURCE: wszText = wszText & L(213, "Resource")
         case NODETYPE_HEADER:   wszText = wszText & L(175, "Header")
         case NODETYPE_MODULE:   wszText = wszText & L(211, "Module")
         case NODETYPE_NORMAL:   wszText = wszText & L(210, "Normal")
      end select
      
      p = Listbox_AddString( hList, @wszText )

      ' Check the global array gExplorerNodeShow to see if these types
      ' of files should be loaded into the listbox.
      if gExplorerNodeShow(ii) = false then continue for
      
      Treeview_SortChildren( hTree, hRootNodes(ii), 0 )
      hParent = Treeview_GetChild( hTree, hRootNodes(ii))
      do until hParent = 0 
         ' Filename
         Treeview_GetItemText(hTree, hParent, @wszText, MAX_PATH)
         dim as clsDocument ptr pDoc = cast(clsDocument ptr, TreeView_GetlParam(hTree, hParent))
         p = Listbox_AddString( hList, @wszText )
         ' Treeview node item is stored in ListBox ItemData
         ListBox_SetItemData( hList, p, hParent )    
         if pDoc then pDoc->hNodeExplorer = hParent
         if hParent = hCurrentSelection then idxRestore = p
         ' Get the next filename
         hParent = Treeview_GetNextSibling(hTree, hParent)
      loop 
   
   next

   ' Restore the top index so the list displays like it did before being reset
   SendMessage( hList, LB_SETTOPINDEX, nTopIndex, 0 ) 

   ' Ensure that Listbox is now properly sized and then show 
   ' the listbox now that it is fully populated
   ShowWindow( hList, SW_SHOW )
   frmFunctionList_PositionWindows()

   ' Set current selection to the item that was selected  prior
   ' to us reloading the listbox
   ListBox_SetCurSel( hList, idxRestore )
   AfxRedrawWindow(hList)

   ' Determine if the VScroll bar has changed size or is now hidden/shown
   frmExplorerVScroll_PositionWindows( SW_HIDE )

   function = 0
end function

   
 

' ========================================================================================
' ShowFunctionList
' This is the function that the user calls to actually show/hide the function list
' Function returns True if the window is shown
' ========================================================================================
public function ShowFunctionList() as boolean

   if IsWindow(HWND_FRMFUNCTIONLIST) = 0 then
      frmFunctionList_Show(HWND_FRMMAIN)
   else
      ' Calling this function works as a toggle to display/close the popup form.
      if IsWindowVisible(HWND_FRMFUNCTIONLIST) then
         ShowWindow(HWND_FRMFUNCTIONLIST, SW_HIDE)
         exit function
      end if
   end if

   frmFunctionList_ReparseFiles()
   frmFunctionList_UpdateListBox()
   frmFunctionList_SetListBoxPosition()

   AfxCenterWindow( HWND_FRMFUNCTIONLIST, HWND_FRMMAIN )
   ShowWindow HWND_FRMFUNCTIONLIST, SW_SHOW
   SetFocus GetDlgItem(HWND_FRMFUNCTIONLIST, IDC_FRMFUNCTIONLIST_LISTBOX)
     
   function = true
end function
