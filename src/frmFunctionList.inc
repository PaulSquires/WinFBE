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

#include once "frmFunctionList.bi"
#include once "clsDocument.bi"
#include once "clsTopTabCtl.bi"


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
' Process WM_COMMAND message for window/dialog: FunctionList
' ========================================================================================
private Function frmFunctionList_OnCommand( ByVal HWnd As HWnd, _
                                            ByVal id As Long, _
                                            ByVal hwndCtl As HWnd, _
                                            ByVal codeNotify As UINT _
                                            ) As LRESULT

   static as long nPrevSel
   
   Select Case id
      Case IDC_FRMFUNCTIONLIST_LISTBOX
         if codeNotify = LBN_SETFOCUS then
            nPrevSel = ListBox_GetCurSel( hwndCtl )
         end if   
         
         if codeNotify = LBN_SELCHANGE then
            dim as long nCurSel = ListBox_GetCurSel( hwndCtl )
            dim as RECT rc
            ' only repaint the items that have changed to reduce lag
            ListBox_GetItemRect( hwndCtl, nPrevSel, @rc )
            InvalidateRect( hwndCtl, @rc, true )
            ListBox_GetItemRect( hwndCtl, nCurSel, @rc )
            InvalidateRect( hwndCtl, @rc, true )
            nPrevSel = nCurSel
         end if

         If codeNotify = LBN_DBLCLK Then
            Dim As Long nCurSel = Listbox_GetCurSel(hwndCtl)
            If nCurSel = -1 Then Exit Function

            If gConfig.CloseFuncList Then PostMessage( hWnd, WM_CLOSE, 0, 0 ) 

            ' The ItemData holds the Explorer Treeview node item handle
            dim as CWSTR wszFilename, wszFunctionName
            Dim As HTREEITEM hNode = cast(HTREEITEM, Listbox_GetItemData(hWndCtl, nCurSel))
            frmFunctionList_GetFileNameFunctionName( hNode, wszFilename, wszFunctionName )
            OpenSelectedDocument( wszFilename, wszFunctionName )
         end if
         
      Case IDC_FRMFUNCTIONLIST_TXTSEARCH
         If codeNotify = EN_CHANGE Then
            frmFunctionList_UpdateListBox
         end if      

   End Select
   
   Function = 0
End Function


' ========================================================================================
' Position all child windows. Called manually and/or by WM_SIZE
' ========================================================================================
private Function frmFunctionList_PositionWindows() As LRESULT
   
   ' Get the entire client area
   Dim As Rect rc
   GetClientRect(HWND_FRMFUNCTIONLIST, @rc)
   
   dim as HWND hTextBox = GetDlgItem(HWND_FRMFUNCTIONLIST, IDC_FRMFUNCTIONLIST_TXTSEARCH)
   dim as HWND hListBox = GetDlgItem(HWND_FRMFUNCTIONLIST, IDC_FRMFUNCTIONLIST_LISTBOX )
   
   dim as long nHeight = AfxGetWindowHeight(hTextBox)
   SetWindowPos( hTextBox, 0, 0, 0, rc.Right-rc.Left, nHeight, SWP_NOZORDER )
   SetWindowPos( hListBox, 0, 0, nHeight, rc.Right-rc.Left, rc.Bottom-rc.top-nHeight, SWP_NOZORDER )

   Function = 0
End Function


' ========================================================================================
' Process WM_SIZE message for window/dialog: FunctionList
' ========================================================================================
private Function frmFunctionList_OnSize( ByVal HWnd As HWnd, _
                                         ByVal state As UINT, _
                                         ByVal cx As Long, _
                                         ByVal cy As Long _
                                         ) As LRESULT
   If state <> SIZE_MINIMIZED Then 
      frmFunctionList_PositionWindows
   End If   
   Function = 0
End Function


' ========================================================================================
' Process WM_CLOSE message for window/dialog: FunctionList
' ========================================================================================
private Function frmFunctionList_OnClose( ByVal HWnd As HWnd ) As LRESULT
   ShowWindow(HWnd, SW_HIDE)
   Function = 0
End Function


' ========================================================================================
' Process WM_MEASUREITEM message for window/dialog: FunctionList
' ========================================================================================
private Function frmFunctionList_OnMeasureItem( ByVal HWnd As HWnd, _
                                                ByVal lpmis As MEASUREITEMSTRUCT Ptr _
                                                ) As Long
   ' Set the height of the List box items. 
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(HWnd)

   If lpmis->CtlId <> IDC_FRMFUNCTIONLIST_LISTBOX Then Exit Function
   lpmis->itemHeight = pWindow->ScaleY(18)
      
   Function = 0
End Function


' ========================================================================================
' Process WM_DRAWITEM message for window/dialog: FunctionList
' ========================================================================================
private function frmFunctionList_OnDrawItem( byval HWnd as HWND, _
                                             byval lpdis as const DRAWITEMSTRUCT ptr _
                                             ) as long

   Dim As HBRUSH hBrush 
   Dim As HFONT hFont, hFontOld
   dim as BOOLEAN bIsHeader
   Dim As Rect rc 
   Dim wzText As WString * MAX_PATH
   
   If lpdis->itemID = -1 Then Exit Function
   
   Dim pWindow As CWindow Ptr = AfxCWindowOwnerPtr(HWnd)
 
   Select Case lpdis->itemAction
      Case ODA_DRAWENTIRE, ODA_SELECT
   
         ' CLEAR BACKGROUND
         hBrush = GetSysColorBrush(COLOR_WINDOW) 
         SelectObject(lpdis->hDC, hBrush)      
         FillRect(lpdis->hDC, @lpdis->rcItem, hBrush)  
   
         ' DRAW TEXT
         SetBkColor(lpdis->hDC, GetSysColor(COLOR_WINDOW))   
         SetTextColor(lpdis->hDC, GetSysColor(COLOR_WINDOWTEXT))
         SendMessage(lpdis->hwndItem, LB_GETTEXT, lpdis->itemID, Cast(LPARAM, @wzText))
         
         ' Determine what font to use and if itès a header line then only show
         ' the filename's name and extension (not full path)
         dim as hwnd hTree = GetDlgItem(HWND_FRMFUNCTIONLIST, IDC_FRMFUNCTIONLIST_TREE)
         bIsHeader = iif( TreeView_GetlParam( hTree, cast(HTREEITEM, lpdis->itemData)), true, false)

         if bIsHeader then wzText = AfxStrPathname( "NAMEX", wzText )
         hFont = pWindow->CreateFont("", 9, iif(bIsHeader, FW_BOLD, FW_NORMAL) )

         hFontOld = SelectObject(lpdis->hDC, hFont)
         DrawTextW( lpdis->hDC, wzText, -1, Cast(lpRect, @lpdis->rcItem), _
                        DT_LEFT Or DT_SINGLELINE Or Iif(bIsHeader, DT_BOTTOM, DT_VCENTER) )
         DeleteObject(SelectObject(lpdis->hDC, hFontOld))
   
         ' SELECTED ITEM
         If (lpdis->itemState And ODS_SELECTED) Then     
            rc.Left   = lpdis->rcItem.Left + 1
            rc.Right  = lpdis->rcItem.Right 
            rc.top    = lpdis->rcItem.top
            rc.bottom = lpdis->rcItem.bottom
            InvertRect(lpdis->hDC, @rc) 
         End If
   
         Function = True : Exit Function
   
      Case ODA_FOCUS
         DrawFocusRect(lpdis->hDC, @lpdis->rcItem) 
   
   End Select
   
   function = true
end function


' ========================================================================================
' Only paint the areas of the multicolumn listbox that does not have items. This can
' be quite challenging for non-integral ownerdraw multicolumn listboxes but it does
' help prevent flicker. This function is called from WM_ERASEBKGND.
' ========================================================================================
private function frmFunctionList_PaintNonItemAreas( byval hDC as HDC ) as long
   dim as RECT rc: GetClientRect( HWND_FRMFUNCTIONLIST, @rc )
   DIM pWindow AS CWindow PTR = AfxCWindowPtr(HWND_FRMFUNCTIONLIST)
   If pWindow Then rc.top = rc.bottom - pWindow->ScaleY(18)
   FillRect( hDC, @rc, ghPanel.hPanelBrush )  
   function = 0
end function

   
' ========================================================================================
' Process WM_DESTROY message for window/dialog: FunctionList
' ========================================================================================
private Function frmFunctionList_OnDestroy( byval HWnd As HWnd ) As LRESULT
   HWND_FRMFUNCTIONLIST = 0
   Function = 0
End Function


' ========================================================================================
' Processes messages for the subclassed ListBox window.
' ========================================================================================
private Function frmFunctionList_ListBox_SubclassProc ( _
                  ByVal HWnd   As HWnd, _                 ' // Control window handle
                  ByVal uMsg   As UINT, _                 ' // Type of message
                  ByVal wParam As WPARAM, _               ' // First message parameter
                  ByVal lParam As LPARAM, _               ' // Second message parameter
                  ByVal uIdSubclass As UINT_PTR, _        ' // The subclass ID
                  ByVal dwRefData As DWORD_PTR _          ' // Pointer to reference data
                  ) As LRESULT

   Select Case uMsg
                        
      Case WM_GETDLGCODE
         ' All keyboard input
         Function = DLGC_WANTALLKEYS
         Exit Function

      Case WM_KEYUP
         Select Case Loword(wParam)
            Case VK_RETURN
               ' Simulate the sending of a LBN_DBLCLK to the control.
               SendMessage( GetParent(HWnd), WM_COMMAND, MAKEWPARAM(IDC_FRMFUNCTIONLIST_LISTBOX, LBN_DBLCLK), Cast(LPARAM,HWnd) )
               Exit Function
            Case VK_ESCAPE
               SendMessage( GetParent(hWnd), WM_CLOSE, 0, 0 ) 
         End Select

      case WM_ERASEBKGND
         ' we paint the listbox here because multicolumn ownerdraw listboxes are
         ' quite hard to always get every portion of the background painted exactly.
         ' Only paint the non-listbox text portions in order to prevent too much flicker.
         frmFunctionList_PaintNonItemAreas( cast(HDC, wParam) )
         return true

      Case WM_DESTROY
         ' REQUIRED: Remove control subclassing
         RemoveWindowSubclass HWnd, @frmFunctionList_ListBox_SubclassProc, uIdSubclass

   End Select

   ' Default processing of Windows messages
   Function = DefSubclassProc(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' FunctionList Window procedure
' ========================================================================================
private Function frmFunctionList_WndProc( ByVal HWnd   As HWND, _
                                          ByVal uMsg   As UINT, _
                                          ByVal wParam As WPARAM, _
                                          ByVal lParam As LPARAM _
                                          ) As LRESULT

   Select Case uMsg
      HANDLE_MSG (HWnd, WM_COMMAND,     frmFunctionList_OnCommand)
      HANDLE_MSG (HWnd, WM_SIZE,        frmFunctionList_OnSize)
      HANDLE_MSG (HWnd, WM_CLOSE,       frmFunctionList_OnClose)
      HANDLE_MSG (HWnd, WM_DESTROY,     frmFunctionList_OnDestroy)
      HANDLE_MSG (HWnd, WM_MEASUREITEM, frmFunctionList_OnMeasureItem)
      HANDLE_MSG (HWnd, WM_DRAWITEM,    frmFunctionList_OnDrawItem)
   
'      case WM_CTLCOLOREDIT
'         dim as HDC hDC = cast( HDC, wParam )
'         SetBkColor( hDC, ghPanel.TextBoxBackColor )   
'         SetTextColor( hDC, ghPanel.TextBoxForeColor )
'         return cast( LRESULT, ghPanel.hBackBrushTextBox )

'      case WM_NCPAINT
'         dim as HDC hDC = cast( HDC, wParam ) 
'         'GetDCEx(hwnd, (HRGN)wParam, DCX_WINDOW|DCX_INTERSECTRGN);
'          auto result = DefWindowProc(hWnd, WM_NCPAINT, wParam, lParam);
'          HDC hdc = GetWindowDC(hWnd); 
'          RECT r = GetNonclientMenuBorderRect(hWnd);
'          HBRUSH red = CreateSolidBrush(RGB(255, 0, 0));
'          FillRect(hdc, &r, red);
'          DeleteObject(red);
'          ReleaseDC(hWnd, hdc);

'    return result;         
'         ReleaseDC(HWnd, hDC)
'         return 0

       case WM_ERASEBKGND
         ' we paint the listbox via WM_DRAWITEM so no need to erase and paint
         ' the background here causing unnecessary flicker.
         return true
   End Select

   ' for messages that we don't deal with
   Function = DefWindowProc(HWnd, uMsg, wParam, lParam)

End Function


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

   dim as hwnd hTV = GetDlgItem(HWND_FRMFUNCTIONLIST, IDC_FRMFUNCTIONLIST_TREE)

   Treeview_DeleteAllItems(hTV)

   gApp.hExplorerRootNode = TreeView_AppendItem( hTV, TVI_ROOT, "" )

   ' Add the subnodes for the various file types
   gApp.hExplorerFilesNode = _
      TreeView_AppendItem( hTV, gApp.hExplorerRootNode, "Files", cast(LPARAM, 0) )
      gConfig.Cat(NODETYPE_FILES).wszDescription = L(2, "File")
      gConfig.Cat(NODETYPE_FILES).hNodeExplorer = gApp.hExplorerFilesNode
      gConfig.Cat(NODETYPE_FILES).bShow = true
      
   gApp.hExplorerMainNode = _
      TreeView_AppendItem( hTV, gApp.hExplorerRootNode, "Main", cast(LPARAM, 0) )
      gConfig.Cat(NODETYPE_MAIN).wszDescription = L(212, "Main")
      gConfig.Cat(NODETYPE_MAIN).hNodeExplorer = gApp.hExplorerMainNode
      gConfig.Cat(NODETYPE_MAIN).bShow = true

   gApp.hExplorerResourceNode = _
      TreeView_AppendItem( hTV, gApp.hExplorerRootNode, "Resource", cast(LPARAM, 0) )
      gConfig.Cat(NODETYPE_RESOURCE).wszDescription = L(213, "Resource")
      gConfig.Cat(NODETYPE_RESOURCE).hNodeExplorer = gApp.hExplorerResourceNode
      gConfig.Cat(NODETYPE_RESOURCE).bShow = true

   gApp.hExplorerHeaderNode = _
      TreeView_AppendItem( hTV, gApp.hExplorerRootNode, "Header", cast(LPARAM, 0) )
      gConfig.Cat(NODETYPE_HEADER).wszDescription = L(175, "Header")
      gConfig.Cat(NODETYPE_HEADER).hNodeExplorer = gApp.hExplorerHeaderNode
      gConfig.Cat(NODETYPE_HEADER).bShow = true

   gApp.hExplorerModuleNode = _
      TreeView_AppendItem( hTV, gApp.hExplorerRootNode, "Module", cast(LPARAM, 0) )
      gConfig.Cat(NODETYPE_MODULE).wszDescription = L(211, "Module")
      gConfig.Cat(NODETYPE_MODULE).hNodeExplorer = gApp.hExplorerModuleNode
      gConfig.Cat(NODETYPE_MODULE).bShow = true

   gApp.hExplorerNormalNode = _
      TreeView_AppendItem( hTV, gApp.hExplorerRootNode, "Normal", cast(LPARAM, 0) )
      gConfig.Cat(NODETYPE_NORMAL).wszDescription = L(210, "Normal")
      gConfig.Cat(NODETYPE_NORMAL).hNodeExplorer = gApp.hExplorerNormalNode
      gConfig.Cat(NODETYPE_NORMAL).bShow = true

   for i as long = lbound(gConfig.Cat) to ubound(gConfig.Cat)
      ' User defined categories will have an GUID id. Starts with an {
      if left(gConfig.Cat(i).id, 1) = "{" then
         gConfig.Cat(i).hNodeExplorer = _
            TreeView_AppendItem( hTV, gApp.hExplorerRootNode, gConfig.Cat(i).wszDescription, cast(LPARAM, 0) )
         gConfig.Cat(i).hNodeExplorer = gConfig.Cat(i).hNodeExplorer
         gConfig.Cat(i).bShow = true
      end if
   next

   function = 0
end function


' ========================================================================================
' Get the special subnode based on the filetype
' ========================================================================================
public function frmFunctionList_GetSpecialNode( byval wszFileType as CWSTR ) as HTREEITEM
   
   dim hRootNode as HTREEITEM = 0
   
   ' If project is not active then we retrieve the data from the Files node
   if gApp.IsProjectActive = false then hRootNode = gApp.hExplorerFilesNode

   ' Add the pDoc to the node depending on what file type it is.
   select case wszFileType
      case FILETYPE_UNDEFINED
         if ( gApp.IsProjectActive = true ) orelse ( gApp.IsProjectLoading = true ) then
            hRootNode = gApp.hExplorerNormalNode
         end if   
      case FILETYPE_MAIN:      hRootNode = gApp.hExplorerMainNode
      case FILETYPE_MODULE:    hRootNode = gApp.hExplorerModuleNode
      case FILETYPE_NORMAL:    hRootNode = gApp.hExplorerNormalNode
      case FILETYPE_RESOURCE:  hRootNode = gApp.hExplorerResourceNode
      case FILETYPE_HEADER:    hRootNode = gApp.hExplorerHeaderNode
      case else
         for i as long = lbound(gConfig.Cat) to ubound(gConfig.Cat)
            if gConfig.Cat(i).id = wszFileType then
               hRootNode = gConfig.Cat(i).hNodeExplorer
               exit for
            end if 
         next
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
' frmFunctionList_Show
' ========================================================================================
public Function frmFunctionList_Show( ByVal hWndParent As HWnd ) as LRESULT

   '  Create the main window and child controls
   Dim pWindow As CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowOwnerPtr(hwndParent)->DPI

   Dim rcWork  As Rect = pWindow->GetWorkArea
   Dim nHeight As Long = (rcWork.Bottom - rcWork.Top) * .60
   Dim nWidth  As Long = (rcWork.Right - rcWork.Left) * .80

   HWND_FRMFUNCTIONLIST = _
   pWindow->Create( hWndParent, L(223,"Function List"), @frmFunctionList_WndProc, _
        0, 0, nWidth, nHeight, WS_POPUP Or WS_BORDER or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, 0 )
   pWindow->Center(pWindow->hWindow, hWndParent)
   ' Disable background erasing by only assigning the one style
   pWindow->ClassStyle = CS_DBLCLKS
   
   Dim As HWnd hTextBox = _
   pWindow->AddControl("TEXTBOX", , IDC_FRMFUNCTIONLIST_TXTSEARCH, "", 0, 0, 0, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL, _
        WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   dim wszText as wstring * MAX_PATH 
   wszText = L(42,"Search"): wszText = AfxStrRemove(wszText, "&") + "..."
   SendMessage hTextBox, EM_SETCUEBANNER, 1, CAST(LPARAM, @wszText)

   dim as hwnd hList = _ 
        pWindow->AddControl("LISTBOX", , IDC_FRMFUNCTIONLIST_LISTBOX, "", _
        0, 0, 0, 0, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or WS_HSCROLL Or _
        LBS_MULTICOLUMN Or LBS_NOINTEGRALHEIGHT Or _
        LBS_OWNERDRAWFIXED Or LBS_HASSTRINGS Or LBS_NOTIFY, 0, , _
        Cast(SUBCLASSPROC, @frmFunctionList_ListBox_SubclassProc), _
        IDC_FRMFUNCTIONLIST_LISTBOX, Cast(DWORD_PTR, @pWindow))
   
   ' Can not use a variable height multicolumn listbox because Microsoft says the following:
   ' "  The LBS_MULTICOLUMN and LBS_OWNERDRAWVARIABLE styles cannot be combined. If both are "
   '    specified, LBS_OWNERDRAWVARIABLE is ignored.  "

   SendMessage( hList, LB_SETCOLUMNWIDTH, pWindow->ScaleX(250), 0 )
   
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

