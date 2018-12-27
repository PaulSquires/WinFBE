'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2019 Paul Squires, PlanetSquires Software
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


' ========================================================================================
' Set the Listbox selected item to the currently selected open document (if applicable)
' ========================================================================================
Function frmFunctionList_SetListBoxPosition() As Long

   Dim nCurSel As Long = -1

   Dim pDocActive As clsDocument Ptr = gTTabCtl.GetActiveDocumentPtr() 
   Dim pDoc As clsDocument Ptr
   Dim As HTREEITEM hItem

   Dim As Hwnd hTree = GetDlgItem(HWND_FRMEXPLORER, IDC_FRMEXPLORER_TREE)
   Dim As HWnd hList  = GetDlgItem(HWND_FRMFUNCTIONLIST, IDC_FRMFUNCTIONLIST_LISTBOX)
   Dim As Long nCount = Listbox_GetCount(hList)
   
   ' Search for a match
   For i As Long = 0 To nCount - 1
      hItem = cast(HTREEITEM, ListBox_GetItemData(hList, i))
      pDoc = Cast(clsDocument ptr, FF_TreeView_GetlParam(hTree, hItem))
      if pDoc = pDocActive THEN
         nCurSel = i: Exit For
      END IF
   Next
   
   Listbox_SetCurSel(hList, nCurSel)
   
   Function = 0
End Function


' ========================================================================================
' Search child nodes to see if search text exists. Needed to allow output of parent node.
' ========================================================================================
function frmFunctionList_SearchChildNodes( byval hParent as HTREEITEM, _
                                           byval hChild as HTREEITEM, _
                                           byref wszSearch as wstring _
                                           ) as Boolean
   dim as Boolean bFound = true
   
   ' If a search term exists then only add the header if a valid found
   ' sub/function matches in one of the child nodes.
   if len(wszSearch) then
      Dim As Hwnd hTree = GetDlgItem(HWND_FRMEXPLORER, IDC_FRMEXPLORER_TREE)
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
Function frmFunctionList_UpdateListBox() As Long

   Dim As HWnd hList = GetDlgItem(HWND_FRMFUNCTIONLIST, IDC_FRMFUNCTIONLIST_LISTBOX)
   
   ' Clear the listbox and add the new items
   Listbox_ResetContent(hList)
   
   Dim As Hwnd hTree = GetDlgItem(HWND_FRMEXPLORER, IDC_FRMEXPLORER_TREE)
   Dim As HTREEITEM hChild, hChildSearch, hParent
   dim as long p, pDoc
   dim as Boolean bFound
   dim wszText as wstring * MAX_PATH
   dim as CWSTR wszSearch = AfxGetWindowText(GetDlgItem(HWND_FRMFUNCTIONLIST, IDC_FRMFUNCTIONLIST_TXTSEARCH))
   
   LockWindowUpdate(hList)
   hParent = Treeview_GetChild( hTree, gApp.hExplorerRootNode)
   do until hParent = 0 
      ' Filename
      Treeview_GetItemText(hTree, hParent, @wszText, MAX_PATH)
      pDoc = FF_TreeView_GetlParam(hTree, hParent)
      wszText = ucase(wszText)

      bFound = frmFunctionList_SearchChildNodes(hParent, hChild, wszSearch) 
      if bFound then   
         p = Listbox_AddString( hList, strptr(wszText) )
         ListBox_SetItemData( hList, p, hParent )    ' Treeview node item is stored in ListBox ItemData
      end if
      
      ' Get all functions in the file
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
   LockWindowUpdate(null)

   Function = 0
End Function



' ========================================================================================
' Process WM_COMMAND message for window/dialog: FunctionList
' ========================================================================================
Function frmFunctionList_OnCommand(ByVal HWnd As HWnd, ByVal id As Long, ByVal hwndCtl As HWnd, ByVal codeNotify As UINT) As LRESULT

   Select Case id
      Case IDC_FRMFUNCTIONLIST_LISTBOX
         If codeNotify = LBN_DBLCLK Then
            Dim As Long nCurSel = Listbox_GetCurSel(hwndCtl)
            If nCurSel = -1 Then Exit Function

            ' The ItemData holds the Explorer Treeview node item handle
            dim as CWSTR wszFilename, wszFunctionName
            Dim As HTREEITEM hNode = cast(HTREEITEM, Listbox_GetItemData(hWndCtl, nCurSel))
            frmExplorer_GetFileNameFunctionName(hNode, wszFilename, wszFunctionName)

            If gConfig.CloseFuncList Then 
               PostMessage( hWnd, WM_CLOSE, 0, 0 ) 
            end if
            OpenSelectedDocument(wszFilename, wszFunctionName)
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
Function frmFunctionList_PositionWindows( ByVal HWnd As HWnd ) As LRESULT
   
   ' Get the entire client area
   Dim As Rect rc
   GetClientRect(HWnd, @rc)
   
   dim as long nHeight = AfxGetWindowHeight(GetDlgItem(HWnd, IDC_FRMFUNCTIONLIST_TXTSEARCH))
   SetWindowPos GetDlgItem(HWnd, IDC_FRMFUNCTIONLIST_TXTSEARCH ), 0, _
                  0, 0, rc.Right-rc.Left, nHeight, SWP_NOZORDER

   SetWindowPos GetDlgItem(HWnd, IDC_FRMFUNCTIONLIST_LISTBOX ), 0, _
                  0, nHeight, rc.Right-rc.Left, rc.Bottom-rc.top-nHeight, SWP_NOZORDER

   Function = 0
End Function


' ========================================================================================
' Process WM_SIZE message for window/dialog: FunctionList
' ========================================================================================
Function frmFunctionList_OnSize(ByVal HWnd As HWnd, ByVal state As UINT, ByVal cx As Long, ByVal cy As Long) As LRESULT
   If state <> SIZE_MINIMIZED Then 
      frmFunctionList_PositionWindows(HWnd)
   End If   
   Function = 0
End Function


' ========================================================================================
' Process WM_CLOSE message for window/dialog: FunctionList
' ========================================================================================
Function frmFunctionList_OnClose( ByVal HWnd As HWnd ) As LRESULT
   DestroyWindow(HWnd)
   Function = 0
End Function


' ========================================================================================
' Process WM_MEASUREITEM message for window/dialog: FunctionList
' ========================================================================================
Function frmFunctionList_OnMeasureItem( ByVal HWnd As HWnd, ByVal lpmis As MEASUREITEMSTRUCT Ptr ) As Long
   ' Set the height of the List box items. 
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(HWnd)

   If lpmis->CtlId <> IDC_FRMFUNCTIONLIST_LISTBOX Then Exit Function
   lpmis->itemHeight = 18 * pWindow->ryRatio
      
   Function = 0
End Function


' ========================================================================================
' Process WM_DRAWITEM message for window/dialog: FunctionList
' ========================================================================================
Function frmFunctionList_OnDrawItem( ByVal HWnd As HWnd, ByVal lpdis As Const DRAWITEMSTRUCT Ptr ) As Long

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
         dim as hwnd hTree = GetDlgItem(HWND_FRMEXPLORER, IDC_FRMEXPLORER_TREE)
         bIsHeader = iif( FF_TreeView_GetlParam( hTree, cast(HTREEITEM, lpdis->itemData)), true, false)

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

   Function = 0
End Function

   
' ========================================================================================
' Process WM_DESTROY message for window/dialog: FunctionList
' ========================================================================================
Function frmFunctionList_OnDestroy(HWnd As HWnd) As LRESULT
   HWND_FRMFUNCTIONLIST = 0
   Function = 0
End Function


' ========================================================================================
' Processes messages for the subclassed ListBox window.
' ========================================================================================
Function frmFunctionList_ListBox_SubclassProc ( _
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
Function frmFunctionList_WndProc( ByVal HWnd   As HWnd, _
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
   End Select

   ' for messages that we don't deal with
   Function = DefWindowProc(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' Reparse any dirty files
' ========================================================================================
function frmFunctionList_ReparseFiles() as Long
   '  Need to re-parse any dirty files. This is especially important for QuickRun scenarios
   '  where the file is never physically saved.
   Dim pDoc As clsDocument Ptr = gApp.pDocList
   do until pDoc = 0
      If cbool(SciExec( pDoc->hWindow(0), SCI_GETMODIFY, 0, 0 )) or pDoc->UserModified Then
         ParseDocument(pDoc, "")
         AddFunctionsToExplorerTreeview(pDoc, true)
      end if
      pDoc = pDoc->pDocNext
   loop   
   function = 0
end function

' ========================================================================================
' frmFunctionList_Show
' ========================================================================================
Function frmFunctionList_Show( ByVal hWndParent As HWnd, _
                         ByVal nCmdShow   As Long = 0 _
                         ) As Long

   ' Calling this function works as a toggle to display/close the popup form.
   if IsWindowVisible(HWND_FRMFUNCTIONLIST) then
      DestroyWindow(HWND_FRMFUNCTIONLIST)
      exit function
   end if
   
   '  Create the main window and child controls
   Dim pWindow As CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowOwnerPtr(hwndParent)->DPI

   Dim rcWork  As Rect = pWindow->GetWorkArea
   Dim nHeight As Long = (rcWork.Bottom - rcWork.Top) * .60
   Dim nWidth  As Long = (rcWork.Right - rcWork.Left) * .80

   HWND_FRMFUNCTIONLIST = _
   pWindow->Create( hWndParent, L(223,"Function List"), @frmFunctionList_WndProc, 0, 0, nWidth, nHeight, _
        WS_POPUP Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, _
        WS_EX_CONTROLPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->Center(pWindow->hWindow, hWndParent)
   
   
   ' Set the small and large icon for the main window (must be set after main window is created)
   pWindow->BigIcon   =  LoadImage( pWindow->InstanceHandle, "IMAGE_AAA_MAINICON", IMAGE_ICON, 32, 32, LR_SHARED)
   pWindow->SmallIcon =  LoadImage( pWindow->InstanceHandle, "IMAGE_AAA_MAINICON", IMAGE_ICON, 16, 16, LR_SHARED)

   Dim As HWnd hTextBox = _
   pWindow->AddControl("TEXTBOX", , IDC_FRMFUNCTIONLIST_TXTSEARCH, "", 0, 0, 0, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   dim wszText as wstring * MAX_PATH 
   wszText = L(42,"Search"): wszText = AfxStrRemove(wszText, "&") + "..."
   SendMessage hTextBox, EM_SETCUEBANNER, 1, CAST(LPARAM, @wszText)

   dim as hwnd hList = _ 
        pWindow->AddControl("LISTBOX", , IDC_FRMFUNCTIONLIST_LISTBOX, "", 0, 0, 0, 0, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or WS_HSCROLL Or LBS_MULTICOLUMN Or LBS_NOINTEGRALHEIGHT Or _
        LBS_OWNERDRAWFIXED Or LBS_HASSTRINGS Or LBS_NOTIFY, WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_RIGHTSCROLLBAR, , _
        Cast(SUBCLASSPROC, @frmFunctionList_ListBox_SubclassProc), IDC_FRMFUNCTIONLIST_LISTBOX, Cast(DWORD_PTR, @pWindow))
   
   ' Can not use a variable height multicolumn listbox because Microsoft says the following:
   ' "  The LBS_MULTICOLUMN and LBS_OWNERDRAWVARIABLE styles cannot be combined. If both are "
   '    specified, LBS_OWNERDRAWVARIABLE is ignored.  "

   SendMessageW( hList, LB_SETCOLUMNWIDTH, 250 * pWindow->rxRatio, 0 )
   
   frmFunctionList_PositionWindows(HWND_FRMFUNCTIONLIST)

   LockWindowUpdate(hList)
   frmFunctionList_ReparseFiles
   frmFunctionList_UpdateListBox()
   frmFunctionList_SetListBoxPosition()
   LockWindowUpdate(null)
   AfxRedrawWindow(hList)

   ShowWindow HWND_FRMFUNCTIONLIST, SW_SHOW
   SetFocus hList
   
   Function = 0
End Function





