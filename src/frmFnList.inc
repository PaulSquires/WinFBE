'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2018 Paul Squires, PlanetSquires Software
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

Function Listbox_ResetContentEx() As Long
   Dim As HWnd hList  = GetDlgItem(HWND_FRMFNLIST, IDC_FRMFNLIST_LISTBOX)
   dim as Type4lParam ptr pType4lParam
   
   ' Search for a match
   For i As Long = 0 To Listbox_GetCount(hList) - 1
      pType4lParam = Cast(Type4lParam ptr, ListBox_GetItemData(hList, i))
      if pType4lParam THEN
         Delete pType4lParam 
      END IF
   Next
   ' Clear the listbox and add the new items
   RETURN Listbox_ResetContent(hList) 
end function

' ========================================================================================
' Set the Listbox selected item to the currently selected open document (if applicable)
' ========================================================================================
Function frmFnList_SetListBoxPosition() As Long

   Dim nCurSel As Long = -1

   Dim pDocActive As clsDocument Ptr = gTTabCtl.GetActiveDocumentPtr() 
   Dim pDoc As clsDocument Ptr

   Dim As HWnd hList  = GetDlgItem(HWND_FRMFNLIST, IDC_FRMFNLIST_LISTBOX)
   Dim As Long nCount = Listbox_GetCount(hList)
   dim as Type4lParam ptr pType4lParam
   
   ' Search for a match
   For i As Long = 0 To nCount - 1
      pType4lParam = Cast(Type4lParam ptr, ListBox_GetItemData(hList, i))
      if pType4lParam then         
         if pType4lParam->eType = ELP_pclsDocument THEN         
            if pType4lParam->pDoc = pDocActive THEN
               nCurSel = i: Exit For
            END IF
         END IF
      END IF
   Next
   
   Listbox_SetCurSel(hList, nCurSel)
   
   Function = 0
End Function


' ========================================================================================
' Update the Listbox with the list of functions from all loaded documents
' ========================================================================================
Function frmFnList_UpdateListBox() As Long

   Dim As HWnd hList = GetDlgItem(HWND_FRMFNLIST, IDC_FRMFNLIST_LISTBOX)
   
   ' Clear the listbox and add the new items
   Listbox_ResetContentEx()
   
   Dim As Hwnd hTree = GetDlgItem(HWND_FRMEXPLORER, IDC_FRMEXPLORER_TREE)
   Dim As HTREEITEM hChild, hParent,hChild2
   dim as long p, pDoc, idx
   dim as wstring * MAX_PATH  wszText,wszClass
   
   dim as long nFirstProject = iif(gApp.IsProjectActive, 1, 0)
   Dim As Long nProjectCount = gApp.GetProjectCount()
   dim as Type4lParam ptr pType4lParam
   Dim As LPARAM lParam 
   For idx = nFirstProject To nProjectCount 
      if gApp.Projects(idx).InUse = false THEN continue for

      hParent = Treeview_GetChild( hTree, gApp.Projects(gApp.GetActiveProjectIndex).hExplorerRootNode)
      DIM item AS TVITEMW
      with item
         .mask = TVIF_TEXT Or TVIF_IMAGE
         .cchTextMax = MAX_PATH
         do until hParent = 0 
            ' Filename
            Treeview_GetItemText(hTree, hParent, @wszText, MAX_PATH)
             
            lParam = FF_TreeView_GetlParam(hTree, hParent)
            wszText = ucase(wszText)
            p = Listbox_AddString( hList, strptr(wszText) )
            pType4lParam=new Type4lParam
            pType4lParam->eType=ELP_pclsDocument
            pType4lParam->pDoc=cast(clsDocument ptr , lParam)
            ListBox_SetItemData( hList, p, pType4lParam )    ' pDoc is stored for a header
            
            ' Get all functions in the file
            hChild = Treeview_GetChild( hTree, hParent)
            do until hChild = 0
               .hitem = hChild
               .pszText = @wszClass
               TreeView_GetItem(hTree, @item)
               select case .iImage
               case II4E_TYPE4CLASS,II4E_TYPE4CLASS2
                     .pszText = @wszText
                     hChild2 = Treeview_GetChild( hTree, hChild)
                     do until hChild2 = 0
                        .hitem = hChild2
                        TreeView_GetItem(hTree, @item)
                        select case .iImage
                        case II4E_TYPE4FUNCTION,II4E_TYPE4FUNCTION2,II4E_TYPE4SUB, _
                          II4E_TYPE4SUB2, II4E_TYPE4PROPERTY,II4E_TYPE4PROPERTY2
                          wszText = wszClass +"." + wszText
                          p = Listbox_AddString( hList, strptr(wszText) )
                          pType4lParam=new Type4lParam
                          pType4lParam->eType=ELP_hTreeItem
                          pType4lParam->hNode=hChild2
                          ListBox_SetItemData( hList, p, pType4lParam )
                        end select
                        hChild2 = Treeview_GetNextSibling(hTree, hChild2)
                     loop
               case II4E_TYPE4FUNCTION,II4E_TYPE4FUNCTION2,II4E_TYPE4SUB, _
                    II4E_TYPE4SUB2, II4E_TYPE4PROPERTY,II4E_TYPE4PROPERTY2
                     p = Listbox_AddString( hList, strptr(wszClass) )
                     pType4lParam=new Type4lParam
                     pType4lParam->eType=ELP_hTreeItem
                     pType4lParam->hNode=hChild
                     ListBox_SetItemData( hList, p, pType4lParam )
               end select
               hChild = Treeview_GetNextSibling(hTree, hChild)
            loop
            
            ' Get the next filename
            hParent = Treeview_GetNextSibling(hTree, hParent)
         loop
      end with 

   Next
   
   Function = 0
End Function



' ========================================================================================
' Process WM_COMMAND message for window/dialog: FnList
' ========================================================================================
Function frmFnList_OnCommand(ByVal HWnd As HWnd, ByVal id As Long, ByVal hwndCtl As HWnd, ByVal codeNotify As UINT) As LRESULT

   Select Case id
      Case IDC_FRMFNLIST_LISTBOX
         If codeNotify = LBN_DBLCLK Then
            Dim As Long nCurSel = Listbox_GetCurSel(hwndCtl)
            If nCurSel = -1 Then Exit Function
            ' Open the document based on the filename or if it is a header line
            ' then simply use the pDoc.
            dim pDoc as clsDocument ptr
            dim wszText as wstring * MAX_PATH
            ListBox_GetText( hwndCtl, nCurSel, @wszText )
            dim as Type4lParam ptr pType4lParam =cast(Type4lParam ptr, Listbox_GetItemData(hWndCtl, nCurSel))
            if pType4lParam then
               select case pType4lParam->eType
                  case ELP_hTreeItem 
                     OpenSelectedDocument(0, wszText, pType4lParam->hNode)
                  case ELP_pclsDocument 
                     OpenSelectedDocument(pType4lParam->pDoc, wszText,0)
               end select
            end if
            If gConfig.CloseFuncList Then ShowWindow( HWND_FRMFNLIST, SW_HIDE )
         end if      
   End Select
   
   Function = 0
End Function


' ========================================================================================
' Position all child windows. Called manually and/or by WM_SIZE
' ========================================================================================
Function frmFnList_PositionWindows( ByVal HWnd As HWnd ) As LRESULT
   
   ' Get the entire client area
   Dim As Rect rc
   GetClientRect(HWnd, @rc)
   
   Dim As HWnd hList = GetDlgItem(HWnd, IDC_FRMFNLIST_LISTBOX )
   SetWindowPos hList, 0, 0, 0, rc.Right-rc.Left, rc.Bottom-rc.top, SWP_SHOWWINDOW Or SWP_NOZORDER

   Function = 0
End Function


' ========================================================================================
' Process WM_SIZE message for window/dialog: FnList
' ========================================================================================
Function frmFnList_OnSize(ByVal HWnd As HWnd, ByVal state As UINT, ByVal cx As Long, ByVal cy As Long) As LRESULT

   If state <> SIZE_MINIMIZED Then 
      frmFnList_PositionWindows(HWnd)
   End If   

   Function = 0
End Function


' ========================================================================================
' Process WM_CLOSE message for window/dialog: FnList
' ========================================================================================
Function frmFnList_OnClose( ByVal HWnd As HWnd ) As LRESULT
   ' Never close the window; simply hide it.
   ShowWindow( HWnd, SW_HIDE )
   Function = 0
End Function


' ========================================================================================
' Process WM_MEASUREITEM message for window/dialog: FnList
' ========================================================================================
Function frmFnList_OnMeasureItem( ByVal HWnd As HWnd, ByVal lpmis As MEASUREITEMSTRUCT Ptr ) As Long
   ' Set the height of the List box items. 
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(HWnd)

   If lpmis->CtlId <> IDC_FRMFNLIST_LISTBOX Then Exit Function
   lpmis->itemHeight = 18 * pWindow->ryRatio
      
   Function = 0
End Function


' ========================================================================================
' Process WM_DRAWITEM message for window/dialog: FnList
' ========================================================================================
Function frmFnList_OnDrawItem( ByVal HWnd As HWnd, ByVal lpdis As Const DRAWITEMSTRUCT Ptr ) As Long

   Dim As HBRUSH hBrush 
   Dim As HFONT hFont, hFontOld
   dim as BOOLEAN bIsHeader = false
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
         dim as Type4lParam ptr pType4lParam =cast(Type4lParam ptr,lpdis->itemData)
         if pType4lParam then         
            if pType4lParam->eType = ELP_pclsDocument THEN
               bIsHeader =  true
            END IF
         END IF
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
' Process WM_DESTROY message for window/dialog: FnList
' ========================================================================================
Function frmFnList_OnDestroy(HWnd As HWnd) As LRESULT
   Listbox_ResetContentEx()
   HWND_FRMFNLIST = 0
   Function = 0
End Function


' ========================================================================================
' Processes messages for the subclassed ListBox window.
' ========================================================================================
Function frmFnList_ListBox_SubclassProc ( _
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
               SendMessage( GetParent(HWnd), WM_COMMAND, MAKEWPARAM(IDC_FRMFNLIST_LISTBOX, LBN_DBLCLK), Cast(LPARAM,HWnd) )
               Exit Function
            Case VK_ESCAPE
               SendMessage( GetParent(hWnd), WM_CLOSE, 0, 0 ) 
         End Select

      Case WM_DESTROY
         ' REQUIRED: Remove control subclassing
         RemoveWindowSubclass HWnd, @frmFnList_ListBox_SubclassProc, uIdSubclass

   End Select

   ' Default processing of Windows messages
   Function = DefSubclassProc(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' FnList Window procedure
' ========================================================================================
Function frmFnList_WndProc( ByVal HWnd   As HWnd, _
                            ByVal uMsg   As UINT, _
                            ByVal wParam As WPARAM, _
                            ByVal lParam As LPARAM _
                            ) As LRESULT

   Select Case uMsg
      HANDLE_MSG (HWnd, WM_COMMAND,     frmFnList_OnCommand)
      HANDLE_MSG (HWnd, WM_SIZE,        frmFnList_OnSize)
      HANDLE_MSG (HWnd, WM_CLOSE,       frmFnList_OnClose)
      HANDLE_MSG (HWnd, WM_DESTROY,     frmFnList_OnDestroy)
      HANDLE_MSG (HWnd, WM_MEASUREITEM, frmFnList_OnMeasureItem)
      HANDLE_MSG (HWnd, WM_DRAWITEM,    frmFnList_OnDrawItem)
   End Select

   ' for messages that we don't deal with
   Function = DefWindowProc(HWnd, uMsg, wParam, lParam)

End Function


''
''  Need to re-parse any dirty files. This is especially important for QuickRun scenarios
''  where the file is never physically saved.
''
function frmFnList_ReparseFiles() as Long
   dim as long idx = gApp.GetActiveProjectIndex()
   Dim pDoc As clsDocument Ptr
   Dim As Long i, nCount
   nCount = gApp.Projects(idx).GetDocumentCount 
   For i = 0 To nCount - 1
      pDoc = gApp.Projects(idx).GetDocumentPtr(i)
      If pDoc Then 
         If cbool(SciExec( pDoc->hWindow(0), SCI_GETMODIFY, 0, 0 )) or pDoc->UserModified Then
            ParseDocument(idx, pDoc, "")
            AddFunctionsToExplorerTreeview(pDoc, true)
         end if
      end if
   next      
   function = 0
end function

' ========================================================================================
' frmFnList_Show
' ========================================================================================
Function frmFnList_Show( ByVal hWndParent As HWnd, _
                         ByVal nCmdShow   As Long = 0 _
                         ) As Long

   ' If function list already exists then toggle its visibility
   ' unless nCmdShow is explicitly set to show it.
   If IsWindow(HWND_FRMFNLIST) Then
      If nCmdShow <> SW_SHOW Then nCmdShow = Iif(IsWindowVisible(HWND_FRMFNLIST), SW_HIDE, SW_SHOW)
      ShowWindow HWND_FRMFNLIST, nCmdShow
      If nCmdShow = SW_SHOW Then
         ' Need to re-parse any dirty files. This is especially important for QuickRun scenarios
         ' where the file is never physically saved.
         frmFnList_ReparseFiles
         frmFnList_UpdateListBox()
         frmFnList_SetListBoxPosition()
         SetFocus GetDlgItem(HWND_FRMFNLIST, IDC_FRMFNLIST_LISTBOX)
      End If
      Exit Function 
   Else 
      ' If the window does not exist yet then ensure that it becomes visible after creation.
      nCmdShow = SW_SHOW
   End If
      

   '  Create the main window and child controls
   Dim pWindow As CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowOwnerPtr(hwndParent)->DPI

   Dim rcWork  As Rect = pWindow->GetWorkArea
   Dim nHeight As Long = (rcWork.Bottom - rcWork.Top) * .60
   Dim nWidth  As Long = (rcWork.Right - rcWork.Left) * .80

   HWND_FRMFNLIST = _
   pWindow->Create( hWndParent, L(223,"Function List"), @frmFnList_WndProc, 0, 0, nWidth, nHeight, _
        WS_POPUP Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, _
        WS_EX_CONTROLPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->Center(pWindow->hWindow, hWndParent)
   
   
   ' Set the small and large icon for the main window (must be set after main window is created)
   pWindow->BigIcon   =  LoadImage( pWindow->InstanceHandle, "IMAGE_AAA_MAINICON", IMAGE_ICON, 32, 32, LR_SHARED)
   pWindow->SmallIcon =  LoadImage( pWindow->InstanceHandle, "IMAGE_AAA_MAINICON", IMAGE_ICON, 16, 16, LR_SHARED)

   Dim As HWnd hList = _ 
        pWindow->AddControl("LISTBOX", , IDC_FRMFNLIST_LISTBOX, "", 0, 0, 0, 0, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or WS_HSCROLL Or LBS_MULTICOLUMN Or LBS_NOINTEGRALHEIGHT Or _
        LBS_OWNERDRAWFIXED Or LBS_HASSTRINGS Or LBS_NOTIFY, WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_RIGHTSCROLLBAR, , _
        Cast(SUBCLASSPROC, @frmFnList_ListBox_SubclassProc), IDC_FRMFNLIST_LISTBOX, Cast(DWORD_PTR, @pWindow))
   
   ' Can not use a variable height multicolumn listbox because Microsoft says the following:
   ' "  The LBS_MULTICOLUMN and LBS_OWNERDRAWVARIABLE styles cannot be combined. If both are "
   '    specified, LBS_OWNERDRAWVARIABLE is ignored.  "

   SendMessageW( hList, LB_SETCOLUMNWIDTH, 250 * pWindow->rxRatio, 0 )
   
   frmFnList_ReparseFiles
   frmFnList_UpdateListBox()
   frmFnList_SetListBoxPosition()

   frmFnList_PositionWindows(HWND_FRMFNLIST)

   ShowWindow HWND_FRMFNLIST, nCmdShow
   SetFocus hList
   
   Function = 0
End Function





