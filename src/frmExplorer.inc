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


''
''  frmExplorer.inc
''   
''  This is the left hand side panel that houses the Explorer treeview listing all of
''  the project files. When no project open or files to show, then the Explorer will
''  the Open/Recent treeviews.
''

#include once "frmExplorer.bi"
#include once "frmRecent.bi"
#include once "clsDocument.bi"
#include once "clsDB2.bi"


' ========================================================================================
' Position all child windows. Called manually and/or by WM_SIZE
' ========================================================================================
private Function PositionExplorerWindows( ByVal HWnd As HWnd ) As LRESULT
   
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(HWnd)
   If pWindow = 0 Then Exit Function
   
   ' Get the entire client area
   Dim As Rect rc
   GetClientRect(HWnd, @rc)

   Dim As HWnd hCtl, hTree
   Dim As Long nLeft, nTop, nWidth, nHeight
   
   ' The Explorer label was painted in the top area so start our treeviews immediately after that
   nWidth  = pWindow->ScaleX(24)
   nHeight = pWindow->ScaleY(24)
   nTop    = pWindow->ScaleY(10)
   
   hCtl = GetDlgItem(HWND, IDC_FRMEXPLORER_BTNCLOSE)
   SetWindowPos( hCtl, 0, rc.Right - (nWidth*1.4), nTop, nWidth, nHeight, SWP_SHOWWINDOW Or SWP_NOZORDER )

   nTop = nTop + nHeight + pWindow->ScaleY(12)

   hTree = GetDlgItem(HWND_FRMEXPLORER, IDC_FRMEXPLORER_TREE)
   
   If gApp.IsProjectActive Then
      ' Project is active to continue to show the Explorer tree regardless of whether 
      ' any files are actually open.
      ShowWindow HWND_FRMRECENT, SW_HIDE
      ShowWindow hTree, SW_SHOW
   else
      ' If no project open then show Recent dialog
      if gApp.GetDocumentCount = 0 then
         ' Ensure that the Recent lists are the most current
         If IsWindowVisible(HWND_FRMRECENT) = 0 Then frmRecent_LoadTreeview(HWND_FRMRECENT)
         ShowWindow hTree, SW_HIDE
         ShowWindow HWND_FRMRECENT, SW_SHOW
      else
         ShowWindow HWND_FRMRECENT, SW_HIDE
         ShowWindow hTree, SW_SHOW
      end if
   End If   

   ' Allow a small space at the right edge to ensure area for splitter
   SetWindowPos( IIF(IsWindowVisible(hTree), hTree, HWND_FRMRECENT), _
                 0, 0, nTop, rc.Right - pWindow->ScaleX(2), rc.bottom-nTop, SWP_SHOWWINDOW Or SWP_NOZORDER )

   AfxRedrawWindow(HWND_FRMEXPLORER)
   
   Function = 0
End Function
 

' ========================================================================================
' Add the names of all subs/functions belonging to this document to the Explorer treeview
' ========================================================================================
private Function AddFunctionsToExplorerTreeview( ByVal pDoc As clsDocument Ptr, _
                                                 ByVal fUpdateNodes As BOOLEAN _
                                                 ) As Long
 
   If pDoc = 0 Then Exit Function
   
   Dim As CWSTR wzFile = AfxStrPathname("NAMEX", pDoc->DiskFilename)
   Dim As Hwnd hTree   = GetDlgItem(HWND_FRMEXPLORER, IDC_FRMEXPLORER_TREE)
   Dim As HTREEITEM hChild, hRootNode
   Dim As BOOLEAN fAddChildren = False
   dim as long idx
   
   hRootNode = gApp.hExplorerRootNode

   ' If hNodeExplorer does not exist then we must be adding a new document. If it does
   ' exist then we simply need to update the child nodes.
   If fUpdateNodes = False Then
      If pDoc->hNodeExplorer = 0 Then
         pDoc->hNodeExplorer = FF_TreeView_InsertItem (hTree, hRootNode, wzFile, cast(LPARAM, pDoc) )
         fAddChildren = True
      End If
   Else
      ' Delete any existing children sub/function nodes and then update with new pointer values
      hChild = TreeView_GetChild( hTree, pDoc->hNodeExplorer )
      Do Until hChild = 0
         TreeView_DeleteItem( hTree, hChild )
         hChild = TreeView_GetChild( hTree, pDoc->hNodeExplorer )
      Loop
      fAddChildren = True
   End If
 
   ' Need to ensure the Explorer treeview is sorted because the filename may have changed.
   Treeview_SortChildren( hTree, hRootNode, 0 )
      
   If fAddChildren Then
      dim wszText as wstring * MAX_PATH
      dim as HTREEITEM hNode
      dim pData as DB2_DATA ptr
      dim as string sFilename = ucase(pDoc->DiskFilename)
      
      gdb2.dbRewind()
      do 
         pData = gdb2.dbGetNext
         if pData = 0 THEN exit do
         select case pData->id 
            case DB2_FUNCTION, DB2_SUB
               if sFilename = ucase(pData->fileName) THEN
                  hNode = FF_TreeView_InsertItem(hTree, pDoc->hNodeExplorer, pData->ElementName, 0 )
               END IF
         END SELECT
      loop      
      Treeview_SortChildren(hTree, pDoc->hNodeExplorer, 0)
   End If
   AfxRedrawWindow(hTree)
   
   Function = 0
End Function


' ========================================================================================
' Process WM_SIZE message for window/dialog: frmExplorer
' ========================================================================================
private Function frmExplorer_OnSize( ByVal HWnd As HWnd, _
                                     ByVal state As UINT, _
                                     ByVal cx As Long, _
                                     ByVal cy As Long _
                                     ) As LRESULT
   If state <> SIZE_MINIMIZED Then
      ' Position all of the child windows
      PositionExplorerWindows(HWnd)
   End If
   Function = 0
End Function


' ========================================================================================
' Process WM_NOTIFY message for window/dialog: frmExplorer
' ========================================================================================
private Function frmExplorer_OnNotify( ByVal HWnd As HWnd, _
                                       ByVal id As Long, _
                                       ByVal pNMHDR As NMHDR Ptr _
                                       ) As LRESULT
 
   Dim lpNMTV As NM_TREEVIEW Ptr = Cast(NM_TREEVIEW Ptr, pNMHDR)
   Dim As HTREEITEM hItem
   Dim As HWND hTree 
   Dim As LPARAM lParam
   Dim pDoc As clsDocument Ptr
   Dim ht As TVHITTESTINFO
   Dim As POINT pt
   
   ' Determine what node was clicked on
   Select Case pNMHDR->code 
      
      case TVN_GETINFOTIP
         ' Display a tooltip when a treeview node is hovered over. This will normally
         ' be the full filename for the project or document.
         dim pnmTip as NMTVGETINFOTIPW ptr = cast(NMTVGETINFOTIPW ptr, pNMHDR)
         if gApp.hExplorerRootNode = pnmTip->hItem then
            pnmTip->pszText = cast(WSTRING Ptr, @**gApp.ProjectFilename)
            pnmTip->cchTextMax = len(gApp.ProjectFilename)
            exit function
         end if   
         ' Second, search for a matching document
         pDoc = Cast(clsDocument ptr, pnmTip->lParam)
         if pDoc THEN
            pnmTip->pszText = @pDoc->DiskFilename
            pnmTip->cchTextMax = len(pDoc->DiskFilename)
         END IF
         exit function   

      Case NM_CLICK, NM_RCLICK
         if id = IDC_FRMEXPLORER_TREE then
            ' If we have clicked on a folder item then issue an expand. Must use HitTest because the 
            ' the current node has not been selected yet.
            hTree = lpNMTV->hdr.hWndFrom
            GetCursorPos(@pt)  
            ScreenToClient(hTree, @pt)
            
            ht.pt = pt 
            Treeview_HitTest(hTree, @ht )
            hItem = ht.hItem
            If hItem Then Treeview_Select( hTree, hItem, TVGN_CARET)
            
            If pNMHDR->code = NM_CLICK Then
               ' Ensure that the correct Project name is showing
               frmMain_PositionWindows(HWND_FRMMAIN)
            end if
            
            If pNMHDR->code = NM_RCLICK Then
               ' Create the popup menu
               lParam = FF_TreeView_GetlParam(hTree, hItem)
               pDoc = Cast(clsDocument ptr, lParam)
               ' Select the node and show the document
               if pDoc then SendMessage(hTree, WM_LBUTTONDBLCLK, 0, 0)
               Dim As HMENU hPopupMenu = CreateExplorerContextMenu(pDoc)
               GetCursorPos @pt
               TrackPopupMenu(hPopUpMenu, 0, pt.x, pt.y, 0, HWND_FRMMAIN, ByVal Null)
               DestroyMenu hPopUpMenu
               Function = True
            End If
         end if
         
         if id = IDC_FRMEXPLORER_BTNCLOSE then
            If pNMHDR->code = NM_CLICK Then
               ShowWindow HWND_FRMEXPLORER, SW_HIDE
               frmMain_PositionWindows(HWND_FRMMAIN)
            end if
         end if
         
   End Select
 
   Function = 0
End Function


' ========================================================================================
' Retrieve the filename and function name of the specified Treeview node
' ========================================================================================
private function frmExplorer_GetFileNameFunctionName( byval hItem as HTREEITEM, _
                                                      byref wszFilename as CWSTR, _
                                                      byref wszFunctionName as CWSTR _
                                                      ) as long

   Dim As Hwnd hTree = GetDlgItem(HWND_FRMEXPLORER, IDC_FRMEXPLORER_TREE)
   Dim pDoc As clsDocument Ptr
   dim as LPARAM lParam 
   dim zFilename as wstring * MAX_PATH
   dim zFunctionName as wstring * MAX_PATH
            
   lParam = FF_TreeView_GetlParam(hTree, hItem)
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
      lParam = FF_TreeView_GetlParam(hTree, hItem)
      pDoc = Cast(clsDocument Ptr, lParam)
      if pDoc then zFilename = pDoc->DiskFilename
   end if

   wszFilename = zFilename
   wszFunctionName = zFunctionName
   
   function = 0
end function


' ========================================================================================
' Processes messages for the subclassed frmExplorer treeview control.
' ========================================================================================
private Function frmExplorer_Tree_SubclassProc ( _
                  ByVal HWnd   As HWnd, _                 ' // Control window handle
                  ByVal uMsg   As UINT, _                 ' // Type of message
                  ByVal wParam As WPARAM, _               ' // First message parameter
                  ByVal lParam As LPARAM, _               ' // Second message parameter
                  ByVal uIdSubclass As UINT_PTR, _        ' // The subclass ID
                  ByVal dwRefData As DWORD_PTR _          ' // Pointer to reference data
                  ) As LRESULT

   Dim As HWND hTree 
   Dim As POINT pt
   Dim As HTREEITEM hItem
   Dim As BOOLEAN bIsFolder
   Dim As BOOLEAN bIsExpanded

   ' Convert our ENTER key presses into LBUTTONDBLCLK to process them similarly
   If (uMsg = WM_KEYUP) And (Loword(wParam) = VK_RETURN) Then uMsg = WM_LBUTTONDBLCLK

      
   Select Case uMsg

      Case WM_GETDLGCODE
         ' All keyboard input
         Function = DLGC_WANTALLKEYS
         Exit Function

      Case WM_LBUTTONDBLCLK
         hTree = hWnd
         hItem = Treeview_GetSelection(hTree)
         dim as CWSTR wszFilename, wszFunctionName
         frmExplorer_GetFileNameFunctionName(hItem, wszFilename, wszFunctionName)
         OpenSelectedDocument(wszFilename, wszFunctionName)
         TreeView_SelectItem(hTree, hItem)
         Exit Function
         
      Case WM_KEYUP
         Select Case Loword(wParam)
            Case VK_RETURN  ' already processed in WM_LBUTTONDBLCLK 
         End Select
         Exit Function

      Case WM_CHAR   ' prevent the annoying beep!
         If wParam = VK_RETURN Then Return 0

      Case WM_DESTROY
         ' REQUIRED: Remove control subclassing
         RemoveWindowSubclass HWnd, @frmExplorer_Tree_SubclassProc, uIdSubclass

   End Select

   ' Default processing of Windows messages
   Function = DefSubclassProc(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' Process WM_PAINT message for window/dialog: frmExplorer
' ========================================================================================
private Function frmExplorer_OnPaint( ByVal HWnd As HWnd ) As LRESULT
            
   Dim As PAINTSTRUCT ps
   Dim As HPEN hPen
   Dim As HDC hDc
   Dim As Rect rc
   
   Dim pWindow As CWindow Ptr = AfxCWindowOwnerPtr(HWND)
   if pWindow = 0 THEN exit function

   hDC = BeginPaint(hWnd, @ps)
   GetClientRect HWnd, @rc

   SaveDC hDC

   FillRect( hDC, @rc, GetSysColorBrush( COLOR_BTNFACE) )
   
   ' Outline the right edge of dialog with a dull gray pen
   hPen = CreatePen( PS_SOLID, 1, GetSysColor(COLOR_ACTIVEBORDER) )
   SelectPen( hDC, hPen )
   MoveToEx( hDC, rc.right - 1, rc.top, null)
   LineTo( hDC, rc.right - 1, rc.bottom-1 )
   RestoreDC hDC, -1 
   DeletePen(hPen)
   
   EndPaint hWnd, @ps
   
   Function = 0
End Function


' ========================================================================================
' frmExplorer Window procedure
' ========================================================================================
private Function frmExplorer_WndProc( ByVal HWnd   As HWnd, _
                                      ByVal uMsg   As UINT, _
                                      ByVal wParam As WPARAM, _
                                      ByVal lParam As LPARAM _
                                      ) As LRESULT

   Select Case uMsg
      HANDLE_MSG (HWnd, WM_SIZE,    frmExplorer_OnSize)
      HANDLE_MSG (HWnd, WM_NOTIFY,  frmExplorer_OnNotify)
      HANDLE_MSG (HWnd, WM_PAINT,   frmExplorer_OnPaint)
   
   case WM_DESTROY
      Dim As HWnd hTree = GetDlgItem( hWnd, IDC_FRMEXPLORER_TREE)
      dim as HFONT hFont = AfxGetWindowFont(hTree)
      DIM pImageCtx AS CImageCtx ptr = AfxCImageCtxPtr(hwnd, IDC_FRMEXPLORER_BTNCLOSE)
      if pImageCtx then delete pImageCtx
      DeleteFont(hFont)
      
   case WM_ERASEBKGND
      return true

   case WM_CTLCOLORSTATIC
      ' Background color for the Explorer label
      SetBkColor Cast(HDC,wParam), GetSysColor(COLOR_BTNFACE)
      return Cast( LRESULT, GetSysColorBrush(COLOR_BTNFACE))

   End Select

   ' for messages that we don't deal with
   Function = DefWindowProc(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' frmExplorer_Show
' ========================================================================================
public Function frmExplorer_Show( ByVal hWndParent As HWnd ) As LRESULT

   '  Create the main window and child controls
   Dim pWindow As CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowPtr(hwndParent)->DPI

   ' Only make the Explorer initially visable if it was already visible
   ' when the most previous instance of the program closed. Also, set the width of
   ' the window to the last used visible width.
   dim as long nWidth = iif(gConfig.ShowExplorer, gConfig.ShowExplorerWidth, 250)
   
   HWND_FRMEXPLORER = pWindow->Create( hWndParent, "Explorer Window", @frmExplorer_WndProc, 0, 0, nWidth, 0, _
        WS_CHILD Or iif(gConfig.ShowExplorer, WS_VISIBLE, 0) Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, _
        WS_EX_CONTROLPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)

   ' Disable background erasing by only assigning the one style
   pWindow->ClassStyle = CS_DBLCLKS
   pWindow->SetFont( "Segoe UI Semibold", 9)

   Dim As HWnd hBtn = _
   pWindow->AddControl("LABEL", , IDC_FRMEXPLORER_LBLEXPLORER, L(247,"EXPLORER"), 10, 12, 75, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT, _
        WS_EX_LEFT Or WS_EX_LTRREADING)

   DIM pImageCtx AS CImageCtx ptr
   pImageCtx = new CImageCtx(pWindow, IDC_FRMEXPLORER_BTNCLOSE, "", 0, 0, 16, 16)
   pImageCtx->SetBkColor(GetSysColor(COLOR_BTNFACE))
   pImageCtx->SetBkColorHot(GetSysColor(COLOR_BTNHIGHLIGHT))
   pImageCtx->LoadImageFromResource (pWindow->InstanceHandle, wstr("IMAGE_CLOSE"))
   AfxAddTooltip(pImageCtx->hWindow, L(161, "Close"))

   Dim As HWnd hTree = _
        pWindow->AddControl("TREEVIEW", , IDC_FRMEXPLORER_TREE, "", 0, 0, 0, 0, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP or _
        TVS_INFOTIP or TVS_SHOWSELALWAYS Or TVS_FULLROWSELECT Or TVS_TRACKSELECT or _
        TVS_HASBUTTONS or TVS_HASLINES or TVS_LINESATROOT, _   
        WS_EX_LEFT Or WS_EX_RIGHTSCROLLBAR, , _
        Cast(SUBCLASSPROC, @frmExplorer_Tree_SubclassProc), IDC_FRMEXPLORER_TREE, Cast(DWORD_PTR, @pWindow))

      ' Use the new style Explorer Treeview (triangles instead of boxes).
      ' Set an undocumented extended style that enables the treeview glyphs to resize
      ' according to the high dpi setting.
      ' https://stackoverflow.com/questions/38772670/ctreectrl-with-explorer-theme-not-dpi-aware
      SendMessage(hTree, TVM_SETEXTENDEDSTYLE, &H1000, &H1000)
      SetWindowTheme(hTree, @wstr("EXPLORER"), 0)

      gApp.hExplorerRootNode = FF_TreeView_InsertItem(hTree, TVI_ROOT, wstr("FILES") )
      Treeview_SetBold(hTree, gApp.hExplorerRootNode, TRUE)

   Dim As HFONT hFont = pWindow->CreateFont("Segoe UI", 9) 
   SendMessage( hTree, WM_SETFONT, Cast(WPARAM, hFont), CTRUE )
   SendMessage( hTree, TVM_SETBKCOLOR, 0, Cast(LPARAM, GetSysColor(COLOR_BTNFACE)) )
   SendMessage( hTree, TVM_SETEXTENDEDSTYLE, TVS_EX_DOUBLEBUFFER, TVS_EX_DOUBLEBUFFER)   

   PositionExplorerWindows(HWND_FRMEXPLORER)

   Function = 0
   
End Function

