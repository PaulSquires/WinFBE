' ########################################################################################
' Microsoft Windows
' File: TreeView.bas
' Compiler: Free Basic
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#define UNICODE
#INCLUDE ONCE "Afx/CWindow.inc"
#INCLUDE ONCE "Afx/AfxGdiPlus.inc"
USING Afx

#define IDC_TREEVIEW 1001

DECLARE FUNCTION WinMain (BYVAL hInstance AS HINSTANCE, _
                          BYVAL hPrevInstance AS HINSTANCE, _
                          BYVAL szCmdLine AS ZSTRING PTR, _
                          BYVAL nCmdShow AS LONG) AS LONG

   END WinMain(GetModuleHandleW(NULL), NULL, COMMAND(), SW_NORMAL)

' // Forward declaration
DECLARE FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

' ========================================================================================
' Main
' ========================================================================================
FUNCTION WinMain (BYVAL hInstance AS HINSTANCE, _
                  BYVAL hPrevInstance AS HINSTANCE, _
                  BYVAL szCmdLine AS ZSTRING PTR, _
                  BYVAL nCmdShow AS LONG) AS LONG

   ' // Set process DPI aware
   ' // The recommended way is to use a manifest
'   AfxSetProcessDPIAware

   ' // Create the main window
   DIM pWindow AS CWindow
   pWindow.Create(NULL, "CWindow with a TreeView", @WndProc)
   ' // Change the window style
   pWindow.WindowStyle = WS_OVERLAPPED OR WS_CAPTION OR WS_SYSMENU
   ' // Change the class style to avoid flicker
   pWindow.ClassStyle = CS_DBLCLKS
   ' // Size it by setting the wanted width and height of its client area
   pWindow.SetClientSize(320, 370)
   ' // Center the window
   pWindow.Center

   ' // Add a TreeView
   DIM hTreeView AS HWND
   hTreeView = pWindow.AddControl("TreeView", , IDC_TREEVIEW, "")
   pWindow.SetWindowPos hTreeView, NULL, 8, 8, 300, 320, SWP_NOZORDER

   ' // To add checkboxes, use:
   ' AfxAddWindowStyle hTreeView, TVS_CHECKBOXES

   ' // Calculate the size of the icon according the DPI
   DIM cx AS LONG = 16 * pWindow.DPI \ 96

   ' // Create an image list for the treeview
   DIM hImageList AS HIMAGELIST
   hImageList = ImageList_Create(cx, cx, ILC_COLOR32 OR ILC_MASK, 3, 0)
   IF hImageList THEN
      AfxGdipAddIconFromRes(hImageList, hInstance, "IDI_FOLDER_CLOSED")
      AfxGdipAddIconFromRes(hImageList, hInstance, "IDI_FOLDER_OPEN")
      AfxGdipAddIconFromRes(hImageList, hInstance, "IDI_ARROW_RIGHT")
      ' // Set the image list
      TreeView_SetImageList(hTreeView, hImageList, TVSIL_NORMAL)
   END IF

   ' // Add items to the TreeView
   DIM AS HTREEITEM hRoot, hNode, hItem
   ' // Create the root node
   hRoot = TreeView_AddRootItem(hTreeView, "Root", , , 2)
   ' // Create a node
   hNode = TreeView_AppendItem(hTreeView, hRoot, "Node 1", , , 2)
   ' // Insert items in the node
   hItem = TreeView_AppendItem(hTreeView, hNode, "Node 1 Item 1", , , 2)
   hItem = TreeView_AppendItem(hTreeView, hNode, "Node 1 Item 2", , , 2)
   hItem = TreeView_AppendItem(hTreeView, hNode, "Node 1 Item 3", , , 2)
   hItem = TreeView_AppendItem(hTreeView, hNode, "Node 1 Item 4", , , 2)
   hItem = TreeView_AppendItem(hTreeView, hNode, "Node 1 Item 5", , , 2)
   hItem = TreeView_AppendItem(hTreeView, hNode, "Node 1 Item 6", , , 2)
   ' // Expand the node
   TreeView_Expand(hTreeView, hNode, TVM_EXPAND)
   ' // Create another node
   hNode = TreeView_AppendItem(hTreeView, hRoot, "Node 2", , , 2)
   ' // Insert items in the node
   hItem = TreeView_AppendItem(hTreeView, hNode, "Node 2 Item 1", , , 2)
   hItem = TreeView_AppendItem(hTreeView, hNode, "Node 2 Item 2", , , 2)
   hItem = TreeView_AppendItem(hTreeView, hNode, "Node 2 Item 3", , , 2)
   hItem = TreeView_AppendItem(hTreeView, hNode, "Node 2 Item 4", , , 2)
   hItem = TreeView_AppendItem(hTreeView, hNode, "Node 2 Item 5", , , 2)
   hItem = TreeView_AppendItem(hTreeView, hNode, "Node 2 Item 6", , , 2)
   ' // Expand the node
   TreeView_Expand(hTreeView, hNode, TVM_EXPAND)
   ' // Create another node
   hNode = TreeView_AppendItem(hTreeView, hRoot, "Node 3", , , 2)
   ' // Insert items in the node
   hItem = TreeView_AppendItem(hTreeView, hNode, "Node 3 Item 1", , , 2)
   hItem = TreeView_AppendItem(hTreeView, hNode, "Node 3 Item 2", , , 2)
   hItem = TreeView_AppendItem(hTreeView, hNode, "Node 3 Item 3", , , 2)
   hItem = TreeView_AppendItem(hTreeView, hNode, "Node 3 Item 4", , , 2)
   hItem = TreeView_AppendItem(hTreeView, hNode, "Node 3 Item 5", , , 2)
   hItem = TreeView_AppendItem(hTreeView, hNode, "Node 3 Item 6", , , 2)
   ' // Expand the node
   TreeView_Expand(hTreeView, hNode, TVM_EXPAND)
   ' // Create another node
   hNode = TreeView_AppendItem(hTreeView, hRoot, "Node 4", , , 2)
   ' // Insert items in the node
   hItem = TreeView_AppendItem(hTreeView, hNode, "Node 4 Item 1", , , 2)
   hItem = TreeView_AppendItem(hTreeView, hNode, "Node 4 Item 2", , , 2)
   hItem = TreeView_AppendItem(hTreeView, hNode, "Node 4 Item 3", , , 2)
   hItem = TreeView_AppendItem(hTreeView, hNode, "Node 4 Item 4", , , 2)
   hItem = TreeView_AppendItem(hTreeView, hNode, "Node 4 Item 5", , , 2)
   hItem = TreeView_AppendItem(hTreeView, hNode, "Node 4 Item 6", , , 2)
   ' // Expand the node
   TreeView_Expand(hTreeView, hNode, TVM_EXPAND)

   ' // Expand the root node
   TreeView_Expand(hTreeView, hRoot, TVE_EXPAND)

   ' // Add a cancel button
   pWindow.AddControl("Button", , IDCANCEL, "&Cancel", 233, 338, 75, 23)

   ' // Dispatch Windows messages
   FUNCTION = pWindow.DoEvents(nCmdShow)

END FUNCTION
' ========================================================================================

' ========================================================================================
' Window procedure
' ================================================================e========================
FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

   SELECT CASE uMsg

      CASE WM_COMMAND
         SELECT CASE GET_WM_COMMAND_ID(wParam, lParam)
            CASE IDCANCEL
               ' // If ESC key pressed, close the application by sending an WM_CLOSE message
               IF GET_WM_COMMAND_CMD(wParam, lParam) = BN_CLICKED THEN
                  SendMessageW hwnd, WM_CLOSE, 0, 0
                  EXIT FUNCTION
               END IF
         END SELECT

      CASE WM_NOTIFY
         DIM ptnmhdr AS NMHDR PTR = CAST(NMHDR PTR, lParam)
         SELECT CASE ptnmhdr->idFrom
            CASE IDC_TREEVIEW
               SELECT CASE ptnmhdr->code
               CASE NM_DBLCLK
                  ' // Retrieve the handle of the TreeView
                  DIM hTreeView AS HWND = GetDlgItem(hwnd, IDC_TREEVIEW)
                  ' // Retrieve the selected item
                  DIM hItem AS HTREEITEM = TreeView_GetSelection(hTreeView)
                  ' // Retrieve the text of the selected item
                  DIM wszText AS WSTRING * 260
                  TreeView_GetItemText(hTreeView, hItem, @wszText, 260)
                  MessageBox hwnd, wszText, "", MB_OK
                  EXIT FUNCTION
               CASE TVN_ITEMEXPANDED
                  ' // Get a pointer to the NMTREEVIEW structure
                  DIM ptnmtv AS NMTREEVIEW PTR = cast(NMTREEVIEW PTR, lParam)
                  ' // Check the item state and select the appropriate image index
                  DIM tvi AS TVITEM = ptnmtv->itemNew
                  IF (ptnmtv->itemNew.state AND TVIS_EXPANDED) = TVIS_EXPANDED THEN
                     tvi.iImage = 1
                  ELSE
                     tvi.iImage = 0
                  END IF
                  ' // Set the item's new attributes
                  TreeView_SetItem(ptnmtv->hdr.hwndFrom, @tvi)
            END SELECT
         END SELECT

    	CASE WM_DESTROY
          ' // Destroy the image list
         DIM hTreeView AS HWND = GetDlgItem(hwnd, IDC_TREEVIEW)
         ImageList_Destroy(TreeView_SetImageList(hTreeView, NULL, TVSIL_NORMAL))
         ' // End the application by sending a WM_QUIT message
         PostQuitMessage(0)
         EXIT FUNCTION

   END SELECT

   ' // Default processing of Windows messages
   FUNCTION = DefWindowProcW(hWnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================
