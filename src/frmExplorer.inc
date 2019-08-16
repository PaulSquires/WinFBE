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

' Drag/drop for Explorer Treeview
dim shared treedragger as DRAG_STATE


' ========================================================================================
' Position all child windows. Called manually and/or by WM_SIZE
' ========================================================================================
private Function frmExplorer_PositionWindows() As LRESULT
   
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(HWND_FRMEXPLORER)
   If pWindow = 0 Then Exit Function
   
   ' Get the entire client area
   Dim As Rect rc
   GetClientRect(HWND_FRMEXPLORER, @rc)

   Dim As HWnd hCtl, hTree
   Dim As Long nLeft, nTop, nWidth, nHeight
   
   ' The Explorer label was painted in the top area so start our treeviews immediately after that
   nWidth  = pWindow->ScaleX(24)
   nHeight = pWindow->ScaleY(24)
   nTop    = pWindow->ScaleY(10)
   
   hCtl = GetDlgItem(HWND_FRMEXPLORER, IDC_FRMEXPLORER_BTNCLOSE)
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
         If IsWindowVisible(HWND_FRMRECENT) = 0 Then frmRecent_LoadTreeview
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
' Get the special subnode based on the filetype
' ========================================================================================
public function GetExplorerSpecialSubNode( byval nFileType as long ) as HTREEITEM
   
   dim hRootNode as HTREEITEM
   
   ' Add the pDoc to the node depending on what file type it is.
   select case nFileType
      case FILETYPE_UNDEFINED
         if ( gApp.IsProjectActive = true ) orelse ( gApp.IsProjectLoading = true ) then
            hRootNode = gApp.hExplorerNormalNode
         else   
            hRootNode = gApp.hExplorerRootNode
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
' Add the names of all subs/functions belonging to this document to the Explorer treeview
' ========================================================================================
public Function AddFunctionsToExplorerTreeview( ByVal pDoc As clsDocument Ptr, _
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

      if pDoc->ProjectFileType = FILETYPE_UNDEFINED then
         if pDoc->IsDesigner then pDoc->ProjectFileType = FILETYPE_NORMAL 
      end if

      if ( gApp.IsProjectActive = true ) orelse ( gApp.IsProjectLoading = true ) then
         ' Give this document a default project type depending on its file extension
         if pDoc->ProjectFileType = FILETYPE_UNDEFINED then
            gApp.ProjectSetFileType( pDoc, pDoc->ProjectFileType)
         end if
      end if

      If pDoc->hNodeExplorer = 0 Then
         hRootNode = GetExplorerSpecialSubNode( pDoc->ProjectFileType )
         pDoc->hNodeExplorer = FF_TreeView_InsertItem( hTree, hRootNode, wzFile, cast(LPARAM, pDoc), 1, 1 )
         frmMain_AddToComboFiles( pDoc )
         fAddChildren = True
      else
         hRootNode = TreeView_GetParent( hTree, pDoc->hNodeExplorer )
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
   ' We never want to sort the ExplorerRootNode if a project is active because it will move
   ' our special subnode branches (Main/Resource/Normal/Module/Header).
   if ( gApp.IsProjectActive = true ) orelse ( gApp.IsProjectLoading = true ) then
      if hRootNode <> gApp.hExplorerRootNode then
         Treeview_SortChildren( hTree, hRootNode, 0 )
      end if   
   else
      Treeview_SortChildren( hTree, hRootNode, 0 )
   end if
      
   If fAddChildren Then
      dim wszText as wstring * MAX_PATH
      dim as HTREEITEM hNode
      dim pData as DB2_DATA ptr
      dim as string sFilename = ucase(pDoc->DiskFilename)
      dim as string sNodeName
      
      gdb2.dbRewind()
      do 
         pData = gdb2.dbGetNext
         if pData = 0 THEN exit do
         select case pData->id 
            case DB2_FUNCTION, DB2_SUB, DB2_PROPERTY
               if sFilename = ucase(pData->fileName) THEN
                  sNodeName = pData->ElementName
                  select case pData->GetSet 
                     case ClassProperty.Getter: sNodeName = sNodeName & " (get)"
                     case ClassProperty.Setter: sNodeName = sNodeName & " (set)"
                     case ClassProperty.ctor:   sNodeName = sNodeName & " (ctor)"
                     case ClassProperty.dtor:   sNodeName = sNodeName & " (dtor)"
                  end select      
                  hNode = FF_TreeView_InsertItem(hTree, pDoc->hNodeExplorer, sNodeName, 0, 2, 2 )
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
      frmExplorer_PositionWindows
   End If
   Function = 0
End Function



' ========================================================================================
' Timer callback to handle scrolling of the Explorer treeview
' ========================================================================================
private function frmExplorer_timerProc( byval hwnd as hwnd, _ 
                                        byval msg as long, _
                                        byval id as long, _
                                        byval ticks as long _
                                         ) as LRESULT
   PostMessage( hwnd, UWM_SCROLL, 0, 0 )
   function = 0
end function


' ========================================================================================
' Logic to determine if we can allow a drop on the specified node
' Testing is called from WM_MOUSEMOVE
' ========================================================================================
private function frmExplorer_AllowDrop( byval target as HTREEITEM, _
                                        byref msgFileType as long _
                                        ) as boolean

   if (target <> null) andalso (target <> treedragger.DragItem ) then
      ' Determine what type of subnode the target drop node belongs to. 
      dim as HTREEITEM parent = TreeView_GetParent( treedragger.hTree, target )
      
      if target = gApp.hExplorerRootNode     then msgFileType = IDM_SETFILENORMAL
      if target = gApp.hExplorerResourceNode then msgFileType = IDM_SETFILERESOURCE
      if target = gApp.hExplorerHeaderNode   then msgFileType = IDM_SETFILEHEADER
      if target = gApp.hExplorerModuleNode   then msgFileType = IDM_SETFILEMODULE
      if target = gApp.hExplorerNormalNode   then msgFileType = IDM_SETFILENORMAL
      if target = gApp.hExplorerMainNode     then msgFileType = IDM_SETFILEMAIN

      if parent = gApp.hExplorerResourceNode then msgFileType = IDM_SETFILERESOURCE
      if parent = gApp.hExplorerHeaderNode   then msgFileType = IDM_SETFILEHEADER
      if parent = gApp.hExplorerModuleNode   then msgFileType = IDM_SETFILEMODULE
      if parent = gApp.hExplorerNormalNode   then msgFileType = IDM_SETFILENORMAL
      if parent = gApp.hExplorerMainNode     then msgFileType = IDM_SETFILEMAIN
      
   end if

   ' Allow the drop if msgFileType has positive value
   function = cast(boolean, msgFileType)
end function


' ========================================================================================
' Send info to TIMER to handle scrolling of the Explorer treeview
' ========================================================================================
private function frmExplorer_AutoScroll( byval hwnd as hwnd, _ 
                                         byval SBcode as long, _
                                         byval scrollspeed as long _
                                         ) as LRESULT
   
   if scrollspeed = 0 then
      ' kill any active timer 
      if treedragger.idTimer then
         KillTimer( hwnd, treedragger.idTimer )
         treedragger.idTimer = 0
      end if
      exit function
   end if
   
   ' create timer based on the new speed
   treedragger.ScrollDirection = SBcode
   
   ' If the hWnd parameter is not NULL and the window specified by hWnd 
   ' already has a timer with the value nIDEvent, then the existing timer 
   ' is replaced by the new timer. When SetTimer replaces a timer, the timer is reset. 
   treedragger.idTimer = SetTimer( hwnd, 1, scrollspeed, _
                                   cast(TIMERPROC, @frmExplorer_timerProc) )
   
   function = 0
end function


' ========================================================================================
' Handle initiating scroll for Explorer treeview
' ========================================================================================
private function frmExplorer_CheckAutoScroll( byval hwnd as hwnd, _ 
                                              byval x as long, _
                                              byval y as long _
                                              ) as LRESULT
   
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(HWND_FRMEXPLORER)
   If pWindow = 0 Then Exit Function

   dim as long ScrollBarHeight
   dim as Boolean hasscroll

   dim as SCROLLINFO si
   dim as long SCROLL_OFF    = 0
   dim as long SCROLL_SLOWLY = 50
   dim as long SCROLL_FASTER = 20

   si.cbSize = sizeof(SCROLLINFO)
   si.fMask = SIF_RANGE
   hasscroll = GetScrollInfo( treedragger.hTree, SB_VERT, @si )
   
   ' GetSystemMetrics is not dpi aware so we need to scale it here
   ScrollBarHeight = pWindow->ScaleY( GetSystemMetrics(SM_CYHSCROLL) ) * 2
   
   if hasscroll andalso (si.nMin <> si.nMax) then
      ' Treeview has a scrollbar, now see if the mouse is in the window
      dim as POINT pt
      dim as RECT rc
      
      pt.x = x
      pt.y = y
      GetClientRect( treedragger.hTree, @rc )
      
      if PtInRect( @rc, pt ) then
         ' We are within the treeview window
         ' See if we are in the scroll zone
         if pt.y < rc.top + ScrollBarHeight then
            ' Scroll up
            frmExplorer_AutoScroll( hwnd, SB_LINEUP, SCROLL_SLOWLY )
         elseif pt.y > rc.bottom - ScrollBarHeight then
            ' Scroll down
            frmExplorer_AutoScroll( hwnd, SB_LINEDOWN, SCROLL_SLOWLY )
         else
            frmExplorer_AutoScroll( hwnd, 0, SCROLL_OFF )
         end if   
      
      else
         ' Outside the treeview window
         if pt.y < rc.top then
            ' fast up
            frmExplorer_AutoScroll( hwnd, SB_LINEUP, SCROLL_FASTER )
         elseif pt.y > rc.bottom then
            ' fast down
            frmExplorer_AutoScroll( hwnd, SB_LINEDOWN, SCROLL_FASTER )
         end if
      end if      

   end if


   function = 0
end function


' ========================================================================================
' Start the drag/drop action for the Explorer treeview
' ========================================================================================
private function frmExplorer_StartDrag( byval hwnd as hwnd, _
                                        byval tvh As NM_TREEVIEW Ptr _
                                        ) as LRESULT
   
   ' Only allow to start a drag on a node that is not the root node
   ' or one of the special subnodes.
   select case tvh->itemNew.hItem
      case gApp.hExplorerRootNode, gApp.hExplorerMainNode, _
           gApp.hExplorerModuleNode, gApp.hExplorerNormalNode, _
           gApp.hExplorerResourceNode, gApp.hExplorerHeaderNode
         ' Returning TRUE will exit the WM_NOTIFY message cracker
         ' and cancel the drag initiation.
         return true
      case else
         ' If the node does not have a valid lParam then it means that 
         ' the node does not have a pDoc and would therefore be a child
         ' node representing a function name. We only want to move
         ' valid pDoc's because they are the filename node.
         if tvh->itemNew.lParam = 0 then return true
   end select
       
   dim as RECT rc
   treedragger.hTree = tvh->hdr.hwndFrom
   treedragger.DragItem = tvh->itemNew.hItem
   treedragger.DragImage = TreeView_CreateDragImage( tvh->hdr.hwndFrom, tvh->itemNew.hItem )
   
   treedragger.pDoc = Cast( clsDocument Ptr, tvh->itemNew.lParam )

   TreeView_SelectItem( tvh->hdr.hwndFrom, null )   ' remove active selections
   
   TreeView_GetItemRect( tvh->hdr.hwndFrom, tvh->itemNew.hItem, @rc, true )  ' text only
   
   ImageList_BeginDrag( treedragger.DragImage, 0, _    ' use first image
                        tvh->ptDrag.x - rc.left, _     ' coords of image hot spot
                        tvh->ptDrag.y - rc.top )
                        
   ImageList_DragEnter( tvh->hdr.hwndFrom, _
                        tvh->ptDrag.x, tvh->ptDrag.y )
                                             
   ShowCursor( false )
   SetCapture( hwnd )
   treedragger.dragging = true
   
   function = 0
end function



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
      
      case TVN_DELETEITEM
         ' A node in the treeview is being removed.
         
      case TVN_BEGINDRAG
         ' Start a drag and drop operation on the Explorer treeview. If
         ' StartDrag returns TRUE then it effectively will cancel the drag
         ' initiation because this message cracker will return TRUE.
         return frmExplorer_StartDrag( hwnd, lpNMTV )
         
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

      case NM_DBLCLK
         hTree = lpNMTV->hdr.hWndFrom
         hItem = Treeview_GetSelection(hTree)
         dim as CWSTR wszFilename, wszFunctionName
         frmExplorer_GetFileNameFunctionName(hItem, wszFilename, wszFunctionName)
         pDoc = OpenSelectedDocument( wszFilename, wszFunctionName )
         TreeView_SelectItem( hTree, hItem )
         ' Get the pDoc and check to see if it is a visual designer. If it is, then
         ' initiate the toggle between the code editor and visual designer.
         if pDoc then
            if pDoc->IsDesigner then
               ' Show the code editor.
               TabCtrl_SetCurSel( pDoc->hDesignTabCtrl, 1 )  ' 1 = code tab
               frmMain_PositionWindows
            end if
         END IF
         return true    ' prevent branch from expanding
         
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
               frmMain_PositionWindows
            end if
            
            If pNMHDR->code = NM_RCLICK Then
               ' Create the popup menu
               lParam = FF_TreeView_GetlParam(hTree, hItem)
               pDoc = Cast(clsDocument ptr, lParam)
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
               frmMain_PositionWindows
            end if
         end if
         
   End Select
 
   Function = 0
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
' Process WM_MOUSEMOVE message for window/dialog: frmExplorer
' ========================================================================================
private Function frmExplorer_OnMouseMove( ByVal HWnd As HWnd, _
                                          ByVal x as long, _
                                          byval y as long, _
                                          byval keyflags as UINT _
                                          ) As Long

   DIM pWindow AS CWindow PTR = AfxCWindowPtr(HWND_FRMMAIN)
   If pWindow = 0 Then Exit Function

   ' These are relative to <hWnd>
   dim as POINT pt = (x, y)

   if treedragger.dragging then
      dim as TV_HITTESTINFO hti
      dim as HTREEITEM target
      
      ' Convert to tree co-ordinates...
      MapWindowPoints( hWnd, treedragger.hTree, cast(POINT ptr, @pt), 1 )

      ImageList_DragMove( pt.x, pt.y )
      
      hti.pt = pt
      target = TreeView_HitTest( treedragger.hTree, @hti )
      
      if frmExplorer_AllowDrop( target, 0 ) = false then
         target = null
      end if
      
      ImageList_DragShowNoLock( false )   ' hide the image
      Treeview_SelectDropTarget( treedragger.hTree, target )   
      ImageList_DragShowNoLock( true )    ' make it visible again
      
      frmExplorer_CheckAutoScroll( hwnd, pt.x, pt.y )
      exit function
      
   else
   
      ' PANELS LEFT/RIGHT SPLITTER
      Dim As Rect rc
      GetWindowRect HWND_FRMEXPLORER, @rc
      rc.Left = rc.Right - pWindow->ScaleX(3)
      GetCursorPos(@pt)
      If PtInRect( @rc, pt ) Then
         if WindowFromPoint(pt) = HWND_FRMEXPLORER then
            SetCursor LoadImage( Null, MAKEINTRESOURCEW(OCR_SIZEWE), IMAGE_CURSOR, 0, 0, LR_DEFAULTSIZE Or LR_SHARED )
         end if
      End If
      If gApp.bDragActive Then
         If gApp.hWndPanel = HWND_FRMEXPLORER Then
            ' Get the current rect of the frmExplorer and compare right side to the current cursor position
            ' and then move the right side to equal the cursor position.
            Dim As Long nDiff = pt.x - rc.Right
            GetClientRect HWND_FRMEXPLORER, @rc
            Dim As Long nWidth = (rc.Right - rc.Left + nDiff)
            SetWindowPos HWND_FRMEXPLORER, 0, 0, 0, nWidth, rc.Bottom - rc.Top, SWP_NOMOVE Or SWP_NOZORDER
            frmMain_PositionWindows
            AfxDoEvents   ' allow screen to repaint
            Exit Function
         End If
      End If
   end if
   
   ' Turn off the autoscroll
   frmExplorer_AutoScroll( hwnd, 0, 0 )
      
   function = 0
end function


' ========================================================================================
' Process WM_LBUTTONDOWN message for window/dialog: frmExplorer
' ========================================================================================
private Function frmExplorer_OnLButtonDown( ByVal HWnd As HWnd, _
                                            byval fDoubleClick as Boolean, _
                                            ByVal x as long, _
                                            byval y as long, _
                                            byval keyflags as UINT _
                                            ) As Long

   DIM pWindow AS CWindow PTR = AfxCWindowPtr(HWND_FRMMAIN)
   If pWindow = 0 Then Exit Function

   ' PANELS LEFT/RIGHT SPLITTER
   Dim As Rect rc
   Dim As Point pt 

   gApp.bDragActive = False 
   GetWindowRect HWND_FRMEXPLORER, @rc
   rc.Left = rc.Right - pWindow->ScaleX(3)
   GetCursorPos(@pt)
   If PtInRect( @rc, pt ) Then
      if WindowFromPoint(pt) = HWND_FRMEXPLORER then
         SetCursor LoadImage( Null, MAKEINTRESOURCEW(OCR_SIZEWE), IMAGE_CURSOR, 0, 0, LR_DEFAULTSIZE Or LR_SHARED )
         gApp.bDragActive = True 
         gApp.hWndPanel   = HWND_FRMEXPLORER
         SetCapture( HWND_FRMEXPLORER )
      end if
   End If

   function = 0
end function


' ========================================================================================
' Process WM_LBUTTONUP message for window/dialog: frmExplorer
' ========================================================================================
private Function frmExplorer_OnLButtonUp( ByVal HWnd As HWnd, _
                                          ByVal x as long, _
                                          byval y as long, _
                                          byval keyflags as UINT _
                                          ) As Long
   if treedragger.dragging then
      ' Handle the drop
      dim as HTREEITEM target
      dim as TV_ITEM tvi
      dim as long msgFileType

      ImageList_EndDrag
      ImageList_DragLeave( treedragger.hTree )
      
      ShowCursor( true )    ' restore mouse cursor
      
      ReleaseCapture
      
      ImageList_Destroy( treedragger.DragImage )
      treedragger.DragImage = null
      
      ' Proceed with the drop...
      target = TreeView_GetDropHilight( treedragger.hTree )
      
      ' Unselect the drop hilight item...
      TreeView_SelectDropTarget( treedragger.hTree , NULL )

      if frmExplorer_AllowDrop( target, msgFileType ) then
         SendMessage( HWND_FRMMAIN, WM_COMMAND, MAKELONG(msgFileType, 0), 0 )
      end if
      
      treedragger.dragging = false
      frmExplorer_AutoScroll( hwnd, 0, 0 )
      AfxRedrawWindow( treedragger.hTree )
      exit function

   elseif gApp.bDragActive then
      gApp.bDragActive = False 
      gApp.hWndPanel = 0
      ReleaseCapture()
   end if
   
   function = 0
end function


' ========================================================================================
' frmExplorer Window procedure
' ========================================================================================
private Function frmExplorer_WndProc( ByVal HWnd   As HWnd, _
                                      ByVal uMsg   As UINT, _
                                      ByVal wParam As WPARAM, _
                                      ByVal lParam As LPARAM _
                                      ) As LRESULT

   Select Case uMsg
      HANDLE_MSG (HWnd, WM_SIZE,        frmExplorer_OnSize)
      HANDLE_MSG (HWnd, WM_NOTIFY,      frmExplorer_OnNotify)
      HANDLE_MSG (HWnd, WM_PAINT,       frmExplorer_OnPaint)
      HANDLE_MSG (HWnd, WM_LBUTTONUP,   frmExplorer_OnLButtonUp)
      HANDLE_MSG (HWnd, WM_LBUTTONDOWN, frmExplorer_OnLButtonDown)
      HANDLE_MSG (HWnd, WM_MOUSEMOVE,   frmExplorer_OnMouseMove)
   
   case WM_DESTROY
      Dim As HWnd hTree = GetDlgItem( hWnd, IDC_FRMEXPLORER_TREE)
      DeleteFont( AfxGetWindowFont(hTree) )
      DIM pImageCtx AS CImageCtx ptr = AfxCImageCtxPtr(hwnd, IDC_FRMEXPLORER_BTNCLOSE)
      if pImageCtx then delete pImageCtx
      
   case WM_ERASEBKGND
      return true

   case WM_CTLCOLORSTATIC
      ' Background color for the Explorer label
      SetBkColor Cast(HDC,wParam), GetSysColor(COLOR_BTNFACE)
      return Cast( LRESULT, GetSysColorBrush(COLOR_BTNFACE))

   case UWM_SCROLL
      ' User defined message to handle slow/fast Explorer treeview scrolling
      dim as POINT pt
      
      ImageList_DragShowNoLock( false )
      SendMessage( treedragger.hTree, WM_VSCROLL, _
                   MAKELONG( treedragger.ScrollDirection, 0 ), _
                   cast(LPARAM, null) )
      ImageList_DragShowNoLock( true )
      
      ' Get mouse position and convert to tree co-ordinates
      GetCursorPos( @pt )
      MapWindowPoints( hWnd, treedragger.hTree, cast(POINT ptr, @pt), 1 )
      frmExplorer_CheckAutoScroll( hwnd, pt.x, pt.y )
      
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
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN or _
        TVS_INFOTIP or TVS_SHOWSELALWAYS Or TVS_FULLROWSELECT Or TVS_TRACKSELECT or _
        TVS_HASBUTTONS or TVS_HASLINES or TVS_LINESATROOT, _   
        WS_EX_LEFT Or WS_EX_RIGHTSCROLLBAR )  
        
   ' Use the new style Explorer Treeview (triangles instead of boxes).
   ' Set an undocumented extended style that enables the treeview glyphs to resize
   ' according to the high dpi setting.
   ' https://stackoverflow.com/questions/38772670/ctreectrl-with-explorer-theme-not-dpi-aware
   TreeView_SetExtendedStyle( hTree, &H1000, &H1000 )
   'SendMessage(hTree, TVM_SETEXTENDEDSTYLE, &H1000, &H1000)
   SetWindowTheme(hTree, @wstr("EXPLORER"), 0)
   TreeView_SetExtendedStyle( hTree, TVS_EX_DOUBLEBUFFER, TVS_EX_DOUBLEBUFFER )

   Dim As HFONT hFont = pWindow->CreateFont("Segoe UI", 9) 
   SendMessage( hTree, WM_SETFONT, Cast(WPARAM, hFont), CTRUE )
   SendMessage( hTree, TVM_SETBKCOLOR, 0, Cast(LPARAM, GetSysColor(COLOR_BTNFACE)) )

   gApp.hExplorerRootNode = FF_TreeView_InsertItem(hTree, TVI_ROOT, wstr("FILES") )
   Treeview_SetBold( hTree, gApp.hExplorerRootNode, TRUE )


   ' Create the imagelist for Explorer drag/drop
   dim as long cx = 16 * (pWindow->DPI \ 96)
   Dim As HANDLE hDragImageList = ImageList_Create(cx, cx, ILC_MASK Or ILC_COLOR32, 3, 0)
   AfxGdipAddIconFromRes( hDragImageList, pWindow->InstanceHandle, "IMAGE_WINFBE")
   AfxGdipAddIconFromRes( hDragImageList, pWindow->InstanceHandle, "IMAGE_DOCUMENT")
   AfxGdipAddIconFromRes( hDragImageList, pWindow->InstanceHandle, "IMAGE_BASETYPE")
   TreeView_SetImageList( hTree, hDragImageList, TVSIL_NORMAL )

   frmExplorer_PositionWindows

   Function = 0
   
End Function

