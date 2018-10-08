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


' ========================================================================================
'  Get Image Index of Treeview 
' ========================================================================================
Function ImageIndex4ExplorerTreeview( ByVal pDoc As clsDocument Ptr ) As long
   dim  As WString * MAX_PATH wzFileExt  = any
   DIM AS LONG iImage=II4E_FILE4NORMAL
   SELECT CASE pDoc->ProjectFileType
   CASE FILETYPE_MAIN
        iImage=II4E_FILE4MAINBAS
   CASE FILETYPE_MODULE
        iImage=II4E_FILE4MODULEBAS
   CASE FILETYPE_RESOURCE
        iImage=II4E_FILE4RESOURCE
   'CASE FILETYPE_NORMAL
   'CASE FILETYPE_UNDEFINED
   CASE ELSE
        wzFileExt=AfxStrPathname( "EXTN",  pDoc->DiskFilename) 
        Select Case Ucase(wzFileExt)
            Case ".BAS":  iImage = II4E_FILE4BAS
            Case ".RC":   if(pDoc->ProjectFileType<> FILETYPE_NORMAL) then  iImage = II4E_FILE4RESOURCE
            Case ".FRM":  iImage = II4E_FILE4FRM
            Case ".INC":  iImage = II4E_FILE4INC
            Case ".BI":   iImage = II4E_FILE4BI
            Case ".TXT":  iImage = II4E_FILE4TXT
            Case ".MASM": iImage = II4E_FILE4MASM
        End Select
   END SELECT
   
   Function = iImage
End Function

' ========================================================================================
'  is a Document ptr ?
' ========================================================================================
function CheckIsDoc4ExplorerTreeview( ByVal hWndControl As HWnd, byVal hNode As HTREEITEM ) as Boolean
   DIM item AS TVITEMW
   with item
      .hitem = hNode
      .mask =  TVIF_IMAGE            
      TreeView_GetItem(hWndControl, @item)
      
      IF .iImage<=II4E_FLAGS4MARK and .iImage>II4E_PROJECT then 
         return true
      end if            
   end with
   return false
End function

' ========================================================================================
' 
' ========================================================================================
function GetProjecthNode4ExplorerTreeview( ByVal hWndControl As HWnd, byVal hNode As HTREEITEM ) as HTREEITEM
   DIM item AS TVITEMW
   dim as Project_DATA ptr  pProjectDATA=any
   dim as DB2_DATA ptr  pDB2_DATA=any
   with item
      .mask =  TVIF_IMAGE      
      while hNode<>0
         .hitem = hNode            
         TreeView_GetItem(hWndControl, @item)
         if .iImage=II4E_PROJECT then
            RETURN hNode
         end if
         hNode = TreeView_GetParent(hWndControl,hNode)
      wend            
   end with
   function = 0
End function

' ========================================================================================
' 
' ========================================================================================
function GetProjectIndex4ExplorerTreeview( ByVal hWndControl As HWnd, byVal hNode As HTREEITEM ) as long
   DIM item AS TVITEMW
   dim as Project_DATA ptr  pProjectDATA=any
   dim as DB2_DATA ptr  pDB2_DATA=any
   with item
      .mask =  TVIF_IMAGE or LVIF_PARAM      
      while hNode<>0
         .hitem = hNode            
         TreeView_GetItem(hWndControl, @item)
         if .lParam then
            select case .iImage
               case is> II4E_FLAGS4MARK
                   pDB2_DATA= cast(DB2_DATA ptr , .lParam)
                   if pDB2_DATA->deleted=false then
                      pProjectDATA =pDB2_DATA->pProjectDATA
                     if pProjectDATA then                   
                        return pProjectDATA->ProjIndex
                     endif 
                  end if
               case is> II4E_PROJECT
                  if gApp.DocumentPtrExists(cast(clsDocument ptr , .lParam))=true then
                     return cast(clsDocument ptr , .lParam)->ProjectIndex
                  else
                     exit while  
                  endif
               case else
                  exit while
            end select
         end if
         hNode = TreeView_GetParent(hWndControl,hNode)
      wend            
   end with
   function = -1
End function
' ========================================================================================
' 
' ========================================================================================
function GetClsDocument4ExplorerTreeview( ByVal hWndControl As HWnd, byVal hNode As HTREEITEM ) as clsDocument ptr
   DIM item AS TVITEMW
   dim as Project_DATA ptr  pProjectDATA=any
   dim as DB2_DATA ptr  pDB2_DATA=any
   with item
      .mask =  TVIF_IMAGE or LVIF_PARAM      
      while hNode<>0
         .hitem = hNode            
         TreeView_GetItem(hWndControl, @item)
         if .lParam then
            select case .iImage
               case is> II4E_FLAGS4MARK
               case is> II4E_PROJECT
                  if gApp.DocumentPtrExists(cast(clsDocument ptr , .lParam))=true then
                     return cast(clsDocument ptr , .lParam)
                  else
                     exit while  
                  endif
               case else
                  exit while
            end select
         end if
         hNode = TreeView_GetParent(hWndControl,hNode)
      wend
   end with
   function = 0
End function

' ========================================================================================
' Change Image Index of Treeview  for  Document
' ========================================================================================
sub ChangeImage4ExplorerTreeview2doc( ByVal pDoc As clsDocument Ptr )
   Dim ti As TV_ITEM    
   Dim As Hwnd hTree = GetDlgItem(HWND_FRMEXPLORER, IDC_FRMEXPLORER_TREE) 
   ti.hItem = pDoc->hNodeExplorer
   ti.mask=TVIF_SELECTEDIMAGE Or TVIF_IMAGE 
   ti.iImage=ImageIndex4ExplorerTreeview(pDoc)
   ti.iSelectedImage=ti.iImage
   TreeView_SetItem(hTree,Varptr(ti))
End sub

' ========================================================================================
' Change Image Index of Treeview   for database
' ========================================================================================
sub ChangeImage4ExplorerTreeview2db(ByVal hWndControl As HWnd, _
                                 ByVal hNode As HTREEITEM, _
                                 ByVal DB2ID as LONG )
   Dim ti As TV_ITEM  
   with ti
      .hItem = hNode
      .mask=TVIF_SELECTEDIMAGE Or TVIF_IMAGE
      select case DB2ID
         case DB2_STANDARDDATATYPE,DB2_TYPE
            .iImage=II4E_TYPE4TYPE
            .iSelectedImage=II4E_TYPE4TYPE2
         case DB2_FUNCTION
            .iImage=II4E_TYPE4FUNCTION
            .iSelectedImage=II4E_TYPE4FUNCTION2
         case DB2_SUB  
            .iImage=II4E_TYPE4SUB
            .iSelectedImage=II4E_TYPE4SUB2
         case DB2_PROPERTY
            .iImage=II4E_TYPE4PROPERTY
            .iSelectedImage=II4E_TYPE4PROPERTY2
         case DB2_VARIABLE 
            .iImage=II4E_TYPE4VARIABLE
            .iSelectedImage=II4E_TYPE4VARIABLE2
         case DB2_SUBTYPE4ENUM 
            .iImage=II4E_TYPE4ENUM
            .iSelectedImage=II4E_TYPE4ENUM2
         case DB2_CLASS 
            .iImage=II4E_TYPE4CLASS
            .iSelectedImage=II4E_TYPE4CLASS2
      end select
   end with
   TreeView_SetItem(hWndControl,Varptr(ti))
End sub

' ========================================================================================
' Get the location level of the current Node  of ExplorerTreeview
' ========================================================================================
function GetLevels4ExplorerTreeview(ByVal hWndControl As HWnd, _
                               ByVal hCurrentNode As HTREEITEM) as long
      dim  as long Index=0 
      DIM item AS TVITEMW

      WITH item
        .mask = TVIF_IMAGE
        while hCurrentNode 
            .hitem = hCurrentNode       
            TreeView_GetItem(hWndControl, @item)
            if .iImage<=II4E_FLAGS4MARK THEN
                exit While     
            end if
            Index+=1        
            hCurrentNode=TreeView_GetParent(hWndControl, hCurrentNode)
         wend 
      end with
      return Index 
 
end function

' ========================================================================================
' Get the location Index of the current Node  of ExplorerTreeview
' ========================================================================================
function GetPrevIndexOfPrevSibling4Treeview(ByVal hWndControl As HWnd, _
                               ByVal hCurrentNode As HTREEITEM) as long
      DIM item AS TVITEMW 
      dim as long Index=0 ,iImage=any
      DIM AS WString* MAX_PATH wszText="",wszName =""    
      
      WITH  item
            .mask = TVIF_TEXT Or TVIF_IMAGE
            .hitem = hCurrentNode
            .pszText = @wszText
            .cchTextMax = MAX_PATH
            TreeView_GetItem(hWndControl, @item) 
             
            iImage=.iImage
            .pszText = @wszName   
           while hCurrentNode
               .hitem = hCurrentNode
               TreeView_GetItem(hWndControl, @item)
               if WStrIsEqual(wszText,wszName,false)=true and iImage=.iImage THEN
                  Index+=1     
               end if        
               hCurrentNode=TreeView_GetPrevSibling(hWndControl, hCurrentNode)
            wend 
      end with
      function = Index 
end function

' ========================================================================================
' Create Name of Treeview , according to the data types.
' ========================================================================================
function CreateName4ExplorerTreeview(ByVal hWndControl As HWnd, _
                               byref ReturnValue as WSTRing , _
                               ByVal hCurrentNode As HTREEITEM =0) as LONG
     DIM item AS TVITEMW 
     DIM AS WString* MAX_PATH wszText="",wszName =""
     DIM AS LONG Index=GetPrevIndexOfPrevSibling4Treeview(hWndControl, hCurrentNode)
     dim AS ImageIndex4Explorer ImageIndex=FF_TreeView_GetiImage(hWndControl,hCurrentNode)
     TreeView_GetItemText(hWndControl, hCurrentNode, @wszName, MAX_PATH)
     select case GetLevels4ExplorerTreeview(hWndControl,hCurrentNode)
           case 1
               select case ImageIndex  
               case II4E_TYPE4VARIABLE, II4E_TYPE4VARIABLE2
                  wszName = ".." & wszName
               end select 
           case 2    
               WITH  item
                  hCurrentNode=TreeView_GetParent(hWndControl, hCurrentNode)
                  .mask = TVIF_TEXT Or TVIF_IMAGE
                  .hitem = hCurrentNode
                  .pszText = @wszText
                  .cchTextMax = MAX_PATH
                  TreeView_GetItem(hWndControl, @item)
                  select case ImageIndex  
                  case II4E_TYPE4VARIABLE, II4E_TYPE4VARIABLE2
                     select case .iImage
                     case II4E_TYPE4FUNCTION, II4E_TYPE4FUNCTION2,  II4E_TYPE4SUB, II4E_TYPE4SUB2, _
                          II4E_TYPE4PROPERTY, II4E_TYPE4PROPERTY2
                          wszName = "." & wszText & "." & wszName
                     case else ' II4E_TYPE4TYPE, II4E_TYPE4TYPE2, II4E_TYPE4CLASS, II4E_TYPE4CLASS2
                          wszName = wszText & ".." & wszName
                     end select 
                  case II4E_TYPE4FUNCTION, II4E_TYPE4FUNCTION2,  II4E_TYPE4SUB, II4E_TYPE4SUB2, _
                     II4E_TYPE4PROPERTY, II4E_TYPE4PROPERTY2
                     wszName = wszText & "." & wszName
                  case else
                     MessageBox NULL, wszText & "?" & wszName,"Unknown(CreateName4ExplorerTreeview)", MB_OK
                  end select
               end with
           case 3
               for I2 as long=1 to 2
                 hCurrentNode=TreeView_GetParent(hWndControl, hCurrentNode)
                 TreeView_GetItemText(hWndControl, hCurrentNode, @wszText, MAX_PATH) 
                 wszName = wszText & "." & wszName 
               next
     end select

   ReturnValue=wszName
   function=Index

end function

' ========================================================================================
' Position all child windows. Called manually and/or by WM_SIZE
' ========================================================================================
Function PositionExplorerWindows( ByVal HWnd As HWnd ) As LRESULT
   
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
   nTop    = pWindow->ScaleY(8)
   
   hCtl = GetDlgItem(HWND, IDC_FRMEXPLORER_BTNCLOSE)
   SetWindowPos( hCtl, 0, rc.Right - (nWidth*1.2), nTop, nWidth, nHeight, SWP_SHOWWINDOW Or SWP_NOZORDER )

   nTop = nTop + nHeight + pWindow->ScaleY(12)

   hTree = GetDlgItem(HWND_FRMEXPLORER, IDC_FRMEXPLORER_TREE)
   
   If gApp.IsProjectActive Then
      ' Project is active to continue to show the Explorer tree regardless of whether 
      ' any files are actually open.
      ShowWindow HWND_FRMRECENT, SW_HIDE
      ShowWindow hTree, SW_SHOW
   else
      ' If no project open then show Recent dialog
      if gApp.Projects(0).GetDocumentCount = 0 then
         ' Ensure that the Recent lists are the most current
         If IsWindowVisible(HWND_FRMRECENT) = 0 Then LoadRecentFilesTreeview(HWND_FRMRECENT)
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
' Find Item of Treeview
' ========================================================================================
Function FindTreeViewItem( ByVal hWndControl As HWnd, _
                                 ByVal hParent As HANDLE, _
                                 ByRef TheText As WString, _
                                 ByRef SearchItem4Last As boolean=false _
                                 ) As HTREEITEM
      DIM AS WString* MAX_PATH wszText
      Dim As HTREEITEM hNode=TreeView_GetChild(hWndControl, hParent)
      Dim As HTREEITEM tmphNode=0
      do while hNode<>0
         TreeView_GetItemText(hWndControl, hNode, @wszText, MAX_PATH)
         if WStrIsEqual(wszText,TheText,false) then 
            tmphNode= hNode
            if SearchItem4Last=false then                  
               EXIT DO
            end if
         end if
         hNode=TreeView_GetNextSibling(hWndControl,hNode)
      loop
      Function = tmphNode
End Function

' ========================================================================================
' Update Item of Treeview
' ========================================================================================
Function UpdateItem4ExploreTreeview( ByVal hWndControl As HWnd, _
                                 ByVal hParent As HANDLE, _
                                 Byval pData as DB2_DATA ptr, _
                                 ByVal lParam As LPARAM = 0 _
                                 ) As HANDLE
      DIM AS LONG INDEX,INDEX2
      Dim As HTREEITEM hNode
      dim as LONG DB2ID=pData->id
      DIM AS WString* MAX_PATH wzText(0 to 2)
      DIM AS WString* MAX_PATH TheText=pData->ElementName
      dim as ImageIndex4Explorer TYPE4CLASS,TYPE4CLASS2
      for index=2 to 0 step -1 
         if TheText="" then exit for     
         INDEX2=instrRev(TheText,".") 
         if INDEX2=0 then
            wzText(index)=TheText
            exit for
         else
            wzText(index)=mid(TheText,INDEX2+1)
            TheText=left(TheText,INDEX2-1) 
         end if
      next

      select CASE DB2ID
         case DB2_STANDARDDATATYPE
               
         case DB2_TYPE
               hNode =FindTreeViewItem(hWndControl, hParent, wzText(2), false) 
               if hNode=0 then
                  hNode=FF_TreeView_InsertItem(hWndControl, hParent, wzText(2),lParam,II4E_TYPE4CLASS,II4E_TYPE4CLASS2)
                  Treeview_RemoveCheckBox(hWndControl, hNode) 
               endif
               if pData->SubType = DB2_SUBTYPE4ENUM then
                  ChangeImage4ExplorerTreeview2db(hWndControl,hNode,DB2_SUBTYPE4ENUM)
               else 
                  ChangeImage4ExplorerTreeview2db(hWndControl,hNode,DB2_CLASS)
               endif
                
         case DB2_FUNCTION,DB2_SUB,DB2_PROPERTY 
            if wzText(1)<>"" then      'Type
               hNode =FindTreeViewItem(hWndControl, hParent, wzText(1), false) 
               if hNode=0 then
                  hNode=FF_TreeView_InsertItem(hWndControl, hParent, wzText(1),0,II4E_TYPE4CLASS,II4E_TYPE4CLASS2)
                  Treeview_RemoveCheckBox(hWndControl, hNode)  
               else
                  ChangeImage4ExplorerTreeview2db(hWndControl,hNode,DB2_CLASS)    
               endif
               hParent=hNode                
            end if             
            if wzText(2)<>"" then     'FUNCTION,SUB,PROPERTY etc.
               'hNode =FindTreeViewItem(hWndControl, hParent, wzText(2), false) 
               'if hNode=0 then
                  hNode=FF_TreeView_InsertItem(hWndControl, hParent, wzText(2),lParam,II4E_TYPE4FUNCTION,II4E_TYPE4FUNCTION2)
                  Treeview_RemoveCheckBox(hWndControl, hNode)
               'endif
               ChangeImage4ExplorerTreeview2db(hWndControl,hNode,DB2ID)              
            end if
         
         case DB2_VARIABLE
            if wzText(0)<>"" then      'Type
               hNode =FindTreeViewItem(hWndControl, hParent, wzText(0), false) 
               if hNode=0 then
                  hNode=FF_TreeView_InsertItem(hWndControl, hParent, wzText(0),0,II4E_TYPE4CLASS,II4E_TYPE4CLASS2)
                  Treeview_RemoveCheckBox(hWndControl, hNode) 
               endif
               'if pData->SubType = DB2_SUBTYPE4ENUM then
               '   ChangeImage4ExplorerTreeview2db(hWndControl,hNode,DB2_SUBTYPE4ENUM)
               'else 
               '   ChangeImage4ExplorerTreeview2db(hWndControl,hNode,DB2_CLASS)
               'endif
               hParent=hNode                
            end if    

            if wzText(1)<>"" then    'FUNCTION,SUB,PROPERTY etc.
               hNode =FindTreeViewItem(hWndControl, hParent, wzText(1), true) 
               if hNode=0 then
                  hNode=FF_TreeView_InsertItem(hWndControl, hParent, wzText(1),0,II4E_TYPE4FUNCTION,II4E_TYPE4FUNCTION2)
                  Treeview_RemoveCheckBox(hWndControl, hNode)
               endif
               hParent=hNode                
            end if

            if wzText(2)<>"" then     'VARIABLE
               hNode=FF_TreeView_InsertItem(hWndControl, hParent, wzText(2),lParam,II4E_TYPE4VARIABLE,II4E_TYPE4VARIABLE2)
               Treeview_RemoveCheckBox(hWndControl, hNode)              
            end if
      END SELECT
      return  hNode
End Function

function UpdateDATAToExplorerTreeview ( byval hTree As Hwnd, ByVal pDoc As clsDocument Ptr, byval Firstptrdb as DB2_DATA ptr ) as boolean
   'ELF_CurrentFile
   if pDoc<>0 then
      while Firstptrdb<>0
         select CASE Firstptrdb->id 
         case DB2_FUNCTION,DB2_SUB,DB2_PROPERTY,DB2_VARIABLE,DB2_TYPE
              UpdateItem4ExploreTreeview(hTree, pDoc->hNodeExplorer,Firstptrdb,cast(LPARAM,Firstptrdb))
         END SELECT
         if Firstptrdb->ChildNode4DB then
            UpdateDATAToExplorerTreeview(hTree, pDoc, Firstptrdb->ChildNode4DB)
         END IF 
         Firstptrdb = Firstptrdb->NextNode
      wend
      function =true
   endif
END FUNCTION

' ========================================================================================
' Add the names of all subs/functions belonging to this document to the Explorer treeview
' ========================================================================================
Function AddFunctionsToExplorerTreeview( ByVal pDoc As clsDocument Ptr, ByVal fUpdateNodes As BOOLEAN ) As Long
 
   If pDoc = 0 Then Exit Function
   
   Dim As CWSTR wzFile = AfxStrPathname("NAMEX", pDoc->DiskFilename)
   Dim As Hwnd hTree   = GetDlgItem(HWND_FRMEXPLORER, IDC_FRMEXPLORER_TREE)
   Dim As HTREEITEM hChild, hRootNode
   Dim As BOOLEAN fAddChildren = False
   dim as long idx
   
   ' Determine what root node to add or update the files from
   idx = gApp.GetActiveProjectIndex()
   hRootNode = gApp.Projects(idx).hExplorerRootNode
   
   ' If hNodeExplorer does not exist then we must be adding a new document. If it does
   ' exist then we simply need to update the child nodes.
   If fUpdateNodes = False Then
      If pDoc->hNodeExplorer = 0 Then
         dim As Long iImage=ImageIndex4ExplorerTreeview(pDoc)
         pDoc->hNodeExplorer = FF_TreeView_InsertItem (hTree, hRootNode, wzFile, cast(LPARAM, pDoc),iImage,iImage)
         fAddChildren = True
         Treeview_RemoveCheckBox(hTree, pDoc->hNodeExplorer)
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
      dim as DB2_DATA ptr pData = any 
      gdb2.dbRewind4File(pDoc->ProjectIndex,pDoc->DiskFilename)
      pData = gdb2.dbRewind4DB()
      UpdateDATAToExplorerTreeview(hTree, pDoc, pData)
   
      Treeview_SortChildren(hTree, pDoc->hNodeExplorer, 0)
   End If
   AfxRedrawWindow(hTree)
   
   Function = 0
End Function


' ========================================================================================
' Process WM_SIZE message for window/dialog: frmExplorer
' ========================================================================================
Function frmExplorer_OnSize(ByVal HWnd As HWnd, ByVal state As UINT, ByVal cx As Long, ByVal cy As Long) As LRESULT
   If state <> SIZE_MINIMIZED Then
      ' Position all of the child windows
      PositionExplorerWindows(HWnd)
   End If
   Function = 0
End Function


' ========================================================================================
' Process WM_COMMAND message for window/dialog: frmExplorer
' ========================================================================================
Function frmExplorer_OnCommand(ByVal HWnd As HWnd, ByVal id As Long, ByVal hwndCtl As HWnd, ByVal codeNotify As UINT) As LRESULT

   Select case codeNotify
      case BN_CLICKED
         If id = IDC_FRMEXPLORER_BTNCLOSE Then
            ShowWindow HWND_FRMEXPLORER, SW_HIDE
            frmMain_PositionWindows(HWND_FRMMAIN)
         end if
   end select
  
   Function = 0
End Function


' ========================================================================================
' Process WM_NOTIFY message for window/dialog: frmExplorer
' ========================================================================================
Function frmExplorer_OnNotify(ByVal HWnd As HWnd, ByVal id As Long, ByVal pNMHDR As NMHDR Ptr) As LRESULT
 
   If (id <> IDC_FRMEXPLORER_TREE) Then Return 0
   
   Dim lpNMTV As NM_TREEVIEW Ptr = Cast(NM_TREEVIEW Ptr, pNMHDR)
   Dim As HTREEITEM hItem
   Dim As HWND hTree 
   Dim As LPARAM lParam
   Dim pDoc As clsDocument Ptr
   Dim ht As TVHITTESTINFO
   Dim As POINT pt
   static As WString * MAX_PATH ElementValue
   ' Determine what node was clicked on
   Select Case pNMHDR->code 
      
      case TVN_GETINFOTIP
         ' Display a tooltip when a treeview node is hovered over. This will normally
         ' be the full filename for the project or document.
         dim pnmTip as NMTVGETINFOTIPW ptr = cast(NMTVGETINFOTIPW ptr, pNMHDR)
         ' First, search if this tooltip belongs to a project root node
         for i as long = 1 to ubound(gApp.Projects)
            if gApp.Projects(i).InUse = true THEN
               if gApp.Projects(i).hExplorerRootNode = pnmTip->hItem then
                  pnmTip->pszText = cast(WSTRING Ptr, @**gApp.Projects(i).ProjectFilename)
                  pnmTip->cchTextMax = len(gApp.Projects(i).ProjectFilename)
                  exit function
               end if   
            END IF
         NEXT
         ' Second, search for a matching document
         if pnmTip->lParam then
            Dim As Hwnd hTree = GetDlgItem(HWND_FRMEXPLORER, IDC_FRMEXPLORER_TREE)
            if CheckIsDoc4ExplorerTreeview(hTree,pnmTip->hItem) then
               pDoc = Cast(clsDocument ptr, pnmTip->lParam)
               pnmTip->pszText = @pDoc->DiskFilename
               pnmTip->cchTextMax = len(pDoc->DiskFilename)
            else 
               ElementValue = Cast(DB2_DATA ptr, pnmTip->lParam)->ElementValue
               pnmTip->pszText = @ElementValue
               pnmTip->cchTextMax = len(ElementValue)
            endif
          endif
         exit function   

      Case NM_CLICK, NM_RCLICK
      ' If we have clicked on a folder item then issue an expand. Must use HitTest because the 
      ' the current node has not been selected yet.
         hTree = lpNMTV->hdr.hWndFrom
         GetCursorPos(@pt)  
         ScreenToClient(hTree, @pt)
         
         ht.pt = pt 
         Treeview_HitTest(hTree, @ht )
         hItem = ht.hItem
         If hItem Then 
            Treeview_Select( hTree, hItem, TVGN_CARET)
            pDoc=GetClsDocument4ExplorerTreeview(hTree, hItem)
            if pDoc then
               gdb2.dbRewind4File(gApp.SetActiveProject(hItem),pDoc->DiskFilename)
            else
               gdb2.dbRewind4Proj(gApp.SetActiveProject(hItem))
            endif
         end if
         
         ' Was the checkbox clicked on. If yes then post user messge to select/deselect
         ' the other checkboxes for the other projects (only one can be active).
         if (ht.flags and TVHT_ONITEMSTATEICON) THEN
            PostMessage( HWND, MSG_USER_TOGGLE_TVCHECKBOXES, 0, Cast(LPARAM, ht.hItem))
         END IF

         If pNMHDR->code = NM_CLICK Then
            ' Ensure that the correct Project name is showing
            frmMain_PositionWindows(HWND_FRMMAIN)
         end if
         
         If pNMHDR->code = NM_RCLICK Then
            ' Create the popup menu
            IF CheckIsDoc4ExplorerTreeview(hTree, hItem) then
               lParam = FF_TreeView_GetlParam(hTree, hItem) 
               pDoc = Cast(clsDocument Ptr, lParam)
            end if
            ' Select the node and show the document
            if pDoc then SendMessage(hTree, WM_LBUTTONDBLCLK, 0, 0)
            Dim As HMENU hPopupMenu = CreateExplorerContextMenu(pDoc)
            GetCursorPos @pt
            TrackPopupMenu(hPopUpMenu, 0, pt.x, pt.y, 0, HWND_FRMMAIN, ByVal Null)
            DestroyMenu hPopUpMenu
            Function = True
         End If

         
'      Case NM_CUSTOMDRAW
'         Dim tvcd As NMTVCUSTOMDRAW Ptr = Cast(NMTVCUSTOMDRAW Ptr, pNMHDR)

'         Select Case tvcd->nmcd.dwDrawStage
'            Case CDDS_PREPAINT
'               Return CDRF_NOTIFYITEMDRAW 

'            Case CDDS_ITEMPREPAINT 
'               Dim As HFONT hFont = AfxGetWindowFont(GetDlgItem(hWnd, id))
'               SelectFont(tvcd->nmcd.hdc, hFont)
''               If (tvcd->nmcd.uItemState And ODS_SELECTED) And (GetFocus() = tvcd->nmcd.hdr.hwndFrom) Then
''                  tvcd->clrTextBk = gTheme.exlistbox_selected_active_bg
''                  tvcd->clrText   = gTheme.exlistbox_selected_active_fg
''               ElseIf (tvcd->nmcd.uItemState And ODS_SELECTED) Then
''                  tvcd->clrTextBk = gTheme.exlistbox_selected_inactive_bg
''                  tvcd->clrText   = gTheme.exlistbox_selected_inactive_fg
''               ElseIf (tvcd->nmcd.uItemState And CDIS_HOT) Then
''                  tvcd->clrTextBk = gTheme.exlistbox_hot_bg
''                  tvcd->clrText   = gTheme.exlistbox_hot_fg
''               Else
'                  tvcd->clrText   = gConfig.Themes(gConfig.idxTheme).colors(CLR_PANELTEXT).nFg
'                  tvcd->clrTextBk = gConfig.Themes(gConfig.idxTheme).colors(CLR_PANELTEXT).nBg
''               End If
                                    
'               Return CDRF_NEWFONT 
'
'         End Select

   End Select
 
   Function = 0
End Function


' ========================================================================================
' Processes messages for the subclassed frmExplorer treeview control.
' ========================================================================================
Function frmExplorer_Tree_SubclassProc ( _
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

         dim wszText as wstring * MAX_PATH
         Dim pDoc As clsDocument Ptr
         DIM item AS TVITEMW

         IF CheckIsDoc4ExplorerTreeview(hTree, hitem) then
            lParam = FF_TreeView_GetlParam(hTree, hitem)
            pDoc = Cast(clsDocument Ptr, lParam)
            gdb2.dbRewind4File(gApp.SetActiveProject(hitem),pDoc->DiskFilename)
         else         
            gdb2.dbRewind4Proj(gApp.SetActiveProject(hitem)) 
         end if
         TreeView_Expand(hTree, hitem,TVE_TOGGLE)
         TreeView_GetItemText(hTree, hitem, @wszText, MAX_PATH)
         OpenSelectedDocument(pDoc, wszText, hItem)
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
Function frmExplorer_OnPaint( ByVal HWnd As HWnd) As LRESULT
            
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
Function frmExplorer_WndProc( ByVal HWnd   As HWnd, _
                              ByVal uMsg   As UINT, _
                              ByVal wParam As WPARAM, _
                              ByVal lParam As LPARAM _
                              ) As LRESULT

   Select Case uMsg
      HANDLE_MSG (HWnd, WM_SIZE,    frmExplorer_OnSize)
      HANDLE_MSG (HWnd, WM_COMMAND, frmExplorer_OnCommand)
      HANDLE_MSG (HWnd, WM_NOTIFY,  frmExplorer_OnNotify)
      HANDLE_MSG (HWnd, WM_PAINT,   frmExplorer_OnPaint)
   
      case WM_DESTROY
         Dim As HWnd hTree = GetDlgItem( hWnd, IDC_FRMEXPLORER_TREE)
         dim as HFONT hFont = AfxGetWindowFont(hTree)
         DeleteFont(hFont)
         ImageList_Destroy(gApp.himl)
         
      case WM_ERASEBKGND
         return true

      case WM_CTLCOLORSTATIC
         ' Background color for the Explorer label
         SetBkColor Cast(HDC,wParam), GetSysColor(COLOR_BTNFACE)
         return Cast( LRESULT, GetSysColorBrush(COLOR_BTNFACE))

      case MSG_USER_TOGGLE_TVCHECKBOXES      
         ' If this is a project root node then ensure that only one of the projects
         ' root node has the checkmark box ticked. If there is only one project
         ' then do not allow it to be unchecked.
         dim as hwnd hTree = GetDlgItem(HWND_FRMEXPLORER, IDC_FRMEXPLORER_TREE)
         dim as HTREEITEM hNode = Cast(HTREEITEM, lParam)
         if gApp.GetProjectCount = 1 THEN
            gApp.SetActiveProject(hNode)
         else         
            if TreeView_IsItemChecked(hTree, hNode) then
               gApp.SetActiveProject(hNode)
            else
               gApp.EnsureDefaultActiveProject(hNode)
            end if   
         END IF

   End Select

   ' for messages that we don't deal with
   Function = DefWindowProc(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' frmExplorer_Show
' ========================================================================================
Function frmExplorer_Show( ByVal hWndParent As HWnd ) As LRESULT

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
   'pWindow->Brush = GetSysColorBrush(COLOR_BTNFACE)

   Dim As HWnd hBtn = _
   pWindow->AddControl("LABEL", , IDC_FRMEXPLORER_LBLEXPLORER, L(247,"EXPLORER"), 10, 12, 75, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT, _
        WS_EX_LEFT Or WS_EX_LTRREADING)

   frmHotImgBtn( HWND_FRMEXPLORER, IDC_FRMEXPLORER_BTNCLOSE, "IMAGE_CLOSE", "IMAGE_CLOSE",  _
                  L(161, "Close"), 24, 24, _
                  GetSysColor(COLOR_BTNFACE), _
                  GetSysColor(COLOR_WINDOW), _
                  0, 0, 24, 24, true, true, true )

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

      ' items that are initially checked will appear checked (must add TVS_CHECKBOXES after control is created)
      SetWindowLongPtr( hTree, GWL_STYLE, GetWindowLongPtr(hTree, GWL_STYLE) OR TVS_CHECKBOXES )
  
   Dim As HFONT hFont = pWindow->CreateFont("Segoe UI", 9) 
   SendMessage( hTree, WM_SETFONT, Cast(WPARAM, hFont), CTRUE )
   SendMessage( hTree, TVM_SETBKCOLOR, 0, Cast(LPARAM, GetSysColor(COLOR_BTNFACE)) )
'   SendMessage( hTree, TVM_SETITEMHEIGHT, pWindow->ScaleY(20), 0 )
   SendMessage( hTree, TVM_SETEXTENDEDSTYLE, TVS_EX_DOUBLEBUFFER, TVS_EX_DOUBLEBUFFER)   

   ' Create the imagelist  
   Dim hBitmap As HBITMAP
   DIM cx AS LONG = 16 * pWindow->DPI \ 96
   gApp.himl=ImageList_Create(cx,cx,ILC_MASK Or ILC_COLOR8,32,0) 
   hBitmap = LoadImage(pWindow->InstanceHandle, "IMAGE_FILES", IMAGE_BITMAP, 0, 0, LR_SHARED)
   ImageList_AddMasked(gApp.himl,hBitmap,&HFFFFFF)
   DeleteObject(hBitmap)
   TreeView_SetImageList(hTree,gApp.himl,TVSIL_NORMAL)
   PositionExplorerWindows(HWND_FRMEXPLORER)

   Function = 0
   
End Function

