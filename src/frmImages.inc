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
' Display a selected image in the Image Manager
' ========================================================================================
Function frmImages_DisplayImage(byval hwnd as HWND, byval nIndex as long) As LRESULT
   dim pImageCtx as CImageCtx PTR = AfxCImageCtxPtr(hwnd, IDC_FRMIMAGES_IMAGECTX)
   if pImageCtx = 0 then exit function
   
   dim as hwnd hImageCtx = GetDlgItem(hwnd, IDC_FRMIMAGES_IMAGECTX)
   
   dim pImageType as IMAGES_TYPE ptr 
   dim wszImageName as wstring * MAX_PATH
   dim as CWSTR wszFilename = L(0002, "File") & ": "
   
   ' Get the disk filename for the selected Image. Do not use the method of storing
   ' a pointer to the IMAGES_TYPE in the ListView data area because the pDoc->AllImages
   ' array gets REDIM PRESERVE which moves the array elements around in memory.
   ShowWindow(hImageCtx, SW_HIDE)
   if nIndex >= 0 then
      FF_ListView_GetItemText(GetDlgItem(hwnd, IDC_FRMIMAGES_LISTVIEW), nIndex, 0, @wszImageName, MAX_PATH)
      pImageType = GetImagesTypePtr(wszImageName)

      if pImageType then 
         pImageCtx->LoadImageFromFile(pImageType->wszFileName)
         if SendMessage( GetDlgItem(hwnd, IDC_FRMIMAGES_ACTUALSIZE), BM_GETCHECK, 0, 0) = BST_CHECKED then
            pImageCtx->SetImageAdjustment( GDIP_IMAGECTX_ACTUALSIZE, CTRUE )
         else   
            pImageCtx->SetImageAdjustment( GDIP_IMAGECTX_AUTOSIZE, CTRUE )
         end if
         wszFilename = wszFilename & pImageType->wszFileName
         if AfxFileExists(pImageType->wszFileName) = false then
            wszFilename = wszFilename & " (" & L(81, "File not found") & ")"
         end if   
         ShowWindow(hImageCtx, SW_SHOW)
      end if
   end if         
   AfxSetWindowText( GetDlgItem(hwnd, IDC_FRMIMAGES_LBLFILENAME), wszFilename)

   Function = 0
End Function


' ========================================================================================
' Generate an IMAGE name based on the filename.
' ========================================================================================
function frmImages_GenerateImageName( byref wszFilename as wstring ) as CWSTR
   dim as CWSTR wszImageName
   dim as Boolean fDuplicateName
   dim as long counter = 0
   
   ' Search all existing image names to ensure there are no duplicates.
   do 
      ' Construct an IMAGE name to test
      fDuplicateName = false
      wszImageName = "IMAGE_" & ucase(AfxStrPathname("NAME", wszFilename))
      wszImageName = AfxStrRemoveAny(wszImageName, " .,") 
      if counter > 0 then wszImageName = wszImageName & counter

      dim pDoc as clsDocument ptr = gApp.pDocList
      do until pDoc = 0
         if (gApp.IsProjectActive = false) andalso (pDoc <> gTTabCtl.GetActiveDocumentPtr) then 
            ' For non-projects only deal with images related to the active form.
         else
            for i as long = lbound(pDoc->AllImages) to ubound(pDoc->AllImages)
               if pDoc->AllImages(i).wszImageName = wszImageName then
                  fDuplicateName = true         
                  counter = counter + 1: exit do
               end if
            next   
         end if
         pDoc = pDoc->pDocNext
      loop

      if fDuplicateName = false then exit do
   loop
   
   function = wszImageName
end function



' ========================================================================================
' Add an image to the Image Manager
' ========================================================================================
Function frmImages_AddImage(ByVal HWnd As HWnd) As LRESULT
   ' Display the Open File Dialog
   Dim pItems As IShellItemArray Ptr = AfxIFileOpenDialogMultiple(HWnd, IDM_ADDIMAGE)
   If pItems = Null Then Exit Function
   Dim dwItemCount As Long, i As Long, pItem As IShellItem Ptr, pwszName As WString Ptr
   pItems->lpVtbl->GetCount(pItems, @dwItemCount)
   
   dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr
   dim as hwnd hLV = GetDlgItem(HWND, IDC_FRMIMAGES_LISTVIEW)
   dim as CWSTR wszText 
   dim as Boolean fDuplicate
   
   dim as long nSelLine = ListView_GetSelection(hLV)
   
   For i = 0 To dwItemCount - 1
      pItems->lpVtbl->GetItemAt(pItems, i, @pItem)
      If pItem Then
         pItem->lpVtbl->GetDisplayName(pItem, SIGDN_FILESYSPATH, @pwszName)
         If pwszName Then 
            ' Add the entry to the listview only if the generated name does
            ' not already exist in the AllImages array in order to avoid
            ' adding duplicate entries.
            fDuplicate = false
            for ii as long = lbound(pDoc->AllImages) to ubound(pDoc->AllImages)
               IF ucase(pDoc->AllImages(ii).wszFileName) = ucase(*pwszName) then
                  fDuplicate = true: exit for
               end if
            next
            if fDuplicate = false then 
               dim as long ub = ubound(pDoc->AllImages) + 1
               redim preserve pDoc->AllImages(ub)
               pDoc->AllImages(ub).wszFileName  = *pwszName
               pDoc->AllImages(ub).wszImageName = frmImages_GenerateImageName(*pwszName)
               pDoc->AllImages(ub).wszFormat    = wstr("RCDATA")
               pDoc->AllImages(ub).pDoc         = pDoc
               dim as long count = ListView_GetItemCount(hLV)
               FF_ListView_InsertItem( hLV, count, 0, pDoc->AllImages(ub).wszImageName )
               FF_ListView_InsertItem( hLV, count, 1, pDoc->AllImages(ub).wszFormat)
               nSelLine = count
            end if
            
            CoTaskMemFree(pwszName)
            pwszName = Null
         End If
         pItem->lpVtbl->Release(pItem)
         pItem = Null
      End If
   Next
   pItems->lpVtbl->Release(pItems)

   ' Save the changes to the pDoc
   if pDoc then pDoc->SaveFile

   ListView_SelectItem(hLV, nSelLine)
   frmImages_DisplayImage(hwnd, nSelLine)
   SetFocus hLV

   Function = 0
End Function


' ========================================================================================
' Remove an image from the Image Manager
' ========================================================================================
Function frmImages_RemoveImage(ByVal HWnd As HWnd) As LRESULT

   dim pDoc  as clsDocument ptr 
   dim pCtrl as clsControl ptr 
   dim pProp as clsProperty ptr 
   dim pImageType as IMAGES_TYPE ptr 
   
   dim as hwnd hLV = GetDlgItem(HWND, IDC_FRMIMAGES_LISTVIEW)
   dim as long nCurSel = ListView_GetSelection(hLV)
   if nCurSel = -1 then exit function 
   
   dim wszImageName as wstring * MAX_PATH
   dim as CWSTR wszMsg
   
   ' Get the pDoc related to this IMAGE_NAME being deleted in order to ensure that
   ' we remove it from the current document.
   FF_ListView_GetItemText(hLV, nCurSel, 0, @wszImageName, MAX_PATH)
   pImageType = GetImagesTypePtr(wszImageName)
   

   ' Count the number of times this IMAGE_NAME is used by controls/properties
   ' in the file/project. Ask the user if he wants to remove the image and
   ' thereby remove all references in controls that rely on the image.
   dim pProps(any) as clsProperty ptr
   pDoc = gApp.pDocList
   do until pDoc = 0
      ' Loop through all controls on the Form
      for i as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
         pCtrl = pDoc->Controls.ItemAt(i)
         ' Loop through all properties for each control
         for ii as long = lbound(pCtrl->Properties) to ubound(pCtrl->Properties)
            pProp = @pCtrl->Properties(ii)
            if pProp = 0 then continue for
            if pProp->PropType = PropertyType.ImagePicker then
               if pProp->wszPropValue = wszImageName then
                  dim as long ub = ubound(pProps) + 1
                  redim preserve pProps(ub)
                  pProps(ub) = pProp
               end if
            end if   
         next
      next   
      pDoc = pDoc->pDocNext
   loop
   
   dim as long numRefs = ubound(pProps) - lbound(pProps) + 1
   wszMsg = L(366,"Are you sure you want to delete?") & vbcrlf & _
            wszImageName & vbcrlf & _
            "(" & numRefs & " " & L(379,"references") & ")"
   If MessageBox( hwnd, wszMsg, L(276,"Confirm"), MB_ICONQUESTION Or MB_YESNOCANCEL ) <> IDYES Then exit function

   ' Do the actual delete from the AllImages array for the pDoc that holds the image.
   if pImageType then pDoc = pImageType->pDoc
   if pDoc then
      for i as long = nCurSel to ubound(pDoc->AllImages) - 1
         pDoc->AllImages(i) = pDoc->AllImages(i+1) 
      next
      if Ubound(pDoc->AllImages)-1 < 0 then
         erase pDoc->AllImages
      else   
         Redim Preserve pDoc->AllImages(Ubound(pDoc->AllImages)-1)
      end if
      ' Save the changes to the pDoc
      pDoc->SaveFile
   end if
   
   ' Ensure that the IMAGE_NAME is removed from all Properties that rely on the image.
   for i as long = lbound(pProps) to ubound(pProps)
      pProps(i)->wszPropValue = ""
   next
   DisplayPropertyList(gTTabCtl.GetActiveDocumentPtr)
   
   ' Remove the Listview line and reposition the selection.
   ListView_DeleteItem(hLV, nCurSel)
   
   ' Reposition Listview to next closest line.
   if ListView_GetItemCount(hLV) > 0 then
      nCurSel = nCurSel - 1
      if nCurSel < 0 then nCurSel = 0
      ListView_SelectItem(hLV, nCurSel)
   else
      nCurSel = -1   
   end if
   frmImages_DisplayImage(hwnd, nCurSel)
   SetFocus hLV
   
   Function = 0
End Function


' ========================================================================================
' Process WM_SIZE message for window/dialog: frmImages
' ========================================================================================
Function frmImages_OnSize(ByVal HWnd As HWnd, ByVal state As UINT, ByVal cx As Long, ByVal cy As Long) As LRESULT
   If state <> SIZE_MINIMIZED Then
      ' Size the toolbar
      Toolbar_AutoSize GetDlgItem(hwnd, IDC_FRMIMAGES_TOOLBAR)
   End If
   Function = 0
End Function


' ========================================================================================
' Process WM_CREATE message for window/dialog: frmImages
' ========================================================================================
Function frmImages_OnCreate(ByVal HWnd As HWnd, ByVal lpCreateStructPtr As LPCREATESTRUCT) As BOOLEAN

   ' This is a modal popup window so disable the parent window
   DisableAllModeless()

   '  Message cracker macro expects a True to be returned for a successful
   '  OnCreate handler even though returning -1 from a standard WM_CREATE
   '  call would stop creating the window. This is just one of those Windows
   '  inconsistencies.
   Return True
End Function


' ========================================================================================
' Process WM_COMMAND message for window/dialog: frmImages
' ========================================================================================
Function frmImages_OnCommand(ByVal HWnd As HWnd, ByVal id As Long, ByVal hwndCtl As HWnd, ByVal codeNotify As UINT) As LRESULT

   dim as hwnd hLV = GetDlgItem(HWND, IDC_FRMIMAGES_LISTVIEW)
   dim as long nCurSel 
   dim wszImageName as wstring * MAX_PATH
   dim wszText as wstring * MAX_PATH

   dim pImageCtx  as CImageCtx PTR
   dim pImageType as IMAGES_TYPE ptr
   dim pProp      as clsProperty ptr

   if IsWindow(hLV) then
      nCurSel = ListView_GetSelection(hLV)
      FF_ListView_GetItemText(hLV, nCurSel, 0, @wszImageName, MAX_PATH)
      pImageType = GetImagesTypePtr(wszImageName)
   end if
   
   
   Select Case id
      Case IDCANCEL
         If codeNotify = BN_CLICKED Then
            SendMessage HWnd, WM_CLOSE, 0, 0
            Exit Function
         End If
         
      case IDM_ADDIMAGE
         frmImages_AddImage(HWND)
         exit function
     
      case IDM_REMOVEIMAGE   
         frmImages_RemoveImage(HWND)
         exit function
      
      case IDM_FORMATIMAGE
         if nCurSel = -1 then exit function
         Dim pt As Point   
         Dim as HMENU hPopUpMenu = CreateImagesResourceMenu()
         GetCursorPos @pt
         TrackPopupMenu(hPopUpMenu, 0, pt.x, pt.y, 0, HWnd, ByVal Null)
         DestroyMenu hPopUpMenu
         exit function
         
      case IDC_FRMIMAGES_FORMATBITMAP, IDC_FRMIMAGES_FORMATICON, _
           IDC_FRMIMAGES_FORMATCURSOR, IDC_FRMIMAGES_FORMATRCDATA
         if id = IDC_FRMIMAGES_FORMATBITMAP then wszText = "BITMAP"
         if id = IDC_FRMIMAGES_FORMATICON   then wszText = "ICON"
         if id = IDC_FRMIMAGES_FORMATCURSOR then wszText = "CURSOR"
         if id = IDC_FRMIMAGES_FORMATRCDATA then wszText = "RCDATA"
         if pImageType then 
            pImageType->wszFormat = wszText
            FF_ListView_SetItemText( hLV, nCurSel, 1, @wszText, MAX_PATH )
            pImageType->pDoc->SaveFile
         end if
         exit function
         
      case IDM_ATTACHIMAGE, IDM_DETACHIMAGE
         if nCurSel = -1 then exit function
         if len(wszImageName) = 0 then exit function
         if id = IDM_DETACHIMAGE then wszImageName = ""
         pProp = GetActivePropertyPtr()
         if pProp then pProp->wszPropValue = wszImageName
         DisplayPropertyList(gTTabCtl.GetActiveDocumentPtr)
         dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr
         dim pCtrl as clsControl ptr = pDoc->Controls.GetActiveControl
         ApplyControlProperties( pDoc, pCtrl )
         pDoc->UserModified = true
         SendMessage( HWnd, WM_CLOSE, 0, 0)
         exit function
         
      CASE IDC_FRMIMAGES_AUTOSIZE
         If codeNotify = BN_CLICKED Then
            pImageCtx = AfxCImageCtxPtr(hwnd, IDC_FRMIMAGES_IMAGECTX)
            if pImageCtx then pImageCtx->SetImageAdjustment( GDIP_IMAGECTX_AUTOSIZE, CTRUE )
            EXIT FUNCTION
         END IF

      CASE IDC_FRMIMAGES_ACTUALSIZE
         If codeNotify = BN_CLICKED Then
            pImageCtx = AfxCImageCtxPtr(hwnd, IDC_FRMIMAGES_IMAGECTX)
            if pImageCtx then pImageCtx->SetImageAdjustment( GDIP_IMAGECTX_ACTUALSIZE, CTRUE )
            EXIT FUNCTION
         END IF
   End Select

   Function = 0
End Function


' ========================================================================================
' Process WM_NOTIFY message for window/dialog: frmImages
' ========================================================================================
Function frmImages_OnNotify(ByVal HWnd As HWnd, ByVal id As Long, ByVal pNMHDR As NMHDR Ptr) As LRESULT

   Dim pWindow As CWindow Ptr = AfxCWindowPtr(HWnd)
    
   Select Case id
      Case IDC_FRMIMAGES_LISTVIEW
         If pNMHDR->code = LVN_ITEMCHANGED Then
            dim as long nSelLine = ListView_GetSelection(pNMHDR->hwndFrom)
            frmImages_DisplayImage(hwnd, nSelLine)
         End If
   End Select

   Function = 0
End Function


' ========================================================================================
' Process WM_CLOSE message for window/dialog: frmImages
' ========================================================================================
Function frmImages_OnClose(HWnd As HWnd) As LRESULT
   ' Enables parent window keeping parent's zorder
   EnableAllModeless()
   DestroyWindow HWnd
   Function = 0
End Function


' ========================================================================================
' Process WM_DESTROY message for window/dialog: frmImages
' ========================================================================================
Function frmImages_OnDestroy(HWnd As HWnd) As LRESULT
   ' Destroy the toolbar image lists
   ImageList_Destroy CAST(HIMAGELIST, SendMessage(GetDlgItem(hwnd, IDC_FRMIMAGES_TOOLBAR), TB_SETIMAGELIST, 0, 0))
   ImageList_Destroy CAST(HIMAGELIST, SendMessage(GetDlgItem(hwnd, IDC_FRMIMAGES_TOOLBAR), TB_SETHOTIMAGELIST, 0, 0))
   ImageList_Destroy CAST(HIMAGELIST, SendMessage(GetDlgItem(hwnd, IDC_FRMIMAGES_TOOLBAR), TB_SETDISABLEDIMAGELIST, 0, 0))
   HWND_FRMIMAGES = 0
   PostQuitMessage(0)
   Function = 0
End Function

' ========================================================================================
' Processes messages for the subclassed frmImages listview control.
' ========================================================================================
Function frmImages_Listview_SubclassProc ( _
                  ByVal HWnd   As HWnd, _                 ' // Control window handle
                  ByVal uMsg   As UINT, _                 ' // Type of message
                  ByVal wParam As WPARAM, _               ' // First message parameter
                  ByVal lParam As LPARAM, _               ' // Second message parameter
                  ByVal uIdSubclass As UINT_PTR, _        ' // The subclass ID
                  ByVal dwRefData As DWORD_PTR _          ' // Pointer to reference data
                  ) As LRESULT

   Dim As HWND hLV 
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
         'SetDocumentErrorPosition( HWND, gCompile.CompileID )    
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
         RemoveWindowSubclass HWnd, @frmImages_Listview_SubclassProc, uIdSubclass

   End Select

   ' Default processing of Windows messages
   Function = DefSubclassProc(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' frmImages Window procedure
' ========================================================================================
Function frmImages_WndProc( ByVal HWnd   As HWnd, _
                            ByVal uMsg   As UINT, _
                            ByVal wParam As WPARAM, _
                            ByVal lParam As LPARAM _
                            ) As LRESULT

   Select Case uMsg
      HANDLE_MSG (HWnd, WM_CREATE,   frmImages_OnCreate)
      HANDLE_MSG (HWnd, WM_CLOSE,    frmImages_OnClose)
      HANDLE_MSG (HWnd, WM_SIZE,     frmImages_OnSize)
      HANDLE_MSG (HWnd, WM_DESTROY,  frmImages_OnDestroy)
      HANDLE_MSG (HWnd, WM_COMMAND,  frmImages_OnCommand)
      HANDLE_MSG (HWnd, WM_NOTIFY,   frmImages_OnNotify)
   End Select

   ' for messages that we don't deal with
   Function = DefWindowProc(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' frmImages_Show
' ========================================================================================
Function frmImages_Show( ByVal hWndParent As HWnd, _
                         Byval pProp as clsProperty ptr = 0 _
                         ) As Long

   dim wszImageName as wstring * MAX_PATH
   if pProp then wszImageName = pProp->wszPropValue
   
   '  Create the main window and child controls
   Dim pWindow As CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowPtr(hwndParent)->DPI

   HWND_FRMIMAGES = _
   pWindow->Create( hWndParent, L(368,"Image Manager"), @frmImages_WndProc, 0, 0, 0, 0, _
        WS_POPUP Or WS_CAPTION Or WS_SYSMENU Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, _
        WS_EX_DLGMODALFRAME Or WS_EX_CONTROLPARENT Or WS_EX_LEFT )
   pWindow->SetClientSize(700, 480)
   pWindow->Center(pWindow->hWindow, hWndParent)

   ' Add a tooolbar
   DIM hToolBar AS HWND = pWindow->AddControl("Toolbar", , IDC_FRMIMAGES_TOOLBAR, "", 0, 0, 0, 0, _
                           WS_CHILD or WS_VISIBLE OR WS_CLIPCHILDREN OR WS_CLIPSIBLINGS or _
                           TBSTYLE_FLAT or TBSTYLE_LIST or TBSTYLE_TOOLTIPS or TBSTYLE_AUTOSIZE or _
                           WS_BORDER or CCS_TOP )

   DIM AS LONG nButtonWidth, nButtonHeight
   nButtonWidth = LOWORD(SendMessage(hToolBar, TB_GETBUTTONSIZE, 0, 0)) * pWindow->rxRatio
   nButtonHeight = HIWORD(SendMessage(hToolBar, TB_GETBUTTONSIZE, 0, 0)) * pWindow->ryRatio
   SendMessage hToolBar, TB_SETBUTTONSIZE, 0, MAKELONG(nButtonWidth, nButtonHeight)
   ' Send this message for backward compatibility
   SendMessage hToolBar, TB_BUTTONSTRUCTSIZE, SIZEOF(TBBUTTON), 0

   ' Calculate the size of the icon according the DPI
   DIM cx AS LONG = 20 * pWindow->DPI \ 96
   
   ' Create an image list for the toolbar
   DIM hImageList AS HIMAGELIST
   hImageList = ImageList_Create(cx, cx, ILC_COLOR32 OR ILC_MASK, 5, 0)
   IF hImageList THEN
      ImageList_SetBkColor(hImageList, CLR_NONE)
      AfxGdipAddIconFromRes(hImageList, pWindow->InstanceHandle, "IMAGE_ADDIMAGE")
      AfxGdipAddIconFromRes(hImageList, pWindow->InstanceHandle, "IMAGE_REMOVEIMAGE")
      AfxGdipAddIconFromRes(hImageList, pWindow->InstanceHandle, "IMAGE_FORMATIMAGE")
      AfxGdipAddIconFromRes(hImageList, pWindow->InstanceHandle, "IMAGE_ATTACHIMAGE")
      AfxGdipAddIconFromRes(hImageList, pWindow->InstanceHandle, "IMAGE_DETACHIMAGE")
      ' Set the normal image list
      Toolbar_SetImageList hToolBar, hImageList
      ' Set the hot image list with the same images than the normal one
      Toolbar_SetHotImageList hToolBar, hImageList
   END IF

   ' // Add buttons to the toolbar
   dim as long fStyle = BTNS_BUTTON or BTNS_AUTOSIZE or BTNS_SHOWTEXT
   Toolbar_AddButton hToolBar, 0, IDM_ADDIMAGE,    0, fStyle, 0, L(369, "Add Image")
   Toolbar_AddSeparator hToolBar
   Toolbar_AddButton hToolBar, 1, IDM_REMOVEIMAGE, 0, fStyle, 0, L(370, "Remove Image")
   Toolbar_AddSeparator hToolBar
   Toolbar_AddButton hToolBar, 2, IDM_FORMATIMAGE, 0, fStyle, 0, L(371, "Resource Format")
   Toolbar_AddSeparator hToolBar
   if pProp then
      Toolbar_AddButton hToolBar, 3, IDM_ATTACHIMAGE, 0, fStyle, 0, L(372, "Attach to Control")
      Toolbar_AddSeparator hToolBar
      Toolbar_AddButton hToolBar, 4, IDM_DETACHIMAGE, 0, fStyle, 0, L(373, "Detach from Control")
   end if
   
   ' Size the toolbar
   Toolbar_AutoSize hToolBar

   dim as long nTop = pWindow->UnScaleY(AfxGetWindowHeight(hToolBar)) + 10
   
   dim as hwnd hLV = _
      pWindow->AddControl("LISTVIEW", , IDC_FRMIMAGES_LISTVIEW, "", 10, nTop, 330, 430, _
      WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or LVS_REPORT Or LVS_SHOWSELALWAYS or LVS_SINGLESEL, _
      WS_EX_LEFT Or WS_EX_CLIENTEDGE or WS_EX_RIGHTSCROLLBAR, , _
      Cast(SUBCLASSPROC, @frmOutput_Listview_SubclassProc), IDC_FRMOUTPUT_LISTVIEW, Cast(DWORD_PTR, @pWindow))
                                     
   ' Make the listview header flat
   ListView_MakeHeaderFlat(hLV)

          
   ' Add some extended styles
   Dim dwExStyle As DWORD
   dwExStyle = ListView_GetExtendedListViewStyle(hLV)
   dwExStyle = dwExStyle Or LVS_EX_FULLROWSELECT Or LVS_EX_GRIDLINES Or LVS_EX_DOUBLEBUFFER Or LVS_EX_FLATSB
   ListView_SetExtendedListViewStyle(hLV, dwExStyle)

   ' Configure the ListView
   ListView_AddColumn( hLV, 0, L(364, "Name"), pWindow->ScaleX(210) )
   ListView_AddColumn( hLV, 1, L(374, "File Type"), pWindow->ScaleX(100) )

   ' Add the images to the listview
   dim as long nextLine = 0

   dim pDoc as clsDocument ptr = gApp.pDocList
   dim as long nCount = 0
   dim as long nSelLine = 0
   
   do until pDoc = 0
      if (gApp.IsProjectActive = false) andalso (pDoc <> gTTabCtl.GetActiveDocumentPtr) then 
         ' For non-projects only output images related to the active form.
      else
         for i as long = lbound(pDoc->AllImages) to ubound(pDoc->AllImages)
            FF_ListView_InsertItem( hLV, nextLine, 0, pDoc->AllImages(i).wszImageName )
            FF_ListView_InsertItem( hLV, nextLine, 1, pDoc->AllImages(i).wszFormat )
            ' If the Image being added matches the incoming ImageName then save the
            ' insert line position so that the selection can be set to that line.
            if pDoc->AllImages(i).wszImageName = wszImageName then nSelLine = nextLine
            nextLine = nextLine + 1
         next   
      end if
      pDoc = pDoc->pDocNext
   loop

   
   ' Add an image control                                                         
   DIM pImageCtx AS CImageCtx = CImageCtx(pWindow, IDC_FRMIMAGES_IMAGECTX, , 370, nTop + 30, 306, 306)

   ' Add a button without coordinates (it will be reiszed in WM_SIZE, below)
   DIM hCtl AS HWND = _
      pWindow->AddControl("RadioButton", , IDC_FRMIMAGES_ACTUALSIZE, L(376, "Actual size"), 370, nTop + 354, 120, 20)
      pWindow->AddControl("RadioButton", , IDC_FRMIMAGES_AUTOSIZE, L(375, "Autosize"), 492, nTop + 354, 100, 20)
      SendMessage hCtl, BM_SETCHECK, BST_CHECKED, 0

   Dim As HWnd hLabel = _
      pWindow->AddControl("LABEL", , IDC_FRMIMAGES_LBLFILENAME, "", 370, nTop + 380, 320, 52, _
           WS_CHILD Or WS_VISIBLE Or SS_LEFT, WS_EX_LEFT Or WS_EX_LTRREADING)


   pWindow->AddControl("GROUPBOX",, IDC_FRMIMAGES_FRAMEPREVIEW, L(377,"Image Preview"), 356, nTop, 334, 350, _
        WS_CHILD Or WS_VISIBLE Or BS_TEXT Or BS_LEFT Or BS_NOTIFY Or BS_GROUPBOX, _
        WS_EX_TRANSPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING)

   ' Set the Listview to the selected line and preview the image display
   ListView_SelectItem(hLV, nSelLine)
   frmImages_DisplayImage(pWindow->hWindow, nSelLine)
   
   SetFocus hLV
      
   ' Process Windows messages
   Function = pWindow->DoEvents(SW_SHOW)
   
   ' Delete the frmImages CWindow class manually allocated memory 
   Delete pWindow

End Function





