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

#include once "frmImageManager.bi"
#include once "clsDocument.bi"
#include once "clsTopTabCtl.bi"


' ========================================================================================
' Create the popup menu showing the list of valid image formats
' ========================================================================================
Function frmImageManager_CreateImagesResourceMenu() As HMENU
   
   dim wszImageFormat as wstring * MAX_PATH
   dim as hwnd hLV = GetDlgItem(HWND_FRMIMAGES, IDC_FRMIMAGEMANAGER_LISTVIEW)
   dim as long nCurSel = ListView_GetSelection(hLV)
   FF_ListView_GetItemText(hLV, nCurSel, 1, @wszImageFormat, MAX_PATH)

   dim as long fNormal  = MF_ENABLED or MF_STRING 
   dim as long fChecked = MF_CHECKED or MF_ENABLED or MF_STRING
   dim as long i
   
   select case wszImageFormat
      case "RCDATA": i = 1
      case "BITMAP": i = 2
      case "ICON":   i = 3
      case "CURSOR": i = 4
   end select
   
   Dim hPopUpMenu As HMENU = CreatePopupMenu()
      AppendMenu( hPopUpMenu, iif(i = 1, fChecked, fNormal), IDC_FRMIMAGEMANAGER_FORMATRCDATA,  "RCDATA")
      AppendMenu( hPopUpMenu, iif(i = 2, fChecked, fNormal), IDC_FRMIMAGEMANAGER_FORMATBITMAP,  "BITMAP")
      AppendMenu( hPopUpMenu, iif(i = 3, fChecked, fNormal), IDC_FRMIMAGEMANAGER_FORMATICON,    "ICON")
      AppendMenu( hPopUpMenu, iif(i = 4, fChecked, fNormal), IDC_FRMIMAGEMANAGER_FORMATCURSOR,  "CURSOR")

   Function = hPopupMenu
 
End Function
   

' ========================================================================================
' Display a selected image in the Image Manager
' ========================================================================================
private Function frmImageManager_DisplayImage( byval hwnd as HWND, _
                                               byval nIndex as long _
                                               ) As LRESULT
   
   dim pImageCtx as CImageCtx PTR = AfxCImageCtxPtr(hwnd, IDC_FRMIMAGEMANAGER_IMAGECTX)
   if pImageCtx = 0 then exit function
   
   dim as hwnd hImageCtx = GetDlgItem(hwnd, IDC_FRMIMAGEMANAGER_IMAGECTX)
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(HWnd)
    
   dim pImageType as IMAGES_TYPE ptr 
   dim wszImageName as wstring * MAX_PATH
   dim as CWSTR wszFilename 
   
   ' Get the disk filename for the selected Image. Do not use the method of storing
   ' a pointer to the IMAGES_TYPE in the ListView data area because the pDoc->AllImages
   ' array gets REDIM PRESERVE which moves the array elements around in memory.
   ShowWindow(hImageCtx, SW_HIDE)
   EnableWindow( GetDlgItem(hwnd, IDC_FRMIMAGEMANAGER_CMDFILENAME), false)
   if nIndex >= 0 then
      FF_ListView_GetItemText(GetDlgItem(hwnd, IDC_FRMIMAGEMANAGER_LISTVIEW), nIndex, 0, @wszImageName, MAX_PATH)
      pImageType = GetImagesTypePtr(wszImageName)

      if pImageType then 
         pImageCtx->LoadImageFromFile(pImageType->wszFileName)
         ' Resize the image based on the ComboBox selections
         dim as long nCurSel = ComboBox_GetCurSel( GetDlgItem(hwnd, IDC_FRMIMAGEMANAGER_COMBO1) )
         select case nCurSel
            case 0     ' Actual size
               pImageCtx->SetImageAdjustment( GDIP_IMAGECTX_ACTUALSIZE, CTRUE )
            case 1     ' Apply High DPI scaling
               ' Make the image high dpi aware
               if pWindow then
                  pImageCtx->SetImageWidth( pWindow->ScaleX(pImageCtx->GetImageWidth) )
                  pImageCtx->SetImageHeight( pWindow->ScaleY(pImageCtx->GetImageHeight) )
               end if
               pImageCtx->SetImageAdjustment( GDIP_IMAGECTX_ACTUALSIZE, CTRUE )
            case 2     ' AutoSize
               pImageCtx->SetImageAdjustment( GDIP_IMAGECTX_AUTOSIZE, CTRUE )
            case 3     ' Fit to Width
               pImageCtx->SetImageAdjustment( GDIP_IMAGECTX_FITTOWIDTH, CTRUE )
            case 4     ' Fit to Height
               pImageCtx->SetImageAdjustment( GDIP_IMAGECTX_FITTOHEIGHT, CTRUE )
            case 5     ' Stretch
               pImageCtx->SetImageAdjustment( GDIP_IMAGECTX_STRETCH, CTRUE )
         end select
         
         dim hCtrl as HWND = GetDlgItem( HWND_FRMIMAGES, IDC_FRMIMAGEMANAGER_FRAMEPREVIEW )
         wszFilename = wszFilename & pImageType->wszFileName
         if AfxFileExists(pImageType->wszFileName) = false then
            wszFilename = wszFilename & " (" & L(81, "File not found") & ")"
         end if   
         ShowWindow(hImageCtx, SW_SHOW)
         EnableWindow( GetDlgItem(hwnd, IDC_FRMIMAGEMANAGER_CMDFILENAME), true)
      end if
   end if         
   AfxSetWindowText( GetDlgItem(hwnd, IDC_FRMIMAGEMANAGER_LBLFILENAME), wszFilename)

   Function = 0
End Function


' ========================================================================================
' Generate an IMAGE name based on the filename.
' ========================================================================================
private function frmImageManager_GenerateImageName( byval wszFilename as CWSTR ) as CWSTR
   
   dim as CWSTR wszImageName
   dim as Boolean fDuplicateName
   dim as long counter = 0
   
   ' Search all existing image names to ensure there are no duplicates.
   do 
      ' Construct an IMAGE name to test
      fDuplicateName = false
      wszImageName = "IMAGE_" & ucase(AfxStrPathname("NAME", wszFilename))
      wszImageName = AfxStrRemoveAny( wszImageName, " .," ) 
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
private Function frmImageManager_AddImage( ByVal HWnd As HWnd ) As LRESULT
   
   ' Display the Open File Dialog
   Dim pItems As IShellItemArray Ptr = AfxIFileOpenDialogMultiple(HWnd, IDM_ADDIMAGE)
   If pItems = Null Then Exit Function
   
   Dim as long dwItemCount 
   dim as long i 
   dim pItem As IShellItem Ptr
   dim pwszName As WString Ptr
   
   pItems->lpVtbl->GetCount(pItems, @dwItemCount)
   
   dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr
   dim as hwnd hLV = GetDlgItem(HWND, IDC_FRMIMAGEMANAGER_LISTVIEW)
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
               pDoc->AllImages(ub).wszImageName = frmImageManager_GenerateImageName(*pwszName)
               pDoc->AllImages(ub).wszFormat    = "RCDATA"
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
   frmImageManager_DisplayImage(hwnd, nSelLine)
   SetFocus hLV

   Function = 0
End Function


' ========================================================================================
' Remove an image from the Image Manager
' ========================================================================================
private Function frmImageManager_RemoveImage( ByVal HWnd As HWnd ) As LRESULT

   dim pDoc  as clsDocument ptr 
   dim pCtrl as clsControl ptr 
   dim pProp as clsProperty ptr 
   dim pImageType as IMAGES_TYPE ptr 
   
   dim as hwnd hLV = GetDlgItem(HWND, IDC_FRMIMAGEMANAGER_LISTVIEW)
   dim as long nCurSel = ListView_GetSelection(hLV)
   if nCurSel = -1 then exit function 
   
   dim wszImageName as wstring * MAX_PATH
   dim as CWSTR wszMsg
   
   ' Get the pDoc related to this IMAGE_NAME being deleted in order to ensure that
   ' we remove it from the correct document.
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
   if pImageType->pDoc then
      for i as long = nCurSel to ubound(pImageType->pDoc->AllImages) - 1
         pImageType->pDoc->AllImages(i) = pImageType->pDoc->AllImages(i+1) 
      next
      if Ubound(pImageType->pDoc->AllImages)-1 < 0 then
         erase pImageType->pDoc->AllImages
      else               
         Redim Preserve pImageType->pDoc->AllImages(Ubound(pImageType->pDoc->AllImages)-1)
      end if
      ' Save the changes to the pDoc
      pImageType->pDoc->SaveFile
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
   frmImageManager_DisplayImage(hwnd, nCurSel)
   SetFocus hLV
   
   Function = 0
End Function


' ========================================================================================
' Process WM_CREATE message for window/dialog: frmImages
' ========================================================================================
private Function frmImageManager_OnCreate( ByVal HWnd As HWnd, _
                                           ByVal lpCreateStructPtr As LPCREATESTRUCT _
                                           ) As BOOLEAN

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
private Function frmImageManager_OnCommand( ByVal HWnd As HWnd, _
                                            ByVal id As Long, _
                                            ByVal hwndCtl As HWnd, _
                                            ByVal codeNotify As UINT _
                                            ) As LRESULT

   dim as hwnd hLV = GetDlgItem(HWND, IDC_FRMIMAGEMANAGER_LISTVIEW)
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
         
      case IDC_FRMIMAGEMANAGER_CMDFILENAME
         If codeNotify = BN_CLICKED Then
            ' Popup an inputbox to get corrected image filename
            dim as CWSTR wszResult
            dim as CWSTR wszFilename 
            if pImageType then wszFilename = pImageType->wszFileName
            wszResult = AfxInputBox( HWND_FRMIMAGES, 0, 0, _
                                     L(368,"Image Manager"), _
                                     L(405,"Enter valid image file path") & ":", _
                                     wszFilename, 1000 )
            if len(wszResult) then
               if ucase(wszResult) <> ucase(wszFilename) then
                  if pImageType then
                     pImageType->wszFileName = wszResult
                     if pImageType->pDoc then pImageType->pDoc->SaveFile
                     frmImageManager_DisplayImage( HWND, nCurSel )
                  end if                     
               end if
            end if   
            Exit Function
         End If
         
      case IDM_ADDIMAGE
         frmImageManager_AddImage(HWND)
         exit function
     
      case IDM_REMOVEIMAGE   
         frmImageManager_RemoveImage(HWND)
         exit function
      
      case IDM_FORMATIMAGE
         if nCurSel = -1 then exit function
         Dim pt As Point   
         Dim as HMENU hPopUpMenu = frmImageManager_CreateImagesResourceMenu()
         GetCursorPos @pt
         TrackPopupMenu(hPopUpMenu, 0, pt.x, pt.y, 0, HWnd, ByVal Null)
         DestroyMenu hPopUpMenu
         exit function
         
      case IDC_FRMIMAGEMANAGER_FORMATBITMAP, IDC_FRMIMAGEMANAGER_FORMATICON, _
           IDC_FRMIMAGEMANAGER_FORMATCURSOR, IDC_FRMIMAGEMANAGER_FORMATRCDATA
         if id = IDC_FRMIMAGEMANAGER_FORMATBITMAP then wszText = "BITMAP"
         if id = IDC_FRMIMAGEMANAGER_FORMATICON   then wszText = "ICON"
         if id = IDC_FRMIMAGEMANAGER_FORMATCURSOR then wszText = "CURSOR"
         if id = IDC_FRMIMAGEMANAGER_FORMATRCDATA then wszText = "RCDATA"
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
         
         ' If an editor called this ImageManager then save the selected
         ' image name into the pProp of the currently selected 
         if (GetParent( HWND_FRMIMAGES ) = HWND_FRMSTATUSBAREDITOR) or _
            (GetParent( HWND_FRMIMAGES ) = HWND_FRMTOOLBAREDITOR) or _
            (GetParent( HWND_FRMIMAGES ) = HWND_FRMVDTABCHILD) then
            Dim pWindow As CWindow Ptr = AfxCWindowPtr(HWND_FRMIMAGES)
            pProp = cast( clsProperty ptr, pWindow->UserData(0) )
            if pProp then pProp->wszPropValue = wszImageName
         else   
            pProp = GetActivePropertyPtr()
            if pProp then pProp->wszPropValue = wszImageName
            DisplayPropertyList(gTTabCtl.GetActiveDocumentPtr)
            dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr
            dim pCtrl as clsControl ptr = pDoc->Controls.GetActiveControl
            ApplyControlProperties( pDoc, pCtrl )
            pDoc->UserModified = true
         end if
         
         SendMessage( HWnd, WM_CLOSE, 0, 0)
         exit function
         
      CASE IDC_FRMIMAGEMANAGER_COMBO1
         If codeNotify = CBN_SELCHANGE Then
            frmImageManager_DisplayImage( hwnd, nCurSel )
            EXIT FUNCTION
         END IF

   End Select

   Function = 0
End Function


' ========================================================================================
' Process WM_NOTIFY message for window/dialog: frmImages
' ========================================================================================
private Function frmImageManager_OnNotify( ByVal HWnd As HWnd, _
                                           ByVal id As Long, _
                                           ByVal pNMHDR As NMHDR Ptr _
                                           ) As LRESULT
   Select Case id
      Case IDC_FRMIMAGEMANAGER_LISTVIEW
         If pNMHDR->code = LVN_ITEMCHANGED Then
            dim as long nSelLine = ListView_GetSelection(pNMHDR->hwndFrom)
            frmImageManager_DisplayImage(hwnd, nSelLine)
         End If
   End Select

   Function = 0
End Function


' ========================================================================================
' Process WM_CLOSE message for window/dialog: frmImages
' ========================================================================================
private Function frmImageManager_OnClose( byval HWnd As HWnd ) As LRESULT
   ' Enables parent window keeping parent's zorder
   EnableAllModeless()
   DestroyWindow HWnd
   Function = 0
End Function


' ========================================================================================
' Process WM_DESTROY message for window/dialog: frmImages
' ========================================================================================
private Function frmImageManager_OnDestroy( byval HWnd As HWnd ) As LRESULT
   ' Destroy the toolbar image lists
   ImageList_Destroy CAST(HIMAGELIST, SendMessage(GetDlgItem(hwnd, IDC_FRMIMAGEMANAGER_TOOLBAR), TB_SETIMAGELIST, 0, 0))
   ImageList_Destroy CAST(HIMAGELIST, SendMessage(GetDlgItem(hwnd, IDC_FRMIMAGEMANAGER_TOOLBAR), TB_SETHOTIMAGELIST, 0, 0))
   ImageList_Destroy CAST(HIMAGELIST, SendMessage(GetDlgItem(hwnd, IDC_FRMIMAGEMANAGER_TOOLBAR), TB_SETDISABLEDIMAGELIST, 0, 0))
   HWND_FRMIMAGES = 0
   PostQuitMessage(0)
   Function = 0
End Function


' ========================================================================================
' Processes messages for the subclassed frmImages listview control.
' ========================================================================================
private Function frmImageManager_Listview_SubclassProc ( _
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
         RemoveWindowSubclass HWnd, @frmImageManager_Listview_SubclassProc, uIdSubclass

   End Select

   ' Default processing of Windows messages
   Function = DefSubclassProc(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' frmImages Window procedure
' ========================================================================================
private Function frmImageManager_WndProc( ByVal HWnd   As HWnd, _
                                          ByVal uMsg   As UINT, _
                                          ByVal wParam As WPARAM, _
                                          ByVal lParam As LPARAM _
                                          ) As LRESULT

   Select Case uMsg
      HANDLE_MSG (HWnd, WM_CREATE,   frmImageManager_OnCreate)
      HANDLE_MSG (HWnd, WM_CLOSE,    frmImageManager_OnClose)
      HANDLE_MSG (HWnd, WM_DESTROY,  frmImageManager_OnDestroy)
      HANDLE_MSG (HWnd, WM_COMMAND,  frmImageManager_OnCommand)
      HANDLE_MSG (HWnd, WM_NOTIFY,   frmImageManager_OnNotify)
   End Select

   Function = DefWindowProc(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' frmImageManager_Show
' ========================================================================================
public Function frmImageManager_Show( ByVal hWndParent As HWnd, _
                                      Byval pProp as clsProperty ptr = 0 _
                                      ) As LRESULT

   dim wszImageName as wstring * MAX_PATH
   if pProp then wszImageName = pProp->wszPropValue
   
   '  Create the main window and child controls
   Dim pWindow As CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowPtr(hwndParent)->DPI

   HWND_FRMIMAGES = _
   pWindow->Create( hWndParent, L(368,"Image Manager"), @frmImageManager_WndProc, 0, 0, 0, 0, _
        WS_POPUP Or WS_CAPTION Or WS_SYSMENU Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, _
        WS_EX_DLGMODALFRAME Or WS_EX_CONTROLPARENT Or WS_EX_LEFT )
   if ucase(gConfig.LocalizationFile) = "ENGLISH.LANG" then     
      pWindow->SetClientSize(700, 480)
   else   
      pWindow->SetClientSize(780, 480)
   end if
   pWindow->Center(pWindow->hWindow, hWndParent)
   
   ' Save the incoming pProp value to the ImageManager window. We do this because if the StatusBar
   ' Editor or ToolBar Editor or TabControl Custom called this ImageManager then we need to save 
   ' the Image Name into the pProp.
   pWindow->UserData(0) = cast( LONG_PTR, pProp )
   
   ' Add a tooolbar
   DIM hToolBar AS HWND = pWindow->AddControl("Toolbar", , IDC_FRMIMAGEMANAGER_TOOLBAR, "", 0, 0, 0, 0, _
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

   ' Add buttons to the toolbar
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
      pWindow->AddControl("LISTVIEW", , IDC_FRMIMAGEMANAGER_LISTVIEW, "", 10, nTop, 330, 430, _
      WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or LVS_REPORT Or LVS_SHOWSELALWAYS or LVS_SINGLESEL, _
      WS_EX_LEFT Or WS_EX_CLIENTEDGE or WS_EX_RIGHTSCROLLBAR, , _
      Cast(SUBCLASSPROC, @frmImageManager_Listview_SubclassProc), IDC_FRMIMAGEMANAGER_LISTVIEW, Cast(DWORD_PTR, @pWindow))
                                     
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
   dim pDoc as clsDocument ptr 
   dim as long nSelLine = 0
   dim as long nextLine = 0
   
   pDoc = gApp.pDocList
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
   DIM pImageCtx AS CImageCtx = CImageCtx(pWindow, IDC_FRMIMAGEMANAGER_IMAGECTX, , 370, nTop + 30, 306, 306)

   ' Add a combobox for the various image resize options.
   
   ' Add a button without coordinates (it will be resized in WM_SIZE, below)
   DIM hCtl AS HWND = _
   pWindow->AddControl("COMBOBOX", , IDC_FRMIMAGEMANAGER_COMBO1, "", 356, nTop + 354, 334, 20, _
           WS_CHILD OR WS_VISIBLE OR WS_VSCROLL OR WS_BORDER OR WS_TABSTOP OR CBS_DROPDOWNLIST OR CBS_HASSTRINGS, _
           WS_EX_CLIENTEDGE or WS_EX_LEFT Or WS_EX_LTRREADING)
   ComboBox_AddString( hCtl, @L(376, "Actual Size") )
   ComboBox_AddString( hCtl, @L(165, "Apply High DPI Scaling") )
   ComboBox_AddString( hCtl, @L(375, "Autosize") )
   ComboBox_AddString( hCtl, @L(151, "Fit to Width") )
   ComboBox_AddString( hCtl, @L(160, "Fit to Height") )
   ComboBox_AddString( hCtl, @L(164, "Stretch") )
   if pWindow->DPI > 96 then
      ComboBox_SetCurSel( hCtl, 1 )   ' Use High DPI Scaling
   else
      ComboBox_SetCurSel( hCtl, 0 )   ' Actual Size
   end if
   
   pWindow->AddControl("BUTTON", , IDC_FRMIMAGEMANAGER_CMDFILENAME, L(0014, "Edit") & "...", 356, nTop + 380, 75, 20, _
        WS_CHILD Or WS_VISIBLE or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)

   pWindow->AddControl("LABEL", , IDC_FRMIMAGEMANAGER_LBLFILENAME, "", 356, nTop + 400, 300, 20, _
           WS_CHILD Or WS_VISIBLE Or SS_LEFT or SS_PATHELLIPSIS, WS_EX_LEFT Or WS_EX_LTRREADING)

   pWindow->AddControl("GROUPBOX",, IDC_FRMIMAGEMANAGER_FRAMEPREVIEW, L(377,"Image Preview"), 356, nTop, 334, 350, _
        WS_CHILD Or WS_VISIBLE Or BS_TEXT Or BS_LEFT Or BS_NOTIFY Or BS_GROUPBOX, _
        WS_EX_TRANSPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING)

   ' Set the Listview to the selected line and preview the image display
   ListView_SelectItem(hLV, nSelLine)
   frmImageManager_DisplayImage(pWindow->hWindow, nSelLine)
   
   SetFocus hLV
      
   ' Process Windows messages
   Function = pWindow->DoEvents(SW_SHOW)
   
   ' Delete the frmImages CWindow class manually allocated memory 
   Delete pWindow

End Function





