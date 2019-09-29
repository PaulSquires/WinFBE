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

#include once "frmStatusBarEditor.bi"
#include once "clsPanelItem.bi"
#include once "clsConfig.bi"


' Temporary PanelItem array to hold items while they are being
' edited in the StatusBar Editor. 
dim shared gPanelItems(any) as clsPanelItem


' ========================================================================================
' Create the statusbar to display on the Form
' ========================================================================================
public Function frmStatusBarEditor_CreateFakeStatusBar( ByVal pDoc as clsDocument ptr ) As Long

   if pDoc = 0 then exit function
   
   Dim pWindow As CWindow Ptr = AfxCWindowPtr( pDoc->hWndForm )
   if pWindow = 0 then exit function
   
   ' Only create the StatusBar if PanelItems actually exist.
   if pDoc->StatusBarExists = false then
      ShowWindow( pDoc->hWndStatusBar, SW_HIDE )
      exit function
   end if
   
   dim rc as RECT 
   GetClientRect( pDoc->hWndForm, @rc )
   
   dim as long nHeight = AfxGetWindowHeight( pDoc->hWndStatusBar )
   SetWindowPos( pDoc->hWndStatusBar, 0, _
                  0, rc.bottom - nHeight, rc.right, nHeight, _
                  SWP_NOZORDER )

   
   ' Size the panels in order to accommodate autosize
   dim as long lb = 0 
   dim as long ub = ubound(pDoc->PanelItems)
   dim as long nPanelWidth, nPanelMinWidth, nBorderStyle
   Dim As Long nWidths(ub)
   dim hIcon as HICON
   dim wszImageName as wstring * MAX_PATH
   

   Dim pWindowMain As CWindow Ptr = AfxCWindowPtr( HWND_FRMMAIN )
   SendMessage( pDoc->hWndStatusBar, WM_SETFONT, cast(WPARAM, pWindowMain->font), cast(LPARAM, CTRUE))

   dim as single cx = pWindow->rxRatio
   
   for i as long = lb to ub
      dim as long nPanelWidth    = val( pDoc->PanelItems(i).wszWidth ) * cx
      dim as long nPanelMinWidth = val( pDoc->PanelItems(i).wszMinWidth ) * cx
      
      dim as single cx = pWindow->rxRatio
   
      select case ucase( pDoc->PanelItems(i).wszAutoSize )
         case "STATUSBARPANELAUTOSIZE.NONE"
            nPanelWidth    = val( pDoc->PanelItems(i).wszWidth ) * cx
            nPanelMinWidth = val( pDoc->PanelItems(i).wszMinWidth ) * cx
            if nPanelWidth < nPanelMinWidth then nPanelWidth = nPanelMinWidth
         case "STATUSBARPANELAUTOSIZE.CONTENTS"
            dim as HDC _hdc = GetDC(pDoc->hWndStatusBar)
            dim as SIZE _size 
            dim as HFONT oldfont = SelectObject( _hdc, pWindowMain->font )
            dim as WString * MAX_PATH wszBuffer = " " & pDoc->PanelItems(i).wszText & " "
            GetTextExtentPoint32( _hdc, wszBuffer, len(wszBuffer), @_size)
            SelectObject( _hdc, oldfont )
            ReleaseDC( pDoc->hWndStatusBar, _hdc )
            nPanelWidth = _size.cx
            nPanelMinWidth = val( pDoc->PanelItems(i).wszMinWidth ) * cx
            if nPanelWidth < nPanelMinWidth then nPanelWidth = nPanelMinWidth
         case "STATUSBARPANELAUTOSIZE.SPRING"
            nPanelWidth = iif(i, rc.right - nWidths(i-1), rc.right)
      end select
      nWidths(i) = iif(i, nWidths(i-1), 0) + nPanelWidth
   NEXT
   StatusBar_SetParts( pDoc->hWndStatusBar, (ub-lb)+1, @nWidths(0))

   ' Set the text/icons for the panels. Need to do this after the widths are set.
   for i as long = lb to ub
      dim as CWSTR wszPanelText = pDoc->PanelItems(i).wszText
      select case ucase( pDoc->PanelItems(i).wszAlignment )
         case "STATUSBARPANELALIGNMENT.LEFT":   ' do nothing
         case "STATUSBARPANELALIGNMENT.CENTER": wszPanelText = chr(9) & wszPanelText
         case "STATUSBARPANELALIGNMENT.RIGHT":  wszPanelText = chr(9,9) & wszPanelText
      end select
      select case ucase( pDoc->PanelItems(i).wszBorderStyle )
         case "STATUSBARPANELBORDERSTYLE.NONE":   nBorderStyle = SBT_NOBORDERS
         case "STATUSBARPANELBORDERSTYLE.SUNKEN": nBorderStyle = 0
         case "STATUSBARPANELBORDERSTYLE.RAISED": nBorderStyle = SBT_POPOUT
      end select

      StatusBar_SetText( pDoc->hWndStatusBar, i, wszPanelText, nBorderStyle ) 

      wszImageName = pDoc->PanelItems(i).pProp.wszPropValue
      hIcon = 0
      if len(wszImageName) then
         dim pImageType as IMAGES_TYPE ptr = GetImagesTypePtr(wszImageName)
         DestroyIcon( StatusBar_GetIcon( pDoc->hWndStatusBar, i ) )
         if pImageType then
            hIcon = AfxGdipIconFromFile( pImageType->wszFileName )
         end if
      end if

      ' If no hIcon defined then any previous icon will be removed (hIcon = 0)
      StatusBar_SetIcon( pDoc->hWndStatusBar, i, hIcon ) 
   next

   ShowWindow( pDoc->hWndStatusBar, SW_SHOW )

   function = 0 
End Function
                  

' ========================================================================================
' Display the entries in the StatusBarEditor Listbox
' ========================================================================================
private Function frmStatusBarEditor_DisplayListBox( ByVal nIndex As Long ) As Long

   ' Display the panel items in the listbox
   dim hList as hwnd = GetDlgItem( HWND_FRMSTATUSBAREDITOR, IDC_FRMSTATUSBAREDITOR_LSTPANELS )

   ListBox_ResetContent( hList )
   for i as long = 0 to ubound(gPanelItems)
      gPanelItems(i).wszName = "StatusBar Panel " & i
      ListBox_AddString( hList, gPanelItems(i).wszName.sptr )
   next
   ListBox_SetCurSel( hList, nIndex )

   function = 0
End Function


' ========================================================================================
' Swap two entries in the StatusBarEditor Listbox
' ========================================================================================
private function frmStatusBarEditor_SwapListBoxItems( byval Item1 as long, _
                                                      Byval Item2 as long _
                                                      ) as Long
   ' Do not swap the Panel Name
   dim as CWSTR wszPanelName1 = gPanelItems(Item1).wszName
   dim as CWSTR wszPanelName2 = gPanelItems(Item2).wszName
   
   ' Swap the array values
   swap gPanelItems(Item1), gPanelItems(Item2)
    
   gPanelItems(Item1).wszName = wszPanelName1
   gPanelItems(Item2).wszName = wszPanelName2 

   frmStatusBarEditor_DisplayListBox( Item2 )
   
   function = 0
end function


' ========================================================================================
' Set the frmStatusBarEditor textboxes and options depending on what listbox entry is selected
' ========================================================================================
private function frmStatusBarEditor_SetTextboxes() as long
   dim as HWND hListBox = GetDlgItem(HWND_FRMSTATUSBAREDITOR, IDC_FRMSTATUSBAREDITOR_LSTPANELS)
   dim as long nCurSel = ListBox_GetCurSel( hListBox )
   dim as Boolean fEnabled = iif( nCurSel < 0, false, true)
   if ListBox_GetCount( hListBox ) = 0 then fEnabled = false

   dim wszText as WString * MAX_PATH
   dim as HWND hCtrl
   dim as long nIndex
   
   hCtrl = GetDlgItem(HWND_FRMSTATUSBAREDITOR, IDC_FRMSTATUSBAREDITOR_TXTTEXT)
   EnableWindow( hCtrl, fEnabled)
   if nCurSel <> -1 then wszText = gPanelItems(nCurSel).wszText else wszText = ""
   AfxSetWindowText( hCtrl, wszText )

   hCtrl = GetDlgItem(HWND_FRMSTATUSBAREDITOR, IDC_FRMSTATUSBAREDITOR_TXTTOOLTIP)
   EnableWindow( hCtrl, fEnabled)
   if nCurSel <> -1 then wszText = gPanelItems(nCurSel).wszTooltip else wszText = ""
   AfxSetWindowText( hCtrl, wszText )

   hCtrl = GetDlgItem(HWND_FRMSTATUSBAREDITOR, IDC_FRMSTATUSBAREDITOR_COMBOALIGNMENT)
   if nCurSel <> -1 then wszText = gPanelItems(nCurSel).wszAlignment else wszText = ""
   nIndex = ComboBox_FindStringExact( hCtrl, -1, @wszText ) 
   ComboBox_SetCurSel( hCtrl, nIndex)
   EnableWindow( hCtrl, fEnabled)

   hCtrl = GetDlgItem(HWND_FRMSTATUSBAREDITOR, IDC_FRMSTATUSBAREDITOR_COMBOBORDERSTYLE)
   if nCurSel <> -1 then wszText = gPanelItems(nCurSel).wszBorderStyle else wszText = ""
   nIndex = ComboBox_FindStringExact( hCtrl, -1, @wszText ) 
   ComboBox_SetCurSel( hCtrl, nIndex)
   EnableWindow( hCtrl, fEnabled)

   hCtrl = GetDlgItem(HWND_FRMSTATUSBAREDITOR, IDC_FRMSTATUSBAREDITOR_COMBOAUTOSIZE)
   if nCurSel <> -1 then wszText = gPanelItems(nCurSel).wszAutoSize else wszText = ""
   nIndex = ComboBox_FindStringExact( hCtrl, -1, @wszText ) 
   ComboBox_SetCurSel( hCtrl, nIndex)
   EnableWindow( hCtrl, fEnabled)

   hCtrl = GetDlgItem(HWND_FRMSTATUSBAREDITOR, IDC_FRMSTATUSBAREDITOR_TXTWIDTH)
   EnableWindow( hCtrl, fEnabled)
   if nCurSel <> -1 then wszText = gPanelItems(nCurSel).wszWidth else wszText = ""
   AfxSetWindowText( hCtrl, wszText )

   hCtrl = GetDlgItem(HWND_FRMSTATUSBAREDITOR, IDC_FRMSTATUSBAREDITOR_TXTMINWIDTH)
   EnableWindow( hCtrl, fEnabled)
   if nCurSel <> -1 then wszText = gPanelItems(nCurSel).wszMinWidth else wszText = ""
   AfxSetWindowText( hCtrl, wszText )

   hCtrl = GetDlgItem(HWND_FRMSTATUSBAREDITOR, IDC_FRMSTATUSBAREDITOR_CMDPICSELECT)
   EnableWindow( hCtrl, fEnabled )

   ' Display any panel image
   dim pImageCtx as CImageCtx PTR = AfxCImageCtxPtr(HWND_FRMSTATUSBAREDITOR, IDC_FRMSTATUSBAREDITOR_PICIMAGE)
   if pImageCtx = 0 then exit function
   
   if nCurSel <> -1 then
      dim pImageType as IMAGES_TYPE ptr 
      dim wszImageName as wstring * MAX_PATH
      dim as CWSTR wszFilename 
      wszImageName = gPanelItems(nCurSel).pProp.wszPropValue 
      pImageType = GetImagesTypePtr(wszImageName)

      hCtrl = GetDlgItem( HWND_FRMSTATUSBAREDITOR, IDC_FRMSTATUSBAREDITOR_PICIMAGE )
      if pImageType then 
         wszFilename = pImageType->wszFileName
         pImageCtx->LoadImageFromFile(wszFilename)
         Dim pWindow As CWindow Ptr = AfxCWindowPtr(HWND_FRMSTATUSBAREDITOR)
         if pWindow then
            pImageCtx->SetImageWidth( pWindow->ScaleX(pImageCtx->GetImageWidth) )
            pImageCtx->SetImageHeight( pWindow->ScaleY(pImageCtx->GetImageHeight) )
         end if
         pImageCtx->SetImageAdjustment( GDIP_IMAGECTX_ACTUALSIZE, CTRUE )
         ShowWindow( hCtrl, SW_SHOW )
      else
         ShowWindow( hCtrl, SW_HIDE )
      end if
   end if
   
   function = 0
end function


' ========================================================================================
' Process WM_CREATE message for window/dialog: frmStatusBarEditor
' ========================================================================================
private Function frmStatusBarEditor_OnCreate( ByVal HWnd As HWnd, _
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
' Process WM_COMMAND message for window/dialog: frmStatusBarEditor
' ========================================================================================
private Function frmStatusBarEditor_OnCommand( ByVal HWnd As HWnd, _
                                               ByVal id As Long, _
                                               ByVal hwndCtl As HWnd, _
                                               ByVal codeNotify As UINT _
                                               ) As LRESULT

   dim as hwnd hList1 = GetDlgItem( HWND, IDC_FRMSTATUSBAREDITOR_LSTPANELS)
   dim as long nCurSel = ListBox_GetCurSel(hList1)

   Select Case id
      case IDC_FRMSTATUSBAREDITOR_LSTPANELS
         if codeNotify = LBN_SELCHANGE THEN
            frmStatusBarEditor_SetTextboxes()
         END IF
           
      
      CASE IDC_FRMSTATUSBAREDITOR_TXTTEXT
         if codeNotify = EN_CHANGE THEN
            if nCurSel > -1 THEN gPanelItems(nCurSel).wszText = AfxGetWindowText(hwndCtl)
         end if

      
      CASE IDC_FRMSTATUSBAREDITOR_TXTTOOLTIP
         if codeNotify = EN_CHANGE THEN
            if nCurSel > -1 THEN gPanelItems(nCurSel).wszTooltip = AfxGetWindowText(hwndCtl)
         end if
           
      
      case IDC_FRMSTATUSBAREDITOR_COMBOALIGNMENT
         if codeNotify = CBN_SELCHANGE THEN
            if nCurSel > -1 THEN gPanelItems(nCurSel).wszAlignment = AfxGetWindowText(hwndCtl) 
         END IF

      
      case IDC_FRMSTATUSBAREDITOR_COMBOBORDERSTYLE
         if codeNotify = CBN_SELCHANGE THEN
            if nCurSel > -1 THEN gPanelItems(nCurSel).wszBorderStyle = AfxGetWindowText(hwndCtl) 
         END IF

      
      case IDC_FRMSTATUSBAREDITOR_COMBOAUTOSIZE
         if codeNotify = CBN_SELCHANGE THEN
            if nCurSel > -1 THEN gPanelItems(nCurSel).wszAutoSize = AfxGetWindowText(hwndCtl) 
         END IF

      
      CASE IDC_FRMSTATUSBAREDITOR_TXTWIDTH
         if codeNotify = EN_CHANGE THEN
            if nCurSel > -1 THEN gPanelItems(nCurSel).wszWidth = AfxGetWindowText(hwndCtl)
         end if


      CASE IDC_FRMSTATUSBAREDITOR_TXTMINWIDTH
         if codeNotify = EN_CHANGE THEN
            if nCurSel > -1 THEN gPanelItems(nCurSel).wszMinWidth = AfxGetWindowText(hwndCtl)
         end if

      
      case IDC_FRMSTATUSBAREDITOR_CMDPANELUP
         if codeNotify = BN_CLICKED THEN
            if nCurSel > 0 THEN
               frmStatusBarEditor_SwapListboxItems(nCurSel, nCurSel - 1)
            END IF
         end if

      
      case IDC_FRMSTATUSBAREDITOR_CMDPANELDOWN
         if codeNotify = BN_CLICKED THEN
            if nCurSel < ListBox_GetCount(hList1)-1 THEN
               frmStatusBarEditor_SwapListboxItems(nCurSel, nCurSel + 1)
            END IF
         end if

      
      case IDC_FRMSTATUSBAREDITOR_CMDPANELADD
         if codeNotify = BN_CLICKED THEN
            if ubound(gPanelItems) = -1 THEN
               redim gPanelItems(0)
               nCurSel = 0
            Else
               Redim Preserve gPanelItems(Ubound(gPanelItems)+1)
               nCurSel = Ubound(gPanelItems)
            END IF
            dim newPanelItem as clsPanelItem
            gPanelItems(nCurSel) = newPanelItem
            frmStatusBarEditor_DisplayListBox( nCurSel )
            frmStatusBarEditor_SetTextBoxes
         end if


      case IDC_FRMSTATUSBAREDITOR_CMDPANELINSERT
         if codeNotify = BN_CLICKED THEN
            if ubound(gPanelItems) = -1 THEN
               redim gPanelItems(0)
               nCurSel = 0
            Else
               Redim Preserve gPanelItems(Ubound(gPanelItems)+1)
               if nCurSel = -1 THEN nCurSel = 0
               for i as long = ubound(gPanelItems) to nCurSel + 1 step -1
                  gPanelItems(i) = gPanelItems(i-1)
               NEXT
            END IF
            dim newPanelItem as clsPanelItem
            gPanelItems(nCurSel) = newPanelItem
            frmStatusBarEditor_DisplayListBox( nCurSel )
            frmStatusBarEditor_SetTextBoxes
         end if
         

      case IDC_FRMSTATUSBAREDITOR_CMDPANELDELETE
         If codeNotify = BN_CLICKED Then
            if nCurSel > -1  THEN
               if ubound(gPanelItems) = 0 THEN
                  erase gPanelItems
                  nCurSel = -1
               else
                  ' remove the item from the internal array
                  for i as long = nCurSel to ubound(gPanelItems) - 1
                     gPanelItems(i) = gPanelItems(i+1)
                  NEXT
                  redim preserve gPanelItems(ubound(gPanelItems)-1)
               END IF
               nCurSel = Min(nCurSel, ubound(gPanelItems))
               frmStatusBarEditor_DisplayListBox( nCurSel )
               frmStatusBarEditor_SetTextBoxes
            end if
         end if


      case IDC_FRMSTATUSBAREDITOR_CMDPICSELECT
         If codeNotify = BN_CLICKED Then
            if nCurSel > -1 then
               frmImageManager_Show( HWND, @gPanelItems(nCurSel).pProp )
               frmStatusBarEditor_SetTextBoxes
            end if
         end if
         
         
      Case IDC_FRMSTATUSBAREDITOR_CMDOK    
         If codeNotify = BN_CLICKED Then
            ' Copy the temporary array back to the form's panel items 
            Dim pDoc As clsDocument Ptr = gTTabCtl.GetActiveDocumentPtr()
            If pDoc Then 
               dim as long numItems = ubound(gPanelItems)
               if numItems > -1 then
                  redim pDoc->PanelItems(numItems)
                  for i as long = lbound(gPanelItems) to ubound(gPanelItems)
                     pDoc->PanelItems(i) = gPanelItems(i) 
                  next
               end if
               ' Save the value indicating that we want the statusbar to display in
               ' the generated code for the form.
               pDoc->GenerateStatusBar = Button_GetCheck( GetDlgItem(HWND, IDC_FRMSTATUSBAREDITOR_CHKDISPLAYONFORM) )
               pDoc->UserModified = true
            end if
            SendMessage(HWnd, WM_CLOSE, 0, 0)
            Exit Function
         end if

             
      Case IDCANCEL
         If codeNotify = BN_CLICKED Then
            SendMessage( HWnd, WM_CLOSE, 0, 0 )
            Exit Function
         End If
   
   End Select

   Function = 0
End Function


' ========================================================================================
' Process WM_CLOSE message for window/dialog: frmStatusBarEditor
' ========================================================================================
private Function frmStatusBarEditor_OnClose( byval HWnd As HWnd ) As LRESULT
   ' Reset the global PanelItems array
   erase gPanelItems
   ' Enables parent window keeping parent's zorder
   EnableAllModeless()
   DestroyWindow( HWnd )
   Function = 0
End Function


' ========================================================================================
' Process WM_DESTROY message for window/dialog: frmStatusBarEditor
' ========================================================================================
private Function frmStatusBarEditor_OnDestroy( byval HWnd As HWnd ) As LRESULT
   PostQuitMessage(0)
   Function = 0
End Function


' ========================================================================================
' frmStatusBarEditor Window procedure
' ========================================================================================
private Function frmStatusBarEditor_WndProc( ByVal HWnd   As HWnd, _
                                             ByVal uMsg   As UINT, _
                                             ByVal wParam As WPARAM, _
                                             ByVal lParam As LPARAM _
                                             ) As LRESULT

   Select Case uMsg
      HANDLE_MSG (HWnd, WM_CREATE,   frmStatusBarEditor_OnCreate)
      HANDLE_MSG (HWnd, WM_CLOSE,    frmStatusBarEditor_OnClose)
      HANDLE_MSG (HWnd, WM_DESTROY,  frmStatusBarEditor_OnDestroy)
      HANDLE_MSG (HWnd, WM_COMMAND,  frmStatusBarEditor_OnCommand)
   End Select

   Function = DefWindowProc(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' frmStatusBarEditor_Show
' ========================================================================================
public Function frmStatusBarEditor_Show( ByVal hWndParent As HWnd ) as LRESULT

   DIM hBitmap AS HBITMAP
   dim hCtrl as HWnd
   dim wszImage as wstring * 100
   

   '  Create the main window and child controls
   Dim pWindow as CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowOwnerPtr(hwndParent)->DPI
   
   Dim As HWnd hForm = _
   pWindow->Create( hWndParent, L(314,"Statusbar Editor"), _
        @frmStatusBarEditor_WndProc, 0, 0, 508, 380, _
        WS_POPUP Or WS_CAPTION Or WS_SYSMENU Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, _
        WS_EX_DLGMODALFRAME Or WS_EX_CONTROLPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->SetClientSize(508, 380)
   pWindow->Center
        
   HWND_FRMSTATUSBAREDITOR = hForm

   pWindow->AddControl("LISTBOX", , IDC_FRMSTATUSBAREDITOR_LSTPANELS, "", 21, 14, 465, 84, _
        WS_CHILD Or WS_VISIBLE Or WS_VSCROLL Or WS_TABSTOP Or LBS_NOTIFY Or LBS_NOINTEGRALHEIGHT, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)

   hCtrl = _
   pWindow->AddControl("BITMAPBUTTON", , IDC_FRMSTATUSBAREDITOR_CMDPANELUP, "", 21, 102, 24, 24, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or BS_BITMAP Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER Or LR_DEFAULTCOLOR Or LR_LOADMAP3DCOLORS Or LR_LOADTRANSPARENT Or LR_SHARED, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   wszImage = iif(pWindow->DPI > 96, "IMAGE_ARROWUP", "IMAGE_ARROWUP16")
         hBitmap = AfxGdipImageFromRes(pWindow->InstanceHandle, wszImage, 0, false, IMAGE_BITMAP, 0)
         SendMessage(hCtrl, BM_SETIMAGE, IMAGE_BITMAP, cast(LPARAM, hBitmap))
         IF hBitmap THEN DeleteObject(hBitmap)

   hCtrl = _
   pWindow->AddControl("BITMAPBUTTON", , IDC_FRMSTATUSBAREDITOR_CMDPANELDOWN, "", 48, 102, 24, 24, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or BS_BITMAP Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER Or LR_DEFAULTCOLOR Or LR_LOADMAP3DCOLORS Or LR_LOADTRANSPARENT Or LR_SHARED, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   wszImage = iif(pWindow->DPI > 96, "IMAGE_ARROWDOWN", "IMAGE_ARROWDOWN16")
         hBitmap = AfxGdipImageFromRes(pWindow->InstanceHandle, wszImage, 0, false, IMAGE_BITMAP, 0)
         SendMessage(hCtrl, BM_SETIMAGE, IMAGE_BITMAP, cast(LPARAM, hBitmap))
         IF hBitmap THEN DeleteObject(hBitmap)

   pWindow->AddControl("BUTTON", , IDC_FRMSTATUSBAREDITOR_CMDPANELADD, L(380, "Add"), 77, 102, 70, 24, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_FLAT Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("BUTTON", , IDC_FRMSTATUSBAREDITOR_CMDPANELINSERT, L(281, "Insert"), 152, 102, 70, 24, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_FLAT Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("BUTTON", , IDC_FRMSTATUSBAREDITOR_CMDPANELDELETE, L(282, "Delete"), 227, 102, 70, 24, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_FLAT Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)


   pWindow->AddControl("LABEL", , IDC_FRMSTATUSBAREDITOR_LABEL1, L(150,"Text") & ":", 21, 135, 93, 16, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("TEXTBOX", , IDC_FRMSTATUSBAREDITOR_TXTTEXT, "", 114, 134, 372, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)

   pWindow->AddControl("LABEL", , IDC_FRMSTATUSBAREDITOR_LABEL2, L(235,"Tooltip Text") & ":", 21, 160, 93, 16, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("TEXTBOX", , IDC_FRMSTATUSBAREDITOR_TXTTOOLTIP, "", 114, 158, 372, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)

   pWindow->AddControl("LABEL", , IDC_FRMSTATUSBAREDITOR_LABEL3, L(236,"Alignment") & ":", 21, 194, 93, 21, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("COMBOBOX", , IDC_FRMSTATUSBAREDITOR_COMBOALIGNMENT, "", 114, 190, 207, 21, _
        WS_CHILD Or WS_VISIBLE Or WS_VSCROLL Or WS_TABSTOP Or CBS_DROPDOWNLIST, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)

   pWindow->AddControl("LABEL", , IDC_FRMSTATUSBAREDITOR_LABEL5, L(238,"Border Style") & ":", 21, 222, 93, 21, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("COMBOBOX", , IDC_FRMSTATUSBAREDITOR_COMBOBORDERSTYLE, "", 114, 219, 207, 21, _
        WS_CHILD Or WS_VISIBLE Or WS_VSCROLL Or WS_TABSTOP Or CBS_DROPDOWNLIST, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)

   pWindow->AddControl("LABEL", , IDC_FRMSTATUSBAREDITOR_LABEL6, L(239,"Autosize") & ":", 21, 250, 93, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("COMBOBOX", , IDC_FRMSTATUSBAREDITOR_COMBOAUTOSIZE, "", 114, 247, 207, 21, _
        WS_CHILD Or WS_VISIBLE Or WS_VSCROLL Or WS_TABSTOP Or CBS_DROPDOWNLIST, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)

   pWindow->AddControl("LABEL", , IDC_FRMSTATUSBAREDITOR_LABEL8, L(381,"Width") & ":", 21, 277, 93, 21, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("TEXTBOX", , IDC_FRMSTATUSBAREDITOR_TXTWIDTH, "", 114, 273, 207, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)

   pWindow->AddControl("LABEL", , IDC_FRMSTATUSBAREDITOR_LABEL7, L(245,"Minimum Width") & ":", 21, 302, 93, 21, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("TEXTBOX", , IDC_FRMSTATUSBAREDITOR_TXTMINWIDTH, "", 114, 298, 207, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)


   pWindow->AddControl("BUTTON", , IDC_FRMSTATUSBAREDITOR_CMDPICSELECT, "...", 350, 295, 26, 16, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)

   ' Add an image control                                                         
   DIM pImageCtx AS CImageCtx = CImageCtx(pWindow, IDC_FRMSTATUSBAREDITOR_PICIMAGE, , 384, 224, 60, 60)

   pWindow->AddControl("GROUPBOX", , IDC_FRMSTATUSBAREDITOR_FRAME1, L(246,"Image"), 342, 190, 144, 128, _
        WS_CHILD Or WS_VISIBLE Or BS_TEXT Or BS_LEFT Or BS_NOTIFY Or BS_GROUPBOX, _
        WS_EX_TRANSPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING)

   pWindow->AddControl("CHECKBOX", , IDC_FRMSTATUSBAREDITOR_CHKDISPLAYONFORM, L(237,"Display on Form"), 21, 345, 175, 16, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)

   pWindow->AddControl("BUTTON", , IDC_FRMSTATUSBAREDITOR_CMDOK, L(0,"OK"), 331, 333, 74, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("BUTTON", , IDCANCEL, L(1,"Cancel"), 412, 333, 74, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)


   ' Copy the form's panel items to the temporary array for editing
   Dim pDoc As clsDocument Ptr = gTTabCtl.GetActiveDocumentPtr()
   If pDoc Then 
      hCtrl = GetDlgItem( hForm, IDC_FRMSTATUSBAREDITOR_CHKDISPLAYONFORM ) 
      Button_SetCheck( hCtrl, pDoc->GenerateStatusBar )
      dim as long numItems = ubound(pDoc->PanelItems)
      if numItems > -1 then
         redim gPanelItems(numItems)
         for i as long = 0 to numItems
            gPanelItems(i) = pDoc->PanelItems(i)
         next
      end if
   end if
   

   ' Load the comboboxes
   hCtrl = GetDlgItem( hForm, IDC_FRMSTATUSBAREDITOR_COMBOALIGNMENT )
   ComboBox_AddString( hCtrl, @wstr("StatusBarPanelAlignment.Left") )
   ComboBox_AddString( hCtrl, @wstr("StatusBarPanelAlignment.Center") )
   ComboBox_AddString( hCtrl, @wstr("StatusBarPanelAlignment.Right") )
  
   hCtrl = GetDlgItem( hForm, IDC_FRMSTATUSBAREDITOR_COMBOBORDERSTYLE )
   ComboBox_AddString( hCtrl, @wstr("StatusBarPanelBorderStyle.None") )
   ComboBox_AddString( hCtrl, @wstr("StatusBarPanelBorderStyle.Sunken") )
   ComboBox_AddString( hCtrl, @wstr("StatusBarPanelBorderStyle.Raised") )

   hCtrl = GetDlgItem( hForm, IDC_FRMSTATUSBAREDITOR_COMBOAUTOSIZE )
   ComboBox_AddString( hCtrl, @wstr("StatusBarPanelAutoSize.None") )
   ComboBox_AddString( hCtrl, @wstr("StatusBarPanelAutoSize.Contents") )
   ComboBox_AddString( hCtrl, @wstr("StatusBarPanelAutoSize.Spring") )

   
   frmStatusBarEditor_DisplayListBox( 0 )
   frmStatusBarEditor_SetTextboxes()
   SetFocus( GetDlgItem( HWND_FRMSTATUSBAREDITOR, IDC_FRMSTATUSBAREDITOR_LSTPANELS ) )

   ' Process Windows messages
   Function = pWindow->DoEvents( SW_SHOW )
   
   ' Delete the frmStatusBarEditor CWindow class manually allocated memory 
   Delete pWindow

   function = 0
end function

