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

#include once "frmToolBarEditor.bi"
#include once "clsToolBarItem.bi"
#include once "clsDocument.bi"


' Temporary ToolBarItem array to hold items while they are being
' edited in the Menu Editor. 
dim shared gToolBarItems(any) as clsToolBarItem


' ========================================================================================
' Processes messages for the subclassed Rebar/ToolBar.
' ========================================================================================
public FUNCTION FakeToolBar_SubclassProc( BYVAL hwnd   AS HWND, _                 ' Control window handle
                                          BYVAL uMsg   AS UINT, _                 ' Type of message
                                          BYVAL wParam AS WPARAM, _               ' First message parameter
                                          BYVAL lParam AS LPARAM, _               ' Second message parameter
                                          BYVAL uIdSubclass AS UINT_PTR, _        ' The subclass ID
                                          BYVAL dwRefData AS DWORD_PTR _          ' Pointer to reference data
                                          ) AS LRESULT

   dim pDoc as clsDocument ptr = cast(clsDocument ptr, dwRefData)
   
   SELECT CASE uMsg

      CASE WM_GETDLGCODE
         ' All keyboard input
         FUNCTION = DLGC_WANTALLKEYS
         EXIT FUNCTION

      case WM_LBUTTONUP
         PostMessage( HWND_FRMMAIN, WM_COMMAND, MAKELONG(IDM_TOOLBAREDITOR, 0), 0 )

      Case WM_MOUSEACTIVATE
         ' Defeat this message so that mouse clicks do not activate the control.
         Function = MA_NOACTIVATE: uMsg = WM_NULL
         Exit Function

      Case WM_SETCURSOR
         Function = CTRUE: uMsg = WM_NULL
         Exit Function
          
      Case WM_SETFOCUS
         ' Defeat the caret activation, for some
         ' reason MA_NOACTIVATE does not work for right clicks.
         Function = 0: uMsg = WM_NULL
         Exit Function

      CASE WM_DESTROY
         ' REQUIRED: Remove control subclassing
         RemoveWindowSubclass( hwnd, @FakeToolBar_SubclassProc, uIdSubclass )

   END SELECT

   ' Default processing of Windows messages
   FUNCTION = DefSubclassProc(hwnd, uMsg, wParam, lParam)

END FUNCTION


' ========================================================================================
' Create the "fake" toolbar to display on the Form
' ========================================================================================
public Function frmToolBarEditor_CreateFakeToolBar( ByVal pDoc as clsDocument ptr ) As Long

   if pDoc = 0 then exit function
   
   ' If Rebar and ToolBar already exist then recreate them
   if (pDoc->hWndRebar) andalso (pDoc->hWndToolBar) then 
      DestroyWindow pDoc->hWndRebar:   pDoc->hWndRebar = 0
      DestroyWindow pDoc->hWndToolBar: pDoc->hWndToolBar = 0
   end if
   if pDoc->ToolBarExists = false then exit function
   
    
   Dim pWindow As CWindow Ptr = AfxCWindowPtr( pDoc->hWndForm )
   if pWindow = 0 then exit function
    
   dim rc as RECT 
   GetClientRect( pDoc->hWndForm, @rc )
   
   ' Create the Rebar and ToolBar
   Dim As HWnd hRebar = _
      pWindow->AddControl("Rebar", pDoc->hWndForm, IDC_FAKEREBAR, "", 0, 0, 0, 0, _
      WS_CHILD or WS_VISIBLE OR WS_BORDER OR WS_CLIPCHILDREN OR _
      WS_CLIPSIBLINGS OR CCS_NODIVIDER OR RBS_VARHEIGHT OR RBS_BANDBORDERS Or CCS_NOPARENTALIGN)
   
   Dim As HWnd hToolBar = _
      pWindow->AddControl("Toolbar", pDoc->hWndForm, IDC_FAKETOOLBAR, "", 0, 0, 0, 0, _
      WS_CHILD Or WS_VISIBLE Or TBSTYLE_TOOLTIPS Or TBSTYLE_FLAT Or CCS_NODIVIDER Or CCS_NOPARENTALIGN)

   pDoc->hWndRebar = hRebar
   pDoc->hWndToolBar = hToolBar

   ' Allow drop down arrows
   SendMessage( hToolBar, TB_SETEXTENDEDSTYLE, 0, TBSTYLE_EX_DRAWDDARROWS )

   ' Determine the Image size
   dim as long nImageSize = 24
   select case **pDoc->wszToolBarSize
      case "SIZE_16":  nImageSize = 16
      case "SIZE_24":  nImageSize = 24
      case "SIZE_32":  nImageSize = 32
      case "SIZE_48":  nImageSize = 48
   end select

   ' We only need to create a Normal ImageList for purposes of displaying in the Designer.
   Dim As HIMAGELIST hImageListNormal
   Dim cx As Long = nImageSize * pWindow->DPI \ 96
   hImageListNormal = ImageList_Create( cx, cx, ILC_COLOR32 Or ILC_MASK, 20, 1)
   dim wszImageName as wstring * 1024

   ' Add buttons to the toolbar
   Dim As Long i, ii
   
   dim as long lb = 0 
   dim as long ub = ubound(pDoc->ToolBarItems)
    
   for i as long = lb to ub
       
      ii = -1
      wszImageName = pDoc->ToolBarItems(i).pPropNormalImage.wszPropValue 
      if len(wszImageName) then
         dim pImageType as IMAGES_TYPE ptr = GetImagesTypePtr(wszImageName)
         if pImageType then
            ii = AfxGdipAddIconFromFile(hImageListNormal, pImageType->wszFileName)
         end if
      end if

      select case **pDoc->ToolBarItems(i).wszButtonType 
         case "ToolBarButton.Button"
            Toolbar_AddButton hToolBar, ii, 0
         
         case "ToolBarButton.Separator"
            Toolbar_AddSeparator hToolBar
         
         case "ToolBarButton.DropDown"
            Toolbar_AddButton hToolBar, ii, 0, 0, TBSTYLE_DROPDOWN
         
         case "ToolBarButton.WholeDropDown"
            Toolbar_AddButton hToolBar, ii, 0, 0, BTNS_WHOLEDROPDOWN
      end select      

   next

   ' Attach the ImageList to the toolbar
   SendMessage hToolBar, TB_SETIMAGELIST, 0, Cast(LPARAM, hImageListNormal)

   ' Size the toolbar
   SendMessage hToolBar, TB_AUTOSIZE, 0, 0

   ' Add the band containing the toolbar control to the rebar
   ' The size of the REBARBANDINFOW is different in Vista/Windows 7
   Dim rbbi As REBARBANDINFOW
   If (AfxWindowsVersion >= 600) AndAlso (AfxComCtlVersion >= 600) Then
      rbbi.cbSize  = REBARBANDINFO_V6_SIZE
   Else
      rbbi.cbSize  = REBARBANDINFO_V3_SIZE
   End If

   ' Insert the toolbar in the rebar control
   rbbi.fMask      = RBBIM_STYLE Or RBBIM_CHILD Or RBBIM_CHILDSIZE Or _
                     RBBIM_SIZE Or RBBIM_ID Or RBBIM_IDEALSIZE Or RBBIM_COLORS
   rbbi.fStyle     = RBBS_CHILDEDGE Or RBBS_GRIPPERALWAYS
   rbbi.hwndChild  = hToolbar
   rbbi.cxMinChild = 270 * pWindow->rxRatio
   rbbi.cyMinChild = Hiword(SendMessageW(hToolBar, TB_GETBUTTONSIZE, 0, 0))
   rbbi.cx         = 270 * pWindow->rxRatio
   rbbi.cxIdeal    = 270 * pWindow->rxRatio
   rbbi.clrFore    = GetSysColor(COLOR_BTNTEXT)
   rbbi.clrBack    = GetSysColor(COLOR_BTNFACE)

   ' Insert band into rebar
   SendMessage hRebar, RB_INSERTBANDW, -1, Cast(LPARAM, @rbbi)

   SetWindowSubclass( hRebar, CAST(SUBCLASSPROC, @FakeToolBar_SubclassProc), 0, CAST(DWORD_PTR, pDoc))
   SetWindowSubclass( hToolBar, CAST(SUBCLASSPROC, @FakeToolBar_SubclassProc), 0, CAST(DWORD_PTR, pDoc))

   ShowWindow( hToolBar, SW_SHOW )
   ShowWindow( hRebar, SW_SHOW )

    ' Position the Rebar below any existing fake topmenu.
   dim as long nTopMenu_Height, nToolBar_Height 
   nToolBar_Height = AfxGetWindowHeight( hRebar )  
   if pDoc->hWndFakeMenu then
      nTopMenu_Height = AfxGetWindowHeight( pDoc->hWndFakeMenu )  
   end if 
   SetWindowPos hRebar, 0, 0, nTopMenu_Height, rc.right, nToolBar_Height, SWP_NOZORDER
    
   function = 0 
End Function 


' ========================================================================================
' Display the details of this ToolBarItem
' ========================================================================================
private Function frmToolBarEditor_DisplayToolBarItem() As Long

   dim as long nCurSel, nIndex
   dim as Boolean bEnabled = true
    
   nCurSel    = Listbox_GetCurSel( GetDlgItem(HWND_FRMTOOLBAREDITOR, IDC_FRMTOOLBAREDITOR_LSTBUTTONS) )
   If nCurSel = -1 Then Exit Function
    
   ' Button type
   nIndex = ComboBox_FindStringExact( GetDlgItem(HWND_FRMTOOLBAREDITOR, IDC_FRMTOOLBAREDITOR_COMBOBUTTONTYPE), -1, gToolBarItems(nCurSel).wszButtonType.sptr) 
   ComboBox_SetCurSel( GetDlgItem(HWND_FRMTOOLBAREDITOR, IDC_FRMTOOLBAREDITOR_COMBOBUTTONTYPE), nIndex)
   ' If this is a Separator button then disable the Images and Tooltip
   If nIndex = 1 Then bEnabled = false   ' Separator index is 1

   AfxSetWindowText( GetDlgItem(HWND_FRMTOOLBAREDITOR, IDC_FRMTOOLBAREDITOR_TXTTOOLTIP), gToolBarItems(nCurSel).wszToolTip)
   AfxSetWindowText( GetDlgItem(HWND_FRMTOOLBAREDITOR, IDC_FRMTOOLBAREDITOR_TXTNORMALIMAGE), gToolBarItems(nCurSel).pPropNormalImage.wszPropValue)
   AfxSetWindowText( GetDlgItem(HWND_FRMTOOLBAREDITOR, IDC_FRMTOOLBAREDITOR_TXTHOTIMAGE), gToolBarItems(nCurSel).pPropHotImage.wszPropValue)
   AfxSetWindowText( GetDlgItem(HWND_FRMTOOLBAREDITOR, IDC_FRMTOOLBAREDITOR_TXTDISABLEDIMAGE), gToolBarItems(nCurSel).pPropDisabledImage.wszPropValue)

   EnableWindow( GetDlgItem(HWND_FRMTOOLBAREDITOR, IDC_FRMTOOLBAREDITOR_TXTTOOLTIP), bEnabled )
   EnableWindow( GetDlgItem(HWND_FRMTOOLBAREDITOR, IDC_FRMTOOLBAREDITOR_TXTNORMALIMAGE), bEnabled )
   EnableWindow( GetDlgItem(HWND_FRMTOOLBAREDITOR, IDC_FRMTOOLBAREDITOR_TXTHOTIMAGE), bEnabled )
   EnableWindow( GetDlgItem(HWND_FRMTOOLBAREDITOR, IDC_FRMTOOLBAREDITOR_TXTDISABLEDIMAGE), bEnabled )

   EnableWindow( GetDlgItem(HWND_FRMTOOLBAREDITOR, IDC_FRMTOOLBAREDITOR_CMDNORMALIMAGE), bEnabled )
   EnableWindow( GetDlgItem(HWND_FRMTOOLBAREDITOR, IDC_FRMTOOLBAREDITOR_CMDHOTIMAGE), bEnabled )
   EnableWindow( GetDlgItem(HWND_FRMTOOLBAREDITOR, IDC_FRMTOOLBAREDITOR_CMDDISABLEDIMAGE), bEnabled )

   SetFocus( GetDlgItem(HWND_FRMTOOLBAREDITOR, IDC_FRMTOOLBAREDITOR_COMBOBUTTONTYPE) )

   function = 0
End Function


' ========================================================================================
' Display the entries in the ToolBarEditor Listbox
' ========================================================================================
private Function frmToolBarEditor_DisplayListBox( ByVal nIndex As Long ) As Long

   ' Display the button items in the listbox
   dim hList as hwnd = GetDlgItem( HWND_FRMTOOLBAREDITOR, IDC_FRMTOOLBAREDITOR_LSTBUTTONS )

   ListBox_ResetContent( hList )
   for i as long = 0 to ubound(gToolBarItems)
      gToolBarItems(i).wszName = "ToolBar Button " & i
      ListBox_AddString( hList, gToolBarItems(i).wszName.sptr )
   next
   ListBox_SetCurSel( hList, nIndex )

   function = 0
End Function


' ========================================================================================
' Swap two entries in the ToolBarEditor Listbox
' ========================================================================================
private function frmToolBarEditor_SwapListBoxItems( byval Item1 as long, _
                                                    Byval Item2 as long _
                                                    ) as Long
   ' Do not swap the Button Name
   dim as CWSTR wszButtonName1 = gToolBarItems(Item1).wszName
   dim as CWSTR wszButtonName2 = gToolBarItems(Item2).wszName
   
   ' Swap the array values
   swap gToolBarItems(Item1), gToolBarItems(Item2)
    
   gToolBarItems(Item1).wszName = wszButtonName1
   gToolBarItems(Item2).wszName = wszButtonName2 

   frmToolBarEditor_DisplayListBox( Item2 )
   
   function = 0
end function


' ========================================================================================
' Process WM_CREATE message for window/dialog: frmToolBarEditor
' ========================================================================================
private Function frmToolBarEditor_OnCreate( ByVal HWnd As HWnd, _
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
' Process WM_COMMAND message for window/dialog: frmToolBarEditor
' ========================================================================================
private Function frmToolBarEditor_OnCommand( ByVal HWnd As HWnd, _
                                             ByVal id As Long, _
                                             ByVal hwndCtl As HWnd, _
                                             ByVal codeNotify As UINT _
                                             ) As LRESULT

   dim as HWND hListBox = GetDlgItem( HWND_FRMTOOLBAREDITOR, IDC_FRMTOOLBAREDITOR_LSTBUTTONS)
   dim as long nCurSel = ListBox_GetCurSel(hListBox)
    
   Select Case id
          
      case IDC_FRMTOOLBAREDITOR_LSTBUTTONS
         If codeNotify = LBN_SELCHANGE Then
            frmToolBarEditor_DisplayToolBarItem
         end if


      Case IDC_FRMTOOLBAREDITOR_CHKDISPLAYONFORM
         If codeNotify = BN_CLICKED Then
            ' Don't need to store the value. We will check it when the 
            ' ToolBar Editor is closed.
         End If

      
      case IDC_FRMTOOLBAREDITOR_COMBOSIZE
         if codeNotify = CBN_SELCHANGE THEN
            ' Don't need to store the value. We will check it when the 
            ' ToolBar Editor is closed.
         END IF


      case IDC_FRMTOOLBAREDITOR_COMBOBUTTONTYPE
         if codeNotify = CBN_SELCHANGE THEN
            if nCurSel > -1 THEN 
               gToolBarItems(nCurSel).wszButtonType = AfxGetWindowText(hwndCtl) 
               frmToolBarEditor_DisplayToolBarItem
            end if
         END IF


      CASE IDC_FRMTOOLBAREDITOR_TXTTOOLTIP
         if codeNotify = EN_CHANGE THEN
            if nCurSel > -1 THEN gToolBarItems(nCurSel).wszTooltip = AfxGetWindowText(hwndCtl)
         end if
           
      
      case IDC_FRMTOOLBAREDITOR_CMDBUTTONUP
         if codeNotify = BN_CLICKED THEN
            if nCurSel > 0 THEN
               frmToolBarEditor_SwapListboxItems(nCurSel, nCurSel - 1)
            END IF
         end if

      
      case IDC_FRMTOOLBAREDITOR_CMDBUTTONDOWN
         if codeNotify = BN_CLICKED THEN
            if nCurSel < ListBox_GetCount(hListBox)-1 THEN
               frmToolBarEditor_SwapListboxItems(nCurSel, nCurSel + 1)
            END IF
         end if

      
      case IDC_FRMTOOLBAREDITOR_CMDBUTTONADD
         if codeNotify = BN_CLICKED THEN
            if ubound(gToolBarItems) = -1 THEN
               redim gToolBarItems(0)
               nCurSel = 0
            Else
               Redim Preserve gToolBarItems(Ubound(gToolBarItems)+1)
               nCurSel = Ubound(gToolBarItems)
            END IF
            dim newToolBarItem as clsToolBarItem
            gToolBarItems(nCurSel) = newToolBarItem
            frmToolBarEditor_DisplayListBox( nCurSel )
            frmToolBarEditor_DisplayToolBarItem
         end if


      case IDC_FRMTOOLBAREDITOR_CMDBUTTONINSERT
         if codeNotify = BN_CLICKED THEN
            if ubound(gToolBarItems) = -1 THEN
               redim gToolBarItems(0)
               nCurSel = 0
            Else
               Redim Preserve gToolBarItems(Ubound(gToolBarItems)+1)
               if nCurSel = -1 THEN nCurSel = 0
               for i as long = ubound(gToolBarItems) to nCurSel + 1 step -1
                  gToolBarItems(i) = gToolBarItems(i-1)
               NEXT
            END IF
            dim newToolBarItem as clsToolBarItem
            gToolBarItems(nCurSel) = newToolBarItem
            frmToolBarEditor_DisplayListBox( nCurSel )
            frmToolBarEditor_DisplayToolBarItem
         end if
          

      case IDC_FRMTOOLBAREDITOR_CMDBUTTONDELETE
         If codeNotify = BN_CLICKED Then
            if nCurSel > -1  THEN
               if ubound(gToolBarItems) = 0 THEN
                  erase gToolBarItems
                  nCurSel = -1
               else
                  ' remove the item from the internal array
                  for i as long = nCurSel to ubound(gToolBarItems) - 1
                     gToolBarItems(i) = gToolBarItems(i+1)
                  NEXT
                  redim preserve gToolBarItems(ubound(gToolBarItems)-1)
               END IF
               nCurSel = Min(nCurSel, ubound(gToolBarItems))
               frmToolBarEditor_DisplayListBox( nCurSel )
               frmToolBarEditor_DisplayToolBarItem
            end if
         end if


      case IDC_FRMTOOLBAREDITOR_CMDNORMALIMAGE
         If codeNotify = BN_CLICKED Then
            if nCurSel > -1 then
               frmImageManager_Show( HWND, @gToolBarItems(nCurSel).pPropNormalImage )
               if gToolBarItems(nCurSel).pPropHotImage.wszPropValue = "" then
                  gToolBarItems(nCurSel).pPropHotImage.wszPropValue = _
                     gToolBarItems(nCurSel).pPropNormalImage.wszPropValue
               end if      
               if gToolBarItems(nCurSel).pPropDisabledImage.wszPropValue = "" then
                  gToolBarItems(nCurSel).pPropDisabledImage.wszPropValue = _
                     gToolBarItems(nCurSel).pPropNormalImage.wszPropValue
               end if      
               frmToolBarEditor_DisplayToolBarItem
            end if
         end if
         
         
      case IDC_FRMTOOLBAREDITOR_CMDHOTIMAGE
         If codeNotify = BN_CLICKED Then
            if nCurSel > -1 then
               frmImageManager_Show( HWND, @gToolBarItems(nCurSel).pPropHotImage )
               if gToolBarItems(nCurSel).pPropNormalImage.wszPropValue = "" then
                  gToolBarItems(nCurSel).pPropNormalImage.wszPropValue = _
                     gToolBarItems(nCurSel).pPropHotImage.wszPropValue
               end if      
               if gToolBarItems(nCurSel).pPropDisabledImage.wszPropValue = "" then
                  gToolBarItems(nCurSel).pPropDisabledImage.wszPropValue = _
                     gToolBarItems(nCurSel).pPropHotImage.wszPropValue
               end if      
               frmToolBarEditor_DisplayToolBarItem
            end if
         end if


      case IDC_FRMTOOLBAREDITOR_CMDDISABLEDIMAGE
         If codeNotify = BN_CLICKED Then
            if nCurSel > -1 then
               frmImageManager_Show( HWND, @gToolBarItems(nCurSel).pPropDisabledImage )
               if gToolBarItems(nCurSel).pPropNormalImage.wszPropValue = "" then
                  gToolBarItems(nCurSel).pPropNormalImage.wszPropValue = _
                     gToolBarItems(nCurSel).pPropDisabledImage.wszPropValue
               end if      
               if gToolBarItems(nCurSel).pPropHotImage.wszPropValue = "" then
                  gToolBarItems(nCurSel).pPropHotImage.wszPropValue = _
                     gToolBarItems(nCurSel).pPropDisabledImage.wszPropValue
               end if      
               frmToolBarEditor_DisplayToolBarItem
            end if
         end if


      Case IDC_FRMTOOLBAREDITOR_CMDOK    
         If codeNotify = BN_CLICKED Then
            ' Copy the temporary array back to the form's button items 
            Dim pDoc As clsDocument Ptr = gTTabCtl.GetActiveDocumentPtr()
            If pDoc Then 
               dim as long numItems = ubound(gToolBarItems)
               if numItems > -1 then
                  redim pDoc->ToolBarItems(numItems)
                  for i as long = lbound(gToolBarItems) to ubound(gToolBarItems)
                     pDoc->ToolBarItems(i) = gToolBarItems(i) 
                  next
               else   
                  erase pDoc->ToolBarItems
               end if
               ' Save the value indicating that we want the toolbar to display in
               ' the generated code for the form.
               pDoc->GenerateToolBar = _
                  iif(Button_GetCheck(GetDlgItem(HWND, IDC_FRMTOOLBAREDITOR_CHKDISPLAYONFORM)) = BST_CHECKED, true, false)
               
               dim as long nIndex = ComboBox_GetCurSel( GetDlgItem(HWND, IDC_FRMTOOLBAREDITOR_COMBOSIZE) )
               pDoc->wszToolBarSize = AfxGetComboBoxText( GetDlgItem(HWND, IDC_FRMTOOLBAREDITOR_COMBOSIZE), nIndex )
               pDoc->UserModified = true
            end if
            SendMessage(HWnd, WM_CLOSE, 0, 0)
            Exit Function
         end if


      Case IDC_FRMTOOLBAREDITOR_CMDCANCEL
         If codeNotify = BN_CLICKED Then
            SendMessage(HWnd, WM_CLOSE, 0, 0)
            Exit Function
         End If
    
   End Select

   Function = 0
End Function


' ========================================================================================
' Process WM_CLOSE message for window/dialog: frmToolBarEditor
' ========================================================================================
private Function frmToolBarEditor_OnClose( byval HWnd As HWnd ) As LRESULT
   ' Reset the global ToolBarItems array
   erase gToolBarItems
   ' Enables parent window keeping parent's zorder
   EnableAllModeless()
   DestroyWindow( HWnd )
   Function = 0
End Function


' ========================================================================================
' Process WM_DESTROY message for window/dialog: frmToolBarEditor
' ========================================================================================
private Function frmToolBarEditor_OnDestroy( byval HWnd As HWnd ) As LRESULT
   PostQuitMessage(0)
   Function = 0
End Function


' ========================================================================================
' frmToolBarEditor Window procedure
' ========================================================================================
private Function frmToolBarEditor_WndProc( ByVal HWnd   As HWnd, _
                                           ByVal uMsg   As UINT, _
                                           ByVal wParam As WPARAM, _
                                           ByVal lParam As LPARAM _
                                           ) As LRESULT
   Select Case uMsg
      HANDLE_MSG (HWnd, WM_CREATE,   frmToolBarEditor_OnCreate)
      HANDLE_MSG (HWnd, WM_CLOSE,    frmToolBarEditor_OnClose)
      HANDLE_MSG (HWnd, WM_DESTROY,  frmToolBarEditor_OnDestroy)
      HANDLE_MSG (HWnd, WM_COMMAND,  frmToolBarEditor_OnCommand)
   End Select

   Function = DefWindowProc(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' frmToolBarEditor_Show
' ========================================================================================
public Function frmToolBarEditor_Show( ByVal hWndParent As HWnd ) as LRESULT

   DIM hBitmap AS HBITMAP
   dim hCtrl as HWnd
   dim wszImage as wstring * 100
   

   '  Create the main window and child controls
   Dim pWindow as CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowOwnerPtr(hwndParent)->DPI
   
   Dim As HWnd hForm = _
   pWindow->Create( hWndParent, L(313,"ToolBar Editor"), _
        @frmToolBarEditor_WndProc, 0, 0, 508, 380, _
        WS_POPUP Or WS_CAPTION Or WS_SYSMENU Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, _
        WS_EX_DLGMODALFRAME Or WS_EX_CONTROLPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->SetClientSize(508, 380)
   pWindow->Center
        
   HWND_FRMTOOLBAREDITOR = hForm

   pWindow->AddControl("LISTBOX", , IDC_FRMTOOLBAREDITOR_LSTBUTTONS, "", 21, 14, 465, 84, _
        WS_CHILD Or WS_VISIBLE Or WS_VSCROLL Or WS_TABSTOP Or LBS_NOTIFY Or LBS_NOINTEGRALHEIGHT, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)

   hCtrl = _
   pWindow->AddControl("BITMAPBUTTON", , IDC_FRMTOOLBAREDITOR_CMDBUTTONUP, "", 21, 102, 24, 24, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or BS_BITMAP Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER Or LR_DEFAULTCOLOR Or LR_LOADMAP3DCOLORS Or LR_LOADTRANSPARENT Or LR_SHARED, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   wszImage = iif(pWindow->DPI > 96, "IMAGE_ARROWUP", "IMAGE_ARROWUP16")
         hBitmap = AfxGdipImageFromRes(pWindow->InstanceHandle, wszImage, 0, false, IMAGE_BITMAP, 0)
         SendMessage(hCtrl, BM_SETIMAGE, IMAGE_BITMAP, cast(LPARAM, hBitmap))
         IF hBitmap THEN DeleteObject(hBitmap)

   hCtrl = _
   pWindow->AddControl("BITMAPBUTTON", , IDC_FRMTOOLBAREDITOR_CMDBUTTONDOWN, "", 48, 102, 24, 24, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or BS_BITMAP Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER Or LR_DEFAULTCOLOR Or LR_LOADMAP3DCOLORS Or LR_LOADTRANSPARENT Or LR_SHARED, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   wszImage = iif(pWindow->DPI > 96, "IMAGE_ARROWDOWN", "IMAGE_ARROWDOWN16")
         hBitmap = AfxGdipImageFromRes(pWindow->InstanceHandle, wszImage, 0, false, IMAGE_BITMAP, 0)
         SendMessage(hCtrl, BM_SETIMAGE, IMAGE_BITMAP, cast(LPARAM, hBitmap))
         IF hBitmap THEN DeleteObject(hBitmap)

   pWindow->AddControl("BUTTON", , IDC_FRMTOOLBAREDITOR_CMDBUTTONADD, L(380, "Add"), 77, 102, 70, 24, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_FLAT Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("BUTTON", , IDC_FRMTOOLBAREDITOR_CMDBUTTONINSERT, L(281, "Insert"), 152, 102, 70, 24, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_FLAT Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("BUTTON", , IDC_FRMTOOLBAREDITOR_CMDBUTTONDELETE, L(282, "Delete"), 227, 102, 70, 24, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_FLAT Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)

   pWindow->AddControl("LABEL", , IDC_FRMTOOLBAREDITOR_LABEL1, L(378,"Images") & ":", 305, 107, 100, 16, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_RIGHT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("COMBOBOX", , IDC_FRMTOOLBAREDITOR_COMBOSIZE, "", 411, 102, 75, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_VSCROLL Or WS_TABSTOP Or CBS_DROPDOWNLIST, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)


   pWindow->AddControl("LABEL", , IDC_FRMTOOLBAREDITOR_LABEL2, L(383,"Button Type") & ":", 21, 145, 93, 16, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("COMBOBOX", , IDC_FRMTOOLBAREDITOR_COMBOBUTTONTYPE, "", 134, 142, 352, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_VSCROLL Or WS_TABSTOP Or CBS_DROPDOWNLIST, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)

   pWindow->AddControl("LABEL", , IDC_FRMTOOLBAREDITOR_LABEL3, L(235,"Tooltip Text") & ":", 21, 181, 93, 16, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("TEXTBOX", , IDC_FRMTOOLBAREDITOR_TXTTOOLTIP, "", 134, 177, 352, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)

   
   pWindow->AddControl("LABEL", , IDC_FRMTOOLBAREDITOR_LABEL4, L(384,"Normal Image") & ":", 21, 209, 93, 21, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("TEXTBOX", , IDC_FRMTOOLBAREDITOR_TXTNORMALIMAGE, "", 134, 205, 320, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   hCtrl = _
   pWindow->AddControl("BITMAPBUTTON", , IDC_FRMTOOLBAREDITOR_CMDNORMALIMAGE, "", 462, 204, 22, 22, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or BS_BITMAP Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER Or LR_DEFAULTCOLOR Or LR_LOADMAP3DCOLORS Or LR_LOADTRANSPARENT Or LR_SHARED, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   wszImage = iif(pWindow->DPI > 96, "IMAGE_ARROWDOWN", "IMAGE_ARROWDOWN16")
         hBitmap = AfxGdipImageFromRes(pWindow->InstanceHandle, wszImage, 0, false, IMAGE_BITMAP, 0)
         SendMessage(hCtrl, BM_SETIMAGE, IMAGE_BITMAP, cast(LPARAM, hBitmap))
         IF hBitmap THEN DeleteObject(hBitmap)


   pWindow->AddControl("LABEL", , IDC_FRMTOOLBAREDITOR_LABEL5, L(385,"Hot Image") & ":", 21, 237, 93, 21, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("TEXTBOX", , IDC_FRMTOOLBAREDITOR_TXTHOTIMAGE, "", 134, 233, 320, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   hCtrl = _
   pWindow->AddControl("BITMAPBUTTON", , IDC_FRMTOOLBAREDITOR_CMDHOTIMAGE, "", 462, 232, 22, 22, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or BS_BITMAP Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER Or LR_DEFAULTCOLOR Or LR_LOADMAP3DCOLORS Or LR_LOADTRANSPARENT Or LR_SHARED, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   wszImage = iif(pWindow->DPI > 96, "IMAGE_ARROWDOWN", "IMAGE_ARROWDOWN16")
         hBitmap = AfxGdipImageFromRes(pWindow->InstanceHandle, wszImage, 0, false, IMAGE_BITMAP, 0)
         SendMessage(hCtrl, BM_SETIMAGE, IMAGE_BITMAP, cast(LPARAM, hBitmap))
         IF hBitmap THEN DeleteObject(hBitmap)


   pWindow->AddControl("LABEL", , IDC_FRMTOOLBAREDITOR_LABEL6, L(386,"Disabled Image") & ":", 21, 265, 93, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("TEXTBOX", , IDC_FRMTOOLBAREDITOR_TXTDISABLEDIMAGE, "", 134, 261, 320, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   hCtrl = _
   pWindow->AddControl("BITMAPBUTTON", , IDC_FRMTOOLBAREDITOR_CMDDISABLEDIMAGE, "", 462, 260, 22, 22, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or BS_BITMAP Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER Or LR_DEFAULTCOLOR Or LR_LOADMAP3DCOLORS Or LR_LOADTRANSPARENT Or LR_SHARED, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   wszImage = iif(pWindow->DPI > 96, "IMAGE_ARROWDOWN", "IMAGE_ARROWDOWN16")
         hBitmap = AfxGdipImageFromRes(pWindow->InstanceHandle, wszImage, 0, false, IMAGE_BITMAP, 0)
         SendMessage(hCtrl, BM_SETIMAGE, IMAGE_BITMAP, cast(LPARAM, hBitmap))
         IF hBitmap THEN DeleteObject(hBitmap)


   pWindow->AddControl("CHECKBOX", , IDC_FRMTOOLBAREDITOR_CHKDISPLAYONFORM, L(237,"Display on Form"), 21, 345, 175, 16, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)

   pWindow->AddControl("BUTTON", , IDC_FRMTOOLBAREDITOR_CMDOK, L(0,"OK"), 331, 333, 74, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("BUTTON", , IDC_FRMTOOLBAREDITOR_CMDCANCEL, L(1,"Cancel"), 412, 333, 74, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)


   ' Copy the form's toolbar items to the temporary array for editing
   Dim pDoc As clsDocument Ptr = gTTabCtl.GetActiveDocumentPtr()
   If pDoc Then 
      hCtrl = GetDlgItem( hForm, IDC_FRMTOOLBAREDITOR_CHKDISPLAYONFORM ) 
      Button_SetCheck( hCtrl, iif(pDoc->GenerateToolBar, BST_CHECKED, BST_UNCHECKED))
      dim as long numItems = ubound(pDoc->ToolBarItems)
      if numItems > -1 then
         redim gToolBarItems(numItems)
         for i as long = 0 to numItems
            gToolBarItems(i) = pDoc->ToolBarItems(i)
         next
      end if
   end if
    

   ' Load the comboboxes
   hCtrl = GetDlgItem( hForm, IDC_FRMTOOLBAREDITOR_COMBOSIZE )
   ComboBox_AddString( hCtrl, @wstr("SIZE_16") )
   ComboBox_AddString( hCtrl, @wstr("SIZE_24") )
   ComboBox_AddString( hCtrl, @wstr("SIZE_32") )
   ComboBox_AddString( hCtrl, @wstr("SIZE_48") )
   dim as long nIndex = ComboBox_FindStringExact( hCtrl, -1, pDoc->wszToolBarSize.sptr ) 
   ComboBox_SetCurSel( hCtrl, nIndex)
   
   hCtrl = GetDlgItem( hForm, IDC_FRMTOOLBAREDITOR_COMBOBUTTONTYPE )
   ComboBox_AddString( hCtrl, @wstr("ToolBarButton.Button") )
   ComboBox_AddString( hCtrl, @wstr("ToolBarButton.Separator") )
   ComboBox_AddString( hCtrl, @wstr("ToolBarButton.DropDown") )
   ComboBox_AddString( hCtrl, @wstr("ToolBarButton.WholeDropDown") )

   frmToolBarEditor_DisplayListBox( 0 )
   frmToolBarEditor_DisplayToolBarItem
   SetFocus( GetDlgItem( HWND_FRMTOOLBAREDITOR, IDC_FRMTOOLBAREDITOR_LSTBUTTONS ) )

   ' Process Windows messages
   Function = pWindow->DoEvents( SW_SHOW )

   ' Delete the frmToolBarEditor CWindow class manually allocated memory 
   Delete pWindow

End Function





