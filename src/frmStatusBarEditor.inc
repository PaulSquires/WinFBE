'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2025 Paul Squires, PlanetSquires Software
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
   dim as long nPanelWidth, nPanelMinWidth, nTotalPanelWidths
   dim as long nSpringIndex = -1
   Dim As Long nTemp(ub)   ' panel widths (-1 if Spring)
   Dim As Long nWidths(ub) ' final panel cummulative widths array 

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
            dim as WString * MAX_PATH wszBuffer = " " & pDoc->PanelItems(i).wszText & " "
            dim as long nTextWidth = GetTextWidthPixels( pDoc->hWndStatusBar, wszBuffer )  
            ' Also need to calculate the width of any defined panel image.
            ' ImageName is held in pProp.wszPropValue
            dim as long nImageWidth 
            if len(pDoc->PanelItems(i).pProp.wszPropValue) then
               nImageWidth = AfxScaleX(20)  ' 16 + 4 padding
            end if
            nPanelWidth = nTextWidth + nImageWidth 
            if nPanelWidth < nPanelMinWidth then nPanelWidth = nPanelMinWidth
         case "STATUSBARPANELAUTOSIZE.SPRING"
            nSpringIndex = i
            nPanelWidth = 0
      end select
      nTemp(i) = nPanelWidth
      nTotalPanelWidths = nTotalPanelWidths + nPanelWidth
   NEXT
   
   ' Handle the Panel that may have been designated as AutoSize = Spring
   if nSpringIndex <> -1 then
      nTemp(nSpringIndex) = MAX(0, rc.right - nTotalPanelWidths)
   end if
   
   ' Build the cumulative Panel Widths array that gets sent to the StatusBar
   for i as long = lb to ub
      nWidths(i) = iif(i, nWidths(i-1), 0) + nTemp(i)
   next
   StatusBar_SetParts( pDoc->hWndStatusBar, (ub-lb)+1, @nWidths(0))

   ' Set the text/icons for the panels. Need to do this after the widths are set.
   ' Basically, a call to SetText with SBT_OWNERDRAW will fire WM_DRAWITEM in the
   ' parent form. That is where the images and icons are actually drawn.
   for i as long = lb to ub
      StatusBar_SetText( pDoc->hWndStatusBar, i, "", SBT_OWNERDRAW ) 
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

   ' BorderStyle is deprecated as of v2.0.4 as it has no effect
   ' in WinFBE programs where Windows Themes are enabled.
   'hCtrl = GetDlgItem(HWND_FRMSTATUSBAREDITOR, IDC_FRMSTATUSBAREDITOR_COMBOBORDERSTYLE)
   'if nCurSel <> -1 then wszText = gPanelItems(nCurSel).wszBorderStyle else wszText = ""
   'nIndex = ComboBox_FindStringExact( hCtrl, -1, @wszText ) 
   'ComboBox_SetCurSel( hCtrl, nIndex)
   'EnableWindow( hCtrl, fEnabled)

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

   hCtrl = GetDlgItem(HWND_FRMSTATUSBAREDITOR, IDC_FRMSTATUSBAREDITOR_COMBOBACKCOLOR)
   if nCurSel <> -1 then wszText = gPanelItems(nCurSel).wszBackColor else wszText = ""
   AfxSetWindowText( hCtrl, wszText )
   EnableWindow( hCtrl, fEnabled)
   AfxRedrawWindow(hCtrl)

   hCtrl = GetDlgItem(HWND_FRMSTATUSBAREDITOR, IDC_FRMSTATUSBAREDITOR_COMBOFORECOLOR)
   if nCurSel <> -1 then wszText = gPanelItems(nCurSel).wszForeColor else wszText = ""
   AfxSetWindowText( hCtrl, wszText )
   EnableWindow( hCtrl, fEnabled)
   AfxRedrawWindow(hCtrl)

   hCtrl = GetDlgItem(HWND_FRMSTATUSBAREDITOR, IDC_FRMSTATUSBAREDITOR_COMBOBACKCOLORHOT)
   if nCurSel <> -1 then wszText = gPanelItems(nCurSel).wszBackColorHot else wszText = ""
   AfxSetWindowText( hCtrl, wszText )
   EnableWindow( hCtrl, fEnabled)
   AfxRedrawWindow(hCtrl)
   
   hCtrl = GetDlgItem(HWND_FRMSTATUSBAREDITOR, IDC_FRMSTATUSBAREDITOR_COMBOFORECOLORHOT)
   if nCurSel <> -1 then wszText = gPanelItems(nCurSel).wszForeColorHot else wszText = ""
   AfxSetWindowText( hCtrl, wszText )
   EnableWindow( hCtrl, fEnabled)
   AfxRedrawWindow(hCtrl)

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

      
      ' BorderStyle is deprecated as of v2.0.4 as it has no effect
      ' in WinFBE programs where Windows Themes are enabled.
      'case IDC_FRMSTATUSBAREDITOR_COMBOBORDERSTYLE
      '   if codeNotify = CBN_SELCHANGE THEN
      '      if nCurSel > -1 THEN gPanelItems(nCurSel).wszBorderStyle = AfxGetWindowText(hwndCtl) 
      '   END IF

      
      case IDC_FRMSTATUSBAREDITOR_COMBOAUTOSIZE
         if codeNotify = CBN_SELCHANGE THEN
            if nCurSel > -1 THEN
               dim as CWSTR wszAutoSize = AfxGetWindowText(hwndCtl)
               ' Only one panel can have AutoSize set to Spring. Reset all panels
               ' that have Spring set and then set this panel to have the Spring.
               if wszAutoSize = "StatusBarPanelAutoSize.Spring" then
                  for i as long = lbound(gPanelItems) to ubound(gPanelItems)
                     if gPanelItems(i).wszAutoSize = "StatusBarPanelAutoSize.Spring" then
                        gPanelItems(i).wszAutoSize = "StatusBarPanelAutoSize.None"
                     end if   
                  next
               end if
               gPanelItems(nCurSel).wszAutoSize = wszAutoSize
            end if
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
               else
                  erase pDoc->PanelItems
               end if
               ' Save the value indicating that we want the statusbar to display in
               ' the generated code for the form.
               pDoc->GenerateStatusBar = _
                  iif(Button_GetCheck(GetDlgItem(HWND, IDC_FRMSTATUSBAREDITOR_CHKDISPLAYONFORM)) = BST_CHECKED, true, false)
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
' Process WM_MEASUREITEM message for window/dialog: frmStatusBarEditor
' ========================================================================================
private Function frmStatusBarEditor_OnMeasureItem( ByVal HWnd As HWnd, _
                                                   ByVal lpmis As MEASUREITEMSTRUCT Ptr _
                                                   ) As Long
   ' Set the height of the List box items. 
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(HWnd)
   lpmis->itemHeight = pWindow->ScaleY(20)
   Function = 0
End Function


' ========================================================================================
' Process WM_DRAWITEM message for window/dialog: frmStatusBarEditor
' ========================================================================================
private Function frmStatusBarEditor_OnDrawItem( ByVal HWnd As HWnd, _
                                                ByVal lpdis As Const DRAWITEMSTRUCT Ptr _
                                                ) As Long
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(HWnd)
   if pWindow = 0 THEN exit function

   dim as hwnd hList1 = GetDlgItem( HWND, IDC_FRMSTATUSBAREDITOR_LSTPANELS)
   dim as long nCurSel = ListBox_GetCurSel(hList1)

   Dim As HBRUSH hBrush, hBrushOld
   Dim As RECT rc, rc2 
   dim as CWSTR wszPropValue
   Dim wszText As WString * MAX_PATH
   
      
   Select Case lpdis->itemAction
      Case ODA_DRAWENTIRE, ODA_SELECT
         
         SaveDC(lpdis->hDC)
         
         ' CLEAR BACKGROUND
         If (lpdis->itemState And ODS_SELECTED) Then     
            SetBkColor(lpdis->hDC, GetSysColor(COLOR_HIGHLIGHT))   
            SetTextColor(lpdis->hDC, GetSysColor(COLOR_HIGHLIGHTTEXT))
            hBrush = GetSysColorBrush(COLOR_HIGHLIGHT) 
         else
            SetBkColor(lpdis->hDC, GetSysColor(COLOR_WINDOW))   
            SetTextColor(lpdis->hDC, GetSysColor(COLOR_WINDOWTEXT))
            hBrush = GetSysColorBrush(COLOR_WINDOW) 
         end if
         SelectObject(lpdis->hDC, hBrush)      
         FillRect(lpdis->hDC, @rc, hBrush)  

         ' Exit if no panel is selected
         if nCurSel = -1 then exit function
      
         ' COMBOBOX FOR BACK/FORE COLORS
         select case lpdis->CtlID 
            case IDC_FRMSTATUSBAREDITOR_COMBOBACKCOLOR
               wszPropValue = gPanelItems(nCurSel).wszBackColor
            case IDC_FRMSTATUSBAREDITOR_COMBOBACKCOLORHOT
               wszPropValue = gPanelItems(nCurSel).wszBackColorHot
            case IDC_FRMSTATUSBAREDITOR_COMBOFORECOLOR
               wszPropValue = gPanelItems(nCurSel).wszForeColor
            case IDC_FRMSTATUSBAREDITOR_COMBOFORECOLORHOT
               wszPropValue = gPanelItems(nCurSel).wszForeColorHot
         end select

         rc = lpdis->rcItem
         
         dim as CWSTR wszList, wszValue
         wszList = AfxStrParse(wszPropValue, 1, "|")
         wszValue = AfxStrParse(wszPropValue, 2, "|")
         wszText = wszValue
         if wszList = "CUSTOM" then wszText = "Custom Color"
            
         rc.left   = pWindow->ScaleX(6)
         rc.top    = rc.top + pWindow->ScaleY(2)
         rc.right  = rc.left + pWindow->ScaleX(16)
         rc.bottom = rc.bottom - pWindow->ScaleY(2)
            
         dim as COLORREF rgbClr = GetRGBColorFromProperty(wszPropValue)
         hBrush = CreateSolidBrush(rgbClr)

         ' DRAW COLOR RECT
         hBrushOld = SelectObject(lpdis->hDC, hBrush)
         Rectangle( lpdis->hDC, rc.Left, rc.Top, rc.Right, rc.Bottom)
         SelectObject( lpdis->hDC, hBrushOld)
         rc.left = rc.right + pWindow->ScaleX(4)
         rc.right = lpdis->rcItem.right
         DeleteObject(hBrush)

         SelectObject(lpdis->hDC, AfxGetWindowFont(hwnd) )
         DrawText( lpdis->hDC, _
                   wszText, _
                   -1, Cast(lpRect, @rc), _
                   DT_LEFT Or DT_SINGLELINE Or DT_VCENTER or DT_NOPREFIX)
   
   end select
   
   RestoreDC(lpdis->hDC, -1)

   return True 
End Function


' ========================================================================================
' Processes messages for the subclassed color comboboxes.
' ========================================================================================
private Function frmStatusBarEditor_ComboBox_SubclassProc ( _
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


      case WM_SYSKEYDOWN
         'This should suppress the Alt+Down arrow that invokes the combolist dropdown.
         return true
         
         
      case WM_KEYDOWN
         ' Only allow the TAB key to pass through. 
         Select Case wParam
            Case VK_TAB
               ' move input fcous to the next/prev control in the tab order
               dim as Boolean bBackward = iif( GetKeyState(VK_SHIFT) < 0, true, false )
               Dim As HWnd hCtrl = GetNextDlgTabItem( GetParent(hwnd), HWND, bBackward)
               SetFocus(hCtrl)

            case else
               return true   
         end select
         
         
      case WM_LBUTTONDOWN, WM_LBUTTONDBLCLK
         ' Catch the mouse down click in order to prevent the subsequent CBN_DROPDOWN and CBN_SELCHANGE
         ' messages. We don't want the dropdown to show. We will show the popup color selector
         ' instead.
         SetFocus(HWND)
         
         dim as hwnd hList1 = GetDlgItem( HWND_FRMSTATUSBAREDITOR, IDC_FRMSTATUSBAREDITOR_LSTPANELS)
         dim as long nCurSel = ListBox_GetCurSel(hList1)
         if nCurSel = -1 then return true
         
         dim as long nID = GetDlgCtrlID( hwnd )
         select case nID
            case IDC_FRMSTATUSBAREDITOR_COMBOBACKCOLOR
               gPanelItems(nCurSel).pPropColor.wszPropValue = gPanelItems(nCurSel).wszBackColor
            case IDC_FRMSTATUSBAREDITOR_COMBOBACKCOLORHOT
               gPanelItems(nCurSel).pPropColor.wszPropValue = gPanelItems(nCurSel).wszBackColorHot
            case IDC_FRMSTATUSBAREDITOR_COMBOFORECOLOR
               gPanelItems(nCurSel).pPropColor.wszPropValue = gPanelItems(nCurSel).wszForeColor
            case IDC_FRMSTATUSBAREDITOR_COMBOFORECOLORHOT
               gPanelItems(nCurSel).pPropColor.wszPropValue = gPanelItems(nCurSel).wszForeColorHot
         end select
         gPanelItems(nCurSel).idColorCombo = nID
           
         dim as RECT rc
         GetWindowRect( hwnd, @rc )
         ' initialize the color popup if not already done so
         frmVDColors_Show( hwnd, gPanelItems(nCurSel).pPropColor.wszPropValue )  
         dim as long nWidth = AfxGetWindowWidth(HWND_FRMVDCOLORS)
         SetWindowPos( HWND_FRMVDCOLORS, HWND_TOP, _
                       rc.right - nWidth, rc.bottom, _
                       0, 0, SWP_NOSIZE or SWP_SHOWWINDOW)

         return true
         

      Case WM_DESTROY
         ' REQUIRED: Remove control subclassing
         RemoveWindowSubclass HWnd, @frmStatusBarEditor_ComboBox_SubclassProc, uIdSubclass

   End Select

   ' Default processing of Windows messages
   Function = DefSubclassProc(HWnd, uMsg, wParam, lParam)

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
      HANDLE_MSG (HWnd, WM_CREATE,      frmStatusBarEditor_OnCreate)
      HANDLE_MSG (HWnd, WM_CLOSE,       frmStatusBarEditor_OnClose)
      HANDLE_MSG (HWnd, WM_DESTROY,     frmStatusBarEditor_OnDestroy)
      HANDLE_MSG (HWnd, WM_COMMAND,     frmStatusBarEditor_OnCommand)
      HANDLE_MSG (HWnd, WM_MEASUREITEM, frmStatusBarEditor_OnMeasureItem)
      HANDLE_MSG (HWnd, WM_DRAWITEM,    frmStatusBarEditor_OnDrawItem)
   End Select

   Function = DefWindowProc(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' frmStatusBarEditor_Show
' ========================================================================================
public Function frmStatusBarEditor_Show( ByVal hWndParent As HWnd, byval nDefaultPanel as Long ) as LRESULT

   DIM hBitmap AS HBITMAP
   dim hCtrl as HWnd
   dim wszImage as wstring * 100
   

   '  Create the main window and child controls
   Dim pWindow as CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowOwnerPtr(hwndParent)->DPI
   
   Dim As HWnd hForm = _
   pWindow->Create( hWndParent, L(314,"Statusbar Editor"), _
        @frmStatusBarEditor_WndProc, 0, 0, 508, 478, _
        WS_POPUP Or WS_CAPTION Or WS_SYSMENU Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, _
        WS_EX_DLGMODALFRAME Or WS_EX_CONTROLPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->SetClientSize(508, 478)
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

   pWindow->AddControl("LABEL", , IDC_FRMSTATUSBAREDITOR_LABEL6, L(239,"Autosize") & ":", 21, 222, 93, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("COMBOBOX", , IDC_FRMSTATUSBAREDITOR_COMBOAUTOSIZE, "", 114, 219, 207, 21, _
        WS_CHILD Or WS_VISIBLE Or WS_VSCROLL Or WS_TABSTOP Or CBS_DROPDOWNLIST, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)

   pWindow->AddControl("LABEL", , IDC_FRMSTATUSBAREDITOR_LABEL8, L(381,"Width") & ":", 21, 250, 93, 21, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("TEXTBOX", , IDC_FRMSTATUSBAREDITOR_TXTWIDTH, "", 114, 247, 207, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)

   pWindow->AddControl("LABEL", , IDC_FRMSTATUSBAREDITOR_LABEL7, L(245,"Minimum Width") & ":", 21, 277, 93, 21, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("TEXTBOX", , IDC_FRMSTATUSBAREDITOR_TXTMINWIDTH, "", 114, 273, 207, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)

   pWindow->AddControl("LABEL", , IDC_FRMSTATUSBAREDITOR_LABEL9, L(397,"BackColor") & ":", 21, 304, 93, 21, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("COMBOBOX", , IDC_FRMSTATUSBAREDITOR_COMBOBACKCOLOR, "", 114, 300, 207, 21, _
        WS_CHILD or WS_VISIBLE or WS_TABSTOP Or CBS_DROPDOWNLIST or CBS_OWNERDRAWFIXED or CBS_HASSTRINGS, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_RIGHTSCROLLBAR, , _
        Cast(SUBCLASSPROC, @frmStatusBarEditor_ComboBox_SubclassProc), IDC_FRMSTATUSBAREDITOR_COMBOBACKCOLOR, Cast(DWORD_PTR, @pWindow))

   pWindow->AddControl("LABEL", , IDC_FRMSTATUSBAREDITOR_LABEL10, L(399,"ForeColor") & ":", 21, 331, 93, 21, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("COMBOBOX", , IDC_FRMSTATUSBAREDITOR_COMBOFORECOLOR, "", 114, 327, 207, 21, _
        WS_CHILD or WS_VISIBLE or WS_TABSTOP Or CBS_DROPDOWNLIST or CBS_OWNERDRAWFIXED or CBS_HASSTRINGS, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_RIGHTSCROLLBAR, , _
        Cast(SUBCLASSPROC, @frmStatusBarEditor_ComboBox_SubclassProc), IDC_FRMSTATUSBAREDITOR_COMBOFORECOLOR, Cast(DWORD_PTR, @pWindow))

   pWindow->AddControl("LABEL", , IDC_FRMSTATUSBAREDITOR_LABEL11, L(398,"BackColorHot") & ":", 21, 358, 93, 21, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("COMBOBOX", , IDC_FRMSTATUSBAREDITOR_COMBOBACKCOLORHOT, "", 114, 354, 207, 21, _
        WS_CHILD or WS_VISIBLE or WS_TABSTOP Or CBS_DROPDOWNLIST or CBS_OWNERDRAWFIXED or CBS_HASSTRINGS, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_RIGHTSCROLLBAR, , _
        Cast(SUBCLASSPROC, @frmStatusBarEditor_ComboBox_SubclassProc), IDC_FRMSTATUSBAREDITOR_COMBOBACKCOLORHOT, Cast(DWORD_PTR, @pWindow))

   pWindow->AddControl("LABEL", , IDC_FRMSTATUSBAREDITOR_LABEL11, L(400,"ForeColorHot") & ":", 21, 385, 93, 21, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("COMBOBOX", , IDC_FRMSTATUSBAREDITOR_COMBOFORECOLORHOT, "", 114, 381, 207, 21, _
        WS_CHILD or WS_VISIBLE or WS_TABSTOP Or CBS_DROPDOWNLIST or CBS_OWNERDRAWFIXED or CBS_HASSTRINGS, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_RIGHTSCROLLBAR, , _
        Cast(SUBCLASSPROC, @frmStatusBarEditor_ComboBox_SubclassProc), IDC_FRMSTATUSBAREDITOR_COMBOFORECOLORHOT, Cast(DWORD_PTR, @pWindow))


   pWindow->AddControl("BUTTON", , IDC_FRMSTATUSBAREDITOR_CMDPICSELECT, "...", 350, 290, 26, 16, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)

   ' Add an image control                                                         
   DIM pImageCtx AS CImageCtx = CImageCtx(pWindow, IDC_FRMSTATUSBAREDITOR_PICIMAGE, , 384, 219, 60, 60)

   pWindow->AddControl("GROUPBOX", , IDC_FRMSTATUSBAREDITOR_FRAME1, L(246,"Image"), 342, 185, 144, 128, _
        WS_CHILD Or WS_VISIBLE Or BS_TEXT Or BS_LEFT Or BS_NOTIFY Or BS_GROUPBOX, _
        WS_EX_TRANSPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING)


   pWindow->AddControl("CHECKBOX", , IDC_FRMSTATUSBAREDITOR_CHKDISPLAYONFORM, L(237,"Display on Form"), 21, 443, 175, 16, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   
   pWindow->AddControl("BUTTON", , IDC_FRMSTATUSBAREDITOR_CMDOK, L(0,"OK"), 331, 431, 74, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("BUTTON", , IDCANCEL, L(1,"Cancel"), 412, 431, 74, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)


   ' Copy the form's panel items to the temporary array for editing
   Dim pDoc As clsDocument Ptr = gTTabCtl.GetActiveDocumentPtr()
   If pDoc Then 
      hCtrl = GetDlgItem( hForm, IDC_FRMSTATUSBAREDITOR_CHKDISPLAYONFORM ) 
      Button_SetCheck( hCtrl, iif(pDoc->GenerateStatusBar, BST_CHECKED, BST_UNCHECKED))
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
  
   ' BorderStyle is deprecated as of v2.0.4 as it has no effect
   ' in WinFBE programs where Windows Themes are enabled.
   'hCtrl = GetDlgItem( hForm, IDC_FRMSTATUSBAREDITOR_COMBOBORDERSTYLE )
   'ComboBox_AddString( hCtrl, @wstr("StatusBarPanelBorderStyle.None") )
   'ComboBox_AddString( hCtrl, @wstr("StatusBarPanelBorderStyle.Sunken") )
   'ComboBox_AddString( hCtrl, @wstr("StatusBarPanelBorderStyle.Raised") )

   hCtrl = GetDlgItem( hForm, IDC_FRMSTATUSBAREDITOR_COMBOAUTOSIZE )
   ComboBox_AddString( hCtrl, @wstr("StatusBarPanelAutoSize.None") )
   ComboBox_AddString( hCtrl, @wstr("StatusBarPanelAutoSize.Contents") )
   ComboBox_AddString( hCtrl, @wstr("StatusBarPanelAutoSize.Spring") )

   
   frmStatusBarEditor_DisplayListBox( nDefaultPanel )
   frmStatusBarEditor_SetTextboxes()
   SetFocus( GetDlgItem( HWND_FRMSTATUSBAREDITOR, IDC_FRMSTATUSBAREDITOR_LSTPANELS ) )

   ' Process Windows messages
   Function = pWindow->DoEvents( SW_SHOW )
   
   ' Delete the frmStatusBarEditor CWindow class manually allocated memory 
   Delete pWindow

   function = 0
end function

