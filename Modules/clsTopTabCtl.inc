

''
''
Property clsTopTabCtl.hWindow( ByVal HWnd As HWnd)
   m_hWindow = HWnd
End Property

Property clsTopTabCtl.hWindow() As HWnd
   Property = m_hWindow
End Property


''
''
Function clsTopTabCtl.GetTabIndexFromFilename( ByVal pwszName As WString Ptr ) As Long
   ' Get the index of tab that holds the incoming filename. This is used
   ' mostly for positioning the editor to the correct document when an
   ' error occurs during compiling. 
   ' Returns -1 if not found, otherwise zero based index of tab.
   
   If TabCtrl_GetItemCount(m_hWindow) = 0 Then Return -1 
   If pwszName = 0 Then Return -1
   
   Dim As Long i
   Dim pDoc As clsDocument Ptr 
   Dim As Long nCount = TabCtrl_GetItemCount(m_hWindow) 

   Dim TB As TC_ITEM 
   TB.Mask = TCIF_PARAM
   
   For i = 0 To nCount - 1
      TabCtrl_GetItem( m_hWindow, i, @TB )
      pDoc = Cast(clsDocument Ptr, TB.lParam)
      If pDoc Then 
         If Ucase(*pDoc->DiskFilename) = Ucase(*pwszName) Then Return i
      End If   
   Next

   Function = -1  ' if not found
End Function


''
''
Function clsTopTabCtl.SetTabIndexByDocumentPtr( ByVal pDocIn As clsDocument Ptr ) As Long
   ' Sets to tab that holds the incoming pDocIn ptr.
   If TabCtrl_GetItemCount(m_hWindow) = 0 Then Return -1 
   If pDocIn = 0 Then Return -1
   
   Dim As Long i
   Dim pDoc As clsDocument Ptr 
   Dim As Long nCount = TabCtrl_GetItemCount(m_hWindow) 

   Dim TB As TC_ITEM 
   TB.Mask = TCIF_PARAM
   
   For i = 0 To nCount - 1
      TabCtrl_GetItem( m_hWindow, i, @TB )
      pDoc = Cast(clsDocument Ptr, TB.lParam)
      If pDoc = pDocIn Then 
         this.SetFocusTab(i)
         Exit Function
      End If   
   Next

   Function = -1  ' if not found
End Function


''
''
Function clsTopTabCtl.SetFocusTab( ByVal idx As Long ) As Long
   ' Set the tab with focus (this sends the TCN_SELCHANGING and TCN_SELCHANGE 
   ' notification codes to its parent window). 
   Dim As Long nCount = TabCtrl_GetItemCount(m_hWindow) 
   If idx < 0 Then idx = nCount - 1
   If TabCtrl_GetCurSel(m_hWindow) <> idx Then
      Function = TabCtrl_SetCurFocus(m_hWindow, idx)
   End If
End Function

''
''
Function clsTopTabCtl.GetActiveDocumentPtr() As clsDocument Ptr
   ' Return pointer to clsDocument class for the current active tab
   If TabCtrl_GetItemCount(m_hWindow) = 0 Then Return 0 
      
   Dim TB As TC_ITEM 
   TB.Mask = TCIF_PARAM
   
   Dim As Long idx = TabCtrl_GetCurSel(m_hWindow)
   TabCtrl_GetItem( m_hWindow, idx, @TB )
   
   Function = Cast(clsDocument Ptr, TB.lParam)
End Function

''
''
Function clsTopTabCtl.GetDocumentPtr( ByVal idx As Long ) As clsDocument Ptr
   ' Return pointer to clsDocument class for the current active tab
   Dim TB As TC_ITEM 
   TB.Mask = TCIF_PARAM
   
   TabCtrl_GetItem( m_hWindow, idx, @TB )
   
   Function = Cast(clsDocument Ptr, TB.lParam)
End Function

''
''
Function clsTopTabCtl.DisplayScintilla( ByVal idx As Long, ByVal bShow As BOOLEAN ) As Long
   ' Show/Hide the Scintilla editing window for the incoming tab index
   Dim TB As TC_ITEM 
   TB.Mask = TCIF_PARAM
   
   TabCtrl_GetItem( m_hWindow, idx, @TB )
   
   Dim pDoc As clsDocument Ptr = Cast(clsDocument Ptr, TB.lParam)
   If pDoc Then 
      ShowWindow(pDoc->hWindow, Iif(bShow, SW_SHOWNORMAL, SW_HIDE))
   End If   
   
   Function = 0
End Function
   
''
''
Function clsTopTabCtl.AddTab( ByVal pDoc As clsDocument Ptr ) As Long
   Dim idx As Long
   Dim TB  As TC_ITEM 
   
   TB.Mask   = TCIF_PARAM
   TB.lParam = Cast(LPARAM, pDoc)
   
   Dim As Long NumItems = TabCtrl_GetItemCount(m_hWindow)             
   
   ' Return the index of the added tab
   idx = SendMessageW( m_hWindow, TCM_INSERTITEM, NumItems, Cast(LPARAM, @TB) )     
   
   this.SetTabText(idx)
   
   Function = idx
End Function

''
''
Function clsTopTabCtl.SetTabText( ByVal idx As Long ) As Long
   Dim TB    As TC_ITEM 
   Dim wText As WString * MAX_PATH  

   If idx = -1 Then idx = TabCtrl_GetCurSel(m_hWindow)

   ' Get the pointer to the clsDocument stored in lParam
   TB.Mask = TCIF_PARAM
   TabCtrl_GetItem( m_hWindow, idx, @TB )
   
   Dim pDoc As clsDocument Ptr = Cast(clsDocument Ptr, TB.lParam)
   If pDoc Then 
      ' Set the text that displays on the tab
      FF_PATHNAME( @WStr("NAMEX"), pDoc->DiskFilename, @wText)
      If SendMessageW( pDoc->hWindow, SCI_GETMODIFY, 0, 0 ) Then
         wText = wText & "*"
      End If   

      TB.Mask    = TCIF_TEXT 
      TB.pszText = @wText  
      Function = TabCtrl_SetItem( m_hWindow, idx, @TB )
   End If
   
End Function


' Global class to manage the top TabControl
Dim Shared gTTabCtl As clsTopTabCtl
