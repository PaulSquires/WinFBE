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
''  clsDocument (Class to handle everything related to a Scintilla editing window)
''

#include once "clsDocument.bi"
#include once "modVDToolbox.bi"
#include once "frmFindReplace.bi"
#include once "modVDRoutines.bi"
#include once "modVDControls.bi"
#include once "modVDDesignFrame.bi"
#include once "modVDDesignMain.bi"
#include once "modVDApplyProperties.bi"
#include once "modVDProperties.bi"
#include once "modGenerateCode.bi"
#include once "modMRU.bi"

   
''
''
Destructor clsDocument
   ' Delete any manually allocated CWindows in the designer
   for i as long = lbound(this.hWindow) to ubound(this.hWindow)
      DestroyWindow this.hWindow(i)    ' destroy the scintilla/designer windows
   NEXT
   DestroyWindow this.hScrollbar       ' destroy the scrollbar
   if this.IsDesigner THEN
      ' Remove all controls
      dim pCtrl as clsControl ptr 
      do until this.Controls.Count <= 1 
         for i as long = this.Controls.ItemFirst to this.Controls.ItemLast
            pCtrl = this.Controls.ItemAt(i)
            this.Controls.Remove(pCtrl)
         NEXT
      loop   
      for i as long = 0 to 3
         DestroyWindow(this.hSnapLine(i))
      next
      DestroyWindow(this.hWndForm)
      DestroyWindow(this.hWndFrame)
      DestroyWindow(this.hWndDesigner)
      DestroyWindow(this.hDesignTabCtrl)
      DeleteObject(this.hFontFakeMenu)
   END IF

   ' Repaint the main area because we don't want any splitter to show
   AfxRedrawWindow(HWND_FRMMAIN)
   
End Destructor

''
''
function clsDocument.GetActiveScintillaPtr() as any ptr
   dim as hwnd hEdit = this.hWndActiveScintilla
   for i as long = lbound(this.hWindow) to ubound(this.hWindow)
      if this.hWindow(i) = hEdit THEN return m_pSci(i)
   NEXT
   ' If no other matches then return 
   function = m_pSci(0)
end function


property clsDocument.hWndActiveScintilla() as hwnd
   if m_hWndActiveScintilla = 0 THEN m_hWndActiveScintilla = this.hWindow(0)
   property = m_hWndActiveScintilla
end property

property clsDocument.hWndActiveScintilla(byval hWindow as hwnd)
   m_hWndActiveScintilla = hWindow
end property

''
''  Returns True/False indicating that a valid top mainmenu exists for this form.
function clsDocument.MainMenuExists() as Boolean
   dim as long numItems = ubound(this.MenuItems) - 1
   if (this.GenerateMenu = BST_CHECKED) andalso (numItems > -1) then
      return true
   else   
      return false
   end if
end function

''
''  Returns True/False indicating that a valid top toolbar exists for this form.
function clsDocument.ToolBarExists() as Boolean
   dim as long numItems = ubound(this.ToolBarItems) - 1
   if (this.GenerateToolBar = BST_CHECKED) andalso (numItems > -1) then
      return true
   else   
      return false
   end if
end function

''
''  Returns True/False indicating that a valid top statusbar exists for this form.
function clsDocument.StatusBarExists() as Boolean
   dim as long numItems = ubound(this.PanelItems) - 1
   if (this.GenerateStatusBar = BST_CHECKED) andalso (numItems > -1) then
      return true
   else   
      return false
   end if
end function

''
''
Function clsDocument.CreateDesignerWindow( ByVal hWndParent As HWnd ) as hwnd 
   
   ' For the Visual Designer, there exists three (3) levels of windows:
   ' (1) The DesignMain (hWindow) used by the top tabcontrol to display the document. 
   '     This is the highest level window and is basically just the container for 
   '     the other two windows. Contains tab control to switch between design/code views.
   ' (2) The DesignFrame window. This is the scrollable window.
   ' (3) The DesignForm window. This is the actual visual for that we manipulate by
   '     adding controls to it, etc.
   ' (*4) There is a 4th window in the sense that the Scintilla code window swaps
   '      out the DesignFrame whenever the tabcontrol switches between design/code view.
   '
   this.IsDesigner = true
   this.FileEncoding = FILE_ENCODING_UTF16_BOM
   
   dim rc as RECT
   
   ' (1) Create the DesignMain window
   dim pMain as CWindow ptr = New CWindow
   pMain->DPI = AfxCWindowPtr(hwndParent)->DPI
   this.hWndDesigner = _
   pMain->Create( hWndParent, "", @DesignerMain_WndProc, 0, 0, 0, 0, _
        WS_CHILD or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, _
        WS_EX_CONTROLPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pMain->Brush = GetSysColorBrush(COLOR_WINDOW)
   pMain->SetClientSize(2500, 2500)
   ' We will set our own mouse pointer as needed
   SetClassLongPtr(this.hWindow(0), GCLP_HCURSOR, 0)   
   ' Allow the Designer Frame window to scroll. This allows us to create Forms that
   ' are larger than the current viewable screen area.
   DIM pScrollWindow AS CScrollWindow PTR = NEW CScrollWindow(pMain->hWindow)
   pMain->ScrollWindowPtr = pScrollWindow
   
   ' Create a tabcontrol that allows us to switch between design/code views.
   this.hDesignTabCtrl = _
      pMain->AddControl("TABCONTROL", hWndParent, _
         IDC_DESIGNTABCTRL, "", 0, 0, 0, 24, _
         WS_CHILD Or WS_TABSTOP Or TCS_SINGLELINE Or TCS_RAGGEDRIGHT Or TCS_HOTTRACK Or _
         TCS_TABS Or TCS_FOCUSNEVER, WS_EX_LEFT Or WS_EX_LTRREADING )
   
   dim wszTemp as wstring * MAX_PATH
   Dim TB As TC_ITEM 
   TB.Mask = TCIF_TEXT
   
   ' Add the two tabs. Design and Code views.
   wszTemp = L(327, "Design")
   TB.pszText = @wszTemp
   SendMessage( this.hDesignTabCtrl, TCM_INSERTITEM, 0, Cast(LPARAM, @TB) )     

   wszTemp = L(328, "Code")
   TB.pszText = @wszTemp
   SendMessage( this.hDesignTabCtrl, TCM_INSERTITEM, 1, Cast(LPARAM, @TB) )     

   TabCtrl_SetCurSel( this.hDesignTabCtrl, 0 )

   
   ' (2) Create the Design Frame window (child of the Main)
   dim pFrame as CWindow ptr = New CWindow
   pFrame->DPI = AfxCWindowPtr(hwndParent)->DPI
   this.hWndFrame = _
   pFrame->Create( pMain->hWindow, "", @DesignerFrame_WndProc, 0, 0, 0, 0, _
        WS_CHILD or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN or WS_VISIBLE, _
        WS_EX_CONTROLPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   'pFrame->ClassStyle = CS_DBLCLKS
   SetWindowLongPtr( pFrame->hWindow, GWLP_ID, IDC_DESIGNFRAME )
   ' We will set our own mouse pointer as needed
   SetClassLongPtr(this.hWndFrame, GCLP_HCURSOR, 0)   
   pFrame->Brush = GetSysColorBrush(COLOR_WINDOW)
   
   
   ' (3) Create the Design Form (child of the Frame)
   dim pCtrl as clsControl ptr
   if this.IsNewFlag then
      this.UserModified = true
      SetRect(@rc, 10, 10, 510, 310)
      pCtrl = CreateToolboxControl( @this, CTRL_FORM, rc )
   end if
   
   
   function = this.hWindow(0)
end function


''
''
Function clsDocument.CreateCodeWindow( ByVal hWndParent As HWnd, _
                                       ByVal IsNewFile  As BOOLEAN, _     
                                       ByVal IsTemplate As BOOLEAN = False, _
                                       ByVal pwszFile   As WString Ptr = 0 _
                                       ) As HWnd   

   ' Creates a Scintilla editing window (initially not visible). Optionally, load a diskfile
   ' into the window and apply properties to it.
   for i as long = lbound(this.hWindow) to ubound(this.hWindow)
      this.hWindow(i) = CreateWindowEx( 0, "Scintilla", "", _
                        WS_CHILD Or WS_TABSTOP Or WS_CLIPCHILDREN, _
                        0,0,0,0,hWndParent, _
                        Cast(HMENU, IDC_SCINTILLA+i), GetModuleHandle(Null), Null)
   
      ' Initialize our direct access to the Scintilla code windows. This is much faster than
      ' using SendMessage to the window. Only need to initialize once no matter how many
      ' code windows that are eventually opened.
      If IsWindow(this.hWindow(i)) Then
         If SciMsg = 0 Then
            SciMsg = Cast( Scintilla_DirectFunction, SendMessage( this.hWindow(0), SCI_GETDIRECTFUNCTION, 0, 0 ) )
         End If
         ' Call the direct function for speed purposes rather than relying on the traditional SendMessage method.
         m_pSci(i) = Cast(Any Ptr, SendMessage( this.hWindow(i), SCI_GETDIRECTPOINTER, 0, 0 )) 
      End If
   next
   
   ' Disable main scintilla control's vertical scroll bar (wParam = 1 to enable)
   SciMsg( m_pSci(0), SCI_SETVSCROLLBAR, 0, 0 )

   ' Create our substitute vertical scrollbar
   dim pWindow as CWindow
   this.hScrollbar = pWindow.AddControl("VSCROLLBAR", hWndParent, IDC_SCROLLV)
   ShowWindow(this.hScrollbar, SW_HIDE)
   
   ' Initialize scroll bar information
   this.ScrInfo.cbSize  = SizeOf(SCROLLINFO)
   this.ScrInfo.fMask   = SIF_ALL


   ' Get the document pointer from our main control and assign it to the other split windows
   dim as any ptr pDoc = cast(any ptr, SciMsg(m_pSci(0), SCI_GETDOCPOINTER, 0, 0))
   If pDoc Then SciMsg( m_pSci(1), SCI_SETDOCPOINTER, 0, cast(LPARAM, pDoc)) 
   
   Dim nResult As Long = IS_TEXT_UNICODE_SIGNATURE
      
   ' If a disk file was specified then open it and load it into the editor
   this.IsNewFlag = IsNewFile
   If (IsNewFile = true) orelse (IsTemplate = true) Then
      this.NextFileNum = this.NextFileNum + 1
      this.DiskFilename = "Untitled" & this.NextFileNum
      if this.IsDesigner then
         this.CreateDesignerWindow(hWndParent)  ' Create the new visual designer window
      end if   
   End If
   this.ProjectFileType = FILETYPE_UNDEFINED
   
   If pwszFile Then   
      ' Do not use Dir() > "" here b/c if incoming file originated from a Do/Loop 
      ' of files using Dir() then there will be problems.

      If AfxFileExists(*pwszFile) Then     
         Dim As String st, sText
         Dim As Long f = Freefile
         Dim As Long idx
         
         If IsTemplate Then
            dim pStream as CTextStream
            if pStream.Open(*pwszFile) = S_OK then
               ' Look at the first 4 lines
               ' Line 3 tells us the file type (bas or xml)
               Do Until pStream.EOS
                  st = pStream.ReadLine
                  idx = idx + 1
                  Select Case idx
                     Case 1
                     Case 2
                     Case 3: this.DiskFilename = this.DiskFilename & Trim(st)    
                     Case 4 
                     Case Else
                        sText = sText & st & vbCrLf
                  End Select   
               Loop
               pStream.Close
               If IsTextUnicode(StrPtr(sText), 2, Cast(LPINT, @nResult) ) Then
                  sText = Mid(sText, 3)
                  sText = AfxACode( Cast(WSTRING Ptr, StrPtr(sText)) )
               End If   
               this.SetText( sText ) 
            End If
            ' Force the template file to be considered "new" so it will be saved.
            this.IsNewFlag = true
            ' Search for "|", replace it with an empty space "" and position the caret in that place
            this.FindReplace( "|", "" )
            ' don't set SAVEPOINT for newly loaded Template files because we want the document to display as dirty

         Else

            ' Set a flag that we are loading code from a file. This is important because if the file
            ' contains visual designer code then we only want to apply control properties after all of 
            ' the file is read, otherwise it would get applied immediately after each call to CreateToolboxControl.
            this.LoadingFromFile = true
            dim sText as string  ' this will be an UTF-8 encoded string
            ' GetFileToString has a call to pDoc->ParseFormMetaData which will create/recreate
            ' the designer form.
            if GetFileToString(*pwszFile, sText, @this) = false then
               this.SetText( sText ) 
               this.DateFileTime = AfxGetFileLastWriteTime( *pwszFile )
            else
               print "Error opening: "; *pwszFile
            End If
            this.LoadingFromFile = false
            this.DiskFilename = *pwszFile
            SciMsg( m_pSci(0), SCI_SETSAVEPOINT, 0, 0)
            SciMsg( m_pSci(0), SCI_EMPTYUNDOBUFFER, 0, 0)
            ' Update the most recently used file list (only for non-IsNewFlag files)
            ' Only add file to MRU list if it is not part of an active Project.
            If gApp.IsProjectActive = false Then 
               UpdateMRUList(*pwszFile)
            end if
         End If
      End If   
   End If

   if this.IsDesigner then
      ' Select the Form as the default focus control
      this.Controls.SelectControl( this.hWndForm )
      this.Controls.SetActiveControl( this.hWndForm )

      ' If the data that was read contained anything that caused the form to have
      ' to be code regenerated then we need to save the file back to disk with the
      ' new meta data and code generation.
      if this.bRegenerateCode then 
         If this.IsNewFlag = false then this.SaveFile
      end if   

      ' In any event, regenerate code for the visual designer 
      this.bRegenerateCode = true
      GenerateFormCode( @this )
   end if   

   this.ApplyProperties()

   Function = this.hWindow(0)
End Function

''
''
Function clsDocument.FindReplace( ByVal strFindText As String, ByVal strReplaceText As String ) As Long
   ' Return Value: New position if successful; -1 if text not found.
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   Dim As Long findFlags = SCFIND_MATCHCASE Or SCFIND_WHOLEWORD
   Dim As Long startPos  = SciMsg( pSci, SCI_GETCURRENTPOS, 0, 0)
   Dim As Long endPos    = SciMsg( pSci, SCI_GETTEXTLENGTH, 0, 0)
   Dim As Long newPos

   ' Set the start and end positions, and search flags, and finally do the search
   SciMsg( pSci, SCI_SETTARGETSTART, startPos, 0)
   SciMsg( pSci, SCI_SETTARGETEND, endPos, 0)
   SciMsg( pSci, SCI_SETSEARCHFLAGS, findFlags, 0)

   ' Search the text to replace
   newPos = SciMsg( pSci, SCI_SEARCHINTARGET, Len(strFindText), Cast(LPARAM, Strptr(strFindText)) )
   
   ' Return -1 if not found
   If newPos = - 1 Then Return -1
   
   gApp.SuppressNotify = true
   ' Position the caret and select the text
   SciMsg( pSci, SCI_SETCURRENTPOS, newPos, 0)
   SciMsg( pSci, SCI_GOTOPOS, newPos, 0)
   SciMsg( pSci, SCI_SETSELECTIONSTART, newPos, 0)
   SciMsg( pSci, SCI_SETSELECTIONEND, newPos + Len(strFindText), 0)

   ' Replace the selection (SCI_REPLACESEL fails if text is "" so use Cut instead for that scenario)                                
   If Len(strReplaceText) = 0 Then
      SciMsg( pSci, SCI_CUT, 0, 0 )
   Else
      SciMsg( pSci, SCI_REPLACESEL, 0, Cast(LPARAM, Strptr(strReplaceText)) )
   End If
   gApp.SuppressNotify = false
   
   ' Return the new position
   Function = newPos

End Function

''
''
Function clsDocument.InsertFile() As BOOLEAN
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   ' Display the Open File Dialog
   Dim pwszName As WString Ptr = AfxIFileOpenDialogW(HWND_FRMMAIN, IDM_INSERTFILE)
   If pwszName Then
      ' save the main file encoding because GetFileToString may change it
      Dim As String sText
      GetFileToString(*pwszName, sText, @this)
      SciMsg( pSci, SCI_INSERTTEXT, -1, Cast(LPARAM, Strptr(sText)) )
      CoTaskMemFree pwszName
      this.ApplyProperties()
   End If
   
   Function = 0
End Function


' ========================================================================================
' Parse all of the file's code to remove and process any visual designer specific code.
' Returns a string representing just the code only (visual designer metastatements removed).
' ========================================================================================
function clsDocument.ParseFormMetaData( ByVal hWndParent As HWnd, byref wszAllText as wstring ) as CWSTR
   
   ' NOTE: The incoming wszAllText string is UTF-16 encoded. The resulting string that
   ' is returned from this function is also UTF-16 encoded. All WinFBE form files
   ' must be unicode encoded.
   
   ' This function filters out all form metadata as well as any code generated code between
   ' the two codegen tags:
   '    ' WINFBE_CODEGEN_START
   '    ' WINFBE_CODEGEN_END

   
   dim pCtrl as clsControl ptr
   dim pCtrlActive as clsControl ptr
   
   dim as RECT rc, rcCtrl
   dim as CWSTR wszControlType, wszPropName, wszPropValue, wszEventName
   dim as CWSTR wst
   dim as Long nControlType, numLines, numOffsetLines
   dim as Boolean bIsValidControl = true
   
   ' The first line MUST contain the identifier that this is a form file.

   if left(wszAllText, 13) = "' WINFBE FORM" then
      this.IsDesigner = true
      ' Remove any previously created controls, etc otherwise reloading the form
      ' file will result in duplicates.
      if this.hWndForm then DestroyWindow(this.hWndForm)
      if this.hWndFrame then DestroyWindow(this.hWndFrame)
      if this.hWndDesigner then DestroyWindow(this.hWndDesigner)
      if this.hDesignTabCtrl then DestroyWindow(this.hDesignTabCtrl)
      this.CreateDesignerWindow(hWndParent)
      numOffsetLines = 1
   else
      ' This is just a regular code file with no visual designer
      this.IsDesigner = false
      return wszAllText   
   END IF
   

   ' Iterate all of the lines related to the visual designer
   dim as long iLineStart = 1
   dim as long iLineEnd, NextMenuItem, NextToolBarItem, NextPanelItem, NextImageItem

   dim as boolean bReadingMenuItem, bReadingToolBarItem, bReadingPanelItem, bReadingImageItem

   do until iLineStart >= len(wszAllText)
      
      iLineEnd = instr(iLineStart, wszAllText, vbcrlf)
      if iLineEnd = 0 then iLineEnd = len(wszAllText)  ' cr/lf not found
      wst = mid(wszAllText, iLineStart, iLineEnd - iLineStart)   
      iLineStart = iLineStart + len(wst) + len(vbcrlf)
      
      numOffsetLines = numOffsetLines + 1
      
      if left(wst, 16) = "' WINFBE VERSION" then
         wszPropValue = trim(mid(**wst, 17))
         if ConvertWinFBEversion(wszPropValue) < ConvertWinFBEversion(APPVERSION) then
            this.bRegenerateCode = true
            this.UserModified = true
         end if
         
      elseif left(wst, 15) = "' LOCKCONTROLS=" then
         wszPropValue = mid(**wst, 16)  ' default is False
         if wszPropValue = "True" then this.bLockControls = true

      elseif left(wst, 12) = "' SNAPLINES=" then
         wszPropValue = mid(**wst, 13)  ' default is True
         if wszPropValue = "False" then this.bSnapLines = false

      elseif left(wst, 17) = "' WINFBE FORM_END" then
         this.ErrorOffset = numOffsetLines 
      
      'elseif left(wst, 23) = "' WINFBE MAINMENU_START" then
      'elseif left(wst, 21) = "' WINFBE IMAGES_START" then
      
      elseif left(wst, 18) = "'   MENUITEM_START" then
         dim as long ub = ubound(this.MenuItems)
         redim preserve this.MenuItems(ub + 1)
         NextMenuItem = ub + 1
         bReadingMenuItem = true
      elseif left(wst, 21) = "'   MAINMENU_DISPLAY=" then
         this.GenerateMenu = val(mid(**wst, 22))
      elseif left(wst, 11) = "'     NAME=" then
         if bReadingMenuItem then this.MenuItems(NextMenuItem).wszName = mid(**wst, 12)
      elseif left(wst, 14) = "'     CAPTION=" then
         if bReadingMenuItem then this.MenuItems(NextMenuItem).wszCaption = mid(**wst, 15)
      elseif left(wst, 13) = "'     INDENT=" then
         if bReadingMenuItem then this.MenuItems(NextMenuItem).nIndent = val(mid(**wst, 14))
      elseif left(wst, 10) = "'     ALT=" then
         if bReadingMenuItem then this.MenuItems(NextMenuItem).chkAlt = val(mid(**wst, 11))
      elseif left(wst, 12) = "'     SHIFT=" then
         if bReadingMenuItem then this.MenuItems(NextMenuItem).chkShift = val(mid(**wst, 13))
      elseif left(wst, 11) = "'     CTRL=" then
         if bReadingMenuItem then this.MenuItems(NextMenuItem).chkCtrl = val(mid(**wst, 12))
      elseif left(wst, 15) = "'     SHORTCUT=" then
         if bReadingMenuItem then this.MenuItems(NextMenuItem).wszShortcut = mid(**wst, 16)
      elseif left(wst, 14) = "'     CHECKED=" then
         if bReadingMenuItem then this.MenuItems(NextMenuItem).chkChecked = val(mid(**wst, 15))
      elseif left(wst, 13) = "'     GRAYED=" then
         if bReadingMenuItem then this.MenuItems(NextMenuItem).chkGrayed = val(mid(**wst, 14))
      elseif left(wst, 16) = "'   MENUITEM_END" then
         bReadingMenuItem = false   
      
      
      elseif left(wst, 21) = "'   TOOLBARITEM_START" then
         dim as long ub = ubound(this.ToolBarItems)
         redim preserve this.ToolBarItems(ub + 1)
         NextToolBarItem = ub + 1
         bReadingToolBarItem = true
      elseif left(wst, 20) = "'   TOOLBAR_DISPLAY=" then
         this.GenerateToolBar = val(mid(**wst, 21))
      elseif left(wst, 17) = "'   TOOLBAR_SIZE=" then
         this.wszToolBarSize = mid(**wst, 18)
      elseif left(wst, 17) = "'     BUTTONNAME=" then
         if bReadingToolBarItem then this.ToolBarItems(NextToolBarItem).wszName = mid(**wst, 18)
      elseif left(wst, 17) = "'     BUTTONTYPE=" then
         if bReadingToolBarItem then this.ToolBarItems(NextToolBarItem).wszButtonType = mid(**wst, 18)
      elseif left(wst, 20) = "'     BUTTONTOOLTIP=" then
         if bReadingToolBarItem then this.ToolBarItems(NextToolBarItem).wszTooltip = mid(**wst, 21)
      elseif left(wst, 24) = "'     BUTTONNORMALIMAGE=" then
         if bReadingToolBarItem then this.ToolBarItems(NextToolBarItem).pPropNormalImage.wszPropValue = mid(**wst, 25)
      elseif left(wst, 21) = "'     BUTTONHOTIMAGE=" then
         if bReadingToolBarItem then this.ToolBarItems(NextToolBarItem).pPropHotImage.wszPropValue = mid(**wst, 22)
      elseif left(wst, 26) = "'     BUTTONDISABLEDIMAGE=" then
         if bReadingToolBarItem then this.ToolBarItems(NextToolBarItem).pPropDisabledImage.wszPropValue = mid(**wst, 27)
      elseif left(wst, 19) = "'   TOOLBARITEM_END" then
         bReadingToolBarItem = false   


      elseif left(wst, 19) = "'   PANELITEM_START" then
         dim as long ub = ubound(this.PanelItems)
         redim preserve this.PanelItems(ub + 1)
         NextPanelItem = ub + 1
         bReadingPanelItem = true
      elseif left(wst, 22) = "'   STATUSBAR_DISPLAY=" then
         this.GenerateStatusBar = val(mid(**wst, 23))
      elseif left(wst, 16) = "'     PANELNAME=" then
         if bReadingPanelItem then this.PanelItems(NextPanelItem).wszName = mid(**wst, 17)
      elseif left(wst, 16) = "'     PANELTEXT=" then
         if bReadingPanelItem then this.PanelItems(NextPanelItem).wszText = mid(**wst, 17)
      elseif left(wst, 19) = "'     PANELTOOLTIP=" then
         if bReadingPanelItem then this.PanelItems(NextPanelItem).wszTooltip = mid(**wst, 20)
      elseif left(wst, 21) = "'     PANELALIGNMENT=" then
         if bReadingPanelItem then this.PanelItems(NextPanelItem).wszAlignment = mid(**wst, 22)
      elseif left(wst, 23) = "'     PANELBORDERSTYLE=" then
         if bReadingPanelItem then this.PanelItems(NextPanelItem).wszBorderStyle = mid(**wst, 24)
      elseif left(wst, 20) = "'     PANELAUTOSIZE=" then
         if bReadingPanelItem then this.PanelItems(NextPanelItem).wszAutoSize = mid(**wst, 21)
      elseif left(wst, 17) = "'     PANELWIDTH=" then
         if bReadingPanelItem then this.PanelItems(NextPanelItem).wszWidth = mid(**wst, 18)
      elseif left(wst, 20) = "'     PANELMINWIDTH=" then
         if bReadingPanelItem then this.PanelItems(NextPanelItem).wszMinWidth = mid(**wst, 21)
      elseif left(wst, 17) = "'     PANELIMAGE=" then
         if bReadingPanelItem then this.PanelItems(NextPanelItem).pProp.wszPropValue = mid(**wst, 18)
      elseif left(wst, 17) = "'   PANELITEM_END" then
         bReadingPanelItem = false   

      
      elseif left(wst, 15) = "'   IMAGE_START" then
         dim as long ub = ubound(this.AllImages)
         redim preserve this.AllImages(ub + 1)
         NextImageItem = ub + 1
         this.AllImages(NextImageItem).pDoc = @this
         bReadingImageItem = true
      elseif left(wst, 16) = "'     IMAGENAME=" then
         if bReadingImageItem then this.AllImages(NextImageItem).wszImageName = mid(**wst, 17)
      elseif left(wst, 15) = "'     FILENAME=" then
         if bReadingImageItem then this.AllImages(NextImageItem).wszFileName = ProcessFromCurdrive(mid(**wst, 16))
      elseif left(wst, 19) = "'     RESOURCETYPE=" then
         if bReadingImageItem then this.AllImages(NextImageItem).wszFormat = mid(**wst, 20)
      elseif left(wst, 13) = "'   IMAGE_END" then
         bReadingImageItem = false   

      'elseif left(wst, 16) = "'   EVENTS_START" then
      'elseif left(wst, 14) = "'   EVENTS_END" then
      'elseif left(wst, 19) = "' WINFBE IMAGES_END" then
      'elseif left(wst, 21) = "' WINFBE MAINMENU_END" then

      elseif left(wst, 20) = "' WINFBE_CODEGEN_END" then
         this.ErrorOffset = numOffsetLines 
         wszAllText = ltrim(AfxStrRemain( wszAllText, "' WINFBE_CODEGEN_END" ), vbcrlf)
         return wszAllText
         
      elseif left(wst, 23) = "' WINFBE CONTROL_START " then
         ' The control type name is parse #4 (blank space)
         wszControlType = AfxStrParse(wst, 4, " ")
         nControlType = GetControlType(wszControlType)
         if nControlType = 0 then   
            ' no longer a valid toolbox control
            bIsValidControl = false
         else
            bIsValidControl = true
            pCtrl = CreateToolboxControl( @this, nControlType, rc )
            pCtrlActive = pCtrl
         end if
         
      elseif left(wst, 20) = "' WINFBE CONTROL_END" then
         if bIsValidControl then
            pCtrl->SuspendLayout = true
            ApplyControlProperties( @this, pCtrl )
            pCtrl->SuspendLayout = false
         end if
      'elseif left(sLine, 20) = "'   PROPERTIES_START" then
      'elseif left(sLine, 18) = "'   PROPERTIES_END" then

      elseif left(wst, 16) = "'     PROP_NAME=" then
         if bIsValidControl then
            wszPropName = mid(**wst, 17)
         end if
         
      elseif left(wst, 17) = "'     PROP_VALUE=" then
         if bIsValidControl then
            wszPropValue = mid(**wst, 18)   ' utf8 encoded 
            ' Only set the loading property if it exists in the current property
            ' listing. We do this otherwise older now unused properties will continue
            ' to get loaded when we no longer want them to.
            if IsPropertyExists(pCtrl, wszPropName) = false then
               ' Set the flag to regenerate code otherwise a compile time error
               ' will occur because the old property could exist in previously generated code.
               this.bRegenerateCode = true
            else   
               SetControlProperty(pCtrl, wszPropName, wszPropValue)
            end if
         end if
         
      'elseif left(sLine, 16) = "'   EVENTS_START" then
      'elseif left(sLine, 14) = "'   EVENTS_END" then

      elseif left(wst, 17) = "'     EVENT_NAME=" then
         if bIsValidControl then
            wszEventName = mid(**wst, 18)
            SetControlEvent(pCtrl, wszEventName, true)
         end if
      end if   

   loop

   return wszAllText
   
end function

''
''
Function clsDocument.SaveFile( ByVal bSaveAs As BOOLEAN = False ) As BOOLEAN

   Dim sText     As String
   Dim wFilename As WString * MAX_PATH
   Dim wText     As WString * MAX_PATH  
   Dim f         As Long
   
   ' If this is a new Untitled document then set flag to display SaveAs dialog.
   If this.IsNewFlag Then bSaveAs = true
      
   wFilename = this.Diskfilename
   
   If bSaveAs Then
      ' Display the Save File Dialog
      Dim pwszName As WString Ptr = AfxIFileSaveDialog(HWND_FRMMAIN, @wFilename, "", IDM_FILESAVEAS)
      If pwszName Then
         wFilename = *pwszName
         CoTaskMemFree(pwszName)
      Else
         Return False
      End If
   End If

   dim as any ptr pSci = this.GetActiveScintillaPtr()
   ' If this is a FORM document then we need to create the string of controls and properties that gets
   ' saved at the start of the file before any of the code is displayed.
   if this.IsDesigner then
      this.bRegenerateCode = true
      GenerateFormMetaData(@this)     ' recreate the metadata
      GenerateFormCode(@this)         ' regenerate visual designer code
      this.ErrorOffset = 0
      this.ErrorOffset = this.ErrorOffset + AfxStrTally(this.wszFormMetaData, vbcrlf) 
      this.ErrorOffset = this.ErrorOffset + AfxStrTally(this.wszFormCodegen, vbcrlf) 
   END IF
   ' Save text buffer to disk by directly accessing buffer rather
   ' saving it to an intermediary string variable first.
   Dim As ZString Ptr psz = Cast( ZString Ptr, SciExec(this.hWindow(0), SCI_GETCHARACTERPOINTER, 0, 0) )
   
   dim as long sciCodePage = SciMsg(pSci, SCI_GETCODEPAGE, 0, 0)   ' 0 or SC_CP_UTF8 
   
   if AfxFileExists(wFilename) THEN Kill(wFilename)
      
   dim pStream as CFileStream
   if pStream.Open(wFilename, STGM_CREATE or STGM_WRITE) = S_OK then
      dim as string st
      select CASE this.FileEncoding
         case FILE_ENCODING_ANSI
            if sciCodePage = 0 THEN    
               pStream.Write psz, len(*psz)     ' no conversion necessary
            else
               ' need to convert
               st = Utf8ToAscii(*psz)
               pStream.Write strptr(st), len(st) 
            end if    

         case FILE_ENCODING_UTF8_BOM
            ' Output the BOM first
            st = chr(&HEF, &HBB, &HBF)
            pStream.Write strptr(st), len(st)
            if sciCodePage = SC_CP_UTF8 THEN    
               ' no conversion necessary
               pStream.Write psz, len(*psz)     ' no conversion necessary
            else
               ' need to convert
               st = AnsiToUtf8(*psz)
               pStream.Write strptr(st), len(st)
            end if    

         case FILE_ENCODING_UTF16_BOM
            ' Output the BOM first
            st = chr(&HFF, &HFE)
            pStream.Write strptr(st), len(st)
            if sciCodePage = SC_CP_UTF8 THEN    
               pStream.Write this.wszFormMetaData.m_pBuffer, this.wszFormMetaData.m_BufferLen
               pStream.Write this.wszFormCodeGen.m_pBuffer, this.wszFormCodeGen.m_BufferLen
               ' convert utf8 to utf16
               st = Utf8ToUnicode(*psz)   ' use a regular string b/c it is a byte buffer
               pStream.Write strptr(st), len(st)
            else
               pStream.Write this.wszFormMetaData.m_pBuffer, this.wszFormMetaData.m_BufferLen
               pStream.Write this.wszFormCodeGen.m_pBuffer, this.wszFormCodeGen.m_BufferLen
               ' need to convert ansi to unicode
               dim as CWSTR wst = WStr(*psz)
               pStream.Write wst.m_pBuffer, wst.m_BufferLen
            end if    
            
      END SELECT
   end if
   pStream.Close

   this.DiskFilename = wFilename
   this.DateFileTime = AfxGetFileLastWriteTime( wFilename )
   this.UserModified = false
   
   ' If this was a new document then it needs to be saved to Recent File list.
   If this.IsNewFlag Then 
      If gApp.IsProjectActive = false Then 
         UpdateMRUList(wFilename)
      end if
   end if
   this.IsNewFlag = False
   
   ' Set the current state of the document to unmodified
   SciMsg( pSci, SCI_SETSAVEPOINT, 0, 0)
   
   Function = True
End Function


''
''
Function clsDocument.GetTextRange( ByVal cpMin As Long, ByVal cpMax As Long) As String
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   Dim p      As Long
   Dim buffer As String
   Dim txtrg  As SCI_TEXTRANGE
   txtrg.chrg.cpMin = cpMin
   txtrg.chrg.cpMax = cpMax
   buffer = Space(cpMax - cpMin + 1)
   txtrg.lpstrText = Strptr(buffer)
   SciMsg(pSci, SCI_GETTEXTRANGE, 0, cast(LPARAM, @txtrg))
   p = Instr(buffer, Chr(0))
   If p Then buffer = Left(buffer, p - 1)
   Function = buffer
End Function

''
''
Function clsDocument.ChangeSelectionCase( ByVal fCase As Long) As Long 

   dim as any ptr pSci = this.GetActiveScintillaPtr()
   Dim startSelPos As Long     ' Starting position
   Dim endSelPos   As Long     ' Ending position
   Dim strText     As String   ' Selected text
   Dim i           As Long 

   ' fCase = 1 (upper case), 2 (lower case), 3 (mixed case)
   If (fCase < 1) Or (fCase > 3) Then Exit Function
   
   ' If startSelPos and endSelPos are the same there is not selection,
   startSelPos = SciMsg( pSci, SCI_GETSELECTIONSTART, 0, 0)
   endSelPos   = SciMsg( pSci, SCI_GETSELECTIONEND, 0, 0)
   If startSelPos = endSelPos Then Exit Function
   
   ' Retrieve the text
   strText = this.GetTextRange(startSelPos, endSelPos)
   
   ' Convert it to upper or lower case
   If fCase = 1 Then
      strText = Ucase(strText)
   ElseIf fCase = 2 Then
      strText = LCase(strText)
   ElseIf fCase = 3 Then
      ' Always uppercase the first character regardless
      Mid(strText,1,1) = Ucase(Left(strText,1))
      Do
          i = Instr(i+1, strText, " ")
          If i = 0 Then Exit Do
          Mid(strText,i+1,1) = Ucase(Mid(strText,i+1,1))
      Loop
   End If
   
   ' Replace the selected text
   SciMsg( pSci, SCI_REPLACESEL, 0, Cast(LPARAM, Strptr(strText)))

   Function = 0
End Function

''
''
Function clsDocument.SetMarkerHighlight() As Long
   ' Set a marker that will highlight the background of the current selection. This
   ' is used when we are attempting to search a selection. We want the current search
   ' area to be a different color than the regular highlighted text because any
   ' search results are colored using the normal highlight colors so we need them
   ' to stand out from the selected range.
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   Dim As Long startPos, endPos, startLine, endLine
   this.GetSelectedLineRange( startLine, endLine, startPos, endPos )
   if endLine <> startLine THEN
      SciMsg( pSci, SCI_MARKERDEFINE, 10, SC_MARK_BACKGROUND )  ' define as marker #10
      SciMsg( pSci, SCI_SETMARGINMASKN, 4, &H400 )   ' set margin mask to allow SC_MARK_BACKGROUND 
      for i as long = startLine to endLine
         function = SciMsg( pSci, SCI_MARKERADD, i, 10)    ' add defined marker #10 to each line
      NEXT
      SciMsg( pSci, SCI_MARKERSETBACK, 10, BGR(220,220,220)) ' set backcolor of marker #10 
   END IF
End Function

''
''
Function clsDocument.RemoveMarkerHighlight() As Long
   ' Remove any markers that were set in the document that signify a highlighted range.
   ' This is used when we are attempting to search a selection. 
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   function = SciMsg( pSci, SCI_MARKERDELETEALL, 10, 0)  ' delete all marker #10
End Function

''
''
Function clsDocument.FirstMarkerHighlight() As long
   ' Get the first line with marker #10 highlight
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   dim as long markerMask 
   markerMask = Bitset(markerMask, 10)
   function = SciMsg( pSci, SCI_MARKERNEXT, 0, markerMask)
End Function

''
''
Function clsDocument.LastMarkerHighlight() As long
   ' Get the first line with marker #10 highlight
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   dim as long nLastPos = SciMsg( pSci, SCI_GETTEXTLENGTH, 0, 0)
   dim as long nLastLine = SciMsg( pSci, SCI_LINEFROMPOSITION, nLastPos, 0)
   dim as long markerMask 
   markerMask = Bitset(markerMask, 10)
   function = SciMsg( pSci, SCI_MARKERPREVIOUS, nLastLine, markerMask)
End Function

''
''
Function clsDocument.HasMarkerHighlight() As BOOLEAN
   ' True/False if selection markers exist in the document search for marker #10
   function = iif(this.FirstMarkerHighlight = -1, false, true)
End Function

''
''
Function clsDocument.GetCurrentLineNumber() As Long
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   Dim nPos As Long = SciMsg( pSci, SCI_GETCURRENTPOS, 0, 0)
   Function = SciMsg( pSci, SCI_LINEFROMPOSITION, nPos, 0)
End Function

''
''
Function clsDocument.SelectLine( ByVal nLineNum As Long ) As Long
   ' Select the incoming nLineNum. If nLineNum is negative then select the current line
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   nLineNum = Iif( nLineNum < 0, this.GetCurrentLineNumber, nLineNum)
   Dim nStartPos As Long = SciMsg( pSci, SCI_POSITIONFROMLINE, nLineNum, 0)
   Dim nEndPos   As Long = SciMsg( pSci, SCI_GETLINEENDPOSITION, nLineNum, 0)
   SciMsg( pSci, SCI_SETSELECTIONSTART, nStartPos, 0)
   SciMsg( pSci, SCI_SETSELECTIONEND, nEndPos, 0)
   Function = 0
End Function

''
''
Function clsDocument.GetLine( ByVal nLine As Long) As String
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   Dim nLen   As Long
   Dim buffer As String
   nLen = SciMsg( pSci, SCI_LINELENGTH, nLine , 0)
   If nLen < 1 Then Exit Function
   buffer = Space(nLen)
   SciMsg( pSci, SCI_GETLINE, nLine, Cast(LPARAM, Strptr(buffer)))
   Function = Rtrim(buffer, Any Chr(13,10,0))
End Function

''
''
Function clsDocument.SetLine( ByVal nLineNum As Long, byval sText as string) As long
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   Dim nStartPos As Long = SciMsg( pSci, SCI_POSITIONFROMLINE, nLineNum, 0)
   Dim nEndPos   As Long = SciMsg( pSci, SCI_GETLINEENDPOSITION, nLineNum, 0)
   SciMsg( pSci, SCI_SETTARGETSTART, nStartPos, 0)
   SciMsg( pSci, SCI_SETTARGETEND, nEndPos, 0)
   SciMsg( pSci, SCI_REPLACETARGET, Len(sText), Cast(LPARAM, Strptr(sText))) 
   function = 0
End Function

''
''
Function clsDocument.GetSelText() As String
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   Dim nLen As Long
   Dim buffer As String
   nLen = SciMsg( pSci, SCI_GETSELTEXT, 0, 0)
   If nLen < 1 Then Exit Function
   buffer = Space(nLen)
   SciMsg( pSci, SCI_GETSELTEXT, 0, Cast(LPARAM, Strptr(buffer)))
   Function = Trim(buffer, Chr(0))
End Function

''
''
Function clsDocument.GetText() As String
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   Dim nLen   As Long
   Dim buffer As String
   nLen = SciMsg( pSci, SCI_GETLENGTH, 0 , 0)
   If nLen < 1 Then Exit Function
   buffer = Space(nLen+1)
   SciMsg( pSci, SCI_GETTEXT, nLen+1, Cast(LPARAM, Strptr(buffer)) )
   Function = Trim(buffer, Chr(0))
End Function

''
''
Function clsDocument.SetText( ByRef sText As Const String ) As Long 
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   SciMsg( pSci, SCI_SETTEXT, 0, Cast(LPARAM, Strptr(sText)) )
   SciMsg( pSci, SCI_COLOURISE, 0, -1 )
   Function = 0
End Function


''
''
Function clsDocument.AppendText( ByRef sText As Const String ) As Long 
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   SciMsg( pSci, SCI_APPENDTEXT, len(sText), Cast(LPARAM, Strptr(sText)) )
   SciMsg( pSci, SCI_COLOURISE, 0, -1 )
   Function = 0
End Function


''
''  IsFunctionLine
''
''  Determine if incoming line number is the start of a valid sub/function.
''  By default we assume every line will be a valid function in order to
''  make the comparisons easier. The value returned is the position 
''  immediately after the declaration (0 if not found). This is useful
''  because this function can then also be used for determining the 
''  actual sub/function name rather than having to code a separate function
''  to do the same comparisons.
''  
function clsDocument.IsFunctionLine( byval lineNum as long ) as long
   
   ' default that no function declaration found
   dim as long position = 0

   ' trim and remove double spaces and replace TABs with single space
   dim as string lineText = this.GetLine( lineNum )
   if len( lineText ) < 4 then exit function
   lineText = ltrim(ucase( AfxStrShrink(lineText, chr(32,9)) ))
   if len( lineText ) < 4 then exit function
      
   If Left( lineText, 9 ) = "FUNCTION " then
      ' bypass any line with = that is a "FUNCTION = " line.
      dim as string subText = ltrim( mid(lineText, 9) )
      if left( subText, 1 ) <> "=" then position = 10
   elseif Left(lineText, 17) = "PRIVATE FUNCTION " then
      position = 18   
   elseif Left(lineText, 16) = "PUBLIC FUNCTION " then
      position = 17
   elseif Left(lineText, 4)  = "SUB " then
      position = 5
   elseif Left(lineText, 12) = "PRIVATE SUB " Then
      position = 13
   elseif Left(lineText, 11) = "PUBLIC SUB " then
      position = 12
   ElseIf Left(lineText, 9)  = "PROPERTY " Then
      ' bypass any line with = that is a "PROPERTY = " line.
      dim as string subText = ltrim( mid(lineText, 9) )
      if left( subText, 1 ) <> "=" then position = 10
   ElseIf Left(lineText, 17) = "PRIVATE PROPERTY " Then
      position = 18
   ElseIf Left(lineText, 12) = "CONSTRUCTOR " Then
      position = 13
   ElseIf Left(lineText, 11) = "DESTRUCTOR " Then
      position = 12
      
   ' If we encounter one of these END statements then our cursor must 
   ' be positioned between functions (ie we're at the module level).   
   ElseIf Left(lineText, 12) = "END FUNCTION" Then
      position = -1
   ElseIf Left(lineText, 7)  = "END SUB" Then
      position = -1
   ElseIf Left(lineText, 15) = "END CONSTRUCTOR" Then
      position = -1
   ElseIf Left(lineText, 14) = "END DESTRUCTOR" Then
      position = -1
   ElseIf Left(lineText, 12) = "END PROPERTY" Then
      position = -1

   End If

   function = position
End Function


' ========================================================================================
' Determine the sub/function name based on the current editing position 
' within the file. This is needed by CodeTips (DereferenceLine) and for code navigation
' in the editor to set the Functions ComboBox to the correct item.
' ========================================================================================
Function clsDocument.GetCurrentFunctionName( byref sFunctionName as string, _
                                             byref nGetSet as ClassProperty _
                                             ) As long
    
   dim as string lineText
   dim as string funcName 
   dim as string funcParams
   
   dim as long curLine
   dim as long position
   
   ' search up the file until we find the start of a sub/function or start of file.
   curLine = this.GetCurrentLineNumber
   nGetSet = ClassProperty.None
   
   for i as long = curLine to 0 step -1
      position = this.IsFunctionLine( i )
      
      if position = -1 then 
         ' We found an "END SUB", "END FUNCTION", "END PROPERTY" line. Keep looking for 
         ' the very next instance of SUB/FUNCTION and then break out of loop
         ' with that function name.
         if i = this.GetCurrentLineNumber then 
            continue for
         else
            exit for
         end if
            
      elseif position = 0 then
         ' Just a regular line... keep looking...
      
      elseif position > 0 then
         ' We found a valid SUB/FUNCTION line so process it
         ' trim and remove double spaces and replace TABs with single space
         lineText = ltrim(ucase( AfxStrShrink(this.GetLine(i), chr(32,9)) ))
         funcName = ltrim(mid(lineText, position))
         funcName = AfxStrParseAny( funcName, 1, " (" )
         
         ' If this is a Property then we need to differentiate between a Get/Set
         If ( Left(lineText, 9)  = "PROPERTY " ) orelse _
            ( Left(lineText, 17) = "PRIVATE PROPERTY " ) Then
            ' if funcParams exist then this must be a Set property. Need to sanitize the
            ' property parameters first. Need to get the starting ( and the ending ) and
            ' evaluate the text between it. Not as easy as using Parse because there could
            ' be embedded array() parameters and the property could end in something like 
            ' AS LONG, etc.
            dim as String st
            dim as string sFuncParams = lineText
            dim as long p1, p2
            p1 = instr( sFuncParams, "(" )
            p2 = InstrRev( sFuncParams, ")" )
            if ( p1 = 0 ) orelse (p2 = 0 ) then
               st = ""
            elseif p2 > p1 then
               st = mid( sFuncParams, p1, p2 - p1 )
            end if
            st = trim( st, any "( )" )
            if len( st ) then
               nGetSet = ClassProperty.Setter
            else   
               nGetSet = ClassProperty.Getter
            end if

         elseIf ( Left(lineText, 12)  = "CONSTRUCTOR " ) then
            nGetSet = ClassProperty.ctor
            
         elseIf ( Left(lineText, 11)  = "DESTRUCTOR " ) then
            nGetSet = ClassProperty.dtor
         end if   

         exit for
      end if
   NEXT

   sFunctionName = funcName
   nGetSet = nGetSet
   
   function = 0
End Function


''
''  GotoNextFunction
''
''  Go to the next sub/function in the document
''
function clsDocument.GotoNextFunction() as long
   
   dim as long curLine  = this.GetCurrentLineNumber
   dim as long maxLines = this.GetLineCount - 1
   dim as long newLine  = curLine
   
   for i as long = curLine + 1 to maxLines
      if this.IsFunctionLine( i ) > 0 then
         newLine = i
         exit for
      end if
   next

   ' if we have found a new line then reposition to that line
   if newLine <> curLine then
      dim as any ptr pSci = this.GetActiveScintillaPtr()
      SciMsg( pSci, SCI_SETFIRSTVISIBLELINE, newLine - 5, 0) 
      SciMsg( pSci, SCI_GOTOLINE, newLine, 0) 
   end if
   
   function = 0
end function


''
''  GotoPrevFunction
''
''  Go to the previous sub/function in the document
''
function clsDocument.GotoPrevFunction() as long
   
   dim as long curLine  = this.GetCurrentLineNumber
   dim as long newLine  = curLine
   
   for i as long = curLine - 1 to 0 step -1
      if this.IsFunctionLine( i ) > 0 then
         newLine = i
         exit for
      end if
   next

   ' if we have found a new line then reposition to that line
   if newLine <> curLine then
      dim as any ptr pSci = this.GetActiveScintillaPtr()
      SciMsg( pSci, SCI_SETFIRSTVISIBLELINE, newLine - 5, 0) 
      SciMsg( pSci, SCI_GOTOLINE, newLine, 0) 
   end if
   
   function = 0
end function


''
''
Function clsDocument.GetSelectedLineRange( ByRef startLine As Long, _
                                           ByRef endLine   As Long, _
                                           ByRef startPos  As Long, _
                                           ByRef endPos    As Long _
                                           ) As Long 
   
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   startPos  = SciMsg( pSci, SCI_GETSELECTIONSTART, 0, 0) 
   endPos    = SciMsg( pSci, SCI_GETSELECTIONEND, 0, 0) 
   startLine = SciMsg( pSci, SCI_LINEFROMPOSITION, startPos, 0) 
   endLine   = SciMsg( pSci, SCI_LINEFROMPOSITION, endPos, 0) 

   Dim nCol As Long = SciMsg( pSci, SCI_GETCOLUMN, endPos, 0)
   If (nCol = 0) And (endLine > startLine) Then endLine = endLine - 1

   Function = 0
End Function


''
''
Function clsDocument.BlockComment( ByVal flagBlock As BOOLEAN ) As Long

   Dim i           As Long        ' Loop counter
   Dim startPos    As Long        ' Starting position
   Dim endPos      As Long        ' Ending position
   Dim startLine   As Long        ' Starting line
   Dim endLine     As Long        ' Ending line
   Dim nPos        As Long        ' Position
   Dim strText     As String      ' Portion of the line to replace
   Dim nCount      As Long        ' Number of "'" added or removed
   
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   ' flagBlock = True for Blocking comment, False for UnBlocking comment
   this.GetSelectedLineRange( startLine, endLine, startPos, endPos )
   SciMsg( pSci, SCI_GOTOLINE, startLine, 0) 
   
   gApp.SuppressNotify = True
   For i = startLine To endLine
      strText = this.GetLine(i)   
      If flagBlock = False Then   ' unblock comment
         If Left(strText, 1) <> "'" Then
            Continue For
         Else   
            strText = Iif(Len(strText) > 1, Mid(strText, 2), " ")
         End If   
      Else
         If Len(rtrim(strText)) Then
            strText = "'" & strText
         end if   
      End If
      If Len(strText) Then
         nPos = SciMsg( pSci, SCI_POSITIONFROMLINE, i, 0)   ' starting position of the line
         SciMsg( pSci, SCI_SETTARGETSTART, nPos, 0)
         SciMsg( pSci, SCI_SETTARGETEND, nPos + Len(strText) + Iif(flagBlock, -1, 1), 0)
         SciMsg( pSci, SCI_REPLACETARGET, Len(strText), Cast(LPARAM, Strptr(strText))) 
         nCount += 1
      End If
   Next
   gApp.SuppressNotify = False

   If startPos <> endPos Then
      SciMsg( pSci, SCI_SETSELECTIONSTART, startPos, 0)
      SciMsg( pSci, SCI_SETSELECTIONEND, endPos + Iif(flagBlock, nCount, -nCount), 0)
   Else
      SciMsg( pSci, SCI_SETSELECTIONSTART, endPos + Iif(flagBlock, nCount, -nCount), 0)
      SciMsg( pSci, SCI_SETSELECTIONEND, endPos + Iif(flagBlock, nCount, -nCount), 0)
   End If
   
   Function = 0
End Function


''
''
Function clsDocument.CurrentLineUp() As Long
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   Dim currentLine As Long = this.GetCurrentLineNumber()
   If (currentLine <> 0) Then
      SciMsg( pSci, SCI_BEGINUNDOACTION, 0, 0)
      currentLine = currentLine -1
      SciMsg( pSci, SCI_LINETRANSPOSE, 0, 0)
      SciMsg( pSci, SCI_GOTOLINE, currentLine, 0)
      SciMsg( pSci, SCI_ENDUNDOACTION, 0, 0)
   End If
   Function = 0
End Function


''
''
Function clsDocument.GetLineCount() As Long
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   function = SciMsg( pSci, SCI_GETLINECOUNT, 0, 0)
end function


''
''
Function clsDocument.CurrentLineDown() As Long
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   Dim currentLine As Long = this.GetCurrentLineNumber()
   If currentLine <> SciMsg( pSci, SCI_GETLINECOUNT, 0, 0) - 1 Then
      SciMsg( pSci, SCI_BEGINUNDOACTION, 0, 0)
      currentLine = currentLine + 1
      SciMsg( pSci, SCI_GOTOLINE, currentLine, 0)
      SciMsg( pSci, SCI_LINETRANSPOSE, 0, 0)
      SciMsg( pSci, SCI_ENDUNDOACTION, 0, 0)
   End If
   Function = 0
End Function
  
''
''
Function clsDocument.MoveCurrentLines( ByVal flagMoveDown As BOOLEAN ) As Long

   Dim startPos    As Long     ' Starting position
   Dim endPos      As Long     ' Ending position
   Dim startLine   As Long     ' Starting line
   Dim endLine     As Long     ' Ending line
   Dim line2swap   As Long 
   Dim nbChar      As Long 
   Dim nTextLen    As Long
   
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   this.GetSelectedLineRange( startLine, endLine, startPos, endPos )

   Dim noSel      As Long = (startLine = endLine)
   Dim nbSelLines As Long = (endLine - startLine + 1)

   If flagMoveDown Then
      nTextLen  = SciMsg( pSci, SCI_GETTEXTLENGTH, 0, 0)
      If (startLine = -1) Or (endLine >= SciMsg( pSci, SCI_LINEFROMPOSITION, nTextLen, 0)) Then Exit Function
      line2swap = endLine + 1
      If (line2swap + 1) = SciMsg( pSci, SCI_GETLINECOUNT, 0, 0) Then nbChar += 2    ' cr/lf
   Else
      If (startLine <= 0) Then Exit Function
      line2swap = startLine - 1
   End If
   
   nbChar = nbChar + SciMsg( pSci, SCI_LINELENGTH, line2swap, 0)
   
   SciMsg( pSci, SCI_BEGINUNDOACTION, 0, 0)
   SciMsg( pSci, SCI_GOTOLINE, line2swap, 0)

   gApp.SuppressNotify = True
   For i As Long = 0 To nbSelLines - 1
      If flagMoveDown Then
         this.CurrentLineUp()
      Else
         this.CurrentLineDown()
      End If   
   Next
   gApp.SuppressNotify = False
   SciMsg( pSci, SCI_ENDUNDOACTION, 0, 0)

   If flagMoveDown Then
      SciMsg( pSci, SCI_SETSELECTIONSTART, startPos + nbChar, 0)
      SciMsg( pSci, SCI_SETSELECTIONEND, Iif(noSel, startPos + nbChar, endPos + nbChar), 0)
   Else
      SciMsg( pSci, SCI_SETSELECTIONSTART, startPos - nbChar, 0)
      SciMsg( pSci, SCI_SETSELECTIONEND, Iif(noSel, startPos - nbChar, endPos - nbChar), 0)
   End If
   
   Function = 0
End Function

''
''
Function clsDocument.ToggleBookmark( ByVal nLine As Long ) As Long
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   Dim fMark As Long  ' must be a 32 bit value
   fMark = SciMsg( pSci, SCI_MARKERGET, nLine, 0) 
   If Bit(fMark, 0) = -1 Then
      SciMsg( pSci, SCI_MARKERDELETE, nLine, 0)
   Else
      SciMsg( pSci, SCI_MARKERADD, nLine, 0)
   End If
   Function = 0
End Function

''
''
Function clsDocument.NextBookmark() As Long 
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   Dim fMark As Long  ' 32 bit value
   Dim nLine As Long = this.GetCurrentLineNumber() + 1
   fMark = BitSet(fMark, 0)
   nLine = SciMsg( pSci, SCI_MARKERNEXT, nLine, fMark)
   If nLine > -1 Then
      SciMsg( pSci, SCI_GOTOLINE, nLine, 0)
   Else
      nLine = SciMsg( pSci, SCI_MARKERNEXT, nLine, fMark)
      If nLine > -1 Then
         SciMsg( pSci, SCI_GOTOLINE, nLine, 0)
      End If
   End If
   Function = 0 
End Function

''
''
Function clsDocument.PrevBookmark() As Long 
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   Dim fMark  As Long  ' 32 bit value
   Dim nLine  As Long = this.GetCurrentLineNumber() - 1
   Dim nLines As Long = SciMsg( pSci, SCI_GETLINECOUNT, 0, 0) - 1
   fMark = BitSet(fMark, 0)
   nLine = SciMsg( pSci, SCI_MARKERPREVIOUS, nLine, fMark)
   If nLine > -1 Then
      SciMsg( pSci, SCI_GOTOLINE, nLine, 0)
   Else
      nLine = SciMsg( pSci, SCI_MARKERPREVIOUS, nLines, fMark)
      If nLine > -1 Then
         SciMsg( pSci, SCI_GOTOLINE, nLine, 0)
      End If
   End If
   Function = 0 
End Function

''
''   
Function clsDocument.FoldToggle( ByVal nLine As Long ) As Long
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   Dim nFoldLevel As Long = SciMsg( pSci, SCI_GETFOLDLEVEL, nLine, 0)

   If (nFoldLevel And SC_FOLDLEVELHEADERFLAG) = 0 Then
      ' Get the number of the head line of the procedure or function
      nLine = SciMsg( pSci, SCI_GETFOLDPARENT, nLine, 0) 
   End If
   If nLine > -1 Then
      SciMsg( pSci, SCI_TOGGLEFOLD, nLine, 0) 
      SciMsg( pSci, SCI_GOTOLINE, nLine, 0)
   End If

   Function = nLine
End Function

''
''
Function clsDocument.FoldAll() As Long

   Dim i          As Long    
   Dim nLines     As Long    
   Dim nFoldLevel As Long
   
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   ' Force the lexer to style the whole document
   SciMsg( pSci, SCI_COLOURISE, -1, 0)

   nLines = SciMsg( pSci, SCI_GETLINECOUNT, 0, 0)

   For i = 0 To nLines
      ' If we are in the head line ...
      nFoldLevel = SciMsg( pSci, SCI_GETFOLDLEVEL, i, 0)
      If (nFoldLevel And SC_FOLDLEVELNUMBERMASK) = SC_FOLDLEVELBASE Then
         If SciMsg( pSci, SCI_GETFOLDEXPANDED, i, 0) Then
            SciMsg( pSci, SCI_TOGGLEFOLD, i, 0) 
         End If
      End If
   Next

   Function = 0
End Function

''
''
Function clsDocument.UnFoldAll() As Long

   Dim i          As Long    
   Dim nLines     As Long    
   Dim nFoldLevel As Long
   
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   ' Force the lexer to style the whole document
   SciMsg( pSci, SCI_COLOURISE, -1, 0 )

   nLines = SciMsg( pSci, SCI_GETLINECOUNT, 0, 0)

   For i = 0 To nLines
      ' If we are in the head line ...
      nFoldLevel = SciMsg( pSci, SCI_GETFOLDLEVEL, i, 0)
      If (nFoldLevel And SC_FOLDLEVELNUMBERMASK) = SC_FOLDLEVELBASE Then
         If SciMsg( pSci, SCI_GETFOLDEXPANDED, i, 0) = 0 Then
            SciMsg( pSci, SCI_TOGGLEFOLD, i, 0) 
         End If
      End If
   Next

   Function = 0
End Function

''
''
Function clsDocument.FoldToggleOnwards( ByVal nLine As Long) As Long

   ' Toggles the curent fold point and, if it is a base level, all the base level fold points below.
   Dim i          As Long    
   Dim nLines     As Long    
   Dim nFoldLevel As Long
   Dim FoldState  As Long
   
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   ' Force the lexer to style the whole document
   SciMsg( pSci, SCI_COLOURISE, -1, 0 )

   nLines = SciMsg( pSci, SCI_GETLINECOUNT, 0, 0)

   ' Toggle the first sub or function
   nLine = this.FoldToggle(nLine)
   
   ' Determine whether the fold is expanded or not
   FoldState = SciMsg( pSci, SCI_GETFOLDEXPANDED, nLine, 0)

   For i = nLine To nLines
      ' If we are in the head line
      nFoldLevel = SciMsg( pSci, SCI_GETFOLDLEVEL, i, 0)
      If (nFoldLevel And SC_FOLDLEVELNUMBERMASK) = SC_FOLDLEVELBASE Then
         ' If the state is different ...
         If SciMsg( pSci, SCI_GETFOLDEXPANDED, i, 0) <> FoldState Then
            SciMsg( pSci, SCI_TOGGLEFOLD, i, 0) 
         End If
      End If
   Next

   Function = 0
End Function

''
''
Function clsDocument.ConvertEOL( ByVal nMode As Long) As Long
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   Function = SciMsg( pSci, SCI_CONVERTEOLS, nMode, 0)
End Function

''
''
Function clsDocument.TabsToSpaces() As Long

   dim as any ptr pSci = this.GetActiveScintillaPtr()
   Dim As Long i, n, nLen, nLines, TabSize, nLineNumber 
   Dim As String strText, strBuffer

   ' Get the current line
   nLineNumber = this.GetCurrentLineNumber
   ' Get the tab size
   TabSize = SciMsg( pSci, SCI_GETTABWIDTH, 0, 0)
   If TabSize < 1 Then Exit Function
   ' Get the length of the text
   nLen  = SciMsg( pSci, SCI_GETTEXTLENGTH, 0, 0) 
   ' Get the number of lines
   nLines = SciMsg( pSci, SCI_GETLINECOUNT, 0, 0)
   ' Remove tabs, line by line
   For i = 0 To nLines - 1
      strText = this.GetLine(i)
      n = 1
      Do
         n = Instr(n, strText, Chr(9))
         If n > 0 Then 
            strText = Left(strText, n - 1) & Space(TabSize) & Mid(strText, n + 1)
            n += 1
         End If   
      Loop Until n = 0
      strBuffer = strBuffer & strText & Chr(13,10)
   Next
   ' Set the new text
   this.SetText(strBuffer)
   ' Set the caret position
   SciMsg( pSci, SCI_GOTOLINE, nLineNumber, 0) 

   Function = 0
End Function


''
''
Function clsDocument.IsMultiLineSelection() As boolean
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   dim as long startPos, endPos, startLine, endLine
   startPos  = SciMsg( pSci, SCI_GETSELECTIONSTART, 0, 0) 
   endPos    = SciMsg( pSci, SCI_GETSELECTIONEND, 0, 0) 
   startLine = SciMsg( pSci, SCI_LINEFROMPOSITION, startPos, 0) 
   endLine   = SciMsg( pSci, SCI_LINEFROMPOSITION, endPos, 0) 
   if endLine <> startLine THEN return true
end function


''
''
Function clsDocument.ApplyProperties() As Long

   Dim nCount      As Long 
   Dim i           As Long
   Dim nPixels     As Long  
   Dim bitsNeeded  As Long 
   Dim wFileExt    As WString * MAX_PATH
   Dim strFontName As String 
   Dim nFontSize   As Long  
   Dim nFontCase   As Long 
   Dim rxRatio     As Single = 1
   Dim ryRatio     As Single = 1
   dim pTheme      as TYPE_THEMES ptr
   
   ' Determine the pWindow parent of the Scintilla window in order
   ' to ensure that DPI ratios are correctly used.
   Dim pWindow As CWindow Ptr = AfxCWindowOwnerPtr(this.hWindow(0))
   If pWindow Then
      rxRatio = pWindow->rxRatio
      ryRatio = pWindow->ryRatio
   End If   
   
   ' If this is a read-only file then set the flag so that the document can not be edited.
   if AfxIsReadOnlyFile( this.DiskFilename ) then
      SciMsg( m_pSci(0), SCI_SETREADONLY, 1, 0)
   else
      SciMsg( m_pSci(0), SCI_SETREADONLY, 0, 0)
   end if

   strFontName = Str(gConfig.EditorFontname)
   nFontSize   = Val(**gConfig.EditorFontsize)
   
   Select Case gConfig.KeywordCase
      Case 0:  nFontCase = SC_CASE_LOWER
      Case 1:  nFontCase = SC_CASE_UPPER
      Case 2:  nFontCase = SC_CASE_CAMEL   
      Case 3:  nFontCase = SC_CASE_MIXED    ' original case
   End Select            
   
   If m_pSci(0) = 0 Then Exit Function

   pTheme = gConfig.GetThemePtr
   

   ' Must apply all settings/styles to each Scintilla split window to ensure that
   ' they all appear and act the same.
   for i as long = lbound(m_pSci) to ubound(m_pSci)
  
      SciMsg( m_pSci(i), SCI_STYLESETFONT, STYLE_DEFAULT, Cast(LPARAM, Strptr(strFontName)) )
      SciMsg( m_pSci(i), SCI_STYLESETSIZE, STYLE_DEFAULT, nFontSize )
      SciMsg( m_pSci(i), SCI_STYLESETCHARACTERSET, STYLE_DEFAULT, GetFontCharSetID(gConfig.EditorFontCharset) )

      SciMsg( m_pSci(i), SCI_STYLESETFORE, STYLE_DEFAULT, pTheme->colors(CLR_TEXT).nFg)
      SciMsg( m_pSci(i), SCI_STYLESETBACK, STYLE_DEFAULT, pTheme->colors(CLR_WINDOW).nFg)
      SciMsg( m_pSci(i), SCI_STYLESETBOLD, STYLE_DEFAULT, pTheme->colors(CLR_TEXT).bFontBold )
      SciMsg( m_pSci(i), SCI_STYLESETITALIC, STYLE_DEFAULT, pTheme->colors(CLR_TEXT).bFontItalic )
      SciMsg( m_pSci(i), SCI_STYLESETUNDERLINE, STYLE_DEFAULT, pTheme->colors(CLR_TEXT).bFontUnderline )
      SciMsg( m_pSci(i), SCI_STYLECLEARALL, 0, 0 )  ' Copies global style to all others
          
      ' Set the style for the AutoComplete popup list
      SciMsg( m_pSci(i), SCI_STYLESETFONT, STYLE_AUTOCOMPLETE, Cast(LPARAM, pWindow->DefaultFontName) )
      SciMsg( m_pSci(i), SCI_STYLESETSIZE, STYLE_AUTOCOMPLETE, pWindow->DefaultFontSize)
      SciMsg( m_pSci(i), SCI_STYLESETCHARACTERSET, STYLE_AUTOCOMPLETE, GetFontCharSetID(gConfig.EditorFontCharset) )

      ''
      ''  MARGIN 0: Line Numbering (defaults to width 0)
      nPixels = SciMsg( m_pSci(i), SCI_TEXTWIDTH, 0, Cast(LPARAM, @"_99999"))
      SciMsg( m_pSci(i), SCI_SETMARGINTYPEN, 0, SC_MARGIN_NUMBER )
      SciMsg( m_pSci(i), SCI_STYLESETFORE, STYLE_LINENUMBER, pTheme->colors(CLR_LINENUMBERS).nFg )
      SciMsg( m_pSci(i), SCI_STYLESETBACK, STYLE_LINENUMBER, pTheme->colors(CLR_LINENUMBERS).nBg )
      SciMsg( m_pSci(i), SCI_SETMARGINWIDTHN, 0, Iif(gConfig.LineNumbering, nPixels, 0) )
      SciMsg( m_pSci(i), SCI_STYLESETBOLD, STYLE_LINENUMBER, pTheme->colors(CLR_LINENUMBERS).bFontBold )
      SciMsg( m_pSci(i), SCI_STYLESETITALIC, STYLE_LINENUMBER, pTheme->colors(CLR_LINENUMBERS).bFontItalic )
      SciMsg( m_pSci(i), SCI_STYLESETUNDERLINE, STYLE_LINENUMBER, pTheme->colors(CLR_LINENUMBERS).bFontUnderline )
   
      ''
      ''  MARGIN 1: Non-Folding symbols (defaults to width 16) (Bookmark symbol, etc) (will be same color as line numbering)
      SciMsg( m_pSci(i), SCI_SETMARGINTYPEN, 1, SC_MARGIN_TEXT )
      SciMsg( m_pSci(i), SCI_SETMARGINSENSITIVEN, 1, 1 )
      SciMsg( m_pSci(i), SCI_SETMARGINWIDTHN, 1, Iif(gConfig.LeftMargin, 16 * rxRatio, 0) )
                 
      ''
      ''  MARGIN 2: Folding symbols (defaults to width 0)
      SciMsg( m_pSci(i), SCI_SETMARGINTYPEN, 2, SC_MARGIN_SYMBOL )
      SciMsg( m_pSci(i), SCI_SETMARGINMASKN, 2, SC_MASK_FOLDERS )
      SciMsg( m_pSci(i), SCI_SETFOLDMARGINCOLOUR, CTRUE, pTheme->colors(CLR_FOLDMARGIN).nFg )
      SciMsg( m_pSci(i), SCI_SETFOLDMARGINHICOLOUR, CTRUE, pTheme->colors(CLR_FOLDMARGIN).nFg )
      SciMsg( m_pSci(i), SCI_SETMARGINSENSITIVEN, 2, 1 )
      SciMsg( m_pSci(i), SCI_SETMARGINWIDTHN, 2, Iif(gConfig.FoldMargin, 16 * rxRatio, 0) )

      ''
      ''  MARGIN 3: Small margin to offset left margins from actual text (4 pixels)
      SciMsg( m_pSci(i), SCI_SETMARGINTYPEN, 3, SC_MARGIN_TEXT )
      SciMsg( m_pSci(i), SCI_SETMARGINWIDTHN, 3, 4 * rxRatio )
      SciMsg( m_pSci(i), SCI_SETMARGINTYPEN, 3, SC_MARGIN_FORE )
      SciMsg( m_pSci(i), SCI_SETMARGINTYPEN, 3, SC_MARGIN_BACK )
             
      '' 
      ''  FONT QUALITY
      ' Commented out v1.7.4 to see if bitmap .fon fonts will now work
      'SciMsg( m_pSci(i), SCI_SETFONTQUALITY, SC_EFF_QUALITY_DEFAULT, 0 )
      'SC_EFF_QUALITY_DEFAULT (backward compatible), SC_EFF_QUALITY_NON_ANTIALIASED, 
      'SC_EFF_QUALITY_ANTIALIASED, SC_EFF_QUALITY_LCD_OPTIMIZED
      
      ' Commented out v1.7.4 to see if bitmap .fon fonts will now work
      'SciMsg( m_pSci(i), SCI_SETTECHNOLOGY, SC_TECHNOLOGY_DIRECTWRITE, 0 )
      'SciMsg( m_pSci(i), SCI_SETBUFFEREDDRAW, 0, 0 ) ' turn off b/c we are using Vista+ Direct Draw
      
      'SC_TECHNOLOGY_DEFAULT (0). 
      'On Windows Vista or later, 
      'SC_TECHNOLOGY_DIRECTWRITE (1), 
      'SC_TECHNOLOGY_DIRECTWRITERETAIN (2), or 
      'SC_TECHNOLOGY_DIRECTWRITEDC (3) 
      
      ''
      ''  CONFINE CARET TO TEXT
      If gConfig.ConfineCaret Then
         SciMsg( m_pSci(i), SCI_SETVIRTUALSPACEOPTIONS, SCVS_RECTANGULARSELECTION, 0 )
      Else
         SciMsg( m_pSci(i), SCI_SETVIRTUALSPACEOPTIONS, SCVS_RECTANGULARSELECTION Or SCVS_USERACCESSIBLE, 0 )
      End If
   
      ''
      ''  TABS AS SPACES
      If gConfig.TabIndentSpaces Then
         SciMsg( m_pSci(i), SCI_SETUSETABS, False, 0 )
      Else
         SciMsg( m_pSci(i), SCI_SETUSETABS, CTRUE, 0 )
      End If

      ''
      ''  SELECTIONS FILL ENTIRE SCREEN SPACE
      SciMsg( m_pSci(i), SCI_SETSELEOLFILLED, CTRUE, 0 )
          
      ''
      ''  TAB WIDTH 
      SciMsg( m_pSci(i), SCI_SETTABWIDTH, Val(**gConfig.TabSize), 0 )
      SciMsg( m_pSci(i), SCI_SETINDENT, Val(**gConfig.TabSize), 0 )
 
      ''
      ''  INDENTATION GUIDES
      If gConfig.IndentGuides Then
         SciMsg( m_pSci(i), SCI_SETINDENTATIONGUIDES, CTRUE, 0)
      Else
         SciMsg( m_pSci(i), SCI_SETINDENTATIONGUIDES, False, 0)
      End If
      SciMsg( m_pSci(i), SCI_SETSELFORE, STYLE_INDENTGUIDE, pTheme->colors(CLR_INDENTGUIDES).nFg )
      SciMsg( m_pSci(i), SCI_SETSELBACK, STYLE_INDENTGUIDE, pTheme->colors(CLR_FOLDMARGIN).nBg )
      
      ''
      ''  CARET
      SciMsg( m_pSci(i), SCI_SETCARETFORE, pTheme->colors(CLR_CARET).nFg, 0 )
      SciMsg( m_pSci(i), SCI_SETCARETWIDTH, 2, 0 )       ' 2 pixels

      ''
      ''  SHOW CARET LINE
      If gConfig.HighlightCurrentLine Then
         SciMsg( m_pSci(i), SCI_SETCARETLINEVISIBLE, CTRUE, 0 )
      Else
         SciMsg( m_pSci(i), SCI_SETCARETLINEVISIBLE, False, 0 )
      End If

      ''
      ''  CARET LINE COLOR
      SciMsg( m_pSci(i), SCI_SETCARETLINEBACK, pTheme->colors(CLR_HIGHLIGHTED).nFg, 0 )
      
      ''
      ''  SELECTION COLORS
      SciMsg( m_pSci(i), SCI_SETSELFORE, CTRUE, pTheme->colors(CLR_SELECTION).nFg)
      SciMsg( m_pSci(i), SCI_SETSELBACK, CTRUE, pTheme->colors(CLR_SELECTION).nBg)
   
      ''
      ''  MULTIPLE SELECTIONS
      SciMsg( m_pSci(i), SCI_SETMULTIPLESELECTION, FALSE, 0 ) 
               
      ''
      ''  ALWAYS KEEP THE CARET LINE VISIBLE
      SciMsg( m_pSci(i), SCI_SETCARETLINEVISIBLEALWAYS, CTRUE, 0 )
      
      ''
      ''  DISABLE RIGHT CLICK POPUP MENU
      SciMsg( m_pSci(i), SCI_USEPOPUP, False, 0 )

      ''
      ''  IDENTIFY CHARACTERS TO BE USED IN WORDS
      SciMsg( m_pSci(i), SCI_SETWORDCHARS, 0, Cast(LPARAM, @"~_:\abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789") )

      ''  UNICODE (UTF-8 encoding)
      IF this.FileEncoding = FILE_ENCODING_ANSI THEN
         SciMsg( m_pSci(i), SCI_SETCODEPAGE, 0, 0 )
      ELSE
         ' UTF8 or UTF16 would have been converted to UTF8 in order to display in the editor.
         SciMsg( m_pSci(i), SCI_SETCODEPAGE, SC_CP_UTF8, 0 )
      END IF
   
      '' RIGHT EDGE COLUMN
      SciMsg( m_pSci(i), SCI_SETEDGEMODE, iif(gConfig.RightEdge, EDGE_LINE, EDGE_NONE), 0 )
      SciMsg( m_pSci(i), SCI_SETEDGECOLUMN, val(**gConfig.RightEdgePosition), 0 )
      'SciMsg( m_pSci, SCI_SETEDGECOLOR, iif(gConfig.RightEdge, EDGE_LINE, EDGE_NONE) )

      ''
      ''  OTHER
      SciMsg( m_pSci(i), SCI_SETADDITIONALSELECTIONTYPING, True, 0 )
   

      ''
      ''  APPLY ALL LANGUAGE SPECIFIC SYNTAX COLORING
      wFileExt = AfxStrPathname( "EXTN", this.DiskFilename )
      wFileExt = Ucase(wFileExt)

      If cbool(wFileExt = ".BAS") OrElse cbool(wFileExt = ".INC") _
            OrElse cbool(wFileExt = ".BI") OrElse (this.IsNewFlag = True) _
            orelse cbool(wFileExt = ".FBTPL") Then

         SciMsg( m_pSci(i), SCI_SETLEXER, SCLEX_VB, 0 )
         bitsNeeded = SciMsg( m_pSci(i), SCI_GETSTYLEBITSNEEDED, 0, 0)
         SciMsg( m_pSci(i), SCI_SETSTYLEBITS, bitsNeeded, 0 )

         ' Set FreeBASIC Keywords
         If Len(gConfig.FBKeywords) Then
            SciMsg( m_pSci(i), SCI_SETKEYWORDS, 0, Cast(LPARAM, Strptr(gConfig.FBKeywords)) )
         End If
         
         If gConfig.SyntaxHighlighting Then
       
            ' Set the Multiline Comments style
            SciMsg( m_pSci(i), SCI_STYLESETFORE, SCE_B_MULTILINECOMMENT, pTheme->colors(CLR_COMMENTS).nFg)
            SciMsg( m_pSci(i), SCI_STYLESETBACK, SCE_B_MULTILINECOMMENT, pTheme->colors(CLR_COMMENTS).nBg)
            SciMsg( m_pSci(i), SCI_STYLESETBOLD, SCE_B_MULTILINECOMMENT, pTheme->colors(CLR_COMMENTS).bFontBold )
            SciMsg( m_pSci(i), SCI_STYLESETITALIC, SCE_B_MULTILINECOMMENT, pTheme->colors(CLR_COMMENTS).bFontItalic )
            SciMsg( m_pSci(i), SCI_STYLESETUNDERLINE, SCE_B_MULTILINECOMMENT, pTheme->colors(CLR_COMMENTS).bFontUnderline )

            ' Set the Comments style
            SciMsg( m_pSci(i), SCI_STYLESETFORE, SCE_B_COMMENT, pTheme->colors(CLR_COMMENTS).nFg)
            SciMsg( m_pSci(i), SCI_STYLESETBACK, SCE_B_COMMENT, pTheme->colors(CLR_COMMENTS).nBg)
            SciMsg( m_pSci(i), SCI_STYLESETBOLD, SCE_B_COMMENT, pTheme->colors(CLR_COMMENTS).bFontBold )
            SciMsg( m_pSci(i), SCI_STYLESETITALIC, SCE_B_COMMENT, pTheme->colors(CLR_COMMENTS).bFontItalic )
            SciMsg( m_pSci(i), SCI_STYLESETUNDERLINE, SCE_B_COMMENT, pTheme->colors(CLR_COMMENTS).bFontUnderline )

            ' Set the Keywords style
            SciMsg( m_pSci(i), SCI_STYLESETFORE, SCE_B_KEYWORD, pTheme->colors(CLR_KEYWORD).nFg)
            SciMsg( m_pSci(i), SCI_STYLESETBACK, SCE_B_KEYWORD, pTheme->colors(CLR_KEYWORD).nBg)
            SciMsg( m_pSci(i), SCI_STYLESETCASE, SCE_B_KEYWORD, nFontCase )
            SciMsg( m_pSci(i), SCI_STYLESETBOLD, SCE_B_KEYWORD, pTheme->colors(CLR_KEYWORD).bFontBold )
            SciMsg( m_pSci(i), SCI_STYLESETITALIC, SCE_B_KEYWORD, pTheme->colors(CLR_KEYWORD).bFontItalic )
            SciMsg( m_pSci(i), SCI_STYLESETUNDERLINE, SCE_B_KEYWORD, pTheme->colors(CLR_KEYWORD).bFontUnderline )

            ' Set the Identifiers style
            SciMsg( m_pSci(i), SCI_STYLESETFORE, SCE_B_IDENTIFIER, pTheme->colors(CLR_TEXT).nFg)
            SciMsg( m_pSci(i), SCI_STYLESETBACK, SCE_B_IDENTIFIER, pTheme->colors(CLR_TEXT).nBg)

            ' Set the Numbers style
            SciMsg( m_pSci(i), SCI_STYLESETFORE, SCE_B_NUMBER, pTheme->colors(CLR_TEXT).nFg)
            SciMsg( m_pSci(i), SCI_STYLESETBACK, SCE_B_NUMBER, pTheme->colors(CLR_TEXT).nBg)

            ' Set the Operators style
            SciMsg( m_pSci(i), SCI_STYLESETFORE, SCE_B_OPERATOR, pTheme->colors(CLR_OPERATORS).nFg)
            SciMsg( m_pSci(i), SCI_STYLESETBACK, SCE_B_OPERATOR, pTheme->colors(CLR_OPERATORS).nBg)
            SciMsg( m_pSci(i), SCI_STYLESETBOLD, SCE_B_OPERATOR, pTheme->colors(CLR_OPERATORS).bFontBold )
            SciMsg( m_pSci(i), SCI_STYLESETITALIC, SCE_B_OPERATOR, pTheme->colors(CLR_OPERATORS).bFontItalic )
            SciMsg( m_pSci(i), SCI_STYLESETUNDERLINE, SCE_B_OPERATOR, pTheme->colors(CLR_OPERATORS).bFontUnderline )

            ' Set the Preprocessor style
            SciMsg( m_pSci(i), SCI_STYLESETFORE, SCE_B_PREPROCESSOR, pTheme->colors(CLR_PREPROCESSOR).nFg)
            SciMsg( m_pSci(i), SCI_STYLESETBACK, SCE_B_PREPROCESSOR, pTheme->colors(CLR_PREPROCESSOR).nBg)
            SciMsg( m_pSci(i), SCI_STYLESETBOLD, SCE_B_PREPROCESSOR, pTheme->colors(CLR_PREPROCESSOR).bFontBold )
            SciMsg( m_pSci(i), SCI_STYLESETITALIC, SCE_B_PREPROCESSOR, pTheme->colors(CLR_PREPROCESSOR).bFontItalic )
            SciMsg( m_pSci(i), SCI_STYLESETUNDERLINE, SCE_B_PREPROCESSOR, pTheme->colors(CLR_PREPROCESSOR).bFontUnderline )

            ' Set the Strings style
            SciMsg( m_pSci(i), SCI_STYLESETFORE, SCE_B_STRING, pTheme->colors(CLR_STRINGS).nFg)
            SciMsg( m_pSci(i), SCI_STYLESETBACK, SCE_B_STRING, pTheme->colors(CLR_STRINGS).nBg)
            SciMsg( m_pSci(i), SCI_STYLESETBOLD, SCE_B_STRING, pTheme->colors(CLR_STRINGS).bFontBold )
            SciMsg( m_pSci(i), SCI_STYLESETITALIC, SCE_B_STRING, pTheme->colors(CLR_STRINGS).bFontItalic )
            SciMsg( m_pSci(i), SCI_STYLESETUNDERLINE, SCE_B_STRING, pTheme->colors(CLR_STRINGS).bFontUnderline )

         End If

      End If

      ''
      ''  CODE FOLDING
      If gConfig.FoldMargin Then
         ' Enable folding of the procedures and functions
         SciMsg( m_pSci(i), SCI_SETPROPERTY, Cast(WPARAM, @"fold"), Cast(LPARAM, @"1") )

         ' Initialize fold symbols for folding - Box tree
         SciMsg( m_pSci(i), SCI_MARKERDEFINE, SC_MARKNUM_FOLDEROPEN,    SC_MARK_BOXMINUS )
         SciMsg( m_pSci(i), SCI_MARKERDEFINE, SC_MARKNUM_FOLDER,        SC_MARK_BOXPLUS )
         SciMsg( m_pSci(i), SCI_MARKERDEFINE, SC_MARKNUM_FOLDERSUB,     SC_MARK_VLINE)
         SciMsg( m_pSci(i), SCI_MARKERDEFINE, SC_MARKNUM_FOLDERTAIL,    SC_MARK_LCORNER)
         SciMsg( m_pSci(i), SCI_MARKERDEFINE, SC_MARKNUM_FOLDEREND,     SC_MARK_BOXPLUSCONNECTED)
         SciMsg( m_pSci(i), SCI_MARKERDEFINE, SC_MARKNUM_FOLDEROPENMID, SC_MARK_EMPTY)   ' SC_MARK_BOXMINUSCONNECTED
         SciMsg( m_pSci(i), SCI_MARKERDEFINE, SC_MARKNUM_FOLDERMIDTAIL, SC_MARK_TCORNER)

         ' Draw line below if not expanded
         SciMsg( m_pSci(i), SCI_SETFOLDFLAGS, 16, 0 )

         ' Colors for folders closed and folders opened
         SciMsg( m_pSci(i), SCI_MARKERSETFORE, SC_MARKNUM_FOLDER,     pTheme->colors(CLR_FOLDSYMBOL).nFg)
         SciMsg( m_pSci(i), SCI_MARKERSETBACK, SC_MARKNUM_FOLDER,     pTheme->colors(CLR_FOLDSYMBOL).nBg)
         SciMsg( m_pSci(i), SCI_MARKERSETFORE, SC_MARKNUM_FOLDEROPEN, pTheme->colors(CLR_FOLDSYMBOL).nFg)
         SciMsg( m_pSci(i), SCI_MARKERSETBACK, SC_MARKNUM_FOLDEROPEN, pTheme->colors(CLR_FOLDSYMBOL).nBg)
      
      Else
         ' Disable folding of the procedures and functions
         SciMsg( m_pSci(i), SCI_SETPROPERTY, Cast(WPARAM, @"fold"), Cast(LPARAM, @"0") )
      End If
   
   next
   
   Function = 0
End Function



''
''
Function clsDocument.GetWord( ByVal curPos As Long = -1 ) As String

   dim as any ptr pSci = this.GetActiveScintillaPtr()
   ' Get word at the specified location or under the cursor
   Dim As Long x, y, p
   Dim As String buffer 

   ' Retrieve the current position
   If curPos = -1 Then curPos = SciMsg( pSci, SCI_GETCURRENTPOS, 0, 0)
   ' Retrieve the starting and ending position of the word
   x = SciMsg( pSci, SCI_WORDSTARTPOSITION, curPos, True)
   y = SciMsg( pSci, SCI_WORDENDPOSITION, curPos, False)
   If y > x Then
      ' Text range
      buffer = this.GetTextRange(x, y)
      ' Remove the $NUL
      p = Instr(buffer, Chr(0))
      If p Then buffer = Left(buffer, p - 1)
   End If
   buffer = AfxStrRemoveAny( buffer, Chr(13, 10, 34) & "()%," )
   Function = buffer

End Function

''
''
Function clsDocument.GetBookmarks() As String

   dim as any ptr pSci = this.GetActiveScintillaPtr()
   ' Retrieve all bookmark positions in the document and return it
   ' as a comma delimited string to be saved to project file.
   Dim As String buffer
   Dim As Long fMark  ' 32 bit value
   Dim As Long nLines = SciMsg( pSci, SCI_GETLINECOUNT, 0, 0) 

   For i As Long = 0 To nLines - 1
      fMark = SciMsg( pSci, SCI_MARKERGET, i, 0)
      If Bit(fMark, 0) Then
         buffer = buffer & i & ","
      End If
   Next

   Function = RTrim(buffer, ",")

End Function

''
''
Function clsDocument.SetBookmarks( ByVal sBookmarks As String ) As Long

   dim as any ptr pSci = this.GetActiveScintillaPtr()
   ' Set all bookmark positions in the document and based on the 
   ' incoming comma delimited string retrieved from a project file.
   sBookmarks = trim(sBookmarks)
   If Len(sBookmarks) = 0 Then Exit Function
   
   Dim As Long nCount = AfxStrParseCount(sBookmarks, ",")
   Dim As Long nLine
   
   For i As Long = 1 To nCount
      nLine = Val( **AfxStrParse(sBookmarks, i, ",") )
      SciMsg( pSci, SCI_MARKERADD, nLine, 0)
   Next

   Function = 0

End Function

''
''
Function clsDocument.LineDuplicate() As Long
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   ' Duplicate the current caret line, or an entire block of
   ' code should a selection be active.
   Dim startSelPos As Long     ' Starting position
   Dim endSelPos   As Long     ' Ending position
   Dim strText     As String   ' Selected text

   ' If startSelPos and endSelPos are the same there is no selection,
   startSelPos = SciMsg( pSci, SCI_GETSELECTIONSTART, 0, 0)
   endSelPos   = SciMsg( pSci, SCI_GETSELECTIONEND, 0, 0)
   
   If startSelPos = endSelPos Then   ' no selection
      ' Simply duplicate the line that the caret is on.
      SciMsg( pSci, SCI_LINEDUPLICATE, 0, 0)
   else
      ' Retrieve the text
      strText = this.GetTextRange(startSelPos, endSelPos)
      SciMsg( pSci, SCI_INSERTTEXT, -1, Cast(LPARAM, Strptr(strText)) )
      SciMsg( pSci, SCI_SETSELECTIONSTART, startSelPos, 0)
      SciMsg( pSci, SCI_SETSELECTIONEND, endSelPos, 0)
   end if

   Function = 0

End Function


''
''
function clsDocument.CompileDirectives( Directives() as COMPILE_DIRECTIVES ) as Long
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   ' Search the source code for any user embedded compiler directives.
   Dim ub     As Long    
   Dim i      As Long    
   Dim nLines As Long    
   dim st     as String
   
   nLines = SciMsg( pSci, SCI_GETLINECOUNT, 0, 0)

   For i = 0 To nLines
      st = ltrim(this.GetLine(i))
      if left(st, 1) <> "'" THEN continue for
      st = ltrim(mid(st, 2))

      if len(st) < 11 THEN continue for
      st = ucase(st)
      
      ub = ubound(Directives)
      
      ' '#CONSOLE ON|OFF
      if left(st, 11) = "#CONSOLE ON" THEN
         redim preserve Directives(ub+1)
         Directives(ub+1).DirectiveFlag = IDM_CONSOLE
      elseif left(st, 12) = "#CONSOLE OFF" THEN   
         redim preserve Directives(ub+1)
         Directives(ub+1).DirectiveFlag = IDM_GUI
      END IF

      ' '#RESOURCE "filename.rc"
      if left(st, 10) = "#RESOURCE " THEN
         redim preserve Directives(ub+1)
         Directives(ub+1).DirectiveFlag = IDM_RESOURCE
         st = mid(st, 11)
         Directives(ub+1).DirectiveText = AfxStrExtract(st, chr(34), chr(34))
      end if
         
   Next

   function = 0
END FUNCTION


