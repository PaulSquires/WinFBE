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
destructor clsDocument
   ' Delete any manually allocated CWindows in the designer
   for i as long = lbound(this.hWindow) to ubound(this.hWindow)
      DestroyWindow this.hWindow(i)    ' destroy the scintilla/designer windows
   next
   if this.IsDesigner then
      ' Remove all controls
      dim pCtrl as clsControl ptr 
      do until this.Controls.Count <= 1 
         for i as long = this.Controls.ItemFirst to this.Controls.ItemLast
            pCtrl = this.Controls.ItemAt(i)
            this.Controls.Remove(pCtrl)
         next
      loop   
      for i as long = 0 to 3
         DestroyWindow(this.hSnapLine(i))
      next
      DestroyWindow(this.hWndForm)
      DestroyWindow(this.hWndFrame)
      DestroyWindow(this.hWndDesigner)
      DeleteObject(this.hFontFakeMenu)
   end if

   ' Repaint the main area because we don't want any splitter to show
   AfxRedrawWindow(HWND_FRMMAIN)

end destructor


''
''  Returns true/False indicating whether user manual changes were made
''  outside of whatever edit changes Scintilla makes.
''
property clsDocument.UserModified() as boolean
   return m_UserModified
end property

property clsDocument.UserModified( byval nModified as boolean ) 
   m_UserModified = nModified
   if nModified = true then this.AutoSaveRequired = true
end property


' ========================================================================================
' Parse the code for this document. Invokes the ctxParser which does all of
' the actual parse work and database updating.
' ========================================================================================
function clsDocument.ParseDocument() as boolean
   if this.bNeedsParsing = false then 
      return false
   end if

   static as boolean bInParseDocument
   
   ' Prevent any potential recursive parsing 
   if bInParseDocument then exit function
   bInParseDocument = true
   
   ' Delete any existing data related to this pDoc
   gdb2.dbDeleteByDocumentPtr(@this)
   
   dim parser as ctxParser
   parser.nFileType = DB2_FILETYPE_USERCODE
   parser.parse(@this)
   frmOutput_UpdateToDoListview()
   
   bInParseDocument = false
   return true
end function


' ========================================================================================
' Determines if the IDC_SCINTILLA id being checked belongs to any of the 
' currently defined Scintilla windows (needed for splitter windows).
' ========================================================================================
function clsDocument.IsValidScintillaID( byval idScintilla as long ) as boolean
   for i as long = lbound(this.hWindow) to ubound(this.hWindow)
      if IDC_SCINTILLA + i = idScintilla then
         return true
      end if
   next
end function


''
''
function clsDocument.GetActiveScintillaPtr() as any ptr
   dim as hwnd hEdit = this.hWndActiveScintilla
   for i as long = lbound(this.hWindow) to ubound(this.hWindow)
      if this.hWindow(i) = hEdit then return m_pSci(i)
   next
   ' if no other matches then return 
   function = m_pSci(0)
end function


property clsDocument.hWndActiveScintilla() as hwnd
   if m_hWndActiveScintilla = 0 then m_hWndActiveScintilla = this.hWindow(0)
   property = m_hWndActiveScintilla
end property

property clsDocument.hWndActiveScintilla(byval hWindow as hwnd)
   m_hWndActiveScintilla = hWindow
end property

''
''  Returns true/False indicating that a valid top mainmenu exists for this form.
function clsDocument.MainMenuExists() as boolean
   dim as long numItems = (ubound(this.MenuItems) - lbound(this.MenuItems)) + 1
   if (this.GenerateMenu = true) andalso (numItems > 0) then
      return true
   else   
      return false
   end if
end function

''
''  Returns true/False indicating that a valid top toolbar exists for this form.
function clsDocument.ToolBarExists() as boolean
   dim as long numItems = (ubound(this.ToolBarItems) - lbound(this.ToolBarItems)) + 1
   if (this.GenerateToolBar = true) andalso (numItems > 0) then
      return true
   else   
      return false
   end if
end function

''
''  Returns true/False indicating that a valid top statusbar exists for this form.
function clsDocument.StatusBarExists() as boolean
   dim as long numItems = (ubound(this.PanelItems) - lbound(this.PanelItems)) + 1
   if (this.GenerateStatusBar = true) andalso (numItems > 0) then
      return true
   else   
      return false
   end if
end function

''
''
function clsDocument.CreateDesignerWindow( byval hWndParent as HWnd ) as hwnd 
   
   ' for the Visual Designer, there exists three (3) levels of windows:
   ' (1) The DesignMain (hWindow) used by the top tabcontrol to display the document. 
   '     This is the highest level window and is basically just the container for 
   '     the other two windows. Contains tab control to switch between design/code views.
   ' (2) The DesignFrame window. This is the scrollable window.
   ' (3) The DesignForm window. This is the actual visual form that we manipulate by
   '     adding controls to it, etc.
   ' (*4) There is a 4th window in the sense that the Scintilla code window swaps
   '      out the DesignFrame whenever the tabcontrol switches between design/code view.
   '
   this.IsDesigner = true
   
   ' In version 3.02+ form file encoding does not have to be unicode because the form meta
   ' data is saved to a separate .design file (UTF-8 encoded no sigature). Therefore, unless
   ' the user specifically changes the code file encoding then we will default to ANSI. That
   ' ANSI setting for the file was set on pDoc creation. Do not reset the value here otherwise
   ' it will affect any existing user files and essentially make all Form files as ANSI.
   ' this.FileEncoding = FILE_ENCODING_UTF16_BOM
   
   dim rc as RECT
   
   ' (1) Create the DesignMain window
   dim pMain as CWindow ptr = new CWindow
   pMain->DPI = AfxCWindowPtr(hwndParent)->DPI
   this.hWndDesigner = _
   pMain->Create( hWndParent, "", @DesignerMain_WndProc, 0, 0, 0, 0, _
        WS_CHILD or WS_CLIPSIBLINGS or WS_CLIPCHILDREN, _
        WS_EX_CONTROLPARENT or WS_EX_LEFT or WS_EX_LTRREADING or WS_EX_RIGHTSCROLLBAR)
   pMain->Brush = ghDesigner.hPanelBrush
   pMain->SetClientSize(3000, 3000)
   ' We will set our own mouse pointer as needed
   SetClassLongPtr(this.hWindow(0), GCLP_HCURSOR, 0)   
   ' Allow the Designer Frame window to scroll. This allows us to create Forms that
   ' are larger than the current viewable screen area.
   dim pScrollWindow as CScrollWindow ptr = new CScrollWindow(pMain->hWindow)
   pMain->ScrollWindowPtr = pScrollWindow
   

   ' (2) Create the Design Frame window (child of the Main)
   dim pFrame as CWindow ptr = new CWindow
   pFrame->DPI = AfxCWindowPtr(hwndParent)->DPI
   this.hWndFrame = _
   pFrame->Create( pMain->hWindow, "", @DesignerFrame_WndProc, 0, 0, 0, 0, _
        WS_CHILD or WS_VISIBLE or WS_CLIPSIBLINGS or WS_CLIPCHILDREN, _ 
        WS_EX_CONTROLPARENT or WS_EX_LEFT or WS_EX_LTRREADING or WS_EX_RIGHTSCROLLBAR)
   pFrame->ClassStyle = CS_DBLCLKS
   SetWindowLongPtr( pFrame->hWindow, GWLP_ID, IDC_DESIGNFRAME )
   ' We will set our own mouse pointer as needed
   SetClassLongPtr(this.hWndFrame, GCLP_HCURSOR, 0)   
   pFrame->Brush = ghDesigner.hPanelBrush
   
   
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
function clsDocument.CreateCodeWindow( _
            byval hWndParent as HWnd, _
            byval IsNewFile  as boolean, _     
            byval IsTemplate as boolean = false, _
            byref wszFile    as wstring = "" _
            ) as HWnd   

   ' Creates a Scintilla editing window (initially not visible). Optionally, load a diskfile
   ' into the window and apply properties to it.
   for i as long = lbound(this.hWindow) to ubound(this.hWindow)
      this.hWindow(i) = CreateWindowEx( 0, "Scintilla", "", _
                        WS_CHILD or WS_TABSTOP or WS_CLIPCHILDREN, _
                        0,0,0,0,hWndParent, _
                        cast(HMENU, IDC_SCINTILLA+i), GetModuleHandle(null), null)

       SendMessage( this.hWindow(i), SCI_SETMODEVENTMASK, _
                      SC_MOD_INSERTTEXT or SC_MOD_DELETETEXT, 0 )
   
      ' Initialize our direct access to the Scintilla code windows. This is much faster than
      ' using SendMessage to the window. Only need to initialize once no matter how many
      ' code windows that are eventually opened.
      if IsWindow(this.hWindow(i)) then
         ' NOTE: In my testing, need to only set the Scintilla lexer to the base editing
         ' window only and NOT both split windows. Also need to do this immediately after
         ' the window is created and do not send the message again afterwards.
         ' Also, every window must have a separate new call to CreateLexer. We can not
         ' just get one lexer and then try to share it amongst multiple new windows. When
         ' a window is destroyed then the pointer would be as well causing other existing
         ' windows to GPF.
         if i = 0 then
            ' Load the FB lexer from Lexilla and feed it into Scintilla
            dim as any ptr pLexer = gApp.pfnCreateLexerfn( "winfbe" )
            SendMessage( this.hWindow(i), SCI_SETILEXER, 0, cast(LPARAM, pLexer) )
         end if
         if SciMsg = 0 then
            SciMsg = cast( Scintilla_Directfunction, SendMessage( this.hWindow(0), SCI_GETDIRECTFUNCTION, 0, 0 ) )
         end if
         ' Call the direct function for speed purposes rather than relying on the traditional SendMessage method.
         m_pSci(i) = cast(any ptr, SendMessage( this.hWindow(i), SCI_GETDIRECTPOINTER, 0, 0 )) 
      end if
   next

   ' Disable scintilla vertical scroll bar (wParam = 1 to enable)
   SciMsg( m_pSci(0), SCI_SETVSCROLLBAR, 0, 0 )
   SciMsg( m_pSci(0), SCI_SETHSCROLLBAR, 0, 0 )
   SciMsg( m_pSci(1), SCI_SETVSCROLLBAR, 0, 0 )
   SciMsg( m_pSci(1), SCI_SETHSCROLLBAR, 0, 0 )
   
   ' Get the document pointer from our main control and assign it to the other split windows
   dim as any ptr pDoc = cast(any ptr, SciMsg(m_pSci(0), SCI_GETDOCPOINTER, 0, 0))
   if pDoc then SciMsg( m_pSci(1), SCI_SETDOCPOINTER, 0, cast(LPARAM, pDoc)) 
   
   dim nResult as long = IS_TEXT_UNICODE_SIGNATURE
      
   ' if a disk file was specified then open it and load it into the editor
   this.IsNewFlag = IsNewFile
   if (IsNewFile = true) orelse (IsTemplate = true) then
      this.nextFileNum = this.nextFileNum + 1
      this.DiskFilename = "Untitled" & this.nextFileNum
      if this.IsDesigner then
         this.CreateDesignerWindow(hWndParent)  ' Create the new visual designer window
      end if   
   end if
   this.ProjectFileType = FILETYPE_UNDEFINED

   if len(wszFile) then   
      ' do not use Dir() > "" here b/c if incoming file originated from a Do/loop 
      ' of files using Dir() then there will be problems.

      if AfxFileExists(wszFile) then     
         dim as string st, sText
         dim as long idx
         
         if IsTemplate then
            dim pStream as CTextStream
            if pStream.Open(wszFile) = S_OK then
               ' Look at the first 4 lines
               ' Line 3 tells us the file type (bas or xml)
               do until pStream.EOS
                  st = pStream.ReadLine
                  idx = idx + 1
                  select case idx
                     case 1
                     case 2
                     case 3: this.DiskFilename = this.DiskFilename & trim(st)    
                     case 4 
                     case else
                        sText = sText & st & vbCrLf
                  end select   
               loop
               pStream.Close
               if IsTextUnicode(strptr(sText), 2, cast(LPINT, @nResult) ) then
                  sText = mid(sText, 3)
                  sText = AfxACode( cast(wstring ptr, strptr(sText)) )
               end if   
               this.SetText( sText ) 
            end if
            ' Force the template file to be considered "new" so it will be saved.
            this.IsNewFlag = true
            ' Search for "|", replace it with an empty space "" and position the caret in that place
            this.FindReplace( "|", "" )
            ' don't set SAVEPOINT for newly loaded Template files because we want the document to display as dirty

         else

            ' Set a flag that we are loading code from a file. This is important because if the file
            ' contains visual designer code then we only want to apply control properties after all of 
            ' the file is read, otherwise it would get applied immediately after each call to CreateToolboxControl.
            this.LoadingFromFile = true
            this.DiskFilename = wszFile   ' assign before GetFileTostring because needed for RelativeFile calculation.
            this.DesignerFilename = wszFile & ".design"
            
            ' GetFileTostring has a call to pDoc->ParseFormMetaData which will create/recreate the designer form.
            ' We continue to use this function in post version 3.02 form file format in order to ensure
            ' that older form files can be imported. 
            dim sText as string  ' this will be an UTF-8 encoded string

            if GetFileToString(wszFile, sText, @this) = false then
               ' Take this opportunity to determine the text line endings
               if instr(sText, chr(13,10)) then
                  SciMsg( m_pSci(0), SCI_SETEOLMODE, SC_EOL_CRLF, 0)
               elseif instr(sText, chr(10)) then
                  SciMsg( m_pSci(0), SCI_SETEOLMODE, SC_EOL_LF, 0)
               elseif instr(sText, chr(13)) then
                  SciMsg( m_pSci(0), SCI_SETEOLMODE, SC_EOL_CR, 0)
               end if
               this.SetText( sText ) 
               this.DateFileTime = AfxGetFileLastWriteTime( wszFile )
            else
               print "Error opening: "; wszFile
            end if

            ' Do a check for IsDesigner. New post version 3.02+ files have new form JSON file
            ' format and require a different parsing routine in order to load. They are also
            ' stored in UTF-8 (no signature).
            if AfxFileExists(this.DesignerFilename) then 
               this.IsDesigner = true
               if (ConvertWinFBEversion(this.wszFormVersion) >= ConvertWinFBEversion("3.0.2")) orelse _
                  (ConvertWinFBEversion(this.wszFormVersion) = 0) then  'b/c upgraded form files do not have version loaded yet
                  sText = ""
                  ' Need to save and then restore the File Encoding for this pDoc because the call to GetFileString
                  ' for the designer file could reset the encoding of the main file.
                  dim as long savedEncoding = this.FileEncoding
                  if GetFileToString(this.DesignerFilename, sText, @this) = false then
                     this.CreateDesignerWindow(hWndParent)        ' Create the new visual designer window
                     this.LoadFormJSONdata(HWND_FRMMAIN, sText)   ' this should now be UTF-8
                  end if   
                  this.FileEncoding = savedEncoding
               end if
            end if

            this.LoadingFromFile = false
            SciMsg( m_pSci(0), SCI_SETSAVEPOINT, 0, 0)
            SciMsg( m_pSci(0), SCI_EMPTYUNDOBUFFER, 0, 0)
            this.UserModified = false
            this.AutoSaveRequired = false
            this.AutoSaveFilename = OnCommand_FileAutoSaveGenerateFilename(wszFile)
            ' Update the most recently used file list (only for non-IsNewFlag files)
            ' Only add file to MRU list if it is not part of an active Project.
            if gApp.IsProjectActive = false then 
               UpdateMRUList(wszFile)
            end if
         end if
      end if   
   end if

   if this.IsDesigner then
      ' select the Form as the default focus control
      this.Controls.SelectControl( this.hWndForm )
      this.Controls.SetActiveControl( this.hWndForm )

      ' if the data that was read contained anything that caused the form to have
      ' to be code regenerated then we need to save the file back to disk with the
      ' new meta data and code generation.
      if this.bRegenerateCode then 
         if this.IsNewFlag = false then this.SaveFile
      end if   

   end if   

   ' Apply code editor properties to the edit window
   this.ApplyProperties

   function = this.hWindow(0)
end function

''
''
function clsDocument.FindReplace( _
            byval strFindText as string, _
            byval strReplaceText as string _
            ) as long

   ' return Value: new position if successful; -1 if text not found.
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   dim as long findFlags = SCFIND_MATCHcase or SCFIND_WHOLEWORD
   dim as long startPos  = SciMsg( pSci, SCI_GETCURRENTPOS, 0, 0)
   dim as long endPos    = SciMsg( pSci, SCI_GETTEXTLENGTH, 0, 0)
   dim as long newPos

   ' Set the start and end positions, and search flags, and finally do the search
   SciMsg( pSci, SCI_SETTARGETSTART, startPos, 0)
   SciMsg( pSci, SCI_SETTARGETEND, endPos, 0)
   SciMsg( pSci, SCI_SETSEARCHFLAGS, findFlags, 0)

   ' Search the text to replace
   newPos = SciMsg( pSci, SCI_SEARCHINTARGET, len(strFindText), cast(LPARAM, strptr(strFindText)) )
   
   ' return -1 if not found
   if newPos = - 1 then return -1
   
   gApp.SuppressNotify = true
   ' Position the caret and select the text
   SciMsg( pSci, SCI_SETCURRENTPOS, newPos, 0)
   SciMsg( pSci, SCI_GOTOPOS, newPos, 0)
   SciMsg( pSci, SCI_SETSELECTIONSTART, newPos, 0)
   SciMsg( pSci, SCI_SETSELECTIONEND, newPos + len(strFindText), 0)

   ' Replace the selection (SCI_REPLACESEL fails if text is "" so use Cut instead for that scenario)                                
   if len(strReplaceText) = 0 then
      SciMsg( pSci, SCI_CUT, 0, 0 )
   else
      SciMsg( pSci, SCI_REPLACESEL, 0, cast(LPARAM, strptr(strReplaceText)) )
   end if
   gApp.SuppressNotify = false
   
   ' return the new position
   function = newPos

end function

''
''
function clsDocument.InsertFile() as boolean
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   ' Display the Open File Dialog
   dim pwszName as wstring ptr = AfxIFileOpenDialogW(HWND_FRMMAIN, IDM_INSERTFILE)
   if pwszName then
      ' save the main file encoding because GetFileTostring may change it
      dim as string sText
      GetFileTostring(*pwszName, sText, @this)
      SciMsg( pSci, SCI_INSERTTEXT, -1, cast(LPARAM, strptr(sText)) )
      CoTaskMemFree pwszName
      this.ApplyProperties()
   end if
   
   function = 0
end function


' ========================================================================================
' Load the JSON data to populate the various Form's data structures.
' ========================================================================================
function clsDocument.LoadFormJSONdata( _
            byval hWndParent as HWnd, _
            byref JSONtext as string, _    ' this will be UTF-8 text
            byval bLoadOnly  as boolean = false _
            ) as long

   dim pCtrl as clsControl ptr
   dim pCtrlActive as clsControl ptr
   
   dim as RECT rc
   dim as CWSTR wst, wszPropName, wszPropValue, wszEventName
   dim as boolean bIsValidControl = true
   
   dim as long numItems
   
   dim as cJSON ptr filePtr, keyPtr
   dim as cJSON ptr itemsPtr, itemPtr
   dim as cJSON ptr ctrlsPtr, ctrlPtr
   dim as cJSON ptr propsPtr, propPtr
   dim as cJSON ptr eventsPtr, eventPtr
   
   dim json as CJSON_TYPE
   
   filePtr = json.parse(JSONtext)
   if filePtr = NULL then exit function
   
   ' First, make sure that this is actually a Form file!
   if json.getbool( filePtr, "form" ) <> true then exit function
   
   this.IsDesigner = true

   ' Remove any previously created controls, etc otherwise reloading the form
   ' file will result in duplicates.
   if bLoadOnly = false then
      if this.hWndForm then DestroyWindow(this.hWndForm)
      if this.hWndFrame then DestroyWindow(this.hWndFrame)
      if this.hWndDesigner then DestroyWindow(this.hWndDesigner)
      this.CreateDesignerWindow(hWndParent)
   end if

   this.wszFormVersion.Utf8 = json.gettext(filePtr, "version")
   if this.wszFormVersion = "" then this.wszFormVersion = APPVERSION
         
   this.bLockControls = json.getbool(filePtr, "lockcontrols")
   this.bSnapLines = json.getbool(filePtr, "snaplines")

   ' Retrieve data related to IMAGES
   itemsPtr = json.getptr(filePtr, "images")
   numItems = json.arraycount(itemsPtr)
   if numItems then
      redim preserve this.AllImages(numItems - 1)
      dim as long i = 0
      itemPtr = json.arrayitem(itemsPtr, i)
      do until itemPtr = 0
         this.AllImages(i).pDoc = @this
         this.AllImages(i).wszImageName.Utf8 = json.gettext(itemPtr, "imagename")
         wst.Utf8 = json.gettext(itemPtr, "filename")
         ' if this is a relative filename then convert it back.
         dim as CWSTR wszImageFilename = ProcessFromCurdriveProject(wst)
         if AfxPathIsRelative(wszImageFilename) then 
            wszImageFilename = AfxPathCombine( AfxStrPathName("PATH", this.DiskFilename), wszImageFilename)
         end if
         this.AllImages(i).wszFileName = wszImageFilename
         
         this.AllImages(i).wszFormat.Utf8 = json.gettext(itemPtr, "resourcetype")
         i += 1
         itemPtr = itemPtr->next
      loop   
   end if   

   ' Retrieve data related to MAINMENU 
   keyPtr = json.getptr(filePtr, "mainmenu")
   if keyPtr then
      this.GenerateMenu = json.getbool(keyPtr, "display")
      itemsPtr = json.getptr(keyPtr, "items")
      numItems = json.arraycount(itemsPtr)
      if numItems then
         redim preserve this.MenuItems(numItems - 1)
         dim as long i = 0
         itemPtr = json.arrayitem(itemsPtr, i)
         do until itemPtr = 0
            dim as CWSTR wszName
            wszName.Utf8 = json.gettext(itemPtr, "name")
            if len(rtrim(wszName)) > 0 then 
               this.MenuItems(i).wszName = wszName
               this.MenuItems(i).wszCaption.Utf8 = json.gettext(itemPtr, "caption")
               this.MenuItems(i).nIndent = json.getnumber(itemPtr, "indent") 
               this.MenuItems(i).chkAlt = json.getnumber(itemPtr, "alt")
               this.MenuItems(i).chkShift = json.getnumber(itemPtr, "shift")
               this.MenuItems(i).chkCtrl = json.getnumber(itemPtr, "ctrl")
               this.MenuItems(i).wszShortcut.Utf8 = json.gettext(itemPtr, "shortcut")
               this.MenuItems(i).chkChecked = json.getnumber(itemPtr, "checked")
               this.MenuItems(i).chkGrayed = json.getnumber(itemPtr, "grayed")
               i += 1
            end if   
            itemPtr = itemPtr->next
         loop
      end if
   end if
   
   ' Retrieve data related to STATUSBAR
   keyPtr = json.getptr(filePtr, "statusbar")
   if keyPtr then
      this.GenerateStatusBar = json.getbool(keyPtr, "display")
      itemsPtr = json.getptr(keyPtr, "items")
      numItems = json.arraycount(itemsPtr)
      if numItems then
         redim preserve this.PanelItems(numItems - 1)
         dim as long i = 0
         itemPtr = json.arrayitem(itemsPtr, i)
         do until itemPtr = 0
            this.PanelItems(i).wszName.Utf8 = json.gettext(itemPtr, "name") 
            this.PanelItems(i).wszText.Utf8 = json.gettext(itemPtr, "text")
            this.PanelItems(i).wszTooltip.Utf8 = json.gettext(itemPtr, "tooltip")
            this.PanelItems(i).wszAlignment.Utf8 = json.gettext(itemPtr, "alignment")
            this.PanelItems(i).wszAutoSize.Utf8 = json.gettext(itemPtr, "autosize")
            this.PanelItems(i).wszWidth.Utf8 = json.gettext(itemPtr, "width")
            this.PanelItems(i).wszMinWidth.Utf8 = json.gettext(itemPtr, "minwidth")
            this.PanelItems(i).pProp.wszPropValue.Utf8 = json.gettext(itemPtr, "image")
            this.PanelItems(i).wszBackColor.Utf8 = json.gettext(itemPtr, "backcolor")
            this.PanelItems(i).wszBackColorHot.Utf8 = json.gettext(itemPtr, "backcolorhot")
            this.PanelItems(i).wszForeColor.Utf8 = json.gettext(itemPtr, "forecolor")
            this.PanelItems(i).wszForeColorHot.Utf8 = json.gettext(itemPtr, "forecolorhot")
            ' BorderStyle is deprecated as of v2.0.4 as it has no effect
            ' in WinFBE programs where Windows Themes are enabled.
            i += 1
            itemPtr = itemPtr->next
         loop
      end if
   end if
   
   ' Retrieve data related to TOOLBAR
   keyPtr = json.getptr(filePtr, "toolbar")
   if keyPtr then
      this.GenerateToolbar = json.getbool(keyPtr, "display")
      this.wszToolBarSize.Utf8 = json.gettext(keyPtr, "size")
      itemsPtr = json.getptr(keyPtr, "items")
      numItems = json.arraycount(itemsPtr)
      if numItems then
         redim preserve this.ToolBarItems(numItems - 1)
         dim as long i = 0
         itemPtr = json.arrayitem(itemsPtr, i)
         do until itemPtr = 0
            this.ToolBarItems(i).wszName.Utf8 = json.gettext(itemPtr, "name") 
            this.ToolBarItems(i).wszButtonType.Utf8 = json.gettext(itemPtr, "type")
            this.ToolBarItems(i).wszTooltip.Utf8 = json.gettext(itemPtr, "tooltip")
            this.ToolBarItems(i).pPropNormalImage.wszPropValue.Utf8 = json.gettext(itemPtr, "normalimage")
            this.ToolBarItems(i).pPropHotImage.wszPropValue.Utf8 = json.gettext(itemPtr, "hotimage")
            this.ToolBarItems(i).pPropDisabledImage.wszPropValue.Utf8 = json.gettext(itemPtr, "disabledimage")
            i += 1
            itemPtr = itemPtr->next
         loop
      end if
   end if

   ctrlsPtr = json.getptr(filePtr, "controls")
   dim as long numControls = json.arraycount(ctrlsPtr)
   if numControls then
      dim wszControlType as CWSTR
      dim nControlType as long

      ctrlPtr = json.arrayitem(ctrlsPtr, 0)
      do until ctrlPtr = 0
         wszControlType.Utf8 = json.gettext(ctrlPtr, "type") 
         nControlType = GetControlType( wszControlType)

         ' Control Start
         if nControlType = 0 then   
            ' no longer a valid toolbox control
            ctrlPtr = ctrlPtr->next
            continue do
         else
            if bLoadOnly = true then
               dim pCtrlTemp as clsControl ptr = new clsControl
               pCtrlTemp->ControlType = nControlType
               this.Controls.Add(pCtrlTemp)
            else
               pCtrl = CreateToolboxControl( @this, nControlType, rc )
               pCtrlActive = pCtrl
            end if   
         end if
         
         propsPtr = json.getptr(ctrlPtr, "properties")
         dim as long numProps = json.arraycount(propsPtr)
         if numProps then
            propPtr = json.arrayitem(propsPtr, 0)
            do until propPtr = 0
               wszPropName.Utf8 = json.gettext(propPtr, "name")
               wszPropValue.Utf8 = json.gettext(propPtr, "value")
               ' Only set the loading property if it exists in the current property
               ' listing. We do this otherwise older now unused properties will continue
               ' to get loaded when we no longer want them to.
               if IsPropertyExists( pCtrl, wszPropName ) then
                  SetControlProperty( pCtrl, wszPropName, wszPropValue ) 
               end if
               propPtr = propPtr->next
            loop
         end if   

         eventsPtr = json.getptr(ctrlPtr, "events")
         dim as long numEvents = json.arraycount(eventsPtr)
         if numEvents then
            eventPtr = json.arrayitem(eventsPtr, 0)
            do until eventPtr = 0
               wszEventName.Utf8 = json.gettext(eventPtr, "name")
               SetControlEvent( pCtrl, wszEventName, true )
               eventPtr = eventPtr->next
            loop
         end if   

         ' Control End
         if bLoadOnly = false then
            pCtrl->SuspendLayout = true
            ApplyControlProperties( @this, pCtrl )
            pCtrl->SuspendLayout = false
         end if

         ctrlPtr = ctrlPtr->next
      loop

   end if

   if filePtr then json.deleteptr(filePtr)

   return 0
   
end function


' ========================================================================================
' Parse all of the file's code to remove and process any visual designer specific code.
' Returns a string representing just the code only (visual designer metastatements removed).
' ========================================================================================
function clsDocument.ParseFormMetaData( _
            byval hWndParent as HWnd, _
            byref wszAllText as wstring, _
            byval bLoadOnly  as boolean = false _
            ) as CWSTR
            
   ' NOTE: THIS FUNCTION IS STILL REQUIRED IN ORDER TO LOAD PRE-VERSION 3.02 FORM FILE
   ' THAT ARE THEN UPGRADED BY WINFBE AFTER THEY ARE LOADED. 
   
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
   dim as long nControlType, numLines, numOffsetLines
   dim as boolean bIsValidControl = true
   
   ' The first line MUST contain the identifier that this is a form file.

   if left(wszAllText, 13) = "' WINFBE FORM" then
      this.IsDesigner = true
      ' Remove any previously created controls, etc otherwise reloading the form
      ' file will result in duplicates.
      if bLoadOnly = false then
         if this.hWndForm then DestroyWindow(this.hWndForm)
         if this.hWndFrame then DestroyWindow(this.hWndFrame)
         if this.hWndDesigner then DestroyWindow(this.hWndDesigner)
         this.CreateDesignerWindow(hWndParent)
      end if
      numOffsetLines = 1
   else
      ' This is just a regular code file with no visual designer
      this.IsDesigner = false
      return wszAllText   
   end if
   

   ' Iterate all of the lines related to the visual designer
   dim as long iLineStart = 1
   dim as long iLineEnd, nextMenuItem, nextToolBarItem, nextPanelItem, nextImageItem

   dim as boolean bReadingMenuItem, bReadingToolBarItem, bReadingPanelItem, bReadingImageItem

   do until iLineStart >= len(wszAllText)
      
      iLineend = instr(iLineStart, wszAllText, vbcrlf)
      if iLineend = 0 then iLineend = len(wszAllText)  ' cr/lf not found
      wst = mid(wszAllText, iLineStart, iLineend - iLineStart)   
      iLineStart = iLineStart + len(wst) + len(vbcrlf)
      
      numOffsetLines = numOffsetLines + 1
      
      if left(wst, 16) = "' WINFBE VERSION" then
         wszPropValue = trim(mid(wst, 17))
         if ConvertWinFBEversion(wszPropValue) < ConvertWinFBEversion(APPVERSION) then
            this.bRegenerateCode = true
            this.UserModified = true
         end if
         this.wszFormVersion = wszPropValue
         
      elseif left(wst, 15) = "' LOCKCONTROLS=" then
         wszPropValue = mid(wst, 16)  ' default is false
         if wszPropValue = "true" then this.bLockControls = true

      elseif left(wst, 12) = "' SNAPLINES=" then
         wszPropValue = mid(wst, 13)  ' default is true
         if wszPropValue = "False" then this.bSnapLines = false

      elseif left(wst, 17) = "' WINFBE FORM_END" then
      
      elseif left(wst, 18) = "'   MENUITEM_START" then
         dim as long ub = ubound(this.MenuItems)
         redim preserve this.MenuItems(ub + 1)
         nextMenuItem = ub + 1
         bReadingMenuItem = true
      elseif left(wst, 21) = "'   MAINMENU_DISPLAY=" then
         this.GenerateMenu = val(mid(wst, 22))
      elseif left(wst, 11) = "'     NAME=" then
         if bReadingMenuItem then this.MenuItems(nextMenuItem).wszName = mid(wst, 12)
      elseif left(wst, 14) = "'     CAPTION=" then
         if bReadingMenuItem then this.MenuItems(nextMenuItem).wszCaption = mid(wst, 15)
      elseif left(wst, 13) = "'     INDENT=" then
         if bReadingMenuItem then this.MenuItems(nextMenuItem).nIndent = val(mid(wst, 14))
      elseif left(wst, 10) = "'     ALT=" then
         if bReadingMenuItem then this.MenuItems(nextMenuItem).chkAlt = val(mid(wst, 11))
      elseif left(wst, 12) = "'     SHIFT=" then
         if bReadingMenuItem then this.MenuItems(nextMenuItem).chkShift = val(mid(wst, 13))
      elseif left(wst, 11) = "'     CTRL=" then
         if bReadingMenuItem then this.MenuItems(nextMenuItem).chkCtrl = val(mid(wst, 12))
      elseif left(wst, 15) = "'     SHORTCUT=" then
         if bReadingMenuItem then this.MenuItems(nextMenuItem).wszShortcut = mid(wst, 16)
      elseif left(wst, 14) = "'     CHECKED=" then
         if bReadingMenuItem then this.MenuItems(nextMenuItem).chkChecked = val(mid(wst, 15))
      elseif left(wst, 13) = "'     GRAYED=" then
         if bReadingMenuItem then this.MenuItems(nextMenuItem).chkGrayed = val(mid(wst, 14))
      elseif left(wst, 16) = "'   MENUITEM_END" then
         bReadingMenuItem = false   
      
      elseif left(wst, 21) = "'   TOOLBARITEM_START" then
         dim as long ub = ubound(this.ToolBarItems)
         redim preserve this.ToolBarItems(ub + 1)
         nextToolBarItem = ub + 1
         bReadingToolBarItem = true
      elseif left(wst, 20) = "'   TOOLBAR_DISPLAY=" then
         this.GenerateToolBar = val(mid(wst, 21))
      elseif left(wst, 17) = "'   TOOLBAR_SIZE=" then
         this.wszToolBarSize = mid(wst, 18)
      elseif left(wst, 17) = "'     BUTTONNAME=" then
         if bReadingToolBarItem then this.ToolBarItems(nextToolBarItem).wszName = mid(wst, 18)
      elseif left(wst, 17) = "'     BUTTONTYPE=" then
         if bReadingToolBarItem then this.ToolBarItems(nextToolBarItem).wszButtonType = mid(wst, 18)
      elseif left(wst, 20) = "'     BUTTONTOOLTIP=" then
         if bReadingToolBarItem then this.ToolBarItems(nextToolBarItem).wszTooltip = mid(wst, 21)
      elseif left(wst, 24) = "'     BUTTONNORMALIMAGE=" then
         if bReadingToolBarItem then this.ToolBarItems(nextToolBarItem).pPropNormalImage.wszPropValue = mid(wst, 25)
      elseif left(wst, 21) = "'     BUTTONHOTIMAGE=" then
         if bReadingToolBarItem then this.ToolBarItems(nextToolBarItem).pPropHotImage.wszPropValue = mid(wst, 22)
      elseif left(wst, 26) = "'     BUTTONDISABLEDIMAGE=" then
         if bReadingToolBarItem then this.ToolBarItems(nextToolBarItem).pPropDisabledImage.wszPropValue = mid(wst, 27)
      elseif left(wst, 19) = "'   TOOLBARITEM_END" then
         bReadingToolBarItem = false   


      elseif left(wst, 19) = "'   PANELITEM_START" then
         dim as long ub = ubound(this.PanelItems)
         redim preserve this.PanelItems(ub + 1)
         nextPanelItem = ub + 1
         bReadingPanelItem = true
      elseif left(wst, 22) = "'   STATUSBAR_DISPLAY=" then
         this.GenerateStatusBar = val(mid(wst, 23))
      elseif left(wst, 16) = "'     PANELNAME=" then
         if bReadingPanelItem then this.PanelItems(nextPanelItem).wszName = mid(wst, 17)
      elseif left(wst, 16) = "'     PANELTEXT=" then
         if bReadingPanelItem then this.PanelItems(nextPanelItem).wszText = mid(wst, 17)
      elseif left(wst, 19) = "'     PANELTOOLTIP=" then
         if bReadingPanelItem then this.PanelItems(nextPanelItem).wszTooltip = mid(wst, 20)
      elseif left(wst, 21) = "'     PANELALIGNMENT=" then
         if bReadingPanelItem then this.PanelItems(nextPanelItem).wszAlignment = mid(wst, 22)
      ' BorderStyle is deprecated as of v2.0.4 as it has no effect
      ' in WinFBE programs where Windows Themes are enabled.
      'elseif left(wst, 23) = "'     PANELBORDERSTYLE=" then
      '   if bReadingPanelItem then this.PanelItems(nextPanelItem).wszBorderStyle = mid(wst, 24)
      elseif left(wst, 20) = "'     PANELAUTOSIZE=" then
         if bReadingPanelItem then this.PanelItems(nextPanelItem).wszAutoSize = mid(wst, 21)
      elseif left(wst, 17) = "'     PANELWIDTH=" then
         if bReadingPanelItem then this.PanelItems(nextPanelItem).wszWidth = mid(wst, 18)
      elseif left(wst, 20) = "'     PANELMINWIDTH=" then
         if bReadingPanelItem then this.PanelItems(nextPanelItem).wszMinWidth = mid(wst, 21)
      elseif left(wst, 17) = "'     PANELIMAGE=" then
         if bReadingPanelItem then this.PanelItems(nextPanelItem).pProp.wszPropValue = mid(wst, 18)
      elseif left(wst, 21) = "'     PANELBACKCOLOR=" then
         if bReadingPanelItem then this.PanelItems(nextPanelItem).wszBackColor = mid(wst, 22)
      elseif left(wst, 24) = "'     PANELBACKCOLORHOT=" then
         if bReadingPanelItem then this.PanelItems(nextPanelItem).wszBackColorHot = mid(wst, 25)
      elseif left(wst, 21) = "'     PANELFORECOLOR=" then
         if bReadingPanelItem then this.PanelItems(nextPanelItem).wszForeColor = mid(wst, 22)
      elseif left(wst, 24) = "'     PANELFORECOLORHOT=" then
         if bReadingPanelItem then this.PanelItems(nextPanelItem).wszForeColorHot = mid(wst, 25)
      elseif left(wst, 17) = "'   PANELITEM_END" then
         bReadingPanelItem = false   
      
      elseif left(wst, 15) = "'   IMAGE_START" then
         dim as long ub = ubound(this.AllImages)
         redim preserve this.AllImages(ub + 1)
         nextImageItem = ub + 1
         this.AllImages(nextImageItem).pDoc = @this
         bReadingImageItem = true
      elseif left(wst, 16) = "'     IMAGENAME=" then
         if bReadingImageItem then this.AllImages(nextImageItem).wszImageName = mid(wst, 17)
      elseif left(wst, 15) = "'     FILENAME=" then
         if bReadingImageItem then 
            ' if this is a relative filename then convert it back.
            dim as CWSTR wszImageFilename = ProcessFromCurdriveProject(mid(wst, 16))
            if AfxPathIsRelative(wszImageFilename) then 
               wszImageFilename = AfxPathCombine( AfxStrPathName("PATH", this.DiskFilename), wszImageFilename)
            end if
            this.AllImages(nextImageItem).wszFileName = wszImageFilename
         end if
      elseif left(wst, 19) = "'     RESOURCETYPE=" then
         if bReadingImageItem then this.AllImages(nextImageItem).wszFormat = mid(wst, 20)
      elseif left(wst, 13) = "'   IMAGE_END" then
         bReadingImageItem = false   

      elseif left(wst, 20) = "' WINFBE_CODEGEN_END" then
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
            if bLoadOnly = true then
               dim pCtrlTemp as clsControl ptr = new clsControl
               pCtrlTemp->ControlType = nControlType
               this.Controls.Add(pCtrlTemp)
            else
               pCtrl = CreateToolboxControl( @this, nControlType, rc )
               pCtrlActive = pCtrl
            end if   
         end if
         
      elseif left(wst, 20) = "' WINFBE CONTROL_END" then
         if bIsValidControl then
            if bLoadOnly = false then
               pCtrl->SuspendLayout = true
               ApplyControlProperties( @this, pCtrl )
               pCtrl->SuspendLayout = false
            end if
         end if

      elseif left(wst, 16) = "'     PROP_NAME=" then
         if bIsValidControl then
            wszPropName = mid(wst, 17)
         end if
         
      elseif left(wst, 17) = "'     PROP_VALUE=" then
         if bIsValidControl then
            wszPropValue = mid(wst, 18)   ' utf8 encoded 
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
         
      elseif left(wst, 17) = "'     EVENT_NAME=" then
         if bIsValidControl then
            wszEventName = mid(wst, 18)
            SetControlEvent(pCtrl, wszEventName, true)
         end if
      end if   

   loop

   return wszAllText
   
end function


''
''
function clsDocument.SaveFormJSONdata() as boolean
   if this.IsDesigner = false then return true
   
   dim pCtrl as clsControl ptr
   dim as CWSTR wst 
   dim as CWSTR wcomma = ","
   
   wst = _
   qnum("form:true")      & wcomma & _
   qstr("version:3.0.2")  & wcomma & _
   qnum("lockcontrols:"   & iif(this.bLockControls, "true", "false")) & wcomma & _
   qnum("snaplines:"      & iif(this.bSnapLines, "true", "false")) 

   ' Save Images(if applicable)
   dim as long numImageItems = ubound(this.AllImages) - lbound(this.AllImages) + 1
   if numImageItems > 0 then
      wst = wst & wcomma & _
      qnum("images: [")
      for ii as long = lbound(this.AllImages) to ubound(this.AllImages)
         dim as CWSTR wszRelative
         dim as CWSTR wszImageFilename = this.AllImages(ii).wszFilename
         
         ' Attempt to convert the image file name to relative path
         if AfxFileExists( this.DiskFilename ) then
            wszRelative = AfxPathRelativePathTo( this.DiskFilename, FILE_ATTRIBUTE_NORMAL, wszImageFilename, FILE_ATTRIBUTE_NORMAL)
            if AfxPathIsRelative(wszRelative) then wszImageFilename = wszRelative
         end if
         
         wst = wst & _
         "{" & _
         qstr("imagename:"        & this.AllImages(ii).wszImageName) & wcomma & _ 
         qstr("filename:"         & AfxStrReplace(ProcessToCurdriveProject(wszImageFilename), "\", "\\")) & wcomma & _ 
         qstr("resourcetype:"     & this.AllImages(ii).wszFormat) & _ 
         "},"
      next
      wst = rtrim(wst, ",") & "]"
   end if

   ' Save MainMenu (if applicable)
   dim as long numMenuItems = ubound(this.MenuItems) - lbound(this.MenuItems) + 1
   if numMenuItems > 0 then
      wst = wst & wcomma & _
      qnum("mainmenu:{") & _
      qnum("display:" & iif(this.GenerateMenu, "true", "false")) & wcomma 

      wst = wst & _
      qnum("items: [")
      for ii as long = lbound(this.MenuItems) to ubound(this.MenuItems)
         ' Skip any "blank" lines that may have been added from the Menu Editor
         if rtrim(this.MenuItems(ii).wszName) = "" then continue for
         
         wst = wst & _
         "{" & _
         qstr("name:"      & this.MenuItems(ii).wszName) & wcomma & _
         qstr("caption:"   & this.MenuItems(ii).wszCaption) & wcomma & _
         qnum("indent:"    & this.MenuItems(ii).nIndent) & wcomma & _
         qnum("alt:"       & this.MenuItems(ii).chkAlt) & wcomma & _
         qnum("shift:"     & this.MenuItems(ii).chkShift) & wcomma & _
         qnum("ctrl:"      & this.MenuItems(ii).chkCtrl) & wcomma & _
         qstr("shortcut:"  & this.MenuItems(ii).wszShortcut) & wcomma & _
         qnum("checked:"   & this.MenuItems(ii).chkChecked) & wcomma & _
         qnum("grayed:"    & this.MenuItems(ii).chkGrayed) & _
         "}," 
      next
      wst = rtrim(wst, ",") & "]}"
   end if

   ' Save ToolBar items (if applicable)
   dim as long numToolBarItems = Ubound(this.ToolBarItems) - lbound(this.ToolBarItems) + 1
   if numToolBarItems > 0 then
      wst = wst & wcomma & _
      qnum("toolbar:{") & _
      qnum("display:" & iif(this.GenerateToolBar, "true", "false")) & wcomma & _
      qstr("size:" & this.wszToolBarSize) & wcomma 

      wst = wst & _
      qnum("items: [")
      for ii as long = lbound(this.ToolBarItems) to ubound(this.ToolBarItems)
         wst = wst & _
         "{" & _
         qstr("name:"           & this.ToolBarItems(ii).wszName) & wcomma & _
         qstr("type:"           & this.ToolBarItems(ii).wszButtonType) & wcomma & _
         qstr("tooltip:"        & this.ToolBarItems(ii).wszTooltip) & wcomma & _
         qstr("hotimage:"       & this.ToolBarItems(ii).pPropHotImage.wszPropValue) & wcomma & _
         qstr("normalimage:"    & this.ToolBarItems(ii).pPropNormalImage.wszPropValue) & wcomma & _
         qstr("disabledimage:"  & this.ToolBarItems(ii).pPropDisabledImage.wszPropValue) & _
         "}," 
      next
      wst = rtrim(wst, ",") & "]}"
   end if


   ' Save StatusBar Panels (if applicable)
   dim as long numPanelItems = Ubound(this.PanelItems) - lbound(this.PanelItems) + 1
   if numPanelItems > 0 then
      wst = wst & wcomma & _
      qnum("statusbar:{") & _
      qnum("display:" & iif(this.GenerateStatusBar, "true", "false")) & wcomma 

      wst = wst & _
      qnum("items: [")
      for ii as long = lbound(this.PanelItems) to ubound(this.PanelItems)
         wst = wst & _
         "{" & _
         qstr("name:"           & this.PanelItems(ii).wszName) & wcomma & _
         qstr("text:"           & this.PanelItems(ii).wszText) & wcomma & _
         qstr("tooltip:"        & this.PanelItems(ii).wszTooltip) & wcomma & _
         qstr("alignment:"      & this.PanelItems(ii).wszAlignment) & wcomma & _
         qstr("autosize:"       & this.PanelItems(ii).wszAutoSize) & wcomma & _
         qstr("width:"          & this.PanelItems(ii).wszWidth) & wcomma & _
         qstr("minwidth:"       & this.PanelItems(ii).wszMinWidth) & wcomma & _
         qstr("image:"          & this.PanelItems(ii).pProp.wszPropValue) & wcomma & _
         qstr("backcolor:"      & this.PanelItems(ii).wszBackColor) & wcomma & _
         qstr("backcolorhot:"   & this.PanelItems(ii).wszBackColorHot) & wcomma & _
         qstr("forecolor:"      & this.PanelItems(ii).wszForeColor) & wcomma & _
         qstr("forecolorhot:"   & this.PanelItems(ii).wszForeColorHot) & _
         "}," 
              ' BorderStyle is deprecated as of v2.0.4 as it has no effect
              ' in WinFBE programs where Windows Themes are enabled.
              '"'     PANELBORDERSTYLE=" & this.PanelItems(ii).wszBorderStyle & vbcrlf & _
      next
      wst = rtrim(wst, ",") & "]}"
   end if
   
   
   ' Iterate all of the controls on the form
   if this.Controls.Count then
      wst = wst & wcomma & _
      qnum("controls:[") 
            
      ' We want to ensure that we output code for the Form first because
      ' when the form is loaded the controls need to attach themselves
      ' to the form that must exist otherwise a GPF will occur.
      dim bOutputForm as boolean = true
      for numLoops as long = 1 to 2
         ' First time through the loop we get the Form data. The second time through
         ' we get everything except for the form.
         for i as long = this.Controls.ItemFirst to this.Controls.ItemLast
            pCtrl = this.Controls.ItemAt(i)
            if pCtrl then
               if bOutputForm = true then
                  if pCtrl->ControlType <> CTRL_FORM then continue for
               elseif bOutputForm = false then
                  if pCtrl->ControlType = CTRL_FORM then continue for
               end if

               wst = wst & _
               "{" & _
               qstr("type:" & GetToolBoxName(pCtrl->ControlType)) & wcomma & _
               qnum("properties: [") 

               for ii as long = lbound(pCtrl->Properties) to ubound(pCtrl->Properties)
                  wst = wst & "{" & _
                  qstr("name:" & pCtrl->Properties(ii).wszPropName) & wcomma & _
                  qstr("value:" & pCtrl->Properties(ii).wszPropValue) & _
                  "},"
               next
               
               wst = rtrim(wst, ",") & "]," & _
               qnum("events: [") 

               for ii as long = lbound(pCtrl->Events) to ubound(pCtrl->Events)
                  ' Only need to output the names of the Events that have been ticked as being in use.
                  if pCtrl->Events(ii).bIsSelected then
                     if rtrim(pCtrl->Events(ii).wszEventName) > "" then
                        wst = wst & "{" & _
                        qstr("name:" & pCtrl->Events(ii).wszEventName) & _
                        "},"
                     end if   
                  end if   
               next
               wst = rtrim(wst, ",") & "]"
         
            end if
            
            wst = wst & "},"
         next

         bOutputForm = false
      next
      
      wst = rtrim(wst, ",") & "]"
   end if

   ' Convert UTF-16 to UTF-8 encoding
   wst = "{" & wst & "}"

'   dim pStream as CTextStream  
'   if pStream.Create(this.DesignerFilename, true, false) <> S_OK then return true   ' error
'   pStream.WriteLine( UnicodeToUtf8(wst) )
'   pStream.Close

   ' Per email from "Allan" on Jan 14, 2023, changes I made to UnicodeToUtf8
   ' function (eliminating the ending Trim(0), and use BytesWritten instead) seems to be
   ' the reason for the Chinese characters now writing 100% correctly.
   ' Use the native FreeBasic file handling commands to save the data to disk. This is a test
   ' to see if the CTextStream approach is somehow causing problems with the UTF-8 output. Allan
   ' is still experiencing on his computer whereas I am not. Everything seems to work correctly
   ' on my side. 2023-01-27, Allan confirmed that using the native FB file operations has
   ' corrected his issue with garbled Chinese characters.
   dim as long f = freefile
   if open (this.DesignerFilename for output as #f) = 0 then
      print #f, UnicodeToUtf8(wst)
      close #f
   else
      return true   ' error
   end if
   
   function = false
end function


''
''
function clsDocument.SaveFile( _
            byval bSaveAs as boolean = false, _
            byval bAutoSaveOnly as boolean = false _
            ) as boolean

   dim sText        as string
   dim wszFilename  as wstring * MAX_PATH
   dim wszAutoSaveFilename  as wstring * MAX_PATH
   dim wszExtension as wstring * MAX_PATH
   dim wszText      as wstring * MAX_PATH  
   dim wst          as CWSTR
   
   ' don't allow the AutoSave timer to be re-entrant on this function
   ' while a save is currently taking place.
   static bAutoSaveIsActive as boolean
   if bAutoSaveOnly andalso bAutoSaveIsActive then exit function
   bAutoSaveIsActive = true
   
   ' if this is a new Untitled document then set flag to display SaveAs dialog.
   if this.IsNewFlag then bSaveAs = true
      
   ' Don't let AutoSave allow bSaveAs
   if bAutoSaveOnly then bSaveAs = false

   wszFilename = this.Diskfilename
   wszAutoSaveFilename = this.AutoSaveFilename
      
   if bSaveAs then
      ' if this is a new Form being saved for the first time then the extension should
      ' be .bas if no project is active, and .inc if a project is active.
      if this.IsNewFlag then
         if this.IsDesigner then
            wszExtension = iif(gApp.IsProjectActive, "inc", "bas")
         end if   
      end if   
      ' Display the Save File Dialog
      dim pwszName as wstring ptr = AfxIFileSaveDialog(HWND_FRMMAIN, @wszFilename, @wszExtension, IDM_FILESAVEAS)
      if pwszName then
         wszFilename = *pwszName
         CoTaskMemFree(pwszName)
         this.bNeedsParsing = true
      else
         return false
      end if
   end if

   dim as any ptr pSci = this.GetActiveScintillaPtr()
   
   ' if pSci does not exist then this file exists in the Explorer but has not yet been
   ' displayed in order to create a Scintilla window. Simply need to exit.
   if pSci = 0 then return true
   
   ' Save text buffer to disk by directly accessing buffer rather
   ' saving it to an intermediary string variable first.
   dim as zstring ptr psz = cast( zstring ptr, SciExec(this.hWindow(0), SCI_GETCHARACTERPOINTER, 0, 0) )
   dim as long sciCodePage = SciMsg(pSci, SCI_GETCODEPAGE, 0, 0)   ' 0 or SC_CP_UTF8 
   
   dim as CWSTR wszSaveFilename
   if bAutoSaveOnly then 
      if AfxFileExists(wszAutoSaveFilename) then AfxDeleteFile(wszAutoSaveFilename)
      wszSaveFilename = wszAutoSaveFilename
   else   
      ' If we are upgrading to v3.02+ then save the original file because the original
      ' file format has visual designer and code all in one file.
      ' Save a copy of the original file to the backup in the event that there is
      ' issues with the upgrade process. Fail if backup already exists.
      if this.IsDesigner then
         if (ConvertWinFBEversion(this.wszFormVersion) >= ConvertWinFBEversion("3.0.2")) orelse _
            (ConvertWinFBEversion(this.wszFormVersion) = 0) then  'b/c upgraded form files do not have version loaded yet
            ' We are already using new form file format
         else   
            dim as CWSTR wszBackupFilename = this.DiskFilename & ".backup-before-upgrade"
            AfxCopyFile( this.DiskFilename, wszBackupFilename, CTRUE )   
         end if
      end if   

      ' ensure any previously created AutoSave files are delete should 
      ' the user had chosen a new filename
      if AfxFileExists(wszFilename) then AfxDeleteFile(wszFilename)
      if AfxFileExists(wszAutoSaveFilename) then AfxDeleteFile(wszAutoSaveFilename)
      wszSaveFilename = wszFilename
   end if   
      
   
   dim pStream as CFileStream
   if pStream.Open(wszSaveFilename, STGM_CREATE or STGM_WRITE) = S_OK then
      dim as string st
      select case this.FileEncoding
         case FILE_ENCODING_ANSI
            if sciCodePage = 0 then    
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
            if sciCodePage = SC_CP_UTF8 then    
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
            if sciCodePage = SC_CP_UTF8 then    
               ' convert utf8 to utf16
               wst.Utf8 = *psz  ' convert the utf *psz from Utf8
               pStream.Write wst.m_pBuffer, wst.m_BufferLen
            else
               ' need to convert ansi to unicode
               dim as CWSTR wst = Wstr(*psz)
               pStream.Write wst.m_pBuffer, wst.m_BufferLen
            end if    
            
      end SELECT
   end if
   pStream.Close

   this.DiskFilename = wszFilename
   this.DesignerFilename = wszFilename & ".design"
   this.AutoSaveFilename = OnCommand_FileAutoSaveGenerateFilename(wszFilename)
   this.DateFileTime = AfxGetFileLastWriteTime( wszFilename )
   this.AutoSaveRequired = false
   
   if this.IsDesigner then this.SaveFormJSONdata()
   
   ' if this was a new document then it needs to be saved to Recent File list.
   if this.IsNewFlag then 
      if gApp.IsProjectActive = false then 
         UpdateMRUList(wszFilename)
      end if
   end if
   this.IsNewFlag = false
   
   ' Set the current state of the document to unmodified
   if bAutoSaveOnly = false then 
      this.UserModified = false   
      SciMsg( pSci, SCI_SETSAVEPOINT, 0, 0)
   end if
   
   bAutoSaveIsActive = false
   
   function = true
end function


''
''
function clsDocument.GetTextRange( _
            byval cpMin as long, _
            byval cpMax as long _
            ) as string
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   dim p      as long
   dim buffer as string
   dim txtrg  as SCI_TEXTRANGE
   txtrg.chrg.cpMin = cpMin
   txtrg.chrg.cpMax = cpMax
   buffer = space(cpMax - cpMin + 1)
   txtrg.lpstrText = strptr(buffer)
   SciMsg(pSci, SCI_GETTEXTRANGE, 0, cast(LPARAM, @txtrg))
   p = instr(buffer, chr(0))
   if p then buffer = left(buffer, p - 1)
   function = buffer
end function

''
''
function clsDocument.ChangeSelectionCase( byval fcase as long) as long 
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   dim startSelPos as long     ' Starting position
   dim endSelPos   as long     ' Ending position
   dim strText     as string   ' Selected text
   dim i           as long 

   ' fcase = 1 (upper case), 2 (lower case), 3 (mixed case)
   if (fcase < 1) or (fcase > 3) then exit function
   
   ' if startSelPos and endSelPos are the same there is not selection,
   startSelPos = SciMsg( pSci, SCI_GETSELECTIONSTART, 0, 0)
   endSelPos   = SciMsg( pSci, SCI_GETSELECTIONEND, 0, 0)
   if startSelPos = endSelPos then exit function
   
   ' Retrieve the text
   strText = this.GetTextRange(startSelPos, endSelPos)
   
   ' Convert it to upper, lower case, or mixed case
   if fcase = 1 then
      strText = ucase(strText)
   elseif fcase = 2 then
      strText = LCase(strText)
   elseif fcase = 3 then
      ' Always uppercase the first character regardless
      strText = LCase(strText)
      mid(strText,1,1) = ucase(mid(strText,1,1))
      dim as string prevChar
      for i as long = 2 to len(strText)
         prevChar = mid(strText,i-1,1)
         select case prevChar
            case chr(13), chr(10), " "
               mid(strText,i,1) = ucase(mid(strText,i,1))
         end select
      next
   end if
   
   ' Replace the selected text
   SciMsg( pSci, SCI_REPLACESEL, 0, cast(LPARAM, strptr(strText)))

   function = 0
end function

''
''
function clsDocument.SetMarkerHighlight() as long
   ' Set a marker that will highlight the background of the current selection. This
   ' is used when we are attempting to search a selection. We want the current search
   ' area to be a different color than the regular highlighted text because any
   ' search results are colored using the normal highlight colors so we need them
   ' to stand out from the selected range.
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   SciMsg( pSci, SCI_MARKERDELETEALL, 10, 0) ' Remove any existing before drawing the new highlight
   dim as long startPos, endPos, startLine, endLine
   this.GetSelectedLineRange( startLine, endLine, startPos, endPos )
   if endLine <> startLine then
      SciMsg( pSci, SCI_MARKERDEFINE, 10, SC_MARK_BACKGROUND )  ' define as marker #10
      SciMsg( pSci, SCI_SETMARGINMASKN, 4, &H400 )   ' set margin mask to allow SC_MARK_BACKGROUND 
      for i as long = startLine to endLine
         function = SciMsg( pSci, SCI_MARKERADD, i, 10)    ' add defined marker #10 to each line
      next
      ' set backcolor of marker #10 
      dim as COLORREF clr = ghEditor.BackColorSelection
      SciMsg( pSci, SCI_MARKERSETBACK, 10, clr) 
   end if
end function

''
''
function clsDocument.RemoveMarkerHighlight() as long
   ' Remove any markers that were set in the document that signify a highlighted range.
   ' This is used when we are attempting to search a selection. 
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   function = SciMsg( pSci, SCI_MARKERDELETEALL, 10, 0)  ' delete all marker #10
end function

''
''
function clsDocument.FirstMarkerHighlight() as long
   ' Get the first line with marker #10 highlight
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   dim as long markerMask 
   markerMask = bitset(markerMask, 10)
   function = SciMsg( pSci, SCI_MARKERNEXT, 0, markerMask)
end function

''
''
function clsDocument.LastMarkerHighlight() as long
   ' Get the first line with marker #10 highlight
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   dim as long nLastPos = SciMsg( pSci, SCI_GETTEXTLENGTH, 0, 0)
   dim as long nLastLine = SciMsg( pSci, SCI_LINEFROMPOSITION, nLastPos, 0)
   dim as long markerMask 
   markerMask = bitset(markerMask, 10)
   function = SciMsg( pSci, SCI_MARKERPREVIOUS, nLastLine, markerMask)
end function

''
''
function clsDocument.HasMarkerHighlight() as boolean
   ' true/False if selection markers exist in the document search for marker #10
   function = iif(this.FirstMarkerHighlight = -1, false, true)
end function

''
''
function clsDocument.GetCurrentLineNumber() as long
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   dim nPos as long = SciMsg( pSci, SCI_GETCURRENTPOS, 0, 0)
   function = SciMsg( pSci, SCI_LINEFROMPOSITION, nPos, 0)
end function

''
''
function clsDocument.SelectLine( byval nLineNum as long ) as long
   ' select the incoming nLineNum. if nLineNum is negative then select the current line
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   nLineNum = iif( nLineNum < 0, this.GetCurrentLineNumber, nLineNum)
   dim nStartPos as long = SciMsg( pSci, SCI_POSITIONFROMLINE, nLineNum, 0)
   dim nEndPos   as long = SciMsg( pSci, SCI_GETLINEENDPOSITION, nLineNum, 0)
   SciMsg( pSci, SCI_SETSELECTIONSTART, nStartPos, 0)
   SciMsg( pSci, SCI_SETSELECTIONEND, nEndPos, 0)
   function = 0
end function

''
''
function clsDocument.GetLine( byval nLine as long ) as string
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   dim nLen   as long
   dim buffer as string
   nLen = SciMsg( pSci, SCI_LINELENGTH, nLine , 0)
   if nLen < 1 then exit function
   buffer = space(nLen)
   SciMsg( pSci, SCI_GETLINE, nLine, cast(LPARAM, strptr(buffer)))
   function = rtrim(buffer, any chr(13,10))
end function

''
''
function clsDocument.SetLine( byval nLineNum as long, byval sText as string) as long
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   dim nStartPos as long = SciMsg( pSci, SCI_POSITIONFROMLINE, nLineNum, 0)
   dim nEndPos   as long = SciMsg( pSci, SCI_GETLINEENDPOSITION, nLineNum, 0)
   SciMsg( pSci, SCI_SETTARGETSTART, nStartPos, 0)
   SciMsg( pSci, SCI_SETTARGETEND, nEndPos, 0)
   SciMsg( pSci, SCI_REPLACETARGET, len(sText), cast(LPARAM, strptr(sText))) 
   function = 0
end function

''
''
function clsDocument.GetSelText() as string
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   dim nLen as long
   dim buffer as string
   nLen = SciMsg( pSci, SCI_GETSELTEXT, 0, 0)
   if nLen < 1 then exit function
   buffer = space(nLen)
   SciMsg( pSci, SCI_GETSELTEXT, 0, cast(LPARAM, strptr(buffer)))
   function = buffer
end function

''
''
function clsDocument.GetText() as string
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   dim nLen   as long
   dim buffer as string
   nLen = SciMsg( pSci, SCI_GETLENGTH, 0 , 0)
   if nLen < 1 then exit function
   buffer = space(nLen+1)
   SciMsg( pSci, SCI_GETTEXT, nLen+1, cast(LPARAM, strptr(buffer)) )
   function = buffer
end function

''
''
function clsDocument.SetText( byref sText as const string ) as long 
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   SciMsg( pSci, SCI_SETTEXT, 0, cast(LPARAM, strptr(sText)) )
   SciMsg( pSci, SCI_COLOURISE, 0, -1 )
   function = 0
end function


''
''
function clsDocument.AppendText( byref sText as const string ) as long 
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   SciMsg( pSci, SCI_APPENDTEXT, len(sText), cast(LPARAM, strptr(sText)) )
   SciMsg( pSci, SCI_COLOURISE, 0, -1 )
   function = 0
end function


''
''
function clsDocument.CenterCurrentLine() as long 
   ' Center the current line to the middle of the visible screen. This is useful
   ' when searching and finding text. The found text will always display in the 
   ' middle of the screen. 
   if gConfig.PositionMiddle = 0 then exit function
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   dim as long numLinesOnScreen = SciMsg( pSci, SCI_LINESONSCREEN, 0, 0)
   dim as long nMiddle = (this.GetCurrentLineNumber - (numLinesOnScreen / 2) + 1)
   SciMsg( pSci, SCI_SETFIRSTVISIBLELINE, nMiddle, 0)
   function = 0
end function



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
      
   if left( lineText, 9 ) = "FUNCTION " then
      ' bypass any line with = that is a "function = " line.
      dim as string subText = ltrim( mid(lineText, 9) )
      if left( subText, 1 ) <> "=" then position = 10
   elseif left(lineText, 17) = "PRIVATE FUNCTION " then
      position = 18   
   elseif left(lineText, 16) = "PUBLIC FUNCTION " then
      position = 17
   elseif left(lineText, 4)  = "SUB " then
      position = 5
   elseif left(lineText, 12) = "PRIVATE SUB " then
      position = 13
   elseif left(lineText, 11) = "PUBLIC SUB " then
      position = 12
   elseif left(lineText, 9)  = "PROPERTY " then
      ' bypass any line with = that is a "PROPERTY = " line.
      dim as string subText = ltrim( mid(lineText, 9) )
      if left( subText, 1 ) <> "=" then position = 10
   elseif left(lineText, 17) = "PRIVATE PROPERTY " then
      position = 18
   elseif left(lineText, 12) = "CONSTRUCTOR " then
      position = 13
   elseif left(lineText, 11) = "DESTRUCTOR " then
      position = 12
      
   ' if we encounter one of these end statements then our cursor must 
   ' be positioned between functions (ie we're at the module level).   
   elseif left(lineText, 12) = "END FUNCTION" then
      position = -1
   elseif left(lineText, 7)  = "END SUB" then
      position = -1
   elseif left(lineText, 15) = "END CONSTRUCTOR" then
      position = -1
   elseif left(lineText, 14) = "END DESTRUCTOR" then
      position = -1
   elseif left(lineText, 12) = "END PROPERTY" then
      position = -1

   end if

   function = position
end function


' ========================================================================================
' Determine the sub/function name based on the current editing position 
' within the file. This is needed by CodeTips (DereferenceLine) and for code navigation
' in the editor to set the functions ComboBox to the correct item.
' ========================================================================================
function clsDocument.GetCurrentfunctionName( _
            byref sfunctionName as string, _
            byref nGetSet as ClassProperty _
            ) as long
    
   dim as string lineText
   dim as string funcName 
   dim as string funcParams
   
   dim as long curLine
   dim as long position
   
   ' search up the file until we find the start of a sub/function or start of file.
   curLine = this.GetCurrentLineNumber
   nGetSet = ClassProperty.None
   
   for i as long = curLine to 0 step -1
      position = this.IsfunctionLine( i )
      
      if position = -1 then 
         ' We found an "end SUB", "end function", "end PROPERTY" line. Keep looking for 
         ' the very next instance of SUB/function and then break out of loop
         ' with that function name.
         if i = this.GetCurrentLineNumber then 
            continue for
         else
            exit for
         end if
            
      elseif position = 0 then
         ' Just a regular line... keep looking...
      
      elseif position > 0 then
         ' We found a valid SUB/function line so process it
         ' trim and remove double spaces and replace TABs with single space
         lineText = ltrim(ucase( AfxStrShrink(this.GetLine(i), chr(32,9)) ))
         funcName = ltrim(mid(lineText, position))
         funcName = AfxStrParseAny( funcName, 1, " (" )
         
         ' if this is a Property then we need to differentiate between a Get/Set
         if ( left(lineText, 9)  = "PROPERTY " ) orelse _
            ( left(lineText, 17) = "PRIVATE PROPERTY " ) then
            ' if funcParams exist then this must be a Set property. Need to sanitize the
            ' property parameters first. Need to get the starting ( and the ending ) and
            ' evaluate the text between it. Not as easy as using Parse because there could
            ' be embedded array() parameters and the property could end in something like 
            ' as long, etc.
            dim as string st
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

         elseif ( left(lineText, 12)  = "CONSTRUCTOR " ) then
            nGetSet = ClassProperty.ctor
            
         elseif ( left(lineText, 11)  = "DESTRUCTOR " ) then
            nGetSet = ClassProperty.dtor
         end if   

         exit for
      end if
   next

   sfunctionName = funcName
   nGetSet = nGetSet
   
   function = 0
end function


''
''  GotoNextfunction
''
''  Go to the next sub/function in the document
''
function clsDocument.GotoNextFunction() as long
   dim as long curLine  = this.GetCurrentLineNumber
   dim as long maxLines = this.GetLineCount - 1
   dim as long newLine  = curLine
   
   for i as long = curLine + 1 to maxLines
      if this.IsfunctionLine( i ) > 0 then
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
''  GotoPrevfunction
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
function clsDocument.GetSelectedLineRange( _
            byref startLine as long, _
            byref endLine   as long, _
            byref startPos  as long, _
            byref endPos    as long _
            ) as long 
   
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   startPos  = SciMsg( pSci, SCI_GETSELECTIONSTART, 0, 0) 
   endPos    = SciMsg( pSci, SCI_GETSELECTIONEND, 0, 0) 
   startLine = SciMsg( pSci, SCI_LINEFROMPOSITION, startPos, 0) 
   endLine   = SciMsg( pSci, SCI_LINEFROMPOSITION, endPos, 0) 

   dim nCol as long = SciMsg( pSci, SCI_GETCOLUMN, endPos, 0)
   if (nCol = 0) and (endLine > startLine) then endLine = endLine - 1

   function = 0
end function


''
''
function clsDocument.BlockComment( byval flagBlock as boolean ) as long
   dim i           as long        ' loop counter
   dim firstPos    as long        ' Starting position of the line
   dim startPos    as long        ' Starting position of selection
   dim endPos      as long        ' Ending position of selection
   dim startLine   as long        ' Starting line number
   dim endLine     as long        ' Ending line number
   dim nPos        as long        ' Position
   dim strText     as string      ' Portion of the line to replace
   dim strNewText  as string      ' new text to insert or delete
   dim nCount      as long        ' Number of "'" added or removed
   
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   ' flagBlock = true for Blocking comment, false for UnBlocking comment
   this.GetSelectedLineRange( startLine, endLine, startPos, endPos )
   
   gApp.SuppressNotify = true
   SetWindowRedraw( this.hWndActiveScintilla, false )
   SciMsg( pSci, SCI_BEGINUNDOACTION, 0, 0)
   for i = startLine to endLine
      strText = this.GetLine(i)   
      if flagBlock = false then   
         ' unblock comment
         if left(strText, 1) <> "'" then
            continue For
         else   
            firstpos = SciMsg( pSci, SCI_POSITIONFROMLINE, i, 0 )
            SciMsg( pSci, SCI_DELETERANGE, firstpos, cast(LPARAM, 1) )
         end if
         nCount += 1
      else
         ' block comment
         if len(rtrim(strText)) then
            strNewText = "'"
            firstpos = SciMsg( pSci, SCI_POSITIONFROMLINE, i, 0 )
            SciMsg( pSci, SCI_INSERTTEXT, firstpos, cast(LPARAM, strptr(strNewText)) )
            nCount += 1
         end if   
      end if
   next
   SciMsg( pSci, SCI_ENDUNDOACTION, 0, 0)

   if startPos <> endPos then
      SciMsg( pSci, SCI_SETSELECTIONSTART, startPos, 0)
      SciMsg( pSci, SCI_SETSELECTIONEND, endPos + iif(flagBlock, nCount, -nCount), 0)
   else
      SciMsg( pSci, SCI_SETSELECTIONSTART, endPos + iif(flagBlock, nCount, -nCount), 0)
      SciMsg( pSci, SCI_SETSELECTIONEND, endPos + iif(flagBlock, nCount, -nCount), 0)
   end if

   SetWindowRedraw( this.hWndActiveScintilla, true )
   gApp.SuppressNotify = false
   AfxRedrawWindow( this.hWndActiveScintilla )
   
   function = 0
end function


''
''
function clsDocument.CurrentLineUp() as long
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   dim currentLine as long = this.GetCurrentLineNumber()
   if (currentLine <> 0) then
      SciMsg( pSci, SCI_BEGINUNDOACTION, 0, 0)
      currentLine = currentLine -1
      SciMsg( pSci, SCI_LINETRANSPOSE, 0, 0)
      SciMsg( pSci, SCI_GOTOLINE, currentLine, 0)
      SciMsg( pSci, SCI_ENDUNDOACTION, 0, 0)
   end if
   function = 0
end function


''
''
function clsDocument.GetLineCount() as long
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   function = SciMsg( pSci, SCI_GETLINECOUNT, 0, 0)
end function


''
''
function clsDocument.CurrentLineDown() as long
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   dim currentLine as long = this.GetCurrentLineNumber()
   if currentLine <> SciMsg( pSci, SCI_GETLINECOUNT, 0, 0) - 1 then
      SciMsg( pSci, SCI_BEGINUNDOACTION, 0, 0)
      currentLine = currentLine + 1
      SciMsg( pSci, SCI_GOTOLINE, currentLine, 0)
      SciMsg( pSci, SCI_LINETRANSPOSE, 0, 0)
      SciMsg( pSci, SCI_ENDUNDOACTION, 0, 0)
   end if
   function = 0
end function
  
''
''
function clsDocument.MoveCurrentLines( byval flagMoveDown as boolean ) as long
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   if flagMoveDown then
      SciMsg( pSci, SCI_MOVESELECTEDLINESDOWN, 0, 0)
   else   
      SciMsg( pSci, SCI_MOVESELECTEDLINESUP, 0, 0)
   end if
   function = 0
end function


''
''
function clsDocument.NewLineBelowCurrent() as long
   ' From anywhere on the current line creates a new line immediately
   ' below the current line and positions the cursor to the start of
   ' that new line.
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   SciMsg( pSci, SCI_LINEEND, 0, 0)
   SciMsg( pSci, SCI_NEWLINE, 0, 0)
   function = 0
end function


''
''
function clsDocument.ToggleBookmark( byval nLine as long ) as long
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   dim fMark as long  ' must be a 32 bit value
   fMark = SciMsg( pSci, SCI_MARKERGET, nLine, 0 ) 
   if bit(fMark, 0) = -1 then
      SciMsg( pSci, SCI_MARKERDELETE, nLine, 0 )
   else
      SciMsg( pSci, SCI_MARKERADD, nLine, 0 )
   end if
   function = 0
end function

''
''
function clsDocument.NextBookmark() as long 
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   dim fMark as long  ' 32 bit value
   dim nCurLine as long = this.GetCurrentLineNumber() 
   dim nLine as long = this.GetCurrentLineNumber() + 1  ' start line
   fMark = bitset(fMark, 0)
   nLine = SciMsg( pSci, SCI_MARKERNEXT, nLine, fMark)
   if nLine > -1 then
      SciMsg( pSci, SCI_GOTOLINE, nLine, 0)
   else
      nLine = SciMsg( pSci, SCI_MARKERNEXT, nLine, fMark)
      if nLine > -1 then
         SciMsg( pSci, SCI_GOTOLINE, nLine, 0)
      end if
   end if
   if nLine <> -1 then
      if nLine <> nCurLine then this.CenterCurrentLine()
   end if
   function = 0 
end function

''
''
function clsDocument.PrevBookmark() as long 
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   dim fMark  as long  ' 32 bit value
   dim nCurLine as long = this.GetCurrentLineNumber() 
   dim nLine  as long = this.GetCurrentLineNumber() - 1   ' start line
   dim nLines as long = SciMsg( pSci, SCI_GETLINECOUNT, 0, 0) - 1
   fMark = bitset(fMark, 0)
   nLine = SciMsg( pSci, SCI_MARKERPREVIOUS, nLine, fMark)
   if nLine > -1 then
      SciMsg( pSci, SCI_GOTOLINE, nLine, 0)
   else
      nLine = SciMsg( pSci, SCI_MARKERPREVIOUS, nLines, fMark)
      if nLine > -1 then
         SciMsg( pSci, SCI_GOTOLINE, nLine, 0)
      end if
   end if
   if nLine <> -1 then
      if nLine <> nCurLine then this.CenterCurrentLine()
   end if
   function = 0 
end function

''
''   
function clsDocument.FoldToggle( byval nLine as long ) as long
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   dim nFoldLevel as long = SciMsg( pSci, SCI_GETFOLDLEVEL, nLine, 0)

   if (nFoldLevel and SC_FOLDLEVELHEADERFLAG) = 0 then
      ' Get the number of the head line of the procedure or function
      nLine = SciMsg( pSci, SCI_GETFOLDPARENT, nLine, 0) 
   end if
   if nLine > -1 then
      SciMsg( pSci, SCI_TOGGLEFOLD, nLine, 0) 
      SciMsg( pSci, SCI_GOTOLINE, nLine, 0)
   end if

   function = nLine
end function

''
''
function clsDocument.FoldAll() as long

   dim i          as long    
   dim nLines     as long    
   dim nFoldLevel as long
   
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   ' Force the lexer to style the whole document
   SciMsg( pSci, SCI_COLOURISE, -1, 0)

   nLines = SciMsg( pSci, SCI_GETLINECOUNT, 0, 0)

   for i = 0 to nLines
      ' if we are in the head line ...
      nFoldLevel = SciMsg( pSci, SCI_GETFOLDLEVEL, i, 0)
      if (nFoldLevel and SC_FOLDLEVELNUMBERMASK) = SC_FOLDLEVELBASE then
         if SciMsg( pSci, SCI_GETFOLDEXPANDED, i, 0) then
            SciMsg( pSci, SCI_TOGGLEFOLD, i, 0) 
         end if
      end if
   next

   function = 0
end function

''
''
function clsDocument.UnFoldAll() as long
   dim i          as long    
   dim nLines     as long    
   dim nFoldLevel as long
   
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   ' Force the lexer to style the whole document
   SciMsg( pSci, SCI_COLOURISE, -1, 0 )

   nLines = SciMsg( pSci, SCI_GETLINECOUNT, 0, 0)

   for i = 0 to nLines
      ' if we are in the head line ...
      nFoldLevel = SciMsg( pSci, SCI_GETFOLDLEVEL, i, 0)
      if (nFoldLevel and SC_FOLDLEVELNUMBERMASK) = SC_FOLDLEVELBASE then
         if SciMsg( pSci, SCI_GETFOLDEXPANDED, i, 0) = 0 then
            SciMsg( pSci, SCI_TOGGLEFOLD, i, 0) 
         end if
      end if
   next

   function = 0
end function

''
''
function clsDocument.FoldToggleOnwards( byval nLine as long) as long
   ' Toggles the curent fold point and all folds below it that are of greater depth
   dim i          as long    
   dim nLines     as long    
   dim nFoldLevel as long
   dim nFoldLevelBase as long
   dim FoldState  as long
   
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   ' Force the lexer to style the whole document
   SciMsg( pSci, SCI_COLOURISE, -1, 0 )

   nLines = SciMsg( pSci, SCI_GETLINECOUNT, 0, 0)

   ' Toggle the first sub or function
   nLine = this.FoldToggle(nLine)
   nFoldLevelBase = SciMsg( pSci, SCI_GETFOLDLEVEL, nLine, 0)
   
   ' Determine whether the fold is expanded or not
   FoldState = SciMsg( pSci, SCI_GETFOLDEXPANDED, nLine, 0)

   for i = nLine to nLines
      nFoldLevel = SciMsg( pSci, SCI_GETFOLDLEVEL, i, 0)
      if nFoldLevel > nFoldLevelBase then
         ' if the state is different ...
         if SciMsg( pSci, SCI_GETFOLDEXPANDED, i, 0) <> FoldState then
            SciMsg( pSci, SCI_TOGGLEFOLD, i, 0) 
         end if
      end if
   next

   function = 0
end function

''
''
function clsDocument.ConvertEOL( byval nMode as long) as long
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   function = SciMsg( pSci, SCI_CONVERTEOLS, nMode, 0)
   SciMsg( pSci, SCI_SETEOLMODE, nMode, 0)
end function

''
''
function clsDocument.TabsToSpaces() as long
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   dim as long i, n, nLen, nLines, TabSize, nLineNumber 
   dim as string strText, strBuffer

   ' Get the current line
   nLineNumber = this.GetCurrentLineNumber
   ' Get the tab size
   TabSize = SciMsg( pSci, SCI_GETTABWIDTH, 0, 0)
   if TabSize < 1 then exit function
   ' Get the length of the text
   nLen  = SciMsg( pSci, SCI_GETTEXTLENGTH, 0, 0) 
   ' Get the number of lines
   nLines = SciMsg( pSci, SCI_GETLINECOUNT, 0, 0)
   ' Remove tabs, line by line
   for i = 0 to nLines - 1
      strText = this.GetLine(i)
      n = 1
      Do
         n = instr(n, strText, chr(9))
         if n > 0 then 
            strText = left(strText, n - 1) & space(TabSize) & mid(strText, n + 1)
            n += 1
         end if   
      loop until n = 0
      strBuffer = strBuffer & strText & chr(13,10)
   next
   ' Set the new text
   this.SetText(strBuffer)
   ' Set the caret position
   SciMsg( pSci, SCI_GOTOLINE, nLineNumber, 0) 

   function = 0
end function


''
''
function clsDocument.IsMultiLineSelection() as boolean
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   dim as long startPos, endPos, startLine, endLine
   startPos  = SciMsg( pSci, SCI_GETSELECTIONSTART, 0, 0) 
   endPos    = SciMsg( pSci, SCI_GETSELECTIONEND, 0, 0) 
   startLine = SciMsg( pSci, SCI_LINEFROMPOSITION, startPos, 0) 
   endLine   = SciMsg( pSci, SCI_LINEFROMPOSITION, endPos, 0) 
   if endLine <> startLine then return true
end function


''
''
function clsDocument.ApplyProperties() as long
   dim nCount      as long 
   dim i           as long
   dim nPixels     as long  
   dim bitsNeeded  as long 
   dim wFileExt    as wstring * MAX_PATH
   dim strFontName as string 
   dim nFontExtraSpace as long  
   dim nFontSize   as long  
   dim nFontcase   as long 
   dim rxRatio     as single = 1
   dim ryRatio     as single = 1

   
   if m_pSci(0) = 0 then exit function

   ' Determine the pWindow parent of the Scintilla window in order
   ' to ensure that DPI ratios are correctly used.
   dim pWindow as CWindow ptr = AfxCWindowOwnerPtr(this.hWindow(0))
   if pWindow then
      rxRatio = pWindow->rxRatio
      ryRatio = pWindow->ryRatio
   end if   
   
   ' if this is a read-only file then set the flag so that the document can not be edited.
   if AfxIsReadOnlyFile( this.DiskFilename ) then
      SciMsg( m_pSci(0), SCI_SETREADONLY, 1, 0)
   else
      SciMsg( m_pSci(0), SCI_SETREADONLY, 0, 0)
   end if

   strFontName = str(gConfig.EditorFontname)
   nFontSize   = val(gConfig.EditorFontsize)
   nFontExtraSpace = val(gConfig.FontExtraSpace)
   
   select case gConfig.KeywordCase
      case 0:  nFontcase = SC_CASE_LOWER
      case 1:  nFontcase = SC_CASE_UPPER
      case 2:  nFontcase = SC_CASE_CAMEL   
      case 3:  nFontcase = SC_CASE_MIXED    ' original case
   end select            
   
   
   ' Must apply all settings/styles to each Scintilla split window to ensure that
   ' they all appear and act the same. 
   
   for i as long = lbound(m_pSci) to ubound(m_pSci)

      ' Add Brace Highlighting functionality:
      ' First, we set the style of the indicator number we want to use to the Box style
      SciMsg( m_pSci(i), SCI_INDICSETSTYLE, 9, INDIC_STRAIGHTBOX )
      ' then, we make that indicator the current one
      SciMsg( m_pSci(i), SCI_SETINDICATORCURRENT, 9, 0 )
      ' Lastly, we apply the indicator (in this case), to a single char
      SciMsg( m_pSci(i), SCI_BRACEHIGHLIGHTINDICATOR, 1, 9 )  ' use indicator for SCI_BRACEHIGHLIGHTINDICATOR...
      SciMsg( m_pSci(i), SCI_BRACEBADLIGHTINDICATOR, 1, 9 )   ' and SCI_BRACEBADLIGHTINDICATOR

      SciMsg( m_pSci(i), SCI_STYLESETFONT, STYLE_DEFAULT, cast(LPARAM, strptr(strFontName)) )
      SciMsg( m_pSci(i), SCI_STYLESETSIZE, STYLE_DEFAULT, nFontSize )
      SciMsg( m_pSci(i), SCI_SETEXTRAASCENT, nFontExtraSpace, 0 )
      SciMsg( m_pSci(i), SCI_SETEXTRADESCENT, nFontExtraSpace, 0 )
      SciMsg( m_pSci(i), SCI_STYLESETCHARACTERSET, STYLE_DEFAULT, GetFontCharSetID(gConfig.EditorFontCharset) )

      SciMsg( m_pSci(i), SCI_STYLESETFORE, STYLE_DEFAULT, ghEditor.ForeColorText)
      SciMsg( m_pSci(i), SCI_STYLESETBACK, STYLE_DEFAULT, ghEditor.BackColorText)
      SciMsg( m_pSci(i), SCI_STYLESETBOLD, STYLE_DEFAULT, ghEditor.TextBold )
      SciMsg( m_pSci(i), SCI_STYLESETITALIC, STYLE_DEFAULT, ghEditor.TextItalic )
      SciMsg( m_pSci(i), SCI_STYLESETUNDERLINE, STYLE_DEFAULT, ghEditor.TextUnderline )
      SciMsg( m_pSci(i), SCI_STYLECLEARALL, 0, 0 )  ' Copies global style to all others
          
      ' Set the style for the AutoComplete popup list
      SciMsg( m_pSci(i), SCI_STYLESETFONT, STYLE_AUTOCOMPLETE, cast(LPARAM, pWindow->DefaultFontName) )
      SciMsg( m_pSci(i), SCI_STYLESETSIZE, STYLE_AUTOCOMPLETE, pWindow->DefaultFontSize)
      SciMsg( m_pSci(i), SCI_STYLESETCHARACTERSET, STYLE_AUTOCOMPLETE, GetFontCharSetID(gConfig.EditorFontCharset) )

      ''
      ''  MARGIN 0: Line Numbering (defaults to width 0)
      nPixels = SciMsg( m_pSci(i), SCI_TEXTWIDTH, 0, cast(LPARAM, @"_99999"))
      SciMsg( m_pSci(i), SCI_SETMARGINTYPEN, 0, SC_MARGIN_NUMBER )
      SciMsg( m_pSci(i), SCI_STYLESETFORE, STYLE_LINENUMBER, ghEditor.ForeColorLinenumbers )
      SciMsg( m_pSci(i), SCI_STYLESETBACK, STYLE_LINENUMBER, ghEditor.BackColorLinenumbers )
      SciMsg( m_pSci(i), SCI_SETMARGINWIDTHN, 0, iif(gConfig.LineNumbering, nPixels, 0) )
      SciMsg( m_pSci(i), SCI_STYLESETBOLD, STYLE_LINENUMBER, ghEditor.LineNumbersBold )
      SciMsg( m_pSci(i), SCI_STYLESETITALIC, STYLE_LINENUMBER, ghEditor.LineNumbersItalic )
      SciMsg( m_pSci(i), SCI_STYLESETUNDERLINE, STYLE_LINENUMBER, ghEditor.LineNumbersUnderline )
   
      ''
      ''  MARGIN 1: Non-Folding symbols (defaults to width 16) (Bookmark symbol, etc) (will be same color as line numbering)
      SciMsg( m_pSci(i), SCI_SETMARGINTYPEN, 1, SC_MARGIN_TEXT )
      SciMsg( m_pSci(i), SCI_SETMARGINSENSITIVEN, 1, 1 )
      SciMsg( m_pSci(i), SCI_SETMARGINWIDTHN, 1, iif(gConfig.LeftMargin, 16 * rxRatio, 0) )
                 
      ''
      ''  MARGIN 2: Folding symbols (defaults to width 0)
      SciMsg( m_pSci(i), SCI_SETMARGINTYPEN, 2, SC_MARGIN_SYMBOL )
      SciMsg( m_pSci(i), SCI_SETMARGINMASKN, 2, SC_MASK_FOLDERS )
      SciMsg( m_pSci(i), SCI_SETFOLDMARGINCOLOUR, Ctrue, ghEditor.ForeColorFoldmargin )
      SciMsg( m_pSci(i), SCI_SETFOLDMARGINHICOLOUR, Ctrue, ghEditor.ForeColorFoldmargin )
      SciMsg( m_pSci(i), SCI_SETMARGINSENSITIVEN, 2, 1 )
      SciMsg( m_pSci(i), SCI_SETMARGINWIDTHN, 2, iif(gConfig.FoldMargin, 16 * rxRatio, 0) )

      ''
      ''  MARGIN 3: Small margin to offset left margins from actual text (4 pixels)
      SciMsg( m_pSci(i), SCI_SETMARGINTYPEN, 3, SC_MARGIN_TEXT )
      SciMsg( m_pSci(i), SCI_SETMARGINWIDTHN, 3, 4 * rxRatio )
      SciMsg( m_pSci(i), SCI_SETMARGINTYPEN, 3, SC_MARGIN_FORE )
      SciMsg( m_pSci(i), SCI_SETMARGINTYPEN, 3, SC_MARGIN_BACK )
             
      ''
      ''  CONFINE CARET to TEXT
      if gConfig.ConfineCaret then
         SciMsg( m_pSci(i), SCI_SETVIRTUALSPACEOPTIONS, SCVS_RECTANGULARSELECTION, 0 )
      else
         SciMsg( m_pSci(i), SCI_SETVIRTUALSPACEOPTIONS, SCVS_RECTANGULARSELECTION or SCVS_USERACCESSIBLE, 0 )
      end if
   
      ''
      ''  TABS as SPACES
      if gConfig.TabIndentSpaces then
         SciMsg( m_pSci(i), SCI_SETUSETABS, false, 0 )
      else
         SciMsg( m_pSci(i), SCI_SETUSETABS, Ctrue, 0 )
      end if
        
      ''
      ''  SELECTIONS FILL ENTIRE SCREEN SPACE
      SciMsg( m_pSci(i), SCI_SETSELEOLFILLED, Ctrue, 0 )
          
      ''
      ''  TAB WIDTH 
      SciMsg( m_pSci(i), SCI_SETTABWIDTH, val(gConfig.TabSize), 0 )
      SciMsg( m_pSci(i), SCI_SETINDENT, val(gConfig.TabSize), 0 )
 
      ''
      ''  INDENTATION GUIDES
      if gConfig.IndentGuides then
         SciMsg( m_pSci(i), SCI_SETINDENTATIONGUIDES, Ctrue, 0)
      else
         SciMsg( m_pSci(i), SCI_SETINDENTATIONGUIDES, false, 0)
      end if
      SciMsg( m_pSci(i), SCI_STYLESETFORE, STYLE_INDENTGUIDE, ghEditor.ForeColorIndentguides )
      SciMsg( m_pSci(i), SCI_STYLESETBACK, STYLE_INDENTGUIDE, ghEditor.BackColorIndentguides )
      
      ''
      ''  CARET
      SciMsg( m_pSci(i), SCI_SETCARETFORE, ghEditor.ForeColorCaret, 0 )
      SciMsg( m_pSci(i), SCI_SETCARETWIDTH, 2, 0 )       ' 2 pixels

      ''
      ''  SHOW CARET LINE
      if gConfig.HighlightCurrentLine then
         SciMsg( m_pSci(i), SCI_SETCARETLINEVISIBLE, Ctrue, 0 )
         SciMsg( m_pSci(i), SCI_SETCARETLINEBACK, ghEditor.BackColorCurrentline, 0 )
      else
         SciMsg( m_pSci(i), SCI_SETCARETLINEVISIBLE, false, 0 )
      end if

      ''
      ''  SELECTION COLORS
      SciMsg( m_pSci(i), SCI_SETSELFORE, Ctrue, ghEditor.ForeColorSelection )
      SciMsg( m_pSci(i), SCI_SETSELBACK, Ctrue, ghEditor.BackColorSelection )
   
      ''
      ''  MULTIPLE SELECTIONS
      SciMsg( m_pSci(i), SCI_SETMULTIPLESELECTION, Ctrue, 0 ) 
               
      ''
      ''  ALWAYS KEEP THE CARET LINE VISIBLE
      SciMsg( m_pSci(i), SCI_SETCARETLINEVISIBLEALWAYS, Ctrue, 0 )
      
      ''
      ''  DISABLE RIGHT CLICK POPUP MENU
      SciMsg( m_pSci(i), SCI_USEPOPUP, false, 0 )

      ''
      ''  IDENTIFY CHARACTERS to BE USED IN WORDS
      SciMsg( m_pSci(i), SCI_SETWORDCHARS, 0, cast(LPARAM, @"~_\abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789") )

      ''  UNICODE (UTF-8 encoding)
      if this.FileEncoding = FILE_ENCODING_ANSI then
         SciMsg( m_pSci(i), SCI_SETCODEPAGE, 0, 0 )
      else
         ' UTF8 or UTF16 would have been converted to UTF8 in order to display in the editor.
         SciMsg( m_pSci(i), SCI_SETCODEPAGE, SC_CP_UTF8, 0 )
      end if
   
      '' RIGHT EDGE COLUMN
      SciMsg( m_pSci(i), SCI_SETEDGEMODE, iif(gConfig.RightEdge, EDGE_LINE, EDGE_NONE), 0 )
      SciMsg( m_pSci(i), SCI_SETEDGECOLUMN, val(gConfig.RightEdgePosition), 0 )

      ''
      ''  OTHER
      SciMsg( m_pSci(i), SCI_SETADDITIONALSELECTIONTYPING, true, 0 )
   

      ''
      ''  APPLY ALL LANGUAGE SPECIFIC SYNTAX COLORING
      wFileExt = AfxStrPathname( "EXTN", this.DiskFilename )
      wFileExt = ucase(wFileExt)

      if cbool(wFileExt = ".BAS") orelse cbool(wFileExt = ".INC") _
            orelse cbool(wFileExt = ".BI") orelse (this.IsNewFlag = true) _
            orelse cbool(wFileExt = ".FBTPL") then

         bitsNeeded = SciMsg( m_pSci(i), SCI_GETSTYLEBITSNEEDED, 0, 0)
         SciMsg( m_pSci(i), SCI_SETSTYLEBITS, bitsNeeded, 0 )

         ' Set FreeBASIC Keywords
         if len(gConfig.FBKeywords) then
            SciMsg( m_pSci(i), SCI_SETKEYWORDS, 0, cast(LPARAM, strptr(gConfig.FBKeywords)) )
         end if

         ' Set Windows Api Keywords
         if len(gConfig.WinApiKeywords) then
            SciMsg( m_pSci(i), SCI_SETKEYWORDS, 1, cast(LPARAM, strptr(gConfig.WinApiKeywords)) )
         end if
         
         if gConfig.SyntaxHighlighting then
            ' Set the Default text colors
            SciMsg( m_pSci(i), SCI_STYLESETFORE, SCE_B_DEFAULT, ghEditor.ForeColorText)
            SciMsg( m_pSci(i), SCI_STYLESETBACK, SCE_B_DEFAULT, ghEditor.BackColorText)
            SciMsg( m_pSci(i), SCI_STYLESETBOLD, SCE_B_DEFAULT, ghEditor.TextBold )
            SciMsg( m_pSci(i), SCI_STYLESETITALIC, SCE_B_DEFAULT, ghEditor.TextItalic )
            SciMsg( m_pSci(i), SCI_STYLESETUNDERLINE, SCE_B_DEFAULT, ghEditor.TextUnderline )

            ' Set the Multiline Comments style
            SciMsg( m_pSci(i), SCI_STYLESETFORE, SCE_B_MULTILINECOMMENT, ghEditor.ForeColorComments )
            SciMsg( m_pSci(i), SCI_STYLESETBACK, SCE_B_MULTILINECOMMENT, ghEditor.BackColorComments )
            SciMsg( m_pSci(i), SCI_STYLESETBOLD, SCE_B_MULTILINECOMMENT, ghEditor.CommentsBold )
            SciMsg( m_pSci(i), SCI_STYLESETITALIC, SCE_B_MULTILINECOMMENT, ghEditor.CommentsItalic )
            SciMsg( m_pSci(i), SCI_STYLESETUNDERLINE, SCE_B_MULTILINECOMMENT, ghEditor.CommentsUnderline )

            ' Set the Comments style
            SciMsg( m_pSci(i), SCI_STYLESETFORE, SCE_B_COMMENT, ghEditor.ForeColorComments )
            SciMsg( m_pSci(i), SCI_STYLESETBACK, SCE_B_COMMENT, ghEditor.BackColorComments )
            SciMsg( m_pSci(i), SCI_STYLESETBOLD, SCE_B_COMMENT, ghEditor.CommentsBold )
            SciMsg( m_pSci(i), SCI_STYLESETITALIC, SCE_B_COMMENT, ghEditor.CommentsItalic )
            SciMsg( m_pSci(i), SCI_STYLESETUNDERLINE, SCE_B_COMMENT, ghEditor.CommentsUnderline )

            ' Set the Keywords style (FreeBasic)
            SciMsg( m_pSci(i), SCI_STYLESETFORE, SCE_B_KEYWORD, ghEditor.ForeColorKeyword )
            SciMsg( m_pSci(i), SCI_STYLESETBACK, SCE_B_KEYWORD, ghEditor.BackColorKeyword )
            SciMsg( m_pSci(i), SCI_STYLESETCASE, SCE_B_KEYWORD, nFontcase )
            SciMsg( m_pSci(i), SCI_STYLESETBOLD, SCE_B_KEYWORD, ghEditor.KeywordBold )
            SciMsg( m_pSci(i), SCI_STYLESETITALIC, SCE_B_KEYWORD, ghEditor.KeywordItalic )
            SciMsg( m_pSci(i), SCI_STYLESETUNDERLINE, SCE_B_KEYWORD, ghEditor.KeywordUnderline )

            ' Set the Keywords style (Windows Api)
            SciMsg( m_pSci(i), SCI_STYLESETFORE, SCE_B_KEYWORD2, ghEditor.ForeColorKeyword2 )
            SciMsg( m_pSci(i), SCI_STYLESETBACK, SCE_B_KEYWORD2, ghEditor.BackColorKeyword2 )
            SciMsg( m_pSci(i), SCI_STYLESETCASE, SCE_B_KEYWORD2, nFontcase )
            SciMsg( m_pSci(i), SCI_STYLESETBOLD, SCE_B_KEYWORD2, ghEditor.KeywordBold2 )
            SciMsg( m_pSci(i), SCI_STYLESETITALIC, SCE_B_KEYWORD2, ghEditor.KeywordItalic2 )
            SciMsg( m_pSci(i), SCI_STYLESETUNDERLINE, SCE_B_KEYWORD2, ghEditor.KeywordUnderline2 )

            ' Set the Numbers style
            SciMsg( m_pSci(i), SCI_STYLESETFORE, SCE_B_NUMBER, ghEditor.ForeColorNumbers )
            SciMsg( m_pSci(i), SCI_STYLESETBACK, SCE_B_NUMBER, ghEditor.BackColorNumbers )
            SciMsg( m_pSci(i), SCI_STYLESETBOLD, SCE_B_NUMBER, ghEditor.NumbersBold )
            SciMsg( m_pSci(i), SCI_STYLESETITALIC, SCE_B_NUMBER, ghEditor.NumbersItalic)
            SciMsg( m_pSci(i), SCI_STYLESETUNDERLINE, SCE_B_NUMBER, ghEditor.NumbersUnderline)

            ' Set the Operators style
            SciMsg( m_pSci(i), SCI_STYLESETFORE, SCE_B_OPERATOR, ghEditor.ForeColorOperators )
            SciMsg( m_pSci(i), SCI_STYLESETBACK, SCE_B_OPERATOR, ghEditor.BackColorOperators )
            SciMsg( m_pSci(i), SCI_STYLESETBOLD, SCE_B_OPERATOR, ghEditor.OperatorsBold)
            SciMsg( m_pSci(i), SCI_STYLESETITALIC, SCE_B_OPERATOR, ghEditor.OperatorsItalic)
            SciMsg( m_pSci(i), SCI_STYLESETUNDERLINE, SCE_B_OPERATOR, ghEditor.OperatorsUnderline)

            ' Set the Preprocessor style
            SciMsg( m_pSci(i), SCI_STYLESETFORE, SCE_B_PREPROCESSOR, ghEditor.ForeColorPreprocessor )
            SciMsg( m_pSci(i), SCI_STYLESETBACK, SCE_B_PREPROCESSOR, ghEditor.BackColorPreprocessor )
            SciMsg( m_pSci(i), SCI_STYLESETBOLD, SCE_B_PREPROCESSOR, ghEditor.PreprocessorBold)
            SciMsg( m_pSci(i), SCI_STYLESETITALIC, SCE_B_PREPROCESSOR, ghEditor.PreprocessorItalic)
            SciMsg( m_pSci(i), SCI_STYLESETUNDERLINE, SCE_B_PREPROCESSOR, ghEditor.PreprocessorUnderline)

            ' Set the strings style
            SciMsg( m_pSci(i), SCI_STYLESETFORE, SCE_B_STRING, ghEditor.ForeColorStrings )
            SciMsg( m_pSci(i), SCI_STYLESETBACK, SCE_B_STRING, ghEditor.BackColorStrings )
            SciMsg( m_pSci(i), SCI_STYLESETBOLD, SCE_B_STRING, ghEditor.StringsBold)
            SciMsg( m_pSci(i), SCI_STYLESETITALIC, SCE_B_STRING, ghEditor.StringsItalic)
            SciMsg( m_pSci(i), SCI_STYLESETUNDERLINE, SCE_B_STRING, ghEditor.StringsUnderline)

         end if
      end if
  
      ''
      ''  CODE FOLDING
      if gConfig.FoldMargin then
         ' Enable folding of the procedures and functions
         SciMsg( m_pSci(i), SCI_SETPROPERTY, cast(WPARAM, @"fold"), cast(LPARAM, @"1") )

         ' Initialize fold symbols for folding - Box tree
         SciMsg( m_pSci(i), SCI_MARKERDEFINE, SC_MARKNUM_FOLDEROPEN,    SC_MARK_BOXMINUS )
         SciMsg( m_pSci(i), SCI_MARKERDEFINE, SC_MARKNUM_FOLDER,        SC_MARK_BOXPLUS )
         SciMsg( m_pSci(i), SCI_MARKERDEFINE, SC_MARKNUM_FOLDERSUB,     SC_MARK_VLINE)
         SciMsg( m_pSci(i), SCI_MARKERDEFINE, SC_MARKNUM_FOLDERTAIL,    SC_MARK_LCORNER)
         SciMsg( m_pSci(i), SCI_MARKERDEFINE, SC_MARKNUM_FOLDEREND,     SC_MARK_BOXPLUSCONNECTED)
         'SciMsg( m_pSci(i), SCI_MARKERDEFINE, SC_MARKNUM_FOLDEROPENMID, SC_MARK_EMPTY)   ' SC_MARK_BOXMINUSCONNECTED
         SciMsg( m_pSci(i), SCI_MARKERDEFINE, SC_MARKNUM_FOLDEROPENMID, SC_MARK_BOXMINUSCONNECTED)
         SciMsg( m_pSci(i), SCI_MARKERDEFINE, SC_MARKNUM_FOLDERMIDTAIL, SC_MARK_TCORNER)

         ' Draw line below if not expanded
         SciMsg( m_pSci(i), SCI_SETFOLDFLAGS, 16, 0 )

         ' Colors for folders closed and folders opened
         dim as COLORREF clrFore = ghEditor.ForeColorFoldsymbol
         dim as COLORREF clrBack = ghEditor.BackColorFoldsymbol

         SciMsg( m_pSci(i), SCI_MARKERSETFORE, SC_MARKNUM_FOLDER, clrBack)
         SciMsg( m_pSci(i), SCI_MARKERSETBACK, SC_MARKNUM_FOLDER, clrFore)

         SciMsg( m_pSci(i), SCI_MARKERSETFORE, SC_MARKNUM_FOLDEROPEN, clrBack)
         SciMsg( m_pSci(i), SCI_MARKERSETBACK, SC_MARKNUM_FOLDEROPEN, clrFore)

         SciMsg( m_pSci(i), SCI_MARKERSETFORE, SC_MARKNUM_FOLDERSUB, clrBack)
         SciMsg( m_pSci(i), SCI_MARKERSETBACK, SC_MARKNUM_FOLDERSUB, clrFore)

         SciMsg( m_pSci(i), SCI_MARKERSETFORE, SC_MARKNUM_FOLDERTAIL, clrBack)
         SciMsg( m_pSci(i), SCI_MARKERSETBACK, SC_MARKNUM_FOLDERTAIL, clrFore)

         SciMsg( m_pSci(i), SCI_MARKERSETFORE, SC_MARKNUM_FOLDEREND, clrBack)
         SciMsg( m_pSci(i), SCI_MARKERSETBACK, SC_MARKNUM_FOLDEREND, clrFore)

         SciMsg( m_pSci(i), SCI_MARKERSETFORE, SC_MARKNUM_FOLDEROPENMID, clrBack)
         SciMsg( m_pSci(i), SCI_MARKERSETBACK, SC_MARKNUM_FOLDEROPENMID, clrFore)

         SciMsg( m_pSci(i), SCI_MARKERSETFORE, SC_MARKNUM_FOLDERMIDTAIL, clrBack)
         SciMsg( m_pSci(i), SCI_MARKERSETBACK, SC_MARKNUM_FOLDERMIDTAIL, clrFore)

         SciMsg( m_pSci(i), SCI_MARKERENABLEHIGHLIGHT, false, 0)
      
      else
         ' Disable folding of the procedures and functions
         SciMsg( m_pSci(i), SCI_SETPROPERTY, cast(WPARAM, @"fold"), cast(LPARAM, @"0") )
      end if
   
   next
   
   function = 0
end function


''
''
function clsDocument.GetWord( byval curPos as long = -1 ) as string
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   ' Get word at the specified location or under the cursor
   dim as long x, y, p
   dim as string buffer 

   ' Retrieve the current position
   if curPos = -1 then curPos = SciMsg( pSci, SCI_GETCURRENTPOS, 0, 0)
   ' Retrieve the starting and ending position of the word
   x = SciMsg( pSci, SCI_WORDSTARTPOSITION, curPos, true)
   y = SciMsg( pSci, SCI_WORDENDPOSITION, curPos, false)
   if y > x then
      ' Text range
      buffer = this.GetTextRange(x, y)
      ' Remove the $NUL
      p = instr(buffer, chr(0))
      if p then buffer = left(buffer, p - 1)
   end if
   buffer = AfxStrRemoveAny( buffer, chr(13, 10, 34) & "()%," )
   function = buffer

end function

''
''
function clsDocument.GetBookmarks() as string
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   ' Retrieve all bookmark positions in the document and return it
   ' as a comma delimited string to be saved to project file.
   dim as string buffer
   dim as long fMark  ' 32 bit value
   dim as long nLines = SciMsg( pSci, SCI_GETLINECOUNT, 0, 0) 

   for i as long = 0 to nLines - 1
      fMark = SciMsg( pSci, SCI_MARKERGET, i, 0)
      if bit(fMark, 0) then
         buffer = buffer & i & ","
      end if
   next

   function = rtrim(buffer, ",")

end function

''
''
function clsDocument.SetBookmarks( byval sBookmarks as string ) as long
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   ' Set all bookmark positions in the document and based on the 
   ' incoming comma delimited string retrieved from a project file.
   sBookmarks = trim(sBookmarks)
   if len(sBookmarks) = 0 then exit function
   
   dim as long nCount = AfxStrParseCount(sBookmarks, ",")
   dim as long nLine
   
   for i as long = 1 to nCount
      nLine = val( AfxStrParse(sBookmarks, i, ",") )
      SciMsg( pSci, SCI_MARKERADD, nLine, 0)
   next

   function = 0

end function

''
''
function clsDocument.GetFoldPoints() as string
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   ' Retrieve all folding positions in the document and return it
   ' as a comma delimited string to be saved to project file.
   dim as string buffer
   dim as long nLines = SciMsg( pSci, SCI_GETLINECOUNT, 0, 0) 

   for i as long = 0 to nLines - 1
      dim nFoldLevel as long = SciMsg( pSci, SCI_GETFOLDLEVEL, i, 0)
      if (nFoldLevel and SC_FOLDLEVELHEADERFLAG) then
         if SciMsg( pSci, SCI_GETFOLDEXPANDED, i, 0) = 0 then
            buffer = buffer & i & ","
         end if   
      end if
   next

   function = rtrim(buffer, ",")

end function

''
''
function clsDocument.SetFoldPoints( byval sFoldPoints as string ) as long
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   ' Set all folding positions in the document and based on the 
   ' incoming comma delimited string retrieved from a project file.
   sFoldPoints = trim(sFoldPoints)
   if len(sFoldPoints) = 0 then exit function
   
   dim as long nCount = AfxStrParseCount(sFoldPoints, ",")
   dim as long nLine
   
   for i as long = 1 to nCount
      nLine = val( AfxStrParse(sFoldPoints, i, ",") )
      SciMsg( pSci, SCI_FOLDLINE, nLine, SC_FOLDACTION_CONTRACT) 
   next

   function = 0

end function

''
''
function clsDocument.LineDuplicate() as long
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   ' Duplicate the current caret line, or an entire block of
   ' code should a selection be active.
   dim as long startSelPos = SciMsg( pSci, SCI_GETSELECTIONSTART, 0, 0)
   dim as long endSelPos   = SciMsg( pSci, SCI_GETSELECTIONEND, 0, 0)
   if startSelPos = endSelPos then   ' no selection
      ' Simply duplicate the line that the caret is on.
      SciMsg( pSci, SCI_LINEDUPLICATE, 0, 0)
   else
      SciMsg( pSci, SCI_SELECTIONDUPLICATE, 0, 0)
   end if
   function = 0
end function


''
''
function clsDocument.LinesPerPage( byval idxWindow as long ) as long
   dim as any ptr pSci = this.GetActiveScintillaPtr()
   ' SCI_LINESONSCREEN is not a reliable count of lines per screen because
   ' it does not take into account any extra ascent or descent value.
   dim as long nLineHeight = SciMsg( pSci, SCI_TEXTHEIGHT, 0, 0 )
   dim as RECT rc = AfxGetWindowRect(this.hWindow(idxWindow))
   MapWindowPoints( HWND_DESKTOP, HWND_FRMMAIN, cast(POINT ptr, @rc), 2 )
   dim as long nHeight = rc.bottom - rc.top
   function = (nHeight / nLineHeight)
end function


''
''
function clsDocument.CompileDirectives( Directives() as COMPILE_DIRECTIVES ) as long
   ' Search the source code for any user embedded compiler directives.
   dim ub     as long    
   dim i      as long    
   dim nLines as long    
   dim st     as string
   
   dim sText as string  ' this will be an UTF-8 encoded string
   
   dim as any ptr pSci = this.GetActiveScintillaPtr()

   if pSci = 0 then exit function

   nLines = SciMsg( pSci, SCI_GETLINECOUNT, 0, 0)

   for i = 0 to nLines - 1
      st = ltrim(this.GetLine(i))
         
      if left(st, 1) <> "'" then continue for
      st = ltrim(mid(st, 2))

      if len(st) < 11 then continue for
      st = ucase(st)
      
      ub = ubound(Directives)
      
      ' '#CONSOLE ON|OFF
      if left(st, 11) = "#CONSOLE ON" then
         redim preserve Directives(ub+1)
         Directives(ub+1).DirectiveFlag = IDM_CONSOLE
      elseif left(st, 12) = "#CONSOLE OFF" then   
         redim preserve Directives(ub+1)
         Directives(ub+1).DirectiveFlag = IDM_GUI
      end if

      ' '#RESOURCE "filename.rc"
      if left(st, 10) = "#RESOURCE " then
         redim preserve Directives(ub+1)
         Directives(ub+1).DirectiveFlag = IDM_RESOURCE
         st = mid(st, 11)
         Directives(ub+1).DirectiveText = AfxStrExtract(st, chr(34), chr(34))
      end if
         
   next

   function = 0
end function


