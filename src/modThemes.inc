'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2023 Paul Squires, PlanetSquires Software
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

type THEME_TYPE
   keyValue as CWSTR
   keyData as CWSTR 
   foreground as COLORREF
   background as COLORREF
   bold as long
   italic as long
   underline as long
end type
dim shared gTheme(any) as THEME_TYPE
dim shared gThemeFilenames(any) as CWSTR

type GENERAL_HANDLES_TYPE
   description as CWSTR
   version as CWSTR
   iconfail as COLORREF      
   iconsuccess as COLORREF      
end type
dim shared ghGeneral as GENERAL_HANDLES_TYPE

type EDITOR_HANDLES_TYPE
   ForeColorCaret as COLORREF
   BackColorCurrentLine as COLORREF
   ForeColorText as COLORREF
   BackColorText as COLORREF
   TextBold as boolean
   TextItalic as boolean
   TextUnderline as boolean
   ForeColorNumbers as COLORREF
   BackColorNumbers as COLORREF
   NumbersBold as boolean
   NumbersItalic as boolean
   NumbersUnderLine as boolean
   ForeColorComments as COLORREF
   BackColorComments as COLORREF
   CommentsBold as boolean
   CommentsItalic as boolean
   CommentsUnderline as boolean
   ForeColorKeyword as COLORREF
   BackColorKeyword as COLORREF
   KeywordBold as boolean
   KeywordItalic as boolean
   KeywordUnderline as boolean
   ForeColorKeyword2 as COLORREF
   BackColorKeyword2 as COLORREF
   KeywordBold2 as boolean
   KeywordItalic2 as boolean
   KeywordUnderline2 as boolean
   ForeColorFoldmargin as COLORREF
   ForeColorFoldsymbol as COLORREF
   BackColorFoldsymbol as COLORREF
   ForeColorLinenumbers as COLORREF
   BackColorLinenumbers as COLORREF
   LinenumbersBold as boolean
   LinenumbersItalic as boolean
   LinenumbersUnderline as boolean
   ForeColorOperators as COLORREF
   BackColorOperators as COLORREF
   OperatorsBold as boolean
   OperatorsItalic as boolean
   OperatorsUnderline as boolean
   ForeColorIndentguides as COLORREF
   BackColorIndentguides as COLORREF
   ForeColorPreprocessor as COLORREF
   BackColorPreprocessor as COLORREF
   PreprocessorBold as boolean
   PreprocessorItalic as boolean
   PreprocessorUnderline as boolean
   ForeColorSelection as COLORREF
   BackColorSelection as COLORREF
   ForeColorStrings as COLORREF
   BackColorStrings as COLORREF
   StringsBold as boolean
   StringsItalic as boolean
   StringsUnderline as boolean
   ForeColorBracegood as COLORREF
   ForeColorBracebad as COLORREF
   ForeColorOccurrence as COLORREF
   hBackBrushScrollBar as HBRUSH
   BackColorScrollBar as COLORREF
   ForeColorScrollBar as COLORREF
   hBackBrushThumb as HBRUSH
   Divider as COLORREF
end type
dim shared ghEditor as EDITOR_HANDLES_TYPE

type PANEL_HANDLES_TYPE
   hPanelBrush as HBRUSH
   hBackBrushScrollBar as HBRUSH
   BackColorScrollBar as COLORREF
   ForeColorScrollBar as COLORREF
   hBackBrushThumb as HBRUSH
   hBackBrush as HBRUSH
   BackColor as COLORREF
   ForeColor as COLORREF
   hBackBrushHot as HBRUSH
   BackColorHot as COLORREF
   ForeColorHot as COLORREF
   hBackBrushButton as HBRUSH
   BackColorButton as COLORREF
   ForeColorButton as COLORREF
   hBackBrushButtonHot as HBRUSH
   BackColorButtonHot as COLORREF
   ForeColorButtonHot as COLORREF
end type
dim shared ghPanel as PANEL_HANDLES_TYPE

type MENUBAR_HANDLES_TYPE
   hPanelBrush as HBRUSH
   hBackBrush as HBRUSH
   hBackBrushHot as HBRUSH
   BackColor as COLORREF
   BackColorHot as COLORREF
   ForeColor as COLORREF
   ForeColorHot as COLORREF
   hFontMenuBar as HFONT
   hFontSymbol as HFONT
   hFontSymbolSmall as HFONT
   hFontSymbolLargeBold as HFONT
end type
dim shared ghMenuBar as MENUBAR_HANDLES_TYPE

type POPUPMENU_HANDLES_TYPE
   PanelColor as COLORREF
   hPanelBrush as HBRUSH
   hBackBrush as HBRUSH
   hBackBrushHot as HBRUSH
   hBackBrushDisabled as HBRUSH
   BackColor as COLORREF
   BackColorHot as COLORREF
   BackColorDisabled as COLORREF
   ForeColor as COLORREF
   ForeColorHot as COLORREF
   ForeColorDisabled as COLORREF
end type
dim shared ghPopup as POPUPMENU_HANDLES_TYPE

type TOPTABS_HANDLES_TYPE
   hPanelBrush as HBRUSH
   hBackBrush as HBRUSH
   hBackBrushHot as HBRUSH
   BackColor as COLORREF
   BackColorHot as COLORREF
   ForeColor as COLORREF
   ForeColorHot as COLORREF
   Divider as COLORREF
   hCloseBrushHot as HBRUSH
   CloseBackColorHot as COLORREF
end type
dim shared ghTopTabs as TOPTABS_HANDLES_TYPE

type STATUSBAR_HANDLES_TYPE
   hPanelBrush as HBRUSH
   hBackBrush as HBRUSH
   hBackBrushHot as HBRUSH
   BackColor as COLORREF
   BackColorHot as COLORREF
   ForeColor as COLORREF
   ForeColorHot as COLORREF
   hFontStatusBar as HFONT
   hFontStatusBarBold as HFONT
end type
dim shared ghStatusBar as STATUSBAR_HANDLES_TYPE

type FINDREPLACE_HANDLES_TYPE
   hPanelBrush as HBRUSH
   hBackBrush as HBRUSH
   hBackBrushHot as HBRUSH
   BackColor as COLORREF
   BackColorHot as COLORREF
   ForeColor as COLORREF
   ForeColorHot as COLORREF
   leftedge as COLORREF
   hIconBrushHot as HBRUSH
   IconBackColorHot as COLORREF
   NotFoundForeColor as COLORREF
   hBackBrushTextBox as HBRUSH
   TextBoxForeColor as COLORREF
   TextBoxBackColor as COLORREF
   hBackBrushIconSelected as HBRUSH
   IconSelectedBackColor as COLORREF
   CueBannerForeColor as COLORREF
end type
dim shared ghFindReplace as FINDREPLACE_HANDLES_TYPE

type OUTPUT_HANDLES_TYPE
   hPanelBrush as HBRUSH
   hBackBrush as HBRUSH
   hBackBrushHot as HBRUSH
   BackColor as COLORREF
   BackColorHot as COLORREF
   ForeColor as COLORREF
   ForeColorHot as COLORREF
   Divider as COLORREF
   hCloseBrushHot as HBRUSH
   CloseBackColorHot as COLORREF
end type
dim shared ghOutput as OUTPUT_HANDLES_TYPE

type DESIGNER_HANDLES_TYPE
   hPanelBrush as HBRUSH
end type
dim shared ghDesigner as DESIGNER_HANDLES_TYPE

''
''  RETRIEVE THEME ITEM FROM ARRAY
''
function getThemePtr( byval lookupKey as string ) as THEME_TYPE ptr
   ' we do a simple linear search for the theme because there will be a small
   ' number of theme elements. maybe in the future put this in a hash array for
   ' faster lookups should we need the performance boost.
   ' in case search fails return pointer to this dummy type to avoid null ptr
   static dummy as THEME_TYPE    
   for i as long = lbound(gTheme) to ubound(gTheme)
      if gTheme(i).keyValue = lookupKey then
         return @gTheme(i)
      end if
   next
   return @dummy
end function


''
''  LOAD THEME DISK FILE
''
function LoadThemeFile() as boolean

   dim as CWSTR wszFilename = AfxGetExePathName & "themes\" & gConfig.ThemeFilename
   if AfxFileExists(wszFilename) = 0 then
      ' default to the visual studio dark theme
      wszFilename = AfxGetExePathName & "themes\winfbe_default_dark.theme"
      if AfxFileExists(wszFilename) = 0 then
         ' If the theme still does not exist then most likely the Themes folder is
         ' missing or the theme was renamed. Use the backup version in the settings folder. 
         wszFilename = AfxGetExePathName & "settings\backup_winfbe_default_dark.theme"
         ' Attempt to copy the file for next time
         dim as CWSTR wszThemeDirFilename = AfxGetExePathName & "themes\winfbe_default_dark.theme"
         if AfxCopyFile( wszFilename, wszThemeDirFilename ) then
            gConfig.ThemeFilename = "winfbe_default_dark.theme"
         end if   
      end if
   end if

   dim pStream AS CTextStream   
   if pStream.Open(wszFilename) <> S_OK then return true   ' error

   ' Ensure to delete any existing gTheme data
   erase gTheme
   
   dim param(100) as CWSTR
   dim param_value(100) as CWSTR
   dim as long nextParam = 0
   
   dim as CWSTR wst, keyValue, keyData
      
   Do Until pStream.EOS
      wst = pStream.ReadLine

      wst = trim(AfxStrExtract( 1, wst, "#"))     ' remove comments
      If Len(wst) = 0 Then Continue Do

      keyValue = trim(AfxStrParse(wst, 1, ":"))
      keyData = trim(AfxStrParse(wst, 2, ":"))
      
      ' is this a replaceable parameter
      if left(wst, 1) = "%" then
         param(nextParam) = keyValue
         param_value(nextParam) = keyData
         nextParam += 1
         continue do
      end if

      ' if keyData is a replaceable parameter then substitute it now
      if left(keyData, 1) = "%" then
         for i as long = lbound(param) to ubound(param)
            if param(i) = keyData then
               keyData = param_value(i)
               exit for
            end if   
         next
      end if
      
      ' is the KeyData a true or false value
      if keyData = "true" then keyData = "1"
      if keyData = "false" then keyData = "0"
      
      ' determine number of keys in the keyvalue. There can be 2 or 3.
      ' if 2 then simply store the data in keyData. if 3 then break the
      ' data into the component properties.
      dim keys(3) as string   ' always have maximum of 3 keys
      dim as long numKeys = AfxStrParseCount(keyValue, ".")  ' actual number of keys retrieved
      for i as long = 1 to numKeys
         keys(i) = AfxStrParse(keyValue, i, ".")
      next
      ' there will always be at least 2 elements to the key and
      ' we use that to lookup the theme type structure.
      keyValue = keys(1) & "." & keys(2)
      
      ' generate the color value
      dim as long r, g, b, clr
      r = val(AfxStrParse(keyData, 1, ","))
      g = val(AfxStrParse(keyData, 2, ","))
      b = val(AfxStrParse(keyData, 3, ","))
      clr = BGR(r, g, b)
      
      ' determine if the key already exists and just needs
      ' the properties added to it, or if the structure needs
      ' to be created.
      dim nFound as long = -1
      for i as long = lbound(gTheme) to ubound(gTheme)
         if gTheme(i).keyValue = keyValue then
            nFound = i: exit for
         end if
      next
      
      ' add the new key/value pair if not found
      if nFound = -1 then
         nFound = ubound(gTheme) + 1
         redim preserve gTheme(nFound) as THEME_TYPE
      end if   

      gTheme(nFound).keyValue = keyValue
      gTheme(nFound).keyData = keyData
      select case keys(3)
         case "foreground": gTheme(nFound).foreground = clr
         case "background": gTheme(nFound).background = clr
         case "bold":       gTheme(nFound).bold = val(keyData)
         case "italic":     gTheme(nFound).italic = val(keyData)
         case "underline":  gTheme(nFound).underline = val(keyData)         
      end select
   loop
   
   pStream.Close
   
   ' Load the TYPE's for each of the forms that require theme data   
   DIM pWindow AS CWindow PTR = AfxCWindowPtr(HWND_FRMMAIN)

   dim as THEME_TYPE ptr pTheme
   
   pTheme = getThemePtr("general.description")
   ghGeneral.description = pTheme->keyData
   pTheme = getThemePtr("general.version")
   ghGeneral.version = pTheme->keyData
   pTheme = getThemePtr("compile.iconsuccess")
   ghGeneral.iconsuccess = pTheme->foreground
   pTheme = getThemePtr("compile.iconfail")
   ghGeneral.iconfail = pTheme->foreground

   pTheme = getThemePtr("editor.text")
   ghEditor.BackColorText = pTheme->background
   ghEditor.ForeColorText = pTheme->foreground
   ghEditor.TextBold = pTheme->bold
   ghEditor.TextItalic = pTheme->italic
   ghEditor.TextUnderline = pTheme->underline
   pTheme = getThemePtr("editor.linenumbers")
   ghEditor.BackColorLinenumbers = pTheme->background
   ghEditor.ForeColorLinenumbers = pTheme->foreground
   ghEditor.LinenumbersBold = pTheme->bold
   ghEditor.LinenumbersItalic = pTheme->italic
   ghEditor.LinenumbersUnderline = pTheme->underline
   pTheme = getThemePtr("editor.foldmargin")
   ghEditor.ForeColorFoldmargin = pTheme->foreground
   pTheme = getThemePtr("editor.indentguides")
   ghEditor.BackColorIndentguides = pTheme->background
   ghEditor.ForeColorIndentguides = pTheme->foreground
   pTheme = getThemePtr("editor.caret")
   ghEditor.ForeColorCaret = pTheme->foreground
   pTheme = getThemePtr("editor.currentline")
   ghEditor.BackColorCurrentline = pTheme->background
   pTheme = getThemePtr("editor.selection")
   ghEditor.BackColorSelection = pTheme->background
   ghEditor.ForeColorSelection = pTheme->foreground
   pTheme = getThemePtr("editor.comments")
   ghEditor.BackColorComments = pTheme->background
   ghEditor.ForeColorComments = pTheme->foreground
   ghEditor.CommentsBold = pTheme->bold
   ghEditor.CommentsItalic = pTheme->italic
   ghEditor.CommentsUnderline = pTheme->underline
   pTheme = getThemePtr("editor.keyword")
   ghEditor.BackColorKeyword = pTheme->background
   ghEditor.ForeColorKeyword = pTheme->foreground
   ghEditor.KeywordBold = pTheme->bold
   ghEditor.KeywordItalic = pTheme->italic
   ghEditor.KeywordUnderline = pTheme->underline
   pTheme = getThemePtr("editor.keyword2")
   ghEditor.BackColorKeyword2 = pTheme->background
   ghEditor.ForeColorKeyword2 = pTheme->foreground
   ghEditor.KeywordBold2 = pTheme->bold
   ghEditor.KeywordItalic2 = pTheme->italic
   ghEditor.KeywordUnderline2 = pTheme->underline
   pTheme = getThemePtr("editor.numbers")
   ghEditor.BackColorNumbers = pTheme->background
   ghEditor.ForeColorNumbers = pTheme->foreground
   ghEditor.NumbersBold = pTheme->bold
   ghEditor.NumbersItalic = pTheme->italic
   ghEditor.NumbersUnderline = pTheme->underline
   pTheme = getThemePtr("editor.operators")
   ghEditor.BackColorOperators = pTheme->background
   ghEditor.ForeColorOperators = pTheme->foreground
   ghEditor.OperatorsBold = pTheme->bold
   ghEditor.OperatorsItalic = pTheme->italic
   ghEditor.OperatorsUnderline = pTheme->underline
   pTheme = getThemePtr("editor.preprocessor")
   ghEditor.BackColorPreprocessor = pTheme->background
   ghEditor.ForeColorPreprocessor = pTheme->foreground
   ghEditor.PreprocessorBold = pTheme->bold
   ghEditor.PreprocessorItalic = pTheme->italic
   ghEditor.PreprocessorUnderline = pTheme->underline
   pTheme = getThemePtr("editor.strings")
   ghEditor.BackColorStrings = pTheme->background
   ghEditor.ForeColorStrings = pTheme->foreground
   ghEditor.StringsBold = pTheme->bold
   ghEditor.StringsItalic = pTheme->italic
   ghEditor.StringsUnderline = pTheme->underline
   pTheme = getThemePtr("editor.foldsymbol")
   ghEditor.BackColorFoldsymbol = pTheme->background
   ghEditor.ForeColorFoldsymbol = pTheme->foreground
   pTheme = getThemePtr("editor.scrollbar")
   if ghEditor.hBackBrushScrollBar then DeleteObject(ghEditor.hBackBrushScrollBar)
   ghEditor.hBackBrushScrollBar = CreateSolidBrush(pTheme->background)
   if ghEditor.hBackBrushThumb then DeleteObject(ghEditor.hBackBrushThumb)
   ghEditor.hBackBrushThumb = CreateSolidBrush(pTheme->foreground)
   ghEditor.ForeColorScrollBar = pTheme->foreground
   ghEditor.BackColorScrollBar = pTheme->background
   pTheme = getThemePtr("editor.bracegood")
   ghEditor.ForeColorBracegood = pTheme->foreground
   pTheme = getThemePtr("editor.bracebad")
   ghEditor.ForeColorBracebad = pTheme->foreground
   pTheme = getThemePtr("editor.occurrence")
   ghEditor.ForeColorOccurrence = pTheme->foreground
   pTheme = getThemePtr("editor.divider")
   ghEditor.Divider = pTheme->foreground

   pTheme = getThemePtr("panel.panel")
   if ghPanel.hPanelBrush then DeleteObject(ghPanel.hPanelBrush)
   ghPanel.hPanelBrush = CreateSolidBrush(pTheme->background)
   pTheme = getThemePtr("panel.text")
   if ghPanel.hBackBrush then DeleteObject(ghPanel.hBackBrush)
   ghPanel.hBackBrush = CreateSolidBrush(pTheme->background)
   ghPanel.BackColor = pTheme->background
   ghPanel.ForeColor = pTheme->foreground
   pTheme = getThemePtr("panel.texthot")
   if ghPanel.hBackBrushHot then DeleteObject(ghPanel.hBackBrushHot)
   ghPanel.hBackBrushHot = CreateSolidBrush(pTheme->background)
   ghPanel.BackColorHot = pTheme->background
   ghPanel.ForeColorHot = pTheme->foreground
   pTheme = getThemePtr("panel.scrollbar")
   if ghPanel.hBackBrushScrollBar then DeleteObject(ghPanel.hBackBrushScrollBar)
   ghPanel.hBackBrushScrollBar = CreateSolidBrush(pTheme->background)
   if ghPanel.hBackBrushThumb then DeleteObject(ghPanel.hBackBrushThumb)
   ghPanel.hBackBrushThumb = CreateSolidBrush(pTheme->foreground)
   ghPanel.BackColorScrollBar = pTheme->background
   ghPanel.ForeColorScrollBar = pTheme->foreground
   pTheme = getThemePtr("panel.button")
   if ghPanel.hBackBrushButton then DeleteObject(ghPanel.hBackBrushButton)
   ghPanel.hBackBrushButton = CreateSolidBrush(pTheme->background)
   ghPanel.BackColorButton = pTheme->background
   ghPanel.ForeColorButton = pTheme->foreground
   pTheme = getThemePtr("panel.buttonhot")
   if ghPanel.hBackBrushButtonHot then DeleteObject(ghPanel.hBackBrushButtonHot)
   ghPanel.hBackBrushButtonHot = CreateSolidBrush(pTheme->background)
   ghPanel.BackColorButtonHot = pTheme->background
   ghPanel.ForeColorButtonHot = pTheme->foreground

   pTheme = getThemePtr("topmenu.panel")
   ghPopup.PanelColor = pTheme->background
   if ghPopup.hPanelBrush then DeleteObject(ghPopup.hPanelBrush)
   ghPopup.hPanelBrush = CreateSolidBrush(ghPopup.PanelColor)
   pTheme = getThemePtr("topmenu.text")
   if ghPopup.hBackBrush then DeleteObject(ghPopup.hBackBrush)
   ghPopup.hBackBrush = CreateSolidBrush(pTheme->background)
   ghPopup.BackColor = pTheme->background
   ghPopup.ForeColor = pTheme->foreground
   pTheme = getThemePtr("topmenu.texthot")
   if ghPopup.hBackBrushHot then DeleteObject(ghPopup.hBackBrushHot)
   ghPopup.hBackBrushHot = CreateSolidBrush(pTheme->background)
   ghPopup.BackColorHot = pTheme->background
   ghPopup.ForeColorHot = pTheme->foreground
   pTheme = getThemePtr("topmenu.textdisabled")
   if ghPopup.hBackBrushDisabled then DeleteObject(ghPopup.hBackBrushDisabled)
   ghPopup.hBackBrushDisabled = CreateSolidBrush(pTheme->background)
   ghPopup.BackColorDisabled = pTheme->background
   ghPopup.ForeColorDisabled = pTheme->foreground
      
   pTheme = getThemePtr("menubar.panel")
   if ghMenuBar.hPanelBrush then DeleteObject(ghMenuBar.hPanelBrush)
   ghMenuBar.hPanelBrush = CreateSolidBrush(pTheme->background)
   pTheme = getThemePtr("menubar.text")
   if ghMenuBar.hBackBrush then DeleteObject(ghMenuBar.hBackBrush)
   ghMenuBar.hBackBrush = CreateSolidBrush(pTheme->background)
   ghMenuBar.BackColor = pTheme->background
   ghMenuBar.ForeColor = pTheme->foreground
   pTheme = getThemePtr("menubar.texthot")
   if ghMenuBar.hBackBrushHot then DeleteObject(ghMenuBar.hBackBrushHot)
   ghMenuBar.hBackBrushHot = CreateSolidBrush(pTheme->background)
   ghMenuBar.BackColorHot = pTheme->background
   ghMenuBar.ForeColorHot = pTheme->foreground
   
   pTheme = getThemePtr("toptabs.panel")
   if ghTopTabs.hPanelBrush then DeleteObject(ghTopTabs.hPanelBrush)
   ghTopTabs.hPanelBrush = CreateSolidBrush(pTheme->background)
   pTheme = getThemePtr("toptabs.text")
   if ghTopTabs.hBackBrush then DeleteObject(ghTopTabs.hBackBrush)
   ghTopTabs.hBackBrush = CreateSolidBrush(pTheme->background)
   ghTopTabs.BackColor = pTheme->background
   ghTopTabs.ForeColor = pTheme->foreground
   pTheme = getThemePtr("toptabs.texthot")
   if ghTopTabs.hBackBrushHot then DeleteObject(ghTopTabs.hBackBrushHot)
   ghTopTabs.hBackBrushHot = CreateSolidBrush(pTheme->background)
   ghTopTabs.BackColorHot = pTheme->background
   ghTopTabs.ForeColorHot = pTheme->foreground
   pTheme = getThemePtr("toptabs.divider")
   ghTopTabs.Divider = pTheme->foreground
   pTheme = getThemePtr("toptabs.closehot")
   if ghTopTabs.hCloseBrushHot then DeleteObject(ghTopTabs.hCloseBrushHot)
   ghTopTabs.hCloseBrushHot = CreateSolidBrush(pTheme->background)
   ghTopTabs.CloseBackColorHot = pTheme->background

   pTheme = getThemePtr("statusbar.panel")
   if ghStatusBar.hPanelBrush then DeleteObject(ghStatusBar.hPanelBrush)
   ghStatusBar.hPanelBrush = CreateSolidBrush(pTheme->background)
   pTheme = getThemePtr("statusbar.text")
   if ghStatusBar.hBackBrush then DeleteObject(ghStatusBar.hBackBrush)
   ghStatusBar.hBackBrush = CreateSolidBrush(pTheme->background)
   ghStatusBar.BackColor = pTheme->background
   ghStatusBar.ForeColor = pTheme->foreground
   pTheme = getThemePtr("statusbar.texthot")
   if ghStatusBar.hBackBrushHot then DeleteObject(ghStatusBar.hBackBrushHot)
   ghStatusBar.hBackBrushHot = CreateSolidBrush(pTheme->background)
   ghStatusBar.BackColorHot = pTheme->background
   ghStatusBar.ForeColorHot = pTheme->foreground

   pTheme = getThemePtr("findreplace.panel")
   if ghFindReplace.hPanelBrush then DeleteObject(ghFindReplace.hPanelBrush)
   ghFindReplace.hPanelBrush = CreateSolidBrush(pTheme->background)
   pTheme = getThemePtr("findreplace.text")
   if ghFindReplace.hBackBrush then DeleteObject(ghFindReplace.hBackBrush)
   ghFindReplace.hBackBrush = CreateSolidBrush(pTheme->background)
   ghFindReplace.BackColor = pTheme->background
   ghFindReplace.ForeColor = pTheme->foreground
   pTheme = getThemePtr("findreplace.texthot")
   if ghFindReplace.hBackBrushHot then DeleteObject(ghFindReplace.hBackBrushHot)
   ghFindReplace.hBackBrushHot = CreateSolidBrush(pTheme->background)
   ghFindReplace.BackColorHot = pTheme->background
   ghFindReplace.ForeColorHot = pTheme->foreground
   pTheme = getThemePtr("findreplace.leftedge")
   ghFindReplace.LeftEdge = pTheme->foreground
   pTheme = getThemePtr("findreplace.iconhot")
   if ghFindReplace.hIconBrushHot then DeleteObject(ghFindReplace.hIconBrushHot)
   ghFindReplace.hIconBrushHot = CreateSolidBrush(pTheme->background)
   ghFindReplace.IconBackColorHot = pTheme->background
   pTheme = getThemePtr("findreplace.textnotfound")
   ghFindReplace.NotFoundForeColor = pTheme->foreground
   pTheme = getThemePtr("findreplace.textbox")
   if ghFindReplace.hBackBrushTextBox then DeleteObject(ghFindReplace.hBackBrushTextBox)
   ghFindReplace.hBackBrushTextBox = CreateSolidBrush(pTheme->background)
   ghFindReplace.TextBoxBackColor = pTheme->background
   ghFindReplace.TextBoxForeColor = pTheme->foreground
   pTheme = getThemePtr("findreplace.iconselected")
   if ghFindReplace.hBackBrushIconSelected then DeleteObject(ghFindReplace.hBackBrushIconSelected)
   ghFindReplace.hBackBrushIconSelected = CreateSolidBrush(pTheme->background)
   ghFindReplace.IconSelectedBackColor = pTheme->background
   pTheme = getThemePtr("findreplace.cuebanner")
   ghFindReplace.CueBannerForeColor = pTheme->foreground

   pTheme = getThemePtr("output.panel")
   if ghOutput.hPanelBrush then DeleteObject(ghOutput.hPanelBrush)
   ghOutput.hPanelBrush = CreateSolidBrush(pTheme->background)
   pTheme = getThemePtr("output.text")
   if ghOutput.hBackBrush then DeleteObject(ghOutput.hBackBrush)
   ghOutput.hBackBrush = CreateSolidBrush(pTheme->background)
   ghOutput.BackColor = pTheme->background
   ghOutput.ForeColor = pTheme->foreground
   pTheme = getThemePtr("output.texthot")
   if ghOutput.hBackBrushHot then DeleteObject(ghOutput.hBackBrushHot)
   ghOutput.hBackBrushHot = CreateSolidBrush(pTheme->background)
   ghOutput.BackColorHot = pTheme->background
   ghOutput.ForeColorHot = pTheme->foreground
   pTheme = getThemePtr("output.divider")
   ghOutput.Divider = pTheme->foreground
   pTheme = getThemePtr("output.closehot")
   if ghOutput.hCloseBrushHot then DeleteObject(ghOutput.hCloseBrushHot)
   ghOutput.hCloseBrushHot = CreateSolidBrush(pTheme->background)
   ghOutput.CloseBackColorHot = pTheme->background

   pTheme = getThemePtr("designer.panel")
   if ghDesigner.hPanelBrush then DeleteObject(ghDesigner.hPanelBrush)
   ghDesigner.hPanelBrush = CreateSolidBrush(pTheme->background)

   ' try to set the default GUI font the application depending on Windows/Wine 
   ' and use fallback options. Start with Windows options because sometimes
   ' the user will have Windows fonts installed on their Linux system.
   dim as CWSTR wszFontName, wszSymbolFontName
   
   ' Segue UI *may* exist on Wine/Linux but most likely Segui UI Symbol
   ' will not, so set the regular gui and symbol fonts separately.
   if isFontNameExist("Segoe UI") then 
      wszFontName = "Segoe UI"
   elseif isFontNameExist("Noto Sans") then 
      wszFontName = "Noto Sans"
   elseif isFontNameExist("Open Sans") then 
      wszFontName = "Open Sans"
   elseif isFontNameExist("Tahoma") then 
      wszFontName = "Tahoma"
   else 
      wszFontName = "Arial"
   end if

   if isFontNameExist("Segoe UI Symbol") then 
      wszSymbolFontName = "Segoe UI Symbol"
   elseif isFontNameExist("Noto Sans Symbols2") then 
      wszSymbolFontName = "Noto Sans Symbols2"
   else 
      wszSymbolFontName = wszFontName
   end if
 
   if ghMenuBar.hFontMenuBar = 0 then
      ghMenuBar.hFontMenuBar = pWindow->CreateFont(wszFontName, 10, FW_NORMAL, , , , DEFAULT_CHARSET)
   end if
   if ghStatusBar.hFontStatusBar = 0 then
      ghStatusBar.hFontStatusBar = pWindow->CreateFont(wszFontName, 9, FW_NORMAL, , , , DEFAULT_CHARSET)
   end if
   if ghStatusBar.hFontStatusBarBold = 0 then
      ghStatusBar.hFontStatusBarBold = pWindow->CreateFont(wszFontName, 9, FW_BOLD, , , , DEFAULT_CHARSET)
   end if
   if ghMenuBar.hFontSymbolLargeBold = 0 then
      ghMenuBar.hFontSymbolLargeBold = pWindow->CreateFont(wszSymbolFontName, 12, FW_BOLD, , , , DEFAULT_CHARSET)
   end if
   if ghMenuBar.hFontSymbol = 0 then
      ghMenuBar.hFontSymbol = pWindow->CreateFont(wszSymbolFontName, 10, FW_NORMAL, , , , DEFAULT_CHARSET)
   end if
   if ghMenuBar.hFontSymbolSmall = 0 then
      ghMenuBar.hFontSymbolSmall = pWindow->CreateFont(wszSymbolFontName, 9, FW_NORMAL, , , , DEFAULT_CHARSET)
   end if
      
   ' set some global brush handles so that we don't always have to be
   ' recreating them every time they are used.
   pTheme = getThemePtr("main.panel")
   if ghBrushMainBackground then DeleteObject(ghBrushMainBackground)
   ghBrushMainBackground = CreateSolidBrush(pTheme->background)

   function = false
end function

