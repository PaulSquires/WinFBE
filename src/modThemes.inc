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

type EXPLORER_HANDLES_TYPE
   hPanelBrush as HBRUSH
   hBackBrushScrollBar as HBRUSH
   BackColorScrollBar as COLORREF
   ForeColorScrollBar as COLORREF
   hBackBrushThumb as HBRUSH
   hBackBrushNode as HBRUSH
   BackColorNode as COLORREF
   ForeColorNode as COLORREF
   hBackBrush as HBRUSH
   BackColor as COLORREF
   ForeColor as COLORREF
   hBackBrushHot as HBRUSH
   BackColorHot as COLORREF
   ForeColorHot as COLORREF
end type
dim shared ghExplorer as EXPLORER_HANDLES_TYPE

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

type FUNCTIONLIST_HANDLES_TYPE
   PanelColor as COLORREF
   hPanelBrush as HBRUSH
   hBackBrushScrollBar as HBRUSH
   BackColorScrollBar as COLORREF
   ForeColorScrollBar as COLORREF
   hBackBrushThumb as HBRUSH
   hBackBrushNode as HBRUSH
   BackColorNode as COLORREF
   ForeColorNode as COLORREF
   hBackBrush as HBRUSH
   BackColor as COLORREF
   ForeColor as COLORREF
   hBackBrushHot as HBRUSH
   BackColorHot as COLORREF
   ForeColorHot as COLORREF
   hBackBrushTextBox as HBRUSH
   TextBoxForeColor as COLORREF
   TextBoxBackColor as COLORREF
   CueBannerForeColor as COLORREF
end type
dim shared ghFunctionList as FUNCTIONLIST_HANDLES_TYPE

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
      wszFilename = AfxGetExePathName & "themes\visual_studio_dark.theme"
      if AfxFileExists(wszFilename) = 0 then
         ' If the theme still does not exist then most likely the Themes folder is
         ' missing or the theme was renamed. Use the backup version in the settings folder. 
         wszFilename = AfxGetExePathName & "settings\backup_visual_studio_dark.theme"
      end if
   end if

   dim pStream AS CTextStream   
   if pStream.Open(wszFilename) <> S_OK then return true   ' error

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
         keys(i) = AfxstrParse(keyValue, i, ".")
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
      ' the proporties added to it, or if the structure needs
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

   pTheme = getThemePtr("explorer.panel")
   if ghExplorer.hPanelBrush then DeleteObject(ghExplorer.hPanelBrush)
   ghExplorer.hPanelBrush = CreateSolidBrush(pTheme->background)
   pTheme = getThemePtr("explorer.text")
   if ghExplorer.hBackBrush then DeleteObject(ghExplorer.hBackBrush)
   ghExplorer.hBackBrush = CreateSolidBrush(pTheme->background)
   ghExplorer.BackColor = pTheme->background
   ghExplorer.ForeColor = pTheme->foreground
   pTheme = getThemePtr("explorer.texthot")
   if ghExplorer.hBackBrushHot then DeleteObject(ghExplorer.hBackBrushHot)
   ghExplorer.hBackBrushHot = CreateSolidBrush(pTheme->background)
   ghExplorer.BackColorHot = pTheme->background
   ghExplorer.ForeColorHot = pTheme->foreground
   pTheme = getThemePtr("explorer.nodetext")
   if ghExplorer.hBackBrushNode then DeleteObject(ghExplorer.hBackBrushNode)
   ghExplorer.hBackBrushNode = CreateSolidBrush(pTheme->background)
   ghExplorer.BackColorNode = pTheme->background
   ghExplorer.ForeColorNode = pTheme->foreground
   pTheme = getThemePtr("explorer.scrollbar")
   if ghExplorer.hBackBrushScrollBar then DeleteObject(ghExplorer.hBackBrushScrollBar)
   ghExplorer.hBackBrushScrollBar = CreateSolidBrush(pTheme->background)
   if ghExplorer.hBackBrushThumb then DeleteObject(ghExplorer.hBackBrushThumb)
   ghExplorer.hBackBrushThumb = CreateSolidBrush(pTheme->foreground)
   ghExplorer.BackColorScrollBar = pTheme->background
   ghExplorer.ForeColorScrollBar = pTheme->foreground
      
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

   pTheme = getThemePtr("functionlist.panel")
   ghFunctionList.PanelColor = pTheme->background
   if ghFunctionList.hPanelBrush then DeleteObject(ghFunctionList.hPanelBrush)
   ghFunctionList.hPanelBrush = CreateSolidBrush(ghFunctionList.PanelColor)
   pTheme = getThemePtr("functionlist.nodetext")
   if ghFunctionList.hBackBrushNode then DeleteObject(ghFunctionList.hBackBrushNode)
   ghFunctionList.hBackBrushNode = CreateSolidBrush(pTheme->background)
   ghFunctionList.BackColorNode = pTheme->background
   ghFunctionList.ForeColorNode = pTheme->foreground
   pTheme = getThemePtr("functionlist.text")
   if ghFunctionList.hBackBrush then DeleteObject(ghFunctionList.hBackBrush)
   ghFunctionList.hBackBrush = CreateSolidBrush(pTheme->background)
   ghFunctionList.BackColor = pTheme->background
   ghFunctionList.ForeColor = pTheme->foreground
   pTheme = getThemePtr("functionlist.texthot")
   if ghFunctionList.hBackBrushHot then DeleteObject(ghFunctionList.hBackBrushHot)
   ghFunctionList.hBackBrushHot = CreateSolidBrush(pTheme->background)
   ghFunctionList.BackColorHot = pTheme->background
   ghFunctionList.ForeColorHot = pTheme->foreground
   pTheme = getThemePtr("functionlist.textbox")
   if ghFunctionList.hBackBrushTextBox then DeleteObject(ghFunctionList.hBackBrushTextBox)
   ghFunctionList.hBackBrushTextBox = CreateSolidBrush(pTheme->background)
   ghFunctionList.TextBoxBackColor = pTheme->background
   ghFunctionList.TextBoxForeColor = pTheme->foreground
   pTheme = getThemePtr("functionlist.cuebanner")
   ghFunctionList.CueBannerForeColor = pTheme->foreground
   pTheme = getThemePtr("functionlist.scrollbar")
   if ghFunctionList.hBackBrushScrollBar then DeleteObject(ghFunctionList.hBackBrushScrollBar)
   ghExplorer.hBackBrushScrollBar = CreateSolidBrush(pTheme->background)
   if ghFunctionList.hBackBrushThumb then DeleteObject(ghFunctionList.hBackBrushThumb)
   ghFunctionList.hBackBrushThumb = CreateSolidBrush(pTheme->foreground)
   ghFunctionList.BackColorScrollBar = pTheme->background
   ghFunctionList.ForeColorScrollBar = pTheme->foreground

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

   if ghMenuBar.hFontMenuBar then DeleteObject(ghMenuBar.hFontMenuBar)
   ghMenuBar.hFontMenuBar = pWindow->CreateFont(wszFontName, 10, FW_NORMAL, , , , DEFAULT_CHARSET)
   
   if ghStatusBar.hFontStatusBar then DeleteObject(ghStatusBar.hFontStatusBar)
   ghStatusBar.hFontStatusBar = pWindow->CreateFont(wszFontName, 9, FW_NORMAL, , , , DEFAULT_CHARSET)

   if ghStatusBar.hFontStatusBarBold then DeleteObject(ghStatusBar.hFontStatusBarBold)
   ghStatusBar.hFontStatusBarBold = pWindow->CreateFont(wszFontName, 9, FW_BOLD, , , , DEFAULT_CHARSET)

   if ghMenuBar.hFontSymbolLargeBold then DeleteObject(ghMenuBar.hFontSymbolLargeBold)
   ghMenuBar.hFontSymbolLargeBold = pWindow->CreateFont(wszSymbolFontName, 12, FW_BOLD, , , , DEFAULT_CHARSET)

   if ghMenuBar.hFontSymbol then DeleteObject(ghMenuBar.hFontSymbol)
   ghMenuBar.hFontSymbol = pWindow->CreateFont(wszSymbolFontName, 10, FW_NORMAL, , , , DEFAULT_CHARSET)

   if ghMenuBar.hFontSymbolSmall then DeleteObject(ghMenuBar.hFontSymbolSmall)
   ghMenuBar.hFontSymbolSmall = pWindow->CreateFont(wszSymbolFontName, 9, FW_NORMAL, , , , DEFAULT_CHARSET)
      
   ' set some global brush handles so that we don't always have to be
   ' recreating them every time they are used.
   pTheme = getThemePtr("main.panel")
   if ghBrushMainBackground then DeleteObject(ghBrushMainBackground)
   ghBrushMainBackground = CreateSolidBrush(pTheme->background)

   function = false
end function

