' ########################################################################################
' Microsoft Windows
' File: AfxMenu.inc
' Contents: Menu wrapper functions
' Compiler: Free Basic 32 & 64 bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#INCLUDE ONCE "win/uxtheme.bi"
#INCLUDE ONCE "Afx/AfxWin.inc"
#INCLUDE ONCE "Afx/CWStr.inc"
USING Afx

' ========================================================================================
' Checks a menu item.
' - hMenu = A handle to the menu that contains the menu item.
' - uItem = The identifier or position of the menu item to get information about.
'           The meaning of this parameter depends on the value of fByPosition.
' - fByPosition = The meaning of uItem. If this parameter is FALSE, uItem is a menu item
'           identifier. Otherwise, it is a menu item position.
' Return Value: The return value specifies the previous state of the menu item (either
' MF_CHECKED or MF_UNCHECKED). If the menu item does not exist, the return value is -1.
' ========================================================================================
PRIVATE FUNCTION AfxCheckMenuItem (BYVAL hMenu AS HMENU, BYVAL uItem AS DWORD, BYVAL fByPosition AS LONG = FALSE) AS DWORD
   DIM dwFlags AS DWORD
   IF fByPosition THEN dwFlags = MF_BYPOSITION ELSE dwFlags = MF_BYCOMMAND
   dwFlags = dwFlags OR MF_CHECKED
   RETURN CheckMenuItem(hMenu, uItem, dwFlags)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Unchecks a menu item.
' - hMenu = A handle to the menu that contains the menu item.
' - uItem = The identifier or position of the menu item to get information about.
'           The meaning of this parameter depends on the value of fByPosition.
' - fByPosition = The meaning of uItem. If this parameter is FALSE, uItem is a menu item
'           identifier. Otherwise, it is a menu item position.
' Return Value: The return value specifies the previous state of the menu item (either
' MF_CHECKED or MF_UNCHECKED). If the menu item does not exist, the return value is -1.
' ========================================================================================
PRIVATE FUNCTION AfxUnCheckMenuItem (BYVAL hMenu AS HMENU, BYVAL uItem AS DWORD, BYVAL fByPosition AS LONG = FALSE) AS DWORD
   DIM dwFlags AS DWORD
   IF fByPosition THEN dwFlags = MF_BYPOSITION ELSE dwFlags = MF_BYCOMMAND
   dwFlags = dwFlags OR MF_UNCHECKED
   RETURN CheckMenuItem(hMenu, uItem, dwFlags)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Toggles the checked state of a menu item.
' ========================================================================================
PRIVATE FUNCTION AfxToggleMenuItem (BYVAL hMenu AS HMENU, BYVAL uItem AS DWORD, BYVAL fByPosition AS LONG = FALSE) AS DWORD
   DIM dwFlags AS DWORD
   IF fByPosition THEN dwFlags = MF_BYPOSITION ELSE dwFlags = MF_BYCOMMAND
   IF GetMenuState(hMenu, uItem, dwFlags) AND MF_CHECKED = MF_CHECKED THEN
      dwFlags = dwFlags OR MF_UNCHECKED
   ELSE
      dwFlags = dwFlags OR MF_CHECKED
   END IF
   RETURN CheckMenuItem(hMenu, uItem, dwFlags)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns TRUE if the specified menu item is checked; FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION AfxIsMenuItemChecked (BYVAL hMenu AS HMENU, BYVAL uItem AS DWORD, BYVAL fByPosition AS LONG = FALSE) AS BOOLEAN
   DIM dwFlags AS DWORD
   IF fByPosition THEN dwFlags = MF_BYPOSITION ELSE dwFlags = MF_BYCOMMAND
   IF GetMenuState(hMenu, uItem, dwFlags) AND MF_CHECKED = MF_CHECKED THEN RETURN TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns TRUE if the specified menu item is enabled; FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION AfxIsMenuItemEnabled (BYVAL hMenu AS HMENU, BYVAL uItem AS DWORD, BYVAL fByPosition AS LONG = FALSE) AS BOOLEAN
   DIM dwFlags AS DWORD
   IF fByPosition THEN dwFlags = MF_BYPOSITION ELSE dwFlags = MF_BYCOMMAND
   DIM dwRes AS DWORD = GetMenuState(hMenu, uItem, dwFlags)
   IF ((dwRes AND MF_DISABLED) <> MF_DISABLED) AND ((dwRes AND MF_GRAYED) <> MF_GRAYED) THEN RETURN TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns TRUE if the specified menu item is disabled; FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION AfxIsMenuItemDisabled (BYVAL hMenu AS HMENU, BYVAL uItem AS DWORD, BYVAL fByPosition AS LONG = FALSE) AS BOOLEAN
   DIM dwFlags AS DWORD
   IF fByPosition THEN dwFlags = MF_BYPOSITION ELSE dwFlags = MF_BYCOMMAND
   DIM dwRes AS DWORD = GetMenuState(hMenu, uItem, dwFlags)
   IF ((dwRes AND MF_DISABLED) = MF_DISABLED) THEN RETURN TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns TRUE if the specified menu item is grayed; FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION AfxIsMenuItemGrayed (BYVAL hMenu AS HMENU, BYVAL uItem AS DWORD, BYVAL fByPosition AS LONG = FALSE) AS BOOLEAN
   DIM dwFlags AS DWORD
   IF fByPosition THEN dwFlags = MF_BYPOSITION ELSE dwFlags = MF_BYCOMMAND
   DIM dwRes AS DWORD = GetMenuState(hMenu, uItem, dwFlags)
   IF ((dwRes AND MF_GRAYED) = MF_GRAYED) THEN RETURN TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns TRUE if the specified menu item is highlighted; FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION AfxIsMenuItemHighlighted (BYVAL hMenu AS HMENU, BYVAL uItem AS DWORD, BYVAL fByPosition AS LONG = FALSE) AS BOOLEAN
   DIM dwFlags AS DWORD
   IF fByPosition THEN dwFlags = MF_BYPOSITION ELSE dwFlags = MF_BYCOMMAND
   DIM dwRes AS DWORD = GetMenuState(hMenu, uItem, dwFlags)
   IF ((dwRes AND MF_HILITE) = MF_HILITE) THEN RETURN TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns TRUE if the specified menu item is a separator; FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION AfxIsMenuItemSeparator (BYVAL hMenu AS HMENU, BYVAL uItem AS DWORD, BYVAL fByPosition AS LONG = FALSE) AS BOOLEAN
   DIM dwFlags AS DWORD
   IF fByPosition THEN dwFlags = MF_BYPOSITION ELSE dwFlags = MF_BYCOMMAND
   DIM dwRes AS DWORD = GetMenuState(hMenu, uItem, dwFlags)
   IF ((dwRes AND MF_SEPARATOR) = MF_SEPARATOR) THEN RETURN TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns TRUE if the specified menu item is a submenu; FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION AfxIsMenuItemPopup (BYVAL hMenu AS HMENU, BYVAL uItem AS DWORD, BYVAL fByPosition AS LONG = FALSE) AS BOOLEAN
   DIM dwFlags AS DWORD
   IF fByPosition THEN dwFlags = MF_BYPOSITION ELSE dwFlags = MF_BYCOMMAND
   DIM dwRes AS DWORD = GetMenuState(hMenu, uItem, dwFlags)
   IF ((dwRes AND MF_POPUP) = MF_POPUP) THEN RETURN TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns TRUE if the specified menu item is ownerdraw; FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION AfxIsMenuItemOwnerDraw (BYVAL hMenu AS HMENU, BYVAL uItem AS DWORD, BYVAL fByPosition AS LONG = FALSE) AS BOOLEAN
   DIM dwFlags AS DWORD
   IF fByPosition THEN dwFlags = MF_BYPOSITION ELSE dwFlags = MF_BYCOMMAND
   DIM dwRes AS DWORD = GetMenuState(hMenu, uItem, dwFlags)
   IF ((dwRes AND MF_OWNERDRAW) = MF_OWNERDRAW) THEN RETURN TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the state of the specified menu item.
' - hMenu = A handle to the menu that contains the menu item.
' - uItem = The identifier or position of the menu item to get information about.
'           The meaning of this parameter depends on the value of fByPosition.
' - fByPosition = The meaning of uItem. If this parameter is FALSE, uItem is a menu item
'           identifier. Otherwise, it is a menu item position.
' Return Value: 0 on failure or one or more of the following values:
' - MFS_CHECKED   The item is checked
' - MFS_DEFAULT   The menu item is the default.
' - MFS_DISABLED  The item is disabled.
' - MFS_ENABLED   The item is enabled.
' - MFS_GRAYED    The item is grayed.
' - MFS_HILITE    The item is highlighted
' - MFS_UNCHECKED The item is unchecked.
' - MFS_UNHILITE  The item is not highlighed.
' Note: To get extended error information, use the GetLastError function.
' ========================================================================================
PRIVATE FUNCTION AfxGetMenuItemState (BYVAL hMenu AS HMENU, BYVAL uItem AS DWORD, BYVAL fByPosition AS LONG = FALSE) AS DWORD
   DIM mii AS MENUITEMINFOW
   mii.cbSize = SIZEOF(mii)
   mii.fMask = MIIM_STATE
   IF GetMenuItemInfoW(hMenu, uItem, fByPosition, @mii) = 0 THEN RETURN 0
   RETURN mii.fState
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the state of the specified menu item.
' - hMenu = A handle to the menu that contains the menu item.
' - uItem = The identifier or position of the menu item to get information about.
'           The meaning of this parameter depends on the value of fByPosition.
' - fState = The menu item state. It can be one or more of these values:
' - MFS_CHECKED   Checks the menu item.
' - MFS_DEFAULT   Specifies that the menu item is the default.
' - MFS_DISABLED  Disables the menu item and grays it so that it cannot be selected.
' - MFS_ENABLED   Enables the menu item so that it can be selected. This is the default state.
' - MFS_GRAYED    Disables the menu item and grays it so that it cannot be selected.
' - MFS_HILITE    Highlights the menu item.
' - MFS_UNCHECKED Unchecks the menu item.
' - MFS_UNHILITE  Removes the highlight from the menu item. This is the default state.
' - fByPosition = The meaning of uItem. If this parameter is FALSE, uItem is a menu item
'                 identifier. Otherwise, it is a menu item position.
' Return Value: TRUE or FALSE. To get extended error information, use the GetLastError function.
' Note: The application must call the DrawMenuBar function whenever a menu changes,
' whether or not the menu is in a displayed window.
' ========================================================================================
PRIVATE FUNCTION AfxSetMenuItemState (BYVAL hMenu AS HMENU, BYVAL uItem AS DWORD, BYVAL fState AS DWORD, BYVAL fByPosition AS LONG = FALSE) AS BOOLEAN
   DIM mii AS MENUITEMINFOW
   mii.cbSize = SIZEOF(mii)
   mii.fMask = MIIM_STATE
   mii.fState = fState
   RETURN SetMenuItemInfoW(hMenu, uItem, fByPosition, @mii)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Enables the specified menu item.
' - hMenu = A handle to the menu that contains the menu item.
' - uItem = The identifier or position of the menu item to get information about.
'           The meaning of this parameter depends on the value of fByPosition.
' - fByPosition = The meaning of uItem. If this parameter is FALSE, uItem is a menu item
'                 identifier. Otherwise, it is a menu item position.
' Return Value: TRUE or FALSE. To get extended error information, use the GetLastError function.
' Note: The application must call the DrawMenuBar function whenever a menu changes,
' whether or not the menu is in a displayed window.
' ========================================================================================
PRIVATE FUNCTION AfxEnableMenuItem (BYVAL hMenu AS HMENU, BYVAL uItem AS DWORD, BYVAL fByPosition AS LONG = FALSE) AS BOOLEAN
   DIM mii AS MENUITEMINFOW
   mii.cbSize = SIZEOF(mii)
   mii.fMask = MIIM_STATE
   mii.fState = MFS_ENABLED
   RETURN SetMenuItemInfoW(hMenu, uItem, fByPosition, @mii)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Disables the specified menu item.
' - hMenu = A handle to the menu that contains the menu item.
' - uItem = The identifier or position of the menu item to get information about.
'           The meaning of this parameter depends on the value of fByPosition.
' - fByPosition = The meaning of uItem. If this parameter is FALSE, uItem is a menu item
'                 identifier. Otherwise, it is a menu item position.
' Return Value: TRUE or FALSE. To get extended error information, use the GetLastError function.
' Note: The application must call the DrawMenuBar function whenever a menu changes,
' whether or not the menu is in a displayed window.
' ========================================================================================
PRIVATE FUNCTION AfxDisableMenuItem (BYVAL hMenu AS HMENU, BYVAL uItem AS DWORD, BYVAL fByPosition AS LONG = FALSE) AS BOOLEAN
   DIM mii AS MENUITEMINFOW
   mii.cbSize = SIZEOF(mii)
   mii.fMask = MIIM_STATE
   mii.fState = MFS_DISABLED
   RETURN SetMenuItemInfoW(hMenu, uItem, fByPosition, @mii)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Grays the specified menu item.
' - hMenu = A handle to the menu that contains the menu item.
' - uItem = The identifier or position of the menu item to get information about.
'           The meaning of this parameter depends on the value of fByPosition.
' - fByPosition = The meaning of uItem. If this parameter is FALSE, uItem is a menu item
'                 identifier. Otherwise, it is a menu item position.
' Return Value: TRUE or FALSE. To get extended error information, use the GetLastError function.
' Note: The application must call the DrawMenuBar function whenever a menu changes,
' whether or not the menu is in a displayed window.
' ========================================================================================
PRIVATE FUNCTION AfxGrayMenuItem (BYVAL hMenu AS HMENU, BYVAL uItem AS DWORD, BYVAL fByPosition AS LONG = FALSE) AS BOOLEAN
   DIM mii AS MENUITEMINFOW
   mii.cbSize = SIZEOF(mii)
   mii.fMask = MIIM_STATE
   mii.fState = MFS_GRAYED
   RETURN SetMenuItemInfoW(hMenu, uItem, fByPosition, @mii)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Highlights the specified menu item.
' - hMenu = A handle to the menu that contains the menu item.
' - uItem = The identifier or position of the menu item to get information about.
'           The meaning of this parameter depends on the value of fByPosition.
' - fByPosition = The meaning of uItem. If this parameter is FALSE, uItem is a menu item
'                 identifier. Otherwise, it is a menu item position.
' Return Value: TRUE or FALSE. To get extended error information, use the GetLastError function.
' Note: The application must call the DrawMenuBar function whenever a menu changes,
' whether or not the menu is in a displayed window.
' ========================================================================================
PRIVATE FUNCTION AfxHiliteMenuItem (BYVAL hMenu AS HMENU, BYVAL uItem AS DWORD, BYVAL fByPosition AS LONG = FALSE) AS BOOLEAN
   DIM mii AS MENUITEMINFOW
   mii.cbSize = SIZEOF(mii)
   mii.fMask = MIIM_STATE
   mii.fState = MFS_HILITE
   RETURN SetMenuItemInfoW(hMenu, uItem, fByPosition, @mii)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Removes the system menu close option and disables the X button.
' Parameter: hwnd = Handle of the window that owns the menu.
' Return value: TRUE of FALSE.
' ========================================================================================
PRIVATE FUNCTION AfxRemoveCloseMenu (BYVAL hwnd AS HWND) AS BOOLEAN
   ' // Get the system menu handle
   DIM hMenu AS HMENU = GetSystemMenu(hwnd, 0)
   IF hMenu = NULL THEN RETURN FALSE
   ' // Get the number of menu items
   DIM cbItems AS LONG = GetMenuItemCount(hMenu)
   IF cbItems = 0 THEN RETURN FALSE
   ' // Remove the close menu item
   IF RemoveMenu(hMenu, cbItems - 1, MF_REMOVE OR MF_BYPOSITION) = 0 THEN RETURN FALSE
   ' // Remove the separator line
   IF RemoveMenu(hMenu, cbItems - 2, MF_REMOVE OR MF_BYPOSITION) = 0 THEN RETURN FALSE
   ' // Redraw the menu (this refreshes the caption bar, dimming the X button)
   DrawMenuBar(hwnd)
   RETURN TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Right justifies a top level menu item. This is usually used to have the Help menu item
' right-justified on the menu bar.
' - hwnd  = [in] A handle to the menu that contains the menu item.
' - uItem = [in] The zero-based position of the menu item to change.
' Return value: TRUE or FALSE.
' ========================================================================================
PRIVATE FUNCTION AfxRightJustifyMenuItem (BYVAL hMenu AS HMENU, BYVAL uItem AS DWORD) AS BOOLEAN
   DIM mii AS MENUITEMINFOW, buffer AS WSTRING * MAX_PATH + 1
   mii.cbSize = SIZEOF(MENUITEMINFOW)
   mii.dwTypeData = @buffer
   mii.cch = MAX_PATH
   mii.fType = MF_STRING
   mii.fState = MFS_DEFAULT
   mii.fMask = MIIM_ID OR MIIM_DATA OR MIIM_TYPE OR MIIM_SUBMENU
   IF GetMenuItemInfoW(hMenu, uItem, CTRUE, @mii) THEN
      mii.fType = mii.fType OR MF_HELP
      RETURN SetMenuItemInfoW(hMenu, uItem, CTRUE, @mii)
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Changes the text of a menu item to bold.
' - hwnd  = [in] A handle to the menu that contains the menu item.
' - uItem = [in] The zero-based position of the menu item to change.
' Return value: TRUE or FALSE.
' ========================================================================================
PRIVATE FUNCTION AfxSetMenuItemBold (BYVAL hMenu AS HMENU, BYVAL uItem AS DWORD) AS BOOLEAN
   DIM mii AS MENUITEMINFOW, buffer AS WSTRING * MAX_PATH + 1
   mii.cbSize = SIZEOF(MENUITEMINFOW)
   mii.dwTypeData = @buffer
   mii.cch = MAX_PATH
   mii.fType = MF_STRING
   mii.fMask = MIIM_ID OR MIIM_DATA OR MIIM_TYPE OR MIIM_SUBMENU OR MIIM_STATE
   IF GetMenuItemInfoW(hMenu, uItem, TRUE, @mii) THEN
      mii.fState = mii.fState OR &H1000
      RETURN SetMenuItemInfoW(hMenu, uItem, CTRUE, @mii)
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the text of the specified menu item.
' - hMenu = A handle to the menu that contains the menu item.
' - uItem = The identifier or position of the menu item to get information about.
'           The meaning of this parameter depends on the value of fByPosition.
' - wszText = Text to set.
' - fByPosition = The meaning of uItem. If this parameter is FALSE, uItem is a menu item
'           identifier. Otherwise, it is a menu item position.
' Return value: TRUE or FALSE.
' ========================================================================================
PRIVATE FUNCTION AfxSetMenuItemText (BYVAL hMenu AS HMENU, BYVAL uItem AS DWORD, BYREF wszText AS WSTRING, BYVAL fByPosition AS LONG = FALSE) AS BOOLEAN
   DIM mii AS MENUITEMINFOW
   mii.cbSize = SIZEOF(mii)
   mii.fMask = MIIM_STRING
   mii.dwTypeData = @wszText
   RETURN SetMenuItemInfoW(hMenu, uItem, fByPosition, @mii)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the lengnth of the specified menu item.
' ========================================================================================
PRIVATE FUNCTION AfxGetMenuItemTextLen (BYVAL hMenu AS HMENU, BYVAL uItem AS DWORD, BYVAL fByPosition AS LONG = FALSE) AS LONG
   ' // Fills the MENUITEMINFOW structure
   DIM mii AS MENUITEMINFOW
   mii.cbSize = SIZEOF(mii)
   mii.fMask = MIIM_STRING
   mii.dwTypeData = NULL
   ' // Get the needed size of the buffer
   IF GetMenuItemInfoW(hMenu, uItem, fByPosition, @mii) = 0 THEN RETURN 0
   RETURN mii.cch
END FUNCTION

' ========================================================================================
' Retrieves the text of the specified menu item.
' - hMenu = Handle to the menu that contains the menu item.
' - uItem = The identifier or position of the menu item to get information about.
'           The meaning of this parameter depends on the value of fByPosition.
' - fByPosition = The meaning of uItem. If this parameter is FALSE, uItem is a menu item
'           identifier. Otherwise, it is a menu item position.
' - pwszText: A pointer to a buffer to receive the retrieved text.
' - cchTextMax : Maximum number of characters to return. Both this value and the size of the
'      buffer pointed by pwszText must be one character bigger that the wanted length of the text
'      to return to make room for the null character terminator.
' Return value: TRUE or FALSE.
' Usage example:
' DIM wsz AS WSTRING * 260
' AfxGetMenuItemText(hSubMenu, 1, TRUE, @wsz, 260)
' AfxMsg wsz
' ========================================================================================
PRIVATE FUNCTION AfxGetMenuItemText OVERLOAD (BYVAL hMenu AS HMENU, BYVAL uItem AS DWORD, BYVAL fByPosition AS LONG, BYVAL pwszText AS WSTRING PTR, BYVAL cchTextMax AS LONG) AS BOOLEAN
   IF pwszText = NULL THEN RETURN FALSE
   ' // Fills the MENUITEMINFOW structure
   DIM mii AS MENUITEMINFOW
   mii.cbSize = SIZEOF(mii)
   mii.fMask = MIIM_STRING
   mii.dwTypeData = pwszText
   mii.cch = cchTextMax
   RETURN GetMenuItemInfoW(hMenu, uItem, fByPosition, @mii)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION AfxGetMenuItemText OVERLOAD (BYVAL hMenu AS HMENU, BYVAL uItem AS DWORD, BYVAL fByPosition AS LONG) AS CWSTR
   DIM mii AS MENUITEMINFOW
   mii.cbSize = SIZEOF(mii)
   mii.fMask = MIIM_STRING
   mii.dwTypeData = NULL
   ' // Get the needed size of the buffer
   IF GetMenuItemInfoW(hMenu, uItem, fByPosition, @mii) = 0 THEN RETURN ""
   ' // Make room for the trailing null
   mii.cch += 1
   ' // Allocate the buffer
   DIM buffer AS CWSTR = SPACE(mii.cch)
   ' // Get the menu string
   mii.dwTypeData = cast(WSTRING PTR, @buffer)
   IF GetMenuItemInfoW(hMenu, uItem, fByPosition, @mii) THEN
      RETURN RTRIM(buffer, CHR(0))
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves information about the font used in menu bars.
' Return value: TRUE on succes or FALSE on failure.
' If the function fails, the return value is zero.
' To get extended error information, call GetLastError.
' ========================================================================================
PRIVATE FUNCTION AfxGetMenuFont (BYVAL plfw AS LOGFONTW PTR) AS BOOLEAN
   DIM ncm AS NONCLIENTMETRICSW
   IF plfw = NULL THEN RETURN FALSE
   IF AfxWindowsVersion >= 6 THEN ncm.cbSize = 504 ELSE ncm.cbSize = 500
   IF SystemParametersInfoW(SPI_GETNONCLIENTMETRICS, SIZEOF(ncm), @ncm, 0) = 0 THEN RETURN FALSE
   *plfw = ncm.lfMenuFont
   RETURN TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the point size of the font used in menu bars.
' If the function fails, the return value is 0.
' ========================================================================================
PRIVATE FUNCTION AfxGetMenuFontPointSize () AS LONG
   DIM ncm AS NONCLIENTMETRICSW
   IF AfxWindowsVersion >= 6 THEN ncm.cbSize = 504 ELSE ncm.cbSize = 500
   IF SystemParametersInfoW(SPI_GETNONCLIENTMETRICS, SIZEOF(ncm), @ncm, 0) = 0 THEN RETURN 0
   DIM hDC AS HDC = CreateDCW("DISPLAY", NULL, NULL, NULL)
   IF hDC = NULL THEN RETURN 0
   DIM cyPixelsPerInch AS LONG = GetDeviceCaps(hDC, LOGPIXELSY)
   DeleteDC hDC
   DIM nPointSize AS LONG = MulDiv(ncm.lfMenuFont.lfHeight, 72, cyPixelsPerInch)
   IF nPointSize < 0 THEN nPointSize = -nPointSize
   RETURN nPointSize
END FUNCTION
' ========================================================================================

' ========================================================================================
' Calculates the size of a menu bar or a drop-down menu.
' - hwnd = Handle of the window that owns the menu.
' - hmenu = Handle of the menu.
' - prcmenu = Pointer to a variable of type RECT where to return the retrieved values.
' Return Value:
' If the function succeeds, the return value is 0.
' If the function fails, the return value is a system error code.
' ========================================================================================
PRIVATE FUNCTION AfxGetMenuRect (BYVAL hwnd AS HWND, BYVAL hmenu AS HMENU, BYVAL prcmenu AS RECT PTR) AS LONG
   DIM i AS LONG, nRes AS LONG, rc AS RECT
   FOR i = 1 TO GetMenuItemCount(hmenu)
      nRes = GetMenuItemRect(hwnd, hmenu, i, @rc)
      IF nRes = -1 THEN nRes = GetLastError : EXIT FOR
      UnionRect prcmenu, prcmenu, @rc
   NEXT
   RETURN nRes
END FUNCTION
' ========================================================================================

' ########################################################################################
'                              *** VISUAL STYLE MENUS ***
' ########################################################################################

' Windows Vista and posterior Windows versions provide menus that are part of the visual
' schema. These menus are rendered using visual styles, which can be added to existing
' applications. Adding code for new features to existing code must be done carefully to
' avoid breaking existing application behavior. Certain situations can cause visual styling
' to be disabled in an application. These situations include:
'    - Customizing menus using owner-draw menu items (MFT_OWNERDRAW)
'    - Using menu breaks (MFT_MENUBREAK or MFT_MENUBARBREAK)
'    - Using HBMMENU_CALLBACK to defer bitmap rendering
'    - Using a destroyed menu handle
' These situations prevent visual style menus from being rendered. Owner-draw menus can be
' used in Windows Vista and posterior Windows versions, but the menus will not be visually
' styled.
' Windows Vista and posterior Windows versions provide alpha-blended bitmaps, which enables
' menu items to be shown without using owner-draw menu items.
' Requirements:
'    - The bitmap is a 32bpp DIB section.
'    - The DIB section has BI_RGB compression.
'    - The bitmap contains pre-multiplied alpha pixels.
'    - The bitmap is stored in hbmpChecked, hbmpUnchecked, or hbmpItem fields.
' Note: MFT_BITMAP items do not support PARGB32 bitmaps.
' The following functions use the the Graphics Device Interface (GDI) to convert icons to
' bitmaps. Another solution is to use the Windows Imaging Component (WIC).
' Usage example:
'    DIM hSubMenu AS HMENU = GetSubMenu(hMenu, 1)
'    DIM hIcon AS HICON
'    hIcon = LoadImageW(NULL, "MyIcon.ico", IMAGE_ICON, 32, 32, LR_LOADFROMFILE)
'    IF hIcon THEN AfxAddIconToMenuItem(hSubMenu, 0, TRUE, hIcon)
' PNG icons can be used by converting them to an icon with AfxGdipImageFromFileEx:
'    hIcon = AfxGdipImageFromFileEx("MyIcon.png")
'    IF hIcon THEN AfxAddIconToMenuItem(hSubMenu, 0, TRUE, hIcon)

#if _WIN32_WINNT = &h0602
' ========================================================================================
' Initializes the BITMAPINFO structure
' Return value: TRUE or FALSE.
' ========================================================================================
PRIVATE FUNCTION AfxInitBitmapInfo (BYVAL pbmi AS BITMAPINFO PTR, BYVAL cbInfo AS ULONG, BYVAL cx AS LONG, BYVAL cy AS LONG, BYVAL bpp AS WORD) AS BOOLEAN
   IF pbmi = NULL THEN RETURN FALSE
   memset(pbmi, 0, cbInfo)
   pbmi->bmiHeader.biSize = SIZEOF(BITMAPINFOHEADER)
   pbmi->bmiHeader.biPlanes = 1
   pbmi->bmiHeader.biCompression = BI_RGB
   pbmi->bmiHeader.biWidth = cx
   pbmi->bmiHeader.biHeight = cy
   pbmi->bmiHeader.biBitCount = bpp
   RETURN TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates a 32bpp and a BI_RGB DIB section.
' Paramaters:
' - hDC     = Handle to the device context.
' - psize   = Size of the icon
' - ppvBits = A pointer to a variable that receives a pointer to the location of the
'             DIB bit values. Can be NULL.
' - phbmp   = A pointer to a variable that receives the handle to the newly created DIB.
' Return value: TRUE or FALSE.
' ========================================================================================
PRIVATE FUNCTION AfxCreate32BitHBITMAP(BYVAL hDC AS HDC, BYVAL psize AS SIZE PTR, BYVAL ppvBits AS ANY PTR PTR, BYVAL phBmp AS HBITMAP PTR) AS BOOLEAN
   IF psize = NULL OR phBmp = NULL THEN RETURN FALSE
   *phBmp = NULL
   DIM bmi AS BITMAPINFO
   AfxInitBitmapInfo(@bmi, SIZEOF(bmi), psize->cx, psize->cy, 32)
   DIM hdcUsed AS HDC = IIF(hDC <> NULL, hDC, GetDC(NULL))
   IF hdcUsed = NULL THEN RETURN FALSE
   *phBmp = CreateDIBSection(hdcUsed, @bmi, DIB_RGB_COLORS, ppvBits, NULL, 0)
   ReleaseDC(NULL, hdcUsed)
   RETURN TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Adds a bitmap to the menu item.
' Parameters:
' - hMenu       = A handle to the menu that contains the menu item.
' - nMenuItem   = The identifier or position of the menu item to change.
'                 The meaning of this parameter depends on the value of fByPosition.
' - fByPosition = The meaning of nMenuItem. If this parameter is FALSE, nMenuItem is a
'                 menu item identifier. Otherwise, it is a menu item position.
' - hbmp        = Bitmap handle.
' Return value: TRUE or FALSE.
' ========================================================================================
PRIVATE FUNCTION AfxAddBitmapToMenuItem (BYVAL hMenu AS HMENU, BYVAL nMenuItem AS LONG, BYVAL fByPosition AS BOOLEAN, BYVAL hbmp AS HBITMAP) AS BOOLEAN
   DIM mii AS MENUITEMINFOW
   mii.cbSize = SIZEOF(mii)
   mii.fMask = MIIM_BITMAP
   mii.hbmpItem = hbmp
   IF SetMenuItemInfoW(hMenu, nMenuItem, fByPosition, @mii) = 0 THEN RETURN FALSE
   RETURN TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Convert a bitmap to a PARGB32 bitmap.
' Parameters:
' - hDC       = Handle of the device context.
' - pargb     = Pointer to the address of the buffer bitmap pixels.
' - hbmp      = Bitmap handle.
' - sizeImage = Size of the bitmap.
' - cxRow     = The width, in pixels, of the buffer bitmap.
'               This value is not necessarily equal to the buffer width. It may be larger.
' Return value: TRUE or FALSE.
' Remarks: MFT_BITMAP items do not support PARGB32 bitmaps.
' ========================================================================================
PRIVATE FUNCTION AfxConvertToPARGB32 (BYVAL hDC AS HDC, BYVAL pargb AS DWORD PTR, BYVAL hbmp AS HBITMAP, BYVAL sizeImage AS SIZE, BYVAL cxRow AS LONG) AS BOOLEAN
   IF hDC = NULL OR pargb = NULL OR hbmp = NULL THEN RETURN FALSE
   DIM bmi AS BITMAPINFO
   AfxInitBitmapInfo(@bmi, SIZEOF(bmi), sizeImage.cx, sizeImage.cy, 32)
   DIM hHeap AS HANDLE = GetProcessHeap
   DIM pvBits AS ANY PTR = HeapAlloc(hHeap, 0, bmi.bmiHeader.biWidth * 4 * bmi.bmiHeader.biHeight)
   IF pvBits THEN
      IF GetDIBits(hDC, hbmp, 0, bmi.bmiHeader.biHeight, pvBits, @bmi, DIB_RGB_COLORS) = bmi.bmiHeader.biHeight THEN
         DIM cxDelta AS ULONG = cxRow - bmi.bmiHeader.biWidth
         DIM pargbMask AS DWORD PTR = CAST(DWORD PTR, pvBits)
         DIM x AS LONG, y AS LONG
         FOR y = bmi.bmiHeader.biHeight TO 1 STEP -1
            FOR x = bmi.bmiHeader.biWidth TO 1 STEP -1
               IF *pargbMask THEN
                  ' // transparent pixel
                  *pargb = 0
                  pargbMask +=1
                  pargb += 1
               ELSE
                  ' // opaque pixel
                  *pargb OR= &hFF000000
                  pargb +=1
               END IF
               pargb += cxDelta
            NEXT
         NEXT
      END IF
   END IF
   HeapFree(hHeap, 0, pvBits)
   RETURN TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Checks if the bitmap has the alpha channel set.
' Parameters:
' - pargb     = Pointer to the address of the buffer bitmap pixels.
' - sizeImage = Size of the bitmap.
' - cxRow     = The width, in pixels, of the buffer bitmap.
'               This value is not necessarily equal to the buffer width. It may be larger.
' Return value: TRUE or FALSE.
' ========================================================================================
PRIVATE FUNCTION AfxHasAlpha (BYVAL pargb AS DWORD PTR, BYVAL sizeImage AS SIZE, BYVAL cxRow AS LONG) AS BOOLEAN
   IF pargb = NULL THEN RETURN FALSE
   DIM cxDelta AS ULONG = cxRow - sizeImage.cx
   DIM x AS LONG, y AS LONG
   FOR y = sizeImage.cy TO 1 STEP -1
      FOR x = sizeImage.cx TO 1 STEP -1
         IF (*pargb AND &hFF000000) THEN RETURN TRUE
         pargb += 1
      NEXT
      pargb += cxDelta
   NEXT
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts a buffered bitmap to an PARGB32 bitmap.
' Parameters:
' - hPaintBuffer = The handle of the buffered paint context, obtained through BeginBufferedPaint.
' - hDC          = The Handle of the device context.
' - hIcon        = The icon handle.
' - sizeIcon     = The size of the icon.
' Return value: TRUE or FALSE.
' ========================================================================================
PRIVATE FUNCTION AfxConvertBufferToPARGB32 (BYVAL hPaintBuffer AS HPAINTBUFFER, BYVAL hDC AS HDC, BYVAL hIcon AS HICON, BYVAL sizeIcon AS SIZE) AS BOOLEAN
   IF hPaintBuffer = NULL OR hDC = NULL OR hIcon = NULL THEN RETURN FALSE
   DIM prgbQuad AS RGBQUAD PTR, cxRow AS LONG
   DIM hr AS HRESULT
   hr = GetBufferedPaintBits(hPaintBuffer, @prgbQuad, @cxRow)
   IF SUCCEEDED(hr) THEN
      DIM pargb AS DWORD PTR = CAST(DWORD PTR, prgbQuad)
      IF NOT AfxHasAlpha(pargb, sizeIcon, cxRow) THEN
         DIM info AS ICONINFO
         IF GetIconInfo(hicon, @info) THEN
            IF info.hbmMask THEN
               hr = AfxConvertToPARGB32(hDC, pargb, info.hbmMask, sizeIcon, cxRow)
               DeleteObject(info.hbmColor)
               DeleteObject(info.hbmMask)
            END IF
         END IF
      END IF
   END IF
   IF SUCCEEDED(hr) THEN RETURN TRUE ELSE RETURN FALSE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts a hIcon to a bitmap and adds it to the specified hbmpItem field of HMENU item.
' The caller is responsible for destroying the bitmap generated. The icon will be destroyed
' if fAutoDestroy is set to true. The hbmpItem field of the menu item can be used to keep
' track of the bitmap by passing NULL to phbmp.
' Parameters:
' - hMenu        = Menu handle that contains the item to which an icon will be added.
' - uItem        = The identifier or position of the menu item to change.
'                  The meaning of this parameter depends on the value of fByPosition.
' - fByPosition  = The meaning of nMenuItem. If this parameter is FALSE, nMenuItem is a
'                  menu item identifier. Otherwise, it is a menu item position.
' - hIcon        = Handle of the icon to add to the menu.
' - fAutoDestroy = TRUE (the default) or FALSE.
'                  If TRUE, AddIconToMenuItem destroys the icon before returning.
' - phbmp        = Location where the bitmap representation of the icon is stored. Can be NULL.
' ========================================================================================
PRIVATE FUNCTION AfxAddIconToMenuItem (BYVAL hMenu AS HMENU, BYVAL uItem AS DWORD, BYVAL fByPosition AS BOOLEAN, BYVAL hIcon AS HICON, BYVAL fAutoDestroy AS BOOLEAN = TRUE, BYVAL phbmp AS HBITMAP PTR = NULL) AS BOOLEAN
   IF hMenu = NULL OR hIcon = NULL THEN RETURN FALSE
   DIM hbmp AS HBITMAP, sizIcon AS SIZE, rcIcon AS RECT
   sizIcon.cx = GetSystemMetrics(SM_CXSMICON)
   sizIcon.cy = GetSystemMetrics(SM_CYSMICON)
   SetRect(@rcIcon, 0, 0, sizIcon.cx, sizIcon.cy)
   DIM hdcDest AS HDC = CreateCompatibleDC(NULL)
   IF hdcDest = NULL THEN RETURN FALSE
   DIM hr AS LONG = (AfxCreate32BitHBITMAP(hdcDest, @sizIcon, NULL, @hbmp))
   IF hr THEN
      DIM hbmpOld AS HBITMAP = CAST(HBITMAP, SelectObject(hdcDest, hbmp))
      IF hbmpOld THEN
         DIM bfAlpha AS BLENDFUNCTION = (AC_SRC_OVER, 0, 255, AC_SRC_ALPHA)
         DIM paintParams AS BP_PAINTPARAMS
         paintParams.cbSize = SIZEOF(paintParams)
         paintParams.dwFlags = BPPF_ERASE
         paintParams.pBlendFunction = @bfAlpha
         DIM hdcBuffer AS HDC
         DIM hPaintBuffer AS HPAINTBUFFER = BeginBufferedPaint(hdcDest, @rcIcon, BPBF_DIB, @paintParams, @hdcBuffer)
         IF hPaintBuffer THEN
            IF DrawIconEx(hdcBuffer, 0, 0, hIcon, sizIcon.cx, sizIcon.cy, 0, NULL, DI_NORMAL) THEN
               ' // If icon did not have an alpha channel, we need to convert buffer to PARGB32.
               hr = AfxConvertBufferToPARGB32(hPaintBuffer, hdcDest, hIcon, sizIcon)
            END IF
            ' // This will write the buffer contents to the destination bitmap.
            EndBufferedPaint(hPaintBuffer, TRUE)
         END IF
         SelectObject(hdcDest, hbmpOld)
      END IF
   END IF
   DeleteDC(hdcDest)
   IF hr THEN hr = AfxAddBitmapToMenuItem(hMenu, uItem, fByPosition, hbmp)
   IF hr = FALSE THEN DeleteObject(hbmp): hbmp = NULL
   IF fAutoDestroy THEN DestroyIcon(hIcon)
   IF phbmp THEN *phbmp = hbmp
   RETURN hr
END FUNCTION
#endif
' ========================================================================================

' ########################################################################################
