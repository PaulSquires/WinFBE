' ########################################################################################
' Microsoft Windows
' File: AfxPrinter.inc
' Contents: Printer wrapper functions.
' Compiler: FreeBasic 32 & 64-bit, Unicode.
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#INCLUDE ONCE "windows.bi"
' // Needed if user has not defined WIN_INCLUDEALL
#INCLUDE ONCE "win/winspool.bi"
#INCLUDE ONCE "win/shellapi.bi"
#INCLUDE ONCE "win/commdlg.bi"
' // ---------------------------------------------
#INCLUDE ONCE "Afx/AfxStr.inc"
USING Afx

' ========================================================================================
' Retrieves the name of the default printer.
' ========================================================================================
PRIVATE FUNCTION AfxGetDefaultPrinter () AS CWSTR
   DIM buffer AS WSTRING * MAX_PATH
   GetProfileStringW "WINDOWS", "DEVICE", "", buffer, SIZEOF(buffer)
   RETURN AfxStrParse(buffer, 1, ",")
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the name of the default printer driver.
' ========================================================================================
PRIVATE FUNCTION AfxGetDefaultPrinterDriver () AS CWSTR
   DIM buffer AS WSTRING * MAX_PATH
   GetProfileStringW "WINDOWS", "DEVICE", "", buffer, SIZEOF(buffer)
   RETURN AfxStrParse(buffer, 2, ",")
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the name of the default printer port.
' Use AfxGetPrinterPort to retrieve the port name for a named printer.
' ========================================================================================
PRIVATE FUNCTION AfxGetDefaultPrinterPort () AS CWSTR
   DIM buffer AS WSTRING * MAX_PATH
   GetProfileStringW "WINDOWS", "DEVICE", "", buffer, SIZEOF(buffer)
   RETURN AfxStrParse(buffer, 3, ",")
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the handle of the printer device context.
' The caller is resposible of deleting the returned handle with DeleteDC.
' Note: The same code using PRINTDLGEXW and PrintDlgExW returns the E_HANDLE error.
' ========================================================================================
PRIVATE FUNCTION AfxGetPrinterDC () AS HDC
   DIM pd AS PRINTDLGW
   pd.lStructSize = SIZEOF(PRINTDLGW)
   pd.Flags = PD_RETURNDEFAULT OR PD_RETURNDC
   DIM hr AS HRESULT = PrintDlgW(@pd)
   IF pd.hDevMode THEN GlobalFree(pd.hDevMode)
   IF pd.hDevNames THEN GlobalFree(pd.hDevNames)
   RETURN pd.hDC
END FUNCTION
' ========================================================================================

' ========================================================================================
' Displays the printer dialog. Returns TRUE or FALSE.
' The caller is resposible of deleting the returned HDC handle with DeleteDC.
' DIM flags AS DWORD, hdc AS HDC, nCopies AS WORD, nFromPage AS WORD, nToPage AS WORD, nMinPage AS WORD, nMaxPage AS WORD
' IF AfxPrinterDialog(hwnd, flags, hdc, nCopies, nFromPage, nToPage, nMinPage, nMaxPage) THEN
'   ' // do printing
'   IF hdc THEN DeleteDC hdc
' END IF
' ========================================================================================
PRIVATE FUNCTION AfxPrinterDialog (BYVAL hwndOwner AS HWND, BYREF flags AS DWORD, BYREF hdc AS HDC, BYREF nCopies AS WORD, _
   BYREF nFromPage AS WORD, BYREF nToPage AS WORD, BYREF nMinPage AS WORD, BYREF nMaxPage AS WORD) AS BOOLEAN
   DIM pd AS PRINTDLGW
   pd.lStructSize = SIZEOF(pd)
   pd.hwndOwner = hwndOwner
   pd.flags = flags OR PD_SHOWHELP
   pd.nCopies = nCopies
   pd.nFromPage = MAX(nFromPage, nMinPage)
   pd.nToPage = MAX(MIN(nToPage, nMaxPage), nMaxPage)
   pd.nMinPage = nMinPage
   pd.nMaxPage = nMaxPage
   IF PrintDlgW(@pd) THEN
      IF pd.hDevMode THEN GlobalFree(pd.hDevMode)
      IF pd.hDevNames THEN GlobalFree(pd.hDevNames)
      RETURN TRUE
   END IF
   RETURN FALSE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Printer setup dialog
' ========================================================================================
PRIVATE FUNCTION AfxPrinterSetup (BYVAL hwndOwner AS HWND) AS BOOLEAN
   DIM pd AS PRINTDLGW
   pd.lStructSize = SIZEOF(pd)
   pd.hwndOwner = hwndOwner
   pd.Flags = PD_SHOWHELP OR PD_PRINTSETUP
   IF PrintDlgW(@pd) THEN
      IF pd.hDevMode THEN GlobalFree(pd.hDevMode)
      IF pd.hDevNames THEN GlobalFree(pd.hDevNames)
      RETURN TRUE
   END IF
   RETURN FALSE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Opens an instance of Explorer with the Printers and Faxes folder selected.
' ========================================================================================
PRIVATE FUNCTION AfxOpenPrintersFolder () AS HINSTANCE
   RETURN ShellExecuteW(0, "open", "explorer.exe", "/e,::{2227A280-3AEA-1069-A2DE-08002B30309D}", NULL, SW_SHOWNORMAL)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a list with the ports that available for printing on a specified server.
' If pwszServerName is NULL, the function enumerates the local machine's printer ports.
' Names are separated with a carriage return and a line feed characters.
' ========================================================================================
PRIVATE FUNCTION AfxEnumPrinterPorts (BYVAL pwszServerName AS WSTRING PTR = NULL) AS CWSTR
   DIM i AS LONG, cbNeeded AS DWORD, cbReturned AS DWORD
   DIM Ports(ANY) AS PORT_INFO_1W
   DIM wstrNames AS CWSTR
   DIM dwLevel AS DWORD = 1
   EnumPortsW(pwszServerName, dwLevel, NULL, 0, @cbNeeded, @cbReturned)
   IF cbNeeded THEN
      REDIM Ports(0 TO cbNeeded \ SIZEOF(PORT_INFO_1))
      EnumPortsW pwszServerName, dwLevel, cast(BYTE PTR, VARPTR(Ports(0))), _
         SIZEOF(PORT_INFO_1W) * (UBOUND(Ports) + 1), @cbNeeded, @cbReturned
      FOR i = 0 TO cbReturned - 1
          wstrNames += *Ports(i).pName
          IF i < cbReturned - 1 THEN wstrNames += CHR(13, 10)
      NEXT
   END IF
   FUNCTION = wstrNames
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a list with port names of the locally installed printers.
' Names are separated with a carriage return and a line feed characters.
' ========================================================================================
PRIVATE FUNCTION AfxEnumLocalPrinterPorts () AS CWSTR
   DIM i AS LONG, cbNeeded AS DWORD, cbReturned AS DWORD
   DIM Pi5(ANY) AS PRINTER_INFO_5W
   DIM wstrNames AS CWSTR
   DIM dwLevel AS DWORD = 5
   EnumPrintersW PRINTER_ENUM_LOCAL, NULL, dwLevel, NULL, 0, @cbNeeded, @cbReturned
   REDIM Pi5(0 TO cbNeeded \ SIZEOF(PRINTER_INFO_5W))
   EnumPrintersW PRINTER_ENUM_LOCAL, "", dwLevel, cast(BYTE PTR, VARPTR(Pi5(0))), _
         SIZEOF(PRINTER_INFO_5W) * (UBOUND(Pi5) + 1), @cbNeeded, @cbReturned
   FOR i = 0 TO cbReturned - 1
      wstrNames += *Pi5(i).pPortName
      IF i < cbReturned - 1 THEN wstrNames += CHR(13, 10)
   NEXT
   RETURN wstrNames
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a list with the available printers, print servers, domains, or print providers.
' Names are separated with a carriage return and a line feed characters.
' ========================================================================================
PRIVATE FUNCTION AfxEnumPrinterNames () AS CWSTR
   DIM i AS LONG, cbNeeded AS DWORD, cbReturned AS DWORD
   DIM Pi5(ANY) AS PRINTER_INFO_5W
   DIM wstrNames AS CWSTR
   DIM dwLevel AS DWORD = 5
   EnumPrintersW PRINTER_ENUM_LOCAL, NULL, dwLevel, NULL, 0, @cbNeeded, @cbReturned
   REDIM Pi5(0 TO cbNeeded \ SIZEOF(PRINTER_INFO_5W))
   EnumPrintersW PRINTER_ENUM_LOCAL, "", dwLevel, cast(BYTE PTR, VARPTR(Pi5(0))), _
         SIZEOF(PRINTER_INFO_5W) * (UBOUND(Pi5) + 1), @cbNeeded, @cbReturned
   FOR i = 0 TO cbReturned - 1
      wstrNames += *Pi5(i).pPrinterName
      IF i < cbReturned - 1 THEN wstrNames += CHR(13, 10)
   NEXT
   RETURN wstrNames
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the port name for a given printer name.
' ========================================================================================
PRIVATE FUNCTION AfxGetPrinterPort (BYREF wszPrinterName AS WSTRING) AS CWSTR
   DIM i AS LONG, cbNeeded AS DWORD, cbReturned AS DWORD
   DIM Pi5(ANY) AS PRINTER_INFO_5W
   DIM dwLevel AS DWORD = 5
   EnumPrintersW PRINTER_ENUM_LOCAL, NULL, dwLevel, NULL, 0, @cbNeeded, @cbReturned
   REDIM Pi5(0 TO cbNeeded \ SIZEOF(PRINTER_INFO_5W))
   EnumPrintersW PRINTER_ENUM_LOCAL, "", dwLevel, cast(BYTE PTR, VARPTR(Pi5(0))), _
         SIZEOF(PRINTER_INFO_5W) * (UBOUND(Pi5) + 1), @cbNeeded, @cbReturned
   FOR i = 0 TO cbReturned - 1
      IF UCASE(*Pi5(i).pPrinterName) = UCASE(wszPrinterName) THEN RETURN *Pi5(i).pPortName
   NEXT
   RETURN ""
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the printer name for a given port name.
' Example: AfxGetPrinterFromPort(AfxGetPrinterPort(AfxGetDefaultPrinter))
' ========================================================================================
PRIVATE FUNCTION AfxGetPrinterFromPort (BYREF wszPortName AS WSTRING) AS CWSTR
   DIM i AS LONG, cbNeeded AS DWORD, cbReturned AS DWORD
   DIM wstrNames AS CWSTR
   DIM Pi5(ANY) AS PRINTER_INFO_5W
   DIM dwLevel AS DWORD = 5
   EnumPrintersW PRINTER_ENUM_LOCAL, NULL, dwLevel, NULL, 0, @cbNeeded, @cbReturned
   REDIM Pi5(0 TO cbNeeded \ SIZEOF(PRINTER_INFO_5W))
   EnumPrintersW PRINTER_ENUM_LOCAL, "", dwLevel, cast(BYTE PTR, VARPTR(Pi5(0))), _
         SIZEOF(PRINTER_INFO_5W) * (UBOUND(Pi5) + 1), @cbNeeded, @cbReturned
   FOR i = 0 TO cbReturned - 1
      IF UCASE(*Pi5(i).pPortName) = UCASE(wszPortName) THEN RETURN *Pi5(i).pPrinterName
   NEXT
   RETURN ""
END FUNCTION
' ========================================================================================

' ========================================================================================
' If the printer supports collating, the return value is TRUE; otherwise, the return value is FALSE.
' If TRUE, the pages that are printed should be collated. To collate is to print out the
' entire document before printing the next copy, as opposed to printing out each page of
' the document the required number of times.
' ========================================================================================
PRIVATE FUNCTION AfxGetPrinterCollate (BYREF wszPrinterName AS WSTRING) AS BOOLEAN
   RETURN DeviceCapabilitiesW(wszPrinterName, NULL, DC_COLLATE, NULL, NULL)
END FUNCTION
' ========================================================================================

' ========================================================================================
' If the printer supports color printing, the return value is TRUE; otherwise, the return value is FALSE.
' CSome color printers have the capability to print using true black instead of a combination
' of cyan, magenta, and yellow (CMY). This usually creates darker and sharper text for
' documents. This option is only useful for color printers that support true black printing.
' ========================================================================================
PRIVATE FUNCTION AfxGetPrinterColorMode (BYREF wszPrinterName AS WSTRING) AS BOOLEAN
   RETURN DeviceCapabilitiesW(wszPrinterName, NULL, DC_COLORDEVICE, NULL, NULL)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the version number of the Windows-based printer driver.
' The version numbers are created and maintained by the driver manufacturer.
' ========================================================================================
PRIVATE FUNCTION AfxGetPrinterDriverVersion (BYREF wszPrinterName AS WSTRING) AS LONG
   RETURN DeviceCapabilitiesW(wszPrinterName, NULL, DC_DRIVER, NULL, NULL)
END FUNCTION
' ========================================================================================

' ========================================================================================
' If the printer supports duplex printing, the return value is TRUE; otherwise, the return value is FALSE.
' ========================================================================================
PRIVATE FUNCTION AfxGetPrinterDuplexMode (BYREF wszPrinterName AS WSTRING) AS BOOLEAN
   RETURN DeviceCapabilitiesW(wszPrinterName, NULL, DC_DUPLEX, NULL, NULL)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the maximum number of copies the device can print.
' If the function returns -1, this may mean either that the capability is not supported or
' there was a general function failure.
' ========================================================================================
PRIVATE FUNCTION AfxGetPrinterMaxCopies (BYREF wszPrinterName AS WSTRING) AS LONG
   RETURN DeviceCapabilitiesW(wszPrinterName, NULL, DC_COPIES, NULL, NULL)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the names of the paper forms that are currently available for use.
' ========================================================================================
PRIVATE FUNCTION AfxGetPrinterMediaReady (BYREF wszPrinterName AS WSTRING) AS CWSTR
   DIM wstrNames AS CWSTR, wszNames(ANY) AS WSTRING * 64
   DIM r AS LONG = DeviceCapabilitiesW(wszPrinterName, NULL, DC_MEDIAREADY, NULL, NULL)
   IF r = -1 THEN RETURN ""
   REDIM wszNames(r - 1)
   r = DeviceCapabilitiesW(wszPrinterName, NULL, DC_MEDIAREADY, @wszNames(0), NULL)
   IF r < 1 THEN RETURN ""
   DIM i AS LONG
   FOR i = 0 TO r - 1
      wstrNames += wszNames(i)
      IF i < r - 1 THEN wstrNames += CHR(13, 10)
   NEXT
   RETURN wstrNames
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the maximum paper width in tenths of a millimeter.
' If the function returns -1, this may mean either that the capability is not supported or
' there was a general function failure.
' ========================================================================================
PRIVATE FUNCTION AfxGetPrinterMaxPaperHeight (BYREF wszPrinterName AS WSTRING) AS LONG
   DIM r AS LONG = DeviceCapabilitiesW(wszPrinterName, NULL, DC_MAXEXTENT, NULL, NULL)
   IF r <> -1 THEN RETURN HIWORD(r)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the maximum paper width in tenths of a millimeter.
' If the function returns -1, this may mean either that the capability is not supported or
' there was a general function failure.
' ========================================================================================
PRIVATE FUNCTION AfxGetPrinterMaxPaperWidth (BYREF wszPrinterName AS WSTRING) AS LONG
   DIM r AS LONG = DeviceCapabilitiesW(wszPrinterName, NULL, DC_MAXEXTENT, NULL, NULL)
   IF r <> -1 THEN RETURN LOWORD(r)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the minimum paper width in tenths of a millimeter.
' If the function returns -1, this may mean either that the capability is not supported or
' there was a general function failure.
' ========================================================================================
PRIVATE FUNCTION AfxGetPrinterMinPaperHeight (BYREF wszPrinterName AS WSTRING) AS LONG
   DIM r AS LONG = DeviceCapabilitiesW(wszPrinterName, NULL, DC_MINEXTENT, NULL, NULL)
   IF r <> -1 THEN RETURN HIWORD(r)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the minimum paper width in tenths of a millimeter.
' If the function returns -1, this may mean either that the capability is not supported or
' there was a general function failure.
' ========================================================================================
PRIVATE FUNCTION AfxGetPrinterMinPaperWidth (BYREF wszPrinterName AS WSTRING) AS LONG
   DIM r AS LONG = DeviceCapabilitiesW(wszPrinterName, NULL, DC_MINEXTENT, NULL, NULL)
   IF r <> -1 THEN RETURN LOWORD(r)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the relationship between portrait and landscape orientations for a device, in
' terms of the number of degrees that portrait orientation is rotated counterclockwise to
' produce landscape orientation. The return value can be one of the following:
'   0  No landscape orientation.
'  90  Portrait is rotated 90 degrees to produce landscape.
' 180  Portrait is rotated 180 degrees to produce landscape.
' 270  Portrait is rotated 270 degrees to produce landscape.
' If the function returns -1, this may mean either that the capability is not supported or
' there was a general function failure.
' ========================================================================================
PRIVATE FUNCTION AfxGetPrinterOrientationDegrees (BYREF wszPrinterName AS WSTRING) AS LONG
   RETURN DeviceCapabilitiesW(wszPrinterName, NULL, DC_ORIENTATION, NULL, NULL)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a list of supported paper names (for example, Letter or Legal).
' The names are separated by a carriage return and a line feed characters.
' ========================================================================================
PRIVATE FUNCTION AfxGetPrinterPaperNames (BYREF wszPrinterName AS WSTRING) AS CWSTR
   DIM wstrNames AS CWSTR, wszNames(ANY) AS WSTRING * 64
   DIM r AS LONG = DeviceCapabilitiesW(wszPrinterName, NULL, DC_PAPERNAMES, NULL, NULL)
   IF r = -1 THEN RETURN ""
   REDIM wszNames(r - 1) AS WSTRING * 64
   r = DeviceCapabilitiesW(wszPrinterName, NULL, DC_PAPERNAMES, @wszNames(0), NULL)
   IF r < 1 THEN RETURN ""
   DIM i AS LONG
   FOR i = 0 TO r - 1
      wstrNames += wszNames(i)
      IF i < r - 1 THEN wstrNames += CHR(13, 10)
   NEXT
   RETURN wstrNames
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a list with the names of the printer's paper bins.
' The names are separated by a carriage return and a line feed characters.
' ========================================================================================
PRIVATE FUNCTION AfxGetPrinterTrayNames (BYREF wszPrinterName AS WSTRING) AS CWSTR
   DIM wstrNames AS CWSTR, wszNames (ANY) AS WSTRING * 64
   DIM r AS LONG = DeviceCapabilitiesW(wszPrinterName, NULL, DC_BINNAMES, NULL, NULL)
   IF r = -1 THEN RETURN ""
   REDIM wszNames(r - 1)
   r = DeviceCapabilitiesW(wszPrinterName, NULL, DC_BINNAMES, @wszNames(0), NULL)
   IF r < 1 THEN RETURN ""
   DIM i AS LONG
   FOR i = 0 TO r - 1
      wstrNames += wszNames(i)
      IF i < r - 1 THEN wstrNames += CHR(13, 10)
   NEXT
   RETURN wstrNames
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a list of each supported paper sizes, in tenths of a millimeter.
' Each entry if formatted as "<width> x <height>" and separated by a carriage return and a
' line feed characters.
' ========================================================================================
PRIVATE FUNCTION AfxGetPrinterPaperSizes (BYREF wszPrinterName AS WSTRING) AS CWSTR
   DIM wstrSizes AS CWSTR
   DIM pSizes(ANY) AS POINT
   DIM r AS LONG = DeviceCapabilitiesW(wszPrinterName, NULL, DC_PAPERSIZE, NULL, NULL)
   IF r = -1 THEN RETURN ""
   REDIM pSizes(r - 1)
   r = DeviceCapabilitiesW(wszPrinterName, NULL, DC_PAPERSIZE, CAST(WSTRING PTR, VARPTR(pSizes(0))), NULL)
   IF r < 1 THEN RETURN ""
   DIM i AS LONG
   FOR i = 0 TO r - 1
      wstrSizes += STR(pSizes(i).x) & " x " & STR(pSizes(i).y)
      IF i < r - 1 THEN wstrSizes += CHR(13, 10)
   NEXT
   RETURN wstrSizes
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the printer's print rate.
' If the function returns -1, this may mean either that the capability is not supported or
' there was a general function failure.
' ========================================================================================
PRIVATE FUNCTION AfxGetPrinterRate (BYREF wszPrinterName AS WSTRING) AS LONG
   RETURN DeviceCapabilitiesW(wszPrinterName, NULL, DC_PRINTRATE, NULL, NULL)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the printer's print rate, in pages per minute.
' If the function returns -1, this may mean either that the capability is not supported or
' there was a general function failure.
' ========================================================================================
PRIVATE FUNCTION AfxGetPrinterRatePPM (BYREF wszPrinterName AS WSTRING) AS LONG
   RETURN DeviceCapabilitiesW(wszPrinterName, NULL, DC_PRINTRATEPPM, NULL, NULL)
END FUNCTION
' ========================================================================================

' ========================================================================================
' The printer's rate. Returns one of the following values:
' PRINTRATEUNIT_CPS   Characters per second.
' PRINTRATEUNIT_IPM   Inches per minute.
' PRINTRATEUNIT_LPM   Lines per minute.
' PRINTRATEUNIT_PPM   Pages per minute.
' If the function returns -1, this may mean either that the capability is not supported or
' there was a general function failure.
' ========================================================================================
PRIVATE FUNCTION AfxGetPrinterRateUnit (BYREF wszPrinterName AS WSTRING) AS LONG
   RETURN DeviceCapabilitiesW(wszPrinterName, NULL, DC_PRINTRATEUNIT, NULL, NULL)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the abilities of the driver to use TrueType fonts.
' The return value can be one or more of the following:
' DCTT_BITMAP   Device can print TrueType fonts as graphics.
' DCTT_DOWNLOAD Device can download TrueType fonts.
' DCTT_SUBDEV   Device can substitute device fonts for TrueType fonts.
' If the function returns -1, this may mean either that the capability is not supported or
' there was a general function failure.
' ========================================================================================
PRIVATE FUNCTION AfxGetPrinterTrueType (BYREF wszPrinterName AS WSTRING) AS LONG
   RETURN DeviceCapabilitiesW(wszPrinterName, NULL, DC_TRUETYPE, NULL, NULL)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves printer initialization information.
' ========================================================================================
PRIVATE FUNCTION AfxGetDocumentProperties (BYREF wszPrinterName AS WSTRING, BYVAL dmField AS DWORD) AS LONG
   ' // Start by opening the printer
   DIM hPrinter AS HANDLE
   IF OpenPrinterW(wszPrinterName, @hPrinter, NULL) = FALSE THEN RETURN 0
   ' // Allocate a buffer of the correct size
   DIM cbNeeded AS DWORD = DocumentPropertiesW(NULL, hPrinter, wszPrinterName, NULL, NULL, 0)
   DIM bufferDoc AS STRING = SPACE(cbNeeded)
   ' // Retrieve the printer configuration data
   DIM nRet AS LONG = DocumentPropertiesW(NULL, hPrinter, wszPrinterName, _
      cast(DEVMODEW PTR, STRPTR(bufferDoc)), NULL, DM_OUT_BUFFER)
   IF nRet = IDOK THEN
      ' // Cast it to a DEVMODEW structure
      DIM pDevMode AS DEVMODEW PTR = cast(DEVMODEW PTR, STRPTR(bufferDoc))
      ' // Finished with the printer
      ClosePrinter(hPrinter)
      ' // Return the requested value
      SELECT CASE dmField
         CASE DM_COLLATE       : RETURN pDevMode->dmCollate
         CASE DM_COPIES        : RETURN pDevMode->dmCopies
         CASE DM_ORIENTATION   : RETURN pDevMode->dmOrientation
         CASE DM_PAPERSIZE     : RETURN pDevMode->dmPaperSize
         CASE DM_PRINTQUALITY  : RETURN pDevMode->dmPrintQuality
         CASE DM_SCALE         : RETURN pDevMode->dmScale
         CASE DM_DEFAULTSOURCE : RETURN pDevMode->dmDefaultSource
         CASE DM_PAPERLENGTH   : RETURN pDevMode->dmPaperLength
         CASE DM_PAPERWIDTH    : RETURN pDevMode->dmPaperWidth
      END SELECT
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the printer collate status.
' DMCOLLATE_FALSE = Collate is turned off
' DMCOLLATE_TRUE = Collate is turned on
' ========================================================================================
PRIVATE FUNCTION AfxGetPrinterCollateStatus (BYREF wszPrinterName AS WSTRING) AS LONG
   RETURN AfxGetDocumentProperties(wszPrinterName, DM_COLLATE)
END FUNCTION

' ========================================================================================
' Returns the number of copies printed if the device supports multiple-page copies.
' ========================================================================================
PRIVATE FUNCTION AfxGetPrinterCopies (BYREF wszPrinterName AS WSTRING) AS LONG
   RETURN AfxGetDocumentProperties(wszPrinterName, DM_COPIES)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the printer orientation.
' The return value can be one of the following:
' - DMORIENT_PORTRAIT = Portrait
' - DMORIENT_LANDSCAPE = Landscape
' ========================================================================================
PRIVATE FUNCTION AfxGetPrinterOrientation (BYREF wszPrinterName AS WSTRING) AS LONG
   RETURN AfxGetDocumentProperties(wszPrinterName, DM_ORIENTATION)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Specifies the paper size for which the printer is currently configured.
' DMPAPER_LETTER = 1; DMPAPER_FIRST = DMPAPER_LETTER; DMPAPER_LETTERSMALL = 2; DMPAPER_TABLOID = 3;
' DMPAPER_LEDGER = 4; DMPAPER_LEGAL = 5; DMPAPER_STATEMENT = 6; DMPAPER_EXECUTIVE = 7; DMPAPER_A3 = 8;
' DMPAPER_A4 = 9; DMPAPER_A4SMALL = 10; DMPAPER_A5 = 11; DMPAPER_B4 = 12; DMPAPER_B5 = 13;
' DMPAPER_FOLIO = 14; DMPAPER_QUARTO = 15; DMPAPER_10X14 = 16; DMPAPER_11X17 = 17; DMPAPER_NOTE = 18;
' DMPAPER_ENV_9 = 19; DMPAPER_ENV_10 = 20; DMPAPER_ENV_11 = 21; DMPAPER_ENV_12 = 22; DMPAPER_ENV_14 = 23;
' DMPAPER_CSHEET = 24; DMPAPER_DSHEET = 25; DMPAPER_ESHEET = 26; DMPAPER_ENV_DL = 27; DMPAPER_ENV_C5 = 28;
' DMPAPER_ENV_C3 = 29; DMPAPER_ENV_C4 = 30; DMPAPER_ENV_C6 = 31; DMPAPER_ENV_C65 = 32; DMPAPER_ENV_B4 = 33;
' DMPAPER_ENV_B5 = 34; DMPAPER_ENV_B6 = 35; DMPAPER_ENV_ITALY = 36; DMPAPER_ENV_MONARCH = 37;
' DMPAPER_ENV_PERSONAL = 38; DMPAPER_FANFOLD_US = 39; DMPAPER_FANFOLD_STD_GERMAN = 40;
' DMPAPER_FANFOLD_LGL_GERMAN = 41; DMPAPER_ISO_B4 = 42; DMPAPER_JAPANESE_POSTCARD = 43; DMPAPER_9X11 = 44;
' DMPAPER_10X11 = 45; DMPAPER_15X11 = 46; DMPAPER_ENV_INVITE = 47; DMPAPER_RESERVED_48 = 48;
' DMPAPER_RESERVED_49 = 49; DMPAPER_LETTER_EXTRA = 50; DMPAPER_LEGAL_EXTRA = 51; DMPAPER_TABLOID_EXTRA = 52;
' DMPAPER_A4_EXTRA = 53; DMPAPER_LETTER_TRANSVERSE = 54; DMPAPER_A4_TRANSVERSE = 55;
' DMPAPER_LETTER_EXTRA_TRANSVERSE = 56; DMPAPER_A_PLUS = 57; DMPAPER_B_PLUS = 58; DMPAPER_LETTER_PLUS = 59;
' DMPAPER_A4_PLUS = 60; DMPAPER_A5_TRANSVERSE = 61; DMPAPER_B5_TRANSVERSE = 62; DMPAPER_A3_EXTRA = 63;
' DMPAPER_A5_EXTRA = 64; DMPAPER_B5_EXTRA = 65; DMPAPER_A2 = 66; DMPAPER_A3_TRANSVERSE = 67;
' DMPAPER_A3_EXTRA_TRANSVERSE = 68; DMPAPER_DBL_JAPANESE_POSTCARD = 69; DMPAPER_A6 = 70;
' DMPAPER_JENV_KAKU2 = 71; DMPAPER_JENV_KAKU3 = 72; DMPAPER_JENV_CHOU3 = 73; DMPAPER_JENV_CHOU4 = 74;
' DMPAPER_LETTER_ROTATED = 75; DMPAPER_A3_ROTATED = 76; DMPAPER_A4_ROTATED = 77; DMPAPER_A5_ROTATED = 78;
' DMPAPER_B4_JIS_ROTATED = 79; DMPAPER_B5_JIS_ROTATED = 80; DMPAPER_JAPANESE_POSTCARD_ROTATED = 81;
' DMPAPER_DBL_JAPANESE_POSTCARD_ROTATED = 82; DMPAPER_A6_ROTATED = 83; DMPAPER_JENV_KAKU2_ROTATED = 84;
' DMPAPER_JENV_KAKU3_ROTATED = 85; DMPAPER_JENV_CHOU3_ROTATED = 86; DMPAPER_JENV_CHOU4_ROTATED = 87;
' DMPAPER_B6_JIS = 88; DMPAPER_B6_JIS_ROTATED = 89; DMPAPER_12X11 = 90; DMPAPER_JENV_YOU4 = 91;
' DMPAPER_JENV_YOU4_ROTATED = 92; DMPAPER_P16K = 93; DMPAPER_P32K = 94; DMPAPER_P32KBIG = 95;
' DMPAPER_PENV_1 = 96; DMPAPER_PENV_2 = 97; DMPAPER_PENV_3 = 98; DMPAPER_PENV_4 = 99; DMPAPER_PENV_5 = 100;
' DMPAPER_PENV_6 = 101; DMPAPER_PENV_7 = 102; DMPAPER_PENV_8 = 103; DMPAPER_PENV_9 = 104;
' DMPAPER_PENV_10 = 105; DMPAPER_P16K_ROTATED = 106; DMPAPER_P32K_ROTATED = 107; DMPAPER_P32KBIG_ROTATED = 108;
' DMPAPER_PENV_1_ROTATED = 109; DMPAPER_PENV_2_ROTATED = 110; DMPAPER_PENV_3_ROTATED = 111;
' DMPAPER_PENV_4_ROTATED = 112; DMPAPER_PENV_5_ROTATED = 113; DMPAPER_PENV_6_ROTATED = 114;
' DMPAPER_PENV_7_ROTATED = 115; DMPAPER_PENV_8_ROTATED = 116; DMPAPER_PENV_9_ROTATED = 117;
' DMPAPER_PENV_10_ROTATED = 118; DMPAPER_LAST = DMPAPER_PENV_10_ROTATED; DMPAPER_USER = 256
' ========================================================================================
PRIVATE FUNCTION AfxGetPrinterPaperSize (BYREF wszPrinterName AS WSTRING) AS LONG
   RETURN AfxGetDocumentProperties(wszPrinterName, DM_PAPERSIZE)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the printer print quality mode.
' The return value can be one of the following:
' - DMRES_DRAFT  = Draft
' - DMRES_LOW    = Low
' - DMRES_MEDIUM = Medium
' - DMRES_HIGH   = High
' ========================================================================================
PRIVATE FUNCTION AfxGetPrinterQuality (BYREF wszPrinterName AS WSTRING) AS LONG
   RETURN AfxGetDocumentProperties(wszPrinterName, DM_PRINTQUALITY)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Specifies the factor by which the printed output is to be scaled. The apparent page size
' is scaled from the physical page size by a factor of dmScale /100. For example, a
' letter-sized page with a dmScale value of 50 would contain as much data as a page of
' 17- by 22-inches because the output text and graphics would be half their original
' height and width.
' ========================================================================================
PRIVATE FUNCTION AfxGetPrinterScale (BYREF wszPrinterName AS WSTRING) AS LONG
   RETURN AfxGetDocumentProperties(wszPrinterName, DM_SCALE)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Specifies the paper source.
' Returns one of the following values: DMBIN_UPPER = 1; DMBIN_LOWER = 2; DMBIN_MIDDLE = 3;
' DMBIN_MANUAL = 4; DMBIN_ENVELOPE = 5; DMBIN_ENVMANUAL = 6; DMBIN_AUTO = 7;
' DMBIN_TRACTOR = 8; DMBIN_SMALLFMT = 9; DMBIN_LARGEFMT = 10; DMBIN_LARGECAPACITY = 11;
' DMBIN_CASSETTE = 14; DMBIN_FORMSOURCE = 15
' ========================================================================================
PRIVATE FUNCTION AfxGetPrinterTray (BYREF wszPrinterName AS WSTRING) AS LONG
   RETURN AfxGetDocumentProperties(wszPrinterName, DM_DEFAULTSOURCE)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the paper length.
' ========================================================================================
PRIVATE FUNCTION AfxGetPrinterPaperLength (BYREF wszPrinterName AS WSTRING) AS LONG
   RETURN AfxGetDocumentProperties(wszPrinterName, DM_PAPERLENGTH)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Specifies the length of the paper, in units of 1/10 of a millimeter.
' ========================================================================================
PRIVATE FUNCTION AfxGetPrinterPaperWidth (BYREF wszPrinterName AS WSTRING) AS LONG
   RETURN AfxGetDocumentProperties(wszPrinterName, DM_PAPERWIDTH)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Width, in pixels, of the printable area of the page.
' ========================================================================================
PRIVATE FUNCTION AfxGetPrinterHorizontalResolution (BYREF wszPrinterName AS WSTRING) AS LONG
   DIM hdc AS HDC = CreateICW(NULL, wszPrinterName, NULL, NULL)
   IF hdc = NULL THEN EXIT FUNCTION
   DIM nResult AS LONG = GetDeviceCaps(hdc, HORZRES)
   DeleteDC hdc
   RETURN nResult
END FUNCTION
' ========================================================================================

' ========================================================================================
' Width, in pixels, of the printable area of the page.
' ========================================================================================
PRIVATE FUNCTION AfxGetPrinterVerticalResolution (BYREF wszPrinterName AS WSTRING) AS LONG
   DIM hdc AS HDC = CreateICW(NULL, wszPrinterName, NULL, NULL)
   IF hdc = NULL THEN RETURN 0
   DIM nResult AS LONG = GetDeviceCaps(hdc, VERTRES)
   DeleteDC hdc
   RETURN nResult
END FUNCTION
' ========================================================================================

' ========================================================================================
' The height of the physical page, in device units. For example, a printer set to print at
' 600 dpi on 8.5-by-11-inch paper has a physical height value of 6600 device units. Note
' that the physical page is almost always greater than the printable area of the page,
' and never smaller.
' ========================================================================================
PRIVATE FUNCTION AfxGetPrinterPhysicalHeight (BYREF wszPrinterName AS WSTRING) AS LONG
   DIM hdc AS HDC = CreateICW(NULL, wszPrinterName, NULL, NULL)
   IF hdc = NULL THEN RETURN 0
   DIM nResult AS LONG = GetDeviceCaps(hdc, PHYSICALHEIGHT)
   DeleteDC hdc
   RETURN nResult
END FUNCTION
' ========================================================================================

' ========================================================================================
' The distance from the left edge of the physical page to the left edge of the printable area,
' in device units. For example, a printer set to print at 600 dpi on 8.5-by-11-inch paper,
' that cannot print on the leftmost 0.25-inch of paper, has a horizontal physical offset
' of 150 device units.
' ========================================================================================
PRIVATE FUNCTION AfxGetPrinterPhysicalOffsetX (BYREF wszPrinterName AS WSTRING) AS LONG
   DIM hdc AS HDC = CreateICW(NULL, wszPrinterName, NULL, NULL)
   IF hdc = NULL THEN RETURN 0
   DIM nResult AS LONG = GetDeviceCaps(hdc, PHYSICALOFFSETX)
   DeleteDC hdc
   RETURN nResult
END FUNCTION
' ========================================================================================

' ========================================================================================
' The distance from the top edge of the physical page to the top edge of the printable
' area, in device units. For example, a printer set to print at 600 dpi on 8.5-by-11-inch
' paper, that cannot print on the topmost 0.5-inch of paper, has a vertical physical
' offset of 300 device units.
' ========================================================================================
PRIVATE FUNCTION AfxGetPrinterPhysicalOffsetY (BYREF wszPrinterName AS WSTRING) AS LONG
   DIM hdc AS HDC = CreateICW(NULL, wszPrinterName, NULL, NULL)
   IF hdc = NULL THEN RETURN 0
   DIM nResult AS LONG = GetDeviceCaps(hdc, PHYSICALOFFSETY)
   DeleteDC hdc
   RETURN nResult
END FUNCTION
' ========================================================================================

' ========================================================================================
' The width of the physical page, in device units. For example, a printer set to print at
' 600 dpi on 8.5-x11-inch paper has a physical width value of 5100 device units. Note that
' the physical page is almost always greater than the printable area of the page, and
' never smaller.
' ========================================================================================
PRIVATE FUNCTION AfxGetPrinterPhysicalWidth (BYREF wszPrinterName AS WSTRING) AS LONG
   DIM hdc AS HDC = CreateICW(NULL, wszPrinterName, NULL, NULL)
   IF hdc = NULL THEN RETURN 0
   DIM nResult AS LONG = GetDeviceCaps(hdc, PHYSICALWIDTH)
   DeleteDC hdc
   RETURN nResult
END FUNCTION
' ========================================================================================

' ========================================================================================
' Scaling factor for the x-axis of the printer.
' ========================================================================================
PRIVATE FUNCTION AfxGetPrinterScalingFactorX (BYREF wszPrinterName AS WSTRING) AS LONG
   DIM hdc AS HDC = CreateICW(NULL, wszPrinterName, NULL, NULL)
   IF hdc = NULL THEN EXIT FUNCTION
   DIM nResult AS LONG = GetDeviceCaps(hdc, SCALINGFACTORX)
   DeleteDC hdc
   RETURN nResult
END FUNCTION
' ========================================================================================

' ========================================================================================
' Scaling factor for the y-axis of the printer.
' ========================================================================================
PRIVATE FUNCTION AfxGetPrinterScalingFactorY (BYREF wszPrinterName AS WSTRING) AS LONG
   DIM hdc AS HDC = CreateICW(NULL, wszPrinterName, NULL, NULL)
   IF hdc = NULL THEN RETURN 0
   DIM nResult AS LONG = GetDeviceCaps(hdc, SCALINGFACTORY)
   DeleteDC hdc
   RETURN nResult
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets data for a specified printer.
' - wszPrinterName : The name of the printer.
' - dmField : Field to change
' - nValue  : New value
' ========================================================================================
PRIVATE FUNCTION AfxSetPrinterInfo (BYREF wszPrinterName AS WSTRING, BYVAL dmField AS DWORD, BYVAL nValue AS LONG) AS BOOLEAN

   ' // Start by opening the printer
   DIM hPrinter AS HANDLE
   DIM pd AS PRINTER_DEFAULTSW
   pd.DesiredAccess = PRINTER_ALL_ACCESS
   IF OpenPrinterW(wszPrinterName, @hPrinter, @pd) = FALSE THEN RETURN FALSE

   ' // The first GetPrinterW cll tells you how big the buffer should be in
   ' // order to hold all of PRINTER_INFO_2. Note that this should fail with
   ' // ERROR_INSUFFICIENT_BUFFER. If GetPrinterW fails for any other reason
   ' // or cbNeeded isn't set for some reason, then there is a problem...
   DIM dwLevel AS DWORD = 2, cbNeeded AS DWORD
   DIM nRet AS LONG = GetPrinterW(hPrinter, dwLevel, NULL, 0, @cbNeeded)
   IF nRet = 0 AND GetLastError <> ERROR_INSUFFICIENT_BUFFER THEN
      ClosePrinter(hPrinter)
      RETURN FALSE
   END IF

   ' // Allocate enough space for PRINTER_INFO_2...
   DIM bufferPrn AS STRING
   bufferPrn = SPACE(cbNeeded)
   ' // The second GetPrinterW fills in all the current settings, so all you
   ' // need to do is modify what you're interested in...
   nRet = GetPrinterW(hPrinter, dwLevel, CAST(BYTE PTR, STRPTR(bufferPrn)), cbNeeded, @cbNeeded)
   IF nRet = 0 THEN
      ClosePrinter(hPrinter)
      RETURN FALSE
   END IF

   ' // If GetPrinterW didn't fill in the DEVMODE, try to get it by calling DocumentProperties...
   DIM pi2 AS PRINTER_INFO_2W PTR
   pi2 = CAST(PRINTER_INFO_2W PTR, STRPTR(bufferPrn))
   DIM bufferDoc AS STRING
   DIM pDevMode AS DEVMODEW PTR
   IF pi2->pDevMode = NULL THEN
      ' // Allocate a buffer of the correct size
      cbNeeded = DocumentPropertiesW(NULL, hPrinter, wszPrinterName, NULL, NULL, 0)
      bufferDoc = SPACE(cbNeeded)
      ' // Retrieve the printer configuration data
      nRet = DocumentPropertiesW(NULL, hPrinter, wszPrinterName, _
         cast(DEVMODEW PTR, STRPTR(bufferDoc)), NULL, DM_OUT_BUFFER)
      IF nRet <> IDOK THEN
         ClosePrinter(hPrinter)
         RETURN FALSE
      END IF
      ' // Cast it to a DEVMODEW structure
      pDevMode = cast(DEVMODEW PTR, STRPTR(bufferDoc))
      pi2->pDevMode = pDevMode
   END IF

   ' // Specify exactly what we are attempting to change...
   SELECT CASE dmField
      CASE DM_COLLATE       : pi2->pDevMode->dmCollate = nValue
      CASE DM_COLOR         : pi2->pDevMode->dmColor = nValue
      CASE DM_COPIES        : pi2->pDevMode->dmCopies = nValue
      CASE DM_DUPLEX        : pi2->pDevMode->dmDuplex = nValue
      CASE DM_ORIENTATION   : pi2->pDevMode->dmOrientation = nValue
      CASE DM_PAPERSIZE     : pi2->pDevMode->dmPaperSize = nValue
      CASE DM_PRINTQUALITY  : pi2->pDevMode->dmPrintQuality = nValue
      CASE DM_DEFAULTSOURCE : pi2->pDevMode->dmDefaultSource = nValue
      CASE DM_SCALE         : pi2->pDevMode->dmScale = nValue
   END SELECT

   ' // Do not attempt to set security descriptor...
   pi2->pSecurityDescriptor = NULL

   ' // Make sure the driver-dependent part of devmode is updated...
   nRet = DocumentPropertiesW(NULL, hPrinter, wszPrinterName, _
             pi2->pDevMode, pi2->pDevMode, DM_IN_BUFFER OR DM_OUT_BUFFER)
   IF nRet <> IDOK THEN
      ClosePrinter(hPrinter)
      RETURN FALSE
   END IF

   ' // Update printer information...
   nRet = SetPrinterW(hPrinter, dwLevel, cast(BYTE PTR, pi2), 0)
   IF nRet = 0 THEN
      ' // The driver doesn't support, or it is unable to make the change...
      ClosePrinter(hPrinter)
      EXIT FUNCTION
   END IF

   ' // Finished with the printer
   ClosePrinter(hPrinter)

   IF nRet <> IDOK THEN RETURN FALSE
   RETURN TRUE

END FUNCTION
' ========================================================================================

' ========================================================================================
' Switches between color and monochrome on color printers.
' The following are the possible values:
'   DMCOLOR_COLOR
'   DMCOLOR_MONOCHROME
' Example: AfxSetPrinterColorMode(AfxGetDefaultPrinter, DMCOLOR_COLOR)
' ========================================================================================
PRIVATE FUNCTION AfxSetPrinterColorMode (BYREF wszPrinterName AS WSTRING, BYVAL nMode AS LONG) AS BOOLEAN
   RETURN AfxSetPrinterInfo(wszPrinterName, DM_COLOR, nMode)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Selects the number of copies printed if the device supports multiple-page copies.
' Example: AfxSetPrinterCopies(AfxGetDefaultPrinter, 2)
' ========================================================================================
PRIVATE FUNCTION AfxSetPrinterCopies (BYREF wszPrinterName AS WSTRING, BYVAL nCopies AS LONG) AS BOOLEAN
   RETURN AfxSetPrinterInfo(wszPrinterName, DM_COPIES, nCopies)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the printer duplex mode
' DMDUP_SIMPLEX = Single sided printing
' DMDUP_VERTICAL = Page flipped on the vertical edge
' DMDUP_HORIZONTAL = Page flipped on the horizontal edge
' Example: AfxSetPrinterDuplexMode(AfxGetDefaultPrinter, DMDUP_VERTICAL)
' ========================================================================================
PRIVATE FUNCTION AfxSetPrinterDuplexMode (BYREF wszPrinterName AS WSTRING, BYVAL nDuplexMode AS LONG) AS BOOLEAN
   RETURN AfxSetPrinterInfo(wszPrinterName, DM_DUPLEX, nDuplexMode)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the printer orientation.
' DMORIENT_PORTRAIT = Portrait
' DMORIENT_LANDSCAPE = Landscape
' Example: AfxSetPrinterOrientation(AfxGetDefaultPrinter, DMORIENT_LANDSCAPE)
' ========================================================================================
PRIVATE FUNCTION AfxSetPrinterOrientation (BYREF wszPrinterName AS WSTRING, BYVAL nOrientation AS LONG) AS BOOLEAN
   RETURN AfxSetPrinterInfo(wszPrinterName, DM_ORIENTATION, nOrientation)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the printer paper size.
' Example: AfxSetPrinterPaperSize(AfxGetDefaultPrinter, DMPAPER_A4)
' ========================================================================================
PRIVATE FUNCTION AfxSetPrinterPaperSize (BYREF wszPrinterName AS WSTRING, BYVAL nSize AS LONG) AS BOOLEAN
   RETURN AfxSetPrinterInfo(wszPrinterName, DM_PAPERSIZE, nSize)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Specifies the printer resolution.
' Example: AfxSetPrinterQuality(AfxGetDefaultPrinter, DMRES_HIGH)
' ========================================================================================
PRIVATE FUNCTION AfxSetPrinterQuality (BYREF wszPrinterName AS WSTRING, BYVAL nMode AS LONG) AS BOOLEAN
   RETURN AfxSetPrinterInfo(wszPrinterName, DM_PRINTQUALITY, nMode)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the paper source. Can be one of the following values, or it can be a device-specific
' value greater than or equal to DMBIN_USER.
' DMBIN_UPPER = 1; DMBIN_LOWER = 2; DMBIN_MIDDLE = 3; DMBIN_MANUAL = 4; DMBIN_ENVELOPE = 5;
' DMBIN_ENVMANUAL = 6; DMBIN_AUTO = 7;  DMBIN_TRACTOR = 8; DMBIN_SMALLFMT = 9;
' DMBIN_LARGEFMT = 10; DMBIN_LARGECAPACITY = 11; DMBIN_CASSETTE = 14; DMBIN_FORMSOURCE = 15
' Example: AfxSetPrinterTray(AfxGetDefaultPrinter, DMBIN_AUTO)
' ========================================================================================
PRIVATE FUNCTION AfxSetPrinterTray (BYREF wszPrinterName AS WSTRING, BYVAL nTray AS LONG) AS BOOLEAN
   RETURN AfxSetPrinterInfo(wszPrinterName, DM_DEFAULTSOURCE, nTray)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Specifies whether collation should be used when printing multiple copies.
' The following are the possible values:
'   DMCOLLATE_TRUE
'   DMCOLLATE_FALSE
' Example: AfxSetPrinterCollateStatus(AfxGetDefaultPrinter, DMCOLLATE_TRUE)
' ========================================================================================
PRIVATE FUNCTION AfxSetPrinterCollateStatus (BYREF wszPrinterName AS WSTRING, BYVAL nMode AS LONG) AS BOOLEAN
   RETURN AfxSetPrinterInfo(wszPrinterName, DM_COLLATE, nMode)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Specifies the factor by which the printed output is to be scaled. The apparent page size
' is scaled from the physical page size by a factor of dmScale /100. For example, a
' letter-sized page with a dmScale value of 50 would contain as much data as a page of
' 17- by 22-inches because the output text and graphics would be half their original
' height and width.
' Example: AfxSetPrinterScale(AfxGetDefaultPrinter, 100)
' ========================================================================================
PRIVATE FUNCTION AfxSetPrinterScale (BYREF wszPrinterName AS WSTRING, BYVAL nScale AS LONG) AS BOOLEAN
   RETURN AfxSetPrinterInfo(wszPrinterName, DM_SCALE, nScale)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the printer status.
' Returns 0 on failure or any reasonable combination of the following values:
' PRINTER_STATUS_PAUSED = &h00000001; PRINTER_STATUS_ERROR = &h00000002;
' PRINTER_STATUS_PENDING_DELETION = &h00000004; PRINTER_STATUS_PAPER_JAM = &h00000008;
' PRINTER_STATUS_PAPER_OUT = &h00000010; PRINTER_STATUS_MANUAL_FEED = &h00000020;
' PRINTER_STATUS_PAPER_PROBLEM = &h00000040; PRINTER_STATUS_OFFLINE = &h00000080;
' PRINTER_STATUS_IO_ACTIVE = &h00000100; PRINTER_STATUS_BUSY = &h00000200;
' PRINTER_STATUS_PRINTING = &h00000400; PRINTER_STATUS_OUTPUT_BIN_FULL = &h00000800;
' PRINTER_STATUS_NOT_AVAILABLE = &h00001000; PRINTER_STATUS_WAITING = &h00002000;
' PRINTER_STATUS_PROCESSING = &h00004000; PRINTER_STATUS_INITIALIZING = &h00008000;
' PRINTER_STATUS_WARMING_UP = &h00010000; PRINTER_STATUS_TONER_LOW = &h00020000;
' PRINTER_STATUS_NO_TONER = &h00040000; PRINTER_STATUS_PAGE_PUNT = &h00080000;
' PRINTER_STATUS_USER_INTERVENTION = &h00100000; PRINTER_STATUS_OUT_OF_MEMORY = &h00200000;
' PRINTER_STATUS_DOOR_OPEN = &h00400000; PRINTER_STATUS_SERVER_UNKNOWN = &h00800000;
' PRINTER_STATUS_POWER_SAVE = &h01000000
' Remarks: Not reliable. pi2->Status returns 0 when the printer is off line.
' ========================================================================================
'PRIVATE FUNCTION AfxGetPrinterStatus (BYREF wszPrinterName AS WSTRING) AS DWORD
'   ' // Open the printer
'   DIM hPrinter AS HANDLE
'   IF OpenPrinterW(wszPrinterName, @hPrinter, NULL) = FALSE THEN RETURN 0
'   ' // Get printer information
'   DIM dwLevel AS DWORD = 2, cbNeeded AS DWORD
'   DIM nRet AS LONG = GetPrinterW(hPrinter, dwLevel, NULL, 0, @cbNeeded)
'   IF nRet = 0 AND GetLastError <> ERROR_INSUFFICIENT_BUFFER THEN
'      ClosePrinter(hPrinter)
'      RETURN 0
'   END IF
'   ' // Allocate enough space for PRINTER_INFO_2...
'   DIM bufferPrn AS STRING
'   bufferPrn = SPACE(cbNeeded)
'   ' // The second GetPrinterW fills in all the current settings
'   nRet = GetPrinterW(hPrinter, dwLevel, CAST(BYTE PTR, STRPTR(bufferPrn)), cbNeeded, @cbNeeded)
'   IF nRet = 0 THEN
'      ClosePrinter(hPrinter)
'      RETURN 0
'   END IF
'   ' // Cast the buffer to a PRINTER_INFO_2W structure
'   DIM pi2 AS PRINTER_INFO_2W PTR = CAST(PRINTER_INFO_2W PTR, STRPTR(bufferPrn))
'   ' // Finished with the printer
'   ClosePrinter(hPrinter)
'   ' // Return the status value
'   RETURN pi2->Status
'END FUNCTION
' ========================================================================================
