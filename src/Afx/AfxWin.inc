' ########################################################################################
' Microsoft Windows
' File: AfxWin.inc
' Contents: Windows wrapper functions
' Compiler: Free Basic 32 & 64 bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#INCLUDE ONCE "windows.bi"
#INCLUDE ONCE "win/commctrl.bi"
#INCLUDE ONCE "win/ole2.bi"
#INCLUDE ONCE "win/psapi.bi"
#INCLUDE ONCE "win/shlobj.bi"
#INCLUDE ONCE "win/KnownFolders.bi"
#INCLUDE ONCE "win/commdlg.bi"
#INCLUDE ONCE "win/cderr.bi"

#INCLUDE ONCE "Afx/AfxPath.inc"
#INCLUDE ONCE "Afx/AfxStr.inc"

USING Afx

' ========================================================================================
' Helper procedures to display feedback and errors
' ========================================================================================
PRIVATE FUNCTION AfxMsg OVERLOAD (BYVAL pwszText AS WSTRING PTR, BYVAL uType AS DWORD = 0) AS LONG
   IF pwszText <> NULL THEN FUNCTION = MessageBoxW(GetActiveWindow, pwszText, "Message", MB_APPLMODAL OR uType) ELSE FUNCTION = MessageBoxW(GetActiveWindow, "", "", MB_APPLMODAL OR uType)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION AfxMsg OVERLOAD (BYREF cws AS CWSTR, BYVAL uType AS DWORD = 0) AS LONG
   FUNCTION = MessageBoxW(GetActiveWindow, *cws, "Message", MB_APPLMODAL OR uType)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Processes pending Windows messages.
' Call this procedure if you are performing a tight FOR/NEXT or DO/LOOP and need to allow
' your application to be responsive to user input.
' ========================================================================================
PRIVATE SUB AfxDoEvents (BYVAL hwnd AS HWND = NULL)
   DIM uMsg AS MSG
   IF hwnd = NULL THEN hwnd = GetActiveWindow
   WHILE PeekMessage(@uMsg, NULL, 0, 0, PM_REMOVE)
      ' // Determines whether a message is intended for the specified
      ' // dialog box and, if it is, processes the message.
      IF IsDialogMessageW(hwnd, @uMsg) = 0 THEN
         ' // Translates virtual-key messages into character messages.
         TranslateMessage @uMsg
         ' // Dispatches a message to a window procedure.
         DispatchMessageW @uMsg
      END IF
   WEND
END SUB
' ========================================================================================

' ========================================================================================
' Processes pending Windows messages.
' Call this procedure if you are performing a tight FOR/NEXT or DO/LOOP and need to allow
' your application to be responsive to user input.
' ========================================================================================
PRIVATE SUB AfxPumpMessages
   DIM uMsg AS MSG
   WHILE PeekMessage(@uMsg, NULL, 0, 0, PM_REMOVE)
      ' // Translates virtual-key messages into character messages.
      TranslateMessage @uMsg
      ' // Dispatches a message to a window procedure.
      DispatchMessageW @uMsg
   WEND
END SUB
' ========================================================================================

' ========================================================================================
' Returns the Windows version
' Platform 1:
'   400 Windows 95
'   410 Windows 98
'   490 Windows ME
' Platform 2:
'   400 Windows NT
'   500 Windows 2000
'   501 Windows XP
'   502 Windows Server 2003
'   600 Windows Vista and Windows Server 2008
'   601 Windows 7
'   602 Windows 8
'   603 Windows 8.1
' Note: As Windows 95 and Windows NT return the same version number, we also need to call
' GetWindowsPlatform to differentiate them.
' ========================================================================================
PRIVATE FUNCTION AfxWindowsVersion () AS LONG
   DIM dwVersion AS DWORD
   DIM AS LONG nMajorVer, nMinorVer
   dwVersion = GetVersion
   nMajorVer = LOBYTE(LOWORD(dwVersion))
   nMinorVer = HIBYTE(LOWORD(dwVersion))
   FUNCTION = (nMajorVer + nMinorVer / 100) * 100
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the Windows platform
'   1 Windows 95/98/ME
'   2 Windows NT/2000/XP/Server/Vista/Windows 7
' ========================================================================================
PRIVATE FUNCTION AfxWindowsPlatform () AS LONG
   DIM dwVersion AS DWORD
   dwVersion = GetVersion
   FUNCTION = IIF&(dwVersion < &H80000000, 2, 1)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the Windows build
' ========================================================================================
PRIVATE FUNCTION AfxWindowsBuild () AS LONG
   DIM dwVersion AS DWORD
   dwVersion = GetVersion
   IF dwVersion < &H80000000 THEN FUNCTION = HIWORD(dwVersion)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns TRUE if the Windows Platform is NT; FALSE, otherwise.
' ========================================================================================
PRIVATE FUNCTION AfxIsPlatformNT () AS BOOLEAN
   DIM osvi AS OSVERSIONINFOW
   osvi.dwOSVersionInfoSize = SIZEOF(osvi)
   GetVersionExW(@osvi)
   FUNCTION = (VER_PLATFORM_WIN32_NT = osvi.dwPlatformId)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the version of specified file multiplied by 100, e.g. 601 for version 6.01.
' Example: DIM ver AS LONG = AfxGetFileVersion("COMCTL32.DLL")
' ========================================================================================
PRIVATE FUNCTION AfxGetFileVersion (BYVAL pwszFileName AS WSTRING PTR) AS LONG
   DIM pvsffi AS VS_FIXEDFILEINFO PTR, dwHandle AS DWORD
   DIM cbLen AS DWORD = GetFileVersionInfoSizeW(pwszFileName, @dwHandle)
   IF cbLen = 0 THEN EXIT FUNCTION
   DIM pVerInfo AS HANDLE = HeapAlloc(GetProcessHeap, HEAP_ZERO_MEMORY, cbLen)
   IF pVerInfo = NULL THEN EXIT FUNCTION
   IF GetFileVersionInfoW(pwszFileName, dwHandle, cbLen, pVerInfo) THEN
      IF VerQueryValueW(pVerInfo, "\", @pvsffi, @cbLen) THEN
         DIM wMajor AS WORD = HIWORD(pvsffi->dwFileVersionMS)
         DIM wMinor AS WORD = LOWORD(pvsffi->dwFileVersionMS)
         FUNCTION = (wMajor + wMinor / 100) * 100
      END IF
   END IF
   HeapFree(GetProcessHeap, 0, pVerInfo)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the version of CommCtl32.dll multiplied by 100, e.g. 582 for version 5.82.
' ========================================================================================
PRIVATE FUNCTION AfxComCtlVersion () AS LONG
   FUNCTION = AfxGetFileVersion("COMCTL32.DLL")
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the path of the Windows directory.
' ========================================================================================
PRIVATE FUNCTION AfxGetWinDir () AS CWSTR
   DIM wszWinDir AS WSTRING * MAX_PATH
   GetWindowsDirectoryW wszWinDir, SIZEOF(wszWinDir)
   RETURN wszWinDir
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the name of the user associated with the current thread.
' ========================================================================================
PRIVATE FUNCTION AfxGetUserName () AS CWSTR
   DIM buffer AS WSTRING * 260
   DIM dwBufLen AS DWORD = SIZEOF(buffer)
   GetUserNameW(buffer, @dwBufLen)
   RETURN buffer
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the computer name of the current system.
' ========================================================================================
PRIVATE FUNCTION AfxGetComputerName () AS CWSTR
   DIM buffer AS WSTRING * MAX_COMPUTERNAME_LENGTH
   DIM dwBufLen AS DWORD = SIZEOF(buffer)
   GetComputerNameW(buffer, @dwBufLen)
   RETURN buffer
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the path of the program which is currently executing.
' ========================================================================================
PRIVATE FUNCTION AfxGetExePathName () AS CWSTR
   DIM buffer AS WSTRING * MAX_PATH, p AS LONG
   GetModuleFileNameW NULL, buffer, SIZEOF(buffer)
   p = INSTRREV(buffer, ANY ":/\")
   RETURN LEFT(buffer, p)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the file name of the program which is currently executing.
' ========================================================================================
PRIVATE FUNCTION AfxGetExeFileName () AS CWSTR
   DIM buffer AS WSTRING * MAX_PATH, p AS LONG
   GetModuleFileNameW NULL, buffer, SIZEOF(buffer)
   p = INSTRREV(buffer, ANY ":/\")
   IF p THEN buffer = MID(buffer, p + 1)
   p = INSTRREV(buffer, ".")
   IF p THEN buffer = LEFT(buffer, p - 1)
   RETURN buffer
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the file name and extension of the program which is currently executing.
' ========================================================================================
PRIVATE FUNCTION AfxGetExeFileNameX () AS CWSTR
   DIM buffer AS WSTRING * MAX_PATH, p AS LONG
   GetModuleFileNameW NULL, buffer, SIZEOF(buffer)
   p = INSTRREV(buffer, ANY ":/\")
   IF p THEN buffer = MID(buffer, p + 1)
   RETURN buffer
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the extension (with a leading period) of the program which is currently executing.
' ========================================================================================
PRIVATE FUNCTION AfxGetExeFileExt () AS CWSTR
   DIM buffer AS WSTRING * MAX_PATH, p AS LONG
   GetModuleFileNameW NULL, buffer, SIZEOF(buffer)
   p = INSTRREV(buffer, ".")
   IF p THEN RETURN MID(buffer, p) ELSE RETURN buffer
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the full path of the program which is currently executing.
' ========================================================================================
PRIVATE FUNCTION AfxGetExeFullPath () AS CWSTR
   DIM buffer AS WSTRING * MAX_PATH, cb AS LONG
   cb = GetModuleFileNameW(NULL, buffer, SIZEOF(buffer))
   RETURN LEFT(buffer, cb)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Parses a path/filename and returns the file name portion. That is the text to the right
' of the last backslash (\) or colon (:), ending just before the last period (.). The file
' extension is excluded.
' ========================================================================================
PRIVATE FUNCTION AfxGetFileName (BYREF wszPath AS WSTRING) AS CWSTR
   DIM cwsPath AS CWSTR = wszPath
   DIM p AS LONG
   p = INSTRREV(cwsPath, ANY ":/\")
   IF p THEN cwsPath = MID(cwsPath, p + 1)
   p = INSTRREV(cwsPath, ".")
   IF p THEN
      cwsPath = LEFT(**cwsPath, p - 1)
   ELSE
      cwsPath = ""
   END IF
   RETURN cwsPath
END FUNCTION
' ========================================================================================

' ========================================================================================
' Parses a path/filename and returns the file name portion. That is the text to the right
' of the last backslash (\) or colon (:), ending just before the last period (.).
' ========================================================================================
PRIVATE FUNCTION AfxGetFileNameX (BYREF wszPath AS WSTRING) AS CWSTR
   DIM cwsPath AS CWSTR = wszPath
   DIM p AS LONG
   p = INSTRREV(cwsPath, ANY ":/\")
   IF p THEN RETURN MID(cwsPath, p + 1) ELSE RETURN cwsPath
END FUNCTION
' ========================================================================================

' ========================================================================================
' Parses a path/filename and returns the extension portion of the path/file name. That is
' the last period (.) in the string plus the text to the right of it.
' ========================================================================================
PRIVATE FUNCTION AfxGetFileExt (BYREF wszPath AS WSTRING) AS CWSTR
   DIM cwsPath AS CWSTR = wszPath
   DIM p AS LONG
   p = INSTRREV(cwsPath, ".")
   IF p THEN RETURN MID(cwsPath, p) ELSE RETURN cwsPath
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the current directory for the current process.
' ========================================================================================
PRIVATE FUNCTION AfxGetCurDir () AS CWSTR
   DIM wszCurDir AS WSTRING * MAX_PATH
   GetCurrentDirectoryW(MAX_PATH, wszCurDir)
   RETURN wszCurDir
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts the specified path to its long form.
' ========================================================================================
PRIVATE FUNCTION AfxGetLongPathName (BYREF wszPath AS WSTRING) AS CWSTR
   DIM wszLongPath AS WSTRING * MAX_PATH
   GetLongPathNameW(wszPath, wszLongPath, SIZEOF(wszLongPath))
   RETURN wszLongPath
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the short path form of the specified path.
' ========================================================================================
PRIVATE FUNCTION AfxGetShortPathName (BYREF wszPath AS WSTRING) AS CWSTR
   DIM wszShortPath AS WSTRING * MAX_PATH
   GetShortPathNameW(wszPath, wszShortPath, SIZEOF(wszShortPath))
   RETURN wszShortPath
END FUNCTION
' ========================================================================================

' ========================================================================================
' Parses a path/filename and returns the path portion. That is the text up to and including
' the last backslash (\) or colon (:).
' ========================================================================================
PRIVATE FUNCTION AfxGetPathName (BYREF wszPath AS WSTRING) AS CWSTR
   DIM cwsPath AS CWSTR = wszPath
   DIM p AS LONG
   p = INSTRREV(cwsPath, ANY ":\/")
   IF p THEN RETURN LEFT(**cwsPath, p)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the path of the specified system DLL.
' ========================================================================================
PRIVATE FUNCTION AfxGetSystemDllPath (BYREF wszDllName AS WSTRING) AS CWSTR
   DIM wszPath AS WSTRING * MAX_PATH
   DIM hLib AS HMODULE = LoadLibraryW(wszDllName)
   IF hLib THEN
      GetModuleFileNameW(hLib, wszPath, SIZEOF(wszPath))
      FreeLibrary hLib
   END IF
   RETURN wszPath
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the path of the executable file given its process identifier.
' ========================================================================================
PRIVATE FUNCTION AfxGetPathFromProcessId (BYVAL dwProcessId AS DWORD) AS CWSTR
   DIM hProcess AS HANDLE, wszPath AS WSTRING * MAX_PATH
   hProcess = OpenProcess(PROCESS_QUERY_INFORMATION OR PROCESS_VM_READ, FALSE, dwProcessId)
   IF hProcess <> NULL THEN
      GetModuleFileNameExW(hProcess, NULL, wszPath, SIZEOF(wszPath))
      CloseHandle(hProcess)
   END IF
   RETURN wszPath
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the path of the executable file that created the specified window.
' ========================================================================================
PRIVATE FUNCTION AfxGetPathFromWindowHandle (BYVAL hwnd AS HWND) AS CWSTR
   DIM idProc AS DWORD, hProcess AS HANDLE, wszPath AS WSTRING * MAX_PATH
   GetWindowThreadProcessId(hwnd, @idProc)
   IF idProc THEN
      hProcess = OpenProcess(PROCESS_QUERY_INFORMATION OR PROCESS_VM_READ, FALSE, idProc)
      IF hProcess THEN
         GetModuleFileNameExW(hProcess, NULL, wszPath, SIZEOF(wszPath))
         CloseHandle(hProcess)
         RETURN wszPath
      END IF
   END IF
END FUNCTION
' ========================================================================================

#if _WIN32_WINNT = &h0602
' ========================================================================================
' Retrieves the path of an special folder. Requires Windows Vista/Windows 7 or superior.
' - rfid: A reference to the KNOWNFOLDERID that identifies the folder. The folders associated
'   with the known folder IDs might not exist on a particular system.
' - dwFlags: Flags that specify special retrieval options. This value can be 0; otherwise,
'   it is one or more of the KNOWN_FOLDER_FLAG values.
' - hToken: An access token used to represent a particular user. This parameter is usually
'   set to NULL, in which case the function tries to access the current user's instance of
'   the folder. However, you may need to assign a value to hToken for those folders that can
'   have multiple users but are treated as belonging to a single user. The most commonly used
'   folder of this type is Documents.
'   The calling application is responsible for correct impersonation when hToken is non-null.
'   It must have appropriate security privileges for the particular user, including TOKEN_QUERY
'   and TOKEN_IMPERSONATE, and the user's registry hive must be currently mounted. See Access
'   Control for further discussion of access control issues.
'   https://msdn.microsoft.com/en-us/library/windows/desktop/aa374860(v=vs.85).aspx
'   Assigning the hToken parameter a value of -1 indicates the Default User. This allows
'   clients of SHGetKnownFolderIDList to find folder locations (such as the Desktop folder)
'   for the Default User. The Default User user profile is duplicated when any new user
'   account is created, and includes special folders such as Documents and Desktop.
'   Any items added to the Default User folder also appear in any new user account.
'   Note that access to the Default User folders requires administrator privileges.
' Return value:
'   The path of the requested folder on success, or an empty string on failure.
' Remarks: For a list of KNOWNFOLDERID constants see:
'   https://msdn.microsoft.com/en-us/library/windows/desktop/dd378457(v=vs.85).aspx
' Usage example: AfxGetKnowFolderPath(@FOLDERID_CommonPrograms)
' ========================================================================================
PRIVATE FUNCTION AfxGetKnowFolderPath (BYVAL rfid AS CONST KNOWNFOLDERID CONST PTR, BYVAL dwFlags AS DWORD = 0, BYVAL hToken AS HANDLE = NULL) AS CWSTR
   DIM pidl AS ITEMIDLIST PTR          ' // Pointer to an item identifier list (PIDL)
   DIM wszPath AS WSTRING * MAX_PATH   ' // Folder's path
   IF SHGetKnownFolderIDList(rfid, dwFlags, hToken, @pidl) = S_OK THEN
      SHGetPathFromIDListW pidl, @wszPath
      CoTaskMemFree pidl
      RETURN wszPath
   END IF
END FUNCTION
' ========================================================================================
#endif

' ========================================================================================
' Retrieves the path of an special folder.
' - nFolder: A CSIDL value that identifies the folder of interest.
' For a list of CSIDL values see:
' https://msdn.microsoft.com/en-us/library/windows/desktop/bb762494(v=vs.85).aspx
' ========================================================================================
PRIVATE FUNCTION AfxGetSpecialFolderLocation (BYVAL nFolder AS LONG) AS CWSTR
   DIM pidl AS ITEMIDLIST PTR          ' // Pointer to an item identifier list (PIDL)
   DIM wszPath AS WSTRING * MAX_PATH   ' // Folder's path
   IF SHGetSpecialFolderLocation(0, nFolder, @pidl) = S_OK THEN
      SHGetPathFromIDListW pidl, @wszPath
      CoTaskMemFree pidl
      RETURN wszPath
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' // Control_RunDLL is an undocumented procedure in the Shell32.dll which can be used
' // to launch control panel applications. You’ve to pass the name of the control panel
' // file (.cpl) and the tool represented by it will be launched. For launching some
' // control panel applications, you’ve to provide a valid windows handle (hwnd parameter)
' // and program instance (hinstance parameter).
' // This opens the control panel: AfxControlRunDLL(0, 0, "", SW_SHOWNORMAL)
' // This opens the applications wizard: AfxControlRunDLL(0, 0, "appwiz.cpl", SW_SHOWNORMAL)
' Return value: TRUE or FALSE.
' ========================================================================================
PRIVATE FUNCTION AfxControlRunDLL (BYVAL hwnd AS HWND, BYVAL hInst AS HINSTANCE, BYVAL cmd AS WSTRING PTR, BYVAL nCmdShow AS LONG) AS BOOLEAN
   DIM AS ANY PTR pLib = DyLibLoad("shell32.dll")
   IF pLib = 0 THEN EXIT FUNCTION
   DIM pProc AS FUNCTION (BYVAL hwnd AS HWND, BYVAL hInst AS HINSTANCE, BYVAL cmd AS WSTRING PTR, BYVAL nCmdShow AS LONG) AS BOOLEAN
   pProc = DyLibSymbol(pLib, "Control_RunDLLW")
   IF pProc = 0 THEN EXIT FUNCTION
   FUNCTION = pProc(hwnd, hInst, cmd, nCmdShow)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Saves the contents of a string buffer in a temporary file and returns the name of the file.
' ========================================================================================
PRIVATE FUNCTION AfxSaveTempFile (BYVAL pwszBuffer AS WSTRING PTR, BYREF wszExtension AS WSTRING) AS CWSTR
   DIM wszTmpFileName AS WSTRING * MAX_PATH
   DIM wszTmpPath AS WSTRING * MAX_PATH - 14
   DIM dwRes AS DWORD, fn AS LONG
   dwRes = GetTempPathW(MAX_PATH - 14, @wszTmpPath)
   IF dwRes > 0 AND dwRes <= MAX_PATH - 14 THEN
      dwRes = GetTempFileNameW(@wszTmpPath, "TMP", 0, @wszTmpFileName)
      IF dwRes THEN
         IF LEN(wszExtension) THEN wszTmpFileName = LEFT(wszTmpFileName, LEN(wszTmpFileName) -  3) & wszExtension
         fn = FREEFILE
         OPEN wszTmpFileName FOR OUTPUT AS #fn
         IF ERR = 0 THEN
            PRINT #fn, *pwszBuffer
            CLOSE #fn
            RETURN wszTmpFileName
         END IF
      END IF
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the Internet Explorer version (major.minor).
' ========================================================================================
PRIVATE FUNCTION AfxGetInternetExplorerVersion () AS SINGLE
   DIM hKey AS HKEY, wszData AS WSTRING * 1024, cbData AS DWORD = SIZEOF(wszData)
   IF RegOpenKeyExW(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Internet Explorer", 0, _
                    KEY_QUERY_VALUE, @hKey) <> ERROR_SUCCESS THEN EXIT FUNCTION
   RegQueryValueExW hKey, "Version", NULL, NULL, cast(BYTE PTR, @wszData), @cbData
   RegCloseKey hKey
   FUNCTION = VAL(wszData)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the path of the default browser.
' ========================================================================================
PRIVATE FUNCTION AfxGetDefaultBrowserPath () AS CWSTR
   DIM wszPath AS WSTRING * MAX_PATH, cchOut AS DWORD = SIZEOF(wszPath)
   IF SUCCEEDED(AssocQueryStringW(0, ASSOCSTR_EXECUTABLE, "http", "open", wszPath, @cchOut)) THEN RETURN wszPath
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the name of the default browser.
' ========================================================================================
PRIVATE FUNCTION AfxGetDefaultBrowserName () AS CWSTR
   DIM wszPath AS WSTRING * MAX_PATH, cchOut AS DWORD = SIZEOF(wszPath)
   IF SUCCEEDED(AssocQueryStringW(0, ASSOCSTR_EXECUTABLE, "http", "open", wszPath, @cchOut)) THEN
      RETURN AfxGetFileName(wszPath)
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the path of the default mail client application.
' ========================================================================================
PRIVATE FUNCTION AfxGetDefaultMailClientPath () AS CWSTR
   DIM wszPath AS WSTRING * MAX_PATH, cchOut AS DWORD = SIZEOF(wszPath)
   IF SUCCEEDED(AssocQueryStringW(0, ASSOCSTR_EXECUTABLE, "mailto", "open", wszPath, @cchOut)) THEN RETURN wszPath
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the name of the default mail client application.
' ========================================================================================
PRIVATE FUNCTION AfxGetDefaultMailClientName () AS CWSTR
   DIM wszPath AS WSTRING * MAX_PATH, cchOut AS DWORD = SIZEOF(wszPath)
   IF SUCCEEDED(AssocQueryStringW(0, ASSOCSTR_EXECUTABLE, "mailto", "open", wszPath, @cchOut)) THEN
      RETURN AfxGetFileName(wszPath)
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Displays a dialog box that enables the user to select a folder.
' - pwszTitle       = A string value that represents the title displayed inside the Browse dialog box.
' - pwszStartFolder = The initial folder that the dialog will show.
' - nFlags          = Optional. A LONG value that contains the options for the method. This can be a
'                     combination of the values listed under the ulFlags member of the BROWSEINFO structure.
'                     See: http://msdn.microsoft.com/en-us/library/windows/desktop/bb773205%28v=vs.85%29.aspx
'                     Default value = BIF_RETURNONLYFSDIRS OR BIF_DONTGOBELOWDOMAIN OR BIF_USENEWUI OR BIF_RETURNFSANCESTORS
' Note: To display the old style dialog, pass -1 in the dwFlags parameter.
' ========================================================================================
' // Browse for folder dialog procedure
PRIVATE FUNCTION AfxBrowseForFolderProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

   DIM wszBuffer AS WSTRING * MAX_PATH

   IF uMsg = BFFM_INITIALIZED THEN
      SendMessageW hwnd, BFFM_SETSELECTIONW, CTRUE, cast(LPARAM, lParam)
   ELSEIF uMsg = BFFM_SELCHANGED THEN
      SHGetPathFromIDListW(cast(ITEMIDLIST PTR, wParam), @wszBuffer)
      IF wParam = 0 OR _ ' // No id number
         LEN(wszBuffer) = 0 OR _ ' // No name
         (GetFileAttributesW(wszBuffer) AND FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY OR _ ' // Not a real directory
         MID(wszBuffer, 2, 1) <> ":" THEN ' // Not a local or mapped drive
            SendMessageW hwnd, BFFM_ENABLEOK, FALSE, FALSE
      ELSEIF ((GetFileAttributesW(wszBuffer) AND FILE_ATTRIBUTE_SYSTEM) = FILE_ATTRIBUTE_SYSTEM) AND _
            RIGHT(wszBuffer, 2) <> ":\" THEN   ' // Exclude system folders, allow root directories
         SendMessageW hwnd, BFFM_ENABLEOK, FALSE, FALSE
      END IF
   END IF
   RETURN 0

END FUNCTION

PRIVATE FUNCTION AfxBrowseForFolder (BYVAL hwnd AS HWND, BYVAL pwszTitle AS WSTRING PTR = NULL, _
   BYVAL pwszStartFolder AS WSTRING PTR = NULL, BYVAL nFlags AS LONG = 0) AS CWSTR

   DIM wszBuffer AS WSTRING * MAX_PATH, bi AS BROWSEINFOW, pidl AS ITEMIDLIST PTR
   IF nFlags = 0 THEN nFlags = BIF_RETURNONLYFSDIRS OR BIF_DONTGOBELOWDOMAIN OR BIF_USENEWUI OR BIF_RETURNFSANCESTORS
   IF nFlags = -1 THEN nFlags = BIF_RETURNONLYFSDIRS OR BIF_DONTGOBELOWDOMAIN OR BIF_RETURNFSANCESTORS

   bi.hWndOwner = hwnd
   bi.lpszTitle = pwszTitle
   bi.ulFlags   = nFlags
   bi.lpfn      = cast(BFFCALLBACK, @AfxBrowseForFolderProc)
   bi.lParam    = cast(LPARAM, pwszStartFolder)
   pidl         = SHBrowseForFolderW(@bi)

   IF pidl THEN
      SHGetPathFromIDListW(pidl, @wszBuffer)
      CoTaskMemFree pidl
   END IF
   RETURN wszBuffer

END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates an Open dialog box that lets the user specify the drive, directory, and the name
' of a file or set of files to be opened.
' - hwndOwner: A handle to the window that owns the dialog box. This parameter can be any
'   valid window handle, or it can be NULL if the dialog box has no owner.
' - wszTitle: A string to be placed in the title bar of the dialog box. If this member is NULL,
'   the system uses the default title (that is, Save As or Open).
' - wszFile: The file name used to initialize the File Name edit control. When the GetOpenFileName
'   or GetSaveFileName function returns successfully, this buffer contains the drive designator,
'   path, file name, and extension of the selected file.
'   If the OFN_ALLOWMULTISELECT flag is set and the user selects multiple files, the buffer
'   contains the current directory followed by the file names of the selected files. For
'   Explorer-style dialog boxes, the directory and file name strings are NULL separated,
'   with an extra NULL character after the last file name. For old-style dialog boxes, the
'   strings are space separated and the function uses short file names for file names with
'   spaces. You can use the FindFirstFile function to convert between long and short file
'   names. If the user selects only one file, the lpstrFile string does not have a separator
'   between the path and file name.
' - wszInitialDir: The initial directory.
' - wszFilter: A buffer containing pairs of "|" separated filter strings. The first string
'   in each pair is a display string that describes the filter (for example, "Text Files"),
'   and the second string specifies the filter pattern (for example, "*.TXT"). To specify
'   multiple filter patterns for a single display string, use a semicolon to separate the
'   patterns (for example, "*.TXT;*.DOC;*.BAK"). A pattern string can be a combination of
'   valid file name characters and the asterisk (*) wildcard character. Do not include spaces
'   in the pattern string.
'   The system does not change the order of the filters. It displays them in the File Types
'   combo box in the order specified in wszFilter. If wszFilter is NULL, the dialog box
'   does not display any filters.
' - wszDefExt: The default extension. GetOpenFileName and GetSaveFileName append this
'   extension to the file name if the user fails to type an extension. This string can be
'   any length, but only the first three characters are appended. The string should not
'   contain a period (.). If this member is NULL and the user fails to type an extension,
'   no extension is appended.
' - pdwFlags: A set of bit flags you can use to initialize the dialog box. When the dialog
'   box returns, it sets these flags to indicate the user's input. For example, to check
'   if the user has checked the read only checkbox:
'   IF (pdwFlags AND OFN_READONLY) = OFN_READONLY THEN ...
'   This value can be a combination of the following flags:
'   See complete list and explanations at:
'   https://msdn.microsoft.com/en-us/library/windows/desktop/ms646839(v=vs.85).aspx
' - pdwBufLen: The size of the buffer, in charactersm where the names of the selected
'   files will be returned.
' Return value:
'   An string containing a comma separated list of the selected files.
'   Parse the number of ",". If only one, then the user has selected only a file and the
'   string contains the full path. If more, The first substring contains the path and the
'   others the files.
'   If the user has not selected any file, an empty string is returned.
'   On failure, an empty string is returned and, if not null, the pdwBufLen parameter will
'   be filled by the size of the required buffer in characters.
' Usage example:
'   DIM wszFile AS WSTRING * 260 = "*.*"
'   DIM wszInitialDir AS STRING * 260 = CURDIR
'   DIM wszFilter AS WSTRING * 260 = "BAS files (*.BAS)|*.BAS|" & "All Files (*.*)|*.*|"
'   DIM dwFlags AS DWORD = OFN_EXPLORER OR OFN_FILEMUSTEXIST OR OFN_HIDEREADONLY OR OFN_ALLOWMULTISELECT
'   DIM cws AS CWSTR = AfxOpenFileDialog(hwnd, "", wszFile, wszInitialDir, wszFilter, "BAS", @dwFlags, NULL)
'   AfxMsg cws
' ========================================================================================
PRIVATE FUNCTION AfxOpenFileDialog ( _
   BYVAL hwndOwner AS HWND _                    ' // Parent window
 , BYREF wszTitle AS WSTRING _                  ' // Caption
 , BYREF wszFile AS WSTRING _                   ' // Filename
 , BYREF wszInitialDir AS WSTRING _             ' // Start directory
 , BYREF wszFilter AS WSTRING _                 ' // Filename filter
 , BYREF wszDefExt AS WSTRING _                 ' // Default extension
 , BYVAL pdwFlags AS DWORD PTR = NULL _         ' // Flags
 , BYVAL pdwBufLen AS DWORD PTR = NULL _        ' // Buffer length
 ) AS CWSTR

   DIM dwFlags AS DWORD, dwBufLen AS DWORD
   IF pdwFlags THEN dwFlags = *pdwFlags
   IF pdwBufLen THEN dwBufLen = *pdwBuflen

   ' // Filter is a sequence of WSTRINGs with a final (extra) double null terminator
   ' // The "|" characters are replaced with nulls
   DIM wszMarkers AS WSTRING * 4 = "||"
   IF RIGHT(wszFilter, 1) <> "|" THEN wszMarkers += "|"
   DIM cwsFilter AS CWSTR = wszFilter & wszMarkers
   DIM dwFilterStrSize AS DWORD = LEN(cwsFilter)
   ' // Replace markers("|") with nulls
   DIM pchar AS WCHAR PTR = *cwsFilter
   FOR i AS LONG = 0 TO LEN(cwsFilter) - 1
      IF pchar[i] = ASC("|") THEN pchar[i] = 0
   NEXT

   ' // If the initial directory has not been specified, assume the current directory
   IF LEN(wszInitialDir) = 0 THEN wszInitialDir = CURDIR
   ' // The size of the buffer must be at least MAX_PATH characters
   IF dwBufLen = 0 THEN
      IF (dwFlags AND OFN_ALLOWMULTISELECT = OFN_ALLOWMULTISELECT) THEN dwBufLen = 32768  ' // 64 Kb buffer
   END IF
   IF dwBufLen < 260 THEN dwBufLen = 260   ' // Make room for at least one path
   ' // Allocate the file name and a marker ("|") to be replaced with a null
   DIM cwsFile AS CWSTR = wszFile & "|"
   ' // Store the position of the marker
   DIM cbPos AS LONG = LEN(cwsFile) - 1
   ' // Allocate room for the buffer
   IF LEN(cwsFile) < dwBufLen THEN cwsFile += SPACE(dwBufLen - LEN(cwsFile))
   DIM dwFileStrSize AS DWORD = LEN(cwsFile)
   ' // The filename must be null terminated (replace the marker with a null)
   pchar = *cwsFile
   pchar[cbPos] = 0

   ' // Fill the members of the structure
   DIM ofn AS OPENFILENAMEW
   ofn.lStructSize     = SIZEOF(ofn)
   IF AfxWindowsVersion < 5 THEN ofn.lStructSize = 76
   ofn.hwndOwner       = hwndOwner
   ofn.lpstrFilter     = *cwsFilter
   ofn.nFilterIndex    = 1
   ofn.lpstrFile       = *cwsFile
   ofn.nMaxFile        = dwFileStrSize
   ofn.lpstrInitialDir = @wszInitialDir
   IF LEN(wszTitle) THEN ofn.lpstrTitle = @wszTitle
   ofn.Flags = dwFlags OR OFN_EXPLORER
   IF LEN(wszDefExt) THEN ofn.lpstrDefExt = @wszDefExt

   ' // Call the open file dialog
   IF GetOpenFilenameW(@ofn) THEN
      pchar = *cwsFile
      FOR i AS LONG = 0 TO dwFileStrSize - 1
         ' // If double null, exit
         IF pchar[i] = 0 AND pchar[i + 1] = 0 THEN EXIT FOR
         ' // Replace null with ","
         IF pchar[i] = 0 THEN pchar[i] = ASC(",")
      NEXT
      ' // Trim trailing spaces
      cwsFile = RTRIM(cwsFile, CHR(32))
      IF RIGHT(**cwsFile, 1) = "," THEN cwsFile = LEFT(**cwsFile, LEN(cwsFile) - 1)
   ELSE
      ' // Buffer too small
      IF CommDlgExtendedError = FNERR_BUFFERTOOSMALL THEN
         dwBufLen = ASC(**cwsFile)
      END IF
      cwsFile = ""
   END IF

   ' // Return the retrieved values
   IF pdwFlags THEN *pdwFlags = ofn.Flags
   IF pdwBufLen THEN *pdwBufLen = dwBufLen
   RETURN cwsFile

END FUNCTION
' ========================================================================================

' ========================================================================================
' The parameters are the same that for AfxOpenFileDialog, except the optional pdwBufferLen.
' In the pdwFlags parameter you may add OFN_OVERWRITEPROMPT to be asked if you want to
' overwrite an existing file.
' Usage example:
'   DIM wszFile AS WSTRING * 260 = "*.*"
'   DIM wszInitialDir AS STRING * 260 = CURDIR
'   DIM wszFilter AS WSTRING * 260 = "BAS files (*.BAS)|*.BAS|" & "All Files (*.*)|*.*|"
'   DIM dwFlags AS DWORD = OFN_EXPLORER OR OFN_FILEMUSTEXIST OR OFN_HIDEREADONLY OR OFN_OVERWRITEPROMPT
'   DIM cws AS CWSTR = AfxSaveFileDialog(hwnd, "", wszFile, wszInitialDir, wszFilter, "BAS", @dwFlags)
'   AfxMsg cws
' ========================================================================================
PRIVATE FUNCTION AfxSaveFileDialog ( _
   BYVAL hwndOwner AS HWND _                    ' // Parent window
 , BYREF wszTitle AS WSTRING _                  ' // Caption
 , BYREF wszFileName AS WSTRING _               ' // Filename
 , BYREF wszInitialDir AS WSTRING _             ' // Start directory
 , BYREF wszFilter AS WSTRING _                 ' // Filename filter
 , BYREF wszDefExt AS WSTRING _                 ' // Default extension
 , BYVAL pdwFlags AS DWORD PTR = NULL _         ' // Flags
 ) AS CWSTR

   DIM dwFlags AS DWORD
   IF pdwFlags THEN dwFlags = *pdwFlags

   ' // Filter is a sequence of WSTRINGs with a final (extra) double null terminator
   ' // The "|" characters are replaced with nulls
   DIM wszMarkers AS WSTRING * 4 = "||"
   IF RIGHT(wszFilter, 1) <> "|" THEN wszMarkers += "|"
   DIM cwsFilter AS CWSTR = wszFilter & wszMarkers
   DIM dwFilterStrSize AS DWORD = LEN(cwsFilter)
   ' // Replace markers("|") with nulls
   DIM pchar AS WCHAR PTR = *cwsFilter
   DIM i AS LONG
   FOR i = 0 TO LEN(cwsFilter) - 1
      IF pchar[i] = ASC("|") THEN pchar[i] = 0
   NEXT

   ' // If the initial directory has not been specified, assume the current directory
   IF LEN(wszInitialDir) = 0 THEN wszInitialDir = CURDIR
   DIM wszFile AS WSTRING * MAX_PATH = wszFileName
   DIM cwsFile AS CWSTR = wszFile & "|"
   ' // Store the position of the marker
   DIM cbPos AS LONG = LEN(cwsFile) - 1
   ' // Allocate room for the buffer
   IF LEN(cwsFile) < MAX_PATH THEN cwsFile += SPACE(MAX_PATH - LEN(cwsFile))
   DIM dwFileStrSize AS DWORD = LEN(cwsFile)
   ' // The filename must be null terminated (replace the marker with a null)
   pchar = *cwsFile
   pchar[cbPos] = 0

   ' // Fill the members of the structure
   DIM ofn AS OPENFILENAMEW
   ofn.lStructSize     = SIZEOF(ofn)
   IF AfxWindowsVersion < 5 THEN ofn.lStructSize = 76
   ofn.lpstrFilter     = *cwsFilter
   ofn.nFilterIndex    = 1
   ofn.lpstrFile       = *cwsFile
   ofn.nMaxFile        = dwFileStrSize
   ofn.lpstrInitialDir = @wszInitialDir
   IF LEN(wszTitle) THEN ofn.lpstrTitle = @wszTitle
   ofn.Flags = dwFlags OR OFN_EXPLORER
   IF LEN(wszDefExt) THEN ofn.lpstrDefExt = @wszDefExt

   ' // Call the save filename dialog
   IF GetSaveFilenameW(@ofn) = 0 THEN cwsFile = ""

   ' // Return the retrieved values
   IF pdwFlags THEN *pdwFlags = ofn.Flags
   RETURN cwsFile

END FUNCTION
' ========================================================================================

' ########################################################################################
'                                      *** WINDOW ***
' ########################################################################################

' ========================================================================================
' Gets the class name of the specified window
' ========================================================================================
PRIVATE FUNCTION AfxGetWindowClassName (BYVAL hwnd AS HWND) AS CWSTR
   DIM wszClassName AS WSTRING * 260
   GetClassNameW hwnd, wszClassName, SIZEOF(wszClassName)
   RETURN wszClassName
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the text of a window. This function can also be used to retrieve the text of buttons,
' and edit and static controls.
' Remarks: The function uses the WM_GETTEXT message because GetWindowText cannot retrieve
' the text of a window in another application.
' Example: DIM cws AS CWSTR = AfxGetWindowText(hwnd)
' ========================================================================================
PRIVATE FUNCTION AfxGetWindowText (BYVAL hwnd AS HWND) AS CWSTR
   DIM nLen AS LONG = SendMessageW(hwnd, WM_GETTEXTLENGTH, 0, 0)
   DIM cwsText AS CWSTR = SPACE(nLen + 1)
   SendMessageW(hwnd, WM_GETTEXT, nLen + 1, cast(LPARAM, *cwsText))
   RETURN cwsText
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the length of the text of a window.
' Usage example: nLen = AfxGetWindowTextLength(hwnd)
' ========================================================================================
PRIVATE FUNCTION AfxGetWindowTextLength (BYVAL hwnd AS HWND) AS LONG
   FUNCTION = SendMessageW(hwnd, WM_GETTEXTLENGTH, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the text of a window. This function can also be used to set the text of buttons,
' edit and static controls. The return value is TRUE if the text is set or FALSE if it is not.
' Remarks: The function uses the WM_SETTEXT message because SetWindowText cannot retrieve
' the text of a window in another application.
' Usage examples:
'   DIM wszText AS WSTRING * 260 = "Some text"
'   AfxSetWindowText(hwnd, @wszText)
' -or-
'   AfxSetWindowText(hwnd, "Some text")
' ========================================================================================
PRIVATE FUNCTION AfxSetWindowText (BYVAL hwnd AS HWND, BYVAL pwszText AS WSTRING PTR) AS BOOLEAN
   FUNCTION = SendMessageW(hwnd, WM_SETTEXT, 0, CAST(LPARAM, pwszText))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the window styles.
' ========================================================================================
PRIVATE FUNCTION AfxGetWindowStyle (BYVAL hwnd AS HWND) AS DWORD
   FUNCTION = GetWindowLongPtrW(hwnd, GWL_STYLE)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the extended window styles.
' ========================================================================================
PRIVATE FUNCTION AfxGetWindowExStyle (BYVAL hwnd AS HWND) AS DWORD
   FUNCTION = GetWindowLongPtrW(hwnd, GWL_EXSTYLE)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the window styles
' ========================================================================================
PRIVATE FUNCTION AfxSetWindowStyle (BYVAL hwnd AS HWND, BYVAL dwStyle AS DWORD) AS DWORD
   DIM dwOldStyle AS DWORD = GetWindowLongPtrW(hwnd, GWL_STYLE)
   SetWindowLongPtrW(hwnd, GWL_STYLE, dwStyle)
   FUNCTION = dwOldStyle
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the window extended styles
' ========================================================================================
PRIVATE FUNCTION AfxSetWindowExStyle (BYVAL hwnd AS HWND, BYVAL dwExStyle AS DWORD) AS DWORD
   DIM dwOldExStyle AS DWORD = GetWindowLongPtrW(hwnd, GWL_EXSTYLE)
   SetWindowLongPtrW(hwnd, GWL_EXSTYLE, dwExStyle)
   FUNCTION = dwOldExStyle
END FUNCTION
' ========================================================================================

' ========================================================================================
' Adds a new style to the specified window.
' - hwnd  = Window handle
' - dwStyle = Style to add
' Return value:
'   The previous window styles
' ========================================================================================
PRIVATE FUNCTION AfxAddWindowStyle (BYVAL hwnd AS HWND, BYVAL dwStyle AS DWORD) AS DWORD
   DIM dwOldStyle AS DWORD = GetWindowLongPtrW(hwnd, GWL_STYLE)
   SetWindowLongPtrW(hwnd, GWL_STYLE, dwOldStyle OR dwStyle)
   FUNCTION = dwOldStyle
END FUNCTION
' ========================================================================================

' ========================================================================================
' Adds a new extended style to the specified window.
' - hwnd  = Window handle
' - dwExStyle = Style to add
' Return value:
'   The previous extended window styles
' ========================================================================================
PRIVATE FUNCTION AfxAddWindowExStyle (BYVAL hwnd AS HWND, BYVAL dwExStyle AS DWORD) AS DWORD
   DIM dwOldExStyle AS DWORD = GetWindowLongPtrW(hwnd, GWL_EXSTYLE)
   SetWindowLongPtrW(hwnd, GWL_EXSTYLE, dwOldExStyle OR dwExStyle)
   FUNCTION = dwOldExStyle
END FUNCTION
' ========================================================================================

' ========================================================================================
' Removes an style from the specified window.
' - hwnd  = Window handle
' - dwStyle = Style to remove
' Return value:
'   The previous window styles
' ========================================================================================
PRIVATE FUNCTION AfxRemoveWindowStyle (BYVAL hwnd AS HWND, BYVAL dwStyle AS DWORD) AS DWORD
   DIM dwOldStyle AS DWORD = GetWindowLongPtrW(hwnd, GWL_STYLE)
   SetWindowLongPtrW(hwnd, GWL_STYLE, dwOldStyle AND (NOT dwStyle))
   FUNCTION = dwOldStyle
END FUNCTION
' ========================================================================================

' ========================================================================================
' Removes an extended style from the specified window.
' - hwnd  = Window handle
' - dwExStyle = Style to remove
' Return value:
'   The previous window styles
' ========================================================================================
PRIVATE FUNCTION AfxRemoveWindowExStyle (BYVAL hwnd AS HWND, BYVAL dwExStyle AS DWORD) AS DWORD
   DIM dwOldExStyle AS DWORD = GetWindowLongPtrW(hwnd, GWL_EXSTYLE)
   SetWindowLongPtrW(hwnd, GWL_EXSTYLE, dwOldExStyle AND (NOT dwEXStyle))
   FUNCTION = dwOldEXStyle
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the dimensions of the bounding rectangle of the specified window. The dimensions
' are given in screen coordinates that are relative to the upper-left corner of the screen.
' Note: To retrieve the height of the desktop window pass the handle returned by the
' API function GetDesktopWindow.
' ========================================================================================
PRIVATE FUNCTION AfxGetWindowRect (BYVAL hwnd AS HWND) AS RECT
   DIM rc AS RECT
   GetWindowRect(hwnd, @rc)
   FUNCTION = rc
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the width in pixels of a window.
' Note: To retrieve the height of the desktop window pass the handle returned by the
' API function GetDesktopWindow.
' ========================================================================================
PRIVATE FUNCTION AfxGetWindowWidth (BYVAL hwnd AS HWND) AS LONG
   DIM rc AS RECT
   GetWindowRect(hwnd, @rc)
   FUNCTION = rc.Right - rc.Left
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the height in pixels of a window.
' Note: To retrieve the height of the desktop window pass the handle returned by the
' API function GetDesktopWindow.
' ========================================================================================
PRIVATE FUNCTION AfxGetWindowHeight (BYVAL hwnd AS HWND) AS LONG
   DIM rc AS RECT
   GetWindowRect(hwnd, @rc)
   FUNCTION = rc.Bottom - rc.Top
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the coordinates of a window's client area. The client coordinates specify the
' upper-left and lower-right corners of the client area. Because client coordinates are
' relative to the upper-left corner of a window's client area, the coordinates of the
' upper-left corner are (0,0).
' Note: To retrieve the height of the desktop window pass the handle returned by the
' API function GetDesktopWindow.
' ========================================================================================
PRIVATE FUNCTION AfxGetWindowClientRect (BYVAL hwnd AS HWND) AS RECT
   DIM rc AS RECT
   GetClientRect(hwnd, @rc)
   FUNCTION = rc
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the width in pixels of the client area of a window.
' Note: To retrieve the height of the desktop window pass the handle returned by the
' API function GetDesktopWindow.
' ========================================================================================
PRIVATE FUNCTION AfxGetWindowClientWidth (BYVAL hwnd AS HWND) AS LONG
   DIM rc AS RECT
   GetClientRect(hwnd, @rc)
   FUNCTION = rc.Right - rc.Left
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the height in pixels of the client area of a window.
' Note: To retrieve the height of the desktop window pass the handle returned by the
' API function GetDesktopWindow.
' ========================================================================================
PRIVATE FUNCTION AfxGetWindowClientHeight (BYVAL hwnd AS HWND) AS LONG
   DIM rc AS RECT
   GetClientRect(hwnd, @rc)
   FUNCTION = rc.Bottom - rc.Top
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the coordinates of the work area on the primary display monitor expressed in
' virtual screen coordinates. The work area is the portion of the screen not obscured by
' the system taskbar or by application desktop toolbars. To get the work area of a monitor
' other than the primary display monitor, call the GetMonitorInfo function.
' ========================================================================================
PRIVATE FUNCTION AfxGetWorkAreaRect () AS RECT
   DIM rcWrk AS RECT
   SystemParametersInfo(SPI_GETWORKAREA, 0, @rcWrk, 0)
   FUNCTION = rcWrk
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION AfxGetWorkAreaWidth () AS LONG
   DIM rcWrk AS RECT
   SystemParametersInfo(SPI_GETWORKAREA, 0, @rcWrk, 0)
   FUNCTION = rcWrk.Right - rcWrk.Left
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION AfxGetWorkAreaHeight () AS LONG
   DIM rcWrk AS RECT
   SystemParametersInfo(SPI_GETWORKAREA, 0, @rcWrk, 0)
   FUNCTION = rcWrk.Bottom - rcWrk.Top
END FUNCTION
' ========================================================================================

' ========================================================================================
' Centers a window on the screen or over another window.
' It also ensures that the placement is done within the work area.
' Parameters:
' - hwnd = Handle of the window.
' - hwndParent = [optional] Handle of the parent window.
' ========================================================================================
PRIVATE SUB AfxCenterWindow (BYVAL hwnd AS HWND = NULL, BYVAL hwndParent AS HWND = NULL)

   DIM rc            AS RECT    ' Window coordinates
   DIM rcParent      AS RECT    ' Parent window coordinates
   DIM rcWorkArea    AS RECT    ' Work area coordinates
   DIM pt            AS POINT   ' x and y coordinates of centered window

   ' // Get the coordinates of the window
   GetWindowRect hwnd, @rc
   ' // Calculate the width and height of the window
   DIM nWidth AS LONG = rc.Right - rc.Left
   DIM nHeight AS LONG = rc.Bottom - rc.Top
   ' // Get the coordinates of the work area
   IF SystemParametersInfoW(SPI_GETWORKAREA, SIZEOF(rcWorkArea), @rcWorkArea, 0) = 0 THEN
      rcWorkArea.Right  = GetSystemMetrics(SM_CXSCREEN)
      rcWorkArea.Bottom = GetSystemMetrics(SM_CYSCREEN)
   END IF
   ' // Get the coordinates of the parent window
   IF hwndParent THEN
      GetWindowRect hwndParent, @rcParent
   ELSE
      rcParent.Left   = rcWorkArea.Left
      rcParent.Top    = rcWorkArea.Top
      rcParent.Right  = rcWorkArea.Right
      rcParent.Bottom = rcWorkArea.Bottom
   END IF
   ' // Calculate the width and height of the parent window
   DIM nParentWidth AS LONG = rcParent.Right - rcParent.Left
   DIM nParentHeight AS LONG = rcParent.Bottom - rcParent.Top
   ' // Calculate the new x coordinate and adjust for work area
   pt.x = rcParent.Left + ((nParentWidth - nWidth) \ 2)
   IF (pt.x < rcWorkArea.Left) THEN
      pt.x = rcWorkArea.Left
   ELSEIF ((pt.x + nWidth) > rcWorkArea.Right) THEN
      pt.x = rcWorkArea.Right - nWidth
   END IF
   ' // Calculate the new y coordinate and adjust for work area
   pt.y = rcParent.Top  + ((nParentHeight - nHeight) \ 2)
   IF (pt.y < rcWorkArea.Top) THEN
      pt.y = rcWorkArea.Top
   ELSEIF ((pt.y + nHeight) > rcWorkArea.Bottom) THEN
      pt.y = rcWorkArea.Bottom - nHeight
   END IF
   ' // Convert screen coordinates to client area coordinates
   IF (GetWindowLongPtrW(hwnd, GWL_STYLE) AND WS_CHILD) = WS_CHILD THEN ScreenToClient hwndParent, @pt
   ' // Reposition the window retaining its size and Z order
   SetWindowPos(hwnd, NULL, pt.x, pt.y, 0, 0, SWP_NOSIZE OR SWP_NOZORDER)

END SUB
' ========================================================================================

' ========================================================================================
' Adjusts the bounding rectangle of a window based on the desired size of the client area.
' Parameters:
' - hwnd = The window handle.
' - nWidth and nHeight = The desired size of the client area.
' - rxRatio and ryRatio = Scaling ratios.
' ========================================================================================
PRIVATE SUB AfxSetWindowClientSize (BYVAL hwnd AS HWND, BYVAL nWidth AS LONG, BYVAL nHeight AS LONG, BYVAL rxRatio AS SINGLE = 1, BYVAL ryRatio AS SINGLE = 1)

   DIM rc AS RECT, rcTemp  AS RECT
   ' // Convert the client rectangle to a window rectangle.
   ' // The AdjustWindowRectEx function cannot take menu wrapping into account
   ' // because it doesn't know which menu we are using.
   SetRect(@rc, 0, 0, nWidth * rxRatio, nHeight * ryRatio)
   DIM hMenu AS HANDLE = GetMenu(hwnd)
   DIM dwStyle AS DWORD = GetWindowLongPtrW(hwnd, GWL_STYLE)
   AdjustWindowRectEx(@rc, dwStyle, (hMenu <> NULL), GetWindowLongPtrW(hwnd, GWL_EXSTYLE))
   ' // If there is a menu, we need to check how much wrapping occurs when we set
   ' // the window to the width specified by AdjustWindowRectEX and an infinite
   ' // amount of height. An infinite height allows us to see every single menu wrap.
   IF hMenu <> NULL THEN
      rcTemp = rc
      rcTemp.Bottom = &H7FFF   ' // "Infinite" height
      SendMessageW(hwnd, WM_NCCALCSIZE, 0, CAST(LPARAM, @rcTemp))
      ' // Adjust our previous calculation to compensate for menu wrapping.
      rc.Bottom = rc.Bottom + rcTemp.Top
   END IF
   ' // The AdjustWindowRectEx function does not take the WS_VSCROLL or WS_HSCROLL
   ' // styles into account. To account for the scroll bars, we need to call the
   ' // GetSystemMetrics function with SM_CXVSCROLL or SM_CYHSCROLL.
   IF (dwStyle AND WS_HSCROLL) = WS_HSCROLL THEN
      rc.Bottom = rc.Bottom + GetSystemMetrics(SM_CYHSCROLL)
   END IF
   IF (dwStyle AND WS_VSCROLL) = WS_VSCROLL THEN
      rc.Right = rc.Right + GetSystemMetrics(SM_CXVSCROLL)
   END IF
   DIM cx AS LONG = rc.Right - rc.Left
   DIM cy AS LONG = rc.Bottom - rc.Top
   SetWindowPos(hwnd, NULL, 0, 0, cx, cy, SWP_NOZORDER OR SWP_NOMOVE OR SWP_NOACTIVATE)

END SUB
' ========================================================================================

' ========================================================================================
' Centers a control horizontally.
' hCtrl = Handle of the control.
' ========================================================================================
PRIVATE SUB AfxCenterControlH (BYVAL hCtrl AS HWND)
   DIM rc AS RECT, rcParent AS RECT
   ' // Get the coordinates of the control
   GetWindowRect hCtrl, @rc
   DIM nWidthControl AS LONG = rc.Right - rc.Left
   DIM nHeightControl AS LONG = rc.Bottom - rc.Top
   ' // Get the coordinates of the parent window
   DIM hParent AS HWND = GetParent(hCtrl)
   GetClientRect hParent, @rcParent
   DIM nWidthParent AS LONG = rcParent.Right - rcParent.Left
   ' // Calculate the x coordinate to center the control horizontally
   DIM x AS LONG = (nWidthParent - nWidthControl) \ 2
   ' // Convert the y coordinate of the control to client area coordinate
   DIM pt AS POINT
   pt.y = rc.Top
   ScreenToClient hParent, @pt
   MoveWindow hCtrl, x, pt.y, nWidthControl, nHeightControl, CTRUE
END SUB
' ========================================================================================

' ========================================================================================
' Centers a control vertically
' hCtrl = Handle of the control.
' ========================================================================================
PRIVATE SUB AfxCenterControlV (BYVAL hCtrl AS HWND)
   DIM rc AS RECT, rcParent AS RECT
   ' // Get the coordinates of the control
   GetWindowRect hCtrl, @rc
   DIM nWidthControl AS LONG = rc.Right - rc.Left
   DIM nHeightControl AS LONG = rc.Bottom - rc.Top
   ' // Get the coordinates of the parent window
   DIM hParent AS HWND = GetParent(hCtrl)
   GetClientRect hParent, @rcParent
   DIM nHeightParent AS LONG = rcParent.Bottom - rcParent.Top
   ' // Calculate the x coordinate to center the control vertically
   DIM y AS LONG = (nHeightParent - nHeightControl) \ 2
   ' // Convert the x coordinate of the control to client area coordinate
   DIM pt AS POINT
   pt.x = rc.Left
   ScreenToClient hParent, @pt
   MoveWindow hCtrl, pt.x + (nWidthControl \ 2), y, nWidthControl, nHeightControl, CTRUE
END SUB
' ========================================================================================

' ========================================================================================
' Redraws the specified window.
' Do not use it from within a WM_PAINT message.
' ========================================================================================
PRIVATE SUB AfxRedrawWindow (BYVAL hwnd AS HWND)
   InvalidateRect hwnd, NULL, CTRUE
   UpdateWindow hwnd
END SUB
' ========================================================================================

' ========================================================================================
' Redraws the non-client area of the specified window.
' Return value: TRUE or FALSE. To get extended error information, call GetLastError.
' ========================================================================================
PRIVATE FUNCTION AfxRedrawNonClientArea (BYVAL hwnd AS HWND) AS BOOLEAN
   FUNCTION = SetWindowPos(hwnd, 0, 0, 0, 0, 0, SWP_NOMOVE OR SWP_NOSIZE OR SWP_NOZORDER OR SWP_NOACTIVATE OR SWP_FRAMECHANGED OR SWP_DRAWFRAME)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sends a WM_SIZE message to the specified window.
' Parameters:
' - hwnd = Handle of the window.
' - nResizeType = Type of resizing requested.
' - nWidth = The new width of the client area.
' - nHeight = The new height of the client ara.
' Return value:
' If an application processes this message, it should return zero.
' ========================================================================================
PRIVATE FUNCTION AfxForwardSizeMessage (BYVAL hwnd AS HWND, BYVAL nResizeType AS DWORD, BYVAL nWidth AS LONG, BYVAL nHeight AS LONG) AS LRESULT
   FUNCTION = SendMessageW(hwnd, WM_SIZE, nResizeType, MAKELONG(nWidth, nHeight))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the location of the top left corner of the window, in pixels.
' The location is relative to the upper-left corner of the client area in the parent window.
' Return value: TRUE or FALSE. To get extended error information, call GetLastError.
' ========================================================================================
PRIVATE FUNCTION AfxSetWindowLocation (BYVAL hwnd AS HWND, BYVAL nLeft AS LONG, BYVAL nTop AS LONG) AS BOOLEAN
   FUNCTION = SetWindowPos(hwnd, 0, nLeft, nTop, 0, 0, SWP_NOSIZE OR SWP_NOZORDER)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the size in pixels of the specified window.
' - hwnd = Handle of the window.
' - nWidth = The new width of the client area.
' - nHeight = The new height of the client ara.
' Return value: TRUE or FALSE. To get extended error information, call GetLastError.
' ========================================================================================
PRIVATE FUNCTION AfxSetWindowSize (BYVAL hwnd AS HWND, BYVAL nWidth AS LONG, BYVAL nHeight AS LONG) AS BOOLEAN
   FUNCTION = SetWindowPos(hwnd, NULL, 0, 0, nWidth, nHeight, SWP_NOZORDER OR SWP_NOMOVE)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Change the visible state of a window.
' Parameters:
' - hwnd = The handle of the window.
' - nShowState
'   SW_FORCEMINIMIZE : Minimizes a window, even if the thread that owns the window is not
'   responding. This flag should only be used when minimizing windows from a different thread.
'   SW_HIDE : Hides the window and activates another window.
'   SW_MAXIMIZE : Maximizes the specified window.
'   SW_MINIMIZE : Minimizes the specified window and activates the next top-level window in the Z order.
'   SW_RESTORE : Activates and displays the window. If the window is minimized or maximized,
'   the system restores it to its original size and position. An application should specify
'   this flag when restoring a minimized window.
'   SW_SHOW : Activates the window and displays it in its current size and position.
'   SW_SHOWDEFAULT : Sets the show state based on the SW_ value specified in the STARTUPINFO
'   structure passed to the CreateProcess function by the program that started the application.
'   SW_SHOWMAXIMIZED : Activates the window and displays it as a maximized window.
'   SW_SHOWMINIMIZED : Activates the window and displays it as a minimized window.
'   SW_SHOWMINNOACTIVE : Displays the window as a minimized window. This value is similar
'   to SW_SHOWMINIMIZED, except the window is not activated.
'   SW_SHOWNA : Displays the window in its current size and position. This value is similar
'   to SW_SHOW, except that the window is not activated.
'   SW_SHOWNOACTIVATE : Displays a window in its most recent size and position. This value is
'   similar to SW_SHOWNORMAL, except that the window is not activated.
'   SW_SHOWNORMAL : Activates and displays a window. If the window is minimized or maximized,
'   the system restores it to its original size and position. An application should specify
'   this flag when displaying the window for the first time.
' Return value:
' TRUE if the window was previously visible.
' FALSE if the window was previously hidden.
' ========================================================================================
PRIVATE FUNCTION AfxShowWindowState (BYVAL hwnd AS HWND, BYVAL nShowState AS LONG) AS BOOLEAN
   FUNCTION = ShowWindow(hwnd, nShowState)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Associates a new icon with a window. The system displays the large icon in the
' ALT+TAB dialog box and the small icon in the window caption.
' Parameters:
' - hwnd  = Window handle
' - hIcon = Handle to the new icon. If this parameter is NULL, the icon is removed.
' - nIconType = ICON_BIG (1) or ICON_SMALL (0).
' The return value is a handle to the previous large or small icon, depending on the value of
' nIconType It is NULL if the window previously had no icon of the type indicated by nIconType.
' ========================================================================================
PRIVATE FUNCTION AfxSetWindowIcon (BYVAL hwnd AS HWND, BYVAL nIconType AS LONG, BYVAL hIcon AS HICON) AS HICON
   FUNCTION = CAST(HICON, SendMessageW(hwnd, WM_SETICON, nIconType, CAST(LPARAM, hIcon)))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Brings the thread that created the specified window into the foreground and activates
' the window. Keyboard input is directed to the window, and various visual cues are changed
' for the user. The system assigns a slightly higher priority to the thread that created
' the foreground window than it does to other threads.
' Replacement for the SetForegroundWindow API function, that sometimes fails.
' WARNING: Apparently, AttachThreadInput can lead to a deadlock.
' See: https://blogs.msdn.microsoft.com/oldnewthing/20080801-00/?p=21393
' ========================================================================================
'PRIVATE SUB AfxForceSetForegroundWindow (BYVAL hwnd AS HWND)
'   DIM dwProcessId AS DWORD
'   DIM hwndForeground AS .HWND = GetForegroundWindow
'   DIM dwThreadId AS DWORD = GetWindowThreadProcessId(hwndForeground, @dwProcessId)
'   DIM dwCurThreadId AS DWORD = GetCurrentThreadId
'   AttachThreadInput(dwCurThreadId, dwThreadId, CTRUE)
'   SetForegroundWindow(hwnd)
'   BringWindowToTop(hwnd)
'   SetFocus(hwnd)
'   AttachThreadInput(dwCurThreadId, dwThreadId, FALSE)
'END SUB
' ========================================================================================

' ========================================================================================
' If you use dual (or even triple/quad) displays then you have undoubtedly encountered the
' following situation: You change the physical order of your displays, or otherwise
' reconfigure the logical ordering using your display software. This sometimes has the
' side-effect of changing your desktop coordinates from zero-based to negative starting
' coordinates (i.e. the top-left coordinate of your desktop changes from 0,0 to -1024,-768).
' This effects many Windows programs which restore their last on-screen position whenever
' they are started. Should the user reorder their display configuration this can sometimes
' result in a Windows program subsequently starting in an off-screen position (i.e. at a
' location that used to be visible) - and is now effectively invisible, preventing the
' user from closing it down or otherwise moving it back on-screen.
' The ForceVisibleDisplay function can be called at program start-time right after the
' main window has been created and positioned 'on-screen'. Should the window be positioned
' in an off-screen position, it is forced back onto the nearest display to its last
' position. The user will be unaware this is happening and won't even realise to thank you
' for keeping their user-interface visible, even though they changed their display
' settings.
' Source: http://www.catch22.net/tuts/tips2
' ========================================================================================
PRIVATE SUB AfxForceVisibleDisplay (BYVAL hwnd AS HWND)
   ' // Check if the specified window-rectangle is visible on any display
   DIM rc AS RECT
   GetWindowRect(hwnd, @rc)
   IF MonitorFromRect(@rc, MONITOR_DEFAULTTONULL) <> NULL THEN EXIT SUB
   ' // Find the nearest display to the rectangle
   DIM hMonitor AS HMONITOR
   DIM mi AS MONITORINFO
   mi.cbSize = SIZEOF(mi)
   hMonitor = MonitorFromRect(@rc, MONITOR_DEFAULTTONEAREST)
   GetMonitorInfoW(hMonitor, @mi)
   ' // Center window rectangle
   rc.left = mi.rcWork.left + ((mi.rcWork.right - mi.rcWork.left) - (rc.right-rc.left)) \ 2
   rc.top = mi.rcWork.top + ((mi.rcWork.bottom - mi.rcWork.top) - (rc.bottom-rc.top)) \ 2
   SetWindowPos(hwnd, 0, rc.left, rc.top, 0, 0, SWP_NOACTIVATE OR SWP_NOZORDER OR SWP_NOSIZE)
END SUB
' ========================================================================================

' ========================================================================================
' Retrieves the window's top-level parent or owner window.
' ========================================================================================
PRIVATE FUNCTION AfxGetTopLevelWindow (BYVAL hwnd AS HWND) AS HWND
   DIM hWndParent AS HWND, hWndTmp AS HWND
   IF IsWindow(hwnd) = NULL THEN EXIT FUNCTION
   hWndTmp = hwnd
   DO
      hWndParent = hWndTmp
      hWndTmp = IIF&(GetWindowLongPtrW(hwndParent, GWL_STYLE) AND WS_CHILD, GetParent(hwndParent), GetWindow(hWndParent, GW_OWNER))
      IF hWndTmp = NULL THEN EXIT DO
   LOOP
   FUNCTION = hWndParent
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the window's top-level parent window.
' ========================================================================================
PRIVATE FUNCTION AfxGetTopLevelParent (BYVAL hwnd AS HWND) AS HWND
   DIM hWndParent AS HWND, hWndTmp AS HWND
   IF IsWindow(hwnd) = NULL THEN EXIT FUNCTION
   hWndParent = hwnd
   DO
      hWndTmp = GetParent(hWndParent)
      IF hWndTmp = NULL THEN EXIT DO
      hWndParent = hWndTmp
   LOOP
   FUNCTION = hWndParent
END FUNCTION
' ========================================================================================

' ========================================================================================
' Finds the handle of the top-level window or MDI child window that is the ancestor of the
' specified window handle. The reference handle is the handle of any control on the form.
' Code borrowed from a function written by Dominic Mitchell for Phoenix.
' ========================================================================================
PRIVATE FUNCTION AfxGetFormHandle (BYVAL hwnd AS HWND) AS HWND
   WHILE (GetWindowLongPtrW(hwnd, GWL_STYLE) AND WS_CHILD)
      IF (GetWindowLongPtrW(hwnd, GWL_EXSTYLE) AND WS_EX_MDICHILD) THEN EXIT WHILE
      hwnd = GetParent(hwnd)
   WEND
   FUNCTION = hwnd
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the handle of the control with the specified identifier. The reference handle
' can be the handle of the form or the handle of any other control on the form.
' Parameters:
' - hwnd = Reference window handle.
' - Control identifier.
' Returns the handle of the control or NULL.
' ========================================================================================

' // Enumerates child controls on the specified window.
' // Callback function for AfxGetControlHandle.
' // Parameters:
' // hwnd = Handle of child window.
' // lParam = Address of variable with child window handle
' // Return value: CTRUE or FALSE

PRIVATE FUNCTION AfxGetControlHandle_ChildEnumProc (BYVAL hwnd AS HWND, BYVAL lParam AS LPARAM) AS LONG
   DIM pItem AS HWND PTR
   CAST(LPARAM, pItem) = lParam
   IF GetDlgCtrlID(hwnd) = LOWORD(*pItem) THEN
      *pItem = hwnd
      EXIT FUNCTION
   END IF
   FUNCTION = CTRUE
END FUNCTION

PRIVATE FUNCTION AfxGetControlHandle (BYVAL hwnd AS HWND, BYVAL wCtrlID AS WORD) AS HWND
   DIM hwndChild AS HWND
   hwnd = AfxGetFormHandle(hwnd)
   hwndChild = CAST(HWND, CAST(DWORD_PTR, wCtrlID))
   EnumChildWindows(hwnd, @AfxGetControlHandle_ChildEnumProc, CAST(LPARAM, @hwndChild))
   IF hwndChild = wCtrlID THEN FUNCTION = NULL ELSE FUNCTION = hwndChild
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the handle of the enabled and visible window at the top of the z-order in an
' application.
' Return value: Handle of window at top of z-order or NULL.
' ========================================================================================

' // Callback enumeration procedure for finding the window at the top of the z-order.
' // Parameters:
' // hwnd = Handle of top-level window
' // lParam = Address of variable for window handle

PRIVATE FUNCTION AfxGetTopEnabledWindow_EnumWindowsProc (BYVAL hWnd AS HWND, BYVAL lParam AS LPARAM) AS LONG
   DIM wszClassName AS WSTRING * MAX_PATH, phwndTop AS HWND PTR, dwProcessId  AS DWORD
   IF IsWindowEnabled(hwnd) THEN
      IF IsWindowVisible(hwnd) THEN
         GetClassNameW hwnd, @wszClassName, MAX_PATH
         IF wszClassName <> "tooltips_class32" THEN
            GetWindowThreadProcessId(hwnd, @dwProcessId)
            IF dwProcessId = GetCurrentProcessId THEN
               CAST(LPARAM, phwndTop) = lParam
               *phwndTop = hwnd
               EXIT FUNCTION
            END IF
         END IF
      END IF
   END IF
   FUNCTION = CTRUE
END FUNCTION

PRIVATE FUNCTION AfxGetTopEnabledWindow () AS HWND
   DIM hwndTop AS HWND
   EnumWindows(@AfxGetTopEnabledWindow_EnumWindowsProc, CAST(LPARAM, @hwndTop))
   FUNCTION = hwndTop
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves a window handle given it's process identifier
' ========================================================================================
PRIVATE FUNCTION AfxGethWndFromPID (BYVAL PID AS DWORD) AS HWND
   DIM dwPID AS DWORD, dwThreadID AS DWORD
   ' // Get the first window handle
   DIM hwnd AS HWND = FindWindowW(NULL, NULL)
   ' // Enumerate all the windows
   WHILE hwnd <> NULL
      ' // If the parent window is NULL, it's a top level window
      IF GetParent(hwnd) = NULL THEN
         ' // Get it's process id
         dwThreadID = GetWindowThreadProcessId(hwnd, @dwPID)
         IF dwPID = PID THEN
            ' // We have found it
            FUNCTION = hwnd
            EXIT WHILE
         END IF
      END IF
      ' // Get the next window handle
      hwnd = GetWindow(hwnd, GW_HWNDNEXT)
   WEND
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the location of the top left corner of the window, in pixels.
' The location is relative to the upper-left corner of the client area in the parent window.
' ========================================================================================
PRIVATE SUB AfxGetWindowLocation (BYVAL hwnd AS HWND, BYVAL nLeft AS LONG PTR, BYVAL nTop AS LONG PTR)
   DIM rc AS RECT
   ' // Get the dimensions of the window
   GetWindowRect(hwnd, @rc)
   ' // Convert the coordinates to be relative to the parent
   MapWindowPoints(HWND_DESKTOP, GetParent(hwnd), CAST(POINT PTR, @rc), 2)
   ' // Return the left and top values
   *nLeft = rc.Left
   *nTop = rc.Top
END SUB
' ========================================================================================

' ========================================================================================
' Sets the return value of a message processed in the dialog box procedure.
' Return value: TRUE or FALSE.
' ========================================================================================
PRIVATE FUNCTION AfxSetDlgMsgResult(BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL result AS LONG) AS BOOLEAN
   IF uMsg = WM_CTLCOLORMSGBOX      OR _
      uMsg = WM_CTLCOLOREDIT        OR _
      uMsg = WM_CTLCOLORLISTBOX     OR _
      uMsg = WM_CTLCOLORBTN         OR _
      uMsg = WM_CTLCOLORDLG         OR _
      uMsg = WM_CTLCOLORSCROLLBAR   OR _
      uMsg = WM_CTLCOLORSTATIC      OR _
      uMsg = WM_COMPAREITEM         OR _
      uMsg = WM_VKEYTOITEM          OR _
      uMsg = WM_CHARTOITEM          OR _
      uMsg = WM_QUERYDRAGICON       OR _
      uMsg = WM_INITDIALOG          THEN
      FUNCTION = IIF&(SetWindowLongPtrW(hwnd, DWLP_MSGRESULT, result), TRUE, FALSE)
   END IF
END FUNCTION
' ========================================================================================


' ########################################################################################
'                              *** DPI RELATED PROCEDURES ***
' ########################################################################################

' ========================================================================================
' Sets the current process as dots per inch (dpi) aware.
' Note: SetProcessDPIAware is subject to a possible race condition if a DLL caches dpi
' settings during initialization. For this reason, it is recommended that dpi-aware be set
' through the application (.exe) manifest rather than by calling SetProcessDPIAware.
' Return value: TRUE on success; FALSE on failure.
' ========================================================================================
PRIVATE FUNCTION AfxSetProcessDPIAware () AS BOOLEAN
   DIM AS ANY PTR pLib = DyLibLoad("user32.dll")
   IF pLib = 0 THEN EXIT FUNCTION
   DIM pProc AS FUNCTION () AS LONG
   pProc = DyLibSymbol(pLib, "SetProcessDPIAware")
   IF pProc = 0 THEN EXIT FUNCTION
   FUNCTION = pProc()
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines whether the current process is dots per inch (dpi) aware such that it adjusts
' the sizes of UI elements to compensate for the dpi setting.
' Return value: TRUE or FALSE
' ========================================================================================
PRIVATE FUNCTION AfxIsProcessDPIAware () AS BOOLEAN
   DIM AS ANY PTR pLib = DyLibLoad("user32.dll")
   IF pLib = 0 THEN EXIT FUNCTION
   DIM pProc AS FUNCTION () AS LONG
   pProc = DyLibSymbol(pLib, "IsProcessDPIAware")
   IF pProc = 0 THEN EXIT FUNCTION
   FUNCTION = pProc()
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the value of the UseDpiScaling setting (Vista/Windows 7+).
' Returns TRUE if the OS uses DPI scaling; FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION AfxUseDpiScaling () AS BOOLEAN
   DIM hkRes AS HKEY, dwType AS DWORD, dwData AS DWORD, cbData AS DWORD
   IF RegOpenKeyExW(HKEY_CURRENT_USER, "Software\Microsoft\Windows\DWM", 0, KEY_QUERY_VALUE, @hkRes) = ERROR_SUCCESS THEN
      IF hkRes THEN
         cbData = SIZEOF(cbData)
         DIM hr AS LONG = RegQueryValueExW(hkRes, "UseDpiScaling", 0, @dwType, CPTR(BYTE PTR, @dwData), @cbData)
         RegCloseKey hkRes
         IF hr = ERROR_SUCCESS THEN FUNCTION = (dwData <> 0)
      END IF
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the number of pixels per logical inch along the screen width of the desktop
' window. In a system with multiple display monitors, this value is the same for all monitors.
' ========================================================================================
PRIVATE FUNCTION AfxLogPixelsX () AS LONG
   DIM hDC AS HDC = GetDC(HWND_DESKTOP)
   DIM dpiX AS LONG = GetDeviceCaps(hDC, LOGPIXELSX)
   ReleaseDC HWND_DESKTOP, hDC
   FUNCTION = dpiX
END FUNCTION
' ========================================================================================
#define AfxGetDpi AfxLogPixelsX
#define AfxGetDpiX AfxLogPixelsX

' ========================================================================================
' Retrieves the number of pixels per logical inch along the screen height of the desktop
' window. In a system with multiple display monitors, this value is the same for all monitors.
' ========================================================================================
PRIVATE FUNCTION AfxLogPixelsY () AS LONG
   DIM hDC AS HDC = GetDC(HWND_DESKTOP)
   DIM dpiY AS LONG = GetDeviceCaps(hDC, LOGPIXELSY)
   ReleaseDC HWND_DESKTOP, hDC
   FUNCTION = dpiY
END FUNCTION
' ========================================================================================
#define AfxGetDpiY AfxLogPixelsY

' ========================================================================================
' Retrieves the desktop horizontal scaling ratio.
' ========================================================================================
PRIVATE FUNCTION AfxScaleRatioX () AS SINGLE
   DIM hDC AS HDC = GetDC(HWND_DESKTOP)
   DIM rxRatio AS SINGLE = (GetDeviceCaps(hDC, LOGPIXELSX) / 96)
   ReleaseDC HWND_DESKTOP, hDC
   FUNCTION = rxRatio
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the desktop vertical scaling ratio.
' ========================================================================================
PRIVATE FUNCTION AfxScaleRatioY () AS SINGLE
   DIM hDC AS HDC = GetDC(HWND_DESKTOP)
   DIM ryRatio AS SINGLE = (GetDeviceCaps(hDC, LOGPIXELSY) / 96)
   ReleaseDC HWND_DESKTOP, hDC
   FUNCTION = ryRatio
END FUNCTION
' ========================================================================================

' ========================================================================================
' Scales an horizontal coordinate according the DPI (dots per pixel) being used by the desktop.
' ========================================================================================
PRIVATE FUNCTION AfxScaleX (BYVAL cx AS SINGLE) AS SINGLE
   DIM hDC AS HDC = GetDC(HWND_DESKTOP)
   FUNCTION = cx * (GetDeviceCaps(hDC, LOGPIXELSX) / 96)
   ReleaseDC HWND_DESKTOP, hDC
END FUNCTION
' ========================================================================================

' ========================================================================================
' Scales a vertical coordinate according the DPI (dots per pixel) being used by the desktop.
' ========================================================================================
PRIVATE FUNCTION AfxScaleY (BYVAL cy AS SINGLE) AS SINGLE
   DIM hDC AS HDC = GetDC(HWND_DESKTOP)
   FUNCTION = cy * (GetDeviceCaps(hDC, LOGPIXELSY) / 96)
   ReleaseDC HWND_DESKTOP, hDC
END FUNCTION
' ========================================================================================

' ========================================================================================
' Unscales an horizontal coordinate according the DPI (dots per pixel) being used by the desktop.
' ========================================================================================
PRIVATE FUNCTION AfxUnscaleX (BYVAL cx AS SINGLE) AS SINGLE
   DIM hDC AS HDC = GetDC(HWND_DESKTOP)
   FUNCTION = cx / (GetDeviceCaps(hDC, LOGPIXELSX) / 96)
   ReleaseDC HWND_DESKTOP, hDC
END FUNCTION
' ========================================================================================

' ========================================================================================
' Unscales a vertical coordinate according the DPI (dots per pixel) being used by the desktop.
' ========================================================================================
PRIVATE FUNCTION AfxUnscaleY (BYVAL cy AS SINGLE) AS SINGLE
   DIM hDC AS HDC = GetDC(HWND_DESKTOP)
   FUNCTION = cy / (GetDeviceCaps(hDC, LOGPIXELSY) / 96)
   ReleaseDC HWND_DESKTOP, hDC
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines if screen resolution meets minimum requirements.
' Parameters:
' - cxMin = Minimum screen resolution width in pixels.
' - cxMin = Minimum screen resolution height in pixels.
' Return value: TRUE or FALSE.
' ========================================================================================
PRIVATE FUNCTION AfxIsResolutionAtLeast (BYVAL cxMin AS LONG, BYVAL cyMin AS LONG) AS BOOLEAN
   DIM ScreenWidth AS LONG = GetSystemMetrics(SM_CXSCREEN)
   DIM ScreenHeight AS LONG = GetSystemMetrics(SM_CYSCREEN)
   IF (cxMin <= ScreenWidth) AND (cyMin <= ScreenHeight) THEN FUNCTION = TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines if screen resolution meets minimum requirements in relative pixels,
' e.g. for a screen resolution of 1920x1080 pixels and a DPI of 192 (scaling ratio = 2),
' the maximum relative pixels for a DPI aware application is 960x540.
' - cxMin = Minimum screen resolution width in relative pixels.
' - cxMin = Minimum screen resolution height in relative pixels.
' Return value: TRUE or FALSE.
' ========================================================================================
PRIVATE FUNCTION AfxIsDPIResolutionAtLeast (BYVAL cxMin AS LONG, BYVAL cyMin AS LONG) AS BOOLEAN
   ' // Get de DPI values used by the desktop window
   DIM hDC AS HDC = GetDC(HWND_DESKTOP)
   DIM dpiX AS LONG = GetDeviceCaps(hDC, LOGPIXELSX)
   DIM dpiY AS LONG = GetDeviceCaps(hDC, LOGPIXELSY)
   ReleaseDC HWND_DESKTOP, hDC
   ' // Scale the values
   cxMin = cxMin * dpiX / 96
   cyMin = cyMin * dpiX / 96
   ' // Calculate the width and height of the primary display monitor, in pixels
   DIM ScreenWidth AS LONG = GetSystemMetrics(SM_CXSCREEN)
   DIM ScreenHeight AS LONG = GetSystemMetrics(SM_CYSCREEN)
   IF (cxMin <= ScreenWidth) AND (cyMin <= ScreenHeight) THEN FUNCTION = TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Loads a specified icon resource with a client-specified system metric.
' See: https://msdn.microsoft.com/en-us/library/windows/desktop/bb775701(v=vs.85).aspx
' ========================================================================================
PRIVATE FUNCTION AfxLoadIconMetric (BYVAL hinst AS HINSTANCE, BYVAL pszName AS PCWSTR, BYVAL lims AS LONG, BYVAL phico AS HICON PTR) AS HRESULT
   DIM AS ANY PTR pLib = DyLibLoad("Comctl32.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pLoadIconMetric AS FUNCTION (BYVAL hinst AS HINSTANCE, BYVAL pszName AS PCWSTR, BYVAL lims AS LONG, BYVAL phico AS HICON PTR) AS HRESULT
   pLoadIconMetric = DyLibSymbol(pLib, "LoadIconMetric")
   IF pLoadIconMetric = NULL THEN EXIT FUNCTION
   FUNCTION = pLoadIconMetric(hinst, pszName, lims, phico)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================


' ########################################################################################
'                                *** METRIC CONVERSIONS ***
' ########################################################################################

' ========================================================================================
' Converts from HiMetric to Pixels
' Note: HiMetric is a scaling unit similar to twips used in computing. It is one
' thousandth of a centimeter and is independent of the screen resolution.
' HiMetric per inch = 2540   ' 1 inch = 2.54 mm
' ========================================================================================
' Horizontal resolution
PRIVATE FUNCTION AfxHiMetricToPixelsX (BYVAL hm AS LONG) AS LONG
   DIM hDC AS HDC = GetDC(HWND_DESKTOP)
   DIM nPixelsPerLogicalInchX AS LONG = GetDeviceCaps(hDC, LOGPIXELSX)
   ReleaseDC(HWND_DESKTOP, hDC)
   FUNCTION = MulDiv(hm, nPixelsPerLogicalInchX, 2540)
END FUNCTION
' ========================================================================================
' Vertical resolution
PRIVATE FUNCTION AfxHiMetricToPixelsY (BYVAL hm AS LONG) AS LONG
   DIM hDC AS HDC = GetDC(HWND_DESKTOP)
   DIM nPixelsPerLogicalInchY AS LONG = GetDeviceCaps(hDC, LOGPIXELSY)
   ReleaseDC(HWND_DESKTOP, hDC)
   FUNCTION = MulDiv(hm, nPixelsPerLogicalInchY, 2540)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts from Pixels to HiMetric
' Note: HiMetric is a scaling unit similar to twips used in computing. It is one
' thousandth of a centimeter and is independent of the screen resolution.
' HiMetric per inch = 2540   ' 1 inch = 2.54 mm
' ========================================================================================
' Horizontal resolution
PRIVATE FUNCTION AfxPixelsToHiMetricX (BYVAL cx AS LONG) AS LONG
   DIM hDC AS HDC = GetDC(HWND_DESKTOP)
   DIM nPixelsPerLogicalInchX AS LONG = GetDeviceCaps(hDC, LOGPIXELSX)
   ReleaseDC(HWND_DESKTOP, hDC)
   FUNCTION = MulDiv(cx, 2540, nPixelsPerLogicalInchX)
END FUNCTION
' ========================================================================================
' Vertical resolution
PRIVATE FUNCTION AfxPixelsToHiMetricY (BYVAL cy AS LONG) AS LONG
   DIM hDC AS HDC = GetDC(HWND_DESKTOP)
   DIM nPixelsPerLogicalInchY AS LONG = GetDeviceCaps(hDC, LOGPIXELSY)
   ReleaseDC(HWND_DESKTOP, hDC)
   FUNCTION = MulDiv(cy, 2540, nPixelsPerLogicalInchY)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts pixels to point size (1/72 of an inch).
' ========================================================================================
' Horizontal resolution
PRIVATE FUNCTION AfxPixelsToPointsX (BYVAL pix AS LONG) AS LONG
   DIM hDC AS HDC = GetDC(HWND_DESKTOP)
   DIM LPX AS LONG = GetDeviceCaps(hDC, LOGPIXELSX)
   ReleaseDC(HWND_DESKTOP, hDC)
   FUNCTION = pix * 72 / LPX
END FUNCTION
' ========================================================================================
' Vertical resolution
PRIVATE FUNCTION AfxPixelsToPointsY (BYVAL pix AS LONG) AS LONG
   DIM hDC AS HDC = GetDC(HWND_DESKTOP)
   DIM LPY AS LONG = GetDeviceCaps(hDC, LOGPIXELSY)
   ReleaseDC(HWND_DESKTOP, hDC)
   FUNCTION = pix * 72 / LPY
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts a point size (1/72 of an inch) to pixels. Horizontal resolution.
' ========================================================================================
' Horizontal resolution
PRIVATE FUNCTION AfxPointsToPixelsX (BYVAL pts AS LONG) AS LONG
   DIM hDC AS HDC = GetDC(HWND_DESKTOP)
   DIM LPX AS LONG = GetDeviceCaps(hDC, LOGPIXELSX)
   ReleaseDC(HWND_DESKTOP, hDC)
   FUNCTION = MulDiv(pts, LPX, 72)
END FUNCTION
' ========================================================================================
' Vertical resolution
PRIVATE FUNCTION AfxPointsToPixelsY (BYVAL pts AS LONG) AS LONG
   DIM hDC AS HDC = GetDC(HWND_DESKTOP)
   DIM LPY AS LONG = GetDeviceCaps(hDC, LOGPIXELSY)
   ReleaseDC(HWND_DESKTOP, hDC)
   FUNCTION = MulDiv(pts, LPY, 72)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts pixels to twips.
' Twips are screen-independent units to ensure that the proportion of screen elements are
' the same on all display systems. A twip is defined as being 1/1440 of an inch.
' ========================================================================================
' Horizontal resolution
PRIVATE FUNCTION AfxPixelsToTwipsX (BYVAL nPixels AS LONG) AS LONG
   DIM hDC AS HDC = GetDC(HWND_DESKTOP)
   DIM LPX AS LONG = GetDeviceCaps(hDC, LOGPIXELSX)
   ReleaseDC(HWND_DESKTOP, hDC)
   FUNCTION = (nPixels * 1440) / LPX
END FUNCTION
' ========================================================================================
' Vertical resolution.
PRIVATE FUNCTION AfxPixelsToTwipsY (BYVAL nPixels AS LONG) AS LONG
   DIM hDC AS HDC = GetDC(HWND_DESKTOP)
   DIM LPY AS LONG = GetDeviceCaps(hDC, LOGPIXELSY)
   ReleaseDC(HWND_DESKTOP, hDC)
   FUNCTION = (nPixels * 1440) / LPY
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts twips to pixels.
' Twips are screen-independent units to ensure that the proportion of screen elements are
' the same on all display systems. A twip is defined as being 1/1440 of an inch.
' ========================================================================================
' Horizontal resolution
PRIVATE FUNCTION AfxTwipsToPixelsX (BYVAL nTwips AS LONG) AS LONG
   DIM hDC AS HDC = GetDC(HWND_DESKTOP)
   DIM LPX AS LONG = GetDeviceCaps(hDC, LOGPIXELSX)
   ReleaseDC(HWND_DESKTOP, hDC)
   FUNCTION = (nTwips / 1440) * LPX
END FUNCTION
' ========================================================================================
' Vertical resolution
PRIVATE FUNCTION AfxTwipsToPixelsY (BYVAL nTwips AS LONG) AS LONG
   DIM hDC AS HDC = GetDC(HWND_DESKTOP)
   DIM LPY AS LONG = GetDeviceCaps(hDC, LOGPIXELSY)
   ReleaseDC(HWND_DESKTOP, hDC)
   FUNCTION = (nTwips / 1440) * LPY
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the width of a pixel, in twips.
' Pixel dimensions can vary between systems and may not always be square, so separate
' functions for pixel width and height are required.
' ========================================================================================
' Horizontal resolution
PRIVATE FUNCTION AfxTwipsPerPixelX () AS LONG
   DIM hDC AS HDC = GetDC(HWND_DESKTOP)
   DIM LPX AS LONG = GetDeviceCaps(hDC, LOGPIXELSX)
   ReleaseDC(HWND_DESKTOP, hDC)
   FUNCTION = 1440 / LPX
END FUNCTION
' ========================================================================================
' Vertical resolution
PRIVATE FUNCTION AfxTwipsPerPixelY () AS LONG
   DIM hDC AS HDC = GetDC(HWND_DESKTOP)
   DIM LPY AS LONG = GetDeviceCaps(hDC, LOGPIXELSY)
   ReleaseDC(HWND_DESKTOP, hDC)
   FUNCTION = 1440 / LPY
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts point size to DIP (device independent pixel).
' DIP is defined as 1/96 of an inch and a point is 1/72 of an inch.
' ========================================================================================
PRIVATE FUNCTION AfxPointSizeToDip (BYVAL ptsize AS SINGLE) AS SINGLE
   FUNCTION = (ptsize / 72) * 96
END FUNCTION
' ========================================================================================


' ########################################################################################
'                                    *** CLIPBOARD ***
' ########################################################################################

' ========================================================================================
' Clears the contents of the clipboard.
' Return Value
'   If the function succeeds, the return value is nonzero.
'   If the function fails, the return value is zero.
' ========================================================================================
PRIVATE FUNCTION AfxClearClipboard () AS LONG
   ' // Opens the clipboard
   IF OpenClipboard(NULL) <> 0 THEN
      ' // Empties the clipboard
      FUNCTION = EmptyClipboard
      ' // Closes the clipboard
      CloseClipboard
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves data from the clipboard in the specified format.
' Parameter
'   cfFormat = Clipboard format.
' Return Value
'   If the function succeeds, the return value is the handle to the data.
'   If the function fails, the return value is NULL.
' ========================================================================================
PRIVATE FUNCTION AfxGetClipboardData (BYVAL cfFormat AS DWORD) AS HGLOBAL
   DIM hSource AS HANDLE
   ' // Opens the clipboard
   IF OpenClipboard(NULL) <> 0 THEN
      ' // Retrieves data from the clipboard in the specified format
      hSource = GetClipboardData(cfFormat)
      ' // Closes the clipboard
      CloseClipboard
      ' // Exits on failure
      IF hSource = NULL THEN EXIT FUNCTION
   END IF
   ' // Gets the size of the specified global memory object, in bytes
   DIM dwSize AS SIZE_T_ = GlobalSize(hSource)
   ' // Exits on failure
   IF dwSize = 0 THEN EXIT FUNCTION
   ' // Gets a pointer to the source memory object
   DIM pSource AS LPVOID = GlobalLock(hSource)
   ' // Exits on failure
   IF pSource = NULL THEN EXIT FUNCTION
   ' // Allocates the specified number of bytes from the heap
   DIM hDest AS HGLOBAL = GlobalAlloc(GHND_, dwSize)
   ' // Exits on failure
   IF hDest = NULL THEN
      ' // Unlocks the source memory object
      GlobalUnlock hSource
      EXIT FUNCTION
   END IF
   ' // Gets a pointer to the destination memory object
   DIM pDest AS LPVOID = GlobalLock(hDest)
   ' // Exits on failure
   IF pDest = NULL THEN
      ' // Unlocks the source memory object
      GlobalUnlock hSource
      ' // Frees the allocated memory block
      GlobalFree hDest
      EXIT FUNCTION
   END IF
   ' // Copies the data from the source to the destination
   memcpy pDest, pSource, dwSize
   ' // Unlocks the source memory object
   GlobalUnlock hSource
   ' // Unlocks the destination memory object
   GlobalUnlock hDest
   ' // Returns the handle to the data
   FUNCTION = hDest
END FUNCTION
' ========================================================================================

' ========================================================================================
' Places a data object into the clipboard.
' Parameters
'   cfFormat = Clipboard format.
'   hData    = Handle to the data in the specified format.
' Return Value
'   If the function succeeds, the return value is the handle to the data.
'   If the function fails, the return value is NULL.
' Remarks
'   The application must not use the hData handle once it has called the AfxSetClipboardData function.
' ========================================================================================
PRIVATE FUNCTION AfxSetClipboardData (BYVAL cfFormat AS DWORD, BYVAL hData AS HANDLE) AS HANDLE
   ' // Opens the clipboard
   IF OpenClipboard(NULL) <> 0 THEN
      ' // Empties the clipboard
      EmptyClipboard
      ' // Places the data object in the clipboard
      FUNCTION = SetClipboardData(cfFormat, hData)
      ' // Closes the clipboard
      CloseClipboard
    END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns an unicode text string from the clipboard.
' ========================================================================================
PRIVATE FUNCTION AfxGetClipboardText () AS CWSTR
   DIM cwsText AS CWSTR
   ' // Check if the required format is available
   IF IsClipboardFormatAvailable(CF_UNICODETEXT) = 0 THEN RETURN ""
   ' // Open the clipboard
   IF OpenClipboard(NULL) = 0 THEN RETURN ""
   ' // Get memory object of clipboard text
   DIM hMem AS HANDLE = GetClipboardData(CF_UNICODETEXT)
   IF hMem = NULL THEN
      CloseClipboard
      RETURN ""
   END IF
   ' // Lock it and get a pointer to the data
   DIM pMem AS HGLOBAL = GlobalLock(hMem)
   IF pMem = NULL THEN
      CloseClipboard
      RETURN ""
   END IF
   ' // Get the size of the global lock
   DIM dwSize AS DWORD = GlobalSize(hMem)
   IF dwSize > 0 THEN cwsText.AppendBuffer(pMem, dwSize)
   ' // Releases the memory object
   GlobalUnlock hMem
   ' // Closes the clipboard
   CloseClipboard
   RETURN cwsText
END FUNCTION
' ========================================================================================

' ========================================================================================
' Places an unicode text string into the clipboard.
' Parameter
'   wszText = Text to place in the clipboard.
' Return Value
'   If the function succeeds, the return value is the handle to the data.
'   If the function fails, the return value is NULL.
' ========================================================================================
PRIVATE FUNCTION AfxSetClipboardText (BYREF wszText AS WSTRING) AS HANDLE
   ' // Opens the clipboard
   IF OpenClipboard(NULL) <> 0 THEN
      ' // Empties the clipboard
      EmptyClipboard
      ' // Allocates a global memory block
      DIM hMem AS HGLOBAL = GlobalAlloc(GMEM_MOVEABLE OR GMEM_DDESHARE, (LEN(wszText) + 1) * 2)
      IF hMem <> NULL THEN
         ' // Locks it and gets a pointer to the memory location
         DIM pMem AS LPVOID = GlobalLock(hMem)
         ' // Copies the text into the allocated memory block
         IF pMem <> NULL THEN *CAST(WSTRING PTR, pMem) = wszText & CHR(0, 0)
         ' // Unlocks the memory block
         GlobalUnlock hMem
         ' // Places the text in the clipboard
         DIM hData AS HANDLE = SetClipboardData(CF_UNICODETEXT, hMem)
         IF hData <> NULL THEN
            ' // Returns the handle of the data
            FUNCTION = hData
         ELSE
            ' // Frees the memory block
            GlobalFree hMem
         END IF
      END IF
      ' // Closes the clipboard
      CloseClipboard
   END IF
END FUNCTION
' ========================================================================================


' ########################################################################################
'                                      *** FONTS ***
' ########################################################################################

' ========================================================================================
' Font enum
' ========================================================================================
enum
   AFX_FONT_CAPTION = 1    ' // Caption font
   AFX_FONT_SMALLCAPTION   ' // Small caption font
   AFX_FONT_MENU           ' // Font used in menu bars
   AFX_FONT_STATUS         ' // Font used in status bars and tooltips
   AFX_FONT_MESSAGE        ' // Font used in message boxes
   ' // Font settings
   AFX_FONT_HEIGHT         ' // Font height
   AFX_FONT_WEIGHT         ' // Font weight
   AFX_FONT_ITALIC         ' // Font italic
   AFX_FONT_UNDERLINE      ' // Font underline
   AFX_FONT_STRIKEOUT      ' // Font strikeout
   AFX_FONT_CHARSET        ' // Font charset
end enum
' ========================================================================================

' ========================================================================================
' Gets the font with which the window is currently drawing its text.
' Note: You can also use the GetWindowFont macro (windowsx.bi).
' ========================================================================================
PRIVATE FUNCTION AfxGetWindowFont (BYVAL hwnd AS HWND) AS HFONT
   FUNCTION = CAST(HFONT, SendMessageW(hwnd, WM_GETFONT, 0, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the font with which the window is currently drawing its text.
' hFont
'   Handle to the font. If this parameter is NULL, the control uses the default system
'   font to draw text.
' fRedraw
'   Specifies whether the control should be redrawn immediately upon setting the font. If
'   this parameter is TRUE, the control redraws itself.
' ========================================================================================
PRIVATE SUB AfxSetWindowFont (BYVAL hwnd AS HWND, BYVAL hFont AS HFONT, BYVAL fRedraw AS LONG = CTRUE)
   SendMessageW(hwnd, WM_SETFONT, CAST(WPARAM, hFont), fRedraw)
END SUB
' ========================================================================================

' ========================================================================================
' Retrieves information about the fonts used by Windows.
' Parameters:
' - nType = The type of font:
'   AFX_FONT_CAPTION, AFX_FONT_SMALLCAPTION, AFX_FONT_MENU, AFX_FONT_STATUS, AFX_FONT_MESSAGE
' - plfw = Pointer to a LOGFONTW structure that receives the font information.
' Return value: TRUE on succes or FALSE on failure.
' To get extended error information, call GetLastError.
' ========================================================================================
PRIVATE FUNCTION AfxGetWindowsFontInfo (BYVAL nType AS LONG, BYVAL plfw AS LOGFONTW PTR) AS BOOLEAN
   DIM ncm AS NONCLIENTMETRICSW
   IF plfw = NULL THEN EXIT FUNCTION
   IF AfxWindowsVersion >= 6 THEN ncm.cbSize = 504 ELSE ncm.cbSize = 500
   IF SystemParametersInfoW(SPI_GETNONCLIENTMETRICS, SIZEOF(ncm), @ncm, 0) = 0 THEN EXIT FUNCTION
   SELECT CASE nType
      CASE AFX_FONT_CAPTION      : *plfw = ncm.lfCaptionFont
      CASE AFX_FONT_SMALLCAPTION : *plfw = ncm.lfSmCaptionFont
      CASE AFX_FONT_MENU         : *plfw = ncm.lfMenuFont
      CASE AFX_FONT_STATUS       : *plfw = ncm.lfStatusFont
      CASE AFX_FONT_MESSAGE      : *plfw = ncm.lfMessageFont
      CASE ELSE
         RETURN FALSE
   END SELECT
   FUNCTION = TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the point size of the fonts used by Windows.
' Parameters:
' - nType = The type of font:
'   AFX_FONT_CAPTION, AFX_FONT_SMALLCAPTION, AFX_FONT_MENU, AFX_FONT_STATUS, AFX_FONT_MESSAGE
' Return value: The point size.
' ========================================================================================
PRIVATE FUNCTION AfxGetWindowsFontPointSize (BYVAL nType AS LONG) AS LONG
   DIM ncm AS NONCLIENTMETRICSW
   IF AfxWindowsVersion >= 6 THEN ncm.cbSize = 504 ELSE ncm.cbSize = 500
   IF SystemParametersInfoW(SPI_GETNONCLIENTMETRICS, SIZEOF(ncm), @ncm, 0) = 0 THEN EXIT FUNCTION
   DIM hDC AS HDC = CreateDCW("DISPLAY", NULL, NULL, NULL)
   IF hDC = NULL THEN EXIT FUNCTION
   DIM cyPixelsPerInch AS LONG = GetDeviceCaps(hDC, LOGPIXELSY)
   DeleteDC hDC
   DIM nPointSize AS LONG
   SELECT CASE nType
      CASE AFX_FONT_CAPTION      : nPointSize = MulDiv(ncm.lfCaptionFont.lfHeight, 72, cyPixelsPerInch)
      CASE AFX_FONT_SMALLCAPTION : nPointSize = MulDiv(ncm.lfSmCaptionFont.lfHeight, 72, cyPixelsPerInch)
      CASE AFX_FONT_MENU         : nPointSize = MulDiv(ncm.lfMenuFont.lfHeight, 72, cyPixelsPerInch)
      CASE AFX_FONT_STATUS       : nPointSize = MulDiv(ncm.lfStatusFont.lfHeight, 72, cyPixelsPerInch)
      CASE AFX_FONT_MESSAGE      : nPointSize = MulDiv(ncm.lfMessageFont.lfHeight, 72, cyPixelsPerInch)
   END SELECT
   IF nPointSize < 0 THEN nPointSize = -nPointSize
   FUNCTION = nPointSize
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the point size of a font given its logical height
' ========================================================================================
PRIVATE FUNCTION AfxGetFontPointSize (BYVAL nHeight AS LONG) AS LONG
   DIM hDC AS HDC = CreateDCW("DISPLAY", NULL, NULL, NULL)
   IF hDC = NULL THEN EXIT FUNCTION
   DIM cyPixelsPerInch AS LONG = GetDeviceCaps(hDC, LOGPIXELSY)
   DeleteDC HDC
   DIM nPointSize AS LONG = MulDiv(nHeight, 72, cyPixelsPerInch)
   IF nPointSize < 0 THEN nPointSize = -nPointSize
   FUNCTION = nPointSize
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the logical height of a font given its point size
' ========================================================================================
PRIVATE FUNCTION AfxGetFontHeight (BYVAL nPointSize AS LONG) AS LONG
   DIM hDC AS HDC = CreateDCW("DISPLAY", NULL, NULL, NULL)
   IF hDC = NULL THEN EXIT FUNCTION
   DIM cyPixelsPerInch AS LONG = GetDeviceCaps(hDC, LOGPIXELSY)
   DeleteDC HDC
   DIM nHeight AS LONG = -MulDiv(nPointSize, cyPixelsPerInch, 72)
   FUNCTION = nHeight
END FUNCTION
' ========================================================================================

' ========================================================================================
' Modifies the face name of the font of a window or control.
' Parameters:
' - hwnd = Handle of the window or control.
' - pwszNewFaceName = The new face name of the font.
' Return value: The handle of the new font on success, or NULL on failure.
' To get extended error information, call GetLastError.
' Remarks: The returned font must be destroyed with DeleteObject or the macro DeleteFont
' when no longer needed to prevent memory leaks.
' ========================================================================================
PRIVATE FUNCTION AfxModifyFontFaceName (BYVAL hwnd AS HWND, BYREF wszNewFaceName AS WSTRING) AS HFONT
   DIM lfw AS LOGFONTW
   IF hwnd = NULL OR VARPTR(wszNewFaceName) = NULL THEN EXIT FUNCTION
   ' // Get the handle of the font used by the header
   DIM hCurFont AS HFONT = CAST(HFONT, SendMessageW(hwnd, WM_GETFONT, 0, 0))
   IF hCurFont = 0 THEN EXIT FUNCTION
   ' // Get the LOGFONTW structure
   IF GetObject(hCurFont, SIZEOF(lfw), @lfw) = 0 THEN EXIT FUNCTION
   ' // Change the face name
   lfw.lfFaceName = wszNewFaceName
   ' // Create a new font
   DIM hNewFont AS HFONT = CreateFontIndirectW(@lfw)
   IF hNewFont = 0 THEN EXIT FUNCTION
   ' // Select the new font and delete the old one
   DIM hDC AS HDC = GetDC(hwnd)
   DeleteObject(SelectObject(hDC, CAST(HGDIOBJ, hNewFont)))
   ReleaseDC(hwnd, hDC)
   SendMessageW(hwnd, WM_SETFONT, CAST(WPARAM, hNewFont), CTRUE)
   FUNCTION = hNewFont
END FUNCTION
' ========================================================================================

' ========================================================================================
' Modifies settings of the font used by a window of control.
' Parameters:
' - hwnd = Handle of the window or control.
' - nSetting : One of the AFX_FONT_xxx constants.
' - nValue: Depends of the nSetting value
'   AFX_FONT_HEIGHT
'      The base is 100. To increase the font a 20% pass 120; to reduce it a 20% pass 80%.
'   AFX_FONT_WEIGHT
'      The weight of the font in the range 0 through 1000. For example, 400 is normal and
'      700 is bold. If this value is zero, a default weight is used.
'      The following values are defined for convenience.
'      FW_DONTCARE (0), FW_THIN (100), FW_EXTRALIGHT (200), FW_ULTRALIGHT (200), FW_LIGHT (300),
'      FW_NORMAL (400), FW_REGULAR (400), FW_MEDIUM (500), FW_SEMIBOLD (600), FW_DEMIBOLD (600),
'      FW_BOLD (700), FW_EXTRABOLD (800), FW_ULTRABOLD (800), FW_HEAVY (900), FW_BLACK (900)
'   AFX_FONT_ITALIC : TRUE or FALSE.
'   AFX_FONT_UNDERLINE : TRUE or FALSE.
'   AFX_FONT_STRIKEOUT : TRUE or FALSE.
'   AFX_FONT_CHARSET
'      The following values are predefined: ANSI_CHARSET, BALTIC_CHARSET, CHINESEBIG5_CHARSET,
'      DEFAULT_CHARSET, EASTEUROPE_CHARSET, GB2312_CHARSET, GREEK_CHARSET, HANGUL_CHARSET,
'      MAC_CHARSET, OEM_CHARSET, RUSSIAN_CHARSET, SHIFTJIS_CHARSET, SYMBOL_CHARSET, TURKISH_CHARSET,
'      VIETNAMESE_CHARSET, JOHAB_CHARSET (Korean language edition of Windows), ARABIC_CHARSET and
'      HEBREW_CHARSET (Middle East language edition of Windows), THAI_CHARSET (Thai language
'      edition of Windows).
'      The OEM_CHARSET value specifies a character set that is operating-system dependent.
'      DEFAULT_CHARSET is set to a value based on the current system locale. For example, when
'      the system locale is English (United States), it is set as ANSI_CHARSET.
'      Fonts with other character sets may exist in the operating system. If an application uses
'      a font with an unknown character set, it should not attempt to translate or interpret
'      strings that are rendered with that font.
'      This parameter is important in the font mapping process. To ensure consistent results,
'      specify a specific character set. If you specify a typeface name in the lfFaceName member,
'      make sure that the lfCharSet value matches the character set of the typeface specified in lfFaceName.
' Return value: The handle of the new font on success, or NULL on failure.
' To get extended error information, call GetLastError.
' Remarks: The returned font must be destroyed with DeleteObject or the macro DeleteFont
' when no longer needed to prevent memory leaks.
' ========================================================================================
PRIVATE FUNCTION AfxModifyFontSettings (BYVAL hwnd AS HWND, BYVAL nSetting AS LONG, BYVAL nValue AS LONG) AS HFONT
   DIM lfw AS LOGFONTW
   IF IsWindow(hwnd) = 0 THEN EXIT FUNCTION
   ' // Get the handle of the font used by the header
   DIM hCurFont AS HFONT = CAST(HFONT, SendMessageW(hwnd, WM_GETFONT, 0, 0))
   IF hCurFont = NULL THEN EXIT FUNCTION
   ' // Get the LOGFONTW structure
   IF GetObject(hCurFont, SIZEOF(lfw), @lfw) = 0 THEN EXIT FUNCTION
   ' // Change the specified setting
   SELECT CASE nSetting
      CASE AFX_FONT_HEIGHT
         ' // Change the point size
         DIM lPointSize AS LONG = AfxGetFontPointSize(lfw.lfHeight)
         lPointSize = lPointSize * (nValue / 100)
         lfw.lfHeight = -MulDiv(lPointSize, AfxLogPixelsY, 72)
      CASE AFX_FONT_WEIGHT
         ' // Change the font weight
         lfw.lfWeight = nValue
      CASE AFX_FONT_ITALIC
         ' // Change the italic flag
         lfw.lfItalic = CUBYTE(nValue)
      CASE AFX_FONT_UNDERLINE
         ' // Change the underline flag
         lfw.lfUnderline = CUBYTE(nValue)
      CASE AFX_FONT_STRIKEOUT
         ' // Change the strikeout flag
         lfw.lfStrikeOut = CUBYTE(nValue)
      CASE AFX_FONT_CHARSET
         ' // Change the charset
         lfw.lfCharset = CUBYTE(nValue)
      CASE ELSE
         EXIT FUNCTION
   END SELECT
   ' // Create a new font
   DIM hNewFont AS HFONT = CreateFontIndirectW(@lfw)
   IF hNewFont = NULL THEN EXIT FUNCTION
   ' // Select the new font and delete the old one
   DIM hDC AS HDC = GetDC(hwnd)
   DeleteObject(SelectObject(hDC, CAST(HGDIOBJ, hNewFont)))
   ReleaseDC(hwnd, hDC)
   SendMessageW(hwnd, WM_SETFONT, CAST(WPARAM, hNewFont), CTRUE)
   FUNCTION = hNewFont
END FUNCTION
' ========================================================================================

' ========================================================================================
' Modifies the height of the font used by a window of control.
' Parameters:
' - hwnd = Handle of the window or control.
' - nValue: The base is 100. To increase the font a 20% pass 120; to reduce it a 20% pass 80%.
' Return value: The handle of the new font on success, or NULL on failure.
' To get extended error information, call GetLastError.
' Remarks: The returned font must be destroyed with DeleteObject or the macro DeleteFont
' when no longer needed to prevent memory leaks.
' ========================================================================================
PRIVATE FUNCTION AfxModifyFontHeight (BYVAL hwnd AS HWND, BYVAL nValue AS LONG) AS HFONT
   FUNCTION = AfxModifyFontSettings(hwnd, AFX_FONT_HEIGHT, nValue)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates a logical font.
' Parameters:
' - wszFaceName = The typeface name.
' - lPointSize = The point size.
' - DPI = Dots per inch to calculate scaling. Default value = 96 (no scaling). If you pass -1
'   and the application is DPI aware, the DPI value used by the operating system will be used.
' - lWeight = The weight of the font in the range 0 through 1000. For example, 400 is normal
'      and 700 is bold. If this value is zero, a default weight is used.
'      The following values are defined for convenience.
'      FW_DONTCARE (0), FW_THIN (100), FW_EXTRALIGHT (200), FW_ULTRALIGHT (200), FW_LIGHT (300),
'      FW_NORMAL (400), FW_REGULAR (400), FW_MEDIUM (500), FW_SEMIBOLD (600), FW_DEMIBOLD (600),
'      FW_BOLD (700), FW_EXTRABOLD (800), FW_ULTRABOLD (800), FW_HEAVY (900), FW_BLACK (900)
' - bItalic = Italic flag. CTRUE or FALSE
' - bUnderline = Underline flag. CTRUE or FALSE
' - bStrikeOut = StrikeOut flag. CTRUE or FALSE
' - bCharset = Charset.
'      The following values are predefined: ANSI_CHARSET, BALTIC_CHARSET, CHINESEBIG5_CHARSET,
'      DEFAULT_CHARSET, EASTEUROPE_CHARSET, GB2312_CHARSET, GREEK_CHARSET, HANGUL_CHARSET,
'      MAC_CHARSET, OEM_CHARSET, RUSSIAN_CHARSET, SHIFTJIS_CHARSET, SYMBOL_CHARSET, TURKISH_CHARSET,
'      VIETNAMESE_CHARSET, JOHAB_CHARSET (Korean language edition of Windows), ARABIC_CHARSET and
'      HEBREW_CHARSET (Middle East language edition of Windows), THAI_CHARSET (Thai language
'      edition of Windows).
'      The OEM_CHARSET value specifies a character set that is operating-system dependent.
'      DEFAULT_CHARSET is set to a value based on the current system locale. For example, when
'      the system locale is English (United States), it is set as ANSI_CHARSET.
'      Fonts with other character sets may exist in the operating system. If an application uses
'      a font with an unknown character set, it should not attempt to translate or interpret
'      strings that are rendered with that font.
'      This parameter is important in the font mapping process. To ensure consistent results,
'      specify a specific character set. If you specify a typeface name in the lfFaceName member,
'      make sure that the lfCharSet value matches the character set of the typeface specified in lfFaceName.
' Return value: The handle of the font or NULL on failure.
' Remarks: The returned font must be destroyed with DeleteObject or the macro DeleteFont
' when no longer needed to prevent memory leaks.
' Usage examples:
'   hFont = AfxCreateFont("MS Sans Serif", 8, , FW_NORMAL, , , , DEFAULT_CHARSET)
'   hFont = AfxCreateFont("Courier New", 10, 96 , FW_BOLD, , , , DEFAULT_CHARSET)
'   hFont = AfxCreateFont("Marlett", 8, -1, FW_NORMAL, , , , SYMBOL_CHARSET)
' ========================================================================================
PRIVATE FUNCTION AfxCreateFont (BYREF wszFaceName AS WSTRING, BYVAL lPointSize AS LONG, BYVAL DPI AS LONG = 96, _
   BYVAL lWeight AS LONG = 0, BYVAL bItalic AS UBYTE = FALSE, BYVAL bUnderline AS UBYTE = FALSE, _
   BYVAL bStrikeOut AS UBYTE = FALSE, BYVAL bCharSet AS UBYTE = DEFAULT_CHARSET) AS HFONT

   DIM tlfw AS LOGFONTW
   DIM hDC AS HDC = GetDC(HWND_DESKTOP)

   ' // Font scaling
   IF DPI = -1 THEN DPI = GetDeviceCaps(hDC, LOGPIXELSX)
   IF DPI > 0 THEN lPointSize = (lPointSize * DPI) \ GetDeviceCaps(hDC, LOGPIXELSY)

   tlfw.lfHeight         = -MulDiv(lPointSize, .GetDeviceCaps(hDC, LOGPIXELSY), 72)  ' logical font height
   tlfw.lfWidth          =  0                                                        ' average character width
   tlfw.lfEscapement     =  0                                                        ' escapement
   tlfw.lfOrientation    =  0                                                        ' orientation angles
   tlfw.lfWeight         =  lWeight                                                  ' font weight
   tlfw.lfItalic         =  bItalic                                                  ' italic(CTRUE/FALSE)
   tlfw.lfUnderline      =  bUnderline                                               ' underline(CTRUE/FALSE)
   tlfw.lfStrikeOut      =  bStrikeOut                                               ' strikeout(CTRUE/FALSE)
   tlfw.lfCharSet        =  bCharset                                                 ' character set
   tlfw.lfOutPrecision   =  OUT_TT_PRECIS                                            ' output precision
   tlfw.lfClipPrecision  =  CLIP_DEFAULT_PRECIS                                      ' clipping precision
   tlfw.lfQuality        =  DEFAULT_QUALITY                                          ' output quality
   tlfw.lfPitchAndFamily =  FF_DONTCARE                                              ' pitch and family
   tlfw.lfFaceName       =  wszFaceName                                              ' typeface name

   ReleaseDC(HWND_DESKTOP, hDC)
   FUNCTION = CreateFontIndirectW(@tlfw)

END FUNCTION
' ========================================================================================


' ########################################################################################
'                        *** DIBs (Device Independent Bitmaps) ***
' ########################################################################################

' ========================================================================================
' Creates a DIB that applications can write to directly.
' Parameters:
' - hDC     = A handle to a device context.
' - nWidth  = The width of the bitmap, in pixels.
' - nHeight = The height of the bitmap, in pixels.
' - bpp     = Bits per pixel
' - ppvBits = A pointer to a variable that receives a pointer to the location of the
'             DIB bit values. Can be NULL.
' Return value: If the function succeeds, the return value is a handle to the newly
' created DIB, and *ppvBits points to the bitmap bit values.
' If the function fails, the return value is NULL, and *ppvBits is NULL.
' The function can fail if one or more of the input parameters is invalid.
' Remarks:
'   You must delete the returned bitmap handle with DeleteObject when no longer needed to
'   avoid memory leaks.
'   You cannot paste a DIB section from one application into another application.
'   AfxCreateDIBSection does not use the BITMAPINFOHEADER parameters biXPelsPerMeter or
'   biYPelsPerMeter and will not provide resolution information in the BITMAPINFO structure.
' Usage example:
'   DIM hdcWindow AS HDC, hbmp AS HBITMAP, pvBits AS ANY PTR
'   hdcWindow = GetWindowDC(hwnd)   ' where hwnd is the handle of the wanted window or control
'   hbmp = AfxCreateDIBSection(hdcWindow, 10, 10, @pvBits)
'   ReleaseDC(hwnd, hdcWindow)
' ========================================================================================
PRIVATE FUNCTION AfxCreateDIBSection (BYVAL hDC AS HDC, BYVAL nWidth AS DWORD, BYVAL nHeight AS DWORD, BYVAL bpp AS LONG = 0, BYVAL ppvBits AS ANY PTR PTR = NULL) AS HBITMAP
   DIM bi AS BITMAPINFO
   bi.bmiHeader.biSize = SIZEOF(bi.bmiHeader)
   bi.bmiHeader.biWidth = nWidth
   bi.bmiHeader.biHeight = nHeight
   bi.bmiHeader.biPlanes = 1
   bi.bmiHeader.biBitCount = IIF&(bpp <> 0, bpp, GetDeviceCaps(hDC, BITSPIXEL_))
   bi.bmiHeader.biCompression = BI_RGB
   FUNCTION = CreateDIBSection(hDC, @bi, DIB_RGB_COLORS, ppvBits, NULL, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Loads a DIB in memory and returns a pointer to it.
' Parameter:
' - pwszFileName = Path of the bitmap file.
' Reutn Value:
'   A pointer to the bitmap file header.
'   Release it with CoTaskMemFree when no longer needed.
' ========================================================================================
PRIVATE FUNCTION AfxDibLoadImage (BYVAL pwszFileName AS WSTRING PTR) AS BITMAPFILEHEADER PTR
   DIM bSuccess AS LONG, dwFileSize AS DWORD, dwHighSize AS DWORD, dwBytesRead AS DWORD
   DIM hFile AS HANDLE, pbmfh AS BITMAPFILEHEADER PTR
   hFile = CreateFileW(pwszFileName, GENERIC_READ, FILE_SHARE_READ, NULL, _
           OPEN_EXISTING, FILE_FLAG_SEQUENTIAL_SCAN, NULL)
   IF hFile = INVALID_HANDLE_VALUE THEN EXIT FUNCTION
   ' // Get the size of the file
   dwFileSize = GetFileSize(hFile, @dwHighSize)
   IF dwHighSize THEN
      CloseHandle(hFile)
      EXIT FUNCTION
   END IF
   ' // Read the contents of the file. Notice that pmfh has been cast as
   ' // BITMAPFILEHEADER PTR to be able to read the header.
   pbmfh = CoTaskMemAlloc(dwFileSize)
   bSuccess = ReadFile(hFile, pbmfh, dwFileSize, @dwBytesRead, NULL)
   ' // Check for "BM" (&H4D42, i.e. &H42 = "B", &H4D = "M", they are in reverse order)
   IF bSuccess = 0 OR dwBytesRead <> dwFileSize OR pbmfh->bfType <> &h4D42 THEN
      CoTaskMemFree(pbmfh)
      CloseHandle(hFile)
      EXIT FUNCTION
   END IF
   ' // Close the file handle and return a pointer to the data read
   CloseHandle(hFile)
   FUNCTION = pbmfh
END FUNCTION
' ========================================================================================

' ========================================================================================
' Saves a DIB to a file.
' - pwszFileName = Path of the file.
' - pbmfh = Pointer to the bitmap file header.
' Return Value
'   TRUE if the DIB has been saved successfully; FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION AfxDibSaveImage (BYVAL pwszFileName AS WSTRING PTR, BYVAL pbmfh AS BITMAPFILEHEADER PTR) AS BOOLEAN
   DIM bSuccess AS LONG, dwBytesWritten AS DWORD, hFile  AS HANDLE
   IF pbmfh = NULL THEN EXIT FUNCTION
   hFile = CreateFileW(pwszFileName, GENERIC_WRITE, 0, BYVAL NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL)
   IF hFile = INVALID_HANDLE_VALUE THEN EXIT FUNCTION
   bSuccess = WriteFile(hFile, BYVAL pbmfh, BYVAL pbmfh->bfSize, @dwBytesWritten, NULL)
   CloseHandle(hFile)
   IF bSuccess = 0 OR dwBytesWritten <> pbmfh->bfSize THEN
      DeleteFileW(pwszFileName)
      EXIT FUNCTION
   END IF
   FUNCTION = TRUE
END FUNCTION
' ========================================================================================


' ########################################################################################
'                                   *** BITMAPS ***
' ########################################################################################

' ========================================================================================
' Retrieves the width of a bitmap
' ========================================================================================
PRIVATE FUNCTION AfxGetBitmapWidth (BYVAL hBitmap AS HBITMAP) AS LONG
   DIM bm AS BITMAP
   IF hBitmap = NULL THEN EXIT FUNCTION
   IF GetObject(hBitmap, SIZEOF(bm), @bm) THEN FUNCTION = bm.bmWidth
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the height of a bitmap
' ========================================================================================
PRIVATE FUNCTION AfxGetBitmapHeight (BYVAL hBitmap AS HBITMAP) AS LONG
   DIM bm AS BITMAP
   IF hBitmap = NULL THEN EXIT FUNCTION
   IF GetObject(hBitmap, SIZEOF(bm), @bm) THEN FUNCTION = bm.bmHeight
END FUNCTION
' ========================================================================================

' ========================================================================================
' Captures the display and returns an handle to a bitmap.
' ========================================================================================
PRIVATE FUNCTION AfxCaptureDisplay () AS HBITMAP
   DIM hScreenDC AS HDC = CreateDCW("DISPLAY", "", "", NULL)
   IF hScreenDC = NULL THEN EXIT FUNCTION
   DIM cx AS LONG = GetDeviceCaps(hScreenDC, HORZRES)
   DIM cy AS LONG = GetDeviceCaps(hScreenDC, VERTRES)
   DIM hMemDC AS HDC = CreateCompatibleDC(hScreenDC)
   IF hMemDC = NULL THEN EXIT FUNCTION
   DIM hBitmap AS HBITMAP = CreateCompatibleBitmap(hScreenDC, cx, cy)
   IF hBitmap = NULL THEN EXIT FUNCTION
   DIM hBmpOld AS HBITMAP = SelectObject(hMemDC, hBitmap)
   ' // Note: CAPTUREBLT flag is required to capture layered windows
   BitBlt(hMemDC, 0, 0, cx, cy, hScreenDC, 0, 0, SRCCOPY OR CAPTUREBLT)
   SelectObject(hMemDC, hBmpOld)
   DeleteDC(hMemDC)
   DeleteDC(hScreenDC)
   FUNCTION = hBitmap
END FUNCTION
' ========================================================================================

' ========================================================================================
' Draws a bitmap.
' - hDC     : A handle to the destination device context.
' - xStart  : The x-coordinate, in logical units, of the upper-left corner of the destination rectangle.
' - yStart  : The y-coordinate, in logical units, of the upper-left corner of the destination rectangle.
' - hBitmap : Handle of the bitmap to draw.
' Return value: TRUE or FALSE.
' ========================================================================================
PRIVATE FUNCTION AfxDrawBitmap (BYVAL hDC AS HDC, BYVAL xStart AS LONG, BYVAL yStart AS LONG, BYVAL hBitmap AS HBITMAP) AS BOOLEAN
   DIM bm AS BITMAP, hMemDC AS HDC
   IF hDC = NULL OR hBitmap = NULL THEN EXIT FUNCTION
   hMemDC = CreateCompatibleDC(hDC)
   IF hMemDC = NULL THEN EXIT FUNCTION
   SelectObject hMemDC, hBitmap
   IF GetObject(hBitmap, SIZEOF(BITMAP), @bm) THEN
      BitBlt hDC, xStart, yStart, bm.bmWidth, bm.bmHeight, hMemDC, 0, 0, SRCCOPY
      FUNCTION = TRUE
   END IF
   DeleteDC hMemDC
END FUNCTION
' ========================================================================================

' ########################################################################################
'                                     *** FILES ***
' ########################################################################################

' ========================================================================================
' Searches a directory for a file or subdirectory with a name that matches a specific name
' (or partial name if wildcards are used).
' Parameter:
' - pwszFileSpec: The directory or path, and the file name, which can include wildcard
'   characters, for example, an asterisk (*) or a question mark (?).
'   This parameter should not be NULL, an invalid string (for example, an empty string or a
'   string that is missing the terminating null character), or end in a trailing backslash (\).
'   If the string ends with a wildcard, period (.), or directory name, the user must have
'   access permissions to the root and all subdirectories on the path. To extend the limit
'   of MAX_PATH wide characters to 32,767 wide characters, prepend "\\?\" to the path.
' Return value:
'   Returns TRUE if the specified file exists or FALSE otherwise.
' Remarks:
'   Prepending the string "\\?\" does not allow access to the root directory.
'   On network shares, you can use a pwszFileSpec in the form of the following:
'   "\\server\service\*". However, you cannot use a pwszFileSpec that points to the share
'   itself; for example, "\\server\service" is not valid.
'   To examine a directory that is not a root directory, use the path to that directory,
'   without a trailing backslash. For example, an argument of "C:\Windows" returns information
'   about the directory "C:\Windows", not about a directory or file in "C:\Windows".
'   To examine the files and directories in "C:\Windows", use an pwszFileSpec of "C:\Windows\*".
'   Be aware that some other thread or process could create or delete a file with this name
'   between the time you query for the result and the time you act on the information.
'   If this is a potential concern for your application, one possible solution is to use
'   the CreateFile function with CREATE_NEW (which fails if the file exists) or OPEN_EXISTING
'   (which fails if the file does not exist).
' ========================================================================================
PRIVATE FUNCTION AfxFileExists (BYVAL pwszFileSpec AS WSTRING PTR) AS BOOLEAN
   DIM fd AS WIN32_FIND_DATAW
   IF pwszFileSpec = NULL THEN EXIT FUNCTION
   DIM hFind AS HANDLE = FindFirstFileW(pwszFileSpec, @fd)
   IF hFind = INVALID_HANDLE_VALUE THEN EXIT FUNCTION
   FindClose hFind
   ' // Make sure that it is not a directory or a temporary file
   IF (fd.dwFileAttributes AND FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY AND _
      (fd.dwFileAttributes AND FILE_ATTRIBUTE_TEMPORARY) <> FILE_ATTRIBUTE_TEMPORARY THEN
      FUNCTION = TRUE
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Searches a directory for a file or subdirectory with a name that matches a specific name
' (or partial name if wildcards are used).
' Parameter:
' - pwszFileSpec: The directory or path, and the file name, which can include wildcard
'   characters, for example, an asterisk (*) or a question mark (?).
'   This parameter should not be NULL, an invalid string (for example, an empty string or a
'   string that is missing the terminating null character), or end in a trailing backslash (\).
'   If the string ends with a wildcard, period (.), or directory name, the user must have
'   access permissions to the root and all subdirectories on the path. To extend the limit
'   of MAX_PATH wide characters to 32,767 wide characters, prepend "\\?\" to the path.
' Return value:
'   Returns TRUE if the specified folder exists or FALSE otherwise.
' Remarks
'   Prepending the string "\\?\" does not allow access to the root directory.
'   On network shares, you can use an btrFileName in the form of the following: "\\server\service\*".
'   However, you cannot use an btrFileName that points to the share itself; for example,
'   "\\server\service" is not valid.
'   To examine a directory that is not a root directory, use the path to that directory,
'   without a trailing backslash. For example, an argument of "C:\Windows" returns information
'   about the directory "C:\Windows", not about a directory or file in "C:\Windows".
'   To examine the files and directories in "C:\Windows", use an btrFileName of "C:\Windows\*".
'   Be aware that some other thread or process could create or delete a file with this name
'   between the time you query for the result and the time you act on the information.
'   If this is a potential concern for your application, one possible solution is to use
'   the CreateFile function with CREATE_NEW (which fails if the file exists) or OPEN_EXISTING
'   (which fails if the file does not exist).
' ========================================================================================
PRIVATE FUNCTION AfxFolderExists (BYVAL pwszFileSpec AS WSTRING PTR) AS BOOLEAN
   DIM fd AS WIN32_FIND_DATAW
   IF pwszFileSpec = NULL THEN EXIT FUNCTION
   DIM hFind AS HANDLE = FindFirstFileW(pwszFileSpec, @fd)
   IF hFind = INVALID_HANDLE_VALUE THEN EXIT FUNCTION
   FindClose hFind
   ' // Make sure that it is a directory
   IF (fd.dwFileAttributes AND FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY THEN FUNCTION = TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns TRUE if the specified path is a file; FALSE, otherwise.
' To extend the limit from MAX_PATH to 32,767 wide characters, prepend "\\?\" to the path.
' ========================================================================================
PRIVATE FUNCTION AfxIsFile (BYREF wszFileSpec AS WSTRING) AS BOOLEAN
   DIM dwAttributes AS DWORD = GetFileAttributesW(wszFileSpec)
   IF dwAttributes = INVALID_FILE_ATTRIBUTES THEN RETURN FALSE
   RETURN ((dwAttributes AND FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns TRUE if the specified path is a directory; FALSE, otherwise.
' To extend the limit from MAX_PATH to 32,767 wide characters, prepend "\\?\" to the path.
' ========================================================================================
PRIVATE FUNCTION AfxIsFolder (BYREF wszFileSpec AS WSTRING) AS BOOLEAN
   DIM dwAttributes AS DWORD = GetFileAttributesW(wszFileSpec)
   IF dwAttributes = INVALID_FILE_ATTRIBUTES THEN RETURN FALSE
   RETURN ((dwAttributes AND FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns TRUE if the specified path is a system file; FALSE, otherwise.
' To extend the limit from MAX_PATH to 32,767 wide characters, prepend "\\?\" to the path.
' ========================================================================================
PRIVATE FUNCTION AfxIsSystemFile (BYREF wszFileSpec AS WSTRING) AS BOOLEAN
   DIM dwAttributes AS DWORD = GetFileAttributesW(wszFileSpec)
   IF dwAttributes = INVALID_FILE_ATTRIBUTES THEN RETURN FALSE
   RETURN ((dwAttributes AND FILE_ATTRIBUTE_SYSTEM) = FILE_ATTRIBUTE_SYSTEM)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns TRUE if the specified path is a hidden file; FALSE, otherwise.
' To extend the limit from MAX_PATH to 32,767 wide characters, prepend "\\?\" to the path.
' ========================================================================================
PRIVATE FUNCTION AfxIsHiddenFile (BYREF wszFileSpec AS WSTRING) AS BOOLEAN
   DIM dwAttributes AS DWORD = GetFileAttributesW(wszFileSpec)
   IF dwAttributes = INVALID_FILE_ATTRIBUTES THEN RETURN FALSE
   RETURN ((dwAttributes AND FILE_ATTRIBUTE_HIDDEN) = FILE_ATTRIBUTE_HIDDEN)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns TRUE if the specified path is a read only file; FALSE, otherwise.
' To extend the limit from MAX_PATH to 32,767 wide characters, prepend "\\?\" to the path.
' ========================================================================================
PRIVATE FUNCTION AfxIsReadOnlyFile (BYREF wszFileSpec AS WSTRING) AS BOOLEAN
   DIM dwAttributes AS DWORD = GetFileAttributesW(wszFileSpec)
   IF dwAttributes = INVALID_FILE_ATTRIBUTES THEN RETURN FALSE
   RETURN ((dwAttributes AND FILE_ATTRIBUTE_READONLY) = FILE_ATTRIBUTE_READONLY)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns TRUE if the specified path is a temporary file; FALSE, otherwise.
' To extend the limit from MAX_PATH to 32,767 wide characters, prepend "\\?\" to the path.
' ========================================================================================
PRIVATE FUNCTION AfxIsTemporaryFile (BYREF wszFileSpec AS WSTRING) AS BOOLEAN
   DIM dwAttributes AS DWORD = GetFileAttributesW(wszFileSpec)
   IF dwAttributes = INVALID_FILE_ATTRIBUTES THEN RETURN FALSE
   RETURN ((dwAttributes AND FILE_ATTRIBUTE_TEMPORARY) = FILE_ATTRIBUTE_TEMPORARY)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns TRUE if the specified path is a normal file; FALSE, otherwise.
' To extend the limit from MAX_PATH to 32,767 wide characters, prepend "\\?\" to the path.
' ========================================================================================
PRIVATE FUNCTION AfxIsNormalFile (BYREF wszFileSpec AS WSTRING) AS BOOLEAN
   DIM dwAttributes AS DWORD = GetFileAttributesW(wszFileSpec)
   IF dwAttributes = INVALID_FILE_ATTRIBUTES THEN RETURN FALSE
   RETURN ((dwAttributes AND FILE_ATTRIBUTE_NORMAL) = FILE_ATTRIBUTE_NORMAL)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns TRUE if the specified path is a compressed file or diectory; FALSE, otherwise.
' To extend the limit from MAX_PATH to 32,767 wide characters, prepend "\\?\" to the path.
' ========================================================================================
PRIVATE FUNCTION AfxIsCompressedFile (BYREF wszFileSpec AS WSTRING) AS BOOLEAN
   DIM dwAttributes AS DWORD = GetFileAttributesW(wszFileSpec)
   IF dwAttributes = INVALID_FILE_ATTRIBUTES THEN RETURN FALSE
   RETURN ((dwAttributes AND FILE_ATTRIBUTE_COMPRESSED) = FILE_ATTRIBUTE_COMPRESSED)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns TRUE if the specified path is an encrypted file or diectory; FALSE, otherwise.
' To extend the limit from MAX_PATH to 32,767 wide characters, prepend "\\?\" to the path.
' ========================================================================================
PRIVATE FUNCTION AfxIsEncryptedFile (BYREF wszFileSpec AS WSTRING) AS BOOLEAN
   DIM dwAttributes AS DWORD = GetFileAttributesW(wszFileSpec)
   IF dwAttributes = INVALID_FILE_ATTRIBUTES THEN RETURN FALSE
   RETURN ((dwAttributes AND FILE_ATTRIBUTE_ENCRYPTED) = FILE_ATTRIBUTE_ENCRYPTED)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns TRUE if the specified file or directory is not to be indexed by the content
' indexing service; FALSE, otherwise.
' To extend the limit from MAX_PATH to 32,767 wide characters, prepend "\\?\" to the path.
' ========================================================================================
PRIVATE FUNCTION AfxIsNotContentIndexedFile (BYREF wszFileSpec AS WSTRING) AS BOOLEAN
   DIM dwAttributes AS DWORD = GetFileAttributesW(wszFileSpec)
   IF dwAttributes = INVALID_FILE_ATTRIBUTES THEN RETURN FALSE
   RETURN ((dwAttributes AND FILE_ATTRIBUTE_NOT_CONTENT_INDEXED) = FILE_ATTRIBUTE_NOT_CONTENT_INDEXED)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns TRUE if the specified file is not available immediately; FALSE, otherwise.
' To extend the limit from MAX_PATH to 32,767 wide characters, prepend "\\?\" to the path.
' ========================================================================================
PRIVATE FUNCTION AfxIsOffLineFile (BYREF wszFileSpec AS WSTRING) AS BOOLEAN
   DIM dwAttributes AS DWORD = GetFileAttributesW(wszFileSpec)
   IF dwAttributes = INVALID_FILE_ATTRIBUTES THEN RETURN FALSE
   RETURN ((dwAttributes AND FILE_ATTRIBUTE_OFFLINE) = FILE_ATTRIBUTE_OFFLINE)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns TRUE if the specified path is a file or directory that has an associated reparse
' point, or a file that is a symbolic link; FALSE, otherwise.
' To extend the limit from MAX_PATH to 32,767 wide characters, prepend "\\?\" to the path.
' ========================================================================================
PRIVATE FUNCTION AfxIsReparsePointFile (BYREF wszFileSpec AS WSTRING) AS BOOLEAN
   DIM dwAttributes AS DWORD = GetFileAttributesW(wszFileSpec)
   IF dwAttributes = INVALID_FILE_ATTRIBUTES THEN RETURN FALSE
   RETURN ((dwAttributes AND FILE_ATTRIBUTE_REPARSE_POINT) = FILE_ATTRIBUTE_REPARSE_POINT)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns TRUE if the specified path is a sparse file; FALSE, otherwise.
' To extend the limit from MAX_PATH to 32,767 wide characters, prepend "\\?\" to the path.
' ========================================================================================
PRIVATE FUNCTION AfxIsSparseFile (BYREF wszFileSpec AS WSTRING) AS BOOLEAN
   DIM dwAttributes AS DWORD = GetFileAttributesW(wszFileSpec)
   IF dwAttributes = INVALID_FILE_ATTRIBUTES THEN RETURN FALSE
   RETURN ((dwAttributes AND FILE_ATTRIBUTE_SPARSE_FILE) = FILE_ATTRIBUTE_SPARSE_FILE)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the size in bytes of the specified file. Don't use it with folders.
' ========================================================================================
PRIVATE FUNCTION AfxGetFileSize (BYREF wszFileSpec AS WSTRING) AS ULONGLONG
   DIM fd AS WIN32_FIND_DATAW
   DIM hFind AS HANDLE = FindFirstFileW(wszFileSpec, @fd)
   IF hFind = INVALID_HANDLE_VALUE THEN RETURN 0
   FindClose hFind
   IF (fd.dwFileAttributes AND FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY AND _
      (fd.dwFileAttributes AND FILE_ATTRIBUTE_TEMPORARY) <> FILE_ATTRIBUTE_TEMPORARY THEN
      DIM ullSize AS ULONGLONG = (fd.nFileSizeHigh * (&hFFFFFFFF + 1)) + fd.nFileSizeLow
      RETURN ullSize
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the time the file was created.
' - wszFileSpec: The directory or path, and the file name, which can include wildcard characters,
'   for example, an asterisk (*) or a question mark (?).
'   This parameter should not be NULL, an invalid string (for example, an empty string or a
'   string that is missing the terminating null character), or end in a trailing backslash (\).
'   If the string ends with a wildcard, period (.), or directory name, the user must have access
'   permissions to the root and all subdirectories on the path.
'   To extend the limit from MAX_PATH to 32,767 wide characters, prepend "\\?\" to the path.
' - bUTC: Pass FALSE if you want to get the time in local time (the NTFS file system stores time
'   values in UTC format, so they are not affected by changes in time zone or daylight saving time).
'   FileTimeToLocalFileTime uses the current settings for the time zone and daylight saving time.
'   Therefore, if it is daylight saving time, it takes daylight saving time into account, even
'   if the file time you are converting is in standard time.
' Usage: AfxFileTimeToDateStr(AfxGetFileCreationTime("C:\Tests\test.bas", FALSE), "dd/MM/yyyy")
' ========================================================================================
PRIVATE FUNCTION AfxGetFileCreationTime (BYREF wszFileSpec AS WSTRING, BYVAL bUTC AS BOOLEAN = TRUE) AS FILETIME
   DIM fd AS WIN32_FIND_DATAW
   DIM hFind AS HANDLE = FindFirstFileW(wszFileSpec, @fd)
   IF hFind <> INVALID_HANDLE_VALUE THEN
      FindClose hFind
      IF bUTC = TRUE THEN
         RETURN fd.ftCreationTime
      ELSE
         DIM FT AS FILETIME
         FileTimeToLocalFileTime(@fd.ftCreationTime, @FT)
         RETURN FT
      END IF
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the time the file was last accessed. The NTFS file system delays updates to the
' last access time for a file by up to 1 hour after the last access.
' ========================================================================================
PRIVATE FUNCTION AfxGetFileLastAccessTime (BYREF wszFileSpec AS WSTRING, BYVAL bUTC AS BOOLEAN = TRUE) AS FILETIME
   DIM fd AS WIN32_FIND_DATAW
   DIM hFind AS HANDLE = FindFirstFileW(wszFileSpec, @fd)
   IF hFind <> INVALID_HANDLE_VALUE THEN
      FindClose hFind
      IF bUTC = TRUE THEN
         RETURN fd.ftLastAccessTime
      ELSE
         DIM FT AS FILETIME
         FileTimeToLocalFileTime(@fd.ftLastAccessTime, @FT)
         RETURN FT
      END IF
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the time the file was last written to, truncated, or overwritten.
' When writing to a file, the last write time is not fully updated until all handles that
' are used for writing are closed.
' ========================================================================================
PRIVATE FUNCTION AfxGetFileLastWriteTime (BYREF wszFileSpec AS WSTRING, BYVAL bUTC AS BOOLEAN = TRUE) AS FILETIME
   DIM fd AS WIN32_FIND_DATAW
   DIM hFind AS HANDLE = FindFirstFileW(wszFileSpec, @fd)
   IF hFind <> INVALID_HANDLE_VALUE THEN
      FindClose hFind
      IF bUTC = TRUE THEN
         RETURN fd.ftLastWriteTime
      ELSE
         DIM FT AS FILETIME
         FileTimeToLocalFileTime(@fd.ftLastWriteTime, @FT)
         RETURN FT
      END IF
   END IF
END FUNCTION
' ========================================================================================
