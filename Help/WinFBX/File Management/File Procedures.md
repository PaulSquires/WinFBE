# File Procedures

Assorted file, folder and path procedures.

**Include files**: AfxWin.inc, AfxPath.inc

# Dialogs

| Name       | Description |
| ---------- | ----------- |
| [AfxBrowseForFolder](#AfxBrowseForFolder) | Displays a dialog box that enables the user to select a folder. |
| [AfxOpenFileDialog](#AfxOpenFileDialog) | Creates an Open dialog box that lets the user specify the drive, directory, and the name of a file or set of files to be opened. |
| [AfxSaveFileDialog](#AfxSaveFileDialog) | Creates a Save dialog box that lets the user specify the drive, directory, and name of a file to save. |

# File and Folder Procedures

| Name       | Description |
| ---------- | ----------- |
| [AfxChDir](#AfxChDir) | Changes the current directory for the current process. |
| [AfxCopyFile](#AfxCopyFile) | Copies an existing file to a new file. |
| [AfxCreateDirectory](#AfxMakeDir) | Creates a new directory. |
| [AfxCurDir](#AfxCurDir) | Retrieves the current directory for the current process. |
| [AfxDeleteFile](#AfxDeleteFile) | Deletes the specified file. |
| [AfxExePath](#AfxExePath) | Returns the path of the program which is currently executing. The path has not a trailing backslash except if it is a drive, e.g. C:\. |
| [AfxFileCopy](#AfxCopyFile) | Copies an existing file to a new file. |
| [AfxFileDateTime](#AfxFileDateTime) | Returns the file's last modified date and time as Date Serial. |
| [AfxFileExists](#AfxFileExists) | Searches a directory for a file or subdirectory with a name that matches a specific name (or partial name if wildcards are used). |
| [AfxFileReadAllLines](#AfxFileReadAllLines) | Reads all the lines of the specified file into a safe array. |
| [AfxFileScan](#AfxFileScan) | Scans a text file and returns the number of occurrences of the specified delimiter. |
| [AfxFolderExists](#AfxFolderExists) | Searches a directory for a file or subdirectory with a name that matches a specific name (or partial name if wildcards are used). |
| [AfxGetCurDir](#AfxCurDir) | Retrieves the current directory for the current process. |
| [AfxGetCurrentDirectory](#AfxCurDir) | Retrieves the current directory for the current process. |
| [AfxGetDriveType](#AfxGetDriveType) | Determines whether a disk drive is a removable, fixed, CD-ROM, RAM disk, or network drive. |
| [AfxGetExeFileExt](#AfxGetExeFileExt) | Parses a path/filename and returns the extension portion of the path/file name. That is the last period (.) in the string plus the text to the right of it. |
| [AfxGetExeFileName](#AfxGetExeFileName) | Returns the file name of the program which is currently executing. |
| [AfxGetExeFileNameX](#AfxGetExeFileNameX) | Returns the file name and extension of the program which is currently executing. |
| [AfxGetExeFullPath](#AfxGetExeFullPath) | Returns the complete drive, path, file name, and extension of the program which is currently executing. |
| [AfxGetExePath](#AfxExePath) | Returns the path of the program which is currently executing. The path has not a trailing backslash except if it is a drive, e.g. C:\. |
| [AfxGetExePathName](#AfxGetExePathName) | Returns the path of the program which is currently executing. The path has a trailing backslash. |
| [AfxGetFileCreationTime](#AfxGetFileCreationTime) | Returns the time the file was created, in FILETIME format. |
| [AfxGetFileExt](#AfxGetFileExt) | Parses a path/filename and returns the extension portion of the path/file name. That is the last period (.) in the string plus the text to the right of it. |
| [AfxGetFileLastAccessTime](#AfxGetFileLastAccessTime) | Returns the time the file was last accessed, in FILETIME format. |
| [AfxGetFileLastWriteTime](#AfxGetFileLastWriteTime) | Returns the time the file was last written to, truncated, or overwritten, in FILETIME format. |
| [AfxFileLen](#AfxGetFileSize) | Returns the size in bytes of the specified file. |
| [AfxGetFileName](#AfxGetFileName) | Parses a path/filename and returns the file name portion. That is the text to the right of the last backslash (\) or colon (:), ending just before the last period (.). |
| [AfxGetFileNameX](#AfxGetFileNameX) | Parses a path/filename and returns the file name and extension portion. That is the text to the right of the last backslash (\) or colon (:). |
| [AfxGetFileSize](#AfxGetFileSize) | Returns the size in bytes of the specified file. |
| [AfxGetFileVersion](#AfxGetFileVersion) | Retrieves the version of the specified file multiplied by 100, e.g. 601 for version 6.01. |
| [AfxGetFolderName](#AfxGetFolderName) | Returns a string containing the name of the folder for a specified path, i.e. the path minus the file name. |
| [AfxGetKnowFolderPath](#AfxGetKnowFolderPath) | Retrieves the path of an special folder. Requires Windows Vista/Windows 7 or superior. |
| [AfxGetLongPathName](#AfxGetLongPathName) | Retrieves the short path form of the specified path. |
| [AfxGetPathName](#AfxGetPathName) | Parses a path/filename and returns the path portion. That is the text up to and including the last backslash (\) or colon (:). |
| [AfxGetShortPathName](#AfxGetShortPathName) | Retrieves the short path form of the specified path. |
| [AfxGetSpecialFolderLocation](#AfxGetSpecialFolderLocation) | Retrieves the path of an special folder. |
| [AfxGetSystemDllPath](#AfxGetSystemDllPath) | Retrieves the fully qualified path for the file that contains the specified module. |
| [AfxGetWinDir](#AfxGetWinDir) | Retrieves the path of the Windows directory. |
| [AfxIsCompressedFile](#AfxIsCompressedFile) | Returns True if the specified file or directory is compressed; False if it is not. |
| [AfxIsEncryptedFile](#AfxIsEncryptedFile) | Returns True if the specified file or directory is encrypted; False if it is not. |
| [AfxIsFolder](#AfxIsFolder) | Returns True if the specified path is a folder; False if it is not. |
| [AfxIsHiddenFile](#AfxIsHiddenFile) | Returns True if the specified path is a hidden file or directory; False if it is not. |
| [AfxIsNormalFile](#AfxIsNormalFile) | Returns True if the specified path is a normal file (a file that does not have other attributes set); False if it is not. |
| [AfxIsNotContentIndexedFile](#AfxIsNotContentIndexedFile) | Returns TRUE if the specified file or directory is not to be indexed by the content indexing service; FALSE, otherwise. |
| [AfxIsOfflineFile](#AfxIsOfflineFile) | Returns TRUE if the specified file file is not available immediately; FALSE, otherwise. |
| [AfxIsReadOnlyFile](#AfxIsReadOnlyFile) | Returns True if the specified path is a read only file; False if it is not. |
| [AfxIsReparsePointFile](#AfxIsReparsePointFile) | Returns TRUE if the specified path is a file or directory that has an associated reparse point, or a file that is a symbolic link.; FALSE, otherwise. |
| [AfxIsSparseFile](#AfxIsSparseFile) | Returns TRUE if the specified path is a sparse file; FALSE, otherwise. |
| [AfxIsSystemFile](#AfxIsSystemFile) | Returns True if the specified path is a system file; False if it is not. |
| [AfxIsTemporaryFile](#AfxIsTemporaryFile) | Returns True if the specified path is a temporary file; False if it is not. |
| [AfxKill](#AfxDeleteFile) | Deletes the specified file. |
| [AfxMakeDir](#AfxMakeDir) | Creates a new directory. |
| [AfxMkDir](#AfxMakeDir) | Creates a new directory. |
| [AfxMoveFile](#AfxMoveFile) | Moves an existing file or a directory, including its children. |
| [AfxName](#AfxMoveFile) | Moves an existing file or a directory, including its children. |
| [AfxRemoveDirectory](#AfxRemoveDir) | Deletes an existing empty directory. |
| [AfxRemoveDir](#AfxRemoveDir) | Deletes an existing empty directory. |
| [AfxRenameFile](#AfxMoveFile) | Moves an existing file or a directory, including its children. |
| [AfxRmDir](#AfxRemoveDir) | Deletes an existing empty directory. |
| [AfxSaveTempFile](#AfxSaveTempFile) | Saves the contents of a string buffer in a temporary file. |
| [AfxSetCurDir](#AfxChDir) | Changes the current directory for the current process. |
| [AfxSetCurrentDirectory](#AfxChDir) | Changes the current directory for the current process. |

## Path and Url Procedures

| Name       | Description |
| ---------- | ----------- |
| [AfxPathAddBackSlash](#AfxPathAddBackSlash) | Adds a backslash to the end of a string to create the correct syntax for a path. If the source path already has a trailing backslash, no backslash will be added. |
| [AfxPathAddExtension](#AfxPathAddExtension) | Adds a file name extension to a path string. |
| [AfxPathAppend](#AfxPathAppend) | Appends one path to the end of another. |
| [AfxPathBuildRoot](#AfxPathBuildRoot) | Creates a root path from a given drive number. |
| [AfxPathCanonicalize](#AfxPathCanonicalize) | Removes elements of a file path according to special strings inserted into that path. |
| [AfxPathCombine](#AfxPathCombine) | Concatenates two strings that represent properly formed paths into one path; also concatenates any relative path elements. |
| [AfxPathCommonPrefix](#AfxPathCommonPrefix) | Compares two paths to determine if they share a common prefix. A prefix is one of these types: "C:\\", ".", "..", "..\\". |
| [AfxPathCompactPath](#AfxPathCompactPath) | Truncates a file path to fit within a given pixel width by replacing path components with ellipses. |
| [AfxPathCompactPathEx](#AfxPathCompactPathEx) | Truncates a path to fit within a certain number of characters by replacing path components with ellipses. |
| [AfxPathCreateFromUrl](#AfxPathCreateFromUrl) | Converts a file URL to a Microsoft MS-DOS path. |
| [AfxPathFileExists](#AfxPathFileExists) | Determines whether a path to a file system object such as a file or directory is valid. |
| [AfxPathFindExtension](#AfxPathFindExtension) | Searches a path for an extension. |
| [AfxPathFindFileName](#AfxPathFindFileName) | Searches a path for a file name. |
| [AfxPathFindNextComponent](#AfxPathFindNextComponent) | Parses a path and returns the portion of that path that follows the first backslash. |
| [AfxPathFindOnPath](#AfxPathFindOnPath) | Searches for a file. |
| [AfxPathFindSuffixArray](#AfxPathFindSuffixArray) | Determines whether a given file name has one of a list of suffixes. |
| [AfxPathGetArgs](#AfxPathGetArgs) | Finds the command line arguments within a given path. |
| [AfxPathGetCharType](#AfxPathGetCharType) | Determines the type of character in relation to a path. |
| [AfxPathGetDriveNumber](#AfxPathGetDriveNumber) | Searches a path for a drive letter within the range of 'A' to 'Z' and returns the corresponding drive number. |
| [AfxPathIsContentType](#AfxPathIsContentType) | Determines if a file's registered content type matches the specified content type. |
| [AfxPathIsDirectory](#AfxPathIsDirectory) | Verifies that a path is a valid directory. |
| [AfxPathIsDirectoryEmpty](#AfxPathIsDirectoryEmpty) | Determines whether a specified path is an empty directory. |
| [AfxPathIsFileSpec](#AfxPathIsFileSpec) | Searches a path for any path-delimiting characters (for example, ':' or '\' ). If there are no path-delimiting characters present, the path is considered to be a File Spec path. |
| [AfxPathIsHTMLFile](#AfxPathIsHTMLFile) | Determines if a file is an HTML file. The determination is made based on the content type that is registered for the file's extension. |
| [AfxPathIsLFNFileSpec](#AfxPathIsLFNFileSpec) | Determines whether a file name is in long format. |
| [AfxPathIsNetworkPath](#AfxPathIsNetworkPath) | Determines whether a path string represents a network resource. |
| [AfxPathIsPrefix](#AfxPathIsPrefix) | Searches a path to determine if it contains a valid prefix of the type passed by wszPrefix. A prefix is one of these types: "C:\\", ".", "..", "..\\". |
| [AfxPathIsRelative](#AfxPathIsRelative) | Searches a path and determines if it is relative. |
| [AfxPathIsRoot](#AfxPathIsRoot) | Parses a path to determine if it is a directory root. |
| [AfxPathIsSameRoot](#AfxPathIsSameRoot) | Compares two paths to determine if they have a common root component. |
| [AfxPathIsSystemFolder](#AfxPathIsSystemFolder) | Determines if an existing folder contains the attributes that make it a system folder. Alternately, this function indicates if certain attributes qualify a folder to be a system folder. |
| [AfxPathIsUNC](#AfxPathIsUNC) | Determines if the string is a valid Universal Naming Convention (UNC) for a server and share path. |
| [AfxPathIsUNCServer](#AfxPathIsUNCServer) | Determines if a string is a valid Universal Naming Convention (UNC) for a server path only. |
| [AfxPathIsUNCServerShare](#AfxPathIsUNCServerShare) | Determines if a string is a valid Universal Naming Convention (UNC) share path, \\\\*server*\\*share*. |
| [AfxPathIsURL](#AfxPathIsURL) | Tests a given string to determine if it conforms to a valid URL format. |
| [AfxPathMakePretty](#AfxPathMakePretty) | Converts a path to all lowercase characters to give the path a consistent appearance. |
| [AfxPathMakeSystemFolder](#AfxPathMakeSystemFolder) | Gives an existing folder the proper attributes to become a system folder. |
| [AfxPathMatchSpec](#AfxPathMatchSpec) | Searches a string using a Microsoft MS-DOS wild card match type. |
| [AfxPathMatchSpecEx](#AfxPathMatchSpecEx) | Searches a path to determine whether it contains a file of a specified file type extension. |
| [AfxPathParseIconLocation](#AfxPathParseIconLocation) | Parses a file location string that contains a file location and icon index, and returns separate values. |
| [AfxPathQuoteSpaces](#AfxPathQuoteSpaces) | Searches a path for spaces. If spaces are found, the entire path is enclosed in quotation marks. |
| [AfxPathRelativePathTo](#AfxPathRelativePathTo) | Creates a relative path from one file or folder to another. |
| [AfxPathRemoveArgs](#AfxPathRemoveArgs) | Removes any arguments from a given path. |
| [AfxPathRemoveBackslash](#AfxPathRemoveBackslash) | Removes the trailing backslash from a given path. |
| [AfxPathRemoveBlanks](#AfxPathRemoveBlanks) | Removes all leading and trailing spaces from a string. |
| [AfxPathRemoveExtension](#AfxPathRemoveExtension) | Removes the file name extension from a path, if one is present. |
| [AfxPathRemoveFileSpec](#AfxPathRemoveFileSpec) | Removes the trailing file name and backslash from a path, if they are present. |
| [AfxPathRenameExtension](#AfxPathRenameExtension) | Replaces the extension of a file name with a new extension. If the file name does not contain an extension, the extension will be attached to the end of the string. |
| [AfxPathSearchAndQualify](#AfxPathSearchAndQualify) | Determines if a given path is correctly formatted and fully qualified. |
| [AfxPathSetDlgItemPath](#AfxPathSetDlgItemPath) | Sets the text of a child control in a window or dialog box, using **AfxCompactPath** to ensure the path fits in the control. |
| [AfxPathSkipRoot](#AfxPathSkipRoot) | Parses a path, ignoring the drive letter or Universal Naming Convention (UNC) server/share path elements. |
| [AfxPathStripPath](#AfxPathStripPath) | Removes the path portion of a fully qualified path and file. |
| [AfxPathStripToRoot](#AfxPathStripToRoot) | Removes all parts of the path except for the root information. |
| [AfxPathUndecorate](#AfxPathUndecorate) | Removes the decoration from a path string. |
| [AfxPathUnExpandEnvStrings](#AfxPathUnExpandEnvStrings) | Replaces certain folder names in a fully-qualified path with their associated environment string. |
| [AfxPathUnmakeSystemFolder](#AfxPathUnmakeSystemFolder) | Removes the attributes from a folder that make it a system folder. This folder must actually exist in the file system. |
| [AfxPathUnquoteSpaces](#AfxPathUnquoteSpaces) | Removes quotes from the beginning and end of a path. |
| [AfxUrlApplyScheme](#AfxUrlApplyScheme) | Determines a scheme for a specified URL string, and returns a string with an appropriate prefix. |
| [AfxUrlCanonicalize](#AfxUrlCanonicalize) | Converts a URL string into canonical form. |
| [AfxUrlCombine](#AfxUrlCombine) | When provided with a relative URL and its base, returns a URL in canonical form. |
| [AfxUrlCompare](#AfxUrlCompare) | Makes a case-sensitive comparison of two URL strings. |
| [AfxUrlCreateFromPath](#AfxUrlCreateFromPath) | Converts a Microsoft MS-DOS path to a canonicalized URL. |
| [AfxUrlEscape](#AfxUrlEscape) | Converts characters in a URL that might be altered during transport across the Internet ("unsafe" characters) into their corresponding escape sequences. |
| [AfxUrlEscapeSpaces](#AfxUrlEscapeSpaces) | Converts space characters into their corresponding escape sequence. |
| [AfxUrlFixup](#AfxUrlFixup) | Attempts to correct a URL whose protocol identifier is incorrect. For example, htttp will be changed to http. |
| [AfxUrlGetLocation](#AfxUrlGetLocation) | Retrieves the location from a URL. |
| [AfxUrlGetPart](#AfxUrlGetPart) | Accepts a URL string and returns a specified part of that URL. |
| [AfxUrlHash](#AfxUrlHash) | Hashes a URL string. |
| [AfxUrlIs](#AfxUrlIs) | Tests whether or not a URL is a specified type. |
| [AfxUrlIsFileUrl](#AfxUrlIsFileUrl) | Tests a URL to determine if it is a file URL. |
| [AfxUrlIsNoHistory](#AfxUrlIsNoHistory) | Returns whether a URL is a URL that browsers typically do not include in navigation history. |
| [AfxUrlIsOpaque](#AfxUrlIsOpaque) | Returns whether a URL is opaque. |
| [AfxUrlUnescape](#AfxUrlUnescape) | Converts escape sequences back into ordinary characters. |
| [AfxUrlUnescapeInPlace](#AfxUrlUnescapeInPlace) | Converts escape sequences back into ordinary characters and overwrites the original string. |

# <a name="AfxBrowseForFolder"></a>AfxBrowseForFolder

Displays a dialog box that enables the user to select a folder.

```
FUNCTION AfxBrowseForFolder (BYVAL hwnd AS HWND, BYVAL pwszTitle AS WSTRING PTR = NULL, _
   BYVAL pwszStartFolder AS WSTRING PTR = NULL, BYVAL nFlags AS LONG = 0) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *hwnd* | The handle to the parent window of the dialog box. This value can be zero. |
| *pwszTitle* | Optional. A string value that represents the title displayed inside the Browse dialog box. |
| *pwszStartFolder* | Optional.  The initial folder that the dialog will show. |
| *nFlags* | Optional. A LONG value that contains the options for the method. This can be zero or a combination of the values listed under the *ulFlags* member of the **BROWSEINFO** structure. |

| Flag       | Description |
| ---------- | ----------- |
| BIF_RETURNONLYFSDIRS | Only return file system directories. If the user selects folders that are not part of the file system, the OK button is grayed. |
| BIF_DONTGOBELOWDOMAIN | Do not include network folders below the domain level in the dialog box's tree view control. |
| BIF_STATUSTEXT | Include a status area in the dialog box. The callback function can set the status text by sending messages to the dialog box. This flag is not supported when BIF_NEWDIALOGSTYLE is specified. |
| BIF_RETURNFSANCESTORS | Only return file system ancestors. An ancestor is a subfolder that is beneath the root folder in the namespace hierarchy. If the user selects an ancestor of the root folder that is not part of the file system, the **OK** button is grayed. |
| BIF_EDITBOX | Version 4.71. Include an edit control in the browse dialog box that allows the user to type the name of an item. |
| BIF_NEWDIALOGSTYLE | Version 5.0. Use the new user interface. Setting this flag provides the user with a larger dialog box that can be resized. The dialog box has several new capabilities, including: drag-and-drop capability within the dialog box, reordering, shortcut menus, new folders, delete, and other shortcut menu commands. |
| BIF_USENEWUI | Version 5.0. Use the new user interface, including an edit box. This flag is equivalent to BIF_EDITBOX OR BIF_NEWDIALOGSTYLE. |
| BIF_UAHINT | Version 6.0. When combined with BIF_NEWDIALOGSTYLE, adds a usage hint to the dialog box, in place of the edit box. BIF_EDITBOX overrides this flag. |
| BIF_NONEWFOLDERBUTTON | Version 6.0. Do not include the **New Folder** button in the browse dialog box. |
| BIF_NOTRANSLATETARGETS | Version 6.0. When the selected item is a shortcut, return the PIDL of the shortcut itself rather than its target. |
| BIF_BROWSEFORCOMPUTER | Only return computers. If the user selects anything other than a computer, the **OK** button is grayed. |
| BIF_BROWSEFORPRINTER | Only allow the selection of printers. If the user selects anything other than a printer, the **OK** button is grayed. In Windows XP and later systems, the best practice is to use a Windows XP-style dialog, setting the root of the dialog to the Printers and Faxes folder (CSIDL_PRINTERS). |
| BIF_BROWSEINCLUDEFILES | Version 4.71. The browse dialog box displays files as well as folders. |
| BIF_SHAREABLE | Version 5.0. The browse dialog box can display shareable resources on remote systems. This is intended for applications that want to expose remote shares on a local system. The BIF_NEWDIALOGSTYLE flag must also be set. |
| BIF_BROWSEFILEJUNCTIONS | Windows 7 and later. Allow folder junctions such as a library or a compressed file with a .zip file name extension to be browsed. |

#### Notes

If COM is initialized through CoInitializeEx with the COINIT_MULTITHREADED flag set, **AfxShellBrowserForFolder** fails if BIF_NEWDIALOGSTYLE or BIF_USENEWUI are passed.

#### Return value

The path of the selected folder.

#### Remarks

If you don't pass any flags, the function will use BIF_RETURNONLYFSDIRS OR BIF_DONTGOBELOWDOMAIN OR BIF_USENEWUI OR BIF_RETURNFSANCESTORS.

#### Usage example

```
DIM cws AS CWSTR = AfxBrowseForFolder(hwnd, "C:")
```

# <a name="AfxOpenFileDialog"></a>AfxOpenFileDialog

Creates an **Open** dialog box that lets the user specify the drive, directory, and the name of a file or set of files to be opened. The dialog box uses the Explorer-style user interface.

```
FUNCTION AfxOpenFileDialog (BYVAL hwndOwner AS HWND, BYREF wszTitle AS WSTRING, BYREF wszFile AS WSTRING, _
   BYREF wszInitialDir AS WSTRING, BYREF wszFilter AS WSTRING, BYREF wszDefExt AS WSTRING, _
   BYVAL pdwFlags AS DWORD PTR = NULL, BYVAL pdwBufLen AS DWORD PTR = NULL) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *hwndOwner* | A handle to the window that owns the dialog box. This member can be any valid window handle, or it can be NULL if the dialog box has no owner. |
| *wszTitle* | A string to be placed in the title bar of the dialog box. If this member is NULL, the system uses the default title (that is, **Open**). |
| *wszFile* | The file name used to initialize the **File Name** edit control. |
| *wszInitialDir* | The initial directory.  If no initial directory is specified, the dialog will use the current directory. |
| *wszFilter* | A buffer containing pairs of "\|" separated strings. The first string in each pair is a display string that describes the filter (for example, "Text Files"), and the second string specifies the filter pattern (for example, "\*.TXT"). To specify multiple filter patterns for a single display string, use a semicolon to separate the patterns (for example, "\*.TXT;\*.DOC;\*.BAK"). A pattern string can be a combination of valid file name characters and the asterisk (\*) wildcard character. Do not include spaces in the pattern string. The system does not change the order of the filters. It displays them in the **File Types** combo box in the order specified in *wszFilter*. |
| *wszDefExt* | The default extension. This extension is appended to the file name if the user fails to type an extension. This string can be any length, but only the first three characters are appended. The string should not contain a period (.). If this member is NULL and the user fails to type an extension, no extension is appended. |
| *pdwFlags* | \[in, out, optional] A set of bit flags you can use to initialize the dialog box. When the dialog box returns, it sets these flags to indicate the user's input. This member can be a combination of the following flags. |
| *pdwBufLen* | \[in, out, optional] Maximum length of the returned string containing the selected file or files. |

| Flag       | Description |
| ---------- | ----------- |
| OFN_ALLOWMULTISELECT | The **File Name** list box allows multiple selections. |
| OFN_DONTADDTORECENT | Prevents the system from adding a link to the selected file in the file system directory that contains the user's most recently used documents. To retrieve the location of this directory, call the **SHGetSpecialFolderLocation** function with the CSIDL_RECENT flag. |
| OFN_EXTENSIONDIFFERENT | The user typed a file name extension that differs from the extension specified by *wszDefExt*. The function does not use this flag if *wszDefExt* is NULL. |
| OFN_FILEMUSTEXIST | The user can type only names of existing files in the **File Name** entry field. If this flag is specified and the user enters an invalid name, the dialog box procedure displays a warning in a message box. If this flag is specified, the OFN_PATHMUSTEXIST flag is also used. |
| OFN_FORCESHOWHIDDEN | Forces the showing of system and hidden files, thus overriding the user setting to show or not show hidden files. However, a file that is marked both system and hidden is not shown. |
| OFN_HIDEREADONLY | Hides the **Read Only** check box. |
| OFN_NODEREFERENCELINKS | Directs the dialog box to return the path and file name of the selected shortcut (.LNK) file. If this value is not specified, the dialog box returns the path and file name of the file referenced by the shortcut. |
| OFN_NONETWORKBUTTON | Hides and disables the **Network** button. |
| OFN_NOREADONLYRETURN | The returned file does not have the **Read Only** check box selected and is not in a write-protected directory. |
| OFN_PATHMUSTEXIST | The user can type only valid paths and file names. If this flag is used and the user types an invalid path and file name in the **File Name** entry field, the dialog box function displays a warning in a message box. |
| OFN_READONLY | Causes the **Read Only** check box to be selected initially when the dialog box is created. This flag indicates the state of the **Read Only** check box when the dialog box is closed. |
| OFN_SHOWHELP | Causes the dialog box to display the **Help** button. The *hwndOwner* member must specify the window to receive the HELPMSGSTRING registered messages that the dialog box sends when the user clicks the **Help** button. |

#### Return value

If the OFN_ALLOWMULTISELECT flag is set and the user selects multiple files, the returned string contains the current directory followed by the file names of the selected files. For Explorer-style dialog boxes, the directory and file name strings are separated by semicolons. If the user selects only one file, the returned string does not have a separator between the path and file name.

Parse the number of ",". If only one, then the user has selected only a file and the string contains the full path. If more, The first substring contains the path and the others the files. If the user has not selected any file, an empty string is returned. On failure, an empty string is returned and, if not null, the pdwBufLen parameter will be filled by the size of the required buffer in characters.

#### Usage example:

```
DIM wszFile AS WSTRING * 260 = "*.*"
DIM wszInitialDir AS STRING * 260 = CURDIR
DIM wszFilter AS WSTRING * 260 = "BAS files (*.BAS)|*.BAS|" & "All Files (*.*)|*.*|"
DIM dwFlags AS DWORD = OFN_EXPLORER OR OFN_FILEMUSTEXIST OR OFN_HIDEREADONLY OR OFN_ALLOWMULTISELECT
DIM cws AS CWSTR = AfxOpenFileDialog(hwnd, "", wszFile, wszInitialDir, wszFilter, "BAS", @dwFlags, NULL)
AfxMsg cws
```

# <a name="AfxSaveFileDialog"></a>AfxSaveFileDialog

Creates a **Save** dialog box that lets the user specify the drive, directory, and name of a file to save. The dialog box uses the Explorer-style user interface.

```
FUNCTION AfxSaveFileDialog (BYVAL hwndOwner AS HWND, BYREF wszTitle AS WSTRING, BYREF wszFile AS WSTRING, _
   BYREF wszInitialDir AS WSTRING, BYREF wszFilter AS WSTRING, BYREF wszDefExt AS WSTRING, _
   BYVAL pdwFlags AS DWORD PTR = NULL) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *hwndOwner* | A handle to the window that owns the dialog box. This member can be any valid window handle, or it can be NULL if the dialog box has no owner. |
| *wszTitle* | A string to be placed in the title bar of the dialog box. If this member is NULL, the system uses the default title (that is, **Save As**). |
| *wszFile* | The file name used to initialize the **File Name** edit control. |
| *wszInitialDir* | The initial directory.  If no initial directory is specified, the dialog will use the current directory. |
| *wszFilter* | A buffer containing pairs of "\|" separated strings. The first string in each pair is a display string that describes the filter (for example, "Text Files"), and the second string specifies the filter pattern (for example, "\*.TXT"). To specify multiple filter patterns for a single display string, use a semicolon to separate the patterns (for example, "\*.TXT;\*.DOC;\*.BAK"). A pattern string can be a combination of valid file name characters and the asterisk (\*) wildcard character. Do not include spaces in the pattern string. The system does not change the order of the filters. It displays them in the **File Types** combo box in the order specified in *wszFilter*. |
| *wszDefExt* | The default extension. This extension is appended to the file name if the user fails to type an extension. This string can be any length, but only the first three characters are appended. The string should not contain a period (.). If this member is NULL and the user fails to type an extension, no extension is appended. |
| *pdwFlags* | \[in, out, optional] A set of bit flags you can use to initialize the dialog box. When the dialog box returns, it sets these flags to indicate the user's input. This member can be a combination of the following flags. |

| Flag       | Description |
| ---------- | ----------- |
| OFN_CREATEPROMPT | If the user specifies a file that does not exist, this flag causes the dialog box to prompt the user for permission to create the file. If the user chooses to create the file, the dialog box closes and the function returns the specified name; otherwise, the dialog box remains open. |
| OFN_DONTADDTORECENT | Prevents the system from adding a link to the selected file in the file system directory that contains the user's most recently used documents. To retrieve the location of this directory, call the **SHGetSpecialFolderLocation** function with the CSIDL_RECENT flag. |
| OFN_EXTENSIONDIFFERENT | The user typed a file name extension that differs from the extension specified by *wszDefExt*. The function does not use this flag if *wszDefExt* is NULL. |
| OFN_FORCESHOWHIDDEN | Forces the showing of system and hidden files, thus overriding the user setting to show or not show hidden files. However, a file that is marked both system and hidden is not shown. |
| OFN_HIDEREADONLY | Hides the **Read Only** check box |
| OFN_NOCHANGEDIR | Restores the current directory to its original value if the user changed the directory while searching for files. |
| OFN_NODEREFERENCELINKS | Directs the dialog box to return the path and file name of the selected shortcut (.LNK) file. If this value is not specified, the dialog box returns the path and file name of the file referenced by the shortcut. |
| OFN_NOTESTFILECREATE | The file is not created before the dialog box is closed. This flag should be specified if the application saves the file on a create-nonmodify network share. When an application specifies this flag, the library does not check for write protection, a full disk, an open drive door, or network protection. Applications using this flag must perform file operations carefully, because a file cannot be reopened once it is closed. |
| OFN_NONETWORKBUTTON | Hides and disables the **Network** button. |
| OFN_NOREADONLYRETURN | The returned file does not have the **Read Only** check box selected and is not in a write-protected directory. |
| OFN_OVERWRITEPROMPT | The user can type only valid paths and file names. If this flag is used and the user types an invalid path and file name in the **File Name** entry field, the dialog box function displays a warning in a message box. |
| OFN_PATHMUSTEXIST | The user can type only valid paths and file names. If this flag is used and the user types an invalid path and file name in the **File Name** entry field, the dialog box function displays a warning in a message box. |
| OFN_SHOWHELP | Causes the dialog box to display the **Help** button. The hwndOwner member must specify the window to receive the HELPMSGSTRING registered messages that the dialog box sends when the user clicks the **Help** button. |

#### Return value

The path of the file to be saved.

#### Usage example:

```
DIM wszFile AS WSTRING * 260 = "*.*"
DIM wszInitialDir AS STRING * 260 = CURDIR
DIM wszFilter AS WSTRING * 260 = "BAS files (*.BAS)|*.BAS|" & "All Files (*.*)|*.*|"
DIM dwFlags AS DWORD = OFN_EXPLORER OR OFN_FILEMUSTEXIST OR OFN_HIDEREADONLY OR OFN_OVERWRITEPROMPT
DIM cws AS CWSTR = AfxSaveFileDialog(hwnd, "", wszFile, wszInitialDir, wszFilter, "BAS", @dwFlags)
AfxMsg cws
```

# <a name="AfxChDir"></a>AfxChDir / AfxSetCurDir / AfxSetCurrentDirectory

Changes the current directory for the current process.

```
FUNCTION AfxChDir (BYVAL pwszPathName AS LPCWSTR) AS BOOLEAN
FUNCTION AfxSetCurDir (BYVAL pwszPathName AS LPCWSTR) AS BOOLEAN
FUNCTION AfxSetCurrentDirectory (BYVAL pwszPathName AS LPCWSTR) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *lpPathName* | The path to the new current directory. This parameter may specify a relative path or a full path. In either case, the full path of the specified directory is calculated and stored as the current directory. |

#### Return value

If the function succeeds, the return value is TRUE.<br>
If the function fails, the return value is FALSE.<br>
To get extended error information, call **GetLastError**.

# <a name="AfxCopyFile"></a>AfxCopyFile / AfxFileCopy

Copies an existing file to a new file.

```
FUNCTION AfxCopyFile (BYVAL lpExistingFileName AS LPCWSTR, BYVAL lpNewFileName AS LPCWSTR, _
   BYVAL bFailIfExists AS BOOLEAN = FALSE) AS BOOLEAN
FUNCTION AfxFileCopy (BYVAL lpExistingFileName AS LPCWSTR, BYVAL lpNewFileName AS LPCWSTR, _
   BYVAL bFailIfExists AS BOOLEAN = FALSE) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *lpExistingFileName* | The name of an existing file. To extend the limit of MAX_PATH characters to 32,767 wide characters prepend "\\\\?\\" to the path. If *lpExistingFileName* does not exist, **CopyFile** fails, and **GetLastError** returns ERROR_FILE_NOT_FOUND. |
| *lpNewFileName* | The name of the new file. To extend the limit of MAX_PATH characters to 32,767 wide characters prepend "\\\\?\\" to the path. |
| *bFailIfExists* | If this parameter is TRUE and the new file specified by *lpNewFileName* already exists, the function fails. If this parameter is FALSE and the new file already exists, the function overwrites the existing file and succeeds. |

#### Return value

If the function succeeds, the return value is TRUE.<br>
If the function fails, the return value is FALSE. To get extended error information, call **GetLastError**.

#### Rermarks

**AfxFileCopy** is an unicode replacement for Free Basic's **FileCopy** and returns 0 on success, or 1 if an error occurred.

# <a name="AfxMakeDir"></a>AfxMakeDir / AfxCreateDirectory / AfxMkDir

Creates a new directory.

```
FUNCTION AfxCreateDirectory (BYVAL lpPathName AS LPCWSTR) AS BOOLEAN
FUNCTION AfxMakeDir (BYVAL lpPathName AS LPCWSTR) AS BOOLEAN
FUNCTION AfxMkDir (BYVAL lpPathName AS LPCWSTR) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *lpPathName* | The path of the directory to be created. To extend the limit to 32,767 wide characters, prepend "\\\\?\\" to the path. |

#### Return value:

If the function succeeds, the return value is TRUE.<br>
If the function fails, the return value is FALSE.<br>
To get extended error information, call **GetLastError**.

Possible errors include the following.

| Error      | Description |
| ---------- | ----------- |
| ERROR_ALREADY_EXISTS | The specified directory already exists. |
| ERROR_PATH_NOT_FOUND | One or more intermediate directories do not exist; this function will only create the final directory in the path. |

#### Remarks

**AfxMkDir** is an unicode replacement for Free Basic's **MkDir** and returns 0 on success, or -1 on failure.

# <a name="AfxDeleteFile"></a>AfxDeleteFile / AfxKill

Deletes the specified file.

```
FUNCTION AfxDeleteFile (BYVAL pwszFileSpec AS WSTRING PTR) AS BOOLEAN
FUNCTION AfxKill (BYVAL pwszFileSpec AS WSTRING PTR) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *pwszFileSpec* | The full path and name of the file to delete. To extend the limit to 32,767 wide characters, prepend "\\\\?\\" to the path. |

#### Return value:

If the function succeeds, the return value is TRUE.<br>
If the function fails, the return value is FALSE.<br>
To get extended error information, call **GetLastError**.

#### Remarks

If an application attempts to delete a file that does not exist, this function fails with ERROR_FILE_NOT_FOUND. If the file is a read-only file, the function fails with ERROR_ACCESS_DENIED.

**AfxKill** is an unicode replacement for Free Basic's **Kill** and returns 0 on success, or -1 on failure.

# <a name="AfxFileDateTime"></a>AfxFileDateTime

Returns the file's last modified date and time as Date Serial.

```
FUNCTION AfxFileDateTime (BYREF wszFileName AS WSTRING) AS DOUBLE
```

#### Return value

The date and time as a Date Serial. If it fails, it returns 0.

#### Remark

Unicode replacement for Free Basic's **FileDateTime**.

#### Example

```
#include "windows.bi"
#include "vbcompat.bi"
#include "Afx/AfxWin.bi"
DIM wszFileName AS WSTRING * MAX_PATH = ExePath & "\c2.bas"
DIM dt AS DOUBLE = AfxFileDateTime(wszFileName)
PRINT Format(dt, "yyyy-mm-dd hh:mm AM/PM")
```

# <a name="AfxCurDir"></a>AfxCurdir / AfxGetCurDir / AfxGetCurrentDirectory

Retrieves the current directory for the current process.

```
FUNCTION AfxCurDir () AS CWSTR
FUNCTION AfxGetCurDir () AS CWSTR
FUNCTION AfxGetCurrentDirectory () AS CWSTR
```

#### Return value

The name of the current directory for the current process.

#### Remark

Unicode replacement for Free Basic's **CurDir**.

# <a name="AfxMoveFile"></a>AfxRenameFile / AfxMoveFile / AfxName

Moves an existing file or a directory, including its children.

```
FUNCTION AfxRenameFile (BYVAL lpExistingFileName AS LPCWSTR, BYVAL lpNewFileName AS LPCWSTR) AS BOOLEAN
FUNCTION AfxMoveFile (BYVAL lpExistingFileName AS LPCWSTR, BYVAL lpNewFileName AS LPCWSTR) AS BOOLEAN
FUNCTION AfxName (BYVAL lpExistingFileName AS LPCWSTR, BYVAL lpNewFileName AS LPCWSTR) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *lpExistingFileName* | The name of an existing file. To extend the limit to 32,767 wide characters, prepend "\\\\?\\" to the path. If *lpExistingFileName* does not exist, **AfxRenameFile** fails, and **GetLastError** returns ERROR_FILE_NOT_FOUND. |
| *lpNewFileName* | The name of the new file. To extend the limit to 32,767 wide characters, prepend "\\\\?\\" to the path. |

#### Return value:

If the function succeeds, the return value is TRUE.<br>
If the function fails, the return value is FALSE.<br>
To get extended error information, call **GetLastError**.

#### Remark

**AfxName** is an unicode replacement for Free Basic's **Name** and returns 0 on success, or non-zero on failure.

# <a name="AfxRemoveDir"></a>AfxRemoveDir / AfxRmDir / AfxRemoveDirectory

Deletes an existing empty directory.

```
FUNCTION AfxRemoveDir (BYVAL lpPathName AS LPCWSTR) AS BOOLEAN
FUNCTION AfxRemoveDirectory (BYVAL lpPathName AS LPCWSTR) AS BOOLEAN
FUNCTION AfxRmDir (BYVAL lpPathName AS LPCWSTR) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *lpPathName* | The path of the directory to be removed. This path must specify an empty directory, and the calling process must have delete access to the directory. To extend the limit to 32,767 wide characters, prepend "\\\\?\\" to the path. |

#### Return value:

If the function succeeds, the return value is TRUE.<br>
If the function fails, the return value is FALSE.<br>
To get extended error information, call **GetLastError**.

#### Remaks

The **AfxRemoveDir** function marks a directory for deletion on close. Therefore, the directory is not removed until the last handle to the directory is closed. To recursively delete the files in a directory, use the **SHFileOperation** function. **AfxRemoveDir** removes a directory junction, even if the contents of the target are not empty; the function removes directory junctions regardless of the state of the target object. 

**AfxRmDir** is an unicode replacement for Free Basic's **RmDir** and returns 0 on success, or -1 on failure.

# <a name="AfxFileExists"></a>AfxFileExists

Searches a directory for a file or subdirectory with a name that matches a specific name (or partial name if wildcards are used).

```
FUNCTION AfxFileExists (BYVAL pwszFileSpec AS WSTRING PTR) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *pwszFileSpec* | The directory or path, and the file name, which can include wildcard characters, for example, an asterisk (\*) or a question mark (?). This parameter should not be NULL, an invalid string (for example, an empty string or a string that is missing the terminating null character), or end in a trailing backslash (\\). If the string ends with a wildcard, period (.), or directory name, the user must have access permissions to the root and all subdirectories on the path. To extend the limit of MAX_PATH wide characters to 32,767 wide characters, prepend "\\\\?\\" to the path. |

#### Return value

Boolean. TRUE if the specified file exist or FALSE otherwise.

#### Remarks

Prepending the string "\\\\?\\" does not allow access to the root directory.

On network shares, you can use an pwszFileSpec in the form of the following: "\\\\server\service\\\*". However, you cannot use an pwszFileSpec that points to the share itself; for example, "\\\\server\service" is not valid.

To examine a directory that is not a root directory, use the path to that directory, without a trailing backslash. For example, an argument of "C:\Windows" returns information about the directory "C:\Windows", not about a directory or file in "C:\Windows". To examine the files and directories in "C:\Windows", use an *pwszFileSpec* of "C:\Windows\*".

Be aware that some other thread or process could create or delete a file with this name between the time you query for the result and the time you act on the information. If this is a potential concern for your application, one possible solution is to use the **CreateFile** function with CREATE_NEW (which fails if the file exists) or OPEN_EXISTING (which fails if the file does not exist).

# <a name="AfxFolderExists"></a>AfxFolderExists

Searches a directory for a file or subdirectory with a name that matches a specific name (or partial name if wildcards are used).

```
FUNCTION AfxFolderExists (BYVAL pwszFileSpec AS WSTRING PTR) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *pwszFileSpec* | The directory or path, and the file name, which can include wildcard characters, for example, an asterisk (\*) or a question mark (?). This parameter should not be NULL, an invalid string (for example, an empty string or a string that is missing the terminating null character), or end in a trailing backslash (\\). If the string ends with a wildcard, period (.), or directory name, the user must have access permissions to the root and all subdirectories on the path. To extend the limit of MAX_PATH wide characters to 32,767 wide characters, prepend "\\\\?\\" to the path. |

#### Return value

Boolean. TRUE if the specified file exist or FALSE otherwise.

#### Remarks

Prepending the string "\\\\?\\" does not allow access to the root directory.

On network shares, you can use an *pwszFileSpec* in the form of the following: "\\\\server\service\\\*". However, you cannot use an pwszFileSpec that points to the share itself; for example, "\\\\server\service" is not valid.

To examine a directory that is not a root directory, use the path to that directory, without a trailing backslash. For example, an argument of "C:\Windows" returns information about the directory "C:\Windows", not about a directory or file in "C:\Windows". To examine the files and directories in "C:\Windows", use an pwszFileSpec of "C:\Windows\\\*".

Be aware that some other thread or process could create or delete a file with this name between the time you query for the result and the time you act on the information. If this is a potential concern for your application, one possible solution is to use the **CreateFile** function with CREATE_NEW (which fails if the file exists) or OPEN_EXISTING (which fails if the file does not exist).

# <a name="AfxGetFileSize"></a>AfxGetFileSize / AfxFileLen

Returns the size in bytes of the specified file.

```
FUNCTION AfxGetFileSize (BYREF wszFileSpec AS WSTRING) AS ULONGLONG
FUNCTION AfxFileLen (BYREF wszFileSpec AS WSTRING) AS ULONGLONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszFileSpec* | The path to a file. To extend the limit of MAX_PATH wide characters to 32,767 wide characters, prepend "\\\\?\\" to the path. |

#### Return value

The size in bytes of the file on success, or 0 on failure.

# <a name="AfxExePath"></a>AfxExePath / AfxGetExePath

Returns the path of the program which is currently executing.

```
FUNCTION AfxExePath () AS CWSTR
FUNCTION AfxGetExePath () AS CWSTR
```

#### Remarks

Unicode replacement for Free Basic's **ExePath** function. The path name has not a trailing backslash, except if it is a drive, e.g. "C:\".

# <a name="AfxGetExePathName"></a>AfxGetExePathName

Returns the path of the program which is currently executing.

```
FUNCTION AfxGetExePathName () AS CWSTR
```

#### Remarks

The path name has a trailing backslash.

# <a name="AfxGetDriveType"></a>AfxGetDriveType

Determines whether a disk drive is a removable, fixed, CD-ROM, RAM disk, or network drive.

```
FUNCTION AfxGetDriveType (BYVAL lpRootPathName AS LPCWSTR) as UINT
```

| Parameter  | Description |
| ---------- | ----------- |
| *lpRootPathName* | The root directory for the drive. A trailing backslash is required. If this parameter is NULL, the function uses the root of the current directory. |

#### Return value

DRIVE_UNKNOWN (0), DRIVE_NO_ROOT_DIR (1), DRIVE_REMOVABLE (2), DRIVE_FIXED(3), DRIVE_REMOTE (4), DRIVE_CDROM (5), DRIVE_RAMDISK (6).

# <a name="AfxGetExeFileExt"></a>AfxGetExeFileExt

Parses a path/filename and returns the extension portion of the path/file name.

```
FUNCTION AfxGetExeFileExt () AS CWSTR
```

#### Return value

The extension portion of the file name. That is the last period (.) in the string plus the text to the right of it.

# <a name="AfxGetExeFileName"></a>AfxGetExeFileName

Returns the file name of the program which is currently executing.

```
FUNCTION AfxGetExeFileName () AS CWSTR
```

# <a name="AfxGetExeFileNameX"></a>AfxGetExeFileNameX

Returns the file name and extension of the program which is currently executing.

```
FUNCTION AfxGetExeFileNameX () AS CWSTR
```

# <a name="AfxGetExeFullPath"></a>AfxGetExeFullPath

Returns the complete drive, path, file name, and extension of the program which is currently executing.

```
FUNCTION AfxGetExeFullPath () AS CWSTR
```

# <a name="AfxGetFileExt"></a>AfxGetFileExt

Parses a path/filename and returns the extension portion of the path/file name. That is the last period (.) in the string plus the text to the right of it.

```
FUNCTION AfxGetFileExt (BYREF wszPath AS WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | The path/filename string. |

# <a name="AfxGetFileCreationTime"></a>AfxGetFileCreationTime

Returns the time the file was created, in FILETIME format.

```
FUNCTION AfxGetFileCreationTime (BYREF wszFileSpec AS WSTRING, BYVAL bUTC AS BOOLEAN = TRUE) AS FILETIME
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszFileSpec* | The directory or path, and the file name, which can include wildcard characters, for example, an asterisk (\*) or a question mark (?). This parameter should not be NULL, an invalid string (for example, an empty string or a string that is missing the terminating null character), or end in a trailing backslash (\\). If the string ends with a wildcard, period (.), or directory name, the user must have access permissions to the root and all subdirectories on the path. To extend the limit from MAX_PATH to 32,767 wide characters, prepend "\\\\?\\" to the path. |
| *bUTC* | Optional. Pass FALSE if you want to get the time in local time (the NTFS file system stores time values in UTC format, so they are not affected by changes in time zone or daylight saving time). **FileTimeToLocalFileTime** uses the current settings for the time zone and daylight saving time. Therefore, if it is daylight saving time, it takes daylight saving time into account, even if the file time you are converting is in standard time. |

# <a name="AfxGetFileLastAccessTime"></a>AfxGetFileLastAccessTime

Returns the time the file was accessed, in FILETIME format.

```
FUNCTION AfxGetFileLastAccessTime (BYREF wszFileSpec AS WSTRING, BYVAL bUTC AS BOOLEAN = TRUE) AS FILETIME
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszFileSpec* | The directory or path, and the file name, which can include wildcard characters, for example, an asterisk (\*) or a question mark (?). This parameter should not be NULL, an invalid string (for example, an empty string or a string that is missing the terminating null character), or end in a trailing backslash (\\). If the string ends with a wildcard, period (.), or directory name, the user must have access permissions to the root and all subdirectories on the path. To extend the limit from MAX_PATH to 32,767 wide characters, prepend "\\\\?\\" to the path. |
| *bUTC* | Optional. Pass FALSE if you want to get the time in local time (the NTFS file system stores time values in UTC format, so they are not affected by changes in time zone or daylight saving time). **FileTimeToLocalFileTime** uses the current settings for the time zone and daylight saving time. Therefore, if it is daylight saving time, it takes daylight saving time into account, even if the file time you are converting is in standard time. |

# <a name="AfxGetFileLastWriteTime"></a>AfxGetFileLastWriteTime

Returns the time the file was last written to, truncated, or overwritten, in FILETIME format.

```
FUNCTION AfxGetFileLastWriteTime (BYREF wszFileSpec AS WSTRING, BYVAL bUTC AS BOOLEAN = TRUE) AS FILETIME
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszFileSpec* | The directory or path, and the file name, which can include wildcard characters, for example, an asterisk (\*) or a question mark (?). This parameter should not be NULL, an invalid string (for example, an empty string or a string that is missing the terminating null character), or end in a trailing backslash (\\). If the string ends with a wildcard, period (.), or directory name, the user must have access permissions to the root and all subdirectories on the path. To extend the limit from MAX_PATH to 32,767 wide characters, prepend "\\\\?\\" to the path. |
| *bUTC* | Optional. Pass FALSE if you want to get the time in local time (the NTFS file system stores time values in UTC format, so they are not affected by changes in time zone or daylight saving time). **FileTimeToLocalFileTime** uses the current settings for the time zone and daylight saving time. Therefore, if it is daylight saving time, it takes daylight saving time into account, even if the file time you are converting is in standard time. |

# <a name="AfxGetFileName"></a>AfxGetFileName

Parses a path/filename and returns the file name portion. That is the text to the right of the last backslash (\) or colon (:), ending just before the last period (.).

```
FUNCTION AfxGetFileName (BYREF wszPath AS WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | The path/filename string. |

# <a name="AfxGetFileNameX"></a>AfxGetFileNameX

Parses a path/filename and returns the file name and extension portion. That is the text to the right of the last backslash (\\) or colon (:).

```
FUNCTION AfxGetFileNameX (BYREF wszPath AS WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | The path/filename string. |

# <a name="AfxGetFileVersion"></a>AfxGetFileVersion

Retrieves the version of the specified file multiplied by 100, e.g. 601 for version 6.01.

```
FUNCTION AfxGetFileVersion (BYVAL pwszFileName AS WSTRING PTR) AS DWORD
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | The path/filename string. |

# <a name="AfxGetFolderName"></a>AfxGetFolderName

Returns a string containing the name of the folder for a specified path, i.e. the path minus the file name.

```
FUNCTION AfxGetFolderName (BYREF wszPath AS WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | The path/filename string. |

# <a name="AfxGetKnowFolderPath"></a>AfxGetKnowFolderPath

Retrieves the path of an special folder.

```
FUNCTION AfxGetKnowFolderPath (BYVAL rfid AS CONST KNOWNFOLDERID CONST PTR, _
   BYVAL dwFlags AS DWORD = 0, BYVAL hToken AS HANDLE = NULL) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *rfid* | A reference to the KNOWNFOLDERID that identifies the folder. The folders associated with the known folder IDs might not exist on a particular system. |
| *dwFlags* | Flags that specify special retrieval options. This value can be 0; otherwise, it is one or more of the KNOWN_FOLDER_FLAG values. |
| *hToken* | An access token used to represent a particular user. This parameter is usually set to NULL, in which case the function tries to access the current user's instance of the folder. However, you may need to assign a value to *hToken* for those folders that can have multiple users but are treated as belonging to a single user. The most commonly used folder of this type is Documents. The calling application is responsible for correct impersonation when *hToken* is non-null. It must have appropriate security privileges for the particular user, including TOKEN_QUERY and TOKEN_IMPERSONATE, and the user's registry hive must be currently mounted. See [Access Control](https://msdn.microsoft.com/en-us/library/windows/desktop/aa374860(v=vs.85).aspx) for further discussion of access control issues.<br>Assigning the *hToken* parameter a value of -1 indicates the Default User. This allows clients of **SHGetKnownFolderIDList** to find folder locations (such as the Desktop folder) for the Default User. The Default User user profile is duplicated when any new user account is created, and includes special folders such as Documents and Desktop. Any items added to the Default User folder also appear in any new user account. Note that access to the Default User folders requires administrator privileges. |

#### Return value

The path of the requested folder on success, or an empty string on failure.

#### Remarks

Requires Windows Vista/Windows 7 or superior.

For a list of KNOWNFOLDERID constants see: [KNOWNFOLDERID](https://msdn.microsoft.com/en-us/library/windows/desktop/dd378457(v=vs.85).aspx)

#### Usage example

```
AfxGetKnowFolderPath(@FOLDERID_CommonPrograms)
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | The path/filename string. |

# <a name="AfxGetLongPathName"></a>AfxGetLongPathName

Retrieves the short path form of the specified path.
```
FUNCTION AfxGetLongPathName (BYREF wszPath AS WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | The path/filename string. To extend the limit of MAX_PATH wode characters to 32,767 wide characters, prepend "\\\\?\\" to the path. |

# <a name="AfxGetPathName"></a>AfxGetPathName

Parses a path/filename and returns the path portion. That is the text up to and including the last backslash (\) or colon (:).

```
FUNCTION AfxGetPathName (BYREF wszPath AS WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | The path/filename string. |

# <a name="AfxGetShortPathName"></a>AfxGetShortPathName

Retrieves the short path form of the specified path.

```
FUNCTION AfxGetShortPathName (BYREF wszPath AS WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | The path/filename string. To extend the limit of MAX_PATH wode characters to 32,767 wide characters, prepend "\\\\?\\" to the path. |

# <a name="AfxGetSpecialFolderLocation"></a>AfxGetSpecialFolderLocation

Retrieves the path of a special folder.

```
FUNCTION AfxGetSpecialFolderLocation (BYVAL nFolder AS LONG) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *nFolder* | A CSIDL value that identifies the folder of interest. |

#### Remarks

For a list of CSIDL values see: [CSIDL](https://msdn.microsoft.com/en-us/library/windows/desktop/bb762494(v=vs.85).aspx)

# <a name="AfxGetSystemDllPath"></a>AfxGetSystemDllPath

Retrieves the fully qualified path for the file that contains the specified module.

```
FUNCTION AfxGetSystemDllPath (BYREF wszDllName AS WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszDllName* | The name of the system DLL to find. |

#### Remarks

To locate the file for a module that was loaded by another process, use the **GetModuleFileNameEx** function.

# <a name="AfxGetWinDir"></a>AfxGetWinDir

Retrieves the path of the Windows directory. This path does not end with a backslash unless the Windows directory is the root directory. For example, if the Windows directory is named Windows on drive C, the path of the Windows directory retrieved by this function is C:\Windows. If the system was installed in the root directory of drive C, the path retrieved is C:\\.

```
FUNCTION AfxGetWinDir () AS CWSTR
```

# <a name="AfxIsCompressedFile"></a>AfxIsCompressedFile

Returns True if the specified file or directory is compressed; False if it is not.

```
FUNCTION AfxIsCompressedFile (BYREF wszFileSpec AS WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszFileSpec* | The path to a file. To extend the limit of MAX_PATH wide characters to 32,767 wide characters, prepend "\\\\?\\" to the path. |

# <a name="AfxIsEncryptedFile"></a>AfxIsEncryptedFile

Returns True if the specified file or directory is encrypted; False if it is not.

```
FUNCTION AfxIsEncryptedFile (BYREF wszFileSpec AS WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszFileSpec* | The path to a file. To extend the limit of MAX_PATH wide characters to 32,767 wide characters, prepend "\\\\?\\" to the path. |

# <a name="AfxIsFolder"></a>AfxIsFolder

Returns True if the specified path is a folder; False if it is not.

```
FUNCTION AfxIsFolder (BYREF wszFileSpec AS WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszFileSpec* | The path to a file. To extend the limit of MAX_PATH wide characters to 32,767 wide characters, prepend "\\\\?\\" to the path. |

# <a name="AfxIsHiddenFile"></a>AfxIsHiddenFile

Returns True if the specified path is a hidden file or directory; False if it is not.

```
FUNCTION AfxIsHiddenFile (BYREF wszFileSpec AS WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszFileSpec* | The path to a file. To extend the limit of MAX_PATH wide characters to 32,767 wide characters, prepend "\\\\?\\" to the path. |

# <a name="AfxIsNormalFile"></a>AfxIsNormalFile

Returns True if the specified path is a normal file (a file that does not have other attributes set); False if it is not.

```
FUNCTION AfxIsNormalFile (BYREF wszFileSpec AS WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszFileSpec* | The path to a file. To extend the limit of MAX_PATH wide characters to 32,767 wide characters, prepend "\\\\?\\" to the path. |

# <a name="AfxIsNotContentIndexedFile"></a>AfxIsNotContentIndexedFile

Returns TRUE if the specified file or directory is not to be indexed by the content indexing service; FALSE, otherwise.

```
FUNCTION AfxIsNotContentIndexedFile (BYREF wszFileSpec AS WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszFileSpec* | The path to a file. To extend the limit of MAX_PATH wide characters to 32,767 wide characters, prepend "\\\\?\\" to the path. |

# <a name="AfxIsOfflineFile"></a>AfxIsOfflineFile

Returns TRUE if the specified file file is not available immediately; FALSE, otherwise.

```
FUNCTION AfxIsOfflineFile (BYREF wszFileSpec AS WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszFileSpec* | The path to a file. To extend the limit of MAX_PATH wide characters to 32,767 wide characters, prepend "\\\\?\\" to the path. |

# <a name="AfxIsReadOnlyFile"></a>AfxIsReadOnlyFile

Returns True if the specified path is a read only file; False if it is not.

```
FUNCTION AfxIsReadOnlyFile (BYREF wszFileSpec AS WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszFileSpec* | The path to a file. To extend the limit of MAX_PATH wide characters to 32,767 wide characters, prepend "\\\\?\\" to the path. |

# <a name="AfxIsReparsePointFile"></a>AfxIsReparsePointFile

Returns TRUE if the specified path is a file or directory that has an associated reparse point, or a file that is a symbolic link.; FALSE, otherwise.

```
FUNCTION AfxIsReparsePointFile (BYREF wszFileSpec AS WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszFileSpec* | The path to a file. To extend the limit of MAX_PATH wide characters to 32,767 wide characters, prepend "\\\\?\\" to the path. |

# <a name="AfxIsSparseFile"></a>AfxIsSparseFile

Returns TRUE if the specified path is a sparse file; FALSE, otherwise.

```
FUNCTION AfxIsSparseFile (BYREF wszFileSpec AS WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszFileSpec* | The path to a file. To extend the limit of MAX_PATH wide characters to 32,767 wide characters, prepend "\\\\?\\" to the path. |

# <a name="AfxIsSystemFile"></a>AfxIsSystemFile

Returns True if the specified path is a system file; False if it is not.

```
FUNCTION AfxIsSystemFile (BYREF wszFileSpec AS WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszFileSpec* | The path to a file. To extend the limit of MAX_PATH wide characters to 32,767 wide characters, prepend "\\\\?\\" to the path. |

# <a name="AfxIsTemporaryFile"></a>AfxIsTemporaryFile

Returns True if the specified path is a temporary file; False if it is not.

```
FUNCTION AfxIsTemporaryFile (BYREF wszFileSpec AS WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszFileSpec* | The path to a file. To extend the limit of MAX_PATH wide characters to 32,767 wide characters, prepend "\\\\?\\" to the path. |

# <a name="AfxFileScan"></a>AfxFileScan

Scans a text file ans returns the number of occurrences of the specified delimiter.

```
FUNCTION AfxFileScanA (BYREF wszFileName AS WSTRING, BYREF Delimiter AS ZSTRING = CHR(13, 10)) AS DWORD
FUNCTION AfxFileScanW (BYREF wszFileName AS WSTRING, BYREF Delimiter AS WSTRING = CHR(13, 10)) AS DWORD
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszFileName* | Path of the file to scan. To extend the limit of MAX_PATH wide characters to 32,767 wide characters, prepend "\\\\?\\" to the path. |
| *Delimiter* | Optional. Delimiter to find. Default value is CHR(13, 10), which returns the number of lines. You can use CHR(10) for Linux files, "," for comma delimited cvs files, CHR(9) for spreadshets in tab delimited format, etc. |

#### Remarks

Use **AfxFileScanA** for ansi text files and **AfxFileScanW** for unicode text files.

# <a name="AfxFileReadAllLines"></a>AfxFileReadAllLines

Reads all the lines of the specified file into a safe array.

```
FUNCTION AfxFileReadAllLinesA (BYREF wszFileName AS WSTRING, BYREF Delimiter AS ZSTRING = CHR(13, 10)) AS CSafeArray
FUNCTION AfxFileReadAllLinesW (BYREF wszFileName AS WSTRING, BYREF Delimiter AS WSTRING = CHR(13, 10)) AS CSafeArray
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszFileName* | Path of the file to scan. To extend the limit of MAX_PATH wide characters to 32,767 wide characters, prepend "\\\\?\\" to the path. |
| *Delimiter* | Optional. Delimiter to find. Default value is CHR(13, 10), which returns the number of lines. You can use CHR(10) for Linux files, "," for comma delimited cvs files, CHR(9) for spreadshets in tab delimited format, etc. |

#### Remarks

Use **AfxFileReadAllLinesA** for ansi text files and **AfxFileReadAllLinesW** for unicode text files.

Because it returns a safe array, this function is located in the CSafeArray.inc include file.

# <a name="AfxSaveTempFile"></a>AfxSaveTempFile

Saves the contents of a string buffer in a temporary file.

```
FUNCTION AfxSaveTempFile (BYVAL pwszBuffer AS WSTRING PTR, BYREF wszExtension AS WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *pwszBuffer* | The string buffer to save. |
| *wszExtension* | Optional. The extension of the file name without a colon (e.g. "bas"). If an empty string is passed, the function will use "tmp" as the extension. |

#### Remarks

Temporary files whose names have been created by this function are not automatically deleted. To delete these files call **AfxDeleteFile**.

# <a name="AfxPathAddBackSlash"></a>AfxPathAddBackSlash

Adds a backslash to the end of a string to create the correct syntax for a path. If the source path already has a trailing backslash, no backslash will be added.

```
FUNCTION AfxPathAddBackSlash (BYREF wszPath AS CONST WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that represents a path. |

#### Return value

The changed path.

#### Usage example

```
DIM cws AS CWSTR = AfxPathAddBackSlash("C:\dir_name\dir_name\file_name")
```

# <a name="AfxPathAddExtension"></a>AfxPathAddExtension

Adds a file name extension to a path string.

```
FUNCTION AfxPathAddExtension (BYREF wszPath AS CONST WSTRING, BYVAL pwszExt AS WSTRING PTR = NULL) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that represents a path. |
| *pwszExt* | A string that contains the file name extension. |

#### Return value

The changed path.

#### Remarks

If there is already a file name extension present, no extension will be added. If the *wszPath* is an empty string, the result will be the file name extension only. If *wszExtension* is an empty string, an ".exe" extension will be added.

#### Usage examples

```
DIM cws AS CWSTR = AfxPathAddExtension("file") ' output: file.exe
DIM cws AS CWSTR = AfxPathAddExtension("file.doc") ' output: file.doc
DIM cws AS CWSTR = AfxPathAddExtension("file", ".txt") ' output: file.txt
DIM cws AS CWSTR = AfxPathAddExtension("", ".txt") ' output: .txt
```

# <a name="AfxPathAppend"></a>AfxPathAppend

Appends one path to the end of another.

```
FUNCTION AfxPathAppend (BYREF wszPath AS CONST WSTRING, BYREF wszMore AS CONST WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that represents a path. |
| *wszMore* | A string that contains the path to be appended. |

#### Return value

The changed path.

#### Remarks

This function automatically inserts a backslash between the two strings, if one is not already present.

The path supplied in *wszPath* cannot begin with "..\\" or ".\\" to produce a relative path string. If present, those periods are stripped from the output string. For example, appending "path3" to "..\\path1\\path2" results in an output of "\\path1\\path2\\path3" rather than "..\\path1\\path2\\path3".

#### Usage example

```
DIM cws AS CWSTR = AfxPathAppend("name_1\name_2", "name_3")
```

# <a name="AfxPathBuildRoot"></a>AfxPathBuildRoot

Creates a root path from a given drive number.

```
FUNCTION AfxPathBuildRoot (BYVAL nDrive AS LONG) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *nDrive* | A variable of type LONG that indicates the desired drive number. It should be between 0 and 25. |

#### Return value

The root path. If the call fails for any reason (for example, an invalid drive number), an empty string is returned.

#### Usage example

```
DIM cws AS CWSTR = AfxPathBuildRoot(2) ' output: C:\
```

# <a name="AfxPathCanonicalize"></a>AfxPathCanonicalize

Removes elements of a file path according to special strings inserted into that path.

```
FUNCTION AfxPathCanonicalize (BYREF wszPath AS CONST WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains the path to be canonicalized. |

#### Return value

The canonicalized path.

#### Remarks

This function allows the user to specify what to remove from a path by inserting special character sequences into the path. The ".." sequence indicates to remove a path segment from the current position to the previous path segment. The "." sequence indicates to skip over the next path segment to the following path segment. The root segment of the path cannot be removed.

If there are more ".." sequences than there are path segments, the contents of the returned string contains just the root, "\\".

#### Usage example

```
DIM cws AS CWSTR = AfxPathCanonicalize("A:\name_1\.\name_2\..\sname_3") ' output: A:\name_1\name_3
DIM cws AS CWSTR = AfxPathCanonicalize("A:\name_1\..\name_2\.\name_3") ' output: A:\name_2\name_3
DIM cws AS CWSTR = AfxPathCanonicalize("A:\name_1\name_2\.\name_3\..\name_4") ' output: A:\name_1\name_2\name_4
DIM cws AS CWSTR = AfxPathCanonicalize("A:\name_1\.\name_2\.\name_3\..\name_4\..") ' output: A:\name_1\name_2
DIM cws AS CWSTR = AfxPathCanonicalize("C:\..") ' output: C:\
```

# <a name="AfxPathCombine"></a>AfxPathCombine

Concatenates two strings that represent properly formed paths into one path; also concatenates any relative path elements.

```
FUNCTION AfxPathCombine (BYREF wszDir AS CONST WSTRING, BYREF wszFile AS CONST WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszDir* | A string that contains the directory path. This value can be "". |
| *wszFile* | A string that contains the file path. This value can be "". |

#### Return value

The concatenated path if successful, or an empty string otherwise.

#### Remarks

The directory path should be in the form of A:,B:, ..., Z:. The file path should be in a correct form that represents the file part of the path. The file path must not be null, and if it ends with a backslash, the backslash will be maintained.

#### Usage example

```
DIM cws AS CWSTR = AfxPathCombine("C:", "One\Two\Three") ' output: C:\One\Two\Three
```

# <a name="AfxPathCommonPrefix"></a>AfxPathCommonPrefix

Compares two paths to determine if they share a common prefix. A prefix is one of these types: "C:\\", ".", "..", "..\\".

```
FUNCTION AfxPathCommonPrefix (BYREF wszFile1 AS CONST WSTRING, BYREF wszFile2 AS CONST WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszFile1* | A string that contains the first path name. |
| *wszFile2* | A string that contains the second path name. |

#### Return value

The common prefix.

#### Usage example

```
DIM cws AS CWSTR = AfxPathCommonPrefix("C:\win\desktop\temp.txt", "c:\win\tray\sample.txt")
```

# <a name="AfxPathCompactPath"></a>AfxPathCompactPath

Truncates a file path to fit within a given pixel width by replacing path components with ellipses.

```
FUNCTION AfxPathCompactPath (BYVAL hDC AS HDC, BYREF wszPath AS CONST WSTRING, BYVAL dx AS DWORD) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *hDC* | A handle to the device context used for font metrics. This value can be NULL. |
| *wszPath* | A string that contains the path to be modified. |
| *dx* | The width, in pixels, in which the string must fit. |

#### Return value

The compacted path.

#### Remarks

This function uses the font currently selected in *hDC* to calculate the width of the text. This function will not compact the path beyond the base file name preceded by ellipses.

#### Usage example

```
DIM cws AS CWSTR = AfxPathCompactPath(NULL, "C:\path1\path2\sample.txt", 180)
```

# <a name="AfxPathCompactPathEx"></a>AfxPathCompactPathEx

Truncates a path to fit within a certain number of characters by replacing path components with ellipses.

```
FUNCTION AfxPathCompactPathEx (BYREF wszPath AS CONST WSTRING, BYVAL cchMax AS DWORD) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains the path to be altered. |
| *cchMax* | The maximum number of characters to be contained in the new string, including the terminating null character. For example, if *cchMax* = 8, the resulting string can contain a maximum of 7 characters plus the terminating null character. |

#### Return value

The compacted path.

#### Remarks

The '/' separator will be used instead of '\\' if the original string used it. If *wszPath* points to a file name that is too long, instead of a path, the file name will be truncated to *cchMax* characters, including the ellipsis and the terminating NULL character. For example, if the input file name is "My Filename" and *cchMax* is 10, *AfxPathCompactPathEx* will return "My Fil...".

#### Usage example

```
DIM cws AS CWSTR = AfxPathCompactPathEx("c:\test\My Filename", 18)
```

# <a name="AfxPathCreateFromUrl"></a>AfxPathCreateFromUrl

Converts a file URL to a Microsoft MS-DOS path.

```
FUNCTION AfxPathCreateFromUrl (BYREF wszUrl AS CONST WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszUrl* | A string that contains the URL. |

#### Return value

The MS-DOS path.

#### Usage example

```
DIM cws AS CWSTR = AfxPathCreateFromUrl("file:///C:/Documents%20and%20Settings/URIS.txt")
```

# <a name="AfxPathFileExists"></a>AfxPathFileExists

Determines whether a path to a file system object such as a file or directory is valid.

```
FUNCTION AfxPathFileExists (BYREF wszPath AS WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains the full path of the object to verify. |

#### Return value

Returns TRUE if the file exists, or FALSE otherwise.

#### Remarks

This function tests the validity of the path.

A path specified by Universal Naming Convention (UNC) is limited to a file only; that is, \\\\server\\share\file is permitted. A UNC path to a server or server share is not permitted; that is, \\\\server or \\\\server\\share. This function returns FALSE if a mounted remote drive is out of service.

# <a name="AfxPathFindExtension"></a>AfxPathFindExtension

Searches a path for an extension.

```
FUNCTION AfxPathFindExtension (BYREF wszPath AS CONST WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains the path to search, including the extension being searched for. |

#### Return value

The file name extension.

#### Remarks

Note that a valid file name extension cannot contain a space.

#### Usage example

```
DIM cws AS CWSTR = AfxPathFindExtension("C:\TEST\filetxt")
```

# <a name="AfxPathFindFileName"></a>AfxPathFindFileName

Searches a path for a file name.

```
FUNCTION AfxPathFindFileName (BYREF wszPath AS CONST WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains the path to search. |

#### Return value

The file name.

# <a name="AfxPathFindNextComponent"></a>AfxPathFindNextComponent

Searches a path for a file name.

```
FUNCTION AfxPathFindNextComponent (BYREF wszPath AS CONST WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains the path to parse. This string must not be longer than MAX_PATH characters, plus the terminating null character. Path components are delimited by backslashes. For instance, the path "c:\\path1\\path2\\file.txt" has four components: c:, path1, path2, and file.txt. |

#### Return value

The truncated path.

#### Remarks

**AfxPathFindNextComponent** walks a path string until it encounters a backslash ("\\"), ignores everything up to that point including the backslash, and returns the rest of the path. Therefore, if a path begins with a backslash (such as \\path1\\path2), the function simply removes the initial backslash and returns the rest (path1\\path2).

#### Usage example

```
DIM cws AS CWSTR = AfxPathFindNextComponent("C:\TEST\file.txt")
```

# <a name="AfxPathFindOnPath"></a>AfxPathFindOnPath

Searches for a file.

```
FUNCTION AfxPathFindOnPath (BYREF wszFile AS CONST WSTRING, _
   BYVAL ppwszOtherDirs AS WSTRING PTR PTR = NULL) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszFile* | A string that contains the file name for which to search. If the search is successful, this parameter is used to return the fully-qualified path name. |
| *ppwszOtherDirs* | An optional, null-terminated array of directories to be searched first. This value can be NULL. |

#### Return value

The fully-qualified path name on success or an empty string on failure.

#### Remarks

**AfxPathFindOnPath** searches for the file specified by *wszFile*. If no directories are specified in *ppwszOtherDirs*, it attempts to find the file by searching standard directories such as System32 and the directories specified in the PATH environment variable. To expedite the process or enable **AfxPathFindOnPath** to search a wider range of directories, use the *ppwszOtherDirs* parameter to specify one or more directories to be searched first. If more than one file has the name specified by *wszFile*, **AfxPathFindOnPath** returns the first instance it finds.

# <a name="AfxPathFindSuffixArray"></a>AfxPathFindSuffixArray

Determines whether a given file name has one of a list of suffixes.

```
FUNCTION AfxPathFindSuffixArray (BYREF wszPath AS WSTRING, BYVAL apwszSuffix AS WSTRING PTR PTR, _
   BYVAL iArraySize AS LONG) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains the file name to be tested. A full path can be used. |
| *apwszSuffix* | An array of *iArraySize* string pointers. Each string pointed to is null-terminated and contains one suffix. The strings can be of variable lengths. |
| *iArraySize* | The number of elements in the array pointed to by *apwszSuffix*. |

#### Return value

The matching suffix if successful, or an empty string if bstrPath does not end with one of the specified suffixes.

#### Remarks

This function uses a case-sensitive comparison. The suffix must match exactly.

# <a name="AfxPathGetArgs"></a>AfxPathGetArgs

Finds the command line arguments within a given path.

```
FUNCTION AfxPathGetArgs (BYREF wszPath AS CONST WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains the path to be searched. |

#### Return value

A string that contains the arguments portion of the path if successful or an empty string if there are no arguments in the path.

#### Remarks

This function should not be used on generic command path templates (from users or the registry), but rather should be used only on templates that the application knows to be well formed.

# <a name="AfxPathGetCharType"></a>AfxPathGetCharType

Determines the type of character in relation to a path.

```
FUNCTION AfxPathGetCharType (BYREF wszCh AS CONST WSTRING) AS UINT
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszCh* | The character for which to determine the type. |

#### Return value

Returns one or more of the following values that define the type of character.

| Retrn code | Description |
| ---------- | ----------- |
| GCT_INVALID | The character is not valid in a path. |
| GCT_LFNCHAR | The character is valid in a long file name. |
| GCT_SEPARATOR | The character is a path separator. |
| GCT_SHORTCHAR | The character is valid in a short (8.3) file name. |
| GCT_WILD | The character is a wildcard character. |

# <a name="AfxPathGetDriveNumber"></a>AfxPathGetDriveNumber

Searches a path for a drive letter within the range of 'A' to 'Z' and returns the corresponding drive number.

```
FUNCTION AfxPathGetDriveNumber (BYREF wszPath AS CONST WSTRING) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains the path to be searched. |

#### Return value

Returns 0 through 25 (corresponding to 'A' through 'Z') if the path has a drive letter, or -1 otherwise.

#### Usage example

```
DIM n AS LONG = AfxPathGetDriveNumber("C:\TEST\bar.txt") ' output: 2
```

# <a name="AfxPathIsContentType"></a>AfxPathIsContentType

Determines if a file's registered content type matches the specified content type. This function obtains the content type for the specified file type and compares that string with the *wszContentType*. The comparison is not case-sensitive.

```
FUNCTION AfxPathIsContentType (BYREF wszPath AS CONST WSTRING, _
   BYREF wszContentType AS CONST WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains the file whose content type will be compared. |
| *wszContentType* | A string that contains the content type string to which the file's registered content type will be compared. |

#### Return value

Returns nonzero if the file's registered content type matches *wszContentType*, or zero otherwise.

#### Usage example

```
DIM b AS BOOLEAN = AfxPathIsContentType("C:\TEST\bar.txt", "text/plain") ' output: true
DIM b AS BOOLEAN = AfxPathIsContentType("C:\TEST\bar.txt", "image/gif") ' output: false
```

# <a name="AfxPathIsDirectory"></a>AfxPathIsDirectory

Verifies that a path is a valid directory.

```
FUNCTION AfxPathIsDirectory (BYREF wszPath AS CONST WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains the path to verify. |

#### Return value

Returns TRUE if the path is a valid directory (FILE_ATTRIBUTE_DIRECTORY is set), or FALSE otherwise.

# <a name="AfxPathIsDirectoryEmpty"></a>AfxPathIsDirectoryEmpty

Determines whether a specified path is an empty directory.

```
FUNCTION AfxPathIsDirectoryEmpty (BYREF wszPath AS CONST WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains the path to be tested. |

#### Return value

Returns TRUE if wszPath is an empty directory. Returns FALSE if *wszPath* is not a directory, or if it contains at least one file other than "." or "..".

#### Remarks

"C:\" is considered a directory.

# <a name="AfxPathIsFileSpec"></a>AfxPathIsFileSpec

Searches a path for any path-delimiting characters (for example, ':' or '\\' ). If there are no path-delimiting characters present, the path is considered to be a File Spec path.

```
FUNCTION AfxPathIsFileSpec (BYREF wszPath AS CONST WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains the path to be searched. |

#### Return value

Returns TRUE if there are no path-delimiting characters within the path, or FALSE if there are path-delimiting characters.

# <a name="AfxPathIsHTMLFile"></a>AfxPathIsHTMLFile

Determines if a file is an HTML file. The determination is made based on the content type that is registered for the file's extension.

```
FUNCTION AfxPathIsHTMLFile (BYREF wszFile AS CONST WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszFile* | A string that contains the path and name of the file. |

#### Return value

Returns True if the file is an HTML file, or False otherwise.

# <a name="AfxPathIsLFNFileSpec"></a>AfxPathIsLFNFileSpec

Determines whether a file name is in long format.

```
FUNCTION AfxPathIsLFNFileSpec (BYREF wszName AS CONST WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszName* | A string that contains the file name to be tested. |

#### Return value

Returns True if wszName exceeds the number of characters allowed by the 8.3 format, or False otherwise.

# <a name="AfxPathIsNetworkPath"></a>AfxPathIsNetworkPath

Determines whether a path string represents a network resource.

```
FUNCTION AfxPathIsNetworkPath (BYREF wszPath AS CONST WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains contains the path. |

#### Return value

Returns True if the string represents a network resource, or False otherwise.

# <a name="AfxPathIsPrefix"></a>AfxPathIsPrefix

Searches a path to determine if it contains a valid prefix of the type passed by wszPrefix. A prefix is one of these types: "C:\\", ".", "..", "..\\".

```
FUNCTION AfxPathIsPrefix (BYREF wszPrefix AS CONST WSTRING, BYREF wszPath AS CONST WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPrefix* | A string that contains the prefix for which to search. |
| *wszPath* | A string that contains contains the path to be searched. |

#### Return value

Returns True if the compared path is the full prefix for the path, or False otherwise.

# <a name="AfxPathIsRelative"></a>AfxPathIsRelative

Searches a path and determines if it is relative.

```
FUNCTION AfxPathIsRelative (BYREF wszPath AS CONST WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains contains the path to search. |

#### Return value

Returns True if the path is relative, or False if it is absolute.

# <a name="AfxPathIsRoot"></a>AfxPathIsRoot

Parses a path to determine if it is a directory root.

```
FUNCTION AfxPathIsRoot (BYREF wszPath AS CONST WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains contains the path to be validated. |

#### Return value

Returns True if the specified path is a root, or False otherwise.

#### Remarks

Returns True for paths such as "\\", "X:\\" or "\\\\server\\share". Paths such as "..\\path2" or "\\\\server\\" return FALSE. 

# <a name="AfxPathIsSameRoot"></a>AfxPathIsSameRoot

Compares two paths to determine if they have a common root component.

```
FUNCTION AfxPathIsSameRoot (BYREF wszPath1 AS CONST WSTRING, BYREF wszPath2 AS CONST WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath1* | A string that contains contains the path to be compared. |
| *wszPath2* | A string that contains contains the second path to be compared. |

#### Return value

Returns True if both strings have the same root component, or False otherwise. If *wszPath1* contains only the server and share, this function also returns False.

# <a name="AfxPathIsSystemFolder"></a>AfxPathIsSystemFolder

Determines if an existing folder contains the attributes that make it a system folder. Alternately, this function indicates if certain attributes qualify a folder to be a system folder.

```
FUNCTION AfxPathIsSystemFolder (BYREF wszPath AS CONST WSTRING, BYVAL dwAttrb AS DWORD) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains the name of an existing folder. The attributes for this folder will be retrieved and compared with those that define a system folder. If this folder contains the attributes to make it a system folder, the function returns nonzero. If this value is NULL, this function determines if the attributes passed in dwAttrb qualify it to be a system folder. |
| *dwAttrb* | The file attributes to be compared. Used only if *wszPath* is NULL. In that case, the attributes passed in this value are compared with those that qualify a folder as a system folder. If the attributes are sufficient to make this a system folder, this function returns nonzero. These attributes are the attributes that are returned from **GetFileAttributes**. |

#### Return value

Returns True if the *wszPath* or *dwAttrb* represent a system folder, or False otherwise.

# <a name="AfxPathIsUNC"></a>AfxPathIsUNC

Determines if the string is a valid Universal Naming Convention (UNC) for a server and share path.

```
FUNCTION AfxPathIsUNC (BYREF wszPath AS CONST WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains contains the path to validate. |

#### Return value

Returns True if the string is a valid UNC path, or False otherwise.

# <a name="AfxPathIsUNCServer"></a>AfxPathIsUNCServer

Determines if a string is a valid Universal Naming Convention (UNC) for a server path only.

```
FUNCTION AfxPathIsUNCServer (BYREF wszPath AS CONST WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains contains the path to validate. |

#### Return value

Returns True if the string is a valid UNC path for a server only (no share name), or False otherwise.

# <a name="AfxPathIsUNCServerShare"></a>AfxPathIsUNCServerShare

Determines if a string is a valid Universal Naming Convention (UNC) share path, \\\\server\\share.

```
FUNCTION AfxPathIsUNCServerShare (BYREF wszPath AS CONST WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains contains the path to be validated. |

#### Return value

Returns True if the string is in the form \\\\server\\share, or False otherwise.

# <a name="AfxPathIsURL"></a>AfxPathIsURL

Tests a given string to determine if it conforms to a valid URL format.

```
FUNCTION AfxPathIsURL (BYREF wszPath AS CONST WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains the URL path to validate. |

#### Return value

Returns True if *wszPath* has a valid URL format, or Fañse otherwise.

#### Remarks

This function does not verify that the path points to an existing site—only that it has a valid URL format.


# <a name="AfxPathMakePretty"></a>AfxPathMakePretty

Converts a path to all lowercase characters to give the path a consistent appearance.

```
FUNCTION AfxPathMakePretty (BYREF wszPath AS CONST WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains the path to be converted. |

#### Return value

The changed path.

#### Remarks

This function only operates on paths that are entirely uppercase. For example: C:\WINDOWS will be converted to c:\windows, but c:\Windows will not be changed.


# <a name="AfxPathMakeSystemFolder"></a>AfxPathMakeSystemFolder

Gives an existing folder the proper attributes to become a system folder.

```
FUNCTION AfxPathMakeSystemFolder (BYREF wszPath AS CONST WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains the name of an existing folder that will be made into a system folder. |

#### Return value

Returns True if successful, or False otherwise.

# <a name="AfxPathMatchSpec"></a>AfxPathMatchSpec

Searches a string using a Microsoft MS-DOS wild card match type.

```
FUNCTION AfxPathMatchSpec (BYREF wszFile AS CONST WSTRING, BYREF wszSpec AS CONST WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszFile* | A string that contains the path to be searched. |
| *wszSpec* | A string that contains the file type for which to search. For example, to test whether *wszFile* is a .doc file, *wszSpec* should be set to "\*.doc". |

#### Return value

Returns True if the string matches, or False otherwise.


# <a name="AfxPathMatchSpecEx"></a>AfxPathMatchSpecEx

Searches a path to determine whether it contains a file of a specified file type extension.

```
FUNCTION AfxPathMatchSpecEx (BYREF wszFile AS CONST WSTRING, _
   BYREF wszSpec AS CONST WSTRING, BYVAL dwFlags AS DWORD) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszFile* | A string that contains the path to be searched. |
| *wszSpec* | A string that contains the file type extension or extensions for which to search. If exactly one extension is specified, set the PMSF_NORMAL flag in *dwFlags*. If more than one extension is specified, separate them with semicolons and set the PMSF_MULTIPLE flag. |
| *dwFlags* | Modifies the search condition. The following are valid flags:<br>*PMSF_NORMAL* (&h00000000): The *wszSpec* parameter points to a single file type extension to be matched.<br>*PMSF_MULTIPLE* (&h00000001): The *wszSpec* parameter points to a semicolon-delimited list of file type extensions to be matched.<br>*PMSF_DONT_STRIP_SPACES* (&h00010000): If PMSF_NORMAL is used, ignore leading spaces in the string pointed to by *wszSpec*. If PMSF_MULTIPLE is used, ignore leading spaces in each file type contained in the string pointed to by *wszSpec*. This flag can be combined with PMSF_NORMAL and PMSF_MULTIPLE. |

#### Return value

Returns one of the following values:

| Return code | Description |
| ----------- | ----------- |
| TRUE | A file type extension specified in *wszSpec* was found in the path pointed to by *wszFile*. |
| FALSE | The path pointed to by *wszFile* does not contain any file type extension specified in *wszSpec*. |

# <a name="AfxPathParseIconLocation"></a>AfxPathParseIconLocation

Parses a file location string that contains a file location and icon index, and returns separate values.

```
FUNCTION AfxPathParseIconLocation (BYREF wszIconFile AS CONST WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszIconFile* | A string that contains a file location string. It should be in the form "*path,iconindex*". |

#### Return value

Returns the valid icon index value.

#### Remarks

This function is useful for taking a DefaultIcon value retrieved from the registry by **SHGetValue** and separating the icon index from the path.

#### Usage example

```
DIM wszIconLocation AS CWSTR = AfxPathParseIconLocation("C:\TEST\sample.txt,3")
```

# <a name="AfxPathQuoteSpaces"></a>AfxPathQuoteSpaces

Searches a path for spaces. If spaces are found, the entire path is enclosed in quotation marks.

```
FUNCTION AfxPathQuoteSpaces (BYREF wszPath AS CONST WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains the path to be searched. |

#### Return value

True if spaces were found; otherwise, False.

# <a name="AfxPathRelativePathTo"></a>AfxPathRelativePathTo

Creates a relative path from one file or folder to another.

```
FUNCTION AfxPathRelativePathTo (BYREF wszFrom AS CONST WSTRING, BYVAL dwAttrFrom AS DWORD, _
   BYREF wszTo AS CONST WSTRING, BYVAL dwAttrTo AS DWORD) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszFrom* | A string that contains the path that defines the start of the relative path. |
| *dwAttrFrom* | The file attributes of *wszFrom*. If this value contains FILE_ATTRIBUTE_DIRECTORY, *wszFrom* is assumed to be a directory; otherwise, *wszFrom* is assumed to be a file. |
| *wszTo* | A string that contains the path that defines the endpoint of the relative path. |
| *dwAttrTo* | The file attributes of wszTo. If this value contains FILE_ATTRIBUTE_DIRECTORY, wszTo is assumed to be directory; otherwise, *wszTo* is assumed to be a file. |

#### Return value

Returns True if successful, or False otherwise.

#### Remarks

This function takes a pair of paths and generates a relative path from one to the other. The paths do not have to be fully-qualified, but they must have a common prefix, or the function will fail and return False.

For example, let the starting point, *wszFrom*, be "c:\\FolderA\FolderB\\FolderC", and the ending point, *wszTo*, be "c:\\FolderA\\FolderD\\FolderE". **AfxPathRelativePathTo** will return the relative path from *wszFrom* to *wszTo* as: "..\\..\\FolderD\\FolderE". You will get the same result if you set *wszFrom* to "\\FolderA\\FolderB\\FolderC" and *wszTo* to "\\FolderA\\FolderD\\FolderE". On the other hand, "c:\\FolderA\\FolderB" and "a:\\FolderA\\FolderD do not share a common prefix, and the function will fail. Note that "\\\\" is not considered a prefix and is ignored. If you set *wszFrom* to "\\\\FolderA\\FolderB", and *wszTo* to "\\\\FolderC\\FolderD", the function will fail.


# <a name="AfxPathRemoveArgs"></a>AfxPathRemoveArgs

Removes any arguments from a given path.

```
FUNCTION AfxPathRemoveArgs (BYREF wszPath AS CONST WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains the path from which to remove arguments. |

#### Return value

The changed path.

#### Remarks

This function should not be used on generic command path templates (from users or the registry), but rather it should be used only on templates that the application knows to be well formed.

# <a name="AfxPathRemoveBackslash"></a>AfxPathRemoveBackslash

Removes the trailing backslash from a given path.

```
FUNCTION AfxPathRemoveBackslash (BYREF wszPath AS CONST WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains the path from which to remove the backslash. |

#### Return value

The changed path.

# <a name="AfxPathRemoveBlanks"></a>AfxPathRemoveBlanks

Removes all leading and trailing spaces from a string.

```
FUNCTION AfxPathRemoveBlanks (BYREF wszPath AS CONST WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains the path from which to strip all leading and trailing spaces. |

#### Return value

The changed path.

# <a name="AfxPathRemoveExtension"></a>AfxPathRemoveExtension

Removes the file name extension from a path, if one is present.

```
FUNCTION AfxPathRemoveExtension (BYREF wszPath AS CONST WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains the path from which to remove the extension. |

#### Return value

The changed path.

# <a name="AfxPathRemoveFileSpec"></a>AfxPathRemoveFileSpec

Removes the trailing file name and backslash from a path, if they are present.

```
FUNCTION AfxPathRemoveFileSpec (BYREF wszPath AS CONST WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains the path from which to remove the file name. |

#### Return value

The changed path.

# <a name="AfxPathRenameExtension"></a>AfxPathRenameExtension

Replaces the extension of a file name with a new extension. If the file name does not contain an extension, the extension will be attached to the end of the string.

```
FUNCTION AfxPathRenameExtension (BYREF wszPath AS CONST WSTRING, BYREF wszExt AS CONST WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains the path in which to replace the extension. |
| *wszExt* | A string that contains a '.' character followed by the new extension. |

#### Return value

The new path. If the new path and extension would exceed MAX_PATH characters, the path won't be changed.

# <a name="AfxPathSearchAndQualify"></a>AfxPathSearchAndQualify

Determines if a given path is correctly formatted and fully qualified.

```
FUNCTION AfxPathSearchAndQualify (BYREF wszPath AS CONST WSTRING, BYREF wszFullyQualifiedPath AS CONST WSTRING, _
   BYVAL cchFullyQualifiedPath AS DWORD) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains the path to search. |
| *wszFullyQualifiedPath* | A string of length MAX_PATH that contains the path to be referenced. |
| *cchFullyQualifiedPath* | The size of the buffer pointed to by *wszFullyQualifiedPath*, in characters. |

#### Return value

Returns True if the path is qualified, or False otherwise.

# <a name="AfxPathSetDlgItemPath"></a>AfxPathSetDlgItemPath

Sets the text of a child control in a window or dialog box, using **AfxCompactPath** to ensure the path fits in the control.

```
SUB AfxPathSetDlgItemPath (BYVAL hDlg AS HWND, BYVAL cID AS LONG, BYREF wszPath AS WSTRING)
```

| Parameter  | Description |
| ---------- | ----------- |
| *hDlg* | A handle to the dialog box or window. |
| *cID* | The identifier of the control. |
| *wszPath* | A string that contains the path to set in the control. |

# <a name="AfxPathSkipRoot"></a>AfxPathSkipRoot

Parses a path, ignoring the drive letter or Universal Naming Convention (UNC) server/share path elements.

```
FUNCTION AfxPathSkipRoot (BYREF wszPath AS CONST WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains the path to parse. |

#### Return value

The changed path.

# <a name="AfxPathStripPath"></a>AfxPathStripPath

Removes the path portion of a fully qualified path and file.

```
FUNCTION AfxPathStripPath (BYREF wszPath AS CONST WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains the path and file name. When this method returns successfully, the string contains only the file name, with the path removed. |

#### Return value

The changed path.

# <a name="AfxPathStripToRoot"></a>AfxPathStripToRoot

Removes all parts of the path except for the root information.

```
FUNCTION AfxPathStripToRoot (BYREF wszRoot AS CONST WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszRoot* | A string of length MAX_PATH that contains the path to be converted. |

#### Return value

Returns a string that contains only the root information taken from that path.

# <a name="AfxPathUndecorate"></a>AfxPathUndecorate

Removes the decoration from a path string.

```
FUNCTION AfxPathUndecorate (BYREF wszPath AS CONST WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains the path. |

#### Return value

The undecorated string.

#### Remarks

A decoration consists of a pair of square brackets with one or more digits in between, inserted immediately after the base name and before the file name extension.

#### Example

The following table illustrates how strings are modified by **AfxPathUndecorate**.

| Initial String | Undecorated String |
| -------------- | ------------------ |
| C:\\Path\\File\[5].txt | C:\\Path\\File.txt |
| C:\\Path\\File\[12] | C:\\Path\\File |
| C:\\Path\\File.txt | C:\\Path\\File.txt |
| C:\\Path\\\[3].txt | C:\\Path\\\[3].txt |

# <a name="AfxPathUnExpandEnvStrings"></a>AfxPathUnExpandEnvStrings

Replaces certain folder names in a fully-qualified path with their associated environment string.

```
FUNCTION AfxPathUnExpandEnvStrings (BYREF wszPath AS CONST WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains the path to be unexpanded. |

#### Return value

The unexpanded string.

# <a name="AfxPathUnmakeSystemFolder"></a>AfxPathUnmakeSystemFolder

Removes the attributes from a folder that make it a system folder. This folder must actually exist in the file system.

```
FUNCTION AfxPathUnmakeSystemFolder (BYREF wszPath AS CONST WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains the name of an existing folder that will have the system folder attributes removed. |

#### Return value

Returns nonzero if successful, or zero otherwise.

# <a name="AfxPathUnquoteSpaces"></a>AfxPathUnquoteSpaces

Removes quotes from the beginning and end of a path.

```
FUNCTION AfxPathUnquoteSpaces (BYREF wszPath AS CONST WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains the path. |

#### Return value

The unquoted path.

# <a name="AfxUrlApplyScheme"></a>AfxUrlApplyScheme

Determines a scheme for a specified URL string, and returns a string with an appropriate prefix.

```
FUNCTION AfxUrlApplyScheme (BYREF wszUrl AS WSTRING, BYVAL dwFlags AS DWORD) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszUrl* | A string that contains a URL. |
| *dwFlags* | The flags that specify how to determine the scheme. The flags can be combined. |

| Flag       | Description |
| ---------- | ----------- |
| URL_APPLY_DEFAULT | Apply the default scheme if **AfxUrlApplyScheme** can't determine one. The default prefix is stored in the registry but is typically "http". |
| URL_APPLY_GUESSSCHEME | Attempt to determine the scheme by examining *wszUrl*. |
| URL_APPLY_GUESSFILE | Attempt to determine a file URL from *wszUrl*. |
| URL_APPLY_FORCEAPPLY | Force **AfxUrlApplyScheme** to determine a scheme for *wszUrl*. |
| URL_APPLY_FORCEAPPLY | Force **AfxUrlApplyScheme** to determine a scheme for *wszUrl*. |

#### Return value

The changed url.

#### Remarks

If the URL has a valid scheme, the string will not be modified. However, almost any combination of two or more characters followed by a colon will be parsed as a scheme. Valid characters include some common punctuation marks, such as ".". If your input string fits this description, **AfxUrlApplyScheme** may treat it as valid and not apply a scheme. To force the function to apply a scheme to a URL, set the URL_APPLY_FORCEAPPLY and URL_APPLY_DEFAULT flags in *dwFlags*. This combination of flags forces the function to apply a scheme to the URL. Typically, the function will not be able to determine a valid scheme. The second flag guarantees that, if no valid scheme can be determined, the function will apply the default scheme to the URL.

# <a name="AfxUrlCanonicalize"></a>AfxUrlCanonicalize

Converts a URL string into canonical form.

```
FUNCTION AfxUrlCanonicalize (BYREF wszUrl AS WSTRING, BYVAL dwFlags AS DWORD) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszUrl* | A string that contains a URL. If the string does not refer to a file, it must include a valid scheme such as "http://". |
| *dwFlags* | The flags that specify how the URL is converted to canonical form. The flags can be combined. |

| Flag       | Description |
| ---------- | ----------- |
| URL_UNESCAPE | Un-escape any escape sequences that the URLs contain, with two exceptions. The escape sequences for "?" and "#" are not un-escaped. If one of the URL_ESCAPE_XXX flags is also set, the two URLs are first un-escaped, then combined, then escaped. |
| URL_ESCAPE_UNSAFE | Replace unsafe characters with their escape sequences. Unsafe characters are those characters that may be altered during transport across the Internet, and include the (<, >, ", #, {, }, \|, \\, ^, \[, ], and ') characters. This flag applies to all URLs, including opaque URLs. |
| URL_PLUGGABLE_PROTOCOL | Combine URLs with client-defined pluggable protocols, according to the W3C specification. This flag does not apply to standard protocols such as ftp, http, gopher, and so on. If this flag is set, **AfxUrlCombine** does not simplify URLs, so there is no need to also set URL_DONT_SIMPLIFY. |
| URL_ESCAPE_SPACES_ONLY | Replace only spaces with escape sequences. This flag takes precedence over URL_ESCAPE_UNSAFE, but does not apply to opaque URLs. |
| URL_DONT_SIMPLIFY | Treat "/./" and "/../" in a URL string as literal characters, not as shorthand for navigation. See Remarks for further discussion. |
| URL_NO_META | Defined to be the same as URL_DONT_SIMPLIFY. |
| URL_ESCAPE_PERCENT | Convert any occurrence of "%" to its escape sequence. |
| URL_ESCAPE_AS_UTF8 | Windows 7 and later. Percent-encode all non-ASCII characters as their UTF-8 equivalents. |

#### Return value

The canonicalized url.

#### Remarks

This function performs such tasks as replacing unsafe characters with their escape sequences and collapsing sequences like "..\\...".

If a URL string contains "/../" or "/./", **AfxUrlCanonicalize** normally treats the characters as indicating navigation in the URL hierarchy. The function simplifies the URLs before combining them. For instance "/hello/cruel/../world" is simplified to "/hello/world". If the URL_DONT_SIMPLIFY flag is set in *dwFlags*, the function does not simplify URLs. In this case, "/hello/cruel/../world" is left as it is.

# <a name="AfxUrlCombine"></a>AfxUrlCombine

When provided with a relative URL and its base, returns a URL in canonical form.

```
FUNCTION AfxUrlCombine (BYREF wszBase AS WSTRING, BYREF wszRelative AS WSTRING, BYVAL dwFlags AS DWORD) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszBase* | A string that contains the base url. |
| *wszRelative* | A string that contains the relative url. |
| *dwFlags* | Flags that specify how the URL is converted to canonical form. The flags can be combined. |

| Flag       | Description |
| ---------- | ----------- |
| URL_DONT_SIMPLIFY | Treat "/./" and "/../" in a URL string as literal characters, not as shorthand for navigation. See Remarks for further discussion. |
| URL_ESCAPE_PERCENT | Convert any occurrence of "%" to its escape sequence. |
| URL_ESCAPE_SPACES_ONLY | Replace only spaces with escape sequences. This flag takes precedence over URL_ESCAPE_UNSAFE, but does not apply to opaque URLs. |
| URL_ESCAPE_UNSAFE | Replace unsafe characters with their escape sequences. Unsafe characters are those characters that may be altered during transport across the Internet, and include the (<, >, ", #, {, }, \|, \\, ^, \[, ], and ') characters. This flag applies to all URLs, including opaque URLs. |
| URL_NO_META | Defined to be the same as URL_DONT_SIMPLIFY. |
| URL_PLUGGABLE_PROTOCOL | Combine URLs with client-defined pluggable protocols, according to the W3C specification. This flag does not apply to standard protocols such as ftp, http, gopher, and so on. If this flag is set, **AfxUrlCombine** does not simplify URLs, so there is no need to also set URL_DONT_SIMPLIFY. |
| URL_UNESCAPE | Un-escape any escape sequences that the URLs contain, with two exceptions. The escape sequences for "?" and "#" are not un-escaped. If one of the URL_ESCAPE_XXX flags is also set, the two URLs are first un-escaped, then combined, then escaped. |
| URL_ESCAPE_AS_UTF8 | Windows 7 and later. Percent-encode all non-ASCII characters as their UTF-8 equivalents. |

#### Return value

The canonicalized url.

#### Remarks

Items between slashes are treated as hierarchical identifiers; the last item specifies the document itself. You must enter a slash (/) after the document name to append more items; otherwise, **AfxUrlCombine** changes one document for another.

If a URL string contains '/../' or '/./', **AfxUrlCombine** usually treats the characters as if they indicated navigation in the URL hierarchy. The function simplifies the URLs before combining them. For instance, "/hello/cruel/../world" is simplified to "/hello/world". If the URL_DONT_SIMPLIFY flag is set in dwFlags, the function does not simplify URLs. In this case, "/hello/cruel/../world" is left as it is.

# <a name="AfxUrlCompare"></a>AfxUrlCompare

Makes a case-sensitive comparison of two URL strings.

```
FUNCTION AfxUrlCompare (BYREF wszUrl1 AS CONST WSTRING, BYREF wszUrl2 AS CONST WSTRING, _
   BYVAL fIgnoreSlash AS LONG) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszUrl1* | A string that contains the first URL. |
| *wszUrl2* | A string that contains the second URL. |
| *fIgnoreSlash* |  A value that is set to TRUE to have **AfxUrlCompare** ignore a trailing '/' character on either or both URLs. |

#### Return value

Returns zero if the two strings are equal. The function will also return zero if *fIgnoreSlash* is set to TRUE and one of the strings has a trailing '\\' character. The function returns a negative integer if the string pointed to by *wszUrl1* is less than the string pointed to by *wszUrl2*. Otherwise, it returns a positive integer.

#### Remarks

For best results, you should first canonicalize the URLs with **AfxUrlCanonicalize**. Then, compare the canonicalized URLs with **AfxUrlCompare**.

# <a name="AfxUrlCreateFromPath"></a>AfxUrlCreateFromPath

Converts a Microsoft MS-DOS path to a canonicalized URL.

```
FUNCTION AfxUrlCreateFromPath (BYREF wszPath AS WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains the MS-DOS path. |

#### Return value

The canonicalized URL.

# <a name="AfxUrlEscape"></a>AfxUrlEscape

Converts characters in a URL that might be altered during transport across the Internet ("unsafe" characters) into their corresponding escape sequences.

```
FUNCTION AfxUrlEscape (BYREF wszUrl AS WSTRING, BYVAL dwFlags AS DWORD) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszUrl* | A string that contains the MS-DOS path. |
| *dwFlags* | The flags that indicate which portion of the URL is being provided in *wszURL* and which characters in that string should be converted to their escape sequences. |

| Flag       | Description |
| ---------- | ----------- |
| URL_DONT_ESCAPE_EXTRA_INFO | Used only in conjunction with URL_ESCAPE_SPACES_ONLY to prevent the conversion of characters in the query (the portion of the URL following the first # or ? character encountered in the string). This flag should not be used alone, nor combined with URL_ESCAPE_SEGMENT_ONLY. |
| URL_BROWSER_MODE | Defined to be the same as URL_DONT_ESCAPE_EXTRA_INFO. |
| URL_ESCAPE_SPACES_ONLY | Convert only space characters to their escape sequences, including those space characters in the query portion of the URL. Other unsafe characters are not converted to their escape sequences. This flag assumes that *wszURL* does not contain a full URL. It expects only the portions following the server specification.<br>Combine this flag with URL_DONT_ESCAPE_EXTRA_INFO to prevent the conversion of space characters in the query portion of the URL.<br>This flag cannot be combined with URL_ESCAPE_PERCENT or URL_ESCAPE_SEGMENT_ONLY. |
| URL_ESCAPE_PERCENT | Convert any % character found in the segment section of the URL (that section falling between the server specification and the first # or ? character). By default, the % character is not converted to its escape sequence. Other unsafe characters in the segment are also converted normally.<br>Combining this flag with URL_ESCAPE_SEGMENT_ONLY includes those % characters in the query portion of the URL. However, as the URL_ESCAPE_SEGMENT_ONLY flag causes the entire string to be considered the segment, any # or ? characters are also converted.<br>This flag cannot be combined with URL_ESCAPE_SPACES_ONLY. |
| URL_ESCAPE_SEGMENT_ONLY | Indicates that *wszURL* contains only that section of the URL following the server component but preceding the query. All unsafe characters in the string are converted. If a full URL is provided when this flag is set, all unsafe characters in the entire string are converted, including # and ? characters.<br>Combine this flag with URL_ESCAPE_PERCENT to include that character in the conversion.<br>This flag cannot be combined with URL_ESCAPE_SPACES_ONLY or URL_DONT_ESCAPE_EXTRA_INFO. |
| URL_ESCAPE_AS_UTF8 | Windows 7 and later. Percent-encode all non-ASCII characters as their UTF-8 equivalents. |

#### Return value

The converted URL.

#### Remarks

For the purposes of this document, a typical URL is divided into three sections: the server, the segment, and the query. For example:

http://microsoft.com/test.asp?url=/example/abc.asp?frame=true#fragment

The server portion is "http://microsoft.com/". The trailing forward slash is considered part of the server portion.

The segment portion is any part of the path found following the server portion, but before the first # or ? character, in this case simply "test.asp".

The query portion is the remainder of the path from the first # or ? character (inclusive) to the end. In the example, it is "?url=/example/abc.asp?frame=true#fragment".

Unsafe characters are those characters that might be altered during transport across the Internet. This function converts unsafe characters into their equivalent "%xy" escape sequences. The following table shows unsafe characters and their escape sequences.

| Character  | Escape Sequence |
| ---------- | ----------- |
| ^ | %5E5E |
| & | %26 |
| ` | %60 |
| { | %7B |
| } | %7D |
| \| | %7C |
| ] | %5D |
| \[ | %5B |
| " | %22 |
| < | %3C |
| > | %3E |
| \ | %5C |


Use of the URL_ESCAPE_SEGMENT_ONLY flag also causes the conversion of the # (%23), ? (%3F), and / (%2F) characters.

By default, AfxUrlEscape ignores any text following a # or ? character. The URL_ESCAPE_SEGMENT_ONLY flag overrides this behavior by regarding the entire string as the segment. The URL_ESCAPE_SPACES_ONLY flag overrides this behavior, but only for space characters.

Security Warning: This function should be called from client applications only.

# <a name="AfxUrlEscapeSpaces"></a>AfxUrlEscapeSpaces

Converts space characters into their corresponding escape sequence.

```
FUNCTION AfxUrlEscapeSpaces (BYREF wszUrl AS WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszUrl* | A string that contains the URL to be corrected. This string must not exceed INTERNET_MAX_PATH_LENGTH characters in length, including the terminating NULL character. |

#### Return value

The converted URL.

# <a name="AfxUrlFixup"></a>AfxUrlFixup

Attempts to correct a URL whose protocol identifier is incorrect. For example, htttp will be changed to http.

```
FUNCTION AfxUrlFixup (BYREF wszUrl AS WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszUrl* | A string that contains the URL to be corrected. This string must not exceed INTERNET_MAX_PATH_LENGTH characters in length, including the terminating NULL character. |

#### Return value

The converted URL.

#### Remarks

The **AfxUrlFixup** function recognizes the schemes specified by the URL_SCHEME enumeration.

Priority is given to the first character in the protocol identifier section so htp will be converted to http instead of ftp.

**Notes**

Do not use this function for deterministic data transformation. The heuristics used by **AfxUrlFixup** can change from one release to the next. The function should only be used to correct possibly invalid user input.

This function is available through Windows 7 and Windows Server 2008. It might be altered or unavailable in subsequent versions of Windows.

# <a name="AfxUrlGetLocation"></a>AfxUrlGetLocation

Retrieves the location from a URL.

```
FUNCTION AfxUrlGetLocation (BYREF wszUrl AS CONST WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszUrl* | A string that contains the location. This string must not exceed INTERNET_MAX_PATH_LENGTH characters in length, including the terminating NULL character. |

#### Return value

A string with the location or an empty string.

#### Remarks

The location is the segment of the URL starting with a ? or # character. If a file URL has a query string, the returned string includes the query string.

# <a name="AfxUrlGetPart"></a>AfxUrlGetPart

Accepts a URL string and returns a specified part of that URL.

```
FUNCTION AfxUrlGetPart (BYREF wszUrl AS WSTRING, _
   BYVAL dwPart AS DWORD, BYVAL dwFlags AS DWORD) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszUrl* | A string that contains the URL. This string must not exceed INTERNET_MAX_PATH_LENGTH characters in length, including the terminating NULL character. |
| *dwPart* | The flags that specify which part of the URL to retrieve. It can have one of the following values.<br>*URL_PART_HOSTNAME*: The host name.<br>*URL_PART_PASSWORD*: The password.<br>*URL_PART_PORT*: The port number.<br>*URL_PART_QUERY*: The query portion of the URL.<br>*URL_PART_SCHEME*: The URL scheme.<br>*URL_PART_USERNAME*: The username. |
| *dwFlags* | A flag that can be set to keep the URL scheme, in addition to the part that is specified by *dwPart*.<br>*URL_PARTFLAG_KEEPSCHEME*: Keep the URL scheme. |

#### Return value

A string with the specified part of that URL.

# <a name="AfxUrlHash"></a>AfxUrlHash

Hashes a URL string.

```
FUNCTION AfxUrlHash (BYREF wszUrl AS WSTRING, BYVAL pbHash AS BYTE PTR, _
   BYVAL cbHash AS DWORD) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszUrl* | A string that contains the URL. This string must not exceed INTERNET_MAX_PATH_LENGTH characters in length, including the terminating NULL character. |
| *pbHash* | A pointer to a buffer that, When this method returns successfully, receives the hashed array. |
| *cbHash* | The number of elements in the array at *pbHash*. It should be no larger than 256. |

#### Return value

If the function succeeds, it returns S_OK (0). Otherwise, it returns an HRESULT error code.

#### Remarks

To hash a URL into a single byte, set cbHash = SIZEOF(BYTE) and pbHash = VARPTR(bHashedValue), where bHashedValue is a one-byte buffer.

To hash a URL into a DWORD, set cbHash = SIZEOF(DWORD) and pbHash = VARPTR(dwHashedValue), where dwHashedValue is a DWORD buffer.

# <a name="AfxUrlIs"></a>AfxUrlIs

Tests whether or not a URL is a specified type.

```
FUNCTION AfxUrlIs (BYREF wszUrl AS WSTRING, BYVAL nUrlIs AS URLIS) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszUrl* | A string that contains the URL. This string must not exceed INTERNET_MAX_PATH_LENGTH characters in length, including the terminating NULL character. |
| *nUrlIs* | The type of URL to be tested for. This parameter can take one of the following values:<br>*URLIS_APPLIABLE*: Attempt to determine a valid scheme for the URL.<br>*URLIS_DIRECTORY*: Does the URL string end with a directory?<br>*URLIS_FILEURL*: Is the URL a file URL?<br>*URLIS_HASQUERY*: Does the URL have an appended query string?<br>*URLIS_NOHISTORY*: Is the URL a URL that is not typically tracked in navigation history?<br>*URLIS_OPAQUE*: Is the URL opaque?<br>*URLIS_URL*: Is the URL valid? |

#### Return value

For all but one of the URL types, **AfxUrlIs** returns True if the URL is the specified type, or False if not.

If **AfxUrlIs** is set to URLIS_APPLIABLE, **AfxUrlIs** will attempt to determine the URL scheme. If the function is able to determine a scheme, it returns True, or False otherwise.

# <a name="AfxUrlIsFileUrl"></a>AfxUrlIsFileUrl

Tests a URL to determine if it is a file URL.

```
FUNCTION AfxUrlIsFileUrl (BYREF wszUrl AS WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszUrl* | A string that contains the URL to be corrected. This string must not exceed INTERNET_MAX_PATH_LENGTH characters in length, including the terminating NULL character. |

#### Return value

Returns True if the URL is a file URL, or False otherwise.

#### Remarks

A file URL has the form "File:// xxx".

# <a name="AfxUrlIsNoHistory"></a>AfxUrlIsNoHistory

Returns whether a URL is a URL that browsers typically do not include in navigation history.

```
FUNCTION AfxUrlIsNoHistory (BYREF wszUrl AS CONST WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszUrl* | A string that contains the URL to be corrected. This string must not exceed INTERNET_MAX_PATH_LENGTH characters in length, including the terminating NULL character. |

#### Return value

Returns a True value if the URL is a URL that is not included in navigation history, or False otherwise.

#### Remarks

This function is equivalent to the following: AfxUrlIs(wszURL, URLIS_NOHISTORY)

# <a name="AfxUrlIsOpaque"></a>AfxUrlIsOpaque

Returns whether a URL is opaque.

```
FUNCTION AfxUrlIsOpaque (BYVAL wszUrl AS CONST WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszUrl* | A string that contains the URL to be corrected. This string must not exceed INTERNET_MAX_PATH_LENGTH characters in length, including the terminating NULL character. |

#### Return value

Returns True if the URL is opaque, or False otherwise.

#### Remarks

A URL that has a scheme that is not followed by two slashes (//) is opaque. For example, mailto:xyz@litwareinc.com is an opaque URL. Opaque URLs cannot be separated into the standard URL hierarchy. **AfxUrlIsOpaque** is equivalent to the following: AfxUrlIs(wszURL, URLIS_OPAQUE)

# <a name="AfxUrlUnescape"></a>AfxUrlUnescape

Converts escape sequences back into ordinary characters.

```
FUNCTION AfxUrlUnescape (BYREF wszUrl AS WSTRING, BYVAL dwFlags AS DWORD) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszUrl* | A string that contains the URL. This string must not exceed INTERNET_MAX_PATH_LENGTH characters in length, including the terminating NULL character. |
| *dwFlags* | Flags that control which characters are unescaped. It can be the following flag:<br>*URL_DONT_UNESCAPE_EXTRA_INFO*: Do not convert the # or ? character, or any characters following them in the string. |

#### Return value

Returns S_OK (0) if successful or an HRESULT otherwise.

#### Remarks

An escape sequence has the form "%xy".

Input strings cannot be longer than INTERNET_MAX_URL_LENGTH.

# <a name="AfxUrlUnescapeInPlace"></a>AfxUrlUnescapeInPlace

Converts escape sequences back into ordinary characters and overwrites the original string.

```
FUNCTION AfxUrlUnescapeInPlace (BYREF wszUrl AS WSTRING, BYVAL dwFlags AS DWORD) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszUrl* | A string that contains the URL. This string must not exceed INTERNET_MAX_PATH_LENGTH characters in length, including the terminating NULL character. The converted string is returned through this parameter. |
| *dwFlags* | Flags that control which characters are unescaped. It can be the following flag:<br>*URL_DONT_UNESCAPE_EXTRA_INFO*: Do not convert the # or ? character, or any characters following them in the string. |

#### Return value

Returns S_OK (0) if successful or an HRESULT otherwise.
