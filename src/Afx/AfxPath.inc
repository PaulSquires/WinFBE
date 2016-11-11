' ########################################################################################
' Microsoft Windows
' File: AfxPath.inc
' Contents: Path wrapper functions.
' Compiler: FreeBasic 32 & 64-bit, Unicode.
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "windows.bi"
#include once "win/shlwapi.bi"
#include once "Afx/CWStr.inc"
USING Afx

' // To avoid to include wininet.bi
#ifndef _WININET_
const INTERNET_MAX_PATH_LENGTH = 2048
const INTERNET_MAX_SCHEME_LENGTH = 32
#define INTERNET_MAX_URL_LENGTH ((INTERNET_MAX_SCHEME_LENGTH + sizeof("://")) + INTERNET_MAX_PATH_LENGTH)
#endif

' ========================================================================================
' Adds a backslash to the end of a string to create the correct syntax for a path.
' If the source path already has a trailing backslash, no backslash will be added.
' Usage example: DIM cws AS CWSTR = AfxPathAddBackSlash("C:\dir_name\dir_name\file_name")
' ========================================================================================
PRIVATE FUNCTION AfxPathAddBackSlash (BYREF wszPath AS CONST WSTRING) AS CWSTR
   DIM _wszPath AS WSTRING * MAX_PATH = wszPath
   PathAddBackSlashW(_wszPath)
   RETURN _wszPath
END FUNCTION
' ========================================================================================

' ========================================================================================
' Removes the trailing backslash from a given path.
' Usage eample: DIM cws AS CWSTR = AfxPathRemoveBackslash("c:\a\b\File\")
' ========================================================================================
PRIVATE FUNCTION AfxPathRemoveBackslash (BYREF wszPath AS CONST WSTRING) AS CWSTR
   DIM _wszPath AS WSTRING * MAX_PATH = wszPath
   PathRemoveBackslashW(_wszPath)
   RETURN _wszPath
END FUNCTION
' ========================================================================================

' ========================================================================================
' Adds a file extension to a path string.
' If there is already a file name extension present, no extension will be added.
' If wszPath is an empt string, the result will be the file name extension only.
' If pwszExt points to a NULL string, an ".exe" extension will be added.
' Example: DIM cws AS CWSTR = AfxPathAddExtension("file") - output: file.exe
' Example: DIM cws AS CWSTR = AfxPathAddExtension("file.doc") - output: file.doc
' Example: DIM cws AS CWSTR = AfxPathAddExtension("file", ".txt") - output: file.txt
' Example: DIM cws AS CWSTR = AfxPathAddExtension("", ".txt") - output: .txt
' ========================================================================================
PRIVATE FUNCTION AfxPathAddExtension (BYREF wszPath AS CONST WSTRING, BYVAL pwszExt AS WSTRING PTR = NULL) AS CWSTR
   DIM _wszPath AS WSTRING * MAX_PATH = wszPath
   PathAddExtensionW(_wszPath, pwszExt)
   RETURN _wszPath
END FUNCTION
' ========================================================================================

' ========================================================================================
' Removes the file extension from a path, if there is one.
' Usage example: DIM cws AS CWSTR = AfxPathRemoveExtension("C:\TEST\sample.txt")
' ========================================================================================
PRIVATE FUNCTION AfxPathRemoveExtension (BYREF wszPath AS CONST WSTRING) AS CWSTR
   DIM _wszPath AS WSTRING * MAX_PATH = wszPath
   PathRemoveExtensionW(_wszPath)
   RETURN _wszPath
END FUNCTION
' ========================================================================================

' ========================================================================================
' Replaces the extension of a file name with a new extension. If the file name does not
' contain an extension, the extension will be attached to the end of the string.
' DIM cws AS CWSTR = AfxPathRenameExtension("C:\TEST\sample.bas", ".exe")
' DIM cws AS CWSTR = AfxPathRenameExtension("C:\TEST\sample", ".exe")
' ========================================================================================
PRIVATE FUNCTION AfxPathRenameExtension (BYREF wszPath AS CONST WSTRING, BYREF wszExt AS CONST WSTRING) AS CWSTR
   DIM _wszPath AS WSTRING * MAX_PATH = wszPath
   PathRenameExtensionW(_wszPath, wszExt)
   RETURN _wszPath
END FUNCTION
' ========================================================================================

' ========================================================================================
' Appends one path to the end of another.
' This function automatically inserts a backslash between the two strings, if one is not already present.
' The path supplied in wszPath cannot begin with "..\" or ".\" to produce a relative path
' string. If present, those periods are stripped from the output string. For example,
' appending "path3" to "..\path1\path2" results in an output of "\path1\path2\path3" rather
' than "..\path1\path2\path3".
' Usage exmaple: DIM cws AS CWSTR = AfxPathAppend("name_1\name_2", "name_3")
' ========================================================================================
PRIVATE FUNCTION AfxPathAppend (BYREF wszPath AS CONST WSTRING, BYREF wszMore AS CONST WSTRING) AS CWSTR
   DIM _wszPath AS WSTRING * MAX_PATH = wszPath
   PathAppendW(_wszPath, wszMore)
   RETURN _wszPath
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates a root path from a given drive number.
' nDrive = The desired drive number. It should be between 0 and 25.
' 0 = A, 1 = B, ..., Z = 25.
' Usage example: DIM cws AS CWSTR = AfxPathBuildRoot(2) - output: C:\
' ========================================================================================
PRIVATE FUNCTION AfxPathBuildRoot (BYVAL nDrive AS LONG) AS CWSTR
   DIM _wszRoot AS WSTRING * 4
   PathBuildRootW(_wszRoot, nDrive)
   RETURN _wszRoot
END FUNCTION
' ========================================================================================

' ========================================================================================
' Canonicalizes a path.
' This function allows the user to specify what to remove from a path by inserting special
' character sequences into the path. The ".." sequence indicates to remove a path segment
' from the current position to the previous path segment. The "." sequence indicates to skip
' over the next path segment to the following path segment. The root segment of the path
' cannot be removed.
' If there are more ".." sequences than there are path segments, the function returns just the root.
' Examples:
' DIM cws AS CWSTR = AfxPathCanonicalize("A:\name_1\.\name_2\..\sname_3") - output: A:\name_1\name_3
' DIM cws AS CWSTR = AfxPathCanonicalize("A:\name_1\..\name_2\.\name_3") - output: A:\name_2\name_3
' DIM cws AS CWSTR = AfxPathCanonicalize("A:\name_1\name_2\.\name_3\..\name_4") - output: A:\name_1\name_2\name_4
' DIM cws AS CWSTR = AfxPathCanonicalize("A:\name_1\.\name_2\.\name_3\..\name_4\..") - output: A:\name_1\name_2
' DIM cws AS CWSTR = AfxPathCanonicalize("C:\..") - output: C:\
' ========================================================================================
PRIVATE FUNCTION AfxPathCanonicalize (BYREF wszPath AS CONST WSTRING) AS CWSTR
   DIM _wszDest AS WSTRING * MAX_PATH
   PathCanonicalizeW(_wszDest, wszPath)
   RETURN _wszDest
END FUNCTION
' ========================================================================================

' ========================================================================================
' Concatenates two strings that represent properly formed paths into one path, as well as
' any relative path pieces. The directory path should be in the form of A:,B:, ..., Z:.
' The file path should be in a correct form that represents the file name part of the path.
' If the directory path ends with a backslash, the backslash will be maintained.
' Note that while wszDir and wszFile are both optional parameters, they cannot both be NULL.
' Example: DIM cws AS CWSTR = AfxPathCombine("C:", "One\Two\Three") - output: C:\One\Two\Three
' ========================================================================================
PRIVATE FUNCTION AfxPathCombine (BYREF wszDir AS CONST WSTRING, BYREF wszFile AS CONST WSTRING) AS CWSTR
   DIM _wszDest AS WSTRING * MAX_PATH
   PathCombineW(_wszDest, wszDir, wszFile)
   RETURN _wszDest
END FUNCTION
' ========================================================================================

' ========================================================================================
' Compares two paths to determine if they share a prefix and returns a string with the prefix
' characters, if any. A prefix is one of these types: "C:\", ".", "..", "..\".
' Example: DIM cws AS CWSTR = AfxPathCommonPrefix("C:\win\desktop\temp.txt", "c:\win\tray\sample.txt")
' Output: C:\win
' ========================================================================================
PRIVATE FUNCTION AfxPathCommonPrefix (BYREF wszPath1 AS CONST WSTRING, BYREF wszPath2 AS CONST WSTRING) AS CWSTR
   DIM _wszPath AS WSTRING * MAX_PATH
   PathCommonPrefixW(wszPath1, wszPath2, _wszPath)
   RETURN _wszPath
END FUNCTION
' ========================================================================================

' ========================================================================================
' Truncates a file path to fit within a given pixel width by replacing path components with ellipses.
' This function uses the font currently selected in hDC to calculate the width of the text.
' This function will not compact the path beyond the base file name preceded by ellipses.
' - hdc : A handle to the device context used for font metrics. This value can be NULL.
' - wszPath : The path to be modified.
' - dx : The width, in pixels, in which the string must fit.
' Usage example: DIM cws AS CWSTR = AfxPathCompactPath(NULL, "C:\path1\path2\sample.txt", 180)
' ========================================================================================
PRIVATE FUNCTION AfxPathCompactPath (BYVAL hDC AS HDC, BYREF wszPath AS CONST WSTRING, BYVAL dx AS DWORD) AS CWSTR
   DIM _wszPath AS WSTRING * MAX_PATH = wszPath
   PathCompactPathW(hDC, _wszPath, dx)
   RETURN _wszPath
END FUNCTION
' ========================================================================================

' ========================================================================================
' Truncates a path to fit within a certain number of characters by replacing path components with ellipses.
' - wszPath : The path to be modified.
' - cchMax : The maximum number of characters to be contained in the new string.
' The '/' separator will be used instead of '\' if the original string used it. If wszPath
' points to a file name that is too long, instead of a path, the file name will be truncated
' to cchMax characters. For example, if the input file name is "My Filename" and cchMax is 10,
' PathCompactPathEx will return "My File...".
' Usage example: DIM cws AS CWSTR = AfxPathCompactPathEx("c:\test\My Filename", 18)
' ========================================================================================
PRIVATE FUNCTION AfxPathCompactPathEx (BYREF wszPath AS CONST WSTRING, BYVAL cchMax AS DWORD) AS CWSTR
   DIM _wszOut AS WSTRING * MAX_PATH
   IF cchMax > MAX_PATH THEN cchMax = MAX_PATH - 1
   PathCompactPathExW(_wszOut, wszPath, cchMax + 1, 0)
   RETURN LEFT(_wszOut, cchMax)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Takes a file URL and converts it to a Microsoft MS-DOS path.
' - wszUrl : A null-terminated string of maximum length INTERNET_MAX_URL_LENGTH that contains the URL.
' Example: DIM cws AS CWSTR = AfxPathCreateFromUrl("file:///C:/Documents%20and%20Settings/URIS.txt")
' Output: C:\Documents and Settings\URIS.txt
' ========================================================================================
PRIVATE FUNCTION AfxPathCreateFromUrl (BYREF wszUrl AS CONST WSTRING) AS CWSTR
   DIM _wszPath AS WSTRING * MAX_PATH
   DIM cchPath AS DWORD = SIZEOF(_wszPath)
   PathCreateFromUrlW(wszUrl, _wszPath, @cchPath, 0)
   RETURN _wszPath
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts a path to all lowercase characters to give the path a consistent appearance.
' This function only operates on paths that are entirely uppercase. For example:
' C:\WINDOWS will be converted to c:\windows, but c:\Windows will not be changed.
' Example: DIM cws AS CWSTR = AfxPathMakePretty("C:\TEST\FILE")
' ========================================================================================
PRIVATE FUNCTION AfxPathMakePretty (BYREF wszPath AS CONST WSTRING) AS CWSTR
   DIM _wszPath AS WSTRING * MAX_PATH = wszPath
   PathMakePrettyW(_wszPath)
   RETURN _wszPath
END FUNCTION
' ========================================================================================

' ========================================================================================
' Searches a path for spaces. If spaces are found, the entire path is enclosed in
' quotation marks.
' Example: DIM cws AS CWSTR = AfxPathQuoteSpaces("C:\sample_one\sample two")
' Output: "C:\sample_one\sample two"
' ========================================================================================
PRIVATE FUNCTION AfxPathQuoteSpaces (BYREF wszPath AS CONST WSTRING) AS CWSTR
   DIM _wszPath AS WSTRING * MAX_PATH = wszPath
   PathQuoteSpacesW(_wszPath)
   RETURN _wszPath
END FUNCTION
' ========================================================================================

' ========================================================================================
' Removes quotes from the beginning and end of a path.
' Usage example: DIM cws AS CWSTR = AfxPathUnquoteSpaces("""C:\sample_one\sample two""")
' Output: C:\sample_one\sample two
' ========================================================================================
PRIVATE FUNCTION AfxPathUnquoteSpaces (BYREF wszPath AS CONST WSTRING) AS CWSTR
   DIM _wszPath AS WSTRING * MAX_PATH = wszPath
   PathUnquoteSpacesW(_wszPath)
   RETURN _wszPath
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines whether a path to a file system object such as a file or directory is valid.
' TRUE if the file exists; otherwise, FALSE. Call GetLastError for extended error information.
' This function tests the validity of the path. A path specified by Universal Naming
' Convention (UNC) is limited to a file only; that is, \\server\share\file is permitted.
' A UNC path to a server or server share is not permitted; that is, \\server or \\server\share.
' This function returns FALSE if a mounted remote drive is out of service.
' Example: DIM b AS BOOLEAN = AfxPathFileExists("C:\TEST\file.txt")
' ========================================================================================
PRIVATE FUNCTION AfxPathFileExists (BYREF wszPath AS CONST WSTRING) AS BOOLEAN
   RETURN PathFileExistsW(wszPath)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Searches a path for an extension.
' wszPath : The path to search, including the extension being searched for.
' Note that a valid file name extension cannot contain a space.
' Usage example: DIM cws AS CWSTR = AfxPathFindExtension("C:\TEST\filetxt")
' ========================================================================================
PRIVATE FUNCTION AfxPathFindExtension (BYREF wszPath AS CONST WSTRING) AS CWSTR
   DIM _wszPath AS WSTRING * MAX_PATH = wszPath
   DIM pFileExt AS WSTRING PTR = PathFindExtensionW(_wszPath)
   IF pFileExt THEN *cast(WSTRING PTR, @_wszPath) = *pFileExt
   RETURN _wszPath
END FUNCTION
' ========================================================================================

' ========================================================================================
' Searches a path for a file name.
' ========================================================================================
PRIVATE FUNCTION AfxPathFindFileName (BYREF wszPath AS CONST WSTRING) AS CWSTR
   DIM _wszPath AS WSTRING * MAX_PATH = wszPath
   DIM pFileName AS WSTRING PTR = PathFindFileNameW(_wszPath)
   IF pFileName THEN *cast(WSTRING PTR, @_wszPath) = *pFileName
   RETURN _wszPath
END FUNCTION
' ========================================================================================

' ========================================================================================
' Parses a path and returns the portion of that path that follows the first backslash.
' PathFindNextComponent walks a path string until it encounters a backslash ("\\"), ignores
' everything up to that point including the backslash, and returns the rest of the path.
' Therefore, if a path begins with a backslash (such as \path1\path2), the function simply
' removes the initial backslash and returns the rest (path1\path2).
' Usage example: DIM cws AS CWSTR = PathFindNextComponentW("C:\TEST\file.txt")
' ========================================================================================
PRIVATE FUNCTION AfxPathFindNextComponent (BYREF wszPath AS WSTRING) AS CWSTR
   DIM _wszPath AS WSTRING * MAX_PATH = wszPath
   DIM pwszComp AS WSTRING PTR = PathFindNextComponentW(_wszPath)
   IF pwszComp THEN *cast(WSTRING PTR, @_wszPath) = *pwszComp
   RETURN _wszPath
END FUNCTION
' ========================================================================================

' ========================================================================================
' Parses a path, ignoring the drive letter or Universal Naming Convention (UNC)
' server/share path parts.
' ========================================================================================
PRIVATE FUNCTION AfxPathSkipRoot (BYREF wszPath AS CONST WSTRING) AS CWSTR
   DIM _wszPath AS WSTRING * MAX_PATH = wszPath
   DIM pwszPath AS WSTRING PTR = PathSkipRootW(_wszPath)
   IF pwszPath THEN *cast(WSTRING PTR, @_wszPath) = *pwszPath
   RETURN _wszPath
END FUNCTION
' ========================================================================================

' ========================================================================================
' *** UNTESTED ***
' Determines if a given file name has one of a list of suffixes.
' - wszPath : The file name to be tested. A full path can be used.
' - apwszSuffix : An array of iArraySize string pointers. Each string pointed to is
'   null-terminated and contains one suffix. The strings can be of variable lengths.
' - iArraySize : The number of elements in the array pointed to by apszSuffix.
' This function uses a case-sensitive comparison. The suffix must match exactly.
' ========================================================================================
PRIVATE FUNCTION AfxPathFindSuffixArray (BYREF wszPath AS WSTRING, BYVAL apwszSuffix AS WSTRING PTR PTR, BYVAL iArraySize AS LONG) AS CWSTR
   DIM pwszSuffix AS WSTRING PTR = cast(WSTRING PTR, PathFindSuffixArrayW(wszPath, apwszSuffix, iArraySize))
   IF pwszSuffix THEN RETURN *pwszSuffix
END FUNCTION
' ========================================================================================

' ========================================================================================
' Finds the command line arguments within a given path.
' This function should not be used on generic command path templates (from users or the
' registry), but rather should be used only on templates that the application knows to be
' well formed.
' Example: DIM cws AS CWSTR = AfxPathGetArgs("test.exe temp.txt sample.doc")
' ========================================================================================
PRIVATE FUNCTION AfxPathGetArgs (BYREF wszPath AS CONST WSTRING) AS CWSTR
   DIM pwszArgs AS WSTRING PTR = PathGetArgsW(wszPath)
   IF pwszArgs THEN RETURN *pwszArgs
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines the type of character with respect to a path.
' Returns one or more of the following values that define the type of character.
' GCT_INVALID   : The character is not valid in a path.
' GCT_LFNCHAR   : The character is valid in a long file name.
' GCT_SEPARATOR : The character is a path separator.
' GCT_SHORTCHAR : The character is valid in a short (8.3) file name.
' GCT_WILD      : The character is a wildcard character.
' ========================================================================================
PRIVATE FUNCTION AfxPathGetCharType (BYREF wszCh AS CONST WSTRING) AS UINT
   RETURN PathGetCharTypeW(ASC(wszCh))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Searches a path for a drive letter within the range of 'A' to 'Z' and returns the
' corresponding drive number, 0 through 25 (corresponding to 'A' through 'Z') if the path
' has a drive letter, or -1 otherwise.
' Example: DIM n AS LONG = AfxPathGetDriveNumber("C:\TEST\bar.txt") - output: 2
' ========================================================================================
PRIVATE FUNCTION AfxPathGetDriveNumber (BYREF wszPath AS CONST WSTRING) AS LONG
   RETURN PathGetDriveNumberW(wszPath)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines if a file's registered content type matches the specified content type.
' This function obtains the content type for the specified file type and compares
' that string with the wszContentType The comparison is not case sensitive.
' Returns nonzero if the file's registered content type matches wszContentType, or zero otherwise.
' Example: DIM b AS BOOLEAN = AfxPathIsContentType("C:\TEST\bar.txt", "text/plain") - output: true
' Example: DIM b AS BOOLEAN = AfxPathIsContentType("C:\TEST\bar.txt", "image/gif") - output: false
' ========================================================================================
PRIVATE FUNCTION AfxPathIsContentType (BYREF wszPath AS CONST WSTRING, BYREF wszContentType AS CONST WSTRING) AS BOOLEAN
   RETURN PathIsContentTypeW(wszPath, wszContentType)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Verifies that a path is a valid directory.
' ========================================================================================
PRIVATE FUNCTION AfxPathIsDirectory (BYREF wszPath AS CONST WSTRING) AS BOOLEAN
   RETURN PathIsDirectoryW(wszPath)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines whether or not a specified path is an empty directory.
' ========================================================================================
PRIVATE FUNCTION AfxPathIsDirectoryEmpty (BYREF wszPath AS CONST WSTRING) AS BOOLEAN
   RETURN PathIsDirectoryEmptyW(wszPath)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Searches a path for any path delimiting characters (for example, ':' or '\' ).
' If there are no path delimiting characters present, the path is considered to be
' a File Spec path.
' ========================================================================================
PRIVATE FUNCTION AfxPathIsFileSpec (BYREF wszPath AS CONST WSTRING) AS BOOLEAN
   RETURN PathIsFileSpecW(wszPath)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines if a file is an HTML file. The determination is made based on the
' content type that is registered for the file's extension.
' ========================================================================================
PRIVATE FUNCTION AfxPathIsHTMLFile (BYREF wszFile AS CONST WSTRING) AS BOOLEAN
   RETURN PathIsHTMLFileW(wszFile)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines whether or not a file name is in long format.
' ========================================================================================
PRIVATE FUNCTION AfxPathIsLFNFileSpec (BYREF wszName AS CONST WSTRING) AS BOOLEAN
   RETURN PathIsLFNFileSpecW(wszName)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines whether a path string represents a network resource.
' ========================================================================================
PRIVATE FUNCTION AfxPathIsNetworkPath (BYREF wszPath AS CONST WSTRING) AS BOOLEAN
   RETURN PathIsNetworkPathW(wszPath)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Searches a path to determine if it contains a valid prefix of the type passed by
' szPrefix. A prefix is one of these types: "C:\", ".", "..", "..\".
' Example: DIM b AS BOOLEAN = AfxPathIsPrefix("C:\", "C:\test\Some\sample")
' ========================================================================================
PRIVATE FUNCTION AfxPathIsPrefix (BYREF wszPrefix AS CONST WSTRING, BYREF wszPath AS CONST WSTRING) AS BOOLEAN
   RETURN PathIsPrefixW(wszPrefix, wszPath)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Searches a path and determines if it is relative.
' ========================================================================================
PRIVATE FUNCTION AfxPathIsRelative (BYREF wszPath AS CONST WSTRING) AS BOOLEAN
   RETURN PathIsRelativeW(wszPath)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Parses a path to determine if it is a directory root.
' ========================================================================================
PRIVATE FUNCTION AfxPathIsRoot (BYREF wszPath AS CONST WSTRING) AS BOOLEAN
   RETURN PathIsRootW(wszPath)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Compares two paths to determine if they have a root component.
' ========================================================================================
PRIVATE FUNCTION AfxPathIsSameRoot (BYREF wszPath1 AS CONST WSTRING, BYREF wszPath2 AS CONST WSTRING) AS BOOLEAN
   RETURN PathIsSameRootW(wszPath1, wszPath2)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines if an existing folder contains the attributes that make it a system
' folder. Alternately indicates if certain attributes qualify a folder to be a
' system folder.
' ========================================================================================
PRIVATE FUNCTION AfxPathIsSystemFolder (BYREF wszPath AS CONST WSTRING, BYVAL dwAttrb AS DWORD) AS BOOLEAN
   RETURN PathIsSystemFolderW(wszPath, dwAttrb)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines if the string is a valid Universal Naming Convention (UNC) for a
' server and share path.
' ========================================================================================
PRIVATE FUNCTION AfxPathIsUNC (BYREF wszPath AS CONST WSTRING) AS BOOLEAN
   RETURN PathIsUNCW(wszPath)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines if a string is a valid Universal Naming Convention (UNC) for a server
' path only.
' ========================================================================================
PRIVATE FUNCTION AfxPathIsUNCServer (BYREF wszPath AS CONST WSTRING) AS BOOLEAN
   RETURN PathIsUNCServerW(wszPath)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines if a string is a valid Universal Naming Convention (UNC) share path,
' \\ server\ share.
' ========================================================================================
PRIVATE FUNCTION AfxPathIsUNCServerShare (BYREF wszPath AS CONST WSTRING) AS BOOLEAN
   RETURN PathIsUNCServerShareW(wszPath)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Tests a given string to determine if it conforms to a valid URL format.
' ========================================================================================
PRIVATE FUNCTION AfxPathIsURL (BYREF wszPath AS CONST WSTRING) AS BOOLEAN
   RETURN PathIsURLW(wszPath)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gives an existing folder the proper attributes to become a system folder.
' ========================================================================================
PRIVATE FUNCTION AfxPathMakeSystemFolder (BYREF wszPath AS CONST WSTRING) AS BOOLEAN
   RETURN PathMakeSystemFolderW(wszPath)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Removes the attributes from a folder that make it a system folder. This folder
' must actually exist in the file system.
' ========================================================================================
PRIVATE FUNCTION AfxPathUnmakeSystemFolder (BYREF wszPath AS CONST WSTRING) AS BOOLEAN
   RETURN PathUnmakeSystemFolderW(wszPath)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Searches for a file.
' PathFindOnPath searches for the file specified by wszFile. If no directories are specified
' in ppszOtherDirs, it attempts to find the file by searching standard directories such as
' System32 and the directories specified in the PATH environment variable. To expedite the
' process or enable PathFindOnPath to search a wider range of directories, use the ppszOtherDirs
' parameter to specify one or more directories to be searched first. If more than one file
' has the name specified by pszFile, PathFindOnPath returns the first instance it finds.
' ========================================================================================
PRIVATE FUNCTION AfxPathFindOnPath (BYREF wszFile AS CONST WSTRING, BYVAL ppwszOtherDirs AS WSTRING PTR PTR = NULL) AS CWSTR
   DIM _wszFile AS WSTRING * MAX_PATH = wszFile
   PathFindOnPathW(_wszFile, ppwszOtherDirs)
   RETURN _wszFile
END FUNCTION
' ========================================================================================

' ========================================================================================
' Searches a string using an Microsoft MS-DOS wild card match type.
' ========================================================================================
PRIVATE FUNCTION AfxPathMatchSpec (BYREF wszFile AS CONST WSTRING, BYREF wszSpec AS CONST WSTRING) AS BOOLEAN
   RETURN PathMatchSpecW(wszFile, wszSpec)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Removes any arguments from a given path.
' This function should not be used on generic command path templates (from users or the
' registry), but rather it should be used only on templates that the application knows to
' be well formed.
' Usage: DIM cws AS CWSTR = AfxPathRemoveArgs("c:\a\b\FileA Arg1 Arg2")
' ========================================================================================
PRIVATE FUNCTION AfxPathRemoveArgs (BYREF wszPath AS CONST WSTRING) AS CWSTR
   DIM _wszPath AS WSTRING * MAX_PATH = wszPath
   PathRemoveArgsW(_wszPath)
   RETURN _wszPath
END FUNCTION
' ========================================================================================

' ========================================================================================
' Removes all leading and trailing spaces from a path.
' ========================================================================================
PRIVATE FUNCTION AfxPathRemoveBlanks (BYREF wszPath AS CONST WSTRING) AS CWSTR
   DIM _wszPath AS WSTRING * MAX_PATH = wszPath
   PathRemoveBlanksW(_wszPath)
   RETURN _wszPath
END FUNCTION
' ========================================================================================

' ========================================================================================
' Removes the trailing file name and backslash from a path, if it has them.
' ========================================================================================
PRIVATE FUNCTION AfxPathRemoveFileSpec (BYREF wszPath AS CONST WSTRING) AS CWSTR
   DIM _wszPath AS WSTRING * MAX_PATH = wszPath
   PathRemoveFileSpecW(_wszPath)
   RETURN _wszPath
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines if a given path is correctly formatted and fully qualified.
' ========================================================================================
PRIVATE FUNCTION AfxPathSearchAndQualify (BYREF wszPath AS CONST WSTRING, BYREF wszFullyQualifiedPath AS CONST WSTRING, BYVAL cchFullyQualifiedPath AS DWORD) AS BOOLEAN
   RETURN PathSearchAndQualifyW(wszPath, wszFullyQualifiedPath, cchFullyQualifiedPath)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the text of a child control in a window or dialog box, using
' PathCompactPath to make sure the path fits in the control.
' ========================================================================================
PRIVATE SUB AfxPathSetDlgItemPath (BYVAL hDlg AS HWND, BYVAL cID AS LONG, BYREF wszPath AS CONST WSTRING)
   PathSetDlgItemPathW(hDlg, cID, wszPath)
END SUB
' ========================================================================================

' ========================================================================================
' Removes the path portion of a fully qualified path and file.
' ========================================================================================
PRIVATE FUNCTION AfxPathStripPath (BYREF wszPath AS CONST WSTRING) AS CWSTR
   DIM _wszPath AS WSTRING * MAX_PATH = wszPath
   PathStripPathW(_wszPath)
   RETURN _wszPath
END FUNCTION
' ========================================================================================

' ========================================================================================
' Removes all parts of the path except for the root information.
' ========================================================================================
PRIVATE FUNCTION AfxPathStripToRoot (BYREF wszPath AS CONST WSTRING) AS CWSTR
   DIM _wszPath AS WSTRING * MAX_PATH = wszPath
   PathStripToRootW(_wszPath)
   RETURN _wszPath
END FUNCTION
' ========================================================================================

' ========================================================================================
' Removes the decoration from a path string.
' ========================================================================================
PRIVATE FUNCTION AfxPathUndecorate (BYREF wszPath AS CONST WSTRING) AS CWSTR
   DIM _wszPath AS WSTRING * MAX_PATH = wszPath
   PathUndecorateW(_wszPath)
   RETURN _wszPath
END FUNCTION
' ========================================================================================

' ========================================================================================
' Takes a fully qualified path, and replaces several folder names with their
' associated environment string.
' ========================================================================================
PRIVATE FUNCTION AfxPathUnExpandEnvStrings (BYREF wszPath AS CONST WSTRING) AS CWSTR
   DIM _wszBuf AS WSTRING * MAX_PATH
   PathUnExpandEnvStringsW(wszPath, _wszBuf, SIZEOF(_wszBuf))
   RETURN _wszBuf
END FUNCTION
' ========================================================================================

' ========================================================================================
' Parses a file location string containing a file location and icon index, and
' returns the valid icon index value.
' ========================================================================================
PRIVATE FUNCTION AfxPathParseIconLocation (BYREF wszIconFile AS CONST WSTRING) AS CWSTR
   DIM _wszIconFile AS WSTRING * MAX_PATH = wszIconFile
   PathParseIconLocationW(_wszIconFile)
   RETURN _wszIconFile
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates a relative path from one file or folder to another.
' This function takes a pair of paths and generates a relative path from one to the other.
' The paths do not have to be fully qualified, but they must have a common prefix, or the
' function will fail and return FALSE.
' For example, let the starting point, wszFrom, be "c:\FolderA\FolderB\FolderC", and the
' ending point, wszTo, be "c:\FolderA\FolderD\FolderE". PathRelativePathTo will return the
' relative path from wszFrom to wszTo as: "..\..\FolderD\FolderE". You will get the same
' result if you set wszFrom to "\FolderA\FolderB\FolderC" and wszTo to "\FolderA\FolderD\FolderE".
' On the other hand, "c:\FolderA\FolderB" and "a:\FolderA\FolderD do not share a common
' prefix, and the function will fail. Note that "\\" is not considered a prefix and is ignored.
' If you set wszFrom to "\\FolderA\FolderB", and wszTo to "\\FolderC\FolderD", the function will fail.
' ========================================================================================
PRIVATE FUNCTION AfxPathRelativePathTo (BYREF wszFrom AS CONST WSTRING, BYVAL dwAttrFrom AS DWORD, BYREF wszTo AS CONST WSTRING, BYVAL dwAttrTo AS DWORD) AS CWSTR
   DIM _wszPath AS WSTRING * MAX_PATH
   IF PathRelativePathToW(_wszPath, wszFrom, dwAttrFrom, wszTo, dwAttrto) THEN RETURN _wszPath
END FUNCTION
' ========================================================================================

' ========================================================================================
' Takes a URL string, determines a scheme for it, and returns a string with an
' appropriate prefix.
' ========================================================================================
PRIVATE FUNCTION AfxUrlApplyScheme (BYREF wszUrl AS WSTRING, BYVAL dwFlags AS DWORD) AS CWSTR
   DIM _wszOut AS WSTRING * INTERNET_MAX_URL_LENGTH
   DIM cchOut AS DWORD = SIZEOF(_wszOut)
   DIM hr AS HRESULT = UrlApplySchemeW(wszUrl, _wszOut, @cchOut, dwFlags)
   IF hr = S_OK OR hr = S_FALSE THEN RETURN _wszOut ELSE RETURN wszUrl
END FUNCTION
' ========================================================================================

' ========================================================================================
' Takes a URL string and converts it into canonical form.
' ========================================================================================
PRIVATE FUNCTION AfxUrlCanonicalize (BYREF wszUrl AS WSTRING, BYVAL dwFlags AS DWORD) AS CWSTR
   DIM _wszCanonicalized AS WSTRING * INTERNET_MAX_URL_LENGTH
   DIM cchCanonicalized AS DWORD = SIZEOF(_wszCanonicalized)
   DIM hr AS HRESULT = UrlCanonicalizeW(wszUrl, _wszCanonicalized, @cchCanonicalized, dwFlags)
   IF hr = S_OK THEN RETURN _wszCanonicalized ELSE RETURN wszUrl
END FUNCTION
' ========================================================================================

' ========================================================================================
' When provided with a relative URL and its base, returns a URL in canonical form.
' ========================================================================================
PRIVATE FUNCTION AfxUrlCombine (BYREF wszBase AS WSTRING, BYREF wszRelative AS WSTRING, BYVAL dwFlags AS DWORD) AS CWSTR
   DIM _wszCombined AS WSTRING * INTERNET_MAX_URL_LENGTH
   DIM cchCombined AS DWORD = SIZEOF(_wszCombined)
   DIM hr AS HRESULT = UrlCombineW(wszBase, wszRelative, _wszCombined, @cchCombined, dwFlags)
   IF hr = S_OK THEN RETURN _wszCombined ELSE RETURN wszBase
END FUNCTION
' ========================================================================================

' ========================================================================================
' Makes a case-sensitive comparison of two URL strings.
' ========================================================================================
PRIVATE FUNCTION AfxUrlCompare (BYREF wszURL1 AS CONST WSTRING, BYREF wszURL2 AS CONST WSTRING, BYVAL fIgnoreSlash AS LONG) AS BOOLEAN
   RETURN UrlCompareW(wszURL1, wszURL2, fIgnoreSlash)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Takes a Microsoft MS-DOS path and converts it to a canonicalized URL.
' ========================================================================================
PRIVATE FUNCTION AfxUrlCreateFromPath (BYREF wszPath AS WSTRING) AS CWSTR
   DIM _wszUrl AS WSTRING * INTERNET_MAX_URL_LENGTH
   DIM cchUrl AS DWORD = SIZEOF(_wszUrl)
   DIM hr AS HRESULT = UrlCreateFromPathW(wszPath, _wszUrl, @cchUrl, 0)
   IF hr = S_OK OR hr = S_FALSE THEN RETURN _wszUrl ELSE RETURN wszPath
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts characters in a URL that might be altered during transport across the
' Internet ("unsafe" characters) into their corresponding escape sequences.
' ========================================================================================
PRIVATE FUNCTION AfxUrlEscape (BYREF wszURL AS WSTRING, BYVAL dwFlags AS DWORD) AS CWSTR
   DIM _wszEscaped AS WSTRING * INTERNET_MAX_URL_LENGTH
   DIM cchEscaped AS DWORD = SIZEOF(_wszEscaped)
   DIM hr AS HRESULT = UrlEscapeW(wszURL, _wszEscaped, @cchEscaped, 0)
   IF hr = S_OK THEN RETURN _wszEscaped ELSE RETURN wszURL
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts space characters into their corresponding escape sequence.
' ========================================================================================
PRIVATE FUNCTION AfxUrlEscapeSpaces (BYREF wszURL AS WSTRING) AS CWSTR
   DIM _wszEscaped AS WSTRING * INTERNET_MAX_URL_LENGTH
   DIM cchEscaped AS DWORD = SIZEOF(_wszEscaped)
   DIM hr AS HRESULT = UrlEscapeSpaces(wszURL, _wszEscaped, @cchEscaped)
   IF hr = S_OK THEN RETURN _wszEscaped ELSE RETURN wszURL
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the location from a URL.
' ========================================================================================
PRIVATE FUNCTION AfxUrlGetLocation (BYREF wszUrl AS CONST WSTRING) AS CWSTR
   DIM _wszUrl AS WSTRING * MAX_PATH = wszUrl
   DIM pwszLocation AS WSTRING PTR = cast(WSTRING PTR, UrlGetLocationW(_wszUrl))
   IF pwszLocation THEN *cast(WSTRING PTR, @_wszUrl) = *pwszLocation
   RETURN _wszUrl
END FUNCTION
' ========================================================================================

' ========================================================================================
' Accepts a URL string and returns a specified part of that URL.
' ========================================================================================
PRIVATE FUNCTION AfxUrlGetPart (BYREF wszUrl AS CONST WSTRING, BYVAL dwPart AS DWORD, BYVAL dwFlags AS DWORD) AS CWSTR
   DIM _wszOut AS WSTRING * INTERNET_MAX_URL_LENGTH
   DIM cchOut AS DWORD = SIZEOF(_wszOut)
   DIM hr AS HRESULT = UrlGetPartW(wszUrl, _wszOut, @cchOut, dwPart, dwFlags)
   IF hr = S_OK THEN RETURN _wszOut
END FUNCTION
' ========================================================================================

' ========================================================================================
' Hashes a URL string.
' ========================================================================================
PRIVATE FUNCTION AfxUrlHash (BYREF wszURL AS CONST WSTRING, BYVAL pbHash AS BYTE PTR, BYVAL cbHash AS DWORD) AS HRESULT
   RETURN UrlHashW(wszURL, pbHash, cbHash)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Tests whether or not a URL is a specified type.
' ========================================================================================
PRIVATE FUNCTION AfxUrlIs (BYREF wszURL AS CONST WSTRING, BYVAL nUrlIs AS URLIS) AS BOOLEAN
   RETURN UrlIsW(wszURL, nUrlIs)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Tests whether or not a URL is a specified type.
' ========================================================================================
PRIVATE FUNCTION AfxUrlIsFileUrl (BYREF wszURL AS CONST WSTRING) AS BOOLEAN
   RETURN UrlIsW(wszURL, URLIS_FILEURL)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns whether or not a URL is a No History URL.
' ========================================================================================
PRIVATE FUNCTION AfxUrlIsNoHistory (BYREF wszURL AS CONST WSTRING) AS BOOLEAN
   RETURN UrlIsNoHistoryW(wszURL)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns whether a URL is opaque.
' ========================================================================================
PRIVATE FUNCTION AfxUrlIsOpaque (BYREF wszURL AS CONST WSTRING) AS BOOLEAN
   RETURN UrlIsOpaqueW(wszURL)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts escape sequences back into ordinary characters.
' ========================================================================================
PRIVATE FUNCTION AfxUrlUnescape (BYREF wszURL AS WSTRING, BYVAL dwFlags AS DWORD) AS CWSTR
   DIM _wszUnescaped AS WSTRING * INTERNET_MAX_URL_LENGTH
   DIM cchUnescaped AS DWORD = SIZEOF(_wszUnescaped)
   DIM hr AS HRESULT = UrlUnescapeW(wszURL, _wszUnescaped, @cchUnescaped, 0)
   IF hr = S_OK THEN
      IF (dwFlags AND URL_UNESCAPE_INPLACE) <> URL_UNESCAPE_INPLACE THEN RETURN _wszUnescaped
   ELSE
      RETURN wszURL
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts escape sequences back into ordinary characters and overwrites the
' original string.
' ========================================================================================
PRIVATE FUNCTION AfxUrlUnescapeInPlace (BYREF wszURL AS WSTRING, BYVAL dwFlags AS DWORD) AS HRESULT
   RETURN UrlUnescapeW(wszURL, NULL, NULL, dwFlags OR URL_UNESCAPE_INPLACE)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Searches a path to determine whether it contains a file of a specified file type extension.
' Returns TRUE if a file name pattern specified in wszSpec matched the file name found in
' the string pointed to by pszFile; otherwise, it returns FALSE.
' Windows Vista. Not declared in the FB includes.
' Example: DIM b AS BOOLEAN = AfxPathMatchSpec("c:\test\file.txt", "*.txt")
' ========================================================================================
PRIVATE FUNCTION AfxPathMatchSpecEx (BYREF wszFile AS CONST WSTRING, BYREF wszSpec AS CONST WSTRING, BYVAL dwFlags AS DWORD = 0) AS BOOLEAN
   DIM AS ANY PTR pLib = DyLibLoad("Shlwapi.dll")
   IF pLib = 0 THEN EXIT FUNCTION
   DIM pProc AS FUNCTION (BYREF wszFile AS CONST WSTRING, BYREF wszSpec AS CONST WSTRING, BYVAL dwFlags AS DWORD) AS HRESULT
   pProc = DyLibSymbol(pLib, "PathMatchSpecExW")
   IF pProc = 0 THEN EXIT FUNCTION
   RETURN pProc(wszFile, wszSpec, dwFlags) = S_OK
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' *** DEPRECATED ***
' [UrlFixupW is available for use in the operating systems specified in the Requirements
' section. It may be altered or unavailable in subsequent versions.]
' Attempts to correct a URL whose protocol identifier is incorrect. For example, htttp
' will be changed to http.
' Not declared in the FB includes. Available only in unicode version.
' Requirements: Windows XP, Windows Vista, Windows 7.
' Example: DIM cws AS CWSTR = AfxUrlFixup("htttp://www.planetsquires.com/protect/forum/index.php")
' ========================================================================================
PRIVATE FUNCTION AfxUrlFixup (BYREF wszUrl AS CONST WSTRING) AS CWSTR
   DIM AS ANY PTR pLib = DyLibLoad("Shlwapi.dll")
   IF pLib = 0 THEN RETURN ""
   DIM pProc AS FUNCTION (BYREF wszUrl AS WSTRING, BYREF pszTranslatedUrl AS WSTRING, BYVAL cchMax AS DWORD) AS HRESULT
   pProc = DyLibSymbol(pLib, "UrlFixupW")
   IF pProc = 0 THEN RETURN ""
   DIM _wszUrl AS WSTRING * INTERNET_MAX_URL_LENGTH = wszUrl
   DIM _wszTranslatedUrl AS WSTRING * INTERNET_MAX_URL_LENGTH
   DIM hr AS HRESULT = pProc(_wszUrl, _wszTranslatedUrl, INTERNET_MAX_URL_LENGTH)
   DyLibFree(pLib)
   RETURN _wszTranslatedUrl
END FUNCTION
' ========================================================================================
