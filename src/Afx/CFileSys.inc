' ########################################################################################
' Microsoft Windows
' File: CFileSys.inc
' Contents: File System class
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "windows.bi"
#include once "Afx/AfxScrRun.bi"
#include once "Afx/CWStr.inc"

NAMESPACE Afx

' ########################################################################################
' CFileSys - File system class
' ########################################################################################
TYPE CFileSys

   m_Result AS HRESULT
   m_bUninitCOM AS BOOLEAN
   m_pFileSys AS Afx_IFileSystem3 PTR

   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR
   DECLARE FUNCTION GetLastResult () AS HRESULT
   DECLARE FUNCTION SetResult (BYVAL Result AS HRESULT) AS HRESULT
   DECLARE FUNCTION DriveLetters () AS CWSTR
   DECLARE FUNCTION DriveExists (BYREF cbsDrive AS CWSTR) AS BOOLEAN
   DECLARE FUNCTION FileExists (BYREF cbsFileSpec AS CWSTR) AS BOOLEAN
   DECLARE FUNCTION FolderExists (BYREF cbsFileSpec AS CWSTR) AS BOOLEAN
   DECLARE FUNCTION IsDriveReady (BYREF cbsDrive AS CWSTR) AS BOOLEAN
   DECLARE FUNCTION GetDriveName (BYREF cbsPathSpec AS CWSTR) AS CWSTR
   DECLARE FUNCTION GetFolderName (BYREF cbsPathSpec AS CWSTR) AS CWSTR
   DECLARE FUNCTION GetFileName (BYREF cbsPathSpec AS CWSTR) AS CWSTR
   DECLARE FUNCTION GetExtensionName (BYREF cbsPathSpec AS CWSTR) AS CWSTR
   DECLARE FUNCTION GetBaseName (BYREF cbsPathSpec AS CWSTR) AS CWSTR
   DECLARE FUNCTION GetAbsolutePathName (BYREF cbsPathSpec AS CWSTR) AS CWSTR
   DECLARE FUNCTION GetTempName () AS CWSTR
   DECLARE FUNCTION GetDriveShareName (BYREF cbsDrive AS CWSTR) AS CWSTR
   DECLARE FUNCTION GetDriveFileSystem (BYREF cbsDrive AS CWSTR) AS CWSTR
   DECLARE FUNCTION GetDriveType (BYREF cbsDrive AS CWSTR) AS DRIVETYPECONST
   DECLARE FUNCTION GetFolderType (BYREF cbsFolder AS CWSTR) AS CWSTR
   DECLARE FUNCTION GetFileType (BYREF cbsFile AS CWSTR) AS CWSTR
   DECLARE FUNCTION GetFileVersion (BYREF cbsFileName AS CWSTR) AS CWSTR
   DECLARE FUNCTION GetParentFolderName (BYREF cbsFolder AS CWSTR) AS CWSTR
   DECLARE FUNCTION GetFolderShortPath (BYREF cbsFolder AS CWSTR) AS CWSTR
   DECLARE FUNCTION GetFolderShortName (BYREF cbsFolder AS CWSTR) AS CWSTR
   DECLARE FUNCTION GetFolderDriveLetter (BYREF cbsFolder AS CWSTR) AS CWSTR
   DECLARE FUNCTION SetFolderName (BYREF cbsFolder AS CWSTR, BYREF cbsName AS CWSTR) AS HRESULT
   DECLARE FUNCTION BuildPath (BYREF cbsPath AS CWSTR, BYREF cbsName AS CWSTR) AS CWSTR
   DECLARE FUNCTION GetSerialNumber (BYREF cbsDrive AS CWSTR) AS LONG
   DECLARE FUNCTION GetVolumeName (BYREF cbsDrive AS CWSTR) AS CWSTR
   DECLARE FUNCTION SetVolumeName (BYREF cbsDrive AS CWSTR, BYREF cbsName AS CWSTR) AS HRESULT
   DECLARE FUNCTION GetNumDrives () AS LONG
   DECLARE FUNCTION GetNumSubFolders (BYREF cbsFolder AS CWSTR) AS LONG
   DECLARE FUNCTION GetNumFiles (BYREF cbsFolder AS CWSTR) AS LONG
   DECLARE FUNCTION IsRootFolder (BYREF cbsFolder AS CWSTR) AS BOOLEAN
   DECLARE FUNCTION GetStandardStream (BYVAL StandardStreamType AS STANDARDSTREAMTYPES, BYVAL bUnicode AS VARIANT_BOOL = FALSE) AS Afx_ITextStream PTR
   DECLARE FUNCTION GetFileSize (BYREF cbsFile AS CWSTR) AS LONG
   DECLARE FUNCTION SetFileName (BYREF cbsFile AS CWSTR, BYREF cbsName AS CWSTR) AS HRESULT
   DECLARE FUNCTION GetFileShortPath (BYREF cbsFile AS CWSTR) AS CWSTR
   DECLARE FUNCTION GetFileShortName (BYREF cbsFile AS CWSTR) AS CWSTR
   DECLARE FUNCTION GetFileAttributes (BYREF cbsFile AS CWSTR) AS FILEATTRIBUTE
   DECLARE FUNCTION SetFileAttributes (BYREF cbsFile AS CWSTR, BYVAL lAttr AS FILEATTRIBUTE) AS HRESULT
   DECLARE FUNCTION GetFileDateCreated (BYREF cbsFile AS CWSTR) AS DATE_
   DECLARE FUNCTION GetFileDateLastModified (BYREF cbsFile AS CWSTR) AS DATE_
   DECLARE FUNCTION GetFileDateLastAccessed (BYREF cbsFile AS CWSTR) AS DATE_
   DECLARE FUNCTION GetFolderSize (BYREF cbsFolder AS CWSTR) AS LONG
   DECLARE FUNCTION GetFolderAttributes (BYREF cbsFolder AS CWSTR) AS FILEATTRIBUTE
   DECLARE FUNCTION SetFolderAttributes (BYREF cbsFolder AS CWSTR, BYVAL lAttr AS FILEATTRIBUTE) AS HRESULT
   DECLARE FUNCTION GetFolderDateCreated (BYREF cbsFolder AS CWSTR) AS DATE_
   DECLARE FUNCTION GetFolderDateLastModified (BYREF cbsFolder AS CWSTR) AS DATE_
   DECLARE FUNCTION GetFolderDateLastAccessed (BYREF cbsFolder AS CWSTR) AS DATE_
   DECLARE FUNCTION GetDriveAvailableSpace (BYREF cbsDrive AS CWSTR) AS DOUBLE
   DECLARE FUNCTION GetDriveFreeSpace (BYREF cbsDrive AS CWSTR) AS DOUBLE
   DECLARE FUNCTION GetDriveTotalSize (BYREF cbsDrive AS CWSTR) AS DOUBLE
   DECLARE FUNCTION CreateFolder (BYREF cbsFolder AS CWSTR) AS Afx_IFolder PTR
   DECLARE FUNCTION DeleteFolder (BYREF cbsFolder AS CWSTR, BYVAL bForce AS VARIANT_BOOL = FALSE) AS HRESULT
   DECLARE FUNCTION MoveFolder (BYREF cbsSource AS CWSTR, BYREF cbsDestination AS CWSTR) AS HRESULT
   DECLARE FUNCTION CopyFolder (BYREF cbsSource AS CWSTR, BYREF cbsDestination AS CWSTR, BYVAL OverWriteFiles AS VARIANT_BOOL = -1) AS HRESULT
   DECLARE FUNCTION CopyFile (BYREF cbsSource AS CWSTR, BYREF cbsDestination AS CWSTR, BYVAL OverWriteFiles AS VARIANT_BOOL = -1) AS HRESULT
   DECLARE FUNCTION DeleteFile (BYREF cbsFileSpec AS CWSTR, BYVAL bForce AS VARIANT_BOOL = FALSE) AS HRESULT
   DECLARE FUNCTION MoveFile (BYREF cbsSource AS CWSTR, BYREF cbsDestination AS CWSTR) AS HRESULT

END TYPE
' ########################################################################################

' ========================================================================================
' Constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CFileSys
   ' // Initialize the COM library
   DIM hr AS HRESULT = CoInitialize(NULL)
   IF hr = S_OK OR hr = S_FALSE THEN m_bUninitCOM = TRUE
   ' // Create an instance of the Filesystem object
   DIM CLSID_FileSystemObject_ AS GUID = (&h0D43FE01, &hF093, &h11CF, {&h89, &h40, &h00, &hA0, &hC9, &h05, &h42, &h28})
' DIM IID_IFileSystem AS GUID = (&h0AB5A3D0, &hE5B6, &h11D0, {&hAB, &hF5, &h00, &hA0, &hC9, &h0F, &hFF, &hC0})
   DIM IID_IFileSystem3 AS GUID = (&h2A0B9D10, &h4B87, &h11D3, {&hA9, &h7A, &h00, &h10, &h4B, &h36, &h5C, &h9F})
   SetResult(CoCreateInstance(@CLSID_FileSystemObject_, NULL, CLSCTX_INPROC_SERVER, @IID_IFileSystem3, @m_pFileSys))
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Destructor
' ========================================================================================
PRIVATE DESTRUCTOR CFileSys
   IF m_pFileSys THEN m_pFileSys->Release
   IF m_bUninitCOM THEN CoUninitialize
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the last result code.
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetLastResult () AS HRESULT
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the last result code.
' ========================================================================================
PRIVATE FUNCTION CFileSys.SetResult (BYVAL Result AS HRESULT) AS HRESULT
   m_Result = Result
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Checks if the specified drive exists and returns TRUE or FALSE.
' For drive letters, the root drive is not included. For example, the path for the C drive
' is C:, not C:\.
' Example:
' DIM pFileSys AS CFileSys
' DIM fExists AS BOOLEAN = pFileSys.DriveExists("C:")
' ========================================================================================
PRIVATE FUNCTION CFileSys.DriveExists (BYREF cwsDrive AS CWSTR) AS BOOLEAN
   IF m_pFileSys = NULL THEN RETURN FALSE
   DIM pfExists AS VARIANT_BOOL
   DIM bstrDrive AS AFX_BSTR = SysAllocString(**cwsDrive)
   SetResult(m_pFileSys->DriveExists(bstrDrive, @pfExists))
   SysFreeString bstrDrive
   RETURN pfExists
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns True if the specified folder exists; False if it does not.
' Example:
' DIM pFileSys AS CFileSys
' DIM fExists AS BOOLEAN = pFileSys.FolderExists("C:\MyFolder")
' ========================================================================================
PRIVATE FUNCTION CFileSys.FolderExists (BYREF cwsFolderSpec AS CWSTR) AS BOOLEAN
   IF m_pFileSys = NULL THEN RETURN FALSE
   DIM pfExists AS VARIANT_BOOL
   DIM bstrFolderSpec AS AFX_BSTR = SysAllocString(**cwsFolderSpec)
   SetResult(m_pFileSys->FolderExists(bstrFolderSpec, @pfExists))
   SysFreeString bstrFolderSpec
   RETURN pfExists
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns True (-1) if the specified drive is ready; False (0) if it is not.
' For drive letters, the root drive is not included. For example, the path for the C drive
' is C:, not C:\.
' Example:
' DIM pFileSys AS CFileSys
' DIM fIsReady AS BOOLEAN = pFileSys.IsDriveReady("C:")
' ========================================================================================
PRIVATE FUNCTION CFileSys.IsDriveReady (BYREF cwsDrive AS CWSTR) AS BOOLEAN
   IF m_pFileSys = NULL THEN RETURN FALSE
   ' // Get a reference to the IDrive interface
   DIM pDrive AS Afx_IDrive PTR
   DIM bstrDrive AS AFX_BSTR = SysAllocString(**cwsDrive)
   SetResult(m_pFileSys->GetDrive(bstrDrive, @pDrive))
   SysFreeString bstrDrive
   IF pDrive = NULL THEN RETURN FALSE
   DIM fIsReady AS VARIANT_BOOL
   SetResult(pDrive->get_IsReady(@fIsReady))
   pDrive->Release
   RETURN fIsReady
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns True if the specified file exists; False if it does not.
' Example:
' DIM pFileSys AS CFileSys
' DIM fExists AS BOOLEAN = pFileSys.FileExists("C:\MyFolder\Test.txt")
' ========================================================================================
PRIVATE FUNCTION CFileSys.FileExists (BYREF cwsFileSpec AS CWSTR) AS BOOLEAN
   IF m_pFileSys = NULL THEN RETURN FALSE
   DIM pfExists AS VARIANT_BOOL
   DIM bstrFileSpec AS AFX_BSTR = SysAllocString(**cwsFileSpec)
   SetResult(m_pFileSys->FileExists(bstrFileSpec, @pfExists))
   SysFreeString bstrFileSpec
   RETURN pfExists
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns a string containing the name of the drive for a specified path.
' Example:
' DIM pFileSys AS CFileSys
' DIM cbsName AS CWSTR = pFileSys.GetDriveName("C:\MyFolder\Test.txt")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetDriveName (BYREF cwsPathSpec AS CWSTR) AS CWSTR
   IF m_pFileSys = NULL THEN RETURN ""
   DIM bstrName AS AFX_BSTR
   DIM bstrPathSpec AS AFX_BSTR = SysAllocString(**cwsPathSpec)
   SetResult(m_pFileSys->GetDriveName(bstrPathSpec, @bstrName))
   SysFreeString bstrPathSpec
   DIM cwsName AS CWSTR = *bstrName
   SysFreeString bstrName
   RETURN cwsName
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns a string containing the name of the folder for a specified path, i.e. the
' path minus the file name.
' Example:
' DIM pFileSys AS CFileSys
' DIM cwsName AS CWSTR = pFileSys.GetFolderName("C:\MyFolder\Test.txt")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetFolderName (BYREF cwsPathSpec AS CWSTR) AS CWSTR
   IF m_pFileSys = NULL THEN RETURN ""
   ' // Retrieve the file name
   DIM bstrFileName AS AFX_BSTR
   DIM bstrPathSpec AS AFX_BSTR = SysAllocString(**cwsPathSpec)
   IF INSTRREV(*bstrPathSpec, ".") THEN SetResult(m_pFileSys->GetFileName(bstrPathSpec, @bstrFileName))
   SysFreestring bstrPathSpec
   DIM cwsFileName AS CWSTR = *bstrFileName
   SysFreeString bstrFileName
   ' // Remove the file name from the path
   DIM cwsFolderName AS CWSTR = cwsPathSpec
   IF LEN(cwsFileName) THEN cwsFolderName = LEFT(**cwsFolderName, LEN(cwsFolderName) - LEN(cwsFileName) - 1)
   IF RIGHT(**cwsFolderName, 1) = "\" THEN cwsFolderName = LEFT(**cwsFolderName, LEN(cwsFolderName) - 1)
   RETURN cwsFolderName
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the short path used by programs that require the earlier 8.3 file naming convention.
' Example:
' DIM pFileSys AS CFileSys
' DIM cwsFolderShortPath AS CWSTR = pFileSys.GetFolderShortPath("c:\MyFolder)
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetFolderShortPath (BYREF cwsFolder AS CWSTR) AS CWSTR
   IF m_pFileSys = NULL THEN RETURN ""
   ' // Get a reference to the IFolder interface
   DIM bstrFolder AS AFX_BSTR = SysAllocString(**cwsFolder)
   DIM pFolder AS Afx_IFolder PTR
   SetResult(m_pFileSys->GetFolder(bstrFolder, @pFolder))
   SysFreeString bstrFolder
   IF pFolder = NULL THEN RETURN ""
   ' // Get the short path
   DIM bstrPath AS AFX_BSTR
   SetResult(pFolder->get_ShortPath(@bstrPath))
   pFolder->Release
   DIM cwsPath AS CWSTR = *bstrPath
   SysFreeString bstrPath
   RETURN cwsPath
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the short name used by programs that require the earlier 8.3 file naming convention.
' Example:
' DIM pFileSys AS CFileSys
' DIM cwsFolderShortname AS CWSTR = pFileSys.GetFolderShortName("c:\MyFolder)
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetFolderShortName (BYREF cwsFolder AS CWSTR) AS CWSTR
   IF m_pFileSys = NULL THEN RETURN ""
   ' // Get a reference to the IFolder interface
   DIM bstrFolder AS AFX_BSTR = SysAllocString(**cwsFolder)
   DIM pFolder AS Afx_IFolder PTR
   SetResult(m_pFileSys->GetFolder(bstrFolder, @pFolder))
   SysFreeString bstrFolder
   IF pFolder = NULL THEN RETURN ""
   ' // Get the short name
   DIM bstrName AS AFX_BSTR
   SetResult(pFolder->get_ShortName(@bstrName))
   pFolder->Release
   DIM cwsName AS CWSTR = *bstrName
   SysFreeString bstrName
   RETURN cwsName
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns a string containing the name of the file for a specified path.
' Example:
' DIM pFileSys AS CFileSys
' DIM cwsName AS CWSTR = pFileSys.GetFileName("C:\MyFolder\Test.txt")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetFileName (BYREF cwsPathSpec AS CWSTR) AS CWSTR
   IF m_pFileSys = NULL THEN RETURN ""
   DIM bstrName AS AFX_BSTR
   DIM bstrPathSpec AS AFX_BSTR = SysAllocString(**cwsPathSpec)
   IF INSTRREV(*bstrPathSpec, ".") THEN SetResult(m_pFileSys->GetFileName(bstrPathSpec, @bstrName))
   SysFreeString bstrPathSpec
   DIM cwsName AS CWSTR = *bstrName
   SysFreeString bstrName
   RETURN cwsName
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns a string containing the extension name of the file for a specified path.
' Example:
' DIM pFileSys AS CFileSys
' DIM cwsName AS CWSTR = pFileSys.GetExtensionName("C:\MyFolder\Test.txt")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetExtensionName (BYREF cwsPathSpec AS CWSTR) AS CWSTR
   IF m_pFileSys = NULL THEN RETURN ""
   DIM bstrName AS AFX_BSTR
   DIM bstrPathSpec AS AFX_BSTR = SysAllocString(**cwsPathSpec)
   SetResult(m_pFileSys->GetExtensionName(bstrPathSpec, @bstrName))
   SysFreeString bstrPathSpec
   DIM cwsName AS CWSTR = *bstrName
   SysFreeString bstrName
   RETURN cwsName
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns a string containing the base name of the last component, less any file
' extension, in a path.
' Example:
' DIM pFileSys AS CFileSys
' DIM cwsName AS CWSTR = pFileSys.GetBaseName("C:\MyFolder\Test.txt")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetBaseName (BYREF cwsPathSpec AS CWSTR) AS CWSTR
   IF m_pFileSys = NULL THEN RETURN ""
   DIM bstrName AS AFX_BSTR
   DIM bstrPathSpec AS AFX_BSTR = SysAllocString(**cwsPathSpec)
   IF INSTRREV(*bstrPathSpec, ".") THEN SetResult(m_pFileSys->GetBaseName(bstrPathSpec, @bstrName))
   SysFreeString bstrPathSpec
   DIM cwsName AS CWSTR = *bstrName
   SysFreestring bstrName
   RETURN cwsName
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns complete and unambiguous path from a provided path specification.
' Example:
' DIM pFileSys AS CFileSys
' DIM cwsName AS CWSTR = pFileSys.GetAbsolutePathName("C:\MyFolder\Test.txt")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetAbsolutePathName (BYREF cwsPathSpec AS CWSTR) AS CWSTR
   IF m_pFileSys = NULL THEN RETURN ""
   DIM bstrName AS AFX_BSTR
   DIM bstrPathSpec AS AFX_BSTR = SysAllocString(**cwsPathSpec)
   SetResult(m_pFileSys->GetAbsolutePathName(bstrPathSpec, @bstrName))
   SysFreeString bstrPathSpec
   DIM cwsName AS CWSTR = *bstrName
   SysFreeString bstrName
   RETURN cwsName
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns a randomly generated temporary file or folder name that is useful for performing
' operations that require a temporary file or folder.
' Example:
' DIM pFileSys AS CFileSys
' DIM cwsName AS CWSTR = pFileSys.GetTempName
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetTempName () AS CWSTR
   IF m_pFileSys = NULL THEN RETURN ""
   DIM bstrName AS AFX_BSTR
   SetResult(m_pFileSys->GetTempName(@bstrName))
   DIM cwsName AS CWSTR = *bstrName
   SysFreeString bstrName
   RETURN cwsName
END FUNCTION
' ========================================================================================

' ========================================================================================
' *** UNTESTED ***
' Returns the network share name for a specified drive.
' For drive letters, the root drive is not included. For example, the path for the C drive
' is C:, not C:\.
' DIM pFileSys AS CFileSys
' DIM cwsShareName AS CWSTR = pFileSys.GetDriveShareName("H:")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetDriveShareName (BYREF cwsDrive AS CWSTR) AS CWSTR
   IF m_pFileSys = NULL THEN RETURN ""
   ' // Get a reference to the IDrive interface
   DIM pDrive AS Afx_IDrive PTR
   DIM bstrDrive AS AFX_BSTR = SysAllocString(**cwsDrive)
   SetResult(m_pFileSys->GetDrive(bstrDrive, @pDrive))
   SysFreeString bstrDrive
   IF pDrive = NULL THEN RETURN ""
   ' // Get the name
   DIM bstrName AS AFX_BSTR
   SetResult(pDrive->get_ShareName(@bstrName))
   pDrive->Release
   DIM cwsName AS CWSTR = *bstrName
   SysFreeString bstrName
   RETURN cwsName
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the type of file system in use for the specified drive.
' Available return types include FAT, NTFS, and CDFS.
' For drive letters, the root drive is not included. For example, the path for the C drive
' is C:, not C:\.
' Example:
' DIM pFileSys AS CFileSys
' DIM cwsFileSystem AS CWSTR = pFileSys.GetDriveFileSystem("C:")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetDriveFileSystem (BYREF cwsDrive AS CWSTR) AS CWSTR
   IF m_pFileSys = NULL THEN RETURN ""
   ' // Get a reference to the IDrive interface
   DIM pDrive AS Afx_IDrive PTR
   DIM bstrDrive AS AFX_BSTR = SysAllocString(**cwsDrive)
   SetResult(m_pFileSys->GetDrive(bstrDrive, @pDrive))
   SysFreeString bstrDrive
   IF pDrive = NULL THEN RETURN ""
   ' // Get the file system name
   DIM bstrFileSystem AS AFX_BSTR
   SetResult(pDrive->get_FileSystem(@bstrFileSystem))
   pDrive->Release
   DIM cwsFileSystem AS CWSTR = *bstrFileSystem
   SysFreeString bstrFileSystem
   RETURN cwsFileSystem
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns a value indicating the type of a specified drive.
' Return value:
' DriveType_UnknownType = 0
' DriveType_Removable = 1
' DriveType_Fixed = 2
' DriveType_Remote = 3
' DriveType_CDRom = 4
' DriveType_RamDisk = 5
' Example:
' DIM pFileSys AS CFileSys
' DIM nDriveType AS DRIVETYPECONST = pFileSys.GetDriveType("C:")
' DIM t AS CWSTR
' SELECT CASE pDrive.DriveType
'  CASE 0 : t = "Unknown"
'  CASE 1 : t = "Removable"
'  CASE 2 : t = "Fixed"
'  CASE 3 : t = "Network"
'  CASE 4 : t = "CD-ROM"
'  CASE 5 : t = "RAM Disk"
' END SELECT
' For drive letters, the root drive is not included. For example, the path for the C drive
' is C:, not C:\.
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetDriveType (BYREF cwsDrive AS CWSTR) AS DRIVETYPECONST
   IF m_pFileSys = NULL THEN RETURN 0
   ' // Get a reference to the IDrive interface
   DIM pDrive AS Afx_IDrive PTR
   DIM bstrDrive AS AFX_BSTR = SysAllocString(**cwsDrive)
   SetResult(m_pFileSys->GetDrive(bstrDrive, @pDrive))
   SysFreeString bstrDrive
   IF pDrive = NULL THEN RETURN 0
   ' // Get the drive type
   DIM nDriveType AS DRIVETYPECONST
   SetResult(pDrive->get_DriveType(@nDriveType))
   pDrive->Release
   RETURN nDriveType
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns information about the type of a folder.
' Example:
' DIM pFileSys AS CFileSys
' DIM cwsFolderType AS CWSTR = pFileSys.GetFolderType("c:\MyFolder)
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetFolderType (BYREF cwsFolder AS CWSTR) AS CWSTR
   IF m_pFileSys = NULL THEN RETURN ""
   ' // Get a reference to the IFolder interface
   DIM pFolder AS Afx_IFolder PTR
   DIM bstrFolder AS AFX_BSTR = SysAllocString(**cwsFolder)
   SetResult(m_pFileSys->GetFolder(bstrFolder, @pFolder))
   SysFreeString bstrFolder
   IF pFolder = NULL THEN RETURN ""
   ' // Get the folder type
   DIM bstrType AS AFX_BSTR
   SetResult(pFolder->get_Type(@bstrType))
   pFolder->Release
   DIM cwsType AS CWSTR = *bstrType
   SysFreeString bstrType
   RETURN cwsType
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns information about the type of a file. For example, for files ending in .TXT,
' "Text Document" is returned.
' Example:
' DIM pFileSys AS CFileSys
' DIM cwsFileType AS CWSTR = pFileSys.FileType("C:\MyPath\MyFile.txt")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetFileType (BYREF cwsFile AS CWSTR) AS CWSTR
   IF m_pFileSys = NULL THEN RETURN ""
   ' // Get a reference to the IFile interface
   DIM pFile AS Afx_IFile PTR
   DIM bstrFile AS AFX_BSTR = SysAllocString(**cwsFile)
   SetResult(m_pFileSys->GetFile(bstrFile, @pFile))
   SysFreeString bstrFile
   IF pFile = NULL THEN RETURN ""
   ' // Get the type of the file
   DIM bstrType AS AFX_BSTR
   SetResult(pFile->get_Type(@bstrType))
   pFile->Release
   DIM cwsType AS CWSTR = *bstrType
   SysFreeString bstrType
   RETURN cwsType
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the version number of a specified file.
' Example:
' DIM pFileSys AS CFileSys
' DIM cwsVersion AS CWSTR = pFileSys.GetFileVersion("C:\MyFolder\MyFile.doc")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetFileVersion (BYREF cwsFile AS CWSTR) AS CWSTR
   IF m_pFileSys = NULL THEN RETURN ""
   DIM bstrVersion AS AFX_BSTR
   DIM bstrFile AS AFX_BSTR = SysAllocString(**cwsFile)
   SetResult(m_pFileSys->GetFileVersion(bstrFile, @bstrVersion))
   DIM cwsVersion AS CWSTR = *bstrVersion
   SysFreeString bstrVersion
   RETURN cwsVersion
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the folder name for the parent of the specified folder.
' Example:
' DIM pFileSys AS CFileSys
' DIM cwsParentFolderName AS CWSTR = pFileSys.GetParentFolderName("C:\MyFolder\MySubfolder")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetParentFolderName (BYREF cwsFolder AS CWSTR) AS CWSTR
   IF m_pFileSys = NULL THEN RETURN ""
   DIM bstrName AS AFX_BSTR
   DIM bstrFolder AS AFX_BSTR = SysAllocString(**cwsFolder)
   SetResult(m_pFileSys->GetParentFolderName(bstrFolder, @bstrName))
   DIM cwsName AS CWSTR = *bstrName
   SysFreeString bstrName
   RETURN cwsName
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Sets the name of a specified folder.
' Example:
' DIM pFileSys AS CFileSys
' pFileSys.SetFolderName("c:\MyFolder", "NewName")
' Note: You only have to pass the new name of the folder, not the full path.
' ========================================================================================
PRIVATE FUNCTION CFileSys.SetFolderName (BYREF cwsFolder AS CWSTR, BYREF cwsName AS CWSTR) AS HRESULT
   IF m_pFileSys = NULL THEN RETURN E_POINTER
   ' // Get a reference to the IFolder interface
   DIM pFolder AS Afx_IFolder PTR
   DIM bstrFolder AS AFX_BSTR = SysAllocString(**cwsFolder)
   SetResult(m_pFileSys->GetFolder(bstrFolder, @pFolder))
   SysFreeString bstrFolder
   IF pFolder = NULL THEN RETURN m_Result
   ' // Change the name
   DIM bstrName AS AFX_BSTR = SysAllocString(**cwsName)
   SetResult(pFolder->put_Name(bstrName))
   pFolder->Release
   SysFreeString bstrName
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns a string containing the drive letter for a specified folder.
' Example:
' DIM pFileSys AS CFileSys
' DIM cwsDriveLetter AS CWSTR = pFileSys.GetFolderDriveLetter("c:\MyFolder)
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetFolderDriveLetter (BYREF cwsFolder AS CWSTR) AS CWSTR
   IF m_pFileSys = NULL THEN RETURN ""
   ' // Get a reference to the IFolder interface
   DIM bstrFolder AS AFX_BSTR = SysALlocString(**cwsFolder)
   DIM pFolder AS Afx_IFolder PTR
   SetResult(m_pFileSys->GetFolder(bstrFolder, @pFolder))
   SysFreeString bstrFolder
   IF pFolder = NULL THEN RETURN ""
   ' // Get a reference to the IDrive interfae
   DIM pDrive AS Afx_IDrive PTR
   SetResult(pFolder->get_Drive(@pDrive))
   pFolder->Release
   IF pDrive = NULL THEN RETURN ""
   ' // Get the drive letter
   DIM bstrDriveLetter AS AFX_BSTR
   SetResult(pDrive->get_DriveLetter(@bstrDriveLetter))
   pDrive->Release
   DIM cwsDriveLetter AS CWSTR = *bstrDriveLetter
   SysFreeString bstrDriveLetter
   RETURN cwsDriveLetter
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Appends a name to an existing path.
' Example:
' DIM pFileSys AS CFileSys
' DIM cwsNewPath = fso.BuildPath("C:\MyFolder\", "New Subfolder")
' ========================================================================================
PRIVATE FUNCTION CFileSys.BuildPath (BYREF cwsPath AS CWSTR, BYREF cwsName AS CWSTR) AS CWSTR
   IF m_pFileSys = NULL THEN RETURN ""
   DIM bstrNewPath AS AFX_BSTR
   DIM bstrPath AS AFX_BSTR = SysAllocString(**cwsPath)
   DIM bstrName AS AFX_BSTR = SysAllocString(**cwsName)
   SetResult(m_pFileSys->BuildPath(bstrPath, bstrName, @bstrNewPath))
   SysFreeString bstrPath
   SysFreeString bstrName
   DIM cwsNewPath AS CWSTR = *bstrNewPath
   SysFreeString bstrNewPath
   RETURN cwsNewPath
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the volume serial number that the operating system assigns when a hard disk is formatted.
' For drive letters, the root drive is not included. For example, the path for the C drive
' is C:, not C:\.
' Example:
' DIM pFileSys AS CFileSys
' DIM cwsSerialNumber AS CWSTR = pFileSys.GetSerialNumber("C:")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetSerialNumber (BYREF cwsDrive AS CWSTR) AS LONG
   IF m_pFileSys = NULL THEN RETURN 0
   ' // Get a reference to the IDrive interface
   DIM pDrive AS Afx_IDrive PTR
   DIM bstrDrive AS AFX_BSTR = SysAllocString(**cwsDrive)
   SetResult(m_pFileSys->GetDrive(bstrDrive, @pDrive))
   SysFreeString bstrDrive
   IF pDrive = NULL THEN RETURN 0
   ' // Get the serial number
   DIM nSerialNumber AS LONG
   SetResult(pDrive->get_SerialNumber(@nSerialNumber))
   pDrive->Release
   RETURN nSerialNumber
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the volume name of the specified drive.
' For drive letters, the root drive is not included. For example, the path for the C drive
' is C:, not C:\.
' Example:
' DIM pFileSys AS CFileSys
' PRINT pFileSys.GetVolumeName("C:")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetVolumeName (BYREF cwsDrive AS CWSTR) AS CWSTR
   IF m_pFileSys = NULL THEN RETURN ""
   ' // Get a reference to the IDrive interface
   DIM bstrDrive AS AFX_BSTR = SysAllocString(**cwsDrive)
   DIM pDrive AS Afx_IDrive PTR
   SetResult(m_pFileSys->GetDrive(bstrDrive, @pDrive))
   SysFreeString bstrDrive
   IF pDrive = NULL THEN RETURN ""
   ' // The the volume name
   DIM bstrName AS AFX_BSTR
   SetResult(pDrive->get_VolumeName(@bstrName))
   pDrive->Release
   DIM cwsName AS CWSTR = *bstrName
   SysFreeString bstrName
   RETURN cwsName
END FUNCTION
' ========================================================================================

' ========================================================================================
' *** UNTESTED ***
' Sets the volume name of the specified drive.
' For drive letters, the root drive is not included. For example, the path for the C drive
' is C:, not C:\.
' ========================================================================================
PRIVATE FUNCTION CFileSys.SetVolumeName (BYREF cwsDrive AS CWSTR, BYREF cwsName AS CWSTR) AS HRESULT
   IF m_pFileSys = NULL THEN RETURN E_POINTER
   ' // Get a reference to the IDrive interface
   DIM pDrive AS Afx_IDrive PTR
   DIM bstrDrive AS AFX_BSTR = SysAllocString(**cwsDrive)
   SetResult(m_pFileSys->GetDrive(bstrDrive, @pDrive))
   SysFreeString bstrDrive
   IF pDrive = NULL THEN RETURN m_Result
   ' // Set the volume name
   DIM bstrName AS AFX_BSTR = SysAllocString(**cwsName)
   SetResult(pDrive->put_VolumeName(bstrName))
   SysFreeString bstrName
   pDrive->Release
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the number of drives.
' Example:
' DIM pFileSys AS CFileSys
' DIM numDrives AS LONG = pFileSys.GetNumDrives
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetNumDrives () AS LONG
   IF m_pFileSys = NULL THEN RETURN 0
   ' // Get a reference to the IDriveCollection interface
   DIM pDrives AS Afx_IDriveCollection PTR
   SetResult(m_pFileSys->get_Drives(@pDrives))
   IF pDrives = NULL THEN RETURN 0
   DIM nDrives AS LONG
   SetResult(pDrives->get_Count(@nDrives))
   pDrives->Release
   RETURN nDrives
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the number of folders contained in a specified folder, including those with
' hidden and system file attributes set.
' Example:
' DIM pFileSys AS CFileSys
' DIM numSubFolders AS LONG = pFileSys.GetNumSubFolders("C:\MyFolder")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetNumSubFolders (BYREF cwsFolder AS CWSTR) AS LONG
   IF m_pFileSys = NULL THEN RETURN 0
   ' // Get a reference to the IFolder interface
   DIM pFolder AS Afx_IFolder PTR
   DIM bstrFolder AS AFX_BSTR = SysAllocString(**cwsFolder)
   SetResult(m_pFileSys->GetFolder(bstrFolder, @pFolder))
   SysFreeString bstrFolder
   IF pFolder = NULL THEN RETURN 0
   ' // Get the number of subfolders
   DIM numSubFolders AS LONG
   DIM pSubFolders AS Afx_IFolderCollection PTR
   SetResult(pFolder->get_SubFolders(@pSubFolders))
   pFolder->Release
   IF pSubFolders = NULL THEN RETURN 0
   SetResult(pSubFolders->get_Count(@numSubFolders))
   pSubFolders->Release
   RETURN numSubFolders
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the number of files contained in a specified folder, including those with
' hidden and system file attributes set.
' Example:
' DIM pFileSys AS CFileSys
' DIM numFiles AS LONG = pFileSys.GetNumFiles("C:\MyFolder")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetNumFiles (BYREF cwsFolder AS CWSTR) AS LONG
   IF m_pFileSys = NULL THEN RETURN 0
   ' // Get a reference to the IFolder interface
   DIM pFolder AS Afx_IFolder PTR
   DIM bstrFolder AS AFX_BSTR = SysAllocString(**cwsFolder)
   SetResult(m_pFileSys->GetFolder(bstrFolder, @pFolder))
   IF pFolder = NULL THEN RETURN 0
   ' // Get a reference to the files collection
   DIM pFiles AS Afx_IFileCollection PTR
   SetResult(pFolder->get_Files(@pFiles))
   pFolder->Release
   IF pFiles = NULL THEN RETURN 0
   ' // Get the number of files
   DIM numFiles AS LONG
   SetResult(pFiles->get_Count(@numFiles))
   pFiles->Release
   RETURN numFiles
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns True(-1) if the specified folder is the root folder; False(0) if it is not.
' DIM pFileSys AS CFileSys
' PRINT pFileSys.IsRootFolder("C:\MyFolder")
' ========================================================================================
PRIVATE FUNCTION CFileSys.IsRootFolder (BYREF cwsFolder AS CWSTR) AS BOOLEAN
   IF m_pFileSys = NULL THEN RETURN FALSE
   ' // Get a reference to the IFolder interface
   DIM pFolder AS Afx_IFolder PTR
   DIM bstrFolder AS AFX_BSTR = SysAllocString(**cwsFolder)
   SetResult(m_pFileSys->GetFolder(bstrFolder, @pFolder))
   SysFreeString bstrFolder
   IF pFolder = NULL THEN RETURN FALSE
   ' // Check if it is the root folder
   DIM fRootFolder AS VARIANT_BOOL
   SetResult(pFolder->get_IsRootFolder(@fRootFolder))
   pFolder->Release
   RETURN fRootFolder
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns a semicolon separated list with the driver letters.
' Example:
' DIM pFileSys AS CFileSys
' DIM cwsLetters AS CWSTR = pFileSys.DriveLetters
' ========================================================================================
PRIVATE FUNCTION CFileSys.DriveLetters () AS CWSTR
   IF m_pFileSys = NULL THEN RETURN ""
   ' // Get a reference to the IDriveCollection interface
   DIM pDrives AS Afx_IDriveCollection PTR
   SetResult(m_pFileSys->get_Drives(@pDrives))
   IF pDrives = NULL THEN RETURN ""
   ' // Get a reference to the standard enumerator
   DIM pEnum AS IEnumVARIANT PTR
   SetResult(pDrives->get__NewEnum(@cast(IUnknown PTR, pEnum)))
   pDrives->Release
   IF pEnum = NULL THEN RETURN ""
   ' // Enumerate the collection
   DIM vItem AS VARIANT, celtFetched AS ULONG, pDrive AS Afx_IDrive PTR
   DIM nDrives AS LONG, bstrDriveLetter AS AFX_BSTR, cwsDrives AS CWSTR
   DO
      pEnum->lpvtbl->Next(pEnum, 1, @vItem, @celtFetched)
      IF celtFetched = 0 THEN EXIT DO
      pDrive = cast(Afx_IDrive PTR, cast(ANY PTR, vItem.pdispVal))
      IF pDrive = NULL THEN
         VariantClear @vItem
         EXIT DO
      END IF
      SetResult(pDrive->get_DriveLetter(@bstrDriveLetter))
      IF bstrDriveLetter THEN
         cwsDrives += *bstrDriveLetter & ";"
         SysFreeString bstrDriveLetter
      END IF
      VariantClear @vItem
   LOOP
   pEnum->lpvtbl->Release(pEnum)
   ' // Remove the last ";"
   IF LEN(cwsDrives) THEN cwsDrives = LEFT(**cwsDrives, LEN(cwsDrives) - 1)
   ' // Return the drive letters
   RETURN cwsDrives
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns a TextStream object corresponding to the standard input, output, or error stream.
' DIM pFileSys AS CFileSys
' DIM pStm AS Afx_ITextStream PTR
' pStm = pFileSys.GetStandardStream(StandardStreamTypes_StdOut)
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetStandardStream (BYVAL StandardStreamType AS STANDARDSTREAMTYPES, BYVAL bUnicode AS VARIANT_BOOL = FALSE) AS Afx_ITextStream PTR
   DIM pStm AS Afx_ITextStream PTR
   SetResult(m_pFileSys->GetStandardStream(StandardStreamType, bUnicode, @pStm))
   RETURN pStm
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the size, in bytes, of the specified file.
' Example:
' DIM pFileSys AS CFileSys
' DIM nFileSize AS LONG = pFileSys.GetFileSize("C:\MyPath\MyFile.txt")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetFileSize (BYREF cwsFile AS CWSTR) AS LONG
   IF m_pFileSys = NULL THEN RETURN 0
   ' // Get a reference to the IFile interface
   DIM pFile AS Afx_IFile PTR
   DIM bstrFile AS AFX_BSTR = SysAllocString(**cwsFile)
   SetResult(m_pFileSys->GetFile(bstrFile, @pFile))
   SysFreeString bstrFile
   IF pFile = NULL THEN RETURN 0
   ' // Get the file size
   DIM vSize AS VARIANT
   SetResult(pFile->get_Size(@vSize))
   pFile->Release
   IF m_Result = S_OK THEN RETURN vSize.lVal
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Sets the name of a specified file.
' Example:
' DIM pFileSys AS CFileSys
' pFileSys.SetFileName("c:\MyFolder\Test.txt", "NewName")
' ========================================================================================
PRIVATE FUNCTION CFileSys.SetFileName (BYREF cwsFile AS CWSTR, BYREF cwsName AS CWSTR) AS HRESULT
   IF m_pFileSys = NULL THEN RETURN E_POINTER
   ' // Get a reference to the IFile interface
   DIM pFile AS Afx_IFile PTR
   DIM bstrFile AS AFX_BSTR = SysAllocString(**cwsFile)
   SetResult(m_pFileSys->GetFile(bstrFile, @pFile))
   SysFreeString bstrFile
   IF pFile = NULL THEN RETURN m_Result
   ' // Set the file name
   DIM bstrName AS AFX_BSTR = SysAllocString(**cwsName)
   SetResult(pFile->put_Name(bstrName))
   SysFreeString bstrName
   pFile->Release
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the short path used by programs that require the earlier 8.3 file naming convention.
' Example:
' DIM pFileSys AS CFileSys
' DIM cbsFileName CBSTR = pFileSys.GetFileShortPath("C:\MyPath\MyFile.txt")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetFileShortPath (BYREF cwsFile AS CWSTR) AS CWSTR
   IF m_pFileSys = NULL THEN RETURN ""
   ' // Get a reference to the IFile interface
   DIM bstrFile AS AFX_BSTR = SysAllocString(**cwsFile)
   DIM pFile AS Afx_IFile PTR
   SetResult(m_pFileSys->GetFile(bstrFile, @pFile))
   SysFreeString bstrFile
   IF pFile = NULL THEN RETURN ""
   ' // Get the short path name
   DIM bstrPath AS AFX_BSTR
   SetResult(pFile->get_ShortPath(@bstrPath))
   pFile->Release
   DIM cwsPath AS CWSTR = *bstrPath
   SysFreeString bstrPath
   RETURN cwsPath
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the short name used by programs that require the earlier 8.3 file naming convention.
' Example:
' DIM pFileSys AS CFileSys
' DIM cwsFileName CWSTR = pFileSys.GetFileShortName("C:\MyPath\MyFile.txt")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetFileShortName (BYREF cwsFile AS CWSTR) AS CWSTR
   IF m_pFileSys = NULL THEN RETURN ""
   ' // Get a reference to the IFile interface
   DIM pFile AS Afx_IFile PTR
   DIM bstrFile AS AFX_BSTR = SysAllocString(**cwsFile)
   SetResult(m_pFileSys->GetFile(bstrFile, @pFile))
   SysFreeString bstrFile
   IF pFile = NULL THEN RETURN ""
   ' // Get the file short name
   DIM bstrName AS AFX_BSTR
   SetResult(pFile->get_ShortName(@bstrName))
   pFile->Release
   DIM cwsName AS CWSTR = *bstrName
   SysFreeString bstrName
   RETURN cwsName
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the attributes of the specified file.
' Example:
' DIM pFileSys AS CFileSys
' DIM lAttr FILEATTRIBUTE = pFileSys.GetFileAttributes("C:\MyPath\MyFile.txt")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetFileAttributes (BYREF cwsFile AS CWSTR) AS FILEATTRIBUTE
   DIM lAttr AS FILEATTRIBUTE
   IF m_pFileSys = NULL THEN RETURN lAttr
   ' // Get a reference to the IFile interface
   DIM pFile AS Afx_IFile PTR
   DIM bstrFile AS AFX_BSTR = SysAllocString(**cwsFile)
   SetResult(m_pFileSys->GetFile(bstrFile, @pFile))
   SysFreeString bstrFile
   IF pFile = NULL THEN RETURN 0
   ' // Get the file attributes
   SetResult(pFile->get_Attributes(@lAttr))
   pFile->Release
   RETURN lAttr
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Sets the attributes of the specified file.
' Example:
' DIM pFileSys AS CFileSys
' pFileSys.SetFileAttributes("C:\MyPath\MyFile.txt", 33)
' ========================================================================================
PRIVATE FUNCTION CFileSys.SetFileAttributes (BYREF cwsFile AS CWSTR, BYVAL lAttr AS FILEATTRIBUTE) AS HRESULT
   IF m_pFileSys = NULL THEN RETURN E_POINTER
   ' // Get a reference to the IFile interface
   DIM pFile AS Afx_IFile PTR, bstrFile AS AFX_BSTR
   bstrFile = SysAllocString(**cwsFile)
   SetResult(m_pFileSys->GetFile(bstrFile, @pFile))
   SysFreeString bstrFile
   IF pFile = NULL THEN RETURN m_Result
   ' // Set the file attributes
   SetResult(pFile->put_Attributes(lAttr))
   pFile->Release
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the date and time that the specified file was created.
' Example:
' DIM pFileSys AS CFileSys
' DIM nDate AS DATE_ = pFileSys.GetFileDateCreated("C:\MyPath\MyFile.txt")
' Note: To convert the date to a string, call AfxVariantDateTimeToStr.
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetFileDateCreated (BYREF cwsFile AS CWSTR) AS DATE_
   DIM nDate AS DATE_
   IF m_pFileSys = NULL THEN RETURN nDate
   ' // Get a reference to the IFile interface
   DIM pFile AS Afx_IFile PTR
   DIM bstrFile AS AFX_BSTR = SysAllocString(**cwsFile)
   SetResult(m_pFileSys->GetFile(bstrFile, @pFile))
   SysFreeString bstrFile
   IF pFile = NULL THEN RETURN 0
   ' // Get the date
   SetResult(pFile->get_DateCreated(@nDate))
   pFile->Release
   RETURN nDate
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the date and time that the specified file was last modified.
' Example:
' DIM pFileSys AS CFileSys
' DIM nDate AS DATE_ = pFileSys.GetFileDateLastModified("C:\MyPath\MyFile.txt")
' Note: To convert the date to a string, call AfxVariantDateTimeToStr.
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetFileDateLastModified (BYREF cwsFile AS CWSTR) AS DATE_
   DIM nDate AS DATE_
   IF m_pFileSys = NULL THEN RETURN nDate
   ' // Get a reference to the IFile interface
   DIM pFile AS Afx_IFile PTR
   DIM bstrFile AS AFX_BSTR = SysAllocString(**cwsFile)
   SetResult(m_pFileSys->GetFile(bstrFile, @pFile))
   SysFreeString bstrFile
   IF pFile = NULL THEN RETURN 0
   ' // Get the date
   SetResult(pFile->get_DateLastModified(@nDate))
   pFile->Release
   RETURN nDate
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the date and time that the specified file was last modified.
' Example:
' DIM pFileSys AS CFileSys
' DIM nDate AS DATE_ = pFileSys.GetFileDateLastAccessed("C:\MyPath\MyFile.txt")
' Note: To convert the date to a string, call AfxVariantDateTimeToStr.
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetFileDateLastAccessed (BYREF cwsFile AS CWSTR) AS DATE_
   DIM nDate AS DATE_
   IF m_pFileSys = NULL THEN RETURN nDate
   ' // Get a reference to the IFile interface
   DIM pFile AS Afx_IFile PTR
   DIM bstrFile AS AFX_BSTR = SysAllocString(**cwsFile)
   SetResult(m_pFileSys->GetFile(bstrFile, @pFile))
   SysFreeString bstrFile
   IF pFile = NULL THEN RETURN 0
   ' // Get the date
   SetResult(pFile->get_DateLastAccessed(@nDate))
   pFile->Release
   RETURN nDate
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the size, in bytes, of the specified file.
' Example:
' DIM pFileSys AS CFileSys
' DIM cbsFolderSize AS LONG = pFileSys.GetFolderSize("C:\MyPath")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetFolderSize (BYREF cwsFolder AS CWSTR) AS LONG
   IF m_pFileSys = NULL THEN RETURN 0
   ' // Get a reference to the IFolder interface
   DIM pFolder AS Afx_IFolder PTR
   DIM bstrFolder AS AFX_BSTR = SysAllocString(**cwsFolder)
   SetResult(m_pFileSys->GetFolder(bstrFolder, @pFolder))
   SysFreeString bstrFolder
   IF pFolder = NULL THEN RETURN 0
   ' // Get the size
   DIM vSize AS VARIANT
   SetResult(pFolder->get_Size(@vSize))
   pFolder->Release
   IF m_Result = S_OK THEN RETURN vSize.lVal
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the attributes of the specified folder.
' Example:
' DIM pFileSys AS CFileSys
' DIM lAttr FILEATTRIBUTE = pFileSys.GetFolderAttributes("C:\MyPath")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetFolderAttributes (BYREF cwsFolder AS CWSTR) AS FILEATTRIBUTE
   DIM lAttr AS FILEATTRIBUTE
   IF m_pFileSys = NULL THEN RETURN lAttr
   ' // Get a reference to the IFolder interface
   DIM pFolder AS Afx_IFolder PTR
   DIM bstrFolder AS AFX_BSTR = SysAllocString(**cwsFolder)
   SetResult(m_pFileSys->GetFolder(bstrFolder, @pFolder))
   SysFreeString bstrFolder
   IF pFolder = NULL THEN RETURN 0
   ' // Get the attributes
   SetResult(pFolder->get_Attributes(@lAttr))
   pFolder->Release
   RETURN lAttr
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the attributes of the specified folder.
' Example:
' DIM pFileSys AS CFileSys
' DIM lAttr FILEATTRIBUTE = pFileSys.SetFolderAttributes("C:\MyPath")
' ========================================================================================
PRIVATE FUNCTION CFileSys.SetFolderAttributes (BYREF cwsFolder AS CWSTR, BYVAL lAttr AS FILEATTRIBUTE) AS HRESULT
   IF m_pFileSys = NULL THEN RETURN E_POINTER
   ' // Get a reference to the IFolder interface
   DIM pFolder AS Afx_IFolder PTR
   DIM bstrFolder AS AFX_BSTR = SysAllocString(**cwsFolder)
   SetResult(m_pFileSys->GetFolder(bstrFolder, @pFolder))
   SysFreeString bstrFolder
   IF pFolder = NULL THEN RETURN m_Result
   ' // Set the attributes
   SetResult(pFolder->put_Attributes(lAttr))
   pFolder->Release
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the date and time that the specified folder was created.
' Example:
' DIM pFileSys AS CFileSys
' DIM nDate AS DATE_ = pFileSys.GetFolderDateCreated("C:\MyPath")
' Note: To convert the date to a string, call AfxVariantDateTimeToStr.
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetFolderDateCreated (BYREF cwsFolder AS CWSTR) AS DATE_
   DIM nDate AS DATE_
   IF m_pFileSys = NULL THEN RETURN nDate
   ' // Get a reference to the IFolder interface
   DIM pFolder AS Afx_IFolder PTR
   DIM bstrFolder AS AFX_BSTR = SysAllocString(**cwsFolder)
   SetResult(m_pFileSys->GetFolder(bstrFolder, @pFolder))
   SysFreeString bstrFolder
   IF pFolder = NULL THEN RETURN 0
   ' // Get the date
   SetResult(pFolder->get_DateCreated(@nDate))
   pFolder->Release
   RETURN nDate
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the date and time that the specified file was last modified.
' Example:
' DIM pFileSys AS CFileSys
' DIM nDate AS DATE_ = pFileSys.GetFolderDateLastModified("C:\MyPath")
' Note: To convert the date to a string, call AfxVariantDateTimeToStr.
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetFolderDateLastModified (BYREF cwsFolder AS CWSTR) AS DATE_
   DIM nDate AS DATE_
   IF m_pFileSys = NULL THEN RETURN nDate
   ' // Get a reference to the IFolder interface
   DIM pFolder AS Afx_IFolder PTR
   DIM bstrFolder AS AFX_BSTR = SysAllocString(**cwsFolder)
   SetResult(m_pFileSys->GetFolder(bstrFolder, @pFolder))
   SysFreeString bstrFolder
   IF pFolder = NULL THEN RETURN 0
   ' // Get the date
   SetResult(pFolder->get_DateLastModified(@nDate))
   pFolder->Release
   RETURN nDate
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the date and time that the specified file was last modified.
' Example:
' DIM pFileSys AS CFileSys
' DIM nDate AS DATE_ = pFileSys.GetFolderDateLastAccessed("C:\MyPath")
' Note: To convert the date to a string, call AfxVariantDateTimeToStr.
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetFolderDateLastAccessed (BYREF cwsFolder AS CWSTR) AS DATE_
   DIM nDate AS DATE_
   IF m_pFileSys = NULL THEN RETURN nDate
   ' // Get a reference to the IFolder interface
   DIM pFolder AS Afx_IFolder PTR
   DIM bstrFolder AS AFX_BSTR = SysAllocString(**cwsFolder)
   SetResult(m_pFileSys->GetFolder(bstrFolder, @pFolder))
   SysFreeString bstrFolder
   IF pFolder = NULL THEN RETURN 0
   ' // Get the date
   SetResult(pFolder->get_DateLastAccessed(@nDate))
   pFolder->Release
   RETURN nDate
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the amount of space available to a user on the specified drive or network share.
' Example:
' DIM pFileSys AS CFileSys
' PRINT pFileSys.GetDriveAvailableSpace("C:")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetDriveAvailableSpace (BYREF cwsDrive AS CWSTR) AS DOUBLE
   IF m_pFileSys = NULL THEN RETURN 0
   ' // Get a reference to the IDrive interface
   DIM pDrive AS Afx_IDrive PTR
   DIM bstrDrive AS AFX_BSTR = SysAllocString(**cwsDrive)
   SetResult(m_pFileSys->GetDrive(bstrDrive, @pDrive))
   SysFreeString bstrDrive
   IF pDrive = NULL THEN RETURN 0
   ' // Get the acailable space
   DIM vSize AS VARIANT
   SetResult(pDrive->get_AvailableSpace(@vSize))
   pDrive->Release
   RETURN vSize.dblVal
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the amount of free space available to a user on the specified drive or network share.
' Example:
' DIM pFileSys AS CFileSys
' PRINT pFileSys.GetDriveFreeSpace("C:")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetDriveFreeSpace (BYREF cwsDrive AS CWSTR) AS DOUBLE
   IF m_pFileSys = NULL THEN RETURN 0
   ' // Get a reference to the IDrive interface
   DIM pDrive AS Afx_IDrive PTR
   DIM bstrDrive AS AFX_BSTR = SysAllocString(**cwsDrive)
   SetResult(m_pFileSys->GetDrive(bstrDrive, @pDrive))
   SysFreeString bstrDrive
   IF pDrive = NULL THEN RETURN 0
   ' // Get the free space
   DIM vSize AS VARIANT
   SetResult(pDrive->get_FreeSpace(@vSize))
   pDrive->Release
   RETURN vSize.dblVal
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the total space, in bytes, of a drive or network share.
' Example:
' DIM pFileSys AS CFileSys
' PRINT pFileSys.GetDriveTotalSize("C:")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetDriveTotalSize (BYREF cwsDrive AS CWSTR) AS DOUBLE
   IF m_pFileSys = NULL THEN RETURN 0
   ' // Get a reference to the IDrive interface
   DIM pDrive AS Afx_IDrive PTR
   DIM bstrDrive AS AFX_BSTR = SysAllocString(**cwsDrive)
   SetResult(m_pFileSys->GetDrive(bstrDrive, @pDrive))
   SysFreeString bstrDrive
   IF pDrive = NULL THEN RETURN 0
   ' // Get the total space
   DIM vSize AS VARIANT
   SetResult(pDrive->get_TotalSize(@vSize))
   pDrive->Release
   RETURN vSize.dblVal
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Creates a folder. An error occurs if the specified folder already exists.
' Example:
' DIM pFileSys AS CFileSys
' DIM pFolder AS Afx_Folder PTR
' pFolder = pFileSys.CreateFolder("C:\MyNewFolder")
' IF pFolder THEN
'  ....
'  pFolder-Release
' END IF
' ========================================================================================
PRIVATE FUNCTION CFileSys.CreateFolder (BYREF cwsFolder AS CWSTR) AS Afx_IFolder PTR
   IF m_pFileSys = NULL THEN RETURN NULL
   DIM pFolder AS Afx_IFolder PTR, bstrFolder AS AFX_BSTR
   bstrFolder = SysAllocString(**cwsFolder)
   SetResult(m_pFileSys->CreateFolder(bstrFolder, @pFolder))
   SysFreeString bstrFolder
   RETURN pFolder
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Deletes a specified folder and its contents.
' Example:
' DIM pFileSys AS CFileSys
' pFileSys.DeleteFolder("C:\MyFolder")
' ========================================================================================
PRIVATE FUNCTION CFileSys.DeleteFolder (BYREF cwsFolder AS CWSTR, BYVAL bForce AS VARIANT_BOOL = FALSE) AS HRESULT
   IF m_pFileSys = NULL THEN RETURN E_POINTER
   DIM bstrFolder AS AFX_BSTR = SysAllocString(**cwsFolder)
   SetResult(m_pFileSys->DeleteFolder(bstrFolder, bForce))
   SysFreeString bstrFolder
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Moves one or more folders from one location to another.
' Parameters:
' - cbsSource: The path to the folder or folders to be moved. The cbsSource argument string
' can contain wildcard characters in the last path component only.
' - cbsDestination: The path where the folder or folders are to be moved. The cbsDestination
' argument can't contain wildcard characters.
' Example:
' DIM pFileSys AS CFileSys
' pFileSys.MoveFolder("C:\MyFolder", "C:\MyNewFolder\")
' ========================================================================================
PRIVATE FUNCTION CFileSys.MoveFolder (BYREF cwsSource AS CWSTR, BYREF cwsDestination AS CWSTR) AS HRESULT
   IF m_pFileSys = NULL THEN RETURN E_POINTER
   DIM bstrSource AS AFX_BSTR = SysAllocString(**cwsSource)
   DIM bstrDestination AS AFX_BSTR = SysAllocString(**cwsDestination)
   SetResult(m_pFileSys->MoveFolder(bstrSource, bstrDestination))
   SysFreeString bstrSource
   SysFreeString bstrDestination
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Recursively copies a folder from one location to another.
' Parameters:
' - cbsSource: Character string file specification, which can include wildcard characters,
' for one or more folders to be copied.
' - cbsDestination: Character string destination where the folder and subfolder from source
' are to be copied. Wildcard characters are not allowed.
' - Boolean value that indicates if existing folders are to be overwritten. If true, files
' are overwritten; if false, they are not. The default is true.
' Example:
' DIM pFileSys AS CFileSys
' pFileSys.CopyFolder("C:\MyFolder", "C:\MyOtherFolder\")
' ========================================================================================
PRIVATE FUNCTION CFileSys.CopyFolder (BYREF cwsSource AS CWSTR, BYREF cwsDestination AS CWSTR, BYVAL OverWriteFiles AS VARIANT_BOOL = -1) AS HRESULT
   IF m_pFileSys = NULL THEN RETURN E_POINTER
   DIM bstrSource AS AFX_BSTR = SysAllocString(**cwsSource)
   DIM bstrDestination AS AFX_BSTR = SysAllocString(**cwsDestination)
   SetResult(m_pFileSys->CopyFolder(bstrSource, bstrDestination, OverWriteFiles))
   SysFreeString bstrSource
   SysFreeString bstrDestination
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Copies one or more files from one location to another.
' Parameters:
' - cbsSource: Character string file specification, which can include wildcard characters,
' for one or more files to be copied.
' - cbsDestination: Character string destination where the file or files from source
' are to be copied. Wildcard characters are not allowed.
' - OverWriteFiles: Boolean value that indicates if existing files are to be overwritten.
' If true, files are overwritten; if false, they are not. The default is true. Note that
' CopyFile will fail if destination has the read-only attribute set, regardless of the
' value of overwrite.
' Example:
' DIM pFileSys AS CFileSys
' pFileSys.CopyFile("C:\MyFolder\MyFile.txt", "C:\MyOtherFolder\MyFile.txt")
' ========================================================================================
PRIVATE FUNCTION CFileSys.CopyFile (BYREF cwsSource AS CWSTR, BYREF cwsDestination AS CWSTR, BYVAL OverWriteFiles AS VARIANT_BOOL = -1) AS HRESULT
   IF m_pFileSys = NULL THEN RETURN E_POINTER
   DIM bstrSource AS AFX_BSTR = SysAllocString(**cwsSource)
   DIM bstrDestination AS AFX_BSTR = SysAllocString(**cwsDestination)
   SetResult(m_pFileSys->CopyFile(bstrSource, bstrDestination, OverWriteFiles))
   SysFreeString bstrSource
   SysFreeString bstrDestination
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Deletes a specified file.
' Parameters:
' - cbsFileSpec: The name of the file to delete. cbsFileSpec can contain wildcard characters
' in the last path component.
' - bForce: Boolean value that is true if files with the read-only attribute set are to be
' deleted; false (default) if they are not.
' Example:
' DIM pFileSys AS CFileSys
' pFileSys.DeleteFile("C:\MyFolder\MyFile.txt")
' ========================================================================================
PRIVATE FUNCTION CFileSys.DeleteFile (BYREF cwsFileSpec AS CWSTR, BYVAL bForce AS VARIANT_BOOL = FALSE) AS HRESULT
   IF m_pFileSys = NULL THEN RETURN E_POINTER
   DIM bstrFileSpec AS AFX_BSTR = SysAllocString(**cwsFileSpec)
   SetResult(m_pFileSys->DeleteFile(bstrFileSpec, bForce))
   SysFreeString bstrFileSpec
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Moves one or more files from one location to another.
' Parameters:
' - cbsSource: The path to the file or files to be moved. The cbsSource argument string
' can contain wildcard characters in the last path component only.
' - cbsDestination: The path where the file or files are to be moved. The cbsDestination
' argument can't contain wildcard characters.
' Example:
' DIM pFileSys AS CFileSys
' pFileSys.MoveFile("C:\MyFolder\MyFile.txt", "C:\MyOtherFolder\")
' ========================================================================================
PRIVATE FUNCTION CFileSys.MoveFile (BYREF cwsSource AS CWSTR, BYREF cwsDestination AS CWSTR) AS HRESULT
   IF m_pFileSys = NULL THEN RETURN E_POINTER
   DIM bstrSource AS AFX_BSTR = SysAllocString(**cwsSource)
   DIM bstrDestination AS AFX_BSTR = SysAllocString(**cwsDestination)
   SetResult(m_pFileSys->MoveFile(bstrSource, bstrDestination))
   SysFreeString bstrSource
   SysFreeString bstrDestination
   RETURN m_Result
END FUNCTION
' ========================================================================================

END NAMESPACE
