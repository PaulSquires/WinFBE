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

Public:
   m_Result AS HRESULT
   m_pFileSys AS Afx_IFileSystem3 PTR

Private:
   m_bUninitCOM AS BOOLEAN

Public:
   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR
   DECLARE FUNCTION GetLastResult () AS HRESULT
   DECLARE FUNCTION SetResult (BYVAL Result AS HRESULT) AS HRESULT
   DECLARE FUNCTION DriveLetters () AS CBSTR
   DECLARE FUNCTION DriveExists (BYREF cbsDrive AS CBSTR) AS BOOLEAN
   DECLARE FUNCTION FileExists (BYREF cbsFileSpec AS CBSTR) AS BOOLEAN
   DECLARE FUNCTION FolderExists (BYREF cbsFileSpec AS CBSTR) AS BOOLEAN
   DECLARE FUNCTION IsDriveReady (BYREF cbsDrive AS CBSTR) AS BOOLEAN
   DECLARE FUNCTION GetDriveName (BYREF cbsPathSpec AS CBSTR) AS CBSTR
   DECLARE FUNCTION GetFolderName (BYREF cbsPathSpec AS CBSTR) AS CBSTR
   DECLARE FUNCTION GetFileName (BYREF cbsPathSpec AS CBSTR) AS CBSTR
   DECLARE FUNCTION GetExtensionName (BYREF cbsPathSpec AS CBSTR) AS CBSTR
   DECLARE FUNCTION GetBaseName (BYREF cbsPathSpec AS CBSTR) AS CBSTR
   DECLARE FUNCTION GetAbsolutePathName (BYREF cbsPathSpec AS CBSTR) AS CBSTR
   DECLARE FUNCTION GetTempName () AS CBSTR
   DECLARE FUNCTION GetDriveShareName (BYREF cbsDrive AS CBSTR) AS CBSTR
   DECLARE FUNCTION GetDriveFileSystem (BYREF cbsDrive AS CBSTR) AS CBSTR
   DECLARE FUNCTION GetDriveType (BYREF cbsDrive AS CBSTR) AS DRIVETYPECONST
   DECLARE FUNCTION GetFolderType (BYREF cbsFolder AS CBSTR) AS CBSTR
   DECLARE FUNCTION GetFileType (BYREF cbsFile AS CBSTR) AS CBSTR
   DECLARE FUNCTION GetFileVersion (BYREF cbsFileName AS CBSTR) AS CBSTR
   DECLARE FUNCTION GetParentFolderName (BYREF cbsFolder AS CBSTR) AS CBSTR
   DECLARE FUNCTION GetFolderShortPath (BYREF cbsFolder AS CBSTR) AS CBSTR
   DECLARE FUNCTION GetFolderShortName (BYREF cbsFolder AS CBSTR) AS CBSTR
   DECLARE FUNCTION GetFolderDriveLetter (BYREF cbsFolder AS CBSTR) AS CBSTR
   DECLARE FUNCTION SetFolderName (BYREF cbsFolder AS CBSTR, BYREF cbsName AS CBSTR) AS HRESULT
   DECLARE FUNCTION BuildPath (BYREF cbsPath AS CBSTR, BYREF cbsName AS CBSTR) AS CBSTR
   DECLARE FUNCTION GetSerialNumber (BYREF cbsDrive AS CBSTR) AS LONG
   DECLARE FUNCTION GetVolumeName (BYREF cbsDrive AS CBSTR) AS CBSTR
   DECLARE FUNCTION SetVolumeName (BYREF cbsDrive AS CBSTR, BYREF cbsName AS CBSTR) AS HRESULT
   DECLARE FUNCTION GetNumDrives () AS LONG
   DECLARE FUNCTION GetNumSubFolders (BYREF cbsFolder AS CBSTR) AS LONG
   DECLARE FUNCTION GetNumFiles (BYREF cbsFolder AS CBSTR) AS LONG
   DECLARE FUNCTION IsRootFolder (BYREF cbsFolder AS CBSTR) AS BOOLEAN
   DECLARE FUNCTION GetStandardStream (BYVAL StandardStreamType AS STANDARDSTREAMTYPES, BYVAL bUnicode AS VARIANT_BOOL = FALSE) AS Afx_ITextStream PTR
   DECLARE FUNCTION GetFileSize (BYREF cbsFile AS CBSTR) AS LONG
   DECLARE FUNCTION SetFileName (BYREF cbsFile AS CBSTR, BYREF cbsName AS CBSTR) AS HRESULT
   DECLARE FUNCTION GetFileShortPath (BYREF cbsFile AS CBSTR) AS CBSTR
   DECLARE FUNCTION GetFileShortName (BYREF cbsFile AS CBSTR) AS CBSTR
   DECLARE FUNCTION GetFileAttributes (BYREF cbsFile AS CBSTR) AS FILEATTRIBUTE
   DECLARE FUNCTION SetFileAttributes (BYREF cbsFile AS CBSTR, BYVAL lAttr AS FILEATTRIBUTE) AS HRESULT
   DECLARE FUNCTION GetFileDateCreated (BYREF cbsFile AS CBSTR) AS DATE_
   DECLARE FUNCTION GetFileDateLastModified (BYREF cbsFile AS CBSTR) AS DATE_
   DECLARE FUNCTION GetFileDateLastAccessed (BYREF cbsFile AS CBSTR) AS DATE_
   DECLARE FUNCTION GetFolderSize (BYREF cbsFolder AS CBSTR) AS LONG
   DECLARE FUNCTION GetFolderAttributes (BYREF cbsFolder AS CBSTR) AS FILEATTRIBUTE
   DECLARE FUNCTION SetFolderAttributes (BYREF cbsFolder AS CBSTR, BYVAL lAttr AS FILEATTRIBUTE) AS HRESULT
   DECLARE FUNCTION GetFolderDateCreated (BYREF cbsFolder AS CBSTR) AS DATE_
   DECLARE FUNCTION GetFolderDateLastModified (BYREF cbsFolder AS CBSTR) AS DATE_
   DECLARE FUNCTION GetFolderDateLastAccessed (BYREF cbsFolder AS CBSTR) AS DATE_
   DECLARE FUNCTION GetDriveAvailableSpace (BYREF cbsDrive AS CBSTR) AS DOUBLE
   DECLARE FUNCTION GetDriveFreeSpace (BYREF cbsDrive AS CBSTR) AS DOUBLE
   DECLARE FUNCTION GetDriveTotalSize (BYREF cbsDrive AS CBSTR) AS DOUBLE
   DECLARE FUNCTION CreateFolder (BYREF cbsFolder AS CBSTR) AS Afx_IFolder PTR
   DECLARE FUNCTION DeleteFolder (BYREF cbsFolder AS CBSTR, BYVAL bForce AS VARIANT_BOOL = FALSE) AS HRESULT
   DECLARE FUNCTION MoveFolder (BYREF cbsSource AS CBSTR, BYREF cbsDestination AS CBSTR) AS HRESULT
   DECLARE FUNCTION CopyFolder (BYREF cbsSource AS CBSTR, BYREF cbsDestination AS CBSTR, BYVAL OverWriteFiles AS VARIANT_BOOL = -1) AS HRESULT
   DECLARE FUNCTION CopyFile (BYREF cbsSource AS CBSTR, BYREF cbsDestination AS CBSTR, BYVAL OverWriteFiles AS VARIANT_BOOL = -1) AS HRESULT
   DECLARE FUNCTION DeleteFile (BYREF cbsFileSpec AS CBSTR, BYVAL bForce AS VARIANT_BOOL = FALSE) AS HRESULT
   DECLARE FUNCTION MoveFile (BYREF cbsSource AS CBSTR, BYREF cbsDestination AS CBSTR) AS HRESULT

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
'   DIM IID_IFileSystem AS GUID = (&h0AB5A3D0, &hE5B6, &h11D0, {&hAB, &hF5, &h00, &hA0, &hC9, &h0F, &hFF, &hC0})
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
PRIVATE FUNCTION CFileSys.DriveExists (BYREF cbsDrive AS CBSTR) AS BOOLEAN
   IF m_pFileSys = NULL THEN RETURN FALSE
   DIM pfExists AS VARIANT_BOOL
   SetResult(m_pFileSys->DriveExists(cbsDrive, @pfExists))
   RETURN pfExists
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns True if the specified folder exists; False if it does not.
' Example:
' DIM pFileSys AS CFileSys
' DIM fExists AS BOOLEAN = pFileSys.FolderExists("C:\MyFolder")
' ========================================================================================
PRIVATE FUNCTION CFileSys.FolderExists (BYREF cbsFolderSpec AS CBSTR) AS BOOLEAN
   IF m_pFileSys = NULL THEN RETURN FALSE
   DIM pfExists AS VARIANT_BOOL
   SetResult(m_pFileSys->FolderExists(cbsFolderSpec, @pfExists))
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
PRIVATE FUNCTION CFileSys.IsDriveReady (BYREF cbsDrive AS CBSTR) AS BOOLEAN
   IF m_pFileSys = NULL THEN RETURN FALSE
   ' // Get a reference to the IDrive interface
   DIM pDrive AS Afx_IDrive PTR
   SetResult(m_pFileSys->GetDrive(cbsDrive, @pDrive))
   DIM fIsReady AS VARIANT_BOOL
   IF pDrive THEN
      SetResult(pDrive->get_IsReady(@fIsReady))
      pDrive->Release
   END IF
   RETURN fIsReady
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns True if the specified file exists; False if it does not.
' Example:
' DIM pFileSys AS CFileSys
' DIM fExists AS BOOLEAN = pFileSys.FileExists("C:\MyFolder\Test.txt")
' ========================================================================================
PRIVATE FUNCTION CFileSys.FileExists (BYREF cbsFileSpec AS CBSTR) AS BOOLEAN
   IF m_pFileSys = NULL THEN RETURN FALSE
   DIM pfExists AS VARIANT_BOOL
   SetResult(m_pFileSys->FileExists(cbsFileSpec, @pfExists))
   RETURN pfExists
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns a string containing the name of the drive for a specified path.
' Example:
' DIM pFileSys AS CFileSys
' DIM cbsName AS CBSTR = pFileSys.GetDriveName("C:\MyFolder\Test.txt")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetDriveName (BYREF cbsPathSpec AS CBSTR) AS CBSTR
   IF m_pFileSys = NULL THEN RETURN ""
   DIM bstrName AS AFX_BSTR
   SetResult(m_pFileSys->GetDriveName(cbsPathSpec, @bstrName))
   RETURN bstrName
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns a string containing the name of the folder for a specified path, i.e. the
' path minus the file name.
' Example:
' DIM pFileSys AS CFileSys
' DIM cbsName AS CBSTR = pFileSys.GetFolderName("C:\MyFolder\Test.txt")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetFolderName (BYREF cbsPathSpec AS CBSTR) AS CBSTR
   IF m_pFileSys = NULL THEN RETURN ""
   ' // Remove the file name
   DIM cbsFolderName AS CBSTR = cbsPathSpec
   DIM cbsFileName AS CBSTR, bstrFileName AS AFX_BSTR
   IF INSTRREV(**cbsPathSpec, ".") THEN SetResult(m_pFileSys->GetFileName(cbsPathSpec, @bstrFileName))
   cbsFileName = bstrFileName
   IF LEN(cbsFileName) THEN cbsFolderName = LEFT(**cbsFolderName, LEN(cbsFolderName) - LEN(cbsFileName) - 1)
   IF RIGHT(**cbsFolderName, 1) = "\" THEN cbsFolderName = LEFT(**cbsFolderName, LEN(cbsFolderName) - 1)
   RETURN cbsFolderName
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the short path used by programs that require the earlier 8.3 file naming convention.
' Example:
' DIM pFileSys AS CFileSys
' DIM cbsFolderShortPath AS CBSTR = pFileSys.GetFolderShortPath("c:\MyFolder)
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetFolderShortPath (BYREF cbsFolder AS CBSTR) AS CBSTR
   IF m_pFileSys = NULL THEN RETURN ""
   ' // Get a reference to the IFolder interface
   DIM bstrPath AS AFX_BSTR
   DIM pFolder AS Afx_IFolder PTR
   SetResult(m_pFileSys->GetFolder(cbsFolder, @pFolder))
   IF pFolder THEN
      SetResult(pFolder->get_ShortPath(@bstrPath))
      pFolder->Release
   END IF
   RETURN bstrPath
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the short name used by programs that require the earlier 8.3 file naming convention.
' Example:
' DIM pFileSys AS CFileSys
' DIM cbsFolderShortname AS CBSTR = pFileSys.GetFolderShortName("c:\MyFolder)
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetFolderShortName (BYREF cbsFolder AS CBSTR) AS CBSTR
   IF m_pFileSys = NULL THEN RETURN ""
   ' // Get a reference to the IFolder interface
   DIM bstrName AS AFX_BSTR
   DIM pFolder AS Afx_IFolder PTR
   SetResult(m_pFileSys->GetFolder(cbsFolder, @pFolder))
   IF pFolder THEN
      SetResult(pFolder->get_ShortName(@bstrName))
      pFolder->Release
   END IF
   RETURN bstrName
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns a string containing the name of the file for a specified path.
' Example:
' DIM pFileSys AS CFileSys
' DIM cbsName AS CBSTR = pFileSys.GetFileName("C:\MyFolder\Test.txt")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetFileName (BYREF cbsPathSpec AS CBSTR) AS CBSTR
   IF m_pFileSys = NULL THEN RETURN ""
   DIM bstrName AS AFX_BSTR
   IF INSTRREV(**cbsPathSpec, ".") THEN SetResult(m_pFileSys->GetFileName(cbsPathSpec, @bstrName))
   RETURN bstrName
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns a string containing the extension name of the file for a specified path.
' Example:
' DIM pFileSys AS CFileSys
' DIM cbsName AS CBSTR = pFileSys.GetExtensionName("C:\MyFolder\Test.txt")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetExtensionName (BYREF cbsPathSpec AS CBSTR) AS CBSTR
   IF m_pFileSys = NULL THEN RETURN ""
   DIM bstrName AS AFX_BSTR
   SetResult(m_pFileSys->GetExtensionName(cbsPathSpec, @bstrName))
   RETURN bstrName
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns a string containing the base name of the last component, less any file
' extension, in a path.
' Example:
' DIM pFileSys AS CFileSys
' DIM cbsName AS CBSTR = pFileSys.GetBaseName("C:\MyFolder\Test.txt")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetBaseName (BYREF cbsPathSpec AS CBSTR) AS CBSTR
   IF m_pFileSys = NULL THEN RETURN ""
   DIM bstrName AS AFX_BSTR
   IF INSTRREV(**cbsPathSpec, ".") THEN SetResult(m_pFileSys->GetBaseName(cbsPathSpec, @bstrName))
   RETURN bstrName
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns complete and unambiguous path from a provided path specification.
' Example:
' DIM pFileSys AS CFileSys
' DIM cbsName AS CBSTR = pFileSys.GetAbsolutePathName("C:\MyFolder\Test.txt")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetAbsolutePathName (BYREF cbsPathSpec AS CBSTR) AS CBSTR
   IF m_pFileSys = NULL THEN RETURN ""
   DIM bstrName AS AFX_BSTR
   SetResult(m_pFileSys->GetAbsolutePathName(cbsPathSpec, @bstrName))
   RETURN bstrName
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns a randomly generated temporary file or folder name that is useful for performing
' operations that require a temporary file or folder.
' Example:
' DIM pFileSys AS CFileSys
' DIM cbsName AS CBSTR = pFileSys.GetTempName
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetTempName () AS CBSTR
   IF m_pFileSys = NULL THEN RETURN ""
   DIM bstrName AS AFX_BSTR
   SetResult(m_pFileSys->GetTempName(@bstrName))
   RETURN bstrName
END FUNCTION
' ========================================================================================

' ========================================================================================
' *** UNTESTED ***
' Returns the network share name for a specified drive.
' For drive letters, the root drive is not included. For example, the path for the C drive
' is C:, not C:\.
' DIM pFileSys AS CFileSys
' DIM cbsShareName AS CBSTR = pFileSys.GetDriveShareName("H:")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetDriveShareName (BYREF cbsDrive AS CBSTR) AS CBSTR
   IF m_pFileSys = NULL THEN RETURN ""
   ' // Get a reference to the IDrive interface
   DIM bstrName AS AFX_BSTR
   DIM pDrive AS Afx_IDrive PTR
   SetResult(m_pFileSys->GetDrive(cbsDrive, @pDrive))
   IF pDrive THEN
      SetResult(pDrive->get_ShareName(@bstrName))
      pDrive->Release
   END IF
   RETURN bstrName
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the type of file system in use for the specified drive.
' Available return types include FAT, NTFS, and CDFS.
' For drive letters, the root drive is not included. For example, the path for the C drive
' is C:, not C:\.
' Example:
' DIM pFileSys AS CFileSys
' DIM cbsFileSystem AS CBSTR = pFileSys.GetDriveFileSystem("C:")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetDriveFileSystem (BYREF cbsDrive AS CBSTR) AS CBSTR
   IF m_pFileSys = NULL THEN RETURN ""
   ' // Get a reference to the IDrive interface
   DIM bstrFileSystem AS AFX_BSTR
   DIM pDrive AS Afx_IDrive PTR
   SetResult(m_pFileSys->GetDrive(cbsDrive, @pDrive))
   IF pDrive THEN
      SetResult(pDrive->get_FileSystem(@bstrFileSystem))
      pDrive->Release
   END IF
   RETURN bstrFileSystem
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns a value indicating the type of a specified drive.
' Return value:
'   DriveType_UnknownType = 0
'   DriveType_Removable = 1
'   DriveType_Fixed = 2
'   DriveType_Remote = 3
'   DriveType_CDRom = 4
'   DriveType_RamDisk = 5
' Example:
' DIM pFileSys AS CFileSys
' DIM nDriveType AS CBSTR = pFileSys.GetDriveType("C:")
' DIM t AS WSTRING
' SELECT CASE pDrive.DriveType
'    CASE 0 : t = "Unknown"
'    CASE 1 : t = "Removable"
'    CASE 2 : t = "Fixed"
'    CASE 3 : t = "Network"
'    CASE 4 : t = "CD-ROM"
'    CASE 5 : t = "RAM Disk"
' END SELECT
' For drive letters, the root drive is not included. For example, the path for the C drive
' is C:, not C:\.
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetDriveType (BYREF cbsDrive AS CBSTR) AS DRIVETYPECONST
   IF m_pFileSys = NULL THEN RETURN 0
   ' // Get a reference to the IDrive interface
   DIM pDrive AS Afx_IDrive PTR
   SetResult(m_pFileSys->GetDrive(cbsDrive, @pDrive))
   DIM nDriveType AS DRIVETYPECONST
   IF pDrive THEN
      SetResult(pDrive->get_DriveType(@nDriveType))
      pDrive->Release
   END IF
   RETURN nDriveType
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns information about the type of a folder.
' Example:
' DIM pFileSys AS CFileSys
' DIM cbsFolderType AS CBSTR = pFileSys.GetFolderType("c:\MyFolder)
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetFolderType (BYREF cbsFolder AS CBSTR) AS CBSTR
   IF m_pFileSys = NULL THEN RETURN ""
   ' // Get a reference to the IFolder interface
   DIM bstrType AS AFX_BSTR
   DIM pFolder AS Afx_IFolder PTR
   SetResult(m_pFileSys->GetFolder(cbsFolder, @pFolder))
   IF pFolder THEN
      SetResult(pFolder->get_Type(@bstrType))
      pFolder->Release
   END IF
   RETURN bstrType
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns information about the type of a file. For example, for files ending in .TXT,
' "Text Document" is returned.
' Example:
' DIM pFileSys AS CFileSys
' DIM cbsFileType AS CBSTR = pFileSys.FileType("C:\MyPath\MyFile.txt")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetFileType (BYREF cbsFile AS CBSTR) AS CBSTR
   IF m_pFileSys = NULL THEN RETURN ""
   DIM bstrType AS AFX_BSTR
   ' // Get a reference to the IFile interface
   DIM pFile AS Afx_IFile PTR
   SetResult(m_pFileSys->GetFile(cbsFile, @pFile))
   IF pFile THEN
      SetResult(pFile->get_Type(@bstrType))
      pFile->Release
   END IF
   RETURN bstrType
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the version number of a specified file.
' Example:
' DIM pFileSys AS CFileSys
' DIM cbsVersion AS CBSTR = pFileSys.GetFileVersion("C:\MyFolder\MyFile.doc")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetFileVersion (BYREF cbsFile AS CBSTR) AS CBSTR
   IF m_pFileSys = NULL THEN RETURN ""
   DIM bstrVersion AS AFX_BSTR
   SetResult(m_pFileSys->GetFileVersion(cbsFile, @bstrVersion))
   RETURN bstrVersion
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the folder name for the parent of the specified folder.
' Example:
' DIM pFileSys AS CFileSys
' PRINT pFileSys.GetParentFolderName("C:\MyFolder\MySubfolder")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetParentFolderName (BYREF cbsFolder AS CBSTR) AS CBSTR
   IF m_pFileSys = NULL THEN RETURN ""
   DIM bstrName AS AFX_BSTR
   SetResult(m_pFileSys->GetParentFolderName(cbsFolder, @bstrName))
   RETURN bstrName
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Sets the name of a specified folder.
' Example:
' DIM pFileSys AS CFileSys
' pFileSys.SetFolderName("c:\MyFolder", "NewName")
' Note: You only have to pass the new name of the folder, not the full path.
' ========================================================================================
PRIVATE FUNCTION CFileSys.SetFolderName (BYREF cbsFolder AS CBSTR, BYREF cbsName AS CBSTR) AS HRESULT
   IF m_pFileSys = NULL THEN RETURN E_POINTER
   ' // Get a reference to the IFolder interface
   DIM pFolder AS Afx_IFolder PTR
   SetResult(m_pFileSys->GetFolder(cbsFolder, @pFolder))
   IF pFolder THEN
      SetResult(pFolder->put_Name(cbsName))
      pFolder->Release
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns a string containing the drive letter for a specified folder.
' Example:
' DIM pFileSys AS CFileSys
' DIM cbsDriveLetter AS CBSTR = pFileSys.GetFolderDriveLetter("c:\MyFolder)
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetFolderDriveLetter (BYREF cbsFolder AS CBSTR) AS CBSTR
   IF m_pFileSys = NULL THEN RETURN ""
   ' // Get a reference to the IFolder interface
   DIM bstrDriveLetter AS AFX_BSTR
   DIM pFolder AS Afx_IFolder PTR
   SetResult(m_pFileSys->GetFolder(cbsFolder, @pFolder))
   IF pFolder THEN
      DIM pDrive AS Afx_IDrive PTR
      SetResult(pFolder->get_Drive(@pDrive))
      IF pDrive THEN
         SetResult(pDrive->get_DriveLetter(@bstrDriveLetter))
         pDrive->Release
      END IF
      pFolder->Release
   END IF
   RETURN bstrDriveLetter
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Appends a name to an existing path.
' Example:
' DIM pFileSys AS CFileSys
' bstrNewPath = fso.BuildPath("C:\MyFolder\", "New Subfolder")
' ========================================================================================
PRIVATE FUNCTION CFileSys.BuildPath (BYREF cbsPath AS CBSTR, BYREF cbsName AS CBSTR) AS CBSTR
   IF m_pFileSys = NULL THEN RETURN ""
   DIM bstrNewPath AS AFX_BSTR
   SetResult(m_pFileSys->BuildPath(cbsPath, cbsName, @bstrNewPath))
   RETURN bstrNewPath
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the volume serial number that the operating system assigns when a hard disk is formatted.
' For drive letters, the root drive is not included. For example, the path for the C drive
' is C:, not C:\.
' Example:
' DIM pFileSys AS CFileSys
' PRINT pFileSys.GetSerialNumber("C:")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetSerialNumber (BYREF cbsDrive AS CBSTR) AS LONG
   IF m_pFileSys = NULL THEN RETURN 0
   ' // Get a reference to the IDrive interface
   DIM pDrive AS Afx_IDrive PTR
   SetResult(m_pFileSys->GetDrive(cbsDrive, @pDrive))
   DIM nSerialNumber AS LONG
   IF pDrive THEN
      SetResult(pDrive->get_SerialNumber(@nSerialNumber))
      pDrive->Release
   END IF
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
PRIVATE FUNCTION CFileSys.GetVolumeName (BYREF cbsDrive AS CBSTR) AS CBSTR
   IF m_pFileSys = NULL THEN RETURN ""
   ' // Get a reference to the IDrive interface
   DIM bstrName AS AFX_BSTR
   DIM pDrive AS Afx_IDrive PTR
   SetResult(m_pFileSys->GetDrive(cbsDrive, @pDrive))
   IF pDrive THEN
      SetResult(pDrive->get_VolumeName(@bstrName))
      pDrive->Release
   END IF
   RETURN bstrName
END FUNCTION
' ========================================================================================

' ========================================================================================
' *** UNTESTED ***
' Sets the volume name of the specified drive.
' For drive letters, the root drive is not included. For example, the path for the C drive
' is C:, not C:\.
' ========================================================================================
PRIVATE FUNCTION CFileSys.SetVolumeName (BYREF cbsDrive AS CBSTR, BYREF cbsName AS CBSTR) AS HRESULT
   IF m_pFileSys = NULL THEN RETURN E_POINTER
   ' // Get a reference to the IDrive interface
   DIM pDrive AS Afx_IDrive PTR
   SetResult(m_pFileSys->GetDrive(cbsDrive, @pDrive))
   IF pDrive THEN
      SetResult(pDrive->put_VolumeName(cbsName))
      pDrive->Release
   END IF
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
   DIM nDrives AS LONG
   IF pDrives THEN
      SetResult(pDrives->get_Count(@nDrives))
      pDrives->Release
   END IF
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
PRIVATE FUNCTION CFileSys.GetNumSubFolders (BYREF cbsFolder AS CBSTR) AS LONG
   IF m_pFileSys = NULL THEN RETURN 0
   ' // Get a reference to the IFolder interface
   DIM pFolder AS Afx_IFolder PTR
   DIM numSubFolders AS LONG
   SetResult(m_pFileSys->GetFolder(cbsFolder, @pFolder))
   IF pFolder THEN
      DIM pSubFolders AS Afx_IFolderCollection PTR
      SetResult(pFolder->get_SubFolders(@pSubFolders))
      IF pSubFolders THEN
         SetResult(pSubFolders->get_Count(@numSubFolders))
         pSubFolders->Release
      END IF
      pFolder->Release
   END IF
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
PRIVATE FUNCTION CFileSys.GetNumFiles (BYREF cbsFolder AS CBSTR) AS LONG
   IF m_pFileSys = NULL THEN RETURN 0
   ' // Get a reference to the IFolder interface
   DIM pFolder AS Afx_IFolder PTR
   DIM numFiles AS LONG
   SetResult(m_pFileSys->GetFolder(cbsFolder, @pFolder))
   IF pFolder THEN
      DIM pFiles AS Afx_IFileCollection PTR
      SetResult(pFolder->get_Files(@pFiles))
      IF pFiles THEN
         SetResult(pFiles->get_Count(@numFiles))
         pFiles->Release
      END IF
      pFolder->Release
   END IF
   RETURN numFiles
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns True(-1) if the specified folder is the root folder; False(0) if it is not.
' DIM pFileSys AS CFileSys
' PRINT pFileSys.IsRootFolder("C:\MyFolder")
' ========================================================================================
PRIVATE FUNCTION CFileSys.IsRootFolder (BYREF cbsFolder AS CBSTR) AS BOOLEAN
   IF m_pFileSys = NULL THEN RETURN FALSE
   ' // Get a reference to the IFolder interface
   DIM pFolder AS Afx_IFolder PTR
   SetResult(m_pFileSys->GetFolder(cbsFolder, @pFolder))
   IF pFolder THEN
      DIM fRootFolder AS VARIANT_BOOL
      SetResult(pFolder->get_IsRootFolder(@fRootFolder))
      pFolder->Release
      RETURN fRootFolder
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns a semicolon separated list with the driver letters.
' ========================================================================================
PRIVATE FUNCTION CFileSys.DriveLetters () AS CBSTR
   IF m_pFileSys = NULL THEN RETURN ""
   ' // Get a reference to the IDriveCollection interface
   DIM pDrives AS Afx_IDriveCollection PTR
   SetResult(m_pFileSys->get_Drives(@pDrives))
   DIM nDrives AS LONG, bstrDriveLetter AS AFX_BSTR, cbsDrives AS CBSTR
   IF pDrives THEN
      ' // Get a reference to the standard enumerator
      DIM pEnum AS IEnumVARIANT PTR, pDrive AS Afx_IDrive PTR
      DIM vItem AS VARIANT, celtFetched AS ULONG
      SetResult(pDrives->get__NewEnum(@cast(ANY PTR, pEnum)))
      ' // Enumerate the collection
      IF pEnum THEN
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
               cbsDrives += *bstrDriveLetter & ";"
               SysFreeString bstrDriveLetter
            END IF
            VariantClear @vItem
         LOOP
         pEnum->lpvtbl->Release(pEnum)
      END IF
      pDrives->Release
   END IF
   ' // Remove the last ";"
   IF LEN(cbsDrives) THEN cbsDrives = LEFT(**cbsDrives, LEN(cbsDrives) - 1)
   ' // Return the drive letters
   RETURN cbsDrives
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
' DIM cbsFileSize AS LONG = pFileSys.GetFileSize("C:\MyPath\MyFile.txt")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetFileSize (BYREF cbsFile AS CBSTR) AS LONG
   IF m_pFileSys = NULL THEN RETURN 0
   ' // Get a reference to the IFile interface
   DIM pFile AS Afx_IFile PTR
   SetResult(m_pFileSys->GetFile(cbsFile, @pFile))
   IF pFile THEN
      DIM vSize AS VARIANT
      SetResult(pFile->get_Size(@vSize))
      pFile->Release
      IF m_Result = S_OK THEN RETURN vSize.lVal
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Sets the name of a specified file.
' Example:
' DIM pFileSys AS CFileSys
' pFileSys.SetFileName("c:\MyFolder\Test.txt", "NewName")
' ========================================================================================
PRIVATE FUNCTION CFileSys.SetFileName (BYREF cbsFile AS CBSTR, BYREF cbsName AS CBSTR) AS HRESULT
   IF m_pFileSys = NULL THEN RETURN E_POINTER
   ' // Get a reference to the IFile interface
   DIM pFile AS Afx_IFile PTR
   SetResult(m_pFileSys->GetFile(cbsFile, @pFile))
   IF pFile THEN
      SetResult(pFile->put_Name(cbsName))
      pFile->Release
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the short path used by programs that require the earlier 8.3 file naming convention.
' Example:
' DIM pFileSys AS CFileSys
' DIM cbsFileName CBSTR = pFileSys.GetFileShortPath("C:\MyPath\MyFile.txt")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetFileShortPath (BYREF cbsFile AS CBSTR) AS CBSTR
   IF m_pFileSys = NULL THEN RETURN ""
   ' // Get a reference to the IFile interface
   DIM bstrPath AS AFX_BSTR
   DIM pFile AS Afx_IFile PTR
   SetResult(m_pFileSys->GetFile(cbsFile, @pFile))
   IF pFile THEN
      SetResult(pFile->get_ShortPath(@bstrPath))
      pFile->Release
   END IF
   RETURN bstrPath
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the short name used by programs that require the earlier 8.3 file naming convention.
' Example:
' DIM pFileSys AS CFileSys
' DIM cbsFileName CBSTR = pFileSys.GetFileShortName("C:\MyPath\MyFile.txt")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetFileShortName (BYREF cbsFile AS CBSTR) AS CBSTR
   IF m_pFileSys = NULL THEN RETURN ""
   ' // Get a reference to the IFile interface
   DIM bstrName AS AFX_BSTR
   DIM pFile AS Afx_IFile PTR
   SetResult(m_pFileSys->GetFile(cbsFile, @pFile))
   IF pFile THEN
      SetResult(pFile->get_ShortName(@bstrName))
      pFile->Release
   END IF
   RETURN bstrName
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the attributes of the specified file.
' Example:
' DIM pFileSys AS CFileSys
' DIM lAttr FILEATTRIBUTE = pFileSys.GetFileAttributes("C:\MyPath\MyFile.txt")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetFileAttributes (BYREF cbsFile AS CBSTR) AS FILEATTRIBUTE
   DIM lAttr AS FILEATTRIBUTE
   IF m_pFileSys = NULL THEN RETURN lAttr
   ' // Get a reference to the IFile interface
   DIM pFile AS Afx_IFile PTR
   SetResult(m_pFileSys->GetFile(cbsFile, @pFile))
   IF pFile THEN
      SetResult(pFile->get_Attributes(@lAttr))
      pFile->Release
   END IF
   RETURN lAttr
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Sets the attributes of the specified file.
' Example:
' DIM pFileSys AS CFileSys
' pFileSys.SetFileAttributes("C:\MyPath\MyFile.txt", 33)
' ========================================================================================
PRIVATE FUNCTION CFileSys.SetFileAttributes (BYREF cbsFile AS CBSTR, BYVAL lAttr AS FILEATTRIBUTE) AS HRESULT
   IF m_pFileSys = NULL THEN RETURN E_POINTER
   ' // Get a reference to the IFile interface
   DIM pFile AS Afx_IFile PTR
   SetResult(m_pFileSys->GetFile(cbsFile, @pFile))
   IF pFile THEN
      SetResult(pFile->put_Attributes(lAttr))
      pFile->Release
   END IF
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
PRIVATE FUNCTION CFileSys.GetFileDateCreated (BYREF cbsFile AS CBSTR) AS DATE_
   DIM nDate AS DATE_
   IF m_pFileSys = NULL THEN RETURN nDate
   ' // Get a reference to the IFile interface
   DIM pFile AS Afx_IFile PTR
   SetResult(m_pFileSys->GetFile(cbsFile, @pFile))
   IF pFile THEN
      SetResult(pFile->get_DateCreated(@nDate))
      pFile->Release
   END IF
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
PRIVATE FUNCTION CFileSys.GetFileDateLastModified (BYREF cbsFile AS CBSTR) AS DATE_
   DIM nDate AS DATE_
   IF m_pFileSys = NULL THEN RETURN nDate
   ' // Get a reference to the IFile interface
   DIM pFile AS Afx_IFile PTR
   SetResult(m_pFileSys->GetFile(cbsFile, @pFile))
   IF pFile THEN
      SetResult(pFile->get_DateLastModified(@nDate))
      pFile->Release
   END IF
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
PRIVATE FUNCTION CFileSys.GetFileDateLastAccessed (BYREF cbsFile AS CBSTR) AS DATE_
   DIM nDate AS DATE_
   IF m_pFileSys = NULL THEN RETURN nDate
   ' // Get a reference to the IFile interface
   DIM pFile AS Afx_IFile PTR
   SetResult(m_pFileSys->GetFile(cbsFile, @pFile))
   IF pFile THEN
      SetResult(pFile->get_DateLastAccessed(@nDate))
      pFile->Release
   END IF
   RETURN nDate
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the size, in bytes, of the specified file.
' Example:
' DIM pFileSys AS CFileSys
' DIM cbsFolderSize AS LONG = pFileSys.GetFolderSize("C:\MyPath")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetFolderSize (BYREF cbsFolder AS CBSTR) AS LONG
   IF m_pFileSys = NULL THEN RETURN 0
   ' // Get a reference to the IFolder interface
   DIM pFolder AS Afx_IFolder PTR
   SetResult(m_pFileSys->GetFolder(cbsFolder, @pFolder))
   IF pFolder THEN
      DIM vSize AS VARIANT
      SetResult(pFolder->get_Size(@vSize))
      pFolder->Release
      IF m_Result = S_OK THEN RETURN vSize.lVal
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the attributes of the specified folder.
' Example:
' DIM pFileSys AS CFileSys
' DIM lAttr FILEATTRIBUTE = pFileSys.GetFolderAttributes("C:\MyPath")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetFolderAttributes (BYREF cbsFolder AS CBSTR) AS FILEATTRIBUTE
   DIM lAttr AS FILEATTRIBUTE
   IF m_pFileSys = NULL THEN RETURN lAttr
   ' // Get a reference to the IFolder interface
   DIM pFolder AS Afx_IFolder PTR
   SetResult(m_pFileSys->GetFolder(cbsFolder, @pFolder))
   IF pFolder THEN
      SetResult(pFolder->get_Attributes(@lAttr))
      pFolder->Release
   END IF
   RETURN lAttr
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Sets the attributes of the specified folder.
' Example:
' DIM pFileSys AS CFileSys
' DIM lAttr FILEATTRIBUTE = pFileSys.SetFolderAttributes("C:\MyPath")
' ========================================================================================
PRIVATE FUNCTION CFileSys.SetFolderAttributes (BYREF cbsFolder AS CBSTR, BYVAL lAttr AS FILEATTRIBUTE) AS HRESULT
   IF m_pFileSys = NULL THEN RETURN E_POINTER
   ' // Get a reference to the IFolder interface
   DIM pFolder AS Afx_IFolder PTR
   SetResult(m_pFileSys->GetFolder(cbsFolder, @pFolder))
   IF pFolder THEN
      SetResult(pFolder->put_Attributes(lAttr))
      pFolder->Release
   END IF
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
PRIVATE FUNCTION CFileSys.GetFolderDateCreated (BYREF cbsFolder AS CBSTR) AS DATE_
   DIM nDate AS DATE_
   IF m_pFileSys = NULL THEN RETURN nDate
   ' // Get a reference to the IFolder interface
   DIM pFolder AS Afx_IFolder PTR
   SetResult(m_pFileSys->GetFolder(cbsFolder, @pFolder))
   IF pFolder THEN
      SetResult(pFolder->get_DateCreated(@nDate))
      pFolder->Release
   END IF
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
PRIVATE FUNCTION CFileSys.GetFolderDateLastModified (BYREF cbsFolder AS CBSTR) AS DATE_
   DIM nDate AS DATE_
   IF m_pFileSys = NULL THEN RETURN nDate
   ' // Get a reference to the IFolder interface
   DIM pFolder AS Afx_IFolder PTR
   SetResult(m_pFileSys->GetFolder(cbsFolder, @pFolder))
   IF pFolder THEN
      SetResult(pFolder->get_DateLastModified(@nDate))
      pFolder->Release
   END IF
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
PRIVATE FUNCTION CFileSys.GetFolderDateLastAccessed (BYREF cbsFolder AS CBSTR) AS DATE_
   DIM nDate AS DATE_
   IF m_pFileSys = NULL THEN RETURN nDate
   ' // Get a reference to the IFolder interface
   DIM pFolder AS Afx_IFolder PTR
   SetResult(m_pFileSys->GetFolder(cbsFolder, @pFolder))
   IF pFolder THEN
      SetResult(pFolder->get_DateLastAccessed(@nDate))
      pFolder->Release
   END IF
   RETURN nDate
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the amount of space available to a user on the specified drive or network share.
' Example:
' DIM pFileSys AS CFileSys
' PRINT pFileSys.GetDriveAvailableSpace("C:")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetDriveAvailableSpace (BYREF cbsDrive AS CBSTR) AS DOUBLE
   IF m_pFileSys = NULL THEN RETURN 0
   ' // Get a reference to the IDrive interface
   DIM pDrive AS Afx_IDrive PTR
   SetResult(m_pFileSys->GetDrive(cbsDrive, @pDrive))
   IF pDrive THEN
      DIM vSize AS VARIANT
      SetResult(pDrive->get_AvailableSpace(@vSize))
      pDrive->Release
      RETURN vSize.dblVal
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the amount of free space available to a user on the specified drive or network share.
' Example:
' DIM pFileSys AS CFileSys
' PRINT pFileSys.GetDriveFreeSpace("C:")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetDriveFreeSpace (BYREF cbsDrive AS CBSTR) AS DOUBLE
   IF m_pFileSys = NULL THEN RETURN 0
   ' // Get a reference to the IDrive interface
   DIM pDrive AS Afx_IDrive PTR
   SetResult(m_pFileSys->GetDrive(cbsDrive, @pDrive))
   IF pDrive THEN
      DIM vSize AS VARIANT
      SetResult(pDrive->get_FreeSpace(@vSize))
      pDrive->Release
      RETURN vSize.dblVal
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the total space, in bytes, of a drive or network share.
' Example:
' DIM pFileSys AS CFileSys
' PRINT pFileSys.GetDriveTotalSize("C:")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetDriveTotalSize (BYREF cbsDrive AS CBSTR) AS DOUBLE
   IF m_pFileSys = NULL THEN RETURN 0
   ' // Get a reference to the IDrive interface
   DIM pDrive AS Afx_IDrive PTR
   SetResult(m_pFileSys->GetDrive(cbsDrive, @pDrive))
   IF pDrive THEN
      DIM vSize AS VARIANT
      SetResult(pDrive->get_TotalSize(@vSize))
      pDrive->Release
      RETURN vSize.dblVal
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Creates a folder. An error occurs if the specified folder already exists.
' Example:
' DIM pFileSys AS CFileSys
' DIM pFolder AS Afx_Folder PTR
' pFolder = pFileSys.CreateFolder("C:\MyNewFolder")
' IF pFolder THEN
'    ....
'    pFolder-Release
' END IF
' ========================================================================================
PRIVATE FUNCTION CFileSys.CreateFolder (BYREF cbsFolder AS CBSTR) AS Afx_IFolder PTR
   IF m_pFileSys = NULL THEN RETURN NULL
   DIM pFolder AS Afx_IFolder PTR
   SetResult(m_pFileSys->CreateFolder(cbsFolder, @pFolder))
   RETURN pFolder
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Deletes a specified folder and its contents.
' Example:
' DIM pFileSys AS CFileSys
' pFileSys.DeleteFolder("C:\MyFolder")
' ========================================================================================
PRIVATE FUNCTION CFileSys.DeleteFolder (BYREF cbsFolder AS CBSTR, BYVAL bForce AS VARIANT_BOOL = FALSE) AS HRESULT
   IF m_pFileSys = NULL THEN RETURN E_POINTER
   RETURN SetResult(m_pFileSys->DeleteFolder(cbsFolder, bForce))
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Moves one or more folders from one location to another.
' Parameters:
' - cbsSource: The path to the folder or folders to be moved. The cbsSource argument string
'   can contain wildcard characters in the last path component only.
' - cbsDestination: The path where the folder or folders are to be moved. The cbsDestination
'   argument can't contain wildcard characters.
' Example:
' DIM pFileSys AS CFileSys
' pFileSys.MoveFolder("C:\MyFolder", "C:\MyNewFolder\")
' ========================================================================================
PRIVATE FUNCTION CFileSys.MoveFolder (BYREF cbsSource AS CBSTR, BYREF cbsDestination AS CBSTR) AS HRESULT
   IF m_pFileSys = NULL THEN RETURN E_POINTER
   RETURN SetResult(m_pFileSys->MoveFolder(cbsSource, cbsDestination))
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Recursively copies a folder from one location to another.
' Parameters:
' - cbsSource: Character string file specification, which can include wildcard characters,
'   for one or more folders to be copied.
' - cbsDestination: Character string destination where the folder and subfolder from source
'   are to be copied. Wildcard characters are not allowed.
' - Boolean value that indicates if existing folders are to be overwritten. If true, files
'   are overwritten; if false, they are not. The default is true.
' Example:
' DIM pFileSys AS CFileSys
' pFileSys.CopyFolder("C:\MyFolder", "C:\MyOtherFolder\")
' ========================================================================================
PRIVATE FUNCTION CFileSys.CopyFolder (BYREF cbsSource AS CBSTR, BYREF cbsDestination AS CBSTR, BYVAL OverWriteFiles AS VARIANT_BOOL = -1) AS HRESULT
   IF m_pFileSys = NULL THEN RETURN E_POINTER
   RETURN SetResult(m_pFileSys->CopyFolder(cbsSource, cbsDestination, OverWriteFiles))
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Copies one or more files from one location to another.
' Parameters:
' - cbsSource: Character string file specification, which can include wildcard characters,
'   for one or more files to be copied.
' - cbsDestination: Character string destination where the file or files from source
'   are to be copied. Wildcard characters are not allowed.
' - OverWriteFiles: Boolean value that indicates if existing files are to be overwritten.
'   If true, files are overwritten; if false, they are not. The default is true. Note that
'   CopyFile will fail if destination has the read-only attribute set, regardless of the
'   value of overwrite.
' Example:
' DIM pFileSys AS CFileSys
' pFileSys.CopyFile("C:\MyFolder\MyFile.txt", "C:\MyOtherFolder\MyFile.txt")
' ========================================================================================
PRIVATE FUNCTION CFileSys.CopyFile (BYREF cbsSource AS CBSTR, BYREF cbsDestination AS CBSTR, BYVAL OverWriteFiles AS VARIANT_BOOL = -1) AS HRESULT
   IF m_pFileSys = NULL THEN RETURN E_POINTER
   RETURN SetResult(m_pFileSys->CopyFile(cbsSource, cbsDestination, OverWriteFiles))
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Deletes a specified file.
' Parameters:
' - cbsFileSpec: The name of the file to delete. cbsFileSpec can contain wildcard characters
'   in the last path component.
' - bForce: Boolean value that is true if files with the read-only attribute set are to be
'   deleted; false (default) if they are not.
' Example:
' DIM pFileSys AS CFileSys
' pFileSys.DeleteFile("C:\MyFolder\MyFile.txt")
' ========================================================================================
PRIVATE FUNCTION CFileSys.DeleteFile (BYREF cbsFileSpec AS CBSTR, BYVAL bForce AS VARIANT_BOOL = FALSE) AS HRESULT
   IF m_pFileSys = NULL THEN RETURN E_POINTER
   RETURN SetResult(m_pFileSys->DeleteFile(cbsFileSpec, bForce))
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Moves one or more files from one location to another.
' Parameters:
' - cbsSource: The path to the file or files to be moved. The cbsSource argument string
'   can contain wildcard characters in the last path component only.
' - cbsDestination: The path where the file or files are to be moved. The cbsDestination
'   argument can't contain wildcard characters.
' Example:
' DIM pFileSys AS CFileSys
' pFileSys.MoveFile("C:\MyFolder\MyFile.txt", "C:\MyOtherFolder\")
' ========================================================================================
PRIVATE FUNCTION CFileSys.MoveFile (BYREF cbsSource AS CBSTR, BYREF cbsDestination AS CBSTR) AS HRESULT
   IF m_pFileSys = NULL THEN RETURN E_POINTER
   RETURN SetResult(m_pFileSys->MoveFile(cbsSource, cbsDestination))
END FUNCTION
' ========================================================================================

END NAMESPACE
