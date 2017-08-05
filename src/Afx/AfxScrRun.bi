' ########################################################################################
' Microsoft Windows
' File: AfxScrRun.bi
' Library name: Scripting
' Version: 1.0
' Documentation string: Microsoft Scripting Runtime
' Path: C:\WINDOWS\system32\scrrun.dll
' Library GUID: {420B2830-E718-11CF-893D-00A0C9054228}
' Help file: C:\WINDOWS\system32\VBENLR98.CHM
' Portions Copyright (c) Microsoft Corporation.
' Compiler: Free Basic 32 & 64 bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "windows.bi"
#include once "win/ole2.bi"
#include once "win/unknwnbase.bi"

NAMESPACE Afx

' // The definition for BSTR in the FreeBASIC headers was inconveniently changed to WCHAR
#ifndef AFX_BSTR
   #define AFX_BSTR WSTRING PTR
#endif

' ========================================================================================
' ProgIDs (Program identifiers)
' ========================================================================================

' CLSID = {EE09B103-97E0-11CF-978F-00A02463E06F}
CONST AFX_PROGID_ScriptingDictionary = "Scripting.Dictionary"
' CLSID = {32DA2B15-CFED-11D1-B747-00C04FC2B085}
CONST AFX_PROGID_ScriptingEncoder = "Scripting.Encoder"
' CLSID = {0D43FE01-F093-11CF-8940-00A0C9054228}
CONST AFX_PROGID_ScriptingFileSystemObject = "Scripting.FileSystemObject"

' ========================================================================================
' ClsIDs (Class identifiers)
' ========================================================================================

CONST AFX_CLSID_Drive = "{C7C3F5B1-88A3-11D0-ABCB-00A0C90FFFC0}"
CONST AFX_CLSID_Drives = "{C7C3F5B2-88A3-11D0-ABCB-00A0C90FFFC0}"
CONST AFX_CLSID_File = "{C7C3F5B5-88A3-11D0-ABCB-00A0C90FFFC0}"
CONST AFX_CLSID_Files = "{C7C3F5B6-88A3-11D0-ABCB-00A0C90FFFC0}"
CONST AFX_CLSID_Folder = "{C7C3F5B3-88A3-11D0-ABCB-00A0C90FFFC0}"
CONST AFX_CLSID_Folders = "{C7C3F5B4-88A3-11D0-ABCB-00A0C90FFFC0}"
CONST AFX_CLSID_TextStream = "{0BB02EC0-EF49-11CF-8940-00A0C9054228}"
CONST AFX_CLSID_Dictionary = "{EE09B103-97E0-11CF-978F-00A02463E06F}"
CONST AFX_CLSID_Encoder = "{32DA2B15-CFED-11D1-B747-00C04FC2B085}"
CONST AFX_CLSID_FileSystemObject = "{0D43FE01-F093-11CF-8940-00A0C9054228}"

' ========================================================================================
' IIDs (Interface identifiers)
' ========================================================================================

CONST AFX_IID_IDictionary = "{42C642C1-97E1-11CF-978F-00A02463E06F}"
CONST AFX_IID_IDrive = "{C7C3F5A0-88A3-11D0-ABCB-00A0C90FFFC0}"
CONST AFX_IID_IDriveCollection = "{C7C3F5A1-88A3-11D0-ABCB-00A0C90FFFC0}"
CONST AFX_IID_IFile = "{C7C3F5A4-88A3-11D0-ABCB-00A0C90FFFC0}"
CONST AFX_IID_IFileCollection = "{C7C3F5A5-88A3-11D0-ABCB-00A0C90FFFC0}"
CONST AFX_IID_IFileSystem = "{0AB5A3D0-E5B6-11D0-ABF5-00A0C90FFFC0}"
CONST AFX_IID_IFileSystem3 = "{2A0B9D10-4B87-11D3-A97A-00104B365C9F}"
CONST AFX_IID_IFolder = "{C7C3F5A2-88A3-11D0-ABCB-00A0C90FFFC0}"
CONST AFX_IID_IFolderCollection = "{C7C3F5A3-88A3-11D0-ABCB-00A0C90FFFC0}"
CONST AFX_IID_IScriptEncoder = "{AADC65F6-CFF1-11D1-B747-00C04FC2B085}"
CONST AFX_IID_ITextStream = "{53BAD8C1-E718-11CF-893D-00A0C9054228}"

' ========================================================================================
' CompareMethod enum
' ========================================================================================
TYPE _CompareMethod AS LONG
ENUM
   CompareMethod_BinaryCompare = 0
   CompareMethod_TextCompare = 1
   CompareMethod_DatabaseCompare = 2
END ENUM
TYPE COMPAREMETHOD AS _CompareMethod

' ========================================================================================
' IOMode enum
' ========================================================================================
TYPE _IOMode AS LONG
ENUM
   IOMode_ForReading = 1
   IOMode_ForWriting = 2
   IOMode_ForAppending = 8
END ENUM
TYPE IOMODE AS _IOMode

' ========================================================================================
' Tristate enum
' ========================================================================================
TYPE _TriState AS LONG
ENUM EXPLICIT
   TriState_True = -1
   TriState_False = 0
   TriState_UseDefault = -2
   TriState_Mixed = -2
END ENUM
TYPE TRISTATE AS _TriState

' ========================================================================================
' FileAttribute enum
' ========================================================================================
TYPE _FileAttribute AS LONG
ENUM
   FileAttribute_Normal = 0
   FileAttribute_ReadOnly = 1
   FileAttribute_Hidden = 2
   FileAttribute_System = 4
   FileAttribute_Volume = 8
   FileAttribute_Directory = 16
   FileAttribute_Archive = 32
   FileAttribute_Alias = 1024
   FileAttribute_Compressed = 2048
END ENUM
TYPE FILEATTRIBUTE AS _FileAttribute

' ========================================================================================
' SpecialFolderConst enum
' ========================================================================================
TYPE _SpecialFolderConst AS LONG
ENUM
   SpecialFolder_WindowsFolder = 0
   SpecialFolder_SystemFolder = 1
   SpecialFolder_TemporaryFolder = 2
END ENUM
TYPE SPECIALFOLDERCONST AS _SpecialFolderConst

' ========================================================================================
' DriveTypeConst enum
' ========================================================================================
TYPE _DriveTypeConst AS LONG
ENUM
   DriveType_UnknownType = 0
   DriveType_Removable = 1
   DriveType_Fixed = 2
   DriveType_Remote = 3
   DriveType_CDRom = 4
   DriveType_RamDisk = 5
END ENUM
TYPE DRIVETYPECONST AS _DriveTypeConst

' ========================================================================================
' StandardStreamTypes enum
' ========================================================================================
TYPE _StandardStreamTypes AS LONG
ENUM
   StandardStreamTypes_StdIn = 0
   StandardStreamTypes_StdOut = 1
   StandardStreamTypes_StdErr = 2
END ENUM
TYPE STANDARDSTREAMTYPES AS _StandardStreamTypes

#ifndef __Afx_IUnknown_INTERFACE_DEFINED__
#define __Afx_IUnknown_INTERFACE_DEFINED__
TYPE Afx_IUnknown AS Afx_IUnknown_
TYPE Afx_IUnknown_ EXTENDS OBJECT
	DECLARE ABSTRACT FUNCTION QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObject AS LPVOID PTR) AS HRESULT
	DECLARE ABSTRACT FUNCTION AddRef () AS ULONG
	DECLARE ABSTRACT FUNCTION Release () AS ULONG
END TYPE
TYPE AFX_LPUNKNOWN AS Afx_IUnknown PTR
#endif

#ifndef __Afx_IDispatch_INTERFACE_DEFINED__
#define __Afx_IDispatch_INTERFACE_DEFINED__
TYPE Afx_IDispatch AS Afx_IDispatch_
TYPE Afx_IDispatch_  EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION GetTypeInfoCount (BYVAL pctinfo AS UINT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetTypeInfo (BYVAL iTInfo AS UINT, BYVAL lcid AS LCID, BYVAL ppTInfo AS ITypeInfo PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetIDsOfNames (BYVAL riid AS CONST IID CONST PTR, BYVAL rgszNames AS LPOLESTR PTR, BYVAL cNames AS UINT, BYVAL lcid AS LCID, BYVAL rgDispId AS DISPID PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL dispIdMember AS DISPID, BYVAL riid AS CONST IID CONST PTR, BYVAL lcid AS LCID, BYVAL wFlags AS WORD, BYVAL pDispParams AS DISPPARAMS PTR, BYVAL pVarResult AS VARIANT PTR, BYVAL pExcepInfo AS EXCEPINFO PTR, BYVAL puArgErr AS UINT PTR) AS HRESULT
END TYPE
TYPE AFX_LPDISPATCH AS Afx_IDispatch PTR
#endif

TYPE Afx_IDictionary AS Afx_IDictionary_
TYPE Afx_IFileSystem AS Afx_IFileSystem_
TYPE Afx_IDriveCollection AS Afx_IDriveCollection_
TYPE Afx_IDrive AS Afx_IDrive_
TYPE Afx_IFolder AS Afx_IFolder_
TYPE Afx_IFolderCollection AS Afx_IFolderCollection_
TYPE Afx_IFileCollection AS Afx_IFileCollection_
TYPE Afx_IFile AS Afx_IFile_
TYPE Afx_ITextStream AS Afx_ITextStream_
TYPE Afx_IFileSystem3 AS Afx_IFileSystem3_
TYPE Afx_IScriptEncoder AS Afx_IScriptEncoder_

' ########################################################################################
' Interface name = IDictionary
' IID = {42C642C1-97E1-11CF-978F-00A02463E06F}
' Attributes = 4176 [&H1050] [Hidden] [Dual] [Dispatchable]
' Inherited interface = IDispatch
' ########################################################################################

#ifndef __Afx_IDictionary_INTERFACE_DEFINED__
#define __Afx_IDictionary_INTERFACE_DEFINED__

TYPE Afx_IDictionary_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION putref_Item (BYVAL vKey AS VARIANT PTR, BYVAL vNewItem AS VARIANT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Item (BYVAL vKey AS VARIANT PTR, BYVAL vNewItem AS VARIANT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Item (BYVAL vKey AS VARIANT PTR, BYVAL vRetItem AS VARIANT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Add (BYVAL vKey AS VARIANT PTR, BYVAL vItem AS VARIANT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Count (BYVAL pCount AS long PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Exists (BYVAL vKey AS VARIANT PTR, BYVAL pExists AS SHORT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Items (BYVAL pItemsArray AS VARIANT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Key (BYVAL vKey AS VARIANT PTR, BYVAL vItem AS VARIANT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Keys (BYVAL pKeysArray AS VARIANT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Remove (BYVAL vKey AS VARIANT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION RemoveAll () AS HRESULT
   DECLARE ABSTRACT FUNCTION put_CompareMode (BYVAL pcomp AS COMPAREMETHOD) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_CompareMode (BYVAL pcomp AS COMPAREMETHOD PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get__NewEnum (BYVAL ppunk AS Afx_IUnknown PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_HashVal (BYVAL vKey AS VARIANT PTR, BYVAL HashVal AS VARIANT PTR) AS HRESULT
END TYPE

TYPE AFX_LPDICTIONARY AS Afx_IDictionary PTR

#endif

' ########################################################################################
' Interface name = IFileSystem
' IID = {0AB5A3D0-E5B6-11D0-ABF5-00A0C90FFFC0}
' Attributes = 4304 [&H10D0] [Hidden] [Dual] [Nonextensible] [Dispatchable]
' Inherited interface = IDispatch
' ########################################################################################

#ifndef __Afx_IFileSystem_INTERFACE_DEFINED__
#define __Afx_IFileSystem_INTERFACE_DEFINED__

TYPE Afx_IFileSystem_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION get_Drives (BYVAL ppdrives AS Afx_IDriveCollection PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION BuildPath (BYVAL Path AS AFX_BSTR, BYVAL Name AS AFX_BSTR, BYVAL pbstrResult AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetDriveName (BYVAL Path AS AFX_BSTR, BYVAL pbstrResult AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetParentFolderName (BYVAL Path AS AFX_BSTR, BYVAL pbstrResult AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetFileName (BYVAL Path AS AFX_BSTR, BYVAL pbstrResult AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetBaseName (BYVAL Path AS AFX_BSTR, BYVAL pbstrResult AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetExtensionName (BYVAL Path AS AFX_BSTR, BYVAL pbstrResult AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetAbsolutePathName (BYVAL Path AS AFX_BSTR, BYVAL pbstrResult AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetTempName (BYVAL pbstrResult AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION DriveExists (BYVAL DriveSpec AS AFX_BSTR, BYVAL pfExists AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION FileExists(BYVAL FileSpec AS AFX_BSTR, BYVAL pfExists AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION FolderExists (BYVAL FolderSpec AS AFX_BSTR, BYVAL pfExists AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetDrive (BYVAL DriveSpec AS AFX_BSTR, BYVAL ppdrive AS Afx_IDrive PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetFile (BYVAL FilePath AS AFX_BSTR, BYVAL ppfile AS Afx_IFile PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetFolder (BYVAL FolderPath AS AFX_BSTR, BYVAL ppfolder AS Afx_IFolder PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetSpecialFolder (BYVAL SpecialFolder AS SPECIALFOLDERCONST, BYVAL ppfolder AS Afx_IFolder PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION DeleteFile (BYVAL FileSpec AS AFX_BSTR, BYVAL bForce AS VARIANT_BOOL = FALSE) AS HRESULT
   DECLARE ABSTRACT FUNCTION DeleteFolder (BYVAL FolderSpec AS AFX_BSTR, BYVAL bForce AS VARIANT_BOOL = FALSE) AS HRESULT
   DECLARE ABSTRACT FUNCTION MoveFile (BYVAL Source AS AFX_BSTR, BYVAL Destination AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION MoveFolder (BYVAL Source AS AFX_BSTR, BYVAL Destination AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION CopyFile (BYVAL Source AS AFX_BSTR, BYVAL Destination AS AFX_BSTR, BYVAL OverWriteFiles AS VARIANT_BOOL = -1) AS HRESULT
   DECLARE ABSTRACT FUNCTION CopyFolder (BYVAL Source AS AFX_BSTR, BYVAL Destination AS AFX_BSTR, BYVAL OverWriteFiles AS VARIANT_BOOL = -1) AS HRESULT
   DECLARE ABSTRACT FUNCTION CreateFolder (BYVAL Path AS AFX_BSTR, BYVAL ppfolder AS Afx_IFolder PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION CreateTextFile (BYVAL FileName AS AFX_BSTR, BYVAL Overwrite AS VARIANT_BOOL = -1, BYVAL bUnicode AS VARIANT_BOOL = FALSE, BYVAL ppts AS Afx_ITextStream PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION OpenTextFile (BYVAL FileName AS AFX_BSTR, BYVAL IOMode AS LONG = 1, BYVAL Create AS VARIANT_BOOL = FALSE, BYVAL Format AS LONG = 0, BYVAL ppts AS Afx_ITextStream PTR PTR) AS HRESULT
END TYPE

TYPE AFX_LPFILESYSTEM AS Afx_IFileSystem PTR

#endif

' ########################################################################################
' Interface name = IDriveCollection
' IID = {C7C3F5A1-88A3-11D0-ABCB-00A0C90FFFC0}
' Attributes = 4304 [&H10D0] [Hidden] [Dual] [Nonextensible] [Dispatchable]
' Inherited interface = IDispatch
' ########################################################################################

#ifndef __Afx_IDriveCollection_INTERFACE_DEFINED__
#define __Afx_IDriveCollection_INTERFACE_DEFINED__

TYPE Afx_IDriveCollection_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION get_Item (BYVAL vKey AS VARIANT, BYVAL ppdrive AS Afx_IDrive PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get__NewEnum (BYVAL ppenum AS Afx_IUnknown PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Count (BYVAL plCount AS LONG PTR) AS HRESULT
END TYPE

TYPE AFX_LPDRIVECOLLECTION AS Afx_IDriveCollection PTR

#endif

' ########################################################################################
' Interface name = IDrive
' IID = {C7C3F5A0-88A3-11D0-ABCB-00A0C90FFFC0}
' Attributes = 4304 [&H10D0] [Hidden] [Dual] [Nonextensible] [Dispatchable]
' Inherited interface = IDispatch
' ########################################################################################

#ifndef __Afx_IDrive_INTERFACE_DEFINED__
#define __Afx_IDrive_INTERFACE_DEFINED__

TYPE Afx_IDrive_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION get_Path (BYVAL pbstrPath AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_DriveLetter (BYVAL pbstrLetter AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ShareName (BYVAL pbstrShareName AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_DriveType (BYVAL pdt AS DRIVETYPECONST PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_RootFolder (BYVAL ppfolder AS Afx_IFolder PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_AvailableSpace (BYVAL pvarAvail AS VARIANT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_FreeSpace (BYVAL pvarFree AS VARIANT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_TotalSize (BYVAL pvarTotal AS VARIANT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_VolumeName (BYVAL pbstrName AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_VolumeName (BYVAL pbstrName AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_FileSystem (BYVAL pbstrFileSystem AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_SerialNumber (BYVAL pulSerialNumber AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_IsReady (BYVAL pfReady AS VARIANT_BOOL PTR) AS HRESULT
END TYPE

TYPE AFX_LPDRIVE AS Afx_IDrive PTR

#endif

' ########################################################################################
' Interface name = IFolder
' IID = {C7C3F5A2-88A3-11D0-ABCB-00A0C90FFFC0}
' Attributes = 4304 [&H10D0] [Hidden] [Dual] [Nonextensible] [Dispatchable]
' Inherited interface = IDispatch
' ########################################################################################

#ifndef __Afx_IFolder_INTERFACE_DEFINED__
#define __Afx_IFolder_INTERFACE_DEFINED__

TYPE Afx_IFolder_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION get_Path (BYVAL pbstrPath AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Name (BYVAL pbstrName AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Name (BYVAL pbstrName AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ShortPath (BYVAL pbstrPath AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ShortName (BYVAL pbstrName AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Drive (BYVAL ppdrive AS Afx_IDrive PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ParentFolder (BYVAL ppfolder AS Afx_IFolder PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Attributes (BYVAL pfa AS FILEATTRIBUTE PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Attributes (BYVAL pfa AS FILEATTRIBUTE) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_DateCreated (BYVAL pdate AS DATE_ PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_DateLastModified (BYVAL pdate AS DATE_ PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_DateLastAccessed (BYVAL pdate AS DATE_ PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Type (BYVAL pbstrType AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Delete_ (BYVAL bForce AS VARIANT_BOOL = FALSE) AS HRESULT
   DECLARE ABSTRACT FUNCTION Copy (BYVAL Destination AS AFX_BSTR, BYVAL OverWriteFiles AS VARIANT_BOOL = -1) AS HRESULT
   DECLARE ABSTRACT FUNCTION Move (BYVAL Destination AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_IsRootFolder (BYVAL pfRootFolder AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Size (BYVAL pvarSize AS VARIANT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_SubFolders (BYVAL ppfolders AS Afx_IFolderCollection PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Files (BYVAL ppfiles AS Afx_IFileCollection PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION CreateTextFile (BYVAL FileName AS AFX_BSTR, BYVAL Overwrite AS VARIANT_BOOL = -1, BYVAL bUnicode AS VARIANT_BOOL = FALSE, BYVAL ppts AS Afx_ITextStream PTR PTR) AS HRESULT
END TYPE

TYPE AFX_LPFOLDER AS Afx_IFolder PTR

#endif

' ########################################################################################
' Interface name = IFolderCollection
' IID = {C7C3F5A3-88A3-11D0-ABCB-00A0C90FFFC0}
' Attributes = 4304 [&H10D0] [Hidden] [Dual] [Nonextensible] [Dispatchable]
' Inherited interface = IDispatch
' ########################################################################################

#ifndef __Afx_IFolderCollection_INTERFACE_DEFINED__
#define __Afx_IFolderCollection_INTERFACE_DEFINED__

TYPE Afx_IFolderCollection_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION Add (BYVAL Name AS AFX_BSTR, BYVAL ppfolder AS Afx_IFolder PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Item (BYVAL vKey AS VARIANT, BYVAL ppfolder AS Afx_IFolder PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get__NewEnum (BYVAL ppenum AS Afx_IUnknown PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Count (BYVAL plCount AS LONG PTR) AS HRESULT
END TYPE

TYPE AFX_LPFOLDERCOLLECTION AS Afx_IFolderCollection_ PTR

#endif

' ########################################################################################
' Interface name = IFileCollection
' IID = {C7C3F5A5-88A3-11D0-ABCB-00A0C90FFFC0}
' Attributes = 4304 [&H10D0] [Hidden] [Dual] [Nonextensible] [Dispatchable]
' Inherited interface = IDispatch
' ########################################################################################

#ifndef __Afx_IFileCollection_INTERFACE_DEFINED__
#define __Afx_IFileCollection_INTERFACE_DEFINED__

TYPE Afx_IFileCollection_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION get_Item (BYVAL vKey AS VARIANT, BYVAL ppfile AS Afx_IFile PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get__NewEnum (BYVAL ppenum AS Afx_IUnknown PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Count (BYVAL plCount AS LONG PTR) AS HRESULT
END TYPE

TYPE AFX_LPFILECOLLECTION AS Afx_IFileCollection_ PTR

#endif

' ########################################################################################
' Interface name = IFile
' IID = {C7C3F5A4-88A3-11D0-ABCB-00A0C90FFFC0}
' Attributes = 4304 [&H10D0] [Hidden] [Dual] [Nonextensible] [Dispatchable]
' Inherited interface = IDispatch
' ########################################################################################

#ifndef __Afx_IFile_INTERFACE_DEFINED__
#define __Afx_IFile_INTERFACE_DEFINED__

TYPE Afx_IFile_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION get_Path (BYVAL pbstrPath AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Name (BYVAL pbstrName AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Name (BYVAL pbstrName AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ShortPath (BYVAL pbstrPath AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ShortName (BYVAL pbstrName AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Drive (BYVAL ppdrive AS Afx_IDrive PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ParentFolder (BYVAL ppfolder AS Afx_IFolder PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Attributes (BYVAL pfa AS FILEATTRIBUTE PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Attributes (BYVAL pfa AS FILEATTRIBUTE) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_DateCreated (BYVAL pdate AS DATE_ PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_DateLastModified (BYVAL pdate AS DATE_ PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_DateLastAccessed (BYVAL pdate AS DATE_ PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Size (BYVAL pvarSize AS VARIANT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Type (BYVAL pbstrType AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Delete_ (BYVAL bForce AS VARIANT_BOOL = FALSE) AS HRESULT
   DECLARE ABSTRACT FUNCTION Copy (BYVAL Destination AS AFX_BSTR, BYVAL OverWriteFiles AS VARIANT_BOOL = -1) AS HRESULT
   DECLARE ABSTRACT FUNCTION Move (BYVAL Destination AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION OpenAsTextStream (BYVAL IOMode AS IOMODE = 1, BYVAL Format AS TRISTATE = 0, BYVAL ppts AS Afx_ITextStream PTR PTR) AS HRESULT
END TYPE

TYPE AFX_LPFILE AS Afx_IFile PTR

#endif

' ########################################################################################
' Interface name = ITextStream
' IID = {53BAD8C1-E718-11CF-893D-00A0C9054228}
' Attributes = 4304 [&H10D0] [Hidden] [Dual] [Nonextensible] [Dispatchable]
' Inherited interface = IDispatch
' ########################################################################################

#ifndef __Afx_ITextStream_INTERFACE_DEFINED__
#define __Afx_ITextStream_INTERFACE_DEFINED__

TYPE Afx_ITextStream_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION get_Line (BYVAL Line AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Column (BYVAL Column AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_AtEndOfStream (BYVAL EOS AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_AtEndOfLine (BYVAL EOL AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Read (BYVAL Characters AS LONG, BYVAL Text AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION ReadLine (BYVAL Text AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION ReadAll (BYVAL Text AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Write (BYVAL Text AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION WriteLine (BYVAL Text AS AFX_BSTR = NULL) AS HRESULT
   DECLARE ABSTRACT FUNCTION WriteBlankLines (BYVAL Lines AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION Skip (BYVAL Characters AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION SkipLine () AS HRESULT
   DECLARE ABSTRACT FUNCTION Close () AS HRESULT
END TYPE

TYPE AFX_LPTEXTSTREAM AS Afx_ITextStream PTR

#endif

' ########################################################################################
' Interface name = IFileSystem3
' IID = {2A0B9D10-4B87-11D3-A97A-00104B365C9F}
' Attributes = 4288 [&H10C0] [Dual] [Nonextensible] [Dispatchable]
' Inherited interface = IFileSystem
' ########################################################################################

#ifndef __Afx_IFileSystem3_INTERFACE_DEFINED__
#define __Afx_IFileSystem3_INTERFACE_DEFINED__

TYPE Afx_IFileSystem3_ EXTENDS Afx_IFileSystem
   DECLARE ABSTRACT FUNCTION GetStandardStream (BYVAL StandardStreamType AS STANDARDSTREAMTYPES, BYVAL bUnicode AS VARIANT_BOOL = FALSE, BYVAL ppts AS Afx_ITextStream PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetFileVersion (BYVAL FileName AS AFX_BSTR, BYVAL FileVersion AS AFX_BSTR PTR) AS HRESULT
END TYPE

TYPE AFX_LPFILESYSTEM3 AS Afx_IFileSystem3 PTR

#endif

' ########################################################################################
' Interface name = IScriptEncoder
' IID = {AADC65F6-CFF1-11D1-B747-00C04FC2B085}
' Attributes = 4160 [&H1040] [Dual] [Dispatchable]
' Inherited interface = IDispatch
' ########################################################################################

#ifndef __Afx_IScriptEncoder_INTERFACE_DEFINED__
#define __Afx_IScriptEncoder_INTERFACE_DEFINED__

TYPE Afx_IScriptEncoder_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION EncodeScriptFile (BYVAL szText AS AFX_BSTR, BYVAL bstrStreamIn AS AFX_BSTR, BYVAL cFlags AS LONG, BYVAL bstrDefaultLang AS AFX_BSTR, BYVAL pbstrStreamOut AS AFX_BSTR PTR) AS HRESULT
END TYPE

TYPE AFX_LPSCRIPTENCODER AS Afx_IScriptEncoder PTR

#endif

END NAMESPACE
