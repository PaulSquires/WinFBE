# CFileSys Class

The **CFileSys** class wraps the Microsoft File System Object and provides methods to work with files and folders, giving your application the ability to create, copy, alter, move, and delete files and folders, or to determine if and where particular files or folders exist. It also enables you to get information about files and folders, such as their names and the date they were created or last modified.

**Include file**: CFileSys.inc

### Methods

| Name       | Description |
| ---------- | ----------- |
| [BuildPath](#BuildPath) | Appends a name to an existing path. |
| [CopyFile](#CopyFile) | Copies one or more files from one location to another. |
| [CopyFolder](#CopyFolder) | Recursively copies a folder from one location to another. |
| [CreateFolder](#CreateFolder) | Creates a folder. |
| [DeleteFile](#DeleteFile) | Deletes a specified file. |
| [DeleteFolder](#DeleteFolder) | Deletes a specified folder and its contents. |
| [DriveExists](#DriveExists) | Checks if the specified drive exists. |
| [DriveLetters](#DriveLetters) | Returns a semicolon separated list with the driver letters. |
| [FileExists](#FileExists) | Checks for the existence of the specified file. |
| [FolderExists](#FolderExists) | Checks for the existence of the specified folder. |
| [GetAbsolutePathName](#GetAbsolutePathName) | Returns complete and unambiguous path from a provided path specification. |
| [GetBaseName](#GetBaseName) | Returns a string containing the base name of the last component, less any file extension, in a path. |
| [GetDriveAvailableSpace](#GetDriveAvailableSpace) | Returns the amount of space available to a user on the specified drive or network share. |
| [GetDriveFileSystem](#GetDriveFileSystem) | Returns the type of file system in use for the specified drive or network share. |
| [GetDriveFreeSpace](#GetDriveFreeSpace) | Returns the amount of free space available to a user on the specified drive or network share. |
| [GetDriveName](#GetDriveName) | Returns a string containing the name of the drive for a specified path. |
| [GetDriveShareName](#GetDriveShareName) | Returns the network share name for a specified drive. |
| [GetDriveTotalSize](#GetDriveTotalSize) | Returns the total space, in bytes, of a drive or network share. |
| [GetDriveType](#GetDriveType) | Returns a value indicating the type of a specified drive. |
| [GetExtensionName](#GetExtensionName) | Returns a string containing the extension name of the file for a specified path. |
| [GetFileAttributes](#GetFileAttributes) | Returns the attributes of files. Read/write or read-only, depending on the attribute. |
| [GetFileDateCreated](#GetFileDateCreated) | Returns the date and time that the specified file was created. |
| [GetFileDateLastAccessed](#GetFileDateLastAccessed) | Returns the date and time that the specified file was accessed. |
| [GetFileDateLastModified](#GetFileDateLastModified) | Returns the date and time that the specified file was modified. |
| [GetFileName](#GetFileName) | Returns a string containing the name of the file for a specified path. |
| [GetFileShortName](#GetFileShortName) | Returns the short name used by programs that require the earlier 8.3 file naming convention. |
| [GetFileShortPath](#GetFileShortPath) | Returns the short path used by programs that require the earlier 8.3 file naming convention. |
| [GetFileSize](#GetFileSize) | Returns the size, in bytes, of the specified file. |
| [GetFileType](#GetFileType) | Returns information about the type of a file. |
| [GetFileVersion](#GetFileVersion) | Returns the version number of a specified file. |
| [GetFolderAttributes](#GetFolderAttributes) | Returns the attributes of folders. |
| [GetFolderDateCreated](#GetFolderDateCreated) | Returns the date and time that the specified folder was created. |
| [GetFolderDateLastAccessed](#GetFolderDateLastAccessed) | Returns the date and time that the specified folder was last accessed. |
| [GetFolderDateLastModified](#GetFolderDateLastModified) | Returns the date and time that the specified folder was last modified. |
| [GetFolderDriveLetter](#GetFolderDriveLetter) | Returns a string containing the drive letter for a specified folder. |
| [GetFolderName](#GetFolderName) | Returns a string containing the name of the folder for a specified path, i.e. the path minus the file name. |
| [GetFolderShortName](#GetFolderShortName) | Returns the short name used by programs that require the earlier 8.3 file naming convention. |
| [GetFolderShortPath](#GetFolderShortPath) | Returns the short path used by programs that require the earlier 8.3 file naming convention. |
| [GetFolderSize](#GetFolderSize) | Returns the size, in bytes, of all files and subfolders contained in the folder. |
| [GetFolderType](#GetFolderType) | Returns information about the type of a folder. |
| [GetLastResult](#GetLastResult) | Returns the last result code. |
| [GetNumDrives](#GetNumDrives) | Returns the number of drives. |
| [GetNumFiles](#GetNumFiles) | Returns the number of files contained in a specified folder, including those with hidden and system file attributes set. |
| [GetNumSubFolders](#GetNumSubFolders) | Returns the number of folders contained in a specified folder, including those with hidden and system file attributes set. |
| [GetParentFolderName](#GetParentFolderName) | Returns the folder name for the parent of the specified folder. |
| [GetSerialNumber](#GetSerialNumber) | Returns the decimal serial number used to uniquely identify a disk volume. |
| [GetStandardStream](#GetStandardStream) | Returns a TextStream object corresponding to the standard input, output, or error stream. |
| [GetStdErrStream](#GetStdErrStream) | Returns a TextStream object corresponding to the standard error stream. |
| [GetStdInStream](#GetStdInStream) | Returns a TextStream object corresponding to the standard input stream. |
| [GetStdOutStream](#GetStdOutStream) | Returns a TextStream object corresponding to the standard output stream. |
| [GetTempName](#GetTempName) | Returns a randomly generated temporary file or folder name that is useful for performing operations that require a temporary file or folder. |
| [GetVolumeName](#GetVolumeName) | Returns the volume name of the specified drive. |
| [IsDriveReady](#IsDriveReady) | Returns True if the specified drive is ready; False if it is not. |
| [IsRootFolder](#IsRootFolder) | Returns True(-1) if the specified folder is the root folder; False(0) if it is not. |
| [MoveFile](#MoveFile) | Moves one or more files from one location to another. |
| [MoveFolder](#MoveFolder) | Moves one or more folders from one location to another. |
| [SetFileAttributes](#SetFileAttributes) | Sets the attributes of files. |
| [SetFileName](#SetFileName) | Sets the name of a specified file. |
| [SetFolderAttributes](#SetFolderAttributes) | Sets the attributes of folders. |
| [SetFolderName](#SetFolderName) | Sets the name of a specified folder. |
| [SetVolumeName](#SetVolumeName) | Sets the volume name of the specified drive. |

# <a name="BuildPath"></a>BuildPath

Appends a name to an existing path.

```
FUNCTION BuildPath (BYREF cbsPath AS CBSTR, BYREF cbsName AS CBSTR) AS CWSTR
```

| Name       | Description |
| ---------- | ----------- |
| *cbsPath* | CBSTR. Existing path to which name is appended. Path can be absolute or relative and need not specify an existing folder. |
| *cbsName* | CBSTR. Name being appended to the existing path. |

#### Return value

CBSTR. The new path.

#### Remarks

The **BuildPath** method inserts an additional path separator between the existing path and the name, only if necessary.

#### Usage example

```
#INCLUDE ONCE "Afx/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM cbsNewPath AS CBSTR = pFileSys.BuildPath ("C:\MyFolder", "Text.txt")
```

# <a name="CopyFile"></a>CopyFile

Copies one or more files from one location to another.

```
FUNCTION CopyFile (BYREF cbsSource AS CBSTR, BYREF cbsDestination AS CBSTR, _
   BYVAL OverWriteFiles AS VARIANT_BOOL = -1) AS HRESULT
```

| Name       | Description |
| ---------- | ----------- |
| *cbsSource* | CBSTR. Character string file specification, which can include wildcard characters, for one or more files to be copied. |
| *cbsDestination* | CBSTR. Character string destination where the file or files from source are to be copied. Wildcard characters are not allowed. |
| *OverWriteFiles* | Boolean value that indicates if existing files are to be overwritten. If true, files are overwritten; if false, they are not. The default is true. Note that **CopyFile** will fail if destination has the read-only attribute set, regardless of the value of overwrite. |

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

#### Remarks

Wildcard characters can only be used in the last path component of the *cbSource* argument.

#### Usage example

```
#INCLUDE ONCE "Afx/CFileSys.inc"
DIM pFileSys AS CFileSys
pFileSys.CopyFile("C:\MyFolder\MyFile.txt", "C:\MyOtherFolder\MyFile.txt")
```

# <a name="CopyFolder"></a>CopyFolder

Recursively copies a folder from one location to another.

```
FUNCTION CopyFolder (BYREF cbsSource AS CBSTR, BYREF cbsDestination AS CBSTR, _
   BYVAL OverWriteFiles AS VARIANT_BOOL = -1) AS HRESULT
```

| Name       | Description |
| ---------- | ----------- |
| *cbsSource* | CBSTR. Character string file specification, which can include wildcard characters, for one or more folders to be copied. |
| *cbsDestination* | CBSTR. Character string destination where the folder and subfolders from source are to be copied (must end with a "\"). Wildcard characters are not allowed.  |
| *OverWriteFiles* | Boolean value that indicates if existing folders are to be overwritten. If true, files are overwritten; if false, they are not. The default is true. |

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

#### Remarks

Wildcard characters can only be used in the last path component of the *cbSource* argument.

#### Usage example

```
#INCLUDE ONCE "Afx/CFileSys.inc"
DIM pFileSys AS CFileSys
pFileSys.CopyFolder("C:\MyFolder", "C:\MyOtherFolder\")
```

# <a name="CreateFolder"></a>CreateFolder

Creates a folder.

```
FUNCTION CreateFolder (BYREF cbsFolderSpec AS CBSTR) AS Afx_IFolder PTR
```

| Name       | Description |
| ---------- | ----------- |
| *cbsFolderSpec* | CBSTR. String expression that identifies the folder to create. |

#### Return value

IDispatch. Reference to an **IFolder** object.

#### Remarks

An error occurs if the specified folder already exists.

#### Usage example

```
#INCLUDE ONCE "Afx/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM pFolder AS Afx_Folder PTR
pFolder = pFileSys.CreateFolder("C:\MyNewFolder")
IF pFolder THEN
   ' ....
   pFolder.Release
END IF
```

# <a name="DeleteFile"></a>DeleteFile

Deletes a specified file.

```
FUNCTION DeleteFile (BYREF cbsFileSpec AS CBSTR, BYVAL bForce AS VARIANT_BOOL = FALSE) AS HRESULT
```

| Name       | Description |
| ---------- | ----------- |
| *cbsFileSpec* | CBSTR. The name of the file to delete. cbsFileSpec can contain wildcard characters in the last path component. |
| *bForce* | Boolean value that is true if files with the read-only attribute set are to be deleted; false (default) if they are not. |

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

#### Remarks

An error occurs if no matching files are found. The **DeleteFile** method stops on the first error it encounters. No attempt is made to roll back or undo any changes that were made before an error occurred.

#### Usage example

```
#INCLUDE ONCE "Afx/CFileSys.inc"
DIM pFileSys AS CFileSys
pFileSys.DeleteFile("C:\MyFolder\MyFile.txt")
```

# <a name="DeleteFolder"></a>DeleteFolder

Deletes a specified folder and its contents.

```
FUNCTION DeleteFolder (BYREF cbsFolderSpec AS CBSTR, BYVAL bForce AS VARIANT_BOOL = FALSE) AS HRESULT
```

| Name       | Description |
| ---------- | ----------- |
| *cbsFolderSpec* | CBSTR. The name of the folder to delete. *cbsFolderSpec* can contain wildcard characters in the last path component. |
| *bForce* | Boolean value that is true if folders with the read-only attribute set are to be deleted; false (default) if they are not. |

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

#### Remarks

The **DeleteFolder** method does not distinguish between folders that have contents and those that do not. The specified folder is deleted regardless of whether or not it has contents. 

An error occurs if no matching folders are found. The **DeleteFolder** method stops on the first error it encounters. No attempt is made to roll back or undo any changes that were made before an error occurred.

#### Usage example

```
#INCLUDE ONCE "Afx/CFileSys.inc"
DIM pFileSys AS CFileSys
pFileSys.DeleteFolder("C:\MyFolder")
```

# <a name="DriveExists"></a>DriveExists

Returns True if the specified drive exists; False if it does not.

```
FUNCTION DriveExists (BYREF cbsDriveSpec AS CBSTR) AS BOOLEAN
```

| Name       | Description |
| ---------- | ----------- |
| *cbsDriveSpec* | CBSTR. A drive letter or a complete path specification. |

#### Return value

BOOLEAN. True if the specified drive exists; False if it does not.

#### Remarks

For drives with removable media, the **DriveExists** method returns true even if there are no media present. Use the **IsDriveReady** method to determine if a drive is ready.

#### Usage example

```
#INCLUDE ONCE "Afx/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM fExists AS BOOLEAN = pFileSys.DriveExists("C:")
```

# <a name="DriveLetters"></a>DriveLetters

Returns a semicolon separated list with the driver letters.

```
FUNCTION DriveLetters () AS CBSTR
```

#### Return value

CBSTR. A semicolon separated list with the driver letters.

#### Usage example

```
#INCLUDE ONCE "Afx/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM cbsLetters AS CBSTR = pFileSys.DriveLetters
```

# <a name="FileExists"></a>FileExists

Returns True if the specified file exists; False if it does not.

```
FUNCTION FileExists (BYREF cbsFileSpec AS CBSTR) AS BOOLEAN
```

| Name       | Description |
| ---------- | ----------- |
| *cbsFileSpec* | CBSTR. The name of the file whose existence is to be determined. A complete path specification (either absolute or relative) must be provided if the file isn't expected to exist in the current folder. |

#### Return value

Boolean. True if the specified file exists; False if it does not.

#### Usage example

```
#INCLUDE ONCE "Afx/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM fExists AS BOOLEAN = pFileSys.FileExists("C:\MyFolder\Test.txt")
```

# <a name="FolderExists"></a>FolderExists

Returns True if the specified folder exists; False if it does not.

```
FUNCTION FolderExists (BYREF cbsFolderSpec AS CBSTR) AS BOOLEAN
```

| Name       | Description |
| ---------- | ----------- |
| *cbsFolderSpec* | CBSTR. The name of the folder whose existence is to be determined. A complete path specification (either absolute or relative) must be provided if the folder isn't expected to exist in the current folder. |

#### Return value

Boolean. True if the specified folder exists; False if it does not.

#### Usage example

```
#INCLUDE ONCE "Afx/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM fExists AS BOOLEAN = pFileSys.FolderExists("C:\MyFolder")
```

# <a name="GetAbsolutePathName"></a>GetAbsolutePathName

Returns complete and unambiguous path from a provided path specification.

```
FUNCTION GetAbsolutePathName (BYREF cbsPathSpec AS CBSTR) AS CBSTR
```

| Name       | Description |
| ---------- | ----------- |
| *cbsPathSpec* | CBSTR. Path specification to change to a complete and unambiguous path. |

#### Return value

Boolean. The path name.

#### Remarks

A path is complete and unambiguous if it provides a complete reference from the root of the specified drive. A complete path can only end with a path separator character (\) if it specifies the root folder of a mapped drive.

Assuming the current directory is c:\mydocuments\reports, the following table illustrates the behavior of the **GetAbsolutePathName** method.

| pathspec   | Returned path |
| ---------- | ------------- |
| "c:" | "c:\mydocuments\reports" |
| "c:.." | "c:\mydocuments" |
| "c:\" | "c:\" |
| "c:*.*\may97" | "c:\mydocuments\reports\\\*\.\*\may97" |
| "region1" | "c:\mydocuments\reports\region1" |
| "c:\..\..\mydocuments" | "c:\mydocuments" |

#### Usage example

```
#INCLUDE ONCE "Afx/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM cbsName AS CBSTR = pFileSys.GetAbsolutePathName("C:\MyFolder\Test.txt")
```

# <a name="GetBaseName"></a>GetBaseName

Returns a string containing the base name of the last component, less any file extension, in a path

```
FUNCTION GetBaseName (BYREF cbsPathSpec AS CBSTR) AS CBSTR
```

| Name       | Description |
| ---------- | ----------- |
| *cbsPathSpec* | CBSTR. The path specification for the component whose base name is to be returned. |

#### Return value

Boolean. The base name.

#### Remarks

The **GetBaseName** method returns a zero-length string ("") if no component matches the path argument. The **GetBaseName** method works only on the provided path string. It does not attempt to resolve the path, nor does it check for the existence of the specified path.

#### Usage example

```
#INCLUDE ONCE "Afx/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM cwsName AS CWSTR = pFileSys.GetBaseName("C:\MyFolder\Test.txt")
```

# <a name="GetDriveAvailableSpace"></a>GetDriveAvailableSpace

Returns the amount of space available to a user on the specified drive or network share.

```
FUNCTION GetDriveAvailableSpace (BYREF cbsDrive AS CBSTR) AS DOUBLE
```

| Name       | Description |
| ---------- | ----------- |
| *cbsDrive* | CBSTR. The drive letter. For drive letters, the root drive is not included. For example, the path for the C drive is C:, not C:\\. |

#### Return value

DOUBLE. The amount of available space.

#### Remarks

The value returned by the **GetDriveAvailableSpace** method is typically the same as that returned by the **GetDriveFreeSpace** method. Differences may occur between the two for computer systems that support quotas. 

#### Usage example

```
#INCLUDE ONCE "Afx/CFileSys.inc"
DIM pFileSys AS CFileSys
PRINT pFileSys.GetDriveAvailableSpace("C:")
```

# <a name="GetDriveFileSystem"></a>GetDriveFileSystem

Returns the type of file system in use for the specified drive or network share.

```
FUNCTION GetDriveFileSystem (BYREF cbsDrive AS CBSTR) AS CBSTR
```

| Name       | Description |
| ---------- | ----------- |
| *cbsDrive* | CBSTR. The drive letter. For drive letters, the root drive is not included. For example, the path for the C drive is C:, not C:\\. |

#### Return value

CBSTR. The type of file system in use.

#### Remarks

Available return types include FAT, NTFS, and CDFS.

#### Usage example

```
#INCLUDE ONCE "Afx/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM cbsFileSystem AS CBSTR = pFileSys.GetDriveFileSystem("C:")
```

# <a name="GetDriveFreeSpace"></a>GetDriveFreeSpace

Returns the amount of free space available to a user on the specified drive or network share

```
FUNCTION GetDriveFreeSpace (BYREF cbsDrive AS CBSTR) AS DOUBLE
```

| Name       | Description |
| ---------- | ----------- |
| *cbsDrive* | CBSTR. The drive letter. For drive letters, the root drive is not included. For example, the path for the C drive is C:, not C:\\. |

#### Return value

DOUBLE. The amount of free space.

#### Remarks

The value returned by the **GetDriveFreeSpace** property is typically the same as that returned by the **GetDriveAvailableSpace** property. Differences may occur between the two for computer systems that support quotas. 

#### Usage example

```
#INCLUDE ONCE "Afx/CFileSys.inc"
DIM pFileSys AS CFileSys
PRINT pFileSys.GetDriveFreeSpace("C:")
```

# <a name="GetDriveName"></a>GetDriveName

Returns a string containing the name of the drive for a specified path.

```
FUNCTION GetDriveName (BYREF cbsPathSpec AS CBSTR) AS CBSTR
```

| Name       | Description |
| ---------- | ----------- |
| *cbsPathSpec* | CBSTR. The path. |

#### Return value

CBSTR. The drive name.

#### Usage example

```
#INCLUDE ONCE "Afx/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM cbsName AS CBSTR = pFileSys.GetDriveName("C:\MyFolder\Test.txt")
```

# <a name="GetDriveShareName"></a>GetDriveShareName

Returns the network share name for a specified drive.

```
FUNCTION GetDriveShareName (BYREF cbsDrive AS CBSTR) AS CBSTR
```

| Name       | Description |
| ---------- | ----------- |
| *cbsDrive* | CBSTR. The drive letter. For drive letters, the root drive is not included. For example, the path for the C drive is C:, not C:\\. |

#### Return value

CBSTR. The share name.

#### Remarks

If object is not a network drive, the **GetDriveShareName** method returns a zero-length string ("").

#### Usage example

```
#INCLUDE ONCE "Afx/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM cbsShareName AS CBSTR = pFileSys.GetDriveShareName("H:")
```

# <a name="GetDriveTotalSize"></a>GetDriveTotalSize

Returns the total space, in bytes, of a drive or network share.

```
FUNCTION GetDriveTotalSize (BYREF cbsDrive AS CBSTR) AS DOUBLE
```

| Name       | Description |
| ---------- | ----------- |
| *cbsDrive* | CBSTR. The drive letter. For drive letters, the root drive is not included. For example, the path for the C drive is C:, not C:\\. |

#### Return value

DOUBLE. The total space in bytes.

#### Usage example

```
#INCLUDE ONCE "Afx/CFileSys.inc"
DIM pFileSys AS CFileSys
PRINT pFileSys.GetDriveTotalSize("C:")
```

# <a name="GetDriveType"></a>GetDriveType

Returns a value indicating the type of a specified drive.

```
FUNCTION GetDriveType (BYREF cbsDrive AS CBSTR) AS DRIVETYPECONST
```

| Name       | Description |
| ---------- | ----------- |
| *cbsDrive* | CBSTR. The drive letter. For drive letters, the root drive is not included. For example, the path for the C drive is C:, not C:\\. |

#### Return value

The type of the specified drive.

```
DriveType_UnknownType = 0
DriveType_Removable = 1
DriveType_Fixed = 2
DriveType_Remote = 3
DriveType_CDRom = 4
DriveType_RamDisk = 5
```
#### Usage example

```
#INCLUDE ONCE "Afx/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM nDriveType AS DRIVETYPECONST = pFileSys.GetDriveType("C:")
DIM t AS CBSTR
SELECT CASE pDrive.DriveType
   CASE 0 : t = "Unknown"
   CASE 1 : t = "Removable"
   CASE 2 : t = "Fixed"
   CASE 3 : t = "Network"
   CASE 4 : t = "CD-ROM"
   CASE 5 : t = "RAM Disk"
END SELECT
```

# <a name="GetExtensionName"></a>GetExtensionName

Returns a string containing the extension name of the file for a specified path.

```
FUNCTION GetExtensionName (BYREF cbsPathSpec AS CBSTR) AS CBSTR
```

| Name       | Description |
| ---------- | ----------- |
| *cbsPathSpec* | CBSTR. The extension name. |

#### Usage example

```
#INCLUDE ONCE "Afx/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM cbsName AS CBSTR = pFileSys.GetExtensionName("C:\MyFolder\Test.txt")
```

# <a name="GetFileAttributes"></a>GetFileAttributes

Returns the attributes of files.

```
FUNCTION GetFileAttributes (BYREF cbsFile AS CBSTR) AS FILEATTRIBUTE
```

| Name       | Description |
| ---------- | ----------- |
| *cbsFile* | CBSTR. The path to a specific file. |

#### Return value

The file attributes. Can be any of the following values or any logical combination of the following values:

| Constant   | Value       | Description |
| ---------- | ----------- | ----------- |
| FileAttribute_Normal     | 0 | Normal file. No attributes are set. |
| FileAttribute_ReadOnly   | 1 | Read-only file. Attribute is read/write. |
| FileAttribute_Hidden     | 2 | Hidden file. Attribute is read/write. |
| FileAttribute_System     | 4 | System file. Attribute is read/write. |
| FileAttribute_Volume     | 8 | Disk drive volume label. Attribute is read-only. |
| FileAttribute_Directory | 16 | Folder or directory. Attribute is read-only. |
| FileAttribute_Archive |   32 | File has changed since last backup. Attribute is read/write. |
| FileAttribute_Alias   | 1024 | Link or shortcut. Attribute is read-only. |
| FileAttribute_Compressed | 2048 | Compressed file. Attribute is read-only. |

#### Usage example

```
#INCLUDE ONCE "Afx/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM lAttr FILEATTRIBUTE = pFileSys.GetFileAttributes("C:\MyPath\MyFile.txt")
```

# <a name="GetFileDateCreated"></a>GetFileDateCreated

Returns the date and time that the specified file was created.

```
FUNCTION GetFileDateCreated (BYREF cbsFile AS CBSTR) AS DATE_
```

| Name       | Description |
| ---------- | ----------- |
| *cbsFile* | CBSTR. The path to a specific file. |

#### Return value

DATE_. The date and time of the file creation. This is a *Date Serial* number that can be formatted using the Free Basic's **Format** function. You can also use the wrapper function **AfxVariantDateTimeToStr**.

#### Usage example

```
#INCLUDE ONCE "Afx/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM nDate AS DATE_ = pFileSys.GetFileDateCreated("C:\MyPath\MyFile.txt")
```

# <a name="GetFileDateLastAccessed"></a>GetFileDateLastAccessed

Returns the date and time that the specified file was accesed.

```
FUNCTION GetFileDateLastAccessed (BYREF cbsFile AS CBSTR) AS DATE_
```

| Name       | Description |
| ---------- | ----------- |
| *cbsFile* | CBSTR. The path to a specific file. |

#### Return value

DATE_. The date and time that the file was last accessed. This is a *Date Serial* number that can be formatted using the Free Basic's **Format** function. You can also use the wrapper function **AfxVariantDateTimeToStr**.

#### Usage example

```
#INCLUDE ONCE "Afx/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM nDate AS DATE_ = pFileSys.GetFileDateLastAccessed("C:\MyPath\MyFile.txt")
```

# <a name="GetFileDateLastModified"></a>GetFileDateLastModified

Returns the date and time that the specified file was accessed.

```
FUNCTION GetFileDateLastModified (BYREF cbsFile AS CBSTR) AS DATE_
```

| Name       | Description |
| ---------- | ----------- |
| *cbsFile* | CBSTR. The path to a specific file. |

#### Return value

DATE_. The date and time that the file was last modified. This is a *Date Serial* number that can be formatted using the Free Basic's **Format** function. You can also use the wrapper function **AfxVariantDateTimeToStr**.

#### Usage example

```
#INCLUDE ONCE "Afx/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM nDate AS DATE_ = pFileSys.GetFileDateLastModified("C:\MyPath\MyFile.txt")
```

# <a name="GetFileName"></a>GetFileName

Returns a string containing the name of the file for a specified path.

```
FUNCTION GetFileName (BYREF cbsPathSpec AS CBSTR) AS CBSTR
```

| Name       | Description |
| ---------- | ----------- |
| *cbsPathSpec* | CBSTR. The path to a specific file. |

#### Return value

CBSTR. The file name.

#### Usage example

```
#INCLUDE ONCE "Afx/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM cbsName AS CBSTR = pFileSys.GetFileName("C:\MyFolder\Test.txt")
```

# <a name="GetFileShortName"></a>GetFileShortName

Returns the short name used by programs that require the earlier 8.3 file naming convention.

```
FUNCTION GetFileShortName (BYREF cbsPathSpec AS CBSTR) AS CBSTR
```

| Name       | Description |
| ---------- | ----------- |
| *cbsPathSpec* | CBSTR. The path to a specific file. |

#### Return value

CBSTR. The short name for the specified file.

#### Usage example

```
#INCLUDE ONCE "Afx/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM cbsName AS CBSTR = pFileSys.GetFileShortName("C:\MyFolder\Test.txt")
```

# <a name="GetFileShortPath"></a>GetFileShortPath

Returns the short path used by programs that require the earlier 8.3 file naming convention.

```
FUNCTION GetFileShortPath (BYREF cbsPathSpec AS CBSTR) AS CBSTR
```

| Name       | Description |
| ---------- | ----------- |
| *cbsPathSpec* | CBSTR. The path to a specific file. |

#### Return value

CBSTR. The short path to a specific file.

#### Usage example

```
#INCLUDE ONCE "Afx/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM cbsName AS CBSTR = pFileSys.GetFileShortPath("C:\MyFolder\Test.txt")
```

# <a name="GetFileSize"></a>GetFileSize

Returns the size, in bytes, of the specified file.

```
FUNCTION GetFileSize (BYREF cbsFile AS CBSTR) AS LONG
```

| Name       | Description |
| ---------- | ----------- |
| *cbsFile* | CBSTR. The path to a specific file. |

#### Return value

LONG. The size, in bytes, of the file.

#### Usage example

```
#INCLUDE ONCE "Afx/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM nFileSize AS LONG = pFileSys.GetFileSize("C:\MyPath\MyFile.txt")
```

# <a name="GetFileType"></a>GetFileType

Returns information about the type of a file. For example, for files ending in .TXT, "Text Document" is returned.

```
FUNCTION GetFileType (BYREF cbsFile AS CBSTR) AS CBSTR
```

| Name       | Description |
| ---------- | ----------- |
| *cbsFile* | CBSTR. The path to a specific file. |

#### Return value

CBSTR. The type of file.

#### Usage example

```
#INCLUDE ONCE "Afx/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM cbsFileType AS CBSTR = pFileSys.GetFileType("C:\MyPath\MyFile.txt")
```

# <a name="GetFileVersion"></a>GetFileVersion

Returns the version number of a specified file.

```
FUNCTION GetFileVersion (BYREF cbsFile AS CBSTR) AS CBSTR
```

| Name       | Description |
| ---------- | ----------- |
| *cbsFile* | CBSTR. The path (absolute or relative) to a specific file. |

#### Return value

CBSTR. The version number.

#### Remarks

The **GetFileVersion** method returns a zero-length string ("") if *cbsFile* does not end with the named component. 

Note: The **GetFileVersion** method works only on the provided path string. It does not attempt to resolve the path, nor does it check for the existence of the specified path.

#### Usage example

```
#INCLUDE ONCE "Afx/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM cbsVersion AS CBSTR = pFileSys.GetFileVersion("C:\MyFolder\MyFile.doc")
IF LEN(cbsVersion) THEN
   MSGBOX "File version: " & cbsVersion
ELSE
   MSGBOX "No version information available"
END IF
```
```
DIM pFileSys AS CFileSys
print pFileSys.GetFileVersion("c:\windows\system32\scrrun.dll")
```

# <a name="GetFolderAttributes"></a>GetFolderAttributes

Returns the attributes of folders.

```
FUNCTION GetFolderAttributes (BYREF cbsFolder AS CBSTR) AS FILEATTRIBUTE
```

| Name       | Description |
| ---------- | ----------- |
| *cbsFolder* | CBSTR. The path to a specific folder. |

#### Return value

The attributes. Can be any of the following values or any logical combination of the following values:

| Constant   | Value       | Description |
| ---------- | ----------- | ----------- |
| FileAttribute_Normal     | 0 | Normal file. No attributes are set. |
| FileAttribute_ReadOnly   | 1 | Read-only file. Attribute is read/write. |
| FileAttribute_Hidden     | 2 | Hidden file. Attribute is read/write. |
| FileAttribute_System     | 4 | System file. Attribute is read/write. |
| FileAttribute_Volume     | 8 | Disk drive volume label. Attribute is read-only. |
| FileAttribute_Directory | 16 | Folder or directory. Attribute is read-only. |
| FileAttribute_Archive |   32 | File has changed since last backup. Attribute is read/write. |
| FileAttribute_Alias   | 1024 | Link or shortcut. Attribute is read-only. |
| FileAttribute_Compressed | 2048 | Compressed file. Attribute is read-only. |

#### Usage example

```
#INCLUDE ONCE "Afx/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM lAttr FILEATTRIBUTE = pFileSys.GetFolderAttributes("C:\MyPath")
```

# <a name="GetFolderDateCreated"></a>GetFolderDateCreated

Returns the date and time that the specified folder was created.

```
FUNCTION GetFolderDateCreated (BYREF cbsFolder AS CBSTR) AS DATE_
```

| Name       | Description |
| ---------- | ----------- |
| *cbsFolder* | CBSTR. The path to a specific folder. |

#### Return value

DATE_. The date and time that the folder was created. This is a *Date Serial* number that can be formatted using the Free Basic's **Format** function. You can also use the wrapper function **AfxVariantDateTimeToStr**.

#### Usage example

```
#INCLUDE ONCE "Afx/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM nDate AS DATE_ = pFileSys.GetFolderDateCreated("C:\MyPath")
```

# <a name="GetFolderDateLastAccessed"></a>GetFolderDateLastAccessed

Returns the date and time that the specified folder was last accessed.

```
FUNCTION GetFolderDateLastAccessed (BYREF cbsFolder AS CBSTR) AS DATE_
```

| Name       | Description |
| ---------- | ----------- |
| *cbsFolder* | CBSTR. The path to a specific folder. |

#### Return value

DATE_. The date and time that the folder was last accessed. This is a *Date Serial* number that can be formatted using the Free Basic's **Format** function. You can also use the wrapper function **AfxVariantDateTimeToStr**.

#### Usage example

```
#INCLUDE ONCE "Afx/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM nDate AS DATE_ = pFileSys.GetFolderDateLastAccessed("C:\MyPath)
```

# <a name="GetFolderDateLastModified"></a>GetFolderDateLastModified

Returns the date and time that the specified folder was last modified.

```
FUNCTION GetFolderDateLastModified (BYREF cbsFolder AS CBSTR) AS DATE_
```

| Name       | Description |
| ---------- | ----------- |
| *cbsFolder* | CBSTR. The path to a specific folder. |

#### Return value

DATE_. The date and time that the folder was last modified. This is a *Date Serial* number that can be formatted using the Free Basic's **Format** function. You can also use the wrapper function **AfxVariantDateTimeToStr**.

#### Usage example

```
#INCLUDE ONCE "Afx/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM nDate AS DATE_ = pFileSys.GetFolderDateLastModified("C:\MyPath)
```

# <a name="GetFolderDriveLetter"></a>GetFolderDriveLetter

Returns a string containing the drive letter for a specified folder.

```
FUNCTION GetFolderDriveLetter (BYREF cbsFolder AS CBSTR) AS CBSTR
```

| Name       | Description |
| ---------- | ----------- |
| *cbsFolder* | CBSTR. The path to a specific folder. |

#### Return value

CBSTR. The drive letter of the folder.

#### Usage example

```
#INCLUDE ONCE "Afx/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM cbsDriveLetter AS CBSTR = pFileSys.GetFolderDriveLetter("c:\MyFolder)
```

# <a name="GetFolderName"></a>GetFolderName

Returns a string containing the name of the folder for a specified path, i.e. the path minus the file name.

```
FUNCTION GetFolderName (BYREF cbsPathSpec AS CBSTR) AS CBSTR
```

| Name       | Description |
| ---------- | ----------- |
| *cbsFolder* | CBSTR. The path to a specific folder. |

#### Return value

CBSTR. The name of the folder.

#### Usage example

```
#INCLUDE ONCE "Afx/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM cbsName AS CBSTR = pFileSys.GetFolderName("C:\MyFolder\Test.txt")
```

# <a name="GetFolderShortName"></a>GetFolderShortName

Returns the short name used by programs that require the earlier 8.3 file naming convention.

```
FUNCTION GetFolderShortName (BYREF cbsPathSpec AS CBSTR) AS CBSTR
```

| Name       | Description |
| ---------- | ----------- |
| *cbsFolder* | CBSTR. The path to a specific folder. |

#### Return value

CBSTR. The short name for the specified folder.

#### Usage example

```
#INCLUDE ONCE "Afx/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM cbsFolderShortPath AS CBSTR = pFileSys.GetFolderShortName("c:\MyFolder)
```

# <a name="GetFolderShortPath"></a>GetFolderShortPath

Returns the short path used by programs that require the earlier 8.3 file naming convention.

```
FUNCTION GetFolderShortPath (BYREF cbsPathSpec AS CBSTR) AS CBSTR
```

| Name       | Description |
| ---------- | ----------- |
| *cbsFolder* | CBSTR. The path to a specific folder. |

#### Return value

CBSTR. The short path for the specified folder.

#### Usage example

```
#INCLUDE ONCE "Afx/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM cbsFolderShortPath AS CBSTR = pFileSys.GetFolderShortPath("c:\MyFolder)
```

# <a name="GetFolderSize"></a>GetFolderSize

Returns the size, in bytes, of all files and subfolders contained in the folder.

```
FUNCTION GetFolderSize (BYREF cbsFolder AS CBSTR) AS LONG
```

| Name       | Description |
| ---------- | ----------- |
| *cbsFolder* | CBSTR. The path to a specific folder. |

#### Return value

LONG. The size, in bytes, of all files and subfolders contained in the folder.

#### Usage example

```
#INCLUDE ONCE "Afx/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM nFolderSize AS LONG = pFileSys.GetFolderSize("C:\MyPath")
```

# <a name="GetFolderType"></a>GetFolderType

Returns information about the type of a folder.

```
FUNCTION GetFolderType (BYREF cbsFolder AS CBSTR) AS CBSTR
```

| Name       | Description |
| ---------- | ----------- |
| *cbsFolder* | CBSTR. The path to a specific folder. |

#### Return value

CBSTR. The type of folder.

#### Usage example

```
#INCLUDE ONCE "Afx/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM cbsFolderType AS CBSTR = pFileSys.GetFolderType("c:\MyFolder)
```

# <a name="GetLastResult"></a>GetLastResult

Returns the last result code.

```
FUNCTION GetLastResult () AS HRESULT
```

#### Return value

HRESULT. The result code returned by the last executed method.

# <a name="GetNumDrives"></a>GetNumDrives

Returns the number of drives.

```
FUNCTION GetNumDrives () AS LONG
```

#### Return value

LONG. The number of drives.

#### Usage example

```
#INCLUDE ONCE "Afx/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM numDrives AS LONG = pFileSys.GetNumDrives
```

# <a name="GetNumFiles"></a>GetNumFiles

Returns the number of files contained in a specified folder, including those with hidden and system file attributes set.

```
FUNCTION GetNumFiles (BYREF cbsFolder AS CBSTR) AS LONG
```

| Name       | Description |
| ---------- | ----------- |
| *cbsFolder* | CBSTR. The path to a specific folder. |

#### Return value

LONG. The number of files.

#### Usage example

```
#INCLUDE ONCE "Afx/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM numFiles AS LONG = pFileSys.GetNumFiles("C:\MyFolder")
```

# <a name="GetNumSubFolders"></a>GetNumSubFolders

Returns the number of files contained in a specified folder, including those with hidden and system file attributes set.

```
FUNCTION GetNumSubFolders (BYREF cbsFolder AS CBSTR) AS LONG
```

| Name       | Description |
| ---------- | ----------- |
| *cbsFolder* | CBSTR. The path to a specific folder. |

#### Return value

LONG. The number of subfolders.

#### Usage example

```
#INCLUDE ONCE "Afx/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM numSubFolders AS LONG = pFileSys.GetNumSubFolders("C:\MyFolder")
```

# <a name="GetParentFolderName"></a>GetParentFolderName

Returns the folder name for the parent of the specified folder.

```
FUNCTION GetParentFolderName (BYREF cbsFolder AS CBSTR) AS CBSTR
```

| Name       | Description |
| ---------- | ----------- |
| *cbsFolder* | CBSTR. The path to a specific folder. |

#### Return value

CBSTR. The name of the parent folder, or a empty string if the folder has no parent.

#### Usage example

```
#INCLUDE ONCE "Afx/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM cbsParentFolderName AS CBSTR = pFileSys.GetParentFolderName("C:\MyFolder\MySubfolder")
```

# <a name="GetSerialNumber"></a>GetSerialNumber

Returns the decimal serial number used to uniquely identify a disk volume.

```
FUNCTION GetSerialNumber (BYREF cbsDrive AS CBSTR) AS LONG
```

| Name       | Description |
| ---------- | ----------- |
| *cbsDrive* | CBSTR. The drive letter. For drive letters, the root drive is not included. For example, the path for the C drive is C:, not C:\\. |

#### Return value

LONG. The serial number.

#### Remarks

You can use the **GetSerialNumber** method to ensure that the correct disk is inserted in a drive with removable media. 

#### Usage example

```
#INCLUDE ONCE "Afx/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM nSerialNumber AS LONG = pFileSys.GetSerialNumber("C:")
```

# <a name="GetStandardStream"></a>GetStandardStream

Returns a TextStream object corresponding to the standard input, output, or error stream.

```
FUNCTION GetStandardStream (BYVAL nStreamType AS STANDARDSTREAMTYPES, _
   BYVAL bUnicode AS VARIANT_BOOL = FALSE) AS Afx_ITextStream PTR
```

| Name       | Description |
| ---------- | ----------- |
| *nStreamType* | LONG. Can be one of three constants: *StandardStreamTypes_StdErr*, *StandardStreamTypes_StdIn*, or *StandardStreamTypes_StdOut*. |
| *bUnicode* | Boolean value that indicates whether the file is created as a Unicode or ASCII file. The value is true if the file is created as a Unicode file, false if it is created as an ASCII file. If omitted, an ASCII file is assumed. |

#### Return value

A pointer to the *ITextStream* interface of the requested standard stream.

#### Settings

The nStreamType argument can have any of the following settings:

| Constant   | Constant    | Description |
| ---------- | ----------- | ----------- |
| StandardStreamTypes_StdIn  | 0 | Returns a TextStream object corresponding to the standard input stream. |
| StandardStreamTypes_StdOut | 1 | Returns a TextStream object corresponding to the standard output stream. |
| StandardStreamTypes_StdErr | 2 | Returns a TextStream object corresponding to the standard error stream. |

#### Usage example

```
#INCLUDE ONCE "Afx/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM pStm AS Afx_ITextStream PTR = pFileSys.GetStandardStream(StandardStreamTypes_StdOut)
```

# <a name="GetStdErrStream"></a>GetStdErrStream

Returns a TextStream object corresponding to the standard error stream.

```
FUNCTION GetStdErrStream (BYVAL bUnicode AS VARIANT_BOOL = FALSE) AS Afx_ITextStream PTR
```

| Name       | Description |
| ---------- | ----------- |
| *bUnicode* | Boolean value that indicates whether the file is created as a Unicode or ASCII file. The value is true if the file is created as a Unicode file, false if it is created as an ASCII file. If omitted, an ASCII file is assumed. |

#### Return value

A pointer to the *ITextStream* interface of the standard error stream.

#### Usage example

```
#INCLUDE ONCE "Afx/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM pStm AS Afx_ITextStream PTR = pFileSys.GetStdErrStream
```

# <a name="GetStdInStream"></a>GetStdInStream

Returns a TextStream object corresponding to the standard input stream.

```
FUNCTION GetStdInStream (BYVAL bUnicode AS VARIANT_BOOL = FALSE) AS Afx_ITextStream PTR
```

| Name       | Description |
| ---------- | ----------- |
| *bUnicode* | Boolean value that indicates whether the file is created as a Unicode or ASCII file. The value is true if the file is created as a Unicode file, false if it is created as an ASCII file. If omitted, an ASCII file is assumed. |

#### Return value

A pointer to the *ITextStream* interface of the standard input stream.

#### Usage example

```
#INCLUDE ONCE "Afx/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM pStm AS Afx_ITextStream PTR = pFileSys.GetStdInStream
```

# <a name="GetStdOutStream"></a>GetStdOutStream

Returns a TextStream object corresponding to the standard output stream.

```
FUNCTION GetStdOutStream (BYVAL bUnicode AS VARIANT_BOOL = FALSE) AS Afx_ITextStream PTR
```

| Name       | Description |
| ---------- | ----------- |
| *bUnicode* | Boolean value that indicates whether the file is created as a Unicode or ASCII file. The value is true if the file is created as a Unicode file, false if it is created as an ASCII file. If omitted, an ASCII file is assumed. |

#### Return value

A pointer to the *ITextStream* interface of the standard output stream.

#### Usage example

```
#INCLUDE ONCE "Afx/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM pStm AS Afx_ITextStream PTR = pFileSys.GetStdOutStream
```

# <a name="GetTempName"></a>GetTempName

Returns a randomly generated temporary file or folder name that is useful for performing operations that require a temporary file or folder.

```
FUNCTION GetTempName () AS CBSTR
```

#### Return value

CBSTR. The temporary name.

#### Remarks
The **GetTempName** method does not create a file. It provides only a temporary file name that can be used with **CreateTextFile** to create a file.

#### Usage example

```
#INCLUDE ONCE "Afx/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM cbsName AS CBSTR = pFileSys.GetTempName
```

# <a name="GetVolumeName"></a>GetVolumeName

Returns the volume name of the specified drive.

```
FUNCTION GetVolumeName (BYREF cbsDrive AS CBSTR) AS CBSTR
```

| Name       | Description |
| ---------- | ----------- |
| *cbsDrive* | CBSTR. The drive letter. For drive letters, the root drive is not included. For example, the path for the C drive is C:, not C:\\. |

#### Return value

CBSTR. The volume name.

#### Usage example

```
#INCLUDE ONCE "Afx/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM cbsVolumeName AS CBSTR = pFileSys.GetVolumeName("C:")
```

# <a name="IsDriveReady"></a>IsDriveReady

Returns True if the specified drive is ready; False if it is not.

```
FUNCTION IsDriveReady (BYREF cbsDrive AS CBSTR) AS BOOLEAN
```

| Name       | Description |
| ---------- | ----------- |
| *cbsDrive* | CBSTR. The drive letter. For drive letters, the root drive is not included. For example, the path for the C drive is C:, not C:\\. |

#### Return value

BOOLEAN. True if the specified drive is ready; False if it is not.

#### Remarks

For removable-media drives and CD-ROM drives, **IsDriveReady** returns True only when the appropriate media is inserted and ready for access. 

#### Usage example

```
#INCLUDE ONCE "Afx/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM bIsReady AS BOOLEAN = pFileSys.IsDriveReady("C:")
```

# <a name="IsRootFolder"></a>IsRootFolder

Returns True(-1) if the specified folder is the root folder; False(0) if it is not.

```
FUNCTION IsRootFolder (BYREF cbsFolder AS CBSTR) AS BOOLEAN
```

| Name       | Description |
| ---------- | ----------- |
| *cbsDrive* | CBSTR. The drive letter. For drive letters, the root drive is not included. For example, the path for the C drive is C:, not C:\\. |

#### Return value

BOOLEAN. True(-1) if the specified folder is the root folder; False(0) if it is not.

#### Usage example

```
#INCLUDE ONCE "Afx/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM bIsRoot AS BOOLEAN = pFileSys.IsRootFolder("C:\MyFolder")
```

# <a name="MoveFile"></a>MoveFile

Moves one or more files from one location to another.

```
FUNCTION MoveFile (BYREF cbsSource AS CBSTR, BYREF cbsDestination AS CBSTR) AS HRESULT
```

| Name       | Description |
| ---------- | ----------- |
| *cbsSource* | CBSTR. The path to the file or files to be moved. The *cbsSource* argument string can contain wildcard characters in the last path component only. |
| *cbsDestination* | CBSTR. Destination where the file is to be moved. Wildcard characters are not allowed. |

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

#### Usage example

```
#INCLUDE ONCE "Afx/CFileSys.inc"
DIM pFileSys AS CFileSys
pFileSys.MoveFile("C:\MyFolder\MyFile.txt", "C:\MyOtherFolder\")
```

# <a name="MoveFolder"></a>MoveFolder

Moves one or more folders from one location to another.

```
FUNCTION MoveFolder (BYREF cbsSource AS CBSTR, BYREF cbsDestination AS CBSTR) AS HRESULT
```

| Name       | Description |
| ---------- | ----------- |
| *cbsSource* | CBSTR. The path to the folder or folders to be moved. The *cbsSource* argument string can contain wildcard characters in the last path component only. |
| *cbsDestination* | CBSTR. Destination where the folder is to be moved (must end with a "\\"). Wildcard characters are not allowed. |

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

#### Usage example

```
#INCLUDE ONCE "Afx/CFileSys.inc"
DIM pFileSys AS CFileSys
pFileSys.MoveFolder("C:\MyFolder", "C:\MyNewFolder\")
```

# <a name="SetFileAttributes"></a>SetFileAttributes

Sets the attributes of files.

```
FUNCTION SetFileAttributes (BYREF cbsFile AS CBSTR, BYVAL lAttr AS FILEATTRIBUTE) AS HRESULT
```

| Name       | Description |
| ---------- | ----------- |
| *cbsFile* | CBSTR. The path to a specific file. |
| *lAttr* | LONG. The new value for the attributes of the specified file. |

#### Return value

The file attributes. Can be any of the following values or any logical combination of the following values:

| Constant   | Value       | Description |
| ---------- | ----------- | ----------- |
| FileAttribute_Normal     | 0 | Normal file. No attributes are set. |
| FileAttribute_ReadOnly   | 1 | Read-only file. Attribute is read/write. |
| FileAttribute_Hidden     | 2 | Hidden file. Attribute is read/write. |
| FileAttribute_System     | 4 | System file. Attribute is read/write. |
| FileAttribute_Volume     | 8 | Disk drive volume label. Attribute is read-only. |
| FileAttribute_Directory | 16 | Folder or directory. Attribute is read-only. |
| FileAttribute_Archive |   32 | File has changed since last backup. Attribute is read/write. |
| FileAttribute_Alias   | 1024 | Link or shortcut. Attribute is read-only. |
| FileAttribute_Compressed | 2048 | Compressed file. Attribute is read-only. |

#### Usage example

```
#INCLUDE ONCE "Afx/CFileSys.inc"
DIM pFileSys AS CFileSys
pFileSys.SetFileAttributes("C:\MyPath\MyFile.txt", 33)   ' FileAttribute_Archive OR FileAttribute_Normal
```

# <a name="SetFileName"></a>SetFileName

Sets the name of a specified file.

```
FUNCTION SetFileName (BYREF cbsFile AS CBSTR, BYREF cbsName AS CBSTR) AS HRESULT
```

| Name       | Description |
| ---------- | ----------- |
| *cbsFile* | CBSTR. The path to a specific file. |
| *cbsName* | CBSTR. The new name of the file. |

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

#### Remarks

You only have to pass the new name of the file, not the full path.

#### Usage example

```
#INCLUDE ONCE "Afx/CFileSys.inc"
DIM pFileSys AS CFileSys
pFileSys.SetFileName("c:\MyFolder\Test.txt", "NewName")
```

# <a name="SetFolderAttributes"></a>SetFolderAttributes

Sets the attributes of folders.

```
FUNCTION SetFolderAttributes (BYREF cbsFolder AS CBSTR, BYVAL lAttr AS FILEATTRIBUTE) AS HRESULT
```

| Name       | Description |
| ---------- | ----------- |
| *cbsFolder* | CBSTR. The path to a specific folder. |
| *lAttr* | LONG. The new value for the attributes of the specified folder. |

#### Return value

The *lAttr* argument can have any of the following values or any logical combination of the following values:

| Constant   | Value       | Description |
| ---------- | ----------- | ----------- |
| FileAttribute_Normal     | 0 | Normal file. No attributes are set. |
| FileAttribute_ReadOnly   | 1 | Read-only file. Attribute is read/write. |
| FileAttribute_Hidden     | 2 | Hidden file. Attribute is read/write. |
| FileAttribute_System     | 4 | System file. Attribute is read/write. |
| FileAttribute_Volume     | 8 | Disk drive volume label. Attribute is read-only. |
| FileAttribute_Directory | 16 | Folder or directory. Attribute is read-only. |
| FileAttribute_Archive |   32 | File has changed since last backup. Attribute is read/write. |
| FileAttribute_Alias   | 1024 | Link or shortcut. Attribute is read-only. |
| FileAttribute_Compressed | 2048 | Compressed file. Attribute is read-only. |

#### Usage example

```
#INCLUDE ONCE "Afx/CFileSys.inc"
DIM pFileSys AS CFileSys
pFileSys.SetFolderAttributes("C:\MyPath\MyFile.txt", 17)   ' FileAttribute_Directory OR FileAttribute_ReadOnly
```

# <a name="SetFolderName"></a>SetFolderName

Sets the name of a specified folder.

```
FUNCTION SetFolderName (BYREF cbsFolder AS CBSTR, BYREF cbsName AS CBSTR) AS HRESULT
```

| Name       | Description |
| ---------- | ----------- |
| *cbsFolder* | CBSTR. The path to a specific folder. |
| *cbsName* | CBSTR. The new name of the folder. |

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

#### Remarks

You only have to pass the new name of the folder, not the full path.

#### Usage example

```
#INCLUDE ONCE "Afx/CFileSys.inc"
DIM pFileSys AS CFileSys
pFileSys.SetFolderName("c:\MyFolder", "NewName")
```

# <a name="SetVolumeName"></a>SetVolumeName

Sets the name of a specified drive.

```
FUNCTION SetVolumeName (BYREF cbsDrive AS CBSTR, BYREF cbsName AS CBSTR) AS HRESULT
```

| Name       | Description |
| ---------- | ----------- |
| *cbsDrive* | CBSTR. The drive letter. For drive letters, the root drive is not included. For example, the path for the C drive is C:, not C:\\. |
| *cbsName* | CBSTR. The volume name. |

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

#### Usage example

```
#INCLUDE ONCE "Afx/CFileSys.inc"
DIM pFileSys AS CFileSys
pFileSys.SetVolumeName("C:", "VolumeName")
```
