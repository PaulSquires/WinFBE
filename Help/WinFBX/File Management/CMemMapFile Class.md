# CMemMapFile Class

`CMemMapFile` is a wrapper class that encapsulates the Windows memory-mapped procedures.

A memory-mapped file contains the contents of a file in virtual memory. This mapping between a file and memory space enables an application, including multiple processes, to modify the file by reading and writing directly to the memory.

There are two types of memory-mapped files:

* Persisted memory-mapped files

  Persisted files are memory-mapped files that are associated with a source file on a disk. When the last process has finished working with the file, the data is saved to the source file on the disk. These memory-mapped files are suitable for working with extremely large source files.

* Non-persisted memory-mapped files

  Non-persisted files are memory-mapped files that are not associated with a file on a disk. When the last process has finished working with the file, the data is lost and the file is reclaimed by garbage collection. These files are suitable for creating shared memory for inter-process communications (IPC).

### Processes, Views, and Managing Memory

Memory-mapped files can be shared across multiple processes. Processes can map to the same memory-mapped file by using a common name that is assigned by the process that created the file.

To work with a memory-mapped file, you must create a view of the entire memory-mapped file or a part of it. You can also create multiple views to the same part of the memory-mapped file, thereby creating concurrent memory. For two views to remain concurrent, they have to be created from the same memory-mapped file.

Multiple views may also be necessary if the file is greater than the size of the application’s logical memory space available for memory mapping (2 GB on a 32-bit computer).

Memory-mapped files are accessed through the operating system’s memory manager, so the file is automatically partitioned into a number of pages and accessed as needed. You do not have to handle the memory management yourself.

**Wikipedia article**: [Memory-mapped file](https://en.wikipedia.org/wiki/Memory-mapped_file)

## Methods

| Name       | Description |
| ---------- | ----------- |
| [MapFile](#MapFile) | Maps the specified file. |
| [MapMemory](#MapMemory) | Maps the specified amount of memory. |
| [MapSharedMemory](#MapSharedMemory) | Maps shared memory. |
| [Unmap](#Unmap) | Unmaps the file or memory and closes handles. |
| [AccessData](#AccessData) | Returns a pointer to access the data in the memory mapped file. |
| [UnaccessData](#UnaccessData) | Releases the synchronization object. |
| [Flush](#Flush) | Flushes the data to the disk. |
| [GetFileHandle](#GetFileHandle) | Returns the handle of the underlying disk file. |
| [GetFileMappingHandle](#GetFileMappingHandle) | Returns the file mapping handle. |
| [GetFileSize](#GetFileSize) | Returns the size of the underlying disk file. |

# <a name="MapFile"></a>MapFile

Maps the specified file.

```
FUNCTION CMemMapFile.MapFile (BYVAL pwszFileName AS WSTRING PTR, BYVAL bReadOnly AS BOOLEAN = FALSE, _
   BYVAL dwShareMode AS DWORD = 0, BYVAL bGrowable AS BOOLEAN = FALSE, _
   BYVAL pwszMappingName AS WSTRING PTR = NULL, BYVAL pwszMutexName AS WSTRING PTR = NULL, _
   BYVAL dwStartOffset  AS CONST LONGINT = 0, BYVAL nBytesToMap AS CONST SIZE_T = 0, _
   BYVAL lpSecurityAttributes AS LPSECURITY_ATTRIBUTES = NULL) AS BOOLEAN
```

| Name       | Description |
| ---------- | ----------- |
| *pwszFileName* | The filename to map. |
| *bReadOnly* | Set it to TRUE if you don't intend to modify the file. |
| *dwShareMode* | The sharing mode of the file.<br>**0** : Prevents other processes from opening a file or device if they request delete, read, or write access.<br>**FILE_SHARE_READ** : Enables subsequent open operations on a file to request read access.<br>**FILE_SHARE_WRITE** : Enables subsequent open operations on a file to request write access. |
| *bGrowable* | IF TRUE, the underlying file will be set to be a sparse file and you won't get access violations if you write past the end of the file (Windows will grow the file silently). |
| *pwszMappingName* | The name of the file mapping object. If this parameter is NULL, the file mapping object is created without a name. |
| *pwszMutexName* | The name of the mutex object. The name is limited to MAX_PATH characters. Name comparison is case sensitive. If *pwszMutexName* is NULL, the mutex object is created without a name. |
| *dwStartOffset* | The high-order DWORD is the file offset where the view begins. The low-order DWORD is the file offset where the view is to begin. The combination of the high and low offsets must specify an offset within the file mapping. They must also match the memory allocation granularity of the system. That is, the offset must be a multiple of the allocation granularity. To obtain the memory allocation granularity of the system, use the **GetSystemInfo** function, which fills in the members of a SYSTEM_INFO structure. |
| *nBytesToMap* | The number of bytes of a file mapping to map to the view. All bytes must be within the maximum size specified by **CreateFileMapping**. If this parameter is 0 (zero), the mapping extends from the specified offset to the end of the file mapping. |
| *lpSecurityAttributes* | A pointer to a **SECURITY_ATTRIBUTES** structure that determines whether a returned handle can be inherited by child processes. The *lpSecurityDescriptor* member of the **SECURITY_ATTRIBUTES** structure specifies a security descriptor for a new file mapping object. If *lpAttributes* is NULL, the handle cannot be inherited and the file mapping object gets a default security descriptor. |

#### Return value

TRUE or FALSE.

#### Example

The following code maps the contents of the ansi file "textA.txt", retrieves access to the data and converts it to lower case.

```
#INCLUDE ONCE "windows.bi"
#INCLUDE ONCE "Afx/CMemMapFile.inc"

DIM pMemMap AS CMemMapFile
IF pMemMap.MapFile(AfxGetExePath & "/testA.txt") THEN
   DIM pData AS ANY PTR = pMemMap.AccessData(100)
   IF pData THEN
      CharLowerBuffA(pData, pMemMap.GetFileSize)
      pMemMap.UnaccessData
   END IF
   pMemMap.Unmap
END IF
```

Unicode:

```
#INCLUDE ONCE "windows.bi"
#INCLUDE ONCE "Afx/CMemMapFile.inc"

DIM pMemMap AS CMemMapFile
IF pMemMap.MapFile(AfxGetExePath & "/testW.txt") THEN
   DIM pData AS ANY PTR = pMemMap.AccessData(100)
   IF pData THEN
      CharLowerBuffW(pData, pMemMap.GetFileSize)
      pMemMap.UnaccessData
   END IF
   pMemMap.Unmap
END IF
```

# <a name="MapMemory"></a>MapMemory

Maps the specified amount of memory.

```
FUNCTION CMemMapFile.MapMemory (BYVAL pwszMappingName AS WSTRING PTR, BYVAL pwszMutexName AS WSTRING PTR, _
   BYVAL nBytesToMap AS SIZE_T, BYVAL bReadOnly AS BOOLEAN = FALSE, BYVAL dwStartOffset AS CONST LONGINT = 0, _
   BYVAL lpSecurityAttributes AS LPSECURITY_ATTRIBUTES = NULL) AS BOOLEAN
```

| Name       | Description |
| ---------- | ----------- |
| *pwszMappingName* | The name of the file mapping object. If this parameter is NULL, the file mapping object is created without a name. |
| *pwszMutexName* | The name of the mutex object. The name is limited to MAX_PATH characters. Name comparison is case sensitive. If *pwszMutexName* is NULL, the mutex object is created without a name. |
| *nBytesToMap* | The high-order DWORD is the maximum size of the file mapping object. The low-order DWORD of the maximum size of the file mapping object. If this parameter is 0 (zero), the maximum size of the file mapping object is equal to the current size of the file that hFile identifies. An attempt to map a file with a length of 0 (zero) fails with an error code of ERROR_FILE_INVALID. Applications should test for files with a length of 0 (zero) and reject those files. |
| *bReadOnly* | Set it to TRUE if you don't intend to modify the file. |
| *dwStartOffset* | The high-order DWORD is the file offset where the view begins. The low-order DWORD is the file offset where the view is to begin. The combination of the high and low offsets must specify an offset within the file mapping. They must also match the memory allocation granularity of the system. That is, the offset must be a multiple of the allocation granularity. To obtain the memory allocation granularity of the system, use the **GetSystemInfo** function, which fills in the members of a SYSTEM_INFO structure. |
| *lpSecurityAttributes* | A pointer to a **SECURITY_ATTRIBUTES** structure that determines whether a returned handle can be inherited by child processes. The *lpSecurityDescriptor* member of the **SECURITY_ATTRIBUTES** structure specifies a security descriptor for a new file mapping object. If *lpAttributes* is NULL, the handle cannot be inherited and the file mapping object gets a default security descriptor. |

#### Return value

TRUE or FALSE.

#### Example

The following example maps 256 bytes of memory, using "MyTest" as the name of the mapping object and "MyMutex" as the name of the synchronization object, acceses it with the `AccessData`method (casting the returned pointer to a WSTRING PTR to allow to work with unicode) and changes the content of the fifth character using the *pData* pointer with an index of 4 (indexed pointers are zero based).

```
#define BUF_SIZE 256
DIM pMemMap AS CMemMapFile
IF pMemMap.MapMemory("MyTest", "MyMutex", BUF_SIZE) THEN
   DIM pData AS WSTRING PTR = pMemMap.AccessData
   IF pData THEN
      pData[4] = "x"
      print pData[4]
      pMemMap.UnaccessData
   END IF
   pMemMap.Unmap
END IF
```

# <a name="MapSharedMemory"></a>MapSharedMemory

Maps shared memory.

```
FUNCTION CMemMapFile.MapSharedMemory (BYVAL pwszMappingName AS WSTRING PTR, BYVAL pwszMutexName AS WSTRING PTR, _
   BYVAL nBytesToMap AS SIZE_T, BYVAL bReadOnly AS BOOLEAN = FALSE, _
   BYVAL bInheritHandle AS BOOLEAN = FALSE, BYVAL dwStartOffset AS CONST LONGINT = 0, _
   BYVAL lpSecurityAttributes AS LPSECURITY_ATTRIBUTES = NULL) AS BOOLEAN
```

| Name       | Description |
| ---------- | ----------- |
| *pwszMappingName* | The name of the file mapping object. If this parameter is NULL, the file mapping object is created without a name. |
| *pwszMutexName* | The name of the mutex object. The name is limited to MAX_PATH characters. Name comparison is case sensitive. If pwszMutexName is NULL, the mutex object is created without a name. |
| *nBytesToMap* | The high-order DWORD is the maximum size of the file mapping object. The low-order DWORD of the maximum size of the file mapping object. If this parameter is 0 (zero), the maximum size of the file mapping object is equal to the current size of the file that hFile identifies. An attempt to map a file with a length of 0 (zero) fails with an error code of ERROR_FILE_INVALID. Applications should test for files with a length of 0 (zero) and reject those files. |
| *bReadOnly* | Set it to TRUE if you don't intend to modify the file. |
| *bInheritHandle* | If this parameter is TRUE, a process created by the **CreateProcess** function can inherit the handle; otherwise, the handle cannot be inherited. |
| *dwStartOffset* | The high-order DWORD is the file offset where the view begins. The low-order DWORD is the file offset where the view is to begin. The combination of the high and low offsets must specify an offset within the file mapping. They must also match the memory allocation granularity of the system. That is, the offset must be a multiple of the allocation granularity. To obtain the memory allocation granularity of the system, use the **GetSystemInfo** function, which fills in the members of a SYSTEM_INFO structure. |
| *lpSecurityAttributes* | A pointer to a **SECURITY_ATTRIBUTES** structure that determines whether a returned handle can be inherited by child processes. The *lpSecurityDescriptor* member of the **SECURITY_ATTRIBUTES** structure specifies a security descriptor for a new file mapping object. If *lpAttributes* is NULL, the handle cannot be inherited and the file mapping object gets a default security descriptor. |

#### Return value

TRUE or FALSE.

#### Example

The following example maps 256 bytes of memory, using "MyTest" as the name of the mapping object and "MyMutex" as the name of the synchronization object, acceses it with the `AccessData`method (casting the returned pointer to a WSTRING PTR to allow to work with unicode) and changes the content of the fifth character using the *pData* pointer with an index of 4 (indexed pointers are zero based).

```
#define BUF_SIZE 256
DIM pMemMap AS CMemMapFile
IF pMemMap.MapMemory("MyTest", "MyMutex", BUF_SIZE) THEN
   DIM pData AS WSTRING PTR = pMemMap.AccessData
   IF pData THEN
      pData[4] = "x"
      print pData[4]
      pMemMap.UnaccessData
   END IF
'   pMemMap.Unmap   ' Don't unmap the file if we intend to share the memory
END IF
```

In another process, we can create an instance of `CMemMapFile` using the same parameters and access the shared memory.

```
DIM pSharedMemMap AS CMemMapFile
IF pSharedMemMap.MapSharedMemory("MyTest", "MyMutex", BUF_SIZE, TRUE) THEN
   ' // Access the data
   DIM pData AS WSTRING PTR = pSharedMemMap.AccessData
   IF pData THEN
      print pData[4]
      pSharedMemMap.UnaccessData
   END IF
   pSharedMemMap.Unmap
END IF
```

# <a name="Unmap"></a>Unmap

Unmaps the file or memory and closes handles.

```
SUB Unmap
```

# <a name="AccessData"></a>AccessData

Returns a pointer to access the data in the memory mapped file.

```
FUNCTION AccessData (BYVAL dwTimeout AS DWORD = INFINITE) AS ANY PTR
```

| Name       | Description |
| ---------- | ----------- |
| *dwTimeout* | The time in milliseconds to wait if the mapping is already opened by another thread or process. |

This method uses a mutex to synchronise access to the memory. If you use the default value (INFINITE), your thread will be suspended indefinitely if another thread is using the file.

A successful call to `AccessData` should be matched with a call to `UnaccessData` when you have fnished accessing the data.

# <a name="UnaccessData"></a>UnaccessData

```
FUNCTION UnaccessData () AS BOOLEAN
```

After a successful call to `AccessData` you should call this method when you have fnished accessing the data. After calling this method you can call `AccessData` again to get the pointer to the data.

# <a name="Flush"></a>Flush

Flushes the data to the disk.

```
FUNCTION UnaccessData () AS BOOLEAN
```

#### Remaks

When modifying a file through a mapped view, the last modification timestamp may not be updated automatically. If required, the caller should use SetFileTime to set the timestamp.

# <a name="GetFileHandle"></a>GetFileHandle

Returns the handle of the underlying disk file (it does not appli to objects created with the `MapMemory` or `MapSharedMemory` methods)..

```
FUNCTION GetFileHandle () AS HANDLE
```

# <a name="GetFileMappingHandle"></a>GetFileMappingHandle

Returns the file mapping handle.

```
FUNCTION GetFileMappingHandle () AS HANDLE
```

# <a name="GetFileSize"></a>GetFileSize

Returns the size of the underlying disk file (it does not appli to objects created with the `MapMemory` or `MapSharedMemory` methods).

```
GetFileSize () AS LONGINT
```
