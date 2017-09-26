' ########################################################################################
' Microsoft Windows
' Contents: SQLite3 classes
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2017 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#define unicode
#include once "windows.bi"
#include once "Afx/CWSTR.inc"
#include once "Afx/CDicObj.inc"
using Afx

NAMESPACE Afx

'#define SQLITE_VERSION        "3.20.0"
'#define SQLITE_VERSION_NUMBER 3020000
'#define SQLITE_SOURCE_ID      "2017-08-01 13:24:15 9501e22dfeebdcefa783575e47c60b514d7c2e0cad73b2a496c0bc4b680900a8"

TYPE sqlite_int64 AS LONGINT
TYPE sqlite_uint64 AS ULONGINT
TYPE sqlite3_int64 AS sqlite_int64
TYPE sqlite3_uint64 AS sqlite_uint64
TYPE sqlite3 AS sqlite3_
TYPE sqlite3_stmt AS sqlite3_stmt_
TYPE sqlite3_value as sqlite3_value_
TYPE sqlite3_context AS sqlite3_context_
TYPE sqlite3_vtab AS sqlite3_vtab_
TYPE sqlite3_index_info AS sqlite3_index_info_
TYPE sqlite3_vtab_cursor AS sqlite3_vtab_cursor_
TYPE sqlite3_mutex AS sqlite3_mutex_
TYPE sqlite3_pcache AS sqlite3_pcache_
TYPE sqlite3_backup AS sqlite3_backup_
TYPE sqlite3_rtree_dbl AS double
TYPE sqlite3_rtree_geometry AS sqlite3_rtree_geometry_
TYPE sqlite3_rtree_query_info AS sqlite3_rtree_query_info_
TYPE sqlite3_blob as sqlite3_blob_

TYPE sqlite3_destructor_type AS SUB (BYVAL AS ANY PTR)
CONST SQLITE_STATIC = cast(sqlite3_destructor_type, 0)
CONST SQLITE_TRANSIENT = cast(sqlite3_destructor_type, -1)

TYPE PFNSQLITE3LIBVERSIONPROC AS FUNCTION CDECL () AS CONST ZSTRING PTR
TYPE PFNSQLITE3LIBVERSIONNUMBERPROC AS FUNCTION CDECL () AS LONG
TYPE PFNSQLITE3SOURCEIDPROC AS FUNCTION CDECL () AS CONST ZSTRING PTR
TYPE PFNSQLITE3COMPLETE16PROC AS FUNCTION CDECL (BYVAL sql AS CONST ANY PTR) AS LONG
TYPE PFNSQLITE3THREADSAFEPROC AS FUNCTION CDECL () AS BOOLEAN
TYPE PFNSQLITE3COMPILEOPTIONUSEDPROC AS FUNCTION CDECL (BYVAL szOptName AS CONST ZSTRING PTR) AS BOOLEAN
TYPE PFNSQLITE3COMPILEOPTIONGETPROC AS FUNCTION CDECL (BYVAL N AS LONG) AS CONST ZSTRING PTR
TYPE PFNSQLITE3MALLOCPROC AS FUNCTION CDECL (BYVAL nBytes AS LONG) AS ANY PTR
TYPE PFNSQLITE3MALLOC64PROC AS FUNCTION CDECL (BYVAL nBytes AS sqlite3_uint64) AS ANY PTR
TYPE PFNSQLITE3REALLOCPROC AS FUNCTION CDECL (BYVAL pMem AS ANY PTR, BYVAL nBytes AS LONG) AS ANY PTR
TYPE PFNSQLITE3REALLOC64PROC AS FUNCTION CDECL (BYVAL pMem AS ANY PTR, BYVAL nBytes AS sqlite3_uint64) AS ANY PTR
TYPE PFNSQLITE3FREEPROC AS SUB CDECL (BYVAL pMem AS ANY PTR)
TYPE PFNSQLITE3MSIZEPROC AS FUNCTION CDECL (BYVAL pMem AS ANY PTR) AS sqlite3_uint64
TYPE PFNSQLITE3MEMORYUSEDPROC AS FUNCTION CDECL () AS sqlite3_uint64
TYPE PFNSQLITE3MEMORYHIGHWATERPROC AS FUNCTION CDECL (BYVAL resetFlag AS LONG) AS sqlite3_uint64
TYPE PFNSQLITE3RANDOMNESSPROC AS SUB CDECL (BYVAL nBytes AS LONG, BYVAL pbuffer AS ANY PTR)
TYPE PFNSQLITE3ENABLESHAREDCACHEPROC AS FUNCTION CDECL (BYVAL bSharing AS BOOLEAN) AS LONG
TYPE PFNSQLITE3RELEASEMEMORYPROC AS FUNCTION CDECL (BYVAL nBytes AS LONG) AS LONG
TYPE PFNSQLITE3SLEEPPROC AS FUNCTION CDECL (BYVAL ms AS LONG) AS LONG
TYPE PFNSQLITE3SOFTHEAPLIMIT64PROC AS FUNCTION CDECL (BYVAL nBytes AS LONG) AS LONG
TYPE PFNSQLITE3STATUSPROC AS FUNCTION CDECL (BYVAL op AS LONG, BYVAL pCurrent AS LONG PTR, BYVAL pHighwater AS LONG PTR, BYVAL resetFlag AS LONG) AS LONG
TYPE PFNSQLITE3STATUS64PROC AS FUNCTION CDECL (BYVAL op AS LONG, BYVAL pCurrent AS sqlite3_int64 PTR, BYVAL pHighwater AS sqlite3_int64 PTR, BYVAL resetFlag AS LONG) AS LONG
TYPE PFNSQLITE3ERRSTRPROC AS FUNCTION CDECL (BYVAL nErrorCode AS LONG) AS CONST ZSTRING PTR
TYPE PFNSQLITE3STRGLOBPROC AS FUNCTION CDECL (BYVAL pszGlob AS ZSTRING PTR, BYVAL pszStr AS ZSTRING PTR) AS LONG

' // Result Codes
#define SQLITE_OK           0   ' /* Successful result */
' /* beginning-of-error-codes */
#define SQLITE_ERROR        1   ' /* Generic error */
#define SQLITE_INTERNAL     2   ' /* Internal logic error in SQLite */
#define SQLITE_PERM         3   ' /* Access permission denied */
#define SQLITE_ABORT        4   ' /* Callback routine requested an abort */
#define SQLITE_BUSY         5   ' /* The database file is locked */
#define SQLITE_LOCKED       6   ' /* A table in the database is locked */
#define SQLITE_NOMEM        7   ' /* A malloc() failed */
#define SQLITE_READONLY     8   ' /* Attempt to write a readonly database */
#define SQLITE_INTERRUPT    9   ' /* Operation terminated by sqlite3_interrupt()*/
#define SQLITE_IOERR       10   ' /* Some kind of disk I/O error occurred */
#define SQLITE_CORRUPT     11   ' /* The database disk image is malformed */
#define SQLITE_NOTFOUND    12   ' /* Unknown opcode in sqlite3_file_control() */
#define SQLITE_FULL        13   ' /* Insertion failed because database is full */
#define SQLITE_CANTOPEN    14   ' /* Unable to open the database file */
#define SQLITE_PROTOCOL    15   ' /* Database lock protocol error */
#define SQLITE_EMPTY       16   ' /* Not used */
#define SQLITE_SCHEMA      17   ' /* The database schema changed */
#define SQLITE_TOOBIG      18   ' /* String or BLOB exceeds size limit */
#define SQLITE_CONSTRAINT  19   ' /* Abort due to constraint violation */
#define SQLITE_MISMATCH    20   ' /* Data type mismatch */
#define SQLITE_MISUSE      21   ' /* Library used incorrectly */
#define SQLITE_NOLFS       22   ' /* Uses OS features not supported on host */
#define SQLITE_AUTH        23   ' /* Authorization denied */
#define SQLITE_FORMAT      24   ' /* Not used */
#define SQLITE_RANGE       25   ' /* 2nd parameter to sqlite3_bind out of range */
#define SQLITE_NOTADB      26   ' /* File opened that is not a database file */
#define SQLITE_NOTICE      27   ' /* Notifications from sqlite3_log() */
#define SQLITE_WARNING     28   ' /* Warnings from sqlite3_log() */
#define SQLITE_ROW         100  ' /* sqlite3_step() has another row ready */
#define SQLITE_DONE        101  ' /* sqlite3_step() has finished executing */
' /* end-of-error-codes */

' // Extended Result Codes
#define SQLITE_IOERR_READ              (SQLITE_IOERR OR (1 shl 8))
#define SQLITE_IOERR_SHORT_READ        (SQLITE_IOERR OR (2 shl 8))
#define SQLITE_IOERR_WRITE             (SQLITE_IOERR OR (3 shl 8))
#define SQLITE_IOERR_FSYNC             (SQLITE_IOERR OR (4 shl 8))
#define SQLITE_IOERR_DIR_FSYNC         (SQLITE_IOERR OR (5 shl 8))
#define SQLITE_IOERR_TRUNCATE          (SQLITE_IOERR OR (6 shl 8))
#define SQLITE_IOERR_FSTAT             (SQLITE_IOERR OR (7 shl 8))
#define SQLITE_IOERR_UNLOCK            (SQLITE_IOERR OR (8 shl 8))
#define SQLITE_IOERR_RDLOCK            (SQLITE_IOERR OR (9 shl 8))
#define SQLITE_IOERR_DELETE            (SQLITE_IOERR OR (10 shl 8))
#define SQLITE_IOERR_BLOCKED           (SQLITE_IOERR OR (11 shl 8))
#define SQLITE_IOERR_NOMEM             (SQLITE_IOERR OR (12 shl 8))
#define SQLITE_IOERR_ACCESS            (SQLITE_IOERR OR (13 shl 8))
#define SQLITE_IOERR_CHECKRESERVEDLOCK (SQLITE_IOERR OR (14 shl 8))
#define SQLITE_IOERR_LOCK              (SQLITE_IOERR OR (15 shl 8))
#define SQLITE_IOERR_CLOSE             (SQLITE_IOERR OR (16 shl 8))
#define SQLITE_IOERR_DIR_CLOSE         (SQLITE_IOERR OR (17 shl a8))
#define SQLITE_IOERR_SHMOPEN           (SQLITE_IOERR OR (18 shl 8))
#define SQLITE_IOERR_SHMSIZE           (SQLITE_IOERR OR (19 shl 8))
#define SQLITE_IOERR_SHMLOCK           (SQLITE_IOERR OR (20 shl 8))
#define SQLITE_IOERR_SHMMAP            (SQLITE_IOERR OR (21 shl 8))
#define SQLITE_IOERR_SEEK              (SQLITE_IOERR OR (22 shl 8))
#define SQLITE_IOERR_DELETE_NOENT      (SQLITE_IOERR OR (23 shl 8))
#define SQLITE_IOERR_MMAP              (SQLITE_IOERR OR (24 shl 8))
#define SQLITE_IOERR_GETTEMPPATH       (SQLITE_IOERR OR (25 shl 8))
#define SQLITE_IOERR_CONVPATH          (SQLITE_IOERR OR (26 shl 8))
#define SQLITE_IOERR_VNODE             (SQLITE_IOERR OR (27 shl 8))
#define SQLITE_IOERR_AUTH              (SQLITE_IOERR OR (28 shl 8))
#define SQLITE_LOCKED_SHAREDCACHE      (SQLITE_LOCKED OR  (1 shl 8))
#define SQLITE_BUSY_RECOVERY           (SQLITE_BUSY   OR  (1 shl 8))
#define SQLITE_BUSY_SNAPSHOT           (SQLITE_BUSY   OR  (2 shl 8))
#define SQLITE_CANTOPEN_NOTEMPDIR      (SQLITE_CANTOPEN OR (1 shl 8))
#define SQLITE_CANTOPEN_ISDIR          (SQLITE_CANTOPEN OR (2 shl 8))
#define SQLITE_CANTOPEN_FULLPATH       (SQLITE_CANTOPEN OR (3 shl 8))
#define SQLITE_CANTOPEN_CONVPATH       (SQLITE_CANTOPEN OR (4 shl 8))
#define SQLITE_CORRUPT_VTAB            (SQLITE_CORRUPT OR (1 shl 8))
#define SQLITE_READONLY_RECOVERY       (SQLITE_READONLY OR (1 shl 8))
#define SQLITE_READONLY_CANTLOCK       (SQLITE_READONLY OR (2 shl 8))
#define SQLITE_READONLY_ROLLBACK       (SQLITE_READONLY OR (3 shl 8))
#define SQLITE_READONLY_DBMOVED        (SQLITE_READONLY OR (4 shl 8))
#define SQLITE_ABORT_ROLLBACK          (SQLITE_ABORT OR (2 shl 8))
#define SQLITE_CONSTRAINT_CHECK        (SQLITE_CONSTRAINT OR (1 shl 8))
#define SQLITE_CONSTRAINT_COMMITHOOK   (SQLITE_CONSTRAINT OR (2 shl 8))
#define SQLITE_CONSTRAINT_FOREIGNKEY   (SQLITE_CONSTRAINT OR (3 shl 8))
#define SQLITE_CONSTRAINT_FUNCTION     (SQLITE_CONSTRAINT OR (4 shl 8))
#define SQLITE_CONSTRAINT_NOTNULL      (SQLITE_CONSTRAINT OR (5 shl 8))
#define SQLITE_CONSTRAINT_PRIMARYKEY   (SQLITE_CONSTRAINT OR (6 shl 8))
#define SQLITE_CONSTRAINT_TRIGGER      (SQLITE_CONSTRAINT OR (7 shl 8))
#define SQLITE_CONSTRAINT_UNIQUE       (SQLITE_CONSTRAINT OR (8 shl 8))
#define SQLITE_CONSTRAINT_VTAB         (SQLITE_CONSTRAINT OR (9 shl 8))
#define SQLITE_CONSTRAINT_ROWID        (SQLITE_CONSTRAINT OR (10 shl 8))
#define SQLITE_NOTICE_RECOVER_WAL      (SQLITE_NOTICE OR (1 shl 8))
#define SQLITE_NOTICE_RECOVER_ROLLBACK (SQLITE_NOTICE OR (2 shl 8))
#define SQLITE_WARNING_AUTOINDEX       (SQLITE_WARNING OR (1 shl 8))
#define SQLITE_AUTH_USER               (SQLITE_AUTH OR (1 shl 8))
#define SQLITE_OK_LOAD_PERMANENTLY     (SQLITE_OK OR (1 shl 8))

' // Flags For File Open Operations
#define SQLITE_OPEN_READONLY         &h00000001 ' /* Ok for sqlite3_open_v2() */
#define SQLITE_OPEN_READWRITE        &h00000002 ' /* Ok for sqlite3_open_v2() */
#define SQLITE_OPEN_CREATE           &h00000004 ' /* Ok for sqlite3_open_v2() */
#define SQLITE_OPEN_DELETEONCLOSE    &h00000008 ' /* VFS only */
#define SQLITE_OPEN_EXCLUSIVE        &h00000010 ' /* VFS only */
#define SQLITE_OPEN_AUTOPROXY        &h00000020 ' /* VFS only */
#define SQLITE_OPEN_URI              &h00000040 ' /* Ok for sqlite3_open_v2() */
#define SQLITE_OPEN_MEMORY           &h00000080 ' /* Ok for sqlite3_open_v2() */
#define SQLITE_OPEN_MAIN_DB          &h00000100 ' /* VFS only */
#define SQLITE_OPEN_TEMP_DB          &h00000200 ' /* VFS only */
#define SQLITE_OPEN_TRANSIENT_DB     &h00000400 ' /* VFS only */
#define SQLITE_OPEN_MAIN_JOURNAL     &h00000800 ' /* VFS only */
#define SQLITE_OPEN_TEMP_JOURNAL     &h00001000 ' /* VFS only */
#define SQLITE_OPEN_SUBJOURNAL       &h00002000 ' /* VFS only */
#define SQLITE_OPEN_MASTER_JOURNAL   &h00004000 ' /* VFS only */
#define SQLITE_OPEN_NOMUTEX          &h00008000 ' /* Ok for sqlite3_open_v2() */
#define SQLITE_OPEN_FULLMUTEX        &h00010000 ' /* Ok for sqlite3_open_v2() */
#define SQLITE_OPEN_SHAREDCACHE      &h00020000 ' /* Ok for sqlite3_open_v2() */
#define SQLITE_OPEN_PRIVATECACHE     &h00040000 ' /* Ok for sqlite3_open_v2() */
#define SQLITE_OPEN_WAL              &h00080000 ' /* VFS only */

'/* Reserved:                         &h00F00000 */

' // Device Characteristics
#define SQLITE_IOCAP_ATOMIC                 &h00000001
#define SQLITE_IOCAP_ATOMIC512              &h00000002
#define SQLITE_IOCAP_ATOMIC1K               &h00000004
#define SQLITE_IOCAP_ATOMIC2K               &h00000008
#define SQLITE_IOCAP_ATOMIC4K               &h00000010
#define SQLITE_IOCAP_ATOMIC8K               &h00000020
#define SQLITE_IOCAP_ATOMIC16K              &h00000040
#define SQLITE_IOCAP_ATOMIC32K              &h00000080
#define SQLITE_IOCAP_ATOMIC64K              &h00000100
#define SQLITE_IOCAP_SAFE_APPEND            &h00000200
#define SQLITE_IOCAP_SEQUENTIAL             &h00000400
#define SQLITE_IOCAP_UNDELETABLE_WHEN_OPEN  &h00000800
#define SQLITE_IOCAP_POWERSAFE_OVERWRITE    &h00001000
#define SQLITE_IOCAP_IMMUTABLE              &h00002000

' // File Locking Levels
#define SQLITE_LOCK_NONE          0
#define SQLITE_LOCK_SHARED        1
#define SQLITE_LOCK_RESERVED      2
#define SQLITE_LOCK_PENDING       3
#define SQLITE_LOCK_EXCLUSIVE     4

' // Synchronization Type Flags
#define SQLITE_SYNC_NORMAL        &h00002
#define SQLITE_SYNC_FULL          &h00003
#define SQLITE_SYNC_DATAONLY      &h00010

' // Standard File Control Opcodes
#define SQLITE_FCNTL_LOCKSTATE               1
#define SQLITE_FCNTL_GET_LOCKPROXYFILE       2
#define SQLITE_FCNTL_SET_LOCKPROXYFILE       3
#define SQLITE_FCNTL_LAST_ERRNO              4
#define SQLITE_FCNTL_SIZE_HINT               5
#define SQLITE_FCNTL_CHUNK_SIZE              6
#define SQLITE_FCNTL_FILE_POINTER            7
#define SQLITE_FCNTL_SYNC_OMITTED            8
#define SQLITE_FCNTL_WIN32_AV_RETRY          9
#define SQLITE_FCNTL_PERSIST_WAL            10
#define SQLITE_FCNTL_OVERWRITE              11
#define SQLITE_FCNTL_VFSNAME                12
#define SQLITE_FCNTL_POWERSAFE_OVERWRITE    13
#define SQLITE_FCNTL_PRAGMA                 14
#define SQLITE_FCNTL_BUSYHANDLER            15
#define SQLITE_FCNTL_TEMPFILENAME           16
#define SQLITE_FCNTL_MMAP_SIZE              18
#define SQLITE_FCNTL_TRACE                  19
#define SQLITE_FCNTL_HAS_MOVED              20
#define SQLITE_FCNTL_SYNC                   21
#define SQLITE_FCNTL_COMMIT_PHASETWO        22
#define SQLITE_FCNTL_WIN32_SET_HANDLE       23
#define SQLITE_FCNTL_WAL_BLOCK              24
#define SQLITE_FCNTL_ZIPVFS                 25
#define SQLITE_FCNTL_RBU                    26
#define SQLITE_FCNTL_VFS_POINTER            27
#define SQLITE_FCNTL_JOURNAL_POINTER        28
#define SQLITE_FCNTL_WIN32_GET_HANDLE       29
#define SQLITE_FCNTL_PDB                    30

' /* deprecated names */
#define SQLITE_GET_LOCKPROXYFILE      SQLITE_FCNTL_GET_LOCKPROXYFILE
#define SQLITE_SET_LOCKPROXYFILE      SQLITE_FCNTL_SET_LOCKPROXYFILE
#define SQLITE_LAST_ERRNO             SQLITE_FCNTL_LAST_ERRNO

' // Flags for the xAccess VFS method
#define SQLITE_ACCESS_EXISTS    0
#define SQLITE_ACCESS_READWRITE 1   ' /* Used by PRAGMA temp_store_directory */
#define SQLITE_ACCESS_READ      2   ' /* Unused */

' // Flags for the xShmLock VFS method
#define SQLITE_SHM_UNLOCK       1
#define SQLITE_SHM_LOCK         2
#define SQLITE_SHM_SHARED       4
#define SQLITE_SHM_EXCLUSIVE    8

' // Maximum xShmLock index
#define SQLITE_SHM_NLOCK        8

' // Configuration Options
#define SQLITE_CONFIG_SINGLETHREAD  1  ' /* nil */
#define SQLITE_CONFIG_MULTITHREAD   2  ' /* nil */
#define SQLITE_CONFIG_SERIALIZED    3  ' /* nil */
#define SQLITE_CONFIG_MALLOC        4  ' /* sqlite3_mem_methods* */
#define SQLITE_CONFIG_GETMALLOC     5  ' /* sqlite3_mem_methods* */
#define SQLITE_CONFIG_SCRATCH       6  ' /* void*, int sz, int N */
#define SQLITE_CONFIG_PAGECACHE     7  ' /* void*, int sz, int N */
#define SQLITE_CONFIG_HEAP          8  ' /* void*, int nByte, int min */
#define SQLITE_CONFIG_MEMSTATUS     9  ' /* boolean */
#define SQLITE_CONFIG_MUTEX        10  ' /* sqlite3_mutex_methods* */
#define SQLITE_CONFIG_GETMUTEX     11  ' /* sqlite3_mutex_methods* */
' /* previously SQLITE_CONFIG_CHUNKALLOC 12 which is now unused. */
#define SQLITE_CONFIG_LOOKASIDE    13  ' /* int int */
#define SQLITE_CONFIG_PCACHE       14  ' /* no-op */
#define SQLITE_CONFIG_GETPCACHE    15  ' /* no-op */
#define SQLITE_CONFIG_LOG          16  ' /* xFunc, void* */
#define SQLITE_CONFIG_URI          17  ' /* int */
#define SQLITE_CONFIG_PCACHE2      18  ' /* sqlite3_pcache_methods2* */
#define SQLITE_CONFIG_GETPCACHE2   19  ' /* sqlite3_pcache_methods2* */
#define SQLITE_CONFIG_COVERING_INDEX_SCAN 20  ' /* int */
#define SQLITE_CONFIG_SQLLOG       21  ' /* xSqllog, void* */
#define SQLITE_CONFIG_MMAP_SIZE    22  ' /* sqlite3_int64, sqlite3_int64 */
#define SQLITE_CONFIG_WIN32_HEAPSIZE      23  ' /* int nByte */
#define SQLITE_CONFIG_PCACHE_HDRSZ        24  ' /* int *psz */
#define SQLITE_CONFIG_PMASZ               25  ' /* unsigned int szPma */
#define SQLITE_CONFIG_STMTJRNL_SPILL      26  ' /* int nByte */

'// Database Connection Configuration Options
#define SQLITE_DBCONFIG_MAINDBNAME            1000 ' /* const char* */
#define SQLITE_DBCONFIG_LOOKASIDE             1001 ' /* void* int int */
#define SQLITE_DBCONFIG_ENABLE_FKEY           1002 ' /* int int* */
#define SQLITE_DBCONFIG_ENABLE_TRIGGER        1003 ' /* int int* */
#define SQLITE_DBCONFIG_ENABLE_FTS3_TOKENIZER 1004 ' /* int int* */
#define SQLITE_DBCONFIG_ENABLE_LOAD_EXTENSION 1005 ' /* int int* */
#define SQLITE_DBCONFIG_NO_CKPT_ON_CLOSE      1006 ' /* int int* */
#define SQLITE_DBCONFIG_ENABLE_QPSG           1007 ' /* int int* */

' // Authorizer Return Codes
#define SQLITE_DENY   1   ' /* Abort the SQL statement with an error */
#define SQLITE_IGNORE 2   ' /* Don't allow access, but don't generate an error */

' // Authorizer Action Codes
#define SQLITE_CREATE_INDEX          1   ' /* Index Name      Table Name      */
#define SQLITE_CREATE_TABLE          2   ' /* Table Name      NULL            */
#define SQLITE_CREATE_TEMP_INDEX     3   ' /* Index Name      Table Name      */
#define SQLITE_CREATE_TEMP_TABLE     4   ' /* Table Name      NULL            */
#define SQLITE_CREATE_TEMP_TRIGGER   5   ' /* Trigger Name    Table Name      */
#define SQLITE_CREATE_TEMP_VIEW      6   ' /* View Name       NULL            */
#define SQLITE_CREATE_TRIGGER        7   ' /* Trigger Name    Table Name      */
#define SQLITE_CREATE_VIEW           8   ' /* View Name       NULL            */
#define SQLITE_DELETE                9   ' /* Table Name      NULL            */
#define SQLITE_DROP_INDEX           10   ' /* Index Name      Table Name      */
#define SQLITE_DROP_TABLE           11   ' /* Table Name      NULL            */
#define SQLITE_DROP_TEMP_INDEX      12   ' /* Index Name      Table Name      */
#define SQLITE_DROP_TEMP_TABLE      13   ' /* Table Name      NULL            */
#define SQLITE_DROP_TEMP_TRIGGER    14   ' /* Trigger Name    Table Name      */
#define SQLITE_DROP_TEMP_VIEW       15   ' /* View Name       NULL            */
#define SQLITE_DROP_TRIGGER         16   ' /* Trigger Name    Table Name      */
#define SQLITE_DROP_VIEW            17   ' /* View Name       NULL            */
#define SQLITE_INSERT               18   ' /* Table Name      NULL            */
#define SQLITE_PRAGMA               19   ' /* Pragma Name     1st arg or NULL */
#define SQLITE_READ                 20   ' /* Table Name      Column Name     */
#define SQLITE_SELECT               21   ' /* NULL            NULL            */
#define SQLITE_TRANSACTION          22   ' /* Operation       NULL            */
#define SQLITE_UPDATE               23   ' /* Table Name      Column Name     */
#define SQLITE_ATTACH               24   ' /* Filename        NULL            */
#define SQLITE_DETACH               25   ' /* Database Name   NULL            */
#define SQLITE_ALTER_TABLE          26   ' /* Database Name   Table Name      */
#define SQLITE_REINDEX              27   ' /* Index Name      NULL            */
#define SQLITE_ANALYZE              28   ' /* Table Name      NULL            */
#define SQLITE_CREATE_VTABLE        29   ' /* Table Name      Module Name     */
#define SQLITE_DROP_VTABLE          30   ' /* Table Name      Module Name     */
#define SQLITE_FUNCTION             31   ' /* NULL            Function Name   */
#define SQLITE_SAVEPOINT            32   ' /* Operation       Savepoint Name  */
#define SQLITE_COPY                  0   ' /* No longer used */
#define SQLITE_RECURSIVE            33   ' /* NULL            NULL            */

' // Run-Time Limit Categories
#define SQLITE_LIMIT_LENGTH                    0
#define SQLITE_LIMIT_SQL_LENGTH                1
#define SQLITE_LIMIT_COLUMN                    2
#define SQLITE_LIMIT_EXPR_DEPTH                3
#define SQLITE_LIMIT_COMPOUND_SELECT           4
#define SQLITE_LIMIT_VDBE_OP                   5
#define SQLITE_LIMIT_FUNCTION_ARG              6
#define SQLITE_LIMIT_ATTACHED                  7
#define SQLITE_LIMIT_LIKE_PATTERN_LENGTH       8
#define SQLITE_LIMIT_VARIABLE_NUMBER           9
#define SQLITE_LIMIT_TRIGGER_DEPTH            10
#define SQLITE_LIMIT_WORKER_THREADS           11

' // Prepare Flags
#define SQLITE_PREPARE_PERSISTENT              &h01

' // Fundamental Datatypes
#define SQLITE_INTEGER  1
#define SQLITE_FLOAT    2
#define SQLITE_BLOB     4
#define SQLITE_NULL     5
#define SQLITE_TEXT     3
#define SQLITE3_TEXT    3

' // Text Encodings
#define SQLITE_UTF8           1    ' /* IMP: R-37514-35566 */
#define SQLITE_UTF16LE        2    ' /* IMP: R-03371-37637 */
#define SQLITE_UTF16BE        3    ' /* IMP: R-51971-34154 */
#define SQLITE_UTF16          4    ' /* Use native byte order */
#define SQLITE_ANY            5    ' /* Deprecated */
#define SQLITE_UTF16_ALIGNED  8    ' /* sqlite3_create_collation only */

' // Function Flags
#define SQLITE_DETERMINISTIC    &h800

' // Mutex Types
#define SQLITE_MUTEX_FAST             0
#define SQLITE_MUTEX_RECURSIVE        1
#define SQLITE_MUTEX_STATIC_MASTER    2
#define SQLITE_MUTEX_STATIC_MEM       3  ' /* sqlite3_malloc() */
#define SQLITE_MUTEX_STATIC_MEM2      4  ' /* NOT USED */
#define SQLITE_MUTEX_STATIC_OPEN      4  ' /* sqlite3BtreeOpen() */
#define SQLITE_MUTEX_STATIC_PRNG      5  ' /* sqlite3_randomness() */
#define SQLITE_MUTEX_STATIC_LRU       6  ' /* lru page list */
#define SQLITE_MUTEX_STATIC_LRU2      7  ' /* NOT USED */
#define SQLITE_MUTEX_STATIC_PMEM      7  ' /* sqlite3PageMalloc() */
#define SQLITE_MUTEX_STATIC_APP1      8  ' /* For use by application */
#define SQLITE_MUTEX_STATIC_APP2      9  ' /* For use by application */
#define SQLITE_MUTEX_STATIC_APP3     10  ' /* For use by application */
#define SQLITE_MUTEX_STATIC_VFS1     11  ' /* For use by built-in VFS */
#define SQLITE_MUTEX_STATIC_VFS2     12  ' /* For use by extension VFS */
#define SQLITE_MUTEX_STATIC_VFS3     13  ' /* For use by application VFS */

' // Testing Interface Operation Codes
#define SQLITE_TESTCTRL_FIRST                    5
#define SQLITE_TESTCTRL_PRNG_SAVE                5
#define SQLITE_TESTCTRL_PRNG_RESTORE             6
#define SQLITE_TESTCTRL_PRNG_RESET               7
#define SQLITE_TESTCTRL_BITVEC_TEST              8
#define SQLITE_TESTCTRL_FAULT_INSTALL            9
#define SQLITE_TESTCTRL_BENIGN_MALLOC_HOOKS     10
#define SQLITE_TESTCTRL_PENDING_BYTE            11
#define SQLITE_TESTCTRL_ASSERT                  12
#define SQLITE_TESTCTRL_ALWAYS                  13
#define SQLITE_TESTCTRL_RESERVE                 14
#define SQLITE_TESTCTRL_OPTIMIZATIONS           15
#define SQLITE_TESTCTRL_ISKEYWORD               16
#define SQLITE_TESTCTRL_SCRATCHMALLOC           17
#define SQLITE_TESTCTRL_LOCALTIME_FAULT         18
#define SQLITE_TESTCTRL_EXPLAIN_STMT            19  ' /* NOT USED */
#define SQLITE_TESTCTRL_ONCE_RESET_THRESHOLD    19
#define SQLITE_TESTCTRL_NEVER_CORRUPT           20
#define SQLITE_TESTCTRL_VDBE_COVERAGE           21
#define SQLITE_TESTCTRL_BYTEORDER               22
#define SQLITE_TESTCTRL_ISINIT                  23
#define SQLITE_TESTCTRL_SORTER_MMAP             24
#define SQLITE_TESTCTRL_IMPOSTER                25
#define SQLITE_TESTCTRL_LAST                    25

' // Status Parameters
#define SQLITE_STATUS_MEMORY_USED          0
#define SQLITE_STATUS_PAGECACHE_USED       1
#define SQLITE_STATUS_PAGECACHE_OVERFLOW   2
#define SQLITE_STATUS_SCRATCH_USED         3
#define SQLITE_STATUS_SCRATCH_OVERFLOW     4
#define SQLITE_STATUS_MALLOC_SIZE          5
#define SQLITE_STATUS_PARSER_STACK         6
#define SQLITE_STATUS_PAGECACHE_SIZE       7
#define SQLITE_STATUS_SCRATCH_SIZE         8
#define SQLITE_STATUS_MALLOC_COUNT         9

' // Status Parameters for database connections
#define SQLITE_DBSTATUS_LOOKASIDE_USED       0
#define SQLITE_DBSTATUS_CACHE_USED           1
#define SQLITE_DBSTATUS_SCHEMA_USED          2
#define SQLITE_DBSTATUS_STMT_USED            3
#define SQLITE_DBSTATUS_LOOKASIDE_HIT        4
#define SQLITE_DBSTATUS_LOOKASIDE_MISS_SIZE  5
#define SQLITE_DBSTATUS_LOOKASIDE_MISS_FULL  6
#define SQLITE_DBSTATUS_CACHE_HIT            7
#define SQLITE_DBSTATUS_CACHE_MISS           8
#define SQLITE_DBSTATUS_CACHE_WRITE          9
#define SQLITE_DBSTATUS_DEFERRED_FKS        10
#define SQLITE_DBSTATUS_CACHE_USED_SHARED   11
#define SQLITE_DBSTATUS_MAX                 11   ' /* Largest defined DBSTATUS */

' // Status Parameters for prepared statements
#define SQLITE_STMTSTATUS_FULLSCAN_STEP     1
#define SQLITE_STMTSTATUS_SORT              2
#define SQLITE_STMTSTATUS_AUTOINDEX         3
#define SQLITE_STMTSTATUS_VM_STEP           4
#define SQLITE_STMTSTATUS_REPREPARE         5
#define SQLITE_STMTSTATUS_RUN               6
#define SQLITE_STMTSTATUS_MEMUSED           99

' // Checkpoint Mode Values
#define SQLITE_CHECKPOINT_PASSIVE  0  ' /* Do as much as possible w/o blocking */
#define SQLITE_CHECKPOINT_FULL     1  ' /* Wait for writers, then checkpoint */
#define SQLITE_CHECKPOINT_RESTART  2  ' /* Like FULL but wait for for readers */
#define SQLITE_CHECKPOINT_TRUNCATE 3  ' /* Like RESTART but also truncate WAL */

' // Virtual Table Configuration Options
#define SQLITE_VTAB_CONSTRAINT_SUPPORT 1

' // Conflict resolution modes
#define SQLITE_ROLLBACK 1
'/* #define SQLITE_IGNORE 2 // Also used by sqlite3_authorizer() callback */
#define SQLITE_FAIL     3
'/* #define SQLITE_ABORT 4  // Also an error code */
#define SQLITE_REPLACE  5

' // Prepared Statement Scan Status Opcodes
#define SQLITE_SCANSTAT_NLOOP    0
#define SQLITE_SCANSTAT_NVISIT   1
#define SQLITE_SCANSTAT_EST      2
#define SQLITE_SCANSTAT_NAME     3
#define SQLITE_SCANSTAT_EXPLAIN  4
#define SQLITE_SCANSTAT_SELECTID 5

' ########################################################################################
' CSQLite base class
' ########################################################################################
TYPE CSQLite

Public:
   DIM m_Result AS LONG
   DIM m_hLib AS HMODULE

Public:
   DECLARE CONSTRUCTOR (BYREF wszDllPath AS WSTRING = "sqlite3.dll")
   DECLARE DESTRUCTOR
   DECLARE FUNCTION RefCount (BYVAL bIncrement AS BOOLEAN, BYREF wszDllPath AS WSTRING = "sqlite3.dll") AS LONG
   DECLARE FUNCTION GetLastResult () AS HRESULT
   DECLARE FUNCTION SetResult (BYVAL Result AS HRESULT) AS HRESULT
   DECLARE FUNCTION Version () AS STRING
   DECLARE FUNCTION VersionNumber () AS LONG
   DECLARE FUNCTION SourceID () AS STRING
   DECLARE FUNCTION Complete (BYREF wszSql AS CONST WSTRING) AS LONG
   DECLARE FUNCTION ThreadSafe () AS BOOLEAN
   DECLARE FUNCTION CompileOptionUsed (BYREF szOptName AS ZSTRING) AS BOOLEAN
   DECLARE FUNCTION GetCompileOption (BYVAL nOption AS LONG) AS STRING
   DECLARE FUNCTION Malloc (BYVAL nBytes AS LONG) AS ANY PTR
   DECLARE FUNCTION Malloc64 (BYVAL nBytes AS sqlite3_uint64) AS ANY PTR
   DECLARE FUNCTION Realloc (BYVAL pMem AS ANY PTR, BYVAL nBytes AS LONG) AS ANY PTR
   DECLARE FUNCTION Realloc64 (BYVAL pMem AS ANY PTR, BYVAL nBytes AS sqlite3_uint64) AS ANY PTR
   DECLARE SUB Free (BYVAL pMem AS ANY PTR)
   DECLARE FUNCTION MemorySize (BYVAL pMem AS ANY PTR) AS sqlite3_uint64
   DECLARE FUNCTION MemoryUsed () AS sqlite3_uint64
   DECLARE FUNCTION MemoryHighwater (BYVAL resetFlag AS BOOLEAN) AS sqlite3_uint64
   DECLARE SUB Randomness (BYVAL nBytes AS LONG, BYVAL pbuffer AS ANY PTR)
   DECLARE FUNCTION EnableSharedCache (BYVAL bSharing AS BOOLEAN) AS LONG
   DECLARE FUNCTION ReleaseMemory (BYVAL nBytes AS LONG) AS LONG
   DECLARE FUNCTION Sleep (BYVAL ms AS LONG) AS LONG
   DECLARE FUNCTION SoftHeapLimit64 (BYVAL nBytes AS LONG) AS LONG
   DECLARE FUNCTION Status (BYVAL op AS LONG, BYREF pCurrent AS LONG, BYREF pHighwater AS LONG, BYVAL resetFlag AS BOOLEAN = FALSE) AS LONG
   DECLARE FUNCTION Status64 (BYVAL op AS LONG, BYREF pCurrent AS sqlite3_int64, BYREF pHighwater AS sqlite3_int64, BYVAL resetFlag AS BOOLEAN = FALSE) AS LONG
   DECLARE FUNCTION ErrStr (BYVAL nErrorCode AS LONG) AS STRING
   DECLARE FUNCTION StrGlob (BYREF szGlob AS ZSTRING, BYREF szStr AS ZSTRING) AS LONG

END TYPE
' ########################################################################################

' ========================================================================================
' Constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CSQLite (BYREF wszDllPath AS WSTRING = "sqlite3.dll")
   ' // Increase the reference count
   this.RefCount(TRUE, wszDllPath)
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Increases/decreases the reference count and frees the library when required.
' ========================================================================================
PRIVATE FUNCTION CSQLite.RefCount (BYVAL bIncrement AS BOOLEAN, BYREF wszDllPath AS WSTRING = "sqlite3.dll") AS LONG
   STATIC _hLib AS HMODULE, cRef AS LONG
   IF cRef = 0 THEN
      ' // Load the SQLite library
      _hLib = LoadLibraryW(wszDllPath)
      IF _hLib = NULL THEN MessageBoxW(GetActiveWindow, "Unable to load " & wszDllPath, "Error", MB_APPLMODAL OR MB_ICONERROR)
   END IF
   ' // Store the handle in the public variable of the class
   m_hLib = _hLib
   ' // Increase or decrease the reference count
   IF bIncrement THEN cRef += 1 ELSE cRef -= 1
   ' // If the reference count reaches a value of 0, free the library
   IF cRef = 0  THEN
      ' // Free the SQLite library
      IF _hLib THEN FreeLibrary _hLib
   END IF
   RETURN cRef
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Destructor
' ===========================================================================================
PRIVATE DESTRUCTOR CSQLite
   ' // Decrease the reference count
   this.RefCount(FALSE)
END DESTRUCTOR
' ===========================================================================================

' ========================================================================================
' Returns the last result
' ========================================================================================
PRIVATE FUNCTION CSQLite.GetLastResult () AS HRESULT
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the last status code.
' ========================================================================================
PRIVATE FUNCTION CSQLite.SetResult (BYVAL Result AS HRESULT) AS HRESULT
   m_Result = Result
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the SQLite version
' ========================================================================================
PRIVATE FUNCTION CSQLite.Version () AS STRING
   DIM pProc AS PFNSQLITE3LIBVERSIONPROC = _
   cast(PFNSQLITE3LIBVERSIONPROC, GetProcAddress(m_hLib, "sqlite3_libversion"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   DIM psz AS CONST ZSTRING PTR = pProc()
   IF psz THEN FUNCTION = *psz
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the SQLite version number
' ========================================================================================
PRIVATE FUNCTION CSQLite.VersionNumber () AS LONG
   DIM pProc AS PFNSQLITE3LIBVERSIONNUMBERPROC = _
   cast(PFNSQLITE3LIBVERSIONNUMBERPROC, GetProcAddress(m_hLib, "sqlite3_libversion_number"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   FUNCTION = pProc()
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the SQLite source identifier
' ========================================================================================
PRIVATE FUNCTION CSQLite.SourceID () AS STRING
   DIM pProc AS PFNSQLITE3SOURCEIDPROC = _
   cast(PFNSQLITE3SOURCEIDPROC, GetProcAddress(m_hLib, "sqlite3_sourceid"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   DIM psz AS CONST ZSTRING PTR = pProc()
   IF psz THEN FUNCTION = *psz
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines if the currently entered text seems to form a complete SQL statement or if
' additional input is needed before sending the text into SQLite for parsing.
' Returns 0 if the statement is incomplete or 1 if the input string appears to be a
' complete SQL statement. If a memory allocation fails, then SQLITE_NOMEM is returned.
' It does not parse the SQL statements thus will not detect syntactically incorrect SQL.
' ========================================================================================
PRIVATE FUNCTION CSQLite.Complete (BYREF wszSql AS CONST WSTRING) AS LONG
   DIM pProc AS PFNSQLITE3COMPLETE16PROC = _
   cast(PFNSQLITE3COMPLETE16PROC, GetProcAddress(m_hLib, "sqlite3_complete16"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   FUNCTION = pProc(@wszSql)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns FALSE if and only if SQLite was compiled with mutexing code omitted.
' ========================================================================================
PRIVATE FUNCTION CSQLite.ThreadSafe () AS BOOLEAN
   DIM pProc AS PFNSQLITE3THREADSAFEPROC = _
   cast(PFNSQLITE3THREADSAFEPROC, GetProcAddress(m_hLib, "sqlite3_threadsafe"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   FUNCTION = pProc()
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns FALSE OR TRUE indicating whether the specified option was defined at compile time.
' The SQLITE_ prefix may be omitted from the option name passed to CompileOptionUsed.
' Example: pSql.CompileOptionUsed(("SQLITE_ENABLE_DBSTAT_VTAB").
' ========================================================================================
PRIVATE FUNCTION CSQLite.CompileOptionUsed (BYREF szOptName AS ZSTRING) AS BOOLEAN
   DIM pProc AS PFNSQLITE3COMPILEOPTIONUSEDPROC = _
   cast(PFNSQLITE3COMPILEOPTIONUSEDPROC, GetProcAddress(m_hLib, "sqlite3_compileoption_used"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   FUNCTION = pProc(@szOptName)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Allows iterating over the list of options that were defined at compile time by returning
' the N-th compile time option string. If N is out of range, GetCompileOption() returns a
' NULL pointer. The SQLITE_ prefix is omitted from any strings returned by GetCompileOption.
' Example: pSql.GetCompileOption(0)
' ========================================================================================
PRIVATE FUNCTION CSQLite.GetCompileOption (BYVAL nOption AS LONG) AS STRING
   DIM pProc AS PFNSQLITE3COMPILEOPTIONGETPROC = _
   cast(PFNSQLITE3COMPILEOPTIONGETPROC, GetProcAddress(m_hLib, "sqlite3_compileoption_get"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   DIM psz AS CONST ZSTRING PTR = pProc(nOption)
   IF psz THEN FUNCTION = *psz
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a pointer to a block of memory at least nBytes bytes in length.
' If Malloc() is unable to obtain sufficient free memory, it returns a NULL pointer. If the
' parameter nBytes to Malloc() is zero or negative then Malloc() returns a NULL pointer.
' ========================================================================================
PRIVATE FUNCTION CSQLite.Malloc (BYVAL nBytes AS LONG) AS ANY PTR
   DIM pProc AS PFNSQLITE3MALLOCPROC = _
   cast(PFNSQLITE3MALLOCPROC, GetProcAddress(m_hLib, "sqlite3_malloc"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   FUNCTION = pProc(nBytes)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Same as Malloc except that the nBytes parameter is an unsigned 64 bit integer instead of
' a signed 32-bit integer.
' ========================================================================================
PRIVATE FUNCTION CSQLite.Malloc64 (BYVAL nBytes AS sqlite3_uint64) AS ANY PTR
   DIM pProc AS PFNSQLITE3MALLOC64PROC = _
   cast(PFNSQLITE3MALLOC64PROC, GetProcAddress(m_hLib, "sqlite3_malloc64"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   FUNCTION = pProc(nBytes)
END FUNCTION
' ========================================================================================

' ========================================================================================
' ttempts to resize a prior memory allocation pMem to be at least nBytes bytes.
' If the pMem parameter to Realloc is a NULL pointer then its behavior is identical to
' calling Malloc(nBytes).
' ========================================================================================
PRIVATE FUNCTION CSQLite.Realloc (BYVAL pMem AS ANY PTR, BYVAL nBytes AS LONG) AS ANY PTR
   DIM pProc AS PFNSQLITE3REALLOCPROC = _
   cast(PFNSQLITE3REALLOCPROC, GetProcAddress(m_hLib, "sqlite3_realloc"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   FUNCTION = pProc(pMem, nBytes)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Same as Realloc except that the nBytes parameter is an unsigned 64 bit integer instead of
' a signed 32-bit integer.
' ========================================================================================
PRIVATE FUNCTION CSQLite.Realloc64 (BYVAL pMem AS ANY PTR, BYVAL nBytes AS sqlite3_uint64) AS ANY PTR
   DIM pProc AS PFNSQLITE3REALLOCPROC = _
   cast(PFNSQLITE3REALLOCPROC, GetProcAddress(m_hLib, "sqlite3_realloc64"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   FUNCTION = pProc(pMem, nBytes)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Calling Free with a pointer previously returned by Malloc or Realloc releases that memory
' so that it might be reused. The Free method is a no-op if is called with a NULL pointer.
' Passing a NULL pointer to Free is harmless.
' ========================================================================================
PRIVATE SUB CSQLite.Free (BYVAL pMem AS ANY PTR)
   DIM pProc AS PFNSQLITE3FREEPROC = _
   cast(PFNSQLITE3FREEPROC, GetProcAddress(m_hLib, "sqlite3_free"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT SUB
   pProc(pMem)
END SUB
' ========================================================================================

' ========================================================================================
' If pMem is a memory allocation previously obtained from Malloc, Malloc64, Realloc, or
' Realloc64(), then MemSize returns the size of that memory allocation in bytes. The value
' returned by MemSize might be larger than the number of bytes requested when pMem was
' allocated. If pMem is a NULL pointer then MemSize returns zero. If pMem points to
' something that is not the beginning of memory allocation, or if it points to a formerly
' valid memory allocation that has now been freed, then the behavior of MemSize is undefined
' and possibly harmful.
' ========================================================================================
PRIVATE FUNCTION CSQLite.MemorySize (BYVAL pMem AS ANY PTR) AS sqlite3_uint64
   DIM pProc AS PFNSQLITE3MSIZEPROC = _
   cast(PFNSQLITE3MSIZEPROC, GetProcAddress(m_hLib, "sqlite3_msize"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   FUNCTION = pProc(pMem)
END FUNCTION
' ========================================================================================

' ========================================================================================
' The MemoryUsed method returns the number of bytes of memory currently outstanding
' (malloced but not freed).
' ========================================================================================
PRIVATE FUNCTION CSQLite.MemoryUsed () AS sqlite3_uint64
   DIM pProc AS PFNSQLITE3MEMORYUSEDPROC = _
   cast(PFNSQLITE3MEMORYUSEDPROC, GetProcAddress(m_hLib, "sqlite3_memory_used"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   FUNCTION = pProc()
END FUNCTION
' ========================================================================================

' ========================================================================================
' The MemoryHighwater method returns the maximum value of MemoryUsed since the high-water
' mark was last reset. The memory high-water mark is reset to the current value of
' MemoryUsed if and only if the parameter to MemoryHighwater is true. The value returned
' by MemoryHighwater(TRUE) is the high-water mark prior to the reset.
' ========================================================================================
PRIVATE FUNCTION CSQLite.MemoryHighwater (BYVAL resetFlag AS BOOLEAN) AS sqlite3_uint64
   DIM pProc AS PFNSQLITE3MEMORYHIGHWATERPROC = _
   cast(PFNSQLITE3MEMORYHIGHWATERPROC, GetProcAddress(m_hLib, "sqlite3_memory_highwater"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   DIM _resetFlag AS LONG = IIF(resetFlag = 0, 0, 1)
   FUNCTION = pProc(_resetFlag)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Stores nBytes bytes of randomness into buffer pbuffer. The pbuffer parameter can be a
' NULL pointer.
' ========================================================================================
PRIVATE SUB CSQLite.Randomness (BYVAL nBytes AS LONG, BYVAL pbuffer AS ANY PTR)
   DIM pProc AS PFNSQLITE3RANDOMNESSPROC = _
   cast(PFNSQLITE3RANDOMNESSPROC, GetProcAddress(m_hLib, "sqlite3_randomness"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT SUB
   pProc(nBytes, pbuffer)
END SUB
' ========================================================================================

' ========================================================================================
' This method enables or disables the sharing of the database cache and schema data
' structures between connections to the same database. Sharing is enabled if the argument
' is true and disabled if the argument is false. Returns SQLITE_OK if shared cache was
' enabled or disabled successfully. An error code is returned otherwise.
' ========================================================================================
PRIVATE FUNCTION CSQLite.EnableSharedCache (BYVAL bSharing AS BOOLEAN) AS LONG
   DIM pProc AS PFNSQLITE3ENABLESHAREDCACHEPROC = _
   cast(PFNSQLITE3ENABLESHAREDCACHEPROC, GetProcAddress(m_hLib, "sqlite3_enable_shared_cache"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   DIM _bSharing AS LONG = IIF(bSharing = 0, 0, 1)
   FUNCTION = pProc(_bSharing)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Attempts to free the specified number of bytes of heap memory by deallocating non-essential
' memory allocations held by the database library. Memory used to cache database pages
' to improve performance is an example of non-essential memory. ReleaseMemory returns
' the number of bytes actually freed, which might be more or less than the amount requested.
' The ReleaseMemory method is a no-op returning zero if SQLite is not compiled with
' SQLITE_ENABLE_MEMORY_MANAGEMENT.
' ========================================================================================
PRIVATE FUNCTION CSQLite.ReleaseMemory (BYVAL nBytes AS LONG) AS LONG
   DIM pProc AS PFNSQLITE3RELEASEMEMORYPROC = _
   cast(PFNSQLITE3RELEASEMEMORYPROC, GetProcAddress(m_hLib, "sqlite3_release_memory"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   FUNCTION = pProc(nBytes)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Causes the current thread to suspend execution for at least a number of milliseconds
' specified in its parameter. If the operating system does not support sleep requests with
' millisecond time resolution, then the time will be rounded up to the nearest second. The
' number of milliseconds of sleep actually requested from the operating system is returned.
' ========================================================================================
PRIVATE FUNCTION CSQLite.Sleep (BYVAL ms AS LONG) AS LONG
   DIM pProc AS PFNSQLITE3SLEEPPROC = _
   cast(PFNSQLITE3SLEEPPROC, GetProcAddress(m_hLib, "sqlite3_sleep"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   FUNCTION = pProc(ms)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets and/or queries the soft limit on the amount of heap memory that may be allocated by
' SQLite. The return value from SoftHeapLimit64 is the size of the soft heap limit prior
' to the call, or negative in the case of an error. If the argument N is negative then no
' change is made to the soft heap limit. Hence, the current size of the soft heap limit
' can be determined by invoking SoftHeapLimit64 with a negative argument.
' ========================================================================================
PRIVATE FUNCTION CSQLite.SoftHeapLimit64 (BYVAL nBytes AS LONG) AS LONG
   DIM pProc AS PFNSQLITE3SOFTHEAPLIMIT64PROC = _
   cast(PFNSQLITE3SOFTHEAPLIMIT64PROC, GetProcAddress(m_hLib, "sqlite3_soft_heap_limit64"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   FUNCTION = pProc(nBytes)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves runtime status information about the performance of SQLite, and optionally
' to reset various highwater marks. The first argument is an integer code for the specific
' parameter to measure. Recognized integer codes are of the form SQLITE_STATUS_.... The
' current value of the parameter is returned into pCurrent. The highest recorded value is
' returned in pHighwater. If the resetFlag is true, then the highest record value is reset
' after pHighwater is written. Some parameters do not record the highest value. For those
' parameters nothing is written into pHighwater and the resetFlag is ignored. Other
' parameters record only the highwater mark and not the current value. For these latter
' parameters nothing is written into pCurrent.
' The SQLiteStatus and SQLiteStatus64 methods return SQLITE_OK on success and a non-zero
' error code on failure.
' ========================================================================================
PRIVATE FUNCTION CSQLite.Status (BYVAL op AS LONG, BYREF pCurrent AS LONG, BYREF pHighwater AS LONG, BYVAL resetFlag AS BOOLEAN = FALSE) AS LONG
   DIM pProc AS PFNSQLITE3STATUSPROC = _
   cast(PFNSQLITE3STATUSPROC, GetProcAddress(m_hLib, "sqlite3_status"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   DIM _resetFlag AS LONG = IIF(resetFlag = 0, 0, 1)
   FUNCTION = pProc(op, @pCurrent, @pHighwater, _resetFlag)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CSQLite.Status64 (BYVAL op AS LONG, BYREF pCurrent AS sqlite3_int64, BYREF pHighwater AS sqlite3_int64, BYVAL resetFlag AS BOOLEAN = FALSE) AS LONG
   DIM pProc AS PFNSQLITE3STATUS64PROC = _
   cast(PFNSQLITE3STATUS64PROC, GetProcAddress(m_hLib, "sqlite3_status64"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   DIM _resetFlag AS LONG = IIF(resetFlag = 0, 0, 1)
   FUNCTION = pProc(op, @pCurrent, @pHighwater, _resetFlag)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the English-language text that describes the result code, as UTF-8. Memory to hold
' the error message string is managed internally and must not be freed by the application.
' ========================================================================================
PRIVATE FUNCTION CSQLite.ErrStr (BYVAL nErrorCode AS LONG) AS STRING
   DIM pProc AS PFNSQLITE3ERRSTRPROC = _
   cast(PFNSQLITE3ERRSTRPROC, GetProcAddress(m_hLib, "sqlite3_errstr"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   DIM psz AS CONST ZSTRING PTR = pProc(nErrorCode)
   IF psz THEN FUNCTION = *psz
END FUNCTION
' ========================================================================================

' ========================================================================================
' StrGlob returns zero if and only if string szStr matches the GLOB pattern szGlob. The
' definition of GLOB pattern matching used in StrGlob is the same as for the "X GLOB P"
' operator in the SQL dialect understood by SQLite. The StrGlob method is case sensitive.
' Note that this method returns zero on a match and non-zero if the strings do not match.
' ========================================================================================
PRIVATE FUNCTION CSQLite.StrGlob (BYREF szGlob AS ZSTRING, BYREF szStr AS ZSTRING) AS LONG
   DIM pProc AS PFNSQLITE3STRGLOBPROC = _
   cast(PFNSQLITE3STRGLOBPROC, GetProcAddress(m_hLib, "sqlite3_strglob"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   FUNCTION = pProc(@szGlob, @szStr)
END FUNCTION
' ========================================================================================

TYPE PFNSQLITE3CLOSEPROC AS FUNCTION CDECL (BYVAL hDbc AS sqlite3 PTR) AS LONG
TYPE PFNSQLITE3OPEN16PROC AS FUNCTION CDECL (BYVAL filename AS CONST ANY PTR, BYVAL ppDb AS sqlite3 PTR PTR) AS LONG
TYPE PFNSQLITE3ERRCODEPROC AS FUNCTION CDECL (BYVAL hDbc AS sqlite3 PTR) AS LONG
TYPE PFNSQLITE3EXTENDEDERRCODEPROC AS FUNCTION CDECL (BYVAL hDbc AS sqlite3 PTR) AS LONG
TYPE PFNSQLITE3ERRMSGPROC AS FUNCTION CDECL (BYVAL hDbc AS sqlite3 PTR) AS ZSTRING PTR
TYPE PFNSQLITE3CHANGESPROC AS FUNCTION CDECL (BYVAL hDbc AS sqlite3 PTR) AS LONG
TYPE PFNSQLITE3TOTALCHANGESPROC AS FUNCTION CDECL (BYVAL hDbc AS sqlite3 PTR) AS LONG
TYPE PFNSQLITE3EXTENDEDRESULTCODESPROC AS FUNCTION CDECL (BYVAL hDbc AS sqlite3 PTR, BYVAL onoff AS LONG) AS LONG
TYPE PFNSQLITE3LASTINSERTROWIDPROC AS FUNCTION CDECL (BYVAL hDbc AS sqlite3 PTR) AS sqlite3_int64
TYPE PFNSQLITE3INTERRUPTPROC AS SUB CDECL (BYVAL hDbc AS sqlite3 PTR)
TYPE PFNSQLITE3LIMITPROC AS FUNCTION CDECL (BYVAL hDbc AS sqlite3 PTR, BYVAL id AS LONG, BYVAL newVal AS LONG) AS LONG
TYPE PFNSQLITE3DBRELEASEMEMORYPROC AS FUNCTION CDECL (BYVAL hDbc AS sqlite3 PTR) AS LONG
TYPE PFNSQLITE3PREPARE16V2PROC AS FUNCTION CDECL (BYVAL hDbc AS sqlite3 PTR, BYVAL pSql AS CONST ANY PTR, _
     BYVAL nByte AS LONG, BYVAL ppStmt AS sqlite3_stmt PTR PTR, BYVAL pzTail AS CONST ANY PTR PTR) AS LONG
TYPE PFNSQLITE3DBSTATUSPROC AS FUNCTION CDECL (BYVAL hDbc AS sqlite3 PTR, BYVAL op AS LONG, BYVAL pCurrent AS LONG PTR, BYVAL pHighwater AS LONG PTR, BYVAL resetFlag AS LONG) AS LONG
TYPE PFNSQLITE3PROGRESSHANDLERPROC AS SUB CDECL (BYVAL hStmt AS sqlite3 PTR, BYVAL nOps AS LONG, BYVAL pCallback AS ANY PTR, BYVAL pArg AS ANY PTR)
TYPE PFNSQLITE3UNLOCKNOTIFYPROC AS FUNCTION CDECL (BYVAL hStmt AS sqlite3 PTR, BYVAL pNotifyCallback AS ANY PTR, BYVAL pNotifyArg AS ANY PTR) AS LONG
TYPE PFNSQLITE3BLOBOPENPROC AS FUNCTION CDECL (BYVAL hStmt AS sqlite3 PTR, BYVAL pszDbName AS ZSTRING PTR, _
     BYVAL pszTableName AS ZSTRING PTR, BYVAL pszColumnName AS ZSTRING PTR, BYVAL iRow AS sqlite3_int64, _
     BYVAL flags AS LONG, BYVAL ppBlob AS sqlite3_blob PTR PTR) AS LONG

' // Statements
TYPE PFNSQLITE3COLUMNCOUNTPROC AS FUNCTION CDECL (BYVAL hStmt AS sqlite3_stmt PTR) AS LONG
TYPE PFNSQLITE3FINALIZEPROC AS FUNCTION CDECL (BYVAL hStmt AS sqlite3_stmt PTR) AS LONG
TYPE PFNSQLITE3RESETPROC AS FUNCTION CDECL (BYVAL hStmt AS sqlite3_stmt PTR) AS LONG
TYPE PFNSQLITE3STEPPROC AS FUNCTION CDECL (BYVAL hStmt AS sqlite3_stmt PTR) AS LONG
TYPE PFNSQLITE3COLUMNNAME16PROC AS FUNCTION CDECL (BYVAL hStmt AS sqlite3_stmt PTR, BYVAL nCol AS LONG) AS ANY PTR
TYPE PFNSQLITE3COLUMNTEXT16PROC AS FUNCTION CDECL (BYVAL hStmt AS sqlite3_stmt PTR, BYVAL nCol AS LONG) AS ANY PTR
TYPE PFNSQLITE3DATACOUNTPROC AS FUNCTION CDECL (BYVAL hStmt AS sqlite3_stmt PTR) AS LONG
TYPE PFNSQLITE3DBHANDLEPROC AS FUNCTION CDECL (BYVAL hStmt AS sqlite3_stmt PTR) AS sqlite3 PTR
TYPE PFNSQLITE3SQLPROC AS FUNCTION CDECL (BYVAL hStmt AS sqlite3_stmt PTR) AS ZSTRING PTR
TYPE PFNSQLITE3STMTBUSYPROC AS FUNCTION CDECL (BYVAL hStmt AS sqlite3_stmt PTR) AS LONG
TYPE PFNSQLITE3STMTREADONLYPROC AS FUNCTION CDECL (BYVAL hStmt AS sqlite3_stmt PTR) AS LONG
TYPE PFNSQLITE3BINDBLOBPROC AS FUNCTION CDECL (BYVAL hStmt AS sqlite3_stmt PTR, BYVAL idx AS LONG, _
     BYVAL pValue AS ANY PTR, BYVAL numBytes AS LONG, BYVAL pDestructor AS ANY PTR = NULL) AS LONG
TYPE PFNSQLITE3BINDBLOB64PROC AS FUNCTION CDECL (BYVAL hStmt AS sqlite3_stmt PTR, BYVAL idx AS LONG, _
     BYVAL pValue AS ANY PTR, BYVAL numBytes AS sqlite3_uint64, BYVAL pDestructor AS ANY PTR = NULL) AS LONG
TYPE PFNSQLITE3BINDDOUBLEPROC AS FUNCTION CDECL (BYVAL hStmt AS sqlite3_stmt PTR, BYVAL idx AS LONG, BYVAL dblValue AS DOUBLE) AS LONG
TYPE PFNSQLITE3BINDINTPROC AS FUNCTION CDECL (BYVAL hStmt AS sqlite3_stmt PTR, BYVAL idx AS LONG, BYVAL longValue AS LONG) AS LONG
TYPE PFNSQLITE3BINDINT64PROC AS FUNCTION CDECL (BYVAL hStmt AS sqlite3_stmt PTR, BYVAL idx AS LONG, BYVAL longintValue AS sqlite3_int64) AS LONG
TYPE PFNSQLITE3BINDNULLPROC AS FUNCTION CDECL (BYVAL hStmt AS sqlite3_stmt PTR) AS LONG
TYPE PFNSQLITE3BINDTEXT16PROC AS FUNCTION CDECL (BYVAL hStmt AS sqlite3_stmt PTR, BYVAL idx AS LONG, BYVAL pwszValue AS ANY PTR, BYVAL nBytes AS LONG, BYVAL pDestructor AS ANY PTR = NULL) AS LONG
TYPE PFNSQLITE3BINDZEROBLOBPROC AS FUNCTION CDECL (BYVAL hStmt AS sqlite3_stmt PTR, BYVAL idx AS LONG, BYVAL numBytes AS LONG) AS LONG
TYPE PFNSQLITE3BINDZEROBLOB64PROC AS FUNCTION CDECL (BYVAL hStmt AS sqlite3_stmt PTR, BYVAL idx AS LONG, BYVAL numBytes AS sqlite3_uint64) AS LONG
TYPE PFNSQLITE3BINDPARAMETERCOUNTPROC AS FUNCTION CDECL (BYVAL hStmt AS sqlite3_stmt PTR) AS LONG
TYPE PFNSQLITE3BINDPARAMETERNAMEPROC AS FUNCTION CDECL (BYVAL hStmt AS sqlite3_stmt PTR, BYVAL idx AS LONG) AS ZSTRING PTR
TYPE PFNSQLITE3BINDPARAMETERINDEXPROC AS FUNCTION CDECL (BYVAL hStmt AS sqlite3_stmt PTR, BYVAL pszName AS ZSTRING PTR) AS LONG
TYPE PFNSQLITE3CLEARBINDINGSPROC AS FUNCTION CDECL (BYVAL hStmt AS sqlite3_stmt PTR) AS LONG
TYPE PFNSQLITE3COLUMNDATABASENAMEPROC AS FUNCTION CDECL (BYVAL hStmt AS sqlite3_stmt PTR, BYVAL nCol AS LONG) AS ANY PTR
TYPE PFNSQLITE3COLUMNTABLENAMEPROC AS FUNCTION CDECL (BYVAL hStmt AS sqlite3_stmt PTR, BYVAL nCol AS LONG) AS ANY PTR
TYPE PFNSQLITE3COLUMNORIGINNAMEPROC AS FUNCTION CDECL (BYVAL hStmt AS sqlite3_stmt PTR, BYVAL nCol AS LONG) AS ANY PTR
TYPE PFNSQLITE3COLUMNDECLTYPEPROC AS FUNCTION CDECL (BYVAL hStmt AS sqlite3_stmt PTR, BYVAL nCol AS LONG) AS ANY PTR
TYPE PFNSQLITE3COLUMNBLOBPROC AS FUNCTION CDECL (BYVAL hStmt AS sqlite3_stmt PTR, BYVAL nCol AS LONG) AS ANY PTR
TYPE PFNSQLITE3COLUMNBYTESPROC AS FUNCTION CDECL (BYVAL hStmt AS sqlite3_stmt PTR, BYVAL nCol AS LONG) AS LONG
TYPE PFNSQLITE3COLUMNDOUBLEPROC AS FUNCTION CDECL (BYVAL hStmt AS sqlite3_stmt PTR, BYVAL nCol AS LONG) AS DOUBLE
TYPE PFNSQLITE3COLUMNINTPROC AS FUNCTION CDECL (BYVAL hStmt AS sqlite3_stmt PTR, BYVAL nCol AS LONG) AS LONG
TYPE PFNSQLITE3COLUMNINT64PROC AS FUNCTION CDECL (BYVAL hStmt AS sqlite3_stmt PTR, BYVAL nCol AS LONG) AS sqlite3_int64
TYPE PFNSQLITE3COLUMNTYPEPROC AS FUNCTION CDECL (BYVAL hStmt AS sqlite3_stmt PTR, BYVAL nCol AS LONG) AS LONG

' ########################################################################################
' CSQLiteDb class
' ########################################################################################
TYPE CSQLiteDb EXTENDS CSQLite

Public:
   DIM m_hDbc AS sqlite3 PTR

Public:
   DECLARE CONSTRUCTOR (BYREF wszPath AS CONST WSTRING)
   DECLARE DESTRUCTOR
   DECLARE FUNCTION OpenDb (BYREF wszFileName AS CONST WSTRING) AS LONG
   DECLARE FUNCTION CloseDb () AS LONG
   DECLARE PROPERTY hDbc () AS sqlite3 PTR
   DECLARE PROPERTY hDbc (BYVAL pDbc AS sqlite3 PTR)
   DECLARE FUNCTION ErrCode () AS LONG
   DECLARE FUNCTION ExtendedErrCode () AS LONG
   DECLARE FUNCTION ExtendedResultCodes (BYVAL onoff AS BOOLEAN) AS LONG
   DECLARE FUNCTION ErrMsg () AS STRING
   DECLARE FUNCTION Prepare (BYREF wszSql AS WSTRING) AS sqlite3_stmt PTR
   DECLARE FUNCTION Exec (BYREF wszSql AS WSTRING) AS LONG
   DECLARE FUNCTION Changes () AS LONG
   DECLARE FUNCTION TotalChanges () AS LONG
   DECLARE FUNCTION LastInsertRowId () AS sqlite3_int64
   DECLARE SUB Interrupt
   DECLARE FUNCTION Limit (BYVAL id AS LONG, BYVAL newVal AS LONG) AS LONG
   DECLARE FUNCTION ReleaseMemory () AS LONG
   DECLARE FUNCTION Status (BYVAL op AS LONG, BYREF pCurrent AS LONG, BYREF pHighwater AS LONG, BYVAL resetFlag AS BOOLEAN = FALSE) AS LONG
   DECLARE SUB ProgressHandler (BYVAL nOps AS LONG, BYVAL pCallback AS ANY PTR, BYVAL pArg AS ANY PTR)
   DECLARE FUNCTION UnlockNotify (BYVAL pNotifyCallback AS ANY PTR, BYVAL pNotifyArg AS ANY PTR) AS LONG
   DECLARE FUNCTION OpenBlob (BYREF szDbName AS ZSTRING, BYREF szTableName AS ZSTRING, _
           BYREF szColumnName AS ZSTRING, BYVAL iRow AS sqlite3_int64, BYVAL flags AS LONG = 0) AS sqlite3_blob PTR

END TYPE
' ########################################################################################

' ========================================================================================
' Constructor
' Opens an SQLite database file as specified by the filename argument. The database is
' opened for reading and writing, and is created if it does not already exist. If the
' filename is ":memory:", then a private, temporary in-memory database is created for the
' connection. This in-memory database will vanish when the database connection is closed.
' If the filename is an empty string, then a private, temporary on-disk database will be
' created. This private database will be automatically deleted as soon as the database
' connection is closed. If URI filename interpretation is enabled, and the filename
' argument begins with "file:", then the filename is interpreted as a URI.
' ========================================================================================
PRIVATE CONSTRUCTOR CSQLiteDb (BYREF wszFileName AS CONST WSTRING)
   DIM pProc AS PFNSQLITE3OPEN16PROC = _
   cast(PFNSQLITE3OPEN16PROC, GetProcAddress(this.m_hLib, "sqlite3_open16"))
   IF pProc = NULL THEN MessageBoxW(GetActiveWindow, "Unable to load " & wszFileName, "Error", MB_APPLMODAL OR MB_ICONERROR) : EXIT CONSTRUCTOR
   this.SetResult(pProc(@wszFileName, @m_hDbc))
END CONSTRUCTOR
' ========================================================================================

' ===========================================================================================
' Destructor
' ===========================================================================================
PRIVATE DESTRUCTOR CSQLiteDb
   ' // Closes the database connection
   this.CloseDb
END DESTRUCTOR
' ===========================================================================================

' ========================================================================================
' Opens an SQLite database file as specified by the filename argument. The database is
' opened for reading and writing, and is created if it does not already exist. If the
' filename is ":memory:", then a private, temporary in-memory database is created for the
' connection. This in-memory database will vanish when the database connection is closed.
' If the filename is an empty string, then a private, temporary on-disk database will be
' created. This private database will be automatically deleted as soon as the database
' connection is closed. If URI filename interpretation is enabled, and the filename
' argument begins with "file:", then the filename is interpreted as a URI.
' ========================================================================================
PRIVATE FUNCTION CSQLiteDb.OpenDb (BYREF wszFileName AS CONST WSTRING) AS LONG
   this.CloseDb   ' // Close the database
   DIM pProc AS PFNSQLITE3OPEN16PROC = _
   cast(PFNSQLITE3OPEN16PROC, GetProcAddress(m_hLib, "sqlite3_open16"))
   RETURN this.SetResult(pProc(@wszFileName, @m_hDbc))
END FUNCTION
' ========================================================================================

' ===========================================================================================
' Closes the database. Calling CloseDb with a NULL pointer argument is a harmless no-op.
' ===========================================================================================
PRIVATE FUNCTION CSQLiteDb.CloseDb () AS LONG
   DIM pProc AS PFNSQLITE3CLOSEPROC = _
   cast(PFNSQLITE3CLOSEPROC, GetProcAddress(m_hLib, "SQLite3_close"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   FUNCTION = pProc(m_hDbc)
   m_hDbc = NULL
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Gets the database handle
' ===========================================================================================
PRIVATE PROPERTY CSQLiteDb.hDbc () AS sqlite3 PTR
   PROPERTY = m_hDbc
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Sets the database handle
' ===========================================================================================
PRIVATE PROPERTY CSQLiteDb.hDbc (BYVAL pDbc AS sqlite3 PTR)
   this.CloseDb   ' // Close the database
   m_hDbc = pDbc
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Gets the error code associated with this dabatase connection.
' ===========================================================================================
PRIVATE FUNCTION CSQLiteDb.ErrCode () AS LONG
   DIM pProc AS PFNSQLITE3ERRCODEPROC = _
   cast(PFNSQLITE3ERRCODEPROC, GetProcAddress(m_hLib, "SQLite3_errcode"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   FUNCTION = pProc(m_hDbc)
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Gets the extended error code associated with this dabatase connection.
' ===========================================================================================
PRIVATE FUNCTION CSQLiteDb.ExtendedErrCode () AS LONG
   DIM pProc AS PFNSQLITE3EXTENDEDERRCODEPROC = _
   cast(PFNSQLITE3EXTENDEDERRCODEPROC, GetProcAddress(m_hLib, "sqlite3_extended_errcode"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   FUNCTION = pProc(m_hDbc)
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Enables or disables the extended result codes feature of SQLite. The extended result codes
' are disabled by default for historical compatibility.
' ===========================================================================================
PRIVATE FUNCTION CSQLiteDb.ExtendedResultCodes (BYVAL onoff AS BOOLEAN) AS LONG
   DIM pProc AS PFNSQLITE3EXTENDEDRESULTCODESPROC = _
   cast(PFNSQLITE3EXTENDEDRESULTCODESPROC, GetProcAddress(m_hLib, "sqlite3_extended_result_codes"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   FUNCTION = pProc(m_hDbc, IIF(onoff = TRUE, 1, 0))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Gets the extended error code associated with this dabatase connection.
' ===========================================================================================
PRIVATE FUNCTION CSQLiteDb.ErrMsg () AS STRING
   DIM pProc AS PFNSQLITE3ERRMSGPROC = _
   cast(PFNSQLITE3ERRMSGPROC, GetProcAddress(m_hLib, "sqlite3_errmsg"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   DIM psz AS CONST ZSTRING PTR = pProc(m_hDbc)
   IF psz THEN FUNCTION = *psz
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Returns the number of database rows that were changed or inserted or deleted by the
' most recently completed SQL statement.
' ===========================================================================================
PRIVATE FUNCTION CSQLiteDb.Changes () AS LONG
   DIM pProc AS PFNSQLITE3CHANGESPROC = _
   cast(PFNSQLITE3CHANGESPROC, GetProcAddress(m_hLib, "sqlite3_changes"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   FUNCTION = pProc(m_hDbc)
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Returns the total number of rows inserted, modified or deleted by all INSERT, UPDATE or
' DELETE statements completed since the database connection was opened, including those
' executed as part of trigger programs. Executing any other type of SQL statement does not
' affect the value returned by sqlite3_total_changes().
' ===========================================================================================
PRIVATE FUNCTION CSQLiteDb.TotalChanges () AS LONG
   DIM pProc AS PFNSQLITE3TOTALCHANGESPROC = _
   cast(PFNSQLITE3TOTALCHANGESPROC, GetProcAddress(m_hLib, "sqlite3_total_changes"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   FUNCTION = pProc(m_hDbc)
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Returns the rowid of the most recent successful INSERT into the database from the
' database connection.
' ===========================================================================================
PRIVATE FUNCTION CSQLiteDb.LastInsertRowId () AS sqlite3_int64
   DIM pProc AS PFNSQLITE3LASTINSERTROWIDPROC = _
   cast(PFNSQLITE3LASTINSERTROWIDPROC, GetProcAddress(m_hLib, "sqlite3_last_insert_rowid"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   FUNCTION = pProc(m_hDbc)
END FUNCTION
' ===========================================================================================

' ========================================================================================
' This function causes any pending database operation to abort and return at its earliest
' opportunity. This routine is typically called in response to a user action such as pressing
' "Cancel" or Ctrl-C where the user wants a long query operation to halt immediately.
' ========================================================================================
PRIVATE SUB CSQLiteDb.Interrupt
   DIM pProc AS PFNSQLITE3INTERRUPTPROC = _
   cast(PFNSQLITE3INTERRUPTPROC, GetProcAddress(m_hLib, "sqlite3_interrupt"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT SUB
   pProc(m_hDbc)
END SUB
' ========================================================================================

' ========================================================================================
' This method allows the size of various constructs to be limited on a connection by connection basis.
' ========================================================================================
PRIVATE FUNCTION CSQLiteDb.Limit (BYVAL id AS LONG, BYVAL newVal AS LONG) AS LONG
   DIM pProc AS PFNSQLITE3LIMITPROC = _
   cast(PFNSQLITE3LIMITPROC, GetProcAddress(m_hLib, "sqlite3_limit"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   FUNCTION = pProc(m_hDbc, id, newVal)
END FUNCTION
' ========================================================================================

' ========================================================================================
' This method allows the size of various constructs to be limited on a connection by connection basis.
' ========================================================================================
PRIVATE FUNCTION CSQLiteDb.ReleaseMemory () AS LONG
   DIM pProc AS PFNSQLITE3DBRELEASEMEMORYPROC = _
   cast(PFNSQLITE3DBRELEASEMEMORYPROC, GetProcAddress(m_hLib, "sqlite3_db_release_memory"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   FUNCTION = pProc(m_hDbc)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates a new prepared statement object.
' ========================================================================================
PRIVATE FUNCTION CSQLiteDb.Prepare (BYREF wszSql AS WSTRING) AS sqlite3_stmt PTR
   DIM pProc AS PFNSQLITE3PREPARE16V2PROC = _
   cast(PFNSQLITE3PREPARE16V2PROC, GetProcAddress(m_hLib, "sqlite3_prepare16_v2"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   DIM ppStmt AS sqlite3_stmt PTR
   this.SetResult(pProc(m_hDbc, @wszSql, -1, @ppStmt, NULL))
   FUNCTION = ppStmt
END FUNCTION
' ========================================================================================

' ========================================================================================
' Convenience wrapper for Prepare and Step. To be used with queries that don't return
' result sets, such CREATE, UPDATE and INSERT.
' SQLITE_BUSY means that the database engine was unable to acquire the database locks it needs
' to do its job. If the statement is a COMMIT or occurs outside of an explicit transaction,
' then you can retry the statement. If the statement is not a COMMIT and occurs within an
' explicit transaction then you should rollback the transaction before continuing.
' SQLITE_DONE means that the statement has finished executing successfully. Step should not
' be called again on this virtual machine without first calling Reset to reset the virtual
' machine back to its initial state.
' SQLITE_ERROR means that a run-time error (such as a constraint violation) has occurred.
' Step should not be called again on the cirtual machine. More information may be found by
' calling ErrMsg.
' ========================================================================================
PRIVATE FUNCTION CSQLiteDb.Exec (BYREF wszSql AS WSTRING) AS LONG
   DIM pStmt AS sqlite3_stmt PTR = this.Prepare(wszSql)
   DIM pProc AS PFNSQLITE3STEPPROC = _
   cast(PFNSQLITE3STEPPROC, GetProcAddress(m_hLib, "sqlite3_step"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   FUNCTION = pProc(pStmt)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves runtime status information about a single database connection.
' The current value of the requested parameter is written into pCurrent and the highest
' instantaneous value is written into pHighwater. If the resetFlg is true, then the highest
' instantaneous value is reset back down to the current value.
' Returns SQLITE_OK on success and a non-zero error code on failure.
' ========================================================================================
PRIVATE FUNCTION CSQLiteDb.Status (BYVAL op AS LONG, BYREF pCurrent AS LONG, BYREF pHighwater AS LONG, BYVAL resetFlag AS BOOLEAN = FALSE) AS LONG
   DIM pProc AS PFNSQLITE3DBSTATUSPROC = _
   cast(PFNSQLITE3DBSTATUSPROC, GetProcAddress(m_hLib, "sqlite3_db_status"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   DIM _resetFlag AS LONG = IIF(resetFlag = 0, 0, 1)
   FUNCTION = pProc(m_hDbc, op, @pCurrent, @pHighwater, _resetFlag)
END FUNCTION
' ========================================================================================

' ===========================================================================================
' Causes the callback function to be invoked periodically during long running calls to
' sqlite3_exec(), sqlite3_step() and sqlite3_get_table() for this database connection.
' An example use for this interface is to keep a GUI updated during a large query.
' The parameter pArg is passed through as the only parameter to the callback function. The
' parameter nOps is the approximate number of virtual machine instructions that are evaluated
' between successive invocations of the callback. If nOps is less than one then the progress
' handler is disabled. Only a single progress handler may be defined at one time per database
' connection; setting a new progress handler cancels the old one. Setting parameter pCallback
' to NULL disables the progress handler. The progress handler is also disabled by setting
' nOps to a value less than 1. If the progress callback returns non-zero, the operation is
' interrupted. This feature can be used to implement a "Cancel" button on a GUI progress
' dialog box.
' Callback function prototype: FUNCTION ProgressHandlerCallback (BYVAL pArg AS ANY PTR) AS LONG
' ===========================================================================================
PRIVATE SUB CSQLiteDb.ProgressHandler (BYVAL nOps AS LONG, BYVAL pCallback AS ANY PTR, BYVAL pArg AS ANY PTR)
   DIM pProc AS PFNSQLITE3PROGRESSHANDLERPROC = _
   cast(PFNSQLITE3PROGRESSHANDLERPROC, GetProcAddress(m_hLib, "sqlite3_progress_handler"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT SUB
   pProc(m_hDbc, nOps, pCallback, pArg)
END SUB
' ===========================================================================================

' ===========================================================================================
' When running in shared-cache mode, a database operation may fail with an SQLITE_LOCKED error
' if the required locks on the shared-cache or individual tables within the shared-cache cannot
' be obtained. This function may be used to register a callback that SQLite will invoke when
' the connection currently holding the required lock relinquishes it. This function is only
' available if the library was compiled with the SQLITE_ENABLE_UNLOCK_NOTIFY .
' Shared-cache locks are released when a database connection concludes its current transaction,
' either by committing it or rolling it back.
' ===========================================================================================
PRIVATE FUNCTION CSQLiteDb.UnlockNotify (BYVAL pNotifyCallback AS ANY PTR, BYVAL pNotifyArg AS ANY PTR) AS LONG
   DIM pProc AS PFNSQLITE3UNLOCKNOTIFYPROC = _
   cast(PFNSQLITE3UNLOCKNOTIFYPROC, GetProcAddress(m_hLib, "sqlite3_unlock_notify"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   FUNCTION = pProc(m_hDbc, pNotifyCallback, pNotifyArg)
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' This method opens a handle to the BLOB located in row iRow, column szColumnName, table
' szTableName in database szDbName; in other words, the same BLOB that would be selected by:
'    SELECT szColumnName FROM szDbName.szTableName WHERE rowid = iRow;
' Parameter szDbName is not the filename that contains the database, but rather the symbolic
' name of the database. For attached databases, this is the name that appears after the AS
' keyword in the ATTACH statement. For the main database file, the database name is "main".
' For TEMP tables, the database name is "temp".
' If the flags parameter is non-zero, then the BLOB is opened for read and write access.
' If the flags parameter is zero, the BLOB is opened for read-only access.
' On success, the new BLOB handle is returned. Otherwise NULL is returned.
' To check for an error code, call the GetLastResult method.
' ===========================================================================================
PRIVATE FUNCTION CSQLiteDb.OpenBlob (BYREF szDbName AS ZSTRING, BYREF szTableName AS ZSTRING, _
   BYREF szColumnName AS ZSTRING, BYVAL iRow AS sqlite3_int64, BYVAL flags AS LONG = 0) AS sqlite3_blob PTR
   DIM pProc AS PFNSQLITE3BLOBOPENPROC = _
   cast(PFNSQLITE3BLOBOPENPROC, GetProcAddress(m_hLib, "sqlite3_blob_open"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   DIM ppBlob AS sqlite3_blob PTR
   this.SetResult(pProc(m_hDbc, @szDbName, @szTableName, @szColumnName, iRow, flags, @ppBlob))
   FUNCTION = ppBlob
END FUNCTION
' ===========================================================================================

' ########################################################################################
' CSQLiteStmt class
' ########################################################################################
TYPE CSQLiteStmt EXTENDS CSQLite

Public:
   DIM m_hStmt AS sqlite3_stmt PTR
   DIM m_pColNames AS CDicObj PTR   ' Array of column names

Public:
   DECLARE CONSTRUCTOR (BYVAL pStmt AS sqlite3_stmt PTR)
   DECLARE DESTRUCTOR
   DECLARE PROPERTY hStmt () AS sqlite3_stmt PTR
   DECLARE PROPERTY hStmt (BYVAL pStmt AS sqlite3_stmt PTR)
   DECLARE FUNCTION GetColumNames () AS CDicObj PTR
   DECLARE FUNCTION GetRow () AS LONG
   DECLARE FUNCTION Step_ () AS LONG
   DECLARE FUNCTION Finalize () AS LONG
   DECLARE FUNCTION ColumnCount () AS LONG
   DECLARE FUNCTION Reset () AS LONG
   DECLARE FUNCTION ColumnName (BYVAL nCol AS LONG) AS CWSTR
   DECLARE FUNCTION ColumnText (BYVAL nCol AS LONG) AS CWSTR
   DECLARE FUNCTION ColumnText (BYREF wszColName AS WSTRING) AS CWSTR
   DECLARE FUNCTION DataCount () AS LONG
   DECLARE FUNCTION DbHandle () AS sqlite3 PTR
   DECLARE FUNCTION Sql () AS STRING
   DECLARE FUNCTION Busy () AS BOOLEAN
   DECLARE FUNCTION ReadOnly () AS BOOLEAN
   DECLARE FUNCTION BindBlob (BYVAL idx AS LONG, BYVAL pValue AS ANY PTR, _
           BYVAL numBytes AS LONG, BYVAL pDestructor AS ANY PTR = NULL) AS LONG
   DECLARE FUNCTION BindBlob64 (BYVAL idx AS LONG, BYVAL pValue AS ANY PTR, _
           BYVAL numBytes AS sqlite3_uint64, BYVAL pDestructor AS ANY PTR = NULL) AS LONG
   DECLARE FUNCTION BindDouble (BYVAL idx AS LONG, BYVAL dblValue AS DOUBLE) AS LONG
   DECLARE FUNCTION BindLong (BYVAL idx AS LONG, BYVAL longValue AS LONG) AS LONG
   DECLARE FUNCTION BindLongInt (BYVAL idx AS LONG, BYVAL longintValue AS sqlite3_int64) AS LONG
   DECLARE FUNCTION BindNull (BYVAL idx AS LONG) AS LONG
   DECLARE FUNCTION BindText (BYVAL idx AS LONG, BYREF wszValue AS WSTRING, BYVAL pDestructor AS ANY PTR = NULL) AS LONG
   DECLARE FUNCTION BindZeroBlob (BYVAL idx AS LONG, BYVAL numBytes AS LONG) AS LONG
   DECLARE FUNCTION BindZeroBlob64 (BYVAL idx AS LONG, BYVAL numBytes AS sqlite3_uint64) AS LONG
   DECLARE FUNCTION BindParameterCount () AS LONG
   DECLARE FUNCTION BindParameterName (BYVAL idx AS LONG) AS STRING
   DECLARE FUNCTION BindParameterIndex (BYREF szName AS ZSTRING) AS LONG
   DECLARE FUNCTION ClearBindings () AS LONG
   DECLARE FUNCTION ColumnDatabaseName (BYVAL nCol AS LONG) AS CWSTR
   DECLARE FUNCTION ColumnDatabaseName (BYREF wszColName AS WSTRING) AS CWSTR
   DECLARE FUNCTION ColumnTableName (BYVAL nCol AS LONG) AS CWSTR
   DECLARE FUNCTION ColumnTableName (BYREF wszColName AS WSTRING) AS CWSTR
   DECLARE FUNCTION ColumnOriginName (BYVAL nCol AS LONG) AS CWSTR
   DECLARE FUNCTION ColumnOriginName (BYREF wszColName AS WSTRING) AS CWSTR
   DECLARE FUNCTION ColumnDeclaredType (BYVAL nCol AS LONG) AS CWSTR
   DECLARE FUNCTION ColumnDeclaredType (BYREF wszColName AS WSTRING) AS CWSTR
   DECLARE FUNCTION ColumnBlob (BYVAL nCol AS LONG) AS ANY PTR
   DECLARE FUNCTION ColumnBlob (BYREF wszColName AS WSTRING) AS ANY PTR
   DECLARE FUNCTION ColumnBytes (BYVAL nCol AS LONG) AS LONG
   DECLARE FUNCTION ColumnBytes (BYREF wszColName AS WSTRING) AS LONG
   DECLARE FUNCTION ColumnDouble (BYVAL nCol AS LONG) AS DOUBLE
   DECLARE FUNCTION ColumnDouble (BYREF wszColName AS WSTRING) AS DOUBLE
   DECLARE FUNCTION ColumnLong (BYVAL nCol AS LONG) AS LONG
   DECLARE FUNCTION ColumnLong (BYREF wszColName AS WSTRING) AS LONG
   DECLARE FUNCTION ColumnLongInt (BYVAL nCol AS LONG) AS LONGINT
   DECLARE FUNCTION ColumnLongInt (BYREF wszColName AS WSTRING) AS LONGINT
   DECLARE FUNCTION ColumnType (BYVAL nCol AS LONG) AS LONG
   DECLARE FUNCTION ColumnType (BYREF wszColName AS WSTRING) AS LONG
   DECLARE FUNCTION IsColumnNull (BYVAL nCol AS LONG) AS BOOLEAN
   DECLARE FUNCTION IsColumnNull (BYREF wszColName AS WSTRING) AS BOOLEAN

END TYPE
' ########################################################################################

' ===========================================================================================
' Constructor
' ===========================================================================================
PRIVATE CONSTRUCTOR CSQLiteStmt (BYVAL pStmt AS sqlite3_stmt PTR)
   m_hStmt = pStmt
   ' // Collection of column names
   m_pColNames = NEW CDicObj
   this.GetColumNames
END CONSTRUCTOR
' ===========================================================================================

' ===========================================================================================
' Destructor
' ===========================================================================================
PRIVATE DESTRUCTOR CSQLiteStmt
   ' // Deletes the prepared statement
   IF m_hStmt THEN this.Finalize
   ' // Clears and deletes the collection of column names
   IF m_pColNames THEN
      m_pColNames->RemoveAll
      Delete m_pColNames
   END IF
END DESTRUCTOR
' ===========================================================================================

' ===========================================================================================
' Gets the statement handle
' ===========================================================================================
PRIVATE PROPERTY CSQLiteStmt.hStmt () AS sqlite3_stmt PTR
   PROPERTY = m_hStmt
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Sets the statement handle
' ===========================================================================================
PRIVATE PROPERTY CSQLiteStmt.hStmt (BYVAL pStmt AS sqlite3_stmt PTR)
   ' // Deletes the prepared statement
   IF m_hStmt THEN this.Finalize
   m_hStmt = pStmt
   ' // Get the column names
   this.GetColumNames
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Finalize is called to delete a prepared statement. If the most recent evaluation of the
' statement encountered no errors or if the statement is never been evaluated, then Finalize
' returns SQLITE_OK. If the most recent evaluation of the statement failed, then Finalize
' returns the appropriate error code or extended error code.
' Finalize can be called at any point during the life cycle of a prepared statement: before
' the statement is ever evaluated, after one or more calls Reset, or after any call to
' Step regardless of whether or not the statement has completed execution.
' Calling Finalize on a NULL pointer is a harmless no-op.
' The application must finalize every prepared statement in order to avoid resource leaks.
' It is a grievous error for the application to try to use a prepared statement after it has
' been finalized. Any use of a prepared statement after it has been finalized can result in
' undefined and undesirable behavior such as segfaults and heap corruption.
' ===========================================================================================
PRIVATE FUNCTION CSQLiteStmt.Finalize () AS LONG
   DIM pProc AS PFNSQLITE3FINALIZEPROC = _
   cast(PFNSQLITE3FINALIZEPROC, GetProcAddress(m_hLib, "sqlite3_finalize"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   FUNCTION = pProc(m_hStmt)
   IF m_pColNames THEN m_pColNames->RemoveAll
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Resets a prepared statement object back to its initial state, ready to be re-executed. Any
' SQL statement variables that had values bound to them using the Bind* API retain their values.
' Use ClearBindings to reset the bindings.
' ===========================================================================================
PRIVATE FUNCTION CSQLiteStmt.Reset () AS LONG
   DIM pProc AS PFNSQLITE3RESETPROC = _
   cast(PFNSQLITE3RESETPROC, GetProcAddress(m_hLib, "sqlite3_reset"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   FUNCTION = pProc(m_hStmt)
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' After preparing an statement, this method must be called one or more times to evaluate the
' statement. If the SQL statement being executed returns any data, then SQLITE_ROW is returned
' each time a new row of data is ready for processing by the caller. The values may be accessed
' using the column access functions. Step is called again to retrieve the next row of data.
' SQLITE_BUSY means that the database engine was unable to acquire the database locks it needs
' to do its job. If the statement is a COMMIT or occurs outside of an explicit transaction,
' then you can retry the statement. If the statement is not a COMMIT and occurs within an
' explicit transaction then you should rollback the transaction before continuing.
' SQLITE_DONE means that the statement has finished executing successfully. Step should not
' be called again on this virtual machine without first calling Reset to reset the virtual
' machine back to its initial state.
' SQLITE_ERROR means that a run-time error (such as a constraint violation) has occurred.
' Step should not be called again on the cirtual machine. More information may be found by
' calling ErrMsg.
' Note: Named GetRow instead of Step because Step is a reserved Free Basic keyword.
' ===========================================================================================
PRIVATE FUNCTION CSQLiteStmt.GetRow () AS LONG
   DIM pProc AS PFNSQLITE3STEPPROC = _
   cast(PFNSQLITE3STEPPROC, GetProcAddress(m_hLib, "sqlite3_step"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   FUNCTION = pProc(m_hStmt)
END FUNCTION
' ===========================================================================================
' ===========================================================================================
PRIVATE FUNCTION CSQLiteStmt.Step_ () AS LONG
   DIM pProc AS PFNSQLITE3STEPPROC = _
   cast(PFNSQLITE3STEPPROC, GetProcAddress(m_hLib, "sqlite3_step"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   FUNCTION = pProc(m_hStmt)
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Returns the number of columns in the result set returned by the prepared statement. If
' this routine returns 0, that means the prepared statement returns no data (for example
' an UPDATE).
' ===========================================================================================
PRIVATE FUNCTION CSQLiteStmt.ColumnCount () AS LONG
   DIM pProc AS PFNSQLITE3COLUMNCOUNTPROC = _
   cast(PFNSQLITE3COLUMNCOUNTPROC, GetProcAddress(m_hLib, "sqlite3_column_count"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   FUNCTION = pProc(m_hStmt)
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Returns the name assigned to a particular column in the result set of a SELECT statement.
' The leftmost column is number 0. The name of a result column is the value of the "AS" clause
' for that column, if there is an AS clause.
' ===========================================================================================
PRIVATE FUNCTION CSQLiteStmt.ColumnName (BYVAL nCol AS LONG) AS CWSTR
   FUNCTION = ""
   DIM pProc AS PFNSQLITE3COLUMNNAME16PROC = _
   cast(PFNSQLITE3COLUMNNAME16PROC, GetProcAddress(m_hLib, "sqlite3_column_name16"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   DIM pwsz AS WSTRING PTR = pProc(m_hStmt, nCol)
   IF pwsz THEN FUNCTION = *pwsz
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Retrieves the column names
' ===========================================================================================
PRIVATE FUNCTION CSQLiteStmt.GetColumNames () AS CDicObj PTR
   ' // Clear the collection of column names
   IF m_pColNames THEN m_pColNames->RemoveAll
   ' // Check for a null pointer
   IF m_hStmt = NULL THEN EXIT FUNCTION
   ' // Get the number of columns
   DIM numCols AS LONG = this.ColumnCount
   IF numCols = 0 THEN EXIT FUNCTION
   ' // Get the column names
   DIM cwsName AS CWSTR
   FOR i AS LONG = 0 TO numCols - 1
      cwsName = this.ColumnName(i)
      IF m_pColNames THEN m_pColNames->Add(cwsName, CVAR(i, "LONG"))
   NEXT
   FUNCTION = m_pColNames
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Returns the name assigned to a particular column in the result set of a SELECT statement.
' The leftmost column is number 0. The name of a result column is the value of the "AS" clause
' for that column, if there is an AS clause.
' ===========================================================================================
PRIVATE FUNCTION CSQLiteStmt.ColumnText (BYVAL nCol AS LONG) AS CWSTR
   FUNCTION = ""
   DIM pProc AS PFNSQLITE3COLUMNTEXT16PROC = _
   cast(PFNSQLITE3COLUMNTEXT16PROC, GetProcAddress(m_hLib, "sqlite3_column_text16"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   DIM pwsz AS WSTRING PTR = pProc(m_hStmt, nCol)
   IF pwsz THEN FUNCTION = *pwsz
END FUNCTION
' ===========================================================================================
' ===========================================================================================
PRIVATE FUNCTION CSQLiteStmt.ColumnText (BYREF wszColName AS WSTRING) AS CWSTR
   FUNCTION = ""
   DIM pProc AS PFNSQLITE3COLUMNTEXT16PROC = _
   cast(PFNSQLITE3COLUMNTEXT16PROC, GetProcAddress(m_hLib, "sqlite3_column_text16"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   IF m_pColNames THEN
      IF m_pColNames->Exists(wszColName) THEN
         DIM nCol AS LONG = m_pColNames->Item(wszColName).ValLong
         IF m_pColNames->GetLastResult = S_OK THEN
            DIM pwsz AS WSTRING PTR = pProc(m_hStmt, nCol)
            IF pwsz THEN FUNCTION = *pwsz
         END IF
      END IF
   END IF
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Returns the number of columns in the current row of the result set of prepared statement.
' ===========================================================================================
PRIVATE FUNCTION CSQLiteStmt.DataCount () AS LONG
   DIM pProc AS PFNSQLITE3DATACOUNTPROC = _
   cast(PFNSQLITE3DATACOUNTPROC, GetProcAddress(m_hLib, "sqlite3_data_count"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   FUNCTION = pProc(m_hStmt)
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Returns the database connection handle to which a prepared statement belongs.
' ===========================================================================================
PRIVATE FUNCTION CSQLiteStmt.DbHandle () AS sqlite3 PTR
   DIM pProc AS PFNSQLITE3DBHANDLEPROC = _
   cast(PFNSQLITE3DBHANDLEPROC, GetProcAddress(m_hLib, "sqlite3_db_handle"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   FUNCTION = pProc(m_hStmt)
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Returns a saved copy of the original SQL text used to create a prepared statement.
' ===========================================================================================
PRIVATE FUNCTION CSQLiteStmt.Sql () AS STRING
   DIM pProc AS PFNSQLITE3SQLPROC = _
   cast(PFNSQLITE3SQLPROC, GetProcAddress(m_hLib, "sqlite3_sql"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   DIM psz AS ZSTRING PTR = pProc(m_hStmt)
   IF psz THEN FUNCTION = *psz
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Returns true if the prepared statement has been stepped at least once using GetRow but has
' neither run to completion (returned SQLITE_DONE from GetRow) nor been reset using Reset.
' ===========================================================================================
PRIVATE FUNCTION CSQLiteStmt.Busy () AS BOOLEAN
   DIM pProc AS PFNSQLITE3STMTBUSYPROC = _
   cast(PFNSQLITE3STMTBUSYPROC, GetProcAddress(m_hLib, "sqlite3_stmt_busy"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   FUNCTION = pProc(m_hStmt)
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Returns true if and only if the prepared statement makes no direct changes to the content
' of the database fiile.
' ===========================================================================================
PRIVATE FUNCTION CSQLiteStmt.ReadOnly () AS BOOLEAN
   DIM pProc AS PFNSQLITE3STMTREADONLYPROC = _
   cast(PFNSQLITE3STMTREADONLYPROC, GetProcAddress(m_hLib, "sqlite3_stmt_readonly"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   FUNCTION = pProc(m_hStmt)
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Binds a blob with the statement.
' The leftmost SQL parameter has an index of 1.
' ===========================================================================================
PRIVATE FUNCTION CSQLiteStmt.BindBlob (BYVAL idx AS LONG, BYVAL pValue AS ANY PTR, _
   BYVAL numBytes AS LONG, BYVAL pDestructor AS ANY PTR = NULL) AS LONG
   DIM pProc AS PFNSQLITE3BINDBLOBPROC = _
   cast(PFNSQLITE3BINDBLOBPROC, GetProcAddress(m_hLib, "sqlite3_bind_blob"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   FUNCTION = pProc(m_hStmt, idx, pValue, numBytes, pDestructor)
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Binds a blob with the statement.
' The leftmost SQL parameter has an index of 1.
' ===========================================================================================
PRIVATE FUNCTION CSQLiteStmt.BindBlob64 (BYVAL idx AS LONG, BYVAL pValue AS ANY PTR, _
   BYVAL numBytes AS sqlite3_uint64, BYVAL pDestructor AS ANY PTR = NULL) AS LONG
   DIM pProc AS PFNSQLITE3BINDBLOB64PROC = _
   cast(PFNSQLITE3BINDBLOB64PROC, GetProcAddress(m_hLib, "sqlite3_bind_blob64"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   FUNCTION = pProc(m_hStmt, idx, pValue, numBytes, pDestructor)
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Binds a double value with the statement.
' The leftmost SQL parameter has an index of 1.
' ===========================================================================================
PRIVATE FUNCTION CSQLiteStmt.BindDouble (BYVAL idx AS LONG, BYVAL dblValue AS DOUBLE) AS LONG
   DIM pProc AS PFNSQLITE3BINDDOUBLEPROC = _
   cast(PFNSQLITE3BINDDOUBLEPROC, GetProcAddress(m_hLib, "sqlite3_bind_double"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   FUNCTION = pProc(m_hStmt, idx, dblValue)
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Binds a long value with the statement.
' The leftmost SQL parameter has an index of 1.
' ===========================================================================================
PRIVATE FUNCTION CSQLiteStmt.BindLong (BYVAL idx AS LONG, BYVAL longValue AS LONG) AS LONG
   DIM pProc AS PFNSQLITE3BINDINTPROC = _
   cast(PFNSQLITE3BINDINTPROC, GetProcAddress(m_hLib, "sqlite3_bind_int"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   FUNCTION = pProc(m_hStmt, idx, longValue)
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Binds a longint value with the statement.
' The leftmost SQL parameter has an index of 1.
' ===========================================================================================
PRIVATE FUNCTION CSQLiteStmt.BindLongInt (BYVAL idx AS LONG, BYVAL longintValue AS sqlite3_int64) AS LONG
   DIM pProc AS PFNSQLITE3BINDINT64PROC = _
   cast(PFNSQLITE3BINDINT64PROC, GetProcAddress(m_hLib, "sqlite3_bind_int64"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   FUNCTION = pProc(m_hStmt, idx, longintValue)
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Binds a null value with the statement.
' The leftmost SQL parameter has an index of 1.
' ===========================================================================================
PRIVATE FUNCTION CSQLiteStmt.BindNull (BYVAL idx AS LONG) AS LONG
   DIM pProc AS PFNSQLITE3BINDNULLPROC = _
   cast(PFNSQLITE3BINDNULLPROC, GetProcAddress(m_hLib, "sqlite3_bind_null"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   FUNCTION = pProc(m_hStmt)
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Binds a text value with the statement.
' The leftmost SQL parameter has an index of 1.
' ===========================================================================================
PRIVATE FUNCTION CSQLiteStmt.BindText (BYVAL idx AS LONG, BYREF wszValue AS WSTRING, BYVAL pDestructor AS ANY PTR = NULL) AS LONG
   DIM pProc AS PFNSQLITE3BINDTEXT16PROC = _
   cast(PFNSQLITE3BINDTEXT16PROC, GetProcAddress(m_hLib, "sqlite3_bind_text16"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   FUNCTION = pProc(m_hStmt, idx, @wszValue, -1, pDestructor)
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Binds a blob of length numBytes that is filled with zeroes.
' The leftmost SQL parameter has an index of 1.
' ===========================================================================================
PRIVATE FUNCTION CSQLiteStmt.BindZeroBlob (BYVAL idx AS LONG, BYVAL numBytes AS LONG) AS LONG
   DIM pProc AS PFNSQLITE3BINDZEROBLOBPROC = _
   cast(PFNSQLITE3BINDZEROBLOBPROC, GetProcAddress(m_hLib, "sqlite3_bind_zeroblob"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   FUNCTION = pProc(m_hStmt, idx, numBytes)
END FUNCTION
' ===========================================================================================
' ===========================================================================================
PRIVATE FUNCTION CSQLiteStmt.BindZeroBlob64 (BYVAL idx AS LONG, BYVAL numBytes AS sqlite3_uint64) AS LONG
   DIM pProc AS PFNSQLITE3BINDZEROBLOB64PROC = _
   cast(PFNSQLITE3BINDZEROBLOB64PROC, GetProcAddress(m_hLib, "sqlite3_bind_zeroblob64"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   FUNCTION = pProc(m_hStmt, idx, numBytes)
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' This method actually returns the index of the largest (rightmost) parameter. For all
' forms except ?NNN, this will correspond to the number of unique parameters. If parameters
' of the ?NNN form are used, there may be gaps in the list.
' ===========================================================================================
PRIVATE FUNCTION CSQLiteStmt.BindParameterCount () AS LONG
   DIM pProc AS PFNSQLITE3BINDPARAMETERCOUNTPROC = _
   cast(PFNSQLITE3BINDPARAMETERCOUNTPROC, GetProcAddress(m_hLib, "sqlite3_bind_parameter_count"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   FUNCTION = pProc(m_hStmt)
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Returns the name of the specified SQL parameter in the prepared statement.
' ===========================================================================================
PRIVATE FUNCTION CSQLiteStmt.BindParameterName (BYVAL idx AS LONG) AS STRING
   DIM pProc AS PFNSQLITE3BINDPARAMETERNAMEPROC = _
   cast(PFNSQLITE3BINDPARAMETERNAMEPROC, GetProcAddress(m_hLib, "sqlite3_bind_parameter_name"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   DIM psz AS ZSTRING PTR = pProc(m_hStmt, idx)
   IF psz THEN FUNCTION = *psz
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Returns the name of the specified SQL parameter in the prepared statement.
' ===========================================================================================
PRIVATE FUNCTION CSQLiteStmt.BindParameterIndex (BYREF szName AS ZSTRING) AS LONG
   DIM pProc AS PFNSQLITE3BINDPARAMETERINDEXPROC = _
   cast(PFNSQLITE3BINDPARAMETERINDEXPROC, GetProcAddress(m_hLib, "sqlite3_bind_parameter_index"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   FUNCTION = pProc(m_hStmt, @szName)
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Sets all the parameters in the compiled SQL statement to NULL. Contrary to the intuition
' of many, sqlite3_reset does not reset the bindings on a prepared statement. Use this
' routine to reset all host parameters to NULL.
' ===========================================================================================
PRIVATE FUNCTION CSQLiteStmt.ClearBindings () AS LONG
   DIM pProc AS PFNSQLITE3CLEARBINDINGSPROC = _
   cast(PFNSQLITE3CLEARBINDINGSPROC, GetProcAddress(m_hLib, "sqlite3_clear_bindings"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   FUNCTION = pProc(m_hStmt)
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Returns the database name that is the origin of a particular result column in SELECT statement.
' The leftmost column of the result set has the index 0.
' ===========================================================================================
PRIVATE FUNCTION CSQLiteStmt.ColumnDatabaseName (BYVAL nCol AS LONG) AS CWSTR
   FUNCTION = ""
   DIM pProc AS PFNSQLITE3COLUMNDATABASENAMEPROC = _
   cast(PFNSQLITE3COLUMNDATABASENAMEPROC, GetProcAddress(m_hLib, "sqlite3_column_database_name16"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   DIM pwsz AS WSTRING PTR = pProc(m_hStmt, nCol)
   IF pwsz THEN FUNCTION = *pwsz
END FUNCTION
' ===========================================================================================
' ===========================================================================================
PRIVATE FUNCTION CSQLiteStmt.ColumnDatabaseName (BYREF wszColName AS WSTRING) AS CWSTR
   FUNCTION = ""
   DIM pProc AS PFNSQLITE3COLUMNDATABASENAMEPROC = _
   cast(PFNSQLITE3COLUMNDATABASENAMEPROC, GetProcAddress(m_hLib, "sqlite3_column_database_name16"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   IF m_pColNames THEN
      IF m_pColNames->Exists(wszColName) THEN
         DIM nCol AS LONG = m_pColNames->Item(wszColName).ValLong
         IF m_pColNames->GetLastResult = S_OK THEN
            DIM pwsz AS WSTRING PTR = pProc(m_hStmt, nCol)
            IF pwsz THEN FUNCTION = *pwsz
         END IF
      END IF
   END IF
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Returns the table name that is the origin of a particular result column in SELECT statement.
' The leftmost column of the result set has the index 0.
' ===========================================================================================
PRIVATE FUNCTION CSQLiteStmt.ColumnTableName (BYVAL nCol AS LONG) AS CWSTR
   FUNCTION = ""
   DIM pProc AS PFNSQLITE3COLUMNTABLENAMEPROC = _
   cast(PFNSQLITE3COLUMNTABLENAMEPROC, GetProcAddress(m_hLib, "sqlite3_column_table_name16"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   DIM pwsz AS WSTRING PTR = pProc(m_hStmt, nCol)
   IF pwsz THEN FUNCTION = *pwsz
END FUNCTION
' ===========================================================================================
' ===========================================================================================
PRIVATE FUNCTION CSQLiteStmt.ColumnTableName (BYREF wszColName AS WSTRING) AS CWSTR
   FUNCTION = ""
   DIM pProc AS PFNSQLITE3COLUMNTABLENAMEPROC = _
   cast(PFNSQLITE3COLUMNTABLENAMEPROC, GetProcAddress(m_hLib, "sqlite3_column_table_name16"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   IF m_pColNames THEN
      IF m_pColNames->Exists(wszColName) THEN
         DIM nCol AS LONG = m_pColNames->Item(wszColName).ValLong
         IF m_pColNames->GetLastResult = S_OK THEN
            DIM pwsz AS WSTRING PTR = pProc(m_hStmt, nCol)
            IF pwsz THEN FUNCTION = *pwsz
         END IF
      END IF
   END IF
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Returns the column name that is the origin of a particular result column in SELECT statement.
' The leftmost column of the result set has the index 0.
' ===========================================================================================
PRIVATE FUNCTION CSQLiteStmt.ColumnOriginName (BYVAL nCol AS LONG) AS CWSTR
   FUNCTION = ""
   DIM pProc AS PFNSQLITE3COLUMNORIGINNAMEPROC = _
   cast(PFNSQLITE3COLUMNORIGINNAMEPROC, GetProcAddress(m_hLib, "sqlite3_column_origin_name16"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   DIM pwsz AS WSTRING PTR = pProc(m_hStmt, nCol)
   IF pwsz THEN FUNCTION = *pwsz
END FUNCTION
' ===========================================================================================
' ===========================================================================================
PRIVATE FUNCTION CSQLiteStmt.ColumnOriginName (BYREF wszColName AS WSTRING) AS CWSTR
   FUNCTION = ""
   DIM pProc AS PFNSQLITE3COLUMNORIGINNAMEPROC = _
   cast(PFNSQLITE3COLUMNORIGINNAMEPROC, GetProcAddress(m_hLib, "sqlite3_column_origin_name16"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   IF m_pColNames THEN
      IF m_pColNames->Exists(wszColName) THEN
         DIM nCol AS LONG = m_pColNames->Item(wszColName).ValLong
         IF m_pColNames->GetLastResult = S_OK THEN
            DIM pwsz AS WSTRING PTR = pProc(m_hStmt, nCol)
            IF pwsz THEN FUNCTION = *pwsz
         END IF
      END IF
   END IF
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Returns the column name that is the origin of a particular result column in SELECT statement.
' The leftmost column of the result set has the index 0.
' ===========================================================================================
PRIVATE FUNCTION CSQLiteStmt.ColumnDeclaredType (BYVAL nCol AS LONG) AS CWSTR
   FUNCTION = ""
   DIM pProc AS PFNSQLITE3COLUMNDECLTYPEPROC = _
   cast(PFNSQLITE3COLUMNDECLTYPEPROC, GetProcAddress(m_hLib, "sqlite3_column_decltype16"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   DIM pwsz AS WSTRING PTR = pProc(m_hStmt, nCol)
   IF pwsz THEN FUNCTION = *pwsz
END FUNCTION
' ===========================================================================================
' ===========================================================================================
PRIVATE FUNCTION CSQLiteStmt.ColumnDeclaredType (BYREF wszColName AS WSTRING) AS CWSTR
   FUNCTION = ""
   DIM pProc AS PFNSQLITE3COLUMNDECLTYPEPROC = _
   cast(PFNSQLITE3COLUMNDECLTYPEPROC, GetProcAddress(m_hLib, "sqlite3_column_decltype16"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   IF m_pColNames THEN
      IF m_pColNames->Exists(wszColName) THEN
         DIM nCol AS LONG = m_pColNames->Item(wszColName).ValLong
         IF m_pColNames->GetLastResult = S_OK THEN
            DIM pwsz AS WSTRING PTR = pProc(m_hStmt, nCol)
            IF pwsz THEN FUNCTION = *pwsz
         END IF
      END IF
   END IF
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Returns the column value as a pointer to a blob.
' The leftmost column of the result set has the index 0.
' ===========================================================================================
PRIVATE FUNCTION CSQLiteStmt.ColumnBlob (BYVAL nCol AS LONG) AS ANY PTR
   DIM pProc AS PFNSQLITE3COLUMNBLOBPROC = _
   cast(PFNSQLITE3COLUMNBLOBPROC, GetProcAddress(m_hLib, "sqlite3_column_blob"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   FUNCTION = pProc(m_hStmt, nCol)
END FUNCTION
' ===========================================================================================
' ===========================================================================================
PRIVATE FUNCTION CSQLiteStmt.ColumnBlob (BYREF wszColName AS WSTRING) AS ANY PTR
   DIM pProc AS PFNSQLITE3COLUMNBLOBPROC = _
   cast(PFNSQLITE3COLUMNBLOBPROC, GetProcAddress(m_hLib, "sqlite3_column_blob"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   IF m_pColNames THEN
      IF m_pColNames->Exists(wszColName) THEN
         DIM nCol AS LONG = m_pColNames->Item(wszColName).ValLong
         IF m_pColNames->GetLastResult = S_OK THEN
            FUNCTION = pProc(m_hStmt, nCol)
         END IF
      END IF
   END IF
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Returns the number of bytes of the column value.
' The leftmost column of the result set has the index 0.
' ===========================================================================================
PRIVATE FUNCTION CSQLiteStmt.ColumnBytes (BYVAL nCol AS LONG) AS LONG
   DIM pProc AS PFNSQLITE3COLUMNBYTESPROC = _
   cast(PFNSQLITE3COLUMNBYTESPROC, GetProcAddress(m_hLib, "sqlite3_column_bytes16"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   FUNCTION = pProc(m_hStmt, nCol)
END FUNCTION
' ===========================================================================================
' ===========================================================================================
PRIVATE FUNCTION CSQLiteStmt.ColumnBytes (BYREF wszColName AS WSTRING) AS LONG
   DIM pProc AS PFNSQLITE3COLUMNBYTESPROC = _
   cast(PFNSQLITE3COLUMNBYTESPROC, GetProcAddress(m_hLib, "sqlite3_column_bytes16"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   IF m_pColNames THEN
      IF m_pColNames->Exists(wszColName) THEN
         DIM nCol AS LONG = m_pColNames->Item(wszColName).ValLong
         IF m_pColNames->GetLastResult = S_OK THEN
            FUNCTION = pProc(m_hStmt, nCol)
         END IF
      END IF
   END IF
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Returns the column value as a double.
' The leftmost column of the result set has the index 0.
' ===========================================================================================
PRIVATE FUNCTION CSQLiteStmt.ColumnDouble (BYVAL nCol AS LONG) AS DOUBLE
   DIM pProc AS PFNSQLITE3COLUMNDOUBLEPROC = _
   cast(PFNSQLITE3COLUMNDOUBLEPROC, GetProcAddress(m_hLib, "sqlite3_column_double"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   FUNCTION = pProc(m_hStmt, nCol)
END FUNCTION
' ===========================================================================================
' ===========================================================================================
PRIVATE FUNCTION CSQLiteStmt.ColumnDouble (BYREF wszColName AS WSTRING) AS DOUBLE
   DIM pProc AS PFNSQLITE3COLUMNDOUBLEPROC = _
   cast(PFNSQLITE3COLUMNDOUBLEPROC, GetProcAddress(m_hLib, "sqlite3_column_double"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   IF m_pColNames THEN
      IF m_pColNames->Exists(wszColName) THEN
         DIM nCol AS LONG = m_pColNames->Item(wszColName).ValLong
         IF m_pColNames->GetLastResult = S_OK THEN
            FUNCTION = pProc(m_hStmt, nCol)
         END IF
      END IF
   END IF
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Returns the column value as a long.
' The leftmost column of the result set has the index 0.
' ===========================================================================================
PRIVATE FUNCTION CSQLiteStmt.ColumnLong (BYVAL nCol AS LONG) AS LONG
   DIM pProc AS PFNSQLITE3COLUMNINTPROC = _
   cast(PFNSQLITE3COLUMNINTPROC, GetProcAddress(m_hLib, "sqlite3_column_int"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   FUNCTION = pProc(m_hStmt, nCol)
END FUNCTION
' ===========================================================================================
' ===========================================================================================
PRIVATE FUNCTION CSQLiteStmt.ColumnLong (BYREF wszColName AS WSTRING) AS LONG
   DIM pProc AS PFNSQLITE3COLUMNINTPROC = _
   cast(PFNSQLITE3COLUMNINTPROC, GetProcAddress(m_hLib, "sqlite3_column_int"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   IF m_pColNames THEN
      IF m_pColNames->Exists(wszColName) THEN
         DIM nCol AS LONG = m_pColNames->Item(wszColName).ValLong
         IF m_pColNames->GetLastResult = S_OK THEN
            FUNCTION = pProc(m_hStmt, nCol)
         END IF
      END IF
   END IF
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Returns the column value as a longint.
' The leftmost column of the result set has the index 0.
' ===========================================================================================
PRIVATE FUNCTION CSQLiteStmt.ColumnLongInt (BYVAL nCol AS LONG) AS LONGINT
   DIM pProc AS PFNSQLITE3COLUMNINT64PROC = _
   cast(PFNSQLITE3COLUMNINT64PROC, GetProcAddress(m_hLib, "sqlite3_column_int64"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   FUNCTION = pProc(m_hStmt, nCol)
END FUNCTION
' ===========================================================================================
' ===========================================================================================
PRIVATE FUNCTION CSQLiteStmt.ColumnLongInt (BYREF wszColName AS WSTRING) AS LONGINT
   DIM pProc AS PFNSQLITE3COLUMNINT64PROC = _
   cast(PFNSQLITE3COLUMNINT64PROC, GetProcAddress(m_hLib, "sqlite3_column_int64"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   IF m_pColNames THEN
      IF m_pColNames->Exists(wszColName) THEN
         DIM nCol AS LONG = m_pColNames->Item(wszColName).ValLong
         IF m_pColNames->GetLastResult = S_OK THEN
            FUNCTION = pProc(m_hStmt, nCol)
         END IF
      END IF
   END IF
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Returns the column type.
' The leftmost column of the result set has the index 0.
' ===========================================================================================
PRIVATE FUNCTION CSQLiteStmt.ColumnType (BYVAL nCol AS LONG) AS LONG
   DIM pProc AS PFNSQLITE3COLUMNTYPEPROC = _
   cast(PFNSQLITE3COLUMNTYPEPROC, GetProcAddress(m_hLib, "sqlite3_column_type"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   FUNCTION = pProc(m_hStmt, nCol)
END FUNCTION
' ===========================================================================================
' ===========================================================================================
PRIVATE FUNCTION CSQLiteStmt.ColumnType (BYREF wszColName AS WSTRING) AS LONG
   DIM pProc AS PFNSQLITE3COLUMNTYPEPROC = _
   cast(PFNSQLITE3COLUMNTYPEPROC, GetProcAddress(m_hLib, "sqlite3_column_type"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   IF m_pColNames THEN
      IF m_pColNames->Exists(wszColName) THEN
         DIM nCol AS LONG = m_pColNames->Item(wszColName).ValLong
         IF m_pColNames->GetLastResult = S_OK THEN
            FUNCTION = pProc(m_hStmt, nCol)
         END IF
      END IF
   END IF
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Return TRUE if the column has a null value.
' ===========================================================================================
PRIVATE FUNCTION CSQLiteStmt.IsColumnNull (BYVAL nCol AS LONG) AS BOOLEAN
   DIM pProc AS PFNSQLITE3COLUMNTYPEPROC = _
   cast(PFNSQLITE3COLUMNTYPEPROC, GetProcAddress(m_hLib, "sqlite3_column_type"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   IF pProc(m_hStmt, nCol) = SQLITE_NULL THEN FUNCTION = TRUE
END FUNCTION
' ===========================================================================================
' ===========================================================================================
PRIVATE FUNCTION CSQLiteStmt.IsColumnNull (BYREF wszColName AS WSTRING) AS BOOLEAN
   DIM pProc AS PFNSQLITE3COLUMNTYPEPROC = _
   cast(PFNSQLITE3COLUMNTYPEPROC, GetProcAddress(m_hLib, "sqlite3_column_type"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   IF m_pColNames THEN
      IF m_pColNames->Exists(wszColName) THEN
         DIM nCol AS LONG = m_pColNames->Item(wszColName).ValLong
         IF m_pColNames->GetLastResult = S_OK THEN
            IF pProc(m_hStmt, nCol) = SQLITE_NULL THEN FUNCTION = TRUE
         END IF
      END IF
   END IF
END FUNCTION
' ===========================================================================================

TYPE PFNSQLITE3BLOBCLOSEPROC AS FUNCTION CDECL (BYVAL pBlob AS sqlite3_blob PTR) AS LONG
TYPE PFNSQLITE3BLOBREOPENPROC AS FUNCTION CDECL (BYVAL pBlob AS sqlite3_blob PTR, BYVAL rowId AS sqlite3_int64) AS LONG
TYPE PFNSQLITE3BLOBBYTESPROC AS FUNCTION CDECL (BYVAL pBlob AS sqlite3_blob PTR) AS LONG
TYPE PFNSQLITE3BLOBREADPROC AS FUNCTION CDECL (BYVAL pBlob AS sqlite3_blob PTR, BYVAL pBuffer AS ANY PTR, _
     BYVAL numBytes AS LONG, BYVAL iOffset AS LONG = 0) AS LONG
TYPE PFNSQLITE3BLOBWRITEPROC AS FUNCTION CDECL (BYVAL pBlob AS sqlite3_blob PTR, BYVAL pBuffer AS ANY PTR, _
     BYVAL numBytes AS LONG, BYVAL iOffset AS LONG = 0) AS LONG

' ########################################################################################
' CSQLiteBlob class
' ########################################################################################
TYPE CSQLiteBlob EXTENDS CSQLite

Public:
   DIM m_hBlob AS sqlite3_blob PTR

Public:
   DECLARE CONSTRUCTOR (BYVAL pBlob AS sqlite3_blob PTR)
   DECLARE DESTRUCTOR
   DECLARE PROPERTY hBlob () AS sqlite3_blob PTR
   DECLARE PROPERTY hBlob (BYVAL AS sqlite3_blob PTR)
   DECLARE FUNCTION Close () AS LONG
   DECLARE FUNCTION Reopen (BYVAL rowId AS sqlite3_int64) AS LONG
   DECLARE FUNCTION Bytes () AS LONG
   DECLARE FUNCTION Read (BYVAL pBuffer AS ANY PTR, BYVAL numBytes AS LONG, BYVAL nOffset AS LONG = 0) AS LONG
   DECLARE FUNCTION Write (BYVAL pBuffer AS ANY PTR, BYVAL numBytes AS LONG, BYVAL nOffset AS LONG = 0) AS LONG

END TYPE
' ########################################################################################

' ===========================================================================================
' Constructor
' ===========================================================================================
PRIVATE CONSTRUCTOR CSQLiteBlob (BYVAL pBlob AS sqlite3_blob PTR)
   m_hBlob = pBlob
END CONSTRUCTOR
' ===========================================================================================

' ===========================================================================================
' Destructor
' ===========================================================================================
PRIVATE DESTRUCTOR CSQLiteBlob
   this.Close   ' // Close the blob handle
END DESTRUCTOR
' ===========================================================================================

' ===========================================================================================
' Gets the blob handle
' ===========================================================================================
PRIVATE PROPERTY CSQLiteBlob.hBlob () AS sqlite3_blob PTR
   PROPERTY = m_hBlob
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Sets the blob handle
' ===========================================================================================
PRIVATE PROPERTY CSQLiteBlob.hBlob (BYVAL pBlob AS sqlite3_blob PTR)
   this.Close   ' // Close the blob handle
   m_hBlob = pBlob
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Closes the blob handle. The BLOB handle is closed unconditionally. Even if the call returns
' an error code, the handle is still closed. If the blob handle being closed was opened for
' read-write access, and if the database is in auto-commit mode and there are no other open
' read-write blob handles or active write statements, the current transaction is committed.
' If an error occurs while committing the transaction, an error code is returned and the
' transaction rolled back.
' This method only works on a BLOB handle which has been created by a prior successful call
' to OpenBlob and which has not been closed. Passing any other pointer in to this routine
' results in undefined and probably undesirable behavior.
' ===========================================================================================
PRIVATE FUNCTION CSQLiteBlob.Close () AS LONG
   DIM pProc AS PFNSQLITE3BLOBCLOSEPROC = _
   cast(PFNSQLITE3BLOBCLOSEPROC, GetProcAddress(m_hLib, "sqlite3_blob_close"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   FUNCTION = pProc(m_hBlob)
   m_hBlob = NULL
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Moves an existing BLOB handle so that it points to a different row of the same database
' table. The new row is identified by the rowid value passed as the second argument.
' ===========================================================================================
PRIVATE FUNCTION CSQLiteBlob.Reopen (BYVAL rowId AS sqlite3_int64) AS LONG
   DIM pProc AS PFNSQLITE3BLOBREOPENPROC = _
   cast(PFNSQLITE3BLOBREOPENPROC, GetProcAddress(m_hLib, "sqlite3_blob_reopen"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   FUNCTION = pProc(m_hBlob, rowId)
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Returns the size in bytes of the BLOB accessible via the successfully opened BLOB handle
' in its only argument. The incremental blob I/O routines can only read or overwriting existing
' blob content; they cannot change the size of a blob.
' This method only works on a BLOB handle which has been created by a prior successful call
' to OpenBlob and which has not been closed. Passing any other pointer in to this routine
' results in undefined and probably undesirable behavior.
' ===========================================================================================
PRIVATE FUNCTION CSQLiteBlob.Bytes () AS LONG
   DIM pProc AS PFNSQLITE3BLOBBYTESPROC = _
   cast(PFNSQLITE3BLOBBYTESPROC, GetProcAddress(m_hLib, "sqlite3_blob_bytes"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   FUNCTION = pProc(m_hBlob)
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' This method is used to read data from an open BLOB handle into a caller-supplied buffer.
' numBytes bytes of data are copied into buffer pBuffer from the open BLOB, starting at offset
' iOffset. If offset iOffset is less than numBytes bytes from the end of the BLOB, SQLITE_ERROR
' is returned and no data is read. If numBytes or iOffset is less than zero, SQLITE_ERROR is
' returned and no data is read. The size of the blob (and hence the maximum value of
' numBytes+iOffset) can be determined using the Bytes method. An attempt to read from an
' expired BLOB handle fails with an error code of SQLITE_ABORT. On success, Read returns
' SQLITE_OK. Otherwise, an error code or an extended error code is returned.
' This method only works on a BLOB handle which has been created by a prior successful call
' to OpenBlob and which has not been closed. Passing any other pointer in to this routine
' results in undefined and probably undesirable behavior.
' ===========================================================================================
PRIVATE FUNCTION CSQLiteBlob.Read (BYVAL pBuffer AS ANY PTR, BYVAL numBytes AS LONG, BYVAL iOffset AS LONG = 0) AS LONG
   DIM pProc AS PFNSQLITE3BLOBREADPROC = _
   cast(PFNSQLITE3BLOBREADPROC, GetProcAddress(m_hLib, "sqlite3_blob_read"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   FUNCTION = pProc(m_hBlob, pBuffer, numBytes, iOffset)
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' This function is used to write data into an open BLOB handle from a caller-supplied buffer.
' numBytes bytes of data are copied from the buffer pBuffer into the open BLOB, starting at
' offset iOffset. On success, Write returns SQLITE_OK. Otherwise, an error code or an extended
' error code is returned. Unless SQLITE_MISUSE is returned, this function sets the database
' connection error code and message accessible ErrCode and ErrMsg and related functions.
' If the BLOB handle passed as the first argument was not opened for writing (the flags
' parameter to OpenBlob was zero), this function returns SQLITE_READONLY.
' This method may only modify the contents of the BLOB; it is not possible to increase the
' size of a BLOB using it. If offset iOffset is less than numBytes bytes from the end of the
' BLOB, SQLITE_ERROR is returned and no data is written. The size of the BLOB (and hence the
' maximum value of numBytes+iOffset) can be determined using the Bytes method. If numBytes or
' iOffset are less than zero SQLITE_ERROR is returned and no data is written.
' An attempt to write to an expired BLOB handle fails with an error code of SQLITE_ABORT.
' Writes to the BLOB that occurred before the BLOB handle expired are not rolled back by the
' expiration of the handle, though of course those changes might have been overwritten by the
' statement that expired the BLOB handle or by other independent statements.
' This method only works on a BLOB handle which has been created by a prior successful call
' to OpenBlob and which has not been closed. Passing any other pointer in to this routine
' results in undefined and probably undesirable behavior.
' ===========================================================================================
PRIVATE FUNCTION CSQLiteBlob.Write (BYVAL pBuffer AS ANY PTR, BYVAL numBytes AS LONG, BYVAL iOffset AS LONG = 0) AS LONG
   DIM pProc AS PFNSQLITE3BLOBWRITEPROC = _
   cast(PFNSQLITE3BLOBWRITEPROC, GetProcAddress(m_hLib, "sqlite3_blob_write"))
   IF pProc = 0 THEN this.SetResult(GetLastError) : EXIT FUNCTION
   FUNCTION = pProc(m_hBlob, pBuffer, numBytes, iOffset)
END FUNCTION
' ===========================================================================================

END NAMESPACE

' TODO (or maybe not, because it doesn't allow to use unicode for the database name)
'declare function sqlite3_backup_init(byval pDest as sqlite3 ptr, byval zDestName as const zstring ptr, byval pSource as sqlite3 ptr, byval zSourceName as const zstring ptr) as sqlite3_backup ptr
'declare function sqlite3_backup_step(byval p as sqlite3_backup ptr, byval nPage as long) as long
'declare function sqlite3_backup_finish(byval p as sqlite3_backup ptr) as long
'declare function sqlite3_backup_remaining(byval p as sqlite3_backup ptr) as long
'declare function sqlite3_backup_pagecount(byval p as sqlite3_backup ptr) as long
