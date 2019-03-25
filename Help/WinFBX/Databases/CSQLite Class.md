# CSQLite Class

SQLite is an embedded SQL database engine. Unlike most other SQL databases, SQLite does not have a separate server process. SQLite reads and writes directly to ordinary disk files. A complete SQL database with multiple tables, indices, triggers, and views, is contained in a single disk file. The database file format is cross-platform - you can freely copy a database between 32-bit and 64-bit systems or between big-endian and little-endian architectures. These features make SQLite a popular choice as an Application File Format.

**Include file**: CSQLite.inc

### Constructor

Creates a new instance of the **CSQlite** base class and loads the SQLite3 DLL.

```
CONSTRUCTOR CSQLite (BYREF wszDllPath AS WSTRING = "sqlite3.dll")
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszDllPath* | The path of the SQLite3 DLL. |

### Methods

| Name       | Description |
| ---------- | ----------- |
| [CompileOptionUsed](#CompileOptionUsed) | Determines whether the specified option was defined at compile time. |
| [Complete](#Complete) | Determines if the currently entered text seems to form a complete SQL statement. |
| [EnableSharedCache](#EnableSharedCache) | Enables or disables the sharing of the database cache and schema data structures between connections to the same database. |
| [ErrStr](#ErrStr) | Returns English-language text that describes the result code. |
| [Free](#Free) | Releases memory previously allocated by **Malloc** or **Realloc**. |
| [GetCompileOption](#GetCompileOption) | Allows iterating over the list of options that were defined at compile time by returning the N-th compile time option string. |
| [GetLastResult](#GetLastResult) | Returns the last result code. |
| [Malloc](#Malloc) | Returns a pointer to a block of memory at least N bytes in length, where N is the parameter. |
| [MemoryHighwater](#MemoryHighwater) | Returns the maximum value of **MemoryUsed** since the high-water mark was last reset. |
| [MemorySize](#MemorySize) | Returns the size of that memory allocation in bytes. |
| [MemoryUsed](#MemoryUsed) | Returns the number of bytes of memory currently outstanding (malloced but not freed). |
| [Randomness](#Randomness) | Pseudo-random number generator. |
| [Realloc](#Realloc) | Attempts to resize a prior memory allocation to be at least the specfified number of bytes. |
| [ReleaseMemory](#ReleaseMemory) | Attempts to free the specified number of bytes of heap memory by deallocating non-essential memory allocations held by the database library. |
| [Sleep](#Sleep) | Causes the current thread to suspend execution for at least a number of milliseconds specified in its parameter. |
| [SoftHeapLimit64](#SoftHeapLimit64) | Sets and/or queries the soft limit on the amount of heap memory that may be allocated by SQLite. |
| [SourceID](#SourceID) | Returns the SQLite3 source identifier. |
| [Status](#Status) | Retrieves runtime status information about the performance of SQLite, and optionally to reset various highwater marks. |
| [StrGlob](#StrGlob) | The **StrGlob** method returns zero if string *szStr* matches the glob pattern *szGlob*, and it returns non-zero if string *szStr* does not match the glob pattern *szGlob*. This function is case sensitive. |
| [ThreadSafe](#ThreadSafe) | Returns zero if and only if SQLite was compiled with mutexing code omitted due to the SQLITE_THREADSAFE compile-time option being set to 0. |
| [Version](#Version) | Returns the SQLite3 version. |
| [VersionNumber](#VersionNumber) | Returns the SQLite3 version number. |

# CSQLiteDb Class

Inherits from CSQLite.

### Constructor

Opens an SQLite database file as specified by the filename argument.

```
CONSTRUCTOR CSQLiteDb (BYREF wszPath AS WSTRING)
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | The name of the database. |

#### Remarks

The database is opened for reading and writing, and is created if it does not already exist. If the filename is ":memory:", then a private, temporary in-memory database is created for the connection. This in-memory database will vanish when the database connection is closed. If the filename is an empty string, then a private, temporary on-disk database will be created. This private database will be automatically deleted as soon as the database connection is closed. If URI filename interpretation is enabled, and the filename argument begins with "file:", then the filename is interpreted as a URI.

### Methods

| Name       | Description |
| ---------- | ----------- |
| [Changes](#Changes) | Returns the number of database rows that were changed or inserted or deleted by the most recently completed SQL statement on the current database connection. |
| [CloseDb](#CloseDb) | Closes the database. |
| [ErrCode](#ErrCode) | Returns the numeric result code for the most recent failed sqlite3 call associated with a database connection. |
| [ErrMsg](#ErrMsg) | Returns English-language text that describes the error. |
| [Exec](#Exec) | Convenience wrapper for **Prepare** and **Step_**. |
| [ExtendedErrCode](#ExtendedErrCode) | Gets the extended error code associated with this dabatase connection. |
| [ExtendedResultCodes](#ExtendedResultCodes) | Enables or disables the extended result codes feature of SQLite. |
| [hDbc](#hDbc) | Gets/sets the database handle. |
| [Interrupt](#Interrupt) | This function causes any pending database operation to abort and return at its earliest opportunity. |
| [LastInsertRowId](#LastInsertRowId) | Returns the rowid of the most recent successful INSERT into the database from the database connection in the first argument. |
| [Limit](#Limit) | This function allows the size of various constructs to be limited on a connection by connection basis. |
| [OpenBlob](#OpenBlob) | Opens a handle to a BLOB. |
| [OpenDb](#OpenDb) | Opens an SQLite database file as specified by the filename argument. |
| [Prepare](#Prepare) | Creates a new prepared statement object. |
| [ProgressHandler](#ProgressHandler) | The **ProgressHandler** method causes a callback function to be invoked periodically during long running calls to **Step_** and **GetRow** for a database connection. |
| [ReleaseMemory](#ReleaseMemory2) | Attempts to free as much heap memory as possible from the specified database connection. |
| [Status](#Status2) | Retrieves runtime status information about a single database connection. |
| [TotalChanges](#TotalChanges) | This function returns the number of row changes caused by INSERT, UPDATE or DELETE statements since the database connection was opened. |
| [UnlockNotify](#UnlockNotify) | Registers a callback that SQLite will invoke when the connection currently holding the required lock relinquishes it. |

# CSQLiteStmt Class

Wraps a sqlite3_stmt pointer returned by the **Prepare** method of the **CSQLite** class.

Inherits from CSQLite.

```
CONSTRUCTOR CSQLiteStmt (BYVAL pStmt AS sqlite3_stmt PTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pStmt* | The sqlite3_stmt pointer. |

### Methods and Properties

| Name       | Description |
| ---------- | ----------- |
| [BindBlob](#BindBlob) | Binds a blob with the statement. |
| [BindDouble](#BindDouble) | Binds a double value with the statement. |
| [BindLong](#BindLong) | Binds a long value with the statement. |
| [BindLongInt](#BindLongInt) | Binds a longint value with the statement. |
| [BindNull](#BindNull) | Binds a null value with the statement. |
| [BindParameterCount](#BindParameterCount) | This method can be used to find the number of SQL parameters in a prepared statement. |
| [BindParameterIndex](#BindParameterIndex) | Returns the index of an SQL parameter given its name. |
| [BindParameterName](#BindParameterName) | Returns the name of the N-th SQL parameter in the prepared statement P. |
| [BindText](#BindText) | Binds a text value with the statement. |
| [BindZeroBlob](#BindZeroBlob) | Binds a BLOB that is filled with zeroes. |
| [Busy](#Busy) | Returns true if the prepared statement has been stepped at least once using **Step_** but has not run to completion and/or has not been reset using **Reset**. |
| [ClearBindings](#ClearBindings) | Sets all the parameters in the compiled SQL statement to NULL. |
| [ColumnBlob](#ColumnBlob) | Returns information about a single column of the current result row of a query. |
| [ColumnBytes](#ColumnBytes) | Returns the number of bytes of the column value. |
| [ColumnCount](#ColumnCount) | Returns the number of columns in the result set returned by the prepared statement. |
| [ColumnDatabaseName](#ColumnDatabaseName) | Returns the database name that is the origin of a particular result column in SELECT statement. |
| [ColumnDeclaredType](#ColumnDeclaredType) | Returns the declared data type of a query result. |
| [ColumnDouble](#ColumnDouble) | Returns the column value as a double. |
| [ColumnLong](#ColumnLong) | Returns the column value as a long. |
| [ColumnLongInt](#ColumnLongInt) | Returns the column value as a quad. |
| [ColumnName](#ColumnName) | Returns the name assigned to a particular column in the result set of a SELECT statement. |
| [ColumnOriginName](#ColumnOriginName) | Returns the column name that is the origin of a particular result column in SELECT statement. |
| [ColumnTableName](#ColumnTableName) | Returns the table name that is the origin of a particular result column in SELECT statement. |
| [ColumnText](#ColumnText) | Returns the column value as a UTF-16 string. |
| [ColumnType](#ColumnType) | Returns the column type. |
| [DataCount](#DataCount) | Returns the number of columns in the result set returned by the prepared statement. |
| [DbHandle](#DbHandle) | Returns the database connection handle to which a prepared statement belongs. |
| [Finalize](#Finalize) | Deletes a prepared statement. |
| [GetRow](#GetRow) | After a prepared statement has been prepared using either Prepare this method must be called one or more times to evaluate the statement. **GetRow** is an alias for **Step_**. |
| [hStmt](#hStmt) | Gets/sets the connection handle. |
| [IsColumnNull](#IsColumnNull) | Returns true is the column value is null or false otherwise. |
| [ReadOnly](#ReadOnly) | Rreturns true if and only if the prepared statement makes no direct changes to the content of the database file. |
| [Reset](#Reset) | Resets a prepared statement object back to its initial state, ready to be re-executed. |
| [Sql](#Sql) | Retrieve a saved copy of the original SQL text used to create a prepared statement if that statement was compiled using **Prepare**. |
| [Step_](#Step_) | After a prepared statement has been prepared using **Prepare** this method must be called one or more times to evaluate the statement. |

# <a name="CompileOptionUsed"></a>CompileOptionUsed

Returns FALSE or TRUE indicating whether the specified option was defined at compile time. The SQLITE_ prefix may be omitted from the option name passed to **CompileOptionUsed**.

```
FUNCTION CompileOptionUsed (BYREF szOptName AS ZSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *szOptName* | The option name. |

### Usage example

```
pSql.CompileOptionUsed("SQLITE_ENABLE_DBSTAT_VTAB")
```

# <a name="Complete"></a>Complete

Determines if the currently entered text seems to form a complete SQL statement or if additional input is needed before sending the text into SQLite for parsing. 

```
FUNCTION Complete (BYREF wszSql AS CONST WSTRING) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszSql* | The SQL string to check. |

#### Return value

Returns 0 if the statement is incomplete or 1 if the input string appears to be a complete SQL statement. If a memory allocation fails, then SQLITE_NOMEM is returned.

#### Remarks

This method is useful during command-line input to determine if the currently entered text seems to form a complete SQL statement or if additional input is needed before sending the text into SQLite for parsing. This method returns 1 if the input string appears to be a complete SQL statement. A statement is judged to be complete if it ends with a semicolon token and is not a prefix of a well-formed CREATE TRIGGER statement. Semicolons that are embedded within string literals or quoted identifier names or comments are not independent tokens (they are part of the token in which they are embedded) and thus do not count as a statement terminator. Whitespace and comments that follow the final semicolon are ignored.

This method does not parse the SQL statements thus will not detect syntactically incorrect SQL.

# <a name="EnableSharedCache"></a>EnableSharedCache

Enables or disables the sharing of the database cache and schema data structures between connections to the same database. Sharing is enabled if the argument is true and disabled if the argument is false.. 

```
FUNCTION EnableSharedCache (BYVAL bSharing AS BOOLEAN) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *bSharing* | TRUE or FALSE. |

#### Return value

Returns SQLITE_OK if shared cache was enabled or disabled successfully. An error code is returned otherwise.

# <a name="ErrStr"></a>ErrStr

Returns English-language text that describes the result code.

```
FUNCTION ErrStr (BYVAL nErrorCode AS LONG) AS STRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *nErrorCode* | The result code. |

#### Return value

A string containing a description of the error.

# <a name="Free"></a>Free

Releases memory previously allocated by Malloc or Realloc.

```
SUB Free (BYVAL pMem AS ANY PTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pMem* | The pointer returned by Malloc or Realloc. |

# <a name="GetCompileOption"></a>GetCompileOption

Allows iterating over the list of options that were defined at compile time by returning the N-th compile time option string. If nOption is out of range, **GetCompileOption** returns a NULL pointer. The SQLITE_ prefix is omitted from any strings returned by **GetCompileOption**.

```
FUNCTION GetCompileOption (BYVAL nOption AS LONG) AS STRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *nOption* | The option number. |

#### Uage example

```
pSql.GetCompileOption(0)
```

# <a name="GetLastResult"></a>GetLastResult

Returns the last result code.

```
FUNCTION GetLastResult () AS LONG
```

#### Return value

The result code returned by the last executed method.

#### Remarks

The last result code is not global, but bound to each object, and its purpose is to check the failure or success of the last SQLite operation. To get descriptive information about SQLite errors call **ErrStr**.

# <a name="Malloc"></a>Malloc

Returns a pointer to a block of memory at least N bytes in length, where N is the parameter. If Malloc is unable to obtain sufficient free memory, it returns a NULL pointer. If the parameter *nBytes* to Malloc is zero or negative then returns a NULL pointer.

```
FUNCTION Malloc (BYVAL nBytes AS LONG) AS ANY PTR
FUNCTION Malloc64 (BYVAL nBytes AS sqlite3_uint64) AS ANY PTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *nBytes* | The number of bytes to allocate. |

#### Return value

Pointer to the allocated memory.

#### Remarks

The memory returned by **Malloc** and **Realloc** is always aligned to at least an 8 byte boundary, or to a 4 byte boundary if the SQLITE_4_BYTE_ALIGNED_MALLOC compile-time option is used.

In SQLite version 3.5.0 and 3.5.1, it was possible to define the SQLITE_OMIT_MEMORY_ALLOCATION which would cause the built-in implementation of these functions to be omitted. That capability is no longer provided. Only built-in memory allocators can be used.

Prior to SQLite version 3.7.10, the Windows OS interface layer called the system **malloc** and **free** directly when converting filenames between the UTF-8 encoding used by SQLite and whatever filename encoding is used by the particular Windows installation. Memory allocation errors were detected, but they were reported back as SQLITE_CANTOPEN or SQLITE_IOERR rather than SQLITE_NOMEM.

The pointer arguments to **Free** and **Realloc** must be either NULL or else pointers obtained from a prior invocation of Malloc or **Realloc** that have not yet been released.

The application must not read or write any part of a block of memory after it has been released using **Free** or **Realloc**. 

# <a name="MemoryHighwater"></a>MemoryHighwater

Returns the maximum value of **MemoryUsed** since the high-water mark was last reset.

```
FUNCTION MemoryHighwater (BYVAL resetFlag AS BOOLEAN) AS sqlite3_uint64
```

| Parameter  | Description |
| ---------- | ----------- |
| *resetFlag* | TRUE or FALSE. Reset the high-water mark. |

#### Return value

The number of bytes of memory currently outstanding (malloced but not freed) since the high-water mark was last reset.

#### Remarks

The value returned by **MemoryHighwater** include any overhead added by SQLite in its implementation of **Malloc**, but not overhead added by the any underlying system library functions that **Malloc** may call.

# <a name="MemorySize"></a>MemorySize

Returns the size of that memory allocation in bytes.

```
FUNCTION MemorySize (BYVAL pMem AS ANY PTR) AS sqlite3_uint64
```

| Parameter  | Description |
| ---------- | ----------- |
| *pMem* | Pointer to the memory allocation. |

#### Return value

If *pMem* is a memory allocation previously obtained from **Malloc**, **Malloc64**, **Realloc** or **Realloc64**, then **MemorySize** returns the size of that memory allocation in bytes. The value returned by **MemorySize** might be larger than the number of bytes requested when pMem was allocated. If *pMem* is a NULL pointer then **MemorySize** returns zero. If *pMem* points to something that is not the beginning of memory allocation, or if it points to a formerly valid memory allocation that has now been freed, then the behavior of **MemorySize** is undefined and possibly harmful.

# <a name="MemoryUsed"></a>MemoryUsed

Returns the number of bytes of memory currently outstanding (malloced but not freed).

```
FUNCTION MemoryUsed () AS sqlite3_uint64
```

#### Return value

The number of bytes of memory currently outstanding (malloced but not freed).

#### Remarks

The value returned by **MemoryUsed** include any overhead added by SQLite in its implementation of **Malloc**, but not overhead added by the any underlying system library functions that **Malloc** may call.

# <a name="Randomness"></a>Randomness

Pseudo-random number generator.

```
SUB Randomness (BYVAL nBytes AS LONG, BYVAL pbuffer AS ANY PTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *nBytes* | The number of bytes of randomness to return. |
| *pbuffer* | Pointer where to store the random bytes. It can be null. |

#### Remarks

SQLite contains a high-quality pseudo-random number generator (PRNG) used to select random ROWIDs when inserting new records into a table that already uses the largest possible ROWID. The PRNG is also used for the build-in random() and randomblob() SQL functions. This method allows applications to access the same PRNG for other purposes.

# <a name="Realloc"></a>Realloc

Attempts to resize a prior memory allocation to be at least *nBytes* bytes. *pMem* is the memory allocation to be resized.

```
FUNCTION Realloc (BYVAL pMem AS ANY PTR, BYVAL nBytes AS LONG) AS ANY PTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *pMem* | Pointer to memory allocated by **Malloc**. |
| *nBytes* | The number of bytes to reallocate. |

#### Return value

Pointer to the reallocated memory.

#### Remarks

The memory returned by **Malloc** and **Realloc** is always aligned to at least an 8 byte boundary, or to a 4 byte boundary if the SQLITE_4_BYTE_ALIGNED_MALLOC compile-time option is used.

In SQLite version 3.5.0 and 3.5.1, it was possible to define the SQLITE_OMIT_MEMORY_ALLOCATION which would cause the built-in implementation of these functions to be omitted. That capability is no longer provided. Only built-in memory allocators can be used.

Prior to SQLite version 3.7.10, the Windows OS interface layer called the system **malloc** and **free** directly when converting filenames between the UTF-8 encoding used by SQLite and whatever filename encoding is used by the particular Windows installation. Memory allocation errors were detected, but they were reported back as SQLITE_CANTOPEN or SQLITE_IOERR rather than SQLITE_NOMEM.

The pointer arguments to **Free** and **Realloc** must be either NULL or else pointers obtained from a prior invocation of **Malloc** or **Realloc** that have not yet been released.

The application must not read or write any part of a block of memory after it has been released using **Free** or **Realloc**. 

# <a name="ReleaseMemory"></a>ReleaseMemory

Attempts to free the specified number of bytes of heap memory by deallocating non-essential memory allocations held by the database library. Memory used to cache database pages to improve performance is an example of non-essential memory. **ReleaseMemory** returns the number of bytes actually freed, which might be more or less than the amount requested. The **ReleaseMemory** method is a no-op returning zero if SQLite is not compiled with SQLITE_ENABLE_MEMORY_MANAGEMENT.

```
FUNCTION ReleaseMemory (BYVAL nBytes AS LONG) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *nBytes* | The number of bytes to free. |

#### Return value

The number of bytes freed.

# <a name="Sleep"></a>Sleep

Causes the current thread to suspend execution for at least a number of milliseconds specified in its parameter.

```
FUNCTION Sleep (BYVAL ms AS LONG) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *ms* | The number of milliseconds to sleep. |

#### Return value

The number of milliseconds of sleep actually requested from the operating system is returned.

#### Remarks

If the operating system does not support sleep requests with millisecond time resolution, then the time will be rounded up to the nearest second. The number of milliseconds of sleep actually requested from the operating system is returned.


# <a name="SoftHeapLimit64"></a>SoftHeapLimit64

Sets and/or queries the soft limit on the amount of heap memory that may be allocated by SQLite.

```
FUNCTION SoftHeapLimit64 (BYVAL nBytes AS LONG) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *nBytes* | The number of bytes. If *nBytes* is negative then no change is made to the soft heap limit. Hence, the current size of the soft heap limit can be determined by invoking **SoftHeapLimit64** with a negative argument. If *nBytes* is zero then the soft heap limit is disabled. |

#### Return value

The size of the soft heap limit prior to the call, or negative in the case of an error.

#### Remarks

SQLite strives to keep heap memory utilization below the soft heap limit by reducing the number of pages held in the page cache as heap memory usages approaches the limit. The soft heap limit is "soft" because even though SQLite strives to stay below the limit, it will exceed the limit rather than generate an SQLITE_NOMEM error. In other words, the soft heap limit is advisory only.

The circumstances under which SQLite will enforce the soft heap limit may change in future releases of SQLite.

# <a name="SourceID"></a>SourceID

Returns the SQLite3 source identifier.

```
FUNCTION SourceID () AS STRING
```

#### Return value

A string containing the SQLite3 identifier, e.g. "2012-06-11 02:05:22 f5b5a13f7394dc143aa136f1d4faba6839eaa6dc".

# <a name="Status"></a>Status

Retrieves runtime status information about the performance of SQLite, and optionally to reset various highwater marks.

```
FUNCTION Status (BYVAL op AS LONG, BYREF pCurrent AS LONG, BYREF pHighwater AS LONG, _
   BYVAL resetFlag AS BOOLEAN = FALSE) AS LONG
FUNCTION Status64 (BYVAL op AS LONG, BYREF pCurrent AS sqlite3_int64, BYREF pHighwater AS sqlite3_int64, _
   BYVAL resetFlag AS BOOLEAN = FALSE) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *op* | An integer code for a specific satus parameter. |
| *pCurrent* | Pointer to a variable that receives the current value of the parameter. |
| *pHighWater* | Pointer to a variable that receives the highest recorded value of the parameter. |
| *resetFlag* | TRUE or FALSE. If true, then the counter is reset to zero after this interface call returns. |

#### Return value

Returns SQLITE_OK (0) on success and a non-zero error code on failure.

#### Remarks

This method is used to retrieve runtime status information about the performance of SQLite, and optionally to reset various highwater marks. The first argument is an integer code for the specific parameter to measure. Recognized integer codes are of the form SQLITE_STATUS_.... The current value of the parameter is returned into *pCurrent*. The highest recorded value is returned in *pHighwater*. If the *resetFlag* is true, then the highest record value is reset after *pHighwater* is written. Some parameters do not record the highest value. For those parameters nothing is written into *pHighwater* and the *resetFlag* is ignored. Other parameters record only the highwater mark and not the current value. For these latter parameters nothing is written into *pCurrent*.

This function is threadsafe but is not atomic. This function can be called while other threads are running the same or different SQLite interfaces. However the values returned in pCurrent and pHighwater reflect the status of SQLite at different points in time and it is possible that another thread might change the parameter in between the times when *pCurrent* and *pHighwater* are written.

# <a name="StrGlob"></a>StrGlob

The *StrGlob* method returns zero if string *szStr* matches the glob pattern *szGlob*, and it returns non-zero if string *szStr* does not match the glob pattern *szGlob*. This function is case sensitive.

```
FUNCTION StrGlob (BYREF szGlob AS ZSTRING, BYREF szStr AS ZSTRING) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *szGlob* | The glob pattern to compare. |
| *szStr* | The string to compare. |

# <a name="ThreadSafe"></a>ThreadSafe

Returns zero if and only if SQLite was compiled with mutexing code omitted due to the SQLITE_THREADSAFE compile-time option being set to 0.

```
FUNCTION ThreadSafe () AS BOOLEAN
```

#### Return value

TRUE or FALSE.

#### Remarks

SQLite can be compiled with or without mutexes. When the SQLITE_THREADSAFE C preprocessor macro is 1 or 2, mutexes are enabled and SQLite is threadsafe. When the SQLITE_THREADSAFE macro is 0, the mutexes are omitted. Without the mutexes, it is not safe to use SQLite concurrently from more than one thread.

Enabling mutexes incurs a measurable performance penalty. So if speed is of utmost importance, it makes sense to disable the mutexes. But for maximum safety, mutexes should be enabled. The default behavior is for mutexes to be enabled.

This interface can be used by an application to make sure that the version of SQLite that it is linking against was compiled with the desired setting of the SQLITE_THREADSAFE macro.

# <a name="Version"></a>Version

Returns the SQLite3 version.

```
FUNCTION Version () AS STRING
```

#### Return value

A string containing the SQLite3 version, e.g. "3.7.13".

# <a name="VersionNumber"></a>VersionNumber

Returns the SQLite3 version number.

```
FUNCTION VersionNumber () AS LONG
```

#### Return value

A long integer containing the SQLite3 version number, e.g. 3007013.

# <a name="Changes"></a>Changes

This method returns the number of database rows that were changed or inserted or deleted by the most recently completed SQL statement on the current database connection.

```
FUNCTION Changes () AS LONG
```

#### Return value

The number of changes.

#### ks

Only changes that are directly specified by the INSERT, UPDATE, or DELETE statement are counted. Auxiliary changes caused by triggers or foreign key actions are not counted. Use the **Changes** method to find the total number of changes including changes caused by triggers and foreign key actions.

Changes to a view that are simulated by an INSTEAD OF trigger are not counted. Only real table changes are counted.

A "row change" is a change to a single row of a single table caused by an INSERT, DELETE, or UPDATE statement. Rows that are changed as side effects of REPLACE constraint resolution, rollback, ABORT processing, DROP TABLE, or by any other mechanisms do not count as direct row changes.

A "trigger context" is a scope of execution that begins and ends with the script of a trigger. Most SQL statements are evaluated outside of any trigger. This is the "top level" trigger context. If a trigger fires from the top level, a new trigger context is entered for the duration of that one trigger. Subtriggers create subcontexts for their duration.

Calling **Exec** or **Step_** recursively does not create a new trigger context.

This method returns the number of direct row changes in the most recent INSERT, UPDATE, or DELETE statement within the same trigger context.

Thus, when called from the top level, this function returns the number of changes in the most recent INSERT, UPDATE, or DELETE that also occurred at the top level. Within the body of a trigger, the **Changes** method can be called to find the number of changes in the most recently completed INSERT, UPDATE, or DELETE statement within the body of the same trigger. However, the number returned does not include changes caused by subtriggers since those have their own context.

If a separate thread makes changes on the same database connection while **Changes** is running then the value returned is unpredictable and not meaningful. 

# <a name="CloseDb"></a>CloseDb

Closes the database.

```
FUNCTION CloseDb () AS LONG
```

#### Return value

SQLITE_OK if the sqlite3 object is successfully destroyed and all associated resources are deallocated.

#### Remarks

Applications must finalize all prepared statements and close all BLOB handles associated with the sqlite3 object prior to attempting to close the object. If **CloseDb** is called on a database connection that still has outstanding prepared statements or BLOB handles, then it returns SQLITE_BUSY.

If **CloseDb** is invoked while a transaction is open, the transaction is automatically rolled back.

# <a name="ErrCode"></a>ErrCode

Returns the numeric result code for the most recent failed sqlite3 call associated with a database connection. If a prior API call failed but the most recent API call succeeded, the return value from **ErrCode** is undefined.

```
FUNCTION ErrCode () AS LONG
```

#### Return value

The error code.

#### Remarks

When the serialized threading mode is in use, it might be the case that a second error occurs on a separate thread in between the time of the first error and the call to these functions. When that happens, the second error will be reported since these functions always report the most recent result. To avoid this, each thread can obtain exclusive use of the database connection by invoking **sqlite3_mutex_enter**(*sqlite3_db_mutex(D)*) before beginning to use D and invoking **sqlite3_mutex_leave**(*sqlite3_db_mutex(D)*) after all calls to the functions listed here are completed.

If a function fails with SQLITE_MISUSE, that means the function was invoked incorrectly by the application. In that case, the error code and message may or may not be set. 

# <a name="ErrMsg"></a>ErrMsg

Returns English-language text that describes the error.

```
FUNCTION ErrMsg () AS STRING
```

#### Return value

A string containing a description of the error.

#### Remarks

When the serialized threading mode is in use, it might be the case that a second error occurs on a separate thread in between the time of the first error and the call to these functions. When that happens, the second error will be reported since these functions always report the most recent result. To avoid this, each thread can obtain exclusive use of the database connection by invoking **sqlite3_mutex_enter**(*sqlite3_db_mutex(D)*) before beginning to use D and invoking **sqlite3_mutex_leave**(*sqlite3_db_mutex(D)*) after all calls to the functions listed here are completed.

If a function fails with SQLITE_MISUSE, that means the function was invoked incorrectly by the application. In that case, the error code and message may or may not be set. 

# <a name="Exec"></a>Exec

Convenience wrapper for **Prepare** and **Step_**. To be used with queries that don't return result sets, such CREATE, UPDATE and INSERT.

```
FUNCTION Exec (BYREF wszSql AS WSTRING) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszSql* | The SQL statement. |

#### Return value

SQLITE_BUSY means that the database engine was unable to acquire the database locks it needs to do its job. If the statement is a COMMIT or occurs outside of an explicit transaction, then you can retry the statement. If the statement is not a COMMIT and occurs within an explicit transaction then you should rollback the transaction before continuing.

SQLITE_DONE means that the statement has finished executing successfully. **Step_** should not be called again on this virtual machine without first calling Reset to reset the virtual machine back to its initial state.

SQLITE_ERROR means that a run-time error (such as a constraint violation) has occurred. Step should not be called again on the cirtual machine. More information may be found by calling **ErrMsg**.

# <a name="ExtendedErrCode"></a>ExtendedErrCode

Gets the extended error code associated with this dabatase connection.

```
FUNCTION ExtendedErrCode () AS LONG
```

# <a name="ExtendedResultCodes"></a>ExtendedResultCodes

Enables or disables the extended result codes feature of SQLite. The extended result codes are disabled by default for historical compatibility.

```
FUNCTION ExtendedResultCodes (BYVAL onoff AS BOOLEAN) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *onoff* | TRUE to enable; FALSE to disable. |

#### Return value

SQLITE_OK (0) or an error code.

# <a name="hDbc"></a>hDbc

Gets/sets the database handle.

```
PROPERTY hDbc () AS sqlite3 PTR
PROPERTY hDbc (pDbc AS sqlite3 PTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *hDbc* | The database handle. |

#### Return value

The database handle.

# <a name="Interrupt"></a>Interrupt

This function causes any pending database operation to abort and return at its earliest opportunity. This function is typically called in response to a user action such as pressing "Cancel" or Ctrl-C where the user wants a long query operation to halt immediately.

```
SUB Interrupt
```

#### Remarks

It is safe to call this function from a thread different from the thread that is currently running the database operation. But it is not safe to call this function with a database connection that is closed or might close before **Interrupt** returns.

If an SQL operation is very nearly finished at the time when **Interrupt** is called, then it might not have an opportunity to be interrupted and might continue to completion.

An SQL operation that is interrupted will return SQLITE_INTERRUPT. If the interrupted SQL operation is an INSERT, UPDATE, or DELETE that is inside an explicit transaction, then the entire transaction will be rolled back automatically.

The **Interrupt** call is in effect until all currently running SQL statements on database connection complete. Any new SQL statements that are started after the **Interrupt** call and before the running statements reaches zero are interrupted as if they had been running prior to the **Interrupt** call. New SQL statements that are started after the running statement count reaches zero are not effected by the **Interrupt**. A call to **Interrupt** that occurs when there are no running SQL statements is a no-op and has no effect on SQL statements that are started after the **Interrupt** call returns.

If the database connection closes while **Interrupt** is running then bad things will likely happen.

# <a name="LastInsertRowId"></a>LastInsertRowId

Returns the rowid of the most recent successful INSERT into the database from the database connection in the first argument.

```
FUNCTION LastInsertRowId () AS sqlite3_int64
```

#### Return value

The rowid of the most recent successful INSERT into the database from the database connection in the first argument.

#### Remarks

Each entry in an SQLite table has a unique 64-bit signed integer key called the "rowid". The rowid is always available as an undeclared column named ROWID, OID, or _ROWID_ as long as those names are not also used by explicitly declared columns. If the table has a column of type INTEGER PRIMARY KEY then that column is another alias for the rowid.

This function returns the rowid of the most recent successful INSERT into the database from the database connection in the first argument. As of SQLite version 3.7.7, this functions records the last insert rowid of both ordinary tables and virtual tables. If no successful INSERTs have ever occurred on that database connection, zero is returned.

If an INSERT occurs within a trigger or within a virtual table method, then this function will return the rowid of the inserted row as long as the trigger or virtual table method is running. But once the trigger or virtual table method ends, the value returned by this function reverts to what it was before the trigger or virtual table method began.

An INSERT that fails due to a constraint violation is not a successful INSERT and does not change the value returned by this function. Thus INSERT OR FAIL, INSERT OR IGNORE, INSERT OR ROLLBACK, and INSERT OR ABORT make no changes to the return value of this function when their insertion fails. When INSERT OR REPLACE encounters a constraint violation, it does not fail. The INSERT continues to completion after deleting rows that caused the constraint problem so INSERT OR REPLACE will always change the return value of this interface.

For the purposes of this function, an INSERT is considered to be successful even if it is subsequently rolled back.

This function is accessible to SQL statements via the **last_insert_rowid** SQL function.

If a separate thread performs a new INSERT on the same database connection while the **LastInsertRowid** function is running and thus changes the last insert rowid, then the value returned by **LastInsertRowid** is unpredictable and might not equal either the old or the new last insert rowid. 

# <a name="Limit"></a>Limit

This function allows the size of various constructs to be limited on a connection by connection basis.

```
FUNCTION Limit (BYVAL id AS LONG, BYVAL newVal AS LONG) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *id* | One of the limit categories that define a class of constructs to be size limited. |
| *newVal* | The the new limit for that construct. |

#### Return value

The prior value of the limit.

#### Remarks

If the new limit is a negative number, the limit is unchanged. For each limit category SQLITE_LIMIT_NAME there is a hard upper bound set at compile-time by a C preprocessor macro called SQLITE_MAX_NAME. (The "_LIMIT_" in the name is changed to "_MAX_".) Attempts to increase a limit above its hard upper bound are silently truncated to the hard upper bound.

Regardless of whether or not the limit was changed, the **Limit** interface returns the prior value of the limit. Hence, to find the current value of a limit without changing it, simply invoke this interface with the third parameter set to -1.

Run-time limits are intended for use in applications that manage both their own internal database and also databases that are controlled by untrusted external sources. An example application might be a web browser that has its own databases for storing history and separate databases controlled by JavaScript applications downloaded off the Internet. The internal databases can be given the large, default limits. Databases managed by external sources can be given much smaller limits designed to prevent a denial of service attack.

New run-time limit categories may be added in future releases.

# <a name="OpenBlob"></a>OpenBlob

This function opens a handle to the BLOB located in row *iRow*, column *szColumnName*, table *szTableName* in database *szDbName*; in other words, the same BLOB that would be selected by:

```
SELECT szColumn FROM szDb.szTable WHERE rowid = qRow;
```

```
FUNCTION OpenBlob (BYREF szDbName AS ZSTRING, BYREF szTableName AS ZSTRING, _
   BYREF szColumnName AS ZSTRING, BYVAL iRow AS sqlite3_int64, _
   BYVAL flags AS LONG = 0) AS sqlite3_blob PTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *szDbName* | The database name. Note that the database name is not the filename that contains the database but rather the symbolic name of the database that appears after the AS keyword when the database is connected using ATTACH. For the main database file, the database name is "main". For TEMP tables, the database name is "temp". |
| *szTableName* | The table name. |
| *szColumnName* | The column name. |
| *iRow* | The row number. If the row that a BLOB handle points to is modified by an UPDATE, DELETE, or by ON CONFLICT side-effects then the BLOB handle is marked as "expired". This is true if any column of the row is changed, even a column other than the one the BLOB handle is open on. Calls to **BlobRead** and **BlobWrite** for an expired BLOB handle fail with a return code of SQLITE_ABORT. Changes written into a BLOB prior to the BLOB expiring are not rolled back by the expiration of the BLOB. Such changes will eventually commit if the transaction continues to completion. |
| *flags* | If the flags parameter is non-zero, then the BLOB is opened for read and write access. If it is zero, the BLOB is opened for read access. It is not possible to open a column that is part of an index or primary key for writing. If foreign key constraints are enabled, it is not possible to open a column that is part of a child key for writing. |

#### Return value

On success, the new BLOB handle is returned. On failure, the returned pointer will be null. To get the error code, call **GetLastResult**.

#### Remarks

Use the **BlobBytes** function to determine the size of the opened blob. The size of a blob may not be changed by this function. Use the UPDATE SQL command to change the size of a blob.

The **BindZeroblob** and **ResultZeroblob** functions and the built-in zeroblob SQL function can be used, if desired, to create an empty, zero-filled blob in which to read or write using this function.

To avoid a resource leak, every open BLOB handle should eventually be released by a call to **CloseBlob*. 

# <a name="OpenDb"></a>OpenDb

Opens an SQLite database file as specified by the filename argument.

```
FUNCTION OpenDb (BYREF wszFileName AS CONST WSTRING) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszFileName* | The database filename. |

#### Return value

Returns SQLITE_OK if successful or an error code otherwise. The **ErrMsg** methods can be used to obtain an English language description of the error following a failure of the **OpenDb** method.

#### Remarks

If the filename is ":memory:", then a private, temporary in-memory database is created for the connection. This in-memory database will vanish when the database connection is closed. Future versions of SQLite might make use of additional special filenames that begin with the ":" character. It is recommended that when a database filename actually does begin with a ":" character you should prefix the filename with a pathname such as "./" to avoid ambiguity.

If the filename is an empty string, then a private, temporary on-disk database will be created. This private database will be automatically deleted as soon as the database connection is closed.

# <a name="Prepare"></a>Prepare

Creates a new prepared statement object.

```
FUNCTION Prepare (BYREF wszSql AS WSTRING) AS sqlite3_stmt PTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszSql* | The SQL statement to be compiled. |

#### Return value

An sqlite3_stmt pointer.

# <a name="ProgressHandler"></a>ProgressHandler

The **ProgressHandler** method causes a callback function to be invoked periodically during long running calls to **Step_** and **GetRow** for a database connection. An example use for this interface is to keep a GUI updated during a large query.

```
SUB ProgressHandler (BYVAL nOps AS LONG, BYVAL pCallback AS ANY PTR, BYVAL pArg AS ANY PTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *nOps* | Number of virtual machine instructions that are evaluated between successive invocations of the callback. |
| *pCallback* | Pointer to the callback function. |
| *pArg* | Data that is passed through as the only parameter to the callback function. |

Callback function prototype:
```
FUNCTION sqlite3_progress_handler_callback CDECL (BYVAL pArg AS ANY PTR) AS LONG
```

#### Remarks

Only a single progress handler may be defined at one time per database connection; setting a new progress handler cancels the old one. Setting parameter pCallBack to NULL disables the progress handler. The progress handler is also disabled by setting pCallBack to a value less than 1.

If the progress callback returns non-zero, the operation is interrupted. This feature can be used to implement a "Cancel" button on a GUI progress dialog box.

The progress handler callback must not do anything that will modify the database connection that invoked the progress handler. Note that **Prepare** and **Step_** both modify their database connections for the meaning of "modify" in this paragraph.

# <a name="ReleaseMemory2"></a>ReleaseMemory

Attempts to free as much heap memory as possible from the specified database connection. Unlike **ReleaseMemory**, this function is effect even when then SQLITE_ENABLE_MEMORY_MANAGEMENT compile-time option is omitted.

```
FUNCTION ReleaseMemory () AS LONG
```

#### Return value

The number of bytes freed.

# <a name="Status2"></a>Status

Retrieves runtime status information about a single database connection.

```
FUNCTION Status (BYVAL op AS LONG, BYREF pCurrent AS LONG, BYREF pHighwater AS LONG, _
   BYVAL resetFlag AS BOOLEAN = FALSE) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *op* | An integer constant, taken from the set of SQLITE_DBSTATUS options, that determines the parameter to interrogate. |
| *pCurrent* | The current value of the requested parameter. |
| *pHighwater* | The highest instantaneous value. |
| *resetFlag* | TRUE or FALSE. If true, then the highest instantaneous value is reset back down to the current value. |

#### SQLITE_DBSTATUS options
```
SQLITE_DBSTATUS_LOOKASIDE_USED      = 0
SQLITE_DBSTATUS_CACHE_USED          = 1
SQLITE_DBSTATUS_SCHEMA_USED         = 2
SQLITE_DBSTATUS_STMT_USED           = 3
SQLITE_DBSTATUS_LOOKASIDE_HIT       = 4
SQLITE_DBSTATUS_LOOKASIDE_MISS_SIZE = 5
SQLITE_DBSTATUS_LOOKASIDE_MISS_FULL = 6
SQLITE_DBSTATUS_CACHE_HIT           = 7
SQLITE_DBSTATUS_CACHE_MISS          = 8
SQLITE_DBSTATUS_CACHE_WRITE         = 9
```

#### Return value

SQLITE_OK (0) on success and a non-zero error code on failure.

# <a name="TotalChanges"></a>TotalChanges

This function returns the number of row changes caused by INSERT, UPDATE or DELETE statements since the database connection was opened.

```
FUNCTION TotalChanges () AS LONG
```

#### Return value

The number of changes.

#### Remarks

The count returned by **TotalChanges** includes all changes from all trigger contexts and changes made by foreign key actions. However, the count does not include changes used to implement REPLACE constraints, do rollbacks or ABORT processing, or DROP TABLE processing. The count does not include rows of views that fire an INSTEAD OF trigger, though if the INSTEAD OF trigger makes changes of its own, those changes are counted. The **TotalChanges** function counts the changes as soon as the statement that makes them is completed (when the statement handle is passed to **Reset** or **Finalize**).

See also the **Changes** interface, the **count_changes pragma**, and the **Changes** SQL function.

If a separate thread makes changes on the same database connection while **Changes** is running then the value returned is unpredictable and not meaningful.

# <a name="UnlockNotify"></a>UnlockNotify

When running in shared-cache mode, a database operation may fail with an SQLITE_LOCKED error if the required locks on the shared-cache or individual tables within the shared-cache cannot be obtained. This API may be used to register a callback that SQLite will invoke when the connection currently holding the required lock relinquishes it. This API is only available if the library was compiled with the SQLITE_ENABLE_UNLOCK_NOTIFY C-preprocessor symbol defined.

```
FUNCTION UnlockNotify (BYVAL pNotifyCallback AS ANY PTR, BYVAL pNotifyArg AS ANY PTR) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *pNotifyCallback* | Pointer to the callback function. |
| *pNotifyArg* | Data that is passed through the second parameter to the callback function. |

Return value

Unless deadlock is detected (see below), UnlockNotify always returns SQLITE_OK.

#### Remarks

Shared-cache locks are released when a database connection concludes its current transaction, either by committing it or rolling it back.

When a connection (known as the blocked connection) fails to obtain a shared-cache lock and SQLITE_LOCKED is returned to the caller, the identity of the database connection (the blocking connection) that has locked the required resource is stored internally. After an application receives an SQLITE_LOCKED error, it may call the **UnlockNotify** method with the blocked connection handle as the first argument to register for a callback that will be invoked when the blocking connections current transaction is concluded. The callback is invoked from within the **Step_** or **Close** call that concludes the blocking connections transaction.

If **UnlockNotify** is called in a multi-threaded application, there is a chance that the blocking connection will have already concluded its transaction by the time **UnlockNotify** is invoked. If this happens, then the specified callback is invoked immediately, from within the call to **UnlockNotify**.

If the blocked connection is attempting to obtain a write-lock on a shared-cache table, and more than one other connection currently holds a read-lock on the same table, then SQLite arbitrarily selects one of the other connections to use as the blocking connection.

There may be at most one unlock-notify callback registered by a blocked connection. If **UnlockNotify** is called when the blocked connection already has a registered unlock-notify callback, then the new callback replaces the old. If **UnlockNotify** is called with a NULL pointer as its second argument, then any existing unlock-notify callback is canceled. The blocked connections unlock-notify callback may also be canceled by closing the blocked connection using **Close**.

The unlock-notify callback is not reentrant. If an application invokes any **sqlite3_xxx** API functions from within an unlock-notify callback, a crash or deadlock may be the result.

#### Callback Invocation Details

When an unlock-notify callback is registered, the application provides a single void* pointer that is passed to the callback when it is invoked. However, the signature of the callback function allows SQLite to pass it an array of void* context pointers. The first argument passed to an unlock-notify callback is a pointer to an array of void* pointers, and the second is the number of entries in the array.

When a blocking connections transaction is concluded, there may be more than one blocked connection that has registered for an unlock-notify callback. If two or more such blocked connections have specified the same callback function, then instead of invoking the callback function multiple times, it is invoked once with the set of void* context pointers specified by the blocked connections bundled together into an array. This gives the application an opportunity to prioritize any actions related to the set of unblocked database connections.

#### Deadlock Detection

Assuming that after registering for an unlock-notify callback a database waits for the callback to be issued before taking any further action (a reasonable assumption), then using this API may cause the application to deadlock. For example, if connection X is waiting for connection Y's transaction to be concluded, and similarly connection Y is waiting on connection X's transaction, then neither connection will proceed and the system may remain deadlocked indefinitely.

To avoid this scenario, the **UnlockNotify** performs deadlock detection. If a given call to **UnlockNotify** would put the system in a deadlocked state, then SQLITE_LOCKED is returned and no unlock-notify callback is registered. The system is said to be in a deadlocked state if connection A has registered for an unlock-notify callback on the conclusion of connection B's transaction, and connection B has itself registered for an unlock-notify callback when connection A's transaction is concluded. Indirect deadlock is also detected, so the system is also considered to be deadlocked if connection B has registered for an unlock-notify callback on the conclusion of connection C's transaction, where connection C is waiting on connection A. Any number of levels of indirection are allowed.

#### The "DROP TABLE" Exception

When a call to **Step_** returns SQLITE_LOCKED, it is almost always appropriate to call **UnlockNotify**. There is however, one exception. When executing a "DROP TABLE" or "DROP INDEX" statement, SQLite checks if there are any currently executing SELECT statements that belong to the same connection. If there are, SQLITE_LOCKED is returned. In this case there is no "blocking connection", so invoking **UnlockNotify** results in the unlock-notify callback being invoked immediately. If the application then re-attempts the "DROP TABLE" or "DROP INDEX" query, an infinite loop might be the result.

One way around this problem is to check the extended error code returned by an **Step_** call. If there is a blocking connection, then the extended error code is set to SQLITE_LOCKED_SHAREDCACHE. Otherwise, in the special "DROP TABLE/INDEX" case, the extended error code is just SQLITE_LOCKED.

# <a name="BindBlob"></a>BindBlob

Binds a blob with the statement.

```
FUNCTION BindBlob (BYVAL idx AS LONG, BYVAL pValue AS ANY PTR, BYVAL numBytes AS LONG, _
   BYVAL pDestructor AS ANY PTR = NULL) AS LONG
FUNCTION BindBlob64 (BYVAL idx AS LONG, BYVAL pValue AS ANY PTR, BYVAL numBytes AS sqlite3_uint64, _
   BYVAL pDestructor AS ANY PTR = NULL) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *idx* | Index of the SQL parameter to be set. The leftmost SQL parameter has an index of 1. When the same named SQL parameter is used more than once, second and subsequent occurrences have the same index as the first occurrence. The index for named parameters can be looked up using the **BindParameterIndex** method if desired. The index for "?NNN" parameters is the value of NNN. The NNN value must be between 1 and the **Limit** parameter SQLITE_LIMIT_VARIABLE_NUMBER (default value: 999). |
| *pValue* | The value to bind to the parameter. |
| *numBytes* | The number of bytes in the parameter. To be clear: the value is the number of bytes in the value, not the number of characters. |
| *pDestructor* | A destructor used to dispose of the BLOB after SQLite has finished with it. The destructor is called to dispose of the BLOB even if the call to **BindBlob** fails. If this argument is the special value **SQLITE_STATIC**, then SQLite assumes that the information is in static, unmanaged space and does not need to be freed. If this argument has the value **SQLITE_TRANSIENT**, then SQLite makes its own private copy of the data immediately, before the **BindBlob** function returns. |

#### Return value

SQLITE_OK on success or an error code if anything goes wrong. SQLITE_RANGE is returned if the parameter index is out of range. SQLITE_NOMEM is returned if malloc fails.

#### Remarks

If **BindBlob** is called with a NULL pointer for the prepared statement or with a prepared statement for which **Step_** has been called more recently than **Reset**, then the call will return SQLITE_MISUSE. If **BindBlob** is passed a prepared statement that has been finalized, the result is undefined and probably harmful.

Bindings are not cleared by the **Reset** function. Unbound parameters are interpreted as NULL.

# <a name="BindDouble"></a>BindDouble

Binds a double value with the statement.

```
FUNCTION BindDouble (BYVAL idx AS LONG, BYVAL dblValue AS DOUBLE) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *îdx* | Index of the SQL parameter to be set. The leftmost SQL parameter has an index of 1. When the same named SQL parameter is used more than once, second and subsequent occurrences have the same index as the first occurrence. The index for named parameters can be looked up using the **BindParameterIndex** method if desired. The index for "?NNN" parameters is the value of NNN. The NNN value must be between 1 and the **Limit** parameter SQLITE_LIMIT_VARIABLE_NUMBER (default value: 999). |
| *dblValue* | The value to bind to the parameter. |

#### Return value

SQLITE_OK on success or an error code if anything goes wrong. SQLITE_RANGE is returned if the parameter index is out of range. SQLITE_NOMEM is returned if **malloc** fails.

#### Remarks

Bindings are not cleared by the **Reset** function. Unbound parameters are interpreted as NULL.

# <a name="BindLong"></a>BindLong

Binds a long integer value with the statement.

```
FUNCTION BindLong (BYVAL idx AS LONG, BYVAL longValue AS LONG) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *îdx* | Index of the SQL parameter to be set. The leftmost SQL parameter has an index of 1. When the same named SQL parameter is used more than once, second and subsequent occurrences have the same index as the first occurrence. The index for named parameters can be looked up using the **BindParameterIndex** method if desired. The index for "?NNN" parameters is the value of NNN. The NNN value must be between 1 and the **Limit** parameter SQLITE_LIMIT_VARIABLE_NUMBER (default value: 999). |
| *longValue* | The value to bind to the parameter. |

#### Return value

SQLITE_OK on success or an error code if anything goes wrong. SQLITE_RANGE is returned if the parameter index is out of range. SQLITE_NOMEM is returned if **malloc** fails.

#### Remarks

Bindings are not cleared by the **Reset** function. Unbound parameters are interpreted as NULL.

# <a name="BindLongInt"></a>BindLongInt

Binds a longint value with the statement.

```
FUNCTION BindLongInt (BYVAL idx AS LONG, BYVAL longintValue AS sqlite3_int64) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *îdx* | Index of the SQL parameter to be set. The leftmost SQL parameter has an index of 1. When the same named SQL parameter is used more than once, second and subsequent occurrences have the same index as the first occurrence. The index for named parameters can be looked up using the **BindParameterIndex** method if desired. The index for "?NNN" parameters is the value of NNN. The NNN value must be between 1 and the **Limit** parameter SQLITE_LIMIT_VARIABLE_NUMBER (default value: 999). |
| *longintValue* | The value to bind to the parameter. |

#### Return value

SQLITE_OK on success or an error code if anything goes wrong. SQLITE_RANGE is returned if the parameter index is out of range. SQLITE_NOMEM is returned if **malloc** fails.

#### Remarks

Bindings are not cleared by the **Reset** function. Unbound parameters are interpreted as NULL.

# <a name="BindNull"></a>BindNull

Binds a null value with the statement.

```
FUNCTION BindNull (BYVAL idx AS LONG) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *îdx* | Index of the SQL parameter to be set. The leftmost SQL parameter has an index of 1. When the same named SQL parameter is used more than once, second and subsequent occurrences have the same index as the first occurrence. The index for named parameters can be looked up using the **BindParameterIndex** method if desired. The index for "?NNN" parameters is the value of NNN. The NNN value must be between 1 and the **Limit** parameter SQLITE_LIMIT_VARIABLE_NUMBER (default value: 999). |

#### Return value

SQLITE_OK on success or an error code if anything goes wrong. SQLITE_RANGE is returned if the parameter index is out of range. SQLITE_NOMEM is returned if **malloc** fails.

#### Remarks

Bindings are not cleared by the **Reset** function. Unbound parameters are interpreted as NULL.

# <a name="BindParameterCount"></a>BindParameterCount

This method can be used to find the number of SQL parameters in a prepared statement. SQL parameters are tokens of the form "?", "?NNN", ":AAA", "$AAA", or "@AAA" that serve as placeholders for values that are bound to the parameters at a later time.

```
FUNCTION BindParameterCount () AS LONG
```

#### Return value

The number of SQL parameters in a prepared statement.

#### Remarks

This function actually returns the index of the largest (rightmost) parameter. For all forms except ?NNN, this will correspond to the number of unique parameters. If parameters of the ?NNN form are used, there may be gaps in the list.

# <a name="BindParameterIndex"></a>BindParameterIndex

Returns the index of an SQL parameter given its name. The index value returned is suitable for use as the second parameter to the **Bind_\*** methods. A zero is returned if no matching parameter is found.

```
FUNCTION BindParameterIndex (BYREF szName AS ZSTRING) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *szName* | The parameter name. Must be given in UTF-8 even if the original statement was prepared from UTF-16 text using **Prepare**. |

#### Return value

The index of the parameter.

# <a name="BindParameterName"></a>BindParameterName

Returns the name of the N-th SQL parameter in the prepared statement P. SQL parameters of the form "?NNN" or ":AAA" or "@AAA" or "$AAA" have a name which is the string "?NNN" or ":AAA" or "@AAA" or "$AAA" respectively. In other words, the initial ":" or "$" or "@" or "?" is included as part of the name. Parameters of the form "?" without a following integer have no name and are referred to as "nameless" or "anonymous parameters".

```
FUNCTION BindParameterName (BYVAL idx AS LONG) AS STRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *idx* | The parameter index. The first host parameter has an index of 1, not 0. |

#### Return value

If the value *idx* is out of range or if the idx-th parameter is nameless, then NULL is returned. The returned string is always in UTF-8 encoding even if the named parameter was originally specified as UTF-16 in Prepare.

# <a name="BindText"></a>BindText

Binds a text value with the statement.

```
FUNCTION BindText (BYVAL idx AS LONG, BYREF wszValue AS WSTRING, _
   BYVAL pDestructor AS ANY PTR = NULL) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *îdx* | Index of the SQL parameter to be set. The leftmost SQL parameter has an index of 1. When the same named SQL parameter is used more than once, second and subsequent occurrences have the same index as the first occurrence. The index for named parameters can be looked up using the **BindParameterIndex** method if desired. The index for "?NNN" parameters is the value of NNN. The NNN value must be between 1 and the **Limit** parameter SQLITE_LIMIT_VARIABLE_NUMBER (default value: 999). |
| *wszValue* | The value to bind to the parameter. |
| *pDestructor* | A destructor used to dispose of the BLOB after SQLite has finished with it. The destructor is called to dispose of the BLOB even if the call to **BindBlob** fails. If this argument is the special value **SQLITE_STATIC**, then SQLite assumes that the information is in static, unmanaged space and does not need to be freed. If this argument has the value **SQLITE_TRANSIENT**, then SQLite makes its own private copy of the data immediately, before the **BindBlob** function returns. |

#### Return value

SQLITE_OK on success or an error code if anything goes wrong. SQLITE_RANGE is returned if the parameter index is out of range. SQLITE_NOMEM is returned if malloc fails.

#### Remarks

If **BindText** is called with a NULL pointer for the prepared statement or with a prepared statement for which **Step_** has been called more recently than **Reset**, then the call will return SQLITE_MISUSE. If **BindText** is passed a prepared statement that has been finalized, the result is undefined and probably harmful.

Bindings are not cleared by the Reset function. Unbound parameters are interpreted as NULL.

# <a name="BindZeroBlob"></a>BindZeroBlob

Binds a BLOB of length *numBytes* that is filled with zeroes.

```
FUNCTION BindZeroBlob (BYVAL idx AS LONG, BYVAL numBytes AS LONG) AS LONG
FUNCTION BindZeroBlob64 (BYVAL idx AS LONG, BYVAL numBytes AS sqlite3_uint64) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *îdx* | Index of the SQL parameter to be set. The leftmost SQL parameter has an index of 1. When the same named SQL parameter is used more than once, second and subsequent occurrences have the same index as the first occurrence. The index for named parameters can be looked up using the **BindParameterIndex** method if desired. The index for "?NNN" parameters is the value of NNN. The NNN value must be between 1 and the **Limit** parameter SQLITE_LIMIT_VARIABLE_NUMBER (default value: 999). |
| *numBytes* | The number of bytes. |

#### Return value

SQLITE_OK on success or an error code if anything goes wrong. SQLITE_RANGE is returned if the parameter index is out of range. SQLITE_NOMEM is returned if **malloc** fails.

#### Remarks

The **BindZeroblob** function binds a BLOB of length *numBytes* that is filled with zeroes. A zeroblob uses a fixed amount of memory (just an integer to hold its size) while it is being processed. Zeroblobs are intended to serve as placeholders for BLOBs whose content is later written using incremental BLOB I/O functions. A negative value for the zeroblob results in a zero-length BLOB.

Bindings are not cleared by the **Reset** function. Unbound parameters are interpreted as NULL.

# <a name="Busy"></a>Busy

Returns true if the prepared statement has been stepped at least once using **Step_** but has not run to completion and/or has not been reset using **Reset**.

```
FUNCTION Busy () AS BOOLEAN
```

#### Return value

TRUE or FALSE.

# <a name="ClearBindings"></a>ClearBindings

Sets all the parameters in the compiled SQL statement to NULL. Contrary to the intuition of many, **Reset** does not reset the bindings on a prepared statement. Use this function to reset all host parameters to NULL. 

```
FUNCTION ClearBindings () AS LONG
```

#### Return value

SQLITE_OK on success or an error code if anything goes wrong.

# <a name="ColumnBlob"></a>ColumnBlob

Returns information about a single column of the current result row of a query.

```
FUNCTION ColumnBlob (BYVAL nCol AS LONG) AS ANY PTR
FUNCTION ColumnBlob (BYREF wszColName AS WSTRING) AS ANY PTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *nCol / wszColName* | The index or the name of the column for which information should be returned. The leftmost column of the result set has the index 0. The number of columns in the result can be determined using **ColumnCount**. |

#### Return value

The column value as a pointer to a blob.

#### Remarks

If the SQL statement does not currently point to a valid row, or if the column index is out of range, the result is undefined. **ColumnLong** may only be called when the most recent call to Step_ has returned SQLITE_ROW and neither **Reset** nor **Finalize** have been called subsequently. If any of these functions are called after **Reset** or **Finalize** or after **Step_** has returned something other than SQLITE_ROW, the results are undefined. If **Step_** or **Reset** or **Finalize** are called from a different thread while **ColumnBlob** is pending, then the result is undefined.

The return value from **ColumnBlob** for a zero-length BLOB is a NULL pointer.

**ColumnBlob** attempts to convert the value where appropriate. The following table details the conversions that are applied:
 
| Internal Type  | Requested Type | Conversion |
| -------------- | -------------- | ---------- |
| NULL | BLOB | Result is NULL pointer |
| INTEGER | BLOB | Use atoi() |
| FLOAT | BLOB | Use atof() |
| TEXT | BLOB | No change |

The table above makes reference to standard C library functions atoi() and atof(). SQLite does not really use these functions. It has its own equivalent internal functions. The atoi() and atof() names are used in the table for brevity and because they are familiar to most C programmers.

The pointer returned is valid until **Step_** or **Reset** or **Finalize** is called. The memory space used to hold BLOBs is freed automatically. Do not pass the pointers returned by **ColumnBlob** into **Free**.

If a memory allocation error occurs during the evaluation of any of these functions, a default value is returned. The default value is a NULL pointer. Subsequent calls to **ErrCode** will return SQLITE_NOMEM. 

# <a name="ColumnBytes"></a>ColumnBytes

Returns the number of bytes of the column value.

```
FUNCTION ColumnBytes (BYVAL nCol AS LONG) AS LONG
FUNCTION ColumnBytes (BYREF wszColName AS WSTRING) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *nCol / wszColName* | The index or the name of the column for which information should be returned. The leftmost column of the result set has the index 0. The number of columns in the result can be determined using **ColumnCount**. |

#### Return value

The number of bytes of the column value.

#### Remarks

If the SQL statement does not currently point to a valid row, or if the column index is out of range, the result is undefined. These functions may only be called when the most recent call to **Step_** has returned SQLITE_ROW and neither **Reset** nor **Finalize** have been called subsequently. If any of these functions are called after **Reset** or **Finalize** or after **Step_** has returned something other than SQLITE_ROW, the results are undefined. If **Step_** or **Reset** or **Finalize** are called from a different thread while **ColumnBytes** is pending, then the result is undefined.

If the result is a BLOB or UTF-8 string then the **ColumnBytes** function returns the number of bytes in that BLOB or string. If the result is a UTF-16 string, then **ColumnBytes** converts the string to UTF-8 and then returns the number of bytes. If the result is a numeric value then **ColumnBytes** uses **snprintf** to convert that value to a UTF-8 string and returns the number of bytes in that string. If the result is NULL, then **ColumnBytes** returns zero.

The values returned by **ColumnBytes** do not include the zero terminators at the end of the string. For clarity: the values returned by **ColumnBytes** are the number of bytes in the string, not the number of characters.

Note that when type conversions occur, pointers returned by prior calls to **ColumnBlob** and/or, **ColumnText** may be invalidated.

The safest and easiest to remember policy is to invoke these functions in one of the following ways:

**ColumnText** followed by **ColumnBytes**<br>
**ColumnBlob** followed by **ColumnBytes**

In other words, you should call **ColumnText** or **ColumnBlob** first to force the result into the desired format, then invoke **ColumnBytes** to find the size of the result.

The pointers returned are valid until a type conversion occurs as described above, or until **Step_** or **Reset** or **Finalize** is called. The memory space used to hold strings and BLOBs is freed automatically. Do not pass the pointers returned **ColumnBlob**, **ColumnText**, etc. into **Free**.

If a memory allocation error occurs during the evaluation of any of these functions, a default value is returned. The default value is 0. Subsequent calls to **ErrCode** will return SQLITE_NOMEM. 

# <a name="ColumnCount"></a>ColumnCount

Returns the number of columns in the result set returned by the prepared statement. This function returns 0 if pStmt is an SQL statement that does not return data (for example an UPDATE).

```
FUNCTION ColumnCount () AS LONG
```

# <a name="ColumnDatabaseName"></a>ColumnDatabaseName

Returns the database name that is the origin of a particular result column in SELECT statement.

```
FUNCTION ColumnDatabaseName (BYVAL nCol AS LONG) AS CWSTR
FUNCTION ColumnDatabaseName (BYREF wszColName AS WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *nCol / wszColName* | The column number or name of the result set. The leftmost column of the result set has the index 0. The number of columns in the result can be determined using **ColumnCount**. |

#### Return value

The original un-aliased name of the database. If the column returned by the statement is an expression or subquery and is not a column value, then this functions returns an empty string.

#### Remarks

This function is only available if the library was compiled with the SQLITE_ENABLE_COLUMN_METADATA C-preprocessor symbol.

If two or more threads call this function against the same prepared statement and column at the same time then the results are undefined.

If two or more threads call one or more column metadata interfaces for the same prepared statement and result column at the same time then the results are undefined. 

# <a name="ColumnDeclaredType"></a>ColumnDeclaredType

Returns the declared data type of a query result.

```
FUNCTION ColumnDeclaredType (BYVAL nCol AS LONG) AS CWSTR
FUNCTION ColumnDeclaredType (BYREF wszColName AS WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *nCol / wszColName* | The column number or name of the result set. The leftmost column of the result set has the index 0. The number of columns in the result can be determined using **ColumnCount**. |

#### Return value

The original un-aliased name of the database. If the column returned by the statement is an expression or subquery and is not a column value, then this functions returns an empty string.

#### Remarks

This function is only available if the library was compiled with the SQLITE_ENABLE_COLUMN_METADATA C-preprocessor symbol.

If two or more threads call this function against the same prepared statement and column at the same time then the results are undefined.

If two or more threads call one or more column metadata interfaces for the same prepared statement and result column at the same time then the results are undefined. 

# <a name="ColumnDouble"></a>ColumnDouble

Returns the column value as a double.

```
FUNCTION ColumnDouble (BYVAL nCol AS LONG) AS DOUBLE
FUNCTION ColumnDouble (BYREF wszColName AS WSTRING) AS DOUBLE
```

| Parameter  | Description |
| ---------- | ----------- |
| *nCol / wszColName* | The index or name of the column for which information should be returned. The leftmost column of the result set has the index 0. The number of columns in the result can be determined using **ColumnCount**. |

#### Return value

The column value as a double.

#### Remarks

If the SQL statement does not currently point to a valid row, or if the column index is out of range, the result is undefined. **ColumnDouble** may only be called when the most recent call to **Step_** has returned SQLITE_ROW and neither **Reset** nor **Finalize** have been called subsequently. If any of these functions are called after **Reset** or **Finalize** or after **Step_** has returned something other than SQLITE_ROW, the results are undefined. If **Step_** or **Reset** or **Finalize** are called from a different thread while ColumnBytes is pending, then the result is undefined.

**ColumnDouble** attempts to convert the value where appropriate. The following table details the conversions that are applied:

| Internal Type  | Requested Type | Conversion |
| -------------- | -------------- | ---------- |
| NULL | FLOAT | Result is 0.0 |
| INTEGER | FLOAT | Convert from integer to float |
| TEXT | FLOAT | Use atof() |
| BLOB | FLOAT | Convert to TEXT then use atof() |

The table above makes reference to standard C library functions **atoi**() and **atof**(). SQLite does not really use these functions. It has its own equivalent internal functions. The **atoi**() and **atof**() names are used in the table for brevity and because they are familiar to most C programmers.

If a memory allocation error occurs during the evaluation of any of these functions, a default value is returned. The default value is the floating point number 0.0. Subsequent calls to **ErrCode** will return SQLITE_NOMEM. 

# <a name="ColumnLong"></a>ColumnLong

Returns the column value as a long.

```
FUNCTION ColumnLong (BYVAL nCol AS LONG) AS LONG
FUNCTION ColumnLong (BYREF wszColName AS WSTRING) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *nCol / wszColName* | The index or name of the column for which information should be returned. The leftmost column of the result set has the index 0. The number of columns in the result can be determined using **ColumnCount**. |

#### Return value

The column value as a long.

#### Remarks

If the SQL statement does not currently point to a valid row, or if the column index is out of range, the result is undefined. **ColumnLong** may only be called when the most recent call to **Step_** has returned SQLITE_ROW and neither **Reset** nor **Finalize** have been called subsequently. If any of these functions are called after **Reset** or **Finalize** or after **Step_** has returned something other than SQLITE_ROW, the results are undefined. If **Step_** or **Reset** or **Finalize** are called from a different thread while ColumnBytes is pending, then the result is undefined.

**ColumnLong** attempts to convert the value where appropriate. The following table details the conversions that are applied:

| Internal Type  | Requested Type | Conversion |
| -------------- | -------------- | ---------- |
| NULL | INTEGER | Result is 0 |
| FLOAT | INTEGER | Convert from float to integer |
| TEXT | INTEGER | Use atoi() |
| BLOB | INTEGER | Convert to TEXT then use atoi() |

The table above makes reference to standard C library function **atoi**(). SQLite does not really use these functions. It has its own equivalent internal functions. The **atoi**() name are used in the table for brevity and because they are familiar to most C programmers.

If a memory allocation error occurs during the evaluation of any of these functions, a default value is returned. The default value is the floating point number 0.0. Subsequent calls to **ErrCode** will return SQLITE_NOMEM. 

# <a name="ColumnLongInt"></a>ColumnLongInt

Returns the column value as a longint.

```
FUNCTION ColumnLongInt (BYVAL nCol AS LONG) AS LONGINT
FUNCTION ColumnLongInt (BYREF wszColName AS WSTRING) AS LONGINT
```

| Parameter  | Description |
| ---------- | ----------- |
| *nCol / wszColName* | The index or name of the column for which information should be returned. The leftmost column of the result set has the index 0. The number of columns in the result can be determined using **ColumnCount**. |

#### Return value

The column value as a longint.

#### Remarks

If the SQL statement does not currently point to a valid row, or if the column index is out of range, the result is undefined. **ColumnLongInt** may only be called when the most recent call to **Step_** has returned SQLITE_ROW and neither **Reset** nor **Finalize** have been called subsequently. If any of these functions are called after **Reset** or **Finalize** or after **Step_** has returned something other than SQLITE_ROW, the results are undefined. If **Step_** or **Reset** or **Finalize** are called from a different thread while ColumnBytes is pending, then the result is undefined.

**ColumnLongInt** attempts to convert the value where appropriate. The following table details the conversions that are applied:

| Internal Type  | Requested Type | Conversion |
| -------------- | -------------- | ---------- |
| NULL | INTEGER | Result is 0 |
| FLOAT | INTEGER | Convert from float to integer |
| TEXT | INTEGER | Use atoi() |
| BLOB | INTEGER | Convert to TEXT then use atoi() |

The table above makes reference to standard C library function **atoi**(). SQLite does not really use these functions. It has its own equivalent internal functions. The **atoi**() name are used in the table for brevity and because they are familiar to most C programmers.

If a memory allocation error occurs during the evaluation of any of these functions, a default value is returned. The default value is the floating point number 0.0. Subsequent calls to **ErrCode** will return SQLITE_NOMEM. 

# <a name="ColumnName"></a>ColumnName

Returns the name assigned to a particular column in the result set of a SELECT statement.

```
FUNCTION ColumnName (BYVAL nCol AS LONG) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *nCol* | The column number of the result set. The leftmost column of the result set has the index 0. The number of columns in the result can be determined using **ColumnCount**. |

#### Return value

The name assigned to the specified column. The name of a result column is the value of the "AS" clause for that column, if there is an AS clause. If there is no AS clause then the name of the column is unspecified and may change from one release of SQLite to the next.

# <a name="ColumnOriginName"></a>ColumnOriginName

Returns the column name that is the origin of a particular result column in SELECT statement.

```
FUNCTION ColumnOriginName (BYVAL nCol AS LONG) AS CWSTR
FUNCTION ColumnOriginName (BYREF wszColName AS WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *nCol / wszColName* | The column number or name of the result set. The leftmost column of the result set has the index 0. The number of columns in the result can be determined using **ColumnCount**. |

#### Return value

The original un-aliased name of the column. If the column returned by the statement is an expression or subquery and is not a column value, then this functions returns an empty string.

#### Remarks

This function is only available if the library was compiled with the SQLITE_ENABLE_COLUMN_METADATA C-preprocessor symbol.

If two or more threads call this function against the same prepared statement and column at the same time then the results are undefined.

If two or more threads call one or more column metadata interfaces for the same prepared statement and result column at the same time then the results are undefined. 

# <a name="ColumnTableName"></a>ColumnTableName

Returns the table name that is the origin of a particular result column in SELECT statement.

```
FUNCTION ColumnTableName (BYVAL nCol AS LONG) AS CWSTR
FUNCTION ColumnTableName (BYREF wszColName AS WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *nCol / wszColName* | The column number or name of the result set. The leftmost column of the result set has the index 0. The number of columns in the result can be determined using **ColumnCount**. |

#### Return value

The original un-aliased name of the table. If the column returned by the statement is an expression or subquery and is not a column value, then this functions returns an empty string.

#### Remarks

This function is only available if the library was compiled with the SQLITE_ENABLE_COLUMN_METADATA C-preprocessor symbol.

If two or more threads call this function against the same prepared statement and column at the same time then the results are undefined.

If two or more threads call one or more column metadata interfaces for the same prepared statement and result column at the same time then the results are undefined. 

# <a name="ColumnText"></a>ColumnText

Returns the column value as a UTF-16 string.

```
FUNCTION ColumnText (BYVAL nCol AS LONG) AS CWSTR
FUNCTION ColumnText (BYREF wszColName AS WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *nCol / wszColName* | The index or name of the column for which information should be returned. The leftmost column of the result set has the index 0. The number of columns in the result can be determined using **ColumnCcount**. |

#### Return value

The column value.

#### Remarks

If the SQL statement does not currently point to a valid row, or if the column index is out of range, the result is undefined. **ColumnText** may only be called when the most recent call to **Step_** has returned SQLITE_ROW and neither **Reset** nor **Finalize** have been called subsequently. If any of these functions are called after **Reset** or **Finalize** or after **Step_** has returned something other than SQLITE_ROW, the results are undefined. If **Step_** or **Reset** or **Finalize** are called from a different thread while ColumnBytes is pending, then the result is undefined.

Strings returned by **ColumnText**, even empty strings, are always zero-terminated.

**ColumnText** attempts to convert the value where appropriate. The following table details the conversions that are applied:

| Internal Type  | Requested Type | Conversion |
| -------------- | -------------- | ---------- |
| NULL | TEXT | Result is NULL pointer |
| INTEGER | TEXT | ASCII rendering of the integer |
| FLOAT | TEXT | ASCII rendering of the float |
| BLOB | TEXT | Add a zero terminator if needed |

Note that when type conversions occur, pointers returned by prior calls to **ColumnBlob** and/or **ColumnText** may be invalidated. Type conversions and pointer invalidations might occur in the following cases:

The initial content is a BLOB and **ColumnText** is called. A zero-terminator might need to be added to the string.

The safest and easiest to remember policy is to invoke these functions in one of the following ways:

**ColumnText** followed by **ColumnBytes**<br>
**ColumnBlob** followed by **ColumnBytes**

In other words, you should call **ColumnText** or **ColumnBlob** first to force the result into the desired format, then invoke **ColumnBytes** to find the size of the result.

If a memory allocation error occurs during the evaluation of any of these functions, a default value is returned. The default value is either the integer 0, the floating point number 0.0, or a NULL pointer. Subsequent calls to **ErrCode** will return SQLITE_NOMEM. 

# <a name="ColumnType"></a>ColumnType

Returns the column type.

```
FUNCTION ColumnType (BYVAL nCol AS LONG) AS LONG
FUNCTION ColumnType (BYREF wszColName AS WSTRING) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *nCol / wszColName* | The index or name of the column for which information should be returned. The leftmost column of the result set has the index 0. The number of columns in the result can be determined using **ColumnCcount**. |

#### Return value

The column type.

#### Remarks

If the SQL statement does not currently point to a valid row, or if the column index is out of range, the result is undefined. **ColumnType** may only be called when the most recent call to **Step_** has returned SQLITE_ROW and neither **Reset** nor **Finalize** have been called subsequently. If any of these functions are called after **Reset** or **Finalize** or after **Step_** has returned something other than SQLITE_ROW, the results are undefined. If **Step_** or **Reset** or **Finalize** are called from a different thread while ColumnBytes is pending, then the result is undefined.

The **ColumnType** function returns the datatype code for the initial data type of the result column. The returned value is one of SQLITE_INTEGER, SQLITE_FLOAT, SQLITE_TEXT, SQLITE_BLOB, or SQLITE_NULL. The value returned by **ColumnType** is only meaningful if no type conversions have occurred as described below. After a type conversion, the value returned by **ColumnType** is undefined. Future versions of SQLite may change the behavior of **ColumnType** following a type conversion.

If the result is a BLOB or UTF-8 string then the **ColumnBytes** function returns the number of bytes in that BLOB or string. If the result is a UTF-16 string, then **ColumnBytes** converts the string to UTF-8 and then returns the number of bytes. If the result is a numeric value then **ColumnBytes** uses **snprintf** to convert that value to a UTF-8 string and returns the number of bytes in that string. If the result is NULL, then **ColumnBytes** returns zero.

If the result is a BLOB or UTF-16 string then the **ColumnBytes** function returns the number of bytes in that BLOB or string. If the result is a numeric value then **ColumnBytes** uses **snprintf** to convert that value to a UTF-16 string and returns the number of bytes in that string. If the result is NULL, then **ColumnBytes** returns zero.

The values returned by **ColumnBytes** do not include the zero terminators at the end of the string. For clarity: the values returned by **ColumnBytes** are the number of bytes in the string, not the number of characters.

Strings returned by **ColumnText**, even empty strings, are always zero-terminated. The return value from **ColumnBlob** for a zero-length BLOB is a NULL pointer.

These functions attempt to convert the value where appropriate. For example, if the internal representation is FLOAT and a text result is requested, snprintf is used internally to perform the conversion automatically. The following table details the conversions that are applied:

| Internal Type  | Requested Type | Conversion |
| -------------- | -------------- | ---------- |
| NULL | INTEGER | Result is 0 |
| NULL | FLOAT | Result is 0.0 |
| NULL | TEXT | Result is NULL pointer |
| NULL | BLOB | Result is NULL pointer |
| INTEGER | FLOAT | Convert from integer to float |
| INTEGER | TEXT | ASCII rendering of the integer |
| INTEGER | BLOB | Same as INTEGER->TEXT |
| FLOAT | INTEGER | Convert from float to integer |
| FLOAT | TEXT | ASCII rendering of the float |
| FLOAT | BLOB | Same as FLOAT->TEXT |
| TEXT | INTEGER | Use atoi() |
| TEXT | FLOAT | Use atof() |
| TEXT | BLOB | No change |
| BLOB | INTEGER | Convert to TEXT then use atoi() |
| BLOB | FLOAT | Convert to TEXT then use atof() |
| BLOB | TEXT | Add a zero terminator if needed |

The table above makes reference to standard C library functions **atoi**() and **atof**(). SQLite does not really use these functions. It has its own equivalent internal functions. The **atoi**() and **atof**() names are used in the table for brevity and because they are familiar to most C programmers.

Note that when type conversions occur, pointers returned by prior calls to **ColumnBlob** and/or **ColumnText** may be invalidated. Type conversions and pointer invalidations might occur in the following cases:

The initial content is a BLOB and **ColumnText** is called. A zero-terminator might need to be added to the string.

The initial content is UTF-8 text and **ColumnBytes** or **ColumnText** is called. The content must be converted to UTF-16.

The safest and easiest to remember policy is to invoke these functions in one of the following ways:

**ColumnText** followed by **ColumnBytes**<br>
**ColumnBlob** followed by **ColumnBytes**

In other words, you should call **ColumnText** or **ColumnBlob** first to force the result into the desired format, then invoke **ColumnBytes** to find the size of the result.

If a memory allocation error occurs during the evaluation of any of these functions, a default value is returned. The default value is either the integer 0, the floating point number 0.0, or a NULL pointer. Subsequent calls to **ErrCode** will return SQLITE_NOMEM. 

# <a name="DataCount"></a>DataCount

Returns the number of columns in the result set returned by the prepared statement. This function returns 0 if pStmt is an SQL statement that does not return data (for example an UPDATE).

```
FUNCTION DataCount () AS LONG
```

#### Return value

The number of columns in the result set returned by the prepared statement.

# <a name="DbHandle"></a>DbHandle

Returns the database connection handle to which a prepared statement belongs. The database connection returned by **DbHandle** is the same database connection that was the first argument to the **Prepare** call that was used to create the statement in the first place. 

```
FUNCTION DbHandle () AS sqlite3 PTR
```

#### Return value

The database connection handle to which a prepared statement belongs.

# <a name="Finalize"></a>Finalize

Deletes a prepared statement.

```
FUNCTION Finalize () AS LONG
```

#### Return value

If the most recent evaluation of the statement encountered no errors or if the statement is never been evaluated, then Finalize returns SQLITE_OK. If the most recent evaluation of a statement failed, then **Finalize** returns the appropriate error code or extended error code.

#### Remarks

The **Finalize** method can be called at any point during the life cycle of prepared statement: before statement is ever evaluated, after one or more calls to **Reset**, or after any call to **Step_** regardless of whether or not the statement has completed execution.

Invoking **Finalize** on a NULL pointer is a harmless no-op.

The application must finalize every prepared statement in order to avoid resource leaks. It is a grievous error for the application to try to use a prepared statement after it has been finalized. Any use of a prepared statement after it has been finalized can result in undefined and undesirable behavior such as segfaults and heap corruption.

# <a name="GetRow"></a>GetRow

After a prepared statement has been prepared using either **Prepare** this method must be called one or more times to evaluate the statement. **GetRow** is an alias for **Step_**.

```
FUNCTION GetRow () AS LONG
```

#### Return value

**SQLITE_BUSY** means that the database engine was unable to acquire the database locks it needs to do its job. If the statement is a COMMIT or occurs outside of an explicit transaction, then you can retry the statement. If the statement is not a COMMIT and occurs within an explicit transaction then you should rollback the transaction before continuing.

**SQLITE_DONE** means that the statement has finished executing successfully. **Step_** should not be called again on this virtual machine without first calling **Reset** to reset the virtual machine back to its initial state.

If the SQL statement being executed returns any data, then **SQLITE_ROW** is returned each time a new row of data is ready for processing by the caller. The values may be accessed using the column access functions. **Step_** is called again to retrieve the next row of data.

**SQLITE_ERROR** means that a run-time error (such as a constraint violation) has occurred. **GetRow** should not be called again on the Virtual Machine. More information may be found by calling **ErrMsg**.

**SQLITE_MISUSE** means that the this function was called inappropriately. Perhaps it was called on a prepared statement that has already been finalized or on one that had previously returned SQLITE_ERROR or SQLITE_DONE. Or it could be the case that the same database connection is being used by two or more threads at the same moment in time.

#### Remarks

For all versions of SQLite up to and including 3.6.23.1, a call to **Reset** was required after **Step_** returned anything other than SQLITE_ROW before any subsequent invocation of Step_. Failure to reset the prepared statement using **Reset** would result in an SQLITE_MISUSE return from **Step_**. But after version 3.6.23.1, **Step_** began calling **Reset** automatically in this circumstance rather than returning SQLITE_MISUSE. This is not considered a compatibility break because any application that ever receives an SQLITE_MISUSE error is broken by definition. The SQLITE_OMIT_AUTORESET compile-time option can be used to restore the legacy behavior.

# <a name="hStmt"></a>hStmt

Gets/sets the connection handle.

```
PROPERTY hStmt () AS sqlite3_stmt PTR
PROPERTY hStmt (BYVAL pStmt AS sqlite3_stmt PTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pStmt* | The statement handle. |

# <a name="IsColumnNull"></a>IsColumnNull

Returns true is the column value is null or false otherwise.

```
FUNCTION IsColumnNull () AS BOOLEAN
```

#### Return value

TRUE or FALSE.

# <a name="ReadOnly"></a>ReadOnly

Returns true if and only if the prepared statement makes no direct changes to the content of the database file.

```
FUNCTION ReadOnly () AS BOOLEAN
```

#### Return value

TRUE or FALSE.

# <a name="Reset"></a>Reset

Resets a prepared statement object back to its initial state, ready to be re-executed. Any SQL statement variables that had values bound to them using the **Bind_\***() methods retain their values. Use **ClearBindings** to reset the bindings.

```
FUNCTION Reset () AS LONG
```

#### Return value

If the most recent call to **Step_** for the prepared statement returned SQLITE_ROW or SQLITE_DONE, or if **Step_** has never before been called on the statement, then Reset returns SQLITE_OK.

If the most recent call to **Step_** for the prepared statement indicated an error, then **Reset** returns an appropriate error code.

# <a name="Sql"></a>Sql

Retrieves a saved copy of the original SQL text used to create a prepared statement if that statement was compiled using **Prepare**.

```
FUNCTIOn Sql () AS STRING
```

#### Return value

The SQL string associated with a prepared statement.

# <a name="Step_"></a>Step_

After a prepared statement has been prepared using either **Prepare** this method must be called one or more times to evaluate the statement. **GetRow** is an alias for **Step_**.

```
FUNCTION Step_ () AS LONG
```

#### Return value

**SQLITE_BUSY** means that the database engine was unable to acquire the database locks it needs to do its job. If the statement is a COMMIT or occurs outside of an explicit transaction, then you can retry the statement. If the statement is not a COMMIT and occurs within an explicit transaction then you should rollback the transaction before continuing.

**SQLITE_DONE** means that the statement has finished executing successfully. **Step_** should not be called again on this virtual machine without first calling **Reset** to reset the virtual machine back to its initial state.

If the SQL statement being executed returns any data, then **SQLITE_ROW** is returned each time a new row of data is ready for processing by the caller. The values may be accessed using the column access functions. **Step_** is called again to retrieve the next row of data.

**SQLITE_ERROR** means that a run-time error (such as a constraint violation) has occurred. **Step_** should not be called again on the Virtual Machine. More information may be found by calling **ErrMsg**.

**SQLITE_MISUSE** means that the this function was called inappropriately. Perhaps it was called on a prepared statement that has already been finalized or on one that had previously returned SQLITE_ERROR or SQLITE_DONE. Or it could be the case that the same database connection is being used by two or more threads at the same moment in time.

#### Remarks

For all versions of SQLite up to and including 3.6.23.1, a call to **Reset** was required after **Step_** returned anything other than SQLITE_ROW before any subsequent invocation of Step_. Failure to reset the prepared statement using **Reset** would result in an SQLITE_MISUSE return from **Step_**. But after version 3.6.23.1, **Step_** began calling **Reset** automatically in this circumstance rather than returning SQLITE_MISUSE. This is not considered a compatibility break because any application that ever receives an SQLITE_MISUSE error is broken by definition. The SQLITE_OMIT_AUTORESET compile-time option can be used to restore the legacy behavior.
