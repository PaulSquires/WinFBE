# CODBCStmt Class

Implements methods to create and manage statement objects. Inherits from CODBCBase.

**Include file**: CODBCStmt.inc

### Methods and Properties

| Name       | Description |
| ---------- | ----------- |
| [Constructors](#Constructors) | Allocates an statement handle. |
| [AddRecord](#AddRecord) | Adds a record to the database. |
| [BindCol](#BindCol) | Binds application data buffers to columns in the result set. |
| [BindParameter](#BindParameter) | Binds a buffer to a parameter marker in an SQL statement. |
| [BulkOperations](#BulkOperations) | Performs bulk insertions and bulk bookmark operations. |
| [Cancel](#Cancel) | Cancels the processing on a statement. |
| [CloseCursor](#CloseCursor) | Closes a cursor that has been opened on a statement and discards pending results. |
| [ColAttribute](#ColAttribute) | Returns descriptor information for a column in a result set. |
| [ColAutoUniqueValue](#ColAutoUniqueValue) | Checks if the column is an autoincrementing column or not. |
| [ColBaseColumnName](#ColBaseColumnName) | Returns the base column name for the result set column. |
| [ColBaseTableName](#ColBaseTableName) | Returns the name of the base table that contains the column. |
| [ColCaseSensitive](#ColCaseSensitive) | Checks if the column is treated as case-sensitive for collations and comparisons. |
| [ColCatalogName](#ColCatalogName) | Returns the catalog of the table that contains the column. |
| [ColConciseType](#ColConciseType) | Returns the concise data type. |
| [ColCount](#ColCount) | Returns the number of columns available in the result set. |
| [ColDisplaySize](#ColDisplaySize) | Returns the maximum number of characters required to display data from the column. |
| [ColFixedPrecScale](#ColFixedPrecScale) | Checks if column has a fixed precision and nonzero scale that are data source-specific. |
| [ColIsNull](#ColIsNull) | Checks if the column is null. |
| [ColLabel](#ColLabel) | Returns the column label or title. |
| [ColLength](#ColLength) | Returns the maximum or actual character length of a character string or binary data type. |
| [ColLiteralPrefix](#ColLiteralPrefix) | Returns the character or characters that the driver recognizes as a prefix for a literal of this data type. |
| [ColLiteralSuffix](#ColLiteralSuffix) | Returns the character or characters that the driver recognizes as a suffix for a literal of this data type. |
| [ColLocalTypeName](#ColLocalTypeName) | Returns the localized (native language) name for the data type that may be different from the regular name of the data type. |
| [ColName](#ColName) | Returns the column alias, if it applies. |
| [ColNullable](#ColNullable) | Checks if the column can have NULL values. |
| [ColNumPrecRadix](#ColNumPrecRadix) | Returns the column numeric precision radix. |
| [ColOctetLength](#ColOctetLength) | The length, in bytes, of a character string or binary data type. |
| [ColPrecision](#ColPrecision) | A numeric value that for a numeric data type denotes the applicable precision. |
| [ColScale](#ColScale) | A numeric value that is the applicable scale for a numeric data type. |
| [ColSchemaName](#ColSchemaName) | The schema of the table that contains the column. |
| [ColSearchable](#ColSearchable) | Indicates if the column data type is searchable. |
| [ColTableName](#ColTableName) | The name of the table that contains the column. |
| [ColType](#ColType) | A numeric value that specifies the SQL data type. |
| [ColTypeName](#ColTypeName) | Data source-dependent data type name. |
| [ColUnnamed](#ColUnnamed) | SQL_NAMED or SQL_UNNAMED. If the SQL_DESC_NAME field of the IRD contains a column alias or a column name, SQL_NAMED is returned. |
| [ColUnsigned](#ColUnsigned) | SQL_TRUE if the column is unsigned (or not numeric). SQL_FALSE if the column is signed. |
| [ColUpdatable](#ColUpdatable) | SQL_TRUE if the column is updatable; SQL_FALSE otherwise. |
| [DbcHandle](#DbcHandle) | Returns the connection handle. |
| [DeleteByBookmark](#DeleteByBookmark) | Deletes a set of rows where each row is identified by a bookmark. |
| [DeleteRecord](#DeleteRecord) | Deletes the sepcified row of data. |
| [DescribeCol](#DescribeCol) | Returns the result descriptor for one column in the result set. |
| [DescribeParam](#DescribeParam) | Returns the description of a parameter marker associated with a prepared SQL statement. |
| [ExecDirect](#ExecDirect) | Executes a preparable statement. |
| [Execute](#Execute) | Executes a prepared statement. |
| [ExtendedFetch](#ExtendedFetch) | Fetches the specified rowset of data from the result set and returns data for all bound columns. |
| [Fetch](#Fetch) | Fetches the next rowset of data from the result set and returns data for all bound columns. |
| [FetchByBookmark](#FetchByBookmark) | Fetches a set of rows where each row is identified by a bookmark. |
| [FetchScroll](#FetchScroll) | Fetches the specified rowset of data from the result set and returns data for all bound columns. |
| [GetColumnPrivileges](#GetColumnPrivileges) | Returns a list of columns and associated privileges for the specified table. |
| [GetColumns](#GetColumns) | Returns the list of column names in specified tables. |
| [GetCursorConcurrency](#GetCursorConcurrency) | Gets a SQLUINTEGER value that specifies the cursor concurrency. |
| [GetCursorKeysetSize](#GetCursorKeysetSize) | Gets a SQLUINTEGER value that specifies the number of rows in the keyset-driven cursor. |
| [GetCursorName](#GetCursorName) | Returns the cursor name associated with a specified statement. |
| [GetCursorScrollability](#GetCursorScrollability) | Gets a SQLUINTEGER value that specifies the scrollability type. |
| [GetCursorSensitivity](#GetCursorSensitivity) | Gets a SQLUINTEGER value that specifies whether cursors on the statement handle made to a result set by another cursor. |
| [GetCursorType](#GetCursorType) | Gets SQLUINTEGER value that specifies the cursor type. |
| [GetData](#GetData) | Retrieves data for a single column in the result set. It can be called multiple times to retrieve variable-length data in parts. |
| [GetDiagField](#GetDiagField) | Returns the current value of a field of a record of the diagnostic data structure (associated with an statement handle) that contains error, warning, and status information. |
| [GetDiagRec](#GetDiagRec) | Returns the current values of multiple fields of a diagnostic record that contains error, warning, and status information. |
| [GetErrorInfo](#GetErrorInfo) | Returns a verbose description of the last errors. |
| [GetForeignKeys](#GetForeignKeys) | Returns list of foreign keys of the table. |
| [GetImpParamDesc](#GetImpParamDesc) | Returns the handle of the IPD. |
| [GetImpParamDescField](#GetImpParamDescField) | Returns the current setting or value of a single field of a descriptor record. |
| [GetImpParamDescFieldName](#GetImpParamDescFieldName) | Returns the name of a single field of a descriptor record. |
| [GetImpParamDescFieldNullable](#GetImpParamDescFieldNullable) | Returns the nullable value (TRUE or FALSE) of a single field of a descriptor record. |
| [GetImpParamDescFieldOctetLength](#GetImpParamDescFieldOctetLength) | Returns the octet length of a single field of a descriptor record. |
| [GetImpParamDescFieldPrecision](#GetImpParamDescFieldPrecision) | Returns the precision value of a single field of a descriptor record. |
| [GetImpParamDescFieldScale](#GetImpParamDescFieldScale) | Returns the scale value of a single field of a descriptor record. |
| [GetImpParamDescFieldType](#GetImpParamDescFieldType) | Returns the type of a single field of a descriptor record. |
| [GetImpParamDescRec](#GetImpParamDescRec) | Returns the current settings or values of multiple fields of a descriptor record. |
| [GetImpRowDesc](#GetImpRowDesc) | Returns the handle to the IRD. |
| [GetImpRowDescField](#GetImpRowDescField) | Returns the current setting or value of a single field of a descriptor record. |
| [GetImpRowDescFieldName](#GetImpRowDescFieldName) | Returns the name of a single field of a descriptor record. |
| [GetImpRowDescFieldNullable](#GetImpRowDescFieldNullable) | Returns the nullable value (TRUE or FALSE) of a single field of a descriptor record. |
| [GetImpRowDescFieldOctetLength](#GetImpRowDescFieldOctetLength) | Returns the octet length of a single field of a descriptor record. |
| [GetImpRowDescFieldPrecision](#GetImpRowDescFieldPrecision) | Returns the precision value of a single field of a descriptor record. |
| [GetImpRowDescFieldScale](#GetImpRowDescFieldScale) | Returns the scale value of a single field of a descriptor record. |
| [GetImpRowDescFieldType](#GetImpRowDescFieldType) | Returns the type of a single field of a descriptor record. |
| [GetImpRowDescRec](#GetImpRowDescRec) | Returns the current settings or values of multiple fields of a descriptor record. |
| [GetLongVarcharData](#GetLongVarcharData) | Retrieves long variable char data from the specified column of the result set. |
| [GetPrimaryKeys](#GetPrimaryKeys) | Returns the column names that make up the primary key for a table. |
| [GetProcedureColumns](#GetProcedureColumns) | Returns the list of input and output parameters, as well as the columns that make up the result set for the specified procedures. |
| [GetProcedures](#GetProcedures) | Returns the list of procedure names stored in a specific data source. |
| [GetSpecialColumns](#GetSpecialColumns) | Retrieves information about columns within a specified table. |
| [GetSqlState](#GetSqlState) | Returns the SqlState for the statement handle. |
| [GetStatistics](#GetStatistics) | Retrieves a list of statistics about a single table and the indexes associated with the table. |
| [GetStmtAppParamDesc](#GetStmtAppParamDesc) | Gets the handle to the APD for subsequent calls to **Execute** and **ExecDirect** on the statement handle. |
| [GetStmtAppRowDesc](#GetStmtAppRowDesc) | Gets the handle to the ARD for subsequent fetches on the statement handle. |
| [GetStmtAsyncEnable](#GetStmtAsyncEnable) | Gets an SQLUINTEGER value that specifies whether a function called with the specified statement is executed asynchronously. |
| [GetStmtAttr](#GetStmtAttr) | Returns the current setting of a statement attribute. |
| [GetStmtFetchBookmarkPtr](#GetStmtFetchBookmarkPtr) | Gets a pointer that points to a binary bookmark value. |
| [GetStmtMaxLength](#GetStmtMaxLength) | Gets an SQLUINTEGER value that specifies the maximum amount of data that the driver returns from a character or binary column. |
| [GetStmtMaxRows](#GetStmtMaxRows) | Gets an SQLUINTEGER value corresponding to the maximum number of rows to return to the application for a SELECT statement. |
| [GetStmtNoScan](#GetStmtNoScan) | Gets an SQLUINTEGER value that indicates whether the driver should scan SQL strings for escape sequences. |
| [GetStmtParamBindOffsetPtr](#GetStmtParamBindOffsetPtr) | Gets an SQLUINTEGER value that points to an offset added to pointers to change binding of dynamic parameters. |
| [GetStmtParamBindType](#GetStmtParamBindType) | Gets an SQLUINTEGER value that indicates the binding orientation to be used for dynamic parameters. |
| [GetStmtParamOperationPtr](#GetStmtParamOperationPtr) | Gets an SQLUSMALLINT value that points to an array of SQLUSMALLINT values used to ignore a parameter during execution of an SQL statement. |
| [GetStmtParamsetSize](#GetStmtParamsetSize) | Gets an SQLUINTEGER value that specifies the number of values for each parameter. |
| [GetStmtParamsProcessedPtr](#GetStmtParamsProcessedPtr) | Gets an SQLUINTEGER record field that points to a buffer in which to return the number of sets of parameters that have been processed, including error sets. |
| [GetStmtParamStatusPtr](#GetStmtParamStatusPtr) | Gets an SQLUSMALLINT value that points to an array of SQLUSMALLINT values containing status information for each row of parameter values after a  call to **Execute** or **ExecDirect**. |
| [GetStmtQueryTimeout](#GetStmtQueryTimeout) | Gets an SQLUINTEGER value corresponding to the number of seconds to wait for an SQL statement to execute before returning to the application. |
| [GetStmtRetrieveData](#GetStmtRetrieveData) | Gets an SQLUINTEGER value specifying the data retrieval mode. |
| [GetStmtRowArraySize](#GetStmtRowArraySize) | Gets an SQLUINTEGER value that specifies the number of rows returned by each call to **Fetch** or **FetchScroll**. |
| [GetStmtRowBindOffsetPtr](#GetStmtRowBindOffsetPtr) | Gets an SQLUINTEGER value that points to an offset added to pointers to change binding of column data. |
| [GetStmtRowBindType](#GetStmtRowBindType) | Gets an SQLUINTEGER value that sets the binding orientation to be used when **Fetch** or **FetchScroll** is called on the associated statement. |
| [GetStmtRowNumber](#GetStmtRowNumber) | Returns an SQLUINTEGER value that is the number of the current row in the entire result set. |
| [GetStmtRowOperationPtr](#GetStmtRowOperationPtr) | Gets an SQLUSMALLINT value that points to an array of SQLUSMALLINT values used to ignore a row during a bulk operation using **SetPos**. |
| [GetStmtRowsFetchedPtr](#GetStmtRowsFetchedPtr) | Gets an SQLUINTEGER value that points to a buffer in which to return the number of rows fetched. |
| [GetStmtRowStatusPtr](#GetStmtRowStatusPtr) | Gets an SQLUSMALLINT value that points to an array of SQLUSMALLINT values containing row status values after a call to **Fetch** or **FetchScroll**. |
| [GetStmtSimulateCursor](#GetStmtSimulateCursor) | Gets an SQLUINTEGER value that specifies whether drivers that simulate positioned update and delete statements guarantee that such statements affect only one single row. |
| [GetStmtUseBookmarks](#GetStmtUseBookmarks) | Gets an SQLUINTEGER value that specifies whether an application will use bookmarks with a cursor. |
| [GetTablePrivileges](#GetTablePrivileges) | Returns a list of tables and the privileges associated with each table. |
| [GetTables](#GetTables) | Returns the list of table, catalog, or schema names, and table types, stored in a specific data source. |
| [GetTypeInfo](#GetTypeInfo) | Returns information about data types supported by the data source. |
| [Handle](#Handle) | Returns the statement handle. |
| [LockRecord](#LockRecord) | Sets the cursor position in a rowset and locks the record. |
| [MoreResults](#MoreResults) | Determines whether more results are available on a statement containing SELECT, UPDATE, INSERT, or DELETE statements and, if so, initializes processing for those results. |
| [Move](#Move) | Moves the cursor forward or backward the specified number of rowsets. |
| [MoveFirst](#MoveFirst) | Fetches the first rowset of data from the result set and returns data for all bound columns. |
| [MoveLast](#MoveLast) | Fetches the last rowset of data from the result set and returns data for all bound columns. |
| [MoveNext](#MoveNext) | Fetches the next rowset of data from the result set and returns data for all bound columns. |
| [MovePrevious](#MovePrevious) | Fetches the previous rowset of data from the result set and returns data for all bound columns. |
| [NumParams](#NumParams) | Returns the number of parameters in an SQL statement. |
| [NumResultCols](#NumResultCols) | Returns the number of columns in a result set. |
| [ParamData](#ParamData) | Used in conjunction with **PutData** to supply parameter data at statement execution time. |
| [Prepare](#Prepare) | Prepares an SQL string for execution. |
| [PutData](#PutData) | Allows an application to send data for a parameter or column to the driver at statement execution time. |
| [RecordCount](#RecordCount) | Gets the number of records in the result set. |
| [RefreshRecord](#RefreshRecord) | Sets the cursor position in a rowset and allows to refresh data in the rowset. |
| [ResetParams](#ResetParams) | Releases all parameter buffers set by **BindParameter** for the given statement handle. |
| [RowCount](#RowCount) | Returns the number of rows affected by update, insert or delete statements. |
| [SetAbsolutePosition](#SetAbsolutePosition) | Fetches the rowset starting at the specified row. |
| [SetCursorConcurrency](#SetCursorConcurrency) | Sets a SQLUINTEGER value that specifies the cursor concurrency. |
| [SetCursorKeysetSize](#SetCursorKeysetSize) | Sets a SQLUINTEGER value that specifies the number of rows in the keyset-driven cursor. |
| [SetCursorName](#SetCursorName) | Sets the cursor name associated with a specified statement. |
| [SetCursorScrollability](#SetCursorScrollability) | Sets a SQLUINTEGER value that specifies the scrollability type. |
| [SetCursorSensitivity](#SetCursorSensitivity) | Sets a SQLUINTEGER value that specifies whether cursors on the statement handle made to a result set by another cursor. |
| [SetCursorType](#SetCursorType) | Sets a SQLUINTEGER value that specifies the cursor type. |
| [SetDynamicCursor](#SetDynamicCursor) | Specifies a dynamic cursor. |
| [SetKeysetDrivenCursor](#SetKeysetDrivenCursor) | Specifies a keyset driven cursor. |
| [SetLockConcurrency](#SetLockConcurrency) | Cursor uses the lowest level of locking sufficient to ensure that the row can be updated. |
| [SetMultiuserKeysetCursor](#SetMultiuserKeysetCursor) | Creates a multiuser keyset cursor. |
| [SetOptimisticConcurrency](#SetOptimisticConcurrency) | Cursor uses optimistic concurrency control, comparing values. |
| [SetPos](#SetPos) | Sets the cursor position in a rowset and allows an application to refresh data in the rowset or to update or delete data in the result set. |
| [SetPosition](#SetPosition) | Sets the cursor position in a rowset. |
| [SetReadOnlyConcurrency](#SetReadOnlyConcurrency) | Cursor is read-only. No updates are allowed. |
| [SetRelativePosition](#SetRelativePosition) | Fetches the rowset from from the start of the current rowset. |
| [SetRowVerConcurrency](#SetRowVerConcurrency) | Cursor uses optimistic concurrency control, comparing row versions such as SQLBase ROWID or Sybase TIMESTAMP. |
| [SetStaticCursor](#SetStaticCursor) | Specifies an static cursor. |
| [SetStmtAppParamDesc](#SetStmtAppParamDesc) | Sets the handle to the APD for subsequent calls to **Execute** and **ExecDirect** on the statement handle. |
| [SetStmtAppRowDesc](#SetStmtAppRowDesc) | Sets the handle to the ARD for subsequent fetches on the statement handle. |
| [SetStmtAttr](#SetStmtAttr) | Sets attributes related to a statement. |
| [SetStmtFetchBookmarkPtr](#SetStmtFetchBookmarkPtr) | Sets a pointer that points to a binary bookmark value. |
| [SetStmtMaxLength](#SetStmtMaxLength) | Sets an SQLUINTEGER value that specifies the maximum amount of data that the driver returns from a character or binary column. |
| [SetStmtMaxRows](#SetStmtMaxRows) | Sets an SQLUINTEGER value corresponding to the maximum number of rows to return to the application for a SELECT statement. |
| [SetStmtNoScan](#SetStmtNoScan) | Sets an SQLUINTEGER value that indicates whether the driver should scan SQL strings for escape sequences. |
| [SetStmtParamBindOffsetPtr](#SetStmtParamBindOffsetPtr) | Sets an SQLUINTEGER value that points to an offset added to pointers to change binding of dynamic parameters. |
| [SetStmtParamBindType](#SetStmtParamBindType) | Sets an SQLUINTEGER value that indicates the binding orientation to be used for dynamic parameters. |
| [SetStmtParamOperationPtr](#SetStmtParamOperationPtr) | Sets an SQLUSMALLINT value that points to an array of SQLUSMALLINT values used to ignore a parameter during execution of an SQL statement. |
| [SetStmtParamsetSize](#SetStmtParamsetSize) | Sets an SQLUINTEGER value that specifies the number of values for each parameter. |
| [SetStmtParamsProcessedPtr](#SetStmtParamsProcessedPtr) | Sets an SQLUINTEGER record field that points to a buffer in which to return the number of sets of parameters that have been processed, including error sets. |
| [SetStmtParamStatusPtr](#SetStmtParamStatusPtr) | Sets an SQLUSMALLINT value that points to an array of SQLUSMALLINT values containing status information for each row of parameter values after a  call to Execute or ExecDirect. This field is required only if PARAMSET_SIZE is greater than 1. |
| [SetStmtQueryTimeout](#SetStmtParamStatusPtr) | Sets an SQLUINTEGER value corresponding to the number of seconds to wait for an SQL statement to execute before returning to the application. |
| [SetStmtRetrieveData](#SetStmtRetrieveData) | Sets an SQLUINTEGER value specifying the data retrieval mode. |
| [SetStmtRowArraySize](#SetStmtRowArraySize) | Sets an SQLUINTEGER value that specifies the number of rows returned by each call to **Fetch** or **FetchScroll**. |
| [SetStmtRowBindOffsetPtr](#SetStmtRowBindOffsetPtr) | Sets an SQLUINTEGER value that points to an offset added to pointers to change binding of column data. |
| [SetStmtRowBindType](#SetStmtRowBindType) | Sets an SQLUINTEGER value that sets the binding orientation to be used when **Fetch** or **FetchScroll** is called on the associated statement. |
| [SetStmtRowOperationPtr](#SetStmtRowOperationPtr) | Sets an SQLUSMALLINT value that points to an array of SQLUSMALLINT values used to ignore a row during a bulk operation using **SetPos**. |
| [SetStmtRowsFetchedPtr](#SetStmtRowsFetchedPtr) | Sets an SQLUINTEGER value that points to a buffer in which to return the number of rows fetched. |
| [SetStmtRowStatusPtr](#SetStmtRowStatusPtr) | Sets an SQLUSMALLINT value that points to an array of SQLUSMALLINT values containing row status values after a call to **Fetch** or **FetchScroll**. |
| [SetStmtSimulateCursor](#SetStmtSimulateCursor) | Sets an SQLUINTEGER value that specifies whether drivers that simulate positioned update and delete statements guarantee that such statements affect only one single row. |
| [SetStmtUseBookmarks](#SetStmtUseBookmarks) | Sets an SQLUINTEGER value that specifies whether an application will use bookmarks with a cursor. |
| [SetStmtAsyncEnable](#SetStmtAsyncEnable) | Sets an SQLUINTEGER value that specifies whether a function called with the specified statement is executed asynchronously. |
| [UnbindCol](#UnbindCol) | Unbinds the specified column buffer bound by **BindCol** for the given statement handle. |
| [UnbindColumns](#UnbindColumns) | Unbinds all column buffers bound by **BindCol** for the given statement handle. |
| [UnlockRecord](#UnlockRecord) | Sets the cursor position in a rowset and unlocks the record. |
| [UpdateByBookmark](#UpdateByBookmark) | Updates a set of rows where each row is identified by a bookmark. |
| [UpdateRecord](#UpdateRecord) | Updates a record. |

# <a name="Constructors"></a>Constructors

Allocates an statement handle. A statement handle provides access to statement information, such as error messages, the cursor name, and status information for SQL statement processing.

```
CONSTRUCTOR COdbcStmt (BYVAL pDbc AS COdbc PTR)
CONSTRUCTOR COdbcStmt (BYREF pCDbc AS COdbc)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pDbc* | Pointer to a connection object. |
| *pCDbc* | Reference to an instance of the CODBC class. |

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_INVALID_HANDLE, or SQL_ERROR.

### Example

```
#include once "Afx/COdbc/COdbc.inc"
USING Afx

' // Create a connection object and connect with the database
DIM wszConStr AS WSTRING * 260 = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=biblio.mdb"
DIM pDbc AS CODBC = wszConStr

' // Allocate an statement object
DIM pStmt AS COdbcStmt = pDbc

' // Generate a result set
pStmt.ExecDirect ("SELECT * FROM Authors ORDER BY Author")

' // Parse the result set
DIM cwsOutput AS CWSTR
DO
   ' // Fetch the record
   IF pStmt.Fetch = FALSE THEN EXIT DO
   ' // Get the values of the columns and display them
   cwsOutput = ""
   cwsOutput += pStmt.GetData(1) & " "
   cwsOutput += pStmt.GetData(2) & " "
   cwsOutput += pStmt.GetData(3)
   PRINT cwsOutput
LOOP
```

# <a name="AddRecord"></a>AddRecord

Adds a record to the database.

```
FUNCTION AddRecord () AS SQLRETURN
```

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_NEED_DATA, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

#### Example

```
#include once "Afx/COdbc/COdbc.inc"
USING Afx

' // Create a connection object and connect with the database
DIM wszConStr AS WSTRING * 260 = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=biblio.mdb"
DIM pDbc AS CODBC = wszConStr
IF pDbc.Handle = NULL THEN PRINT "Unable to create the connection handle" : SLEEP : END

' // Allocate an statement object
DIM pStmt AS COdbcStmt = @pDbc
IF pStmt.Handle = NULL THEN PRINT "Unable to create the statement handle" : SLEEP : END

' // Cursor type
pStmt.SetMultiuserKeysetCursor
' // Bind the columns
DIM AS LONG lAuId, cbAuId
pStmt.BindCol(1, @lAuId, @cbAuId)
DIM wszAuthor AS WSTRING * 256, cbAuthor AS LONG
pStmt.BindCol(2, @wszAuthor, SIZEOF(wszAuthor), @cbAuthor)
DIM iYearBorn AS SHORT, cbYearBorn AS LONG
pStmt.BindCol(3, @iYearBorn, @cbYearBorn)

' // Generate a result set
pStmt.ExecDirect ("SELECT * FROM Authors ORDER BY Author")

' // Fill the values of the bounded application variables and its sizes
lAuId     = 999               : cbAuID     = SIZEOF(lAuId)
wszAuthor = "Edgar Allan Poe" : cbAuthor   = LEN(wszAuthor)
iYearBorn = 1809              : cbYearBorn = SIZEOF(iYearBorn)
' // Add the record
pStmt.AddRecord
IF pStmt.Error = FALSE THEN PRINT "Record added" ELSE PRINT pStmt.GetErrorInfo
```

# <a name="BindCol"></a>BindCol

Binds application data buffers to columns in the result set. 

```
FUNCTION BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetType AS SQLSMALLINT, _
   BYVAL TargetValue AS SQLPOINTER, BYVAL BufferLength AS SQLLEN, _
   BYVAL StrLen_or_Ind AS ANY PTR) AS SQLRETURN
FUNCTION BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS BYTE PTR, _
   BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
FUNCTION BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS UBYTE PTR, _
   BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
FUNCTION BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS SHORT PTR, _
   BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
FUNCTION BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS USHORT PTR, _
   BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
FUNCTION BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS LONG PTR, _
   BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
FUNCTION BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS ULONG PTR, _
   BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
FUNCTION BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS SINGLE PTR, _
   BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
FUNCTION BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS DOUBLE PTR, _
   BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
FUNCTION BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS LONGINT PTR, _
   BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
FUNCTION BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS ULONGINT PTR, _
   BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
FUNCTION BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS ZSTRING PTR, _
   BYVAL BufferLenght AS SQLLEN, BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
FUNCTION BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS WSTRING PTR, _
   BYVAL BufferLenght AS SQLLEN, BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
FUNCTION BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS DATE_STRUCT PTR, _
   BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
FUNCTION  BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS TIME_STRUCT PTR, _
   BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
FUNCTION BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS TIMESTAMP_STRUCT PTR, _
   BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
FUNCTION BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS ANY PTR, _
   BYVAL BufferLenght AS SQLLEN, BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
FUNCTION BindColToBit (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS SHORT PTR, _
   BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
FUNCTION BindColToNumeric (BYVAL ColNumber AS SQLUSMALLINT, BYREF TargetValue AS ZSTRING, _
   BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
FUNCTION BindColToDecimal (BYVAL ColNumber AS SQLUSMALLINT, BYREF TargetValue AS ZSTRING, _
   BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColNumber* | Number of the result set column to bind. Columns are numbered in increasing column order starting at 0, where column 0 is the bookmark column. If bookmarks are not used — that is, the SQL_ATTR_USE_BOOKMARKS statement attribute is set to SQL_UB_OFF — then column numbers start at 1. |
| *TargetValue* | Pointer to the data buffer to bind to the column. **Fetch** and **FetchScroll** return data in this buffer. **BulkOperations** returns data in this buffer when *Operation* is SQL_FETCH_BY_BOOKMARK; it retrieves data from this buffer when *Operation* is SQL_ADD or SQL_UPDATE_BY_BOOKMARK. **SetPos** returns data in this buffer when *Operation* is SQL_REFRESH; it retrieves data from this buffer when *Operation* is SQL_UPDATE.<br>If *TargetValue* is a null pointer, the driver unbinds the data buffer for the column. An application can unbind all columns by calling **UnbindColumns** or **FreeStmt** with the SQL_UNBIND option. An application can unbind the data buffer for a column but still have a length/indicator buffer bound for the column, if the *TargetValue* argument in the call to **BindCol** is a null pointer but the *StrLen_or_IndPtr* argument is a valid value. |
| *StrLen_or_IndPtr* | Pointer to the length/indicator buffer to bind to the column. **Fetch** and **FetchScroll** return a value in this buffer. **BulkOperations** retrieves a value from this buffer when *Operation* is SQL_ADD, SQL_UPDATE_BY_BOOKMARK, or SQL_DELETE_BY_BOOKMARK. **BulkOperations** returns a value in this buffer when *Operation* is SQL_FETCH_BY_BOOKMARK. **SetPos** returns a value in this buffer when *Operation* is SQL_REFRESH; it retrieves a value from this buffer when *Operation* is SQL_UPDATE.<br><br>**Fetch**, **FetchScroll**, **BulkOperations**, and **SetPos** can return the following values in the length/indicator buffer:<br><ul><li>The length of the data available to return</li><li>SQL_NO_TOTAL</li><li>SQL_NULL_DATA</li></ul>The application can place the following values in the length/indicator buffer for use with **BulkOperations** or **SetPos**:<ul><li>The length of the data being sent</li><li>SQL_NTS</li><li>SQL_NULL_DATA</li><li>SQL_DATA_AT_EXEC</li><li>The result of the SQL_LEN_DATA_AT_EXEC macro</li><li>SQL_COLUMN_IGNORE</li></ul>If the indicator buffer and the length buffer are separate buffers, the indicator buffer can return only SQL_NULL_DATA, while the length buffer can return all other values.<br><br>If *StrLen_or_IndPtr* is a null pointer, no length or indicator value is used. This is an error when fetching data and the data is NULL.  |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_INVALID_HANDLE, or SQL_ERROR.

#### Example

```
#include once "Afx/COdbc/COdbc.inc"
USING Afx

' // Create a connection object and connect with the database
DIM wszConStr AS WSTRING * 260 = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=biblio.mdb"
DIM pDbc AS CODBC = wszConStr
IF pDbc.Handle = NULL THEN PRINT "Unable to create the connection handle" : SLEEP : END

' // Allocate an statement object
DIM pStmt AS COdbcStmt = @pDbc
IF pStmt.Handle = NULL THEN PRINT "Unable to create the statement handle" : SLEEP : END

' // Cursor type
pStmt.SetMultiuserKeysetCursor
' // Bind the columns
DIM AS LONG lAuId, cbAuId
pStmt.BindCol(1, @lAuId, @cbAuId)
DIM wszAuthor AS WSTRING * 260, cbAuthor AS LONG
pStmt.BindCol(2, @wszAuthor, SIZEOF(wszAuthor), @cbAuthor)
DIM iYearBorn AS SHORT, cbYearBorn AS LONG
pStmt.BindCol(3, @iYearBorn, @cbYearBorn)

' // Generate a result set
pStmt.ExecDirect ("SELECT * FROM Authors ORDER BY Author")

' // Parse the result set
DIM cwsOutput AS CWSTR
DO
   ' // Fetch the record
   IF pStmt.Fetch = FALSE THEN EXIT DO
   ' // Get the values of the columns and display them
   cwsOutput = ""
   cwsOutput += pStmt.GetData(1) & " "
   cwsOutput += pStmt.GetData(2) & " "
   cwsOutput += pStmt.GetData(3)
   PRINT cwsOutput
LOOP
```

# <a name="BindParameter"></a>BindParameter

Binds a buffer to a parameter marker in an SQL statement.

```
FUNCTION BindParameter (BYVAL ParameterNumber AS SQLUSMALLINT, BYVAL InputOutputType AS SQLSMALLINT, _
   BYVAL ValueType AS SQLSMALLINT, BYVAL ParameterType AS SQLSMALLINT, BYVAL ColumnSize AS SQLULEN, _
   BYVAL DecimalDigits AS SQLSMALLINT, BYVAL ParameterValuePtr AS SQLPOINTER, _
   BYVAL BufferLength AS SQLLEN, BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *ParameterNumber* | Parameter number, ordered sequentially in increasing parameter order, starting at 1. |
| *InputOutputType* | The type of the parameter. |
| *ValueType* | The C data type of the parameter. |
| *ParameterType* | The SQL data type of the parameter. |
| *ColumnSize* | The size of the column or expression of the corresponding parameter marker. |
| *DecimalDigits* | The decimal digits of the column or expression of the corresponding parameter marker. |
| *ParameterValuePtr* | A pointer to a buffer for the parameter's data. |
| *BufferLength* | Length of the ParameterValuePtr buffer in bytes. |
| *StrLen_or_IndPtr* | A pointer to a buffer for the parameter's length. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_INVALID_HANDLE, or SQL_ERROR.

# <a name="BulkOperations"></a>BulkOperations

Performs bulk insertions and bulk bookmark operations, including update, delete, and fetch by bookmark. 

```
FUNCTION BulkOperations (BYVAL Operation AS SQLSMALLINT) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *Operation* |Operation to perform: SQL_ADD, SQL_UPDATE_BY_BOOKMARK, SQL_DELETE_BY_BOOKMARK, SQL_FETCH_BY_BOOKMARK |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_NEED_DATA, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="Cancel"></a>Cancel

Cancels the processing on a statement.

```
FUNCTION Cancel () AS SQLRETURN
```

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="CloseCursor"></a>CloseCursor

Closes a cursor that has been opened on a statement and discards pending results. Returns SQLSTATE 24000 (Invalid cursor state) if no cursor is open in the statement.

```
FUNCTION CloseCursor () AS SQLRETURN
```

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.


# <a name="ColAttribute"></a>ColAttribute

Returns descriptor information for a column in a result set. Descriptor information is returned as a character string, a 32-bit descriptor-dependent value, or an integer value.

```
FUNCTION ColAttribute (BYVAL ColumnNumber AS SQLUSMALLINT, BYVAL FieldIdentifier AS SQLUSMALLINT, _
   BYVAL CharacterAttribute AS SQLPOINTER, BYVAL BufferLength AS SQLSMALLINT, _
   BYVAL StringLength AS SQLSMALLINT PTR, BYVAL NumericAttribute AS ANY PTR) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColumnNumber* | The number of the record in the IRD from which the field value is to be retrieved. This argument corresponds to the column number of result data, ordered sequentially in increasing column order, starting at 1. Columns can be described in any order. Column 0 can be specified in this argument, but all values except SQL_DESC_TYPE and SQL_DESC_OCTET_LENGTH will return undefined values. | 
| *FieldIdentifier* | The field in row *ColumnNumber* of the IRD that is to be returned. |
| *CharacterAttribute* | Pointer to a buffer in which to return the value in the *FieldIdentifier* field of the *ColumnNumber* row of the IRD, if the field is a character string. Otherwise, the field is unused. |
| *BufferLength* | If *FieldIdentifier* is an ODBC-defined field and *CharacterAttribute* points to a character string or binary buffer, this argument should be the length of *CharacterAttribute*. If *FieldIdentifier* is an ODBC-defined field and *CharacterAttribute* is an integer, this field is ignored. If *FieldIdentifier* is a driver-defined field, the application indicates the nature of the field to the Driver Manager by setting the *BufferLength* argument. *BufferLength* can have the following values:<br><br><ul><li>If *CharacterAttribute* is a pointer to a pointer, *BufferLength* should have the value SQL_IS_POINTER.</li><li>If *CharacterAttribute* is a pointer to a character string, the *BufferLength* is the length of the buffer.</li><li>If *CharacterAttribute* is a pointer to a binary buffer, the application places the result of the SQL_LEN_BINARY_ATTR(length) macro in *BufferLength*. This places a negative value in *BufferLength*.</li><li>If *CharacterAttribute* is a pointer to a fixed-length data type, *BufferLength* must be one of the following: SQL_IS_INTEGER, SQL_IS_UNINTEGER, SQL_SMALLINT, or %SQLUSMALLINT.</li></ul> |
| *StringLength* | Pointer to a buffer in which to return the total number of bytes (excluding the null-termination byte for character data) available to return in *CharacterAttribute*.<br>For character data, if the number of bytes available to return is greater than or equal to *BufferLength*, the descriptor information in *CharacterAttribute* is truncated to *BufferLength* minus the length of a null-termination character and is null-terminated by the driver.<br>For all other types of data, the value of *BufferLength* is ignored and the driver assumes the size of *CharacterAttribute* is 32 bits or 64 bits. |
| *NumericAttribute* | Pointer to an integer buffer in which to return the value in the *FieldIdentifier* field of the *ColumnNumber* row of the IRD, if the field is a numeric descriptor type, such as SQL_DESC_COLUMN_LENGTH. Otherwise, the field is unused. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

#### Example

```
#include once "Afx/COdbc/COdbc.inc"
USING Afx

' // Create a connection object and connect with the database
DIM wszConStr AS WSTRING * 260 = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=biblio.mdb"
DIM pDbc AS CODBC = wszConStr
IF pDbc.Handle = NULL THEN PRINT "Unable to create the connection handle" : SLEEP : END

' // Allocate an statement object
DIM pStmt AS COdbcStmt = @pDbc
IF pStmt.Handle = NULL THEN PRINT "Unable to create the statement handle" : SLEEP : END

' // Cursor type
pStmt.SetMultiuserKeysetCursor
' // Generate a result set
pStmt.ExecDirect ("SELECT * FROM Authors ORDER BY Author")

' // Retrieve the number of columns
DIM numCols AS SHORT = pStmt.NumResultCols
PRINT "Number of columns:" & STR(numCols)
IF numCols = 0 THEN PRINT "There are no columns": SLEEP : END
' // Retrieve the names of the fields (columns)
FOR idx AS LONG = 1 TO numCols
   PRINT "Field #" & STR(idx) & " name: " & pStmt.ColName(idx)
NEXT
' // Parse the result set
DO
   ' // Fetch the record
   IF pStmt.Fetch = FALSE THEN EXIT DO
   ' // Get the values of the columns and display them
   FOR idx AS LONG = 1 TO numCols
      PRINT pStmt.GetData(idx)
   NEXT
LOOP
```

# <a name="ColAutoUniqueValue"></a>ColAutoUniqueValue

Returns SQL_TRUE if the column is an autoincrementing column, SQL_FALSE if the column is not an autoincrementing column or is not numeric. This field is valid for numeric data type columns only. An application can insert values into a row containing an autoincrement number, but typically cannot update values in the column. When an insert is made into an autoincrement column, a unique value is inserted into the column at insert time. The increment is not defined, but is data source-specific. An application should not assume that an autoincrement column starts at any particular point or increments by any particular value.

```
FUNCTION ColAutoUniqueValue (BYVAL ColNum AS SQLUSMALLINT) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColNum* | Column number. |

#### Return value

TRUE or FALSE.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="ColBaseColumnName"></a>ColBaseColumnName

Returns the base column name for the result set column. If a base column name does not exist (as in the case of columns that are expressions), then this variable contains an empty string. This information is returned from the SQL_DESC_BASE_COLUMN_NAME record field of the IRD, which is a read-only field.

```
FUNCTION ColBaseColumnName (BYVAL ColNum AS SQLUSMALLINT) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColNum* | Column number. |

#### Return value

The base column name for the result set column.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="ColBaseTableName"></a>ColBaseTableName

Returns the name of the base table that contains the column. If the base table name cannot be defined or is not applicable, then this variable contains an empty string. This information is returned from the SQL_DESC_BASE_TABLE_NAME record field of the IRD, which is a read-only field.

```
FUNCTION ColBaseTableName (BYVAL ColNum AS SQLUSMALLINT) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColNum* | Column number. |

#### Return value

The name of the base table that contains the column.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="ColCaseSensitive"></a>ColCaseSensitive

Returns SQL_TRUE if the column is treated as case-sensitive for collations and comparisons. Returns SQL_FALSE if the column is not treated as case-sensitive for collations and comparisons or is noncharacter.

```
FUNCTION ColCaseSensitive (BYVAL ColNum AS SQLUSMALLINT) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColNum* | Column number. |

#### Return value

TRUE or FALSE

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="ColCatalogName"></a>ColCatalogName

Returns the catalog of the table that contains the column. The returned value is implementation-defined if the column is an expression or if the column is part of a view. If the data source does not support catalogs or the catalog name cannot be determined, an empty string is returned. This VARCHAR record field is not limited to 128 characters.

```
FUNCTION ColCatalogName (BYVAL ColNum AS SQLUSMALLINT) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColNum* | Column number. |

#### Return value

The catalog of the table that contains the column.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="ColConciseType"></a>ColConciseType

Returns the concise data type.

For the datetime and interval data types, this field returns the concise data type; for example, SQL_TYPE_TIME or SQL_INTERVAL_YEAR. This information is returned from the SQL_DESC_CONCISE_TYPE record field of the IRD.

```
FUNCTION ColConciseType (BYVAL ColNum AS SQLUSMALLINT) AS SQLINTEGER
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColNum* | Column number. |

#### Return value

The concise data type.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="ColCount"></a>ColCount

Returns the number of columns available in the result set.

```
FUNCTION ColCount () AS SQLINTEGER
```

#### Return value

The number of columns available in the result set. It returns 0 if there are no columns in the result set.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="ColDisplaySize"></a>ColDisplaySize

Returns the maximum number of characters required to display data from the column.

```
FUNCTION ColDisplaySize (BYVAL ColNum AS SQLUSMALLINT) AS SQLINTEGER
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColNum* | Column number. |

#### Return value

The maximum number of characters required to display data from the column.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="ColFixedPrecScale"></a>ColFixedPrecScale

Returns SQL_TRUE if the column has a fixed precision and nonzero scale that are data source-specific. Returns SQL_FALSE if the column does not have a fixed precision and nonzero scale that are data source-specific.

```
FUNCTION ColFixedPrecScale (BYVAL ColNum AS SQLUSMALLINT) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColNum* | Column number. |

#### Return value

TRUE or FALSE

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="ColIsNull"></a>ColIsNull

Checks if the column is null. Should not be used with a column that is currently binded to a variable or buffer or it will return an error. The binding functions already return an indicator in the last parameter.

```
FUNCTION ColIsNull (BYVAL ColNum AS SQLUSMALLINT) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColNum* | Column number. |

#### Return value

TRUE or FALSE

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_NO_DATA, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="ColLabel"></a>ColLabel

Returns the column label or title. For example, a column named EmpName might be labeled Employee Name or might be labeled with an alias. If a column does not have a label, the column name is returned. If the column is unlabeled and unnamed, an empty string is returned.

```
FUNCTION ColLabel (BYVAL ColNum AS SQLUSMALLINT) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColNum* | Column number. |

#### Return value

The column label or title.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="ColLength"></a>ColLength

Returns a numeric value that is either the maximum or actual character length of a character string or binary data type. It is the maximum character length for a fixed-length data type, or the actual character length for a variable-length data type. Its value always excludes the null-termination byte that ends the character string. This information is returned from the SQL_DESC_LENGTH record field of the IRD.

```
FUNCTION ColLength (BYVAL ColNum AS SQLUSMALLINT) AS SQLINTEGER
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColNum* | Column number. |

#### Return value

The maximum or actual character length of a character string or binary data type.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="ColLiteralPrefix"></a>ColLiteralPrefix

This VARCHAR(128) record field contains the character or characters that the driver recognizes as a prefix for a literal of this data type. This field contains an empty string for a data type for which a literal prefix is not applicable.

```
FUNCTION ColLiteralPrefix (BYVAL ColNum AS SQLUSMALLINT) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColNum* | Column number. |

#### Return value

The character or characters that the driver recognizes as a prefix for a literal of this data type.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="ColLiteralSuffix"></a>ColLiteralSuffix

This VARCHAR(128) record field contains the character or characters that the driver recognizes as a suffix for a literal of this data type. This field contains an empty string for a data type for which a literal suffix is not applicable.

```
FUNCTION ColLiteralSuffix (BYVAL ColNum AS SQLUSMALLINT) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColNum* | Column number. |

#### Return value

The character or characters that the driver recognizes as a suffix for a literal of this data type. 

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="ColLocalTypeName"></a>ColLocalTypeName

This VARCHAR(128) record field contains any localized (native language) name for the data type that may be different from the regular name of the data type. If there is no localized name, then an empty string is returned. This field is for display purposes only. The character set of the string is locale-dependent and is typically the default character set of the server.

```
FUNCTION ColLocalTypeName (BYVAL ColNum AS SQLUSMALLINT) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColNum* | Column number. |

#### Return value

The localized (native language) name for the data type that may be different from the regular name of the data type.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="ColName"></a>ColName

The column alias, if it applies. If the column alias does not apply, the column name is returned. In either case, SQL_DESC_UNNAMED is set to SQL_NAMED. If there is no column name or a column alias, an empty string is returned and SQL_DESC_UNNAMED is set to SQL_UNNAMED. This information is returned from the SQL_DESC_NAME record field of the IRD.

```
FUNCTION ColName (BYVAL ColNum AS SQLUSMALLINT) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColNum* | Column number. |

#### Return value

The column alias or name.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="ColNullable"></a>ColNullable

Returns SQL_NULLABLE if the column can have NULL values; SQL_NO_NULLS if the column does not have NULL values; or SQL_NULLABLE_UNKNOWN if it is not known whether the column accepts NULL values. This information is returned from the SQL_DESC_NULLABLE record field of the IRD.

```
FUNCTION ColNullable (BYVAL ColNum AS SQLUSMALLINT) AS SQLINTEGER
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColNum* | Column number. |

#### Return value

SQL_NULLABLE, SQL_NO_NULLS or SQL_NULLABLE_UNKNOWN.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="ColNumPrecRadix"></a>ColNumPrecRadix

If the data type in the SQL_DESC_TYPE field is an approximate numeric data type, this SQLINTEGER field contains a value of 2 because the SQL_DESC_PRECISION field contains the number of bits. If the data type in the SQL_DESC_TYPE field is an exact numeric data type, this field contains a value of 10 because the SQL_DESC_PRECISION field contains the number of decimal digits. This field is set to 0 for all non-numeric data types.

```
FUNCTION ColNumPrecRadix (BYVAL ColNum AS SQLUSMALLINT) AS SQLINTEGER
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColNum* | Column number. |

#### Return value

The column numeric precision radix.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="ColOctetLength"></a>ColOctetLength

The length, in bytes, of a character string or binary data type. For fixed-length character or binary types, this is the actual length in bytes. For variable-length character or binary types, this is the maximum length in bytes. This value includes the null terminator. This information is returned from the SQL_DESC_OCTET_LENGTH record field of the IRD.

```
FUNCTION ColOctetLength (BYVAL ColNum AS SQLUSMALLINT) AS SQLINTEGER
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColNum* | Column number. |

#### Return value

The length, in bytes, of a character string or binary data type.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="ColPrecision"></a>ColPrecision

A numeric value that for a numeric data type denotes the applicable precision. For data types SQL_TYPE_TIME, SQL_TYPE_TIMESTAMP, and all the interval data types that represent a time interval, its value is the applicable precision of the fractional seconds component. This information is returned from the SQL_DESC_PRECISION record field of the IRD.

```
FUNCTION ColPrecision (BYVAL ColNum AS SQLUSMALLINT) AS SQLINTEGER
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColNum* | Column number. |

#### Return value

The column applicable precision.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="ColScale"></a>ColScale

A numeric value that is the applicable scale for a numeric data type. For DECIMAL and NUMERIC data types, this is the defined scale. It is undefined for all other data types. This information is returned from the SCALE record field of the IRD.

```
FUNCTION ColScale (BYVAL ColNum AS SQLUSMALLINT) AS SQLINTEGER
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColNum* | Column number. |

#### Return value

The column applicable scale for a numeric data type.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="ColSchemaName"></a>ColSchemaName

The schema of the table that contains the column. The returned value is implementation-defined if the column is an expression or if the column is part of a view. If the data source does not support schemas or the schema name cannot be determined, an empty string is returned. This VARCHAR record field is not limited to 128 characters.

```
FUNCTION ColSchemaName (BYVAL ColNum AS SQLUSMALLINT) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColNum* | Column number. |

#### Return value

The schema of the table that contains the column.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="ColSearchable"></a>ColSearchable

Indicates if the column data type is searchable.

```
FUNCTION ColSearchable (BYVAL ColNum AS SQLUSMALLINT) AS SQLINTEGER
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColNum* | Column number. |

#### Return value

SQL_PRED_NONE if the column cannot be used in a WHERE clause. (This is the same as the SQL_UNSEARCHABLE value in ODBC 2.x.)

SQL_PRED_CHAR if the column can be used in a WHERE clause but only with the LIKE predicate. (This is the same as the SQL_LIKE_ONLY value in ODBC 2.x.)

SQL_PRED_BASIC if the column can be used in a WHERE clause with all the comparison operators except LIKE. (This is the same as the SQL_EXCEPT_LIKE value in ODBC 2.x.)

SQL_PRED_SEARCHABLE if the column can be used in a WHERE clause with any comparison operator.

Columns of type SQL_LONGVARCHAR and SQL_LONGVARBINARY usually return SQL_PRED_CHAR.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="ColTableName"></a>ColTableName

The name of the table that contains the column. The returned value is implementation-defined if the column is an expression or if the column is part of a view. If the table name cannot be determined, an empty string is returned.

```
FUNCTION ColTableName (BYVAL ColNum AS SQLUSMALLINT) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColNum* | Column number. |

#### Return value

The name of the table that contains the column.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="ColType"></a>ColType

A numeric value that specifies the SQL data type. When colNum is equal to 0, SQL_BINARY is returned for variable-length bookmarks and SQL_INTEGER is returned for fixed-length bookmarks. For the datetime and interval data types, this field returns the verbose data type: SQL_DATETIME or SQL_INTERVAL. This information is returned from the SQL_DESC_TYPE record field of the IRD.

```
FUNCTION ColType (BYVAL ColNum AS SQLUSMALLINT) AS SQLINTEGER
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColNum* | Column number. |

#### Return value

A numeric value that specifies the SQL data type.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="ColTypeName"></a>ColTypeName

Data source-dependent data type name; for example, "CHAR", "VARCHAR", "MONEY", "LONG VARBINARY", or "CHAR ( ) FOR BIT DATA". If the type is unknown, an empty string is returned.

```
FUNCTION ColTypeName (BYVAL ColNum AS SQLUSMALLINT) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColNum* | Column number. |

#### Return value

Data source-dependent data type name.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="ColUnnamed"></a>ColUnnamed

SQL_NAMED or SQL_UNNAMED. If the SQL_DESC_NAME field of the IRD contains a column alias or a column name, SQL_NAMED is returned. If there is no column name or column alias, SQL_UNNAMED is returned. This information is returned from the SQL_DESC_UNNAMED record field of the IRD.

```
FUNCTION ColUnnamed (BYVAL ColNum AS SQLUSMALLINT) AS SQLINTEGER
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColNum* | Column number. |

#### Return value

SQL_NAMED or SQL_UNNAMED.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="ColUnsigned"></a>ColUnsigned

SQL_TRUE if the column is unsigned (or not numeric). SQL_FALSE if the column is signed.

```
FUNCTION ColUnsigned (BYVAL ColNum AS SQLUSMALLINT) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColNum* | Column number. |

#### Return value

TRUE or FALSE.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="ColUpdatable"></a>ColUpdatable

SQL_TRUE if the column is updatable; SQL_FALSE otherwise.

```
FUNCTION ColUpdatable (BYVAL ColNum AS SQLUSMALLINT) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColNum* | Column number. |

#### Return value

TRUE or FALSE.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="DbcHandle"></a>DbcHandle

Returns the connection handle.

```
FUNCTION DbcHandle () AS SQLHANDLE
```

# <a name="DeleteByBookmark"></a>DeleteByBookmark

Deletes a set of rows where each row is identified by a bookmark.

```
FUNCTION DeleteByBookmark () AS SQLRETURN
```

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_NEED_DATA, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="DeleteRecord"></a>DeleteRecord

The driver positions the cursor on the row specified by RowNumber and deletes the underlying row of data. It changes the corresponding element of the row status array to SQL_ROW_DELETED. After the row has been deleted, the following are not valid for the row: positioned update and delete statements, calls to **GetData** and calls to **SetPos** with *Operation* set to anything except SQL_POSITION. For drivers that support packing, the row is deleted from the cursor when new data is retrieved from the data source. Whether the row remains visible depends on the cursor type. For example, deleted rows are visible to static and keyset-driven cursors but invisible to dynamic cursors. The row operation array pointed to by the SQL_ATTR_ROW_OPERATION_PTR statement attribute can be used to indicate that a row in the current rowset should be ignored during a bulk delete.

```
FUNCTION DeleteRecord (BYVAL wRow AS SQLSETPOSIROW = 1) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *RowNumber* | Row number inside the rowset. Note: *RowNumber* is the row number inside the rowset (if it is a single row rowset, RowNumber must be always 1). |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_NEED_DATA, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

#### Example

```
#include once "Afx/COdbc/COdbc.inc"
USING Afx

' // Connect with the database
DIM wszConStr AS WSTRING * 260 = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=biblio.mdb"
DIM pDbc AS CODBC = wszConStr

' // Allocate an statement object
DIM pStmt AS COdbcStmt = @pDbc
IF pStmt.Handle = NULL THEN PRINT "Failed to allocate the statement handle": SLEEP : END

' // Cursor type
pStmt.SetMultiuserKeysetCursor

' // Bind the columns
DIM AS LONG lAuId, cbAuId
pStmt.BindCol(1, @lAuId, @cbAuId)
DIM wszAuthor AS WSTRING * 256, cbAuthor AS LONG
pStmt.BindCol(2, @wszAuthor, 256, @cbAuthor)
DIM iYearBorn AS SHORT, cbYearBorn AS LONG
pStmt.BindCol(3, @iYearBorn, @cbYearBorn)

' // Generate a result set
pStmt.ExecDirect ("SELECT * FROM Authors WHERE Au_Id=999")

' // Fetch the record
pstmt.Fetch

' // Fill the values of the binded application variables and its sizes
cbAuID = SQL_COLUMN_IGNORE
wszAuthor = "Félix Lope de Vega Carpio"
cbAuthor = LEN(wszAuthor) * 2   ' Unicode uses 2 bytes per character
iYearBorn = 1562
cbYearBorn = SIZEOF(iYearBorn)

' // Delete the record
pStmt.DeleteRecord
IF pStmt.Error = FALSE THEN PRINT "Record deleted" ELSE PRINT pStmt.GetErrorInfo
```

# <a name="DescribeCol"></a>DescribeCol

Returns the result descriptor — column name, type, column size, decimal digits, and nullability — for one column in the result set. This information also is available in the fields of the IRD. 

```
FUNCTION DescribeCol (BYVAL ColumnNumber AS SQLUSMALLINT, BYVAL pwszColumnName AS WSTRING PTR, _
   BYVAL BufferLength AS SQLSMALLINT, BYVAL NameLength AS SQLSMALLINT PTR, _
   BYVAL DataType AS SQLSMALLINT PTR, BYVAL ColumnSizePtr AS SQLULEN PTR, _
   BYVAL DecimalDigits AS SQLSMALLINT PTR, BYVAL Nullable AS SQLSMALLINT PTR) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColNumber* | Column number of result data, ordered sequentially in increasing column order, starting at 1. The *ColNumber* argument can also be set to 0 to describe the bookmark column. |
| *ColName* | Pointer to a null-terminated buffer in which to return the column name. This value is read from the SQL_DESC_NAME field of the IRD. If the column is unnamed or the column name cannot be determined, the driver returns an empty string. |
| *DataType* | Pointer to a buffer in which to return the SQL data type of the column. This value is read from the SQL_DESC_CONCISE_TYPE field of the IRD. This will be one of the values in SQL Data Types, or a driver-specific SQL data type. If the data type cannot be determined, the driver returns SQL_UNKNOWN_TYPE.<br>In ODBC 3.x, SQL_TYPE_DATE, SQL_TYPE_TIME, or SQL_TYPE_TIMESTAMP is returned in DataType for date, time, or timestamp data, respectively; in ODBC 2.x, SQL_DATE, SQL_TIME, or SQL_TIMESTAMP is returned. The Driver Manager performs the required mappings when an ODBC 2.x application is working with an ODBC 3.x driver or when an ODBC 3.x application is working with an ODBC 2.x driver.<br>When ColNumber is equal to 0 (for a bookmark column), SQL_BINARY is returned in DataType for variable-length bookmarks. (SQL_INTEGER is returned if bookmarks are used by an ODBC 3.x application working with an ODBC 2.x driver or by an ODBC 2.x application working with an ODBC 3.x driver.) |
| *ColSize* | Pointer to a buffer in which to return the size (in characters) of the column on the data source. If the column size cannot be determined, the driver returns 0. |
| *DecimalDigits* | Pointer to a buffer in which to return the number of decimal digits of the column on the data source. If the number of decimal digits cannot be determined or is not applicable, the driver returns 0. |
| *Nullable* | Pointer to a buffer in which to return a value that indicates whether the column allows NULL values. This value is read from the SQL_DESC_NULLABLE field of the IRD. The value is one of the following:<br>SQL_NO_NULLS: The column does not allow NULL values.<br>SQL_NULLABLE: The column allows NULL values.<br>QL_NULLABLE_UNKNOWN: The driver cannot determine if the column allows NULL values.  |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

#### Example

```
#include once "Afx/COdbc/COdbc.inc"
USING Afx

' // Connect with the database
DIM wszConStr AS WSTRING * 260 = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=biblio.mdb"
DIM pDbc AS CODBC = wszConStr

' // Allocate an statement object
DIM pStmt AS COdbcStmt = @pDbc
IF pStmt.Handle = NULL THEN PRINT "Failed to allocate the statement handle": SLEEP : END

' // Cursor type
pStmt.SetMultiuserKeysetCursor

' // Generate a result set
pStmt.ExecDirect ("SELECT TOP 20 * FROM Authors ORDER BY Author")

' -------------------------------------------------------------------------------------
' Use DescribeCol to retrieve information about column 2
' -------------------------------------------------------------------------------------
DIM wszColName AS WSTRING * 260
DIM iNameLength AS SHORT
DIM iDataType AS SHORT
DIM dwColumnSize AS DWORD
DIM iDecimalDigits AS SHORT
DIM iNullable AS SHORT

pStmt.DescribeCol(2, @wszColName, 260, @iNameLength, @iDataType, @dwColumnSize, @iDecimalDigits, @iNullable)

? "Column name: " & wszColName
? "Name length: " & STR(iNameLength)
? "Data type: " & STR(iDataType)
? "Column size: " & STR(dwColumnSize)
? "Decimal digits: " & STR(iDecimalDigits)
? "Nullable: " & STR(iNullable) & " - " & IIF(iNullable, "TRUE", "FALSE")
```

# <a name="DescribeParam"></a>DescribeParam

Returns the description of a parameter marker associated with a prepared SQL statement. This information is also available in the fields of the IPD.

```
FUNCTION DescribeParam (BYVAL ParameterNumber AS SQLUSMALLINT, BYVAL DataType AS SQLSMALLINT PTR, _
   BYVAL ParameterSize AS SQLULEN PTR, BYVAL DecimalDigits AS SQLSMALLINT PTR, _
   BYVAL Nullable AS SQLSMALLINT PTR) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *ParameterNumber* | Parameter marker number ordered sequentially in increasing parameter order, starting at 1. |
| *DataType* | Pointer to a buffer in which to return the SQL data type of the parameter. This value is read from the SQL_DESC_CONCISE_TYPE record field of the IPD.<br>In ODBC 3.x, SQL_TYPE_DATE, SQL_TYPE_TIME, or SQL_TYPE_TIMESTAMP will be returned in DataType for date, time, or timestamp data, respectively; in ODBC 2.x, SQL_DATE, SQL_TIME, or SQL_TIMESTAMP will be returned. The Driver Manager performs the required mappings when an ODBC 2.x application is working with an ODBC 3.x driver or when an ODBC 3.x application is working with an ODBC 2.x driver.<br>When *ColumnNumber* is equal to 0 (for a bookmark column), SQL_BINARY is returned in *DataType* for variable-length bookmarks. (SQL_INTEGER is returned if bookmarks are used by an ODBC 3.x application working with an ODBC 2.x driver or by an ODBC 2.x application working with an ODBC 3.x driver.) |
| *ParameterSize* | Pointer to a buffer in which to return the size of the column or expression of the corresponding parameter marker as defined by the data source. |
| *DecimalDigits* | Pointer to a buffer in which to return the number of decimal digits of the column or expression of the corresponding parameter as defined by the data source. |
| *Nullable* | Pointer to a buffer in which to return a value that indicates whether the parameter allows NULL values. This value is read from the SQL_DESC_NULLABLE field of the IPD. One of the following:<br>SQL_NO_NULLS: The parameter does not allow NULL values (this is the default value).<br>SQL_NULLABLE: The parameter allows NULL values.<br>SQL_NULLABLE_UNKNOWN: The driver cannot determine if the parameter allows NULL values. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="ExecDirect"></a>ExecDirect

Executes a preparable statement, using the current values of the parameter marker variables if any parameters exist in the statement. **ExecDirect** is the fastest way to submit an SQL statement for one-time execution.

```
FUNCTION ExecDirect (BYREF wszSqlStr AS WSTRING) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszSqlStr* | SQL statement to be executed. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_NEED_DATA, SQL_STILL_EXECUTING, SQL_ERROR, SQL_NO_DATA, or SQL_INVALID_HANDLE.

#### Example

```
#include once "Afx/COdbc/COdbc.inc"
USING Afx

' // Connect with the database
DIM wszConStr AS WSTRING * 260 = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=biblio.mdb"
DIM pDbc AS CODBC = wszConStr

' // Allocate an statement object
DIM pStmt AS COdbcStmt = @pDbc
IF pStmt.Handle = NULL THEN PRINT "Failed to allocate the statement handle": SLEEP : END

' // Insert a new record
pStmt.ExecDirect ("INSERT INTO Authors (Au_ID, Author, [Year Born]) VALUES ('999', 'José Roca', 1950)")
IF pStmt.Error = FALSE THEN PRINT "Record added" ELSE PRINT pStmt.GetErrorInfo
```

# <a name="Execute"></a>Execute

Executes a prepared statement, using the current values of the parameter marker variables if any parameter markers exist in the statement.

```
FUNCTION Execute () AS SQLRETURN
```

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_NEED_DATA, SQL_STILL_EXECUTING, SQL_ERROR, SQL_NO_DATA, or SQL_INVALID_HANDLE.

# <a name="ExtendedFetch"></a>ExtendedFetch

Fetches the specified rowset of data from the result set and returns data for all bound columns. Rowsets can be specified at an absolute or relative position or by bookmark.

**Note**: In ODBC 3.x, **ExtendedFetch** has been replaced by FetchScroll. ODBC 3.x applications should not call **ExtendedFetch**; instead they should call **FetchScroll**. The Driver Manager maps **FetchScroll** to **ExtendedFetch** when working with an ODBC 2.x driver. 

```
FUNCTION ExtendedFetch (BYVAL FetchOrientation AS SQLUSMALLINT, BYVAL FetchOffset AS SQLLEN, _
   BYVAL RowCountPtr AS ANY PTR, BYVAL RowStatusArray AS SQLUSMALLINT PTR) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *FetchOrientation* | Type of fetch. This is the same as *FetchOrientation* in **FetchScroll**. |
| *FetchOffset* | Number of the row to fetch. This is the same as *FetchOffset* in **FetchScroll**, with one exception. When *FetchOrientation* is SQL_FETCH_BOOKMARK, *FetchOffset* is a fixed-length bookmark, not an offset from a bookmark. In other words, **ExtendedFetch** retrieves the bookmark from this argument, not the SQL_ATTR_FETCH_BOOKMARK_PTR statement attribute. It does not support variable-length bookmarks and does not support fetching a rowset at an offset (other than 0) from a bookmark. |
| *RowCountPtr* | Pointer to a buffer in which to return the number of rows actually fetched. This buffer is used in the same manner as the buffer specified by the SQL_ATTR_ROWS_FETCHED_PTR statement attribute. This buffer is used only by ExtendedFetch. It is not used by **Fetch** or **FetchScroll**. |
| *RowStatusArray* | Pointer to an array in which to return the status of each row. This array is used in the same manner as the array specified by the SQL_ATTR_ROW_STATUS_PTR statement attribute.<br>However, the address of this array is not stored in the SQL_DESC_STATUS_ARRAY_PTR field in the IRD. Furthermore, this array is used only by **ExtendedFetch** and by **BulkOperations** with an *Operation* of SQL_ADD or **SetPos** when it is called after **ExtendedFetch**. It is not used by **Fetch** or **FetchScroll**, and it is not used by **BulkOperations** or **SetPos** when they are called after **Fetch** or **FetchScroll**. It is also not used when **BulkOperations** with an *Operation* of SQL_ADD is called before any fetch function is called. In other words, it is used only in statement state S7. It is not used in statement states S5 or S6.<br>Applications should provide a valid pointer in the **RowStatusArray* argument; if not, the behavior of **ExtendedFetch** and the behavior of calls to **BulkOperations** or **SetPos** after a cursor has been positioned by **ExtendedFetch** are undefined. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_NO_DATA, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="Fetch"></a>Fetch

Fetches the next rowset of data from the result set and returns data for all bound columns. 

```
FUNCTION Fetch () AS BOOLEAN
```

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_NO_DATA, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="FetchByBookmark"></a>FetchByBookmark

Fetches a set of rows where each row is identified by a bookmark.

```
FUNCTION FetchByBookmark () AS SQLRETURN
```

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_NEED_DATA, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="FetchScroll"></a>FetchScroll

Fetches the specified rowset of data from the result set and returns data for all bound columns. Rowsets can be specified at an absolute or relative position or by bookmark.

When working with an ODBC 2.x driver, the Driver Manager maps this function to **ExtendedFetch**. 

```
FUNCTION FetchScroll (BYVAL FetchOrientation AS SQLSMALLINT, BYVAL FetchOffset AS SQLLEN) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *FetchOrientation* | Type of fetch: SQL_FETCH_NEXT, SQL_FETCH_PRIOR, SQL_FETCH_FIRST, SQL_FETCH_LAST, SQL_FETCH_ABSOLUTE, SQL_FETCH_RELATIVE, SQL_FETCH_BOOKMARK |
| *FetchOffset* | Number of the row to fetch. The interpretation of this argument depends on the value of the *FetchOrientation* argument. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_NO_DATA, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="GetColumnPrivileges"></a>GetColumnPrivileges

Returns a list of columns and associated privileges for the specified table. The driver returns the information as a result set on the specified statement handle.

```
FUNCTION GetColumnPrivileges (BYREF wszCatalogName AS WSTRING, BYVAL CatalogNameLength AS SQLSMALLINT, _
   BYREF wszSchemaName AS WSTRING, BYVAL SchemaNameLength AS SQLSMALLINT, _
   BYREF wszTableName AS WSTRING, BYVAL TableNameLength AS SQLSMALLINT, _
   BYREF wszColumnName AS WSTRING, BYVAL ColumnNameLength AS SQLSMALLINT) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszCatalogName* | Catalog name. If a driver supports names for some catalogs but not for others,such as when the driver retrieves data from different DBMSs, an empty string ("") denotes those catalogs that do not have names. *wszCatalogName* cannot contain a string search pattern. |
| *CatalogNameLength* | Length of *wszCatalogName*. |
| *wszSchemaName* | Schema name. If a driver supports schemas for some tables but not for others, such as when the driver retrieves data from different DBMSs, an empty string ("") denotes those tables that do not have schemas. *wszSchemaName* cannot contain a string search pattern. |
| *SchemaNameLength* | Length of *wszSchemaName*. |
| *wszTableName* | Table name. This argument cannot be a null pointer. *wszTableName* cannot contain a string search pattern. |
| *TableNameLength* | Length of *wszTableName*. |
| *wszColumnName* | String search pattern for column names. |
| *ColumnNameLength* | Length of *wszColumnName*. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

#### Diagnostics

When **ColumnPrivileges** returns SQL_ERROR or SQL_SUCCESS_WITH_INFO, an associated SQLSTATE value may be obtained by calling **GetDiagRec** with a *HandleType* of SQL_HANDLE_STMT and a *Handle* of *hStmt*.

# <a name="GetColumns"></a>GetColumns

Returns the list of column names in specified tables. The driver returns this information as a result set on the specified statement handle.

```
FUNCTION GetColumns (BYREF wszCatalogName AS WSTRING, BYVAL CatalogNameLength AS SQLSMALLINT, _
   BYREF wszSchemaName AS WSTRING, BYVAL SchemaNameLength AS SQLSMALLINT, _
   BYREF wszTableName AS WSTRING, BYVAL TableNameLength AS SQLSMALLINT, _
   BYREF wszColumnName AS WSTRING, BYVAL ColumnNameLength AS SQLSMALLINT) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszCatalogName* | Catalog name. If a driver supports names for some catalogs but not for others,such as when the driver retrieves data from different DBMSs, an empty string ("") denotes those catalogs that do not have names. *wszCatalogName* cannot contain a string search pattern. |
| *CatalogNameLength* | Length of *wszCatalogName*. |
| *wszSchemaName* | String search pattern for schema names. If a driver supports schemas for some tables but not for others, such as when the driver retrieves data from different DBMSs, an empty string ("") indicates those tables that do not have schemas. |
| *SchemaNameLength* | Length of *wszSchemaName*. |
| *wszTableName* | String search pattern for table names. |
| *TableNameLength* | Length of *wszTableName*. |
| *wszColumnName* | String search pattern for column names. |
| *ColumnNameLength* | Length of *wszColumnName*. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

#### Diagnostics

When **Columns** returns SQL_ERROR or SQL_SUCCESS_WITH_INFO, an associated SQLSTATE value can be obtained by calling **GetDiagRec** with a *HandleType* of SQL_HANDLE_STMT and a *Handle* of *hStmt*.

# <a name="GetCursorConcurrency"></a>GetCursorConcurrency

Gets a SQLUINTEGER value that specifies the cursor concurrency.

```
FUNCTION GetCursorConcurrency () AS SQLUINTEGER
```

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="GetCursorKeysetSize"></a>GetCursorKeysetSize

Gets a SQLUINTEGER value that specifies the number of rows in the keyset-driven cursor.

```
FUNCTION GetCursorKeysetSize () AS SQLUINTEGER
```

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.


# <a name="GetCursorName"></a>GetCursorName

Returns the cursor name associated with a specified statement.

```
FUNCTION GetCursorName () AS CWSTR
```

#### Remarks

For efficient processing, the cursor name should not include any leading or trailing spaces in the cursor name, and if the cursor name includes a delimited identifier, the delimiter should be positioned as the first character in the cursor name.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="GetCursorScrollability"></a>GetCursorScrollability

Gets a SQLUINTEGER value that specifies the scrollability type. Optional feature not implemented in Microsoft Access Driver.

```
FUNCTION GetCursorScrollability () AS SQLUINTEGER
```

#### Return value

The scrollability type.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="GetCursorSensitivity"></a>GetCursorSensitivity

Gets a SQLUINTEGER value that specifies whether cursors on the statement handle made to a result set by another cursor. Setting this attribute affects subsequent calls to ExecDirect and SQLExecute. An application can read back the value of this attribute to obtain its initial state or its state  as most recently set by the application.

**Note**: Optional feature not implemented in Microsoft Access Driver.

```
FUNCTION GetCursorSensitivity () AS SQLUINTEGER
```

#### Return value

The type of cursor sensitivity.

**SQL_UNSPECIFIED** = It is unspecified what the cursor type is and whether  cursors on the statement handle make visible the changes made to a result set by another cursor. Cursors on the statement handle may make visible none, some, or all such changes. This is the default.

**SQL_INSENSITIVE** = All cursors on the statement handle show the result set without reflecting any changes made to it by any other cursor, which has a concurrency that is read-only.

**SQL_SENSITIVE** = All cursors on the statement handle make visible all changes made to a result set by another cursor.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="GetCursorType"></a>GetCursorType

Gets SQLUINTEGER value that specifies the cursor type.

**Note**: Microsoft Access Driver changes SQL_CURSOR_DYNAMIC to SQL_CURSOR_KEYSET_DRIVEN.

```
FUNCTION GetCursorType () AS SQLUINTEGER
```

#### Return value

The type of cursor.

**SQL_CURSOR_FORWARD_ONLY** = The cursor only scrolls forward.

**SQL_CURSOR_STATIC** = The data in the result set is static.

**SQL_CURSOR_KEYSET_DRIVEN** = The driver saves and uses only the keys for the number of rows specified in the SQL_ATTR_KEYSET_SIZE statement attribute.

**SQL_CURSOR_DYNAMIC** = The driver saves and uses only the keys for the rows in the rowset.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.


# <a name="GetData"></a>GetData

Retrieves data for a single column in the result set. It can be called multiple times to retrieve variable-length data in parts.

```
FUNCTION GetData (BYVAL ColumnNumber AS SQLUSMALLINT, BYVAL TargetType AS SQLSMALLINT, _
   BYVAL TargetValue AS SQLPOINTER, BYVAL BufferLength AS SQLLEN, _
   BYVAL StrLen_or_Ind AS SQLLEN PTR) AS SQLRETURN
FUNCTION GetData (BYREF ColumnName AS WSTRING, BYVAL TargetType AS SQLSMALLINT, _
   BYVAL TargetValue AS SQLPOINTER, BYVAL BufferLength AS SQLLEN, _
   BYVAL StrLen_or_Ind AS SQLLEN PTR) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColumnNumber* | Number or name of the column for which to return data. Result set columns are numbered in increasing column order starting at 1. The bookmark column is column number 0; this can be specified only if bookmarks are enabled. |
| *TargetType* | The type identifier of the C data type of the *TargetValuePtr* buffer. If *TargetType* is SQL_ARD_TYPE, the driver uses the type identifier specified in the SQL_DESC_CONCISE_TYPE field of the ARD. If it is SQL_C_DEFAULT, the driver selects the default C data type based on the SQL data type of the source. |
| *TargetValue* | Pointer to the buffer in which to return the data. |
| *BufferLength* | Length of the *TargetValue* buffer in bytes.<br>The driver uses *BufferLength* to avoid writing past the end of the TargetValue buffer when returning variable-length data, such as character or binary data. Note that the driver counts the null-termination character when returning character data to *TargetValue*. *TargetValue* must therefore contain space for the null-termination character, or the driver will truncate the data.<br>When the driver returns fixed-length data, such as an integer or a date structure, the driver ignores *BufferLength* and assumes the buffer is large enough to hold the data. It is therefore important for the application to allocate a large enough buffer for fixed-length data or the driver will write past the end of the buffer.<br>**GetData** returns SQLSTATE HY090 (Invalid string or buffer length) when *BufferLength* is less than 0 but not when *BufferLength* is 0.<br>If *TargetValue* is set to a null pointer,*BufferLength* is ignored by the driver.  |
| *StrLen_or_Ind* | Pointer to the buffer in which to return the length or indicator value. If this is a null pointer, no length or indicator value is returned. This returns an error when the data being fetched is NULL.<br>**GetData** can return the following values in the length/indicator buffer:<br>The length of the data available to return: SQL_NO_TOTAL, SQL_NULL_DATA. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_NO_DATA, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

### Overloaded Functions

Retrieves data for a single column in the result set. It can be called multiple times to retrieve variable-length data in parts.

```
FUNCTION GetData (BYVAL ColumnNumber AS SQLUSMALLINT, BYVAL maxChars AS LONG = 256 * 2) AS CWSTR
FUNCTION GetData (BYREF ColumnName AS WSTRING, BYVAL maxChars AS LONG = 256 * 2) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColumnNumber* | Number or name of the column for which to return data. Result set columns are numbered in increasing column order starting at 1. The bookmark column is column number 0; this can be specified only if bookmarks are enabled. |
| *BufferLength* | Length of the *TargetValue* buffer in bytes.<br>The driver uses *BufferLength* to avoid writing past the end of the *TargetValue* buffer when returning variable-length data, such as character or binary data. Note that the driver counts the null-termination character when returning character data to *TargetValue*. *TargetValue* must therefore contain space for the null-termination character, or the driver will truncate the data.<br>When the driver returns fixed-length data, such as an integer or a date structure, the driver ignores *BufferLength* and assumes the buffer is large enough to hold the data. It is therefore important for the application to allocate a large enough buffer for fixed-length data or the driver will write past the end of the buffer.<br>**GetData** returns SQLSTATE HY090 (Invalid string or buffer length) when *BufferLength* is less than 0 but not when *BufferLength* is 0. If *TargetValue* is set to a null pointer, *BufferLength* is ignored by the driver.  |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_NO_DATA, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="GetDiagField"></a>GetDiagField

Returns the current value of a field of a record of the diagnostic data structure (associated with an statement handle) that contains error, warning, and status information.

```
FUNCTION GetDiagField (BYVAL RecNumber AS SQLSMALLINT, BYVAL DiagIdentifier AS SQLSMALLINT, _
   BYVAL DiagInfo AS SQLPOINTER, BYVAL BufferLength AS SQLSMALLINT, _
   BYVAL StringLength AS SQLSMALLINT PTR) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *RecNumber* | Indicates the status record from which the application seeks information. Status records are numbered from 1. If the *DiagIdentifier* argument indicates any field of the diagnostics header, *RecNumber* is ignored. If not, it should be greater than 0. |
| *DiagIdentifier* | Indicates the field of the diagnostic whose value is to be returned. |
| *DiagInfo* | Pointer to a buffer in which to return the diagnostic information. The data type depends on the value of *DiagIdentifier*. |
| *BufferLength* | If *DiagIdentifier* is an ODBC-defined diagnostic and DiagInfoPtr points to a character string or a binary buffer, this argument should be the length of *DiagInfo*. If *DiagIdentifier* is an ODBC-defined field and *DiagInfo* is an integer, *BufferLength* is ignored.<br>If *DiagIdentifier* is a driver-defined field, the application indicates the nature of the field to the Driver Manager by setting the *BufferLength* argument. *BufferLength* can have the following values:<ul><li>If *DiagInfo* is a pointer to a character string, then *BufferLength* is the length of the string or SQL_NTS.</li><li>If *DiagInfo* is a pointer to a binary buffer, then the application places the result of the SQL_LEN_BINARY_ATTR(length) macro in *BufferLength*. This places a negative value in *BufferLength*.</li><li>If *DiagInfoPtr* is a pointer to a value other than a character string or binary string, then *BufferLength* should have the value SQL_IS_POINTER.</li><li>If *DiagInfoPtr* is contains a fixed-length data type, then *BufferLength* is SQL_IS_INTEGER, SQL_IS_UINTEGER, SQL_IS_SMALLINT, or SQL_IS_USMALLINT, as appropriate.</li></ul> |
| *StringLength* | Pointer to a buffer in which to return the total number of bytes (excluding the number of bytes required for the null-termination character) available to return in *DiagInfo*, for character data. If the number of bytes available to return is greater than or equal to *BufferLength*, the text in *DiagInfo* is truncated to *BufferLength* minus the length of a null-termination character. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, SQL_INVALID_HANDLE, or SQL_NO_DATA.

# <a name="GetDiagRec"></a>GetDiagRec

Returns the current values of multiple fields of a diagnostic record that contains error, warning, and status information. Unlike **GetDiagField**, which returns one diagnostic field per call, **GetDiagRec** returns several commonly used fields of a diagnostic record, including the SQLSTATE, the native error code, and the diagnostic message text.

```
FUNCTION GetDiagRec (BYVAL RecNumber AS SQLSMALLINT, BYVAL Sqlstate AS WSTRING PTR, _
   BYVAL NativeError AS SQLINTEGER PTR, BYVAL MessageText AS WSTRING PTR, _
   BYVAL BufferLength AS SQLSMALLINT, BYVAL TextLength AS SQLSMALLINT PTR) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *RecNumber* | Indicates the status record from which the application seeks information. Status records are numbered from 1. |
| *SQLState* | Pointer to a buffer in which to return a five-character SQLSTATE code pertaining to the diagnostic record RecNumber. The first two characters indicate the class; the next three indicate the subclass. This information is contained in the SQL_DIAG_SQLSTATE diagnostic field. |
| *NativeError* | Pointer to a buffer in which to return the native error code, specific to the data source. This information is contained in the SQL_DIAG_NATIVE diagnostic field. |
| *MessageText* | Pointer to a buffer in which to return the diagnostic message text string. This information is contained in the SQL_DIAG_MESSAGE_TEXT diagnostic field. |
| *BufferLength* | Length of the *MessageText* buffer in characters. There is no maximum length of the diagnostic message text. |
| *TextLength* | Pointer to a buffer in which to return the total number of bytes (excluding the number of bytes required for the null-termination character) available to return in *MessageText*. If the number of bytes available to return is greater than *BufferLength*, the diagnostic message text in *MessageText* is truncated to *BufferLength* minus the length of a null-termination character. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="GetErrorInfo"></a>GetErrorInfo

Returns a verbose description of the last error(s).

```
FUNCTION GetErrorInfo (BYVAL iErrorCode AS SQLRETURN = 0) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *iErrorCode* | Optional. The error code returned by **GetLastResult**. |

# <a name="GetForeignKeys"></a>GetForeignKeys

ForeignKeys can return:

* A list of foreign keys in the specified table (columns in the specified table that refer to primary keys in other tables).
* A list of foreign keys in other tables that refer to the primary key in the specified table.

```
FUNCTION GetForeignKeys (BYREF wszPkCatalogName AS WSTRING, BYVAL pkCatalogNameLength AS SQLSMALLINT, _
   BYREF wszPkSchemaName AS WSTRING, BYVAL pkSchemaNameLength AS SQLSMALLINT, _
   BYREF wszPkTableName AS WSTRING, BYVAL pkTableNameLength AS SQLSMALLINT, _
   BYREF wszFkCatalogName AS WSTRING, BYVAL FkCatalogNameLength AS SQLSMALLINT, _
   BYREF wszFkSchemaName AS WSTRING, BYVAL FkSchemaNameLength AS SQLSMALLINT, _
   BYREF wszFkTableName AS WSTRING, BYVAL fkTableNameLength AS SQLSMALLINT) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPkCatalogName* | Primary key table catalog name. If a driver supports catalogs for some tables but not for others, such as when the driver retrieves data from different DBMSs, an empty string ("") denotes those tables that do not have catalogs. *wszPkCatalogName* cannot contain a string search pattern. |
| *CatalogNameLength* | Length of *wszPkCatalogName*. |
| *wszPkSchemaName* | Primary key table schema name. If a driver supports schemas for some tables but not for others, such as when the driver retrieves data from different DBMSs, an empty string ("") denotes those tables that do not have schemas. *wszPKSchemaName* cannot contain a string search pattern. |
| *PkSchemaNameLength* | Length of *wszPkSchemaName*. |
| *wszPkTableName* | Primary key table name. *wszPkTableName* cannot contain a string search pattern. |
| *PkTableNameLength* | Length of *wszPkTableName*. |
| *wszFkCatalogName* | Foreign key table catalog name. If a driver supports catalogs for some tables but not for others, such as when the driver retrieves data from different DBMSs, an empty string ("") denotes those tables that do not have catalogs. *wszFkCatalogName* cannot contain a string search pattern. |
| *FkCatalogNameLength* | Length of *wszFkCatalogName*. |
| *wszFkSchemaName* | Foreign key table schema name. If a driver supports schemas for some tables but not for others, such as when the driver retrieves data from different DBMSs, an empty string ("") denotes those tables that do not have schemas. *wszFkSchemaName* cannot contain a string search pattern. |
| *FkSchemaNameLength* | Length of *wszFkSchemaName*. |
| *wszFkTableName* | Foreign key table name. *wszFkTableName* cannot contain a string search pattern. |
| *FkTableNameLength* | Length of *wszFkTableName*. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

#### Diagnostics

When **ForeignKeys** returns SQL_ERROR or SQL_SUCCESS_WITH_INFO, an associated SQLSTATE value can be obtained by calling **GetDiagRec** with a *HandleType* of SQL_HANDLE_STMT and a *Handle* of *hStmt*.

# <a name="GetImpParamDesc"></a>GetImpParamDesc

Returns the handle to the IPD (Implicitily Parameter Descriptor). The value of this attribute is the descriptor allocated when the statement was initially allocated. The application cannot set this attribute. This attribute can be retrieved by a call to **GetStmtAttr** but not set by a call to **SetStmtAttr**.

```
FUNCTION GetImpParamDesc () AS SQLUINTEGER
```

#### Return value

The handle of IPD.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="GetImpParamDescField"></a>GetImpParamDescField

Returns the current setting or value of a single field of a descriptor record. The field returned describe the name, data type, and storage of column or parameter data.

```
FUNCTION GetImpParamDescField (BYVAL RecNumber AS SQLSMALLINT, BYVAL FieldIdentifier AS SQLSMALLINT, _
   BYVAL ValuePtr AS SQLPOINTER, BYVAL BufferLength AS SQLINTEGER, _
   BYVAL StringLength AS SQLINTEGER PTR) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *RecNumber* | Indicates the descriptor record from which the application seeks information. Descriptor records are numbered from 1, with record number 0 being the bookmark record. The *RecNumber* argument must be less or equal to the value of SQL_DESC_COUNT. If *RecNumber* is less that or equal to SQL_DESC_COUNT but the row does not contain data for a column or parameter, a call to **GetImpRowDesc** will return the default values of the fields. |
| *FieldIdentifier* | Indicates the field of the descriptor whose value is to be returned: SQL_DESC_NAME, SQL_DESC_TYPE, SQL_DESC_OCTET_LENGTH, SQL_DESC_PRECISION, SQL_DESC_SCALE, SQL_DESC_NULLABLE. |
| *ValuePtr* | Pointer to a buffer in which to return the descriptor information. The data type depends on the value of *FieldIdentifier*. |
| *BufferLength* | If *FieldIdentifier* is an ODBC-defined field and *ValuePtr* points to a character string or a binary buffer, this argument should be the length of *ValuePtr*. If *FieldIdentifier* is an ODBC-defined field and *ValuePtr* is an integer, *BufferLength* is ignored.<br><br>If *FieldIdentifier* is a driver-defined field, the application indicates the nature of the field to the Driver Manager by setting the *BufferLength* argument. *BufferLength* can have the following values:<br><br><ul><li>If *ValuePtr* is a pointer to a character string, then *BufferLength* is the length of the string or SQL_NTS.</li><li>If *ValuePtr* is a pointer to a binary buffer, then the application places the result of the SQL_LEN_BINARY_ATTR(length) macro in *BufferLength*. This places a negative value in *BufferLength*.</li><li>If *ValuePtr* is a pointer to a value other than a character string or binary string, then *BufferLength* should have the value SQL_IS_POINTER.</li><li>If *ValuePtr* is contains a fixed-length data type, then *BufferLength* is either SQL_IS_INTEGER, SQL_IS_UINTEGER, SQL_IS_SMALLINT, or SQL_IS_USMALLINT, as appropriate.</li>/ul> |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, SQL_NO_DATA, or SQL_INVALID_HANDLE.

SQL_NO_DATA is returned if *RecNumber* is greater than the current number of descriptor records or the statement is in the prepared or executed state but there was no open cursor associated with it.

# <a name="GetImpParamDescFieldName"></a>GetImpParamDescFieldName

Returns the name of a single field of a descriptor record.

```
FUNCTION GetImpParamDescFieldName (BYVAL RecNumber AS SQLSMALLINT) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *RecNumber* | Indicates the descriptor record from which the application seeks information. Descriptor records are numbered from 1, with record number 0 being the bookmark record. The *RecNumber* argument must be less or equal to the value of SQL_DESC_COUNT. |

#### Return value

The name of the field.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, SQL_NO_DATA, or SQL_INVALID_HANDLE.

SQL_NO_DATA is returned if *RecNumber* is greater than the current number of descriptor records or the statement is in the prepared or executed state but there was no open cursor associated with it.

# <a name="GetImpParamDescFieldNullable"></a>GetImpParamDescFieldNullable

Returns the nullable value (TRUE or FALSE) of a single field of a descriptor record.

```
FUNCTION GetImpParamDescFieldNullable (BYVAL RecNumber AS SQLSMALLINT) AS SQLSMALLINT
```

| Parameter  | Description |
| ---------- | ----------- |
| *RecNumber* | Indicates the descriptor record from which the application seeks information. Descriptor records are numbered from 1, with record number 0 being the bookmark record. The *RecNumber* argument must be less or equal to the value of SQL_DESC_COUNT. |

#### Return value

TRUE or FALSE.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, SQL_NO_DATA, or SQL_INVALID_HANDLE.

SQL_NO_DATA is returned if *RecNumber* is greater than the current number of descriptor records or the statement is in the prepared or executed state but there was no open cursor associated with it.

# <a name="GetImpParamDescFieldOctetLength"></a>GetImpParamDescFieldOctetLength

Returns the octet length of a single field of a descriptor record.

```
FUNCTION GetImpParamDescFieldOctetLength (BYVAL RecNumber AS SQLSMALLINT) AS SQLSMALLINT
```

| Parameter  | Description |
| ---------- | ----------- |
| *RecNumber* | Indicates the descriptor record from which the application seeks information. Descriptor records are numbered from 1, with record number 0 being the bookmark record. The *RecNumber* argument must be less or equal to the value of SQL_DESC_COUNT. |

#### Return value

The octet length.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, SQL_NO_DATA, or SQL_INVALID_HANDLE.

SQL_NO_DATA is returned if *RecNumber* is greater than the current number of descriptor records or the statement is in the prepared or executed state but there was no open cursor associated with it.

# <a name="GetImpParamDescFieldPrecision"></a>GetImpParamDescFieldPrecision

Returns the precision value of a single field of a descriptor record.

```
FUNCTION GetImpParamDescFieldPrecision (BYVAL RecNumber AS SQLSMALLINT) AS SQLSMALLINT
```

| Parameter  | Description |
| ---------- | ----------- |
| *RecNumber* | Indicates the descriptor record from which the application seeks information. Descriptor records are numbered from 1, with record number 0 being the bookmark record. The *RecNumber* argument must be less or equal to the value of SQL_DESC_COUNT. |

#### Return value

The precision value.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, SQL_NO_DATA, or SQL_INVALID_HANDLE.

SQL_NO_DATA is returned if *RecNumber* is greater than the current number of descriptor records or the statement is in the prepared or executed state but there was no open cursor associated with it.

# <a name="GetImpParamDescFieldScale"></a>GetImpParamDescFieldScale

Returns the scale value of a single field of a descriptor record.

```
FUNCTION GetImpParamDescFieldScale (BYVAL RecNumber AS SQLSMALLINT) AS SQLSMALLINT
```

| Parameter  | Description |
| ---------- | ----------- |
| *RecNumber* | Indicates the descriptor record from which the application seeks information. Descriptor records are numbered from 1, with record number 0 being the bookmark record. The *RecNumber* argument must be less or equal to the value of SQL_DESC_COUNT. |

#### Return value

The scale value.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, SQL_NO_DATA, or SQL_INVALID_HANDLE.

SQL_NO_DATA is returned if *RecNumber* is greater than the current number of descriptor records or the statement is in the prepared or executed state but there was no open cursor associated with it.

# <a name="GetImpParamDescFieldType"></a>GetImpParamDescFieldType

Returns the type of a single field of a descriptor record.

```
FUNCTION GetImpParamDescFieldType (BYVAL RecNumber AS SQLSMALLINT) AS SQLSMALLINT
```

| Parameter  | Description |
| ---------- | ----------- |
| *RecNumber* | Indicates the descriptor record from which the application seeks information. Descriptor records are numbered from 1, with record number 0 being the bookmark record. The *RecNumber* argument must be less or equal to the value of SQL_DESC_COUNT. |

#### Return value

The type of the field.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, SQL_NO_DATA, or SQL_INVALID_HANDLE.

SQL_NO_DATA is returned if *RecNumber* is greater than the current number of descriptor records or the statement is in the prepared or executed state but there was no open cursor associated with it.

# <a name="GetImpParamDescRec"></a>GetImpParamDescRec

Returns the current settings or values of multiple fields of a descriptor record. The fields returned describe the name, data type, and storage of column or parameter data.

```
FUNCTION GetImpParamDescRec (BYVAL RecNumber AS SQLSMALLINT, BYVAL pwszName AS WSTRING PTR, _
   BYVAL BufferLength AS SQLSMALLINT, BYVAL StringLength AS SQLSMALLINT PTR, _
   BYVAL TypePtr AS SQLSMALLINT PTR, BYVAL SubTypePtr AS SQLSMALLINT PTR, _
   BYVAL LengthPtr AS SQLLEN PTR, BYVAL PrecisionPtr AS SQLSMALLINT PTR, _
   BYVAL ScalePtr AS SQLSMALLINT PTR, BYVAL NullablePtr AS SQLSMALLINT PTR) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *RecNumber* | Indicates the descriptor record from which the application seeks information. Descriptor records are numbered 1, with record number 0 being the bookmark record. The *RecNumber* argument must be less than or equal to the value of SQL_DESC_COUNT. If RecNumber is less than or equal to SQL_DESC_COUNT but the row does not contain data for a column or parameter, a call to **GetDescRec** will return the default values of the fields.ates the descriptor record from which the application seeks information. Descriptor records are numbered.  |
| *pwszName* | A pointer to a buffer in which to return the SQL_DESC_NAME field for the descriptor record. |
| *TypePtr* | A pointer to a buffer in which to return the value of the SQL_DESC_TYPE field for the descriptor record. |
| *SubTypePtr* | For records whose type is SQL_DATETIME or SQL_INTERVAL, this is a pointer to a buffer in which to return the value of the SQL_DESC_DATETIME_INTERVAL_CODE field. |
| *LengthPtr* | A pointer to a buffer in which to return the value of the SQL_DESC_OCTET_LENGTH field for the descriptor record. |
| *PrecisionPtr* | A pointer to a buffer in which to return the value of the SQL_DESC_PRECISION field for the descriptor record. |
| *ScalePtr* | A pointer to a buffer in which to return the value of the SQL_DESC_SCALE field for the descriptor record. |
| *NullablePtr* | A pointer to a buffer in which to return the value of the SQL_DESC_NULLABLE field for the descriptor record. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, SQL_NO_DATA, or SQL_INVALID_HANDLE.

SQL_NO_DATA is returned if *RecNumber* is greater than the current number of descriptor records or the the statement is in the prepared or executed state but there was no open cursor associated with it.

# <a name="GetImpRowDesc"></a>GetImpRowDesc

Returns the handle to the IRD. The value of this attribute is the descriptor  allocated when the statement was initially allocated. The application  cannot set this attribute.

This attribute can be retrieved by a call to **GetStmtAttr** but not set by a call to **SetStmtAttr**.

```
FUNCTION GetImpRowDesc () AS SQLUINTEGER
```

#### Return value

The handle of the IRD.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="GetImpRowDescField"></a>GetImpRowDescField

Returns the current setting or value of a single field of a descriptor record. The field returned describe the name, data type, and storage of column or parameter data.

```
FUNCTION GetImpRowDescField (BYVAL RecNumber AS SQLSMALLINT, BYVAL FieldIdentifier AS SQLSMALLINT, _
   BYVAL ValuePtr AS SQLPOINTER, BYVAL BufferLength AS SQLINTEGER, _
   BYVAL StringLength AS SQLINTEGER PTR) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *RecNumber* | Indicates the descriptor record from which the application seeks information. Descriptor records are numbered from 1, with record number 0 being the bookmark record. The *RecNumber* argument must be less or equal to the value of SQL_DESC_COUNT. If RecNumber is less that or equal to SQL_DESC_COUNT but the row does not contain data for a column or parameter, a call to **GetImpRowDesc** will return the default values of the fields. |
| *FieldIdentifier* | Indicates the field of the descriptor whose value is to be returned: SQL_DESC_NAME, SQL_DESC_TYPE, SQL_DESC_OCTET_LENGTH, SQL_DESC_PRECISION, SQL_DESC_SCALE, SQL_DESC_NULLABLE. |
| *ValuePtr* | Pointer to a buffer in which to return the descriptor information. The data type depends on the value of *FieldIdentifier*. |
| *BufferLength* | If *FieldIdentifier* is an ODBC-defined field and ValuePtr points to a character string or a binary buffer, this argument should be the length of *ValuePtr*. If *FieldIdentifier* is an ODBC-defined field and ValuePtr is an integer, *BufferLength* is ignored.<br>If *FieldIdentifier* is a driver-defined field, the application indicates the nature of the field to the Driver Manager by setting the BufferLength argument. *BufferLength* can have the following values:<br><br><ul><li>If *ValuePtr* is a pointer to a character string, then  BufferLengthis the length of the string or SQL_NTS.</li><li>If *ValuePtr* is a pointer to a binary buffer, then the application places the result of the SQL_LEN_BINARY_ATTR(length) macro in *BufferLength*. This places a negative value in *BufferLength*.</li><li>If *ValuePtr* is a pointer to a value other than a character string or binary string, then *BufferLength* should have the value SQL_IS_POINTER.</li><li>If *ValuePtr* is contains a fixed-length data type, then *BufferLength* is either SQL_IS_INTEGER, SQL_IS_UINTEGER, SQL_IS_SMALLINT, or SQL_IS_USMALLINT, as appropriate.</li></ul> |

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, SQL_NO_DATA, or SQL_INVALID_HANDLE.

SQL_NO_DATA is returned if *RecNumber* is greater than the current number of descriptor records.

# <a name="GetImpRowDescFieldName"></a>GetImpRowDescFieldName

Returns the name of a single field of a descriptor record.

```
FUNCTION GetImpRowDescFieldName (BYVAL RecNumber AS SQLSMALLINT) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *RecNumber* | Indicates the descriptor record from which the application seeks information. Descriptor records are numbered from 1, with record number 0 being the bookmark record. The *RecNumber* argument must be less or equal to the value of SQL_DESC_COUNT. |

#### Return value

The name of the field.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, SQL_NO_DATA, or SQL_INVALID_HANDLE.

SQL_NO_DATA is returned if *RecNumber* is greater than the current number of descriptor records.

# <a name="GetImpRowDescFieldNullable"></a>GetImpRowDescFieldNullable

Returns the nullable value (TRUE or FALSE) of a single field of a descriptor record.

```
FUNCTION GetImpRowDescFieldNullable (BYVAL RecNumber AS SQLSMALLINT) AS SQLSMALLINT
```

| Parameter  | Description |
| ---------- | ----------- |
| *RecNumber* | Indicates the descriptor record from which the application seeks information. Descriptor records are numbered from 1, with record number 0 being the bookmark record. The *RecNumber* argument must be less or equal to the value of SQL_DESC_COUNT. |

#### Return value

TRUE or FALSE.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, SQL_NO_DATA, or SQL_INVALID_HANDLE.

SQL_NO_DATA is returned if *RecNumber* is greater than the current number of descriptor records.

# <a name="GetImpRowDescFieldOctetLength"></a>GetImpRowDescFieldOctetLength

Returns the octet length of a single field of a descriptor record.

```
FUNCTION GetImpRowDescFieldOctetLength (BYVAL RecNumber AS SQLSMALLINT) AS SQLSMALLINT
```

| Parameter  | Description |
| ---------- | ----------- |
| *RecNumber* | Indicates the descriptor record from which the application seeks information. Descriptor records are numbered from 1, with record number 0 being the bookmark record. The *RecNumber* argument must be less or equal to the value of SQL_DESC_COUNT. |

#### Return value

The octet length.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, SQL_NO_DATA, or SQL_INVALID_HANDLE.

SQL_NO_DATA is returned if *RecNumber* is greater than the current number of descriptor records.

# <a name="GetImpRowDescFieldPrecision"></a>GetImpRowDescFieldPrecision

Returns the precision value of a single field of a descriptor record.

```
FUNCTION GetImpRowDescFieldPrecision (BYVAL RecNumber AS SQLSMALLINT) AS SQLSMALLINT
```

| Parameter  | Description |
| ---------- | ----------- |
| *RecNumber* | Indicates the descriptor record from which the application seeks information. Descriptor records are numbered from 1, with record number 0 being the bookmark record. The *RecNumber* argument must be less or equal to the value of SQL_DESC_COUNT. |

#### Return value

The precision value.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, SQL_NO_DATA, or SQL_INVALID_HANDLE.

SQL_NO_DATA is returned if *RecNumber* is greater than the current number of descriptor records.

# <a name="GetImpRowDescFieldScale"></a>GetImpRowDescFieldScale

Returns the scale value of a single field of a descriptor record.

```
FUNCTION GetImpRowDescFieldScale (BYVAL RecNumber AS SQLSMALLINT) AS SQLSMALLINT
```

| Parameter  | Description |
| ---------- | ----------- |
| *RecNumber* | Indicates the descriptor record from which the application seeks information. Descriptor records are numbered from 1, with record number 0 being the bookmark record. The *RecNumber* argument must be less or equal to the value of SQL_DESC_COUNT. |

#### Return value

The scale value.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, SQL_NO_DATA, or SQL_INVALID_HANDLE.

SQL_NO_DATA is returned if *RecNumber* is greater than the current number of descriptor records.

# <a name="GetImpRowDescFieldType"></a>GetImpRowDescFieldType

Returns the type of a single field of a descriptor record.

```
FUNCTION GetImpRowDescFieldType (BYVAL RecNumber AS SQLSMALLINT) AS SQLSMALLINT
```

| Parameter  | Description |
| ---------- | ----------- |
| *RecNumber* | Indicates the descriptor record from which the application seeks information. Descriptor records are numbered from 1, with record number 0 being the bookmark record. The *RecNumber* argument must be less or equal to the value of SQL_DESC_COUNT. |

#### Return value

The type of the field.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, SQL_NO_DATA, or SQL_INVALID_HANDLE.

SQL_NO_DATA is returned if *RecNumber* is greater than the current number of descriptor records.

# <a name="GetImpRowDescRec"></a>GetImpRowDescRec

Returns the current settings or values of multiple fields of a descriptor record. The fields returned describe the name, data type, and storage of column or parameter data.

```
FUNCTION GetImpRowDescRec (BYVAL RecNumber AS SQLSMALLINT, BYVAL pwszName AS WSTRING PTR, _
   BYVAL BufferLength AS SQLSMALLINT, BYVAL StringLength AS SQLSMALLINT PTR, _
   BYVAL TypePtr AS SQLSMALLINT PTR, BYVAL SubTypePtr AS SQLSMALLINT PTR, _
   BYVAL LengthPtr AS SQLLEN PTR, BYVAL PrecisionPtr AS SQLSMALLINT PTR, _
   BYVAL ScalePtr AS SQLSMALLINT PTR, BYVAL NullablePtr AS SQLSMALLINT PTR) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *RecNumber* | Indicates the descriptor record from which the application seeks information. Descriptor records are numbered from 1, with record number 0 being the bookmark record. The RecNumber argument must be less or equal to the value of SQL_DESC_COUNT. If RecNumber is less that or equal to SQL_DESC_COUNT but the row does not contain data for a column or parameter, a call to **GetImpRowDesc** will return the default values of the fields. |
| *pwszName* | A pointer to a buffer in which to return the SQL_DESC_NAME field for the descriptor record. |
| *TypePtr* | A pointer to a buffer in which to return the value of the SQL_DESC_TYPE field for the descriptor record. |
| *SubTypePtr* | For records whose type is SQL_DATETIME or SQL_INTERVAL, this is a pointer to a buffer in which to return the value of the SQL_DESC_DATETIME_INTERVAL_CODE field. |
| *LengthPtr* | A pointer to a buffer in which to return the value of the SQL_DESC_OCTET_LENGTH field for the descriptor record. |
| *PrecisionPtr* | A pointer to a buffer in which to return the value of the SQL_DESC_PRECISION field for the descriptor record. |
| *ScaleSptr* | A pointer to a buffer in which to return the value of the SQL_DESC_SCALE field for the descriptor record. |
| *NullablePtr* | A pointer to a buffer in which to return the value of the SQL_DESC_NULLABLE field for the descriptor record. |

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, SQL_NO_DATA, or SQL_INVALID_HANDLE.

SQL_NO_DATA is returned if *RecNumber* is greater than the current number of descriptor records or the the statement is in the prepared or executed state but there was no open cursor associated with it.

# <a name="GetLongVarcharData"></a>GetLongVarcharData

Retrieves long variable char data from the specified column of the result set.

```
FUNCTION GetLongVarcharData (BYVAL ColumnNumber AS SQLSMALLINT) AS CWSTR
FUNCTION GetLongVarcharData (BYREF ColumnName AS WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColumnNumber* / *ColumnName* | Number or name of the column for which to return data. Result set columns are numbered in increasing column order starting at 1. The bookmark column is column number 0; this can be specified only if bookmarks are enabled. |

Return value

The retrieved data.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="GetPrimaryKeys"></a>GetPrimaryKeys

Returns the column names that make up the primary key for a table. This function does not support returning primary keys from multiple tables in a single call. Optional feature not implemented by the Microsoft Access Driver.

```
FUNCTION GetPrimaryKeys (BYREF wszCatalogName AS WSTRING, BYVAL CatalogNameLength AS SQLSMALLINT, _
   BYREF wszSchemaName AS WSTRING, BYVAL SchemanameLength AS SQLSMALLINT, _
   BYREF wszTableName AS WSTRING, BYVAL TableNameLength AS SQLSMALLINT) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszCatalogName* | Catalog name. If a driver supports catalogs for some tables but not for others, such as when the driver retrieves data from different DBMSs, an empty string ("") denotes those tables that do not have catalogs. *wszCatalogName* cannot contain a string search pattern.<br>If the SQL_ATTR_METADATA_ID statement attribute is set to SQL_TRUE, *wszCatalogName* is treated as an identifier and its case is not significant. If it is SQL_FALSE, *wszCatalogName* is an ordinary argument; it is treated literally, and its case is significant. |
| *CatalogNameLength* | Length of *wszCatalogName*. |
| *wszSchemaName* | Schema name. If a driver supports schemas for some tables but not for others, such as when the driver retrieves data from different DBMSs, an empty string ("") denotes those tables that do not have schemas. *wszSchemaName* cannot contain a string search pattern. |
| *SchemaNameLength* | Length of *wszSchemaName*. |
| *wszTableName* | Table name. This argument cannot be a null pointer. *wszTableName* cannot contain a string search pattern. |
| *TableNameLength* | Length of *wszTableName*. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

#### Diagnostics

When **PrimaryKeys** returns SQL_ERROR or SQL_SUCCESS_WITH_INFO, an associated SQLSTATE value can be obtained by calling **GetDiagRec** with a *HandleType* of SQL_HANDLE_STMT and a *Handle* of *hStmt*.

# <a name="GetProcedureColumns"></a>GetProcedureColumns

Returns the list of input and output parameters, as well as the columns that make up the result set for the specified procedures.

```
FUNCTION GetProcedureColumns (BYREF wszCatalogName AS WSTRING, BYVAL CatalogNameLength AS SQLSMALLINT, _
   BYREF wszSchemaName AS WSTRING, BYVAL SchemaNameLength AS SQLSMALLINT, _
   BYREF wszProcName AS WSTRING, BYVAL ProcNameLength AS SQLSMALLINT, _
   BYREF wszColumnName AS WSTRING, BYVAL ColumnNameLength AS SQLSMALLINT) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszCatalogName* | Procedure catalog name. If a driver supports catalogs for some procedures but not for others, such as when the driver retrieves data from different DBMSs, an empty string ("") denotes those procedures that do not have catalogs. *wszCatalogName* cannot contain a string search pattern. |
| *CatalogNameLength* | Length of *wszCatalogName*. |
| *wszSchemaName* | String search pattern for procedure schema names. If a driver supports schemas for some procedures but not for others, such as when the driver retrieves data from different DBMSs, an empty string ("") denotes those procedures that do not have schemas. |
| *SchemaNameLength* | Length of *wszSchemaName*. |
| *wszProcName* | String search pattern for procedure names. |
| *ProcNameLength* | Length of *wszProcName*. |
| *wszColumnName* | String search pattern for column names. |
| *ColumnNameLength* | Length of *wszColumnName*. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

#### Diagnostics

When **ProcedureColumns** returns SQL_ERROR or SQL_SUCCESS_WITH_INFO, an associated SQLSTATE value can be obtained by calling **GetDiagRec** with a *HandleType* of SQL_HANDLE_STMT and a *Handle* of *hStmt*.

# <a name="GetProcedures"></a>GetProcedures

Returns the list of procedure names stored in a specific data source. Procedure is a generic term used to describe an executable object, or a named entity that can be invoked using input and output parameters.

```
FUNCTION GetProcedures (BYREF wszCatalogName AS WSTRING, BYVAL CatalogNameLength AS SQLSMALLINT, _
   BYREF wszSchemaName AS WSTRING, BYVAL SchemaNameLength AS SQLSMALLINT, _
   BYREF wszProcName AS WSTRING, BYVAL ProcNameLength AS SQLSMALLINT) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszCatalogName* | Procedure catalog. If a driver supports catalogs for some tables but not for others, such as when the driver retrieves data from different DBMSs, an empty string ("") denotes those tables that do not have catalogs. *wszCatalogName* cannot contain a string search pattern. |
| *CatalogNameLength* | Length of *wszCatalogName*. |
| *wszSchemaName* | String search pattern for procedure schema names. If a driver supports schemas for some procedures but not for others, such as when the driver retrieves data from different DBMSs, an empty string ("") denotes those procedures that do not have schemas. |
| *SchemaNameLength* | Length of *wszSchemaName*. |
| *wszProcName* | String search pattern for procedure names. |
| *ProcNameLength* | Length of *wszProcName*. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

#### Diagnostics

When **Procedures** returns SQL_ERROR or SQL_SUCCESS_WITH_INFO, an associated SQLSTATE value can be obtained by calling **GetDiagRec** with a *HandleType* of SQL_HANDLE_STMT and a *Handle* of *hStmt*.

# <a name="GetSpecialColumns"></a>GetSpecialColumns

Retrieves the following information about columns within a specified table:

* The optimal set of columns that uniquely identifies a row in the table.
* Columns that are automatically updated when any value in the row is updated by a transaction.

```
FUNCTION GetSpecialColumns (BYVAL identifierType AS SQLUSMALLINT, BYREF wszCatalogName AS WSTRING, _
   BYVAL CatalogNameLength AS SQLUSMALLINT, BYREF wszSchemaName AS WSTRING, _
   BYVAL SchemaNameLength AS SQLUSMALLINT, BYREF wszTableName AS WSTRING, _
   BYVAL TableNameLength AS SQLUSMALLINT, BYVAL fScope AS SQLUSMALLINT,_
   BYVAL fNullable AS SQLUSMALLINT) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *IdentifierType* | Type of column to return. Must be one of the following values:<br>SQL_BEST_ROWID: Returns the optimal column or set of columns that, by retrieving values from the column or columns, allows any row in the specified table to be uniquely identified. A column can be either a pseudo-column specifically designed for this purpose (as in Oracle ROWID or Ingres TID) or the column or columns of any unique index for the table.<br>SQL_ROWVER: Returns the column or columns in the specified table, if any, that are automatically updated by the data source when any value in the row is updated by any transaction (as in SQLBase ROWID or Sybase TIMESTAMP). |
| *wszCatalogName* | Procedure catalog. If a driver supports catalogs for some tables but not for others, such as when the driver retrieves data from different DBMSs, an empty string ("") denotes those tables that do not have catalogs. *wszCatalogName* cannot contain a string search pattern. |
| *CatalogNameLength* | Length of *wszCatalogName*. |
| *wszSchemaName* | String search pattern for procedure schema names. If a driver supports schemas for some procedures but not for others, such as when the driver retrieves data from different DBMSs, an empty string ("") denotes those procedures that do not have schemas. |
| *SchemaNameLength* | Length of *wszSchemaName*. |
| *wszTableName* | Table name. This argument cannot be a null pointer. *wszTableName* cannot contain a string search pattern. |
| *TableNameLength* | Length of *wszTableName*. |
| *fScope* | Minimum required scope of the rowid. The returned rowid may be of greater scope. Must be one of the following:<br>SQL_SCOPE_CURROW: The rowid is guaranteed to be valid only while positioned on that row. A later reselect using rowid may not return a row if the row was updated or deleted by another transaction.<br>SQL_SCOPE_TRANSACTION: The rowid is guaranteed to be valid for the duration of the current transaction.<br>SQL_SCOPE_SESSION: The rowid is guaranteed to be valid for the duration of the session (across transaction boundaries). |
| *fNullable* | Determines whether to return special columns that can have a NULL value. Must be one of the following:<br>SQL_NO_NULLS: Exclude special columns that can have NULL values. Some drivers cannot support SQL_NO_NULLS, and these drivers will return an empty result set if SQL_NO_NULLS was specified. Applications should be prepared for this case and request SQL_NO_NULLS only if it is absolutely required.<br>SQL_NULLABLE: return special columns even if they can have NULL values.  |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

#### Diagnostics

When **SpecialColumns** returns SQL_ERROR or SQL_SUCCESS_WITH_INFO, an associated SQLSTATE value may be obtained by calling **GetDiagRec** with a *HandleType* of SQL_HANDLE_STMT and a *Handle* of *hStmt*.

# <a name="GetSqlState"></a>GetSqlState

Returns the SqlState for the statement handle.

```
FUNCTION GetSqlState () AS CWSTR
```

#### Return value

The SqlState value.

# <a name="GetStatistics"></a>GetStatistics

Retrieves a list of statistics about a single table and the indexes associated with the table. The driver returns this information as a result set on the specified statement handle.

```
FUNCTION GetStatistics (BYREF wszCatalogName AS WSTRING, BYVAL CatalogNameLength AS SQLSMALLINT, _
   BYREF wszSchemaName AS WSTRING, BYVAL SchemaNameLength AS SQLSMALLINT, _
   BYREF wszTableName AS WSTRING, BYVAL TableNameLength AS SQLSMALLINT, _
   BYVAL fUnique AS SQLUSMALLINT, BYVAL fCardinality AS SQLUSMALLINT) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *hStmt* | Statement handle. |
| *wszCatalogName* | Procedure catalog. If a driver supports catalogs for some tables but not for others, such as when the driver retrieves data from different DBMSs, an empty string ("") denotes those tables that do not have catalogs. *wszCatalogName* cannot contain a string search pattern. |
| *CatalogNameLength* | Length of *wszCatalogName*. |
| *wszSchemaName* | String search pattern for procedure schema names. If a driver supports schemas for some procedures but not for others, such as when the driver retrieves data from different DBMSs, an empty string ("") denotes those procedures that do not have schemas. *wszSchemaName* cannot contain a string search pattern. |
| *SchemaNameLength* | Length of *wszSchemaName*. |
| *wszTableName* | Table name. *wszTableName* cannot contain a string search pattern.<br>If the SQL_ATTR_METADATA_ID statement attribute is set to SQL_TRUE, *wszTableName* is treated as an identifier and its case is not significant. If it is SQL_FALSE, *wszTableName* is an ordinary argument; it is treated literally, and its case is significant.  |
| *TableNameLength* | Length of *wszTableName*. |
| *fUnique* | Type of index: SQL_INDEX_UNIQUE or SQL_INDEX_ALL. |
| *fCardinality* | Indicates the importance of the CARDINALITY and PAGES columns in the result set. The following options affect the return of the CARDINALITY and PAGES columns only; index information is returned even if CARDINALITY and PAGES are not returned.<br>SQL_ENSURE requests that the driver unconditionally retrieve the statistics. (Drivers that conform only to the X/Open standard and do not support ODBC extensions will not be able to support SQL_ENSURE.)<br>SQL_QUICK requests that the driver retrieve the CARDINALITY and PAGES only if they are readily available from the server. In this case, the driver does not ensure that the values are current. (Applications that are written to the X/Open standard will always get SQL_QUICK behavior from ODBC 3.x-compliant drivers.)  |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

#### Diagnostics

When **GetStatistics** returns SQL_ERROR or SQL_SUCCESS_WITH_INFO, an associated SQLSTATE value can be obtained by calling **GetDiagRec** with a *HandleType* of SQL_HANDLE_STMT and a *Handle* of *hStmt*.

#### Example

```
#include once "Afx/COdbc/COdbc.inc"
USING Afx

' // Connect with the database
DIM wszConStr AS WSTRING * 260 = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=biblio.mdb"
DIM pDbc AS CODBC = wszConStr

' // Allocate an statement object
DIM pStmt AS COdbcStmt = @pDbc
IF pStmt.Handle = NULL THEN PRINT "Failed to allocate the statement handle": SLEEP : END

DIM cbbytes AS LONG
DIM szCatalogName AS ZSTRING * 256
DIM szSchemaName AS ZSTRING * 256
DIM szTableName AS ZSTRING * 129
DIM iNonUnique AS SHORT
DIM szIndexQualifier AS ZSTRING * 129
DIM szIndexName AS ZSTRING * 129
DIM iInfoType AS SHORT
DIM iOrdinalPosition AS SHORT
DIM szColumnName AS ZSTRING * 129
DIM szAscOrDesc AS ZSTRING * 2
DIM lCardinality AS LONG
DIM lPages AS LONG
DIM szFilterCondition AS ZSTRING * 129

' // Note: Despite calling the SQLStatisticsW function, the GetStatistics
' // method fails if we use unicode variables instead of ansi.
szTableName = "Authors"
pStmt.GetStatistics(szCatalogName, LEN(szCatalogName), szSchemaName, LEN(szSchemaName), _
   szTableName, LEN(szTableName), SQL_INDEX_ALL, SQL_ENSURE)

print pStmt.GetLastResult
pStmt.BindCol( 1, @szCatalogName, SIZEOF(szCatalogName), @cbBytes)
pStmt.BindCol( 2, @szSchemaName, SIZEOF(szSchemaName), @cbbytes)
pStmt.BindCol( 3, @szTableName, SIZEOF(szTableName), @cbbytes)
pStmt.BindCol( 4, @iNonUnique, @cbbytes)
pStmt.BindCol( 5, @szIndexQualifier, SIZEOF(szIndexQualifier), @cbbytes)
pStmt.BindCol( 6, @szIndexName, SIZEOF(szIndexName), @cbbytes)
pStmt.BindCol( 7, @iInfoType, @cbbytes)
pStmt.BindCol( 8, @iOrdinalPosition, @cbbytes)
pStmt.BindCol( 9, @szColumnName, SIZEOF(szColumnName), @cbbytes)
pStmt.BindCol(10, @szAscOrDesc, SIZEOF(szAscOrDesc), @cbbytes)
pStmt.BindCol(11, @lCardinality, @cbbytes)
pStmt.BindCol(12, @lPages, @cbbytes)
pStmt.BindCol(13, @szFilterCondition, SIZEOF(szFilterCondition), @cbbytes)
DO
   IF pStmt.Fetch = FALSE THEN EXIT DO
   ? "Table catalog name: " & szCatalogName
   ? "Table schema name: " & szSchemaName
   ? "Table name: " & szTableName
   ? "Non unique: " & STR(iNonUnique)
   ? "Index qualifier: " & szIndexQualifier
   ? "Index name: " & szIndexName
   ? "Info type: " & STR(iInfoType)
   ? "Ordinal position: " & STR(iOrdinalPosition)
   ? "Column name: " & szColumnName
   ? "Asc or desc: " & szAscOrDesc
   ? "Cardinality: " & STR(lCardinality)
   ? "Pages: " & STR(lPages)
   ? "Filter condition: " & szFilterCondition
   PRINT
   PRINT "Press any key..."
   SLEEP
   CLS
LOOP
```

# <a name="GetStmtAppParamDesc"></a>GetStmtAppParamDesc

Gets the handle to the APD for subsequent calls to **Execute** and **ExecDirect** on the statement handle. The initial value of this attribute is the descriptor implicitly allocated when the statement was initially allocated. If the value of this attribute is set to SQL_NULL_DESC or the handle originally allocated for the descriptor, an explicitly allocated APD handle that was previously associated with the statement handle is dissociated from it and the statement handle reverts to the implicitly allocated APD  handle.

This attribute cannot be set to a descriptor handle that was implicitly  allocated for another statement or to another descriptor handle that was implicitly set on the same statement; implicitly allocated descriptor handles cannot be associated with more than one statement or descriptor handle.

```
FUNCTION GetStmtAppParamDesc () AS SQLUINTEGER
```

#### Return value

The handle of the APD.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="GetStmtAppRowDesc"></a>GetStmtAppRowDesc

Gets the handle to the ARD for subsequent fetches on the statement handle. The initial value of this attribute is the descriptor implicitly allocated when the statement was initially allocated. If the value of this attribute is set to SQL_NULL_DESC or the handle originally allocated for the descriptor, an explicitly allocated ARD handle that was previously associated with the statement handle is dissociated from it and the statement handle reverts to the implicitly allocated ARD handle.

This attribute cannot be set to a descriptor handle that was implicitly allocated for another statement or to another descriptor handle that was implicitly set on the same statement; implicitly allocated descriptor handles cannot be associated with more than one statement or descriptor handle.

```
FUNCTION GetStmtAppRowDesc () AS SQLUINTEGER
```

#### Return value

The handle of the ARD.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="GetStmtAsyncEnable"></a>GetStmtAsyncEnable

Gets an SQLUINTEGER value that specifies whether a function called with the specified statement is executed asynchronously.

**Note**: Optional feature not implemented by the Microsoft Access Driver.

```
FUNCTION GetStmtAsyncEnable () AS SQLUINTEGER
```

#### Remarks

Once a function has been called asynchronously, only the original function, **Cancel**, **GetDiagField**, or **GetDiagRec** can be called on the statement, and only the original function, **AllocStmt**, **GetDiagField**, **GetDiagRec**, or **GetFunctions** can be called on the connection associated with the statement, until the original function returns a code other than SQL_STILL_EXECUTING. Any other function called on the statement or the connection associated with the statement returns SQL_ERROR with an SQLSTATE of HY010 (Function sequence error). Functions can be called on other statements.

For drivers with statement level asynchronous execution support, the statement attribute SQL_ATTR_ASYNC_ENABLE may be set. Its initial value is the same as the value of the connection level attribute with the same name at the time the statement handle was allocated.

For drivers with connection-level, asynchronous-execution support, the statement attribute SQL_ATTR_ASYNC_ENABLE is read-only. Its value is the same as the value of the connection level attribute with the same name at the time the statement handle was allocated. Calling **SetStmtAttr** to set SQL_ATTR_ASYNC_ENABLE when the SQL_ASYNC_MODE **InfoType** returns SQL_AM_CONNECTION returns SQLSTATE HYC00 (Optional feature not implemented).

As a standard practice, applications should execute functions asynchronously only on single-thread operating systems. On multithread operating systems, applications should execute functions on separate threads rather than executing them asynchronously on the same thread. No functionality is lost if drivers that operate only on multithread operating systems do not need to support asynchronous execution. 

The following functions can be executed asynchronously:

BulkOperations ColAttribute ColumnPrivileges Columns CopyDesc DescribeCol DescribeParam ExecDirect Execute Fetch FetchScroll ForeignKeys GetData GetDescField\[1] GetDescRec\[1] GetDiagField GetDiagRec GetTypeInfo MoreResults NumParams NumResultCols ParamData Prepare PrimaryKeys ProcedureColumns Procedures PutData SetPos SpecialColumns Statistics TablePrivileges Tables.

\[1] These functions can be called asynchronously only if the descriptor is an implementation descriptor, not an application descriptor.

Return value

SQL_ASYNC_ENABLE_OFF or SQL_ASYNC_ENABLE_ON.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="GetStmtAttr"></a>GetStmtAttr

Returns the current setting of a statement attribute.

```
FUNCTION GetStmtAttr (BYVAL Attribute AS SQLINTEGER, BYVAL ValuePtr AS SQLPOINTER, _
   BYVAL BufferLength AS SQLINTEGER, BYVAL StringLength AS SQLINTEGER PTR) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *Attribute* | Attribute to retrieve.  |
| *ValuePtr* | Pointer to a buffer in which to return the value of the attribute specified in *Attribute*. |
| *BufferLength* |  If *Attribute* is an ODBC-defined attribute and *ValuePtr* points to a character string or a binary buffer, this argument should be the length of *ValuePtr*. If *Attribute* is an ODBC-defined attribute and *ValuePtr* is an integer, *BufferLength* is ignored. If the value returned in *ValuePtr* is a Unicode string, the *BufferLength* argument must be an even number.<br><br>If *Attribute* is a driver-defined attribute, the application indicates the nature of the attribute to the Driver Manager by setting the *BufferLength* argument. *BufferLength* can have the following values:<ul><li>If *ValuePtr* is a pointer to a character string, then *BufferLength* is the length of the string or SQL_NTS.</li><li>If *ValuePtr* is a pointer to a binary buffer, then the application places the result of the SQL_LEN_BINARY_ATTR(length) macro in *BufferLength*. This places a negative value in *BufferLength*.</li><li>If *ValuePtr* is a pointer to a value other than a character string or binary string, then *BufferLength* should have the value SQL_IS_POINTER.<br>If *ValuePtr* is contains a fixed-length data type, then *BufferLength* is either SQL_IS_INTEGER or SQL_IS_UINTEGER, as appropriate.</li></ul> |
| *StringLengthPtr* | A pointer to a buffer in which to return the total number of bytes (excluding the null-termination character) available to return in *ValuePtr*. If *ValuePtr* is a null pointer, no length is returned. If the attribute value is a character string, and the number of bytes available to return is greater than or equal to *BufferLength*, the data in *ValuePtr* is truncated to *BufferLength* minus the length of a null-termination character and is null-terminated by the driver. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="GetStmtFetchBookmarkPtr"></a>GetStmtFetchBookmarkPtr

Gets a pointer that points to a binary bookmark value. When **FetchScroll** is called with *FetchOrientation* equal to SQL_FETCH_BOOKMARK, the driver picks up the bookmark value from this field. This field defaults to a null pointer.

The value pointed to by this field is not used for delete by bookmark, update by bookmark, or fetch by bookmark operations in **BulkOperations**, which use bookmarks cached in rowset buffers.

```
FUNCTION GetStmtFetchBookmarkPtr () AS SQLUINTEGER
```

#### Return value

A pointer that points to a binary bookmark.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="GetStmtMaxLength"></a>GetStmtMaxLength

Gets an SQLUINTEGER value that specifies the maximum amount of data that the driver returns from a character or binary column. Optional feature not implemented by the Microsoft Access Driver.

```
FUNCTION GetStmtMaxLength () AS SQLUINTEGER
```

#### Remarks

If *dwAttr* is less than the length of the available data, **Fetch** or **GetData** truncates the data and returns SQL_SUCCESS. If *dwAttr* is 0 (the default), the driver attempts to return all available data. If the specified length is less than the minimum amount of data that the data source can return or greater than the maximum amount of data that the data source can return, the driver substitutes that value and returns SQLSTATE 01S02 (Option value changed). The value of this attribute can be set on an open cursor; however, the setting might not take effect immediately, in which case the driver will return SQLSTATE 01S02 (Option value changed) and reset the attribute to its original value. This attribute is intended to reduce network traffic and should be supported only when the data source (as opposed to the driver) in a  multiple-tier driver can implement it. This mechanism should not be used by applications to truncate data; to truncate data received, an application should specify the maximum buffer length in the *BufferLength* argument in **BindCol** or **GetData**.

#### Return value

The maximum amount of data that the driver returns from a character or binary column.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="GetStmtMaxRows"></a>GetStmtMaxRows

Gets an SQLUINTEGER value corresponding to the maximum number of rows to return to the application for a SELECT statement.

```
FUNCTION GetStmtMaxRows () AS SQLUINTEGER
```

#### Remarks

If *dwAttr* equals 0 (the default), the driver returns all rows. This attribute is intended to reduce network traffic. Conceptually, it is applied when the result set is created and limits the result set to the first dwAttr rows. If the number of rows in the result set is greater than *dwAttr*, the result set is truncated. SQL_ATTR_MAX_ROWS applies to all result sets on the statement, including those returned by catalog functions. SQL_ATTR_MAX_ROWS establishes a maximum for the value of the cursor row count.

A driver should not emulate SQL_ATTR_MAX_ROWS behavior for *Fetch* or *FetchScroll* (if result set size limitations cannot be implemented at the data source) if it cannot guarantee that SQL_ATTR_MAX_ROWS will be implemented properly. It is driver-defined whether SQL_ATTR_MAX_ROWS applies to statements other than SELECT statements (such as catalog functions).

The value of this attribute can be set on an open cursor; however, the setting might not take effect immediately, in which case the driver will return SQLSTATE 01S02 (Option value changed) and reset the attribute to its original value.

#### Return value

The maximum number of rows to return to the application for a SELECT statement.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="GetStmtNoScan"></a>GetStmtNoScan

Gets an SQLUINTEGER value that indicates whether the driver should scan SQL strings for escape sequences.

```
FUNCTION GetStmtNoScan () AS SQLUINTEGER
```

#### Return value

SQL_NOSCAN_OFF or SQL_NOSCAN_ON.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="GetStmtParamBindOffsetPtr"></a>GetStmtParamBindOffsetPtr

Gets an SQLUINTEGER value that points to an offset added to pointers to change binding of dynamic parameters.

```
FUNCTION GetStmtParamBindOffsetPtr () AS SQLUINTEGER
```

#### Remarks

If this field is non-null, the driver dereferences the pointer, adds the dereferenced value to each of the deferred fields in the descriptor record (SQL_DESC_DATA_PTR, SQL_DESC_INDICATOR_PTR, and SQL_DESC_OCTET_LENGTH_PTR), and uses the new pointer values when binding. It is set to null by default.

The bind offset is always added directly to the SQL_DESC_DATA_PTR, SQL_DESC_INDICATOR_PTR, and SQL_DESC_OCTET_LENGTH_PTR fields. If the offset is changed to a different value, the new value is still added directly to the value in the descriptor field. The new offset is not added to the field value plus any earlier offsets.

Setting this statement attribute sets the SQL_DESC_BIND_OFFSET_PTR field in the APD header.

#### Return value

The offset pointer.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="GetStmtParamBindType"></a>GetStmtParamBindType

Gets an SQLUINTEGER value that indicates the binding orientation to be used for dynamic parameters.

```
FUNCTION GetStmtParamBindType () AS SQLUINTEGER
```

#### Remarks

This field is set to SQL_PARAM_BIND_BY_COLUMN (the default) to select column-wise binding.

To select row-wise binding, this field is set to the length of the structure or an instance of a buffer that will be bound to a set of dynamic parameters. This length must include space for all of the bound parameters and any padding of the structure or buffer to ensure that when the address of a bound parameter is incremented with the specified length, the result will point to the beginning of the same parameter in the next set of parameters. When using the sizeof operator in ANSI C, this behavior is guaranteed.

Setting this statement attribute sets the SQL_DESC_ BIND_TYPE field in the APD header.

#### Return value

The binding orientation.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="GetStmtParamOperationPtr"></a>GetStmtParamOperationPtr

Gets an SQLUSMALLINT value that points to an array of SQLUSMALLINT values used to ignore a parameter during execution of an SQL statement. Optional feature not implemented by the Microsoft Access Driver.

```
FUNCTION GetStmtParamOperationPtr () AS SQLUINTEGER
```

#### Remarks

Each value is set to either SQL_PARAM_PROCEED (for the parameter to be executed) or SQL_PARAM_IGNORE (for the parameter to be ignored).

A set of parameters can be ignored during processing by setting the status value in the array pointed to by SQL_DESC_ARRAY_STATUS_PTR in the APD to SQL_PARAM_IGNORE. A set of parameters is processed if its status value is set to SQL_PARAM_PROCEED or if no elements in the array are set.

This statement attribute can be set to a null pointer, in which case the driver does not return parameter status values. This attribute can be set at any time, but the new value is not used until the next time **ExecDirect** or **Execute** is called.

Setting this statement attribute sets the SQL_DESC_ARRAY_STATUS_PTR field in the APD header.

#### Return value

The pointer to the array.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="GetStmtParamsetSize"></a>GetStmtParamsetSize

Gets an SQLUINTEGER value that specifies the number of values for each parameter.

```
FUNCTION GetStmtParamsetSize () AS SQLUINTEGER
```

#### Remarks

If SQL_ATTR_PARAMSET_SIZE is greater than 1, SQL_DESC_DATA_PTR, SQL_DESC_INDICATOR_PTR, and SQL_DESC_OCTET_LENGTH_PTR of the APD point to arrays. The cardinality of each array is equal to the value of this field.

Setting this statement attribute sets the SQL_DESC_ARRAY_SIZE field in the APD header.

#### Return value

The number of values for each parameter.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="GetStmtParamsProcessedPtr"></a>GetStmtParamsProcessedPtr

Gets an SQLUINTEGER record field that points to a buffer in which to return the number of sets of parameters that have been processed, including error sets.

```
FUNCTION GetStmtParamsProcessedPtr () AS SQLUINTEGER
```

#### Remarks

Setting this statement attribute sets the SQL_DESC_ROWS_PROCESSED_PTR field in the IPD header.

If the call to **ExecDirect** or **Execute** that fills in the buffer pointed to by this attribute does not return SQL_SUCCESS or SQL_SUCCESS_WITH_INFO, the contents of the buffer are undefined.

#### Return value

The pointer to the buffer.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="GetStmtParamStatusPtr"></a>GetStmtParamStatusPtr

Gets an SQLUSMALLINT value that points to an array of SQLUSMALLINT values containing status information for each row of parameter values after a  call to **Execute** or **ExecDirect**. This field is required only if PARAMSET_SIZE is greater than 1.

```
FUNCTION GetStmtParamStatusPtr () AS SQLUINTEGER
```

#### Remarks

The status values can contain the following values:

**SQL_PARAM_SUCCESS**: The SQL statement was successfully executed for this set of parameters.

**SQL_PARAM_SUCCESS_WITH_INFO**: The SQL statement was successfully executed for this set of parameters; however, warning information is available in the diagnostics data structure.

**SQL_PARAM_ERROR**: There was an error in processing this set of parameters. Additional error information is available in the diagnostics data structure.

**SQL_PARAM_UNUSED**: This parameter set was unused, possibly due to the fact that some previous parameter set caused an error that aborted further processing, or because SQL_PARAM_IGNORE was set for that set of parameters in the array specified by the SQL_ATTR_PARAM_OPERATION_PTR.

**SQL_PARAM_DIAG_UNAVAILABLE**: The driver treats arrays of parameters as a monolithic unit and so does not generate this level of error information.

This statement attribute can be set to a null pointer, in which case the driver does not return parameter status values. This attribute can be set at any time, but the new value is not used until the next time **Execute** or **ExecDirect** is called. Note that setting this attribute can affect the output parameter behavior implemented by the driver.

Setting this statement attribute sets the SQL_DESC_ARRAY_STATUS_PTR field in the IPD header.

Return value

The pointer to the array.

Result code (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="GetStmtQueryTimeout"></a>GetStmtQueryTimeout

Gets an SQLUINTEGER value corresponding to the number of seconds to wait for an SQL statement to execute before returning to the application. If *dwAttr* is equal to 0 (default), there is no timeout.

**Note**: Optional feature not implemented by the Microsoft Access Driver.

```
FUNCTION GetStmtQueryTimeout () AS SQLUINTEGER
```

#### Remarks

If the specified timeout exceeds the maximum timeout in the data source or is smaller than the minimum timeout, **SetStmtAttr** substitutes that value and returns SQLSTATE 01S02 (Option value changed).

Note that the application need not call *CloseCursor* to reuse the statement if a SELECT statement timed out.

The query timeout set in this statement attribute is valid in both synchronous and asynchronous modes.

#### Return value

The number of seconds to wait for an SQL statement to execute before returning to the application. 

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="GetStmtRetrieveData"></a>GetStmtRetrieveData

Gets an SQLUINTEGER value specifying the data retrieval mode.

```
FUNCTION GetStmtRetrieveData () AS SQLUINTEGER
```

#### Remarks

By setting SQL_RETRIEVE_DATA to SQL_RD_OFF, an application can verify that a row exists or retrieve a bookmark for the row without incurring the overhead of retrieving rows.

The value of this attribute can be set on an open cursor; however, the setting might not take effect immediately, in which case the driver will return SQLSTATE 01S02 (Option value changed) and reset the attribute to its original value.

#### Return value

SQL_RD_ON or SQL_RD_OFF.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="GetStmtRowArraySize"></a>GetStmtRowArraySize

Gets an SQLUINTEGER value that specifies the number of rows returned by each call to **Fetch** or **FetchScroll**. It is also the number of rows in a bookmark array used in a bulk bookmark operation in **BulkOperations**. The default value is 1.

```
FUNCTION GetStmtRowArraySize () AS SQLUINTEGER
```

#### Remarks

If the specified rowset size exceeds the maximum rowset size supported by the data source, the driver substitutes that value and returns SQLSTATE 01S02 (Option value changed).

Setting this statement attribute sets the SQL_DESC_ARRAY_SIZE field in the ARD header.

#### Return value

The number of rows returned by each call to **Fetch** or **FetchScroll**.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="GetStmtRowBindOffsetPtr"></a>GetStmtRowBindOffsetPtr

Gets an SQLUINTEGER value that points to an offset added to pointers to change binding of column data. If this field is non-null, the driver dereferences the pointer, adds the dereferenced value to each of the deferred fields in the descriptor record (SQL_DESC_DATA_PTR, SQL_DESC_INDICATOR_PTR, and SQL_DESC_OCTET_LENGTH_PTR), and uses the new pointer values when binding. It is set to null by default.

```
FUNCTION GetStmtRowBindOffsetPtr () AS SQLUINTEGER
```

#### Remarks

Setting this statement attribute sets the SQL_DESC_BIND_OFFSET_PTR field in the ARD header.

#### Return value

The offset pointer.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="GetStmtRowBindType"></a>GetStmtRowBindType

Gets an SQLUINTEGER value that sets the binding orientation to be used when **Fetch** or **FetchScroll** is called on the associated statement. Column-wise binding is selected by setting the value to SQL_BIND_BY_COLUMN. Row-wise binding is selected by setting the value to the length of a structure or an instance of a buffer into which result columns will be bound.

```
FUNCTION GetStmtRowBindType () AS SQLUINTEGER
```

#### Remarks

If a length is specified, it must include space for all of the bound columns and any padding of the structure or buffer to ensure that when the address of a bound column is incremented with the specified length, the result will point to the beginning of the same column in the next row. When using the sizeof operator with structures or unions in ANSI C, this behavior is guaranteed.

Column-wise binding is the default binding orientation for **Fetch** and **FetchScroll**.

Setting this statement attribute sets the SQL_DESC_BIND_TYPE field in the ARD header.

#### Return value

The binding orientation.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="GetStmtRowNumber"></a>GetStmtRowNumber

Returns an SQLUINTEGER value that is the number of the current row in the entire result set. If the number of the current row cannot be determined or there is no current row, the driver returns 0.

```
FUNCTION GetStmtRowNumber () AS SQLUINTEGER
```

#### Remarks

This attribute can be retrieved but not set.

**Note**: Microsoft Access Driver returns 0.

#### Return value

The number of the current row in the entire result set.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="GetStmtRowOperationPtr"></a>GetStmtRowOperationPtr

Gets an SQLUSMALLINT value that points to an array of SQLUSMALLINT values used to ignore a row during a bulk operation using SetPos. Each value is set to either SQL_ROW_PROCEED (for the row to be included in the bulk operation) or SQL_ROW_IGNORE (for the row to be excluded from the bulk operation). (Rows cannot be ignored by using this array during calls to **BulkOperations**.)

**Note**: Optional feature not implemented by the Microsoft Access Driver.

```
FUNCTION GetStmtRowOperationPtr () AS SQLUINTEGER
```

#### Remarks

This statement attribute can be set to a null pointer, in which case the driver does not return row status values. This attribute can be set at any time, but the new value is not used until the next time **SetPos** is called.

Setting this statement attribute sets the SQL_DESC_ARRAY_STATUS_PTR field in the ARD.

#### Return value

The pointer to the array.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="GetStmtRowsFetchedPtr"></a>GetStmtRowsFetchedPtr

Gets an SQLUINTEGER value that points to a buffer in which to return the number of rows fetched after a call to **Fetch** or **FetchScroll**; the number of rows affected by a bulk operation performed by a call to **SetPos** with an *Operation* argument of SQL_REFRESH; or the number of rows affected by a bulk operation performed by **BulkOperations**. This number includes error rows.

**Note**: Optional feature not implemented by the Microsoft Access Driver.

```
FUNCTION GetStmtRowsFetchedPtr () AS SQLUINTEGER
```

#### Remarks

Setting this statement attribute sets the SQL_DESC_ROWS_PROCESSED_PTR field in the IRD header.

If the call to **Fetch** or **FetchScroll** that fills in the buffer pointed to by this attribute does not return SQL_SUCCESS or SQL_SUCCESS_WITH_INFO, the contents of the buffer are undefined.

#### Return value

The pointer to the buffer.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="GetStmtRowStatusPtr"></a>GetStmtRowStatusPtr

Gets an SQLUSMALLINT value that points to an array of SQLUSMALLINT values containing row status values after a call to **Fetch** or **FetchScroll**.

```
FUNCTION GetStmtRowStatusPtr () AS SQLUINTEGER
```

#### Remarks

The array has as many elements as there are rows in the rowset. This statement attribute can be set to a null pointer, in which case the driver does not return row status values. This attribute can be set at any time, but the new value is not used until the next time BulkOperations, Fetch, FetchScroll, or SetPos is called.

Setting this statement attribute sets the SQL_DESC_ARRAY_STATUS_PTR field in the IRD header.

This attribute is mapped by an ODBC 2.x driver to the rgbRowStatus array in a call to **ExtendedFetch**.

#### Return value

The pointer to the array.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="GetStmtSimulateCursor"></a>GetStmtSimulateCursor

Gets an SQLUINTEGER value that specifies whether drivers that simulate positioned update and delete statements guarantee that such statements affect only one single row. Optional feature not implemented by the Microsoft Access Driver.

```
FUNCTION GetStmtSimulateCursor () AS SQLUINTEGER
```

#### Remarks

To simulate positioned update and delete statements, most drivers construct a searched UPDATE or DELETE statement containing a WHERE clause that specifies the value of each column in the current row. Unless these columns make up a unique key, such a statement can affect more than one row. To guarantee that such statements affect only one row, the driver determines the columns in a unique key and adds these columns to the result set. If an application guarantees that the columns in the result set make up a unique key, the driver is not required to do so. This may reduce execution time.

**SQL_SC_NON_UNIQUE** = The driver does not guarantee that simulated positioned update or delete statements will affect only one row; it is the application's responsibility to do so. If a statement affects more than one row, **Execute**, **ExecDirect**, or **SetPos** returns SQLSTATE 01001 (Cursor operation conflict).

**SQL_SC_TRY_UNIQUE** = The driver attempts to guarantee that simulated positioned update or delete statements affect only one row. The driver always executes such statements, even if they might affect more than one row, such as when there is no unique key. If a statement affects more than one row, **Execute**, **ExecDirect**, or **SetPos** returns SQLSTATE 01001 (Cursor operation conflict).

**SQL_SC_UNIQUE** = The driver guarantees that simulated positioned update or delete statements affect only one row. If the driver cannot guarantee this for a given statement, **ExecDirect** or **Prepare** returns an error.

If the data source provides native SQL support for positioned update and delete statements and the driver does not simulate cursors, SQL_SUCCESS is returned when SQL_SC_UNIQUE is requested for SQL_SIMULATE_CURSOR. SQL_SUCCESS_WITH_INFO is returned if SQL_SC_TRY_UNIQUE or SQL_SC_NON_UNIQUE is requested. If the data source provides the SQL_SC_TRY_UNIQUE level of support and the driver does not, %QL_SUCCESS is returned for %SL_SC_TRY_UNIQUE and %QL_SUCCESS_WITH_INFO is returned for %QL_SC_NON_UNIQUE.

If the specified cursor simulation type is not supported by the data source, the driver substitutes a different simulation type and returns SQLSTATE 01S02 (Option value changed). For SQL_SC_UNIQUE, the driver substitutes, in order, SQL_SC_TRY_UNIQUE or SQL_SC_NON_UNIQUE. For SQL_SC_TRY_UNIQUE, the driver substitutes SQL_SC_NON_UNIQUE.

#### Return value

SQL_SC_NON_UNIQUE, SQL_SC_TRY_UNIQUE or SQL_SC_UNIQUE.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="GetStmtUseBookmarks"></a>GetStmtUseBookmarks

Gets an SQLUINTEGER value that specifies whether an application will use bookmarks with a cursor.

```
FUNCTION GetStmtUseBookmarks () AS SQLUINTEGER
```

#### Return value

SQL_UB_OFF, SQL_UB_VARIABLE or SQL_UB_FIXED.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="GetTablePrivileges"></a>GetTablePrivileges

Returns a list of tables and the privileges associated with each table.

```
FUNCTION GetTablePrivileges (BYREF wszCatalogName AS WSTRING, BYVAL CatalogNameLength AS SQLSMALLINT, _
   BYREF wszSchemaName AS WSTRING, BYVAL SchemaNameLength AS SQLSMALLINT, _
   BYREF wszTableName AS WSTRING, BYVAL TableNameLength AS SQLSMALLINT) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszCatalogName* | Table catalog. If a driver supports catalogs for some tables but not for others, such as when the driver retrieves data from different DBMSs, an empty string ("") denotes those tables that do not have catalogs. *wszCatalogName* cannot contain a string search pattern. |
| *CatalogNameLength* | Length of *wszCatalogName*. |
| *wszSchemaName* | String search pattern for schema names. If a driver supports schemas for some tables but not for others, such as when the driver retrieves data from different DBMSs, an empty string ("") denotes those tables that do not have schemas. |
| *SchemaNameLength* | Length of *wszSchemaName*. |
| *wszTableName* | String search pattern for table names. Pass an empty string ("") for all tables. |
| *TableNameLength* | Length of *wszTableName*. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

#### Diagnostics

When **TablePrivileges** returns SQL_ERROR or SQL_SUCCESS_WITH_INFO, an associated SQLSTATE value may be obtained by calling **GetDiagRec** with a *HandleType* of SQL_HANDLE_STMT and a *Handle* of *hStmt*.

# <a name="GetTables"></a>GetTables

Returns the list of table, catalog, or schema names, and table types, stored in a specific data source.

```
FUNCTION GetTables (BYREF wszCatalogName AS WSTRING, BYVAL CatalogNameLength AS SQLSMALLINT, _
   BYREF wszSchemaName AS WSTRING, BYVAL SchemaNameLength AS SQLSMALLINT, _
   BYREF wszTableName AS WSTRING, BYVAL TableNameLength AS SQLSMALLINT, _
   BYREF wszTableType AS WSTRING, BYVAL TableTypeLength AS SQLSMALLINT) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszCatalogName* | Catalog name. The *wszCatalogName* argument accepts search patterns if the SQL_ODBC_VERSION environment attribute is SQL_OV_ODBC3; it does not accept search patterns if SQL_OV_ODBC2 is set. If a driver supports catalogs for some tables but not for others, such as when a driver retrieves data from different DBMSs, an empty string ("") indicates those tables that do not have catalogs. |
| *CatalogNameLength* | Length of *wszCatalogName*. |
| *wszSchemaName* | String search pattern for schema names. If a driver supports schemas for some tables but not for others, such as when the driver retrieves data from different DBMSs, an empty string ("") indicates those tables that do not have schemas. |
| *SchemaNameLength* | Length of *wszSchemaName*. |
| *wszTableName* | String search pattern for table names. |
| *TableNameLength* | Length of *wszTableName*. |
| *wszTableType* | List of table types to match. Table type names: 'TABLE', 'VIEW', 'SYSTEM TABLE', 'GLOBAL TEMPORARY', 'LOCAL TEMPORARY', 'ALIAS', 'SYNONIM', or a data source-specific type name. The meaning od 'ALIAS' and 'SYNONIM' are driver specific. Pass an empty string ("") to retrieve all table types.<br>If *wszTableTypes* is not an empty string, it must contain a list of comma-separated values for the types of interest; each value may be enclosed in single quotation marks (') or unquoted — for example, 'TABLE', 'VIEW' or TABLE, VIEW. An application should always specify the table type in uppercase; the driver should convert the table type to whatever case is needed by the data source. If the data source does not support a specified table type, Tables does not return any results for that type.<br>Note that the SQL_ATTR_METADATA_ID statement attribute has no effect upon the *wszTableTypes* argument. *wszTableTypes* is a value list argument, no matter what the setting of SQL_ATTR_METADATA_ID.  |
| *TableTypeLength* | Length of *wszTableType*. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

#### Diagnostics

When **Tables** returns SQL_ERROR or SQL_SUCCESS_WITH_INFO, an associated SQLSTATE value can be obtained by calling **GetDiagRec** with a *HandleType* of SQL_HANDLE_STMT and a *Handle* of *hStmt*.

# <a name="GetTypeInfo"></a>GetTypeInfo

Returns information about data types supported by the data source. The driver returns the information in the form of an SQL result set. The data types are intended for use in Data Definition Language (DDL) statements.

**Important**: Applications must use the type names returned in the TYPE_NAME column of the **GetTypeInfo** result set in ALTER TABLE and CREATE TABLE statements. **GetTypeInfo** may return more than one row with the same value in the DATA_TYPE column.

```
FUNCTION GetTypeInfo (BYVAL DataType AS SQLSMALLINT) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *DataType* | The SQL data type. SQL_ALL_TYPES specifies that information about all data types should be returned. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

#### Diagnostics

When **GetTypeInfo** returns SQL_ERROR or SQL_SUCCESS_WITH_INFO, an associated SQLSTATE value can be obtained by calling **GetDiagRec** with a *HandleType* of SQL_HANDLE_STMT and a *Handle* of *hStmt*.

#### Example

```
#include once "Afx/COdbc/COdbc.inc"
USING Afx

' // Connect with the database
DIM wszConStr AS WSTRING * 260 = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=biblio.mdb"
DIM pDbc AS CODBC = wszConStr

' // Allocate an statement object
DIM pStmt AS COdbcStmt = @pDbc
IF pStmt.Handle = NULL THEN PRINT "Failed to allocate the statement handle": SLEEP : END

DIM cbbytes AS LONG
DIM wszTypeName AS WSTRING * 129
DIM iDataType AS SHORT
DIM lColumnSize AS LONG
DIM wszIntervalPrefix AS WSTRING * 129
DIM wszIntervalSuffix AS WSTRING * 129
DIM wszCreateParams AS WSTRING * 129
DIM iNullable AS SHORT
DIM iCaseSensitive AS SHORT
DIM iSearchable AS SHORT
DIM iUnsignedAttribute AS SHORT
DIM iFixedPrecScale AS SHORT
DIM iAutoUniqueValue AS SHORT
DIM wszLocalTypeName AS WSTRING * 129
DIM iMinimumScale AS SHORT
DIM iMaximumScale AS SHORT
DIM iSqlDataType AS SHORT
DIM iSqlDatetimeSub AS SHORT
DIM lNumPrecRadix AS LONG
DIM iIntervalPrecision AS SHORT

pStmt.GetTypeInfo(SQL_ALL_TYPES)
pStmt.BindCol( 1, @wszTypeName, SIZEOF(wszTypeName), @cbbytes)
pStmt.BindCol( 2, @iDataType, @cbbytes)
pStmt.BindCol( 3, @lColumnSize, @cbbytes)
pStmt.BindCol( 4, @wszIntervalPrefix, SIZEOF(wszIntervalPrefix), @cbbytes)
pStmt.BindCol( 5, @wszIntervalSuffix, SIZEOF(wszIntervalSuffix), @cbbytes)
pStmt.BindCol( 6, @wszCreateParams, SIZEOF(wszCreateParams), @cbbytes)
pStmt.BindCol( 7, @iNullable, @cbbytes)
pStmt.BindCol( 8, @iCasesensitive, @cbbytes)
pStmt.BindCol( 9, @iSearchable, @cbbytes)
pStmt.BindCol(10, @iUnsignedAttribute, @cbbytes)
pStmt.BindCol(11, @iFixedPrecScale, @cbbytes)
pStmt.BindCol(12, @iAutoUniqueValue, @cbbytes)
pStmt.BindCol(13, @wszLocalTypeName, SIZEOF(wszLocalTypeName), @cbbytes)
pStmt.BindCol(14, @iMinimumScale, @cbbytes)
pStmt.BindCol(15, @iMaximumScale, @cbbytes)
pStmt.BindCol(16, @iSqlDataType, @cbbytes)
pStmt.BindCol(17, @iSqlDateTimeSub, @cbbytes)
pStmt.BindCol(18, @lNumPrecRadix, @cbbytes)
pStmt.BindCol(19, @iIntervalPrecision, @cbbytes)

DO
   IF pStmt.Fetch = FALSE THEN EXIT DO
   ? "Type name: " & wszTypeName
   ? "Data type: " & STR(iDataType)
   ? "Column size: " & STR(lColumnSize)
   ? "Interval prefix: " & wszIntervalPrefix
   ? "Interval suffix: " & wszIntervalSuffix
   ? "Create params: " & wszCreateParams
   ? "Nullable: " & STR(iNullable)
   ? "Case sensitive: " & STR(iCaseSensitive)
   ? "Searchable: " & STR(iSearchable)
   ? "Unsigned attribute: " & STR(iUnsignedAttribute)
   ? "Fixed prec scale: " & STR(iFixedPrecScale)
   ? "Auto unique value: " & STR(iAutoUniqueValue)
   ? "Local type name: " & wszLocalTypeName
   ? "Minimum scale: " & STR(iMinimumScale)
   ? "Maximum scale: " & STR(iMaximumScale)
   ? "SQL data type: " & STR(iSqlDataType)
   ? "SQL Datetime sub: " & STR(iSqlDatetimeSub)
   ? "Num prec radix: " & STR(lNumPrecRadix)
   ? "Interval precision: " & STR(iIntervalPrecision)
   PRINT
   PRINT "Press any key..."
   SLEEP
   CLS
LOOP
```

# <a name="Handle"></a>Handle

Returns the statement handle.

```
FUNCTION Handle () AS SQLHANDLE
```

# <a name="LockRecord"></a>LockRecord

Sets the cursor position in a rowset and locks the record.

```
FUNCTION LockRecord (BYVAL wRow AS SQLSETPOSIROW = 1) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wRow* | Optional. Row number inside the rowset. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_NEED_DATA, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="MoreResults"></a>MoreResults

Determines whether more results are available on a statement containing SELECT, UPDATE, INSERT, or DELETE statements and, if so, initializes processing for those results.

```
FUNCTION MoreResults () AS SQLRETURN
```

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_NO_DATA, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="Move"></a>Move

Moves the cursor forward or backward *FetchOffset* rowsets.

```
FUNCTION Move (BYVAL FetchOffset AS SQLLEN) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *FetchOffset* | Return the rowset FetchOffset from the start of the current rowset. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_NO_DATA, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="MoveFirst"></a>MoveFirst

Fetches the first rowset of data from the result set and returns data for all bound columns. 

```
FUNCTION MoveFirst () AS SQLRETURN
```

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_NO_DATA, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="MoveLast"></a>MoveLast

Fetches the last rowset of data from the result set and returns data for all bound columns. 

```
FUNCTION MoveLast () AS SQLRETURN
```

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_NO_DATA, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="MoveNext"></a>MoveNext

Fetches the next rowset of data from the result set and returns data for all bound columns. 

```
FUNCTION MoveNext () AS SQLRETURN
```

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_NO_DATA, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="MovePrevious"></a>MovePrevious

Fetches the previous rowset of data from the result set and returns data for all bound columns. 

```
FUNCTION MovePrevious () AS SQLRETURN
```

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_NO_DATA, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="NumParams"></a>NumParams

Returns the number of parameters in an SQL statement.

```
FUNCTION NumParams () AS SQLSMALLINT
```

#### Return value

The number of parameters in an SQL statement.

**Result value** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="NumResultCols"></a>NumResultCols

Returns the number of columns in a result set.

```
FUNCTION NumResultCols () AS SQLSMALLINT
```

#### Return value

The number of columns in a result set.

**Result value** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

#### Example

```
#include once "Afx/COdbc/COdbc.inc"
USING Afx

' // Connect with the database
DIM wszConStr AS WSTRING * 260 = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=biblio.mdb"
DIM pDbc AS CODBC = wszConStr

' // Allocate an statement object
DIM pStmt AS COdbcStmt = @pDbc
IF pStmt.Handle = NULL THEN PRINT "Failed to allocate the statement handle": SLEEP : END

' // Cursor type
pStmt.SetMultiuserKeysetCursor

' // Generate a result set
pStmt.ExecDirect ("SELECT * FROM Authors ORDER BY Author")

' // Retrieve the number of columns
DIM numCols AS SQLSMALLINT = pStmt.NumResultCols
PRINT "Number of columns:" & STR(numCols)

' // Retrieve the names of the fields (columns)
FOR idx AS LONG = 1 TO numCols
   PRINT "Field #" & STR(idx) & " name: " & pStmt.ColName(idx)
NEXT
```

# <a name="ParamData"></a>ParamData

**ParamData** is used in conjunction with **PutData** to supply parameter data at statement execution time.

```
FUNCTION ParamData (BYVAL ValuePtrPtr AS SQLPOINTER PTR) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *ValuePtrPtr* | Pointer to a buffer in which to return the address of the *ParameterValue* buffer specified in **BindParameter** (for parameter data) or the address of the *TargetValue* buffer specified in **BindCol** (for column data), as contained in the SQL_DESC_DATA_PTR descriptor record field. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_NEED_DATA, SQL_NO_DATA, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="Prepare"></a>Prepare

Prepares an SQL string for execution.

```
FUNCTION Prepare (BYREF StatementText AS WSTRING) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *StatementText* | SQL text string. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

#### Example

```
#include once "Afx/COdbc/COdbc.inc"
USING Afx

' // Connect with the database
DIM wszConStr AS WSTRING * 260 = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=biblio.mdb"
DIM pDbc AS CODBC = wszConStr

' // Allocate an statement object
DIM pStmt AS COdbcStmt = @pDbc
IF pStmt.Handle = NULL THEN PRINT "Failed to allocate the statement handle": SLEEP : END

' // Cursor type
pStmt.SetMultiuserKeysetCursor

' // Prepare the statement
pStmt.Prepare("UPDATE Authors SET Author=? WHERE Au_ID=?")
' // Bind the columns
DIM wszAuthor AS WSTRING * 256, cbAuthor AS LONG
DIM lAuId AS LONG, cbAuId AS LONG
pStmt.BindParameter(1, SQL_PARAM_INPUT, SQL_C_CHAR, SQL_CHAR, 255, 0, @wszAuthor, SIZEOF(wszAuthor), @cbAuthor)
pStmt.BindParameter(2, SQL_PARAM_INPUT, SQL_C_LONG, SQL_INTEGER, 4, 0, @lAuId, 0, @cbAuID)
' // Fill the parameter value
lAuId = 999 : cbAuId = 4
wszAuthor = "William Shakespeare" : cbAuthor = LEN(wszAuthor)
' // Execute the prepared statement
pStmt.Execute
IF pStmt.Error = FALSE THEN  PRINT "Record updated" ELSE PRINT pStmt.GetErrorInfo

```

# <a name="PutData"></a>PutData

Allows an application to send data for a parameter or column to the driver at statement execution time. This function can be used to send character or binary data values in parts to a column with a character, binary, or data source–specific data type (for example, parameters of the SQL_LONGVARBINARY or SQL_LONGVARCHAR types).

```
FUNCTION PutData (BYVAL DataPtr AS SQLPOINTER, BYVAL StrLen_or_Ind AS SQLLEN) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *DataPtr* | Pointer to a buffer containing the actual data for the parameter or column. The data must be in the C data type specified in the *ValueType* argument of **BindParameter** (for parameter data) or the *TargetType* argument of **BindCol** (for column data). |
| *StrLen_or_Ind* | Length of *DataPtr*. Specifies the amount of data sent in a call to **PutData**. The amount of data can vary with each call for a given parameter or column. *StrLen_or_Ind* is ignored unless it meets one of the following conditions:<br>*StrLen_or_Ind* is SQL_NTS, SQL_NULL_DATA, or SQL_DEFAULT_PARAM.<br>The C data type specified in **BindParameter** or **BindCol** is SQL_C_CHAR or SQL_C_BINARY.<br>The C data type is SQL_C_DEFAULT, and the default C data type for the specified SQL data type is SQL_C_CHAR or SQL_C_BINARY.<br>For all other types of C data, if StrLen_or_Ind is not SQL_NULL_DATA or SQL_DEFAULT_PARAM, the driver assumes that the size of the DataPtr buffer is the size of the C data type specified with *ValueType* or *TargetType* and sends the entire data value. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="RecordCount"></a>RecordCount

Gets the number of records in the result set.

```
FUNCTION RecordCount (BYREF wszSqlStr AS WSTRING) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszSqlStr* | SQL string.<br>Uses an instruction like SqlStr = "SELECT COUNT(\*) FROM Customers".<br>This is a know workaround to get the number of records in a result set by using the same SELECT clause that will be used to open the result set but adding the COUNT function. |

#### Return value

The number of records in the result set.

#### Result code

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_NO_DATA, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="RefreshRecord"></a>RefreshRecord

Sets the cursor position in a rowset and allows to refresh data in the rowset.

```
FUNCTION RefreshRecord (BYVAL wRow AS SQLSETPOSIROW = 1, _
   BYVAL fLock AS SQLUSMALLINT = SQL_LOCK_NO_CHANGE) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wRow* | Optional. Row number inside the rowset. |
| *fLock* | Optional. Lock type. SQL_LOCK_NO_CHANGE, SQL_LOCK_EXCLUSIVE, SQL_LOCK_UNLOCK |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_NEED_DATA, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="ResetParams"></a>ResetParams

Releases all parameter buffers set by **BindParameter** for the given statement handle.

```
FUNCTION ResetParams () AS SQLRETURN
```

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_INVALID_HANDLE, or SQL_ERROR.


# <a name="RowCount"></a>RowCount

Returns the number of rows affected by an UPDATE, INSERT, or DELETE statement; an SQL_ADD, SQL_UPDATE_BY_BOOKMARK, or SQL_DELETE_BY_BOOKMARK operation in **BulkOperations**; or an SQL_UPDATE or SQL_DELETE operation in **SetPos**.

```
FUNCTION RowCount () AS SQLLEN
```

#### Return value

Returns the number of rows affected by an UPDATE, INSERT, or DELETE statement, or -1 if the number of affected rows is not available.

#### Remarks

When **Execute**, **ExecDirect**, **BulkOperations*, **SetPos**, or **MoreResults** is called, the SQL_DIAG_ROW_COUNT field of the diagnostic data structure is set to the row count, and the row count is cached in an implementation-dependent way. **RowCount** returns the cached row count value. The cached row count value is valid until the statement handle is set back to the prepared or allocated state, the statement is reexecuted, or **CloseCursor** is called. Note that if a function has been called since the SQL_DIAG_ROW_COUNT field was set, the value returned by **RowCount** might be different from the value in the SQL_DIAG_ROW_COUNT field because the SQL_DIAG_ROW_COUNT field is reset to 0 by any function call.

For other statements and functions, the driver may define the value returned in **RowCountPtr**. For example, some data sources may be able to return the number of rows returned by a SELECT statement or a catalog function before fetching the rows

**Note**: Many data sources cannot return the number of rows in a result set before fetching them; for maximum interoperability, applications should not rely on this behavior.

#### Result code

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="SetAbsolutePosition"></a>SetAbsolutePosition

Fetches the rowset starting at row *FetchOffset*.

```
FUNCTION SetAbsolutePosition (BYVAL FetchOffset AS SQLLEN) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *FetchOffset* | Return the rowset starting at row *FetchOffset*. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_NO_DATA, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="SetCursorConcurrency"></a>SetCursorConcurrency

Sets a SQLUINTEGER value that specifies the cursor concurrency.

```
FUNCTION SetCursorConcurrency (BYVAL LockType AS SQLUINTEGER) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *LockType* | The type of concurrency.<br>SQL_CONCUR_READ_ONLY = Cursor is read-only. No updates are allowed.<br>SQL_CONCUR_LOCK = Cursor uses the lowest level of locking sufficient to ensure that the row can be updated.<br>SQL_CONCUR_ROWVER = Cursor uses optimistic concurrency control,comparing row versions such as SQLBase ROWID or Sybase TIMESTAMP.<br>SQL_CONCUR_VALUES = Cursor uses optimistic concurrency control, comparing values. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="SetCursorKeysetSize"></a>SetCursorKeysetSize

Sets a SQLUINTEGER value that specifies the number of rows in the keyset-driven cursor.

```
FUNCTION SetCursorKeysetSize (BYVAL KeysetSize AS SQLUINTEGER) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *KeysetSize* | The size of the keyset cursor.<br>If the keyset size is 0 (the default), the cursor is fully keyset-driven.<br>If the keyset size is greater than 0, the cursor is mixed (keyset-driven within the keyset and dynamic outside of the keyset). The default keyset size is 0. **Fetch** or **FetchScroll** returns an error if the keyset size is greater than 0 and less than the rowset size.<br>For keyset-driven and mixed cursors, the application can specify the keyset size. It does this with the SQL_ATTR_KEYSET_SIZE statement attribute. If  the keyset size is set to 0, which is the default, the keyset size is set to the result size and a keyset-driven cursor is used. The keyset size can be changed after the cursor has been opened. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="SetCursorName"></a>SetCursorName

Sets the cursor name associated with a specified statement.

```
FUNCTION SetCursorName (BYREF wszCursorName AS WSTRING) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszCursorName* | Cursor name. For efficient processing, the cursor name should not include any leading or trailing spaces in the cursor name, and if the cursor name includes a delimited identifier, the delimiter should be positioned as the first character in the cursor name. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

#### Example

```
#include once "Afx/COdbc/COdbc.inc"
USING Afx

' // Connect with the database
DIM wszConStr AS WSTRING * 260 = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=biblio.mdb"
DIM pDbc AS CODBC = wszConStr

' // Allocate an statement object
DIM pStmt AS COdbcStmt = @pDbc
IF pStmt.Handle = NULL THEN PRINT "Failed to allocate the statement handle": SLEEP : END

' // Cursor type
pStmt.SetMultiuserKeysetCursor
' // Sets the cursor name
pStmt.SetCursorName "MyCursor"
' // Gets the cursor name
PRINT "Cursor name: " & pStmt.GetCursorName
```

# <a name="SetCursorScrollability"></a>SetCursorScrollability

Sets a SQLUINTEGER value that specifies the scrollability type. Optional feature not implemented in Microsoft Access Driver.

```
FUNCTION SetCursorScrollability (BYVAL CursorScrollability AS SQLUINTEGER) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *CursorScrollability* | The scrollability type.<br>**SQL_NONSCROLLABLE** = Scrollable cursors are not required on the statement handle. If the application calls FetchScroll on this handle, the only value of *FetchOrientation* is SQL_FETCH_NEXT. This is the default.<br>**SQL_SCROLLABLE** = Scrollable cursors are required on the statement handle. When calling **FetchScroll** the application may specify any valid value of *FetchOrientation*, achieving cursor positioning in modes other than the sequential mode. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="SetCursorSensitivity"></a>SetCursorSensitivity

Sets a SQLUINTEGER value that specifies whether cursors on the statement handle made to a result set by another cursor. Setting this attribute affects subsequent calls to **ExecDirect** and **Execute**. An application can read back the value of this attribute to obtain its initial state or its state  as most recently set by the application.

**Note**: Optional feature not implemented in Microsoft Access Driver.

```
FUNCTION SetCursorSensitivity (BYVAL CursorSensitivity AS SQLUINTEGER) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *CursorSensitivity* | The type of cursor sensitivity.<br>**SQL_UNSPECIFIED** = It is unspecified what the cursor type is and whether  cursors on the statement handle make visible the changes made to a result set by another cursor. Cursors on the statement handle may make visible none, some, or all such changes. This is the default.<br>**SQL_INSENSITIVE** = All cursors on the statement handle show the result set without reflecting any changes made to it by any other cursor, which has a concurrency that is read-only.<br>**SQL_SENSITIVE** = All cursors on the statement handle make visible all changes made to a result set by another cursor. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="SetCursorType"></a>SetCursorType

Sets a SQLUINTEGER value that specifies the cursor type.

**Note**: Microsoft Access Driver changes SQL_CURSOR_DYNAMIC to SQL_CURSOR_KEYSET_DRIVEN.

```
FUNCTION SetCursorType (BYVAL CursorType AS DWORD) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *CursorType* | The type of cursor.<br>**SQL_CURSOR_FORWARD_ONLY** = The cursor only scrolls forward.<br>**SQL_CURSOR_STATIC** = The data in the result set is static.<br>**SQL_CURSOR_KEYSET_DRIVEN** = The driver saves and uses only the keys for the number of rows specified in the SQL_ATTR_KEYSET_SIZE statement attribute.<br>**SQL_CURSOR_DYNAMIC** = The driver saves and uses only the keys for the rows in the rowset.<br>The default value is SQL_FORWARD_ONLY. This attribute cannot be specified after the SQL statement has been prepared.<br>The application can specify the cursor type before executing a statement that creates a result set. It does this with the SQL_ATTR_CURSOR_TYPE statement attribute. If the application does not explicitly specify a type, a forward-only cursor will be used. To get a mixed cursor, an application specifies a keyset-driven cursor but declares a keyset size less than the result set size. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="SetDynamicCursor"></a>SetDynamicCursor

Specifies a dynamic cursor. Note: Microsoft Access Driver changes it to a Keyset cursor.

```
FUNCTION SetDynamicCursor () AS SQLRETURN
```

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_NO_DATA, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="SetForwardOnlyCursor"></a>SetForwardOnlyCursor

Specifies a forward only cursor.

```
FUNCTION SetForwardOnlyCursor () AS SQLRETURN
```

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_NO_DATA, SQL_ERROR, or SQL_INVALID_HANDLE.


# <a name="SetKeysetDrivenCursor"></a>SetKeysetDrivenCursor

Specifies a keyset driven cursor.

```
FUNCTION SetKeysetDrivenCursor () AS SQLRETURN
```

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_NO_DATA, SQL_ERROR, or SQL_INVALID_HANDLE.


# <a name="SetLockConcurrency"></a>SetLockConcurrency

Cursor uses the lowest level of locking sufficient to ensure that the row can be updated.

```
FUNCTION SetLockConcurrency () AS SQLRETURN
```

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="SetMultiuserKeysetCursor"></a>SetMultiuserKeysetCursor

Creates a multiuser keyset cursor.

```
FUNCTION SetMultiuserKeysetCursor (BYVAL pwszCursorName AS WSTRING PTR = NULL) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *pwszCursorName* | Name of the cursor. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="SetOptimisticConcurrency"></a>SetOptimisticConcurrency

Cursor uses optimistic concurrency control, comparing values.

```
FUNCTION SetOptimisticConcurrency () AS SQLRETURN
```

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="SetPos"></a>SetPos

Sets the cursor position in a rowset and allows an application to refresh data in the rowset or to update or delete data in the result set.

```
FUNCTION SetPos (BYVAL wRow AS SQLSETPOSIROW, BYVAL fOption AS SQLUSMALLINT, _
   BYVAL fLock AS SQLUSMALLINT) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wRow* | Position of the row in the rowset on which to perform the operation specified with the fOption argument. If wRow is 0, the operation applies to every row in the rowset. |
| *fOption* | Operation to perform: SQL_POSITION, SQL_REFRESH, SQL_UPDATE, SQL_DELETE. |
| *fLock* | Specifies how to lock the row after performing the operation specified in the Operation argument. SQL_LOCK_NO_CHANGE, SQL_LOCK_EXCLUSIVE, SQL_LOCK_UNLOCK. |

#### Diagnostics

When **SetPos** returns SQL_ERROR or SQL_SUCCESS_WITH_INFO, an associated SQLSTATE value may be obtained by calling **GetDiagRec** with a *HandleType* of SQL_HANDLE_STMT and a *Handle* of *hStmt*.

# <a name="SetPosition"></a>SetPosition

Sets the cursor position in a rowset. Uses SQL_POSITION as the operation type and SQL_LOCK_NO_CHANGE as the lock type.

```
FUNCTION SetPosition (BYVAL wRow AS SQLSETPOSIROW) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wRow* | Position of the row in the rowset. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_NEED_DATA, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="SetReadOnlyConcurrency"></a>SetReadOnlyConcurrency

Cursor is read-only. No updates are allowed.

```
FUNCTION SetReadOnlyConcurrency () AS SQLRETURN
```

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="SetRelativePosition"></a>SetRelativePosition

Fetches the rowset *FetchOffset* from the start of the current rowset.

```
FUNCTION SetRelativePosition (BYVAL FetchOffset AS SQLLEN) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *FetchOffset* | The rowset from from the start of the current rowset. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_NO_DATA, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="SetRowVerConcurrency"></a>SetRowVerConcurrency

Cursor uses optimistic concurrency control, comparing row versions such as SQLBase ROWID or Sybase TIMESTAMP.

```
FUNCTION SetRowVerConcurrency () AS SQLRETURN
```

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="SetStaticCursor"></a>SetStaticCursor

Specifies an static cursor.

```
FUNCTION SetStaticCursor () AS SQLRETURN
```

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_NO_DATA, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="SetStmtAppParamDesc"></a>SetStmtAppParamDesc

Sets the handle to the APD for subsequent calls to **Execute** and **ExecDirect** on the statement handle. The initial value of this attribute is the descriptor implicitly allocated when the statement was initially allocated. If the value of this attribute is set to SQL_NULL_DESC or the handle originally allocated for the descriptor, an explicitly allocated APD handle that was previously associated with the statement handle is dissociated from it and the statement handle reverts to the implicitly allocated APD  handle.

This attribute cannot be set to a descriptor handle that was implicitly  allocated for another statement or to another descriptor handle that was implicitly set on the same statement; implicitly allocated descriptor handles cannot be associated with more than one statement or descriptor handle.

```
FUNCTION SetStmtAppParamDesc (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *dwAttr* | Attribute value. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="SetStmtAppRowDesc"></a>SetStmtAppRowDesc

Sets the handle to the ARD for subsequent fetches on the statement handle. The initial value of this attribute is the descriptor implicitly allocated when the statement was initially allocated. If the value of this attribute is set to SQL_NULL_DESC or the handle originally allocated for the descriptor, an explicitly allocated ARD handle that was previously associated with the statement handle is dissociated from it and the statement handle reverts to the implicitly allocated ARD handle.

This attribute cannot be set to a descriptor handle that was implicitly allocated for another statement or to another descriptor handle that was implicitly set on the same statement; implicitly allocated descriptor handles cannot be associated with more than one statement or descriptor handle.

```
FUNCTION SetStmtAppRowDesc (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *dwAttr* | Attribute value. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.


# <a name="SetStmtAttr"></a>SetStmtAttr

Sets the current setting of a statement attribute.

**Note**: Additional procedures, one foe each attribute, have been developed to be used instead of this generic procedure.

```
FUNCTION SetStmtAttr (BYVAL Attribute AS SQLINTEGER, BYVAL ValuePtr AS SQLPOINTER, _
   BYVAL StringLength AS SQLINTEGER) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *Attribute* | Attribute to set. |
| *ValuePtr* | Pointer to the value to be associated with *Attribute*. Depending on the value of *Attribute*, *ValuePtr* will be a 32-bit unsigned integer value or a pointer to a null-terminated character string, a binary buffer, or a driver-defined value. If the *Attribute* argument is a driver-specific value, *ValuePtr* may be a signed integer. |
| *StringLength* |  If *Attribute* is an ODBC-defined attribute and *ValuePtr* points to a character string or a binary buffer, this argument should be the length of *ValuePtr*. If *Attribute* is an ODBC-defined attribute and *ValuePtr* is an integer, *StringLength* is ignored.<br><br>If *Attribute* is a driver-defined attribute, the application indicates the nature of the attribute to the Driver Manager by setting the StringLength argument. StringLength can have the following values:<br><br><ul><li>If *ValuePtr* is a pointer to a character string, then *StringLength* is the length of the string or SQL_NTS.</li><li>If *ValuePtr* is a pointer to a binary buffer, then the application places the result of the SQL_LEN_BINARY_ATTR(length) macro in *StringLength*. This places a negative value in StringLength.</li><li>If *ValuePtr* is a pointer to a value other than a character string or a binary string, then *StringLength* should have the value SQL_IS_POINTER.</li><li>If *ValuePtr* contains a fixed-length value, then *StringLength* is either SQL_IS_INTEGER or SQL_IS_UINTEGER, as appropriate.</li></ul> |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="SetStmtFetchBookmarkPtr"></a>SetStmtFetchBookmarkPtr

Sets a pointer that points to a binary bookmark value. When **FetchScroll** is called with *FetchOrientation* equal to SQL_FETCH_BOOKMARK, the driver picks up the bookmark value from this field. This field defaults to a null pointer.

The value pointed to by this field is not used for delete by bookmark, update by bookmark, or fetch by bookmark operations in **BulkOperations**, which use bookmarks cached in rowset buffers.

```
FUNCTION SetStmtFetchBookmarkPtr (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *dwAttr* | Pointer to a binary bookmark value. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="SetStmtMaxLength"></a>SetStmtMaxLength

Sets an SQLUINTEGER value that specifies the maximum amount of data that the driver returns from a character or binary column. Note Optional feature not implemented by the Microsoft Access Driver.

```
FUNCTION SetStmtMaxLength (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *dwAttr* | Maximum amount of data that the driver returns from a character or binary column. |

#### Remarks

If *dwAttr* is less than the length of the available data, **Fetch** or **GetData** truncates the data and returns SQL_SUCCESS. If dwAttr is 0 (the default), the driver attempts to return all available data. If the specified length is less than the minimum amount of data that the data source can return or greater than the maximum amount of data that the data source can return, the driver substitutes that value and returns SQLSTATE 01S02 (Option value changed). The value of this attribute can be set on an open cursor; however, the setting might not take effect immediately, in which case the driver will return SQLSTATE 01S02 (Option value changed) and reset the attribute to its original value. This attribute is intended to reduce network traffic and should be supported only when the data source (as opposed to the driver) in a  multiple-tier driver can implement it. This mechanism should not be used by applications to truncate data; to truncate data received, an application should specify the maximum buffer length in the BufferLength argument in **BindCol** or **GetDat**.

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="SetStmtMaxRows"></a>SetStmtMaxRows

Sets an SQLUINTEGER value corresponding to the maximum number of rows to return to the application for a SELECT statement.

```
FUNCTION SetStmtMaxRows (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *dwAttr* | Maximum number of rows to return to the application for a SELECT statement. |

#### Remarks

If *dwAttr* equals 0 (the default), the driver returns all rows. This attribute is intended to reduce network traffic. Conceptually, it is applied when the result set is created and limits the result set to the first *dwAttr* rows. If the number of rows in the result set is greater than *dwAttr*, the result set is truncated. SQL_ATTR_MAX_ROWS applies to all result sets on the statement, including those returned by catalog functions. SQL_ATTR_MAX_ROWS establishes a maximum for the value of the cursor row count.

A driver should not emulate SQL_ATTR_MAX_ROWS behavior for **Fetch** or **FetchScroll** (if result set size limitations cannot be implemented at the data source) if it cannot guarantee that SQL_ATTR_MAX_ROWS will be implemented properly. It is driver-defined whether SQL_ATTR_MAX_ROWS applies to statements other than SELECT statements (such as catalog functions).

The value of this attribute can be set on an open cursor; however, the setting might not take effect immediately, in which case the driver will return SQLSTATE 01S02 (Option value changed) and reset the attribute to its original value.

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="SetStmtNoScan"></a>SetStmtNoScan

Sets an SQLUINTEGER value that indicates whether the driver should scan SQL strings for escape sequences.

```
FUNCTION SetStmtNoScan (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *dwAttr* | The scan value.<br>**SQL_NOSCAN_OFF** = The driver scans SQL strings for escape sequences (the default).<br>**SQL_NOSCAN_ON** = The driver does not scan SQL strings for escape sequences. Instead, the driver sends the statement directly to the data source. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="SetStmtParamBindOffsetPtr"></a>SetStmtParamBindOffsetPtr

Sets an SQLUINTEGER value that points to an offset added to pointers to change binding of dynamic parameters.

```
FUNCTION SetStmtParamBindOffsetPtr (BYVAL dwAttr AS DWORD) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *dwAttr* | Offset value that points to an offset added to pointers to change binding of dynamic parameters. |

#### Remarks

If this field is non-null, the driver dereferences the pointer, adds the dereferenced value to each of the deferred fields in the descriptor record (SQL_DESC_DATA_PTR, SQL_DESC_INDICATOR_PTR, and SQL_DESC_OCTET_LENGTH_PTR), and uses the new pointer values when binding. It is set to null by default.

The bind offset is always added directly to the SQL_DESC_DATA_PTR, SQL_DESC_INDICATOR_PTR, and SQL_DESC_OCTET_LENGTH_PTR fields. If the offset is changed to a different value, the new value is still added directly to the value in the descriptor field. The new offset is not added to the field value plus any earlier offsets.

Setting this statement attribute sets the SQL_DESC_BIND_OFFSET_PTR field in the APD header.

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="SetStmtParamBindType"></a>SetStmtParamBindType

Sets an SQLUINTEGER value that indicates the binding orientation to be used for dynamic parameters.

```
FUNCTION SetStmtParamBindType (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *dwAttr* | The binding operation. |

#### Remarks

This field is set to SQL_PARAM_BIND_BY_COLUMN (the default) to select column-wise binding.

To select row-wise binding, this field is set to the length of the structure or an instance of a buffer that will be bound to a set of dynamic parameters. This length must include space for all of the bound parameters and any padding of the structure or buffer to ensure that when the address of a bound parameter is incremented with the specified length, the result will point to the beginning of the same parameter in the next set of parameters. When using the sizeof operator in ANSI C, this behavior is guaranteed.

Setting this statement attribute sets the SQL_DESC_ BIND_TYPE field in the APD header.

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="SetStmtParamOperationPtr"></a>SetStmtParamOperationPtr

Sets an SQLUSMALLINT value that points to an array of SQLUSMALLINT values used to ignore a parameter during execution of an SQL statement. Optional feature not implemented by the Microsoft Access Driver.

```
FUNCTION SetStmtParamOperationPtr (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *dwAttr* | Value of the attribute. |

#### Remarks

Each value is set to either SQL_PARAM_PROCEED (for the parameter to be executed) or SQL_PARAM_IGNORE (for the parameter to be ignored).

A set of parameters can be ignored during processing by setting the status value in the array pointed to by SQL_DESC_ARRAY_STATUS_PTR in the APD to SQL_PARAM_IGNORE. A set of parameters is processed if its status value is set to SQL_PARAM_PROCEED or if no elements in the array are set.

This statement attribute can be set to a null pointer, in which case the driver does not return parameter status values. This attribute can be set at any time, but the new value is not used until the next time **ExecDirect** or **Execute** is called.

Setting this statement attribute sets the SQL_DESC_ARRAY_STATUS_PTR field in the APD header.

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="SetStmtParamsetSize"></a>SetStmtParamsetSize

Sets an SQLUINTEGER value that specifies the number of values for each parameter.

```
FUNCTION SetStmtParamsetSize (dwAttr AS DWORD) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *dwAttr* | Value of the attribute. |

#### Remarks

If SQL_ATTR_PARAMSET_SIZE is greater than 1, SQL_DESC_DATA_PTR, SQL_DESC_INDICATOR_PTR, and SQL_DESC_OCTET_LENGTH_PTR of the APD point to arrays. The cardinality of each array is equal to the value of this field.

Setting this statement attribute sets the SQL_DESC_ARRAY_SIZE field in the APD header.

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="SetStmtParamsProcessedPtr"></a>SetStmtParamsProcessedPtr

Sets an SQLUINTEGER record field that points to a buffer in which to return the number of sets of parameters that have been processed, including error sets.

```
FUNCTION SetStmtParamsProcessedPtr (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *dwAttr* | Value of the attribute. |

#### Remarks

Setting this statement attribute sets the SQL_DESC_ROWS_PROCESSED_PTR field in the IPD header.

If the call to **ExecDirect** or **Execute** that fills in the buffer pointed to by this attribute does not return SQL_SUCCESS or SQL_SUCCESS_WITH_INFO, the contents of the buffer are undefined.

#### Return value

The pointer to the buffer.

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="SetStmtParamStatusPtr"></a>SetStmtParamStatusPtr

Sets an SQLUSMALLINT value that points to an array of SQLUSMALLINT values containing status information for each row of parameter values after a  call to **Execute** or **ExecDirect**. This field is required only if PARAMSET_SIZE is greater than 1.

```
FUNCTION SetStmtParamStatusPtr (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *dwAttr* | Value of the attribute. |

#### Remarks

The status values can contain the following values:

**SQL_PARAM_SUCCESS**: The SQL statement was successfully executed for this set of parameters.

**SQL_PARAM_SUCCESS_WITH_INFO**: The SQL statement was successfully executed for this set of parameters; however, warning information is available in the diagnostics data structure.

**SQL_PARAM_ERROR**: There was an error in processing this set of parameters. Additional error information is available in the diagnostics data structure.

**SQL_PARAM_UNUSED**: This parameter set was unused, possibly due to the fact that some previous parameter set caused an error that aborted further processing, or because SQL_PARAM_IGNORE was set for that set of parameters in the array specified by the SQL_ATTR_PARAM_OPERATION_PTR.

**SQL_PARAM_DIAG_UNAVAILABLE**: The driver treats arrays of parameters as a monolithic unit and so does not generate this level of error information.

This statement attribute can be set to a null pointer, in which case the driver does not return parameter status values. This attribute can be set at any time, but the new value is not used until the next time **Execute** or **ExecDirect** is called. Note that setting this attribute can affect the output parameter behavior implemented by the driver.

Setting this statement attribute sets the SQL_DESC_ARRAY_STATUS_PTR field in the IPD header.

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="SetStmtQueryTimeout"></a>SetStmtQueryTimeout

Sets an SQLUINTEGER value corresponding to the number of seconds to wait for an SQL statement to execute before returning to the application. If dwAttr is equal to 0 (default), there is no timeout.

**Note**: Optional feature not implemented by the Microsoft Access Driver.

```
FUNCTION SetStmtQueryTimeout (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *dwAttr* | Value of the attribute. |

#### Remarks

If the specified timeout exceeds the maximum timeout in the data source or is smaller than the minimum timeout, **SetStmtAttr** substitutes that value and returns SQLSTATE 01S02 (Option value changed).

Note that the application need not call **CloseCursor** to reuse the statement if a SELECT statement timed out.

The query timeout set in this statement attribute is valid in both synchronous and asynchronous modes.

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="SetStmtRetrieveData"></a>SetStmtRetrieveData

Sets an SQLUINTEGER value specifying the data retrieval mode.

```
FUNCTION SetStmtRetrieveData (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *dwAttr* | Value of the attribute.<br>**SQL_RD_ON** = **FetchScroll** and, in ODBC 3.x, **Fetch** retrieve data after it positions the cursor to the specified location. This is the default.<br>**SQL_RD_OFF** = **FetchScroll** and, in ODBC 3.x, **Fetch** do not retrieve data after it positions the cursor. |

#### Remarks

By setting SQL_RETRIEVE_DATA to SQL_RD_OFF, an application can verify that a row exists or retrieve a bookmark for the row without incurring the overhead of retrieving rows.

The value of this attribute can be set on an open cursor; however, the setting might not take effect immediately, in which case the driver will return SQLSTATE 01S02 (Option value changed) and reset the attribute to its original value.

#### Return value

SQL_RD_ON or SQL_RD_OFF.

**Result value** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="SetStmtRowArraySize"></a>SetStmtRowArraySize

Sets an SQLUINTEGER value that specifies the number of rows returned by each call to **Fetch** or **FetchScroll**. It is also the number of rows in a bookmark array used in a bulk bookmark operation in **BulkOperations**. The default value is 1.

```
FUNCTION SetStmtRowArraySize (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *dwAttr* | The number of rows. |

#### Remarks

If the specified rowset size exceeds the maximum rowset size supported by the data source, the driver substitutes that value and returns SQLSTATE 01S02 (Option value changed).

Setting this statement attribute sets the SQL_DESC_ARRAY_SIZE field in the ARD header.

#### Return value

The number of rows returned by each call to Fetch or FetchScroll.

**Result value** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="SetStmtRowBindOffsetPtr"></a>SetStmtRowBindOffsetPtr

Sets an SQLUINTEGER value that points to an offset added to pointers to change binding of column data. If this field is non-null, the driver dereferences the pointer, adds the dereferenced value to each of the deferred fields in the descriptor record (SQL_DESC_DATA_PTR, SQL_DESC_INDICATOR_PTR, and SQL_DESC_OCTET_LENGTH_PTR), and uses the new pointer values when binding. It is set to null by default.

```
FUNCTION SetStmtRowBindOffsetPtr (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *dwAttr* | Value of the attribute. |

#### Remarks

Setting this statement attribute sets the SQL_DESC_BIND_OFFSET_PTR field in the ARD header.

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="SetStmtRowBindType"></a>SetStmtRowBindType

Sets an SQLUINTEGER value that sets the binding orientation to be used when **Fetch** or **FetchScroll** is called on the associated statement. Column-wise binding is selected by setting the value to SQL_BIND_BY_COLUMN. Row-wise binding is selected by setting the value to the length of a structure or an instance of a buffer into which result columns will be bound.

```
FUNCTION SetStmtRowBindType (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *dwAttr* | Value of the attribute. |

#### Remarks

If a length is specified, it must include space for all of the bound columns and any padding of the structure or buffer to ensure that when the address of a bound column is incremented with the specified length, the result will point to the beginning of the same column in the next row. When using the sizeof operator with structures or unions in ANSI C, this behavior is guaranteed.

Column-wise binding is the default binding orientation for **Fetch** and **FetchScroll**.

Setting this statement attribute sets the SQL_DESC_BIND_TYPE field in the ARD header.

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="SetStmtRowOperationPtr"></a>SetStmtRowOperationPtr

Sets an SQLUSMALLINT value that points to an array of SQLUSMALLINT values used to ignore a row during a bulk operation using SetPos. Each value is set to either SQL_ROW_PROCEED (for the row to be included in the bulk operation) or SQL_ROW_IGNORE (for the row to be excluded from the bulk operation). (Rows cannot be ignored by using this array during calls to **BulkOperations**.)

**Note**: Optional feature not implemented by the Microsoft Access Driver.

```
FUNCTION SetStmtRowOperationPtr (BYVAL dwAttr AS DWORD) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *dwAttr* | Value of the attribute. |

#### Remarks

This statement attribute can be set to a null pointer, in which case the driver does not return row status values. This attribute can be set at any time, but the new value is not used until the next time **SetPos** is called.

Setting this statement attribute sets the SQL_DESC_ARRAY_STATUS_PTR field in the ARD.

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="SetStmtRowsFetchedPtr"></a>SetStmtRowsFetchedPtr

Sets an SQLUINTEGER value that points to a buffer in which to return the number of rows fetched after a call to **Fetch** or **FetchScroll**; the number of rows affected by a bulk operation performed by a call to **SetPos** with an *Operation* argument of SQL_REFRESH; or the number of rows affected by a bulk operation performed by **BulkOperations**. This number includes error rows.

**Note**: Optional feature not implemented by the Microsoft Access Driver.

```
FUNCTION SetStmtRowsFetchedPtr (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *dwAttr* | Value of the attribute. |

#### Remarks

Setting this statement attribute sets the SQL_DESC_ROWS_PROCESSED_PTR field in the IRD header.

If the call to **Fetch** or **FetchScroll** that fills in the buffer pointed to by this attribute does not return SQL_SUCCESS or SQL_SUCCESS_WITH_INFO, the contents of the buffer are undefined.

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="SetStmtRowStatusPtr"></a>SetStmtRowStatusPtr

Sets an SQLUSMALLINT value that points to an array of SQLUSMALLINT values containing row status values after a call to Fetch or **FetchScroll**.

```
FUNCTION SetStmtRowStatusPtr (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *dwAttr* | Value of the attribute. |

#### Remarks

The array has as many elements as there are rows in the rowset. This statement attribute can be set to a null pointer, in which case the driver does not return row status values. This attribute can be set at any time, but the new value is not used until the next time **BulkOperations**, **Fetch**, **FetchScroll**, or **SetPos** is called.

Setting this statement attribute sets the SQL_DESC_ARRAY_STATUS_PTR field in the IRD header.

This attribute is mapped by an ODBC 2.x driver to the rgbRowStatus array in a call to **ExtendedFetch**.

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="SetStmtSimulateCursor"></a>SetStmtSimulateCursor

Sets an SQLUINTEGER value that specifies whether drivers that simulate positioned update and delete statements guarantee that such statements affect only one single row. Optional feature not implemented by the Microsoft Access Driver.

```
FUNCTION SetStmtSimulateCursor (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *dwAttr* | Value of the attribute. |

#### Remarks

To simulate positioned update and delete statements, most drivers construct a searched UPDATE or DELETE statement containing a WHERE clause that specifies the value of each column in the current row. Unless these columns make up a unique key, such a statement can affect more than one row. To guarantee that such statements affect only one row, the driver determines the columns in a unique key and adds these columns to the result set. If an application guarantees that the columns in the result set make up a unique key, the driver is not required to do so. This may reduce execution time.

**SQL_SC_NON_UNIQUE** = The driver does not guarantee that simulated positioned update or delete statements will affect only one row; it is the application's responsibility to do so. If a statement affects more than one row, **Execute**, **ExecDirect**, or **SetPos** returns SQLSTATE 01001 (Cursor operation conflict).

**SQL_SC_TRY_UNIQUE** = The driver attempts to guarantee that simulated positioned update or delete statements affect only one row. The driver always executes such statements, even if they might affect more than one row, such as when there is no unique key. If a statement affects more than one row, **Execute**, **ExecDirect**, or **SetPos** returns SQLSTATE 01001 (Cursor operation conflict).

**SQL_SC_UNIQUE** = The driver guarantees that simulated positioned update or delete statements affect only one row. If the driver cannot guarantee this for a given statement, **ExecDirect** or **Prepare** returns an error.

If the data source provides native SQL support for positioned update and delete statements and the driver does not simulate cursors, SQL_SUCCESS is returned when SQL_SC_UNIQUE is requested for SQL_SIMULATE_CURSOR. SQL_SUCCESS_WITH_INFO is returned if SQL_SC_TRY_UNIQUE or SQL_SC_NON_UNIQUE is requested. If the data source provides the SQL_SC_TRY_UNIQUE level of support and the driver does not, %QL_SUCCESS is returned for %SL_SC_TRY_UNIQUE and %QL_SUCCESS_WITH_INFO is returned for %QL_SC_NON_UNIQUE.

If the specified cursor simulation type is not supported by the data source, the driver substitutes a different simulation type and returns SQLSTATE 01S02 (Option value changed). For SQL_SC_UNIQUE, the driver substitutes, in order, SQL_SC_TRY_UNIQUE or SQL_SC_NON_UNIQUE. For SQL_SC_TRY_UNIQUE, the driver substitutes SQL_SC_NON_UNIQUE.

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="SetStmtUseBookmarks"></a>SetStmtUseBookmarks

Sets an SQLUINTEGER value that specifies whether an application will use bookmarks with a cursor.

```
FUNCTION SetStmtUseBookmarks (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *dwAttr* | Value of the attribute.<br>**SQL_UB_OFF** = Off (the default)<br>**SQL_UB_VARIABLE** = An application will use bookmarks with a cursor, and the driver will provide variable-length bookmarks if they are supported.<br>**SQL_UB_FIXED** is deprecated in ODBC 3.x. ODBC 3.x applications should always use variable-length bookmarks, even when working with ODBC 2.x drivers (which supported only 4-byte, fixed-length bookmarks). This is because a fixed-length bookmark is just a special case of a variable-length bookmark. When working with an ODBC 2.x driver, the Driver Manager maps SQL_UB_VARIABLE to SQL_UB_FIXED.<br>To use bookmarks with a cursor, the application must specify this attribute with the SQL_UB_VARIABLE value before opening the cursor. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="SetStmtAsyncEnable"></a>SetStmtAsyncEnable

Sets an SQLUINTEGER value that specifies whether a function called with the specified statement is executed asynchronously.

**Note**: Optional feature not implemented by the Microsoft Access Driver.

```
FUNCTION SetStmtAsyncEnable (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *dwAttr* | Value of the attribute.<br>**SQL_ASYNC_ENABLE_OFF** = Off (the default)<br>**SQL_ASYNC_ENABLE_ON** = On |

#### Remarks

Once a function has been called asynchronously, only the original function, **Cancel**, **GetDiagField**, or **GetDiagRec** can be called on the statement, and only the original function, **AllocStmt**, **GetDiagField**, **GetDiagRec**, or **GetFunctions** can be called on the connection associated with the statement, until the original function returns a code other than SQL_STILL_EXECUTING. Any other function called on the statement or the connection associated with the statement returns SQL_ERROR with an SQLSTATE of HY010 (Function sequence error). Functions can be called on other statements.

For drivers with statement level asynchronous execution support, the statement attribute SQL_ATTR_ASYNC_ENABLE may be set. Its initial value is the same as the value of the connection level attribute with the same name at the time the statement handle was allocated.

For drivers with connection-level, asynchronous-execution support, the statement attribute SQL_ATTR_ASYNC_ENABLE is read-only. Its value is the same as the value of the connection level attribute with the same name at the time the statement handle was allocated. Calling **SetStmtAttr** to set SQL_ATTR_ASYNC_ENABLE when the SQL_ASYNC_MODE InfoType returns SQL_AM_CONNECTION returns SQLSTATE HYC00 (Optional feature not implemented).

As a standard practice, applications should execute functions asynchronously only on single-thread operating systems. On multithread operating systems, applications should execute functions on separate threads rather than executing them asynchronously on the same thread. No functionality is lost if drivers that operate only on multithread operating systems do not need to support asynchronous execution. 

The following functions can be executed asynchronously:

BulkOperations ColAttribute ColumnPrivileges Columns CopyDesc DescribeCol DescribeParam ExecDirect Execute Fetch FetchScroll ForeignKeys GetData GetDescField\[1] GetDescRec\[1] GetDiagField GetDiagRec GetTypeInfo MoreResults NumParams NumResultCols ParamData Prepare PrimaryKeys ProcedureColumns Procedures PutData SetPos SpecialColumns Statistics TablePrivileges Tables.

\[1] These functions can be called asynchronously only if the descriptor is an implementation descriptor, not an application descriptor.

#### Return value

SQL_ASYNC_ENABLE_OFF or SQL_ASYNC_ENABLE_ON.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="UnbindCol"></a>UnbindCol

Unbinds the specified column buffer bound by **BindCol** for the given statement handle.

```
FUNCTION UnbindCol (BYVAL ColNum AS SQLUSMALLINT) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *ColNum* | Column number. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_INVALID_HANDLE, or SQL_ERROR.

# <a name="UnbindColumns"></a>UnbindColumns

Unbinds all column buffers bound by **BindCol** for the given statement handle.

```
FUNCTION UnbindColumns () AS SQLRETURN
```

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_INVALID_HANDLE, or SQL_ERROR.

# <a name="UnlockRecord"></a>UnlockRecord

Sets the cursor position in a rowset and unlocks the record.

```
FUNCTION UnlockRecord (BYVAL wRow AS SQLSETPOSIROW = 1) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wRow* | Optional. Row number inside the rowset. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_NEED_DATA, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="UpdateByBookmark"></a>UpdateByBookmark

Updates a set of rows where each row is identified by a bookmark.

```
FUNCTION UpdateByBookmark () AS SQLRETURN
```

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_NEED_DATA, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="UpdateRecord"></a>UpdateRecord

The driver positions the cursor in the row specified by *wRow* and updates the underlying row of data with the values in the rowset buffers (the *TargetValue* argument in **BindCol**). It retrieves the lengths of the data from the length indicator buffers (the StrLen_or_Ind argument in **BindCol**). If the length of any column is SQL_COLUMN_IGNORE, the column is not updated. After updating the row, the driver changes the corresponding element of the row status array to SQL_ROW_UPDATED or SQL_ROW_ACCESS_WITH_INFO (if the row status array exists).

```
FUNCTION UpdateRecord (BYVAL wRow AS SQLSETPOSIROW = 1) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wRow* | Optional. Row number inside the rowset.<br>Note: *wRow* is the row number inside the rowset (if the rowset has only one row, then *wRow* must be always 1). |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_NEED_DATA, SQL_STILL_EXECUTING, SQL_ERROR, or SQL_INVALID_HANDLE.

#### Example

```
#include once "Afx/COdbc/COdbc.inc"
USING Afx

' // Connect with the database
DIM wszConStr AS WSTRING * 260 = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=biblio.mdb"
DIM pDbc AS CODBC = wszConStr

' // Allocate an statement object
DIM pStmt AS COdbcStmt = @pDbc
IF pStmt.Handle = NULL THEN PRINT "Failed to allocate the statement handle": SLEEP : END

' // Cursor type
pStmt.SetMultiuserKeysetCursor

' // Bind the columns
DIM AS LONG lAuId, cbAuId
pStmt.BindCol(1, @lAuId, @cbAuId)
DIM wszAuthor AS WSTRING * 256, cbAuthor AS LONG
pStmt.BindCol(2, @wszAuthor, 256, @cbAuthor)
DIM iYearBorn AS SHORT, cbYearBorn AS LONG
pStmt.BindCol(3, @iYearBorn, @cbYearBorn)

' // Generate a result set
pStmt.ExecDirect ("SELECT * FROM Authors WHERE Au_Id=999")

' // Fetch the record
pstmt.Fetch

' // Fill the values of the binded application variables and its sizes
cbAuID = SQL_COLUMN_IGNORE
wszAuthor = "Félix Lope de Vega Carpio"
cbAuthor = LEN(wszAuthor) * 2   ' Unicode uses 2 bytes per character
iYearBorn = 1562
cbYearBorn = SIZEOF(iYearBorn)

' // Update the record
pStmt.UpdateRecord
IF pStmt.Error = FALSE THEN PRINT "Record updated" ELSE PRINT pStmt.GetErrorInfo
```
