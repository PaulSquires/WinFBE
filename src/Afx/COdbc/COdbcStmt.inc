' ########################################################################################
' Microsoft Windows
' File: COdbcStmt.inc
' Contents: ODBC wrapper classes (unicode)
' Implementation of the ODBC statement class.
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

' ========================================================================================
' Allocates an statement handle
' ========================================================================================
PRIVATE CONSTRUCTOR COdbcStmt (BYVAL pDbc AS COdbc PTR)
   IF pDbc THEN
      SetResult(SQLAllocHandle(SQL_HANDLE_STMT, pDbc->Handle, @m_hStmt))
      m_hDbc = pDbc->Handle
   END IF
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Cleanup
' ========================================================================================
PRIVATE DESTRUCTOR COdbcStmt
   IF m_hStmt THEN
      SQLCloseCursor(m_hStmt)
      SQLFreeHandle(SQL_HANDLE_STMT, m_hStmt)
   END IF
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the connection handle
' ========================================================================================
PRIVATE FUNCTION COdbcStmt.Handle () AS SQLHANDLE
   RETURN m_hStmt
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the connection handle
' ========================================================================================
PRIVATE FUNCTION COdbcStmt.DbcHandle () AS SQLHANDLE
   RETURN m_hDbc
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the current value of a field of a record of the diagnostic data structure
' (associated with a specific handle) that contains error, warning, and status information.
' ========================================================================================
PRIVATE FUNCTION COdbcStmt.GetDiagField (BYVAL RecNumber AS SQLSMALLINT, BYVAL DiagIdentifier AS SQLSMALLINT, _
BYVAL DiagInfoPtr AS SQLPOINTER, BYVAL BufferLength AS SQLSMALLINT, BYVAL StringLengthPtr AS SQLSMALLINT PTR) AS SQLRETURN
   RETURN SetResult(SQLGetDiagFieldW(SQL_HANDLE_STMT, m_hStmt, RecNumber, DiagIdentifier, DiagInfoPtr, BufferLength, StringLengthPtr), "COdbcStmt.GetDiagField", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the current values of multiple fields of a diagnostic record that contains
' error, warning and status information. Unlike GetDiagField, which returns one
' diagnostic call per call, GetDiagRec returns several commonly used fields of a
' diagnostic record, including the SQLSTATE, the native error code, and the
' diagnostic message text.
' ========================================================================================
PRIVATE FUNCTION COdbcStmt.GetDiagRec (BYVAL RecNumber AS SQLSMALLINT, BYVAL Sqlstate AS WSTRING PTR, BYVAL NativeError AS SQLINTEGER PTR, _
BYVAL MessageText AS WSTRING PTR, BYVAL BufferLength AS SQLSMALLINT, BYVAL TextLength AS SQLSMALLINT PTR) AS SQLRETURN
   RETURN SetResult(SQLGetDiagRecW(SQL_HANDLE_DBC, m_hStmt, RecNumber, Sqlstate, NativeError, MessageText, BufferLength, TextLength), "COdbcStmt.GetDiagRec", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' ========================================================================================

' ===========================================================================================
' Uses SQLGetDiagRec to retrieve an error description.
' iErrorCode: Optional. The error code returned by some of the methods of this interface.
' ===========================================================================================
PRIVATE FUNCTION COdbcStmt.GetErrorInfo (BYVAL iErrorCode AS SQLRETURN = 0) AS CWSTR
   RETURN Base.GetErrorInfo(SQL_HANDLE_STMT, m_hStmt, iErrorCode)
END FUNCTION
' ===========================================================================================

' ========================================================================================
' Gets the SqlState for the environment handle
' ========================================================================================
PRIVATE FUNCTION COdbcStmt.GetSqlState () AS CWSTR
   RETURN Base.GetSqlState(SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Executes the specified statement.
' ========================================================================================
PRIVATE FUNCTION COdbcStmt.ExecDirect (BYREF wszSqlStr AS WSTRING) AS SQLRETURN
   m_Result = SetResult(SQLExecDirectW(m_hStmt, @wszSqlStr, SQL_NTS), "COdbcStmt.ExecDirect", SQL_HANDLE_STMT, m_hStmt)
   ' // Retrieve the names of the fields (columns)
   ERASE rgColNames
   DIM numCols AS SQLSMALLINT = this.NumResultCols
   IF numCols THEN
      REDIM rgColNames(1 TO numCols)
      FOR idx AS LONG = 1 TO numCols
         rgColNames(idx) = this.ColName(idx)
      NEXT
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' =====================================================================================
' Gets the number of records in the result set.
' Uses an instruction like wszSqlStr = "SELECT COUNT(*) FROM Customers".
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.RecordCount (BYREF wszSqlStr AS WSTRING) AS LONG
   DIM numCols AS SQLSMALLINT, rCount AS SQLINTEGER
   m_Result = SetResult(SQLExecDirectW(m_hStmt, @wszSqlStr, SQL_NTS), "COdbcStmt.RecordCount", SQL_HANDLE_STMT, m_hStmt)
   IF SQL_SUCCEEDED(m_Result) THEN
      m_Result = SetResult(SQLNumResultCols(m_hStmt, @numCols), "COdbcStmt.RecordCount", SQL_HANDLE_STMT, m_hStmt)
      IF SQL_SUCCEEDED(m_Result) THEN
         IF numCols = 0 THEN RETURN 0
         m_Result = SetResult(SQLFetch(m_hStmt), "COdbcStmt.RecordCount", SQL_HANDLE_STMT, m_hStmt)
         IF SQL_SUCCEEDED(m_Result) = FALSE THEN RETURN 0
         m_Result = SetResult(SQLGetData(m_hStmt, 1, SQL_C_LONG, @rCount, 0, NULL), "COdbcStmt.RecordCount", SQL_HANDLE_STMT, m_hStmt)
         IF SQL_SUCCEEDED(m_Result) = 0 THEN RETURN 0
      END IF
   END IF
   RETURN rCount
END FUNCTION
' =====================================================================================

' ========================================================================================
' Cancels the processing on a statement.
' ========================================================================================
PRIVATE FUNCTION COdbcStmt.Cancel () AS SQLRETURN
   RETURN SetResult(SQLCancel(m_hStmt), "COdbcStmt.Cancel", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' ========================================================================================

' =====================================================================================
' Returns the number of columns of the result set.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.NumResultCols () AS SQLSMALLINT
   DIM numCols AS SQLSMALLINT
   SetResult(SQLNumResultCols(m_hStmt, @numCols), "COdbcStmt.NumResultCols", SQL_HANDLE_STMT, m_hStmt)
   RETURN numCols
END FUNCTION
' =====================================================================================

' =====================================================================================
' Determines whether more results are available on a statement containing SELECT,
' UPDATE, INSERT, or DELETE statements and, if so, initializes processing for those
' results.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.MoreResults () AS SQLRETURN
   RETURN SetResult(SQLMoreResults(m_hStmt), "COdbcStmt.MoreResults", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Returns the number of rows affected by an UPDATE, INSERT, or DELETE statement; an
' SQL_ADD, SQL_UPDATE_BY_BOOKMARK or SQL_DELETE_BY_BOOKMARK operation in
' BulkOperations; or an SQL_UPDATE or SQL_DELETE operation in SetPos.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.RowCount () AS SQLLEN
   DIM nCount AS SQLLEN
   SetResult(SQLRowCount(m_hStmt, @nCount), "COdbcStmt.RowCount", SQL_HANDLE_STMT, m_hStmt)
   RETURN nCount
END FUNCTION
' =====================================================================================

' ========================================================================================
' Fetches the next rowset of data from the result set and returns data for all bound columns.
' ========================================================================================
PRIVATE FUNCTION COdbcStmt.Fetch () AS BOOLEAN
   m_Result = SetResult(SqlFetch(m_hStmt), "COdbcStmt.Fetch", SQL_HANDLE_STMT, m_hStmt)
   RETURN SQL_SUCCEEDED(m_Result)
END FUNCTION
' ========================================================================================

' =====================================================================================
' Retrieves data for a single column in the result set.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.GetData (BYVAL ColumnNumber AS SQLUSMALLINT, BYVAL TargetType AS SQLSMALLINT, _
BYVAL TargetValue AS SQLPOINTER, BYVAL BufferLength AS SQLLEN, BYVAL StrLen_or_Ind AS SQLLEN PTR) AS SQLRETURN
   RETURN SetResult(SQLGetData(m_hStmt, ColumnNumber, TargetType, TargetValue, BufferLength, StrLen_or_Ind), "COdbcStmt.GetData", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.GetData (BYREF ColumnName AS WSTRING, BYVAL TargetType AS SQLSMALLINT, _
BYVAL TargetValue AS SQLPOINTER, BYVAL BufferLength AS SQLLEN, BYVAL StrLen_or_Ind AS SQLLEN PTR) AS SQLRETURN
   FOR i AS LONG = LBOUND(rgColNames) TO UBOUND (rgColNames)
      IF UCASE(rgColNames(i)) = UCASE(ColumnName) THEN
         RETURN this.GetData(i, TargetType, TargetValue, BufferLength, StrLen_or_Ind)
      END IF
   NEXT
END FUNCTION
' =====================================================================================

' ========================================================================================
' Returns the data in a specefied column as a string.
' Note: The size of the string must be specified in bytes, not characters.
' ========================================================================================
PRIVATE FUNCTION COdbcStmt.GetData (BYVAL ColumnNumber AS SQLUSMALLINT, BYVAL nMaxChars AS LONG = 256 * 2) AS CWSTR
   DIM cbLen AS SQLLEN
   DIM cws AS CWSTR = SPACE(nMaxChars + 1)   ' // make room for the null character
   m_Result = SetResult(SQLGetData(m_hStmt, ColumnNumber, SQL_C_WCHAR, cws, LEN(cws) * 2, @cbLen), "COdbcStmt.GetData", SQL_HANDLE_STMT, m_hStmt)
   ' // Had problems returning LEFT(**cws, cbLen) if the column has no data
   ' // because cbLen returns -1 and LEFT does not support negative numbers.
   DIM cwsOut AS CWSTR
   IF SQL_SUCCEEDED(m_Result) THEN
      IF cbLen <> SQL_NULL_DATA THEN cwsOut = LEFT(**cws, cbLen)
   END IF
   RETURN cwsOut
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION COdbcStmt.GetData (BYREF ColumnName AS WSTRING, BYVAL nMaxChars AS LONG = 256 * 2) AS CWSTR
   FOR i AS LONG = LBOUND(rgColNames) TO UBOUND (rgColNames)
      IF UCASE(rgColNames(i)) = UCASE(ColumnName) THEN
         RETURN this.GetData(i, nMaxChars)
      END IF
   NEXT
END FUNCTION
' ========================================================================================

' =====================================================================================
' Reads long variable char data
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.GetLongVarcharData (BYVAL ColumnNumber AS SQLSMALLINT) AS CWSTR

   DIM cbMemo AS SQLLEN
   DIM MemoBuf AS WSTRING * 32769         ' 32 Kb buffer + CHR(0)
   DIM cws AS CWSTR
   DIM wszSqlState AS WSTRING * 6
   DIM lNativeError AS SQLINTEGER
   DIM wszErrMsg AS WSTRING * SQL_MAX_MESSAGE_LENGTH + 1
   DIM cbbytes AS SQLSMALLINT

   DO
      m_Result = SetResult(SQLGetData(m_hStmt, ColumnNumber, SQL_C_WCHAR, @MemoBuf, 32769, @cbMemo), "COdbcStmt.GetLongVarcharData", SQL_HANDLE_STMT, m_hStmt)
      IF m_Result = SQL_SUCCESS THEN                  ' We have read all the content
         cws += MemoBuf
         EXIT DO
      ELSEIF m_Result = SQL_SUCCESS_WITH_INFO THEN
         m_Result = SetResult(SQLGetDiagRecW(SQL_HANDLE_STMT, m_hStmt, 1, @wszSqlState, @lNativeError, @wszErrMsg, SIZEOF(wszErrMsg), @cbbytes), "COdbcStmt.GetLongVarcharData", SQL_HANDLE_STMT, m_hStmt)
         IF m_Result <> SQL_SUCCESS THEN EXIT DO      ' Error in SQLGetDiagRec
         IF wszSqlState = "01004" THEN           ' Insufficient buffer
            cws += MemoBuf
            CONTINUE DO
         ELSE                                   ' Error
            EXIT DO
         END IF
      ELSE                                      ' Error
         IF m_Result = SQL_ERROR OR m_Result = SQL_INVALID_HANDLE THEN
            SetResult(m_Result, "COdbcStmt.GetLongVarcharData", SQL_HANDLE_STMT, m_hStmt)   ' Error
         END IF
      END IF
   LOOP

   DIM cbPos AS SQLINTEGER = INSTR(cws, CHR(13, 10))
   IF cbPos THEN cws = LEFT(**cws, cbPos - 1)
   RETURN cws

END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.GetLongVarcharData (BYREF ColumnName AS WSTRING) AS CWSTR
   FOR i AS LONG = LBOUND(rgColNames) TO UBOUND (rgColNames)
      IF UCASE(rgColNames(i)) = UCASE(ColumnName) THEN
         RETURN this.GetLongVarcharData(i)
      END IF
   NEXT
END FUNCTION
' =====================================================================================

' =====================================================================================
' Allows an application to send data for a parameter or column to the driver at
' statement execution time
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.PutData (BYVAL pData AS SQLPOINTER, BYVAL StrLen_or_Ind AS SQLLEN) AS SQLRETURN
   RETURN SetResult(SQLPutData(m_hStmt, pData, StrLen_or_Ind), "COdbcStmt.PutData", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Used in conjuntion with PutData to supply parameter data at statement
' execution time.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.ParamData (BYVAL ValuePtrPtr AS SQLPOINTER PTR) AS SQLRETURN
   RETURN SetResult(SQLParamData(m_hStmt, ValuePtrPtr), "COdbcStmt.ParamData", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Releases all parameter buffers set by BindParameter for the given statement handle.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.ResetParams () AS SQLRETURN
   RETURN SetResult(SQLFreeStmt(m_hStmt, SQL_RESET_PARAMS), "COdbcStmt.ResetParams", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Prepares an SQL string for execution.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.Prepare (BYREF StatementText AS WSTRING) AS SQLRETURN
   RETURN SetResult(SQLPrepareW(m_hStmt, @StatementText, LEN(StatementText)), "COdbcStmt.Prepare", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Executes a prepared statement, using the current values of the parameter marker
' variables if any parameter markers exist in the statement.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.Execute () AS SQLRETURN
   m_Result = SetResult(SQLExecute(m_hStmt), "COdbcStmt.Execute", SQL_HANDLE_STMT, m_hStmt)
   ' // Retrieve the names of the fields (columns)
   ERASE rgColNames
   DIM numCols AS SQLSMALLINT = this.NumResultCols
   IF numCols THEN
      REDIM rgColNames(1 TO numCols)
      FOR idx AS LONG = 1 TO numCols
         rgColNames(idx) = this.ColName(idx)
      NEXT
   END IF
   RETURN m_Result
END FUNCTION
' =====================================================================================

' =====================================================================================
' Returns the result descriptor --column name, type, column size, decimal
' digits, and nullability-- for one column in the result set. This information
' also is available in the fields of the IRD.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.DescribeCol (BYVAL ColumnNumber AS SQLUSMALLINT, BYVAL pwszColumnName AS WSTRING PTR, _
BYVAL BufferLength AS SQLSMALLINT, BYVAL NameLengthPtr AS SQLSMALLINT PTR, BYVAL DataTypePtr AS SQLSMALLINT PTR, _
BYVAL ColumnSizePtr AS SQLULEN PTR, BYVAL DecimalDigitsPtr AS SQLSMALLINT PTR, BYVAL Nullable AS SQLSMALLINT PTR) AS SQLRETURN
   RETURN SetResult(SQLDescribeColW(m_hStmt, ColumnNumber, pwszColumnName, BufferLength, NameLengthPtr, _
          DataTypePtr, ColumnSizePtr, DecimalDigitsPtr, Nullable), "COdbcStmt.DescribeCol", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Returns the number of parameters in an SQL statement.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.NumParams () AS SQLSMALLINT
   DIM cParams AS SQLSMALLINT
   SetResult(SQLNumParams(m_hStmt, @cParams), "COdbcStmt.NumParams", SQL_HANDLE_STMT, m_hStmt)
   RETURN cParams
END FUNCTION
' =====================================================================================

' =====================================================================================
' Returns the description of a parameter marker associated with a prepared SQL
' statement. This information is also available in the fields of the IRD.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.DescribeParam (BYVAL ParameterNumber AS SQLUSMALLINT, BYVAL DataTypePtr AS SQLSMALLINT PTR, _
BYVAL ParameterSizePtr AS SQLULEN PTR, BYVAL DecimalDigitsPtr AS SQLSMALLINT PTR, BYVAL NullablePtr AS SQLSMALLINT PTR) AS SQLRETURN
   RETURN SetResult(SQLDescribeParam(m_hStmt, ParameterNumber, DataTypePtr, ParameterSizePtr, DecimalDigitsPtr, NullablePtr), "COdbcStmt.DescribeParam", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Checks if the column is null.
' Note: Should not be used with a column that is currently binded to a variable or
' buffer or it will return an error. The binding functions already return an
' indicator in the last parameter.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.ColIsNull (BYVAL ColumnNumber AS SQLUSMALLINT) AS BOOLEAN
   DIM buffer AS WSTRING * 256, ind AS SQLLEN
   m_Result = SetResult(SQLGetData(m_hStmt, ColumnNumber, SQL_C_CHAR, @buffer, 256, @ind), "COdbcStmt.ColIsNull", SQL_HANDLE_STMT, m_hStmt)
   IF SQL_SUCCEEDED(m_Result) THEN
      IF ind = SQL_NULL_DATA THEN RETURN TRUE ELSE RETURN FALSE
   END IF
END FUNCTION
' =====================================================================================

' #####################################################################################
' BINDING
' #####################################################################################

' ========================================================================================
' Binds application data buffers to columns in the result set.
' ========================================================================================
PRIVATE FUNCTION COdbcStmt.BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetType AS SQLSMALLINT, _
BYVAL TargetValue AS SQLPOINTER, BYVAL BufferLength AS SQLLEN, BYVAL StrLen_or_Ind AS ANY PTR) AS SQLRETURN
   RETURN SetResult(SQLBindCol(m_hStmt, ColNumber, TargetType, TargetValue, BufferLength, StrLen_or_Ind), "COdbcStmt.BindCol", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' ========================================================================================

' =====================================================================================
' A BYTE variable is used to bind the column.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS BYTE PTR, BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
   RETURN SetResult(SQLBindCol(m_hStmt, ColNumber, SQL_C_STINYINT, TargetValue, 3, StrLen_or_IndPtr), "COdbcStmt.BindCol", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' An UBYTE variable is used to bind the column.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS UBYTE PTR, BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
   RETURN SetResult(SQLBindCol(m_hStmt, ColNumber, SQL_C_UTINYINT, TargetValue, 3, StrLen_or_IndPtr), "COdbcStmt.BindCol", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' A SHORT variable is used to bind the column.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS SHORT PTR, BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
   RETURN SetResult(SQLBindCol(m_hStmt, ColNumber, SQL_C_SSHORT, TargetValue, 0, StrLen_or_IndPtr), "COdbcStmt.BindCol", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' An USHORT variable is used to bind the column.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS USHORT PTR, BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
   RETURN SetResult(SQLBindCol(m_hStmt, ColNumber, SQL_C_USHORT, TargetValue, 0, StrLen_or_IndPtr), "COdbcStmt.BindCol", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' A LONG variable is used to bind the column.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS LONG PTR, BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
   RETURN SetResult(SQLBindCol(m_hStmt, ColNumber, SQL_C_LONG, TargetValue, 0, StrLen_or_IndPtr), "COdbcStmt.BindCol", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' An ULONG variable is used to bind the column.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS ULONG PTR, BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
   RETURN SetResult(SQLBindCol(m_hStmt, ColNumber, SQL_C_ULONG, TargetValue, 0, StrLen_or_IndPtr), "COdbcStmt.BindCol", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' A SINGLE variable is used to bind the column.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS SINGLE PTR, BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
   RETURN SetResult(SQLBindCol(m_hStmt, ColNumber, SQL_C_FLOAT, TargetValue, 0, StrLen_or_IndPtr), "COdbcStmt.BindCol", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' A DOUBLE variable is used to bind the column.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS DOUBLE PTR, BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
   RETURN SetResult(SQLBindCol(m_hStmt, ColNumber, SQL_C_DOUBLE, TargetValue, 0, StrLen_or_IndPtr), "COdbcStmt.BindCol", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' A LONGINT variable is used to bind the column.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS LONGINT PTR, BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
   RETURN SetResult(SQLBindCol(m_hStmt, ColNumber, SQL_C_SBIGINT, TargetValue, 0, StrLen_or_IndPtr), "COdbcStmt.BindCol", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' An ULONGINT variable is used to bind the column.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS ULONGINT PTR, BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
   RETURN SetResult(SQLBindCol(m_hStmt, ColNumber, SQL_C_UBIGINT, TargetValue, 0, StrLen_or_IndPtr), "COdbcStmt.BindCol", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' A ZSTRING is used to bind the column.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS ZSTRING PTR, BYVAL BufferLenght AS SQLLEN, BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
   RETURN SetResult(SQLBindCol(m_hStmt, ColNumber, SQL_C_CHAR, TargetValue, BufferLenght, StrLen_or_IndPtr), "COdbcStmt.BindCol", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' A WSTRING is used to bind the column.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS WSTRING PTR, BYVAL BufferLenght AS SQLLEN, BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
   RETURN SetResult(SQLBindCol(m_hStmt, ColNumber, SQL_C_WCHAR, TargetValue, BufferLenght, StrLen_or_IndPtr), "COdbcStmt.BindCol", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' A DATE_STRUCT is used to bind the column.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS DATE_STRUCT PTR, BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
   RETURN SetResult(SQLBindCol(m_hStmt, ColNumber, SQL_C_DATE, TargetValue, 0, StrLen_or_IndPtr), "COdbcStmt.BindCol", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' A TIME_STRUCT is used to bind the column.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS TIME_STRUCT PTR, BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
   RETURN SetResult(SQLBindCol(m_hStmt, ColNumber, SQL_C_TIME, TargetValue, 0, StrLen_or_IndPtr), "COdbcStmt.BindCol", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' A TIMESTAMP_STRUCT is used to bind the column.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS TIMESTAMP_STRUCT PTR, BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
   RETURN SetResult(SQLBindCol(m_hStmt, ColNumber, SQL_C_TIMESTAMP, TargetValue, 0, StrLen_or_IndPtr), "COdbcStmt.BindCol", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' A memory buffer is used to bind the column.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS ANY PTR, BYVAL BufferLenght AS SQLLEN, BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
   RETURN SetResult(SQLBindCol(m_hStmt, ColNumber, SQL_C_BINARY, TargetValue, BufferLenght, StrLen_or_IndPtr), "COdbcStmt.BindCol", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' An integer variable is used to bind the column.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.BindColToBit (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS SHORT PTR, BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
   RETURN SetResult(SQLBindCol(m_hStmt, ColNumber, SQL_C_BIT, TargetValue, 1, StrLen_or_IndPtr), "COdbcStmt.BindColToBit", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' An ASCIIZ * 22 string is used to bind the column.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.BindColToNumeric (BYVAL ColNumber AS SQLUSMALLINT, BYREF TargetValue AS ZSTRING, BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
   RETURN SetResult(SQLBindCol(m_hStmt, ColNumber, SQL_C_CHAR, @TargetValue, 21, StrLen_or_IndPtr), "COdbcStmt.BindColToNumeric", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' An ASCIIZ * 22 string is used to bind the column.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.BindColToDecimal (BYVAL ColNumber AS SQLUSMALLINT, BYREF TargetValue AS ZSTRING, BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
   RETURN SetResult(SQLBindCol(m_hStmt, ColNumber, SQL_C_CHAR, @TargetValue, 21, StrLen_or_IndPtr), "COdbcStmt.BindColToDecimal", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' ========================================================================================
' Releases the specified column buffer for the given statement handle.
' ========================================================================================
PRIVATE FUNCTION COdbcStmt.UnbindCol (BYVAL nCol AS SQLUSMALLINT) AS SQLRETURN
   RETURN SetResult(SQLBindCol (m_hStmt, nCol, SQL_C_DEFAULT, NULL, 0, NULL), "COdbcStmt.UnbindCol", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Releases all column buffers bound to the specified statement handle.
' ========================================================================================
PRIVATE FUNCTION COdbcStmt.UnbindColumns () AS SQLRETURN
   RETURN SetResult(SQLFreeStmt (m_hStmt, SQL_UNBIND), "COdbcStmt.UnbindColumns", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Binds a buffer to a parameter marker in a SQL statement. Supports binding to a Unicode
' C data type, even if the underlying driver does not support Unicode data.
' ========================================================================================
PRIVATE FUNCTION COdbcStmt.BindParameter (BYVAL ParameterNumber AS SQLUSMALLINT, BYVAL InputOutputType AS SQLSMALLINT, _
BYVAL ValueType AS SQLSMALLINT, BYVAL ParameterType AS SQLSMALLINT, BYVAL ColumnSize AS SQLULEN, BYVAL DecimalDigits AS SQLSMALLINT, _
BYVAL ParameterValuePtr AS SQLPOINTER, BYVAL BufferLength AS SQLLEN, BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
   RETURN SetResult(SQLBindParameter(m_hStmt, ParameterNumber, InputOutputType, ValueType, ParameterType, ColumnSize, DecimalDigits, BYVAL ParameterValuePtr, BufferLength, StrLen_or_IndPtr), "COdbcStmt.BindParameter", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' ========================================================================================

' #####################################################################################
' CURSOR RELATED METHODS
' #####################################################################################

' =====================================================================================
' Closes the cursor associated with m_hStmt (if one was defined) and discards all pending
' reults. Returns SQLSTATE 24000 (Invalid cursor state) if no cursor is open in the statement.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.CloseCursor () AS SQLRETURN
   RETURN SetResult(SQLCloseCursor(m_hStmt), "COdbcStmt.CloseCursor", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Returns the cursor name associated with a specified statement.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.GetCursorName () AS CWSTR
   DIM wszCursorName AS WSTRING * 64
   SetResult(SQLGetCursorNameW(m_hStmt, @wszCursorName, 64, NULL), "COdbcStmt.GetCursorName", SQL_HANDLE_STMT, m_hStmt)
   RETURN wszCursorName
END FUNCTION
' =====================================================================================

' =====================================================================================
' Associates a cursor name with an active statement.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.SetCursorName (BYREF wszCursorName AS WSTRING) AS SQLRETURN
   RETURN SetResult(SQLSetCursorNameW(m_hStmt, @wszCursorName, SQL_NTS), "COdbcStmt.SetCursorName", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Fetches the specified rowset of data from the result set and returns data for all
' bound columns. Rowsets can be specified at an absolute position or by bookmark.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.FetchScroll (BYVAL FetchOrientation AS SQLSMALLINT, BYVAL FetchOffset AS SQLLEN) AS SQLRETURN
   RETURN SEtResult(SqlFetchScroll(m_hStmt, FetchOrientation, FetchOffset), "COdbcStmt.FetchScroll", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Fetches the specified rowset of data from the result set and returns data for all
' bound columns. Rowsets can be specified at an absolute or relative position or
' by bookmark.
' NOTE: Since the called function, ExtendedFetch, has been deprecated and replaced by
' FetchScroll, users must use the FetchScroll method instead of this one.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.ExtendedFetch (BYVAL FetchOrientation AS SQLUSMALLINT, BYVAL FetchOffset AS SQLLEN, BYVAL RowCountPtr AS ANY PTR, BYVAL RowStatusArray AS SQLUSMALLINT PTR) AS SQLRETURN
   RETURN SetResult(SQLExtendedFetch(m_hStmt, FetchOrientation, FetchOffset, RowCountPtr, BYVAL RowStatusArray), "COdbcStmt.ExtendedFetch", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Returns the first rowset in the result set.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.MoveFirst () AS SQLRETURN
   RETURN SetResult(SQLFetchScroll(m_hStmt, SQL_FETCH_FIRST, 0), "COdbcStmt.MoveFirst", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Returns the last rowset in the result set.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.MoveLast () AS SQLRETURN
   RETURN SetResult(SQLFetchScroll(m_hStmt, SQL_FETCH_LAST, 0), "COdbcStmt.MoveLast", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Returns the next rowset.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.MoveNext () AS SQLRETURN
   RETURN SetResult(SQLFetchScroll(m_hStmt, SQL_FETCH_NEXT, 0), "COdbcStmt.MoveNext", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Returns the prior rowset.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.MovePrevious () AS SQLRETURN
   RETURN SetResult(SQLFetchScroll(m_hStmt, SQL_FETCH_PRIOR, 0), "COdbcStmt.MovePrevious", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Moves the cursor forward or backward FetchOffset rowsets.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.Move (BYVAL FetchOffset AS SQLLEN) AS SQLRETURN
   RETURN SetResult(SQLFetchScroll(m_hStmt, SQL_FETCH_RELATIVE, FetchOffset), "COdbcStmt.Move", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Returns the rowset starting at row FetchOffset.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.SetAbsolutePosition (BYVAL FetchOffset AS SQLLEN) AS SQLRETURN
   RETURN SetResult(SQLFetchScroll(m_hStmt, SQL_FETCH_ABSOLUTE, FetchOffset), "COdbcStmt.SetAbsolutePosition", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Returns the rowset FetchOffset from the start of the current rowset.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.SetRelativePosition (BYVAL FetchOffset AS SQLLEN) AS SQLRETURN
   RETURN SetResult(SQLFetchScroll(m_hStmt, SQL_FETCH_RELATIVE, FetchOffset), "COdbcStmt.SetRelativePosition", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Sets the cursor position in a rowset.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.SetPos (BYVAL wRow AS SQLSETPOSIROW, BYVAL fOption AS SQLUSMALLINT, BYVAL fLock AS SQLUSMALLINT) AS SQLRETURN
   RETURN SetResult(SqlSetPos(m_hStmt, wRow, fOption, fLock), "COdbcStmt.SetPos", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Sets the cursor position in a rowset.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.SetPosition (BYVAL wRow AS SQLSETPOSIROW) AS SQLRETURN
   RETURN SetResult(SQLSetPos(m_hStmt, wRow, SQL_POSITION, SQL_LOCK_NO_CHANGE), "COdbcStmt.SetPosition", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' The driver positions the cursor in the row specified by RowNumber and updates the
' underlying row of data with the values in the rowset buffers (the TargetValuePtr
' argument in BindCol). It retrieves the lengths of the data from the length
' indicator buffers (the StrLen_or_IndPtr argument in BindCol). If the length of
' any column is SQL_COLUMN_IGNORE, the column is not updated. After updating the row,
' the driver changes the corresponding element of the row status array to
' SQL_ROW_UPDATED or SQL_ROW_ACCESS_WITH_INFO (if the row status array exists).
' Note: irow is the row number inside the rowset (if the rowset has only one row, then
' irow must be always 1).
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.UpdateRecord (BYVAL wRow AS SQLSETPOSIROW = 1) AS SQLRETURN
   RETURN SetResult(SQLSetPos(m_hStmt, wRow, SQL_UPDATE, SQL_LOCK_NO_CHANGE), "COdbcStmt.UpdateRecord", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Sets the cursor position in a rowset and allows to refresh data in the rowset.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.RefreshRecord (BYVAL wRow AS SQLSETPOSIROW = 1, BYVAL fLock AS SQLUSMALLINT = SQL_LOCK_NO_CHANGE) AS SQLRETURN
   RETURN SetResult(SQLSetPos(m_hStmt, wRow, SQL_REFRESH, fLock), "COdbcStmt.RefreshRecord", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Sets the cursor position in a rowset and locks the record.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.LockRecord (BYVAL wRow AS SQLSETPOSIROW = 1) AS SQLRETURN
   RETURN SetResult(SQLSetPos(m_hStmt, wRow, SQL_POSITION, SQL_LOCK_EXCLUSIVE), "COdbcStmt.LockRecord", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Sets the cursor position in a rowset and locks the record.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.UnlockRecord (BYVAL wRow AS SQLSETPOSIROW = 1) AS SQLRETURN
   RETURN SetResult(SQLSetPos(m_hStmt, wRow, SQL_POSITION, SQL_LOCK_UNLOCK), "COdbcStmt.UnlockRecord", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' The driver positions the cursor on the row specified by RowNumber and deletes the
' underlying row of data. It changes the corresponding element of the row status array
' to SQL_ROW_DELETED. After the row has been deleted, the following are not valid for
' the row: positioned update and delete statements, calls to GetData and calls to
' SetPos with Operation set to anything except SQL_POSITION. For drivers that support
' packing, the row is deleted from the cursor when new data is retrieved from the data
' source. Whether the row remains visible depends on the cursor type. For example,
' deleted rows are visible to static and keyset-driven cursors but invisible to dynamic
' cursors.
' The row operation array pointed to by the SQL_ATTR_ROW_OPERATION_PTR statement
' attribute can be used to indicate that a row in the current rowset should be ignored
' during a bulk delete.
' Note: irow is the row number inside the rowset (if it is a single row rowset, irow
' must be always 1).
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.DeleteRecord (BYVAL wRow AS SQLSETPOSIROW = 1) AS SQLRETURN
   RETURN SetResult(SQLSetPos(m_hStmt, wRow, SQL_DELETE, SQL_LOCK_NO_CHANGE), "COdbcStmt.DeleteRecord", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Performs bulk insertions and bulk bookmark operations, including update, delete, and
' fetch by bookmark.
' Operation:
'   SQL_ADD
'   SQL_UPDATE_BY_BOOKMARK
'   SQL_DELETE_BY_BOOKMARK
'   SQL_FETCH_BY_BOOKMARK
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.BulkOperations (BYVAL Operation AS SQLSMALLINT) AS SQLRETURN
   RETURN SetResult(SQLBulkOperations(m_hStmt, Operation), "COdbcStmt.BulkOperations", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Adds a record in the the database
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.AddRecord () AS SQLRETURN
   RETURN SetResult(SQLBulkOperations(m_hStmt, SQL_ADD), "COdbcStmt.AddRecord", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Updates a set of rows where each row os identified by a bookmark.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.UpdateByBookmark () AS SQLRETURN
   RETURN SEtResult(SQLBulkOperations(m_hStmt, SQL_UPDATE_BY_BOOKMARK), "COdbcStmt.UpdateByBookmark", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Deletes a set of rows where each row os identified by a bookmark.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.DeleteByBookmark () AS SQLRETURN
   RETURN SetResult(SQLBulkOperations(m_hStmt, SQL_DELETE_BY_BOOKMARK), "COdbcStmt.DeleteByBookmark", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Fetches a set of rows where each row os identified by a bookmark.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.FetchByBookmark () AS SQLRETURN
   RETURN SetResult(SQLBulkOperations(m_hStmt, SQL_FETCH_BY_BOOKMARK), "COdbcStmt.FetchByBookmark", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' ColAttribute returns descriptor information for a column in a result set. Descriptor
' information is returned as a character string, a 32-bit descriptor-dependent value, or
' an integer value.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.ColAttribute (BYVAL ColumnNumber AS SQLUSMALLINT, BYVAL FieldIdentifier AS SQLUSMALLINT, _
BYVAL CharacterAttribute AS SQLPOINTER, BYVAL BufferLength AS SQLSMALLINT, BYVAL StringLength AS SQLSMALLINT PTR, _
BYVAL NumericAttribute AS ANY PTR) AS SQLRETURN
   RETURN SetResult(SQLColAttributeW(m_hStmt, ColumnNumber, FieldIdentifier, CharacterAttribute, BufferLength, StringLength, NumericAttribute), "COdbcStmt.ColAttribute", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Returns SQL_TRUE if the column is an autoincrementing column, SQL_FALSE if the column
' is not an autoincrementing column or is not numeric.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.ColAutoUniqueValue (BYVAL ColNum AS SQLUSMALLINT) AS BOOLEAN
   DIM lAttr AS SQLINTEGER
   SetResult(SQLColAttributeW(m_hStmt, ColNum, SQL_DESC_AUTO_UNIQUE_VALUE, NULL, 0, NULL, cast(ANY PTR, @lAttr)), "COdbcStmt.ColAutoUniqueValue", SQL_HANDLE_STMT, m_hStmt)
   RETURN lAttr
END FUNCTION
' =====================================================================================

' =====================================================================================
' The base column name for the result set column. If a base column name does not exist
' (as in the case of columns that are expressions), then this variable contains an empty
' string.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.ColBaseColumnName (BYVAL ColNum AS SQLUSMALLINT) AS CWSTR
   DIM wszCharAttr AS WSTRING * 255, CharAttrLen AS SQLSMALLINT
   SetResult(SQLColAttributeW(m_hStmt, ColNum, SQL_DESC_BASE_COLUMN_NAME, @wszCharAttr, 255, @CharAttrLen, NULL), "COdbcStmt.ColBaseColumnName", SQL_HANDLE_STMT, m_hStmt)
   RETURN wszCharAttr
END FUNCTION
' =====================================================================================

' =====================================================================================
' The name of the base table that contains the column. If the base table name cannot be
' defined or is not applicable, then this variable contains an empty string.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.ColBaseTableName (BYVAL ColNum AS SQLUSMALLINT) AS CWSTR
   DIM wszCharAttr AS WSTRING * 255, CharAttrLen AS SQLSMALLINT
   SetResult(SQLColAttributeW(m_hStmt, ColNum, SQL_DESC_BASE_TABLE_NAME, @wszCharAttr, 255, @CharAttrLen, NULL), "COdbcStmt.ColBaseTableName", SQL_HANDLE_STMT, m_hStmt)
   RETURN wszCharAttr
END FUNCTION
' =====================================================================================

' =====================================================================================
' Returns SQL_TRUE if the column is treated as case-sensitive for collations and
' comparisons.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.ColCaseSensitive (BYVAL ColNum AS SQLUSMALLINT) AS BOOLEAN
   DIM lAttr AS SQLINTEGER
   SetResult(SQLColAttributeW(m_hStmt, ColNum, SQL_DESC_CASE_SENSITIVE, NULL, 0, NULL, cast(ANY PTR, @lAttr)), "COdbcStmt.ColCaseSensitive", SQL_HANDLE_STMT, m_hStmt)
   RETURN lAttr
END FUNCTION
' =====================================================================================

' =====================================================================================
' The catalog of the table that contains the column. The returned value is
' implementation-defined if the column is an expression or if the column is part of a
' view. If the data source does not support catalogs or the catalog name cannot be
' determined, an empty string is returned. This VARCHAR record field is not limited to
' 128 characters.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.ColCatalogName (BYVAL ColNum AS SQLUSMALLINT) AS CWSTR
   DIM wszCharAttr AS WSTRING * 255, CharAttrLen AS SQLSMALLINT
   SetResult(SQLColAttributeW(m_hStmt, ColNum, SQL_DESC_CATALOG_NAME, @wszCharAttr, 255, @CharAttrLen, NULL), "COdbcStmt.ColCatalogName", SQL_HANDLE_STMT, m_hStmt)
   RETURN wszCharAttr
END FUNCTION
' =====================================================================================

' =====================================================================================
' Returns The concise data type.
' For the datetime and interval data types, this field returns the concise data type; for
' example, SQL_TYPE_TIME or SQL_INTERVAL_YEAR.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.ColConciseType (BYVAL ColNum AS SQLUSMALLINT) AS SQLINTEGER
   DIM lAttr AS SQLINTEGER
   SetResult(SQLColAttributeW(m_hStmt, ColNum, SQL_DESC_CONCISE_TYPE, NULL, 0, NULL, cast(ANY PTR, @lAttr)), "COdbcStmt.ColConciseType", SQL_HANDLE_STMT, m_hStmt)
   RETURN lAttr
END FUNCTION
' =====================================================================================

' =====================================================================================
' The number of columns available in the result set. This returns 0 if there are no
' columns in the result set. The value in the ColumnNumber argument is ignored.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.ColCount () AS SQLINTEGER
   DIM lAttr AS SQLINTEGER
   SetResult(SQLColAttributeW(m_hStmt, 0, SQL_DESC_COUNT, NULL, 0, NULL, cast(ANY PTR, @lAttr)), "COdbcStmt.ColCount", SQL_HANDLE_STMT, m_hStmt)
   RETURN lAttr
END FUNCTION
' =====================================================================================

' =====================================================================================
' Maximum number of characters required to display data from the column.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.ColDisplaySize (BYVAL ColNum AS SQLUSMALLINT) AS SQLINTEGER
   DIM lAttr AS SQLINTEGER
   SetResult(SQLColAttributeW(m_hStmt, ColNum, SQL_DESC_DISPLAY_SIZE, NULL, 0, NULL, cast(ANY PTR, @lAttr)), "COdbcStmt.ColDisplaySize", SQL_HANDLE_STMT, m_hStmt)
   RETURN lAttr
END FUNCTION
' =====================================================================================

' =====================================================================================
' Returns SQL_TRUE if the column has a fixed precision and nonzero scale that are data
' source-specific.
' Returns SQL_FALSE if the column does not have a fixed precision and nonzero scale that
' are data source-specific.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.ColFixedPrecScale (BYVAL ColNum AS SQLUSMALLINT) AS BOOLEAN
   DIM lAttr AS SQLINTEGER
   SetResult(SQLColAttributeW(m_hStmt, ColNum, SQL_DESC_FIXED_PREC_SCALE, NULL, 0, NULL, cast(ANY PTR, @lAttr)), "COdbcStmt.ColFixedPrecScale", SQL_HANDLE_STMT, m_hStmt)
   RETURN lAttr
END FUNCTION
' =====================================================================================

' =====================================================================================
' The column label or title. For example, a column named EmpName might be labeled
' Employee Name or might be labeled with an alias.
' If a column does not have a label, the column name is returned. If the column is
' unlabeled and unnamed, an empty string is returned.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.ColLabel (BYVAL ColNum AS SQLUSMALLINT) AS CWSTR
   DIM wszCharAttr AS WSTRING * 255, CharAttrLen AS SQLSMALLINT
   SetResult(SQLColAttributeW(m_hStmt, ColNum, SQL_DESC_LABEL, @wszCharAttr, 255, @CharAttrLen, NULL), "COdbcStmt.ColLabel", SQL_HANDLE_STMT, m_hStmt)
   RETURN wszCharAttr
END FUNCTION
' =====================================================================================

' =====================================================================================
' A numeric value that is either the maximum or actual character length of a character
' string or binary data type. It is the maximum character length for a fixed-length data
' type, or the actual character length for a variable-length data type. Its value always
' excludes the null-termination byte that ends the character string.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.ColLength (BYVAL ColNum AS SQLUSMALLINT) AS SQLINTEGER
   DIM lAttr AS SQLINTEGER
   SetResult(SQLColAttributeW(m_hStmt, ColNum, SQL_DESC_LENGTH, NULL, 0, NULL, cast(ANY PTR, @lAttr)), "COdbcStmt.ColLength", SQL_HANDLE_STMT, m_hStmt)
   RETURN lAttr
END FUNCTION
' =====================================================================================

' =====================================================================================
' This VARCHAR(128) record field contains the character or characters that the driver
' recognizes as a prefix for a literal of this data type.
' This field contains an empty string for a data type for which a literal prefix is not
' applicable.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.ColLiteralPrefix (BYVAL ColNum AS SQLUSMALLINT) AS CWSTR
   DIM wszCharAttr AS WSTRING * 255, CharAttrLen AS SQLSMALLINT
   SetResult(SQLColAttributeW(m_hStmt, ColNum, SQL_DESC_LITERAL_PREFIX, @wszCharAttr, 255, @CharAttrLen, NULL), "COdbcStmt.ColLiteralPrefix", SQL_HANDLE_STMT, m_hStmt)
   RETURN wszCharAttr
END FUNCTION
' =====================================================================================

' =====================================================================================
' This VARCHAR(128) record field contains the character or characters that the driver
' recognizes as a suffix for a literal of this data type.
' This field contains an empty string for a data type for which a literal suffix is not
' applicable.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.ColLiteralSuffix (BYVAL ColNum AS SQLUSMALLINT) AS CWSTR
   DIM wszCharAttr AS WSTRING * 255, CharAttrLen AS SQLSMALLINT
   SetResult(SQLColAttributeW(m_hStmt, ColNum, SQL_DESC_LITERAL_SUFFIX, @wszCharAttr, 255, @CharAttrLen, NULL), "COdbcStmt.ColLiteralSuffix", SQL_HANDLE_STMT, m_hStmt)
   RETURN wszCharAttr
END FUNCTION
' =====================================================================================

' =====================================================================================
' This VARCHAR(128) record field contains any localized (native language) name for the
' data type that may be different from the regular name of the data type. If there is no
' localized name, then an empty string is returned. This field is for display purposes
' only. The character set of the string is locale-dependent and is typically the default
' character set of the server.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.ColLocalTypeName (BYVAL ColNum AS SQLUSMALLINT) AS CWSTR
   DIM wszCharAttr AS WSTRING * 255, CharAttrLen AS SQLSMALLINT
   SetResult(SQLColAttributeW(m_hStmt, ColNum, SQL_DESC_LOCAL_TYPE_NAME, @wszCharAttr, 255, @CharAttrLen, NULL), "COdbcStmt.ColLocalTypeName", SQL_HANDLE_STMT, m_hStmt)
   RETURN wszCharAttr
END FUNCTION
' =====================================================================================

' =====================================================================================
' The column alias, if it applies. If the column alias does not apply, the column name is
' returned. In either case, SQL_DESC_UNNAMED is set to SQL_NAMED. If there is no column
' name or a column alias, an empty string is returned and SQL_DESC_UNNAMED is set to
' SQL_UNNAMED.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.ColName (BYVAL ColNum AS SQLUSMALLINT) AS CWSTR
   DIM wszCharAttr AS WSTRING * 255, CharAttrLen AS SQLSMALLINT
   SetResult(SQLColAttributeW(m_hStmt, ColNum, SQL_DESC_NAME, @wszCharAttr, 255, @CharAttrLen, NULL), "COdbcStmt.ColName", SQL_HANDLE_STMT, m_hStmt)
   RETURN wszCharAttr
END FUNCTION
' =====================================================================================

' =====================================================================================
' Returns SQL_NULLABLE if the column can have NULL values; SQL_NO_NULLS if the column
' does not have NULL values; or SQL_NULLABLE_UNKNOWN if it is not known whether the
' column accepts NULL values.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.ColNullable (BYVAL ColNum AS SQLUSMALLINT) AS SQLINTEGER
   DIM lAttr AS SQLINTEGER
   SetResult(SQLColAttributeW(m_hStmt, ColNum, SQL_DESC_NULLABLE, NULL, 0, NULL, cast(ANY PTR, @lAttr)), "COdbcStmt.ColNullable", SQL_HANDLE_STMT, m_hStmt)
   RETURN lAttr
END FUNCTION
' =====================================================================================

' =====================================================================================
' If the data type in the SQL_DESC_TYPE field is an approximate numeric data type, this
' SQLINTEGER field contains a value of 2 because the SQL_DESC_PRECISION field contains
' the number of bits.
' If the data type in the SQL_DESC_TYPE field is an exact numeric data type, this field
' contains a value of 10 because the SQL_DESC_PRECISION field contains the number of
' decimal digits.
' This field is set to 0 for all non-numeric data types.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.ColNumPrecRadix (BYVAL ColNum AS SQLUSMALLINT) AS SQLINTEGER
   DIM lAttr AS SQLINTEGER
   SetResult(SQLColAttributeW(m_hStmt, ColNum, SQL_DESC_NUM_PREC_RADIX, NULL, 0, NULL, cast(ANY PTR, @lAttr)), "COdbcStmt.ColNumPrecRadix", SQL_HANDLE_STMT, m_hStmt)
   RETURN lAttr
END FUNCTION
' =====================================================================================

' =====================================================================================
' The length, in bytes, of a character string or binary data type. For fixed-length
' character or binary types, this is the actual length in bytes. For variable-length
' character or binary types, this is the maximum length in bytes. This value includes the
' null terminator.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.ColOctetLength (BYVAL ColNum AS SQLUSMALLINT) AS SQLINTEGER
   DIM lAttr AS SQLINTEGER
   SetResult(SQLColAttributeW(m_hStmt, ColNum, SQL_DESC_OCTET_LENGTH, NULL, 0, NULL, cast(ANY PTR, @lAttr)), "COdbcStmt.ColOctetLength", SQL_HANDLE_STMT, m_hStmt)
   RETURN lAttr
END FUNCTION
' =====================================================================================

' =====================================================================================
' A numeric value that for a numeric data type denotes the applicable precision. For data
' types SQL_TYPE_TIME, SQL_TYPE_TIMESTAMP, and all the interval data types that
' represent a time interval, its value is the applicable precision of the fractional
' seconds component.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.ColPrecision (BYVAL ColNum AS SQLUSMALLINT) AS SQLINTEGER
   DIM lAttr AS SQLINTEGER
   SetResult(SQLColAttributeW(m_hStmt, ColNum, SQL_DESC_PRECISION, NULL, 0, NULL, cast(ANY PTR, @lAttr)), "COdbcStmt.ColPrecision", SQL_HANDLE_STMT, m_hStmt)
   RETURN lAttr
END FUNCTION
' =====================================================================================

' =====================================================================================
' A numeric value that is the applicable scale for a numeric data type.
' For DECIMAL and NUMERIC data types, this is the defined scale. It is undefined for all
' other data types.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.ColScale (BYVAL ColNum AS SQLUSMALLINT) AS SQLINTEGER
   DIM lAttr AS SQLINTEGER
   SetResult(SQLColAttributeW(m_hStmt, ColNum, SQL_DESC_SCALE, NULL, 0, NULL, cast(ANY PTR, @lAttr)), "COdbcStmt.ColScale", SQL_HANDLE_STMT, m_hStmt)
   RETURN lAttr
END FUNCTION
' =====================================================================================

' =====================================================================================
' The schema of the table that contains the column. The returned value is implementation-
' defined if the column is an expression or if the column is part of a view. If the data
' source does not support schemas or the schema name cannot be determined, an empty string
' is returned. This VARCHAR record field is not limited to 128 characters.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.ColSchemaName (BYVAL ColNum AS SQLUSMALLINT) AS CWSTR
   DIM wszCharAttr AS WSTRING * 255, CharAttrLen AS SQLSMALLINT
   SetResult(SQLColAttributeW(m_hStmt, ColNum, SQL_DESC_SCHEMA_NAME, @wszCharAttr, 255, @CharAttrLen, NULL), "COdbcStmt.ColSchemaName", SQL_HANDLE_STMT, m_hStmt)
   RETURN wszCharAttr
END FUNCTION
' =====================================================================================

' =====================================================================================
' SQL_PRED_NONE if the column cannot be used in a WHERE clause. (This is the same as the
' SQL_UNSEARCHABLE value in ODBC 2.x.)
' SQL_PRED_CHAR if the column can be used in a WHERE clause but only with the LIKE
' predicate. (This is the same as the SQL_LIKE_ONLY value in ODBC 2.x.)
' SQL_PRED_BASIC if the column can be used in a WHERE clause with all the comparison
' operators except LIKE. (This is the same as the SQL_EXCEPT_LIKE value in ODBC 2.x.)
' SQL_PRED_SEARCHABLE if the column can be used in a WHERE clause with any comparison
' operator.
' Columns of type SQL_LONGVARCHAR and SQL_LONGVARBINARY usually return SQL_PRED_CHAR.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.ColSearchable (BYVAL ColNum AS SQLUSMALLINT) AS SQLINTEGER
   DIM lAttr AS SQLINTEGER
   SetResult(SQLColAttributeW(m_hStmt, ColNum, SQL_DESC_SEARCHABLE, NULL, 0, NULL, cast(ANY PTR, @lAttr)), "COdbcStmt.ColSearchable", SQL_HANDLE_STMT, m_hStmt)
   RETURN lAttr
END FUNCTION
' =====================================================================================

' =====================================================================================
' The name of the table that contains the column. The returned value is implementation-
' defined if the column is an expression or if the column is part of a view.
' If the table name cannot be determined, an empty string is returned.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.ColTableName (BYVAL ColNum AS SQLUSMALLINT) AS CWSTR
   DIM wszCharAttr AS WSTRING * 255, CharAttrLen AS SQLSMALLINT
   SetResult(SQLColAttributeW(m_hStmt, ColNum, SQL_DESC_TABLE_NAME, @wszCharAttr, 255, @CharAttrLen, NULL), "COdbcStmt.ColTableName", SQL_HANDLE_STMT, m_hStmt)
   RETURN wszCharAttr
END FUNCTION
' =====================================================================================

' =====================================================================================
' A numeric value that specifies the SQL data type.
' When ColumnNumber is equal to 0, SQL_BINARY is returned for variable-length bookmarks
' and SQL_INTEGER is returned for fixed-length bookmarks.
' For the datetime and interval data types, this field returns the verbose data type:
' SQL_DATETIME or SQL_INTERVAL.
' This information is returned from the SQL_DESC_TYPE record field of the IRD.
' Note: To work against ODBC 2.x drivers, use SQL_DESC_CONCISE_TYPE instead.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.ColType (BYVAL ColNum AS SQLUSMALLINT) AS SQLINTEGER
   DIM lAttr AS SQLINTEGER
   SetResult(SQLColAttributeW(m_hStmt, ColNum, SQL_DESC_TYPE, NULL, 0, NULL, cast(ANY PTR, @lAttr)), "COdbcStmt.ColType", SQL_HANDLE_STMT, m_hStmt)
   RETURN lAttr
END FUNCTION
' =====================================================================================

' =====================================================================================
' Data source-dependent data type name; for example, "CHAR", "VARCHAR", "MONEY",
' "LONG VARBINARY", or "CHAR ( ) FOR BIT DATA".
' If the type is unknown, an empty string is returned.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.ColTypeName (BYVAL ColNum AS SQLUSMALLINT) AS CWSTR
   DIM wszCharAttr AS WSTRING * 255, CharAttrLen AS SQLSMALLINT
   SetResult(SQLColAttributeW(m_hStmt, ColNum, SQL_DESC_TYPE_NAME, @wszCharAttr, 255, @CharAttrLen, NULL), "COdbcStmt.ColTypeName", SQL_HANDLE_STMT, m_hStmt)
   RETURN wszCharAttr
END FUNCTION
' =====================================================================================

' =====================================================================================
' SQL_NAMED or SQL_UNNAMED. If the SQL_DESC_NAME field of the IRD contains a column
' alias or a column name, SQL_NAMED is returned. If there is no column name or column
' alias, SQL_UNNAMED is returned.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.ColUnnamed (BYVAL ColNum AS SQLUSMALLINT) AS SQLINTEGER
   DIM lAttr AS SQLINTEGER
   SetResult(SQLColAttributeW(m_hStmt, ColNum, SQL_DESC_UNNAMED, NULL, 0, NULL, cast(ANY PTR, @lAttr)), "COdbcStmt.ColUnnamed", SQL_HANDLE_STMT, m_hStmt)
   RETURN lAttr
END FUNCTION
' =====================================================================================

' =====================================================================================
' SQL_TRUE if the column is unsigned (or not numeric).
' SQL_FALSE if the column is signed.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.ColUnsigned (BYVAL ColNum AS SQLUSMALLINT) AS BOOLEAN
   DIM lAttr AS SQLINTEGER
   SetResult(SQLColAttributeW(m_hStmt, ColNum, SQL_DESC_UNSIGNED, NULL, 0, NULL, cast(ANY PTR, @lAttr)), "COdbcStmt.ColUnsigned", SQL_HANDLE_STMT, m_hStmt)
   RETURN lAttr
END FUNCTION
' =====================================================================================

' =====================================================================================
' SQL_TRUE if the column is updatable; SQL_FALSE otherwise.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.ColUpdatable (BYVAL ColNum AS SQLUSMALLINT) AS BOOLEAN
   DIM lAttr AS SQLINTEGER
   SetResult(SQLColAttributeW(m_hStmt, ColNum, SQL_DESC_UPDATABLE, NULL, 0, NULL, cast(ANY PTR, @lAttr)), "COdbcStmt.ColUpdatable", SQL_HANDLE_STMT, m_hStmt)
   RETURN lAttr
END FUNCTION
' =====================================================================================

' #####################################################################################
' STATEMENT ATTRIBUTES
' #####################################################################################

' =====================================================================================
' Returns the current setting of a statement attribute.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.GetStmtAttr (BYVAL Attribute AS SQLINTEGER, BYVAL ValuePtr AS SQLPOINTER, _
BYVAL BufferLength AS SQLINTEGER, BYVAL StringLength AS SQLINTEGER PTR) AS SQLRETURN
   RETURN SetResult(SQLGetStmtAttrW(m_hStmt, Attribute, ValuePtr, BufferLength, StringLength), "COdbcStmt.GetStmtAttr", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Sets the current setting of a statement attribute.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.SetStmtAttr (BYVAL Attribute AS SQLINTEGER, BYVAL ValuePtr AS SQLPOINTER, BYVAL StringLength AS SQLINTEGER) AS SQLRETURN
   RETURN SetResult(SQLSetStmtAttrW(m_hStmt, Attribute, ValuePtr, StringLength), "COdbcStmt.SetStmtAttr", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Sets or returns an integer value that specifies the cursor concurrency.
'   SQL_CONCUR_READ_ONLY = Cursor is read-only. No updates are allowed.
'   SQL_CONCUR_LOCK = Cursor uses the lowest level of locking sufficient to ensure that
'      the row can be updated.
'   SQL_CONCUR_ROWVER = Cursor uses optimistic concurrency control, comparing row
'      versions such as SQLBase ROWID or Sybase TIMESTAMP.
'   SQL_CONCUR_VALUES = Cursor uses optimistic concurrency control, comparing values.
' The default attribute is SQL_CONCUR_READ_ONLY. This attribute cannot be specified for
' an open cursor.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.GetCursorConcurrency () AS SQLUINTEGER
   DIM dwCursorConcurrency AS SQLUINTEGER
   SetResult(SQLGetStmtAttrW(m_hStmt, SQL_ATTR_CONCURRENCY, @dwCursorConcurrency, SQL_IS_UINTEGER, NULL), "COdbcStmt.GetCursorConcurrency", SQL_HANDLE_STMT, m_hStmt)
   RETURN dwCursorConcurrency
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.SetCursorConcurrency (BYVAL LockType AS SQLUINTEGER) AS SQLRETURN
   RETURN SetResult(SQLSetStmtAttrW(m_hStmt, SQL_ATTR_CONCURRENCY, cast(SQLPOINTER, cast(ULONG_PTR, LockType)), SQL_IS_UINTEGER), "COdbcStmt.SetCursorConcurrency", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.SetOptimisticConcurrency () AS SQLRETURN
   RETURN SetResult(SQLSetStmtAttrW(m_hStmt, SQL_ATTR_CONCURRENCY, cast(SQLPOINTER, SQL_CONCUR_VALUES), SQL_IS_UINTEGER), "COdbcStmt.SetOptimisticConcurrency", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.SetLockConcurrency () AS SQLRETURN
   RETURN SetResult(SQLSetStmtAttrW(m_hStmt, SQL_ATTR_CONCURRENCY, cast(SQLPOINTER, SQL_CONCUR_LOCK), SQL_IS_UINTEGER), "COdbcStmt.SetLockConcurrency", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.SetReadOnlyConcurrency () AS SQLRETURN
   RETURN SetResult(SQLSetStmtAttrW(m_hStmt, SQL_ATTR_CONCURRENCY, cast(SQLPOINTER, SQL_CONCUR_READ_ONLY), SQL_IS_UINTEGER), "COdbcStmt.SetReadOnlyConcurrency", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.SetRowVerConcurrency () AS SQLRETURN
   RETURN SetResult(SQLSetStmtAttrW(m_hStmt, SQL_ATTR_CONCURRENCY, cast(SQLPOINTER, SQL_CONCUR_ROWVER), SQL_IS_UINTEGER), "COdbcStmt.SetRowVerConcurrency", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Sets or returns a LONG value that specifies the cursor type:
'   SQL_CURSOR_FORWARD_ONLY = The cursor only scrolls forward.
'   SQL_CURSOR_STATIC = The data in the result set is static.
'   SQL_CURSOR_KEYSET_DRIVEN = The driver saves and uses only the keys for the number
'     of rows specified in the SQL_ATTR_KEYSET_SIZE statement attribute.
'   SQL_CURSOR_DYNAMIC = The driver saves and uses only the keys for the rows in the
'     rowset.
' The default value is SQL_FORWARD_ONLY. This attribute cannot be specified after the
'     SQL statement has been prepared.
' The application can specify the cursor type before executing a statement that creates a
' result set. It does this with the SQL_ATTR_CURSOR_TYPE statement attribute. If the
' application does not explicitly specify a type, a forward-only cursor will be used.
' To get a mixed cursor, an application specifies a keyset-driven cursor but declares a
' keyset size less than the result set size.
' Note: Microsoft Access Driver changes SQL_CURSOR_DYNAMIC to SQL_CURSOR_KEYSET_DRIVEN.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.GetCursorType () AS SQLUINTEGER
   DIM CursorType AS SQLUINTEGER
   SetResult(SQLGetStmtAttrW(m_hStmt, SQL_ATTR_CURSOR_TYPE, @CursorType, SQL_IS_UINTEGER, NULL), "COdbcStmt.GetCursorType", SQL_HANDLE_STMT, m_hStmt)
   RETURN CursorType
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.SetCursorType (BYVAL CursorType AS DWORD) AS SQLRETURN
   RETURN SetResult(SQLSetStmtAttrW(m_hStmt, SQL_ATTR_CURSOR_TYPE, cast(SQLPOINTER, cast(ULONG_PTR, CursorType)), SQL_IS_UINTEGER), "COdbcStmt.SetCursorType", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.SetForwardOnlyCursor () AS SQLRETURN
   RETURN SetResult(SQLSetStmtAttrW(m_hStmt, SQL_ATTR_CURSOR_TYPE, cast(SQLPOINTER, SQL_CURSOR_FORWARD_ONLY), SQL_IS_UINTEGER), "COdbcStmt.SetForwardOnlyCursor", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.SetKeysetDrivenCursor () AS SQLRETURN
   RETURN SetResult(SQLSetStmtAttrW(m_hStmt, SQL_ATTR_CURSOR_TYPE, cast(SQLPOINTER, SQL_CURSOR_KEYSET_DRIVEN), SQL_IS_UINTEGER), "COdbcStmt.SetKeysetDrivenCursor", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.SetStaticCursor () AS SQLRETURN
   RETURN SetResult(SQLSetStmtAttrW(m_hStmt, SQL_ATTR_CURSOR_TYPE, cast(SQLPOINTER, SQL_CURSOR_STATIC), SQL_IS_UINTEGER), "COdbcStmt.SetStaticCursor", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.SetDynamicCursor () AS SQLRETURN
   RETURN SetResult(SQLSetStmtAttrW(m_hStmt, SQL_ATTR_CURSOR_TYPE, cast(SQLPOINTER, SQL_CURSOR_DYNAMIC), SQL_IS_UINTEGER), "COdbcStmt.SetDynamicCursor", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Creates a multiuser keyset cursor.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.SetMultiuserKeysetCursor (BYVAL pwszCursorName AS WSTRING PTR = NULL) AS SQLRETURN
   m_Result = SetResult(SQLSetStmtAttrW(m_hStmt, SQL_ATTR_CURSOR_TYPE, cast(SQLPOINTER, SQL_CURSOR_KEYSET_DRIVEN), NULL), "COdbcStmt.SetMultiuserKeysetCursor", SQL_HANDLE_STMT, m_hStmt)
   IF SQL_SUCCEEDED(m_Result) THEN
      m_Result = SetResult(SQLSetStmtAttrW(m_hStmt, SQL_ATTR_CONCURRENCY, cast(SQLPOINTER, SQL_CONCUR_VALUES), NULL), "COdbcStmt.SetMultiuserKeysetCursor", SQL_HANDLE_STMT, m_hStmt)
      IF SQL_SUCCEEDED(m_Result) THEN
         IF pwszCursorName THEN SetResult(SQLSetCursorNameW(m_hStmt, pwszCursorName, SQL_NTS), "COdbcStmt.SetMultiuserKeysetCursor", SQL_HANDLE_STMT, m_hStmt)
      END IF
   END IF
   RETURN m_Result
END FUNCTION
' =====================================================================================

' =====================================================================================
' Sets or returns LONG value that specifies the number of rows in the keyset-driven cursor.
' If the keyset size is 0 (the default), the cursor is fully keyset-driven.
' If the keyset size is greater than 0, the cursor is mixed (keyset-driven within the
' keyset and dynamic outside of the keyset). The default keyset size is 0. Fetch or
' FetchScroll returns an error if the keyset size is greater than 0 and less than the
' rowset size.
' For keyset-driven and mixed cursors, the application can specify the keyset size. It
' does this with the SQL_ATTR_KEYSET_SIZE statement attribute. If the keyset size is set
' to 0, which is the default, the keyset size is set to the result size and a
' keyset-driven cursor is used. The keyset size can be changed after the cursor has been
' opened.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.GetCursorKeysetSize () AS SQLUINTEGER
   DIM KeysetSize AS SQLUINTEGER
   SetResult(SQLGetStmtAttrW(m_hStmt, SQL_ATTR_KEYSET_SIZE, @KeysetSize, SQL_IS_UINTEGER, NULL), "COdbcStmt.GetCursorKeysetSize", SQL_HANDLE_STMT, m_hStmt)
   RETURN KeysetSize
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.SetCursorKeysetSize (BYVAL KeysetSize AS SQLUINTEGER) AS SQLRETURN
   RETURN SetResult(SQLSetStmtAttrW(m_hStmt, SQL_ATTR_KEYSET_SIZE, cast(SQLPOINTER, cast(ULONG_PTR, KeysetSize)), SQL_IS_UINTEGER), "COdbcStmt.SetCursorKeysetSize", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Sets or returns a LONG value that specifies the level of support that the application
' requires. Setting this attribute affects subsequent calls to ExecDirect and Execute.
'    SQL_NONSCROLLABLE = Scrollable cursors are not required on the statement handle. If
'      the application calls FetchScroll on this handle, the only value of
'      FetchOrientation is SQL_FETCH_NEXT. This is the default.
'    SQL_SCROLLABLE = Scrollable cursors are required on the statement handle. When
'      calling FetchScroll the application may specify any valid value of
'      FetchOrientation, achieving cursor positioning in modes other than the sequential
'      mode.
' Note: Optional function not implemented in Microsoft Access Driver.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.GetCursorScrollability () AS SQLUINTEGER
   DIM CursorScrollability AS SQLUINTEGER
   SetResult(SQLGetStmtAttrW(m_hStmt, SQL_ATTR_CURSOR_SCROLLABLE, @CursorScrollability, SQL_IS_UINTEGER, NULL), "COdbcStmt.GetCursorScrollability", SQL_HANDLE_STMT, m_hStmt)
   RETURN CursorScrollability
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.SetCursorScrollability (BYVAL CursorScrollability AS SQLUINTEGER) AS SQLRETURN
   RETURN SetResult(SQLSetStmtAttrW(m_hStmt, SQL_ATTR_CURSOR_SCROLLABLE, cast(SQLPOINTER, cast(ULONG_PTR, CursorScrollability)), SQL_IS_UINTEGER), "COdbcStmt.SetCursorScrollability", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Sets or returns a LONG value that specifies whether cursors on the statement handle
' made to a result set by another cursor. Setting this attribute affects subsequent
' calls to ExecDirect and Execute. An application can read back the value of this
' attribute to obtain its initial state or its state as most recently set by the application.
'   SQL_UNSPECIFIED = It is unspecified what the cursor type is and whether cursors on
'      the statement handle make visible the changes made to a result set by another
'      cursor. Cursors on the statement handle may make visible none, some, or all such
'      changes. This is the default.
'   SQL_INSENSITIVE = All cursors on the statement handle show the result set without
'      reflecting any changes made to it by any other cursor, which has a concurrency that
'      is read-only.
'   SQL_SENSITIVE = All cursors on the statement handle make visible all changes made to
'      a result set by another cursor.
' Note: Optional functions not implemented in Microsoft Access Driver.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.GetCursorSensitivity () AS SQLUINTEGER
   DIM CursorSensitivity AS SQLUINTEGER
   SetResult(SQLGetStmtAttrW(m_hStmt, SQL_ATTR_CURSOR_SENSITIVITY, @CursorSensitivity, SQL_IS_UINTEGER, NULL), "COdbcStmt.GetCursorSensitivity", SQL_HANDLE_STMT, m_hStmt)
   RETURN CursorSensitivity
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.SetCursorSensitivity (BYVAL CursorSensitivity AS SQLUINTEGER) AS SQLRETURN
   RETURN SetResult(SQLSetStmtAttrW(m_hStmt, SQL_ATTR_CURSOR_SENSITIVITY, cast(SQLPOINTER, cast(ULONG_PTR, CursorSensitivity)), SQL_IS_UINTEGER), "COdbcStmt.SetCursorSensitivity", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' The handle to the APD for subsequent calls to Execute and ExecDirect on the
' statement handle. The initial value of this attribute is the descriptor implicitly
' allocated when the statement was initially allocated.
' If the value of this attribute is set to SQL_NULL_DESC or the handle originally
' allocated for the descriptor, an explicitly allocated APD handle that was previously
' associated with the statement handle is dissociated from it and the statement handle
' reverts to the implicitly allocated APD handle.
' This attribute cannot be set to a descriptor handle that was implicitly allocated for
' another statement or to another descriptor handle that was implicitly set on the same
' statement; implicitly allocated descriptor handles cannot be associated with more than
' one statement or descriptor handle.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.GetStmtAppParamDesc () AS SQLUINTEGER
   DIM dwAttr AS SQLUINTEGER
   SetResult(SQLGetStmtAttrW(m_hStmt, SQL_ATTR_APP_PARAM_DESC, @dwAttr, SQL_IS_UINTEGER, NULL), "COdbcStmt.GetStmtAppParamDesc", SQL_HANDLE_STMT, m_hStmt)
   RETURN dwAttr
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.SetStmtAppParamDesc (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
   RETURN SetResult(SQLSetStmtAttrW(m_hStmt, SQL_ATTR_APP_PARAM_DESC, cast(SQLPOINTER, cast(ULONG_PTR, dwAttr)), SQL_IS_UINTEGER), "COdbcStmt.SetStmtAppParamDesc", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' The handle to the ARD for subsequent fetches on the statement handle. The initial value
' of this attribute is the descriptor implicitly allocated when the statement was
' initially allocated. If the value of this attribute is set to SQL_NULL_DESC or the
' handle originally allocated for the descriptor, an explicitly allocated ARD handle that
' was previously associated with the statement handle is dissociated from it and the
' statement handle reverts to the implicitly allocated ARD handle.
' This attribute cannot be set to a descriptor handle that was implicitly allocated for
' another statement or to another descriptor handle that was implicitly set on the same
' statement; implicitly allocated descriptor handles cannot be associated with more than
' one statement or descriptor handle.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.GetStmtAppRowDesc () AS SQLUINTEGER
   DIM dwAttr AS SQLUINTEGER
   SetResult(SQLGetStmtAttrW(m_hStmt, SQL_ATTR_APP_ROW_DESC, @dwAttr, SQL_IS_UINTEGER, NULL), "COdbcStmt.GetStmtAppRowDesc", SQL_HANDLE_STMT, m_hStmt)
   RETURN dwAttr
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.SetStmtAppRowDesc (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
   RETURN SetResult(SQLSetStmtAttrW(m_hStmt, SQL_ATTR_APP_ROW_DESC, cast(SQLPOINTER, cast(ULONG_PTR, dwAttr)), SQL_IS_UINTEGER), "COdbcStmt.SetStmtAppRowDesc", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' An SQLUINTEGER value that specifies whether a function called with the specified
' statement is executed asynchronously:
'   SQL_ASYNC_ENABLE_OFF = Off (the default)
'   SQL_ASYNC_ENABLE_ON = On
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.GetStmtAsyncEnable () AS SQLUINTEGER
   DIM dwAttr AS SQLUINTEGER
   SetResult(SQLGetStmtAttrW(m_hStmt, SQL_ATTR_ASYNC_ENABLE, @dwAttr, SQL_IS_UINTEGER, NULL), "COdbcStmt.GetStmtAsyncEnable", SQL_HANDLE_STMT, m_hStmt)
   RETURN dwAttr
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.SetStmtAsyncEnable (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
   RETURN SetResult(SQLSetStmtAttrW(m_hStmt, SQL_ATTR_ASYNC_ENABLE, cast(SQLPOINTER, cast(ULONG_PTR, dwAttr)), SQL_IS_UINTEGER), "COdbcStmt.SetStmtAsyncEnable", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' A pointer that points to a binary bookmark value. When FetchScroll is called with
' fFetchOrientation equal to SQL_FETCH_BOOKMARK, the driver picks up the bookmark value
' from this field. This field defaults to a null pointer.
' The value pointed to by this field is not used for delete by bookmark, update by
' bookmark, or fetch by bookmark operations in BulkOperations, which use bookmarks
' cached in rowset buffers.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.GetStmtFetchBookmarkPtr () AS SQLUINTEGER
   DIM dwAttr AS SQLUINTEGER
   SetResult(SQLGetStmtAttrW(m_hStmt, SQL_ATTR_FETCH_BOOKMARK_PTR, @dwAttr, SQL_IS_UINTEGER, NULL), "COdbcStmt.GetStmtFetchBookmarkPtr", SQL_HANDLE_STMT, m_hStmt)
   RETURN dwAttr
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.SetStmtFetchBookmarkPtr (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
   RETURN SetResult(SQLSetStmtAttrW(m_hStmt, SQL_ATTR_FETCH_BOOKMARK_PTR, cast(SQLPOINTER, cast(ULONG_PTR, dwAttr)), SQL_IS_UINTEGER), "COdbcStmt.SetStmtFetchBookmarkPtr", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' The handle to the IPD. The value of this attribute is the descriptor allocated when the
' statement was initially allocated. The application cannot set this attribute.
' This attribute can be retrieved by a call to GetStmtAttr but not set by a call to
' SetStmtAttr.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.GetImpParamDesc () AS SQLUINTEGER
   DIM dwAttr AS SQLUINTEGER
   SetResult(SQLGetStmtAttrW(m_hStmt, SQL_ATTR_IMP_PARAM_DESC, @dwAttr, SQL_IS_UINTEGER, NULL), "COdbcStmt.GetStmtImpParamDesc", SQL_HANDLE_STMT, m_hStmt)
   RETURN dwAttr
END FUNCTION
' =====================================================================================

' =====================================================================================
' ImpParamDescRec (Implicitily Parameter Descriptor).
' Returns the current settings or values of multiple fields of a descriptor record. The
' fields returned describe the name, data type, and storage of column or parameter data.
' - m_hStmt : Statement handle
' - RecNumber : Indicates the descriptor record from which the application seeks information.
' Descriptor records are numbered from 1, with record number 0 being the bookmark record.
' The RecNumber argument must be less or equal to the value of SQL_DESC_COUNT. If RecNumber
' is less that or equal to SQL_DESC_COUNT but the row does not contain data for a column or
' parameter, a call to GetImpRowDesc will return the default values of the fields.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.GetImpParamDescRec (BYVAL RecNumber AS SQLSMALLINT, BYVAL pwszName AS WSTRING PTR, _
BYVAL BufferLength AS SQLSMALLINT, BYVAL StringLength AS SQLSMALLINT PTR, BYVAL TypePtr AS SQLSMALLINT PTR, _
BYVAL SubTypePtr AS SQLSMALLINT PTR, BYVAL LengthPtr AS SQLLEN PTR, BYVAL PrecisionPtr AS SQLSMALLINT PTR, _
BYVAL ScalePtr AS SQLSMALLINT PTR, BYVAL NullablePtr AS SQLSMALLINT PTR) AS SQLRETURN
   ' Gets the handle of the IPD (implicitily allocated descriptor).
   DIM hIpd AS SQLUINTEGER
   m_Result = SetResult(SQLGetStmtAttrW(m_hStmt, SQL_ATTR_IMP_PARAM_DESC, @hIpd, SQL_IS_UINTEGER, NULL), "COdbcStmt.GetImpParamDescRec", SQL_HANDLE_STMT, m_hStmt)
   IF SQL_SUCCEEDED(m_Result) THEN
      m_Result = SetResult(SQLGetDescRecW(cast(SQLHDESC, cast(LONG_PTR, hIpd)), RecNumber, pwszName, BufferLength, StringLength, TypePtr, _
                 SubTypePtr, LengthPtr, PrecisionPtr, ScalePtr, ScalePtr), "COdbcStmt.GetImpParamDescRec", SQL_HANDLE_STMT, m_hStmt)
   END IF
   RETURN m_Result
END FUNCTION
' =====================================================================================
' ========================================================================================
PRIVATE FUNCTION COdbcStmt.GetImpParamDescField (BYVAL RecNumber AS SQLSMALLINT, BYVAL FieldIdentifier AS SQLSMALLINT, _
BYVAL ValuePtr AS SQLPOINTER, BYVAL BufferLength AS SQLINTEGER, BYVAL StringLength AS SQLINTEGER PTR) AS SQLRETURN
   DIM hIpd AS SQLUINTEGER
   ' Gets the handle of the IPD (implicitily allocated descriptor).
   m_Result = SetResult(SQLGetStmtAttrW(m_hStmt, SQL_ATTR_IMP_PARAM_DESC, @hIpd, SQL_IS_UINTEGER, NULL), "COdbcStmt.GetImpParamDescField", SQL_HANDLE_STMT, m_hStmt)
   IF SQL_SUCCEEDED(m_Result) THEN
      m_Result = SetResult(SQLGetDescFieldW(cast(SQLHDESC, cast(LONG_PTR, hIpd)), RecNumber, FieldIdentifier, ValuePtr, BufferLength, StringLength), "COdbcStmt.GetImpParamDescField", SQL_HANDLE_STMT, m_hStmt)
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION COdbcStmt.GetImpParamDescFieldName (BYVAL RecNumber AS SQLSMALLINT) AS CWSTR
   DIM hIpd AS SQLUINTEGER, wszName AS WSTRING * 256, StringLength AS SQLINTEGER
   ' Gets the handle of the IPD (implicitily allocated descriptor).
   m_Result = SetResult(SQLGetStmtAttrW(m_hStmt, SQL_ATTR_IMP_PARAM_DESC, @hIpd, SQL_IS_UINTEGER, NULL), "COdbcStmt.GetImpParamDescFieldName", SQL_HANDLE_STMT, m_hStmt)
   IF hIpd THEN
      m_Result = SetResult(SQLGetDescFieldW(cast(SQLHDESC, cast(LONG_PTR, hIpd)), RecNumber, SQL_DESC_NAME, @wszName, 256, @StringLength), "COdbcStmt.GetImpParamDescFieldName", SQL_HANDLE_STMT, m_hStmt)
      IF SQL_SUCCEEDED(m_Result) THEN RETURN wszName
   END IF
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION COdbcStmt.GetImpParamDescFieldType (BYVAL RecNumber AS SQLSMALLINT) AS SQLSMALLINT
   DIM hIpd AS SQLUINTEGER, nType AS SQLUINTEGER, StringLength AS SQLINTEGER
   ' Gets the handle of the IPD (implicitily allocated descriptor).
   m_Result = SetResult(SQLGetStmtAttrW(m_hStmt, SQL_ATTR_IMP_PARAM_DESC, @hIpd, SQL_IS_UINTEGER, NULL), "COdbcStmt.GetImpParamDescFieldType", SQL_HANDLE_STMT, m_hStmt)
   IF hIpd = NULL THEN RETURN 0
   SetResult(SQLGetDescFieldW(cast(SQLHDESC, cast(LONG_PTR, hIpd)), RecNumber, SQL_DESC_TYPE, @nType, SQL_IS_SMALLINT, @StringLength), "COdbcStmt.GetImpParamDescFieldType", SQL_HANDLE_STMT, m_hStmt)
   RETURN nType
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION COdbcStmt.GetImpParamDescFieldOctetLength (BYVAL RecNumber AS SQLSMALLINT) AS SQLSMALLINT
   DIM hIpd AS SQLUINTEGER, nOctetLength AS SQLSMALLINT, StringLength AS SQLINTEGER
   ' Gets the handle of the IPD (implicitily allocated descriptor).
   m_Result = SetResult(SQLGetStmtAttrW(m_hStmt, SQL_ATTR_IMP_PARAM_DESC, @hIpd, SQL_IS_UINTEGER, NULL), "COdbcStmt.GetImpParamDescFieldOctetLength", SQL_HANDLE_STMT, m_hStmt)
   IF hIpd = NULL THEN RETURN 0
   SetResult(SQLGetDescFieldW(cast(SQLHDESC, cast(LONG_PTR, hIpd)), RecNumber, SQL_DESC_OCTET_LENGTH, @nOctetLength, SQL_IS_SMALLINT, @StringLength), "COdbcStmt.GetImpParamDescFieldOctetLength", SQL_HANDLE_STMT, m_hStmt)
   RETURN nOctetLength
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION COdbcStmt.GetImpParamDescFieldPrecision (BYVAL RecNumber AS SQLSMALLINT) AS SQLSMALLINT
   DIM hIpd AS SQLUINTEGER, nPrecision AS SQLSMALLINT, StringLength AS SQLINTEGER
   ' Gets the handle of the IPD (implicitily allocated descriptor).
   m_Result = SetResult(SQLGetStmtAttrW(m_hStmt, SQL_ATTR_IMP_PARAM_DESC, @hIpd, SQL_IS_UINTEGER, NULL), "COdbcStmt.GetImpParamDescFieldPrecision", SQL_HANDLE_STMT, m_hStmt)
   IF hIpd = NULL THEN RETURN 0
   SetResult(SQLGetDescFieldW(cast(SQLHDESC, cast(LONG_PTR, hIpd)), RecNumber, SQL_DESC_PRECISION, @nPrecision, SQL_IS_SMALLINT, @StringLength), "COdbcStmt.GetImpParamDescFieldPrecision", SQL_HANDLE_STMT, m_hStmt)
   RETURN nPrecision
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION COdbcStmt.GetImpParamDescFieldScale (BYVAL RecNumber AS SQLSMALLINT) AS SQLSMALLINT
   DIM hIpd AS SQLUINTEGER, nScale AS SQLSMALLINT, StringLength AS SQLINTEGER
   ' Gets the handle of the IPD (implicitily allocated descriptor).
   m_Result = SetResult(SQLGetStmtAttrW(m_hStmt, SQL_ATTR_IMP_PARAM_DESC, @hIpd, SQL_IS_UINTEGER, NULL), "COdbcStmt.GetImpParamDescFieldScale", SQL_HANDLE_STMT, m_hStmt)
   IF hIpd = NULL THEN RETURN 0
   SetResult(SQLGetDescFieldW(cast(SQLHDESC, cast(LONG_PTR, hIpd)), RecNumber, SQL_DESC_SCALE, @nScale, SQL_IS_SMALLINT, @StringLength), "COdbcStmt.GetImpParamDescFieldScale", SQL_HANDLE_STMT, m_hStmt)
   RETURN nScale
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION COdbcStmt.GetImpParamDescFieldNullable (BYVAL RecNumber AS SQLSMALLINT) AS SQLSMALLINT
   DIM hIpd AS SQLUINTEGER, nNullable AS SQLSMALLINT, StringLength AS SQLINTEGER
   ' Gets the handle of the IPD (implicitily allocated descriptor).
   m_Result = SetResult(SQLGetStmtAttrW(m_hStmt, SQL_ATTR_IMP_PARAM_DESC, @hIpd, SQL_IS_UINTEGER, NULL), "COdbcStmt.GetImpParamDescFieldNullable", SQL_HANDLE_STMT, m_hStmt)
   IF hIpd = NULL THEN RETURN 0
   SetResult(SQLGetDescFieldW(cast(SQLHDESC, cast(LONG_PTR, hIpd)), RecNumber, SQL_DESC_NULLABLE, @nNullable, SQL_IS_SMALLINT, @StringLength), "COdbcStmt.GetImpParamDescFieldNullable", SQL_HANDLE_STMT, m_hStmt)
   RETURN nNullable
END FUNCTION
' ========================================================================================

' ========================================================================================
' The handle to the IRD. The value of this attribute is the descriptor allocated when the
' statement was initially allocated. The application cannot set this attribute.
' This attribute can be retrieved by a call to GetStmtAttr but not set by a call to
' SetStmtAttr.
' ========================================================================================
PRIVATE FUNCTION COdbcStmt.GetImpRowDesc () AS SQLUINTEGER
   DIM dwAttr AS SQLUINTEGER
   SetResult(SQLGetStmtAttrW(m_hStmt, SQL_ATTR_IMP_ROW_DESC, @dwAttr, SQL_IS_UINTEGER, NULL), "COdbcStmt.GetStmtImpRowDesc", SQL_HANDLE_STMT, m_hStmt)
   RETURN dwAttr
END FUNCTION
' ========================================================================================

' =====================================================================================
' ImpRowDescRec (Implicitily Row Descriptor).
' Returns the current settings or values of multiple fields of a descriptor record. The fields
' returned describe the name, data type, and storage of column or parameter data.
' m_hStmt : Statement handle
' RecNumber : Indicates the descriptor record from which the application seeks information.
' Descriptor records are numbered from 1, with record number 0 being the bookmark record.
' The RecNumber argument must be less or equal to the value of SQL_DESC_COUNT. If RecNumber
' is less that or equal to SQL_DESC_COUNT but the row does not contain data for a column or
' parameter, a call to GetImpRowDesc will return the default values of the fields.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.GetImpRowDescRec (BYVAL RecNumber AS SQLSMALLINT, BYVAL pwszName AS WSTRING PTR, _
BYVAL BufferLength AS SQLSMALLINT, BYVAL StringLength AS SQLSMALLINT PTR, BYVAL TypePtr AS SQLSMALLINT PTR, _
BYVAL SubTypePtr AS SQLSMALLINT PTR, BYVAL LengthPtr AS SQLLEN PTR, BYVAL PrecisionPtr AS SQLSMALLINT PTR, _
BYVAL ScalePtr AS SQLSMALLINT PTR, BYVAL NullablePtr AS SQLSMALLINT PTR) AS SQLRETURN
   ' Gets the handle of the IPD (implicitily allocated descriptor).
   DIM hIpd AS SQLUINTEGER
   m_Result = SetResult(SQLGetStmtAttrW(m_hStmt, SQL_ATTR_IMP_PARAM_DESC, @hIpd, SQL_IS_UINTEGER, NULL), "COdbcStmt.GetImpRowDescRec", SQL_HANDLE_STMT, m_hStmt)
   IF SQL_SUCCEEDED(m_Result) THEN
      m_Result = SetResult(SQLGetDescRecW(cast(SQLHDESC, cast(LONG_PTR, hIpd)), RecNumber, pwszName, BufferLength, StringLength, TypePtr, _
                 SubTypePtr, LengthPtr, PrecisionPtr, ScalePtr, ScalePtr), "COdbcStmt.GetImpRowDescRec", SQL_HANDLE_STMT, m_hStmt)
   END IF
   RETURN m_Result
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.GetImpRowDescField (BYVAL RecNumber AS SQLSMALLINT, BYVAL FieldIdentifier AS SQLSMALLINT, _
BYVAL ValuePtr AS SQLPOINTER, BYVAL BufferLength AS SQLINTEGER, BYVAL StringLength AS SQLINTEGER PTR) AS SQLRETURN
   DIM hIpd AS SQLUINTEGER
   ' Gets the handle of the IPD (implicitily allocated descriptor).
   m_Result = SetResult(SQLGetStmtAttrW(m_hStmt, SQL_ATTR_IMP_ROW_DESC, @hIpd, SQL_IS_UINTEGER, NULL), "COdbcStmt.GetImpRowDescField", SQL_HANDLE_STMT, m_hStmt)
   IF SQL_SUCCEEDED(m_Result) THEN
      m_Result = SetResult(SQLGetDescFieldW(cast(SQLHDESC, cast(LONG_PTR, hIpd)), RecNumber, FieldIdentifier, ValuePtr, BufferLength, StringLength), "COdbcStmt.GetImpRowDescField", SQL_HANDLE_STMT, m_hStmt)
   END IF
   RETURN m_Result
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.GetImpRowDescFieldName (BYVAL RecNumber AS SQLSMALLINT) AS CWSTR
   DIM hIpd AS SQLUINTEGER, wszName AS WSTRING * 256, StringLength AS SQLINTEGER
   ' Gets the handle of the IPD (implicitily allocated descriptor).
   m_Result = SetResult(SQLGetStmtAttrW(m_hStmt, SQL_ATTR_IMP_ROW_DESC, @hIpd, SQL_IS_UINTEGER, NULL), "COdbcStmt.GetImpRowDescFieldName", SQL_HANDLE_STMT, m_hStmt)
   IF hIpd THEN SetResult(SQLGetDescFieldW(cast(SQLHDESC, cast(LONG_PTR, hIpd)), RecNumber, SQL_DESC_NAME, @wszName, 256, @StringLength), "COdbcStmt.GetImpRowDescFieldName", SQL_HANDLE_STMT, m_hStmt)
   RETURN wszName
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.GetImpRowDescFieldType (BYVAL RecNumber AS SQLSMALLINT) AS SQLSMALLINT
   DIM hIpd AS SQLUINTEGER, nType AS SQLSMALLINT, StringLength AS SQLINTEGER
   ' Gets the handle of the IPD (implicitily allocated descriptor).
   m_Result = SetResult(SQLGetStmtAttrW(m_hStmt, SQL_ATTR_IMP_ROW_DESC,@ hIpd, SQL_IS_UINTEGER, NULL), "COdbcStmt.GetImpRowDescFieldType", SQL_HANDLE_STMT, m_hStmt)
   IF hIpd THEN SetResult(SQLGetDescFieldW(cast(SQLHDESC, cast(LONG_PTR, hIpd)), RecNumber, SQL_DESC_TYPE, @nType, SQL_IS_SMALLINT, @StringLength), "COdbcStmt.GetImpRowDescFieldType", SQL_HANDLE_STMT, m_hStmt)
   RETURN nType
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.GetImpRowDescFieldOctetLength (BYVAL RecNumber AS SQLSMALLINT) AS SQLSMALLINT
   DIM hIpd AS SQLUINTEGER, nOctetLength AS SQLSMALLINT, StringLength AS SQLINTEGER
   ' Gets the handle of the IPD (implicitily allocated descriptor).
   m_Result = SetResult(SQLGetStmtAttrW(m_hStmt, SQL_ATTR_IMP_ROW_DESC, @hIpd, SQL_IS_UINTEGER, NULL), "COdbcStmt.GetImpRowDescFieldOctetLength", SQL_HANDLE_STMT, m_hStmt)
   IF hIpd THEN SetResult(SQLGetDescFieldW(cast(SQLHDESC, cast(LONG_PTR, hIpd)), RecNumber, SQL_DESC_OCTET_LENGTH, @nOctetLength, SQL_IS_SMALLINT, @StringLength), "COdbcStmt.GetImpRowDescFieldOctetLength", SQL_HANDLE_STMT, m_hStmt)
   RETURN nOctetLength
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.GetImpRowDescFieldPrecision (BYVAL RecNumber AS SQLSMALLINT) AS SQLSMALLINT
   DIM hIpd AS SQLUINTEGER, nPrecision AS SQLSMALLINT, StringLength AS SQLINTEGER
   ' Gets the handle of the IPD (implicitily allocated descriptor).
   m_Result = SetResult(SQLGetStmtAttrW(m_hStmt, SQL_ATTR_IMP_ROW_DESC, @hIpd, SQL_IS_UINTEGER, NULL), "COdbcStmt.GetImpRowDescFieldPrecision", SQL_HANDLE_STMT, m_hStmt)
   IF hIpd THEN SetResult(SQLGetDescFieldW(cast(SQLHDESC, cast(LONG_PTR, hIpd)), RecNumber, SQL_DESC_PRECISION, @nPrecision, SQL_IS_SMALLINT, @StringLength), "COdbcStmt.GetImpRowDescFieldPrecision", SQL_HANDLE_STMT, m_hStmt)
   RETURN nPrecision
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.GetImpRowDescFieldScale (BYVAL RecNumber AS SQLSMALLINT) AS SQLSMALLINT
   DIM hIpd AS SQLUINTEGER, nScale AS SQLSMALLINT, StringLength AS SQLINTEGER
   ' Gets the handle of the IPD (implicitily allocated descriptor).
   m_Result = SetResult(SQLGetStmtAttrW(m_hStmt, SQL_ATTR_IMP_ROW_DESC, @hIpd, SQL_IS_UINTEGER, NULL), "COdbcStmt.GetImpRowDescFieldScale", SQL_HANDLE_STMT, m_hStmt)
   IF hIpd THEN SetResult(SQLGetDescFieldW(cast(SQLHDESC, cast(LONG_PTR, hIpd)), RecNumber, SQL_DESC_SCALE, @nScale, SQL_IS_SMALLINT, @StringLength), "COdbcStmt.GetImpRowDescFieldScale", SQL_HANDLE_STMT, m_hStmt)
   RETURN nScale
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.GetImpRowDescFieldNullable (BYVAL RecNumber AS SQLSMALLINT) AS SQLSMALLINT
   DIM hIpd AS SQLUINTEGER, nNullable AS SQLSMALLINT, StringLength AS SQLINTEGER
   ' Gets the handle of the IPD (implicitily allocated descriptor).
   m_Result = SetResult(SQLGetStmtAttrW(m_hStmt, SQL_ATTR_IMP_ROW_DESC, @hIpd, SQL_IS_UINTEGER, NULL), "COdbcStmt.GetImpRowDescFieldNullable", SQL_HANDLE_STMT, m_hStmt)
   IF hIpd THEN SetResult(SQLGetDescFieldW(cast(SQLHDESC, cast(LONG_PTR, hIpd)), RecNumber, SQL_DESC_NULLABLE, @nNullable, SQL_IS_SMALLINT, @StringLength), "COdbcStmt.GetImpRowDescFieldNullable", SQL_HANDLE_STMT, m_hStmt)
   RETURN nNullable
END FUNCTION
' =====================================================================================

' =====================================================================================
PRIVATE FUNCTION COdbcStmt.GetStmtMaxLength () AS SQLUINTEGER
   DIM dwAttr AS SQLUINTEGER
   SetResult(SQLGetStmtAttrW(m_hStmt, SQL_ATTR_MAX_LENGTH, @dwAttr, SQL_IS_UINTEGER, NULL), "COdbcStmt.GetStmtMaxLength", SQL_HANDLE_STMT, m_hStmt)
   RETURN dwAttr
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.SetStmtMaxLength (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
   RETURN SetResult(SQLSetStmtAttrW(m_hStmt, SQL_ATTR_MAX_LENGTH, cast(SQLPOINTER, cast(ULONG_PTR, dwAttr)), SQL_IS_UINTEGER), "COdbcStmt.SetStmtMaxLength", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
PRIVATE FUNCTION COdbcStmt.GetStmtMaxRows () AS SQLUINTEGER
   DIM dwAttr AS SQLUINTEGER
   SetResult(SQLGetStmtAttrW(m_hStmt, SQL_ATTR_MAX_ROWS, @dwAttr, SQL_IS_UINTEGER, NULL), "COdbcStmt.GetStmtMaxRows", SQL_HANDLE_STMT, m_hStmt)
   RETURN dwAttr
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.SetStmtMaxRows (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
   RETURN SetResult(SQLSetStmtAttrW(m_hStmt, SQL_ATTR_MAX_ROWS, cast(SQLPOINTER, cast(ULONG_PTR, dwAttr)), SQL_IS_UINTEGER), "COdbcStmt.SetStmtMaxRows", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
PRIVATE FUNCTION COdbcStmt.GetStmtNoScan () AS SQLUINTEGER
   DIM dwAttr AS SQLUINTEGER
   SetResult(SQLGetStmtAttrW(m_hStmt, SQL_ATTR_NOSCAN, @dwAttr, SQL_IS_UINTEGER, NULL), "COdbcStmt.GetStmtNoScan", SQL_HANDLE_STMT, m_hStmt)
   RETURN dwAttr
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.SetStmtNoScan (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
   RETURN SetResult(SQLSetStmtAttrW(m_hStmt, SQL_ATTR_NOSCAN, cast(SQLPOINTER, cast(ULONG_PTR, dwAttr)), SQL_IS_UINTEGER), "COdbcStmt.SetStmtNoScan", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
PRIVATE FUNCTION COdbcStmt.GetStmtParamBindOffsetPtr () AS SQLUINTEGER
   DIM dwAttr AS SQLUINTEGER
   SetResult(SQLGetStmtAttrW(m_hStmt, SQL_ATTR_PARAM_BIND_OFFSET_PTR, @dwAttr, SQL_IS_UINTEGER, NULL), "COdbcStmt.GetStmtParamBindOffsetPtr", SQL_HANDLE_STMT, m_hStmt)
   RETURN dwAttr
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.SetStmtParamBindOffsetPtr (BYVAL dwAttr AS DWORD) AS SQLRETURN
   RETURN SetResult(SQLSetStmtAttrW(m_hStmt, SQL_ATTR_PARAM_BIND_OFFSET_PTR, cast(SQLPOINTER, cast(ULONG_PTR, dwAttr)), SQL_IS_UINTEGER), "COdbcStmt.SetStmtParamBindOffsetPtr", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
PRIVATE FUNCTION COdbcStmt.GetStmtParamBindType () AS SQLUINTEGER
   DIM dwAttr AS SQLUINTEGER
   SetResult(SQLGetStmtAttrW(m_hStmt, SQL_ATTR_PARAM_BIND_TYPE, @dwAttr, SQL_IS_UINTEGER, NULL), "COdbcStmt.GetStmtParamBindType", SQL_HANDLE_STMT, m_hStmt)
   RETURN dwAttr
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.SetStmtParamBindType (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
   RETURN SetResult(SQLSetStmtAttrW(m_hStmt, SQL_ATTR_PARAM_BIND_TYPE, cast(SQLPOINTER, cast(ULONG_PTR, dwAttr)), SQL_IS_UINTEGER), "COdbcStmt.SetStmtParamBindType", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
PRIVATE FUNCTION COdbcStmt.GetStmtParamOperationPtr () AS SQLUINTEGER
   DIM dwAttr AS DWORD
   SetResult(SQLGetStmtAttrW(m_hStmt, SQL_ATTR_PARAM_OPERATION_PTR, @dwAttr, SQL_IS_UINTEGER, NULL), "COdbcStmt.GetStmtParamOperationPtr", SQL_HANDLE_STMT, m_hStmt)
   RETURN dwAttr
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.SetStmtParamOperationPtr (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
   RETURN SetResult(SQLSetStmtAttrW(m_hStmt, SQL_ATTR_PARAM_OPERATION_PTR, cast(SQLPOINTER, cast(ULONG_PTR, dwAttr)), SQL_IS_UINTEGER), "COdbcStmt.SetStmtParamOperationPtr", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
PRIVATE FUNCTION COdbcStmt.GetStmtParamsProcessedPtr () AS SQLUINTEGER
   DIM dwAttr AS SQLUINTEGER
   SetResult(SQLGetStmtAttrW(m_hStmt, SQL_ATTR_PARAMS_PROCESSED_PTR, @dwAttr, SQL_IS_UINTEGER, NULL), "COdbcStmt.GetStmtParamsProcessedPtr", SQL_HANDLE_STMT, m_hStmt)
   RETURN dwAttr
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.SetStmtParamsProcessedPtr (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
   RETURN SetResult(SQLSetStmtAttrW(m_hStmt, SQL_ATTR_PARAMS_PROCESSED_PTR, cast(SQLPOINTER, cast(ULONG_PTR, dwAttr)), SQL_IS_UINTEGER), "COdbcStmt.SetStmtParamsProcessedPtr", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
PRIVATE FUNCTION COdbcStmt.GetStmtParamsetSize () AS SQLUINTEGER
   DIM dwAttr AS SQLUINTEGER
   SetResult(SQLGetStmtAttrW(m_hStmt, SQL_ATTR_PARAMSET_SIZE, @dwAttr, SQL_IS_UINTEGER, NULL), "COdbcStmt.GetStmtParamsetSize", SQL_HANDLE_STMT, m_hStmt)
   RETURN dwAttr
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.SetStmtParamsetSize (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
   RETURN SetResult(SQLSetStmtAttrW(m_hStmt, SQL_ATTR_PARAMSET_SIZE, cast(SQLPOINTER, cast(ULONG_PTR, dwAttr)), SQL_IS_UINTEGER), "COdbcStmt.SetStmtParamsetSize", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
PRIVATE FUNCTION COdbcStmt.GetStmtQueryTimeout () AS SQLUINTEGER
   DIM dwAttr AS SQLUINTEGER
   SetResult(SQLGetStmtAttrW(m_hStmt, SQL_ATTR_QUERY_TIMEOUT, @dwAttr, SQL_IS_UINTEGER, NULL), "COdbcStmt.GetStmtQueryTimeout", SQL_HANDLE_STMT, m_hStmt)
   RETURN dwAttr
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.SetStmtQueryTimeout (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
   RETURN SetResult(SQLSetStmtAttrW(m_hStmt, SQL_ATTR_QUERY_TIMEOUT, cast(SQLPOINTER, cast(ULONG_PTR, dwAttr)), SQL_IS_UINTEGER), "COdbcStmt.SetStmtQueryTimeout", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
PRIVATE FUNCTION COdbcStmt.GetStmtRetrieveData () AS SQLUINTEGER
   DIM dwAttr AS SQLUINTEGER
   SetResult(SQLGetStmtAttrW(m_hStmt, SQL_ATTR_RETRIEVE_DATA, @dwAttr, SQL_IS_UINTEGER, NULL), "COdbcStmt.GetStmtRetrieveData", SQL_HANDLE_STMT, m_hStmt)
   RETURN dwAttr
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.SetStmtRetrieveData (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
   RETURN SetResult(SQLSetStmtAttrW(m_hStmt, SQL_ATTR_RETRIEVE_DATA, cast(SQLPOINTER, cast(ULONG_PTR, dwAttr)), SQL_IS_UINTEGER), "COdbcStmt.SetStmtRetrieveData", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
PRIVATE FUNCTION COdbcStmt.GetStmtParamStatusPtr () AS SQLUINTEGER
   DIM dwAttr AS SQLUINTEGER
   SetResult(SQLGetStmtAttrW(m_hStmt, SQL_ATTR_PARAM_STATUS_PTR, @dwAttr, SQL_IS_UINTEGER, NULL), "COdbcStmt.GetStmtParamStatusPtr", SQL_HANDLE_STMT, m_hStmt)
   RETURN dwAttr
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.SetStmtParamStatusPtr (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
   RETURN SetResult(SQLSetStmtAttrW(m_hStmt, SQL_ATTR_PARAM_STATUS_PTR, cast(SQLPOINTER, cast(ULONG_PTR, dwAttr)), SQL_IS_UINTEGER), "COdbcStmt.SetStmtParamStatusPtr", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
PRIVATE FUNCTION COdbcStmt.GetStmtRowArraySize () AS SQLUINTEGER
   DIM dwAttr AS SQLUINTEGER
   SetResult(SQLGetStmtAttrW(m_hStmt, SQL_ATTR_ROW_ARRAY_SIZE, @dwAttr, SQL_IS_UINTEGER, NULL), "COdbcStmt.GetStmtRowArraySize", SQL_HANDLE_STMT, m_hStmt)
   RETURN dwAttr
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.SetStmtRowArraySize (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
   RETURN SetResult(SQLSetStmtAttrW(m_hStmt, SQL_ATTR_ROW_ARRAY_SIZE, cast(SQLPOINTER, cast(ULONG_PTR, dwAttr)), SQL_IS_UINTEGER), "COdbcStmt.SetStmtRowArraySize", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
PRIVATE FUNCTION COdbcStmt.GetStmtRowBindOffsetPtr () AS SQLUINTEGER
   DIM dwAttr AS SQLUINTEGER
   SetResult(SQLGetStmtAttrW(m_hStmt, SQL_ATTR_ROW_BIND_OFFSET_PTR, @dwAttr, SQL_IS_UINTEGER, NULL), "COdbcStmt.GetStmtRowBindOffsetPtr", SQL_HANDLE_STMT, m_hStmt)
   RETURN dwAttr
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.SetStmtRowBindOffsetPtr (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
   RETURN SetResult(SQLSetStmtAttrW(m_hStmt, SQL_ATTR_ROW_BIND_OFFSET_PTR, cast(SQLPOINTER, cast(ULONG_PTR, dwAttr)), SQL_IS_UINTEGER), "COdbcStmt.SetStmtRowBindOffsetPtr", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
PRIVATE FUNCTION COdbcStmt.GetStmtRowBindType () AS SQLUINTEGER
   DIM dwAttr AS SQLUINTEGER
   SetResult(SQLGetStmtAttrW(m_hStmt, SQL_ATTR_ROW_BIND_TYPE, @dwAttr, SQL_IS_UINTEGER, NULL), "COdbcStmt.GetStmtRowBindType", SQL_HANDLE_STMT, m_hStmt)
   RETURN dwAttr
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.SetStmtRowBindType (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
   RETURN SetResult(SQLSetStmtAttrW(m_hStmt, SQL_ATTR_ROW_BIND_TYPE, cast(SQLPOINTER, cast(ULONG_PTR, dwAttr)), SQL_IS_UINTEGER), "COdbcStmt.SetStmtRowBindType", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' An SQLUINTEGER value that is the number of the current row in the entire result set.
' If the number of the current row cannot be determined or there is no current row, the
' driver returns 0.
' This attribute can be retrieved by a call to GetStmtAttr but not set by a call to
' SetStmtAttr.
' Note: Microsoft Access Driver returns 0.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.GetStmtRowNumber () AS SQLUINTEGER
   DIM dwAttr AS SQLUINTEGER
   SetResult(SQLGetStmtAttrW(m_hStmt, SQL_ATTR_ROW_NUMBER, @dwAttr, SQL_IS_UINTEGER, NULL), "COdbcStmt.GetStmtRowNumber", SQL_HANDLE_STMT, m_hStmt)
   RETURN dwAttr
END FUNCTION
' =====================================================================================

' =====================================================================================
PRIVATE FUNCTION COdbcStmt.GetStmtRowOperationPtr () AS SQLUINTEGER
   DIM dwAttr AS SQLUINTEGER
   SetResult(SQLGetStmtAttrW(m_hStmt, SQL_ATTR_ROW_OPERATION_PTR, @dwAttr, SQL_IS_UINTEGER, NULL), "COdbcStmt.GetStmtRowOperationPtr", SQL_HANDLE_STMT, m_hStmt)
   RETURN dwAttr
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.SettmtRowOperationPtr (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
   RETURN SetResult(SQLSetStmtAttrW(m_hStmt, SQL_ATTR_ROW_OPERATION_PTR, cast(SQLPOINTER, cast(ULONG_PTR, dwAttr)), SQL_IS_UINTEGER), "COdbcStmt.SettmtRowOperationPtr", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
PRIVATE FUNCTION COdbcStmt.GetStmtRowStatusPtr () AS SQLUINTEGER
   DIM dwAttr AS SQLUINTEGER
   SetResult(SQLGetStmtAttrW(m_hStmt, SQL_ATTR_ROW_STATUS_PTR, @dwAttr, SQL_IS_UINTEGER, NULL), "COdbcStmt.GetStmtRowStatusPtr", SQL_HANDLE_STMT, m_hStmt)
   RETURN dwAttr
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.SetStmtRowStatusPtr (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
   RETURN SetResult(SQLSetStmtAttrW(m_hStmt, SQL_ATTR_ROW_STATUS_PTR, cast(SQLPOINTER, cast(ULONG_PTR, dwAttr)), SQL_IS_UINTEGER), "COdbcStmt.SetStmtRowStatusPtr", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
PRIVATE FUNCTION COdbcStmt.GetStmtRowsFetchedPtr () AS SQLUINTEGER
   DIM dwAttr AS SQLUINTEGER
   SetResult(SQLGetStmtAttrW(m_hStmt, SQL_ATTR_ROWS_FETCHED_PTR, @dwAttr, SQL_IS_UINTEGER, NULL), "COdbcStmt.GetStmtRowsFetchedPtr", SQL_HANDLE_STMT, m_hStmt)
   RETURN dwAttr
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.SetStmtRowsFetchedPtr (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
   RETURN SetResult(SQLSetStmtAttrW(m_hStmt, SQL_ATTR_ROWS_FETCHED_PTR, cast(SQLPOINTER, cast(ULONG_PTR, dwAttr)), SQL_IS_UINTEGER), "COdbcStmt.SetStmtRowsFetchedPtr", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
PRIVATE FUNCTION COdbcStmt.GetStmtSimulateCursor () AS SQLUINTEGER
   DIM dwAttr AS SQLUINTEGER
   SetResult(SQLGetStmtAttrW(m_hStmt, SQL_ATTR_SIMULATE_CURSOR, @dwAttr, SQL_IS_UINTEGER, NULL), "COdbcStmt.GetStmtSimulateCursor", SQL_HANDLE_STMT, m_hStmt)
   RETURN dwAttr
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.SetStmtSimulateCursor (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
   RETURN SetResult(SQLSetStmtAttrW(m_hStmt, SQL_ATTR_SIMULATE_CURSOR, cast(SQLPOINTER, cast(ULONG_PTR, dwAttr)), SQL_IS_UINTEGER), "COdbcStmt.SetStmtSimulateCursor", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
PRIVATE FUNCTION COdbcStmt.GetStmtUseBookmarks () AS SQLUINTEGER
   DIM dwAttr AS SQLUINTEGER
   SetResult(SQLGetStmtAttrW(m_hStmt, SQL_ATTR_USE_BOOKMARKS, @dwAttr, SQL_IS_UINTEGER, NULL), "COdbcStmt.GetStmtUseBookmarks", SQL_HANDLE_STMT, m_hStmt)
   RETURN dwAttr
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.SetStmtUseBookmarks (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
   RETURN SetResult(SQLSetStmtAttrW(m_hStmt, SQL_ATTR_USE_BOOKMARKS, cast(SQLPOINTER, cast(ULONG_PTR, dwAttr)), SQL_IS_UINTEGER), "COdbcStmt.SetStmtUseBookmarks", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' #####################################################################################
'                               MISCELLANEOUS FUNCTIONS
' #####################################################################################

' =====================================================================================
' Retrieves a list of statistics about a single table and the indexes associated with
' the table. The driver returns the information as a result set on the specified
' statement handle.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.GetStatistics (BYREF wszCatalogName AS WSTRING, BYVAL CatalogNameLength AS SQLSMALLINT, _
BYREF wszSchemaName AS WSTRING, BYVAL SchemaNameLength AS SQLSMALLINT, BYREF wszTableName AS WSTRING, _
BYVAL TableNameLength AS SQLSMALLINT, BYVAL fUnique AS SQLUSMALLINT, BYVAL fCardinality AS SQLUSMALLINT) AS SQLRETURN
   RETURN SetResult(SQLStatisticsW(m_hStmt, @wszCatalogName, CatalogNameLength, @wszSchemaName, SchemaNameLength, _
          @wszTableName, TableNameLength, fUnique, fCardinality), "COdbcStmt.GetStatistics", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Returns the list of column names in specified tables. The driver returns this
' information as a result set on the specified statement handle.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.GetColumns (BYREF wszCatalogName AS WSTRING, BYVAL CatalogNameLength AS SQLSMALLINT, _
BYREF wszSchemaName AS WSTRING, BYVAL SchemaNameLength AS SQLSMALLINT, BYREF wszTableName AS WSTRING, _
BYVAL TableNameLength AS SQLSMALLINT, BYREF wszColumnName AS WSTRING, BYVAL ColumnNameLength AS SQLSMALLINT) AS SQLRETURN
   RETURN SetResult(SQLColumnsW(m_hStmt, @wszCatalogName, CatalogNameLength, @wszSchemaName, SchemaNameLength, _
          @wszTableName, TableNameLength, @wszColumnName, ColumnNameLength), "COdbcStmt.GetColumns", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Returns information about data types supported by the data source. The driver returns
' the information as a result set on the specified statement handle.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.GetTypeInfo (BYVAL DataType AS SQLSMALLINT) AS SQLRETURN
   RETURN SetResult(SQLGetTypeInfo (m_hStmt, DataType), "COdbcStmt.GetTypeInfo", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Returns the list of table, catalog, or schema names, and table types, stored in a
' specific data source. The driver returns the information as a result set on the
' specified statement handle.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.GetTables (BYREF wszCatalogName AS WSTRING, BYVAL CatalogNameLength AS SQLSMALLINT, _
BYREF wszSchemaName AS WSTRING, BYVAL SchemaNameLength AS SQLSMALLINT, BYREF wszTableName AS WSTRING, _
BYVAL TableNameLength AS SQLSMALLINT, BYREF wszTableType AS WSTRING, BYVAL TableTypeLength AS SQLSMALLINT) AS SQLRETURN
   RETURN SetResult(SQLTablesW(m_hStmt, @wszCatalogName, CatalogNameLength, @wszSchemaName, SchemaNameLength, _
          @wszTableName, TableNameLength, @wszTableType, TableTypeLength), "COdbcStmt.GetTables", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Returns a list of tables and the privileges associated with each table. The driver
' returns the information as a result set on the specified statement handle.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.GetTablePrivileges (BYREF wszCatalogName AS WSTRING, BYVAL CatalogNameLength AS SQLSMALLINT, _
BYREF wszSchemaName AS WSTRING, BYVAL SchemaNameLength AS SQLSMALLINT, BYREF wszTableName AS WSTRING, _
BYVAL TableNameLength AS SQLSMALLINT) AS SQLRETURN
   RETURN SetResult(SQLTablePrivilegesW(m_hStmt, @wszCatalogName, CatalogNameLength, @wszSchemaName, SchemaNameLength, @wszTableName, TableNameLength), "COdbcStmt.GetTablePrivileges", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Returns the column names that make up the primary key for a table. This function does
' not support returning primary keys from multiple tables in a single call. The driver
' returns the information as a result set on the specified statement handle.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.GetPrimaryKeys (BYREF wszCatalogName AS WSTRING, BYVAL CatalogNameLength AS SQLSMALLINT, _
BYREF wszSchemaName AS WSTRING, BYVAL SchemanameLength AS SQLSMALLINT, BYREF wszTableName AS WSTRING, _
BYVAL TableNameLength AS SQLSMALLINT) AS SQLRETURN
   RETURN SetResult(SQLPrimaryKeysW(m_hStmt, @wszCatalogName, CatalogNameLength, @wszSchemaName, SchemaNameLength, _
          @wszTableName, TableNameLength), "COdbcStmt.GetPrimaryKeys", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Returns the list of procedure names stored in a specific data source. Procedure is a
' generic term used to describe an executable object, or a named entity that can be
' invoked using input and output parameters. The driver returns the information as a
' result set on the specified statement handle.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.GetProcedures (BYREF wszCatalogName AS WSTRING, BYVAL CatalogNameLength AS SQLSMALLINT, _
BYREF wszSchemaName AS WSTRING, BYVAL SchemaNameLength AS SQLSMALLINT, BYREF wszProcName AS WSTRING, _
BYVAL ProcNameLength AS SQLSMALLINT) AS SQLRETURN
   RETURN SetResult(SQLProceduresW(m_hStmt, @wszCatalogName, CatalogNameLength, @wszSchemaName, SchemaNameLength, _
          @wszProcName, ProcNameLength), "COdbcStmt.GetProcedures", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Retrieves the following information about columns within a specified table.
' The driver returns the information as a result set on the specified statement handle.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.GetSpecialColumns (BYVAL identifierType AS SQLUSMALLINT, BYREF wszCatalogName AS WSTRING, _
BYVAL CatalogNameLength AS SQLUSMALLINT, BYREF wszSchemaName AS WSTRING, BYVAL SchemaNameLength AS SQLUSMALLINT, _
BYREF wszTableName AS WSTRING, BYVAL TableNameLength AS SQLUSMALLINT, BYVAL fScope AS SQLUSMALLINT,_
BYVAL fNullable AS SQLUSMALLINT) AS SQLRETURN
   RETURN SetResult(SQLSpecialColumnsW(m_hStmt, identifierType, @wszCatalogName, CatalogNameLength, _
          @wszSchemaName, SchemaNameLength, @wszTableName, TableNameLength, fScope, fNullable), "COdbcStmt.GetSpecialColumns", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Returns a list of columns and associated privileges for the specified table. The
' driver returns the information as a result set on the specified statement handle.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.GetColumnPrivileges (BYREF wszCatalogName AS WSTRING, BYVAL CatalogNameLength AS SQLSMALLINT, _
BYREF wszSchemaName AS WSTRING, BYVAL SchemaNameLength AS SQLSMALLINT, BYREF wszTableName AS WSTRING, _
BYVAL TableNameLength AS SQLSMALLINT, BYREF wszColumnName AS WSTRING, BYVAL ColumnNameLength AS SQLSMALLINT) AS SQLRETURN
   RETURN SetResult(SQLColumnPrivilegesW(m_hStmt, @wszCatalogName, CatalogNameLength, _
       @wszSchemaName, SchemaNameLength, @wszTableName, TableNameLength, @wszColumnName, ColumnNameLength), "COdbcStmt.GetColumnPrivileges", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Returns a list of foreign keys. The driver returns the information as a result set
' on the specified statement handle.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.GetForeignKeys (BYREF wszPkCatalogName AS WSTRING, BYVAL pkCatalogNameLength AS SQLSMALLINT, _
BYREF wszPkSchemaName AS WSTRING, BYVAL pkSchemaNameLength AS SQLSMALLINT, BYREF wszPkTableName AS WSTRING, _
BYVAL pkTableNameLength AS SQLSMALLINT, BYREF wszFkCatalogName AS WSTRING, BYVAL FkCatalogNameLength AS SQLSMALLINT, _
BYREF wszFkSchemaName AS WSTRING, BYVAL FkSchemaNameLength AS SQLSMALLINT, BYREF wszFkTableName AS WSTRING, BYVAL fkTableNameLength AS SQLSMALLINT) AS SQLRETURN
   RETURN SetResult(SQLForeignKeysW(m_hStmt, @wszPkCatalogName, PkCatalogNameLength, _
          @wszPkSchemaName, PkSchemaNameLength, @wszPkTableName, PkTableNameLength, _
          @wszFkCatalogName, FkCatalogNameLength, @wszFkSchemaName, FkSchemaNameLength, _
          @wszFkTableName, FkTableNameLength), "COdbcStmt.GetForeignKeys", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Returns the list of input and output parameters, as well as the columns that make up
' the result set for the specified procedures. The driver returns the information as a
' result set on the specified statement.
' =====================================================================================
PRIVATE FUNCTION COdbcStmt.GetProcedureColumns (BYREF wszCatalogName AS WSTRING, BYVAL CatalogNameLength AS SQLSMALLINT, _
BYREF wszSchemaName AS WSTRING, BYVAL SchemaNameLength AS SQLSMALLINT, BYREF wszProcName AS WSTRING, _
BYVAL ProcNameLength AS SQLSMALLINT, BYREF wszColumnName AS WSTRING, BYVAL ColumnNameLength AS SQLSMALLINT) AS SQLRETURN
   RETURN SetResult(SQLProcedureColumnsW(m_hStmt, @wszCatalogName, CatalogNameLength, _
          @wszSchemaName, SchemaNameLength, @wszProcName, ProcNameLength, @wszColumnName, ColumnNameLength), "COdbcStmt.GetProcedureColumns", SQL_HANDLE_STMT, m_hStmt)
END FUNCTION
' =====================================================================================
