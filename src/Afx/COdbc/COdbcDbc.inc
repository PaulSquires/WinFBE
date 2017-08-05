' ########################################################################################
' Microsoft Windows
' File: COdbcDbc.inc
' Contents: ODBC wrapper classes (unicode)
' Implementation of the ODBC connection class.
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2016-2017 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

' ========================================================================================
' Allocates a connection handle and opens the database.
' In Windows 7, the Driver Manager version is 03.80. In Windows 8, the Driver Manager
' version is 03.81 via the SQLGetInfo SQL_DM_VER (InfoType parameter). SQL_ODBC_VER reports
' the version as 03.80 in both Windows 7 and Windows 8.
' ========================================================================================
CONSTRUCTOR COdbc (BYREF wszConnectionString AS WSTRING, BYVAL nODbcVersion AS SQLINTEGER = SQL_OV_ODBC3, BYVAL ConnectionPoolingAttr AS SQLUINTEGER = 0)
   IF AFX_ODBC_hEnv = NULL THEN
      ' // Connection pooling
      IF ConnectionPoolingAttr = SQL_CP_ONE_PER_DRIVER OR ConnectionPoolingAttr = SQL_CP_ONE_PER_HENV THEN
         SQLSetEnvAttr(NULL, SQL_ATTR_CONNECTION_POOLING, cast(SQLPOINTER, cast(LONG_PTR, ConnectionPoolingAttr)), SQL_IS_UINTEGER)
      END IF
      ' // Allocates the environment handle
      SetResult(SQLAllocHandle(SQL_HANDLE_ENV, SQL_NULL_HENV, @AFX_ODBC_hEnv))
      ' // Sets the ODBC version
      IF AFX_ODBC_hEnv THEN
         SetResult(SQLSetEnvAttr(AFX_ODBC_hEnv, SQL_ATTR_ODBC_VERSION, cast(SQLPOINTER, cast(LONG_PTR, nODbcVersion)), SQL_IS_INTEGER))
      END IF
   END IF
   IF AFX_ODBC_hEnv THEN
      ' // Allocate a connection handle
      SetResult(SQLAllocHandle (SQL_HANDLE_DBC, AFX_ODBC_hEnv, @m_hDbc))
      IF m_Result <> SQL_ERROR AND m_Result <> SQL_INVALID_HANDLE THEN AFX_ODBC_numConnections += 1
   END IF
   IF m_hDbc THEN
      DIM StringLength2Ptr AS SQLSMALLINT
      DIM wszOutConnectionString AS WSTRING * 1024
      ' Note: The value of the 6th parameter is the number of characters (SIZEOF returns the number of bytes).
      SetResult(SQLDriverConnectW(m_hDbc, HWND_DESKTOP, @wszConnectionString, SQL_NTS, _
          @wszOutConnectionString, 1024, @StringLength2Ptr, SQL_DRIVER_COMPLETE))
   END IF
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Establishes connections to a driver and a data source.
' ========================================================================================
CONSTRUCTOR COdbc (BYREF wszServerName AS WSTRING, BYREF wszUserName AS WSTRING, BYREF wszAuthentication AS WSTRING, BYVAL nODbcVersion AS SQLINTEGER = SQL_OV_ODBC3, BYVAL ConnectionPoolingAttr AS SQLUINTEGER = 0)
   IF AFX_ODBC_hEnv = NULL THEN
      ' // Connection pooling
      IF ConnectionPoolingAttr = SQL_CP_ONE_PER_DRIVER OR ConnectionPoolingAttr = SQL_CP_ONE_PER_HENV THEN
         SQLSetEnvAttr(NULL, SQL_ATTR_CONNECTION_POOLING, cast(SQLPOINTER, cast(LONG_PTR, ConnectionPoolingAttr)), SQL_IS_UINTEGER)
      END IF
      ' // Allocates the environment handle
      SetResult(SQLAllocHandle(SQL_HANDLE_ENV, SQL_NULL_HENV, @AFX_ODBC_hEnv))
      ' // Sets the ODBC version
      IF AFX_ODBC_hEnv THEN
         SetResult(SQLSetEnvAttr(AFX_ODBC_hEnv, SQL_ATTR_ODBC_VERSION, cast(SQLPOINTER, cast(LONG_PTR, nODbcVersion)), SQL_IS_INTEGER))
      END IF
   END IF
   IF AFX_ODBC_hEnv THEN
      ' // Allocate a connection handle
      SetResult(SQLAllocHandle (SQL_HANDLE_DBC, AFX_ODBC_hEnv, @m_hDbc))
      IF m_Result <> SQL_ERROR AND m_Result <> SQL_INVALID_HANDLE THEN AFX_ODBC_numConnections += 1
   END IF
   IF m_hDbc THEN
      SetResult(SQLConnectW(m_hDbc, wszServerName, LEN(wszServerName), wszUserName, LEN(wszUserName), wszAuthentication, LEN(wszAuthentication)))
   END IF
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Cleanup
' ========================================================================================
PRIVATE DESTRUCTOR COdbc
   IF m_hDbc THEN
      SQLDisconnect(m_hDbc)
      SQLFreeHandle(SQL_HANDLE_DBC, m_hDbc)
      AFX_ODBC_numConnections -= 1
      IF AFX_ODBC_numConnections = 0 THEN
         SQLFreeHandle(SQL_HANDLE_ENV, AFX_ODBC_hEnv)
         AFX_ODBC_hEnv = NULL
      END IF
   END IF
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the connection handle
' ========================================================================================
PRIVATE FUNCTION COdbc.Handle () AS SQLHANDLE
   RETURN m_hDbc
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the environment handle
' ========================================================================================
PRIVATE FUNCTION COdbc.EnvHandle () AS SQLHANDLE
   RETURN AFX_ODBC_hEnv
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the current value of a field of a record of the diagnostic data structure
' (associated with a specific handle) that contains error, warning, and status information.
' ========================================================================================
PRIVATE FUNCTION COdbc.GetDiagField (BYVAL RecNumber AS SQLSMALLINT, BYVAL DiagIdentifier AS SQLSMALLINT, _
BYVAL DiagInfoPtr AS SQLPOINTER, BYVAL BufferLength AS SQLSMALLINT, BYVAL StringLengthPtr AS SQLSMALLINT PTR) AS SQLRETURN
   RETURN SetResult(SQLGetDiagFieldW(SQL_HANDLE_DBC, m_hDbc, RecNumber, DiagIdentifier, DiagInfoPtr, _
          BufferLength, StringLengthPtr), "COdbc.GetDiagField", SQL_HANDLE_DBC, m_hDbc)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the current values of multiple fields of a diagnostic record that contains
' error, warning and status information. Unlike GetDiagField, which returns one
' diagnostic call per call, GetDiagRec returns several commonly used fields of a
' diagnostic record, including the SQLSTATE, the native error code, and the
' diagnostic message text.
' ========================================================================================
PRIVATE FUNCTION COdbc.GetDiagRec (BYVAL RecNumber AS SQLSMALLINT, BYVAL Sqlstate AS WSTRING PTR, BYVAL NativeError AS SQLINTEGER PTR, _
BYVAL MessageText AS WSTRING PTR, BYVAL BufferLength AS SQLSMALLINT, BYVAL TextLength AS SQLSMALLINT PTR) AS SQLRETURN
   RETURN SetResult(SQLGetDiagRecW(SQL_HANDLE_DBC, m_hDbc, RecNumber, Sqlstate, NativeError, MessageText, _
          BufferLength, TextLength), "COdbc.GetDiagRec", SQL_HANDLE_DBC, m_hDbc)
END FUNCTION
' ========================================================================================

' ===========================================================================================
' Uses SQLGetDiagRec to retrieve an error description.
' iErrorCode: Optional. The error code returned by some of the methods of this interface.
' ===========================================================================================
PRIVATE FUNCTION COdbc.GetErrorInfo (BYVAL iErrorCode AS SQLRETURN = 0) AS CWSTR
   RETURN Base.GetErrorInfo(SQL_HANDLE_DBC, m_hDbc, iErrorCode)
END FUNCTION
' ===========================================================================================

' ========================================================================================
' Gets the SqlState for the environment handle
' ========================================================================================
PRIVATE FUNCTION COdbc.GetSqlState () AS CWSTR
   RETURN Base.GetSqlState(SQL_HANDLE_DBC, m_hDbc)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the SQL string as modified by the driver.
' ========================================================================================
PRIVATE FUNCTION COdbc.NativeSql (BYREF wszInText AS WSTRING) AS CWSTR
   DIM wszOutText AS WSTRING * 1024, cbbytes AS SQLINTEGER
   SetResult(SQLNativeSqlW(m_hDbc, @wszInText, LEN(wszInText), @wszOutText, 1024, @cbbytes),_
      "COdbc.NativeSql", SQL_HANDLE_DBC, m_hDbc)
   RETURN wszOutText
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns information about whether a driver supports a specific ODBC function.
' ========================================================================================
PRIVATE FUNCTION COdbc.Functions (BYVAL FunctionId AS SQLUSMALLINT) AS BOOLEAN
   DIM Supported AS SQLUSMALLINT
   SetResult(SQLGetFunctions(m_hDbc, FunctionId, @Supported), "COdbc.Functions", SQL_HANDLE_DBC, m_hDbc)
   RETURN Supported
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns information about whether a driver supports a specific ODBC function.
' It is an alias for the above function.
' ========================================================================================
PRIVATE FUNCTION COdbc.Supports (BYVAL FunctionId AS SQLUSMALLINT) AS BOOLEAN
   DIM Supported AS SQLUSMALLINT
   SetResult(SQLGetFunctions(m_hDbc, FunctionId, @Supported), "COdbc.Supports", SQL_HANDLE_DBC, m_hDbc)
   RETURN Supported
END FUNCTION
' ========================================================================================

' ########################################################################################
' TRANSACTIONS
' ########################################################################################

' ========================================================================================
' Commits a transaction
' ========================================================================================
PRIVATE FUNCTION COdbc.CommitTran () AS SQLRETURN
   RETURN SetResult(SQLEndTran(SQL_HANDLE_DBC, m_hDbc, SQL_COMMIT), "COdbc.CommitTran", SQL_HANDLE_DBC, m_hDbc)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Rollbacks a transaction
' ========================================================================================
PRIVATE FUNCTION COdbc.RollbackTran () AS SQLRETURN
   RETURN SetResult(SQLEndTran(SQL_HANDLE_DBC, m_hDbc, SQL_ROLLBACK), "COdbc.RollbackTran", SQL_HANDLE_DBC, m_hDbc)
END FUNCTION
' ========================================================================================

' ########################################################################################
' CONNECTION ATTRIBUTES
' ########################################################################################

' ========================================================================================
' Returns the current setting of a connection attribute.
' ========================================================================================
PRIVATE FUNCTION COdbc.GetConnectAttr (BYVAL Attribute AS SQLINTEGER, BYVAL ValuePtr AS SQLPOINTER, BYVAL _
BufferLength AS SQLINTEGER, BYVAL StringLength AS SQLINTEGER PTR) AS SQLRETURN
   RETURN SetResult(SQLGetConnectAttrW(m_hDbc, Attribute, BYVAL ValuePtr, BufferLength, StringLength), _
          "COdbc.GetConnectAttr", SQL_HANDLE_DBC, m_hDbc)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the current setting of a connection attribute.
' ========================================================================================
PRIVATE FUNCTION COdbc.SetConnectAttr (BYVAL Attribute AS SQLINTEGER, BYVAL ValuePtr AS SQLPOINTER, BYVAL StringLength AS SQLINTEGER) AS SQLRETURN
   RETURN SetResult(SQLSetConnectAttrW(m_hDbc, Attribute, BYVAL ValuePtr, StringLength), _
          "COdbc.SetConnectAttr", SQL_HANDLE_DBC, m_hDbc)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the current setting of the specified connection attribute.
' ========================================================================================
PRIVATE FUNCTION COdbc.GetConnectAttr (BYREF wszAttribute AS WSTRING) AS SQLUINTEGER
   DIM dwAttr AS SQLUINTEGER
   SELECT CASE LCASE(wszAttribute)
      CASE "accessmode" : SetResult(SQLGetConnectAttrW(m_hDbc, SQL_ATTR_ACCESS_MODE, @dwAttr, SQL_IS_UINTEGER, NULL), "COdbc.GetConnectAttr", SQL_HANDLE_DBC, m_hDbc)
      CASE "asyncenable" : SetResult(SQLGetConnectAttrW(m_hDbc, SQL_ATTR_ASYNC_ENABLE, @dwAttr, SQL_IS_UINTEGER, NULL), "COdbc.GetConnectAttr", SQL_HANDLE_DBC, m_hDbc)
      CASE "autoipd" : SetResult(SQLGetConnectAttrW(m_hDbc, SQL_ATTR_AUTO_IPD, @dwAttr, SQL_IS_UINTEGER, NULL), "COdbc.GetConnectAttr", SQL_HANDLE_DBC, m_hDbc)
      CASE "autocommit" : SetResult(SQLGetConnectAttrW(m_hDbc, SQL_ATTR_AUTOCOMMIT, @dwAttr, SQL_IS_UINTEGER, NULL), "COdbc.GetConnectAttr", SQL_HANDLE_DBC, m_hDbc)
      CASE "connectiondead" : SetResult(SQLGetConnectAttrW(m_hDbc, SQL_ATTR_CONNECTION_DEAD, @dwAttr, SQL_IS_UINTEGER, NULL), "COdbc.GetConnectAttr", SQL_HANDLE_DBC, m_hDbc)
      CASE "connectiontimeout" : SetResult(SQLGetConnectAttrW(m_hDbc, SQL_ATTR_CONNECTION_TIMEOUT, @dwAttr, SQL_IS_UINTEGER, NULL), "COdbc.GetConnectAttr", SQL_HANDLE_DBC, m_hDbc)
      CASE "logintimeout" : SetResult(SQLGetConnectAttrW(m_hDbc, SQL_ATTR_LOGIN_TIMEOUT, @dwAttr, SQL_IS_UINTEGER, NULL), "COdbc.GetConnectAttr", SQL_HANDLE_DBC, m_hDbc)
      CASE "metadataid" : SetResult(SQLGetConnectAttrW(m_hDbc, SQL_ATTR_METADATA_ID, @dwAttr, SQL_IS_UINTEGER, NULL), "COdbc.GetConnectAttr", SQL_HANDLE_DBC, m_hDbc)
      CASE "cursors" : SetResult(SQLGetConnectAttrW(m_hDbc, SQL_ATTR_ODBC_CURSORS, @dwAttr, SQL_IS_UINTEGER, NULL), "COdbc.GetConnectAttr", SQL_HANDLE_DBC, m_hDbc)
      CASE "packetsize" : SetResult(SQLGetConnectAttrW(m_hDbc, SQL_ATTR_PACKET_SIZE, @dwAttr, SQL_IS_UINTEGER, NULL), "COdbc.GetConnectAttr", SQL_HANDLE_DBC, m_hDbc)
      CASE "quietmode" : SetResult(SQLGetConnectAttrW(m_hDbc, SQL_ATTR_QUIET_MODE, @dwAttr, SQL_IS_UINTEGER, NULL), "COdbc.GetConnectAttr", SQL_HANDLE_DBC, m_hDbc)
      CASE "trace" : SetResult(SQLGetConnectAttrW(m_hDbc, SQL_ATTR_TRACE, @dwAttr, SQL_IS_UINTEGER, NULL), "COdbc.GetConnectAttr", SQL_HANDLE_DBC, m_hDbc)
      CASE "txnisolation" : SetResult(SQLGetConnectAttrW(m_hDbc, SQL_ATTR_TXN_ISOLATION, @dwAttr, SQL_IS_UINTEGER, NULL), "COdbc.GetConnectAttr", SQL_HANDLE_DBC, m_hDbc)
   END SELECT
   RETURN dwAttr
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION COdbc.GetConnectAttrStr (BYREF wszAttribute AS WSTRING) AS CWSTR
   DIM wszAttr AS WSTRING * MAX_PATH + 1
   SELECT CASE LCASE(wszAttribute)
      CASE "currentcatalog" : SetResult(SQLGetConnectAttrW(m_hDbc, SQL_ATTR_CURRENT_CATALOG, @wszAttr, MAX_PATH + 1, NULL), "COdbc.GetConnectAttrStr", SQL_HANDLE_DBC, m_hDbc)
      CASE "tracefile" : SetResult(SQLGetConnectAttrW(m_hDbc, SQL_ATTR_TRACEFILE, @wszAttr, MAX_PATH + 1, NULL), "COdbc.GetConnectAttrStr", SQL_HANDLE_DBC, m_hDbc)
      CASE "translatelib" : SetResult(SQLGetConnectAttrW(m_hDbc, SQL_ATTR_TRANSLATE_LIB, @wszAttr, MAX_PATH + 1, NULL), "COdbc.GetConnectAttrStr", SQL_HANDLE_DBC, m_hDbc)
   END SELECT
   RETURN wszAttr
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the setting of the specified connection attribute.
' ========================================================================================
PRIVATE FUNCTION COdbc.SetConnectAttr (BYREF wszAttribute AS WSTRING, BYVAL dwAttribute AS SQLUINTEGER) AS SQLRETURN
   DIM dwAttr AS SQLUINTEGER
   SELECT CASE LCASE(wszAttribute)
      CASE "accessmode" : RETURN SetResult(SQLSetConnectAttrW(m_hDbc, SQL_ATTR_ACCESS_MODE, @dwAttr, SQL_IS_UINTEGER), "COdbc.SetConnectAttr", SQL_HANDLE_DBC, m_hDbc)
      CASE "asyncenable" : RETURN SetResult(SQLSetConnectAttrW(m_hDbc, SQL_ATTR_ASYNC_ENABLE, @dwAttr, SQL_IS_UINTEGER), "COdbc.SetConnectAttr", SQL_HANDLE_DBC, m_hDbc)
      CASE "autocommit" : RETURN SetResult(SQLSetConnectAttrW(m_hDbc, SQL_ATTR_AUTOCOMMIT, @dwAttr, SQL_IS_UINTEGER), "COdbc.SetConnectAttr", SQL_HANDLE_DBC, m_hDbc)
      CASE "connectiontimeout" : RETURN SetResult(SQLSetConnectAttrW(m_hDbc, SQL_ATTR_CONNECTION_TIMEOUT, @dwAttr, SQL_IS_UINTEGER), "COdbc.SetConnectAttr", SQL_HANDLE_DBC, m_hDbc)
      CASE "logintimeout" : RETURN SetResult(SQLSetConnectAttrW(m_hDbc, SQL_ATTR_LOGIN_TIMEOUT, @dwAttr, SQL_IS_UINTEGER), "COdbc.SetConnectAttr", SQL_HANDLE_DBC, m_hDbc)
      CASE "metadataid" : RETURN SetResult(SQLSetConnectAttrW(m_hDbc, SQL_ATTR_METADATA_ID, @dwAttr, SQL_IS_UINTEGER), "COdbc.SetConnectAttr", SQL_HANDLE_DBC, m_hDbc)
      CASE "cursors" : RETURN SetResult(SQLSetConnectAttrW(m_hDbc, SQL_ATTR_ODBC_CURSORS, @dwAttr, SQL_IS_UINTEGER), "COdbc.SetConnectAttr", SQL_HANDLE_DBC, m_hDbc)
      CASE "packetsize" : RETURN SetResult(SQLSetConnectAttrW(m_hDbc, SQL_ATTR_PACKET_SIZE, @dwAttr, SQL_IS_UINTEGER), "COdbc.SetConnectAttr", SQL_HANDLE_DBC, m_hDbc)
      CASE "quietmode" : RETURN SetResult(SQLSetConnectAttrW(m_hDbc, SQL_ATTR_QUIET_MODE, @dwAttr, SQL_IS_UINTEGER), "COdbc.SetConnectAttr", SQL_HANDLE_DBC, m_hDbc)
      CASE "trace" : RETURN SetResult(SQLSetConnectAttrW(m_hDbc, SQL_ATTR_TRACE, @dwAttr, SQL_IS_UINTEGER), "COdbc.SetConnectAttr", SQL_HANDLE_DBC, m_hDbc)
      CASE "txnisolation" : RETURN SetResult(SQLSetConnectAttrW(m_hDbc, SQL_ATTR_TXN_ISOLATION, @dwAttr, SQL_IS_UINTEGER), "COdbc.SetConnectAttr", SQL_HANDLE_DBC, m_hDbc)
   END SELECT
   RETURN dwAttr
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION COdbc.SetConnectAttr (BYREF wszAttribute AS WSTRING, BYREF wszAttrValue AS WSTRING) AS SQLRETURN
   SELECT CASE LCASE(wszAttribute)
      CASE "currentcatalog" : RETURN SetResult(SQLSetConnectAttrW(m_hDbc, SQL_ATTR_CURRENT_CATALOG, @wszAttrValue, SQL_NTS), "COdbc.SetConnectAttr", SQL_HANDLE_DBC, m_hDbc)
   END SELECT
END FUNCTION
' ========================================================================================


' ########################################################################################
' GENERAL INFORMATION
' ########################################################################################

' ========================================================================================
' Returns general information about the driver and data source associated with a connection.
' ========================================================================================
PRIVATE FUNCTION COdbc.GetInfo (BYVAL InfoType AS SQLUSMALLINT, BYVAL InfoValuePtr AS SQLPOINTER, _
BYVAL BufferLength AS SQLSMALLINT, BYVAL StringLength AS SQLSMALLINT PTR) AS SQLRETURN
   RETURN SetResult(SQLGetInfoW(m_hDbc, InfoType, BYVAL InfoValuePtr, BufferLength, StringLength), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION COdbc.GetInfoStr (BYREF wszInfoType AS WSTRING) AS CWSTR
   DIM wszInfoValue AS WSTRING * 1024
   SELECT CASE LCASE(wszInfoType)
      CASE "dmver" : SetResult(SQLGetInfoW(m_hDbc, SQL_DM_VER, @wszInfoValue, 1024, NULL), "COdbc.GetInfoStr", SQL_HANDLE_DBC, m_hDbc)
      CASE "xopencliyear" : SetResult(SQLGetInfoW(m_hDbc, SQL_XOPEN_CLI_YEAR, @wszInfoValue, 1024, NULL), "COdbc.GetInfoStr", SQL_HANDLE_DBC, m_hDbc)
      CASE "datasourcename" : SetResult(SQLGetInfoW(m_hDbc, SQL_DATA_SOURCE_NAME, @wszInfoValue, 1024, NULL), "COdbc.GetInfoStr", SQL_HANDLE_DBC, m_hDbc)
      CASE "drivername" : SetResult(SQLGetInfoW(m_hDbc, SQL_DRIVER_NAME, @wszInfoValue, 1024, NULL), "COdbc.GetInfoStr", SQL_HANDLE_DBC, m_hDbc)
      CASE "driverodbcver" : SetResult(SQLGetInfoW(m_hDbc, SQL_DRIVER_ODBC_VER, @wszInfoValue, 1024, NULL), "COdbc.GetInfoStr", SQL_HANDLE_DBC, m_hDbc)
      CASE "driverver" : SetResult(SQLGetInfoW(m_hDbc, SQL_DRIVER_VER, @wszInfoValue, 1024, NULL), "COdbc.GetInfoStr", SQL_HANDLE_DBC, m_hDbc)
      CASE "odbcver" : SetResult(SQLGetInfoW(m_hDbc, SQL_ODBC_VER, @wszInfoValue, 1024, NULL), "COdbc.GetInfoStr", SQL_HANDLE_DBC, m_hDbc)
      CASE "rowupdates" : SetResult(SQLGetInfoW(m_hDbc, SQL_ROW_UPDATES, @wszInfoValue, 1024, NULL), "COdbc.GetInfoStr", SQL_HANDLE_DBC, m_hDbc)
      CASE "searchpatternescape" : SetResult(SQLGetInfoW(m_hDbc, SQL_SEARCH_PATTERN_ESCAPE, @wszInfoValue, 1024, NULL), "COdbc.GetInfoStr", SQL_HANDLE_DBC, m_hDbc)
      CASE "servername" : SetResult(SQLGetInfoW(m_hDbc, SQL_SERVER_NAME, @wszInfoValue, 1024, NULL), "COdbc.GetInfoStr", SQL_HANDLE_DBC, m_hDbc)
      CASE "databasename" : SetResult(SQLGetInfoW(m_hDbc, SQL_DATABASE_NAME, @wszInfoValue, 1024, NULL), "COdbc.GetInfoStr", SQL_HANDLE_DBC, m_hDbc)
      CASE "dbmsname" : SetResult(SQLGetInfoW(m_hDbc, SQL_DBMS_NAME, @wszInfoValue, 1024, NULL), "COdbc.GetInfoStr", SQL_HANDLE_DBC, m_hDbc)
      CASE "dbmsver" : SetResult(SQLGetInfoW(m_hDbc, SQL_DBMS_VER, @wszInfoValue, 1024, NULL), "COdbc.GetInfoStr", SQL_HANDLE_DBC, m_hDbc)
      CASE "accessibleprocedures" : SetResult(SQLGetInfoW(m_hDbc, SQL_ACCESSIBLE_PROCEDURES, @wszInfoValue, 1024, NULL), "COdbc.GetInfoStr", SQL_HANDLE_DBC, m_hDbc)
      CASE "accessibletables" : SetResult(SQLGetInfoW(m_hDbc, SQL_ACCESSIBLE_TABLES, @wszInfoValue, 1024, NULL), "COdbc.GetInfoStr", SQL_HANDLE_DBC, m_hDbc)
      CASE "catalogterm" : SetResult(SQLGetInfoW(m_hDbc, SQL_CATALOG_TERM, @wszInfoValue, 1024, NULL), "COdbc.GetInfoStr", SQL_HANDLE_DBC, m_hDbc)
      CASE "collationseq" : SetResult(SQLGetInfoW(m_hDbc, SQL_COLLATION_SEQ, @wszInfoValue, 1024, NULL), "COdbc.GetInfoStr", SQL_HANDLE_DBC, m_hDbc)
      CASE "datasourcereadonly" : SetResult(SQLGetInfoW(m_hDbc, SQL_DATA_SOURCE_READ_ONLY, @wszInfoValue, 1024, NULL), "COdbc.GetInfoStr", SQL_HANDLE_DBC, m_hDbc)
      CASE "describeparameter" : SetResult(SQLGetInfoW(m_hDbc, SQL_DESCRIBE_PARAMETER, @wszInfoValue, 1024, NULL), "COdbc.GetInfoStr", SQL_HANDLE_DBC, m_hDbc)
      CASE "multresultsets" : SetResult(SQLGetInfoW(m_hDbc, SQL_MULT_RESULT_SETS, @wszInfoValue, 1024, NULL), "COdbc.GetInfoStr", SQL_HANDLE_DBC, m_hDbc)
      CASE "multipleactivetxn" : SetResult(SQLGetInfoW(m_hDbc, SQL_MULTIPLE_ACTIVE_TXN, @wszInfoValue, 1024, NULL), "COdbc.GetInfoStr", SQL_HANDLE_DBC, m_hDbc)
      CASE "needlongdatalen" : SetResult(SQLGetInfoW(m_hDbc, SQL_NEED_LONG_DATA_LEN, @wszInfoValue, 1024, NULL), "COdbc.GetInfoStr", SQL_HANDLE_DBC, m_hDbc)
      CASE "procedureterm" : SetResult(SQLGetInfoW(m_hDbc, SQL_PROCEDURE_TERM, @wszInfoValue, 1024, NULL), "COdbc.GetInfoStr", SQL_HANDLE_DBC, m_hDbc)
      CASE "schematerm" : SetResult(SQLGetInfoW(m_hDbc, SQL_SCHEMA_TERM, @wszInfoValue, 1024, NULL), "COdbc.GetInfoStr", SQL_HANDLE_DBC, m_hDbc)
      CASE "tableterm" : SetResult(SQLGetInfoW(m_hDbc, SQL_TABLE_TERM, @wszInfoValue, 1024, NULL), "COdbc.GetInfoStr", SQL_HANDLE_DBC, m_hDbc)
      CASE "username" : SetResult(SQLGetInfoW(m_hDbc, SQL_USER_NAME, @wszInfoValue, 1024, NULL), "COdbc.GetInfoStr", SQL_HANDLE_DBC, m_hDbc)
      CASE "catalogname" : SetResult(SQLGetInfoW(m_hDbc, SQL_CATALOG_NAME, @wszInfoValue, 1024, NULL), "COdbc.GetInfoStr", SQL_HANDLE_DBC, m_hDbc)
      CASE "catalognameseparator" : SetResult(SQLGetInfoW(m_hDbc, SQL_CATALOG_NAME_SEPARATOR, @wszInfoValue, 1024, NULL), "COdbc.GetInfoStr", SQL_HANDLE_DBC, m_hDbc)
      CASE "columnalias" : SetResult(SQLGetInfoW(m_hDbc, SQL_COLUMN_ALIAS, @wszInfoValue, 1024, NULL), "COdbc.GetInfoStr", SQL_HANDLE_DBC, m_hDbc)
      CASE "expressionsinorderby" : SetResult(SQLGetInfoW(m_hDbc, SQL_EXPRESSIONS_IN_ORDERBY, @wszInfoValue, 1024, NULL), "COdbc.GetInfoStr", SQL_HANDLE_DBC, m_hDbc)
      CASE "identifierquotechar" : SetResult(SQLGetInfoW(m_hDbc, SQL_IDENTIFIER_QUOTE_CHAR, @wszInfoValue, 1024, NULL), "COdbc.GetInfoStr", SQL_HANDLE_DBC, m_hDbc)
      CASE "integrity" : SetResult(SQLGetInfoW(m_hDbc, SQL_INTEGRITY, @wszInfoValue, 1024, NULL), "COdbc.GetInfoStr", SQL_HANDLE_DBC, m_hDbc)
      CASE "keywords" : SetResult(SQLGetInfoW(m_hDbc, SQL_KEYWORDS, @wszInfoValue, 1024, NULL), "COdbc.GetInfoStr", SQL_HANDLE_DBC, m_hDbc)
      CASE "likeescapeclause" : SetResult(SQLGetInfoW(m_hDbc, SQL_LIKE_ESCAPE_CLAUSE, @wszInfoValue, 1024, NULL), "COdbc.GetInfoStr", SQL_HANDLE_DBC, m_hDbc)
      CASE "outerjoins" : SetResult(SQLGetInfoW(m_hDbc, SQL_OUTER_JOINS, @wszInfoValue, 1024, NULL), "COdbc.GetInfoStr", SQL_HANDLE_DBC, m_hDbc)
      CASE "orderbycolumnsinselect" : SetResult(SQLGetInfoW(m_hDbc, SQL_ORDER_BY_COLUMNS_IN_SELECT, @wszInfoValue, 1024, NULL), "COdbc.GetInfoStr", SQL_HANDLE_DBC, m_hDbc)
      CASE "procedures" : SetResult(SQLGetInfoW(m_hDbc, SQL_PROCEDURES, @wszInfoValue, 1024, NULL), "COdbc.GetInfoStr", SQL_HANDLE_DBC, m_hDbc)
      CASE "specialcharacters" : SetResult(SQLGetInfoW(m_hDbc, SQL_SPECIAL_CHARACTERS, @wszInfoValue, 1024, NULL), "COdbc.GetInfoStr", SQL_HANDLE_DBC, m_hDbc)
      CASE "maxrowsizeincludeslong" : SetResult(SQLGetInfoW(m_hDbc, SQL_MAX_ROW_SIZE_INCLUDES_LONG, @wszInfoValue, 1024, NULL), "COdbc.GetInfoStr", SQL_HANDLE_DBC, m_hDbc)
   END SELECT
   RETURN wszInfoValue
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION COdbc.GetInfo (BYREF wszInfoType AS WSTRING) AS SQLINTEGER
   DIM info AS SQLINTEGER
   SELECT CASE LCASE(wszInfoType)
      CASE "activeenvironments" : SetResult(SQLGetInfoW(m_hDbc, SQL_ACTIVE_ENVIRONMENTS, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "asyncmode" : SetResult(SQLGetInfoW(m_hDbc, SQL_ASYNC_MODE, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "batchrowcount" : SetResult(SQLGetInfoW(m_hDbc, SQL_BATCH_ROW_COUNT, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "batchsupport" : SetResult(SQLGetInfoW(m_hDbc, SQL_BATCH_SUPPORT, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "driverhenv" : SetResult(SQLGetInfoW(m_hDbc, SQL_DRIVER_HENV, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "driverhdbc" : SetResult(SQLGetInfoW(m_hDbc, SQL_DRIVER_HDBC, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "driverhstmt" : SetResult(SQLGetInfoW(m_hDbc, SQL_DRIVER_HSTMT, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "driverhdesc" : SetResult(SQLGetInfoW(m_hDbc, SQL_DRIVER_HDESC, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "driverhlib" : SetResult(SQLGetInfoW(m_hDbc, SQL_DRIVER_HLIB, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "dynamiccursorattributes1" : SetResult(SQLGetInfoW(m_hDbc, SQL_DYNAMIC_CURSOR_ATTRIBUTES1, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "dynamiccursorattributes2" : SetResult(SQLGetInfoW(m_hDbc, SQL_DYNAMIC_CURSOR_ATTRIBUTES2, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "forwardonlycursorattributes1" : SetResult(SQLGetInfoW(m_hDbc, SQL_FORWARD_ONLY_CURSOR_ATTRIBUTES1, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "forwardonlycursorattributes2" : SetResult(SQLGetInfoW(m_hDbc, SQL_FORWARD_ONLY_CURSOR_ATTRIBUTES2, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "fileusage" : SetResult(SQLGetInfoW(m_hDbc, SQL_FILE_USAGE, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "dataextensions" : SetResult(SQLGetInfoW(m_hDbc, SQL_GETDATA_EXTENSIONS, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "infoschemaviews" : SetResult(SQLGetInfoW(m_hDbc, SQL_INFO_SCHEMA_VIEWS, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "keysetcursorattributes1" : SetResult(SQLGetInfoW(m_hDbc, SQL_KEYSET_CURSOR_ATTRIBUTES1, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "keysetcursorattributes2" : SetResult(SQLGetInfoW(m_hDbc, SQL_KEYSET_CURSOR_ATTRIBUTES2, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "maxasyncconcurrentstatements" : SetResult(SQLGetInfoW(m_hDbc, SQL_MAX_ASYNC_CONCURRENT_STATEMENTS, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "maxconcurrentactivities" : SetResult(SQLGetInfoW(m_hDbc, SQL_MAX_CONCURRENT_ACTIVITIES, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "maxdriverconnections" : SetResult(SQLGetInfoW(m_hDbc, SQL_MAX_DRIVER_CONNECTIONS, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "odbcinterfaceconformance" : SetResult(SQLGetInfoW(m_hDbc, SQL_ODBC_INTERFACE_CONFORMANCE, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "paramarrayrowcounts" : SetResult(SQLGetInfoW(m_hDbc, SQL_PARAM_ARRAY_ROW_COUNTS, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "paramarrayselects" : SetResult(SQLGetInfoW(m_hDbc, SQL_PARAM_ARRAY_SELECTS, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "posoperations" : SetResult(SQLGetInfoW(m_hDbc, SQL_POS_OPERATIONS, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "standardcliconformance" : SetResult(SQLGetInfoW(m_hDbc, SQL_STANDARD_CLI_CONFORMANCE, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "staticcursorattributes1" : SetResult(SQLGetInfoW(m_hDbc, SQL_STATIC_CURSOR_ATTRIBUTES1, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "staticcursorattributes2" : SetResult(SQLGetInfoW(m_hDbc, SQL_STATIC_CURSOR_ATTRIBUTES2, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "bookmarkpersistence" : SetResult(SQLGetInfoW(m_hDbc, SQL_BOOKMARK_PERSISTENCE, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "concatnullbehavior" : SetResult(SQLGetInfoW(m_hDbc, SQL_CONCAT_NULL_BEHAVIOR, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "cursorcommitbehavior" : SetResult(SQLGetInfoW(m_hDbc, SQL_CURSOR_COMMIT_BEHAVIOR, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "cursorrollbackbehavior" : SetResult(SQLGetInfoW(m_hDbc, SQL_CURSOR_ROLLBACK_BEHAVIOR, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "cursorsensitivitysupport" : SetResult(SQLGetInfoW(m_hDbc, SQL_CURSOR_SENSITIVITY, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "defaulttxnisolation" : SetResult(SQLGetInfoW(m_hDbc, SQL_DEFAULT_TXN_ISOLATION, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "nullcollation" : SetResult(SQLGetInfoW(m_hDbc, SQL_NULL_COLLATION, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "scrolloptions" : SetResult(SQLGetInfoW(m_hDbc, SQL_SCROLL_OPTIONS, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "txncapable" : SetResult(SQLGetInfoW(m_hDbc, SQL_TXN_CAPABLE, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "txnisolationoption" : SetResult(SQLGetInfoW(m_hDbc, SQL_TXN_ISOLATION_OPTION, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "aggregatefunctions" : SetResult(SQLGetInfoW(m_hDbc, SQL_AGGREGATE_FUNCTIONS, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "alterdomain" : SetResult(SQLGetInfoW(m_hDbc, SQL_ALTER_DOMAIN, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "altertable" : SetResult(SQLGetInfoW(m_hDbc, SQL_ALTER_TABLE, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "cataloglocation" : SetResult(SQLGetInfoW(m_hDbc, SQL_CATALOG_LOCATION, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "catalogusage" : SetResult(SQLGetInfoW(m_hDbc, SQL_CATALOG_USAGE, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "correlationname" : SetResult(SQLGetInfoW(m_hDbc, SQL_CORRELATION_NAME, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "createassertion" : SetResult(SQLGetInfoW(m_hDbc, SQL_CREATE_ASSERTION, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "createcharacterset" : SetResult(SQLGetInfoW(m_hDbc, SQL_CREATE_CHARACTER_SET, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "createcollation" : SetResult(SQLGetInfoW(m_hDbc, SQL_CREATE_COLLATION, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "createdomain" : SetResult(SQLGetInfoW(m_hDbc, SQL_CREATE_DOMAIN, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "createschema" : SetResult(SQLGetInfoW(m_hDbc, SQL_CREATE_SCHEMA, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "createtable" : SetResult(SQLGetInfoW(m_hDbc, SQL_CREATE_TABLE, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "createtranslation" : SetResult(SQLGetInfoW(m_hDbc, SQL_CREATE_TRANSLATION, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "createview" : SetResult(SQLGetInfoW(m_hDbc, SQL_CREATE_VIEW, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "datetimeliterals" : SetResult(SQLGetInfoW(m_hDbc, SQL_DATETIME_LITERALS, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "ddlindex" : SetResult(SQLGetInfoW(m_hDbc, SQL_DDL_INDEX, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "dropassertion" : SetResult(SQLGetInfoW(m_hDbc, SQL_DROP_ASSERTION, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "dropcharacterset" : SetResult(SQLGetInfoW(m_hDbc, SQL_DROP_CHARACTER_SET, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "dropcollation" : SetResult(SQLGetInfoW(m_hDbc, SQL_DROP_COLLATION, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "dropdomain" : SetResult(SQLGetInfoW(m_hDbc, SQL_DROP_DOMAIN, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "dropschema" : SetResult(SQLGetInfoW(m_hDbc, SQL_DROP_SCHEMA, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "droptable" : SetResult(SQLGetInfoW(m_hDbc, SQL_DROP_TABLE, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "droptranslation" : SetResult(SQLGetInfoW(m_hDbc, SQL_DROP_TRANSLATION, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "dropview" : SetResult(SQLGetInfoW(m_hDbc, SQL_DROP_VIEW, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "groupby" : SetResult(SQLGetInfoW(m_hDbc, SQL_GROUP_BY, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "identifiercase" : SetResult(SQLGetInfoW(m_hDbc, SQL_IDENTIFIER_CASE, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "indexkeywords" : SetResult(SQLGetInfoW(m_hDbc, SQL_INDEX_KEYWORDS, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "insertstatement" : SetResult(SQLGetInfoW(m_hDbc, SQL_INSERT_STATEMENT, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "nonnullablecolumns" : SetResult(SQLGetInfoW(m_hDbc, SQL_NON_NULLABLE_COLUMNS, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "sqlconformance" : SetResult(SQLGetInfoW(m_hDbc, SQL_SQL_CONFORMANCE, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "ojcapabilities" : SetResult(SQLGetInfoW(m_hDbc, SQL_OJ_CAPABILITIES, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "quotedidentifiercase" : SetResult(SQLGetInfoW(m_hDbc, SQL_QUOTED_IDENTIFIER_CASE, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "schemausage" : SetResult(SQLGetInfoW(m_hDbc, SQL_SCHEMA_USAGE, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "subqueries" : SetResult(SQLGetInfoW(m_hDbc, SQL_SUBQUERIES, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "union" : SetResult(SQLGetInfoW(m_hDbc, SQL_UNION, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "maxbinaryliterallen" : SetResult(SQLGetInfoW(m_hDbc, SQL_MAX_BINARY_LITERAL_LEN, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "maxcatalognamelen" : SetResult(SQLGetInfoW(m_hDbc, SQL_MAX_CATALOG_NAME_LEN, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "maxcharliterallen" : SetResult(SQLGetInfoW(m_hDbc, SQL_MAX_CHAR_LITERAL_LEN, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "maxcolumnnamelen" : SetResult(SQLGetInfoW(m_hDbc, SQL_MAX_COLUMN_NAME_LEN, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "maxcolumnsingroupby" : SetResult(SQLGetInfoW(m_hDbc, SQL_MAX_COLUMNS_IN_GROUP_BY, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "maxcolumnsinindex" : SetResult(SQLGetInfoW(m_hDbc, SQL_MAX_COLUMNS_IN_INDEX, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "maxcolumnsinorderby" : SetResult(SQLGetInfoW(m_hDbc, SQL_MAX_COLUMNS_IN_ORDER_BY, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "maxcolumnsinselect" : SetResult(SQLGetInfoW(m_hDbc, SQL_MAX_COLUMNS_IN_SELECT, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "maxcolumnsintable" : SetResult(SQLGetInfoW(m_hDbc, SQL_MAX_COLUMNS_IN_TABLE, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "maxcursornamelen" : SetResult(SQLGetInfoW(m_hDbc, SQL_MAX_CURSOR_NAME_LEN, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "maxidentifierlen" : SetResult(SQLGetInfoW(m_hDbc, SQL_MAX_IDENTIFIER_LEN, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "maxindexsize" : SetResult(SQLGetInfoW(m_hDbc, SQL_MAX_INDEX_SIZE, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "maxprocedurenamelen" : SetResult(SQLGetInfoW(m_hDbc, SQL_MAX_PROCEDURE_NAME_LEN, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "maxrowsize" : SetResult(SQLGetInfoW(m_hDbc, SQL_MAX_ROW_SIZE, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "maxschemanamelen" : SetResult(SQLGetInfoW(m_hDbc, SQL_MAX_SCHEMA_NAME_LEN, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "maxstatementlen" : SetResult(SQLGetInfoW(m_hDbc, SQL_MAX_STATEMENT_LEN, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "maxtablenamelen" : SetResult(SQLGetInfoW(m_hDbc, SQL_MAX_TABLE_NAME_LEN, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "maxtablesinselect" : SetResult(SQLGetInfoW(m_hDbc, SQL_MAX_TABLES_IN_SELECT, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "maxusernamelen" : SetResult(SQLGetInfoW(m_hDbc, SQL_MAX_USER_NAME_LEN, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "convertfunctions" : SetResult(SQLGetInfoW(m_hDbc, SQL_CONVERT_FUNCTIONS, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "numericfunctions" : SetResult(SQLGetInfoW(m_hDbc, SQL_NUMERIC_FUNCTIONS, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "stringfunctions" : SetResult(SQLGetInfoW(m_hDbc, SQL_STRING_FUNCTIONS, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "systemfunctions" : SetResult(SQLGetInfoW(m_hDbc, SQL_SYSTEM_FUNCTIONS, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "timedateaddintervals" : SetResult(SQLGetInfoW(m_hDbc, SQL_TIMEDATE_ADD_INTERVALS, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "timedatediffintervals" : SetResult(SQLGetInfoW(m_hDbc, SQL_TIMEDATE_DIFF_INTERVALS, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "timedatefunctions" : SetResult(SQLGetInfoW(m_hDbc, SQL_TIMEDATE_FUNCTIONS, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "convertbigint" : SetResult(SQLGetInfoW(m_hDbc, SQL_CONVERT_BIGINT, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "convertbinary" : SetResult(SQLGetInfoW(m_hDbc, SQL_CONVERT_BINARY, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "convertbit" : SetResult(SQLGetInfoW(m_hDbc, SQL_CONVERT_BIT, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "convertchar" : SetResult(SQLGetInfoW(m_hDbc, SQL_CONVERT_CHAR, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "convertdate" : SetResult(SQLGetInfoW(m_hDbc, SQL_CONVERT_DATE, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "convertdecimal" : SetResult(SQLGetInfoW(m_hDbc, SQL_CONVERT_DECIMAL, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "convertdouble" : SetResult(SQLGetInfoW(m_hDbc, SQL_CONVERT_DOUBLE, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "convertfloat" : SetResult(SQLGetInfoW(m_hDbc, SQL_CONVERT_FLOAT, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "convertinteger" : SetResult(SQLGetInfoW(m_hDbc, SQL_CONVERT_INTEGER, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "convertintervalyearmonth" : SetResult(SQLGetInfoW(m_hDbc, SQL_CONVERT_INTERVAL_YEAR_MONTH, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "convertintervaldaytime" : SetResult(SQLGetInfoW(m_hDbc, SQL_CONVERT_INTERVAL_DAY_TIME, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "convertlongvarbinary" : SetResult(SQLGetInfoW(m_hDbc, SQL_CONVERT_LONGVARBINARY, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "convertlongvarchar" : SetResult(SQLGetInfoW(m_hDbc, SQL_CONVERT_LONGVARCHAR, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "convertnumeric" : SetResult(SQLGetInfoW(m_hDbc, SQL_CONVERT_NUMERIC, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "convertreal" : SetResult(SQLGetInfoW(m_hDbc, SQL_CONVERT_REAL, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "convertsmallint" : SetResult(SQLGetInfoW(m_hDbc, SQL_CONVERT_SMALLINT, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "converttime" : SetResult(SQLGetInfoW(m_hDbc, SQL_CONVERT_TIME, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "converttimestamp" : SetResult(SQLGetInfoW(m_hDbc, SQL_CONVERT_TIMESTAMP, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "converttinyint" : SetResult(SQLGetInfoW(m_hDbc, SQL_CONVERT_TINYINT, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "convertvarbinary" : SetResult(SQLGetInfoW(m_hDbc, SQL_CONVERT_VARBINARY, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "convertvarchar" : SetResult(SQLGetInfoW(m_hDbc, SQL_CONVERT_VARCHAR, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "sql92datetimefunctions" : SetResult(SQLGetInfoW(m_hDbc, SQL_SQL92_DATETIME_FUNCTIONS, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "sql92foreignkeydeleterule" : SetResult(SQLGetInfoW(m_hDbc, SQL_SQL92_FOREIGN_KEY_DELETE_RULE, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "sql92foreignkeyupdaterule" : SetResult(SQLGetInfoW(m_hDbc, SQL_SQL92_FOREIGN_KEY_UPDATE_RULE, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "sql92grant" : SetResult(SQLGetInfoW(m_hDbc, SQL_SQL92_GRANT, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "sql92numericvaluefunctions" : SetResult(SQLGetInfoW(m_hDbc, SQL_SQL92_NUMERIC_VALUE_FUNCTIONS, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "sql92predicates" : SetResult(SQLGetInfoW(m_hDbc, SQL_SQL92_PREDICATES, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "sql92relationaljoinoperators" : SetResult(SQLGetInfoW(m_hDbc, SQL_SQL92_RELATIONAL_JOIN_OPERATORS, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "sql92revoke" : SetResult(SQLGetInfoW(m_hDbc, SQL_SQL92_REVOKE, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "sql92rowvaluevonstructor" : SetResult(SQLGetInfoW(m_hDbc, SQL_SQL92_ROW_VALUE_CONSTRUCTOR, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "sql92stringfunctions" : SetResult(SQLGetInfoW(m_hDbc, SQL_SQL92_STRING_FUNCTIONS, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
      CASE "sql92valueexpressions" : SetResult(SQLGetInfoW(m_hDbc, SQL_SQL92_VALUE_EXPRESSIONS, @info, 4, NULL), "COdbc.GetInfo", SQL_HANDLE_DBC, m_hDbc)
   END SELECT
   RETURN info
END FUNCTION
' ========================================================================================

