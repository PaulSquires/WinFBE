' ########################################################################################
' Microsoft Windows
' File: CODBC.inc
' Contents: ODBC wrapper classes (unicode)
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "windows.bi"
#include once "win/sql.bi"
#include once "win/sqlucode.bi"
#include once "Afx/CWStr.inc"

NAMESPACE Afx

'// new values for SQL_ATTR_ODBC_VERSION
'// From ODBC 3.8 onwards, we should use <major version> * 100 + <minor version>
#ifndef SQL_OV_ODBC3_80
#define SQL_OV_ODBC3_80 380u
#endif

' ========================================================================================
' Macro for debug
' To allow debugging, define _CBODBC_DEBUG_ 1 in your application before including this file.
' ========================================================================================
#ifndef _CODBC_DEBUG_
   #define _CODBC_DEBUG_ 0
#ENDIF
#ifndef _CODBC_DP_
   #define _CODBC_DP_ 1
   #MACRO CODBC_DP(st)
      #IF (_CODBC_DEBUG_ = 1)
         OutputDebugStringW(st)
      #ENDIF
   #ENDMACRO
#ENDIF
' ========================================================================================

DIM AFX_ODBC_hEnv AS SQLHANDLE        ' Environment handle
DIM AFX_ODBC_numConnections AS LONG   ' Number of connections

' ########################################################################################
' COdbcBase - ODBC base class
' ########################################################################################
TYPE COdbcBase

   m_Result AS SQLRETURN   ' Last return code
   m_ErrorProc AS SUB (BYVAL nResult AS SQLRETURN, BYREF wszSrc AS WSTRING, BYREF wszErrorMsg AS WSTRING)
   m_Warnings AS BOOLEAN   ' Send warnings to the error callback

   DECLARE FUNCTION GetLastResult () AS SQLRETURN
   DECLARE FUNCTION SetResult (BYVAL Result AS SQLRETURN, BYREF wszSrc AS WSTRING = "", BYVAL HandleType AS SQLSMALLINT = 0, BYVAL Handle AS SQLHANDLE = NULL) AS SQLRETURN
   DECLARE FUNCTION Error () AS BOOLEAN
   DECLARE SUB SetErrorProc (BYVAL pProc AS ANY PTR, BYVAL reportWarnings AS BOOLEAN = FALSE)
   DECLARE FUNCTION GetErrorInfo (BYVAL HandleType AS SQLSMALLINT, BYVAL Handle AS SQLHANDLE, BYVAL iErrorCode AS SQLRETURN = 0) AS CWSTR
   DECLARE FUNCTION GetSqlState (BYVAL HandleType AS SQLSMALLINT, BYVAL Handle AS SQLHANDLE) AS CWSTR
   DECLARE FUNCTION ODBCVersion () AS SQLUINTEGER
   DECLARE FUNCTION GetEnvAttr (BYVAL Attribute AS SQLINTEGER, BYVAL ValuePtr AS SQLPOINTER, BYVAL BufferLength AS SQLINTEGER, BYVAL StringLength AS SQLINTEGER PTR) AS SQLRETURN
   DECLARE FUNCTION SetEnvAttr (BYVAL Attribute AS SQLINTEGER, BYVAL ValuePtr AS SQLPOINTER, BYVAL StringLength AS SQLINTEGER) AS SQLRETURN
   DECLARE FUNCTION GetCPMatch () AS SQLUINTEGER
   DECLARE FUNCTION SetCPMatch (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
   DECLARE FUNCTION GetOutputNTS () AS SQLUINTEGER
   DECLARE FUNCTION SetOutputNTS (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
   DECLARE FUNCTION GetDataSources (BYVAL Direction AS SQLUSMALLINT, BYREF cwsServerName AS CWSTR, BYREF cwsDescription AS CWSTR) AS SQLRETURN
   DECLARE FUNCTION GetDrivers (BYVAL Direction AS SQLUSMALLINT, BYREF cwsDriverDesc AS CWSTR, BYREF cwsDriverAttributes AS CWSTR) AS SQLRETURN

END TYPE
' ########################################################################################

' ========================================================================================
' Returns the last status code.
' ========================================================================================
PRIVATE FUNCTION COdbcBase.GetLastResult () AS SQLRETURN
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the last status code.
' ========================================================================================
PRIVATE FUNCTION COdbcBase.SetResult (BYVAL Result AS SQLRETURN, BYREF wszSrc AS WSTRING = "", _
BYVAL HandleType AS SQLSMALLINT = 0, BYVAL Handle AS SQLHANDLE = NULL) AS SQLRETURN
   m_Result = Result
   IF m_ErrorProc THEN
      IF m_Result = SQL_ERROR OR m_Result = SQL_INVALID_HANDLE THEN
         m_ErrorProc(m_Result, wszSrc, this.GetErrorInfo(HandleType, Handle, m_Result))
      ELSEIF m_Result = SQL_SUCCESS_WITH_INFO THEN
         IF m_Warnings THEN m_ErrorProc(m_Result, wszSrc, this.GetErrorInfo(HandleType, Handle, m_Result))
      END IF
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' TRUE if there has been an error; FALSE, otherwise.
' ========================================================================================
PRIVATE FUNCTION COdbcBase.Error () AS BOOLEAN
   IF m_Result = SQL_ERROR OR m_Result = SQL_INVALID_HANDLE THEN RETURN TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the address of an application defined error callback
' Example of error callbak:
' SUB ODBC_ErrorCallback (BYVAL nResult AS SQLRETURN, BYREF wszSrc AS WSTRING, BYREF wszErrorMsg AS WSTRING)
'   PRINT "Error: " & WSTR(nResult) & " - Source: " & wszSrc
'   IF LEN(wszErrorMsg) THEN PRINT wszErrorMsg
' END SUB
' pDbc.SetErrorProc(@ODBC_ErrorCallback)   ' // Sets the error callback for the connection object
' pStmt->SetErrorProc(@ODBC_ErrorCallback)   ' // Sets the error callback for the statement object
' ========================================================================================
PRIVATE SUB COdbcBase.SetErrorProc (BYVAL pProc AS ANY PTR, BYVAL reportWarnings AS BOOLEAN = FALSE)
   m_ErrorProc = pProc
   m_Warnings = reportWarnings
END SUB
' ========================================================================================

' ===========================================================================================
' Uses SQLGetDiagRec to retrieve an error description.
' iErrorCode: Optional. The error code returned by some of the methods of this interface.
' ===========================================================================================
PRIVATE FUNCTION COdbcBase.GetErrorInfo (BYVAL HandleType AS SQLSMALLINT, BYVAL Handle AS SQLHANDLE, BYVAL iErrorCode AS SQLRETURN = 0) AS CWSTR

   DIM iResult AS SQLRETURN
   DIM lNativeError AS SQLINTEGER
   DIM cbbytes AS SQLSMALLINT
   DIM ErrorCount AS SQLINTEGER
   DIM wszSqlState AS WSTRING * 6
   DIM wszErrMsg AS WSTRING * SQL_MAX_MESSAGE_LENGTH + 1
   DIM cwsErrMsg AS CWSTR

   IF HandleType <> 0 AND Handle <> NULL THEN
      iResult = SQLGetDiagFieldW(HandleType, Handle, 0, SQL_DIAG_NUMBER, cast(SQLPOINTER, @ErrorCount), SQL_IS_INTEGER, @cbbytes)
   END IF
   IF ErrorCount THEN
      FOR i AS LONG = 1 TO ErrorCount
         iResult = SQLGetDiagRecW(HandleType, Handle, 1, @wszSqlState, @lNativeError, @wszErrMsg, SIZEOF(wszErrMsg), @cbbytes)
         IF iResult <> SQL_SUCCESS AND iResult <> SQL_SUCCESS_WITH_INFO THEN EXIT FOR
         cwsErrMsg = wszErrMsg & CHR(13, 10) & "SqlState: " & wszSqlState & "; Native error: " & STR(lNativeError) & CHR(13, 10)
      NEXT
   ELSEIF iErrorCode THEN
      SELECT CASE iErrorCode
         CASE  -1 : cwsErrMsg = "SQL error" ' "SQL_ERROR"
         CASE   2 : cwsErrMsg = "Still executing" ' "SQL_STILL_EXECUTING"
         CASE  -2 : cwsErrMsg = "Invalid handle" '"SQL_INVALID_HANDLE"
         CASE  99 : cwsErrMsg = "Need data" ' "SQL_NEED_DATA"
         CASE 100 : cwsErrMsg = "No data" '"SQL_NO_DATA"
         CASE ELSE
            cwsErrMsg = "Error " & STR(iErrorCode)
      END SELECT
   END IF

   RETURN cwsErrMsg

END FUNCTION
' ===========================================================================================

' ========================================================================================
' Gets the SqlState for the specified handle
' ========================================================================================
PRIVATE FUNCTION COdbcBase.GetSqlState (BYVAL HandleType AS SQLSMALLINT, BYVAL Handle AS SQLHANDLE) AS CWSTR
   DIM lNativeError AS SQLINTEGER, cbbytes AS SQLSMALLINT
   DIM wszSqlState AS WSTRING * 6, wszErrMsg AS WSTRING * SQL_MAX_MESSAGE_LENGTH + 1
   SetResult(SQLGetDiagRecW(HandleType, Handle, 1, @wszSqlState, @lNativeError, @wszErrMsg, _
      SIZEOF(wszErrMsg), @cbbytes), "COdbcBase.GetSqlState", HandleType, Handle)
   RETURN wszSqlState
END FUNCTION
' ========================================================================================

' ========================================================================================
PRIVATE FUNCTION COdbcBase.ODBCVersion () AS SQLUINTEGER
   DIM dwAttr AS DWORD
   SetResult(SQLGetEnvAttr(AFX_ODBC_hEnv, SQL_ATTR_ODBC_VERSION, @dwAttr, SQL_IS_UINTEGER, NULL))
   RETURN dwAttr
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns or sets the current setting of an environment attribute.
' ========================================================================================
PRIVATE FUNCTION COdbcBase.GetEnvAttr (BYVAL Attribute AS SQLINTEGER, BYVAL ValuePtr AS SQLPOINTER, BYVAL BufferLength AS SQLINTEGER, BYVAL StringLength AS SQLINTEGER PTR) AS SQLRETURN
   RETURN SetResult(SQLGetEnvAttr(AFX_ODBC_hEnv, Attribute, ValuePtr, BufferLength, StringLength))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION COdbcBase.SetEnvAttr (BYVAL Attribute AS SQLINTEGER, BYVAL ValuePtr AS SQLPOINTER, BYVAL StringLength AS SQLINTEGER) AS SQLRETURN
   RETURN SetResult(SQLSetEnvAttr(AFX_ODBC_hEnv, Attribute, BYVAL ValuePtr, StringLength))
END FUNCTION
' ========================================================================================

' ========================================================================================
PRIVATE FUNCTION COdbcBase.GetCPMatch () AS SQLUINTEGER
   DIM dwAttr AS SQLUINTEGER
   SetResult(SQLGetEnvAttr(AFX_ODBC_hEnv, SQL_ATTR_CP_MATCH, @dwAttr, SQL_IS_UINTEGER, NULL))
   RETURN dwAttr
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION COdbcBase.SetCPMatch (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
   RETURN SetResult(SQLSetEnvAttr(AFX_ODBC_hEnv, SQL_ATTR_CP_MATCH, cast(SQLPOINTER, cast(LONG_PTR, dwAttr)), SQL_IS_UINTEGER))
END FUNCTION
' ========================================================================================

' ========================================================================================
PRIVATE FUNCTION COdbcBase.GetOutputNTS () AS SQLUINTEGER
   DIM dwAttr AS SQLUINTEGER
   SetResult(SQLGetEnvAttr(AFX_ODBC_hEnv, SQL_ATTR_OUTPUT_NTS, @dwAttr, SQL_IS_UINTEGER, NULL))
   RETURN dwAttr
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION COdbcBase.SetOutputNTS (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
   RETURN SetResult(SQLSetEnvAttr(AFX_ODBC_hEnv, SQL_ATTR_OUTPUT_NTS, cast(SQLPOINTER, cast(LONG_PTR, dwAttr)), SQL_IS_UINTEGER))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Lists available DSN / Drivers installed.
' ========================================================================================
PRIVATE FUNCTION COdbcBase.GetDataSources (BYVAL Direction AS SQLUSMALLINT, BYREF cwsServerName AS CWSTR, BYREF cwsDescription AS CWSTR) AS SQLRETURN
   DIM wszServerName AS WSTRING * SQL_MAX_DSN_LENGTH + 1
   DIM wszDescription AS WSTRING * 255
   DIM cbServerNameLen AS SQLUSMALLINT
   DIM cbDescription AS SQLUSMALLINT
   DIM cbDescriptionLen AS SQLUSMALLINT
   FUNCTION = SetResult(SQLDataSourcesW(AFX_ODBC_hEnv, Direction, @wszServerName, SQL_MAX_DSN_LENGTH + 1, _
      @cbServerNameLen, @wszDescription, 255, @cbDescriptionLen))
   cwsServerName = wszServerName
   cwsDescription = wszDescription
END FUNCTION
' ========================================================================================

' ========================================================================================
' Lists driver descriptions and driver attribute keywords. This function is implemented
' solely by the driver manager.
' ========================================================================================
PRIVATE FUNCTION COdbcBase.GetDrivers (BYVAL Direction AS SQLUSMALLINT, BYREF cwsDriverDesc AS CWSTR, BYREF cwsDriverAttributes AS CWSTR) AS SQLRETURN
   DIM wszDriverAttributes AS WSTRING * 255
   DIM wszDriverDesc AS WSTRING * 255
   DIM cbDriverDesc AS SQLSMALLINT
   DIM cbDrvrAttrLen AS SQLSMALLINT
   FUNCTION = SetResult(SQLDriversW(AFX_ODBC_hEnv, Direction, @wszDriverDesc, 255, @cbDriverDesc, _
       @wszDriverAttributes, 255, @cbDrvrAttrLen))
   DIM pwszDriverAttributes AS WSTRING PTR = @wszDriverAttributes
   FOR i AS LONG = 0 TO cbDrvrAttrLen - 1
      IF pwszDriverAttributes[i] = 0 THEN pwszDriverAttributes[i] = ASC(";")
   NEXT
   DIM nPos AS LONG
   IF INSTR(wszDriverAttributes, ";;") THEN wszDriverAttributes = LEFT(wszDriverAttributes, nPos)
   cwsDriverDesc = wszDriverDesc
   cwsDriverAttributes = wszDriverAttributes
END FUNCTION
' ========================================================================================


' ########################################################################################
' COdbc class
' ########################################################################################
TYPE COdbc EXTENDS COdbcBase

   m_hDbc AS SQLHDBC   ' Connection handle

   DECLARE CONSTRUCTOR (BYREF wszConnectionString AS WSTRING, BYVAL nODbcVersion AS SQLINTEGER = SQL_OV_ODBC3, BYVAL ConnectionPoolingAttr AS SQLUINTEGER = 0)
   DECLARE CONSTRUCTOR (BYREF wszServerName AS WSTRING, BYREF wszUserName AS WSTRING, BYREF wszAuthentication AS WSTRING, BYVAL nODbcVersion AS SQLINTEGER = SQL_OV_ODBC3, BYVAL ConnectionPoolingAttr AS SQLUINTEGER = 0)
   DECLARE DESTRUCTOR
   DECLARE FUNCTION Handle () AS SQLHANDLE
   DECLARE FUNCTION EnvHandle () AS SQLHANDLE
   DECLARE FUNCTION GetDiagField (BYVAL RecNumber AS SQLSMALLINT, BYVAL DiagIdentifier AS SQLSMALLINT, _
           BYVAL DiagInfoPtr AS SQLPOINTER, BYVAL BufferLength AS SQLSMALLINT, BYVAL StringLengthPtr AS SQLSMALLINT PTR) AS SQLRETURN
   DECLARE FUNCTION GetDiagRec (BYVAL RecNumber AS SQLSMALLINT, BYVAL Sqlstate AS WSTRING PTR, BYVAL NativeError AS SQLINTEGER PTR, _
          BYVAL MessageText AS WSTRING PTR, BYVAL BufferLength AS SQLSMALLINT, BYVAL TextLength AS SQLSMALLINT PTR) AS SQLRETURN
   DECLARE FUNCTION GetErrorInfo (BYVAL iErrorCode AS SQLRETURN = 0) AS CWSTR
   DECLARE FUNCTION GetSqlState () AS CWSTR
   DECLARE FUNCTION NativeSql (BYREF wszInText AS WSTRING) AS CWSTR
   DECLARE FUNCTION Functions (BYVAL FunctionId AS SQLUSMALLINT) AS BOOLEAN
   DECLARE FUNCTION Supports (BYVAL FunctionId AS SQLUSMALLINT) AS BOOLEAN
   DECLARE FUNCTION CommitTran () AS SQLRETURN
   DECLARE FUNCTION RollbackTran () AS SQLRETURN
   DECLARE FUNCTION GetConnectAttr (BYVAL Attribute AS SQLINTEGER, BYVAL ValuePtr AS SQLPOINTER, BYVAL _
           BufferLength AS SQLINTEGER, BYVAL StringLength AS SQLINTEGER PTR) AS SQLRETURN
   DECLARE FUNCTION SetConnectAttr (BYVAL Attribute AS SQLINTEGER, BYVAL ValuePtr AS SQLPOINTER, BYVAL StringLength AS SQLINTEGER) AS SQLRETURN
   DECLARE FUNCTION GetInfo (BYVAL InfoType AS SQLUSMALLINT, BYVAL InfoValuePtr AS SQLPOINTER, _
           BYVAL BufferLength AS SQLSMALLINT, BYVAL StringLength AS SQLSMALLINT PTR) AS SQLRETURN
   DECLARE FUNCTION GetConnectAttr (BYREF wszAttribute AS WSTRING) AS SQLUINTEGER
   DECLARE FUNCTION GetConnectAttrStr (BYREF wszAttribute AS WSTRING) AS CWSTR
   DECLARE FUNCTION SetConnectAttr (BYREF wszAttribute AS WSTRING, BYVAL dwAttribute AS SQLUINTEGER) AS SQLRETURN
   DECLARE FUNCTION SetConnectAttr (BYREF wszAttribute AS WSTRING, BYREF wszAttrValue AS WSTRING) AS SQLRETURN
   DECLARE FUNCTION GetInfo (BYREF wszInfoType AS WSTRING) AS SQLINTEGER
   DECLARE FUNCTION GetInfoStr (BYREF wszInfoType AS WSTRING) AS CWSTR

END TYPE
' ########################################################################################

' ########################################################################################
' COdbcStmt class
' ########################################################################################
TYPE COdbcStmt EXTENDS COdbcBase

   m_hDbc  AS SQLHDBC    ' Connection handle
   m_hStmt AS SQLHSTMT   ' Statement handle
   DIM rgColNames(ANY) AS WSTRING * 255   ' Array of column names

   DECLARE CONSTRUCTOR (BYVAL pDbc AS COdbc PTR)
   DECLARE DESTRUCTOR
   DECLARE FUNCTION Handle () AS SQLHANDLE
   DECLARE FUNCTION DbcHandle () AS SQLHANDLE
   DECLARE FUNCTION GetDiagField (BYVAL RecNumber AS SQLSMALLINT, BYVAL DiagIdentifier AS SQLSMALLINT, _
           BYVAL DiagInfoPtr AS SQLPOINTER, BYVAL BufferLength AS SQLSMALLINT, BYVAL StringLengthPtr AS SQLSMALLINT PTR) AS SQLRETURN
   DECLARE FUNCTION GetDiagRec (BYVAL RecNumber AS SQLSMALLINT, BYVAL Sqlstate AS WSTRING PTR, BYVAL NativeError AS SQLINTEGER PTR, _
          BYVAL MessageText AS WSTRING PTR, BYVAL BufferLength AS SQLSMALLINT, BYVAL TextLength AS SQLSMALLINT PTR) AS SQLRETURN
   DECLARE FUNCTION GetErrorInfo (BYVAL iErrorCode AS SQLRETURN = 0) AS CWSTR
   DECLARE FUNCTION GetSqlState () AS CWSTR
   DECLARE FUNCTION ExecDirect (BYREF wszSqlStr AS WSTRING) AS SQLRETURN
   DECLARE FUNCTION RecordCount (BYREF wszSqlStr AS WSTRING) AS SQLINTEGER
   DECLARE FUNCTION Cancel () AS SQLRETURN
   DECLARE FUNCTION NumResultCols () AS SQLSMALLINT
   DECLARE FUNCTION MoreResults () AS SQLRETURN
   DECLARE FUNCTION Fetch () AS BOOLEAN
   DECLARE FUNCTION GetData (BYVAL ColumnNumber AS SQLUSMALLINT, BYVAL TargetType AS SQLSMALLINT, _
                    BYVAL TargetValue AS SQLPOINTER, BYVAL BufferLength AS SQLLEN, BYVAL StrLen_or_Ind AS SQLLEN PTR) AS SQLRETURN
   DECLARE FUNCTION GetData (BYREF ColumnName AS WSTRING, BYVAL TargetType AS SQLSMALLINT, _
                    BYVAL TargetValue AS SQLPOINTER, BYVAL BufferLength AS SQLLEN, BYVAL StrLen_or_Ind AS SQLLEN PTR) AS SQLRETURN
   DECLARE FUNCTION GetData (BYVAL ColumnNumber AS SQLUSMALLINT, BYVAL nMaxChars AS LONG = 256 * 2) AS CWSTR
   DECLARE FUNCTION GetData (BYREF ColumnName AS WSTRING, BYVAL nMaxChars AS LONG = 256 * 2) AS CWSTR
   DECLARE FUNCTION GetLongVarcharData (BYVAL ColumnNumber AS SQLSMALLINT) AS CWSTR
   DECLARE FUNCTION GetLongVarcharData (BYREF ColumnName AS WSTRING) AS CWSTR
   DECLARE FUNCTION PutData (BYVAL pData AS SQLPOINTER, BYVAL StrLen_or_Ind AS SQLLEN) AS SQLRETURN
   DECLARE FUNCTION ParamData (BYVAL Value AS SQLPOINTER PTR) AS SQLRETURN
   DECLARE FUNCTION ResetParams () AS SQLRETURN
   DECLARE FUNCTION Prepare (BYREF StatementText AS WSTRING) AS SQLRETURN
   DECLARE FUNCTION Execute () AS SQLRETURN
   DECLARE FUNCTION DescribeCol (BYVAL ColumnNumber AS SQLUSMALLINT, BYVAL pwszColumnName AS WSTRING PTR, _
                    BYVAL BufferLength AS SQLSMALLINT, BYVAL NameLengthPtr AS SQLSMALLINT PTR, BYVAL DataTypePtr AS SQLSMALLINT PTR, _
                    BYVAL ColumnSizePtr AS SQLULEN PTR, BYVAL DecimalDigitsPtr AS SQLSMALLINT PTR, BYVAL Nullable AS SQLSMALLINT PTR) AS SQLRETURN
   DECLARE FUNCTION NumParams () AS SQLSMALLINT
   DECLARE FUNCTION DescribeParam (BYVAL ParameterNumber AS SQLUSMALLINT, BYVAL DataTypePtr AS SQLSMALLINT PTR, _
                    BYVAL ParameterSizePtr AS SQLULEN PTR, BYVAL DecimalDigitsPtr AS SQLSMALLINT PTR, BYVAL NullablePtr AS SQLSMALLINT PTR) AS SQLRETURN
   DECLARE FUNCTION ColIsNull (BYVAL ColumnNumber AS SQLUSMALLINT) AS BOOLEAN
   DECLARE FUNCTION UnbindCol (BYVAL nCol AS SQLUSMALLINT) AS SQLRETURN
   DECLARE FUNCTION UnbindColumns () AS SQLRETURN
   DECLARE FUNCTION CloseCursor () AS SQLRETURN
   DECLARE FUNCTION GetCursorName () AS CWSTR
   DECLARE FUNCTION SetCursorName (BYREF wszCursorName AS WSTRING) AS SQLRETURN
   DECLARE FUNCTION MoveFirst () AS SQLRETURN
   DECLARE FUNCTION MoveLast () AS SQLRETURN
   DECLARE FUNCTION MoveNext () AS SQLRETURN
   DECLARE FUNCTION MovePrevious () AS SQLRETURN
   DECLARE FUNCTION BulkOperations (BYVAL Operation AS SQLSMALLINT) AS SQLRETURN
   DECLARE FUNCTION AddRecord () AS SQLRETURN
   DECLARE FUNCTION UpdateByBookmark () AS SQLRETURN
   DECLARE FUNCTION DeleteByBookmark () AS SQLRETURN
   DECLARE FUNCTION FetchByBookmark () AS SQLRETURN
   DECLARE FUNCTION RowCount () AS SQLLEN
   DECLARE FUNCTION BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetType AS SQLSMALLINT, _
           BYVAL TargetValue AS SQLPOINTER, BYVAL BufferLength AS SQLLEN, BYVAL StrLen_or_Ind AS ANY PTR) AS SQLRETURN
   DECLARE FUNCTION BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS BYTE PTR, BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
   DECLARE FUNCTION BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS UBYTE PTR, BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
   DECLARE FUNCTION BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS SHORT PTR, BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
   DECLARE FUNCTION BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS USHORT PTR, BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
   DECLARE FUNCTION BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS LONG PTR, BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
   DECLARE FUNCTION BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS ULONG PTR, BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
   DECLARE FUNCTION BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS SINGLE PTR, BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
   DECLARE FUNCTION BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS DOUBLE PTR, BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
   DECLARE FUNCTION BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS LONGINT PTR, BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
   DECLARE FUNCTION BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS ULONGINT PTR, BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
   DECLARE FUNCTION BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS ZSTRING PTR, BYVAL BufferLenght AS SQLLEN, BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
   DECLARE FUNCTION BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS WSTRING PTR, BYVAL BufferLenght AS SQLLEN, BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
   DECLARE FUNCTION BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS BYTE PTR, BYVAL BufferLenght AS SQLLEN, BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
   DECLARE FUNCTION BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS DATE_STRUCT PTR, BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
   DECLARE FUNCTION BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS TIME_STRUCT PTR, BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
   DECLARE FUNCTION BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS TIMESTAMP_STRUCT PTR, BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
   DECLARE FUNCTION BindCol (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS ANY PTR, BYVAL BufferLenght AS SQLLEN, BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
   DECLARE FUNCTION BindColToBit (BYVAL ColNumber AS SQLUSMALLINT, BYVAL TargetValue AS SHORT PTR, BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
   DECLARE FUNCTION BindColToNumeric (BYVAL ColNumber AS SQLUSMALLINT, BYREF TargetValue AS ZSTRING, BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
   DECLARE FUNCTION BindColToDecimal (BYVAL ColNumber AS SQLUSMALLINT, BYREF TargetValue AS ZSTRING, BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
   DECLARE FUNCTION BindParameter (BYVAL ParameterNumber AS SQLUSMALLINT, BYVAL InputOutputType AS SQLSMALLINT, _
           BYVAL ValueType AS SQLSMALLINT, BYVAL ParameterType AS SQLSMALLINT, BYVAL ColumnSize AS SQLULEN, BYVAL DecimalDigits AS SQLSMALLINT, _
           BYVAL ParameterValuePtr AS SQLPOINTER, BYVAL BufferLength AS SQLLEN, BYVAL StrLen_or_IndPtr AS ANY PTR) AS SQLRETURN
   DECLARE FUNCTION FetchScroll (BYVAL FetchOrientation AS SQLSMALLINT, BYVAL FetchOffset AS SQLLEN) AS SQLRETURN
   DECLARE FUNCTION ExtendedFetch (BYVAL FetchOrientation AS SQLUSMALLINT, BYVAL FetchOffset AS SQLLEN, BYVAL RowCountPtr AS ANY PTR, BYVAL RowStatusArray AS SQLUSMALLINT PTR) AS SQLRETURN
   DECLARE FUNCTION Move (BYVAL FetchOffset AS SQLLEN) AS SQLRETURN
   DECLARE FUNCTION SetAbsolutePosition (BYVAL FetchOffset AS SQLLEN) AS SQLRETURN
   DECLARE FUNCTION SetRelativePosition (BYVAL FetchOffset AS SQLLEN) AS SQLRETURN
   DECLARE FUNCTION SetPos (BYVAL wRow AS SQLSETPOSIROW, BYVAL fOption AS SQLUSMALLINT, BYVAL fLock AS SQLUSMALLINT) AS SQLRETURN
   DECLARE FUNCTION SetPosition (BYVAL wRow AS SQLSETPOSIROW) AS SQLRETURN
   DECLARE FUNCTION UpdateRecord (BYVAL wRow AS SQLSETPOSIROW = 1) AS SQLRETURN
   DECLARE FUNCTION RefreshRecord (BYVAL wRow AS SQLSETPOSIROW = 1, BYVAL fLock AS SQLUSMALLINT = SQL_LOCK_NO_CHANGE) AS SQLRETURN
   DECLARE FUNCTION LockRecord (BYVAL wRow AS SQLSETPOSIROW = 1) AS SQLRETURN
   DECLARE FUNCTION UnlockRecord (BYVAL wRow AS SQLSETPOSIROW = 1) AS SQLRETURN
   DECLARE FUNCTION DeleteRecord (BYVAL wRow AS SQLSETPOSIROW = 1) AS SQLRETURN
   DECLARE FUNCTION ColAttribute (BYVAL ColumnNumber AS SQLUSMALLINT, BYVAL FieldIdentifier AS SQLUSMALLINT, _
                    BYVAL CharacterAttribute AS SQLPOINTER, BYVAL BufferLength AS SQLSMALLINT, BYVAL StringLength AS SQLSMALLINT PTR, _
                    BYVAL NumericAttribute AS ANY PTR) AS SQLRETURN
   DECLARE FUNCTION ColAutoUniqueValue (BYVAL ColNum AS SQLUSMALLINT) AS BOOLEAN
   DECLARE FUNCTION ColBaseColumnName (BYVAL ColNum AS SQLUSMALLINT) AS CWSTR
   DECLARE FUNCTION ColBaseTableName (BYVAL ColNum AS SQLUSMALLINT) AS CWSTR
   DECLARE FUNCTION ColCaseSensitive (BYVAL ColNum AS SQLUSMALLINT) AS BOOLEAN
   DECLARE FUNCTION ColCatalogName (BYVAL ColNum AS SQLUSMALLINT) AS CWSTR
   DECLARE FUNCTION ColConciseType (BYVAL ColNum AS SQLUSMALLINT) AS SQLINTEGER
   DECLARE FUNCTION ColCount () AS SQLINTEGER
   DECLARE FUNCTION ColDisplaySize (BYVAL ColNum AS SQLUSMALLINT) AS SQLINTEGER
   DECLARE FUNCTION ColFixedPrecScale (BYVAL ColNum AS SQLUSMALLINT) AS BOOLEAN
   DECLARE FUNCTION ColLabel (BYVAL ColNum AS SQLUSMALLINT) AS CWSTR
   DECLARE FUNCTION ColLength (BYVAL ColNum AS SQLUSMALLINT) AS SQLINTEGER
   DECLARE FUNCTION ColLiteralPrefix (BYVAL ColNum AS SQLUSMALLINT) AS CWSTR
   DECLARE FUNCTION ColLiteralSuffix (BYVAL ColNum AS SQLUSMALLINT) AS CWSTR
   DECLARE FUNCTION ColLocalTypeName (BYVAL ColNum AS SQLUSMALLINT) AS CWSTR
   DECLARE FUNCTION ColName (BYVAL ColNum AS SQLUSMALLINT) AS CWSTR
   DECLARE FUNCTION ColNullable (BYVAL ColNum AS SQLUSMALLINT) AS SQLINTEGER
   DECLARE FUNCTION ColNumPrecRadix (BYVAL ColNum AS SQLUSMALLINT) AS SQLINTEGER
   DECLARE FUNCTION ColOctetLength (BYVAL ColNum AS SQLUSMALLINT) AS SQLINTEGER
   DECLARE FUNCTION ColPrecision (BYVAL ColNum AS SQLUSMALLINT) AS SQLINTEGER
   DECLARE FUNCTION ColScale (BYVAL ColNum AS SQLUSMALLINT) AS SQLINTEGER
   DECLARE FUNCTION ColSchemaName (BYVAL ColNum AS SQLUSMALLINT) AS CWSTR
   DECLARE FUNCTION ColSearchable (BYVAL ColNum AS SQLUSMALLINT) AS SQLINTEGER
   DECLARE FUNCTION ColTableName (BYVAL ColNum AS SQLUSMALLINT) AS CWSTR
   DECLARE FUNCTION ColType (BYVAL ColNum AS SQLUSMALLINT) AS SQLINTEGER
   DECLARE FUNCTION ColTypeName (BYVAL ColNum AS SQLUSMALLINT) AS CWSTR
   DECLARE FUNCTION ColUnnamed (BYVAL ColNum AS SQLUSMALLINT) AS SQLINTEGER
   DECLARE FUNCTION ColUnsigned (BYVAL ColNum AS SQLUSMALLINT) AS BOOLEAN
   DECLARE FUNCTION ColUpdatable (BYVAL ColNum AS SQLUSMALLINT) AS BOOLEAN
   DECLARE FUNCTION GetStmtAttr (BYVAL Attribute AS SQLINTEGER, BYVAL ValuePtr AS SQLPOINTER, _
                    BYVAL BufferLength AS SQLINTEGER, BYVAL StringLength AS SQLINTEGER PTR) AS SQLRETURN
   DECLARE FUNCTION SetStmtAttr (BYVAL Attribute AS SQLINTEGER, BYVAL ValuePtr AS SQLPOINTER, BYVAL StringLength AS SQLINTEGER) AS SQLRETURN
   DECLARE FUNCTION GetCursorConcurrency () AS SQLUINTEGER
   DECLARE FUNCTION SetCursorConcurrency (BYVAL LockType AS SQLUINTEGER) AS SQLRETURN
   DECLARE FUNCTION SetOptimisticConcurrency () AS SQLRETURN
   DECLARE FUNCTION SetLockConcurrency () AS SQLRETURN
   DECLARE FUNCTION SetReadOnlyConcurrency () AS SQLRETURN
   DECLARE FUNCTION SetRowVerConcurrency () AS SQLRETURN
   DECLARE FUNCTION GetCursorType () AS SQLUINTEGER
   DECLARE FUNCTION SetCursorType (BYVAL CursorType AS DWORD) AS SQLRETURN
   DECLARE FUNCTION SetForwardOnlyCursor () AS SQLRETURN
   DECLARE FUNCTION SetKeysetDrivenCursor () AS SQLRETURN
   DECLARE FUNCTION SetStaticCursor () AS SQLRETURN
   DECLARE FUNCTION SetDynamicCursor () AS SQLRETURN
   DECLARE FUNCTION GetCursorKeysetSize () AS SQLUINTEGER
   DECLARE FUNCTION SetCursorKeysetSize (BYVAL KeysetSize AS SQLUINTEGER) AS SQLRETURN
   DECLARE FUNCTION GetCursorScrollability () AS SQLUINTEGER
   DECLARE FUNCTION SetCursorScrollability (BYVAL CursorScrollability AS SQLUINTEGER) AS SQLRETURN
   DECLARE FUNCTION GetCursorSensitivity () AS SQLUINTEGER
   DECLARE FUNCTION SetCursorSensitivity (BYVAL CursorSensitivity AS SQLUINTEGER) AS SQLRETURN
   DECLARE FUNCTION SetMultiuserKeysetCursor (BYVAL pwzCursorName AS WSTRING PTR = NULL) AS SQLRETURN
   DECLARE FUNCTION GetStmtAppParamDesc () AS SQLUINTEGER
   DECLARE FUNCTION SetStmtAppParamDesc (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
   DECLARE FUNCTION GetStmtAppRowDesc () AS SQLUINTEGER
   DECLARE FUNCTION SetStmtAppRowDesc (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
   DECLARE FUNCTION GetStmtAsyncEnable () AS SQLUINTEGER
   DECLARE FUNCTION SetStmtAsyncEnable (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
   DECLARE FUNCTION GetStmtFetchBookmarkPtr () AS SQLUINTEGER
   DECLARE FUNCTION SetStmtFetchBookmarkPtr (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
   DECLARE FUNCTION GetImpParamDesc () AS SQLUINTEGER
   DECLARE FUNCTION GetImpParamDescRec (BYVAL RecNumber AS SQLSMALLINT, BYVAL pwszName AS WSTRING PTR, _
                    BYVAL BufferLength AS SQLSMALLINT, BYVAL StringLength AS SQLSMALLINT PTR, BYVAL TypePtr AS SQLSMALLINT PTR, _
                    BYVAL SubTypePtr AS SQLSMALLINT PTR, BYVAL LengthPtr AS SQLLEN PTR, BYVAL PrecisionPtr AS SQLSMALLINT PTR, _
                    BYVAL ScalePtr AS SQLSMALLINT PTR, BYVAL NullablePtr AS SQLSMALLINT PTR) AS SQLRETURN
   DECLARE FUNCTION GetImpParamDescField (BYVAL RecNumber AS SQLSMALLINT, BYVAL FieldIdentifier AS SQLSMALLINT, _
                    BYVAL ValuePtr AS SQLPOINTER, BYVAL BufferLength AS SQLINTEGER, BYVAL StringLength AS SQLINTEGER PTR) AS SQLRETURN
   DECLARE FUNCTION GetImpParamDescFieldName (BYVAL RecNumber AS SQLSMALLINT) AS CWSTR
   DECLARE FUNCTION GetImpParamDescFieldType (BYVAL RecNumber AS SQLSMALLINT) AS SQLSMALLINT
   DECLARE FUNCTION GetImpParamDescFieldOctetLength (BYVAL RecNumber AS SQLSMALLINT) AS SQLSMALLINT
   DECLARE FUNCTION GetImpParamDescFieldPrecision (BYVAL RecNumber AS SQLSMALLINT) AS SQLSMALLINT
   DECLARE FUNCTION GetImpParamDescFieldScale (BYVAL RecNumber AS SQLSMALLINT) AS SQLSMALLINT
   DECLARE FUNCTION GetImpParamDescFieldNullable (BYVAL RecNumber AS SQLSMALLINT) AS SQLSMALLINT
   DECLARE FUNCTION GetImpRowDesc () AS SQLUINTEGER
   DECLARE FUNCTION GetImpRowDescRec (BYVAL RecNumber AS SQLSMALLINT, BYVAL pwszName AS WSTRING PTR, _
                    BYVAL BufferLength AS SQLSMALLINT, BYVAL StringLength AS SQLSMALLINT PTR, BYVAL TypePtr AS SQLSMALLINT PTR, _
                    BYVAL SubTypePtr AS SQLSMALLINT PTR, BYVAL LengthPtr AS SQLLEN PTR, BYVAL PrecisionPtr AS SQLSMALLINT PTR, _
                    BYVAL ScalePtr AS SQLSMALLINT PTR, BYVAL NullablePtr AS SQLSMALLINT PTR) AS SQLRETURN
   DECLARE FUNCTION GetImpRowDescField (BYVAL RecNumber AS SQLSMALLINT, BYVAL FieldIdentifier AS SQLSMALLINT, _
                    BYVAL ValuePtr AS SQLPOINTER, BYVAL BufferLength AS SQLINTEGER, BYVAL StringLength AS SQLINTEGER PTR) AS SQLRETURN
   DECLARE FUNCTION GetImpRowDescFieldName (BYVAL RecNumber AS SQLSMALLINT) AS CWSTR
   DECLARE FUNCTION GetImpRowDescFieldType (BYVAL RecNumber AS SQLSMALLINT) AS SQLSMALLINT
   DECLARE FUNCTION GetImpRowDescFieldOctetLength (BYVAL RecNumber AS SQLSMALLINT) AS SQLSMALLINT
   DECLARE FUNCTION GetImpRowDescFieldPrecision (BYVAL RecNumber AS SQLSMALLINT) AS SQLSMALLINT
   DECLARE FUNCTION GetImpRowDescFieldScale (BYVAL RecNumber AS SQLSMALLINT) AS SQLSMALLINT
   DECLARE FUNCTION GetImpRowDescFieldNullable (BYVAL RecNumber AS SQLSMALLINT) AS SQLSMALLINT
   DECLARE FUNCTION GetStmtMaxLength () AS SQLUINTEGER
   DECLARE FUNCTION SetStmtMaxLength (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
   DECLARE FUNCTION GetStmtMaxRows () AS SQLUINTEGER
   DECLARE FUNCTION SetStmtMaxRows (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
   DECLARE FUNCTION GetStmtNoScan () AS SQLUINTEGER
   DECLARE FUNCTION SetStmtNoScan (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
   DECLARE FUNCTION GetStmtParamBindOffsetPtr () AS SQLUINTEGER
   DECLARE FUNCTION SetStmtParamBindOffsetPtr (BYVAL dwAttr AS DWORD) AS SQLRETURN
   DECLARE FUNCTION GetStmtParamBindType () AS SQLUINTEGER
   DECLARE FUNCTION SetStmtParamBindType (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
   DECLARE FUNCTION GetStmtParamOperationPtr () AS SQLUINTEGER
   DECLARE FUNCTION SetStmtParamOperationPtr (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
   DECLARE FUNCTION GetStmtParamsProcessedPtr () AS SQLUINTEGER
   DECLARE FUNCTION SetStmtParamsProcessedPtr (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
   DECLARE FUNCTION GetStmtParamsetSize () AS SQLUINTEGER
   DECLARE FUNCTION SetStmtParamsetSize (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
   DECLARE FUNCTION GetStmtQueryTimeout () AS SQLUINTEGER
   DECLARE FUNCTION SetStmtQueryTimeout (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
   DECLARE FUNCTION GetStmtRetrieveData () AS SQLUINTEGER
   DECLARE FUNCTION SetStmtRetrieveData (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
   DECLARE FUNCTION GetStmtParamStatusPtr () AS SQLUINTEGER
   DECLARE FUNCTION SetStmtParamStatusPtr (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
   DECLARE FUNCTION GetStmtRowArraySize () AS SQLUINTEGER
   DECLARE FUNCTION SetStmtRowArraySize (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
   DECLARE FUNCTION GetStmtRowBindOffsetPtr () AS SQLUINTEGER
   DECLARE FUNCTION SetStmtRowBindOffsetPtr (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
   DECLARE FUNCTION GetStmtRowBindType () AS SQLUINTEGER
   DECLARE FUNCTION SetStmtRowBindType (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
   DECLARE FUNCTION GetStmtRowNumber () AS SQLUINTEGER
   DECLARE FUNCTION GetStmtRowOperationPtr () AS SQLUINTEGER
   DECLARE FUNCTION SettmtRowOperationPtr (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
   DECLARE FUNCTION GetStmtRowStatusPtr () AS SQLUINTEGER
   DECLARE FUNCTION SetStmtRowStatusPtr (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
   DECLARE FUNCTION GetStmtRowsFetchedPtr () AS SQLUINTEGER
   DECLARE FUNCTION SetStmtRowsFetchedPtr (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
   DECLARE FUNCTION GetStmtSimulateCursor () AS SQLUINTEGER
   DECLARE FUNCTION SetStmtSimulateCursor (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
   DECLARE FUNCTION GetStmtUseBookmarks () AS SQLUINTEGER
   DECLARE FUNCTION SetStmtUseBookmarks (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
   DECLARE FUNCTION GetStatistics (BYREF wszCatalogName AS WSTRING, BYVAL CatalogNameLength AS SQLSMALLINT, _
                    BYREF wszSchemaName AS WSTRING, BYVAL SchemaNameLength AS SQLSMALLINT, BYREF wszTableName AS WSTRING, _
                    BYVAL TableNameLength AS SQLSMALLINT, BYVAL fUnique AS SQLUSMALLINT, BYVAL fCardinality AS SQLUSMALLINT) AS SQLRETURN
   DECLARE FUNCTION GetColumns (BYREF wszCatalogName AS WSTRING, BYVAL CatalogNameLength AS SQLSMALLINT, _
                    BYREF wszSchemaName AS WSTRING, BYVAL SchemaNameLength AS SQLSMALLINT, BYREF wszTableName AS WSTRING, _
                    BYVAL TableNameLength AS SQLSMALLINT, BYREF wszColumnName AS WSTRING, BYVAL ColumnNameLength AS SQLSMALLINT) AS SQLRETURN
   DECLARE FUNCTION GetTypeInfo (BYVAL DataType AS SQLSMALLINT) AS SQLRETURN
   DECLARE FUNCTION GetTables (BYREF wszCatalogName AS WSTRING, BYVAL CatalogNameLength AS SQLSMALLINT, _
                    BYREF wszSchemaName AS WSTRING, BYVAL SchemaNameLength AS SQLSMALLINT, BYREF wszTableName AS WSTRING, _
                    BYVAL TableNameLength AS SQLSMALLINT, BYREF wszTableType AS WSTRING, BYVAL TableTypeLength AS SQLSMALLINT) AS SQLRETURN
   DECLARE FUNCTION GetTablePrivileges (BYREF wszCatalogName AS WSTRING, BYVAL CatalogNameLength AS SQLSMALLINT, _
                    BYREF wszSchemaName AS WSTRING, BYVAL SchemaNameLength AS SQLSMALLINT, BYREF wszTableName AS WSTRING, _
                    BYVAL TableNameLength AS SQLSMALLINT) AS SQLRETURN
   DECLARE FUNCTION GetPrimaryKeys (BYREF wszCatalogName AS WSTRING, BYVAL CatalogNameLength AS SQLSMALLINT, _
                    BYREF wszSchemaName AS WSTRING, BYVAL SchemanameLength AS SQLSMALLINT, BYREF wszTableName AS WSTRING, _
                    BYVAL TableNameLength AS SQLSMALLINT) AS SQLRETURN
   DECLARE FUNCTION GetProcedures (BYREF wszCatalogName AS WSTRING, BYVAL CatalogNameLength AS SQLSMALLINT, _
                    BYREF wszSchemaName AS WSTRING, BYVAL SchemaNameLength AS SQLSMALLINT, BYREF wszProcName AS WSTRING, _
                    BYVAL ProcNameLength AS SQLSMALLINT) AS SQLRETURN
   DECLARE FUNCTION GetSpecialColumns (BYVAL identifierType AS SQLUSMALLINT, BYREF wszCatalogName AS WSTRING, _
                    BYVAL CatalogNameLength AS SQLUSMALLINT, BYREF wszSchemaName AS WSTRING, BYVAL SchemaNameLength AS SQLUSMALLINT, _
                    BYREF wszTableName AS WSTRING, BYVAL TableNameLength AS SQLUSMALLINT, BYVAL fScope AS SQLUSMALLINT,_
                    BYVAL fNullable AS SQLUSMALLINT) AS SQLRETURN
   DECLARE FUNCTION GetColumnPrivileges (BYREF wszCatalogName AS WSTRING, BYVAL CatalogNameLength AS SQLSMALLINT, _
                    BYREF wszSchemaName AS WSTRING, BYVAL SchemaNameLength AS SQLSMALLINT, BYREF wszTableName AS WSTRING, _
                    BYVAL TableNameLength AS SQLSMALLINT, BYREF wszColumnName AS WSTRING, BYVAL ColumnNameLength AS SQLSMALLINT) AS SQLRETURN
   DECLARE FUNCTION GetForeignKeys (BYREF wszPkCatalogName AS WSTRING, BYVAL pkCatalogNameLength AS SQLSMALLINT, _
                    BYREF wszPkSchemaName AS WSTRING, BYVAL pkSchemaNameLength AS SQLSMALLINT, BYREF wszPkTableName AS WSTRING, _
                    BYVAL pkTableNameLength AS SQLSMALLINT, BYREF wszFkCatalogName AS WSTRING, BYVAL FkCatalogNameLength AS SQLSMALLINT, _
                    BYREF wszFkSchemaName AS WSTRING, BYVAL FkSchemaNameLength AS SQLSMALLINT, BYREF wszFkTableName AS WSTRING, BYVAL fkTableNameLength AS SQLSMALLINT) AS SQLRETURN
   DECLARE FUNCTION GetProcedureColumns (BYREF wszCatalogName AS WSTRING, BYVAL CatalogNameLength AS SQLSMALLINT, _
                    BYREF wszSchemaName AS WSTRING, BYVAL SchemaNameLength AS SQLSMALLINT, BYREF wszProcName AS WSTRING, _
                    BYVAL ProcNameLength AS SQLSMALLINT, BYREF wszColumnName AS WSTRING, BYVAL ColumnNameLength AS SQLSMALLINT) AS SQLRETURN


END TYPE
' ########################################################################################

#include once "Afx/COdbc/COdbcDbc.inc"
#include once "Afx/COdbc/COdbcStmt.inc"

END NAMESPACE
