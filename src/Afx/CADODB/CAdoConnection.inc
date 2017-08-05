' ########################################################################################
' Microsoft Windows
' Contents: ADO Connection class
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once

NAMESPACE Afx

' ########################################################################################
' CAdoConnection class
' ########################################################################################
TYPE CAdoConnection EXTENDS CAdoBase

   m_Result AS HRESULT
   m_pConnection AS Afx_ADOConnection PTR

   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYVAL pConnection AS Afx_ADOConnection PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   DECLARE DESTRUCTOR
'   DECLARE OPERATOR @ () AS Afx_ADOConnection PTR PTR
   DECLARE FUNCTION vptr () AS Afx_ADOConnection PTR PTR

   ' // Attaches a reference to an Afx_ADOCommand object to the class
   DECLARE SUB Attach (BYVAL pConnection AS Afx_ADOConnection PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   ' // * Sets or returns the connection string
   DECLARE PROPERTY ConnectionString () AS CBSTR
   DECLARE PROPERTY ConnectionString (BYREF cbsConstr AS CBSTR)
   ' // * Opens a connection to a data source
   DECLARE FUNCTION Open (BYREF cbsConStr AS CBSTR = "", BYREF cbsUserID AS CBSTR = "", BYREF cbsPassword AS CBSTR = "", BYVAL Options AS LONG = adOptionUnspecified) AS HRESULT
   ' // * Indicates if a connection is open or closed
   DECLARE PROPERTY State () AS LONG
   ' // * Indicates the ADO version number
   DECLARE PROPERTY Version () AS CBSTR
   ' // * Closes a Connection object and any dependent objects
   DECLARE FUNCTION Close () AS HRESULT
   ' // * Sets or returns a Long value that indicates, in seconds, how long to wait for a command to execute
   DECLARE PROPERTY CommandTimeout () AS LONG
   DECLARE PROPERTY CommandTimeout (BYVAL lTimeout AS LONG)
   ' // * Indicates how long to wait while establishing a connection before terminating the attempt and generating an error
   DECLARE PROPERTY ConnectionTimeout () AS LONG
   DECLARE PROPERTY ConnectionTimeout (BYVAL lTimeout AS LONG)
   ' // Indicates the default database for a Connection object
   DECLARE PROPERTY DefaultDatabase () AS CBSTR
   DECLARE PROPERTY DefaultDatabase (BYREF cbsDatabase AS CBSTR)
   ' // * Indicates the location of the cursor service
   DECLARE PROPERTY CursorLocation () AS CursorLocationEnum
   DECLARE PROPERTY CursorLocation (BYVAL lCursorLoc AS CursorLocationEnum)
   ' // * Indicates the level of isolation for a Connection object
   DECLARE PROPERTY IsolationLevel () AS IsolationLevelEnum
   DECLARE PROPERTY IsolationLevel (BYVAL Level AS IsolationLevelEnum)
   ' // Indicates one or more characteristics of an object
   DECLARE PROPERTY Attributes () AS LONG
   DECLARE PROPERTY Attributes (BYVAL lAttr AS LONG)
   ' // Indicates the available permissions for modifying data in a Connection object
   DECLARE PROPERTY Mode () AS ConnectModeEnum
   DECLARE PROPERTY Mode (BYVAL lMode AS ConnectModeEnum)
   ' // * Indicates the name of the provider for a Connection object.
   DECLARE PROPERTY Provider () AS CBSTR
   DECLARE PROPERTY Provider (BYREF cbsProvider AS CBSTR)
   ' // * Returns a reference to the Errors collection
   DECLARE PROPERTY Errors () AS Afx_ADOErrors PTR
   ' // * Returns a reference to the Properties collection.
   DECLARE PROPERTY Properties () AS Afx_ADOProperties PTR
   ' // Executes the specified query, SQL statement, stored procedure, or provider-specific text
   DECLARE FUNCTION Execute (BYREF cbsCommandText AS CBSTR, BYVAL RecordsAffected AS LONG PTR = NULL, BYVAL Options AS LONG = -1) AS Afx_ADORecordset PTR
   ' // Begins a new transaction
   DECLARE FUNCTION BeginTrans () AS LONG
   ' // Saves any changes and ends the current transaction. It may also start a new transaction.
   DECLARE FUNCTION CommitTrans () AS HRESULT
   ' // Saves any changes and ends the current transaction. It may also start a new transaction.
   DECLARE FUNCTION RollbackTrans () AS HRESULT
   ' // Cancels execution of a pending, asynchronous method call
   DECLARE FUNCTION Cancel () AS HRESULT
   ' // Obtains database schema information from the provider
   DECLARE FUNCTION OpenSchema (BYVAL Schema AS SchemaEnum, BYREF Restrictions AS CVAR = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND), _
           BYREF SchemaID AS CVAR = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND)) AS Afx_ADORecordset PTR
   ' // Retrieves information from the ADO Errors collection
   DECLARE FUNCTION GetErrorInfo (BYVAL dwError AS HRESULT = 0) AS CBSTR

END TYPE
' ########################################################################################

' ========================================================================================
' Default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CAdoConnection
   CADODB_DP("-BEGIN CAdoConnection CONSTRUCTOR")
   ' // Create an instance of the Connection object
'   DIM CLSID_Connection_ AS GUID = (&h00000514, &h0000, &h0010, {&h80, &h00, &h00, &hAA, &h00, &h6D, &h2E, &hA4})
'   DIM IID_Connection_ AS GUID = (&h00001550, &h0000, &h0010, {&h80, &h00, &h00, &hAA, &h00, &h6D, &h2E, &hA4})
'   SetResult(CoCreateInstance(@CLSID_Connection_, NULL, CLSCTX_INPROC_SERVER, @IID_Connection_, @m_pConnection))
   m_pConnection = AfxNewCom("ADODB.Connection")
   CADODB_DP("-END CAdoConnection CONSTRUCTOR " & WSTR(m_pConnection))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Attaches a reference to an Afx_ADOCommand object to the class
' ========================================================================================
PRIVATE CONSTRUCTOR CAdoConnection (BYVAL pConnection AS Afx_ADOConnection PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   CADODB_DP("CAdoConnection CONSTRUCTOR - Afx_ADOConnection")
   m_pConnection = pConnection
   IF m_pConnection THEN
      IF fAddRef THEN m_pConnection->AddRef
   END IF
END CONSTRUCTOR
' ========================================================================================

' ===========================================================================================
' Cleanup
' ===========================================================================================
PRIVATE DESTRUCTOR CAdoConnection
   CADODB_DP("CAdoConnection DESTRUCTOR " & WSTR(m_pConnection))
   ' // Close and release the connection
   IF m_pConnection THEN
      IF this.State = adStateOpen THEN this.Close
      m_pConnection->Release
   END IF
END DESTRUCTOR
' ===========================================================================================

' ========================================================================================
' Returns the address of the connection pointer
' ========================================================================================
'PRIVATE OPERATOR CAdoConnection.@ () AS Afx_ADOConnection PTR PTR
'   OPERATOR = @m_pConnection
'END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CAdoConnection.vptr () AS Afx_ADOConnection PTR PTR
   CVAR_DP("CAdoConnection vptr")
   IF m_pConnection THEN m_pConnection->Release
   RETURN @m_pConnection
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the encapsulated IDispatch pointer
' ========================================================================================
PRIVATE OPERATOR * (BYREF cConnection AS CAdoConnection) AS Afx_ADOConnection PTR
   OPERATOR = cConnection.m_pConnection
END OPERATOR
' ========================================================================================

' ===========================================================================================
' Attaches a reference to an Afx_ADOConnection object to the class
' ===========================================================================================
PRIVATE SUB CAdoConnection.Attach (BYVAL pCommand AS Afx_ADOConnection PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   CADODB_DP("CAdoConnection.Attach")
   IF m_pConnection THEN m_pConnection->Release
   m_pConnection = m_pConnection
   ' // Increase the reference count
   IF fAddRef THEN
      IF m_pConnection THEN m_pConnection->AddRef
   END IF
END SUB
' ===========================================================================================

' ===========================================================================================
' * Sets or returns the connection string
' ===========================================================================================
PRIVATE PROPERTY CAdoConnection.ConnectionString () AS CBSTR
   DIM bsConnectionString AS AFX_BSTR
   IF m_pConnection THEN SetResult(m_pConnection->get_ConnectionString(@bsConnectionString))
   PROPERTY = bsConnectionString
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoConnection.ConnectionString (BYREF cbsConStr AS CBSTR)
   IF m_pConnection THEN SetResult(m_pConnection->put_ConnectionString(cbsConStr.m_bstr))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' * Opens a connection to a data source
' ===========================================================================================
PRIVATE FUNCTION CAdoConnection.Open (BYREF cbsConStr AS CBSTR = "", BYREF cbsUserID AS CBSTR = "", _
BYREF cbsPassword AS CBSTR = "", BYVAL Options AS LONG = adOptionUnspecified) AS HRESULT
   CADODB_DP("CAdoConnection Open - CBSTR " & cbsConStr)
   IF m_pConnection THEN RETURN SetResult(m_pConnection->Open(cbsConStr.m_bstr, cbsUserID.m_bstr, cbsPassword.m_bstr, Options))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' * Indicates if a connection is open or closed
' ===========================================================================================
PRIVATE PROPERTY CAdoConnection.State () AS LONG
   DIM plObjState AS LONG
   IF m_pConnection THEN SetResult(m_pConnection->get_State(@plObjState))
   PROPERTY = plObjState
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' * Indicates the ADO version number
' ===========================================================================================
PRIVATE PROPERTY CAdoConnection.Version () AS CBSTR
   DIM bstrVersion AS AFX_BSTR
   IF m_pConnection THEN SetResult(m_pConnection->get_Version(@bstrVersion))
   PROPERTY = bstrVersion
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' * Closes a Connection object and any dependent objects
' ===========================================================================================
PRIVATE FUNCTION CAdoConnection.Close () AS HRESULT
   IF m_pConnection THEN RETURN(SetResult(m_pConnection->Close))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' * Sets or returns a Long value that indicates, in seconds, how long to wait for a command to
' execute. Default is 30.
' ===========================================================================================
PRIVATE PROPERTY CAdoConnection.CommandTimeout () AS LONG
   DIM lTimeout AS LONG
   IF m_pConnection THEN SetResult(m_pConnection->get_CommandTimeout(@lTimeout))
   PROPERTY = lTimeout
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoConnection.CommandTimeout (BYVAL lTimeout AS LONG)
   IF m_pConnection THEN SetResult(m_pConnection->put_CommandTimeout(lTimeout))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' * Indicates how long to wait while establishing a connection before terminating the attempt
' and generating an error.
' ===========================================================================================
PRIVATE PROPERTY CAdoConnection.ConnectionTimeout () AS LONG
   DIM lTimeout AS LONG
   IF m_pConnection THEN SetResult(m_pConnection->get_ConnectionTimeout(@lTimeout))
   PROPERTY = lTimeout
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoConnection.ConnectionTimeout (BYVAL lTimeout AS LONG)
   IF m_pConnection THEN SetResult(m_pConnection->put_ConnectionTimeout(lTimeout))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the default database for a Connection object
' ===========================================================================================
PRIVATE PROPERTY CAdoConnection.DefaultDatabase () AS CBSTR
   DIM bstrDatabase AS AFX_BSTR
   IF m_pConnection THEN SetResult(m_pConnection->get_DefaultDatabase(@bstrDatabase))
   PROPERTY = bstrDatabase
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoConnection.DefaultDatabase (BYREF cbsDatabase AS CBSTR)
   IF m_pConnection THEN SetResult(m_pConnection->put_DefaultDatabase(cbsDatabase.m_bstr))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' * Indicates the location of the cursor service
' ===========================================================================================
PRIVATE PROPERTY CAdoConnection.CursorLocation () AS CursorLocationEnum
   DIM lCursorLoc AS CursorLocationEnum
   IF m_pConnection THEN SetResult(m_pConnection->get_CursorLocation(@lCursorLoc))
   PROPERTY = lCursorLoc
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoConnection.CursorLocation (BYVAL lCursorLoc AS CursorLocationEnum)
   IF m_pConnection THEN SetResult(m_pConnection->put_CursorLocation(lCursorLoc))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the level of isolation for a Connection object.
' ===========================================================================================
PRIVATE PROPERTY CAdoConnection.IsolationLevel () AS IsolationLevelEnum
   DIM Level AS IsolationLevelEnum
   IF m_pConnection THEN SetResult(m_pConnection->get_IsolationLevel(@Level))
   PROPERTY = Level
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoConnection.IsolationLevel (BYVAL Level AS IsolationLevelEnum)
   IF m_pConnection THEN SetResult(m_pConnection->put_IsolationLevel(Level))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates one or more characteristics of an object
' ===========================================================================================
PRIVATE PROPERTY CAdoConnection.Attributes () AS LONG
   DIM lAttr AS LONG
   IF m_pConnection THEN SetResult(m_pConnection->get_Attributes(@lAttr))
   PROPERTY = lAttr
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoConnection.Attributes (BYVAL lAttr AS LONG)
   IF m_pConnection THEN SetResult(m_pConnection->put_Attributes(lAttr))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the available permissions for modifying data in a Connection object
' ===========================================================================================
PRIVATE PROPERTY CAdoConnection.Mode () AS ConnectModeEnum
   DIM lMode AS ConnectModeEnum
   IF m_pConnection THEN SetResult(m_pConnection->get_Mode(@lMode))
   PROPERTY = lMode
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoConnection.Mode (BYVAL lMode AS ConnectModeEnum)
   IF m_pConnection THEN SetResult(m_pConnection->put_Mode(lMode))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the name of the provider for a Connection object
' ===========================================================================================
PRIVATE PROPERTY CAdoConnection.Provider () AS CBSTR
   DIM bstrProvider AS AFX_BSTR
   IF m_pConnection THEN SetResult(m_pConnection->get_Provider(@bstrProvider))
   PROPERTY = bstrProvider
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoConnection.Provider (BYREF cbsProvider AS CBSTR)
   IF m_pConnection THEN SetResult(m_pConnection->put_Provider(cbsProvider.m_bstr))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Returns a reference to the Errors collection
' ===========================================================================================
PRIVATE PROPERTY CAdoConnection.Errors () AS Afx_ADOErrors PTR
   DIM pErrors AS Afx_ADOErrors PTR
   IF m_pConnection THEN SetResult(m_pConnection->get_Errors(@pErrors))
   PROPERTY = pErrors
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Returns a reference to the Properties collection
' ===========================================================================================
PRIVATE PROPERTY CAdoConnection.Properties () AS Afx_ADOProperties PTR
   DIM pProperties AS Afx_ADOProperties PTR
   IF m_pConnection THEN SetResult(m_pConnection->get_Properties(@pProperties))
   PROPERTY = pProperties
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Executes the specified query, SQL statement, stored procedure, or provider-specific text
' ===========================================================================================
PRIVATE FUNCTION CAdoConnection.Execute (BYREF cbsCommandText AS CBSTR, BYVAL RecordsAffected AS LONG PTR = NULL, BYVAL Options AS LONG = -1) AS Afx_ADORecordset PTR
   DIM vRecordsAffected AS VARIANT
   V_VT(@vRecordsAffected) = VT_I4
   DIM piRset AS Afx_ADORecordset PTR
   IF m_pConnection THEN SetResult(m_pConnection->Execute(cbsCommandText.m_bstr, @vRecordsAffected, Options, @piRset))
   IF RecordsAffected THEN *RecordsAffected = vRecordsAffected.lVal
   RETURN piRset
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Begins a new transaction
' ===========================================================================================
PRIVATE FUNCTION CAdoConnection.BeginTrans () AS LONG
   DIM TransactionLevel AS LONG
   IF m_pConnection THEN SetResult(m_pConnection->BeginTrans(@TransactionLevel))
   FUNCTION = TransactionLevel
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Saves any changes and ends the current transaction. It may also start a new transaction.
' ===========================================================================================
PRIVATE FUNCTION CAdoConnection.CommitTrans () AS HRESULT
   IF m_pConnection THEN RETURN(SetResult(m_pConnection->CommitTrans))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Saves any changes and ends the current transaction. It may also start a new transaction.
' ===========================================================================================
PRIVATE FUNCTION CAdoConnection.RollbackTrans () AS HRESULT
   IF m_pConnection THEN RETURN(SetResult(m_pConnection->RollbackTrans))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Cancels execution of a pending, asynchronous method call
' ===========================================================================================
PRIVATE FUNCTION CAdoConnection.Cancel () AS HRESULT
   IF m_pConnection THEN RETURN(SetResult(m_pConnection->Cancel))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Obtains database schema information from the provider
' ===========================================================================================
PRIVATE FUNCTION CAdoConnection.OpenSchema (BYVAL Schema AS SchemaEnum, BYREF Restrictions AS CVAR = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND), _
BYREF SchemaID AS CVAR = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND)) AS Afx_ADORecordset PTR
   DIM pprset AS Afx_ADORecordset PTR
   IF m_pConnection THEN SetResult(m_pConnection->OpenSchema(Schema, Restrictions, SchemaID, @pprset))
   RETURN pprset
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Retrieves information from the ADO Errors collection
' ===========================================================================================
PRIVATE FUNCTION CAdoConnection.GetErrorInfo (BYVAL nError AS HRESULT = 0) AS CBSTR
   RETURN AfxAdoGetErrorInfo(m_pConnection, nError)
END FUNCTION
' ===========================================================================================

END NAMESPACE
