' ########################################################################################
' Microsoft Windows
' Contents: ADO Command class
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once

NAMESPACE Afx

' ########################################################################################
' CAdoCommand class
' ########################################################################################
TYPE CAdoCommand EXTENDS CAdoBase

   m_Result AS HRESULT
   m_pCommand AS Afx_ADOCommand PTR

   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYVAL pCommand AS Afx_ADOCommand PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   DECLARE DESTRUCTOR
'   DECLARE OPERATOR @ () AS Afx_ADOCommand PTR PTR
   DECLARE FUNCTION vptr () AS Afx_ADOCommand PTR PTR

   ' // Attaches a reference to an Afx_ADOCommand object to the class
   DECLARE SUB Attach (BYVAL pCommand AS Afx_ADOCommand PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   ' // Returns a reference to the Properties collection.
   DECLARE PROPERTY Properties () AS Afx_ADOProperties PTR
   ' // Returns a reference to the Parameters collection.
   DECLARE PROPERTY Parameters () AS Afx_ADOParameters PTR
   ' // * Determines the Connection object over which the specified Command object will execute
   DECLARE PROPERTY ActiveConnection (BYREF vConn AS CVAR)
   DECLARE PROPERTY ActiveConnection (BYVAL pconn AS Afx_ADOConnection PTR)
   DECLARE PROPERTY ActiveConnection (BYREF pconn AS CAdoConnection)
   DECLARE PROPERTY ActiveConnection () AS Afx_ADOConnection PTR
   ' // * Sets or returns a BSTR value that contains a provider command
   DECLARE PROPERTY CommandText (BYREF cbsText AS CBSTR)
   DECLARE PROPERTY CommandText () AS CBSTR
   ' // * Executes the query, SQL statement, or stored procedure specified in the CommandText or CommandStream property
   DECLARE FUNCTION Execute (BYVAL RecordsAffected AS LONG PTR = NULL, BYREF cvParameters AS CVAR = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND), BYVAL Options AS LONG = adCmdUnspecified) AS Afx_ADORecordset PTR
   ' // * Sets or returns a Long value that indicates, in seconds, how long to wait for a command to execute
   DECLARE PROPERTY CommandTimeout () AS LONG
   DECLARE PROPERTY CommandTimeout (BYVAL lTimeout AS LONG)
   ' // Sets or returns one or more CommandTypeEnum values.
   DECLARE PROPERTY CommandType (BYVAL lCmdType AS CommandTypeEnum)
   DECLARE PROPERTY CommandType () AS CommandTypeEnum
   ' // Sets or returns a Boolean value that, if set to True, indicates that the command should be prepared.
   DECLARE PROPERTY Prepared (BYVAL fPrepared AS BOOLEAN)
   DECLARE PROPERTY Prepared () AS BOOLEAN
   ' // Sets or returns a BSTR value that indicates the name of a Command object.
   DECLARE PROPERTY Name () AS CBSTR
   DECLARE PROPERTY Name (BYREF cbsName AS CBSTR)
   ' // Indicates the dialect of the CommandText or CommandStream properties.
   DECLARE PROPERTY Dialect (BYREF cbsDialect AS CBSTR)
   DECLARE PROPERTY Dialect () AS CBSTR
   ' // Indicates whether parameter names should be passed to the provider.
   DECLARE PROPERTY NamedParameters (BYVAL fNamedParameters AS BOOLEAN)
   DECLARE PROPERTY NamedParameters () AS BOOLEAN
   ' // Sets or returns the stream used as the input for a Command object.
   DECLARE PROPERTY CommandStream (BYVAL pStream AS Afx_ADOStream PTR)
   DECLARE PROPERTY CommandStream () AS Afx_ADOStream PTR
   ' // Indicates if the Command object is open or closed.
   DECLARE PROPERTY State () AS LONG
   ' // Cancels execution of a pending, asynchronous method call.
   DECLARE FUNCTION Cancel () AS HRESULT
   ' // Creates a new Parameter object with the specified properties.
   DECLARE FUNCTION CreateParameter (BYREF cbsName AS CBSTR = "", BYVAL nType AS DataTypeEnum = 0, _
           BYVAL Direction AS ParameterDirectionEnum = adParamInput, BYVAL Size AS LONG = 0, _
           BYREF cvValue AS CVAR = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND)) AS Afx_ADOParameter PTR
   ' // Retrieves information from the ADO Errors collection
   DECLARE FUNCTION GetErrorInfo (BYVAL nError AS HRESULT = 0) AS CBSTR

END TYPE

' ========================================================================================
' Default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CAdoCommand
   CADODB_DP("-BEGIN CAdoCommand CONSTRUCTOR")
   ' // Create an instance of the Connection object
'   DIM CLSID_Command_ AS GUID = (&h00000507, &h0000, &h0010, {&h80, &h00, &h00, &hAA, &h00, &h6D, &h2E, &hA4})
'   DIM IID_Command_ AS GUID = (&h986761E8, &h7269, &h4890, {&hAA, &h65, &hAD, &h7C, &h03, &h69, &h7A, &h6D})
'   SetResult(CoCreateInstance(@CLSID_Command_, NULL, CLSCTX_INPROC_SERVER, @IID_Command_, @m_pCommand))
   m_pCommand = AfxNewCom("ADODB.Command")
   CADODB_DP("-END CAdoCommand CONSTRUCTOR " & WSTR(m_pCommand))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Attaches a reference to an Afx_ADOCommand object to the class
' ========================================================================================
PRIVATE CONSTRUCTOR CAdoCommand (BYVAL pCommand AS Afx_ADOCommand PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   CADODB_DP("-BEGIN CAdoCommand CONSTRUCTOR - Afx_ADOCommand")
   m_pCommand = pCommand
   IF m_pCommand THEN
      IF fAddRef THEN m_pCommand->AddRef
   END IF
END CONSTRUCTOR
' ========================================================================================

' ===========================================================================================
' Cleanup
' ===========================================================================================
PRIVATE DESTRUCTOR CAdoCommand
   CADODB_DP("CAdoCommand DESTRUCTOR " & WSTR(m_pCommand))
   IF m_pCommand THEN m_pCommand->Release
END DESTRUCTOR
' ===========================================================================================

' ========================================================================================
' Returns the address of the connection pointer
' ========================================================================================
'PRIVATE OPERATOR CAdoCommand.@ () AS Afx_ADOCommand PTR PTR
'   OPERATOR = @m_pCommand
'END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CAdoCommand.vptr () AS Afx_ADOCommand PTR PTR
   CVAR_DP("CAdoCommand vptr")
   IF m_pCommand THEN m_pCommand->Release
   RETURN @m_pCommand
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the encapsulated IDispatch pointer
' ========================================================================================
PRIVATE OPERATOR * (BYREF cCommand AS CAdoCommand) AS Afx_ADOCommand PTR
   OPERATOR = cCommand.m_pCommand
END OPERATOR
' ========================================================================================

' ===========================================================================================
' Attaches a reference to an Afx_ADOCommand object to the class
' ===========================================================================================
PRIVATE SUB CAdoCommand.Attach (BYVAL pCommand AS Afx_ADOCommand PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   CADODB_DP("CAdoCommand.Attach")
   IF m_pCommand THEN m_pCommand->Release
   m_pCommand = pCommand
   ' // Increase the reference count
   IF fAddRef THEN
      IF m_pCommand THEN m_pCommand->AddRef
   END IF
END SUB
' ===========================================================================================

' ===========================================================================================
' * Returns a reference to the Properties collection.
' ===========================================================================================
PRIVATE PROPERTY CAdoCommand.Properties () AS Afx_ADOProperties PTR
   CADODB_DP("CAdoCommand.Properties")
   DIM pProperties AS Afx_ADOProperties PTR
   IF m_pCommand THEN SetResult(m_pCommand->get_Properties(@pProperties))
   PROPERTY = pProperties
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' * Returns a reference to the Parameters collection.
' ===========================================================================================
PRIVATE PROPERTY CAdoCommand.Parameters () AS Afx_ADOParameters PTR
   CADODB_DP("CAdoCommand.Parameters")
   DIM pParameters AS Afx_ADOParameters PTR
   IF m_pCommand THEN SetResult(m_pCommand->get_Parameters(@pParameters))
   PROPERTY = pParameters
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' * Sets a reference to the active connection
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoCommand.ActiveConnection (BYREF vConn AS CVAR)
   CADODB_DP("CAdoCommand.ActiveConnection - CVAR")
   IF m_pCommand THEN SetResult(m_pCommand->put_ActiveConnection(vConn))
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoCommand.ActiveConnection (BYVAL pconn AS Afx_ADOConnection PTR)
   CADODB_DP("CAdoCommandCAdoCommand.ActiveConnection")
   IF m_pCommand THEN SetResult(m_pCommand->putref_ActiveConnection(pconn))
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoCommand.ActiveConnection (BYREF pconn AS CAdoConnection)
   CADODB_DP("CAdoCommandCAdoCommand.ActiveConnection")
   IF m_pCommand THEN SetResult(m_pCommand->putref_ActiveConnection(pconn.m_pConnection))
END PROPERTY
' ===========================================================================================
' ===========================================================================================
' * Gets a reference to the active connection
' ===========================================================================================
PRIVATE PROPERTY CAdoCommand.ActiveConnection () AS Afx_ADOConnection PTR
   CADODB_DP("CAdoCommandCAdoCommand.ActiveConnection")
   DIM pConnection AS Afx_ADOCOnnection PTR
   IF m_pCommand THEN SetResult(m_pCommand->get_ActiveConnection(@pConnection))
   PROPERTY = pConnection
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' * Sets or returns a BSTR value that contains a provider command
' ===========================================================================================
PRIVATE PROPERTY CAdoCommand.CommandText () AS CBSTR
   CADODB_DP("CAdoCommand.CommandText")
   DIM bstrText AS AFX_BSTR
   IF m_pCommand THEN SetResult(m_pCommand->get_CommandText(@bstrText))
   RETURN bstrText
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoCommand.CommandText (BYREF cbsText AS CBSTR)
   CADODB_DP("CAdoCommand.CommandText")
   IF m_pCommand THEN SetResult(m_pCommand->put_CommandText(cbsText.m_bstr))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' * Executes the specified query, SQL statement, stored procedure, or provider-specific text
' ===========================================================================================
PRIVATE FUNCTION CAdoCommand.Execute (BYVAL RecordsAffected AS LONG PTR = NULL, BYREF cvParameters AS CVAR = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND), BYVAL Options AS LONG = adCmdUnspecified) AS Afx_ADORecordset PTR
   CADODB_DP("CAdoCommand.Execute")
   DIM vRecordsAffected AS VARIANT
   V_VT(@vRecordsAffected) = VT_I4
   DIM ppirs AS Afx_ADORecordset PTR
   IF m_pCommand THEN SetResult(m_pCommand->Execute(@vRecordsAffected, cvParameters, Options, @ppirs))
   IF RecordsAffected THEN *RecordsAffected = vRecordsAffected.lVal
   RETURN ppirs
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' * Sets or returns a Long value that indicates, in seconds, how long to wait for a command to
' execute. Default is 30.
' ===========================================================================================
PRIVATE PROPERTY CAdoCommand.CommandTimeout () AS LONG
   CADODB_DP("CAdoCommand.CommandTimeout")
   DIM lTimeout AS LONG
   IF m_pCommand THEN SetResult(m_pCommand->get_CommandTimeout(@lTimeout))
   PROPERTY = lTimeout
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoCommand.CommandTimeout (BYVAL lTimeout AS LONG)
   CADODB_DP("CAdoCommand.CommandTimeout")
   IF m_pCommand THEN SetResult(m_pCommand->put_CommandTimeout(lTimeout))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Sets or returns one or more CommandTypeEnum values.
' ===========================================================================================
PRIVATE PROPERTY CAdoCommand.CommandType (BYVAL lCmdType AS CommandTypeEnum)
   CADODB_DP("CAdoCommand.CommandType")
   IF m_pCommand THEN SetResult(m_pCommand->put_CommandType(lCmdType))
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoCommand.CommandType () AS CommandTypeEnum
   CADODB_DP("CAdoCommand.CommandType")
   DIM lCmdType AS CommandTypeEnum
   IF m_pCommand THEN SetResult(m_pCommand->get_CommandType(@lCmdType))
   RETURN lCmdType
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Sets or returns a Boolean value that, if set to True, indicates that the command should be prepared.
' ===========================================================================================
PRIVATE PROPERTY CAdoCommand.Prepared (BYVAL fPrepared AS BOOLEAN)
   CADODB_DP("CAdoCommand.Prepared")
   IF m_pCommand THEN SetResult(m_pCommand->put_Prepared(cast(VARIANT_BOOL, fPrepared)))
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoCommand.Prepared () AS BOOLEAN
   CADODB_DP("CAdoCommand.Prepared")
   DIM fPrepared AS VARIANT_BOOL
   IF m_pCommand THEN SetResult(m_pCommand->get_Prepared(@fPrepared))
   RETURN fPrepared
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Sets or returns a BSTR value that indicates the name of a Command object.
' ===========================================================================================
PRIVATE PROPERTY CAdoCommand.Name (BYREF cbsName AS CBSTR)
   CADODB_DP("CAdoCommand.Name")
   IF m_pCommand THEN SetResult(m_pCommand->put_Name(cbsName.m_bstr))
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoCommand.Name () AS CBSTR
   CADODB_DP("CAdoCommand.Name")
   DIM bstrName AS AFX_BSTR
   IF m_pCommand THEN SetResult(m_pCommand->get_Name(@bstrName))
   RETURN bstrName
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the dialect of the CommandText or CommandStream properties.
' ===========================================================================================
PRIVATE PROPERTY CAdoCommand.Dialect (BYREF cbsDialect AS CBSTR)
   CADODB_DP("CAdoCommand.Dialect")
   IF m_pCommand THEN SetResult(m_pCommand->put_Dialect(cbsDialect.m_bstr))
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoCommand.Dialect () AS CBSTR
   CADODB_DP("CAdoCommand.Dialect")
   DIM bstrDialect AS AFX_BSTR
   IF m_pCommand THEN SetResult(m_pCommand->get_Dialect(@bstrDialect))
   RETURN bstrDialect
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates whether parameter names should be passed to the provider.
' ===========================================================================================
PRIVATE PROPERTY CAdoCommand.NamedParameters (BYVAL fNamedParameters AS BOOLEAN)
   CADODB_DP("CAdoCommand.NamedParameters")
   IF m_pCommand THEN SetResult(m_pCommand->put_NamedParameters(fNamedParameters))
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoCommand.NamedParameters () AS BOOLEAN
   CADODB_DP("CAdoCommand.NamedParameters")
   DIM fNamedParameters AS VARIANT_BOOL
   IF m_pCommand THEN SetResult(m_pCommand->get_NamedParameters(@fNamedParameters))
   RETURN fNamedParameters
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates whether parameter names should be passed to the provider.
' ===========================================================================================
PRIVATE PROPERTY CAdoCommand.CommandStream (BYVAL pStream AS Afx_ADOStream PTR)
   CADODB_DP("CAdoCommand.CommandStream")
   IF m_pCommand THEN SetResult(m_pCommand->putref_CommandStream(cast(ANY PTR, pStream)))
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoCommand.CommandStream () AS Afx_ADOStream PTR
   CADODB_DP("CAdoCommand.CommandStream")
   DIM vStream AS VARIANT
   IF m_pCommand THEN SetResult(m_pCommand->get_CommandStream(@vStream))
   RETURN cast(Afx_ADOStream PTR, cast(ULONG_PTR, vStream.pdispVal))
   ' Don't clear the vStream variant becuase we are transfering owenership
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates if the Command object is open or closed.
' ===========================================================================================
PRIVATE PROPERTY CAdoCommand.State () AS LONG
   CADODB_DP("CAdoCommand.State")
   DIM lState AS LONG
   IF m_pCommand THEN SetResult(m_pCommand->get_State(@lState))
   RETURN lState
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Cancels execution of a pending, asynchronous method call.
' ===========================================================================================
PRIVATE FUNCTION CAdoCommand.Cancel () AS HRESULT
   CADODB_DP("CAdoCommand.Cancel")
   IF m_pCommand THEN RETURN(SetResult(m_pCommand->Cancel))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Cancels execution of a pending, asynchronous method call.
' ===========================================================================================
PRIVATE FUNCTION CAdoCommand.CreateParameter (BYREF cbsName AS CBSTR = "", BYVAL nType AS DataTypeEnum = 0, _
BYVAL Direction AS ParameterDirectionEnum = adParamInput, BYVAL Size AS LONG = 0, _
BYREF cvValue AS CVAR = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND)) AS Afx_ADOParameter PTR
   CADODB_DP("CAdoCommand.CreateParameter")
   DIM ppiprm AS Afx_ADOParameter PTR
   IF m_pCommand THEN SetResult(m_pCommand->CreateParameter(cbsName.m_bstr, nType, Direction, Size, cvValue, @ppiprm))
   RETURN ppiprm
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Retrieves information from the ADO Errors collection
' ===========================================================================================
PRIVATE FUNCTION CAdoCommand.GetErrorInfo (BYVAL nError AS HRESULT = 0) AS CBSTR
   CADODB_DP("CAdoCommand.GetErrorInfo")
   DIM pConnection AS Afx_ADOCOnnection PTR
   IF m_pCommand THEN SetResult(m_pCommand->get_ActiveConnection(@pConnection))
   DIM cbsError AS CBSTR
   IF pConnection THEN
      cbsError = AfxAdoGetErrorInfo(pConnection, nError)
      pConnection->Release
   END IF
   RETURN cbsError
END FUNCTION
' ===========================================================================================

END NAMESPACE
