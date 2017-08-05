' ########################################################################################
' Microsoft Windows
' Contents: ADO Record class
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once

NAMESPACE Afx

' ########################################################################################
' CAdoRecord class
' ########################################################################################
TYPE CAdoRecord EXTENDS CAdoBase

   m_Result AS HRESULT
   m_pRecord AS Afx_ADORecord PTR

   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYVAL pRecord AS Afx_ADORecord PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   DECLARE DESTRUCTOR
'   DECLARE OPERATOR @ () AS Afx_ADORecord PTR PTR
   DECLARE FUNCTION vptr () AS Afx_ADORecord PTR PTR

   ' // Attaches a reference to an Afx_ADORecord object to the class
   DECLARE SUB Attach (BYVAL pRecord AS Afx_ADORecord PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   ' // Returns a reference to the Properties collection.
   DECLARE PROPERTY Properties () AS Afx_ADOProperties PTR
   ' // Returns a reference to the Fields collection.
   DECLARE PROPERTY Fields () AS Afx_ADOFields PTR
   ' // * Determines the Connection object over which the specified Command object will execute
   DECLARE PROPERTY ActiveConnection (BYREF cbsConn AS CBSTR)
   DECLARE PROPERTY ActiveConnection (BYVAL pconn AS Afx_ADOConnection PTR)
   DECLARE PROPERTY ActiveConnection (BYREF pconn AS CAdoConnection)
   DECLARE PROPERTY ActiveConnection () AS Afx_ADOConnection PTR
   ' // Indicates for all applicable objects whether the state of the object is open or closed.
   DECLARE PROPERTY State () AS ObjectStateEnum
   ' // Indicates the type of Record object.
   DECLARE PROPERTY RecordType () AS RecordTypeEnum
   ' // Indicates the data source or object represented by the Record.
   DECLARE PROPERTY Source () AS CVAR
   DECLARE PROPERTY Source (BYREF cbsSource AS CBSTR)
   DECLARE PROPERTY Source (BYVAL pSource AS Afx_IDispatch PTR)
   ' // Sets or returns a ConnectModeEnum value. The default value for a Record object is adModeRead.
   DECLARE PROPERTY Mode () AS ConnectModeEnum
   DECLARE PROPERTY Mode (BYVAL Mode AS ConnectModeEnum)
   ' // Indicates an absolute URL string that points to the parent Record of the current Record object.
   DECLARE PROPERTY ParentURL () AS CBSTR
   ' // Closes a Record object and any dependent objects.
   DECLARE FUNCTION Close () AS HRESULT
   ' // Cancels execution of a pending, asynchronous method call.
   DECLARE FUNCTION Cancel () AS HRESULT
   ' // Returns a Recordset whose rows represent the children of a collection Record.
   DECLARE FUNCTION GetChildren () AS Afx_ADORecordset PTR
   ' // Opens an existing Record object, or creates a new item represented by the Record (such as a file or directory).
   DECLARE FUNCTION Open (BYREF cvSource AS CVAR = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND), _
         BYREF cvActiveConnection AS CVAR = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND), _
         BYVAL nMode AS ConnectModeEnum = adModeUnknown, BYVAL CreateOptions AS RecordCreateOptionsEnum = adFailIfNotExists, _
         BYREF cbsUserName AS CBSTR = "", BYREF cbsPassword AS CBSTR = "") AS HRESULT
   ' // Moves a entity represented by a Record to another location.
   DECLARE FUNCTION MoveRecord (BYREF cbsSource AS CBSTR = "", BYREF cbsDestination AS CBSTR = "", _
         BYREF cbsUserName AS CBSTR = "", BYREF cbsPassword AS CBSTR = "", BYVAL Options AS MoveRecordOptionsEnum = adMoveUnspecified, _
         BYVAL Async AS BOOLEAN = FALSE) AS CBSTR
   ' // Copies a entity represented by a Record to another location.
   DECLARE FUNCTION CopyRecord (BYREF cbsSource AS CBSTR = "", BYREF cbsDestination AS CBSTR = "", _
         BYREF cbsUserName AS CBSTR = "", BYREF cbsPassword AS CBSTR = "", BYVAL Options AS MoveRecordOptionsEnum = adCopyUnspecified, _
         BYVAL Async AS BOOLEAN = FALSE) AS CBSTR
   ' // Deletes a the entity represented by a Record.
   DECLARE FUNCTION DeleteRecord (BYREF cbsSource AS CBSTR = "", BYVAL Async AS BOOLEAN = FALSE) AS HRESULT
   ' // Retrieves information from the ADO Errors collection
   DECLARE FUNCTION GetErrorInfo (BYVAL nError AS HRESULT = 0) AS CBSTR

END TYPE

' ========================================================================================
' Default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CAdoRecord
   CADODB_DP("-BEGIN CAdoRecord CONSTRUCTOR")
   ' // Create an instance of the Record object
   m_pRecord = AfxNewCom("ADODB.Record")
   CADODB_DP("-END CAdoRecord CONSTRUCTOR " & WSTR(m_pRecord))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CAdoRecord (BYVAL pRecord AS Afx_ADORecord PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   CADODB_DP("-BEGIN CAdoRecord CONSTRUCTOR")
   m_pRecord = pRecord
   IF fAddRef THEN
      IF m_pRecord THEN m_pRecord->AddRef
   END IF
END CONSTRUCTOR
' ========================================================================================

' ===========================================================================================
' Cleanup
' ===========================================================================================
PRIVATE DESTRUCTOR CAdoRecord
   CADODB_DP("CAdoRecord DESTRUCTOR " & WSTR(m_pRecord))
   IF m_pRecord THEN
      IF this.State = adStateOpen THEN this.Close
      m_pRecord->Release
   END IF
END DESTRUCTOR
' ===========================================================================================

' ========================================================================================
' Returns the address of the connection pointer
' ========================================================================================
'PRIVATE OPERATOR CAdoRecord.@ () AS Afx_ADORecord PTR PTR
'   OPERATOR = @m_pRecord
'END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CAdoRecord.vptr () AS Afx_ADORecord PTR PTR
   CVAR_DP("CAdoRecord vptr")
   IF m_pRecord THEN m_pRecord->Release
   RETURN @m_pRecord
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the encapsulated IDispatch pointer
' ========================================================================================
PRIVATE OPERATOR * (BYREF cRecord AS CAdoRecord) AS Afx_ADORecord PTR
   OPERATOR = cRecord.m_pRecord
END OPERATOR
' ========================================================================================

' ===========================================================================================
' Attaches a reference to an Afx_ADORecord object to the class
' ===========================================================================================
PRIVATE SUB CAdoRecord.Attach (BYVAL pRecord AS Afx_ADORecord PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   CADODB_DP("CAdoRecord.Attach")
   IF m_pRecord THEN m_pRecord->Release
   m_pRecord = pRecord
   ' // Increase the reference count
   IF fAddRef THEN
      IF m_pRecord THEN m_pRecord->AddRef
   END IF
END SUB
' ===========================================================================================

' ===========================================================================================
' * Returns a reference to the Properties collection.
' ===========================================================================================
PRIVATE PROPERTY CAdoRecord.Properties () AS Afx_ADOProperties PTR
   CADODB_DP("CAdoRecord.Properties")
   DIM pProperties AS Afx_ADOProperties PTR
   IF m_pRecord THEN SetResult(m_pRecord->get_Properties(@pProperties))
   PROPERTY = pProperties
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' * Returns a reference to the Fields collection.
' ===========================================================================================
PRIVATE PROPERTY CAdoRecord.Fields () AS Afx_ADOFields PTR
   CADODB_DP("CAdoRecord.Fields")
   DIM pFields AS Afx_ADOFields PTR
   IF m_pRecord THEN SetResult(m_pRecord->get_Fields(@pFields))
   PROPERTY = pFields
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Sets a reference to the active connection
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoRecord.ActiveConnection (BYREF cbsConn AS CBSTR)
   CADODB_DP("CAdoRecord.ActiveConnection - CVAR")
   IF m_pRecord THEN SetResult(m_pRecord->put_ActiveConnection(cbsConn.m_bstr))
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoRecord.ActiveConnection (BYVAL pconn AS Afx_ADOConnection PTR)
   CADODB_DP("CAdoRecordCAdoRecord.ActiveConnection")
   IF m_pRecord THEN SetResult(m_pRecord->putref_ActiveConnection(pconn))
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoRecord.ActiveConnection (BYREF pconn AS CAdoConnection)
   CADODB_DP("CAdoRecordCAdoRecord.ActiveConnection")
   IF m_pRecord THEN SetResult(m_pRecord->putref_ActiveConnection(pconn.m_pConnection))
END PROPERTY
' ===========================================================================================
' ===========================================================================================
' * Gets a reference to the active connection
' ===========================================================================================
PRIVATE PROPERTY CAdoRecord.ActiveConnection () AS Afx_ADOConnection PTR
   CADODB_DP("CAdoRecordCAdoRecord.ActiveConnection")
   DIM vConnection AS VARIANT
   IF m_pRecord THEN SetResult(m_pRecord->get_ActiveConnection(@vConnection))
   DIM pConnection AS Afx_ADOConnection PTR
   IF vConnection.vt = VT_DISPATCH THEN
      pConnection = cast(Afx_ADOConnection PTR, cast(ULONG_PTR, vConnection.pdispVal))
      IF pConnection THEN pConnection->AddRef
   END IF
   PROPERTY = pConnection
   VariantClear(@vConnection)
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates for all applicable objects whether the state of the object is open or closed.
' ===========================================================================================
PRIVATE PROPERTY CAdoRecord.State () AS ObjectStateEnum
   CADODB_DP("CAdoRecord.State")
   DIM lState AS ObjectStateEnum
   IF m_pRecord THEN SetResult(m_pRecord->get_State(@lState))
   RETURN lState
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the type of Record object.
' ===========================================================================================
PRIVATE PROPERTY CAdoRecord.RecordType () AS RecordTypeEnum
   CADODB_DP("CAdoRecord.RecordType")
   DIM lType AS LONG
   IF m_pRecord THEN SetResult(m_pRecord->get_RecordType(@lType))
   RETURN lType
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the data source or object represented by the Record.
' ===========================================================================================
PRIVATE PROPERTY CAdoRecord.Source () AS CVAR
   CADODB_DP("CAdoRecord.Source")
   DIM cvSource AS VARIANT
   IF m_pRecord THEN SetResult(m_pRecord->get_Source(@cvSource))
   RETURN cvSource
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoRecord.Source (BYREF cbSource AS CBSTR)
   CADODB_DP("CAdoRecord.Source")
   IF m_pRecord THEN SetResult(m_pRecord->put_Source(cbSource.m_bstr))
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoRecord.Source (BYVAL pSource AS Afx_IDispatch PTR)
   CADODB_DP("CAdoRecord.Source")
   IF m_pRecord THEN SetResult(m_pRecord->putref_Source(pSource))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Sets or returns a ConnectModeEnum value. The default value for a Record object is adModeRead.
' ===========================================================================================
PRIVATE PROPERTY CAdoRecord.Mode () AS ConnectModeEnum
   CADODB_DP("CAdoRecord.Mode")
   DIM lMode AS LONG
   IF m_pRecord THEN SetResult(m_pRecord->get_Mode(@lMode))
   RETURN lMode
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoRecord.Mode (BYVAL lMode AS ConnectModeEnum)
   CADODB_DP("CAdoRecord.Mode")
   IF m_pRecord THEN SetResult(m_pRecord->put_Mode(lMode))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates an absolute URL string that points to the parent Record of the current Record object.
' ===========================================================================================
PRIVATE PROPERTY CAdoRecord.ParentURL () AS CBSTR
   CADODB_DP("CAdoRecord.ParentURL")
   DIM bstrURL AS AFX_BSTR
   IF m_pRecord THEN SetResult(m_pRecord->get_ParentURL(@bstrURL))
   RETURN bstrURL
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Closes a Record object and any dependent objects.
' ===========================================================================================
PRIVATE FUNCTION CAdoRecord.Close () AS HRESULT
   CADODB_DP("CAdoRecord.Close")
   IF m_pRecord THEN RETURN(SetResult(m_pRecord->Close))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Cancels execution of a pending, asynchronous method call.
' ===========================================================================================
PRIVATE FUNCTION CAdoRecord.Cancel () AS HRESULT
   CADODB_DP("CAdoRecord.Cancel")
   IF m_pRecord THEN RETURN(SetResult(m_pRecord->Cancel))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Returns a Recordset whose rows represent the children of a collection Record.
' ===========================================================================================
PRIVATE FUNCTION CAdoRecord.GetChildren () AS Afx_ADORecordset PTR
   CADODB_DP("CAdoRecord.GetChildren")
   DIM pRSet AS Afx_ADORecordset PTR
   IF m_pRecord THEN SetResult(m_pRecord->GetChildren(@pRSet))
   RETURN pRSet
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Opens an existing Record object, or creates a new item represented by the Record (such as a file or directory).
' ===========================================================================================
PRIVATE FUNCTION CAdoRecord.Open (BYREF cvSource AS CVAR = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND), _
BYREF cvActiveConnection AS CVAR = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND), _
BYVAL nMode AS ConnectModeEnum = adModeUnknown, BYVAL CreateOptions AS RecordCreateOptionsEnum = adFailIfNotExists, _
BYREF cbsUserName AS CBSTR = "", BYREF cbsPassword AS CBSTR = "") AS HRESULT
   CADODB_DP("CAdoRecord.Open")
   IF m_pRecord THEN RETURN(SetResult(m_pRecord->Open(cvSource, cvActiveConnection, nMode, CreateOptions, cbsUserName.m_bstr, cbsPassword.m_bstr)))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Moves a entity represented by a Record to another location.
' ===========================================================================================
PRIVATE FUNCTION CAdoRecord.MoveRecord (BYREF cbsSource AS CBSTR = "", BYREF cbsDestination AS CBSTR = "", _
BYREF cbsUserName AS CBSTR = "", BYREF cbsPassword AS CBSTR = "", BYVAL Options AS MoveRecordOptionsEnum = adMoveUnspecified, _
BYVAL Async AS BOOLEAN = FALSE) AS CBSTR
   CADODB_DP("CAdoRecord.MoveRecord")
   DIM bstrNewURL AS AFX_BSTR
   IF m_pRecord THEN m_pRecord->MoveRecord(cbsSource, cbsDestination, cbsUserName.m_bstr, cbsPassword.m_bstr, Options, Async, @bstrNewURL)
   RETURN bstrNewURL
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Copies a entity represented by a Record to another location.
' ===========================================================================================
PRIVATE FUNCTION CAdoRecord.CopyRecord (BYREF cbsSource AS CBSTR = "", BYREF cbsDestination AS CBSTR = "", _
BYREF cbsUserName AS CBSTR = "", BYREF cbsPassword AS CBSTR = "", BYVAL Options AS MoveRecordOptionsEnum = adMoveUnspecified, _
BYVAL Async AS BOOLEAN = FALSE) AS CBSTR
   CADODB_DP("CAdoRecord.CopyRecord")
   DIM bstrNewURL AS AFX_BSTR
   IF m_pRecord THEN m_pRecord->CopyRecord(cbsSource.m_bstr, cbsDestination.m_bstr, cbsUserName.m_bstr, cbsPassword.m_bstr, Options, Async, @bstrNewURL)
   RETURN bstrNewURL
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Deletes a the entity represented by a Record.
' ===========================================================================================
PRIVATE FUNCTION CAdoRecord.DeleteRecord (BYREF cbsSource AS CBSTR = "", BYVAL Async AS BOOLEAN = FALSE) AS HRESULT
   CADODB_DP("CAdoRecord.DeleteRecord")
   IF m_pRecord THEN RETURN(m_pRecord->DeleteRecord(cbsSource.m_bstr, Async))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Retrieves information from the ADO Errors collection
' ===========================================================================================
PRIVATE FUNCTION CAdoRecord.GetErrorInfo (BYVAL nError AS HRESULT = 0) AS CBSTR
   CADODB_DP("CAdoRecord.GetErrorInfo")
   DIM vConnection AS VARIANT
   IF m_pRecord THEN m_pRecord->get_ActiveConnection(@vConnection)
   DIM cbsError AS CBSTR
   IF vConnection.vt = VT_DISPATCH THEN
      cbsError = AfxAdoGetErrorInfo(cast(Afx_ADOConnection PTR, cast(ULONG_PTR, vConnection.pdispVal)), nError)
   END IF
   VariantClear(@vConnection)
   RETURN cbsError
END FUNCTION
' ===========================================================================================

END NAMESPACE
