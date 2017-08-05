' ########################################################################################
' Microsoft Windows
' Contents: ADO Recordset class
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once

NAMESPACE Afx

' ########################################################################################
' CAdoRecordset class
' ########################################################################################
TYPE CAdoRecordset EXTENDS CAdoBase

   m_Result AS HRESULT
   m_pRecordset AS Afx_ADORecordset PTR

   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYVAL pRecordset AS Afx_AdoRecordset PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   DECLARE DESTRUCTOR
'   DECLARE OPERATOR @ () AS Afx_ADORecordset PTR PTR
   DECLARE FUNCTION vptr () AS Afx_ADORecordset PTR PTR

   DECLARE FUNCTION Open (BYREF cvSource AS CVAR, BYREF cActiveConnection AS CAdoConnection, _
           BYVAL nCursorType AS CursorTypeEnum = adOpenUnspecified, BYVAL nLockType AS LockTypeEnum = adLockUnspecified, _
           BYVAL nOptions AS LONG = adCmdUnspecified) AS HRESULT
   DECLARE FUNCTION Open (BYREF cvSource AS CVAR, BYREF cvActiveConnection AS CVAR = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND), _
           BYVAL nCursorType AS CursorTypeEnum = adOpenUnspecified, BYVAL nLockType AS LockTypeEnum = adLockUnspecified, _
           BYVAL nOptions AS LONG = adCmdUnspecified) AS HRESULT
   DECLARE FUNCTION Open (BYREF cvSource AS CVAR, BYVAL pActiveConnection AS CAdoConnection PTR, _
           BYVAL nCursorType AS CursorTypeEnum = adOpenUnspecified, BYVAL nLockType AS LockTypeEnum = adLockUnspecified, _
           BYVAL nOptions AS LONG = adCmdUnspecified) AS HRESULT
   DECLARE FUNCTION Open (BYVAL nCursorType AS CursorTypeEnum = adOpenUnspecified, _
           BYVAL nLockType AS LockTypeEnum = adLockUnspecified, BYVAL nOptions AS LONG = adCmdUnspecified) AS HRESULT
   ' // Attaches a recordset to the class
   DECLARE FUNCTION Attach (BYVAL pRecordset AS AFX_ADORecordset PTR, BYVAL fAddRef AS BOOLEAN = FALSE) AS HRESULT
   ' // * Returns a reference to the Properties collection
   DECLARE PROPERTY Properties () AS Afx_ADOProperties PTR
   ' // Gets a reference to the Fields collection of a Record object
   DECLARE PROPERTY Fields () AS Afx_ADOFields PTR
   ' // Indicates that the current record position is before the first record in a Recordset object
   DECLARE PROPERTY BOF () AS BOOLEAN
   ' // * Indicates that the current record position is after the last record in a Recordset object
   DECLARE PROPERTY EOF () AS BOOLEAN
   ' // Indicates the Command object that created the associated Recordset object
   DECLARE FUNCTION ActiveCommand () AS Afx_ADOCommand PTR
   ' // * Sets a reference to the active connection
   DECLARE PROPERTY ActiveConnection (BYREF vConn AS CVAR)
   DECLARE PROPERTY ActiveConnection (BYVAL pconn AS Afx_ADOConnection PTR)
   DECLARE PROPERTY ActiveConnection (BYREF pconn AS CAdoConnection)
   ' // * Gets a reference to the active connection
   DECLARE PROPERTY ActiveConnection () AS Afx_ADOConnection PTR
   ' // * Indicates the data source for a Recordset object.
   DECLARE PROPERTY Source (BYREF cbsConn AS CBSTR)
   DECLARE PROPERTY Source (BYVAL pcmd AS Afx_ADOCommand PTR)
   DECLARE PROPERTY Source () AS CVAR
   ' // * Sets or returns a Variant expression that evaluates to a valid bookmark
   DECLARE PROPERTY Bookmark () AS CVAR
   DECLARE PROPERTY Bookmark (BYREF cvBookmark AS CVAR)
   ' // * Sets or returns a Long value that must be greater than 0. Default is 1.
   DECLARE PROPERTY CacheSize () AS LONG
   DECLARE PROPERTY CacheSize (BYVAL size AS LONG)
   ' // * Sets or returns a CursorTypeEnum value. The default value is adOpenForwardOnly.
   DECLARE PROPERTY CursorType () AS CursorTypeEnum
   DECLARE PROPERTY CursorType (BYVAL lCursorType AS CursorTypeEnum)
   ' // * Indicates the maximum number of records to return to a Recordset from a query.
   DECLARE PROPERTY CursorLocation () AS CursorLocationEnum
   DECLARE PROPERTY CursorLocation (BYVAL lCursorLoc AS CursorLocationEnum)
   ' // * Sets or returns the lock type, a Long value that must be greater than 0. Default is 1.
   DECLARE PROPERTY LockType () AS LockTypeEnum
   DECLARE PROPERTY LockType (BYVAL lLockType AS LockTypeEnum)
   ' // Indicates the maximum number of records to return to a Recordset from a query.
   DECLARE PROPERTY MaxRecords () AS LONG
   DECLARE PROPERTY MaxRecords (BYVAL lMaxRecords AS LONG)
   ' // * Indicates how many records constitute one page in the Recordset.
   DECLARE PROPERTY PageSize () AS LONG
   DECLARE PROPERTY PageSize (BYVAL nPageSize AS LONG)
   ' // * Indicates how many pages of data the Recordset object contains
   DECLARE PROPERTY PageCount () AS LONG
   ' // * Indicates the number of records in a Recordset object.
   DECLARE PROPERTY RecordCount () AS LONG
   ' // Indicates which records are to be marshaled back to the server
   DECLARE PROPERTY MarshalOptions () AS MarshalOptionsEnum
   DECLARE PROPERTY MarshalOptions (BYVAL eMarshal AS MarshalOptionsEnum)
   ' // * Sets or returns a Variant value that indicates the value of the object
   DECLARE PROPERTY Collect (BYREF cvIndex AS CVAR) AS CVAR
   DECLARE PROPERTY Collect (BYREF cvIndex AS CVAR, BYREF cvValue AS CVAR)
   ' // Indicates the name of the index currently in effect for a Recordset object
   DECLARE PROPERTY Index () AS CBSTR
   DECLARE PROPERTY Index (BYREF cbsIndex AS CBSTR)
   ' // * Returns the Recordset as a string
   DECLARE FUNCTION GetString (BYVAL StringFormat AS StringFormatEnum = adClipString, BYVAL NumRows AS LONG = adReadAll, _
           BYREF ColumnDelimeter AS CBSTR = "", BYREF RowDelimeter AS CBSTR = "", BYREF NullExpr AS CBSTR = "") AS CBSTR
   ' // Indicates an object that contains data to be represented as a Recordset object
   DECLARE PROPERTY DataSource () AS Afx_IUnknown PTR
   DECLARE PROPERTY DataSource (BYVAL punkDataSource AS Afx_IUnknown PTR)
   ' // Indicates the name of the data member that will be retrieved from the object referenced by the DataSource property
   DECLARE PROPERTY DataMember () AS CBSTR
   DECLARE PROPERTY DataMember (BYREF cbsDataMember AS CBSTR)
   ' // * Sets or returns a value that indicates the current record position within the recordset
   DECLARE PROPERTY AbsolutePosition () AS PositionEnum
   DECLARE PROPERTY AbsolutePosition (BYVAL Position AS PositionEnum)
   ' // * Sets or returns a value that indicates the current page position within the recordset
   DECLARE PROPERTY AbsolutePage () AS PositionEnum
   DECLARE PROPERTY AbsolutePage (BYVAL nPage AS PositionEnum)
   ' // * Moves the position of the current record in a Recordset object
   DECLARE FUNCTION Move (BYVAL NumRecords AS LONG, BYVAL start AS BookmarkEnum = adBookmarkCurrent) AS HRESULT
   DECLARE FUNCTION Move (BYVAL NumRecords AS LONG, BYREF cvStart AS CVAR) AS HRESULT
   ' // * Moves the current record position to the next record in the Recordset
   DECLARE FUNCTION MoveNext () AS HRESULT
   ' // Moves to the previous record in a specified Recordset object and makes that record the current record
   DECLARE FUNCTION MovePrevious () AS HRESULT
   ' // Moves to the first record in a specified Recordset object and makes that record the current record
   DECLARE FUNCTION MoveFirst () AS HRESULT
   ' // Moves to the last record in a specified Recordset object and makes that record the current record
   DECLARE FUNCTION MoveLast () AS HRESULT
   ' // Indicates the status of the current record with respect to batch updates or other bulk operations
   DECLARE PROPERTY Status () AS LONG
   ' // Indicates for whether the state of the Recordset object is open or closed
   DECLARE PROPERTY State () AS LONG
   ' // Indicates the editing status of the current record
   DECLARE PROPERTY EditMode () AS EditModeEnum
   ' // Indicates, in a hierarchical Recordset object, whether the reference to the underlying child records
   ' // (that is, the chapter) changes when the parent row position changes.
   DECLARE PROPERTY StayInSync () AS BOOLEAN
   DECLARE PROPERTY StayInSync (BYVAL bStayInSync AS BOOLEAN)
   ' // Refreshes the data in the current Recordset object from the underlying database
   DECLARE FUNCTION Resync (BYVAL AffectRecords AS AffectEnum = adAffectAll, BYVAL ResyncValues AS ResyncEnum = adResyncAllValues) AS HRESULT
   ' // Updates the data in a Recordset object by re-executing the query on which the object is based
   DECLARE FUNCTION Requery (BYVAL Options AS LONG = adOpenUnspecified) AS HRESULT
   ' // Moves to the previous record in a specified Recordset object and makes that record the current record
   DECLARE FUNCTION NextRecordset (BYVAL RecordsAffected AS LONG PTR = NULL) AS Afx_AdoRecordset PTR
   ' // Determines whether a specified Recordset object supports a particular type of functionality
   DECLARE FUNCTION Supports (BYVAL CursorOptions AS CursorOptionEnum) AS BOOLEAN
   ' // Writes all pending batch updates to disk
   DECLARE FUNCTION UpdateBatch (BYVAL AffectRecords AS AffectEnum = adAffectAll) AS HRESULT
   ' // Cancels a pending batch update
   DECLARE FUNCTION CancelBatch (BYVAL AffectRecords AS AffectEnum = adAffectAll) AS HRESULT
   ' // Cancels any changes made to the current or new row of a Recordset object before calling the Update method
   DECLARE FUNCTION CancelUpdate () AS HRESULT
   ' // Closes a Recordset object and any dependent objects
   DECLARE FUNCTION Close () AS HRESULT
   ' // Creates a duplicate Recordset object from an existing Recordset object. Optionally, specifies that the clone be read-only
   DECLARE FUNCTION Clone (BYVAL nLockType AS LockTypeEnum = adLockUnspecified) AS Afx_AdoRecordset PTR
   ' // Cancels execution of a pending, asynchronous method call
   DECLARE FUNCTION Cancel () AS HRESULT
   ' // Indicates a filter for data in a Recordset
   DECLARE PROPERTY Filter () AS CVAR
   DECLARE PROPERTY Filter (BYVAL cvCriteria AS CVAR)
   ' // Indicates one or more field names on which the Recordset is sorted,
   ' // and whether each field is sorted in ascending or descending order
   DECLARE PROPERTY Sort () AS CBSTR
   DECLARE PROPERTY Sort (BYREF cbsCriteria AS CBSTR)
   ' // Deletes the current record or a group of records.
   DECLARE FUNCTION Delete_ (BYVAL AffectRecords AS AffectEnum = adAffectCurrent) AS HRESULT
   ' // Saves the Recordset in a file
   DECLARE FUNCTION Save (BYREF cvDestination AS CVAR, BYVAL PersistFormat AS PersistFormatEnum = adPersistADTG) AS HRESULT
   ' // Saves the Recordset in a Stream object
   DECLARE FUNCTION Save (BYVAL pDestination AS Afx_AdoStream PTR, BYVAL PersistFormat AS PersistFormatEnum = adPersistADTG) AS HRESULT
   ' // Compares two bookmarks and returns an indication of their relative values.
   DECLARE FUNCTION CompareBookmarks (BYREF cvBookmark1 AS CVAR, BYREF cvBookmark2 AS CVAR) AS CompareEnum
   ' // * Searches a Recordset for the row that satisfies the specified criteria
   DECLARE FUNCTION Find (BYREF cbsCriteria AS CBSTR, BYREF cvStart AS CVAR = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND), _
           BYVAL SkipRecords AS LONG = 0, BYVAL SearchDirection AS SearchDirectionEnum = adSearchForward) AS HRESULT
   ' // * Retrieves multiple records of a Recordset object into an array
   DECLARE FUNCTION GetRows (BYVAL Rows AS LONG = adGetRowsRest, BYREF cvStart AS CVAR = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND), BYREF cvFields AS CVAR = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND)) AS SAFEARRAY PTR
   ' // * Searches the index of a Recordset to quickly locate the row that matches the specified values, and changes the current row position to that row
   DECLARE FUNCTION Seek (BYREF cvKeyValues AS CVAR, BYVAL SeekOption AS SeekEnum = adSeekFirstEQ) AS HRESULT
   ' // Creates a new record for an updatable Recordset object.
   DECLARE FUNCTION AddNew (BYREF cvFieldList AS CVAR = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND), BYREF cvValues AS CVAR = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND)) AS HRESULT
   ' // Saves any changes you make to the current row of a Recordset object
   DECLARE FUNCTION Update (BYREF cvFieldList AS CVAR = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND), BYREF cvValues AS CVAR = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND)) AS HRESULT
   ' // Retrieves information from the ADO Errors collection
   DECLARE FUNCTION GetErrorInfo (BYVAL nError AS HRESULT = 0) AS CBSTR

END TYPE

' ========================================================================================
' Default constructor
' ========================================================================================
CONSTRUCTOR CAdoRecordset
   CADODB_DP("CAdoRecordset CONSTRUCTOR")
   ' // Create an instance of the Recordset object
'   DIM CLSID_Recordset_ AS GUID = (&h00000535, &h0000, &h0010, {&h80, &h00, &h00, &hAA, &h00, &h6D, &h2E, &hA4})
'   DIM IID_Recordset_ AS GUID = (&h00001556, &h0000, &h0010, {&h80, &h00, &h00, &hAA, &h00, &h6D, &h2E, &hA4})
'   SetResult(CoCreateInstance(@CLSID_Recordset_, NULL, CLSCTX_INPROC_SERVER, @IID_Recordset_, @m_pRecordset))
   m_pRecordset = AfxNewCom("ADODB.Recordset")
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
CONSTRUCTOR CAdoRecordset (BYVAL pRecordset AS Afx_ADORecordset PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   CADODB_DP("CAdoRecordset CONSTRUCTOR")
   m_pRecordset = pRecordset
   IF fAddRef THEN
      IF m_pRecordset THEN m_pRecordset->AddRef
   END IF
END CONSTRUCTOR
' ========================================================================================

' ===========================================================================================
' Cleanup
' ===========================================================================================
DESTRUCTOR CAdoRecordset
   CADODB_DP("CAdoRecordset DESTRUCTOR")
   ' // Close and release the recordset
   IF m_pRecordset THEN
      IF this.State = adStateOpen THEN this.Close
      m_pRecordset->Release
   END IF
END DESTRUCTOR
' ===========================================================================================

' ========================================================================================
' Returns the address of the connection pointer
' ========================================================================================
'PRIVATE OPERATOR CAdoRecordset.@ () AS Afx_ADORecordset PTR PTR
'   CADODB_DP("CAdoRecordset Operator @")
'   OPERATOR = @m_pRecordset
'END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CAdoRecordset.vptr () AS Afx_ADORecordset PTR PTR
   CVAR_DP("CAdoRecordset vptr")
   IF m_pRecordset THEN m_pRecordset->Release
   RETURN @m_pRecordset
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the encapsulated IDispatch pointer
' ========================================================================================
PRIVATE OPERATOR * (BYREF crs AS CAdoRecordset) AS Afx_ADORecordset PTR
   CADODB_DP("CAdoRecordset Operator *")
   OPERATOR = crs.m_pRecordset
END OPERATOR
' ========================================================================================

' ===========================================================================================
' Opens a connection to a data source.
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.Open (BYREF cvSource AS CVAR, BYREF cActiveConnection AS CAdoConnection, _
BYVAL nCursorType AS CursorTypeEnum = adOpenUnspecified, BYVAL nLockType AS LockTypeEnum = adLockUnspecified, _
BYVAL nOptions AS LONG = adCmdUnspecified) AS HRESULT
   CADODB_DP("CAdoRecordset.Open - cvSource - cActiveConnection")
   ' // Make a VT_DISPATCH variant with the active connection pointer
   DIM vConnection AS VARIANT
   vConnection.vt = VT_DISPATCH
   vConnection.pdispVal = cast(IDispatch PTR, cast(ULONG_PTR, cActiveConnection.m_pConnection))
   IF vConnection.pdispVal THEN vConnection.pdispVal->lpVtbl->AddRef(vConnection.pdispVal)
   ' // Call the Open method
   IF m_pRecordset THEN SetResult(m_pRecordset->Open(cvSource, vConnection, nCursorType, nLockType, nOptions))
   ' // Clear the variant
   VariantClear(@vConnection)
   ' // Return the result code
   RETURN m_Result
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Opens a connection to a data source.
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.Open (BYREF cvSource AS CVAR, BYREF cvActiveConnection AS CVAR = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND), _
BYVAL nCursorType AS CursorTypeEnum = adOpenUnspecified, BYVAL nLockType AS LockTypeEnum = adLockUnspecified, _
BYVAL nOptions AS LONG = adCmdUnspecified) AS HRESULT
   CADODB_DP("CAdoRecordset.Open - cvSource - cvActiveConnection")
   IF m_pRecordset THEN SetResult(m_pRecordset->Open(cvSource, cvActiveConnection, nCursorType, nLockType, nOptions))
   RETURN m_Result
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' * Opens a connection to a data source.
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.Open (BYREF cvSource AS CVAR, BYVAL pActiveConnection AS CAdoConnection PTR, _
BYVAL nCursorType AS CursorTypeEnum = adOpenUnspecified, BYVAL nLockType AS LockTypeEnum = adLockUnspecified, _
BYVAL nOptions AS LONG = adCmdUnspecified) AS HRESULT
   CADODB_DP("CAdoRecordset.Open - cvSource - pActiveConnection")

   ' // Make a VT_DISPATCH variant with the active connection pointer
   DIM vConnection AS VARIANT
   vConnection.vt = VT_DISPATCH
   vConnection.pdispVal = cast(IDispatch PTR, cast(ULONG_PTR, pActiveConnection->m_pConnection))
   IF vConnection.pdispVal THEN vConnection.pdispVal->lpVtbl->AddRef(vConnection.pdispVal)
   ' // Call the Open method
   IF m_pRecordset THEN SetResult(m_pRecordset->Open(cvSource, vConnection, nCursorType, nLockType, nOptions))
   ' // Clear the variant
   VariantClear(@vConnection)
   ' // Return the result code
   RETURN m_Result

END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Opens a connection to a data source.
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.Open (BYVAL nCursorType AS CursorTypeEnum = adOpenUnspecified, _
BYVAL nLockType AS LockTypeEnum = adLockUnspecified, BYVAL nOptions AS LONG = adCmdUnspecified) AS HRESULT
   CADODB_DP("CAdoRecordset.Open")
   ' // Call the Open method
   DIM vSource AS VARIANT = TYPE(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND)
   DIM vConnection AS VARIANT = TYPE(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND)
   IF m_pRecordset THEN SetResult(m_pRecordset->Open(vSource, vConnection, nCursorType, nLockType, nOptions))
   ' // Return the result code
   RETURN m_Result
END FUNCTION
' ===========================================================================================

' ========================================================================================
' Attaches a recordset to the class.
' Useful to attach to the class a pointer returned by a function.
' ========================================================================================
PRIVATE FUNCTION CAdoRecordset.Attach (BYVAL pRecordset AS AFX_ADORecordset PTR, BYVAL fAddRef AS BOOLEAN = FALSE) AS HRESULT
   CADODB_DP("CAdoRecordset.Attach - pRecordset")
   IF m_pRecordset = NULL THEN RETURN E_INVALIDARG
   ' // Close and release the recordset
   IF this.State = adStateOpen THEN this.Close
   m_pRecordset->Release
   ' // Attach the recordset
   m_pRecordset = pRecordset
   ' // Increase the reference count
   IF fAddRef THEN
      IF m_pRecordset THEN m_pRecordset->AddRef
   END IF
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ===========================================================================================
' * Returns a reference to the Properties collection.
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.Properties () AS Afx_ADOProperties PTR
   CADODB_DP("CAdoRecordset.Properties")
   DIM pProperties AS Afx_ADOProperties PTR
   IF m_pRecordset THEN SetResult(m_pRecordset->get_Properties(@pProperties))
   PROPERTY = pProperties
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' * Returns a reference to the Fields collection.
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.Fields () AS Afx_ADOFields PTR
   CADODB_DP("CAdoRecordset.Fields")
   DIM pFields AS Afx_ADOFields PTR
   IF m_pRecordset THEN SetResult(m_pRecordset->get_Fields(@pFields))
   PROPERTY = pFields
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates that the current record position is before the first record in a Recordset object
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.BOF () AS BOOLEAN
   CADODB_DP("CAdoRecordset.BOF")
   DIM pb AS VARIANT_BOOL
   IF m_pRecordset THEN SetResult(m_pRecordset->get_BOF(@pb))
   PROPERTY = pb
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' * Indicates that the current record position is after the last record in a Recordset object
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.EOF () AS BOOLEAN
   CADODB_DP("CAdoRecordset.EOF")
   DIM pb AS VARIANT_BOOL
   IF m_pRecordset THEN SetResult(m_pRecordset->get_EOF(@pb))
   PROPERTY = pb
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the Command object that created the associated Recordset object
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.ActiveCommand () AS Afx_ADOCommand PTR
   DIM pcmd AS Afx_IDispatch PTR
   IF m_pRecordset THEN SetResult(m_pRecordset->get_ActiveCommand(@pcmd))
   RETURN cast(Afx_ADOCommand PTR, cast(ULONG_PTR, pcmd))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' * Sets a reference to the active connection
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.ActiveConnection (BYVAL pconn AS Afx_ADOConnection PTR)
   CADODB_DP("CAdoRecordset.ActiveConnection - Afx_ADOConnection")
   IF m_pRecordset THEN SetResult(m_pRecordset->putref_ActiveConnection(pconn))
END PROPERTY
' ===========================================================================================
' ===========================================================================================
' * Sets a reference to the active connection
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.ActiveConnection (BYREF pconn AS CAdoConnection)
   CADODB_DP("CAdoRecordset.ActiveConnection - CAdoConnection")
   IF m_pRecordset THEN SetResult(m_pRecordset->putref_ActiveConnection(pconn.m_pConnection))
END PROPERTY
' ===========================================================================================
' ===========================================================================================
' * Sets a reference to the active connection
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.ActiveConnection (BYREF vConn AS CVAR)
   CADODB_DP("CAdoRecordset.ActiveConnection - CVAR")
   IF m_pRecordset THEN SetResult(m_pRecordset->put_ActiveConnection(vConn))
END PROPERTY
' ===========================================================================================
' ===========================================================================================
' * Gets a reference to the active connection
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.ActiveConnection () AS Afx_ADOConnection PTR
   CADODB_DP("CAdoRecordset.ActiveConnection - AS Afx_ADOConnection PTR")
   DIM vConnection AS VARIANT
   IF m_pRecordset THEN SetResult(m_pRecordset->get_ActiveConnection(@vConnection))
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
' * Indicates the data source for a Recordset object
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.Source (BYREF cbsConn AS CBSTR)
   CADODB_DP("CAdoRecordset.Source - CBSTR")
   IF m_pRecordset THEN SetResult(m_pRecordset->put_Source(cbsConn.m_bstr))
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.Source (BYVAL pcmd AS Afx_ADOCommand PTR)
   CADODB_DP("CAdoRecordset.Source - IDispatch PTR")
   IF m_pRecordset THEN SetResult(m_pRecordset->putref_Source(pcmd))
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.Source () AS CVAR
   CADODB_DP("CAdoRecordset.Source - AS CVAR")
   DIM cvSource AS CVAR
   IF m_pRecordset THEN SetResult(m_pRecordset->get_Source(cvSource.vptr))
   PROPERTY = cvSource
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' * Sets or returns a Variant expression that evaluates to a valid bookmark
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.Bookmark () AS CVAR
   CADODB_DP("CAdoRecordset.Bookmark - AS CVAR")
   DIM cvSource AS CVAR
   IF m_pRecordset THEN SetResult(m_pRecordset->get_Bookmark(cvSource.vptr))
   PROPERTY = cvSource
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.Bookmark (BYREF cvBookmark AS CVAR)
   CADODB_DP("CAdoRecordset.Bookmark - CVAR")
   IF m_pRecordset THEN SetResult(m_pRecordset->put_Bookmark(cvBookmark))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' * Sets or returns a Long value that must be greater than 0. Default is 1.
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.CacheSize () AS LONG
   CADODB_DP("CAdoRecordset.CacheSize - AS LONG")
   DIM size AS LONG
   IF m_pRecordset THEN SetResult(m_pRecordset->get_CacheSize(@size))
   PROPERTY = size
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.CacheSize (BYVAL size AS LONG)
   CADODB_DP("CAdoRecordset.CacheSize")
   IF m_pRecordset THEN SetResult(m_pRecordset->put_CacheSize(size))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' * Sets or returns a CursorTypeEnum value. The default value is adOpenForwardOnly.
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.CursorType () AS CursorTypeEnum
   CADODB_DP("CAdoRecordset.CursorType - AS CursorTypeEnum")
   DIM lCursorType AS CursorTypeEnum
   IF m_pRecordset THEN SetResult(m_pRecordset->get_CursorType(@lCursorType))
   PROPERTY = lCursorType
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.CursorType (BYVAL lCursorType AS CursorTypeEnum)
   CADODB_DP("CAdoRecordset.CursorType")
   IF m_pRecordset THEN SetResult(m_pRecordset->put_CursorType(lCursorType))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the location of the cursor service
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.CursorLocation () AS CursorLocationEnum
   CADODB_DP("CAdoRecordset.CursorLocation - AS CursorLocationEnum")
   DIM lCursorLoc AS CursorLocationEnum
   IF m_pRecordset THEN SetResult(m_pRecordset->get_CursorLocation(@lCursorLoc))
   PROPERTY = lCursorLoc
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.CursorLocation (BYVAL lCursorLoc AS CursorLocationEnum)
   CADODB_DP("CAdoRecordset.CursorLocation")
   IF m_pRecordset THEN SetResult(m_pRecordset->put_CursorLocation(lCursorLoc))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Sets or returns a Long value that must be greater than 0. Default is 1.
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.LockType () AS LockTypeEnum
   CADODB_DP("CAdoRecordset.LockType - AS LockTypeEnum")
   DIM lLockType AS LockTypeEnum
   IF m_pRecordset THEN SetResult(m_pRecordset->get_LockType(@lLockType))
   PROPERTY = lLockType
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.LockType (BYVAL lLockType AS LockTypeEnum)
   CADODB_DP("CAdoRecordset.LockType")
   IF m_pRecordset THEN SetResult(m_pRecordset->put_LockType(lLockType))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the maximum number of records to return to a Recordset from a query
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.MaxRecords () AS LONG
   CADODB_DP("CAdoRecordset.MaxRecords - AS LONG")
   DIM lMaxRecords AS LONG
   IF m_pRecordset THEN SetResult(m_pRecordset->get_MaxRecords(@lMaxRecords))
   PROPERTY = lMaxRecords
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.MaxRecords (BYVAL lMaxRecords AS LONG)
   CADODB_DP("CAdoRecordset.MaxRecords")
   IF m_pRecordset THEN SetResult(m_pRecordset->put_MaxRecords(lMaxRecords))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates how many records constitute one page in the Recordset
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.PageSize () AS LONG
   CADODB_DP("CAdoRecordset.PageSize - AS LONG")
   DIM nPageSize AS LONG
   IF m_pRecordset THEN SetResult(m_pRecordset->get_PageSize(@nPageSize))
   PROPERTY = nPageSize
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.PageSize (BYVAL nPageSize AS LONG)
   CADODB_DP("CAdoRecordset.PageSize")
   IF m_pRecordset THEN SetResult(m_pRecordset->put_PageSize(nPageSize))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates how many pages of data the Recordset object contains
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.PageCount () AS LONG
   CADODB_DP("CAdoRecordset.PageCount")
   DIM nPageCount AS LONG
   IF m_pRecordset THEN SetResult(m_pRecordset->get_PageSize(@nPageCount))
   PROPERTY = nPageCount
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates how many pages of data the Recordset object contains
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.RecordCount () AS LONG
   CADODB_DP("CAdoRecordset.RecordCount")
   DIM pl AS LONG
   IF m_pRecordset THEN SetResult(m_pRecordset->get_RecordCount(@pl))
   PROPERTY = pl
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates which records are to be marshaled back to the server
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.MarshalOptions () AS MarshalOptionsEnum
   CADODB_DP("CAdoRecordset.MarshalOptions - AS MarshalOptionsEnum")
   DIM eMarshal AS MarshalOptionsEnum
   IF m_pRecordset THEN SetResult(m_pRecordset->get_MarshalOptions(@eMarshal))
   PROPERTY = eMarshal
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.MarshalOptions (BYVAL eMarshal AS MarshalOptionsEnum)
   CADODB_DP("CAdoRecordset.MarshalOptions")
   IF m_pRecordset THEN SetResult(m_pRecordset->put_MarshalOptions(eMarshal))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' * Sets or returns a Variant value that indicates the value of the object
' The ADO Recorset object exposes a hidden member: the Collect property. This property is
' functionally similar to the Field's Value property, but it doesn't need a reference (explicit
' or implicit) to the Field object. You can pass either a numeric index or a field's name to
' this property.
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.Collect (BYREF cvIndex AS CVAR) AS CVAR
   CADODB_DP("CAdoRecordset.Collect - CVAR")
   DIM cvValue AS CVAR
   IF m_pRecordset THEN SetResult(m_pRecordset->get_Collect(cvIndex, cvValue.vptr))
   PROPERTY = cvValue
END PROPERTY
' ===========================================================================================

' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.Collect (BYREF cvIndex AS CVAR, BYREF cvValue AS CVAR)
   CADODB_DP("CAdoRecordset.Collect - Index, Value")
   IF m_pRecordset THEN SetResult(m_pRecordset->put_Collect(cvIndex, cvValue))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the name of the index currently in effect for a Recordset object
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.Index () AS CBSTR
   CADODB_DP("CAdoRecordset.Index - AS CBSTR")
   DIM bstrIndex AS AFX_BSTR
   IF m_pRecordset THEN SetResult(m_pRecordset->get_Index(@bstrIndex))
   PROPERTY = bstrIndex
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.Index (BYREF cbsIndex AS CBSTR)
   CADODB_DP("CAdoRecordset.Index")
   IF m_pRecordset THEN SetResult(m_pRecordset->put_Index(cbsIndex.m_bstr))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Returns the Recordset as a string
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.GetString (BYVAL StringFormat AS StringFormatEnum = adClipString, BYVAL NumRows AS LONG = adReadAll, _
BYREF cbsColumnDelimeter AS CBSTR = "", BYREF cbsRowDelimeter AS CBSTR = "", BYREF cbsNullExpr AS CBSTR = "") AS CBSTR
   CADODB_DP("CAdoRecordset.GetString")
   DIM bstrOutput AS AFX_BSTR
   IF m_pRecordset THEN SetResult(m_pRecordset->GetString(StringFormat, NumRows, cbsColumnDelimeter, cbsRowDelimeter, cbsNullExpr, @bstrOutput))
   RETURN bstrOutput
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Indicates an object that contains data to be represented as a Recordset object
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.DataSource () AS Afx_IUnknown PTR
   CADODB_DP("CAdoRecordset.DataSource - AS Afx_IUnknown PTR")
   DIM pDataSource AS Afx_IUnknown PTR
   IF m_pRecordset THEN SetResult(m_pRecordset->get_DataSource(@pDataSource))
   PROPERTY = pDataSource
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.DataSource (BYVAL punkDataSource AS Afx_IUnknown PTR)
   CADODB_DP("CAdoRecordset.DataSource - IUnknown PTR")
   IF m_pRecordset THEN SetResult(m_pRecordset->putref_DataSource(punkDataSource))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the name of the data member that will be retrieved from the object referenced by
' the DataSource property.
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.DataMember () AS CBSTR
   CADODB_DP("CAdoRecordset.DataMember - AS CBSTR")
   DIM cbsDataMember AS CBSTR
   IF m_pRecordset THEN SetResult(m_pRecordset->get_DataMember(cbsDataMember.vptr))
   PROPERTY = cbsDataMember
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.DataMember (BYREF cbsDataMember AS CBSTR)
   CADODB_DP("CAdoRecordset.DataMember - CBSTR")
   IF m_pRecordset THEN SetResult(m_pRecordset->put_DataMember(cbsDataMember.m_bstr))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Sets or returns a value that indicates the current record position within the recordset
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.AbsolutePosition () AS PositionEnum
   CADODB_DP("CAdoRecordset.AbsolutePosition - AS PositionEnum")
   DIM Position AS PositionEnum
   IF m_pRecordset THEN SetResult(m_pRecordset->get_AbsolutePosition(@Position))
   PROPERTY = Position
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.AbsolutePosition (BYVAL Position AS PositionEnum)
   CADODB_DP("CAdoRecordset.AbsolutePosition - PositionEnum")
   IF m_pRecordset THEN SetResult(m_pRecordset->put_AbsolutePosition(Position))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Sets or returns a value that indicates the current page position within the recordset
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.AbsolutePage () AS PositionEnum
   CADODB_DP("CAdoRecordset.AbsolutePage - AS PositionEnum")
   DIM nPage AS PositionEnum
   IF m_pRecordset THEN SetResult(m_pRecordset->get_AbsolutePage(@nPage))
   PROPERTY = nPage
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.AbsolutePage (BYVAL nPage AS PositionEnum)
   CADODB_DP("CAdoRecordset.AbsolutePage - PositionEnum")
   IF m_pRecordset THEN SetResult(m_pRecordset->put_AbsolutePage(nPage))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Moves the position of the current record in a Recordset object.
' - NumRecords: A signed Long expression that specifies the number of records that the current
'   record position moves.
' - start: Optional. A BookmarkEnum value.
'   adBookmarkCurrent (0) Starts at the current record.
'   adBookmarkFirst (1) Starts at the first record.
'   adBookmarkLast (2) Starts at the last record.
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.Move (BYVAL NumRecords AS LONG, BYVAL start AS BookmarkEnum = adBookmarkCurrent) AS HRESULT
   CADODB_DP("CAdoRecordset.Move - BookmarkEnum")
   DIM vStart AS VARIANT : V_VT(@vStart) = VT_I4 : V_I4(@vStart) = start
   IF m_pRecordset THEN RETURN(SetResult(m_pRecordset->Move(NumRecords, vStart)))
END FUNCTION
' ===========================================================================================
' ===========================================================================================
' Moves the position of the current record in a Recordset object.
' - NumRecords: A signed Long expression that specifies the number of records that the current
'   record position moves.
' - cvStart: Optional. A Variant value that evaluates to a bookmark.
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.Move (BYVAL NumRecords AS LONG, BYREF cvStart AS CVAR) AS HRESULT
   CADODB_DP("CAdoRecordset.Move - CVAR")
   IF m_pRecordset THEN RETURN(SetResult(m_pRecordset->Move(NumRecords, cvStart)))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' * Moves the current record position to the next record in the Recordset
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.MoveNext () AS HRESULT
   CADODB_DP("CAdoRecordset.MoveNext")
   IF m_pRecordset THEN RETURN(SetResult(m_pRecordset->MoveNext))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Moves the current record position to the previous record in the Recordset
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.MovePrevious () AS HRESULT
   CADODB_DP("CAdoRecordset.MovePrevious")
   IF m_pRecordset THEN RETURN(SetResult(m_pRecordset->MovePrevious))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Moves the current record position to the first record in the Recordset.
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.MoveFirst () AS HRESULT
   CADODB_DP("CAdoRecordset.MoveFirst")
   IF m_pRecordset THEN RETURN(SetResult(m_pRecordset->MoveFirst))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Moves the current record position to the last record in the Recordset.
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.MoveLast () AS HRESULT
   CADODB_DP("CAdoRecordset.MoveLast")
   IF m_pRecordset THEN RETURN(SetResult(m_pRecordset->MoveLast))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Indicates the status of the current record with respect to batch updates or other bulk operations
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.Status () AS LONG
   CADODB_DP("CAdoRecordset.Status")
   DIM lStatus AS LONG
   IF m_pRecordset THEN SetResult(m_pRecordset->get_Status(@lStatus))
   PROPERTY = lStatus
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates for whether the state of the Recordset object is open or closed
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.State () AS LONG
   CADODB_DP("CAdoRecordset.State")
   DIM lState AS LONG
   IF m_pRecordset THEN SetResult(m_pRecordset->get_State(@lState))
   PROPERTY = lState
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the editing status of the current record
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.EditMode () AS EditModeEnum
   CADODB_DP("CAdoRecordset.EditMode")
   DIM nMode AS EditModeEnum
   IF m_pRecordset THEN SetResult(m_pRecordset->get_EditMode(@nMode))
   PROPERTY = nMode
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates, in a hierarchical Recordset object, whether the reference to the underlying child
' records (that is, the chapter) changes when the parent row position changes.
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.StayInSync () AS BOOLEAN
   CADODB_DP("CAdoRecordset.StayInSync")
   DIM bStayInSync AS VARIANT_BOOL
   IF m_pRecordset THEN SetResult(m_pRecordset->get_StayInSync(@bStayInSync))
   PROPERTY = bStayInSync
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.StayInSync (BYVAL bStayInSync AS BOOLEAN)
   CADODB_DP("CAdoRecordset.StayInSync")
   IF m_pRecordset THEN SetResult(m_pRecordset->put_StayInSync(bStayInSync))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Refreshes the data in the current Recordset object from the underlying database
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.Resync (BYVAL AffectRecords AS AffectEnum = adAffectAll, BYVAL ResyncValues AS ResyncEnum = adResyncAllValues) AS HRESULT
   CADODB_DP("CAdoRecordset.Resync")
   IF m_pRecordset THEN RETURN(SetResult(m_pRecordset->Resync(AffectRecords, ResyncValues)))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Updates the data in a Recordset object by re-executing the query on which the object is based
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.Requery (BYVAL Options AS LONG = adOpenUnspecified) AS HRESULT
   CADODB_DP("CAdoRecordset.Requery")
   IF m_pRecordset THEN RETURN(SetResult(m_pRecordset->Requery(Options)))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Moves to the previous record in a specified Recordset object and makes that record the current record
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.NextRecordset (BYVAL RecordsAffected AS LONG PTR = NULL) AS Afx_AdoRecordset PTR
   CADODB_DP("CAdoRecordset.NextRecordset")
   DIM vRecordsAffected AS VARIANT, pRecordset AS Afx_AdoRecordset PTR
   IF m_pRecordset THEN SetResult(m_pRecordset->NextRecordset(@vRecordsAffected, @pRecordset))
   IF RecordsAffected THEN *RecordsAffected = vRecordsAffected.lVal
   RETURN pRecordset
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Determines whether a specified Recordset object supports a particular type of functionality
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.Supports (BYVAL CursorOptions AS CursorOptionEnum) AS BOOLEAN
   CADODB_DP("CAdoRecordset.Supports")
   DIM pb AS VARIANT_BOOL
   IF m_pRecordset THEN SetResult(m_pRecordset->Supports(CursorOptions, @pb))
   RETURN pb
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Writes all pending batch updates to disk
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.UpdateBatch (BYVAL AffectRecords AS AffectEnum = adAffectAll) AS HRESULT
   CADODB_DP("CAdoRecordset.UpdateBatch")
   IF m_pRecordset THEN RETURN(SetResult(m_pRecordset->UpdateBatch))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Cancels a pending batch update
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.CancelBatch (BYVAL AffectRecords AS AffectEnum = adAffectAll) AS HRESULT
   CADODB_DP("CAdoRecordset.CancelBatch")
   IF m_pRecordset THEN RETURN(SetResult(m_pRecordset->CancelBatch))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Cancels any changes made to the current or new row of a Recordset object before calling the Update method
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.CancelUpdate () AS HRESULT
   CADODB_DP("CAdoRecordset.CancelUpdate")
   IF m_pRecordset THEN RETURN(SetResult(m_pRecordset->CancelUpdate))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Closes a Recordset object and any dependent objects
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.Close () AS HRESULT
   CADODB_DP("CAdoRecordset.Close")
   IF m_pRecordset THEN RETURN(SetResult(m_pRecordset->Close))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Closes a Recordset object and any dependent objects
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.Clone (BYVAL nLockType AS LockTypeEnum = adLockUnspecified) AS Afx_AdoRecordset PTR
   CADODB_DP("CAdoRecordset.Clone")
   DIM ppvObject AS Afx_AdoRecordset PTR
   IF m_pRecordset THEN SetResult(m_pRecordset->Clone(nLockType, @ppvObject))
   RETURN ppvObject
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Cancels execution of a pending, asynchronous method call
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.Cancel () AS HRESULT
   CADODB_DP("CAdoRecordset.Cancel")
   IF m_pRecordset THEN RETURN(SetResult(m_pRecordset->Cancel))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Indicates a filter for data in a Recordset
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.Filter () AS CVAR
   CADODB_DP("CAdoRecordset.Filter")
   DIM cvCriteria AS CVAR
   IF m_pRecordset THEN SetResult(m_pRecordset->get_Filter(cvCriteria.vptr))
   PROPERTY = cvCriteria
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.Filter (BYVAL cvCriteria AS CVAR)
   CADODB_DP("CAdoRecordset.Filter")
   IF m_pRecordset THEN SetResult(m_pRecordset->put_Filter(cvCriteria))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates one or more field names on which the Recordset is sorted, and whether each field
' is sorted in ascending or descending order
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.Sort () AS CBSTR
   CADODB_DP("CAdoRecordset.Sort")
   DIM bstrCriteria AS AFX_BSTR
   IF m_pRecordset THEN SetResult(m_pRecordset->get_Sort(@bstrCriteria))
   PROPERTY = bstrCriteria
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.Sort (BYREF cbsCriteria AS CBSTR)
   CADODB_DP("CAdoRecordset.Sort")
   IF m_pRecordset THEN SetResult(m_pRecordset->put_Sort(cbsCriteria.m_bstr))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Deletes the current record or a group of records.
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.Delete_ (BYVAL AffectRecords AS AffectEnum = adAffectCurrent) AS HRESULT
   CADODB_DP("CAdoRecordset.Delete_")
   IF m_pRecordset THEN RETURN(SetResult(m_pRecordset->Delete_(AffectRecords)))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Saves the Recordset in a file
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.Save (BYREF cvDestination AS CVAR, BYVAL PersistFormat AS PersistFormatEnum = adPersistADTG) AS HRESULT
   CADODB_DP("CAdoRecordset.Save - CBSTR")
   IF m_pRecordset THEN RETURN(SetResult(m_pRecordset->Save(cvDestination, PersistFormat)))
END FUNCTION
' ===========================================================================================
' ===========================================================================================
' Saves the Recordset in a Stream object
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.Save (BYVAL pDestination AS Afx_AdoStream PTR, BYVAL PersistFormat AS PersistFormatEnum = adPersistADTG) AS HRESULT
   CADODB_DP("CAdoRecordset.Save - Afx_AdoStream")
   DIM vDestination AS VARIANT : V_VT(@vDestination) = VT_BSTR : V_DISPATCH(@vDestination) = cast(IDispatch PTR, cast(ULONG_PTR, pDestination))
   IF m_pRecordset THEN RETURN(SetResult(m_pRecordset->Save(vDestination, PersistFormat)))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Compares two bookmarks and returns an indication of their relative values.
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.CompareBookmarks (BYREF cvBookmark1 AS CVAR, BYREF cvBookmark2 AS CVAR) AS CompareEnum
   CADODB_DP("CAdoRecordset.CompareBookmarks")
   DIM pCompare AS CompareEnum
   IF m_pRecordset THEN SetResult(m_pRecordset->CompareBookmarks(cvBookmark1, cvBookmark2, @pCompare))
   RETURN pCompare
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' * Searches a Recordset for the row that satisfies the specified criteria
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.Find (BYREF cbsCriteria AS CBSTR, BYREF cvStart AS CVAR = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND), _
BYVAL SkipRecords AS LONG = 0, BYVAL SearchDirection AS SearchDirectionEnum = adSearchForward) AS HRESULT
   CADODB_DP("CAdoRecordset.Find")
   IF m_pRecordset THEN RETURN(SetResult(m_pRecordset->Find(cbsCriteria.m_bstr, SkipRecords, SearchDirection, cvStart)))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' * Retrieves multiple records of a Recordset object into an array
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.GetRows (BYVAL Rows AS LONG = adGetRowsRest, BYREF cvStart AS CVAR = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND), BYREF cvFields AS CVAR = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND)) AS SAFEARRAY PTR
   CADODB_DP("CAdoRecordset.GetRows")
   DIM vRows AS VARIANT
   IF m_pRecordset THEN SetResult(m_pRecordset->GetRows(Rows, cvStart, cvFields, @vRows))
   ' // Don't clear the variant because we are transfering ownership of the safe array
   IF m_Result = S_OK THEN RETURN vRows.parray
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' * Searches the index of a Recordset to quickly locate the row that matches the specified values, and changes the current row position to that row
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.Seek (BYREF cvKeyValues AS CVAR, BYVAL SeekOption AS SeekEnum = adSeekFirstEQ) AS HRESULT
   CADODB_DP("CAdoRecordset.Seek")
   IF m_pRecordset THEN RETURN(SetResult(m_pRecordset->Seek(cvKeyValues, SeekOption)))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Creates a new record for an updatable Recordset object.
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.AddNew (BYREF cvFieldList AS CVAR = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND), BYREF cvValues AS CVAR = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND)) AS HRESULT
   CADODB_DP("CAdoRecordset.AddNew")
   IF m_pRecordset THEN FUNCTION = SetResult(m_pRecordset->AddNew(cvFieldList, cvValues))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Saves any changes you make to the current row of a Recordset object
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.Update (BYREF cvFieldList AS CVAR = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND), BYREF cvValues AS CVAR = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND)) AS HRESULT
   CADODB_DP("CAdoRecordset.Update")
   IF m_pRecordset THEN FUNCTION = SetResult(m_pRecordset->AddNew(cvFieldList, cvValues))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Retrieves information from the ADO Errors collection
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.GetErrorInfo (BYVAL nError AS HRESULT = 0) AS CBSTR
   CADODB_DP("CAdoRecordset.GetErrorInfo")
   DIM vConnection AS VARIANT
   IF m_pRecordset THEN m_pRecordset->get_ActiveConnection(@vConnection)
   DIM cbsError AS CBSTR
   IF vConnection.vt = VT_DISPATCH THEN
      cbsError = AfxAdoGetErrorInfo(cast(Afx_ADOConnection PTR, cast(ULONG_PTR, vConnection.pdispVal)), nError)
   END IF
   VariantClear(@vConnection)
   RETURN cbsError
END FUNCTION
' ===========================================================================================

END NAMESPACE
