' ########################################################################################
' Microsoft Windows
' Contents: ADO classes
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "Afx/AfxAdo.bi"
#include once "Afx/AfxCOM.inc"
#include once "Afx/CWSTR.inc"
#include once "Afx/CVAR.inc"

' ########################################################################################
' CAdoBase class
' ########################################################################################
TYPE CAdoBase

   DIM m_Result AS HRESULT
   DIM m_bUninitCOM AS BOOLEAN

   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR
   DECLARE FUNCTION GetLastResult () AS HRESULT
   DECLARE FUNCTION SetResult (BYVAL Result AS HRESULT) AS HRESULT

END TYPE
' ########################################################################################

' ========================================================================================
' Default constructor
' ========================================================================================
CONSTRUCTOR CAdoBase
   CADODB_DP("*** CAdoBase CONSTRUCTOR - m_bUninitCOM = " & STR(m_bUninitCOM))
   ' // Initialize the COM library
   IF m_bUninitCOM = FALSE THEN
      DIM hr AS HRESULT = CoInitialize(NULL)
      IF hr = S_OK OR hr = S_FALSE THEN m_bUninitCOM = TRUE
   END IF
END CONSTRUCTOR
' ========================================================================================

' ===========================================================================================
' Destructor
' ===========================================================================================
DESTRUCTOR CAdoBase
   CADODB_DP("*** CAdoBase DESTRUCTOR")
   ' // Free the COM library
   IF m_bUninitCOM THEN CoUninitialize
END DESTRUCTOR
' ===========================================================================================

' ========================================================================================
' Returns the last result
' ========================================================================================
PRIVATE FUNCTION CAdoBase.GetLastResult () AS HRESULT
   CADODB_DP("CAdoBase.GetLastResult")
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the last status code.
' ========================================================================================
PRIVATE FUNCTION CAdoBase.SetResult (BYVAL Result AS HRESULT) AS HRESULT
   CADODB_DP("CAdoBase.SetResult")
   m_Result = Result
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ===========================================================================================
' Returns the description of an ADO runtime error
' ===========================================================================================
PRIVATE FUNCTION AfxAdoGetRuntimeErrorInfo (BYVAL wError AS WORD) AS CBSTR

   DIM s AS STRING

   SELECT CASE wError
      CASE 3000 ' &HBB8 - %adErrProviderFailed
         s = "Provider failed to perform the requested operation."
      CASE 3001 ' &HBB9 - %adErrInvalidArgument
         s = "Arguments are of the wrong type, are out of acceptable range, or are in conflict with one another."
      CASE 3002 ' &HBBA - %adErrOpeningFile
         s = "File could not be opened."
      CASE 3003 ' &HBBB - %adErrReadFile
         s = "File could not be read."
      CASE 3004 ' &HBBC - %adErrWriteFile
         s = "Write to file failed."
      CASE 3021 ' &HBCD - %adErrNoCurrentRecord
         s = "Either BOF or EOF is True, or the current record has been deleted. Requested operation requires a current record."
      CASE 3219 ' &HC93 - %adErrIllegalOperation
         s = "Operation is not allowed in this context."
      CASE 3220 ' &HC94 - %adErrCantChangeProvider
         s = "Supplied provider is different from the one already in use."
      CASE 3246 ' &HCAE - %adErrInTransaction
         s = "Connection object cannot be explicitly closed while in a transaction."
      CASE 3251 ' &HCB3 - %adErrFeatureNotAvailable
         s = "Object or provider is not capable of performing requested operation."
      CASE 3265 ' &HCC1 - %adErrItemNotFound
         s = "Item cannot be found in the collection corresponding to the requested name or ordinal."
      CASE 3367 ' &HD27 - %adErrObjectInCollection
         s = "Object is already in collection. Cannot append."
      CASE 3420 ' &HD5C - %adErrObjectNotSet
         s = "Object is no longer valid."
      CASE 3421 ' &HD5D - %adErrDataConversion
         s = "Application uses a value of the wrong type for the current operation."
      CASE 3704 ' &HE78 - %adErrObjectClosed
         s = "Operation is not allowed when the object is closed. "
      CASE 3705 ' &HE79 - %adErrObjectOpen
         s = "Operation is not allowed when the object is open. "
      CASE 3706 ' &HE7A - %adErrProviderNotFound
         s = "Provider cannot be found. It may not be properly installed."
      CASE 3707 ' &HE7B - %adErrBoundToCommand
         s = "Cannot change the ActiveConnection property of a Recordset object which has a Command object as its source."
      CASE 3708 ' &HE7C - %adErrInvalidParamInfo
         s = "Parameter object is improperly defined. Inconsistent or incomplete information was provided."
      CASE 3709 ' &HE7D - %adErrInvalidConnection
         s = "The connection cannot be used to perform this operation. It is either closed or invalid in this context."
      CASE 3710 ' &HE7E - %adErrNotReentrant
         s = "Operation cannot be performed while processing event. "
      CASE 3711 ' &HE7F - %adErrStillExecuting
         s = "Operation cannot be performed while executing asynchronously."
      CASE 3712 ' &HE80 - %adErrOperationCancelled
         s = "Operation has been cancelled by the user."
      CASE 3713 ' &HE81 - %adErrStillConnecting
         s = "Operation cannot be performed while connecting aynchronously. "
      CASE 3714 ' &HE82 - %adErrInvalidTransaction
         s = "Coordinating transaction is invalid or has not started."
      CASE 3715 ' &HE83 - %adErrNotExecuting
         s = "Operation cannot be performed while not executing."
      CASE 3716 ' &HE84 - %adErrUnsafeOperation
         s = "Safety settings on this computer prohibit accessing a data source on another domain."
'      CASE 3717 ' &HE85 - %adwrnSecurityDialog
'         For internal use only. Don't use.
'      CASE 3718 ' &HE86 - %adwrnSecurityDialogHeader
'         For internal use only. Don't use.
      CASE 3719 ' &HE87 - %adErrIntegrityViolation
         s = "Data value conflicts with the integrity constraints of the field."
      CASE 3720 ' &HE88 - %adErrPermissionDenied
         s = "Insufficent permission prevents writing to the field."
      CASE 3721 ' &HE89 - %adErrDataOverflow
         s = "Data value is too large to be represented by the field data type."
      CASE 3722 ' &HE8A - %adErrSchemaViolation
         s = "Data value conflicts with the data type or constraints of the field."
      CASE 3723 ' &HE8B - %adErrSignMismatch
         s = "Conversion failed because the data value was signed and the field data type used by the provider was unsigned."
      CASE 3724 ' &HE8C - %adErrCantConvertvalue
         s = "Data value cannot be converted for reasons other than sign mismatch or data overflow. For example, conversion would have truncated data."
      CASE 3725 ' &HE8D - %adErrCantCreate
         s = "Data value cannot be set or retrieved because the field data type was unknown, or the provider had insufficient resources to perform the operation."
      CASE 3726 ' &HE8E - %adErrColumnNotOnThisRow
         s = "Record does not contain this field."
      CASE 3727 ' &HE8F - %adErrURLDoesNotExist
         s = "Either the source URL or the parent of the destination URL does not exist."
      CASE 3728 ' &HE90 - %adErrTreePermissionDenied
         s = "Permissions are insufficient to access tree or subtree. "
      CASE 3729 ' &HE91 - %adErrInvalidURL
         s = "URL contains invalid characters. Make sure the URL is typed correctly."
      CASE 3730 ' &HE92 - %adErrResourceLocked
         s = "Object represented by the specified URL is locked by one or more other processes. Wait until the process has finished and attempt the operation again."
      CASE 3731 ' &HE93 - %adErrResourceExists
         s = "Copy operation cannot be performed. Object named by destination URL already exists. Specify adCopyOverwrite to replace the object."
      CASE 3732 ' &HE94 - %adErrCannotComplete
         s = "Server cannot complete the operation."
      CASE 3733 ' &HE95 - %adErrVolumeNotFound
         s = "Provider cannot locate the storage device indicated by the URL. Make sure the URL is typed correctly."
      CASE 3734 ' &HE96 - %adErrOutOfSpace
         s = "Operation cannot be performed. Provider cannot obtain enough storage space."
      CASE 3735 ' &HE97 - %adErrResourceOutOfScope
         s = "Source or destination URL is outside the scope of the current record."
      CASE 3736 ' &HE98 - %adErrUnavailable
         s = "Operation failed to complete and the status is unavailable. The field may be unavailable or the operation was not attempted."
      CASE 3737 ' &HE99 - %adErrURLNamedRowDoesNotExist
         s = "Record named by this URL does not exist."
      CASE 3738 ' &HE9A - %adErrDelResOutOfScope
         s = "URL of the object to be deleted is outside the scope of the current record."
      CASE 3739 ' &HE9B - %adErrPropInvalidColumn
         s = "Cannot apply property to field"
      CASE 3740 ' &HE9C - %adErrPropInvalidOption
         s = "Attribute property invalid"
      CASE 3741 ' &HE9D - %adErrPropInvalidValue
         s = "Invalid property value"
      CASE 3742 ' &HE9E - %adErrPropConflicting
         s = "Property values conflict with each other"
      CASE 3743 ' &HE9F - %adErrPropNotAllSettable
         s = "Cannot set property or read-only"
      CASE 3744 ' &HEA0 - %adErrPropNotSet
         s = "Optional property value not set"
      CASE 3745 ' &HEA1 - %adErrPropNotSettable
         s = "Read-only property cannot be set"
      CASE 3746 ' &HEA2 - %adErrPropNotSupported
         s = "Property not supported by provider"
      CASE 3747 ' &HEA3 - %adErrCatalogNotSet
         s = "Operation requires a valid ParentCatalog."
      CASE 3748 ' &HEA4 - %adErrCantChangeConnection
         s = "Connection was denied. New connection you requested has different characteristics than the one already in use."
      CASE 3749 ' &HEA5 - %adErrFieldsUpdateFailed
         s = "Fields update failed. For further information, examine the Status property of individual field objects."
      CASE 3750 ' &HEA6 - %adErrDenyNotSupported
         s = "Provider does not support sharing restrictions."
      CASE 3751 ' &HEA7 - %adErrDenyTypeNotSupported
         s = "Provider does not support the requested kind of sharing restriction."
      CASE 3753 ' 3753 - %adErrProviderNotSpecified
         s = "Provider not specified"
      CASE 3754 ' &HEAA - %adErrConnectionStringTooLong
         s = "Connection string too long"
      CASE ELSE
         s = "Unknown error"
   END SELECT

   FUNCTION = s

END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Retrieves information from the ADO Errors collection
' ===========================================================================================
PRIVATE FUNCTION AfxAdoGetErrorInfo (BYVAL pConnection AS AFX_ADOConnection PTR, BYVAL nError AS HRESULT = 0) AS CBSTR

   DIM hr             AS HRESULT             ' // Result code
   DIM pErrors        AS Afx_ADOErrors PTR   ' // Reference to the Errors collection
   DIM pError         AS Afx_ADOError PTR    ' // Reference to the Error object
   DIM wError         AS WORD                ' // ADO runtime error
   DIM i              AS LONG                ' // Loop counter
   DIM nCount         AS LONG                ' // Number of errors
   DIM eNumber        AS LONG                ' // Error number
   DIM eSource        AS CBSTR               ' // Source of the error
   DIM eDescription   AS CBSTR               ' // Description of the error
   DIM eHelpFile      AS CBSTR               ' // Help file
   DIM eHelpContext   AS LONG                ' // Help context
   DIM eSQLState      AS CBSTR               ' // SQlState
   DIM eNativeError   AS LONG                ' // Native error
   DIM cbsErrorText   AS CBSTR               ' // String to return the result
   DIM cbMsgLen       AS DWORD               ' // Length of the returned message
   DIM pBuffer        AS WSTRING PTR         ' // Pointer to the buffer
   DIM cbsDescription AS CBSTR               ' // Description of the ADO runtime error
   DIM cvIndex        AS CVAR                ' // Index

   IF pConnection = NULL THEN RETURN ""

   ' Get a reference to the errors collection
   pConnection->get_Errors(@pErrors)
   IF pErrors THEN
      ' Get the number of errors
      pErrors->get_Count(@nCount)
      ' The errors collection is zero based
      FOR i = 0 TO nCount - 1
         ' Get a reference to the error object
         DIM cvIndex AS CVAR = i
         pErrors->get_Item(cvIndex, @pError)
         IF pError THEN
            pError->get_Number(@eNumber)
            pError->get_Description(eDescription.vptr)
            pError->get_Source(eSource.vptr)
            pError->get_SQLState(eSqlState.vptr)
            pError->get_NativeError(@eNativeError)
            pError->get_HelpFile(eHelpFile.vptr)
            pError->get_HelpContext(@eHelpContext)
            cbsErrorText += "OLE DB provider error" & CHR(13, 10)
            cbsErrorText += "Error number: " & STR(eNumber) & " [&h" & HEX(eNumber, 8) & "]" & CHR(13, 10)
            IF LEN(eDescription) THEN cbsErrorText += "Description: " & eDescription & CHR(13, 10)
            IF LEN(eSource) THEN cbsErrorText += "Source: " & eSource & CHR(13, 10)
            IF LEN(eSqlState) THEN cbsErrorText += "SQLState: " & eSqlState & CHR(13, 10)
            IF eNativeError THEN cbsErrorText += "Native error: " & STR(eNativeError) & " [&H" & HEX(eNativeError, 8) & "]" & CHR(13, 10)
            IF LEN(eHelpFile) THEN cbsErrorText += "Help file: " & eHelpFile & CHR(13, 10)
            IF eHelpContext THEN cbsErrorText += "Help context: " & STR(eHelpContext)
            ' Release the error object
            pError->Release
         END IF
      NEXT
      ' Clear the errors collection
      pErrors->Clear
      ' Release the errors collection
      pErrors->Release
      IF nCount THEN RETURN cbsErrorText
   END IF

   ' Must be an ADO runtime error
   IF nError = 0 THEN RETURN ""
   wError = nError
   cbsDescription = AfxAdoGetRuntimeErrorInfo(wError)
   IF LEN(cbsDescription) THEN
      cbsErrorText += "ADO runtime error" & CHR(13, 10)
      cbsErrorText += "Error code: " & STR(nError) & " [&h" & HEX$(nError, 8) & "]" & CHR(13, 10)
      cbsErrorText += "Description: " & cbsDescription
   ELSE
      ' If it is not a Provider or an ADO runtime error then must be a Windows error
      cbMsgLen = FormatMessageW(FORMAT_MESSAGE_ALLOCATE_BUFFER OR _
                 FORMAT_MESSAGE_FROM_SYSTEM OR FORMAT_MESSAGE_IGNORE_INSERTS, _
                 NULL, nError, BYVAL MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), _
                 cast(LPWSTR, @pBuffer), 0, NULL)
      IF cbMsgLen THEN
         cbsErrorText += "Windows error" & CHR(13, 10)
         cbsErrorText += "Error code:" & STR(nError) & " [&h" & HEX(nError, 8) & "]"
         cbsErrorText += "Description: " & *pBuffer
         LocalFree pBuffer
      END IF
   END IF

   RETURN cbsErrorText

END FUNCTION
' ===========================================================================================

#include once "Afx/CAdoDb/CAdoConnection.inc"
#include once "Afx/CAdoDb/CAdoRecordset.inc"
#include once "Afx/CAdoDb/CAdoCommand.inc"
#include once "Afx/CAdoDb/CAdoFields.inc"
#include once "Afx/CAdoDb/CAdoProperties.inc"
#include once "Afx/CAdoDb/CAdoParameters.inc"
#include once "Afx/CAdoDb/CAdoRecord.inc"
#include once "Afx/CAdoDb/CAdoStream.inc"
#include once "Afx/CAdoDb/CAdoErrors.inc"
