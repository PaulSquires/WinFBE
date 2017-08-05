' ########################################################################################
' Microsoft Windows
' Contents: ADO Fields
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once

NAMESPACE Afx

' ########################################################################################
' CAdoField class
' ########################################################################################
TYPE CAdoField EXTENDS CAdoBase

   m_Result AS HRESULT
   m_pField AS Afx_ADOField PTR

   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYVAL pField AS Afx_ADOField PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   DECLARE DESTRUCTOR
'   DECLARE OPERATOR @ () AS Afx_ADOField PTR PTR
   DECLARE FUNCTION vptr () AS Afx_ADOField PTR PTR

   ' // Attaches a reference to an Afx_ADOField object to the class
   DECLARE SUB Attach (BYVAL pField AS Afx_ADOField PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   ' // Retrieves the name of a field
   DECLARE PROPERTY Name () AS CBSTR
   ' // Sets or returns a Variant value that indicates the value of the object
   DECLARE PROPERTY Value () AS CVAR
   DECLARE PROPERTY Value (BYREF cvValue AS CVAR)
   ' // Returns a reference to the Properties collection.
   DECLARE PROPERTY Properties () AS Afx_ADOProperties PTR
   ' // Indicates the operational type or data type of a Field object.
   DECLARE PROPERTY Type_ () AS DataTypeEnum
   DECLARE PROPERTY Type_ (BYVAL nType AS DataTypeEnum)
   ' // Retrieves the data capacity of a field
   DECLARE PROPERTY ActualSize () AS LONG
   ' // Indicates the data capacity of a field.
   DECLARE PROPERTY DefinedSize () AS LONG
   DECLARE PROPERTY DefinedSize (BYVAL lSize AS LONG)
   ' // Indicates one or more characteristics of a field.
   DECLARE PROPERTY Attributes () AS LONG
   DECLARE PROPERTY Attributes (BYVAL lAttr AS LONG)
   ' // Indicates the degree of precision for numeric values for numeric fields.
   DECLARE PROPERTY Precision () AS BYTE
   DECLARE PROPERTY Precision (BYVAL bPrecision AS BYTE)
   ' // Indicates the scale of numeric values in a field.
   DECLARE PROPERTY NumericScale () AS BYTE
   DECLARE PROPERTY NumericScale (BYVAL bScale AS BYTE)
   ' // * Indicates the value of a field that existed in the record before any changes were made.
   DECLARE PROPERTY OriginalValue () AS CVAR
   ' // Indicates a field's current value in the database.
   DECLARE PROPERTY UnderlyingValue () AS CVAR
   ' // Appends data to a large text or binary data field.
   DECLARE FUNCTION AppendChunk (BYREF cvData AS CVAR) AS HRESULT
   ' // Returns all, or a portion, of the contents of a large text or binary data field.
   DECLARE FUNCTION GetChunk (BYVAL nLen AS LONG) AS CVAR
   ' // Links the current Field object to a data-bound control.
   DECLARE PROPERTY DataFormat () AS Afx_IUnknown PTR
   DECLARE PROPERTY DataFormat (BYVAL piDF AS Afx_IUnknown PTR)
   ' // * Indicates the status of a field. Returns a FieldStatusEnum value. The default value is adFieldOK.
   DECLARE PROPERTY Status () AS LONG

END TYPE
' ########################################################################################

' ========================================================================================
' Default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CAdoField
   CADODB_DP("CAdoField CONSTRUCTOR Default")
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CAdoField (BYVAL pField AS Afx_ADOField PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   CADODB_DP("-BEGIN CAdoField CONSTRUCTOR")
   IF pField THEN
      m_pField = pField
      IF fAddRef THEN m_pField->AddRef
   END IF
END CONSTRUCTOR
' ========================================================================================

' ===========================================================================================
' Cleanup
' ===========================================================================================
PRIVATE DESTRUCTOR CAdoField
   CADODB_DP("CAdoField DESTRUCTOR " & WSTR(m_pField))
   IF m_pField THEN m_pField->Release
END DESTRUCTOR
' ===========================================================================================

' ========================================================================================
' Returns the address of the connection pointer
' ========================================================================================
'PRIVATE OPERATOR CAdoField.@ () AS Afx_ADOField PTR PTR
'   OPERATOR = @m_pField
'END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CAdoField.vptr () AS Afx_ADOField PTR PTR
   CVAR_DP("CAdoField vptr")
   IF m_pField THEN m_pField->Release
   RETURN @m_pField
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the encapsulated IDispatch pointer
' ========================================================================================
PRIVATE OPERATOR * (BYREF cField AS CAdoField) AS Afx_ADOField PTR
   OPERATOR = cField.m_pField
END OPERATOR
' ========================================================================================

' ===========================================================================================
' Attaches a reference to an Afx_ADOField object to the class
' ===========================================================================================
PRIVATE SUB CAdoField.Attach (BYVAL pField AS Afx_ADOField PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   CADODB_DP("CAdoField.Attach")
   IF m_pField THEN m_pField->Release
   m_pField = pField
   ' // Increase the reference count
   IF fAddRef THEN
      IF m_pField THEN m_pField->AddRef
   END IF
END SUB
' ===========================================================================================

' ===========================================================================================
' Retrieves the name of a field
' ===========================================================================================
PRIVATE PROPERTY CAdoField.Name () AS CBSTR
   CADODB_DP("CAdoField.Name")
   IF m_pField = NULL THEN SetResult E_POINTER : RETURN ""
   DIM bstrName AS AFX_BSTR
   SetResult(m_pField->get_Name(@bstrName))
   RETURN bstrName
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the operational type or data type of a field.
' ===========================================================================================
PRIVATE PROPERTY CAdoField.Value () AS CVAR
   CADODB_DP("CAdoField.Value")
   DIM vValue AS VARIANT
   IF m_pField THEN SetResult(m_pField->get_Value(@vValue))
   RETURN vValue
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoField.Value (BYREF cvValue AS CVAR)
   CADODB_DP("CAdoField.Value")
   IF m_pField THEN SetResult(m_pField->put_Value(cvValue))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Returns a reference to the Properties collection.
' ===========================================================================================
PRIVATE PROPERTY CAdoField.Properties () AS Afx_ADOProperties PTR
   CADODB_DP("CAdoField.Properties")
   DIM pProperties AS Afx_ADOProperties PTR
   IF m_pField THEN SetResult(m_pField->get_Properties(@pProperties))
   RETURN pProperties
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the operational type or data type of a field.
' ===========================================================================================
PRIVATE PROPERTY CAdoField.Type_ () AS DataTypeEnum
   CADODB_DP("CAdoField.Type_")
   DIM nType AS DataTypeEnum
   IF m_pField THEN SetResult(m_pField->get_Type(@nType))
   RETURN nType
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the operational type or data type of a field.
' ===========================================================================================
PRIVATE PROPERTY CAdoField.Type_ (BYVAL nType AS DataTypeEnum)
   CADODB_DP("CAdoField.Type_")
   IF m_pField THEN SetResult(m_pField->put_Type(nType))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the actual length of a field's value.
' ===========================================================================================
PRIVATE PROPERTY CAdoField.ActualSize () AS LONG
   CADODB_DP("CAdoField.ActualSize")
   DIM lSize AS LONG
   IF m_pField THEN SetResult(m_pField->get_ActualSize(@lSize))
   RETURN lSize
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the data capacity of a field.
' ===========================================================================================
PRIVATE PROPERTY CAdoField.DefinedSize () AS LONG
   CADODB_DP("CAdoField.DefinedSize")
   DIM lSize AS LONG
   IF m_pField THEN SetResult(m_pField->get_DefinedSize(@lSize))
   RETURN lSize
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the data capacity of a field.
' ===========================================================================================
PRIVATE PROPERTY CAdoField.DefinedSize (BYVAL lSize AS LONG)
   CADODB_DP("CAdoField.DefinedSize")
   IF m_pField THEN SetResult(m_pField->put_DefinedSize(lSize))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the operational type or data type of a field.
' ===========================================================================================
PRIVATE PROPERTY CAdoField.Attributes () AS LONG
   CADODB_DP("CAdoField.Attributes")
   DIM lAttr AS LONG
   IF m_pField THEN SetResult(m_pField->get_Attributes(@lAttr))
   RETURN lAttr
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the operational type or data type of a field.
' ===========================================================================================
PRIVATE PROPERTY CAdoField.Attributes (BYVAL lAttr AS LONG)
   CADODB_DP("CAdoField.Attributes")
   IF m_pField THEN SetResult(m_pField->put_Attributes(lAttr))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the degree of precision for numeric values for numeric fields.
' ===========================================================================================
PRIVATE PROPERTY CAdoField.Precision () AS BYTE
   CADODB_DP("CAdoField.Precision")
   DIM bPrecision AS BYTE
   IF m_pField THEN SetResult(m_pField->get_Precision(@bPrecision))
   RETURN bPrecision
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the degree of precision for numeric values for numeric fields.
' ===========================================================================================
PRIVATE PROPERTY CAdoField.Precision (BYVAL bPrecision AS BYTE)
   CADODB_DP("CAdoField.Precision")
   IF m_pField THEN SetResult(m_pField->put_Precision(bPrecision))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the data capacity of a field.
' ===========================================================================================
PRIVATE PROPERTY CAdoField.NumericScale () AS BYTE
   CADODB_DP("CAdoField.NumericScale")
   DIM bScale AS BYTE
   IF m_pField THEN SetResult(m_pField->get_NumericScale(@bScale))
   RETURN bScale
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the data capacity of a field.
' ===========================================================================================
PRIVATE PROPERTY CAdoField.NumericScale (BYVAL bScale AS BYTE)
   CADODB_DP("CAdoField.NumericScale")
   IF m_pField THEN SetResult(m_pField->put_NumericScale(bScale))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the value of a field that existed in the record before any changes were made.
' ===========================================================================================
PRIVATE PROPERTY CAdoField.OriginalValue () AS CVAR
   CADODB_DP("CAdoField.OriginalValue")
   DIM vValue AS VARIANT
   IF m_pField THEN SetResult(m_pField->get_OriginalValue(@vValue))
   RETURN vValue
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates a field's current value in the database.
' ===========================================================================================
PRIVATE PROPERTY CAdoField.UnderlyingValue () AS CVAR
   CADODB_DP("CAdoField.UnderlyingValue")
   DIM vValue AS VARIANT
   IF m_pField THEN SetResult(m_pField->get_UnderlyingValue(@vValue))
   RETURN vValue
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Appends data to a large text or binary data Field.
' ===========================================================================================
PRIVATE FUNCTION CAdoField.AppendChunk (BYREF cvData AS CVAR) AS HRESULT
   CADODB_DP("CAdoField.AppendChunk")
   IF m_pField THEN SetResult(m_pField->AppendChunk(cvData))
   RETURN m_Result
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Returns all, or a portion, of the contents of a large text or binary data field.
' ===========================================================================================
PRIVATE FUNCTION CAdoField.GetChunk (BYVAL nLen AS LONG) AS CVAR
   CADODB_DP("CAdoField.GetChunk")
   DIM cvData AS CVAR
   IF m_pField THEN SetResult(m_pField->GetChunk(nLen, cvData.vptr))
   RETURN cvData
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Links the current Field object to a data-bound control
' ===========================================================================================
PRIVATE PROPERTY CAdoField.DataFormat () AS Afx_IUnknown PTR
   CADODB_DP("CAdoField.DataFormat")
   DIM piDF AS Afx_IUnknown PTR
   IF m_pField THEN SetResult(m_pField->get_DataFormat(@piDF))
   RETURN piDF
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoField.DataFormat (BYVAL piDF AS Afx_IUnknown PTR)
   CADODB_DP("CAdoField.DataFormat")
   IF m_pField THEN SetResult(m_pField->putref_DataFormat(piDF))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the status of a field. Returns a FieldStatusEnum value. The default value is adFieldOK.
' ===========================================================================================
PRIVATE PROPERTY CAdoField.Status () AS LONG
   CADODB_DP("CAdoField.Status")
   DIM lStatus AS LONG
   IF m_pField THEN SetResult(m_pField->get_Status(@lStatus))
   RETURN lStatus
END PROPERTY
' ===========================================================================================


' ########################################################################################
' CAdoFields class
' ########################################################################################
TYPE CAdoFields EXTENDS CAdoBase

   m_Result AS HRESULT
   m_pFields AS Afx_ADOFields PTR

   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYVAL pFields AS Afx_ADOFields PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   DECLARE DESTRUCTOR
'   DECLARE OPERATOR @ () AS Afx_ADOFields PTR PTR
   DECLARE FUNCTION vptr () AS Afx_ADOFields PTR PTR

   ' // Attaches a reference to an Afx_ADOFields object to the class
   DECLARE SUB Attach (BYVAL pFields AS Afx_ADOFields PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   ' // Retrieves the number of fields
   DECLARE PROPERTY Count () AS LONG
   ' // Indicates a specific member of the Fields collection, by name or ordinal number.
   DECLARE PROPERTY Item (BYREF cvIndex AS CVAR) AS Afx_ADOField PTR
   ' // Appends an object to a collection.
   DECLARE FUNCTION Append (BYREF cbsName AS CBSTR, BYVAL nType AS DataTypeEnum, BYVAL DefinedSize AS ADO_LONGPTR = 0, _
           BYVAL Attrib AS FieldAttributeEnum = 0, BYREF cvFieldValue AS CVAR = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND)) AS HRESULT
   ' // Deletes an object from the Fields collection.
   DECLARE FUNCTION Delete_ (BYREF cvIndex AS CVAR) AS HRESULT
   ' // Returns a reference to the collection's enumerator.
   DECLARE FUNCTION NewEnum () AS AFX_IUnknown PTR
   ' // Refreshes the contents of the Properties collection.
   DECLARE FUNCTION Refresh () AS HRESULT
   ' // Resynchronizes the contents of the Fields collection.
   DECLARE FUNCTION Resync (BYVAL ResyncValues AS ResyncEnum = adResyncAllValues) AS HRESULT
   ' // Saves any changes you make to the current Fields collection of a Record object.
   DECLARE FUNCTION Update () AS HRESULT
   ' // Cancels any changes made to the Fields collection of a Record object before calling the Update method.
   DECLARE FUNCTION CancelUpdate () AS HRESULT

END TYPE
' ########################################################################################

' ========================================================================================
' Default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CAdoFields
   CADODB_DP("BEGIN CAdoFields CONSTRUCTOR Default")
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CAdoFields (BYVAL pFields AS Afx_ADOFields PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   CADODB_DP("-BEGIN CAdoFields CONSTRUCTOR")
   IF pFields THEN
      m_pFields = pFields
      IF fAddRef THEN m_pFields->AddRef
   END IF
END CONSTRUCTOR
' ========================================================================================

' ===========================================================================================
' Cleanup
' ===========================================================================================
PRIVATE DESTRUCTOR CAdoFields
   CADODB_DP("CAdoFields DESTRUCTOR " & WSTR(m_pFields))
   IF m_pFields THEN m_pFields->Release
END DESTRUCTOR
' ===========================================================================================

' ========================================================================================
' Returns the address of the connection pointer
' ========================================================================================
'PRIVATE OPERATOR CAdoFields.@ () AS Afx_ADOFields PTR PTR
'   OPERATOR = @m_pFields
'END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CAdoFields.vptr () AS Afx_ADOFields PTR PTR
   CVAR_DP("CAdoFields vptr")
   IF m_pFields THEN m_pFields->Release
   RETURN @m_pFields
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the encapsulated IDispatch pointer
' ========================================================================================
PRIVATE OPERATOR * (BYREF cFields AS CAdoFields) AS Afx_ADOFields PTR
   OPERATOR = cFields.m_pFields
END OPERATOR
' ========================================================================================

' ===========================================================================================
' Attaches a reference to an Afx_ADOFields object to the class
' ===========================================================================================
PRIVATE SUB CAdoFields.Attach (BYVAL pFields AS Afx_ADOFields PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   CADODB_DP("CAdoFields.Attach")
   IF m_pFields THEN m_pFields->Release
   m_pFields = pFields
   IF fAddRef THEN
      IF m_pFields THEN m_pFields->AddRef
   END IF
END SUB
' ===========================================================================================

' ===========================================================================================
' Retrieves the number of fields
' ===========================================================================================
PRIVATE PROPERTY CAdoFields.Count () AS LONG
   CADODB_DP("CAdoFields.Count")
   DIM nCount AS LONG
   IF m_pFields THEN SetResult(m_pFields->get_Count(@nCount))
   RETURN nCount
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates a specific member of the Fields collection, by name or ordinal number.
' ===========================================================================================
PRIVATE PROPERTY CAdoFields.Item (BYREF cvIndex AS CVAR) AS Afx_ADOField PTR
   CADODB_DP("CAdoFields.Item")
   DIM pField AS Afx_ADOField PTR
   IF m_pFields THEN SetResult(m_pFields->get_Item(cvIndex, @pField))
   RETURN pField
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Appends an object to a collection.
' ===========================================================================================
PRIVATE FUNCTION CAdoFields.Append (BYREF cbsName AS CBSTR, BYVAL nType AS DataTypeEnum, BYVAL DefinedSize AS ADO_LONGPTR = 0, _
BYVAL Attrib AS FieldAttributeEnum = 0, BYREF cvFieldValue AS CVAR = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND)) AS HRESULT
   CADODB_DP("CAdoFields.Append")
   IF m_pFields THEN RETURN(SetResult(m_pFields->Append(cbsName.m_bstr, nType, DefinedSize, Attrib, cvFieldValue)))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Deletes an object from the Fields collection.
' ===========================================================================================
PRIVATE FUNCTION CAdoFields.Delete_ (BYREF cvIndex AS CVAR) AS HRESULT
   CADODB_DP("CAdoFields.Delete_")
   IF m_pFields THEN RETURN(SetResult(m_pFields->Delete_(cvIndex)))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Returns a reference to the collection's enumerator.
' ===========================================================================================
PRIVATE FUNCTION CAdoFields.NewEnum () AS Afx_IUnknown PTR
   CADODB_DP("CAdoFields.NewEnum")
   DIM pUnk AS Afx_IUnknown PTR
   IF m_pFields THEN SetResult(m_pFields->_NewEnum(@pUnk))
   RETURN pUnk
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Refreshes the contents of the Properties collection.
' ===========================================================================================
PRIVATE FUNCTION CAdoFields.Refresh () AS HRESULT
   CADODB_DP("CAdoFields.Refresh")
   IF m_pFields THEN RETURN(SetResult(m_pFields->Refresh))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Resynchronizes the contents of the Fields collection.
' ===========================================================================================
PRIVATE FUNCTION CAdoFields.Resync (BYVAL ResyncValues AS ResyncEnum = adResyncAllValues) AS HRESULT
   CADODB_DP("CAdoFields.Resync")
   IF m_pFields THEN RETURN(SetResult(m_pFields->Resync))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Saves any changes you make to the current Fields collection of a Record object.
' ===========================================================================================
PRIVATE FUNCTION CAdoFields.Update () AS HRESULT
   CADODB_DP("CAdoFields.Update")
   IF m_pFields THEN RETURN(SetResult(m_pFields->Update))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Cancels any changes made to the Fields collection of a Record object before calling the Update method.
' ===========================================================================================
PRIVATE FUNCTION CAdoFields.CancelUpdate () AS HRESULT
   CADODB_DP("CAdoFields.CancelUpdate")
   IF m_pFields THEN RETURN(SetResult(m_pFields->CancelUpdate))
END FUNCTION
' ===========================================================================================

END NAMESPACE
