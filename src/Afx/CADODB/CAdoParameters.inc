' ########################################################################################
' Microsoft Windows
' Contents: ADO Parameters
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once

NAMESPACE Afx

' ########################################################################################
' CAdoParameter class
' ########################################################################################
TYPE CAdoParameter EXTENDS CAdoBase

   m_Result AS HRESULT
   m_pParameter AS Afx_ADOParameter PTR

   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYVAL pParameter AS Afx_ADOParameter PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   DECLARE DESTRUCTOR
'   DECLARE OPERATOR @ () AS Afx_ADOParameter PTR PTR
   DECLARE FUNCTION vptr () AS Afx_ADOParameter PTR PTR

   ' // Attaches a reference to an Afx_ADOParameter object to the class
   DECLARE SUB Attach (BYVAL pParameter AS Afx_ADOParameter PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   ' // Sets or retrieves the name of a parameter
   DECLARE PROPERTY Name () AS CBSTR
   DECLARE PROPERTY Name (BYVAL cbsName AS CBSTR)
   ' // Sets or returns a Variant value that indicates the value of the object
   DECLARE PROPERTY Value () AS CVAR
   DECLARE PROPERTY Value (BYREF cvValue AS CVAR)
   ' // Indicates the operational type or data type of a Property object.
   DECLARE PROPERTY Type_ () AS DataTypeEnum
   ' // Indicates one or more characteristics of an object.
   DECLARE PROPERTY Attributes () AS LONG
   DECLARE PROPERTY Attributes (BYVAL lAttr AS LONG)
   ' // Appends data to a parameter object
   DECLARE FUNCTION AppendChunk (BYREF cvData AS CVAR) AS HRESULT
   ' // Indicates whether the Parameter represents an input parameter, an output parameter,
   ' // an input and an output parameter, or if the parameter is the Return value from a stored procedure.
   DECLARE PROPERTY Direction (BYVAL lParmDirection AS ParameterDirectionEnum)
   DECLARE PROPERTY Direction () AS ParameterDirectionEnum
   ' // Indicates the scale of numeric values in a parameter.
   DECLARE PROPERTY NumericScale () AS BYTE
   DECLARE PROPERTY NumericScale (BYVAL bScale AS BYTE)
   ' // Indicates the degree of precision for numeric values in a Parameter object.
   DECLARE PROPERTY Precision () AS BYTE
   DECLARE PROPERTY Precision (BYVAL bPrecision AS BYTE)
   ' // Returns a reference to the Properties collection.
   DECLARE PROPERTY Properties () AS Afx_ADOProperties PTR
   ' // Sets or returns a Long value that indicates the maximum size in bytes or characters of a value in a Parameter object.
   DECLARE PROPERTY Size () AS LONG
   DECLARE PROPERTY Size (BYVAL lSize AS LONG)

END TYPE
' ########################################################################################

' ========================================================================================
' Default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CAdoParameter
   CADODB_DP("CAdoParameter CONSTRUCTOR Default")
   ' // Create an instance of the Record object
   m_pParameter = AfxNewCom("ADODB.Parameter")
   CADODB_DP("-END CAdoParameter CONSTRUCTOR " & WSTR(m_pParameter))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CAdoParameter (BYVAL pParameter AS Afx_ADOParameter PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   CADODB_DP("CAdoParameter CONSTRUCTOR Default")
   m_pParameter = pParameter
   IF fAddRef THEN
      IF m_pParameter THEN m_pParameter->AddRef
   END IF
END CONSTRUCTOR
' ========================================================================================

' ===========================================================================================
' Cleanup
' ===========================================================================================
PRIVATE DESTRUCTOR CAdoParameter
   CADODB_DP("CAdoParameter DESTRUCTOR " & WSTR(m_pParameter))
   IF m_pParameter THEN m_pParameter->Release
END DESTRUCTOR
' ===========================================================================================

' ========================================================================================
' Returns the address of the connection pointer
' ========================================================================================
'PRIVATE OPERATOR CAdoParameter.@ () AS Afx_ADOParameter PTR PTR
'   OPERATOR = @m_pParameter
'END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CAdoParameter.vptr () AS Afx_ADOParameter PTR PTR
   CVAR_DP("CAdoParameter vptr")
   IF m_pParameter THEN m_pParameter->Release
   RETURN @m_pParameter
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the encapsulated IDispatch pointer
' ========================================================================================
PRIVATE OPERATOR * (BYREF cProperty AS CAdoParameter) AS Afx_ADOParameter PTR
   OPERATOR = cProperty.m_pParameter
END OPERATOR
' ========================================================================================

' ===========================================================================================
' Attaches a reference to an Afx_ADOCommand object to the class
' ===========================================================================================
PRIVATE SUB CAdoParameter.Attach (BYVAL pParameter AS Afx_ADOParameter PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   CADODB_DP("CAdoParameter.Attach")
   IF m_pParameter THEN m_pParameter->Release
   m_pParameter = pParameter
   ' // Increase the reference count
   IF fAddRef THEN
      IF m_pParameter THEN m_pParameter->AddRef
   END IF
END SUB
' ===========================================================================================

' ===========================================================================================
' Sets or tetrieves the name of a parameter
' ===========================================================================================
PRIVATE PROPERTY CAdoParameter.Name () AS CBSTR
   CADODB_DP("CAdoParameter.Name")
   DIM bstrName AS AFX_BSTR
   IF m_pParameter THEN SetResult(m_pParameter->get_Name(@bstrName))
   RETURN bstrName
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoParameter.Name (BYVAL cbsName AS CBSTR)
   CADODB_DP("CAdoParameter.Name")
   IF m_pParameter THEN SetResult(m_pParameter->put_Name(cbsName.m_bstr))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the operational type or data type of a parameter.
' ===========================================================================================
PRIVATE PROPERTY CAdoParameter.Value () AS CVAR
   CADODB_DP("CAdoParameter.Value")
   DIM cvValue AS CVAR
   IF m_pParameter THEN SetResult(m_pParameter->get_Value(cvValue.vptr))
   RETURN cvValue
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the operational type or data type of a parameter.
' ===========================================================================================
PRIVATE PROPERTY CAdoParameter.Value (BYREF cvValue AS CVAR)
   CADODB_DP("CAdoParameter.Value")
   IF m_pParameter THEN SetResult(m_pParameter->put_Value(cvValue))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the operational type or data type of a Parameter object.
' ===========================================================================================
PRIVATE PROPERTY CAdoParameter.Type_ () AS DataTypeEnum
   CADODB_DP("CAdoParameter.Type")
   DIM nType AS DataTypeEnum
   IF m_pParameter THEN SetResult(m_pParameter->get_Type(@nType))
   RETURN nType
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the operational type or data type of a parameter.
' ===========================================================================================
PRIVATE PROPERTY CAdoParameter.Attributes () AS LONG
   CADODB_DP("CAdoParameter.Attributes")
   DIM lAttr AS LONG
   IF m_pParameter THEN SetResult(m_pParameter->get_Attributes(@lAttr))
   RETURN lAttr
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the operational type or data type of a parameter.
' ===========================================================================================
PRIVATE PROPERTY CAdoParameter.Attributes (BYVAL lAttr AS LONG)
   CADODB_DP("CAdoParameter.Attributes")
   IF m_pParameter THEN SetResult(m_pParameter->put_Attributes(lAttr))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Appends data to a Parameter object.
' ===========================================================================================
PRIVATE FUNCTION CAdoParameter.AppendChunk (BYREF cvData AS CVAR) AS HRESULT
   CADODB_DP("CAdoParameter.AppendChunk")
   IF m_pParameter THEN SetResult(m_pParameter->AppendChunk(cvData))
   RETURN m_Result
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' // Indicates whether the Parameter represents an input parameter, an output parameter,
' // an input and an output parameter, or if the parameter is the Return value from a stored procedure.
' ===========================================================================================
PRIVATE PROPERTY CAdoParameter.Direction (BYVAL lParmDirection AS ParameterDirectionEnum)
   CADODB_DP("CAdoParameter.Direction")
   IF m_pParameter THEN SetResult(m_pParameter->put_Direction(lParmDirection))
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoParameter.Direction () AS ParameterDirectionEnum
   CADODB_DP("CAdoParameter.Direction")
   DIM lParmDirection AS ParameterDirectionEnum
   IF m_pParameter THEN SetResult(m_pParameter->get_Direction(@lParmDirection))
   RETURN lParmDirection
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the data capacity of a parameter.
' ===========================================================================================
PRIVATE PROPERTY CAdoParameter.NumericScale () AS BYTE
   CADODB_DP("CAdoParameter.NumericScale")
   DIM bScale AS BYTE
   IF m_pParameter THEN SetResult(m_pParameter->get_NumericScale(@bScale))
   RETURN bScale
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the data capacity of a parameter.
' ===========================================================================================
PRIVATE PROPERTY CAdoParameter.NumericScale (BYVAL bScale AS BYTE)
   CADODB_DP("CAdoParameter.NumericScale")
   IF m_pParameter THEN SetResult(m_pParameter->put_NumericScale(bScale))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the degree of precision for numeric values in a Parameter object.
' ===========================================================================================
PRIVATE PROPERTY CAdoParameter.Precision () AS BYTE
   CADODB_DP("CAdoParameter.Precision")
   DIM bPrecision AS BYTE
   IF m_pParameter THEN SetResult(m_pParameter->get_Precision(@bPrecision))
   RETURN bPrecision
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the degree of precision for numeric values for numeric parameters.
' ===========================================================================================
PRIVATE PROPERTY CAdoParameter.Precision (BYVAL bPrecision AS BYTE)
   CADODB_DP("CAdoParameter.Precision")
   IF m_pParameter THEN SetResult(m_pParameter->put_Precision(bPrecision))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Returns a reference to the Properties collection.
' ===========================================================================================
PRIVATE PROPERTY CAdoParameter.Properties () AS Afx_ADOProperties PTR
   CADODB_DP("CAdoParameter.Properties")
   DIM pProperties AS Afx_ADOProperties PTR
   IF m_pParameter THEN SetResult(m_pParameter->get_Properties(@pProperties))
   RETURN pProperties
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Sets or returns a Long value that indicates the maximum size in bytes or characters of a
' value in a Parameter object.
' ===========================================================================================
PRIVATE PROPERTY CAdoParameter.Size () AS LONG
   CADODB_DP("CAdoParameter.Size")
   DIM lSize AS LONG
   IF m_pParameter THEN SetResult(m_pParameter->get_Size(@lSize))
   RETURN lSize
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoParameter.Size (BYVAL lSize AS LONG)
   CADODB_DP("CAdoParameter.Size")
   IF m_pParameter THEN SetResult(m_pParameter->put_Size(lSize))
END PROPERTY
' ===========================================================================================


' ########################################################################################
' CAdoParameters class
' ########################################################################################
TYPE CAdoParameters EXTENDS CAdoBase

   m_Result AS HRESULT
   m_pParameters AS Afx_ADOParameters PTR

   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYVAL pParameters AS Afx_ADOParameters PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   DECLARE DESTRUCTOR
'   DECLARE OPERATOR @ () AS Afx_ADOParameters PTR PTR
   DECLARE FUNCTION vptr () AS Afx_ADOParameters PTR PTR

   ' // Attaches a reference to an Afx_ADOParameters object to the class
   DECLARE SUB Attach (BYVAL pParameters AS Afx_ADOParameters PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   ' // Retrieves the number of parameters
   DECLARE PROPERTY Count () AS LONG
   ' // Indicates a specific member of the Parameters collection, by name or ordinal number.
   DECLARE PROPERTY Item (BYREF cvIndex AS CVAR) AS Afx_ADOParameter PTR
   ' // Returns a reference to the collection's enumerator.
   DECLARE FUNCTION NewEnum () AS Afx_IUnknown PTR
   ' // Refreshes the contents of the Properties collection.
   DECLARE FUNCTION Refresh () AS HRESULT
   ' // Appends an object to the Parameters collection.
   DECLARE FUNCTION Append (BYVAL pObject AS Afx_IDispatch PTR) AS HRESULT
   DECLARE FUNCTION Append (BYREF pParam AS CAdoParameter) AS HRESULT
   ' // Deletes an object from the Parameters collection.
   DECLARE FUNCTION Delete_ (BYREF cvIndex AS CVAR) AS HRESULT

END TYPE
' ########################################################################################

' ========================================================================================
' Default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CAdoParameters
   CADODB_DP("BEGIN CAdoParameters CONSTRUCTOR Default")
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CAdoParameters (BYVAL pParameters AS Afx_ADOParameters PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   CADODB_DP("-BEGIN CAdoParameters CONSTRUCTOR")
   IF pParameters THEN
      m_pParameters = pParameters
      IF fAddRef THEN m_pParameters->AddRef
   END IF
END CONSTRUCTOR
' ========================================================================================

' ===========================================================================================
' Cleanup
' ===========================================================================================
PRIVATE DESTRUCTOR CAdoParameters
   CADODB_DP("CAdoParameters DESTRUCTOR " & WSTR(m_pParameters))
   IF m_pParameters THEN m_pParameters->Release
END DESTRUCTOR
' ===========================================================================================

' ========================================================================================
' Returns the address of the connection pointer
' ========================================================================================
'PRIVATE OPERATOR CAdoParameters.@ () AS Afx_ADOParameters PTR PTR
'   OPERATOR = @m_pParameters
'END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CAdoParameters.vptr () AS Afx_ADOParameters PTR PTR
   CVAR_DP("CAdoParameters vptr")
   IF m_pParameters THEN m_pParameters->Release
   RETURN @m_pParameters
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the encapsulated IDispatch pointer
' ========================================================================================
PRIVATE OPERATOR * (BYREF cParameters AS CAdoParameters) AS Afx_ADOParameters PTR
   OPERATOR = cParameters.m_pParameters
END OPERATOR
' ========================================================================================

' ===========================================================================================
' Attaches a reference to an Afx_ADOParameters object to the class
' ===========================================================================================
PRIVATE SUB CAdoParameters.Attach (BYVAL pParameters AS Afx_ADOParameters PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   CADODB_DP("CAdoParameters.Attach")
   IF m_pParameters THEN m_pParameters->Release
   m_pParameters = pParameters
   ' // Increase the reference count
   IF fAddRef THEN
      IF m_pParameters THEN m_pParameters->AddRef
   END IF
END SUB
' ===========================================================================================

' ===========================================================================================
' Retrieves the number of parameters
' ===========================================================================================
PRIVATE PROPERTY CAdoParameters.Count () AS LONG
   CADODB_DP("CAdoParameters.Count")
   DIM nCount AS LONG
   IF m_pParameters THEN SetResult(m_pParameters->get_Count(@nCount))
   RETURN nCount
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates a specific member of the Parameters collection, by name or ordinal number.
' ===========================================================================================
PRIVATE PROPERTY CAdoParameters.Item (BYREF cvIndex AS CVAR) AS Afx_ADOParameter PTR
   CADODB_DP("CAdoParameters.Item")
   DIM pParameter AS Afx_ADOParameter PTR
   IF m_pParameters THEN SetResult(m_pParameters->get_Item(cvIndex, @pParameter))
   RETURN pParameter
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Returns a reference to the collection's enumerator.
' ===========================================================================================
PRIVATE FUNCTION CAdoParameters.NewEnum () AS Afx_IUnknown PTR
   CADODB_DP("CAdoParameters.NewEnum")
   DIM pUnk AS Afx_IUnknown PTR
   IF m_pParameters THEN SetResult(m_pParameters->_NewEnum(@pUnk))
   RETURN pUnk
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Refreshes the contents of the Properties collection.
' ===========================================================================================
PRIVATE FUNCTION CAdoParameters.Refresh () AS HRESULT
   CADODB_DP("CAdoParameters.Refresh")
   IF m_pParameters THEN RETURN(SetResult(m_pParameters->Refresh))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Appends an object to the collection.
' ===========================================================================================
PRIVATE FUNCTION CAdoParameters.Append (BYVAL pObject AS Afx_IDispatch PTR) AS HRESULT
   CADODB_DP("CAdoParameters.Append - pObject")
   IF m_pParameters THEN RETURN(SetResult(m_pParameters->Append(pObject)))
END FUNCTION
' ===========================================================================================
' ===========================================================================================
' Appends a parameter to the collection.
' ===========================================================================================
PRIVATE FUNCTION CAdoParameters.Append (BYREF pParam AS CAdoParameter) AS HRESULT
   CADODB_DP("CAdoParameters.Append - CAdoParameter")
   IF m_pParameters THEN RETURN(SetResult(m_pParameters->Append(cast(ANY PTR, pParam.m_pParameter))))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Deletes an object from the Parameters collection.
' ===========================================================================================
PRIVATE FUNCTION CAdoParameters.Delete_ (BYREF cvIndex AS CVAR) AS HRESULT
   CADODB_DP("CAdoParameters.Delete_")
   IF m_pParameters THEN RETURN(SetResult(m_pParameters->Delete_(cvIndex)))
END FUNCTION
' ===========================================================================================

END NAMESPACE
