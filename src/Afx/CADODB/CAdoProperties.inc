' ########################################################################################
' Microsoft Windows
' Contents: ADO Properties
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once

NAMESPACE Afx

' ########################################################################################
' CAdoProperty class
' ########################################################################################
TYPE CAdoProperty EXTENDS CAdoBase

   m_Result AS HRESULT
   m_pProperty AS Afx_ADOProperty PTR

   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYVAL pProperty AS Afx_ADOProperty PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   DECLARE DESTRUCTOR
'   DECLARE OPERATOR @ () AS Afx_ADOProperty PTR PTR
   DECLARE FUNCTION vptr () AS Afx_ADOProperty PTR PTR

   ' // Attaches a reference to an Afx_ADOProperty object to the class
   DECLARE SUB Attach (BYVAL pProperty AS Afx_ADOProperty PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   ' // Retrieves the name of a field
   DECLARE PROPERTY Name () AS CBSTR
   ' // Sets or returns a Variant value that indicates the value of the object
   DECLARE PROPERTY Value () AS CVAR
   DECLARE PROPERTY Value (BYREF cvValue AS CVAR)
   ' // Indicates the operational type or data type of a Property object.
   DECLARE PROPERTY Type_ () AS DataTypeEnum
   ' // Indicates one or more characteristics of an object.
   DECLARE PROPERTY Attributes () AS LONG
   DECLARE PROPERTY Attributes (BYVAL lAttr AS LONG)

END TYPE
' ########################################################################################

' ========================================================================================
' Default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CAdoProperty
   CADODB_DP("CAdoProperty CONSTRUCTOR Default")
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CAdoProperty (BYVAL pProperty AS Afx_ADOProperty PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   CADODB_DP("-BEGIN CAdoProperty CONSTRUCTOR")
   IF pProperty THEN
      m_pProperty = pProperty
      IF fAddRef THEN m_pProperty->AddRef
   END IF
END CONSTRUCTOR
' ========================================================================================

' ===========================================================================================
' Cleanup
' ===========================================================================================
PRIVATE DESTRUCTOR CAdoProperty
   CADODB_DP("CAdoProperty DESTRUCTOR " & WSTR(m_pProperty))
   IF m_pProperty THEN m_pProperty->Release
END DESTRUCTOR
' ===========================================================================================

' ========================================================================================
' Returns the address of the connection pointer
' ========================================================================================
'PRIVATE OPERATOR CAdoProperty.@ () AS Afx_ADOProperty PTR PTR
'   OPERATOR = @m_pProperty
'END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CAdoProperty.vptr () AS Afx_ADOProperty PTR PTR
   CVAR_DP("CAdoProperty vptr")
   IF m_pProperty THEN m_pProperty->Release
   RETURN @m_pProperty
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the encapsulated IDispatch pointer
' ========================================================================================
PRIVATE OPERATOR * (BYREF cProperty AS CAdoProperty) AS Afx_ADOProperty PTR
   OPERATOR = cProperty.m_pProperty
END OPERATOR
' ========================================================================================

' ===========================================================================================
' Attaches a reference to an Afx_ADOProperty object to the class
' ===========================================================================================
PRIVATE SUB CAdoProperty.Attach (BYVAL pProperty AS Afx_ADOProperty PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   CADODB_DP("CAdoProperty.Attach")
   IF m_pProperty THEN m_pProperty->Release
   m_pProperty = pProperty
   ' // Increase the reference count
   IF fAddRef THEN
      IF m_pProperty THEN m_pProperty->AddRef
   END IF
END SUB
' ===========================================================================================

' ===========================================================================================
' Retrieves the name of a property
' ===========================================================================================
PRIVATE PROPERTY CAdoProperty.Name () AS CBSTR
   CADODB_DP("CAdoProperty.Name")
   DIM bstrName AS AFX_BSTR
   IF m_pProperty THEN SetResult(m_pProperty->get_Name(@bstrName))
   RETURN bstrName
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the operational type or data type of a property.
' ===========================================================================================
PRIVATE PROPERTY CAdoProperty.Value () AS CVAR
   CADODB_DP("CAdoProperty.Value")
   DIM vValue AS VARIANT
   IF m_pProperty THEN SetResult(m_pProperty->get_Value(@vValue))
   RETURN vValue
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the operational type or data type of a property.
' ===========================================================================================
PRIVATE PROPERTY CAdoProperty.Value (BYREF cvValue AS CVAR)
   CADODB_DP("CAdoProperty.Value")
   IF m_pProperty THEN SetResult(m_pProperty->put_Value(cvValue))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the operational type or data type of a Property object.
' ===========================================================================================
PRIVATE PROPERTY CAdoProperty.Type_ () AS DataTypeEnum
   CADODB_DP("CAdoProperty.Type")
   DIM nType AS DataTypeEnum
   IF m_pProperty THEN SetResult(m_pProperty->get_Type(@nType))
   RETURN nType
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the operational type or data type of a property.
' ===========================================================================================
PRIVATE PROPERTY CAdoProperty.Attributes () AS LONG
   CADODB_DP("CAdoProperty.Attributes")
   DIM lAttr AS LONG
   IF m_pProperty THEN SetResult(m_pProperty->get_Attributes(@lAttr))
   RETURN lAttr
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the operational type or data type of a property.
' ===========================================================================================
PRIVATE PROPERTY CAdoProperty.Attributes (BYVAL lAttr AS LONG)
   CADODB_DP("CAdoProperty.Attributes")
   IF m_pProperty THEN SetResult(m_pProperty->put_Attributes(lAttr))
END PROPERTY
' ===========================================================================================

' ########################################################################################
' CAdoProperties class
' ########################################################################################
TYPE CAdoProperties EXTENDS CAdoBase

   m_Result AS HRESULT
   m_pProperties AS Afx_ADOProperties PTR

   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYVAL pProperties AS Afx_ADOProperties PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   DECLARE DESTRUCTOR
'   DECLARE OPERATOR @ () AS Afx_ADOProperties PTR PTR
   DECLARE FUNCTION vptr () AS Afx_ADOProperties PTR PTR

   ' // Attaches a reference to an Afx_ADOProperties object to the class
   DECLARE SUB Attach (BYVAL pProperties AS Afx_ADOProperties PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   ' // Refreshes the contents of the Properties collection.
   DECLARE FUNCTION Refresh () AS HRESULT
   ' // Returns a reference to the collection's enumerator.
   DECLARE FUNCTION NewEnum () AS Afx_IUnknown PTR
   ' // Retrieves the number of properties
   DECLARE PROPERTY Count () AS LONG
   ' // Indicates a specific member of the Properties collection, by name or ordinal number.
   DECLARE PROPERTY Item (BYREF cvIndex AS CVAR) AS Afx_ADOProperty PTR

END TYPE
' ########################################################################################

' ========================================================================================
' Default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CAdoProperties
   CADODB_DP("BEGIN CAdoProperties CONSTRUCTOR Default")
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CAdoProperties (BYVAL pProperties AS Afx_ADOProperties PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   CADODB_DP("-BEGIN CAdoProperties CONSTRUCTOR")
   IF pProperties THEN
      m_pProperties = pProperties
      IF fAddRef THEN m_pProperties->AddRef
   END IF
END CONSTRUCTOR
' ========================================================================================

' ===========================================================================================
' Cleanup
' ===========================================================================================
PRIVATE DESTRUCTOR CAdoProperties
   CADODB_DP("CAdoProperties DESTRUCTOR " & WSTR(m_pProperties))
   IF m_pProperties THEN m_pProperties->Release
END DESTRUCTOR
' ===========================================================================================

' ========================================================================================
' Returns the address of the connection pointer
' ========================================================================================
'PRIVATE OPERATOR CAdoProperties.@ () AS Afx_ADOProperties PTR PTR
'   OPERATOR = @m_pProperties
'END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CAdoProperties.vptr () AS Afx_ADOProperties PTR PTR
   CVAR_DP("CAdoProperties vptr")
   IF m_pProperties THEN m_pProperties->Release
   RETURN @m_pProperties
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the encapsulated IDispatch pointer
' ========================================================================================
PRIVATE OPERATOR * (BYREF cFields AS CAdoProperties) AS Afx_ADOProperties PTR
   OPERATOR = cFields.m_pProperties
END OPERATOR
' ========================================================================================

' ===========================================================================================
' Attaches a reference to an Afx_ADOProperties object to the class
' ===========================================================================================
PRIVATE SUB CAdoProperties.Attach (BYVAL pProperties AS Afx_ADOProperties PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   CADODB_DP("CAdoProperties.Attach")
   IF m_pProperties THEN m_pProperties->Release
   m_pProperties = pProperties
   ' // Increase the reference count
   IF fAddRef THEN
      IF m_pProperties THEN m_pProperties->AddRef
   END IF
END SUB
' ===========================================================================================

' ===========================================================================================
' Refreshes the contents of the Properties collection.
' ===========================================================================================
PRIVATE FUNCTION CAdoProperties.Refresh () AS HRESULT
   CADODB_DP("CAdoProperties.Refresh")
   IF m_pProperties THEN RETURN(SetResult(m_pProperties->Refresh))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Returns a reference to the collection's enumerator.
' ===========================================================================================
PRIVATE FUNCTION CAdoProperties.NewEnum () AS Afx_IUnknown PTR
   CADODB_DP("CAdoProperties.NewEnum")
   DIM pUnk AS Afx_IUnknown PTR
   IF m_pProperties THEN SetResult(m_pProperties->_NewEnum(@pUnk))
   RETURN pUnk
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Retrieves the number of fields
' ===========================================================================================
PRIVATE PROPERTY CAdoProperties.Count () AS LONG
   CADODB_DP("CAdoProperties.Count")
   DIM nCount AS LONG
   IF m_pProperties THEN SetResult(m_pProperties->get_Count(@nCount))
   RETURN nCount
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates a specific member of the Properties collection, by name or ordinal number.
' ===========================================================================================
PRIVATE PROPERTY CAdoProperties.Item (BYREF cvIndex AS CVAR) AS Afx_ADOProperty PTR
   CADODB_DP("CAdoProperties.Item")
   DIM pProperty AS Afx_ADOProperty PTR
   IF m_pProperties THEN SetResult(m_pProperties->get_Item(cvIndex, @pProperty))
   RETURN pProperty
END PROPERTY
' ===========================================================================================

END NAMESPACE
