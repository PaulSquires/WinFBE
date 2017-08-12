' ########################################################################################
' Microsoft Windows
' File: CWstrDic.inc
' Contents: CWSTR dictionary class.
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "windows.bi"
#include once "win/ole2.bi"
#include once "Afx/AfxScrRun.bi"
#include once "Afx/CWStr.inc"
#include once "Afx/CWStrArray.inc"

NAMESPACE Afx

' ========================================================================================
' Macro for debug
' To allow debugging, define _CWSTRDIC_DEBUG_ 1 in your application before including this file.
' ========================================================================================
#ifndef _CWSTRDIC_DEBUG_
   #define _CWSTRDIC_DEBUG_ 0
#ENDIF
#ifndef _CWSTRDIC_DP_
   #define _CWSTRDIC_DP_ 1
   #MACRO CWSTRDIC_DP(st)
      #IF (_CWSTRDIC_DEBUG_ = 1)
         OutputDebugStringW(st)
      #ENDIF
   #ENDMACRO
#ENDIF
' ========================================================================================

' ========================================================================================
' CWStrDic class
' ========================================================================================
TYPE CWStrDic

   Public:
      m_Result AS HRESULT
      m_pDictionary AS Afx_IDictionary PTR

   Private:
      m_bUninitCOM AS BOOLEAN

   Public:
      DECLARE CONSTRUCTOR
      DECLARE DESTRUCTOR
      DECLARE FUNCTION GetLastResult () AS HRESULT
      DECLARE FUNCTION SetResult (BYVAL Result AS HRESULT) AS HRESULT
      DECLARE PROPERTY Item (BYREF cwsKey AS CWSTR, BYREF cwsNewItem AS CWSTR)
      DECLARE PROPERTY Item (BYREF cwsKey AS CWSTR) AS CWSTR
      DECLARE FUNCTION Add (BYREF cwsKey AS CWSTR, BYREF cwsItem AS CWSTR) AS HRESULT
      DECLARE FUNCTION Count () AS LONG
      DECLARE FUNCTION Exists (BYREF cwsKey AS CWSTR) AS BOOLEAN
      DECLARE FUNCTION Items () AS CWstrArray
      DECLARE PROPERTY Key (BYREF cwsKey AS CWSTR, BYREF cwsNewKey AS CWSTR)
      DECLARE FUNCTION Keys () AS CWstrArray
      DECLARE FUNCTION Remove (BYREF cwsKey AS CWSTR) AS HRESULT
      DECLARE FUNCTION RemoveAll() AS HRESULT

END TYPE
' ========================================================================================

' ========================================================================================
' CWSTRDIC constructor
' ========================================================================================
CONSTRUCTOR CWStrDic
   ' // Initialize the COM library
   DIM hr AS HRESULT = CoInitialize(NULL)
   IF hr = S_OK OR hr = S_FALSE THEN m_bUninitCOM = TRUE
   ' // Create an instance of the Dictionary object
   DIM CLSID_Dictionary_ AS CLSID = (&hEE09B103, &h97E0, &h11CF, {&h97, &h8F, &h00, &hA0, &h24, &h63, &hE0, &h6F})
   DIM IID_IDictionary_  AS GUID  = (&h42C642C1, &h97E1, &h11CF, {&h97, &h8F, &h00, &hA0, &h24, &h63, &hE0, &h6F})
   CoCreateInstance(@CLSID_Dictionary_, NULL, CLSCTX_INPROC_SERVER, @IID_IDictionary_, @m_pDictionary)
   CWSTRDIC_DP("CWSTRDIC CONSTRUCTOR Default - m_pDictionary = " & WSTR(m_pDictionary))
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CWSTRDIC destructor
' ========================================================================================
DESTRUCTOR CWStrDic
   CWSTRDIC_DP("CWSTRDIC DESTRUCTOR")
   ' // Release the Dictionary object
   IF m_pDictionary THEN
      m_pDictionary->RemoveAll
      m_pDictionary->Release
   END IF
   ' // Free the COM library
   IF m_bUninitCOM THEN CoUninitialize
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the last result code.
' ========================================================================================
PRIVATE FUNCTION CWstrDic.GetLastResult () AS HRESULT
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the last result code.
' ========================================================================================
PRIVATE FUNCTION CWstrDic.SetResult (BYVAL Result AS HRESULT) AS HRESULT
   m_Result = Result
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Item property
' Sets or returns an item for a specified cwsKey in the associative array.
' Arguments:
' - cwsKey     : Key associated with the item being retrieved or added.
' - cwsNewItem : The new value associated with the specified key.
' Remarks:
' If key is not found when changing an item, a new key is created with the specified cbsNewItem.
' If key is not found when attempting to return an existing item, a new key is created and
' its corresponding item is left empty.
' ========================================================================================
PRIVATE PROPERTY CWstrDic.Item (BYREF cwsKey AS CWSTR, BYREF cwsNewItem AS CWSTR)
   CWSTRDIC_DP("CWSTRDIC PutItem")
   DIM vKey AS VARIANT, vNewItem AS VARIANT
   V_VT(@vKey) = VT_BSTR : V_BSTR(@vKey) = SysAllocString(cwsKey)
   V_VT(@vNewItem) = VT_BSTR : V_BSTR(@vNewItem) = SysAllocString(cwsNewItem)
   SetResult(m_pDictionary->put_Item(@vKey, @vNewItem))
   VariantClear @vKey : VariantClear(@vNewItem)
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CWstrDic.Item (BYREF cwsKey AS CWSTR) AS CWSTR
   CWSTRDIC_DP("CWSTRDIC GetItem")
   DIM vKey AS VARIANT, vRetItem AS VARIANT
   V_VT(@vKey) = VT_BSTR : V_BSTR(@vKey) = SysAllocString(cwsKey)
   SetResult(m_pDictionary->get_Item(@vKey, @vRetItem))
   PROPERTY = *cast(AFX_BSTR, vRetItem.bstrVal)
   VariantClear @vKey : VariantClear @vRetItem
END PROPERTY
' ========================================================================================

' ========================================================================================
' Add method
' Adds a key and item pair to the associative array.
' Arguments:
' - cwsKey  : The key associated with the item being added.
' - cwsItem : The item associated with the key being added.
' Remarks:
' An error occurs if the key already exists.
' ========================================================================================
PRIVATE FUNCTION CWstrDic.Add (BYREF cwsKey AS CWSTR, BYREF cwsItem AS CWSTR) AS HRESULT
   CWSTRDIC_DP("CWSTRDIC Add")
   DIM vKey AS VARIANT, vItem AS VARIANT
   V_VT(@vKey) = VT_BSTR : V_BSTR(@vKey) = SysAllocString(cwsKey)
   V_VT(@vItem) = VT_BSTR : V_BSTR(@vItem) = SysAllocString(cwsItem)
   FUNCTION = SetResult(m_pDictionary->Add(@vKey, @vItem))
   VariantClear @vKey : VariantClear(@vItem)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Count property
' Returns the number of items in the associative array.
' ========================================================================================
PRIVATE FUNCTION CWstrDic.Count () AS LONG
   CWSTRDIC_DP("CWSTRDIC Count")
   DIM nCount AS LONG
   SetResult(m_pDictionary->get_Count(@nCount))
   RETURN nCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' Exists method
' Returns TRUE if a specified key exists in the associative array, FALSE if it does not.
' Argument:
' - cwsKey : Key value being searched for the associative array.
' ========================================================================================
PRIVATE FUNCTION CWstrDic.Exists (BYREF cwsKey AS CWSTR) AS BOOLEAN
   CWSTRDIC_DP("CWSTRDIC Exists")
   DIM vKey AS VARIANT, nExists AS SHORT
   V_VT(@vKey) = VT_BSTR : V_BSTR(@vKey) = SysAllocString(cwsKey)
   SetResult(m_pDictionary->Exists(@vKey, @nExists))
   VariantClear @vKey
   RETURN nExists
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns an array containing all the items in a associative array.
' ========================================================================================
PRIVATE FUNCTION CWstrDic.Items () AS CWstrArray
   CWSTRDIC_DP("CWSTRDIC Items")
   DIM vItems AS VARIANT
   SetResult(m_pDictionary->Items(@vItems))
   IF m_Result = S_OK THEN
      ' // Get the lower and upper bounds
      DIM plLBound AS LONG, plUBound AS LONG
      SafeArrayGetLBound(vItems.parray, 1, @plLBound)
      SafeArrayGetUBound(vItems.parray, 1, @plUBound)
      ' // Create an array of CWSTRs
      DIM cwsa AS CWstrArray
      ' // Access the data directly
      DIM pvData AS VARIANT PTR
      SafeArrayAccessData(vItems.parray, @pvData)
      IF pvData THEN
         FOR i AS LONG = plLBound TO plUBound
            SetResult(cwsa.Append(*cast(AFX_BSTR, pvData[i].bstrVal)))
         NEXT
      END IF
      SafeArrayUnaccessData(vItems.parray)
      RETURN cwsa
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Key property
' Arguments:
' - cwsKey  : Key value being changed.
' - cbsItem : New value that replaces the specified key.
' Remarks:f
' If key is not found when changing a key, a new key is created and its associated item is left empty.
' Usage example: pDic.Key("a") = "m"
' ========================================================================================
PRIVATE PROPERTY CWstrDic.Key (BYREF cwsKey AS CWSTR, BYREF cwsNewKey AS CWSTR)
   CWSTRDIC_DP("CWSTRDIC PutKey")
   DIM vKey AS VARIANT, vNewKey AS VARIANT
   V_VT(@vKey) = VT_BSTR : V_BSTR(@vKey) = SysAllocString(cwsKey)
   V_VT(@vNewKey) = VT_BSTR : V_BSTR(@vNewKey) = SysAllocString(cwsNewKey)
   SetResult(m_pDictionary->put_Key(@vKey, @vNewKey))
   VariantClear @vKey : VariantClear @vNewKey
END PROPERTY
' ========================================================================================

' ========================================================================================
' Keys method
' Returns an array containing all existing keys in the associative array.
' ========================================================================================
PRIVATE FUNCTION CWstrDic.Keys () AS CWstrArray
   CWSTRDIC_DP("CWSTRDIC Keys")
   DIM vKeys AS VARIANT
   SetResult(m_pDictionary->Keys(@vKeys))
   IF m_Result = S_OK THEN
      ' // Get the lower and upper bounds
      DIM plLBound AS LONG, plUBound AS LONG
      SafeArrayGetLBound(vKeys.parray, 1, @plLBound)
      SafeArrayGetUBound(vKeys.parray, 1, @plUBound)
      ' // Create an array of CWSTRs
      DIM cwsa AS CWstrArray
      ' // Access the data directly
      DIM pvData AS VARIANT PTR
      SafeArrayAccessData(vKeys.parray, @pvData)
      IF pvData THEN
         FOR i AS LONG = plLBound TO plUBound
            SetResult(cwsa.Append(*cast(AFX_BSTR, pvData[i].bstrVal)))
         NEXT
      END IF
      SafeArrayUnaccessData(vKeys.parray)
      RETURN cwsa
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Remove method
' Removes a key, item pair from the associative array.
' Argument:
' - cwsKey : Key associated with the key, item pair you want to remove from the associative array.
' Remarks:
' An error occurs if the specified key, item pair does not exist.
' ========================================================================================
PRIVATE FUNCTION CWstrDic.Remove (BYREF cwsKey AS CWSTR) AS HRESULT
   CWSTRDIC_DP("CWSTRDIC Remove")
   DIM vKey AS VARIANT
   V_VT(@vKey) = VT_BSTR : V_BSTR(@vKey) = SysAllocString(cwsKey)
   FUNCTION = SetResult(m_pDictionary->Remove(@vKey))
   VariantClear @vKey
END FUNCTION
' ========================================================================================

' ========================================================================================
' RemoveAll method
' Removes all key, item pairs from the associative array.
' ========================================================================================
PRIVATE FUNCTION CWstrDic.RemoveAll() AS HRESULT
   CWSTRDIC_DP("CWSTRDIC RemoveAll")
   RETURN SetResult(m_pDictionary->RemoveAll)
END FUNCTION
' ========================================================================================

END NAMESPACE
