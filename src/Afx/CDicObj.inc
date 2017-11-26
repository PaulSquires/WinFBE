' ########################################################################################
' Microsoft Windows
' File: CDicObj.inc
' Contents: Dictionary class.
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2017 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "windows.bi"
#include once "win/ole2.bi"
#include once "Afx/AfxWin.inc"
#include once "Afx/AfxScrRun.bi"
#include once "Afx/CVar.inc"

NAMESPACE Afx

' ========================================================================================
' Macro for debug
' To allow debugging, define _CBDICOBJ_DEBUG_ 1 in your application before including this file.
' ========================================================================================
#ifndef _CBDICOBJ_DEBUG_
   #define _CBDICOBJ_DEBUG_ 0
#ENDIF
#ifndef _CBDICOBJ_DP_
   #define _CBDICOBJ_DP_ 1
   #MACRO CDICOBJ_DP(st)
      #IF (_CBDICOBJ_DEBUG_ = 1)
         OutputDebugStringW(st)
      #ENDIF
   #ENDMACRO
#ENDIF
' ========================================================================================

' ========================================================================================
' CDicObj class
' ========================================================================================
TYPE CDicObj

Public:
   m_Result AS HRESULT
   m_pDictionary AS Afx_IDictionary PTR

Private:
   m_bUninitCOM AS BOOLEAN

Public:
   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYREF pDic AS CDicObj)
   DECLARE DESTRUCTOR
   DECLARE FUNCTION DispPtr () AS ANY PTR
'   DECLARE FUNCTION DispObj () AS ANY PTR
   DECLARE FUNCTION GetLastResult () AS HRESULT
   DECLARE FUNCTION SetResult (BYVAL Result AS HRESULT) AS HRESULT
   DECLARE PROPERTY Item (BYREF cvKey AS CVAR, BYREF cvNewItem AS CVAR)
   DECLARE PROPERTY Item (BYREF cvKey AS CVAR) AS CVAR
   DECLARE FUNCTION Add (BYREF cvKey AS CVAR, BYREF cvItem AS CVAR) AS HRESULT
   DECLARE FUNCTION Count () AS LONG
   DECLARE FUNCTION Exists (BYREF cvKey AS CVAR) AS BOOLEAN
   DECLARE FUNCTION Items () AS CVAR
   DECLARE PROPERTY Key (BYREF cvKey AS CVAR, BYREF cbsNewKey AS CVAR)
   DECLARE FUNCTION Keys () AS CVAR
   DECLARE FUNCTION Remove (BYREF cvKey AS CVAR) AS HRESULT
   DECLARE FUNCTION RemoveAll() AS HRESULT
   DECLARE FUNCTION NewEnum () AS IUnknown PTR
   DECLARE FUNCTION HashVal (BYREF cvKey AS CVAR) AS CVAR

END TYPE
' ========================================================================================

' ========================================================================================
' CDicObj constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CDicObj
   ' // Initialize the COM library
   DIM hr AS HRESULT = CoInitialize(NULL)
   IF hr = S_OK OR hr = S_FALSE THEN m_bUninitCOM = TRUE
   ' // Create an instance of the Dictionary object
   DIM CLSID_Dictionary_ AS CLSID = (&hEE09B103, &h97E0, &h11CF, {&h97, &h8F, &h00, &hA0, &h24, &h63, &hE0, &h6F})
   DIM IID_IDictionary_  AS GUID  = (&h42C642C1, &h97E1, &h11CF, {&h97, &h8F, &h00, &hA0, &h24, &h63, &hE0, &h6F})
   CoCreateInstance(@CLSID_Dictionary_, NULL, CLSCTX_INPROC_SERVER, @IID_IDictionary_, @m_pDictionary)
   CDICOBJ_DP("CDICOBJ CONSTRUCTOR Default - m_pDictionary = " & WSTR(m_pDictionary))
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Creates a new instance of the Dictionary object and copies the contents of the
' passed CDicObj.
' ========================================================================================
PRIVATE CONSTRUCTOR CDicObj (BYREF pDic AS CDicObj)
   ' // Initialize the COM library
   DIM hr AS HRESULT = CoInitialize(NULL)
   IF hr = S_OK OR hr = S_FALSE THEN m_bUninitCOM = TRUE
   ' // Create an instance of the Dictionary object
   DIM CLSID_Dictionary_ AS CLSID = (&hEE09B103, &h97E0, &h11CF, {&h97, &h8F, &h00, &hA0, &h24, &h63, &hE0, &h6F})
   DIM IID_IDictionary_  AS GUID  = (&h42C642C1, &h97E1, &h11CF, {&h97, &h8F, &h00, &hA0, &h24, &h63, &hE0, &h6F})
   CoCreateInstance(@CLSID_Dictionary_, NULL, CLSCTX_INPROC_SERVER, @IID_IDictionary_, @m_pDictionary)
   ' // Copy the contents of the passed CDicObj
   DIm cvKey AS CVAR, cvValue AS CVAR
   DIM cvKeys AS CVAR = pDic.Keys
   FOR i AS LONG = cvKeys.GetLBound TO cvKeys.GetUBound
      cvKey = cvKeys.GetVariantElem(i)
      cvValue = pDic.Item(cvKey)
      pDic.Add(cvKey, cvValue)
   NEXT
   CDICOBJ_DP("CDICOBJ CONSTRUCTOR pDic = " & WSTR(m_pDictionary) & " - m_pDictionary = " & WSTR(m_pDictionary))
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CDicObj destructor
' ========================================================================================
PRIVATE DESTRUCTOR CDicObj
   CDICOBJ_DP("CDICOBJ DESTRUCTOR - m_pDictionary = " & WSTR(m_pDictionary))
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
' Returns the underlying dispatch pointer.
' Don't call IUnknown_Release on it.
' ========================================================================================
PRIVATE FUNCTION CDicObj.DispPtr () AS ANY PTR
   RETURN m_pDictionary
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a counted reference of the underlying dispatch pointer.
' You must call IUnknown_Release when no longer need it.
' Note: It is not possible to clone the object this way because the destructor of the
' class removes all the contents.
' ========================================================================================
'PRIVATE FUNCTION CDicObj.DispObj () AS ANY PTR
'   IF m_pDictionary <> NULL THEN
'      m_pDictionary->AddRef
'      RETURN m_pDictionary
'   END IF
'END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the last result code.
' ========================================================================================
PRIVATE FUNCTION CDicObj.GetLastResult () AS HRESULT
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the last result code.
' ========================================================================================
PRIVATE FUNCTION CDicObj.SetResult (BYVAL Result AS HRESULT) AS HRESULT
   m_Result = Result
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Item property
' Sets or returns an item for a specified cbsKey in the associative array.
' Arguments:
' - cbsKey     : Key associated with the item being retrieved or added.
' - cbsNewItem : The new value associated with the specified key.
' Remarks:
' If key is not found when changing an item, a new key is created with the specified cbsNewItem.
' If key is not found when attempting to return an existing item, a new key is created and
' its corresponding item is left empty.
' ========================================================================================
PRIVATE PROPERTY CDicObj.Item (BYREF cvKey AS CVAR, BYREF cvNewItem AS CVAR)
   CDICOBJ_DP("CDICOBJ PutItem")
   SetResult(m_pDictionary->put_Item(cvKey, cvNewItem))
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CDicObj.Item (BYREF cvKey AS CVAR) AS CVAR
   CDICOBJ_DP("CDICOBJ GettItem")
   DIM cvItem AS CVAR
   SetResult(m_pDictionary->get_Item(cvKey, cvItem))
   PROPERTY = cvItem
END PROPERTY
' ========================================================================================

' ========================================================================================
' Add method
' Adds a key and item pair to the associative array.
' Arguments:
' - cbsKey  : The key associated with the item being added.
' - cbsItem : The item associated with the key being added.
' Remarks:
' An error occurs if the key already exists.
' ========================================================================================
PRIVATE FUNCTION CDicObj.Add (BYREF cvKey AS CVAR, BYREF cvItem AS CVAR) AS HRESULT
   CDICOBJ_DP("CDICOBJ Add")
   RETURN SetResult(m_pDictionary->Add(cvKey, cvItem))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Count property
' Returns the number of items in the associative array.
' ========================================================================================
PRIVATE FUNCTION CDicObj.Count () AS LONG
   CDICOBJ_DP("CDICOBJ Count")
   DIM nCount AS LONG
   SetResult(m_pDictionary->get_Count(@nCount))
   RETURN nCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' Exists method
' Returns TRUE if a specified key exists in the associative array, FALSE if it does not.
' Argument:
' - cbsKey : Key value being searched for the associative array.
' ========================================================================================
PRIVATE FUNCTION CDicObj.Exists (BYREF cvKey AS CVAR) AS BOOLEAN
   CDICOBJ_DP("CDICOBJ Exists")
   DIM nExists AS SHORT
   SetResult(m_pDictionary->Exists(cvKey, @nExists))
   RETURN nExists
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns an array containing all the items in a associative array.
' ========================================================================================
PRIVATE FUNCTION CDicObj.Items () AS CVAR
   CDICOBJ_DP("CDICOBJ Items")
   DIM cvItems AS CVAR
   SetResult(m_pDictionary->Items(cvItems))
   RETURN cvItems
END FUNCTION
' ========================================================================================

' ========================================================================================
' Key property
' Arguments:
' - cbsKey  : Key value being changed.
' - cbsItem : New value that replaces the specified key.
' Remarks:
' If key is not found when changing a key, a new key is created and its associated item is left empty.
' Usage example: pDic.Key("a") = "m"
' ========================================================================================
PRIVATE PROPERTY CDicObj.Key (BYREF cvKey AS CVAR, BYREF cvNewKey AS CVAR)
   CDICOBJ_DP("CDICOBJ PutKey")
   SetResult(m_pDictionary->put_Key(cvKey, cvNewKey))
END PROPERTY
' ========================================================================================

' ========================================================================================
' Keys method
' Returns an array containing all existing keys in the associative array.
' ========================================================================================
PRIVATE FUNCTION CDicObj.Keys () AS CVAR
   CDICOBJ_DP("CDICOBJ Keys")
   DIM cvKeys AS CVAR
   SetResult(m_pDictionary->Keys(cvKeys))
   RETURN cvKeys
END FUNCTION
' ========================================================================================

' ========================================================================================
' Remove method
' Removes a key, item pair from the associative array.
' Argument:
' - cbsKey : Key associated with the key, item pair you want to remove from the associative array.
' Remarks:
' An error occurs if the specified key, item pair does not exist.
' ========================================================================================
PRIVATE FUNCTION CDicObj.Remove (BYREF cvKey AS CVAR) AS HRESULT
   CDICOBJ_DP("CDICOBJ Remove")
   RETURN SetResult(m_pDictionary->Remove(cvKey))
END FUNCTION
' ========================================================================================

' ========================================================================================
' RemoveAll method
' Removes all key, item pairs from the associative array.
' ========================================================================================
PRIVATE FUNCTION CDicObj.RemoveAll() AS HRESULT
   CDICOBJ_DP("CDICOBJ RemoveAll")
   RETURN SetResult(m_pDictionary->RemoveAll)
END FUNCTION
' ========================================================================================

' ========================================================================================
' NewEnum method
' Returns a reference to the standard IEnumVARIANT interface.
' ========================================================================================
PRIVATE FUNCTION CDicObj.NewEnum () AS IUnknown PTR
   CDICOBJ_DP("CDICOBJ NewEnum")
   DIM ppUnk AS IUnknown PTR
   SetResult(m_pDictionary->get__NewEnum(cast(ANY PTR, @ppunk)))
   RETURN ppunk
END FUNCTION
' ========================================================================================

' ========================================================================================
' HashVal property
' Returns the hash value for a specified key in the associative array.
' ========================================================================================
PRIVATE FUNCTION CDicObj.HashVal (BYREF cvKey AS CVAR) AS CVAR
   CDICOBJ_DP("CDICOBJ HashVal")
   DIM cvHashVal AS CVAR
   SetResult(m_pDictionary->get_HashVal(cvKey, cvHashVal))
   RETURN cvHashVal
END FUNCTION
' ========================================================================================

END NAMESPACE
