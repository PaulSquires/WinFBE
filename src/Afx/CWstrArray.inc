' ########################################################################################
' Microsoft Windows
' File: CWstrArray.inc
' Contents: One-dimensional safe array class of strings.
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "windows.bi"
#include once "Afx/CWStr.inc"

NAMESPACE Afx

' // The definition for BSTR in the FreeBASIC headers was incorrectly changed to WCHAR
#ifndef AFX_BSTR
   #define AFX_BSTR WSTRING PTR
#endif

' ========================================================================================
' Macro for debug
' To allow debugging, define _CWstrArray_DEBUG_ 1 in your application before including this file.
' ========================================================================================
#ifndef _CWSTRARRAY_DEBUG_
   #define _CWSTRARRAY_DEBUG_ 0
#ENDIF
#ifndef _CWSTRARRAY_DP_
   #define _CWSTRARRAY_DP_ 1
   #MACRO CWSTRARRAY_DP(st)
      #IF (_CWSTRARRAY_DEBUG_ = 1)
         OutputDebugStringW(st)
      #ENDIF
   #ENDMACRO
#ENDIF
' ========================================================================================

' ########################################################################################
' CWstrArray - One-dimensional safe array class of strings
' ########################################################################################
TYPE CWstrArray

   DIM m_psa AS SAFEARRAY PTR

   ' // * Creates an instance of the class with a null safe array pointer
   DECLARE CONSTRUCTOR
   ' // * Creates a one-dimensional BSTR safe array with the specified number of elements and lower bound.
   DECLARE CONSTRUCTOR (BYVAL cElements AS DWORD = 0, BYVAL lLBound AS LONG = 0)
   ' // * Creates a one-dimensional BSTR safe array from another CWstrArray.
   DECLARE CONSTRUCTOR (BYREF cwsa AS CWstrArray)
   ' // * Creates a one-dimensional BSTR safe array from a safe array.
   DECLARE CONSTRUCTOR (BYVAL psaSrc AS SAFEARRAY PTR)
   ' // * Creates a one-dimensional BSTR safe array from a variant.
   DECLARE CONSTRUCTOR (BYVAL pvar AS VARIANT PTR)
   ' // * Destroys the safe array when the class is destroyed.
   DECLARE DESTRUCTOR
   ' // Returns the address of the safe array descriptor.
   DECLARE OPERATOR @ () AS ANY PTR
   ' // * Returns a pointer to the safe array descriptor.
   DECLARE FUNCTION GetPtr () AS SAFEARRAY PTR
   ' // * Assigns a CWstrArray.
   DECLARE OPERATOR LET (BYREF cwsa AS CWstrArray)
   ' // * Assigns a safe array.
   DECLARE OPERATOR LET (BYVAL psa AS SAFEARRAY PTR)
   ' // * Creates a one-dimensional BSTR SafeArray
   DECLARE FUNCTION Create (BYVAL cElements AS DWORD, BYVAL lLBound AS LONG) AS SAFEARRAY PTR
   ' // * Destroys an existing array descriptor and all of the data in the array.
   DECLARE FUNCTION Destroy () AS HRESULT
   ' // * Returns the lower bound of the safe array.
   DECLARE FUNCTION LBound () AS LONG
   ' // * Returns the upper bound of the safe array.
   DECLARE FUNCTION UBound () AS LONG
   ' // * Returns the number of elements in the the array.
   DECLARE FUNCTION Count () AS ULONG
   ' // * Returns the number of times the array has been locked without the corresponding unlock.
   DECLARE FUNCTION LocksCount () AS ULONG
   ' // * Gets an element of the array. If the function fails, it returns an empty string.
   DECLARE PROPERTY Item (BYVAL idx AS LONG) AS CWSTR
   DECLARE OPERATOR [] (BYVAL idx AS LONG) AS CWSTR
   ' // * Puts a string element at a given location in the array.
   DECLARE PROPERTY Item (BYVAL idx AS LONG, BYREF cws AS CWSTR)
   ' // * Destroys all the data in the safe array.
   DECLARE FUNCTION DestroyData () AS HRESULT
   DECLARE FUNCTION Clear () AS HRESULT
   DECLARE FUNCTION Erase () AS HRESULT
   DECLARE FUNCTION Reset () AS HRESULT
   ' // * Increments the lock count of an array, and retrieves a pointer to the array data.
   DECLARE FUNCTION AccessData () AS ANY PTR
   ' // * Decrements the lock count of an array, and invalidates the pointer retrieved by AccessData.
   DECLARE FUNCTION UnaccessData () AS HRESULT
   ' // * Returns a pointer to an array element.
   DECLARE FUNCTION PtrOfIndex (BYVAL idx AS LONG) AS ANY PTR
   ' // * Changes the right-most (least significant) bound of a safe array.
   DECLARE FUNCTION Redim (BYVAL cElements AS DWORD) AS HRESULT
   ' // * Appends a BSTR to the end of the safe array.
   DECLARE FUNCTION Append (BYREF cws AS CWSTR) AS HRESULT
   ' // * Inserts a new string at the specified position.
   DECLARE FUNCTION InsertItem (BYVAL nPos AS LONG, BYREF cwsText AS CWSTR) AS BOOLEAN
   ' // * Deletes the specified element of the array.
   DECLARE FUNCTION DeleteItem (BYVAL nPos AS LONG) AS BOOLEAN
   ' // * Scans the array to search for the specified string.
   DECLARE FUNCTION FindItem (BYREF cwsItem AS CWSTR, BYVAL bNoCase AS BOOLEAN = FALSE) AS LONG
   ' // * Creates a copy of the safe array.
   DECLARE FUNCTION Copy () AS SAFEARRAY PTR
   ' // * Copies the source array to the target array after releasing any resources in the target array.
   DECLARE FUNCTION CopyData (BYVAL psaTarget AS SAFEARRAY PTR) AS HRESULT
   ' // * Attaches a SAFEARRAY descriptor to a CWstrArray array.
   DECLARE FUNCTION Attach (BYVAL psaSrc AS SAFEARRAY PTR) AS HRESULT
   ' // * Detaches the SAFEARRAY descriptor from the CWstrArray array.
   DECLARE FUNCTION Detach () AS SAFEARRAY PTR
   ' // * Copies the contents of a safe array.
   DECLARE FUNCTION CopyFrom (BYVAL psaSrc AS SAFEARRAY PTR) AS HRESULT
   ' // * Copies the contents of a VARIANT of type VT_ARRAY.
   DECLARE FUNCTION CopyFromVariant (BYVAL pvar AS VARIANT PTR) AS HRESULT
   ' // * Returns a variant containing a copy of the safe array.
   DECLARE FUNCTION CopyToVariant (BYVAL pvar AS VARIANT PTR) AS HRESULT
   ' // * Returns a variant containing a copy of the safe array.
   DECLARE FUNCTION MoveFromVariant (BYVAL pvar AS VARIANT PTR) AS HRESULT
   ' // * Transfers ownership of the safe array to a variant and detaches it from the object.
   DECLARE FUNCTION MoveToVariant (BYVAL pvar AS VARIANT PTR) AS HRESULT
   ' // * Sorts the array calling the C qsort function.
   DECLARE SUB Sort (BYVAL bAscend AS BOOLEAN = TRUE)

END TYPE
' ########################################################################################

' ========================================================================================
' CWstrArray default constructor
' ========================================================================================
CONSTRUCTOR CWstrArray
   CWSTRARRAY_DP("CWstrArray Constructor Default")
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Creates a one-dimensional BSTR safe array.
' - lLBound: The lower bound value; that is, the index of the first element in the array.
'   Can be negative.
' - cElements: The number of elements in the array.
' ========================================================================================
PRIVATE CONSTRUCTOR CWstrArray (BYVAL cElements AS DWORD = 0, BYVAL lLBound AS LONG = 0)
   m_psa = this.Create(cElements, lLBound)
   CWSTRARRAY_DP("CWstrArray Constructor - m_psa = " & WSTR(m_psa) & " - Elements: " & WSTR(cElements) & " - LBound: " & WSTR(lLBound))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Creates a one-dimensional BSTR safe array from another CWstrArray.
' ========================================================================================
PRIVATE CONSTRUCTOR CWstrArray (BYREF cwsa AS CWstrArray)
   CWSTRARRAY_DP("CWstrArray Constructor - cwsa.m_psa = " & WSTR(cwsa.m_psa))
   this.CopyFrom(cwsa.m_psa)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Creates a one-dimensional BSTR safe array from a safe array.
' ========================================================================================
PRIVATE CONSTRUCTOR CWstrArray (BYVAL psaSrc AS SAFEARRAY PTR)
   CWSTRARRAY_DP("CWstrArray Constructor - psaSrc = " & WSTR(psaSrc))
   this.CopyFrom(psaSrc)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Creates a one-dimensional BSTR safe array from a variant.
' ========================================================================================
PRIVATE CONSTRUCTOR CWstrArray (BYVAL pvar AS VARIANT PTR)
   CWSTRARRAY_DP("CWstrArray Constructor - VARIANT ")
   IF pvar <> NULL THEN this.CopyFromVariant(pvar)
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Destroys the safe array when the class is destroyed.
' ========================================================================================
DESTRUCTOR CWstrArray
   CWSTRARRAY_DP("CWstrArray Destructor - " & WSTR(m_psa))
   IF m_psa THEN
      SafeArrayUnlock(m_psa)
      SafeArrayDestroy(m_psa)
   END IF
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' * Returns the lower bound of the safe array.
' Note: SafeArrayGetLBound already checks for null pointers.
' ========================================================================================
PRIVATE FUNCTION CWstrArray.LBound () AS LONG
   CWSTRARRAY_DP("CWstrArray LBound")
   DIM plLBound AS LONG
   SafeArrayGetLBound(m_psa, 1, @plLBound)
   RETURN plLBound
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the upper bound of the safe array.
' Note: SafeArrayGetUBound already checks for null pointers.
' ========================================================================================
PRIVATE FUNCTION CWstrArray.UBound () AS LONG
   CWSTRARRAY_DP("CWstrArray UBound")
   DIM plUBound AS LONG
   SafeArrayGetUBound(m_psa, 1, @plUBound)
   RETURN plUBound
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the number of elements in the the array.
' ========================================================================================
PRIVATE FUNCTION CWstrArray.Count () AS ULONG
   CWSTRARRAY_DP("CWstrArray Count")
   DIM plLBound AS LONG, plUBound AS LONG
   SafeArrayGetLBound(m_psa, 1, @plLBound)
   SafeArrayGetUBound(m_psa, 1, @plUBound)
   RETURN plUbound - plLBound + 1
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the number of times the array has been locked without the corresponding unlock.
' ========================================================================================
PRIVATE FUNCTION CWstrArray.LocksCount () AS ULONG
   CWSTRARRAY_DP("CWstrArray LocksCount")
   IF m_psa THEN RETURN m_psa->cLocks
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Increments the lock count of an array, and retrieves a pointer to the array data.
' Return value: A pointer to the array data.
' Remarks: After calling AccessData, you must call the UnaccessData function to unlock the array.
' ========================================================================================
PRIVATE FUNCTION CWstrArray.AccessData () AS ANY PTR
   CWSTRARRAY_DP("CWstrArray AccessData")
   DIM pvData AS ANY PTR
   SafeArrayAccessData(m_psa, @pvData)
   RETURN pvData
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Decrements the lock count of an array, and invalidates the pointer retrieved by AccessData.
' Return value:
'   S_OK Success.
'   E_UNEXPECTED The array could not be unlocked.
' ========================================================================================
PRIVATE FUNCTION CWstrArray.UnaccessData () AS HRESULT
   CWSTRARRAY_DP("CWstrArray UnaccessData")
   RETURN SafeArrayUnaccessData(m_psa)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns a pointer to an array element.
' Parameters:
' - idx: Index value that identifies an element of the array.
' Return Value:
' Pointer to the array element on success or NULL on failure.
' Failure will happen if the index is invalid.
' ========================================================================================
PRIVATE FUNCTION CWstrArray.PtrOfIndex (BYVAL idx AS LONG) AS ANY PTR
   CWSTRARRAY_DP("CWstrArray PtrOfIndex")
   DIM rgidx AS LONG = idx, pvData AS ANY PTR
   IF m_psa THEN SafeArrayPtrOfIndex(m_psa, @rgidx, @pvData)
   RETURN pvData
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Gets an element of the array. If the function fails, it returns an empty string.
' - idx : The index of the array element.
' ========================================================================================
PRIVATE PROPERTY CWstrArray.Item (BYVAL idx AS LONG) AS CWSTR
   CWSTRARRAY_DP("PROPERTY ITEM [GET] - CWSTR")
   IF m_psa = NULL THEN EXIT PROPERTY
   SafeArrayLock(m_psa)
   DIM pvData AS AFX_BSTR PTR = this.PtrOfIndex(idx)
   IF pvData THEN PROPERTY = *pvData
   SafeArrayUnlock(m_psa)
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CWstrArray.[] (BYVAL idx AS LONG) AS CWSTR
   OPERATOR = this.Item(idx)
END OPERATOR
' ========================================================================================

' ========================================================================================
' * Puts a string element at a given location in the array.
' - idx : The index of the array element.
' - cws : The string data to store.
' ========================================================================================
PRIVATE PROPERTY CWstrArray.Item (BYVAL idx AS LONG, BYREF cws AS CWSTR)
   CWSTRARRAY_DP("PROPERTY ITEM [PUT] - CWSTR")
   IF m_psa = NULL THEN EXIT PROPERTY
   SafeArrayLock(m_psa)
   DIM pvData AS AFX_BSTR PTR = this.PtrOfIndex(idx)
   IF pvData THEN *pvData = SysAllocString(**cws)
   SafeArrayUnlock(m_psa)
END PROPERTY
' ========================================================================================

' ========================================================================================
' * Inserts a new string at the specified position.
' - nPos = Index of the array element where the new string will be inserted.
' - cwsText = The string to insert.
' Return value: TRUE or FALSE.
' ========================================================================================
PRIVATE FUNCTION CWstrArray.InsertItem (BYVAL nPos AS LONG, BYREF cws AS CWSTR) AS BOOLEAN
   CWSTRARRAY_DP("CWstrArray InsertItem")
   DIM cElem AS LONG = nPos - this.LBound
   IF nPos < this.LBound OR nPos > this.UBound THEN RETURN FALSE
   DIM cElements AS LONG = this.UBound - this.LBound + 1
   IF this.Redim(cElements + 1) <> S_OK THEN RETURN FALSE
   DIM pvData AS AFX_BSTR PTR = this.AccessData
   IF pvData THEN
      ' // Move all the elements down
      FOR i AS LONG = cElements TO cElem + 1 STEP - 1
         pvData[i] = pvData[i - 1]
      NEXT
      pvData[cElem] = SysAllocString(**cws)
   END IF
   this.UnaccessData
   RETURN TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Deletes the specified element of the array.
' - nPos = Index of the array element to be removed.
' Return value: TRUE or FALSE.
' ========================================================================================
PRIVATE FUNCTION CWstrArray.DeleteItem (BYVAL nPos AS LONG) AS BOOLEAN
   CWSTRARRAY_DP("CWstrArray DeleteItem")
   DIM cElem AS LONG = nPos - this.LBound
   IF nPos < this.LBound OR nPos > this.UBound THEN RETURN FALSE
   DIM cElements AS LONG = this.UBound - this.LBound + 1
   DIM pvData AS AFX_BSTR PTR = this.AccessData
   IF pvData THEN
      ' // Save the element to be deleted
      DIM pTemp AS AFX_BSTR = pvData[cElem]
      ' // Move all the elements up
      FOR i AS LONG = cElem TO cElements - 1 STEP 1
         pvData[i] = pvData[i + 1]
      NEXT
      ' // Copy the element to be deleted to the end of the array
      pvData[cElements - 1] = pTemp
   END IF
   this.UnaccessData
   ' // Shrink the array by one element (will free the last element)
   IF this.Redim(cElements - 1) = S_OK THEN RETURN TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Scans the array to search for the specified string.
' - cwsItem = The item to search.
' - bNoCase = TRUE = Ignore case.
' ========================================================================================
PRIVATE FUNCTION CWstrArray.FindItem (BYREF cwsItem AS CWSTR, BYVAL bNoCase AS BOOLEAN = FALSE) AS LONG
   CWSTRARRAY_DP("CWstrArray FindItem")
   DIM cwsFind AS CWSTR
   IF bNoCase = FALSE THEN cwsFind = cwsItem ELSE cwsFind = UCASE(cwsItem)
   DIM pvData AS AFX_BSTR PTR = this.AccessData
   IF pvData THEN
      FOR i AS LONG = 0 TO this.UBound - this.LBound
         IF IIF(bNoCase = FALSE, *pvData[i], UCASE(*pvData[i])) = **cwsFind THEN
            FUNCTION = i + this.LBound
            EXIT FOR
         END IF
      NEXT
   END IF
   this.UnaccessData
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Creates a one-dimensional BSTR SafeArray
' Parameters:
' - lLBound
'   The lower bound value; that is, the index of the first element in the array.
'   Can be negative.
' - cElements
'   The number of elements in the array.
' Return value:
'    A safe array descriptor, or null if the array could not be created.
' ========================================================================================
PRIVATE FUNCTION CWstrArray.Create (BYVAL cElements AS DWORD, BYVAL lLBound AS LONG) AS SAFEARRAY PTR
   CWSTRARRAY_DP("CWstrArray Create - " & WSTR(cElements) & "- " & WSTR(lLBound))
   DIM rgsabound AS SAFEARRAYBOUND
   rgsabound.cElements = cElements
   rgsabound.lLBound = lLBound
   RETURN SafeArrayCreate(VT_BSTR, 1, @rgsabound)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Changes the right-most (least significant) bound of a safe array.
' Parameter:
' - cElements: Number of elements of the array.
' Remarks:
'   If you reduce the bound of an array, SafeArrayRedim deallocates the array elements
'   outside the new array boundary. If the bound of an array is increased, SafeArrayRedim
'   allocates and initializes the new array elements. The data is preserved for elements
'   that exist in both the old and new array.
' Return value:
'   S_OK Success.
'   E_FAIL Failure.
'   DISP_E_ARRAYISLOCKED The array is locked.
'   E_POINTER The m_psa pointer is null.
' ========================================================================================
PRIVATE FUNCTION CWstrArray.Redim (BYVAL cElements AS DWORD) AS HRESULT
   CWSTRARRAY_DP("CWstrArray Redim")
   IF m_psa = NULL THEN RETURN E_POINTER
   DIM psaboundNew AS SAFEARRAYBOUND
   psaboundNew.cElements = cElements
   psaboundNew.lLBound = m_psa->rgsabound(0).lLBound
   SafeArrayUnlock(m_psa)
   DIM hr AS HRESULT = SafeArrayRedim(m_psa, @psaboundNew)
   SafeArrayLock(m_psa)
   RETURN hr
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Appends a BSTR to the end of the safe array.
' If the safe array doesn't exist, an empty zero-based one-dimensional array is created.
' Return value:
'   S_OK Success.
'   E_FAIL Failure.
'   E_OUTOFMEMORY Memory could not be allocated for the element.
' ========================================================================================
PRIVATE FUNCTION CWstrArray.Append (BYREF cws AS CWSTR) AS HRESULT
   CWSTRARRAY_DP("CWstrArray Append - CWSTR")
   IF m_psa = NULL THEN
      m_psa = this.Create(1, 0)
      IF m_psa = NULL THEN RETURN E_FAIL
      this.Item(0) = cws
   ELSE
      DIM cElements AS DWORD = this.Count + 1
      DIM hr AS HRESULT = this.Redim(cElements)
      IF hr <> S_OK THEN RETURN hr
      DIM idx AS LONG = cElements - 1 + this.LBound
      this.Item(idx) = cws
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Creates a copy of the safe array.
' Return value:
' Pointer of the new array descriptor. You must free this pointer calling the API
' function SafeArrayDestroy.
' ========================================================================================
PRIVATE FUNCTION CWstrArray.Copy () AS SAFEARRAY PTR
   CWSTRARRAY_DP("CWstrArray Copy")
   DIM psaOut AS SAFEARRAY PTR
   IF m_psa THEN SafeArrayCopy(m_psa, @psaOut)
   RETURN psaOut
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Copies the source array to the target array after releasing any resources in the
' target array. This is similar to SafeArrayCopy, except that the target array has
' to be set up by the caller. The target is not allocated or reallocated.
' Parameter:
' - psaTarget
'   The target safe array. On exit, the array referred to by psaTarget contains a
'   copy of the data in m_psa.
' Return value:
'   S_OK Success.
'   E_FAIL Failure.
'   E_INVALIDARG The argument psaTarget was not valid.
'   E_OUTOFMEMORY Insufficient memory to create the copy.
' ========================================================================================
PRIVATE FUNCTION CWstrArray.CopyData (BYVAL psaTarget AS SAFEARRAY PTR) AS HRESULT
   CWSTRARRAY_DP("CWstrArray CopyData")
   RETURN SafeArrayCopyData(m_psa, psaTarget)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Destroys an existing array descriptor and all of the data in the array.
' SysFreeString will be called on each element.
' Return value:
' S_OK Success.
' DISP_E_ARRAYISLOCKED The array is locked.
' ========================================================================================
PRIVATE FUNCTION CWstrArray.Destroy () AS HRESULT
   CWSTRARRAY_DP("CWstrArray Destroy - m_psa = " & WSTR(m_psa))
   SafeArrayUnlock(m_psa)
   DIM hr AS HRESULT = SafeArrayDestroy(m_psa)
   IF SUCCEEDED(hr) THEN m_psa = NULL
   RETURN hr
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Destroys all the data in a safe array.
' SysFreeString will be called on each element.
' Returns S_OK on success, or an error HRESULT on failure.
' E_INVALIDARG: The argument m_psa is not valid.
' DISP_E_ARRAYISLOCKED: The array is locked.
' ========================================================================================
PRIVATE FUNCTION CWstrArray.DestroyData () AS HRESULT
   CWSTRARRAY_DP("CWstrArray DestroyData")
   SafeArrayUnlock(m_psa)
   DIM hr AS HRESULT = SafeArrayDestroyData(m_psa)
   SafeArrayLock(m_psa)
   RETURN hr
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CWstrArray.Clear () AS HRESULT
   RETURN this.DestroyData
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CWstrArray.Erase () AS HRESULT
   RETURN this.DestroyData
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CWstrArray.Reset () AS HRESULT
   RETURN this.DestroyData
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Attaches a SAFEARRAY descriptor to a CWstrArray array.
' Parameter:
' - psaSrc: A pointer to the SAFEARRAY descriptor.
' Returns S_OK on success, or an error HRESULT on failure.
' Remarks: Never attach the same safe array to more than one CWstrArray class because each one
' will try to destroy it when they are deleted or went out of scope.
' ========================================================================================
PRIVATE FUNCTION CWstrArray.Attach (BYVAL psaSrc AS SAFEARRAY PTR) AS HRESULT
   CWSTRARRAY_DP("CWstrArray Attach")
   IF psaSrc = NULL THEN RETURN E_FAIL
   DIM vt AS VARTYPE
   SafeArrayGetVartype(psaSrc, @vt)
   IF vt <> VT_BSTR THEN RETURN E_INVALIDARG
   DIM hr AS HRESULT = this.Destroy
   IF SUCCEEDED(hr) THEN
      m_psa = psaSrc
      hr = SafeArrayLock(m_psa)
   END IF
   RETURN hr
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Detaches the SAFEARRAY descriptor from the CWstrArray array.
' Return value: Returns a pointer to a SAFEARRAY descriptor.
' The caller takes ownership of it and must destroy it when no longer needed.
' ========================================================================================
PRIVATE FUNCTION CWstrArray.Detach () AS SAFEARRAY PTR
   CWSTRARRAY_DP("CWstrArray Detach")
   SafeArrayUnlock(m_psa)
   FUNCTION = m_psa
   m_psa = NULL
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the address of the safe array descriptor.
' ========================================================================================
OPERATOR CWstrArray.@ () AS ANY PTR
   CWSTRARRAY_DP("CWstrArray Operator @")
   OPERATOR = @m_psa
END OPERATOR
' ========================================================================================
' ========================================================================================
' * Returns a pointer to the safe array descriptor.
' ========================================================================================
PRIVATE FUNCTION CWstrArray.GetPtr () AS SAFEARRAY PTR
   CWSTRARRAY_DP("CWstrArray GetPtr")
   RETURN m_psa
END FUNCTION
' ========================================================================================
' ========================================================================================
' Returns a pointer to the safe array descriptor.
' ========================================================================================
OPERATOR * (BYREF cwsa AS CWstrArray) AS SAFEARRAY PTR
   CWSTRARRAY_DP("CWstrArray Operator *")
   OPERATOR = cwsa.m_psa
END OPERATOR
' ========================================================================================

' ========================================================================================
' * Assigns a CWstrArray.
' ========================================================================================
OPERATOR CWstrArray.Let (BYREF cwsa AS CWstrArray)
   CWSTR_DP("CWstrArray LET CWstrArray")
   this.CopyFrom(cwsa.m_psa)
END OPERATOR
' ========================================================================================
' ========================================================================================
' * Assigns a safe array.
' ========================================================================================
OPERATOR CWstrArray.Let (BYVAL psa AS SAFEARRAY PTR)
   CWSTR_DP("CWstrArray LET SAFEARRAY PTR")
   this.CopyFrom(psa)
END OPERATOR
' ========================================================================================

' ========================================================================================
' * Copies the contents of a safe array.
' Parameter:
' - psaSrc: Pointer to an array descriptor created by SafeArrayCreate.
' Return value:
' Returns S_OK on success, or an HRESULT on failure.
' Remarks:
' SafeArrayPutElement, automatically calls SafeArrayLock and SafeArrayUnlock before and
' after assigning the element.
' ========================================================================================
PRIVATE FUNCTION CWstrArray.CopyFrom (BYVAL psaSrc AS SAFEARRAY PTR) AS HRESULT
   CWSTRARRAY_DP("CWstrArray CopyFrom - safe array ptr = " & WSTR(psaSrc))
   IF psaSrc = NULL THEN RETURN E_INVALIDARG
   ' // Get the safe array type and return an error if it is not a BSTR safe array
   DIM vt AS VARTYPE
   SafeArrayGetVartype(psaSrc, @vt)
   IF vt <> VT_BSTR THEN RETURN E_INVALIDARG
   ' // Return an error if it is not a one-dimensional array
   IF SafeArrayGetDim(psaSrc) <> 1 THEN RETURN E_INVALIDARG
   ' // Destroy the safe array
   IF m_psa THEN this.Destroy
   ' // Copy the passed safe array and lock it
   DIM hr AS HRESULT = SafeArrayCopy(psaSrc, @m_psa)
   SafeArrayLock(m_psa)
   CWSTRARRAY_DP("CWstrArray CopyFrom - m_psa = " & WSTR(m_psa))
   RETURN hr
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Copies the contents of a VARIANT of type VT_ARRAY, i.e. containing a safe array, to
' the CWstrArray array. The VARIANT remains unaltered.
' ========================================================================================
PRIVATE FUNCTION CWstrArray.CopyFromVariant (BYVAL pvar AS VARIANT PTR) AS HRESULT
   CWSTRARRAY_DP("CWstrArray CopyFromVariant")
   IF pvar = NULL THEN RETURN E_INVALIDARG
   IF (pvar->vt AND VT_ARRAY) <> VT_ARRAY THEN RETURN HRESULT_FROM_WIN32(ERROR_INVALID_DATA)
   IF (pvar->vt AND VT_BSTR) <> VT_BSTR THEN RETURN HRESULT_FROM_WIN32(ERROR_INVALID_DATA)
   this.CopyFrom(pvar->parray)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns a variant containing a copy of the safe array.
' ========================================================================================
PRIVATE FUNCTION CWstrArray.CopyToVariant (BYVAL pvar AS VARIANT PTR) AS HRESULT
   CWSTRARRAY_DP("CWstrArray CopyToVariant")
   IF m_psa = NULL THEN RETURN E_FAIL
   IF pvar = NULL THEN RETURN E_INVALIDARG
   VariantClear(pvar)
   pvar->vt = VT_ARRAY OR VT_BSTR
   RETURN SafeArrayCopy(m_psa, @pvar->parray)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Transfers ownership of the safe array contained in the variant parameter to this
' object. The variant is then changed to VT_EMPTY.
' ========================================================================================
PRIVATE FUNCTION CWstrArray.MoveFromVariant (BYVAL pvar AS VARIANT PTR) AS HRESULT
   CWSTRARRAY_DP("CWstrArray MoveFromVariant")
   IF (pvar->vt AND VT_ARRAY) <> VT_ARRAY THEN RETURN HRESULT_FROM_WIN32(ERROR_INVALID_DATA)
   IF (pvar->vt AND VT_BSTR) <> VT_BSTR THEN RETURN HRESULT_FROM_WIN32(ERROR_INVALID_DATA)
   this.CopyFrom(pvar->parray)
   pvar->vt = VT_EMPTY
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Transfers ownership of the safe array to a variant and detaches it from the object.
' ========================================================================================
PRIVATE FUNCTION CWstrArray.MoveToVariant (BYVAL pvar AS VARIANT PTR) AS HRESULT
   CWSTRARRAY_DP("CWstrArray MoveToVariant")
   IF m_psa = NULL THEN RETURN E_FAIL
   IF pvar = NULL THEN RETURN E_INVALIDARG
   VariantClear(pvar)
   FUNCTION = SafeArrayUnlock(m_psa)
   pvar->vt = VT_ARRAY OR VT_BSTR
   pvar->parray = m_psa
   m_psa = NULL
END FUNCTION
' ========================================================================================

' ========================================================================================
' qsort CWstrArray comparison function
' ========================================================================================
PRIVATE FUNCTION AfxCWstrArrayCompare CDECL (BYVAL a AS AFX_BSTR PTR, BYVAL b AS AFX_BSTR PTR) AS LONG
   FUNCTION = wcscmp(*a, *b)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Reverse qsort CWstrArray comparison function
' ========================================================================================
PRIVATE FUNCTION AfxCWstrArrayReverseCompare CDECL (BYVAL a AS AFX_BSTR PTR, BYVAL b AS AFX_BSTR PTR) AS LONG
   DIM r AS LONG = wcscmp(*a, *b)
   IF r = 1 THEN r = -1 ELSE IF r = -1 THEN r = 1
   RETURN r
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sorts the array calling the C qsort function.
' Parameter:
' - bAscend: TRUE for sorting in ascending order; FALSE for sorting in descending order.
' ========================================================================================
PRIVATE SUB CWstrArray.Sort (BYVAL bAscend AS BOOLEAN = TRUE)
   CWSTRARRAY_DP("CWstrArray Sort")
   DIM pvData AS AFX_BSTR PTR = this.AccessData
   IF bAscend THEN
      qsort pvData, this.Count, SIZEOF(AFX_BSTR), CPTR(ANY PTR, @AfxCWstrArrayCompare)
   ELSE
      qsort pvData, this.Count, SIZEOF(AFX_BSTR), CPTR(ANY PTR, @AfxCWstrArrayReverseCompare)
   END IF
   this.UnaccessData
END SUB
' ========================================================================================

END NAMESPACE
