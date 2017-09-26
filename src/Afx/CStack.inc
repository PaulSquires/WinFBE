' ########################################################################################
' Microsoft Windows
' File: CStack.inc
' Contents: Stack and Queue collections
' Compiler: FreeBasic 32 & 64-bit
' (c) 2017 by José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "Afx/CSafeArray.inc"
using Afx

NAMESPACE Afx

' ========================================================================================
' CStack class
' A Stack Collection is an ordered set of data items, which are accessed on a LIFO
' (Last-In / First-Out) basis. Each data item is passed and stored as a variant variable,
' using the Push and Pop methods.
' ========================================================================================
TYPE CStack

Private:
   DIM m_psa AS CSafeArray PTR

Public:
   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR
   DECLARE FUNCTION Push (BYREF cvData AS CVAR) AS HRESULT
   DECLARE FUNCTION Pop () AS CVAR
   DECLARE FUNCTION Count () AS UINT
   DECLARE FUNCTION Clear () AS HRESULT

END TYPE
' ========================================================================================

' ========================================================================================
' CStack constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CStack
   ' // Create a safe array of 0 elements and a lower bound of 1
   m_psa = NEW CSafeArray(VT_VARIANT, 0, 1)
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CStack destructor
' ========================================================================================
PRIVATE DESTRUCTOR CStack
   Delete m_psa
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Appends a variant at the end of the array.
' ========================================================================================
PRIVATE FUNCTION CStack.Push (BYREF cvData AS CVAR) AS HRESULT
   RETURN m_psa->AppendElement(cvData)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets and removes the last element of the array.
' ========================================================================================
PRIVATE FUNCTION CStack.Pop () AS CVAR
   DIM cv AS CVAR
   DIM nPos AS UINT = this.Count
   IF nPos = 0 THEN RETURN cv
   cv = m_psa->GetVar(nPos)
   m_psa->DeleteVariantElement(nPos)
   RETURN cv
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the number of elements in the collection.
' ========================================================================================
PRIVATE FUNCTION CStack.Count () AS UINT
   RETURN m_psa->Count
END FUNCTION
' ========================================================================================

' ========================================================================================
' Removes all the elements in the collection.
' ========================================================================================
PRIVATE FUNCTION CStack.Clear () AS HRESULT
   RETURN m_psa->Reset
END FUNCTION
' ========================================================================================

' ========================================================================================
' CQueue class
' A Queue Collection is an ordered set of data items, which are accessed on a FIFO
' (First-In / First-Out) basis. Each data item is passed and stored as a variant variable,
' using the Enqueue and Dequeue methods.
' ========================================================================================
TYPE CQueue

Private:
   DIM m_psa AS CSafeArray PTR

Public:
   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR
   DECLARE FUNCTION Enqueue (BYREF cvData AS CVAR) AS HRESULT
   DECLARE FUNCTION Dequeue () AS CVAR
   DECLARE FUNCTION Count () AS UINT
   DECLARE FUNCTION Clear () AS HRESULT

END TYPE
' ========================================================================================

' ========================================================================================
' CQueue constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CQueue
   ' // Create a safe array of 0 elements and a lower bound of 1
   m_psa = NEW CSafeArray(VT_VARIANT, 0, 1)
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CQueue destructor
' ========================================================================================
PRIVATE DESTRUCTOR CQueue
   Delete m_psa
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Appends a variant at the end of the array.
' ========================================================================================
PRIVATE FUNCTION CQueue.Enqueue (BYREF cvData AS CVAR) AS HRESULT
   RETURN m_psa->AppendElement(cvData)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets and removes the first element of the array.
' ========================================================================================
PRIVATE FUNCTION CQueue.Dequeue () AS CVAR
   DIM cv AS CVAR
   IF this.Count = 0 THEN RETURN cv
   cv = m_psa->GetVar(1)
   m_psa->DeleteVariantElement(1)
   RETURN cv
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the number of elements in the collection.
' ========================================================================================
PRIVATE FUNCTION CQueue.Count () AS UINT
   RETURN m_psa->Count
END FUNCTION
' ========================================================================================

' ========================================================================================
' Removes all the elements in the collection.
' ========================================================================================
PRIVATE FUNCTION CQueue.Clear () AS HRESULT
   RETURN m_psa->Reset
END FUNCTION
' ========================================================================================

END NAMESPACE
