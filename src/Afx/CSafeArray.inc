' ########################################################################################
' Microsoft Windows
' File: CSafeArray.inc
' Contents: Safe array class.
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2016-2017 José Roca. Freeware. Use at your own risk.é
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "windows.bi"
#include once "win/ole2.bi"
#include once "Afx/AfxWin.inc"
#include once "Afx/CVAR.inc"
#include once "Afx/CWSTR.inc"

NAMESPACE Afx

' ========================================================================================
' Macro for debug
' To allow debugging, define _CSAFEARRAY_DEBUG_ 1 in your application before including this file.
' ========================================================================================
#ifndef _CSAFEARRAY_DEBUG_
   #define _CSAFEARRAY_DEBUG_ 0
#ENDIF
#ifndef _CSAFEARRAY_DP_
   #define _CSAFEARRAY_DP_ 1
   #MACRO CSAFEARRAY_DP(st)
      #IF (_CSAFEARRAY_DEBUG_ = 1)
         OutputDebugStringW(st)
      #ENDIF
   #ENDMACRO
#ENDIF
' ========================================================================================

' ########################################################################################
' CSafeArray - Safe array class
' ########################################################################################
TYPE CSafeArray

   m_psa AS SAFEARRAY PTR

   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYVAL vt AS VARTYPE, BYVAL cDims AS UINT, BYVAL prgsabounds AS SAFEARRAYBOUND PTR)
   DECLARE CONSTRUCTOR (BYVAL vt AS VARTYPE, BYVAL cElements AS ULONG = 0, BYVAL lLBound AS LONG = 0)
   DECLARE CONSTRUCTOR (BYVAL vt AS VARTYPE, BYVAL cElements1 AS ULONG, BYVAL lLBound1 AS LONG, BYVAL cElements2 AS ULONG, BYVAL lLBound2 AS LONG)
   DECLARE CONSTRUCTOR (BYREF strType AS STRING, BYVAL cDims AS UINT, BYVAL prgsabounds AS SAFEARRAYBOUND PTR)
   DECLARE CONSTRUCTOR (BYREF strType AS STRING, BYVAL cElements AS ULONG = 0, BYVAL lLBound AS LONG = 0)
   DECLARE CONSTRUCTOR (BYREF strType AS STRING, BYVAL cElements1 AS ULONG, BYVAL lLBound1 AS LONG, BYVAL cElements2 AS ULONG, BYVAL lLBound2 AS LONG)
   DECLARE CONSTRUCTOR (BYREF csa AS CSafeArray)
   DECLARE CONSTRUCTOR (BYVAL psa AS SAFEARRAY PTR)
   DECLARE CONSTRUCTOR (BYVAL psa AS SAFEARRAY PTR, BYVAL fAttach AS BOOLEAN)
   DECLARE CONSTRUCTOR (BYVAL pvar AS VARIANT PTR)
   DECLARE DESTRUCTOR
   DECLARE OPERATOR Let (BYREF csa AS CSafeArray)
   DECLARE OPERATOR Let (BYVAL psa AS SAFEARRAY PTR)
   DECLARE OPERATOR Let (BYVAL pvar AS VARIANT PTR)
   DECLARE FUNCTION GetPtr () AS SAFEARRAY PTR
'   DECLARE OPERATOR @ () AS ANY PTR
'   DECLARE FUNCTION vptr () AS SAFEARRAY PTR
   DECLARE FUNCTION Create (BYVAL vt AS VARTYPE, BYVAL cDims AS UINT, BYVAL prgsabound AS SAFEARRAYBOUND PTR) AS HRESULT
   DECLARE FUNCTION Create (BYVAL vt AS VARTYPE, BYVAL cElements AS ULONG, BYVAL lLBound AS LONG) AS HRESULT
   DECLARE FUNCTION Create (BYVAL vt AS VARTYPE, BYVAL cElements1 AS ULONG, BYVAL lLBound1 AS LONG, BYVAL cElements2 AS ULONG, BYVAL lLBound2 AS LONG) AS HRESULT
   DECLARE FUNCTION CreateEx (BYVAL vt AS VARTYPE, BYVAL cDims AS UINT, BYVAL prgsabound AS SAFEARRAYBOUND PTR, BYVAL pvExtra AS PVOID) AS HRESULT
   DECLARE FUNCTION CreateEx (BYVAL vt AS VARTYPE, BYVAL cElements AS ULONG, BYVAL lLBound AS LONG, BYVAL pvExtra AS ANY PTR) AS HRESULT
   DECLARE FUNCTION CreateEx (BYVAL vt AS VARTYPE, BYVAL cElements1 AS ULONG, BYVAL lLBound1 AS LONG, BYVAL cElements2 AS ULONG, BYVAL lLBound2 AS LONG, BYVAL pvExtra AS ANY PTR) AS HRESULT
   DECLARE FUNCTION CreateVector (BYVAL vt AS VARTYPE, BYVAL cElements AS ULONG, BYVAL lLBound AS LONG) AS HRESULT
   DECLARE FUNCTION CreateVectorEx (BYVAL vt AS VARTYPE, BYVAL cElements AS ULONG, BYVAL lLBound AS LONG, BYVAL pvExtra AS ANY PTR) AS HRESULT
   DECLARE FUNCTION NumDims () AS UINT
   DECLARE FUNCTION LBound (BYVAL nDim AS UINT = 1) AS LONG
   DECLARE FUNCTION UBound (BYVAL nDim AS UINT = 1) AS LONG
   DECLARE FUNCTION Count (BYVAL nDim AS UINT = 1) AS DWORD
   DECLARE FUNCTION ElemSize () AS UINT
   DECLARE FUNCTION IsResizable () AS BOOLEAN
   DECLARE FUNCTION GetType () AS VARTYPE
   DECLARE FUNCTION LocksCount () AS UINT
   DECLARE FUNCTION Flags () AS USHORT
   DECLARE FUNCTION Features () AS USHORT
   DECLARE FUNCTION AccessData () AS ANY PTR
   DECLARE FUNCTION UnaccessData () AS HRESULT
   DECLARE FUNCTION Redim (BYVAL pnewsabounds AS SAFEARRAYBOUND PTR) AS HRESULT
   DECLARE FUNCTION Redim (BYVAL cElements AS DWORD, BYVAL lLBound AS LONG) AS HRESULT
   DECLARE FUNCTION Redim (BYVAL cElements1 AS ULONG, BYVAL lLBound1 AS LONG, BYVAL cElements2 AS ULONG, BYVAL lLBound2 AS LONG) AS HRESULT
   DECLARE FUNCTION Copy () AS SAFEARRAY PTR
   DECLARE FUNCTION CopyData (BYVAL psaTarget AS SAFEARRAY PTR) AS HRESULT
   DECLARE FUNCTION CopyFrom (BYVAL psaSrc AS SAFEARRAY PTR) AS HRESULT
   DECLARE FUNCTION PtrOfIndex (BYVAL prgIndices AS LONG PTR) AS ANY PTR
   DECLARE FUNCTION PtrOfIndex (BYVAL idx AS LONG) AS ANY PTR
   DECLARE FUNCTION PtrOfIndex (BYVAL cElem AS LONG, BYVAL cDim AS LONG) AS ANY PTR
   DECLARE FUNCTION Destroy () AS HRESULT
   DECLARE FUNCTION DestroyData () AS HRESULT
   DECLARE FUNCTION Clear () AS HRESULT
   DECLARE FUNCTION Erase () AS HRESULT
   DECLARE FUNCTION Reset () AS HRESULT
   DECLARE FUNCTION CopyFromVariant (BYVAL pvar AS VARIANT PTR) AS HRESULT
   DECLARE FUNCTION MoveFromVariant (BYVAL pvar AS VARIANT PTR) AS HRESULT
   DECLARE FUNCTION CopyToVariant (BYVAL pvar AS VARIANT PTR) AS HRESULT
   DECLARE FUNCTION MoveToVariant (BYVAL pvar AS VARIANT PTR) AS HRESULT
   DECLARE FUNCTION Attach (BYVAL psaSrc AS SAFEARRAY PTR) AS HRESULT
   DECLARE FUNCTION Detach () AS SAFEARRAY PTR
   DECLARE FUNCTION GetIID () AS GUID
   DECLARE FUNCTION SetIID (BYVAL pguid AS GUID PTR) AS HRESULT
   DECLARE FUNCTION GetRecordInfo () AS IRecordInfo PTR
   DECLARE FUNCTION SetRecordInfo (BYVAL prinfo AS IRecordInfo PTR) AS HRESULT
   DECLARE FUNCTION AppendElement (BYVAL pData AS ANY PTR) AS HRESULT
   DECLARE FUNCTION AppendElement (BYREF cbsData AS CBSTR) AS HRESULT
   DECLARE FUNCTION AppendElement (BYREF cvData AS CVAR) AS HRESULT
   DECLARE FUNCTION AppendElement (BYVAL vData AS VARIANT) AS HRESULT
   DECLARE FUNCTION Append (BYVAL pData AS ANY PTR) AS HRESULT
   DECLARE FUNCTION Append (BYREF cbsData AS CBSTR) AS HRESULT
   DECLARE FUNCTION Append (BYREF cvData AS CVAR) AS HRESULT
   DECLARE FUNCTION Append (BYVAL vData AS VARIANT) AS HRESULT
   DECLARE FUNCTION AppendStr (BYVAL pwszData AS WSTRING PTR) AS HRESULT
   DECLARE FUNCTION InsertElement (BYVAL nPos AS LONG, BYVAL pData AS ANY PTR) AS HRESULT
   DECLARE FUNCTION InsertElement (BYVAL pData AS ANY PTR) AS HRESULT
   DECLARE FUNCTION InsertElement (BYVAL nPos AS LONG, BYREF cbsData AS CBSTR) AS HRESULT
   DECLARE FUNCTION InsertElement (BYREF cbsData AS CBSTR) AS HRESULT
   DECLARE FUNCTION InsertElement (BYVAL nPos AS LONG, BYREF cvData AS CVAR) AS HRESULT
   DECLARE FUNCTION InsertElement (BYVAL nPos AS LONG, BYVAL vData AS VARIANT) AS HRESULT
   DECLARE FUNCTION InsertElement (BYREF cvData AS CVAR) AS HRESULT
   DECLARE FUNCTION InsertElement (BYVAL vData AS VARIANT) AS HRESULT
   DECLARE FUNCTION Insert (BYVAL nPos AS LONG, BYVAL pData AS ANY PTR) AS HRESULT
   DECLARE FUNCTION Insert (BYVAL pData AS ANY PTR) AS HRESULT
   DECLARE FUNCTION Insert (BYVAL nPos AS LONG, BYREF cbsData AS CBSTR) AS HRESULT
   DECLARE FUNCTION Insert (BYREF cbsData AS CBSTR) AS HRESULT
   DECLARE FUNCTION Insert (BYVAL nPos AS LONG, BYREF cvData AS CVAR) AS HRESULT
   DECLARE FUNCTION Insert (BYVAL nPos AS LONG, BYVAL vData AS VARIANT) AS HRESULT
   DECLARE FUNCTION Insert (BYREF cvData AS CVAR) AS HRESULT
   DECLARE FUNCTION Insert (BYVAL vData AS VARIANT) AS HRESULT
   DECLARE FUNCTION InsertStr (BYVAL nPos AS LONG, BYVAL pwszData AS WSTRING PTR) AS HRESULT
   DECLARE FUNCTION InsertStr (BYVAL pwszData AS WSTRING PTR) AS HRESULT
   DECLARE FUNCTION DeleteElement (BYVAL nPos AS LONG) AS HRESULT
   DECLARE FUNCTION Remove (BYVAL nPos AS LONG) AS HRESULT
   DECLARE FUNCTION DeleteStringElement (BYVAL nPos AS LONG) AS HRESULT
   DECLARE FUNCTION DeleteVariantElement (BYVAL nPos AS LONG) AS HRESULT
   DECLARE FUNCTION RemoveStr (BYVAL nPos AS LONG) AS HRESULT
   DECLARE FUNCTION RemoveVar (BYVAL nPos AS LONG) AS HRESULT
   DECLARE FUNCTION FindElement (BYREF wszFind AS WSTRING, BYVAL bNoCase AS BOOLEAN = FALSE) AS LONG
   DECLARE FUNCTION Find (BYREF wszFind AS WSTRING, BYVAL bNoCase AS BOOLEAN = FALSE) AS LONG
   DECLARE FUNCTION Sort (BYVAL bAscend AS BOOLEAN = TRUE) AS HRESULT
   ' // ANY PTR
   DECLARE FUNCTION GetElement (BYVAL prgIndices AS LONG PTR, BYVAL pData AS ANY PTR) AS HRESULT
   DECLARE FUNCTION GetElement (BYVAL idx AS LONG, BYVAL pData AS ANY PTR) AS HRESULT
   DECLARE FUNCTION GetElement (BYVAL cElem AS LONG, BYVAL cDim AS LONG, BYVAL pData AS ANY PTR) AS HRESULT
   DECLARE FUNCTION Get (BYVAL prgIndices AS LONG PTR, BYVAL pData AS ANY PTR) AS HRESULT
   DECLARE FUNCTION Get (BYVAL idx AS LONG, BYVAL pData AS ANY PTR) AS HRESULT
   DECLARE FUNCTION Get (BYVAL cElem AS LONG, BYVAL cDim AS LONG, BYVAL pData AS ANY PTR) AS HRESULT
   ' // BYREF CBSTR
   DECLARE FUNCTION GetElement (BYVAL prgIndices AS LONG PTR, BYREF cbsData AS CBSTR) AS HRESULT
   DECLARE FUNCTION GetElement (BYVAL idx AS LONG, BYREF cbsData AS CBSTR) AS HRESULT
   DECLARE FUNCTION GetElement (BYVAL cElem AS LONG, BYVAL cDim AS LONG, BYREF cbsData AS CBSTR) AS HRESULT
   DECLARE FUNCTION Get (BYVAL prgIndices AS LONG PTR, BYREF cbsData AS CBSTR) AS HRESULT
   DECLARE FUNCTION Get (BYVAL idx AS LONG, BYREF cbsData AS CBSTR) AS HRESULT
   DECLARE FUNCTION Get (BYVAL cElem AS LONG, BYVAL cDim AS LONG, BYREF cbsData AS CBSTR) AS HRESULT
   DECLARE FUNCTION GetStr (BYVAL prgIndices AS LONG PTR) AS CBSTR
   DECLARE FUNCTION GetStr (BYVAL idx AS LONG) AS CBSTR
   DECLARE FUNCTION GetStr (BYVAL cElem AS LONG, BYVAL cDim AS LONG) AS CBSTR
   ' // BYREF CVAR
   DECLARE FUNCTION GetElement (BYVAL prgIndices AS LONG PTR, BYREF cvData AS CVAR) AS HRESULT
   DECLARE FUNCTION GetElement (BYVAL idx AS LONG, BYREF cvData AS CVAR) AS HRESULT
   DECLARE FUNCTION GetElement (BYVAL cElem AS LONG, BYVAL cDim AS LONG, BYREF cvData AS CVAR) AS HRESULT
   DECLARE FUNCTION Get (BYVAL prgIndices AS LONG PTR, BYREF cvData AS CVAR) AS HRESULT
   DECLARE FUNCTION Get (BYVAL idx AS LONG, BYREF cvData AS CVAR) AS HRESULT
   DECLARE FUNCTION Get (BYVAL cElem AS LONG, BYVAL cDim AS LONG, BYREF cvData AS CVAR) AS HRESULT
   DECLARE FUNCTION GetVar (BYVAL prgIndices AS LONG PTR) AS CVAR
   DECLARE FUNCTION GetVar (BYVAL idx AS LONG) AS CVAR
   DECLARE FUNCTION GetVar (BYVAL cElem AS LONG, BYVAL cDim AS LONG) AS CVAR
   ' // ANY PTR
   DECLARE FUNCTION PutElement (BYVAL prgIndices AS LONG PTR, BYVAL pData AS ANY PTR) AS HRESULT
   DECLARE FUNCTION PutElement (BYVAL idx AS LONG, BYVAL pData AS ANY PTR) AS HRESULT
   DECLARE FUNCTION PutElement (BYVAL cElem AS LONG, BYVAL cDim AS LONG, BYVAL pData AS ANY PTR) AS HRESULT
   DECLARE FUNCTION Put (BYVAL prgIndices AS LONG PTR, BYVAL pData AS ANY PTR) AS HRESULT
   DECLARE FUNCTION Put (BYVAL idx AS LONG, BYVAL pData AS ANY PTR) AS HRESULT
   DECLARE FUNCTION Put (BYVAL cElem AS LONG, BYVAL cDim AS LONG, BYVAL pData AS ANY PTR) AS HRESULT
   ' // BYREF CBSTR
   DECLARE FUNCTION PutElement (BYVAL prgIndices AS LONG PTR, BYREF cbsData AS CBSTR) AS HRESULT
   DECLARE FUNCTION PutElement (BYVAL idx AS LONG, BYREF cbsData AS CBSTR) AS HRESULT
   DECLARE FUNCTION PutElement (BYVAL cElem AS LONG, BYVAL cDim AS LONG, BYREF cbsData AS CBSTR) AS HRESULT
   DECLARE FUNCTION PutStr (BYVAL prgIndices AS LONG PTR, BYREF cbsData AS CBSTR) AS HRESULT
   DECLARE FUNCTION PutStr (BYVAL idx AS LONG, BYREF cbsData AS CBSTR) AS HRESULT
   DECLARE FUNCTION PutStr (BYVAL cElem AS LONG, BYVAL cDim AS LONG, BYREF cbsData AS CBSTR) AS HRESULT
   ' // BYREF CVAR
   DECLARE FUNCTION PutElement (BYVAL prgIndices AS LONG PTR, BYREF cvData AS CVAR) AS HRESULT
   DECLARE FUNCTION PutElement (BYVAL idx AS LONG, BYREF cvData AS CVAR) AS HRESULT
   DECLARE FUNCTION PutElement (BYVAL cElem AS LONG, BYVAL cDim AS LONG, BYREF cvData AS CVAR) AS HRESULT
   DECLARE FUNCTION PutVar (BYVAL prgIndices AS LONG PTR, BYREF cvData AS CVAR) AS HRESULT
   DECLARE FUNCTION PutVar (BYVAL idx AS LONG, BYREF cvData AS CVAR) AS HRESULT
   DECLARE FUNCTION PutVar (BYVAL cElem AS LONG, BYVAL cDim AS LONG, BYREF cvData AS CVAR) AS HRESULT

END TYPE
' ########################################################################################

' ========================================================================================
' CSafeArray default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CSafeArray
   CSAFEARRAY_DP("CSafeArray Constructor Default")
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Creates a safe array.
' Parameters:
' - vtType: The base type of the array (the VARTYPE of each element of the array). The
'   VARTYPE is restricted to a subset of the variant types. Neither the VT_ARRAY nor the
'   VT_BYREF flag can be set. VT_EMPTY and VT_NULL are not valid base types for the array.
'   All other types are legal.
' - cDims: The number of dimensions in the array. The number cannot be changed after the
'   array is created.
' - lLBound: The lower bound value; that is, the index of the first element in the array.
'   Can be negative.
' - cElements: The number of elements in the array.
' ========================================================================================
' // Multidimensional array
PRIVATE CONSTRUCTOR CSafeArray (BYVAL vt AS VARTYPE, BYVAL cDims AS UINT, BYVAL prgsabounds AS SAFEARRAYBOUND PTR)
   m_psa = SafeArrayCreate(vt, cDims, prgsabounds)
   CSAFEARRAY_DP("CSafeArray Constructor multidim " & WSTR(m_psa))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' // One-dimensional array
PRIVATE CONSTRUCTOR CSafeArray (BYVAL vt AS VARTYPE, BYVAL cElements AS ULONG = 0, BYVAL lLBound AS LONG = 0)
   DIM rgsabounds(0) AS SAFEARRAYBOUND = {(cElements, lLBound)}
   m_psa = SafeArrayCreate(vt, 1, @rgsabounds(0))
   CSAFEARRAY_DP("CSafeArray Constructor one-dimensional " & WSTR(m_psa))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' // Two-dimensional array
PRIVATE CONSTRUCTOR CSafeArray (BYVAL vt AS VARTYPE, BYVAL cElements1 AS ULONG, BYVAL lLBound1 AS LONG, BYVAL cElements2 AS ULONG, BYVAL lLBound2 AS LONG)
   DIM rgsabounds(1) AS SAFEARRAYBOUND = {(cElements1, lLBound1), (cElements2, lLBound2)}
   m_psa = SafeArrayCreate(vt, 2, @rgsabounds(0))
   CSAFEARRAY_DP("CSafeArray Constructor two-dimensional " & WSTR(m_psa))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' // Multidimensional array
PRIVATE CONSTRUCTOR CSafeArray (BYREF strType AS STRING, BYVAL cDims AS UINT, BYVAL prgsabounds AS SAFEARRAYBOUND PTR)
   DIM vt AS WORD
   SELECT CASE UCASE(strType)
      CASE "STRING"      : vt = VT_BSTR
      CASE "BSTR"        : vt = VT_BSTR
      CASE "BOOL"        : vt = VT_BOOL
      CASE "BOOLEAN"     : vt = VT_BOOL
      CASE "BYTE"        : vt = VT_I1
      CASE "UBYTE"       : vt = VT_UI1
      CASE "SHORT"       : vt = VT_I2
      CASE "USHORT"      : vt = VT_UI2
      CASE "INT"         : vt = VT_INT
      CASE "UINT"        : vt = VT_UINT
      CASE "LONG"        : vt = VT_I4
      CASE "ULONG"       : vt = VT_UI4
      CASE "LONGINT"     : vt = VT_I8
      CASE "ULONGINT"    : vt = VT_UI8
      CASE "SINGLE"      : vt = VT_R4
      CASE "FLOAT"       : vt = VT_R4
      CASE "DOUBLE"      : vt = VT_R8
      CASE "VARIANT"     : vt = VT_VARIANT
      CASE "CY"          : vt = VT_CY
      CASE "CURRENCY"    : vt = VT_CY
      CASE "DECIMAL"     : vt = VT_DECIMAL
      CASE "UNKNOWN"     : vt = VT_UNKNOWN
      CASE "DISPATCH"    : vt = VT_DISPATCH
      CASE "FILETIME"    : vt = VT_FILETIME
      CASE "DATE"        : vt = VT_DATE
      CASE "USERDEFINED" : vt = VT_USERDEFINED
      CASE "PTR"         : vt = VT_PTR
   END SELECT
   m_psa = SafeArrayCreate(vt, cDims, prgsabounds)
   CSAFEARRAY_DP("CSafeArray Constructor multidim " & WSTR(m_psa))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' // One-dimensional array
PRIVATE CONSTRUCTOR CSafeArray (BYREF strType AS STRING, BYVAL cElements AS ULONG = 0, BYVAL lLBound AS LONG = 0)
   DIM vt AS WORD
   SELECT CASE UCASE(strType)
      CASE "STRING"      : vt = VT_BSTR
      CASE "BSTR"        : vt = VT_BSTR
      CASE "BOOL"        : vt = VT_BOOL
      CASE "BOOLEAN"     : vt = VT_BOOL
      CASE "BYTE"        : vt = VT_I1
      CASE "UBYTE"       : vt = VT_UI1
      CASE "SHORT"       : vt = VT_I2
      CASE "USHORT"      : vt = VT_UI2
      CASE "INT"         : vt = VT_INT
      CASE "UINT"        : vt = VT_UINT
      CASE "LONG"        : vt = VT_I4
      CASE "ULONG"       : vt = VT_UI4
      CASE "LONGINT"     : vt = VT_I8
      CASE "ULONGINT"    : vt = VT_UI8
      CASE "SINGLE"      : vt = VT_R4
      CASE "FLOAT"       : vt = VT_R4
      CASE "DOUBLE"      : vt = VT_R8
      CASE "VARIANT"     : vt = VT_VARIANT
      CASE "CY"          : vt = VT_CY
      CASE "CURRENCY"    : vt = VT_CY
      CASE "DECIMAL"     : vt = VT_DECIMAL
      CASE "UNKNOWN"     : vt = VT_UNKNOWN
      CASE "DISPATCH"    : vt = VT_DISPATCH
      CASE "FILETIME"    : vt = VT_FILETIME
      CASE "DATE"        : vt = VT_DATE
      CASE "USERDEFINED" : vt = VT_USERDEFINED
      CASE "PTR"         : vt = VT_PTR
   END SELECT
   DIM rgsabounds(0) AS SAFEARRAYBOUND = {(cElements, lLBound)}
   m_psa = SafeArrayCreate(vt, 1, @rgsabounds(0))
   CSAFEARRAY_DP("CSafeArray Constructor one-dimensional " & WSTR(m_psa))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' // Two-dimensional array
PRIVATE CONSTRUCTOR CSafeArray (BYREF strType AS STRING, BYVAL cElements1 AS ULONG, BYVAL lLBound1 AS LONG, BYVAL cElements2 AS ULONG, BYVAL lLBound2 AS LONG)
   DIM vt AS WORD
   SELECT CASE UCASE(strType)
      CASE "STRING"      : vt = VT_BSTR
      CASE "BSTR"        : vt = VT_BSTR
      CASE "BOOL"        : vt = VT_BOOL
      CASE "BOOLEAN"     : vt = VT_BOOL
      CASE "BYTE"        : vt = VT_I1
      CASE "UBYTE"       : vt = VT_UI1
      CASE "SHORT"       : vt = VT_I2
      CASE "USHORT"      : vt = VT_UI2
      CASE "INT"         : vt = VT_INT
      CASE "UINT"        : vt = VT_UINT
      CASE "LONG"        : vt = VT_I4
      CASE "ULONG"       : vt = VT_UI4
      CASE "LONGINT"     : vt = VT_I8
      CASE "ULONGINT"    : vt = VT_UI8
      CASE "SINGLE"      : vt = VT_R4
      CASE "FLOAT"       : vt = VT_R4
      CASE "DOUBLE"      : vt = VT_R8
      CASE "VARIANT"     : vt = VT_VARIANT
      CASE "CY"          : vt = VT_CY
      CASE "CURRENCY"    : vt = VT_CY
      CASE "DECIMAL"     : vt = VT_DECIMAL
      CASE "UNKNOWN"     : vt = VT_UNKNOWN
      CASE "DISPATCH"    : vt = VT_DISPATCH
      CASE "FILETIME"    : vt = VT_FILETIME
      CASE "DATE"        : vt = VT_DATE
      CASE "USERDEFINED" : vt = VT_USERDEFINED
      CASE "PTR"         : vt = VT_PTR
   END SELECT
   DIM rgsabounds(1) AS SAFEARRAYBOUND = {(cElements1, lLBound1), (cElements2, lLBound2)}
   m_psa = SafeArrayCreate(vt, 2, @rgsabounds(0))
   CSAFEARRAY_DP("CSafeArray Constructor two-dimensional " & WSTR(m_psa))
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Creates a safe array from another CSafeArray.
' ========================================================================================
PRIVATE CONSTRUCTOR CSafeArray (BYREF csa AS CSafeArray)
   CSAFEARRAY_DP("CSafeArray Constructor - csa.m_psa = " & WSTR(csa.m_psa))
   this.CopyFrom(csa.m_psa)
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Creates a safe array from another safe array.
' ========================================================================================
PRIVATE CONSTRUCTOR CSafeArray (BYVAL psa AS SAFEARRAY PTR)
   CSAFEARRAY_DP("CSafeArray Constructor - psa = " & WSTR(psa))
   this.CopyFrom(psa)
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Creates a safe array from another safe array.
' If fAttach = TRUE, the safe array is attached, else a copy is made.
' ========================================================================================
PRIVATE CONSTRUCTOR CSafeArray (BYVAL psa AS SAFEARRAY PTR, BYVAL fAttach AS BOOLEAN)
   CSAFEARRAY_DP("CSafeArray Constructor - psa = " & WSTR(psa))
   IF fAttach THEN this.Attach(psa) ELSE this.CopyFrom(psa)
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Creates a safe array from a variant.
' ========================================================================================
PRIVATE CONSTRUCTOR CSafeArray (BYVAL pvar AS VARIANT PTR)
   CSAFEARRAY_DP("CSafeArray Constructor - VARIANT " & WSTR(pVar))
   this.CopyFromVariant(pvar)
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Destroys the safe array when the class is destroyed.
' The array descriptor and all of the data in the array is destroyed. If objects are stored
' in the array, Release is called on each object in the array.
' Safe arrays of variant will have VariantClear called on each member and safe arrays of
' BSTR will have SysFreeString called on each element. IRecordInfo.RecordClear will be called
' to release object references and other values of a record without deallocating the record.
' ========================================================================================
PRIVATE DESTRUCTOR CSafeArray
   CSAFEARRAY_DP("CSafeArray Destructor - " & WSTR(m_psa))
   IF m_psa THEN
      SafeArrayUnlock(m_psa)
      SafeArrayDestroy(m_psa)
   END IF
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Assigns a CSafeArray.
' ========================================================================================
PRIVATE OPERATOR CSafeArray.Let (BYREF csa AS CSafeArray)
   CSAFEARRAY_DP("CSafeArray LET CSafeArray")
   this.CopyFrom(csa.m_psa)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Assigns a safe array.
' ========================================================================================
PRIVATE OPERATOR CSafeArray.Let (BYVAL psa AS SAFEARRAY PTR)
   CSAFEARRAY_DP("CSafeArray LET SAFEARRAY PTR")
   this.CopyFrom(psa)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Assigns a variant.
' ========================================================================================
PRIVATE OPERATOR CSafeArray.Let (BYVAL pvar AS VARIANT PTR)
   CSAFEARRAY_DP("CSafeArray LET VARIANT ")
   this.CopyFromVariant(pvar)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Returns the address of the safe array descriptor.
' Removed to allow to use @ to get the address of the class.
' ========================================================================================
'PRIVATE OPERATOR CSafeArray.@ () AS ANY PTR
'   CSAFEARRAY_DP("CSafeArray Operator @")
'   OPERATOR = @m_psa
'END OPERATOR
' ========================================================================================

' ========================================================================================
' Destroys the safe array and returns the address of the descriptor.
' To be used with BYVAL OUT safe array parameters to avoid a memory leak.
' Gotcha: The returned safe array won't be locked. Better declare a variable as a safe
' array, e.g. DIM psa AS SAFEARRAY PTR, pass it to that function or method, e.g. @psa,
' and then attach if to the class.
' ========================================================================================
'PRIVATE FUNCTION CSafeArray.vptr () AS SAFEARRAY PTR
'   CSAFEARRAY_DP("CSafeArray vptr")
'   IF m_psa THEN
'      SafeArrayUnlock(m_psa)
'      SafeArrayDestroy(m_psa)
'   END IF
'   RETURN @m_psa
'END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a pointer to the safe array descriptor (same as *).
' ========================================================================================
PRIVATE FUNCTION CSafeArray.GetPtr () AS SAFEARRAY PTR
   CSAFEARRAY_DP("CSafeArray GetPtr")
   RETURN m_psa
END FUNCTION
' ========================================================================================
' ========================================================================================
' Returns a pointer to the safe array descriptor.
' ========================================================================================
PRIVATE OPERATOR * (BYREF csa AS CSafeArray) AS SAFEARRAY PTR
   CSAFEARRAY_DP("CSafeArray Operator *")
   OPERATOR = csa.m_psa
END OPERATOR
' ========================================================================================

' =====================================================================================
' Creates a safe array from the given VARTYPE, number of dimensions and bounds.
' Parameters:
' vt
'   [in] Base type of the array (the VARTYPE of each element of the array).
'   The VARTYPE is restricted to a subset of the variant types.
'   Neither VT_ARRAY nor the VT_BYREF flag can be set.
'   VT_EMPTY and VT_NULL are not valid base types for the array.
'   All other types are legal.
' cDims
'   [in] Number of dimensions in the array.
'   The number cannot be changed after the array is created.
' rgsabound
'   [in] Pointer to a vector of bounds (one for each dimension) to allocate for the array.
' Return value:
'   Returns S_OK on success, or an error HRESULT on failure.
' =====================================================================================
' // Multidimensional array
PRIVATE FUNCTION CSafeArray.Create (BYVAL vt AS VARTYPE, BYVAL cDims AS UINT, BYVAL prgsabound AS SAFEARRAYBOUND PTR) AS HRESULT
   CSAFEARRAY_DP("CSafeArray Create - multidimensional")
   IF m_psa <> NULL THEN RETURN E_FAIL
   IF prgsabound = NULL THEN RETURN E_INVALIDARG
   IF cDims < 1 THEN RETURN E_INVALIDARG
   m_psa = SafeArrayCreate(vt, cDims, prgsabound)
   IF m_psa = NULL THEN RETURN E_OUTOFMEMORY
   RETURN SafeArrayLock(m_psa)
END FUNCTION
' =====================================================================================
' =====================================================================================
' // One-dimensional array
PRIVATE FUNCTION CSafeArray.Create (BYVAL vt AS VARTYPE, BYVAL cElements AS ULONG, BYVAL lLBound AS LONG) AS HRESULT
   CSAFEARRAY_DP("CSafeArray Create - one-dimensional")
   DIM rgsabounds(0) AS SAFEARRAYBOUND = {(cElements, lLBound)}
   RETURN this.Create(vt, 1, @rgsabounds(0))
END FUNCTION
' =====================================================================================
' =====================================================================================
' // Two-dimensional array
PRIVATE FUNCTION CSafeArray.Create (BYVAL vt AS VARTYPE, BYVAL cElements1 AS ULONG, BYVAL lLBound1 AS LONG, BYVAL cElements2 AS ULONG, BYVAL lLBound2 AS LONG) AS HRESULT
   CSAFEARRAY_DP("CSafeArray Create - two-dimensional")
   DIM rgsabounds(1) AS SAFEARRAYBOUND = {(cElements1, lLBound1), (cElements2, lLBound2)}
   RETURN this.Create(vt, 2, @rgsabounds(0))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Creates a safe array from the given VARTYPE, number of dimensions and bounds.
' Parameters:
' vt
'   [in] The base type or the VARTYPE of each element of the array. The FADF_RECORD
'   flag can be set for a variant type VT_RECORD, The FADF_HAVEIID flag can be set
'   for VT_DISPATCH or VT_UNKNOWN, and FADF_HAVEVARTYPE can be set for all other VARTYPEs.
' cDims
'   [in] Number of dimensions in the array.
'   The number cannot be changed after the array is created.
' rgsabound
'   [in] Pointer to a vector of bounds (one for each dimension) to allocate for the array.
' pvExtra
'   Points to the type information of the user-defined type, if you are creating a
'   safe array of user-defined types. If the vt parameter is VT_RECORD, then
'   pvExtra will be a pointer to an IRecordInfo describing the record. If the vt
'   parameter is VT_DISPATCH or VT_UNKNOWN, then pvExtra will contain a pointer to
'   a GUID representing the type of interface being passed to the array.
' Return value:
'   Returns S_OK on success, or an error HRESULT on failure.
' Comments:
'   If the VARTYPE is VT_RECORD then SafeArraySetRecordInfo is called. If the
'   VARTYPE is VT_DISPATCH or VT_UNKNOWN then the elements of the array must contain
'   interfaces of the same type. Part of the process of marshaling this array to
'   other processes does include generating the proxy/stub code of the IID pointed
'   to by pvExtra parameter. To actually pass heterogeneous interfaces one will need
'   to specify either IID_IUnknown or IID_IDispatch in pvExtra and provide some
'   other means for the caller to identify how to query for the actual interface.
' =====================================================================================
' // Multidimensional array
PRIVATE FUNCTION CSafeArray.CreateEx (BYVAL vt AS VARTYPE, BYVAL cDims AS UINT, BYVAL prgsabound AS SAFEARRAYBOUND PTR, BYVAL pvExtra AS PVOID) AS HRESULT
   CSAFEARRAY_DP("CSafeArray CreateEx - multidimensional")
   IF m_psa <> NULL THEN RETURN E_FAIL
   IF prgsabound = NULL THEN RETURN E_INVALIDARG
   IF cDims < 1 THEN RETURN E_INVALIDARG
   m_psa = SafeArrayCreateEx(vt, cDims, prgsabound, pvExtra)
   IF m_psa = NULL THEN RETURN E_OUTOFMEMORY
   RETURN SafeArrayLock(m_psa)
END FUNCTION
' =====================================================================================
' =====================================================================================
' // One-dimensional array
PRIVATE FUNCTION CSafeArray.CreateEx (BYVAL vt AS VARTYPE, BYVAL cElements AS ULONG, BYVAL lLBound AS LONG, BYVAL pvExtra AS ANY PTR) AS HRESULT
   CSAFEARRAY_DP("CSafeArray CreateEx - one-dimensional")
   DIM rgsabounds(0) AS SAFEARRAYBOUND = {(cElements, lLBound)}
   RETURN this.CreateEx(vt, 1, @rgsabounds(0), pvExtra)
END FUNCTION
' =====================================================================================
' =====================================================================================
' // Two-dimensional array
PRIVATE FUNCTION CSafeArray.CreateEx (BYVAL vt AS VARTYPE, BYVAL cElements1 AS ULONG, BYVAL lLBound1 AS LONG, BYVAL cElements2 AS ULONG, BYVAL lLBound2 AS LONG, BYVAL pvExtra AS ANY PTR) AS HRESULT
   CSAFEARRAY_DP("CSafeArray CreateEx - two-dimensional")
   DIM rgsabounds(1) AS SAFEARRAYBOUND = {(cElements1, lLBound1), (cElements2, lLBound2)}
   RETURN this.CreateEx(vt, 2, @rgsabounds(0), pvExtra)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Creates a one-dimensional array. A safe array created with SafeArrayCreateVector
' is a fixed size, so the constant FADF_FIXEDSIZE is always set.
' Parameters:
' vt
'   [in] Base type of the array (the VARTYPE of each element of the array).
'   The VARTYPE is restricted to a subset of the variant types.
'   Neither VT_ARRAY nor the VT_BYREF flag can be set.
'   VT_EMPTY and VT_NULL are not valid base types for the array.
'   All other types are legal.
' cElements
'   The number of elements in the array.
' lLBound
'   The lower bound value; that is, the index of the first element in the array.
'   Can be negative.
' Return value:
'   Returns S_OK on success, or an error HRESULT on failure.
' Comments:
'  SafeArrayCreateVector allocates a single block of memory containing a SAFEARRAY
'  structure for a single-dimension array (24 bytes), immediately followed by the
'  array data. All of the existing safe array functions work correctly for safe
'  arrays that are allocated with SafeArrayCreateVector.
'  A SafeArrayCreateVector is allocated as a single block of memory. Both the
'  SafeArray descriptor and the array data block are allocated contiguously in one
'  allocation, which speeds up array allocation.
' =====================================================================================
PRIVATE FUNCTION CSafeArray.CreateVector (BYVAL vt AS VARTYPE, BYVAL cElements AS ULONG, BYVAL lLBound AS LONG) AS HRESULT
   CSAFEARRAY_DP("CSafeArray CreateVector")
   IF m_psa <> NULL THEN RETURN E_FAIL
   m_psa = SafeArrayCreateVector(vt, lLBound, cElements)
   IF m_psa = NULL THEN RETURN E_OUTOFMEMORY
   RETURN SafeArrayLock(m_psa)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Creates a one-dimensional array of the specified type and bounds.
' Parameters:
' vt
'   The base type of the array (the VARTYPE of each element of the array). The
'   FADF_RECORD flag can be set for VT_RECORD. The FADF_HAVEIID can be set for
'   VT_DISPATCH or VT_UNKNOWN and FADF_HAVEVARTYPE can be set for all other types.
' cElements
'   The number of elements in the array.
' lLBound
'   The lower bound value; that is, the index of the first element in the array.
'   Can be negative.
' pvExtra
'   Points to the type information of the user-defined type, if you are creating a
'   safe array of user-defined types. If the vt parameter is VT_RECORD, then
'   pvExtra will be a pointer to an IRecordInfo describing the record. If the vt
'   parameter is VT_DISPATCH or VT_UNKNOWN, then pvExtra will contain a pointer to
'   a GUID representing the type of interface being passed to the array.
' OBJRESULT:
'   Returns S_OK on success, or an error HRESULT on failure.
' =====================================================================================
PRIVATE FUNCTION CSafeArray.CreateVectorEx (BYVAL vt AS VARTYPE, BYVAL cElements AS ULONG, BYVAL lLBound AS LONG, BYVAL pvExtra AS ANY PTR) AS HRESULT
   CSAFEARRAY_DP("CSafeArray CreateVectorEx")
   IF m_psa <> NULL THEN RETURN E_FAIL
   m_psa = SafeArrayCreateVectorEx(vt, lLBound, cElements, pvExtra)
   IF m_psa = NULL THEN RETURN E_OUTOFMEMORY
   RETURN SafeArrayLock(m_psa)
END FUNCTION
' =====================================================================================

' ========================================================================================
' Returns the number of dimensions in the array.
' ========================================================================================
PRIVATE FUNCTION CSafeArray.NumDims () AS UINT
   CSAFEARRAY_DP("CSafeArray NumDims")
   IF m_psa THEN RETURN SafeArrayGetDim(m_psa)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the lower bound of the safe array.
' ========================================================================================
PRIVATE FUNCTION CSafeArray.LBound (BYVAL nDim AS UINT = 1) AS LONG
   CSAFEARRAY_DP("CSafeArray LBound")
   IF m_psa = NULL THEN RETURN 0
   DIM plLBound AS LONG
   SafeArrayGetLBound(m_psa, nDim, @plLBound)
   RETURN plLBound
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the upper bound of the safe array.
' ========================================================================================
PRIVATE FUNCTION CSafeArray.UBound (BYVAL nDim AS UINT = 1) AS LONG
   CSAFEARRAY_DP("CSafeArray UBound")
   IF m_psa = NULL THEN RETURN 0
   DIM plUBound AS LONG
   SafeArrayGetUBound(m_psa, nDim, @plUBound)
   RETURN plUBound
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the number of elements in the specified dimension of the array.
' Parameter:
' nDim: The array dimension for which to get the number of elements.
' ========================================================================================
PRIVATE FUNCTION CSafeArray.Count (BYVAL nDim AS UINT = 1) AS UINT
   CSAFEARRAY_DP("CSafeArray Count - multidimensional array")
   IF m_psa = NULL THEN RETURN 0
   IF nDim < 1 OR nDim > this.NumDims THEN RETURN 0
   DIM plLBound AS LONG, plUBound AS LONG
   SafeArrayGetLBound(m_psa, nDim, @plLBound)
   SafeArrayGetUBound(m_psa, nDim, @plUBound)
   RETURN plUbound - plLBound + 1
END FUNCTION
' ========================================================================================

' =====================================================================================
' Returns the size (in bytes) of an element of the array.
' Does not include size of pointed-to data.
' =====================================================================================
PRIVATE FUNCTION CSafeArray.ElemSize () AS UINT
   CSAFEARRAY_DP("CSafeArray ElemSize")
   IF m_psa THEN RETURN SafeArrayGetElemSize(m_psa)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Tests if the safe array can be resized.
' =====================================================================================
PRIVATE FUNCTION CSafeArray.IsResizable () AS BOOLEAN
   CSAFEARRAY_DP("CSafeArray IsResizable")
   IF m_psa = NULL THEN RETURN FALSE
   RETURN (m_psa->fFeatures AND FADF_FIXEDSIZE) <> FADF_FIXEDSIZE
END FUNCTION
' =====================================================================================

' =====================================================================================
' Returns the VARTYPE stored in the given safe array.
' =====================================================================================
PRIVATE FUNCTION CSafeArray.GetType () AS VARTYPE
   IF m_psa = NULL THEN RETURN 0
   DIM vt AS VARTYPE
   SafeArrayGetVartype(m_psa, @vt)
   CSAFEARRAY_DP("CSafeArray VarType = " & WSTR(vt))
   RETURN vt
END FUNCTION
' =====================================================================================

' =====================================================================================
' Returns the number of Number of times the array has been locked without
' the corresponding unlock.
' =====================================================================================
PRIVATE FUNCTION CSafeArray.LocksCount () AS UINT
   CSAFEARRAY_DP("CSafeArray LocksCount")
   IF m_psa THEN RETURN m_psa->cLocks
END FUNCTION
' =====================================================================================

' =====================================================================================
' Returns the flags used by the safe array.
' =====================================================================================
PRIVATE FUNCTION CSafeArray.Flags () AS USHORT
   CSAFEARRAY_DP("CSafeArray Flags")
   IF m_psa THEN RETURN m_psa->fFeatures
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CSafeArray.Features () AS USHORT
   CSAFEARRAY_DP("CSafeArray Features")
   IF m_psa THEN RETURN m_psa->fFeatures
END FUNCTION
' =====================================================================================

' ========================================================================================
' Increments the lock count of an array, and retrieves a pointer to the array data.
' Return value: A pointer to the array data.
' Remarks: After calling AccessData, you must call the UnaccessData function to unlock the array.
' ========================================================================================
PRIVATE FUNCTION CSafeArray.AccessData () AS ANY PTR
   CSAFEARRAY_DP("CSafeArray AccessData")
   DIM pvData AS ANY PTR
   IF m_psa THEN SafeArrayAccessData(m_psa, @pvData)
   RETURN pvData
END FUNCTION
' ========================================================================================

' ========================================================================================
' Decrements the lock count of an array, and invalidates the pointer retrieved by AccessData.
' ========================================================================================
PRIVATE FUNCTION CSafeArray.UnaccessData () AS HRESULT
   CSAFEARRAY_DP("CSafeArray UnaccessData")
   RETURN SafeArrayUnaccessData(m_psa)
END FUNCTION
' ========================================================================================

' =====================================================================================
' Changes the right-most (least significant) bound of a safe array.
' Parameter:
' psaboundNew
'   Pointer to a new safe array bound structure that contains the new array boundary.
'   You can change only the least significant dimension of an array.
' Comments:
'   If you reduce the bound of an array, SafeArrayRedim deallocates the array
'   elements outside the new array boundary. If the bound of an array is increased,
'   SafeArrayRedim allocates and initializes the new array elements. The data is
'   preserved for elements that exist in both the old and new array.
' OBJRESULT:
'   S_OK Success.
'   %DISP_E_ARRAYISLOCKED The array is currently locked.
'   E_INVALIDARG psaboundNew is a null pointer.
'   E_FAIL The item pointed to by m_psa is not a safe array descriptor.
'          It is a fixed-size array.
' =====================================================================================
PRIVATE FUNCTION CSafeArray.Redim (BYVAL pnewsabounds AS SAFEARRAYBOUND PTR) AS HRESULT
   CSAFEARRAY_DP("CSafeArray Redim - multidimensional")
   IF m_psa = NULL THEN RETURN E_FAIL
   IF (m_psa->fFeatures AND FADF_FIXEDSIZE) = FADF_FIXEDSIZE THEN RETURN E_FAIL
   DIM hr AS HRESULT = SafeArrayUnlock(m_psa)
   IF SUCCEEDED(hr) THEN
      hr = SafeArrayRedim(m_psa, pnewsabounds)
      SafeArrayLock(m_psa)
   END IF
   RETURN hr
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CSafeArray.Redim (BYVAL cElements AS ULONG, BYVAL lLBound AS LONG) AS HRESULT
   CSAFEARRAY_DP("CSafeArray Redim - one-dimensional")
   DIM sanewbounds(0) AS SAFEARRAYBOUND = {(cElements, lLBound)}
   RETURN SafeArrayRedim(m_psa, @sanewbounds(0))
END FUNCTION
' ========================================================================================
' =====================================================================================
PRIVATE FUNCTION CSafeArray.Redim (BYVAL cElements1 AS ULONG, BYVAL lLBound1 AS LONG, BYVAL cElements2 AS ULONG, BYVAL lLBound2 AS LONG) AS HRESULT
   CSAFEARRAY_DP("CSafeArray Redim - two-dimensional")
   DIM sanewbounds(1) AS SAFEARRAYBOUND = {(cElements1, lLBound1), (cElements2, lLBound2)}
   RETURN SafeArrayRedim(m_psa, @sanewbounds(0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates a copy of the safe array.
' Return value:
' Pointer of the new array descriptor. You must free this pointer calling the API
' function SafeArrayDestroy.
' ========================================================================================
PRIVATE FUNCTION CSafeArray.Copy () AS SAFEARRAY PTR
   CSAFEARRAY_DP("CSafeArray Copy")
   DIM psaOut AS SAFEARRAY PTR
   IF m_psa THEN SafeArrayCopy(m_psa, @psaOut)
   RETURN psaOut
END FUNCTION
' ========================================================================================

' ========================================================================================
' Copies the source array to the target array after releasing any resources in the
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
PRIVATE FUNCTION CSafeArray.CopyData (BYVAL psaTarget AS SAFEARRAY PTR) AS HRESULT
   CSAFEARRAY_DP("CSafeArray CopYData")
   RETURN SafeArrayCopyData(m_psa, psaTarget)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Copies the contents of a safe array.
' Parameter:
' - psaSrc: Pointer to an array descriptor created by SafeArrayCreate.
' Return value:
' Returns S_OK on success, or an HRESULT on failure.
' Remarks:
' SafeArrayPutElement, automatically calls SafeArrayLock and SafeArrayUnlock before and
' after assigning the element.
' ========================================================================================
PRIVATE FUNCTION CSafeArray.CopyFrom (BYVAL psaSrc AS SAFEARRAY PTR) AS HRESULT
   CSAFEARRAY_DP("CSafeArray CopyFrom - safe array ptr = " & WSTR(psaSrc))
   IF psaSrc = NULL THEN RETURN E_INVALIDARG
   ' // Destroy the safe array
   IF m_psa THEN this.Destroy
   ' // Copy the passed safe array and lock it
   DIM hr AS HRESULT = SafeArrayCopy(psaSrc, @m_psa)
   SafeArrayLock(m_psa)
   CSAFEARRAY_DP("CSafeArray CopyFrom - m_psa = " & WSTR(m_psa))
   RETURN hr
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a pointer to an array element.
' Parameters:
' prgIndices
'   An array of index values that identify an element of the array. All indexes for
'   the element must be specified.
' Return Value:
'   Pointer to the array element.
' ========================================================================================
' // Multidimensional array
' // prgIndices: first the element, then the dimension
PRIVATE FUNCTION CSafeArray.PtrOfIndex (BYVAL prgIndices AS LONG PTR) AS ANY PTR
   CSAFEARRAY_DP("CSafeArray PtrOfIndex - multidimensional")
   IF m_psa = NULL THEN RETURN NULL
   DIM pvData AS ANY PTR
   SafeArrayPtrOfIndex(m_psa, prgIndices, @pvData)
   RETURN pvData
END FUNCTION
' ========================================================================================
' =====================================================================================
' // One-dimensional array
PRIVATE FUNCTION CSafeArray.PtrOfIndex (BYVAL idx AS LONG) AS ANY PTR
   CSAFEARRAY_DP("CSafeArray PtrOfIndex - one-dimensional")
   IF m_psa = NULL THEN RETURN NULL
   DIM pvData AS ANY PTR
   SafeArrayPtrOfIndex(m_psa, @idx, @pvData)
   RETURN pvData
END FUNCTION
' =====================================================================================
' =====================================================================================
' // Two-dimensional array
' // First element, then dimension, e.g. 2, 1 (element 2, first dimension), 1, 2 (element 1, 2nd dimension).
PRIVATE FUNCTION CSafeArray.PtrOfIndex (BYVAL cElem AS LONG, BYVAL cDim AS LONG) AS ANY PTR
   CSAFEARRAY_DP("CSafeArray PtrOfIndex - two-dimensional")
   IF m_psa = NULL THEN RETURN NULL
   DIM rgIdx(1) AS LONG = {cElem, cDim}
   DIM pvData AS ANY PTR
   SafeArrayPtrOfIndex(m_psa, @rgIdx(0), @pvData)
   RETURN pvData
END FUNCTION
' =====================================================================================

' ========================================================================================
' Destroys an existing array descriptor and all of the data in the array.
' SysFreeString will be called on each element.
' Return value:
' S_OK Success.
' DISP_E_ARRAYISLOCKED The array is locked.
' ========================================================================================
PRIVATE FUNCTION CSafeArray.Destroy () AS HRESULT
   CSAFEARRAY_DP("CSafeArray Destroy - m_psa = " & WSTR(m_psa))
   DIM hr AS HRESULT
   IF m_psa THEN
      hr = SafeArrayUnlock(m_psa)
      hr = SafeArrayDestroy(m_psa)
      IF SUCCEEDED(hr) THEN m_psa = NULL
   END IF
   RETURN hr
END FUNCTION
' ========================================================================================

' ========================================================================================
' Destroys all the data in a safe array.
' Return value:
'   Returns S_OK on success, or an error HRESULT on failure.
'   E_INVALIDARG: The argument psa is not valid.
'   DISP_E_ARRAYISLOCKED: The array is locked.
' Comments:
'   This method is typically used when freeing safe arrays that contain elements with
'   data types other than variants. If objects are stored in the array, Release is
'   called on each object in the array. Safe arrays of variant will have VariantClear
'   called on each member and safe arrays of BSTR will have SysFreeString called on
'   each element. IRecordInfo::RecordClear will be called to release object references
'   and other values of a record without deallocating the record.
' ========================================================================================
PRIVATE FUNCTION CSafeArray.DestroyData () AS HRESULT
   CSAFEARRAY_DP("CSafeArray DestroyData")
   DIM hr AS HRESULT
   IF m_psa THEN
      SafeArrayUnlock(m_psa)
      hr = SafeArrayDestroyData(m_psa)
'      Don't lock the erased array
'      SafeArrayLock(m_psa)
   END IF
   RETURN hr
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CSafeArray.Clear () AS HRESULT
   RETURN this.Destroy
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CSafeArray.Erase () AS HRESULT
   RETURN this.DestroyData
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CSafeArray.Reset () AS HRESULT
   RETURN this.DestroyData
END FUNCTION
' ========================================================================================

' ========================================================================================
' Copies the contents of a VARIANT of type VT_ARRAY, i.e. containing a safe array, to
' the CSafeArray array. The VARIANT remains unaltered.
' ========================================================================================
PRIVATE FUNCTION CSafeArray.CopyFromVariant (BYVAL pvar AS VARIANT PTR) AS HRESULT
   CSAFEARRAY_DP("CSafeArray CopyFromVariant")
   IF pvar = NULL THEN RETURN E_INVALIDARG
   IF (pvar->vt AND VT_ARRAY) <> VT_ARRAY THEN RETURN HRESULT_FROM_WIN32(ERROR_INVALID_DATA)
   RETURN this.CopyFrom(pvar->parray)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Transfers ownership of the safe array contained in the variant parameter to this
' object. The variant is then changed to VT_EMPTY.
' ========================================================================================
PRIVATE FUNCTION CSafeArray.MoveFromVariant (BYVAL pvar AS VARIANT PTR) AS HRESULT
   CSAFEARRAY_DP("CSafeArray MoveFromVariant")
   IF pvar = NULL THEN RETURN E_INVALIDARG
   IF (pvar->vt AND VT_ARRAY) <> VT_ARRAY THEN RETURN HRESULT_FROM_WIN32(ERROR_INVALID_DATA)
   FUNCTION = this.CopyFrom(pvar->parray)
   pvar->vt = VT_EMPTY
END FUNCTION
' ========================================================================================

' ========================================================================================
' Copies the contents of the safe array to a variant.
' ========================================================================================
PRIVATE FUNCTION CSafeArray.CopyToVariant (BYVAL pvar AS VARIANT PTR) AS HRESULT
   CSAFEARRAY_DP("CSafeArray CopyToVariant")
   IF m_psa = NULL THEN RETURN E_FAIL
   IF pvar = NULL THEN RETURN E_INVALIDARG
   VariantClear(pvar)
   DIM vt AS VARTYPE
   SafeArrayGetVartype(m_psa, @vt)
   pvar->vt = vt OR VT_ARRAY
   RETURN SafeArrayCopy(m_psa, @pvar->parray)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Transfers ownership of the safe array to a variant and detaches it from the object.
' ========================================================================================
PRIVATE FUNCTION CSafeArray.MoveToVariant (BYVAL pvar AS VARIANT PTR) AS HRESULT
   CSAFEARRAY_DP("CSafeArray MoveToVariant")
   IF m_psa = NULL THEN RETURN E_FAIL
   IF pvar = NULL THEN RETURN E_INVALIDARG
   VariantClear(pvar)
   DIM vt AS VARTYPE
   SafeArrayGetVartype(m_psa, @vt)
   pvar->vt = vt OR VT_ARRAY
   FUNCTION = SafeArrayUnlock(m_psa)
   pvar->parray = m_psa
   m_psa = NULL
END FUNCTION
' ========================================================================================

' ========================================================================================
' Attaches a SAFEARRAY descriptor to a CSafeArray array.
' Parameter:
' - psaSrc: A pointer to the SAFEARRAY descriptor.
' Returns S_OK on success, or an error HRESULT on failure.
' Remarks: Never attach the same safe array to more than one CSafeArray class because each one
' will try to destroy it when they are deleted or went out of scope.
' ========================================================================================
PRIVATE FUNCTION CSafeArray.Attach (BYVAL psaSrc AS SAFEARRAY PTR) AS HRESULT
   CSAFEARRAY_DP("CSafeArray Attach")
   IF psaSrc = NULL THEN RETURN E_FAIL
   DIM hr AS HRESULT = this.Destroy
   IF SUCCEEDED(hr) THEN
      m_psa = psaSrc
      hr = SafeArrayLock(m_psa)
   END IF
   RETURN hr
END FUNCTION
' ========================================================================================

' ========================================================================================
' Detaches the SAFEARRAY descriptor from the CSafeArray array.
' Return value: Returns a pointer to a SAFEARRAY descriptor.
' The caller takes ownership of it and must destroy it when no longer needed.
' ========================================================================================
PRIVATE FUNCTION CSafeArray.Detach () AS SAFEARRAY PTR
   CSAFEARRAY_DP("CSafeArray DETACH")
   SafeArrayUnlock(m_psa)
   FUNCTION = m_psa
   m_psa = NULL
END FUNCTION
' ========================================================================================

' =====================================================================================
' Returns the GUID of the interface contained within a given safe array.
' Return Value:
'   The GUID of the interface, on success, or a null guid on failure.
' Return value:
'   If the function succeeds it returns the guid; otherwise, returns a null guid.
' =====================================================================================
PRIVATE FUNCTION CSafeArray.GetIID () AS GUID
   CSAFEARRAY_DP("CSafeArray GetIID")
   DIM hr AS HRESULT, guid_ AS GUID
   IF m_psa = NULL THEN RETURN guid_
   IF (m_psa->fFeatures AND FADF_HAVEIID) <> FADF_HAVEIID THEN RETURN guid_
   hr = SafeArrayGetIID(m_psa, @guid_)
   IF hr = S_OK THEN RETURN guid_
END FUNCTION
' =====================================================================================

' =====================================================================================
' Sets the GUID of the interface contained within a given safe array.
' Return value:
'   S_OK Success.
'   E_FAIL If m_psa is null.
'   E_INVALIDARG If the array descriptor does not have the FADF_HAVEIID flag set.
' =====================================================================================
PRIVATE FUNCTION CSafeArray.SetIID (BYVAL pguid AS GUID PTR) AS HRESULT
   CSAFEARRAY_DP("CSafeArray SetIID")
   IF m_psa = NULL THEN RETURN E_FAIL
   IF (m_psa->fFeatures AND FADF_HAVEIID) <> FADF_HAVEIID THEN RETURN HRESULT_FROM_WIN32(ERROR_INVALID_DATA)
   RETURN SafeArraySetIID(m_psa, pguid)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Retrieves the IRecordInfo interface of the UDT contained in a given safe array.
' Return value:
' If the function succeeds it returns the guid; otherwise, returns a null guid.
' =====================================================================================
PRIVATE FUNCTION CSafeArray.GetRecordInfo () AS IRecordInfo PTR
   CSAFEARRAY_DP("CSafeArray GetRecordInfo")
   IF m_psa = NULL THEN RETURN NULL
   IF (m_psa->fFeatures AND FADF_RECORD) <> FADF_RECORD THEN RETURN NULL
   DIM hr AS HRESULT, prinfo AS IRecordInfo PTR
   hr = SafeArrayGetRecordInfo(m_psa, @prinfo)
   RETURN prinfo
END FUNCTION
' =====================================================================================

' =====================================================================================
' Sets the IRecordInfo Interface of the UDT contained in a given safe array.
' OBJRESULT:
'   S_OK Success.
'   E_FAIL If m_psa is null.
'   E_INVALIDARG If the array descriptor does not have the FADF_RECORD flag set.
' =====================================================================================
PRIVATE FUNCTION CSafeArray.SetRecordInfo (BYVAL prinfo AS IRecordInfo PTR) AS HRESULT
   CSAFEARRAY_DP("CSafeArray SetRecordInfo")
   IF m_psa = NULL THEN RETURN E_FAIL
   IF (m_psa->fFeatures AND FADF_RECORD) <> FADF_RECORD THEN RETURN HRESULT_FROM_WIN32(ERROR_INVALID_DATA)
   RETURN SafeArraySetRecordInfo(m_psa, prinfo)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Retrieves a single element of the array.
' Parameters:
' prgIndices
'   Pointer to a vector of indexes for each dimension of the array. The right-most
'   (least significant) dimension is prgIndices[0]. The left-most dimension is stored
'   at rgIndices[psa->cDims – 1].
' pData
'   Pointer to the location to place the element of the array.
' Comments:
'   This function calls SafeArrayLock and SafeArrayUnlock automatically, before and
'   after retrieving the element. The caller must provide a storage area of the
'   correct size to receive the data. If the data element is a string, object, or
'   variant, the function copies the element in the correct way.
' OBJRESULT:
'   The return value obtained from the returned HRESULT is one of the following.
'   S_OK Success.
'   E_FAIL This safe array has no elements
'   DISP_E_BADINDEX The specified index is invalid.
'   E_INVALIDARG One of the arguments is invalid.
'   E_OUTOFMEMORY Memory could not be allocated for the element.
' =====================================================================================

' =====================================================================================
' // *** BYVAL ANY PTR ***
' =====================================================================================
' // Multidimensional array
' // prgIndices: first the element, then the dimension
PRIVATE FUNCTION CSafeArray.GetElement (BYVAL prgIndices AS LONG PTR, BYVAL pData AS ANY PTR) AS HRESULT
   CSAFEARRAY_DP("CSafeArray GetElement - multi - pData = " & WSTR(pData))
   IF m_psa = NULL THEN RETURN E_FAIL
   RETURN SafeArrayGetElement(m_psa, prgIndices, pData)
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CSafeArray.Get (BYVAL prgIndices AS LONG PTR, BYVAL pData AS ANY PTR) AS HRESULT
   RETURN this.GetElement(prgIndices, pData)
END FUNCTION
' =====================================================================================

' =====================================================================================
' // One-dimensional array
PRIVATE FUNCTION CSafeArray.GetElement (BYVAL idx AS LONG, BYVAL pData AS ANY PTR) AS HRESULT
   CSAFEARRAY_DP("CSafeArray GetElement - 1D - pData = " & WSTR(pData))
   IF m_psa = NULL THEN RETURN E_FAIL
   RETURN SafeArrayGetElement(m_psa, @idx, pData)
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CSafeArray.Get (BYVAL idx AS LONG, BYVAL pData AS ANY PTR) AS HRESULT
   RETURN this.GetElement(idx, pData)
END FUNCTION
' =====================================================================================

' =====================================================================================
' // Two-dimensional array
' // First element, then dimension, e.g. 2, 1 (element 2, first dimension), 1, 2 (element 1, 2nd dimension).
PRIVATE FUNCTION CSafeArray.GetElement (BYVAL cElem AS LONG, BYVAL cDim AS LONG, BYVAL pData AS ANY PTR) AS HRESULT
   CSAFEARRAY_DP("CSafeArray GetElement - 2D - pData = " & WSTR(pData))
   IF m_psa = NULL THEN RETURN E_FAIL
   DIM rgIdx(1) AS LONG = {cElem, cDim}
   RETURN SafeArrayGetElement(m_psa, @rgIdx(0), pData)
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CSafeArray.Get (BYVAL cElem AS LONG, BYVAL cDim AS LONG, BYVAL pData AS ANY PTR) AS HRESULT
   RETURN this.GetElement(cElem, cDim, pData)
END FUNCTION
' =====================================================================================

' =====================================================================================
' ** BYREF CBSTR ***
' =====================================================================================
' // Multidimensional array
' // prgIndices: first the element, then the dimension
PRIVATE FUNCTION CSafeArray.GetElement (BYVAL prgIndices AS LONG PTR, BYREF cbsData AS CBSTR) AS HRESULT
   CSAFEARRAY_DP("CSafeArray GetElement - multi - CBSTR " & WSTR(cbsData.m_bstr))
   IF m_psa = NULL THEN RETURN E_FAIL
   DIM vt AS VARTYPE = this.GetType
   IF vt <> VT_BSTR THEN RETURN E_FAIL
   RETURN SafeArrayGetElement(m_psa, prgIndices, cbsData.vptr)
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CSafeArray.Get (BYVAL prgIndices AS LONG PTR, BYREF cbsData AS CBSTR) AS HRESULT
   RETURN this.GetElement(prgIndices, cbsData)
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CSafeArray.GetStr (BYVAL prgIndices AS LONG PTR) AS CBSTR
   CSAFEARRAY_DP("CSafeArray - multi - GET CBSTR")
   IF m_psa = NULL THEN RETURN E_FAIL
   DIM vt AS VARTYPE = this.GetType
   IF vt <> VT_BSTR THEN RETURN ""
   DIM bs AS AFX_BSTR
   SafeArrayGetElement(m_psa, prgIndices, @bs)
   RETURN bs
END FUNCTION
' =====================================================================================

' =====================================================================================
' // One-dimensional array
PRIVATE FUNCTION CSafeArray.GetElement (BYVAL idx AS LONG, BYREF cbsData AS CBSTR) AS HRESULT
   CSAFEARRAY_DP("CSafeArray GetElement - 1D - CBSTR " & WSTR(cbsData.m_bstr))
   IF m_psa = NULL THEN RETURN E_FAIL
   DIM vt AS VARTYPE = this.GetType
   IF vt <> VT_BSTR THEN RETURN E_FAIL
   RETURN SafeArrayGetElement(m_psa, @idx, cbsData.vptr)
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CSafeArray.Get (BYVAL idx AS LONG, BYREF cbsData AS CBSTR) AS HRESULT
   RETURN this.GetElement(idx, cbsData)
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CSafeArray.GetStr (BYVAL idx AS LONG) AS CBSTR
   CSAFEARRAY_DP("CSafeArray - 1D - GET CBSTR")
   IF m_psa = NULL THEN RETURN E_FAIL
   DIM vt AS VARTYPE = this.GetType
   IF vt <> VT_BSTR THEN RETURN ""
   DIM bs AS AFX_BSTR
   SafeArrayGetElement(m_psa, @idx, @bs)
   RETURN bs
END FUNCTION
' =====================================================================================

' =====================================================================================
' // Two-dimensional array
' // First element, then dimension, e.g. 2, 1 (element 2, first dimension), 1, 2 (element 1, 2nd dimension).
PRIVATE FUNCTION CSafeArray.GetElement (BYVAL cElem AS LONG, BYVAL cDim AS LONG, BYREF cbsData AS CBSTR) AS HRESULT
   CSAFEARRAY_DP("CSafeArray GetElement - 2D - CBSTR " & WSTR(cbsData.m_bstr))
   IF m_psa = NULL THEN RETURN E_FAIL
   DIM vt AS VARTYPE = this.GetType
   IF vt <> VT_BSTR THEN RETURN E_FAIL
   DIM rgIdx(1) AS LONG = {cElem, cDim}
   RETURN SafeArrayGetElement(m_psa, @rgIdx(0), cbsData.vptr)
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CSafeArray.Get (BYVAL cElem AS LONG, BYVAL cDim AS LONG, BYREF cbsData AS CBSTR) AS HRESULT
   RETURN this.GetElement(cElem, cDim, cbsData)
END FUNCTION
' =====================================================================================
' =====================================================================================
' // Two-dimensional array
' // First element, then dimension, e.g. 2, 1 (element 2, first dimension), 1, 2 (element 1, 2nd dimension).
PRIVATE FUNCTION CSafeArray.GetStr (BYVAL cElem AS LONG, BYVAL cDim AS LONG) AS CBSTR
   CSAFEARRAY_DP("CSafeArray - 2D - GET CBSTR")
   IF m_psa = NULL THEN RETURN E_FAIL
   DIM vt AS VARTYPE = this.GetType
   IF vt <> VT_BSTR THEN RETURN ""
   DIM rgIdx(1) AS LONG = {cElem, cDim}
   DIM bs AS AFX_BSTR
   SafeArrayGetElement(m_psa, @rgIdx(0), @bs)
   RETURN bs
END FUNCTION
' =====================================================================================

' =====================================================================================
' ** BYREF CVAR ***
' =====================================================================================
' // Multidimensional array
' // prgIndices: first the element, then the dimension
PRIVATE FUNCTION CSafeArray.GetElement (BYVAL prgIndices AS LONG PTR, BYREF cvData AS CVAR) AS HRESULT
   CSAFEARRAY_DP("CSafeArray GetElement - multi - CVAR")
   IF m_psa = NULL THEN RETURN E_FAIL
   DIM vt AS VARTYPE = this.GetType
   IF vt <> VT_VARIANT THEN RETURN E_FAIL
   ' // No need to clear the variant because SafeArrayGetElement uses VariantCopy to
   ' // copy the data and VariantCopy clears the destination variant.
   RETURN SafeArrayGetElement(m_psa, prgIndices, cvData.vptr)
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CSafeArray.Get (BYVAL prgIndices AS LONG PTR, BYREF cvData AS CVAR) AS HRESULT
   RETURN this.GetElement(prgIndices, cvData)
END FUNCTION
' =====================================================================================
' =====================================================================================
' // Multidimensional array
' // prgIndices: first the element, then the dimension
PRIVATE FUNCTION CSafeArray.GetVar (BYVAL prgIndices AS LONG PTR) AS CVAR
   CSAFEARRAY_DP("CSafeArray - multi - GET CVAR")
   IF m_psa = NULL THEN RETURN E_FAIL
   DIM vt AS VARTYPE = this.GetType
   IF vt <> VT_VARIANT THEN RETURN E_FAIL
   DIM cvData AS CVAR
   SafeArrayGetElement(m_psa, prgIndices, cvData)
   RETURN cvData
END FUNCTION
' =====================================================================================

' =====================================================================================
' // One-dimensional array
PRIVATE FUNCTION CSafeArray.GetElement (BYVAL idx AS LONG, BYREF cvData AS CVAR) AS HRESULT
   CSAFEARRAY_DP("CSafeArray GetElement - 1D - CVAR")
   IF m_psa = NULL THEN RETURN E_FAIL
   DIM vt AS VARTYPE = this.GetType
   IF vt <> VT_VARIANT THEN RETURN E_FAIL
   RETURN SafeArrayGetElement(m_psa, @idx, cvData.vptr)
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CSafeArray.Get (BYVAL idx AS LONG, BYREF cvData AS CVAR) AS HRESULT
   RETURN this.GetElement(idx, cvData)
END FUNCTION
' =====================================================================================
' =====================================================================================
' // One-dimensional array
PRIVATE FUNCTION CSafeArray.GetVar (BYVAL idx AS LONG) AS CVAR
   CSAFEARRAY_DP("CSafeArray - 1D - GET CVAR")
   IF m_psa = NULL THEN RETURN E_FAIL
   DIM vt AS VARTYPE = this.GetType
   IF vt <> VT_VARIANT THEN RETURN E_FAIL
   DIM cvData AS CVAR
   SafeArrayGetElement(m_psa, @idx, cvData)
   RETURN cvData
END FUNCTION
' =====================================================================================

' =====================================================================================
' // Two-dimensional array
' // First element, then dimension, e.g. 2, 1 (element 2, first dimension), 1, 2 (element 1, 2nd dimension).
PRIVATE FUNCTION CSafeArray.GetElement (BYVAL cElem AS LONG, BYVAL cDim AS LONG, BYREF cvData AS CVAR) AS HRESULT
   CSAFEARRAY_DP("CSafeArray GetElement - 2D - CVAR")
   IF m_psa = NULL THEN RETURN E_FAIL
   DIM vt AS VARTYPE = this.GetType
   IF vt <> VT_VARIANT THEN RETURN E_FAIL
   DIM rgIdx(1) AS LONG = {cElem, cDim}
   RETURN SafeArrayGetElement(m_psa, @rgIdx(0), cvData.vptr)
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CSafeArray.Get (BYVAL cElem AS LONG, BYVAL cDim AS LONG, BYREF cvData AS CVAR) AS HRESULT
   RETURN this.GetElement(cElem, cDim, cvData)
END FUNCTION
' =====================================================================================
' =====================================================================================
' // Two-dimensional array
' // First element, then dimension, e.g. 2, 1 (element 2, first dimension), 1, 2 (element 1, 2nd dimension).
PRIVATE FUNCTION CSafeArray.GetVar (BYVAL cElem AS LONG, BYVAL cDim AS LONG) AS CVAR
   CSAFEARRAY_DP("CSafeArray GetElement - 2D - CVAR")
   IF m_psa = NULL THEN RETURN E_FAIL
   DIM vt AS VARTYPE = this.GetType
   IF vt <> VT_VARIANT THEN RETURN E_FAIL
   DIM rgIdx(1) AS LONG = {cElem, cDim}
   DIM cvData AS CVAR
   SafeArrayGetElement(m_psa, @rgIdx(0), cvData)
   RETURN cvData
END FUNCTION
' =====================================================================================

' =====================================================================================
' Stores the data element at a given location in the array.
' Parameters
' prgIndices
'   Pointer to a vector of indexes for each dimension of the array. The right-most
'   (least significant) dimension is prgIndices[0]. The left-most dimension is stored
'   at rgIndices[psa->cDims – 1].
' pData
'   Pointer to the data to assign to the array. The variant types VT_DISPATCH,
'   VT_UNKNOWN, and VT_BSTR are pointers, and do not require another level of indirection.
' Comments:
'   This function automatically calls SafeArrayLock and SafeArrayUnlock before and
'   after assigning the element. If the data element is a string, object, or variant,
'   the function copies it correctly when the safe array is destroyed. If the
'   existing element is a string, object, or variant, it is cleared correctly. If
'   the data element is a VT_DISPATCH or VT_UNKNOWN, AddRef is called to increment
'   the object's reference count.
' Note
'   Multiple locks can be on an array. Elements can be put into an array while the
'   array is locked by other operations.
' OBJRESULT:
'   The return value obtained from the returned HRESULT is one of the following.
'   S_OK Success.
'   E_FAIL This safe array has no elements
'   DISP_E_BADINDEX The specified index was invalid.
'   E_INVALIDARG One of the arguments is invalid.
'   E_OUTOFMEMORY Memory could not be allocated for the element.
' =====================================================================================

' =====================================================================================
' *** ANY PTR ***
' =====================================================================================
' // Multidimensional array
PRIVATE FUNCTION CSafeArray.PutElement (BYVAL prgIndices AS LONG PTR, BYVAL pData AS ANY PTR) AS HRESULT
   CSAFEARRAY_DP("CSafeArray PutElement - multidim - pData = " & WSTR(pData))
   IF m_psa = NULL THEN RETURN E_FAIL
   RETURN SafeArrayPutElement(m_psa, prgIndices, pData)
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CSafeArray.Put (BYVAL prgIndices AS LONG PTR, BYVAL pData AS ANY PTR) AS HRESULT
   RETURN this.PutElement(prgIndices, pData)
END FUNCTION
' =====================================================================================

' =====================================================================================
' // One-dimensional array
PRIVATE FUNCTION CSafeArray.PutElement (BYVAL idx AS LONG, BYVAL pData AS ANY PTR) AS HRESULT
   CSAFEARRAY_DP("CSafeArray PutElement - 1D - pData = " & WSTR(pData))
   IF m_psa = NULL THEN RETURN E_FAIL
   RETURN SafeArrayPutElement(m_psa, @idx, pData)
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CSafeArray.Put (BYVAL idx AS LONG, BYVAL pData AS ANY PTR) AS HRESULT
   RETURN this.PutElement(idx, pData)
END FUNCTION
' =====================================================================================

' =====================================================================================
' // Two-dimensional array
' // First element, then dimension, e.g. 2, 1 (element 2, first dimension), 1, 2 (element 1, 2nd dimension).
PRIVATE FUNCTION CSafeArray.PutElement (BYVAL cElem AS LONG, BYVAL cDim AS LONG, BYVAL pData AS ANY PTR) AS HRESULT
   CSAFEARRAY_DP("CSafeArray PutElement - 2D - pData = " & WSTR(pData))
   IF m_psa = NULL THEN RETURN E_FAIL
   DIM rgIdx(1) AS LONG = {cElem, cDim}
   RETURN SafeArrayPutElement(m_psa, @rgIdx(0), pData)
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CSafeArray.Put (BYVAL cElem AS LONG, BYVAL cDim AS LONG, BYVAL pData AS ANY PTR) AS HRESULT
   RETURN this.PutElement(cElem, cDim, pData)
END FUNCTION
' =====================================================================================

' =====================================================================================
' *** BYREF CBSTR ***
' =====================================================================================
' // Multidimensional array
PRIVATE FUNCTION CSafeArray.PutElement (BYVAL prgIndices AS LONG PTR, BYREF cbsData AS CBSTR) AS HRESULT
   CSAFEARRAY_DP("CSafeArray PutElement - multidim - CBSTR")
   IF m_psa = NULL THEN RETURN E_FAIL
   DIM vt AS VARTYPE = this.GetType
   IF vt <> VT_BSTR THEN RETURN E_FAIL
   RETURN SafeArrayPutElement(m_psa, prgIndices, cbsData.m_bstr)
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CSafeArray.PutStr (BYVAL prgIndices AS LONG PTR, BYREF cbsData AS CBSTR) AS HRESULT
   RETURN this.PutElement(prgIndices, cbsData)
END FUNCTION
' =====================================================================================

' =====================================================================================
' // One-dimensional array
PRIVATE FUNCTION CSafeArray.PutElement (BYVAL idx AS LONG, BYREF cbsData AS CBSTR) AS HRESULT
   CSAFEARRAY_DP("CSafeArray PutElement - 1D -  CBSTR")
   IF m_psa = NULL THEN RETURN E_FAIL
   DIM vt AS VARTYPE = this.GetType
   IF vt <> VT_BSTR THEN RETURN E_FAIL
   RETURN SafeArrayPutElement(m_psa, @idx, cbsData.m_bstr)
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CSafeArray.PutStr (BYVAL idx AS LONG, BYREF cbsData AS CBSTR) AS HRESULT
   RETURN this.PutElement(idx, cbsData)
END FUNCTION
' =====================================================================================

' =====================================================================================
' // Two-dimensional array
' // First element, then dimension, e.g. 2, 1 (element 2, first dimension), 1, 2 (element 1, 2nd dimension).
PRIVATE FUNCTION CSafeArray.PutElement (BYVAL cElem AS LONG, BYVAL cDim AS LONG, BYREF cbsData AS CBSTR) AS HRESULT
   CSAFEARRAY_DP("CSafeArray PutElement - 2D -  CBSTR")
   IF m_psa = NULL THEN RETURN E_FAIL
   DIM rgIdx(1) AS LONG = {cElem, cDim}
   DIM vt AS VARTYPE = this.GetType
   IF vt <> VT_BSTR THEN RETURN E_FAIL
   RETURN SafeArrayPutElement(m_psa, @rgIdx(0), cbsData.m_bstr)
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CSafeArray.PutStr (BYVAL cElem AS LONG, BYVAL cDim AS LONG, BYREF cbsData AS CBSTR) AS HRESULT
   RETURN this.PutElement(cElem, cDim, cbsData)
END FUNCTION
' =====================================================================================

' =====================================================================================
' *** BYREF CVAR ***
' =====================================================================================
' // Multidimensional array
PRIVATE FUNCTION CSafeArray.PutElement (BYVAL prgIndices AS LONG PTR, BYREF cvData AS CVAR) AS HRESULT
   CSAFEARRAY_DP("CSafeArray PutElement - multidim - CVAR")
   IF m_psa = NULL THEN RETURN E_FAIL
   DIM vt AS VARTYPE = this.GetType
   IF vt <> VT_VARIANT THEN RETURN E_FAIL
   RETURN SafeArrayPutElement(m_psa, prgIndices, cvData)
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CSafeArray.PutVar (BYVAL prgIndices AS LONG PTR, BYREF cvData AS CVAR) AS HRESULT
   RETURN this.PutElement(prgIndices, cvData)
END FUNCTION
' =====================================================================================

' =====================================================================================
' // One-dimensional array
PRIVATE FUNCTION CSafeArray.PutElement (BYVAL idx AS LONG, BYREF cvData AS CVAR) AS HRESULT
   CSAFEARRAY_DP("CSafeArray PutElement - 1D -  CVAR")
   IF m_psa = NULL THEN RETURN E_FAIL
   DIM vt AS VARTYPE = this.GetType
   IF vt <> VT_VARIANT THEN RETURN E_FAIL
   RETURN SafeArrayPutElement(m_psa, @idx, cvData)
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CSafeArray.PutVar (BYVAL idx AS LONG, BYREF cvData AS CVAR) AS HRESULT
   RETURN this.PutElement(idx, cvData)
END FUNCTION
' =====================================================================================

' =====================================================================================
' // Two-dimensional array
' // First element, then dimension, e.g. 2, 1 (element 2, first dimension), 1, 2 (element 1, 2nd dimension).
PRIVATE FUNCTION CSafeArray.PutElement (BYVAL cElem AS LONG, BYVAL cDim AS LONG, BYREF cvData AS CVAR) AS HRESULT
   CSAFEARRAY_DP("CSafeArray PutElement - 2D -  CVAR")
   IF m_psa = NULL THEN RETURN E_FAIL
   DIM vt AS VARTYPE = this.GetType
   IF vt <> VT_VARIANT THEN RETURN E_FAIL
   DIM rgIdx(1) AS LONG = {cElem, cDim}
   RETURN SafeArrayPutElement(m_psa, @rgIdx(0), cvData)
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CSafeArray.PutVar (BYVAL cElem AS LONG, BYVAL cDim AS LONG, BYREF cvData AS CVAR) AS HRESULT
   RETURN this.PutElement(cElem, cDim, cvData)
END FUNCTION
' =====================================================================================

' ========================================================================================
' Appends a value to the end of the one-dimensional safe array.
' If the safe array is not a one-dimensional array or it is not resizable it will return E_FAIL.
' Return value:
'   S_OK Success.
'   DISP_E_BADINDEX The specified index is not valid.
'   E_INVALIDARG One of the arguments is not valid.
'   E_OUTOFMEMORY Memory could not be allocated for the element.
'   E_FAIL The item pointed to by m_psa is not a safe array descriptor.
'          It is a fixed-size array.
'          It is not a one-dimensional array.
' ========================================================================================
PRIVATE FUNCTION CSafeArray.AppendElement (BYVAL pData AS ANY PTR) AS HRESULT
   CSAFEARRAY_DP("CSafeArray AppendElement - ANY PTR")
   IF m_psa = NULL THEN RETURN E_FAIL
   DIM nDims AS UINT = SafeArrayGetDim(m_psa)
   IF nDims <> 1 THEN RETURN E_FAIL
   IF this.IsResizable = FALSE THEN RETURN E_FAIL
   DIM cElements AS DWORD = this.Count(1) + 1
   DIM lLBound AS LONG = this.LBound(1)
   DIM sanewbounds(0) AS SAFEARRAYBOUND = {(cElements, lLBound)}
   DIM hr AS HRESULT = SafeArrayRedim(m_psa, @sanewbounds(0))
   IF hr THEN RETURN hr
   DIM idx AS LONG = cElements - 1 + lLBound
   RETURN SafeArrayPutElement(m_psa, @idx, pData)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CSafeArray.Append (BYVAL pData AS ANY PTR) AS HRESULT
   RETURN this.AppendElement(pData)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Appends a BSTR to the end of the one-dimensional VT_BSTR safe array.
' If the safe array is not a one-dimensional VT_BSTR array or it is not resizable it will
' return E_FAIL.
' Return value:
'   S_OK Success.
'   DISP_E_BADINDEX The specified index is not valid.
'   E_INVALIDARG One of the arguments is not valid.
'   E_OUTOFMEMORY Memory could not be allocated for the element.
'   E_FAIL The item pointed to by m_psa is not a safe array descriptor.
'          It is a fixed-size array.
'          It is not a one-dimensional array.
' ========================================================================================
PRIVATE FUNCTION CSafeArray.AppendElement (BYREF cbsData AS CBSTR) AS HRESULT
   CSAFEARRAY_DP("CSafeArray AppendElement - CBSTR")
   IF m_psa = NULL THEN RETURN E_FAIL
   DIM nDims AS UINT = SafeArrayGetDim(m_psa)
   IF nDims <> 1 THEN RETURN E_FAIL
   IF this.IsResizable = FALSE THEN RETURN E_FAIL
   DIM vt AS VARTYPE = this.GetType
   IF vt <> VT_BSTR THEN RETURN E_FAIL
   DIM cElements AS DWORD = this.Count(1) + 1
   DIM lLBound AS LONG = this.LBound(1)
   DIM sanewbounds(0) AS SAFEARRAYBOUND = {(cElements, lLBound)}
   DIM hr AS HRESULT = SafeArrayRedim(m_psa, @sanewbounds(0))
   IF hr THEN RETURN hr
   DIM idx AS LONG = cElements - 1 + lLBound
   RETURN SafeArrayPutElement(m_psa, @idx, cbsData.m_bstr)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CSafeArray.Append (BYREF cbsData AS CBSTR) AS HRESULT
   RETURN this.AppendElement(cbsData)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CSafeArray.AppendStr (BYVAL pwszData AS WSTRING PTR) AS HRESULT
   CSAFEARRAY_DP("CSafeArray AppendStr")
   IF m_psa = NULL THEN RETURN E_FAIL
   DIM nDims AS UINT = SafeArrayGetDim(m_psa)
   IF nDims <> 1 THEN RETURN E_FAIL
   IF this.IsResizable = FALSE THEN RETURN E_FAIL
   DIM vt AS VARTYPE = this.GetType
   IF vt <> VT_BSTR THEN RETURN E_FAIL
   DIM cElements AS DWORD = this.Count(1) + 1
   DIM lLBound AS LONG = this.LBound(1)
   DIM sanewbounds(0) AS SAFEARRAYBOUND = {(cElements, lLBound)}
   DIM hr AS HRESULT = SafeArrayRedim(m_psa, @sanewbounds(0))
   IF hr THEN RETURN hr
   DIM idx AS LONG = cElements - 1 + lLBound
   DIM bs AS AFX_BSTR = SysAllocString(pwszData)
   FUNCTION = SafeArrayPutElement(m_psa, @idx, bs)
   SysFreeString bs
END FUNCTION
' ========================================================================================

' ========================================================================================
' Appends a VARIANT to the end of the one-dimensional VT_VARIANT safe array.
' If the safe array is not a one-dimensional VT_VARIANT array or it is not resizable it will
' return E_FAIL.
' Return value:
'   S_OK Success.
'   DISP_E_BADINDEX The specified index is not valid.
'   E_INVALIDARG One of the arguments is not valid.
'   E_OUTOFMEMORY Memory could not be allocated for the element.
'   E_FAIL The item pointed to by m_psa is not a safe array descriptor.
'          It is a fixed-size array.
'          It is not a one-dimensional array.
' ========================================================================================
PRIVATE FUNCTION CSafeArray.AppendElement (BYREF cvData AS CVAR) AS HRESULT
   CSAFEARRAY_DP("CSafeArray AppendElement - CVAR")
   IF m_psa = NULL THEN RETURN E_FAIL
   DIM nDims AS UINT = SafeArrayGetDim(m_psa)
   IF nDims <> 1 THEN RETURN E_FAIL
   IF this.IsResizable = FALSE THEN RETURN E_FAIL
   DIM vt AS VARTYPE = this.GetType
   IF vt <> VT_VARIANT THEN RETURN E_FAIL
   DIM cElements AS DWORD = this.Count(1) + 1
   DIM lLBound AS LONG = this.LBound(1)
   DIM sanewbounds(0) AS SAFEARRAYBOUND = {(cElements, lLBound)}
   DIM hr AS HRESULT = SafeArrayRedim(m_psa, @sanewbounds(0))
   IF hr THEN RETURN hr
   DIM idx AS LONG = cElements - 1 + lLBound
   RETURN SafeArrayPutElement(m_psa, @idx, cvData)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CSafeArray.Append (BYREF cvData AS CVAR) AS HRESULT
   RETURN this.AppendElement(cvData)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CSafeArray.AppendElement (BYVAL vData AS VARIANT) AS HRESULT
   CSAFEARRAY_DP("CSafeArray AppendElement - CVAR")
   IF m_psa = NULL THEN RETURN E_FAIL
   DIM nDims AS UINT = SafeArrayGetDim(m_psa)
   IF nDims <> 1 THEN RETURN E_FAIL
   IF this.IsResizable = FALSE THEN RETURN E_FAIL
   DIM vt AS VARTYPE = this.GetType
   IF vt <> VT_VARIANT THEN RETURN E_FAIL
   DIM cElements AS DWORD = this.Count(1) + 1
   DIM lLBound AS LONG = this.LBound(1)
   DIM sanewbounds(0) AS SAFEARRAYBOUND = {(cElements, lLBound)}
   DIM hr AS HRESULT = SafeArrayRedim(m_psa, @sanewbounds(0))
   IF hr THEN RETURN hr
   DIM idx AS LONG = cElements - 1 + lLBound
   RETURN SafeArrayPutElement(m_psa, @idx, @vData)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CSafeArray.Append (BYVAL vData AS VARIANT) AS HRESULT
   RETURN this.AppendElement(vData)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Inserts a value in the specified position of the one-dimensional safe array.
' If the safe array is not a one-dimensional array or it is not resizable it will return E_FAIL.
' Return value:
'   S_OK Success.
'   DISP_E_BADINDEX The specified index is not valid.
'   E_INVALIDARG One of the arguments is not valid.
'   E_OUTOFMEMORY Memory could not be allocated for the element.
'   E_FAIL The item pointed to by m_psa is not a safe array descriptor.
'          It is a fixed-size array.
'          It is not a one-dimensional array.
' - nPos = Index of the array element where the new string will be inserted.
' - pData = Pointer to the data to insert.
' ========================================================================================
PRIVATE FUNCTION CSafeArray.InsertElement (BYVAL nPos AS LONG, BYVAL pData AS ANY PTR) AS HRESULT
   CSAFEARRAY_DP("CSafeArray InsertElement - ANY PTR")
   IF m_psa = NULL THEN RETURN E_FAIL
   IF this.Count = 0 AND nPos = this.LBound THEN RETURN this.AppendElement(pData)
   DIM nDims AS UINT = SafeArrayGetDim(m_psa)
   IF nDims <> 1 THEN RETURN E_FAIL
   IF this.IsResizable = FALSE THEN RETURN E_FAIL
   IF nPos < this.LBound OR nPos > this.UBound THEN RETURN DISP_E_BADINDEX
   DIM cElements AS DWORD = this.Count(1) + 1
   DIM lLBound AS LONG = this.LBound(1)
   DIM sanewbounds(0) AS SAFEARRAYBOUND = {(cElements, lLBound)}
   DIM hr AS HRESULT = SafeArrayRedim(m_psa, @sanewbounds(0))
   IF hr THEN RETURN hr
   ' // Move all the elements up to nPos down
   DIM pvData AS ANY PTR, idx AS LONG
   FOR i AS LONG = cElements TO nPos STEP - 1
      idx = i - 1
      hr = SafeArrayGetElement(m_psa, @idx, @pvData)
      IF hr = S_OK THEN
         idx = i
         hr = SafeArrayPutElement(m_psa, @idx, @pvData)
      END IF
   NEXT
   ' // Set the value in the specified position
   idx = nPos
   RETURN SafeArrayPutElement(m_psa, @idx, pData)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CSafeArray.InsertElement (BYVAL pData AS ANY PTR) AS HRESULT
   RETURN this.InsertElement(this.LBound, pData)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CSafeArray.Insert (BYVAL nPos AS LONG, BYVAL pData AS ANY PTR) AS HRESULT
   RETURN this.InsertElement(nPos, pData)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CSafeArray.Insert (BYVAL pData AS ANY PTR) AS HRESULT
   RETURN this.InsertElement(this.LBound, pData)
END FUNCTION
' ========================================================================================

' ========================================================================================
PRIVATE FUNCTION CSafeArray.InsertElement (BYVAL nPos AS LONG, BYREF cbsData AS CBSTR) AS HRESULT
   CSAFEARRAY_DP("CSafeArray InsertElement - CBSTR")
   IF m_psa = NULL THEN RETURN E_FAIL
   IF this.Count = 0 AND nPos = this.LBound THEN RETURN this.AppendElement(cbsData)
   DIM nDims AS UINT = SafeArrayGetDim(m_psa)
   IF nDims <> 1 THEN RETURN E_FAIL
   IF this.IsResizable = FALSE THEN RETURN E_FAIL
   DIM vt AS VARTYPE = this.GetType
   IF vt <> VT_BSTR THEN RETURN E_FAIL
   IF nPos < this.LBound OR nPos > this.UBound THEN RETURN DISP_E_BADINDEX
   DIM cElements AS DWORD = this.Count(1) + 1
   DIM lLBound AS LONG = this.LBound(1)
   DIM sanewbounds(0) AS SAFEARRAYBOUND = {(cElements, lLBound)}
   DIM hr AS HRESULT = SafeArrayRedim(m_psa, @sanewbounds(0))
   IF hr THEN RETURN hr
   ' // Move all the elements down
   DIM cElem AS LONG = nPos - this.LBound
   DIM pvData AS AFX_BSTR PTR = this.AccessData
   IF pvData THEN
      ' // Move all the elements from nPos down
      FOR i AS LONG = cElements - 1 TO cElem + 1 STEP - 1
         pvData[i] = pvData[i - 1]
      NEXT
      pvData[cElem] = SysAllocString(**cbsData)
      this.UnaccessData
   END IF
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CSafeArray.InsertElement (BYREF cbsData AS CBSTR) AS HRESULT
   RETURN this.InsertElement(this.LBound, cbsData)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CSafeArray.Insert (BYVAL nPos AS LONG, BYREF cbsData AS CBSTR) AS HRESULT
   RETURN this.InsertElement(nPos, cbsData)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CSafeArray.Insert (BYREF cbsData AS CBSTR) AS HRESULT
   RETURN this.InsertElement(this.LBound, cbsData)
END FUNCTION
' ========================================================================================

' ========================================================================================
PRIVATE FUNCTION CSafeArray.InsertStr (BYVAL nPos AS LONG, BYVAL pwszData AS WSTRING PTR) AS HRESULT
   CSAFEARRAY_DP("CSafeArray InsertStr - WSTRING")
   IF m_psa = NULL THEN RETURN E_FAIL
   IF this.Count = 0 AND nPos = this.LBound THEN RETURN this.AppendStr(pwszData)
   DIM nDims AS UINT = SafeArrayGetDim(m_psa)
   IF nDims <> 1 THEN RETURN E_FAIL
   IF this.IsResizable = FALSE THEN RETURN E_FAIL
   DIM vt AS VARTYPE = this.GetType
   IF vt <> VT_BSTR THEN RETURN E_FAIL
   IF nPos < this.LBound OR nPos > this.UBound THEN RETURN DISP_E_BADINDEX
   DIM cElements AS DWORD = this.Count(1) + 1
   DIM lLBound AS LONG = this.LBound(1)
   DIM sanewbounds(0) AS SAFEARRAYBOUND = {(cElements, lLBound)}
   DIM hr AS HRESULT = SafeArrayRedim(m_psa, @sanewbounds(0))
   IF hr THEN RETURN hr
   ' // Move all the elements down
   DIM cElem AS LONG = nPos - this.LBound
   DIM pvData AS AFX_BSTR PTR = this.AccessData
   IF pvData THEN
      ' // Move all the elements from nPos down
      FOR i AS LONG = cElements - 1 TO cElem + 1 STEP - 1
         pvData[i] = pvData[i - 1]
      NEXT
      pvData[cElem] = SysAllocString(pwszData)
      this.UnaccessData
   END IF
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CSafeArray.InsertStr (BYVAL pwszData AS WSTRING PTR) AS HRESULT
   RETURN this.InsertStr(this.LBound, pwszData)
END FUNCTION
' ========================================================================================

' ========================================================================================
PRIVATE FUNCTION CSafeArray.InsertElement (BYVAL nPos AS LONG, BYREF cvData AS CVAR) AS HRESULT
   CSAFEARRAY_DP("CSafeArray InsertElement - CVAR")
   IF m_psa = NULL THEN RETURN E_FAIL
   IF this.Count = 0 AND nPos = this.LBound THEN RETURN this.AppendElement(cvData)
   DIM nDims AS UINT = SafeArrayGetDim(m_psa)
   IF nDims <> 1 THEN RETURN E_FAIL
   IF this.IsResizable = FALSE THEN RETURN E_FAIL
   DIM vt AS VARTYPE = this.GetType
   IF vt <> VT_VARIANT THEN RETURN E_FAIL
   IF nPos < this.LBound OR nPos > this.UBound THEN RETURN DISP_E_BADINDEX
   DIM cElements AS DWORD = this.Count(1) + 1
   DIM lLBound AS LONG = this.LBound(1)
   DIM sanewbounds(0) AS SAFEARRAYBOUND = {(cElements, lLBound)}
   DIM hr AS HRESULT = SafeArrayRedim(m_psa, @sanewbounds(0))
   IF hr THEN RETURN hr
   ' // Move all the elements down
   DIM cElem AS LONG = nPos - this.LBound
   DIM pvData AS VARIANT PTR = this.AccessData
   IF pvData THEN
      ' // Move all the elements from nPos down
      FOR i AS LONG = cElements - 1 TO cElem + 1 STEP - 1
         pvData[i] = pvData[i - 1]
      NEXT
      ' // Mark it as empty to avoid that VariantClear,
      ' // called by VariantCopy, will free it.
      pvData[cElem].vt = VT_EMPTY
      VariantCopy(@pvData[cElem], cvData)
      this.UnaccessData
   END IF
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CSafeArray.InsertElement (BYREF cvData AS CVAR) AS HRESULT
   RETURN this.InsertElement(this.LBound, cvData)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CSafeArray.Insert (BYVAL nPos AS LONG, BYREF cvData AS CVAR) AS HRESULT
   RETURN this.InsertElement(nPos, cvData)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CSafeArray.Insert (BYREF cvData AS CVAR) AS HRESULT
   RETURN this.InsertElement(this.LBound, cvData)
END FUNCTION
' ========================================================================================

' ========================================================================================
PRIVATE FUNCTION CSafeArray.InsertElement (BYVAL nPos AS LONG, BYVAL vData AS VARIANT) AS HRESULT
   CSAFEARRAY_DP("CSafeArray InsertElement - CVAR")
   IF m_psa = NULL THEN RETURN E_FAIL
   IF this.Count = 0 AND nPos = this.LBound THEN RETURN this.AppendElement(vData)
   DIM nDims AS UINT = SafeArrayGetDim(m_psa)
   IF nDims <> 1 THEN RETURN E_FAIL
   IF this.IsResizable = FALSE THEN RETURN E_FAIL
   DIM vt AS VARTYPE = this.GetType
   IF vt <> VT_VARIANT THEN RETURN E_FAIL
   IF nPos < this.LBound OR nPos > this.UBound THEN RETURN DISP_E_BADINDEX
   DIM cElements AS DWORD = this.Count(1) + 1
   DIM lLBound AS LONG = this.LBound(1)
   DIM sanewbounds(0) AS SAFEARRAYBOUND = {(cElements, lLBound)}
   DIM hr AS HRESULT = SafeArrayRedim(m_psa, @sanewbounds(0))
   IF hr THEN RETURN hr
   ' // Move all the elements down
   DIM cElem AS LONG = nPos - this.LBound
   DIM pvData AS VARIANT PTR = this.AccessData
   IF pvData THEN
      ' // Move all the elements from nPos down
      FOR i AS LONG = cElements - 1 TO cElem + 1 STEP - 1
         pvData[i] = pvData[i - 1]
      NEXT
      ' // Mark it as empty to avoid that VariantClear,
      ' // called by VariantCopy, will free it.
      pvData[cElem].vt = VT_EMPTY
      VariantCopy(@pvData[cElem], @vData)
      this.UnaccessData
   END IF
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CSafeArray.InsertElement (BYVAL vData AS VARIANT) AS HRESULT
   RETURN this.InsertElement(this.LBound, vData)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CSafeArray.Insert (BYVAL nPos AS LONG, BYVAL vData AS VARIANT) AS HRESULT
   RETURN this.InsertElement(nPos, vData)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CSafeArray.Insert (BYVAL vData AS VARIANT) AS HRESULT
   RETURN this.InsertElement(this.LBound, vData)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Appends a value to the end of the one-dimensional safe array.
' If the safe array is not a one-dimensional array or it is not resizable it will return E_FAIL.
' Return value:
'   S_OK Success.
'   DISP_E_BADINDEX The specified index is not valid.
'   E_INVALIDARG One of the arguments is not valid.
'   E_OUTOFMEMORY Memory could not be allocated for the element.
'   E_FAIL The item pointed to by m_psa is not a safe array descriptor.
'          It is a fixed-size array.
'          It is not a one-dimensional array.
' ========================================================================================
PRIVATE FUNCTION CSafeArray.DeleteElement (BYVAL nPos AS LONG) AS HRESULT
   CSAFEARRAY_DP("CSafeArray DeleteElement")
   IF m_psa = NULL THEN RETURN E_FAIL
   DIM nDims AS UINT = SafeArrayGetDim(m_psa)
   IF nDims <> 1 THEN RETURN E_FAIL
   IF this.IsResizable = FALSE THEN RETURN E_FAIL
   IF nPos < this.LBound OR nPos > this.UBound THEN RETURN E_FAIL
   DIM cElements AS DWORD = this.Count(1)
   ' // Save the element to be deleted
   DIM idx AS LONG, pTemp AS ANY PTR
   idx = nPos
   DIM hr AS HRESULT = SafeArrayGetElement(m_psa, @idx, @pTemp)
   IF hr THEN RETURN hr
   ' // Move all the elements from nPos + 1 up
   DIM pvData AS ANY PTR
   FOR i AS LONG = nPos TO cElements - 1 STEP 1
      idx = i + 1
      hr = SafeArrayGetElement(m_psa, @idx, @pvData)
      IF hr = S_OK THEN
         idx = i
         hr = SafeArrayPutElement(m_psa, @idx, @pvData)
      END IF
   NEXT
   ' // Copy the element to be deleted to the end of the array
   idx = cElements
   hr = SafeArrayPutElement(m_psa, @idx, @pTemp)
   ' // Shrink the array by one element (will free the last element)
   DIM lLBound AS LONG = this.LBound(1)
   DIM sanewbounds(0) AS SAFEARRAYBOUND = {(cElements - 1, lLBound)}
   RETURN SafeArrayRedim(m_psa, @sanewbounds(0))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CSafeArray.Remove (BYVAL nPos AS LONG) AS HRESULT
   RETURN this.DeleteElement(nPos)
END FUNCTION
' ========================================================================================

' ========================================================================================
PRIVATE FUNCTION CSafeArray.DeleteStringElement (BYVAL nPos AS LONG) AS HRESULT
   CSAFEARRAY_DP("CSafeArray DeleteStringElement")
   IF m_psa = NULL THEN RETURN E_FAIL
   DIM nDims AS UINT = SafeArrayGetDim(m_psa)
   IF nDims <> 1 THEN RETURN E_FAIL
   DIM vt AS VARTYPE = this.GetType
   IF vt <> VT_BSTR THEN RETURN E_FAIL
   IF this.IsResizable = FALSE THEN RETURN E_FAIL
   IF nPos < this.LBound OR nPos > this.UBound THEN RETURN E_FAIL
   DIM cElem AS LONG = nPos - this.LBound
   DIM cElements AS DWORD = this.Count(1)
   DIM pvData AS AFX_BSTR PTR = this.AccessData
   IF pvData THEN
      ' // Save the element to be deleted
      DIM pTemp AS AFX_BSTR = pvData[cElem]
      ' // Move all the elements from nPos + 1 up
      FOR i AS LONG = cElem TO cElements - 1 STEP 1
         pvData[i] = pvData[i + 1]
      NEXT
      ' // Copy the element to be deleted to the end of the array
      pvData[cElements - 1] = pTemp
      this.UnaccessData
   END IF
   ' // Shrink the array by one element (will free the last element)
   DIM lLBound AS LONG = this.LBound(1)
   DIM sanewbounds(0) AS SAFEARRAYBOUND = {(cElements - 1, lLBound)}
   RETURN SafeArrayRedim(m_psa, @sanewbounds(0))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CSafeArray.RemoveStr (BYVAL nPos AS LONG) AS HRESULT
   RETURN this.DeleteStringElement(nPos)
END FUNCTION
' ========================================================================================

' ========================================================================================
PRIVATE FUNCTION CSafeArray.DeleteVariantElement (BYVAL nPos AS LONG) AS HRESULT
   CSAFEARRAY_DP("CSafeArray DeleteVariantElement")
   IF m_psa = NULL THEN RETURN E_FAIL
   DIM nDims AS UINT = SafeArrayGetDim(m_psa)
   IF nDims <> 1 THEN RETURN E_FAIL
   DIM vt AS VARTYPE = this.GetType
   IF vt <> VT_VARIANT THEN RETURN E_FAIL
   IF this.IsResizable = FALSE THEN RETURN E_FAIL
   IF nPos < this.LBound OR nPos > this.UBound THEN RETURN E_FAIL
   DIM cElem AS LONG = nPos - this.LBound
   DIM cElements AS DWORD = this.Count(1)
   DIM pvData AS VARIANT PTR = this.AccessData
   IF pvData THEN
      ' // Save the element to be deleted
      DIM pTemp AS VARIANT = pvData[cElem]
      ' // Move all the elements from nPos + 1 up
      FOR i AS LONG = cElem TO cElements - 1 STEP 1
         pvData[i] = pvData[i + 1]
      NEXT
      ' // Copy the element to be deleted to the end of the array
      pvData[cElements - 1] = pTemp
      this.UnaccessData
   END IF
   ' // Shrink the array by one element (will free the last element)
   DIM lLBound AS LONG = this.LBound(1)
   DIM sanewbounds(0) AS SAFEARRAYBOUND = {(cElements - 1, lLBound)}
   RETURN SafeArrayRedim(m_psa, @sanewbounds(0))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CSafeArray.RemoveVar (BYVAL nPos AS LONG) AS HRESULT
   RETURN this.DeleteVariantElement(nPos)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Scans the array to search for the specified string.
' - wszFind = The string to search.
' - bNoCase = TRUE: Ignore case.
' ========================================================================================
PRIVATE FUNCTION CSafeArray.FindElement (BYREF wszFind AS WSTRING, BYVAL bNoCase AS BOOLEAN = FALSE) AS LONG
   CSAFEARRAY_DP("CSafeArray FindElement")
   IF m_psa = NULL THEN RETURN 0
   DIM nDims AS UINT = SafeArrayGetDim(m_psa)
   IF nDims <> 1 THEN RETURN 0
   DIM vt AS VARTYPE = this.GetType
   IF vt <> VT_BSTR THEN RETURN 0
   DIM cwsFind AS CBSTR = IIF(bNoCase = FALSE, wszFind, UCASE(wszFind))
   DIM pvData AS AFX_BSTR PTR = this.AccessData
   IF pvData THEN
      FOR i AS LONG = 0 TO this.UBound - this.LBound
         IF IIF(bNoCase = FALSE, *pvData[i], UCASE(*pvData[i])) = **cwsFind THEN
            FUNCTION = i + this.LBound
            EXIT FOR
         END IF
      NEXT
      this.UnaccessData
   END IF
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CSafeArray.Find (BYREF wszFind AS WSTRING, BYVAL bNoCase AS BOOLEAN = FALSE) AS LONG
   RETURN this.FindElement(wszFind, bNoCase)
END FUNCTION
' ========================================================================================

' ========================================================================================
' qsort CSafeArray comparison function
' ========================================================================================
PRIVATE FUNCTION AfxCSafeArrayCompare CDECL (BYVAL a AS AFX_BSTR PTR, BYVAL b AS AFX_BSTR PTR) AS LONG
   FUNCTION = wcscmp(*a, *b)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Reverse qsort CSafeArray comparison function
' ========================================================================================
PRIVATE FUNCTION AfxCSafeArrayReverseCompare CDECL (BYVAL a AS AFX_BSTR PTR, BYVAL b AS AFX_BSTR PTR) AS LONG
   DIM r AS LONG = wcscmp(*a, *b)
   IF r = 1 THEN r = -1 ELSE IF r = -1 THEN r = 1
   RETURN r
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sorts a one-dimensional VT_BSTR CSafeArray calling the C qsort function.
' Parameter:
' - bAscend: TRUE for sorting in ascending order; FALSE for sorting in descending order.
' ========================================================================================
PRIVATE FUNCTION CSafeArray.Sort (BYVAL bAscend AS BOOLEAN = TRUE) AS HRESULT
   CSAFEARRAY_DP("CSafeArray Sort")
   IF m_psa = NULL THEN RETURN E_FAIL
   DIM nDims AS UINT = SafeArrayGetDim(m_psa)
   IF nDims <> 1 THEN RETURN E_FAIL
   DIM vt AS VARTYPE = this.GetType
   IF vt <> VT_BSTR THEN RETURN E_FAIL
   DIM pvData AS AFX_BSTR PTR = this.AccessData
   IF pvData = NULL THEN RETURN E_UNEXPECTED
   IF bAscend THEN
      qsort pvData, this.Count, SIZEOF(AFX_BSTR), CPTR(ANY PTR, @AfxCSafeArrayCompare)
   ELSE
      qsort pvData, this.Count, SIZEOF(AFX_BSTR), CPTR(ANY PTR, @AfxCSafeArrayReverseCompare)
   END IF
   IF pvData THEN this.UnaccessData
END FUNCTION
' ========================================================================================

' ========================================================================================
'                         *** Safe array related functions ***
' ========================================================================================

' ========================================================================================
' Splits a string into tokens, which are sequences of contiguous characters separated by
' any of the characters that are part of delimiters.
' - wszStr = The string to split.
' - wszDelimiters = The delimiter characters.
' Return value: A CSafeArray containing a token in each element.
' Usage example:
' DIM wsz AS WSTRING * 260 = "- This, a sample string."
' DIM csa AS CSafeArray = AfxStrSplit(wsz, " ,.-")
' FOR i AS LONG = csa.LBound TO csa.UBound
'    DIM cbsOut AS CBSTR
'    csa.GetElement(i, cbsOut)
'    print cbsOut
' NEXT
' -or-
' FOR i AS LONG = csa.LBound TO csa.UBound
'    print csa.GetStr(i)
' NEXT
' ========================================================================================
PRIVATE FUNCTION AfxStrSplit (BYREF wszStr AS WSTRING, BYREF wszDelimiters AS WSTRING) AS CSafeArray
   DIM cws AS CWSTR = wszStr
   DIM csa AS CSafeArray = CSafeArray(VT_BSTR, 0, 1)
   DIM pwsz AS WSTRING PTR = wcstok(cws, @wszDelimiters)
   WHILE pwsz <> NULL
      csa.AppendStr(pwsz)
      pwsz = wcstok(NULL, @wszDelimiters)
   WEND
   RETURN csa
END FUNCTION
' ========================================================================================


' ========================================================================================
' Returns a string consisting of all of the strings in an array, each separated by a delimiter.
' If the delimiter is a null (zero-length) string then no separators are inserted between
' the string sections. If the delimiter expression is the 3-byte value of "," which may be
' expressed in your source code as the string literal """,""" or as Chr(34,44,34) then a
' leading and trailing double-quote is added to each string section. This ensures that the
' returned string contains standard comma-delimited quoted fields that can be easily parsed.
' Usage example:
' DIM csa AS CSafeArray = CSafeArray(VT_BSTR, 3, 1)
' csa.PutElement(1, CBSTR("One"))
' csa.PutElement(2, CBSTR("Two"))
' csa.PutElement(3, CBSTR("Three"))
' DIM cws AS CWSTR = AfxStrJoin(csa, ",")
' PRINT cws   ' ouput: One,Two,Three
' ========================================================================================
PRIVATE FUNCTION AfxStrJoin (BYREF csa AS CSafeArray, BYREF wszDelimiter AS WSTRING) AS CWSTR
   DIM cws AS CWSTR
   IF csa.NumDims <> 1 THEN RETURN cws
   IF csa.GetType <> VT_BSTR THEN RETURN cws
   ' // Use direct access for speed
   DIM nCount AS LONG = csa.Count
   IF nCount = 0 THEN RETURN cws
   DIM pvData AS AFX_BSTR PTR = csa.AccessData
   IF pvData = NULL THEN RETURN cws
   ' // Add a leading ""
   IF wszDelimiter = CHR(34, 44, 34) THEN cws = CHR(34)
   FOR i AS LONG = 0 TO nCount - 1
      cws += *pvData[i]
      IF i <> nCount - 1 AND wszDelimiter <> "" THEN cws += wszDelimiter
   NEXT
   csa.UnaccessData
   ' // Add a trailing ""
   IF wszDelimiter = CHR(34, 44, 34) THEN cws += CHR(34)
   RETURN cws
END FUNCTION
' ========================================================================================

' ========================================================================================
' Base64 is a group of similar encoding schemes that represent binary data in an ASCII
' string format by translating it into a radix-64 representation. The Base64 term
' originates from a specific MIME content transfer encoding.
' Base64 encoding schemes are commonly used when there is a need to encode binary data
' that needs be stored and transferred over media that are designed to deal with textual
' data. This is to ensure that the data remains intact without modification during
' transport. Base64 is used commonly in a number of applications including email via MIME,
' and storing complex data in XML.
' Note: For a non COM version, see AfxBase64Encode in AfxStr.inc
' ========================================================================================

' ========================================================================================
' Base64 mime encoding
' ========================================================================================
PRIVATE FUNCTION AfxXmlBase64Encode (BYREF strData AS STRING) AS STRING

   ' // Create a byte safe array
   DIM rgsabound AS SAFEARRAYBOUND
   rgsabound.cElements = LEN(strData)
   rgsabound.lLBound = 0
   DIM psa AS SAFEARRAY PTR
   psa = SafeArrayCreate(VT_UI1, 1, @rgsabound)
   IF psa = NULL THEN EXIT FUNCTION
   ' // Lock the safearray for access
   DIM pvData AS ANY PTR
   SafeArrayAccessData(psa, @pvData)
   IF pvData THEN
      ' // copy the contents of the string
      memcpy pvData, STRPTR(strData), LEN(strData)
      ' // Unlock the safe array
      SafeArrayUnaccessData(psa)
   END IF

   ' // Create a variant to host the safe array
   DIM vData AS VARIANT
   vData.vt = VT_ARRAY OR VT_UI1
   vData.parray = psa

   ' // Create an instance of the IXMLDOMDocument interface
   DIM pXmlDocument AS IXMLDOMDocument PTR
   CoCreateInstance(@CLSID_DOMDocument, NULL, CLSCTX_INPROC_SERVER, @IID_IXMLDOMDocument, @pXmlDocument)
   IF pXmlDocument THEN
      ' // Create a "b64" element node
      DIM pXmlElement AS IXMLDOMElement PTR
      DIM pbstrTagName AS AFX_BSTR = SysAllocString("b64")
      IF pbstrTagName THEN
         pXmlDocument->lpvtbl->createElement(pXmlDocument, pbstrTagName, @pXmlElement)
         SysFreeString pbstrTagName
      END IF
      ' // Set the data type to binary base 64 encoded
      IF pXmlElement THEN
         DIM pDataType AS AFX_BSTR = SysAllocString("bin.base64")
         IF pDataType THEN
            pXmlElement->lpvtbl->put_dataType(pXmlElement, pDataType)
            SysFreeString pDataType
         END IF
         ' // Put the data in the node
         pXmlElement->lpvtbl->put_nodeTypedValue(pXmlElement, vData)
         ' // Get the data as text
         DIM pbstrText AS AFX_BSTR
         pXmlElement->lpvtbl->get_Text(pXmlElement, @pbstrText)
         ' // Return the base 64 encoded text
         IF pbstrText THEN
            FUNCTION = *pbstrText
            SysFreeString pbstrText
         END IF
         IUnknown_Release(pXmlElement)
      END IF
      IUnknown_Release(pXmlDocument)
   END IF

   ' // Clear the variant
   VariantClear @vData

END FUNCTION
' ========================================================================================

' ========================================================================================
' Base64 mime decoding
' ========================================================================================
PRIVATE FUNCTION AfxXmlBase64Decode (BYREF strData AS STRING) AS STRING

   ' // Create an instance of the IXMLDOMDocument interface
   DIM pXmlDocument AS IXMLDOMDocument PTR
   CoCreateInstance(@CLSID_DOMDocument, NULL, CLSCTX_INPROC_SERVER, @IID_IXMLDOMDocument, @pXmlDocument)
   IF pXmlDocument THEN
      ' // Create a "b64" element node
      DIM pXmlElement AS IXMLDOMElement PTR
      DIM pbstrTagName AS AFX_BSTR = SysAllocString("b64")
      IF pbstrTagName THEN
         pXmlDocument->lpvtbl->createElement(pXmlDocument, pbstrTagName, @pXmlElement)
         SysFreeString pbstrTagName
      END IF
      ' // Set the data type to binary base 64 encoded
      IF pXmlElement THEN
         DIM pDataType AS AFX_BSTR = SysAllocString("bin.base64")
         IF pDataType THEN
            pXmlElement->lpvtbl->put_dataType(pXmlElement, pDataType)
            SysFreeString pDataType
         END IF
         ' // Put the data as text
         DIM pbstrText AS AFX_BSTR = SysAllocString(strData)
         IF pbstrText THEN
            pXmlElement->lpvtbl->put_Text(pXmlElement, pbstrText)
            SysFreeString pbstrText
         END IF
         ' // Get the data as a byte safe array
         DIM vData AS VARIANT
         pXmlElement->lpvtbl->get_nodeTypedValue(pXmlElement, @vData)
         IF vData.parray THEN
            ' // The number of dimensions must be 1
            IF SafeArrayGetDim(vData.parray) = 1 THEN
               ' // Retrieve the number of elements of the array
               DIM nLBound AS LONG, nUBound AS LONG
               SafeArrayGetLBound(vData.parray, 1, @nLBound)
               SafeArrayGetUBound(vData.parray, 1, @nUBound)
               ' // Calculate the number of bytes to read
               DIM nBytes AS LONG = nUbound - nLBound + 1
               IF nBytes THEN
                  ' // Lock the safearray for access
                  DIM pvData AS ANY PTR
                  SafeArrayAccessData(vData.parray, @pvData)
                  IF pvData THEN
                     ' // Read the data
                     DIM buffer AS STRING = SPACE(nBytes)
                     memcpy STRPTR(buffer), pvData, nBytes
                     ' // Unlock the array
                     SafeArrayUnaccessData(vData.parray)
                     ' // Return the data
                     FUNCTION = buffer
                  END IF
               END IF
            END IF
         END IF
         VariantClear @vData
         IUnknown_Release(pXmlElement)
      END IF
      IUnknown_Release(pXmlDocument)
   END IF

END FUNCTION
' ========================================================================================

' ========================================================================================
' Reads all the lines of the specified file into a safe array.
' - wszFileName: Path of the file
' - szDelimiter: Delimiter of the line (CHR(13, 10) in Windows, CHR(10) in Linmux).
' ========================================================================================
PRIVATE FUNCTION AfxFileReadAllLinesA (BYREF wszFileName AS WSTRING, BYREF szDelimiter AS ZSTRING = CHR(13, 10)) AS CSafeArray
   DIM _csa AS CSafeArray = CSafeArray(VT_BSTR, 0, 1)
   DIM dwCount AS DWORD, dwFileSize AS DWORD, dwHighSize AS DWORD, dwBytesRead AS DWORD
   IF LEN(szDelimiter) = 0 THEN RETURN _csa
   ' // Open the file
   DIM hFile AS HANDLE = CreateFileW(@wszFileName, GENERIC_READ, FILE_SHARE_READ, NULL, _
                         OPEN_EXISTING, FILE_FLAG_SEQUENTIAL_SCAN, NULL)
   IF hFile = INVALID_HANDLE_VALUE THEN RETURN _csa
   ' // Get the size of the file
   dwFileSize = GetFileSize(hFile, @dwHighSize)
   DIM pBuffer AS UBYTE PTR
   pBuffer = CAllocate(1, dwFileSize)
   IF pBuffer = NULL THEN RETURN _csa
   DIM bSuccess AS LONG = ReadFile(hFile, pBuffer, dwFileSize, @dwBytesRead, NULL)
   CloseHandle(hFile)
   IF bSuccess = FALSE THEN RETURN _csa
   ' // Get the number of lines
   DIM nLen AS LONG = LEN(szDelimiter)
   DIM pstr AS ANY PTR = pBuffer
   DO
      pstr = strstr(pstr, szDelimiter)
      IF pstr = NULL THEN EXIT DO
      pstr += nLen
      dwCount += 1
   LOOP
' -------------------------------------------------------------------
  ' // Dimension a safe array with dwCount elements
'   DIM csa AS CSafeArray = CSafeArray(VT_BSTR, dwCount, 1)
'   Note: strtok can't be used because it skips empty lines.
'   ' // Fill the array with the lines
'   DIM pwsz AS ZSTRING PTR = strtok(pBuffer, @szDelimiter)
'   DIM idx AS LONG = 1
'   WHILE pwsz <> NULL
'      IF idx < dwCount THEN csa.PutElement(idx, CBSTR(*pwsz))
'      idx += 1
'      pwsz = strtok(NULL, @szDelimiter)
'   WEND
' -------------------------------------------------------------------
  ' // Dimension a safe array with dwCount elements
   DIM csa AS CSafeArray = CSafeArray(VT_BSTR, dwCount, 1)
   DIM s AS STRING, sLen AS LONG
   DIM idx AS LONG = 1
   DIM _pstr AS ANY PTR = pBuffer
   pstr = pBuffer
   ' // Check for UTF-8 BOM
   s = "   "
   strncpy STRPTR(s), _pstr, 3
   IF s = CHR(&hEF, &hBB, &hBF) THEN _pstr += 3
   ' // Parse the buffer
   DO
      pstr = strstr(pstr, szDelimiter)
      IF pstr = NULL THEN EXIT DO
      sLen = pstr - _pstr
      s = ""
      IF sLen > nLen THEN
         s = STRING(sLen, CHR(0))
         strncpy STRPTR(s), _pstr, sLen
      END IF
      IF idx <= dwCount THEN csa.PutElement(idx, CBSTR(s))
      idx += 1
      pstr += nLen
      _pstr = pstr
   LOOP
   DeAllocate(pBuffer)
   RETURN csa
END FUNCTION
' ========================================================================================

' ========================================================================================
' Reads all the lines of the specified file into a safe array.
' - wszFileName: Path of the file
' - szDelimiter: Delimiter of the line (CHR(13, 10) in Windows, CHR(10) in Linmux).
' ========================================================================================
PRIVATE FUNCTION AfxFileReadAllLinesW (BYREF wszFileName AS WSTRING, BYREF wszDelimiter AS WSTRING = CHR(13, 10)) AS CSafeArray
   DIM _csa AS CSafeArray = CSafeArray(VT_BSTR, 0, 1)
   DIM dwCount AS DWORD, dwFileSize AS DWORD, dwHighSize AS DWORD, dwBytesRead AS DWORD
   IF LEN(wszDelimiter) = 0 THEN RETURN _csa
   ' // Open the file
   DIM hFile AS HANDLE = CreateFileW(@wszFileName, GENERIC_READ, FILE_SHARE_READ, NULL, _
                         OPEN_EXISTING, FILE_FLAG_SEQUENTIAL_SCAN, NULL)
   IF hFile = INVALID_HANDLE_VALUE THEN RETURN _csa
   ' // Get the size of the file
   dwFileSize = GetFileSize(hFile, @dwHighSize)
   DIM pBuffer AS UBYTE PTR
   pBuffer = CAllocate(1, dwFileSize)
   IF pBuffer = NULL THEN RETURN _csa
   DIM bSuccess AS LONG = ReadFile(hFile, pBuffer, dwFileSize, @dwBytesRead, NULL)
   CloseHandle(hFile)
   IF bSuccess = FALSE THEN RETURN _csa
   ' // Get the number of lines
   DIM nLen AS LONG = LEN(wszDelimiter) * 2
   DIM pstr AS ANY PTR = pBuffer
   DO
      pstr = wcsstr(pstr, @wszDelimiter)
      IF pstr = NULL THEN EXIT DO
      pstr += nLen
      dwCount += 1
   LOOP
  ' // Dimension a safe array with dwCount elements
   DIM csa AS CSafeArray = CSafeArray(VT_BSTR, dwCount, 1)
   DIM cws AS CWSTR, sLen AS LONG
   DIM idx AS LONG = 1
   DIM _pstr AS ANY PTR = pBuffer
   pstr = pBuffer
   ' / Skip the BOM
   DIM s AS STRING = "  "
   strncpy STRPTR(s), _pstr, 2
   ' // Check for UTF-16 BOM (little endian)
   IF s = CHR(&hFF, &hFE) THEN
      _pstr += 2
   ELSE
      ' // Check for UTF-8 BOM
      s = "   "
      strncpy STRPTR(s), _pstr, 3
      IF s = CHR(&hEF, &hBB, &hBF) THEN _pstr += 3
   END IF
   ' // Parse the buffer
   DO
      pstr = wcsstr(pstr, @wszDelimiter)
      IF pstr = NULL THEN EXIT DO
      sLen = (pstr - _pstr) \ 2
      cws = ""
      IF sLen > nLen THEN
         cws = STRING(sLen, CHR(0))
         wcsncpy cws, _pstr, sLen
      END IF
      IF idx <= dwCount THEN csa.PutElement(idx, CBSTR(cws))
      idx += 1
      pstr += nLen
      _pstr = pstr
   LOOP
   DeAllocate(pBuffer)
   RETURN csa
END FUNCTION
' ========================================================================================

END NAMESPACE
