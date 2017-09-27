' ########################################################################################
' Microsoft Windows
' File: CVar.inc
' Contents: Windows VARIANT class.
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "windows.bi"
#include once "crt/limits.bi"
#include once "Afx/AfxCOM.inc"

NAMESPACE Afx

' ========================================================================================
' Macro for debug
' To allow debugging, define _CVAR_DEBUG_ 1 in your application before including this file.
' ========================================================================================
#ifndef _CVAR_DEBUG_
   #define _CVAR_DEBUG_ 0
#ENDIF
#ifndef _CVAR_DP_
   #define _CVAR_DP_ 1
   #MACRO CVAR_DP(st)
      #IF (_CVAR_DEBUG_ = 1)
         OutputDebugStringW(st)
      #ENDIF
   #ENDMACRO
#ENDIF
' ========================================================================================

'#define AFX_BOOL VT_BOOL
'#define AFX_BYTE VT_I1
'#define AFX_UBYTE VT_UI1
'#define AFX_SHORT VT_I2
'#define AFX_USHORT VT_UI2
'#define AFX_INT VT_INT
'#define AFX_UINT VT_UINT
'#define AFX_LONG VT_I4
'#define AFX_ULONG VT_UI4
'#define AFX_LONGINT VT_I8
'#define AFX_ULONGINT VT_UI8
'#define AFX_SINGLE VT_R4
'#define AFX_FLOAT VT_R4
'#define AFX_DOUBLE VT_R8
'#define AFX_NULL VT_NULL
'#define AFX_ARRAY VT_ARRAY

' ########################################################################################
' CVar - VARIANT class
' ########################################################################################
TYPE CVar

Private:
   cws AS CWSTR          ' // To allow to return a WSTRING by reference

Public:
   vd AS VARIANT         ' // Variant data

   ' // COnstructors
   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR
   DECLARE CONSTRUCTOR (BYREF cv AS CVAR)
   DECLARE CONSTRUCTOR (BYVAL v AS VARIANT)
   DECLARE CONSTRUCTOR (BYREF wsz AS WSTRING)
   DECLARE CONSTRUCTOR (BYREF cws AS CWSTR)
   DECLARE CONSTRUCTOR (BYREF cbs AS CBSTR)
   DECLARE CONSTRUCTOR (BYVAL pvar AS VARIANT PTR)
   DECLARE CONSTRUCTOR (BYVAL cy AS CURRENCY)
   DECLARE CONSTRUCTOR (BYVAL dec AS DECIMAL)
   DECLARE CONSTRUCTOR (BYVAL pdisp AS IDispatch PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   DECLARE CONSTRUCTOR (BYVAL punk AS IUnknown PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   DECLARE CONSTRUCTOR (BYVAL _value AS LONGINT, BYVAL _vType AS WORD = VT_I4)
   DECLARE CONSTRUCTOR (BYVAL _value AS DOUBLE, BYVAL _vType AS WORD = VT_R8)
   DECLARE CONSTRUCTOR (BYVAL _value AS LONGINT, BYREF strType AS STRING)
   DECLARE CONSTRUCTOR (BYVAL _value AS DOUBLE, BYREF strType AS STRING)
   DECLARE CONSTRUCTOR (BYVAL _pvar AS ANY PTR, BYVAL _vType AS WORD)
   DECLARE CONSTRUCTOR (BYVAL _pvar AS ANY PTR, BYREF strType AS STRING)
   ' // Casting
'   DECLARE OPERATOR @ () AS VARIANT PTR
   DECLARE FUNCTION vptr () AS VARIANT PTR
   DECLARE FUNCTION sptr () AS VARIANT PTR
   DECLARE FUNCTION wstr () AS CWSTR
   DECLARE FUNCTION bstr () AS CBSTR
   DECLARE OPERATOR CAST () AS VARIANT
   DECLARE OPERATOR CAST () AS ANY PTR
   DECLARE OPERATOR CAST () BYREF AS WSTRING
   DECLARE FUNCTION ToStr () AS CWSTR
   DECLARE FUNCTION ToWStr () AS CWSTR
   DECLARE FUNCTION ToBStr () AS CBSTR
   DECLARE FUNCTION ToUtf8 () AS STRING
   DECLARE FUNCTION ToBuffer (BYVAL pv AS ANY PTR, BYVAL cb AS UINT) AS HRESULT
   DECLARE FUNCTION ToBuffer () AS STRING
   DECLARE FUNCTION ToUnknown () AS ANY PTR
   DECLARE FUNCTION ToDispatch () AS ANY PTR
   DECLARE FUNCTION DecToDouble () AS DOUBLE
   DECLARE FUNCTION DecToCy () AS CY
   DECLARE FUNCTION ToVbDate () AS DATE_
   DECLARE FUNCTION ToSystemTime () AS SYSTEMTIME
   DECLARE FUNCTION ToGuid () AS GUID
   DECLARE FUNCTION ToGuidStr () AS CWSTR
   DECLARE FUNCTION ToGuidWStr () AS CWSTR
   DECLARE FUNCTION ToGuidBStr () AS CBSTR
   DECLARE FUNCTION ToDosDateTime (BYVAL pwDate AS USHORT PTR, BYVAL pwTime AS USHORT PTR) AS HRESULT
   DECLARE FUNCTION ToFileTime (BYVAL stfOut AS AFX_PSTIME_FLAGS) AS FILETIME
   DECLARE FUNCTION ToStrRet () AS STRRET
   DECLARE FUNCTION ToBooleanArray (BYVAL pprg AS WINBOOL PTR, BYVAL crgn AS ULONG) AS ULONG
   DECLARE FUNCTION ToBooleanArrayAlloc (BYVAL pprgf AS WINBOOL PTR PTR) AS ULONG
   DECLARE FUNCTION ToShortArray (BYVAL prgn AS SHORT PTR, BYVAL crgn AS ULONG) AS ULONG
   DECLARE FUNCTION ToShortArrayAlloc (BYVAL pprgn AS SHORT PTR PTR) AS ULONG
   DECLARE FUNCTION ToUShortArray (BYVAL prgn AS USHORT PTR, BYVAL crgn AS ULONG) AS ULONG
   DECLARE FUNCTION ToUShortArrayAlloc (BYVAL pprgn AS USHORT PTR PTR) AS ULONG
   DECLARE FUNCTION ToLongArray (BYVAL prgn AS LONG PTR, BYVAL crgn AS ULONG) AS ULONG
   DECLARE FUNCTION ToLongArrayAlloc (BYVAL pprgn AS LONG PTR PTR) AS ULONG
   DECLARE FUNCTION ToULongArray (BYVAL prgn AS ULONG PTR, BYVAL crgn AS ULONG) AS ULONG
   DECLARE FUNCTION ToULongArrayAlloc (BYVAL pprgn AS ULONG PTR PTR) AS ULONG
   DECLARE FUNCTION ToLongIntArray (BYVAL prgn AS LONGINT PTR, BYVAL crgn AS ULONG) AS ULONG
   DECLARE FUNCTION ToLongIntArrayAlloc (BYVAL pprgn AS LONGINT PTR PTR) AS ULONG
   DECLARE FUNCTION ToULongIntArray (BYVAL prgn AS ULONGINT PTR, BYVAL crgn AS ULONG) AS ULONG
   DECLARE FUNCTION ToULongIntArrayAlloc (BYVAL pprgn AS ULONGINT PTR PTR) AS ULONG
   DECLARE FUNCTION ToDoubleArray (BYVAL prgn AS DOUBLE PTR, BYVAL crgn AS ULONG) AS ULONG
   DECLARE FUNCTION ToDoubleArrayAlloc (BYVAL pprgn AS DOUBLE PTR PTR) AS ULONG
   DECLARE FUNCTION ToStringArray (BYVAL prgsz AS PWSTR, BYVAL crgsz AS ULONG) AS ULONG
   DECLARE FUNCTION ToStringArrayAlloc (BYVAL pprgsz AS PWSTR PTR) AS ULONG
   ' // LET assignments
   DECLARE OPERATOR Let (BYREF cv AS CVAR)
   DECLARE OPERATOR Let (BYVAL v AS VARIANT)
   DECLARE OPERATOR Let (BYREF wszStr AS WSTRING)
   DECLARE OPERATOR Let (BYREF cws AS CWSTR)
   DECLARE OPERATOR Let (BYREF cbs AS CBSTR)
   DECLARE OPERATOR Let (BYVAL pvar AS VARIANT PTR)
   DECLARE OPERATOR Let (BYVAL cy AS CURRENCY)
   DECLARE OPERATOR Let (BYVAL dec AS DECIMAL)
   DECLARE OPERATOR Let (BYVAL pdisp AS IDispatch PTR)
   DECLARE OPERATOR Let (BYVAL punk AS IUnknown PTR)
   ' // Assignments
   DECLARE SUB Put (BYREF wszStr AS WSTRING)
   DECLARE SUB Put (BYREF cws AS CWSTR)
   DECLARE SUB Put (BYREF cbs AS CBSTR)
   DECLARE FUNCTION Put (BYVAL pdisp AS IDispatch PTR, BYVAL fAddRef AS BOOLEAN = FALSE) AS HRESULT
   DECLARE FUNCTION Put (BYVAL punk AS IUnknown PTR, BYVAL fAddRef AS BOOLEAN = FALSE) AS HRESULT
   DECLARE FUNCTION Put (BYREF cv AS CVAR) AS HRESULT
   DECLARE FUNCTION Put (BYREF v AS VARIANT) AS HRESULT
   DECLARE FUNCTION Put (BYVAL pvar AS VARIANT PTR) AS HRESULT
   DECLARE SUB Put (BYVAL _value AS LONGINT, BYVAL _vType AS WORD = VT_I4)
   DECLARE SUB Put (BYVAL _value AS LONGINT, BYREF strType AS STRING)
   DECLARE SUB Put (BYVAL _value AS DOUBLE, BYVAL _vType AS WORD = VT_R8)
   DECLARE SUB Put (BYVAL _value AS DOUBLE, BYREF strType AS STRING)
   DECLARE SUB PutNull
   DECLARE SUB PutBool (BYVAL _value AS BOOL)
   DECLARE SUB PutBoolean (BYVAL _value AS BOOLEAN)
   DECLARE SUB PutByte (BYVAL _value AS BYTE)
   DECLARE SUB PutUByte (BYVAL _value AS UBYTE)
   DECLARE SUB PutShort (BYVAL _value AS SHORT)
   DECLARE SUB PutUShort (BYVAL _value AS USHORT)
   DECLARE SUB PutInt (BYVAL _value AS INT_)
   DECLARE SUB PutUInt (BYVAL _value AS UINT)
   DECLARE SUB PutLong (BYVAL _value AS LONG)
   DECLARE SUB PutULong (BYVAL _value AS ULONG)
   DECLARE SUB PutLongInt (BYVAL _value AS LONGINT)
   DECLARE SUB PutULongInt (BYVAL _value AS ULONGINT)
   DECLARE SUB PutSingle (BYVAL _value AS SINGLE)
   DECLARE SUB PutFloat (BYVAL _value AS SINGLE)
   DECLARE SUB PutDouble (BYVAL _value AS DOUBLE)
   DECLARE FUNCTION PutBuffer(BYVAL pv AS ANY PTR, BYVAL cb AS UINT) AS HRESULT
   DECLARE FUNCTION PutUtf8 (BYREF strUtf8 AS STRING) AS HRESULT
   DECLARE FUNCTION PutSafeArray (BYVAL parray AS SAFEARRAY PTR, BYVAL fAttach AS BOOLEAN = FALSE) AS HRESULT
   DECLARE FUNCTION PutResource (BYVAL hinst AS HINSTANCE, BYVAL id AS UINT) AS HRESULT
   DECLARE FUNCTION PutRecord (BYVAL pIRecordInfo AS IRecordInfo PTR, BYVAL pRec AS ANY PTR) AS HRESULT
   DECLARE FUNCTION PutDateString (BYVAL pwszDate AS WSTRING PTR, BYVAL lcid AS LCID = 0, BYVAL dwFlags AS ULONG = 0) AS HRESULT
   DECLARE FUNCTION PutVbDate (BYVAL vbDate AS DATE_) AS HRESULT
   DECLARE FUNCTION PutSystemTime (BYVAL st AS SYSTEMTIME PTR) AS BOOLEAN
   DECLARE FUNCTION PutGuid (BYVAL guid AS IID PTR) AS HRESULT
   DECLARE FUNCTION PutFileTime (BYVAL pft AS FILETIME PTR) AS HRESULT
   DECLARE FUNCTION PutFileTimeArray (BYVAL prgft AS FILETIME PTR, BYVAL cElems AS ULONG) AS HRESULT
   DECLARE FUNCTION PutStrRet (BYVAL pstrret AS STRRET PTR, BYVAL pidl AS PCUITEMID_CHILD) AS HRESULT
   DECLARE FUNCTION PutDec (BYVAL dec AS DECIMAL) AS HRESULT
   DECLARE FUNCTION PutDecFromStr (BYVAL pwszIn AS WSTRING PTR, BYVAL lcid AS LCID = 0, BYVAL dwFlags AS ULONG = 0) AS HRESULT
   DECLARE FUNCTION PutDecFromDouble (BYVAL dbIn AS DOUBLE) AS HRESULT
   DECLARE FUNCTION PutDecFromCy (BYVAL cyIn AS CY) AS HRESULT
   DECLARE FUNCTION PutBooleanArray (BYVAL prgf AS WINBOOL PTR, BYVAL cElems AS ULONG) AS HRESULT
   DECLARE FUNCTION PutShortArray (BYVAL prgf AS SHORT PTR, BYVAL cElems AS ULONG) AS HRESULT
   DECLARE FUNCTION PutUShortArray (BYVAL prgf AS USHORT PTR, BYVAL cElems AS ULONG) AS HRESULT
   DECLARE FUNCTION PutLongArray (BYVAL prgn AS LONG PTR, BYVAL cElems AS ULONG) AS CVAR
   DECLARE FUNCTION PutULongArray (BYVAL prgn AS ULONG PTR, BYVAL cElems AS ULONG) AS HRESULT
   DECLARE FUNCTION PutLongIntArray (BYVAL prgn AS LONGINT PTR, BYVAL cElems AS ULONG) AS HRESULT
   DECLARE FUNCTION PutULongIntArray (BYVAL prgn AS ULONGINT PTR, BYVAL cElems AS ULONG) AS HRESULT
   DECLARE FUNCTION PutDoubleArray (BYVAL prgn AS DOUBLE PTR, BYVAL cElems AS ULONG) AS HRESULT
   DECLARE FUNCTION PutStringArray (BYVAL prgsz AS PCWSTR, BYVAL cElems AS ULONG) AS HRESULT
   DECLARE FUNCTION PutPropVariant (BYVAL pPropVar AS PROPVARIANT PTR) AS HRESULT
   DECLARE FUNCTION PutVariantArrayElem (BYVAL pvarIn AS VARIANT PTR, BYVAL iElem AS ULONG) AS CVAR
   ' // Assignments by reference
   DECLARE FUNCTION PutRef (BYVAL _value AS ANY PTR, BYVAL _vType AS WORD = VT_I8) AS HRESULT
   DECLARE FUNCTION PutRef (BYVAL _pvar AS ANY PTR, BYREF strType AS STRING) AS HRESULT
   ' // Safe arrays
   DECLARE FUNCTION GetDim () AS ULONG
   DECLARE FUNCTION GetLBound (BYVAL nDim AS UINT = 1) AS LONG
   DECLARE FUNCTION GetUBound (BYVAL nDim AS UINT = 1) AS LONG
   DECLARE FUNCTION GetVariantElem (BYVAL iElem AS ULONG) AS CVAR
   ' // Arrays
   DECLARE FUNCTION GetElementCount () AS ULONG
   DECLARE FUNCTION GetBooleanElem (BYVAL iElem AS ULONG) AS BOOLEAN
   DECLARE FUNCTION GetShortElem (BYVAL iElem AS ULONG) AS SHORT
   DECLARE FUNCTION GetUShortElem (BYVAL iElem AS ULONG) AS USHORT
   DECLARE FUNCTION GetLongElem (BYVAL iElem AS ULONG) AS LONG
   DECLARE FUNCTION GetULongElem (BYVAL iElem AS ULONG) AS ULONG
   DECLARE FUNCTION GetLongIntElem (BYVAL iElem AS ULONG) AS LONGINT
   DECLARE FUNCTION GetULongIntElem (BYVAL iElem AS ULONG) AS ULONGINT
   DECLARE FUNCTION GetDoubleElem (BYVAL iElem AS ULONG) AS DOUBLE
   DECLARE FUNCTION GetStringElem (BYVAL iElem AS ULONG) AS CWSTR
   ' // Other...
   DECLARE SUB Clear
   DECLARE FUNCTION Attach (BYVAL pvar AS VARIANT PTR) AS HRESULT
   DECLARE FUNCTION Attach (BYREF v AS VARIANT) AS HRESULT
   DECLARE FUNCTION Detach (BYVAL pvar AS VARIANT PTR) AS HRESULT
   DECLARE FUNCTION Detach (BYREF v AS VARIANT) AS HRESULT
   DECLARE FUNCTION vType () AS VARTYPE
   DECLARE FUNCTION ChangeType (BYVAL vtNew AS VARTYPE, BYVAL wFlags AS USHORT = 0) AS HRESULT
   DECLARE FUNCTION ChangeTypeEx (BYVAL vtNew AS VARTYPE, BYVAL lcid AS LCID = 0, BYVAL wFlags AS USHORT = 0) AS HRESULT
   DECLARE FUNCTION Round (BYREF cv AS CVAR, BYVAL cDecimals AS LONG) AS CVAR
   DECLARE FUNCTION Round (BYVAL cDecimals AS LONG) AS CVAR
   DECLARE FUNCTION FormatNumber (BYVAL iNumDig AS LONG = -1, BYVAL ilncLead AS LONG = -2, _
           BYVAL iUseParens AS LONG = -2, BYVAL iGroup AS LONG = -2, BYVAL dwFlags AS DWORD = 0) AS CWSTR
   ' // VAL wrappers
   DECLARE FUNCTION ValInt () AS LONG
   DECLARE FUNCTION ValUInt () AS ULONG
   DECLARE FUNCTION ValLong () AS LONG
   DECLARE FUNCTION ValULong () AS ULONG
   DECLARE FUNCTION ValLongInt () AS LONGINT
   DECLARE FUNCTION ValULongInt () AS ULONGINT
   DECLARE FUNCTION ValDouble () AS DOUBLE
   DECLARE FUNCTION Value () AS DOUBLE
   ' // Math operators
   DECLARE OPERATOR += (BYREF cv AS CVAR)
   DECLARE OPERATOR &= (BYREF cv AS CVAR)
   DECLARE OPERATOR -= (BYREF cv AS CVAR)
   DECLARE OPERATOR *= (BYREF cv AS CVAR)
   DECLARE OPERATOR /= (BYREF cv AS CVAR)
   DECLARE OPERATOR \= (BYREF cv AS CVAR)
   DECLARE OPERATOR ^= (BYREF cv AS CVAR)
   DECLARE OPERATOR And= (BYREF cv AS CVAR)
   DECLARE OPERATOR Or= (BYREF cv AS CVAR)
   DECLARE OPERATOR XOr= (BYREF cv AS CVAR)
   DECLARE OPERATOR Eqv= (BYREF cv AS CVAR)
   DECLARE OPERATOR Imp= (BYREF cv AS CVAR)
   DECLARE OPERATOR Mod= (BYREF cv AS CVAR)

END TYPE
' ########################################################################################

' ========================================================================================
' Note about constructors:
' We can use the constructors to pass values to parameters in procedures without assigning
' them first to a variable, e.g.:
' SUB Foo (BYREF cv AS CVAR)   ' -or- (BYVAL cv AS CVAR PTR)
'    PRINT AfxCVarToStr(cv)
' END SUB
' Foo CVAR("Test string")
' Foo CVAR(12345)
' Foo CVAR(12345, "LONG")
' ========================================================================================

' ========================================================================================
' Default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CVar
   CVAR_DP("CVAR CONSTRUCTOR")
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Destructor
' ========================================================================================
PRIVATE DESTRUCTOR CVar
   CVAR_DP("CVAR DESTRUCTOR")
   VariantClear(@vd)
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Initializes the CVAR from another CVAR.
' ========================================================================================
PRIVATE CONSTRUCTOR CVar (BYREF cv AS CVAR)
   CVAR_DP("CVAR CONSTRUCTOR - BYREF CVAR")
'   VariantCopy(@vd, @cv)   ' // removed @ operator
   VariantCopy(@vd, cv.sptr)
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Initializes the CVAR from a VARIANT.
' ========================================================================================
PRIVATE CONSTRUCTOR CVar (BYVAL v AS VARIANT)
   CVAR_DP("CVAR CONSTRUCTOR - BYVAL VARIANT")
   VariantCopy(@vd, @v)
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Initializes the CVAR from a string.
' ========================================================================================
PRIVATE CONSTRUCTOR CVar (BYREF wsz AS WSTRING)
   CVAR_DP("CVAR CONSTRUCTOR - WSTRING")
   vd.vt = VT_BSTR : vd.bstrVal = SysAllocString(@wsz)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CVar (BYREF cws AS CWSTR)
   CVAR_DP("CVAR CONSTRUCTOR - CWSTR")
   vd.vt = VT_BSTR : vd.bstrVal = SysAllocString(cws)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CVar (BYREF cbs AS CBSTR)
   CVAR_DP("CVAR CONSTRUCTOR - CBSTR")
   vd.vt = VT_BSTR : vd.bstrVal = SysAllocString(cbs)
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Initializes a CVAR from a pointer to a VARIANT.
' ========================================================================================
PRIVATE CONSTRUCTOR CVar (BYVAL pvar AS VARIANT PTR)
   CVAR_DP("CVAR CONSTRUCTOR - VARIANT PTR")
   VariantCopy(@vd, pvar)
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Initializes a CVAR from a CURRENCY structure.
' ========================================================================================
PRIVATE CONSTRUCTOR CVar (BYVAL cy AS CURRENCY)
   CVAR_DP("CVAR CONSTRUCTOR - CURRENCY")
   vd.vt = VT_CY
   vd.cyVal = cy
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Initializes a CVAR from a DECIMAL structure.
' ========================================================================================
PRIVATE CONSTRUCTOR CVar (BYVAL dec AS DECIMAL)
   CVAR_DP("CVAR CONSTRUCTOR - DECIMAL")
   vd.vt = VT_DECIMAL
   vd.decVal = dec
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Initializes a CVAR from an IDispatch pointer
' ========================================================================================
PRIVATE CONSTRUCTOR CVar (BYVAL pdisp AS IDispatch PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   CVAR_DP("CVAR CONSTRUCTOR - IDISPATCH PTR")
   IF fAddRef THEN IDispatch_AddRef(pdisp)
   V_VT(@vd) = VT_DISPATCH : V_DISPATCH(@vd) = pdisp
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Initializes a CVAR from an IUnknown pointer
' ========================================================================================
PRIVATE CONSTRUCTOR CVar (BYVAL punk AS IUnknown PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   CVAR_DP("CVAR CONSTRUCTOR - IUNKNOWN PTR")
   IF fAddRef THEN IUnknown_AddRef(punk)
   V_VT(@vd) = VT_UNKNOWN : V_UNKNOWN(@vd) = punk
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Initializes a CVAR from an integer value
' If _vType is wrong, default to VT_I8 (LongInt).
' ========================================================================================
PRIVATE CONSTRUCTOR CVar (BYVAL _value AS LONGINT, BYVAL _vType AS WORD = VT_I4)
   CVAR_DP("CVAR CONSTRUCTOR - value, type")
   IF (_vType AND VT_ARRAY) = VT_ARRAY THEN
      ' // NULL array : Suitable for parameters that, instead of being optional,
      ' // require a variant with a null safearray of a given type.
      vd.vt = _vType
   ELSE
      ' // If _value exceeds the limits of max long or max ulong, promote to ulongint
      IF SGN(_value) = -1 AND _value > LONG_MAX THEN
         _vType = VT_I8
      ELSEIF _value > ULONG_MAX THEN
         _vType = VT_I8
      END IF
      SELECT CASE _vType
         CASE VT_BOOL  : vd.vt = VT_BOOL : vd.boolVal = IIF(_value = 0, 0, -1)
         CASE VT_I1    : vd.vt = VT_I1   : vd.cVal = CBYTE(_value)
         CASE VT_UI1   : vd.vt = VT_UI1  : vd.bVal = CUBYTE(_value)
         CASE VT_I2    : vd.vt = VT_I2   : vd.iVal = CSHORT(_value)
         CASE VT_UI2   : vd.vt = VT_UI2  : vd.uiVal = CUSHORT(_value)
         CASE VT_INT   : vd.vt = VT_INT  : vd.intVal = CLNG(_value)
         CASE VT_UINT  : vd.vt = VT_UINT : vd.uintVal = CULNG(_value)
         CASE VT_I4    : vd.vt = VT_I4   : vd.lVal = CLNG(_value)
         CASE VT_UI4   : vd.vt = VT_UI4  : vd.ulVal = CULNG(_value)
         CASE VT_I8    : vd.vt = VT_I8   : vd.llVal = CLNGINT(_value)
         CASE VT_UI8   : vd.vt = VT_UI8  : vd.ullVal = CULNGINT(_value)
         ' // If you need to assign an ULongInt value greater than a LongInt, use PutUlongInt.
         CASE VT_NULL  : vd.vt = VT_NULL
         CASE ELSE     : vd.vt = VT_I4   : vd.llVal = CLNG(_value)
      END SELECT
   END IF
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CVar (BYVAL _value AS LONGINT, BYREF strType AS STRING)
   CVAR_DP("CVAR CONSTRUCTOR - strType")
   DIM vt AS WORD
   SELECT CASE UCASE(strType)
      CASE "BOOL"     : vd.vt = VT_BOOL : vd.boolVal = IIF(_value = 0, 0, -1)
      CASE "BOOLEAN"  : vd.vt = VT_BOOL : vd.boolVal = IIF(_value = 0, 0, -1)
      CASE "BYTE"     : vd.vt = VT_I1   : vd.cVal = CBYTE(_value)
      CASE "UBYTE"    : vd.vt = VT_UI1  : vd.bVal = CUBYTE(_value)
      CASE "SHORT"    : vd.vt = VT_I2   : vd.iVal = CSHORT(_value)
      CASE "USHORT"   : vd.vt = VT_UI2  : vd.uiVal = CUSHORT(_value)
      CASE "INT"      : vd.vt = VT_INT  : vd.intVal = CLNG(_value)
      CASE "UINT"     : vd.vt = VT_UINT : vd.uintVal = CULNG(_value)
      CASE "LONG"     : vd.vt = VT_I4   : vd.lVal = CLNG(_value)
      CASE "ULONG"    : vd.vt = VT_UI4  : vd.ulVal = CULNG(_value)
      CASE "LONGINT"  : vd.vt = VT_I8   : vd.llVal = CLNGINT(_value)
      CASE "ULONGINT" : vd.vt = VT_UI8  : vd.ullVal = CULNGINT(_value)
      ' // If you need to assign an ULongInt value greater than a LongInt, use PutUlongInt.
      CASE "NULL"     : vd.vt = VT_NULL
      CASE ELSE       : vd.vt = VT_I4   : vd.llVal = CLNG(_value)
   END SELECT
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Initializes a CVAR from a float value
' If _vType is wrong, default to VT_R8 (double).
' ========================================================================================
PRIVATE CONSTRUCTOR CVar (BYVAL _value AS DOUBLE, BYVAL _vType AS WORD = VT_R8)
   CVAR_DP("CVAR CONSTRUCTOR - FLOAT")
   IF _value > FLT_MAX OR _value < FLT_MIN THEN _vType = VT_R8
   SELECT CASE _vType
      CASE VT_R4    : vd.vt = VT_R4   : vd.fltVal = CSNG(_value)
      CASE VT_R8    : vd.vt = VT_R8   : vd.dblVal = CDBL(_value)
      CASE ELSE     : vd.vt = VT_R8   : vd.dblVal = CDBL(_value)
   END SELECT
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CVar (BYVAL _value AS DOUBLE, BYREF strType AS STRING)
   CVAR_DP("CVAR CONSTRUCTOR - strType")
   SELECT CASE UCASE(strType)
      CASE "SINGLE" : vd.vt = VT_R4   : vd.fltVal = CSNG(_value)
      CASE "FLOAT"  : vd.vt = VT_R4   : vd.fltVal = CSNG(_value)
      CASE "DOUBLE" : vd.vt = VT_R8   : vd.dblVal = CDBL(_value)
      CASE ELSE     : vd.vt = VT_R8   : vd.dblVal = CDBL(_value)
   END SELECT
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Initializes a CVAR from a value by reference (a pointer to a variable), e.g.:
' DIM dblVal AS DOUBLE = 123456.12
' DIM cv AS CVAR = CVAR(@dblVal, "DOUBLE")
' print cv   ' // prints 123456.12
' dblVal = 345.12
' print cv   ' // prints 345.12
' ========================================================================================
PRIVATE CONSTRUCTOR CVar (BYVAL _pvar AS ANY PTR, BYVAL _vType AS WORD)
   CVAR_DP("CVAR CONSTRUCTOR - ANY PTR - vType")
   SELECT CASE _vType
      CASE VT_BOOL      : vd.vt = VT_BOOL OR VT_BYREF      : vd.pboolVal = _pvar
      CASE VT_I1        : vd.vt = VT_I1 OR VT_BYREF        : vd.pcVal = _pvar
      CASE VT_UI1       : vd.vt = VT_UI1 OR VT_BYREF       : vd.pbVal = _pvar
      CASE VT_I2        : vd.vt = VT_I2 OR VT_BYREF        : vd.piVal = _pvar
      CASE VT_UI2       : vd.vt = VT_UI2 OR VT_BYREF       : vd.puiVal = _pvar
      CASE VT_INT       : vd.vt = VT_INT OR VT_BYREF       : vd.pintVal = _pvar
      CASE VT_UINT      : vd.vt = VT_UINT OR VT_BYREF      : vd.puintVal = _pvar
      CASE VT_I4        : vd.vt = VT_I4 OR VT_BYREF        : vd.plVal = _pvar
      CASE VT_UI4       : vd.vt = VT_UI4 OR VT_BYREF       : vd.pulVal = _pvar
      CASE VT_I8        : vd.vt = VT_I8 OR VT_BYREF        : vd.pllVal = _pvar
      CASE VT_UI8       : vd.vt = VT_UI8 OR VT_BYREF       : vd.pullVal = _pvar
      CASE VT_R4        : vd.vt = VT_R4 OR VT_BYREF        : vd.pfltVal = _pvar
      CASE VT_R8        : vd.vt = VT_R8 OR VT_BYREF        : vd.pdblVal = _pvar
      CASE VT_BSTR      : vd.vt = VT_BSTR OR VT_BYREF      : vd.pbstrVal = _pvar
      CASE VT_UNKNOWN   : vd.vt = VT_UNKNOWN OR VT_BYREF   : vd.ppunkVal = _pvar
      CASE VT_DISPATCH  : vd.vt = VT_DISPATCH OR VT_BYREF  : vd.ppdispVal = _pvar
      CASE VT_DECIMAL   : vd.vt = VT_DECIMAL OR VT_BYREF   : vd.pdecVal = _pvar
      CASE VT_CY        : vd.vt = VT_CY OR VT_BYREF        : vd.pcyVal = _pvar
      CASE VT_DATE      : vd.vt = VT_DATE OR VT_BYREF      : vd.pdate = _pvar
      CASE VT_VARIANT   : vd.vt = VT_VARIANT OR VT_BYREF   : vd.pvarVal = _pvar
      CASE VT_SAFEARRAY : vd.vt = VT_SAFEARRAY OR VT_BYREF : vd.pvarVal = _pvar
      CASE VT_ERROR     : vd.vt = VT_ERROR OR VT_BYREF     : vd.pparray = _pvar
   END SELECT
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CVar (BYVAL _pvar AS ANY PTR, BYREF strType AS STRING)
   CVAR_DP("CVAR CONSTRUCTOR - ANY PTR - strType")
   IF _pvar <> NULL THEN
      SELECT CASE UCASE(strType)
         CASE "BOOL"      : vd.vt = VT_BOOL OR VT_BYREF      : vd.pboolVal = _pvar
         CASE "BOOLEAN"   : vd.vt = VT_BOOL OR VT_BYREF      : vd.pboolVal = _pvar
         CASE "BYTE"      : vd.vt = VT_I1 OR VT_BYREF        : vd.pcVal = _pvar
         CASE "UBYTE"     : vd.vt = VT_UI1 OR VT_BYREF       : vd.pbVal = _pvar
         CASE "SHORT"     : vd.vt = VT_I2 OR VT_BYREF        : vd.piVal = _pvar
         CASE "USHORT"    : vd.vt = VT_UI2 OR VT_BYREF       : vd.puiVal = _pvar
         CASE "INT"       : vd.vt = VT_INT OR VT_BYREF       : vd.pintVal = _pvar
         CASE "UINT"      : vd.vt = VT_UINT OR VT_BYREF      : vd.puintVal = _pvar
         CASE "LONG"      : vd.vt = VT_I4 OR VT_BYREF        : vd.plVal = _pvar
         CASE "ULONG"     : vd.vt = VT_UI4 OR VT_BYREF       : vd.pulVal = _pvar
         CASE "LONGINT"   : vd.vt = VT_I8 OR VT_BYREF        : vd.pllVal = _pvar
         CASE "ULONGINT"  : vd.vt = VT_UI8 OR VT_BYREF       : vd.pullVal = _pvar
         CASE "SINGLE"    : vd.vt = VT_R4 OR VT_BYREF        : vd.pfltVal = _pvar
         CASE "FLOAT"     : vd.vt = VT_R4 OR VT_BYREF        : vd.pfltVal = _pvar
         CASE "DOUBLE"    : vd.vt = VT_R8 OR VT_BYREF        : vd.pdblVal = _pvar
         CASE "BSTR"      : vd.vt = VT_BSTR OR VT_BYREF      : vd.pbstrVal = _pvar
         CASE "UNKNOWN"   : vd.vt = VT_UNKNOWN OR VT_BYREF   : vd.ppunkVal = _pvar
         CASE "DISPATCH"  : vd.vt = VT_DISPATCH OR VT_BYREF  : vd.ppdispVal = _pvar
         CASE "DECIMAL"   : vd.vt = VT_DECIMAL OR VT_BYREF   : vd.pdecVal = _pvar
         CASE "CY"        : vd.vt = VT_CY OR VT_BYREF        : vd.pcyVal = _pvar
         CASE "CURRENCY"  : vd.vt = VT_CY OR VT_BYREF        : vd.pcyVal = _pvar
         CASE "DATE"      : vd.vt = VT_DATE OR VT_BYREF      : vd.pdate = _pvar
         CASE "VARIANT"   : vd.vt = VT_VARIANT OR VT_BYREF   : vd.pvarVal = _pvar
         CASE "SAFEARRAY" : vd.vt = VT_SAFEARRAY OR VT_BYREF : vd.pvarVal = _pvar
         CASE "ERROR"     : vd.vt = VT_ERROR OR VT_BYREF     : vd.pparray = _pvar
      END SELECT
   END IF
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the address of the underlying variant.
' Removed to allow to use @ to get the address of the class.
' After removing it, we can now do:
' SUB Foo (BYVAL cv AS CVAR PTR)
'   AfxMsg AfxCVarToStr(cv)
' END SUB
' DIM cv AS CVAR = "Test string"
' Foo @cv
' --Or--:
' SUB Foo (BYREF v AS VARIANT)
'    AfxMsg AfxVarToStr(@v)
' END SUB
' DIM cv AS CVAR = "Test string"
' Foo cv
' ========================================================================================
'PRIVATE OPERATOR CVar.@ () AS VARIANT PTR
'   CVAR_DP("CVAR OPERATOR @ - " & WSTR(@vd))
'   OPERATOR = @vd
'END OPERATOR
' ========================================================================================

' ========================================================================================
' Clears the CVAR and returns the address of the underlying variant.
' To pass the variant to an OUT BYVAL VARIANT PTR parameter.
' If we pass a CVAR to a function with an OUT variant parameter without first clearing the
' contents of the CVAR, we may have a memory leak.
' Example:
' SUB Foo (BYVAL v AS VARIANT PTR)
'    v->vt = VT_I4
'    v->lVal = 12345
' END SUB
' DIM cv AS CVAR = "Test string"
' Foo cv.vptr
' PRINT cv
' Otherwise, you need to clear the underlying variant before passing the CVAR.
' DIM cv AS CVAR = "Test string"
' cv.Clear
' Foo *cv
' PRINT cv
' ========================================================================================
PRIVATE FUNCTION CVar.vptr () AS VARIANT PTR
   CVAR_DP("CVAR FUNCTION vptr")
   VariantClear @vd
   RETURN @vd
END FUNCTION
' ========================================================================================

' ========================================================================================
' To pass the variant to a IN BYVAL VARIANT PTR parameter.
' Usage example:
' SUB Foo (BYVAL v AS VARIANT PTR)
'    PRINT AfxVarToStr(v)
' END SUB
' Using the pointer syntax:
' DIM pcv AS CVAR PTR = NEW CVAR("Test string")
' Foo pcv->sptr
' Delete pcv
' Using the normal syntax:
' DIM cv AS CVAR = "Test string"
' Foo cv.sptr
' But with the normal syntax you can use the * operator instead:
' DIM cv AS CVAR = "Test string"
' Foo *cv
' ========================================================================================
PRIVATE FUNCTION CVar.sptr () AS VARIANT PTR
   CVAR_DP("CVAR FUNCTION sptr")
   RETURN @vd
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the address of the underlying variant.
' One * returns the address of the underlying variant.
' Two ** deferences the variant data.
' Can be used to pass the variant to a BYVAL VARIANT parameter, e.g.
' SUB Foo (BYVAL v AS VARIANT)
'    PRINT AfxVarToStr(@v)
' END SUB
' DIM cv AS CVAR = "Test string"
' Foo **cv
' -or-
' Foo *cv.sptr
' Using the pointer syntax:
' DIM pcv AS CVAR PTR  = NEW CVAR("Test string")
' Foo *pcv
' Delete pcv
' Using the constructors_
' Foo **CVAR(12345, "LONG")
' ========================================================================================
PRIVATE OPERATOR * (BYREF cv AS CVAR) AS VARIANT PTR
   CVAR_DP("CVAR OPERATOR *")
   OPERATOR = @cv.vd
END OPERATOR
' ========================================================================================

' ========================================================================================
' The CAST operators allow to transparently pass the underlying VARIANT to a procedure.
' They aren't called directly.
'
' SUB Foo (BYREF v AS VARIANT)
'    PRINT AfxVarToStr(@v)
' END SUB
' Foo CVAR(12345, "LONG")
'
' SUB Foo2 (BYVAL v AS VARIANT)
'    PRINT AfxVarToStr(@v)
' END SUB
' Foo2 CVAR(12345, "LONG")
'
' SUB Foo3 (BYREF cv AS CVAR)
'    PRINT cv
' END SUB
' Foo3 CVAR(12345, "LONG")
'
' SUB Foo4 (BYVAL cv AS CVAR PTR)
'    PRINT *cv
' END SUB
' Foo4 CVAR(12345, "LONG")
' ========================================================================================

' ========================================================================================
' Returns the variant data.
' ========================================================================================
PRIVATE OPERATOR CVar.CAST () AS VARIANT
   CVAR_DP("CVAR CAST VARIANT")
   OPERATOR = vd
END OPERATOR
' ========================================================================================
' ========================================================================================
' Returns a pointer to the variant data.
' ========================================================================================
PRIVATE OPERATOR CVar.CAST () AS ANY PTR
   CVAR_DP("CVAR CAST ANY PTR")
   OPERATOR = cast(ANY PTR, @vd)
END OPERATOR
' ========================================================================================

' =====================================================================================
' Extracts the contents of a VARIANT and returns them as a CWSTR.
' =====================================================================================
PRIVATE OPERATOR CVar.CAST () BYREF AS WSTRING
   CVAR_DP("CVAR CAST BYREF AS WSTRING")
   cws = AfxVarToStr(@vd)
   OPERATOR = *cast(WSTRING PTR, cws.m_pBuffer)
END OPERATOR
' ========================================================================================
' =====================================================================================
PRIVATE FUNCTION CVar.ToStr () AS CWSTR
   CVAR_DP("CVAR ToStr")
   RETURN AfxVarToStr(@vd)
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CVar.ToWStr () AS CWSTR
   CVAR_DP("CVAR ToWStr")
   RETURN AfxVarToStr(@vd)
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CVar.wstr () AS CWSTR
   CVAR_DP("CVAR wstr")
   RETURN AfxVarToStr(@vd)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Extracts the contents of a VARIANT and returns them as a CBSTR.
' =====================================================================================
PRIVATE FUNCTION CVar.ToBStr () AS CBSTR
   CVAR_DP("CVAR ToBStr")
   RETURN AfxVarToStr(@vd)
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CVar.bstr () AS CBSTR
   CVAR_DP("CVAR bstr")
   RETURN AfxVarToStr(@vd)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Extracts an unicode string contained in a variant and maps it to an encoded UTF8 code page.
' =====================================================================================
PRIVATE FUNCTION CVar.ToUtf8 () AS STRING
   CVAR_DP("CVAR ToUtf8")
   IF vd.vt = VT_BSTR THEN RETURN AfxAcode(vd.bstrVal, CP_UTF8)
   RETURN ""
END FUNCTION
' =====================================================================================

' =====================================================================================
' Extracts the contents of a buffer stored in a CVAR of type VT_ARRRAY OR VT_UI1.
' Note: To retrieve the size of the array call GetElementCount.
' =====================================================================================
PRIVATE FUNCTION CVar.ToBuffer (BYVAL pv AS ANY PTR, BYVAL cb AS UINT) AS HRESULT
   CVAR_DP("CVAR ToBuffer")
   RETURN AfxVariantToBuffer(@vd, pv, cb)
END FUNCTION
' =====================================================================================
' =====================================================================================
' Extracts the contents of a buffer stored in a CVAR of type VT_ARRRAY OR VT_UI1 and
' returns it as string used as a buffer.
' =====================================================================================
PRIVATE FUNCTION CVar.ToBuffer () AS STRING
   CVAR_DP("CVAR ToBuffer - STRING")
   IF (vd.vt = VT_ARRAY OR VT_UI1) THEN
      DIM cb AS LONG = this.GetElementCount
      IF cb THEN
         DIM s AS STRING = SPACE(cb)
         this.ToBuffer(STRPTR(s), cb)
         RETURN s
      END IF
   END IF
   RETURN ""
END FUNCTION
' =====================================================================================

' =====================================================================================
' Checks if the variant if of the type VT_UNKNOWN and returns the unknown pointer.
' It is the responsability of the called of release the returned pointer.
' =====================================================================================
PRIVATE FUNCTION CVar.ToUnknown () AS ANY PTR
   IF vd.vt <> VT_UNKNOWN AND vd.vt <> VT_DISPATCH THEN RETURN NULL
   IF vd.vt = VT_UNKNOWN THEN
      IF vd.punkVal = NULL THEN RETURN NULL
      IUnknown_AddRef(vd.punkVal)
      RETURN vd.punkVal
   END IF
   IF vd.vt = VT_DISPATCH THEN
      IF vd.pdispVal = NULL THEN RETURN NULL
      DIM punk AS IUnknown PTR
      DIM hr AS HRESULT = vd.pdispVal->lpvtbl->QueryInterface(vd.pdispVal, @IID_IUnknown, @punk)
      IF hr = S_OK THEN
         ' // QueryInterface calls IUnknown_AddRef in the pointer it returns
         RETURN punk
      ELSE
         IDispatch_AddRef(vd.pdispVal)
         RETURN vd.pdispVal
      END IF
   END IF
END FUNCTION
' =====================================================================================

' =====================================================================================
' If the variant is of the type VT_UNKNOWN or VT_DISPATCH it returns the dispatch pointer.
' It is the responsability of the called of release the returned pointer.
' =====================================================================================
PRIVATE FUNCTION CVar.ToDispatch () AS ANY PTR
   IF vd.vt <> VT_UNKNOWN AND vd.vt <> VT_DISPATCH THEN RETURN NULL
   IF vd.vt = VT_DISPATCH THEN
      IF vd.pdispVal = NULL THEN RETURN NULL
      IDispatch_AddRef(vd.pdispVal)
      RETURN vd.pdispVal
   END IF
   IF vd.vt = VT_UNKNOWN THEN
      IF vd.punkVal = NULL THEN RETURN NULL
      DIM pdisp AS IDispatch PTR
      DIM hr AS HRESULT = vd.punkVal->lpvtbl->QueryInterface(vd.punkVal, @IID_IDispatch, @pdisp)
      IF hr = S_OK THEN
         ' // QueryInterface calls IUnknown_AddRef in the pointer it returns
         RETURN pdisp
      ELSE
         IUnknown_AddRef(vd.punkVal)
         RETURN vd.punkVal
      END IF
   END IF
END FUNCTION
' =====================================================================================

' =====================================================================================
' Converts a CVAR of type decimal to a double.
' =====================================================================================
PRIVATE FUNCTION CVar.DecToDouble () AS DOUBLE
   CVAR_DP("CVAR DecToDouble")
   DIM dblOut AS DOUBLE
   IF vd.vt = VT_DECIMAL THEN VarR8FromDec(@vd.decVal, @dblOut)
   RETURN dblOut
END FUNCTION
' =====================================================================================

' =====================================================================================
' Converts a variant of type decimal to currency.
' =====================================================================================
PRIVATE FUNCTION CVar.DecToCy () AS CY
   CVAR_DP("CVAR DecToCy")
   DIM cyOut AS CY
   IF vd.vt = VT_DECIMAL THEN VarCyFromDec(@vd.decVal, @cyOut)
   RETURN cyOut
END FUNCTION
' =====================================================================================

' =====================================================================================
' Converts a VT_DATE variant to a variant representation (double)
' =====================================================================================
PRIVATE FUNCTION CVar.ToVbDate () AS DATE_
   CVAR_DP("CVAR ToVbDate")
   IF vd.vt = VT_DATE THEN RETURN vd.date
END FUNCTION
' =====================================================================================

' =====================================================================================
' Extracts a FILETIME structure from a CVAR as a SYSTEMTIME structure.
' =====================================================================================
PRIVATE FUNCTION CVar.ToSystemTime () AS SYSTEMTIME
   CVAR_DP("CVAR ToSystemTime")
   DIM st AS SYSTEMTIME
   IF vd.vt = VT_DATE THEN VariantTimeToSystemTime(vd.date, @st)
   RETURN st
END FUNCTION
' =====================================================================================

' =====================================================================================
' Extracts a GUID property value of a VARIANT.
' =====================================================================================
PRIVATE FUNCTION CVar.ToGuid () AS GUID
   CVAR_DP("CVAR ToGuid")
   DIM guid_ AS GUID : AfxVariantToGUID(@vd, @guid_) : RETURN guid_
END FUNCTION
' =====================================================================================

' =====================================================================================
' Extracts a GUID property value as a CWSTR.
' =====================================================================================
PRIVATE FUNCTION CVar.ToGuidStr () AS CWSTR
   CVAR_DP("CVAR ToGuidStr")
   DIM guid_ AS GUID, cwsGuid AS CWSTR, pwsz AS WSTRING PTR
   DIM hr AS HRESULT = AfxVariantToGUID(@vd, @guid_)
   IF hr = S_OK THEN
      StringFromCLSID(@guid_, CAST(LPOLESTR PTR, @pwsz))
      IF pwsz THEN cwsGuid = *pwsz
      CoTaskMemFree(pwsz)
   END IF
   RETURN cwsGuid
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CVar.ToGuidWStr () AS CWSTR
   CVAR_DP("CVAR ToGuidStr")
   DIM guid_ AS GUID, cwsGuid AS CWSTR, pwsz AS WSTRING PTR
   DIM hr AS HRESULT = AfxVariantToGUID(@vd, @guid_)
   IF hr = S_OK THEN
      StringFromCLSID(@guid_, CAST(LPOLESTR PTR, @pwsz))
      IF pwsz THEN cwsGuid = *pwsz
      CoTaskMemFree(pwsz)
   END IF
   RETURN cwsGuid
END FUNCTION
' =====================================================================================

' =====================================================================================
' Extracts a GUID property value of a CBSTR.
' =====================================================================================
PRIVATE FUNCTION CVar.ToGuidBStr () AS CBSTR
   CVAR_DP("CVAR ToGuidBStr")
   DIM guid_ AS GUID, cbsGuid AS CBSTR, pwsz AS WSTRING PTR
   DIM hr AS HRESULT = AfxVariantToGUID(@vd, @guid_)
   IF hr = S_OK THEN
      StringFromCLSID(@guid_, CAST(LPOLESTR PTR, @pwsz))
      IF pwsz THEN cbsGuid = *pwsz
      CoTaskMemFree(pwsz)
   END IF
   RETURN cbsGuid
END FUNCTION
' =====================================================================================

' =====================================================================================
' Extracts a date and time value in Microsoft MS-DOS format from a VARIANT.
' =====================================================================================
PRIVATE FUNCTION CVar.ToDosDateTime (BYVAL pwDate AS USHORT PTR, BYVAL pwTime AS USHORT PTR) AS HRESULT
   CVAR_DP("CVAR ToDosDateTime")
   RETURN AfxVariantToDosDateTime(@vd, pwDate, pwTime)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Extracts a FILETIME structure from a VARIANT.
' =====================================================================================
PRIVATE FUNCTION CVar.ToFileTime (BYVAL stfOut AS AFX_PSTIME_FLAGS) AS FILETIME
   CVAR_DP("CVAR ToFileTime")
   DIM ft AS FILETIME
   AfxVariantToFileTime(@vd, stfOut, @ft)
   RETURN ft
END FUNCTION
' =====================================================================================

' =====================================================================================
' If the source variant is a VT_BSTR, extracts string and places it into a STRRET structure.
' =====================================================================================
PRIVATE FUNCTION CVar.ToStrRet () AS STRRET
   CVAR_DP("CVAR ToStrRet")
   DIM strret_ AS STRRET
   AfxVariantToStrRet(@vd, @strret_)
   RETURN strret_
END FUNCTION
' =====================================================================================

' =====================================================================================
' Extracts an array of Boolean values from a VARIANT structure.
' Returns the number of elements extracted from the source Variant structure.
' GetLastError: Returns S_OK if successful, or an error value otherwise, including the following:
' TYPE_E_BUFFERTOOSMALL: The source VARIANT contained more than crgn values.
' E_INVALIDARG: The VARIANT was not of the appropriate type.
' To retrieve the number of elements in the array call the GetElementCount method.
' =====================================================================================
PRIVATE FUNCTION CVar.ToBooleanArray (BYVAL prgf AS WINBOOL PTR, BYVAL crgn AS ULONG) AS ULONG
   CVAR_DP("CVAR ToBooleanArray")
   DIM pcElem AS ULONG
   SetLastError AfxVariantToBooleanArray(@vd, prgf, crgn, @pcElem)
   RETURN pcElem
END FUNCTION
' =====================================================================================
' =====================================================================================
' Allocates an array of BOOL values then extracts data from a VARIANT into that array.
' GetLastError: Returns S_OK if successful, or an error value otherwise.
' Note: Use CoTaskMemFree to free the memory allocated for the array.
' =====================================================================================
PRIVATE FUNCTION CVar.ToBooleanArrayAlloc (BYVAL pprgf AS WINBOOL PTR PTR) AS ULONG
   CVAR_DP("CVAR ToBooleanArrayAlloc")
   DIM pcElem AS ULONG
   SetLastError AfxVariantToBooleanArrayAlloc(@vd, pprgf, @pcElem)
   RETURN pcElem
END FUNCTION
' =====================================================================================

' =====================================================================================
' Extracts an array of Int16 values from a VARIANT structure.
' Returns the number of elements extracted from the source Variant structure.
' GetLastError: Returns S_OK if successful, or an error value otherwise, including the following:
' TYPE_E_BUFFERTOOSMALL: The source VARIANT contained more than crgn values.
' E_INVALIDARG: The VARIANT was not of the appropriate type.
' To retrieve the number of elements in the array call the GetElementCount method.
' =====================================================================================
PRIVATE FUNCTION CVar.ToShortArray (BYVAL prgn AS SHORT PTR, BYVAL crgn AS ULONG) AS ULONG
   CVAR_DP("CVAR ToShortArray")
   DIM pcElem AS ULONG
   SetLastError AfxVariantToInt16Array(@vd, prgn, crgn, @pcElem)
   RETURN pcElem
END FUNCTION
' =====================================================================================
' =====================================================================================
' Extracts data from a vector structure into a newly-allocated Int16 array.
' GetLastError: Returns S_OK if successful, or an error value otherwise.
' Note Use CoTaskMemFree to release the memory allocated for the array.
' =====================================================================================
PRIVATE FUNCTION CVar.ToShortArrayAlloc (BYVAL pprgn AS SHORT PTR PTR) AS ULONG
   CVAR_DP("CVAR ToShortArrayAlloc")
   DIM pcElem AS ULONG
   SetLastError AfxVariantToInt16ArrayAlloc(@vd, pprgn, @pcElem)
   RETURN pcElem
END FUNCTION
' =====================================================================================

' =====================================================================================
' Extracts an array of UInt16 values from a VARIANT structure.
' Returns the number of elements extracted from the source Variant structure.
' GetLastError: Returns S_OK if successful, or an error value otherwise, including the following:
' TYPE_E_BUFFERTOOSMALL: The source VARIANT contained more than crgn values.
' E_INVALIDARG: The VARIANT was not of the appropriate type.
' To retrieve the number of elements in the array call the GetElementCount method.
' =====================================================================================
PRIVATE FUNCTION CVar.ToUShortArray (BYVAL prgn AS USHORT PTR, BYVAL crgn AS ULONG) AS ULONG
   CVAR_DP("CVAR ToShortArray")
   DIM pcElem AS ULONG
   SetLastError AfxVariantToUInt16Array(@vd, prgn, crgn, @pcElem)
   RETURN pcElem
END FUNCTION
' =====================================================================================
' =====================================================================================
' Extracts data from a vector structure into a newly-allocated Int16 array.
' GetLastError: Returns S_OK if successful, or an error value otherwise.
' Note Use CoTaskMemFree to release the memory allocated for the array.
' =====================================================================================
PRIVATE FUNCTION CVar.ToUShortArrayAlloc (BYVAL pprgn AS USHORT PTR PTR) AS ULONG
   CVAR_DP("CVAR ToUShortArrayAlloc")
   DIM pcElem AS ULONG
   SetLastError AfxVariantToUInt16ArrayAlloc(@vd, pprgn, @pcElem)
   RETURN pcElem
END FUNCTION
' =====================================================================================

' =====================================================================================
' Extracts an array of Int32 values from a VARIANT structure.
' Returns the number of elements extracted from the source Variant structure.
' GetLastError: Returns S_OK if successful, or an error value otherwise, including the following:
' TYPE_E_BUFFERTOOSMALL: The source VARIANT contained more than crgn values.
' E_INVALIDARG: The VARIANT was not of the appropriate type.
' To retrieve the number of elements in the array call the GetElementCount method.
' =====================================================================================
PRIVATE FUNCTION CVar.ToLongArray (BYVAL prgn AS LONG PTR, BYVAL crgn AS ULONG) AS ULONG
   CVAR_DP("CVAR ToLongArray")
   DIM pcElem AS ULONG
   SetLastError AfxVariantToInt32Array(@vd, prgn, crgn, @pcElem)
   RETURN pcElem
END FUNCTION
' =====================================================================================
' =====================================================================================
' Extracts data from a vector structure into a newly-allocated Int32 array.
' GetLastError: Returns S_OK if successful, or an error value otherwise.
' Note Use CoTaskMemFree to release the memory allocated for the array.
' =====================================================================================
PRIVATE FUNCTION CVar.ToLongArrayAlloc (BYVAL pprgn AS LONG PTR PTR) AS ULONG
   CVAR_DP("CVAR ToLongArrayAlloc")
   DIM pcElem AS ULONG
   SetLastError AfxVariantToInt32ArrayAlloc(@vd, pprgn, @pcElem)
   RETURN pcElem
END FUNCTION
' =====================================================================================

' =====================================================================================
' Extracts an array of UInt32 values from a VARIANT structure.
' Returns the number of elements extracted from the source Variant structure.
' GetLastError: Returns S_OK if successful, or an error value otherwise, including the following:
' TYPE_E_BUFFERTOOSMALL: The source VARIANT contained more than crgn values.
' E_INVALIDARG: The VARIANT was not of the appropriate type.
' To retrieve the number of elements in the array call the GetElementCount method.
' =====================================================================================
PRIVATE FUNCTION CVar.ToULongArray (BYVAL prgn AS ULONG PTR, BYVAL crgn AS ULONG) AS ULONG
   CVAR_DP("CVAR ToULongArray")
   DIM pcElem AS ULONG
   SetLastError AfxVariantToUInt32Array(@vd, prgn, crgn, @pcElem)
   RETURN pcElem
END FUNCTION
' =====================================================================================
' =====================================================================================
' Extracts data from a vector structure into a newly-allocated unsigned Int32 array.
' GetLastError: Returns S_OK if successful, or an error value otherwise.
' Note Use CoTaskMemFree to release the memory allocated for the array.
' =====================================================================================
PRIVATE FUNCTION CVar.ToULongArrayAlloc (BYVAL pprgn AS ULONG PTR PTR) AS ULONG
   CVAR_DP("CVAR ToULongArrayAlloc")
   DIM pcElem AS ULONG
   SetLastError AfxVariantToUInt32ArrayAlloc(@vd, pprgn, @pcElem)
   RETURN pcElem
END FUNCTION
' =====================================================================================

' =====================================================================================
' Extracts an array of Int64 values from a VARIANT structure.
' Returns the number of elements extracted from the source Variant structure.
' GetLastError: Returns S_OK if successful, or an error value otherwise, including the following:
' TYPE_E_BUFFERTOOSMALL: The source VARIANT contained more than crgn values.
' E_INVALIDARG: The VARIANT was not of the appropriate type.
' To retrieve the number of elements in the array call the GetElementCount method.
' =====================================================================================
PRIVATE FUNCTION CVar.ToLongIntArray (BYVAL prgn AS LONGINT PTR, BYVAL crgn AS ULONG) AS ULONG
   CVAR_DP("CVAR ToULongArray")
   DIM pcElem AS ULONG
   SetLastError AfxVariantToInt64Array(@vd, prgn, crgn, @pcElem)
   RETURN pcElem
END FUNCTION
' =====================================================================================
' =====================================================================================
' Extracts data from a vector structure into a newly-allocated Int64 array.
' GetLastError: Returns S_OK if successful, or an error value otherwise.
' Note Use CoTaskMemFree to release the memory allocated for the array.
' =====================================================================================
PRIVATE FUNCTION CVar.ToLongIntArrayAlloc (BYVAL pprgn AS LONGINT PTR PTR) AS ULONG
   CVAR_DP("CVAR ToLongIntArrayAlloc")
   DIM pcElem AS ULONG
   SetLastError AfxVariantToInt64ArrayAlloc(@vd, pprgn, @pcElem)
   RETURN pcElem
END FUNCTION
' =====================================================================================

' =====================================================================================
' Extracts an array of Int64 values from a VARIANT structure.
' Returns the number of elements extracted from the source Variant structure.
' GetLastError: Returns S_OK if successful, or an error value otherwise, including the following:
' TYPE_E_BUFFERTOOSMALL: The source VARIANT contained more than crgn values.
' E_INVALIDARG: The VARIANT was not of the appropriate type.
' To retrieve the number of elements in the array call the GetElementCount method.
' To retrieve the number of elements in the array call the GetElementCount method.
' =====================================================================================
PRIVATE FUNCTION CVar.ToULongIntArray (BYVAL prgn AS ULONGINT PTR, BYVAL crgn AS ULONG) AS ULONG
   CVAR_DP("CVAR ToULongIntArray")
   DIM pcElem AS ULONG
   SetLastError AfxVariantToUInt64Array(@vd, prgn, crgn, @pcElem)
   RETURN pcElem
END FUNCTION
' =====================================================================================
' =====================================================================================
' Extracts data from a vector structure into a newly-allocated unsigned Int64 array.
' GetLastError: Returns S_OK if successful, or an error value otherwise.
' Note Use CoTaskMemFree to release the memory allocated for the array.
' =====================================================================================
PRIVATE FUNCTION CVar.ToULongIntArrayAlloc (BYVAL pprgn AS ULONGINT PTR PTR) AS ULONG
   CVAR_DP("CVAR ToULongIntArrayAlloc")
   DIM pcElem AS ULONG
   SetLastError AfxVariantToUInt64ArrayAlloc(@vd, pprgn, @pcElem)
   RETURN pcElem
END FUNCTION
' =====================================================================================

' =====================================================================================
' Extracts an array of Double values from a VARIANT structure.
' Returns the number of elements extracted from the source Variant structure.
' GetLastError: Returns S_OK if successful, or an error value otherwise, including the following:
' TYPE_E_BUFFERTOOSMALL: The source VARIANT contained more than crgn values.
' E_INVALIDARG: The VARIANT was not of the appropriate type.
' To retrieve the number of elements in the array call the GetElementCount method.
' =====================================================================================
PRIVATE FUNCTION CVar.ToDoubleArray (BYVAL prgn AS DOUBLE PTR, BYVAL crgn AS ULONG) AS ULONG
   CVAR_DP("CVAR ToDoubleArray")
   DIM pcElem AS ULONG
   SetLastError AfxVariantToDoubleArray(@vd, prgn, crgn, @pcElem)
   RETURN pcElem
END FUNCTION
' =====================================================================================
' =====================================================================================
' Allocates an array of DOUBLE values then extracts data from a VARIANT into that array.
' GetLastError: Returns S_OK if successful, or an error value otherwise.
' Note Use CoTaskMemFree to release the memory allocated for the array.
' =====================================================================================
PRIVATE FUNCTION CVar.ToDoubleArrayAlloc (BYVAL pprgn AS DOUBLE PTR PTR) AS ULONG
   CVAR_DP("CVAR ToDoubleArrayAlloc")
   DIM pcElem AS ULONG
   SetLastError AfxVariantToDoubleArrayAlloc(@vd, pprgn, @pcElem)
   RETURN pcElem
END FUNCTION
' =====================================================================================

' =====================================================================================
' Extracts an array of string values from a VARIANT structure.
' Returns the number of elements extracted from the source Variant structure.
' GetLastError: Returns S_OK if successful, or an error value otherwise, including the following:
' TYPE_E_BUFFERTOOSMALL: The source VARIANT contained more than crgn values.
' E_INVALIDARG: The VARIANT was not of the appropriate type.
' To retrieve the number of elements in the array call the GetElementCount method.
' =====================================================================================
PRIVATE FUNCTION CVar.ToStringArray (BYVAL prgsz AS PWSTR, BYVAL crgsz AS ULONG) AS ULONG
   CVAR_DP("CVAR ToStringArray")
   DIM pcElem AS ULONG
   SetLastError AfxVariantToStringArray(@vd, prgsz, crgsz, @pcElem)
   RETURN pcElem
END FUNCTION
' =====================================================================================
' =====================================================================================
' Extracts data from a vector structure into a newly-allocated String array.
' GetLastError: Returns S_OK if successful, or an error value otherwise.
' Note Use CoTaskMemFree to free the memory used by each of the strings and the returned array.
' =====================================================================================
PRIVATE FUNCTION CVar.ToStringArrayAlloc (BYVAL pprgsz AS PWSTR PTR) AS ULONG
   CVAR_DP("CVAR ToStringArrayAlloc")
   DIM pcElem AS ULONG
   SetLastError AfxVariantToStringArrayAlloc(@vd, pprgsz, @pcElem)
   RETURN pcElem
END FUNCTION
' =====================================================================================

' ========================================================================================
' Assigns another CVAR.
' ========================================================================================
PRIVATE OPERATOR CVar.Let (BYREF cv AS CVAR)
   CVAR_DP("CVAR LET = CVAR")
'   VariantClear(@vd) : VariantCopy(@vd, @cv)   ' // removed @ operator
   VariantClear(@vd) : VariantCopy(@vd, *cv)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Assigns a VARIANT.
' ========================================================================================
PRIVATE OPERATOR CVar.Let (BYVAL v AS VARIANT)
   CVAR_DP("CVAR LET = VARIANT")
   VariantClear(@vd) : VariantCopy(@vd, @v)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Assigns a WSTRING.
' ========================================================================================
PRIVATE OPERATOR CVar.Let (BYREF wszStr AS WSTRING)
   CVAR_DP("CVAR LET = WSTRING")
   VariantClear(@vd) : vd.vt = VT_BSTR : vd.bstrVal = SysAllocString(wszStr)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Assigns a CWSTR.
' ========================================================================================
PRIVATE OPERATOR CVar.Let (BYVAL pvar AS VARIANT PTR)
   CVAR_DP("CVAR LET = VARIANT PTR")
   VariantClear(@vd) : VariantCopy(@vd, pvar)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Assigns a CURRENCY structure.
' ========================================================================================
PRIVATE OPERATOR CVar.Let (BYVAL cy AS CURRENCY)
   CVAR_DP("CVAR LET - CURRENCY")
   VariantClear(@vd) : vd.vt = VT_CY : vd.cyVal = cy
END OPERATOR
' ========================================================================================

' ========================================================================================
' Assigns a CURRENCY structure.
' ========================================================================================
PRIVATE OPERATOR CVar.Let (BYVAL dec AS DECIMAL)
   CVAR_DP("CVAR LET - DECIMAL")
   VariantClear(@vd) : vd.vt = VT_DECIMAL : vd.decVal = dec
END OPERATOR
' ========================================================================================

' ========================================================================================
' Assigns a CWSTR.
' ========================================================================================
PRIVATE OPERATOR CVar.Let (BYREF cws AS CWSTR)
   CVAR_DP("CVAR LET = CWSTR")
   VariantClear(@vd) : vd.vt = VT_BSTR : vd.bstrVal = SysAllocString(cws)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Assigns a CBSTR.
' ========================================================================================
PRIVATE OPERATOR CVar.Let (BYREF cbs AS CBSTR)
   CVAR_DP("CVAR LET = CBSTR")
   VariantClear(@vd) : vd.vt = VT_BSTR : vd.bstrVal = SysAllocString(cbs)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Assigns a Dispatch pointer
' ========================================================================================
PRIVATE OPERATOR CVar.Let (BYVAL pdisp AS IDispatch PTR)
   CVAR_DP("CVAR LET = IDispatch PTR")
   VariantClear(@vd) : V_VT(@vd) = VT_DISPATCH : V_DISPATCH(@vd) = pdisp : IDispatch_AddRef(pdisp)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Assigns an IUnknown pointer
' ========================================================================================
PRIVATE OPERATOR CVar.Let (BYVAL punk AS IUnknown PTR)
   CVAR_DP("CVAR LET = IUnknown PTR")
   VariantClear(@vd) : V_VT(@vd) = VT_UNKNOWN : V_UNKNOWN(@vd) = punk : IUnknown_AddRef(punk)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Assigns another CVAR.
' ========================================================================================
PRIVATE FUNCTION CVar.Put (BYREF cv AS CVAR) AS HRESULT
   CVAR_DP("CVAR Put = CVAR")
'   VariantClear(@vd) : VariantCopy(@vd, @cv)   ' // removed @ operator
   VariantClear(@vd)
   RETURN VariantCopy(@vd, *cv)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Assigns a VARIANT.
' ========================================================================================
PRIVATE FUNCTION CVar.Put (BYREF v AS VARIANT) AS HRESULT
   CVAR_DP("CVAR Put = VARIANT")
   VariantClear(@vd)
   RETURN VariantCopy(@vd, @v)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Assigns a WSTRING.
' ========================================================================================
PRIVATE SUB CVar.Put (BYREF wszStr AS WSTRING)
   CVAR_DP("CVAR Put = WSTRING")
   VariantClear(@vd)
   vd.vt = VT_BSTR
   vd.bstrVal = SysAllocString(wszStr)
END SUB
' ========================================================================================

' ========================================================================================
' Assigns a VARIANT.
' ========================================================================================
PRIVATE FUNCTION CVar.Put (BYVAL pvar AS VARIANT PTR) As HRESULT
   CVAR_DP("CVAR Put = VARIANT PTR")
   VariantClear(@vd)
   RETURN VariantCopy(@vd, pvar)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Assigns a CWSTR.
' ========================================================================================
PRIVATE SUB CVar.Put (BYREF cws AS CWSTR)
   CVAR_DP("CVAR Put = CWSTR")
   VariantClear(@vd)
   vd.vt = VT_BSTR
   vd.bstrVal = SysAllocString(cws)
END SUB
' ========================================================================================

' ========================================================================================
' Assigns a CBSTR.
' ========================================================================================
PRIVATE SUB CVar.Put (BYREF cbs AS CBSTR)
   CVAR_DP("CVAR Put = CBSTR")
   VariantClear(@vd)
   vd.vt = VT_BSTR
   vd.bstrVal = SysAllocString(cbs)
END SUB
' ========================================================================================

' ========================================================================================
' Assigns a Dispatch pointer
' ========================================================================================
PRIVATE FUNCTION CVar.Put (BYVAL pdisp AS IDispatch PTR, BYVAL fAddRef AS BOOLEAN = FALSE) AS HRESULT
   CVAR_DP("CVAR Put = IDispatch PTR")
   VariantClear(@vd)
   IF pDisp = NULL THEN RETURN E_INVALIDARG
   V_VT(@vd) = VT_DISPATCH
   V_DISPATCH(@vd) = pdisp
   IF fAddRef THEN IDispatch_AddRef(pdisp)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Assigns an IUnknown pointer
' ========================================================================================
PRIVATE FUNCTION CVar.Put (BYVAL punk AS IUnknown PTR, BYVAL fAddRef AS BOOLEAN = FALSE) AS HRESULT
   CVAR_DP("CVAR Put = IUnknown PTR")
   VariantClear(@vd)
   IF punk = NULL THEN RETURN E_INVALIDARG
   V_VT(@vd) = VT_UNKNOWN
   V_UNKNOWN(@vd) = punk
   IF fAddRef THEN IUnknown_AddRef(punk)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Assigns an integer value to a CVAR
' If _vType is wrong, default to VT_I8 (LongInt).
' ========================================================================================
PRIVATE SUB CVar.Put (BYVAL _value AS LONGINT, BYVAL _vType AS WORD)
   CVAR_DP("CVAR Put - vType")
   VariantClear(@vd)
   IF (_vType AND VT_ARRAY) = VT_ARRAY THEN
      ' // NULL array : Suitable for parameters that, instead of being optional,
      ' // require a variant with a null safearray of a given type.
      vd.vt = _vType
   ELSE
      ' // If _value exceeds the limits of max long or max ulong, promote to ulongint
      IF SGN(_value) = -1 AND _value > LONG_MAX THEN
         _vType = VT_I8
      ELSEIF _value > ULONG_MAX THEN
         _vType = VT_I8
      END IF
      SELECT CASE _vType
         CASE VT_BOOL  : vd.vt = VT_BOOL : vd.boolVal = IIF(_value = 0, 0, -1)
         CASE VT_I1    : vd.vt = VT_I1   : vd.cVal = CBYTE(_value)
         CASE VT_UI1   : vd.vt = VT_UI1  : vd.bVal = CUBYTE(_value)
         CASE VT_I2    : vd.vt = VT_I2   : vd.iVal = CSHORT(_value)
         CASE VT_UI2   : vd.vt = VT_UI2  : vd.uiVal = CUSHORT(_value)
         CASE VT_INT   : vd.vt = VT_INT  : vd.intVal = CLNG(_value)
         CASE VT_UINT  : vd.vt = VT_UINT : vd.uintVal = CULNG(_value)
         CASE VT_I4    : vd.vt = VT_I4   : vd.lVal = CLNG(_value)
         CASE VT_UI4   : vd.vt = VT_UI4  : vd.ulVal = CULNG(_value)
         CASE VT_I8    : vd.vt = VT_I8   : vd.llVal = CLNGINT(_value)
         CASE VT_UI8   : vd.vt = VT_UI8  : vd.ullVal = CULNGINT(_value)
         ' // If you need to assign an ULongInt value greater than a LongInt, use PutUlongInt.
         CASE VT_NULL  : vd.vt = VT_NULL
         CASE ELSE     : vd.vt = VT_I8   : vd.llVal = CLNGINT(_value)
      END SELECT
   END IF
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB CVar.Put (BYVAL _value AS LONGINT, BYREF strType AS STRING)
   CVAR_DP("CVAR Put - strType")
   VariantClear(@vd)
   DIM vt AS WORD
   SELECT CASE UCASE(strType)
      CASE "BOOL"     : vd.vt = VT_BOOL : vd.boolVal = IIF(_value = 0, 0, -1)
      CASE "BOOLEAN"  : vd.vt = VT_BOOL : vd.boolVal = IIF(_value = 0, 0, -1)
      CASE "BYTE"     : vd.vt = VT_I1   : vd.cVal = CBYTE(_value)
      CASE "UBYTE"    : vd.vt = VT_UI1  : vd.bVal = CUBYTE(_value)
      CASE "SHORT"    : vd.vt = VT_I2   : vd.iVal = CSHORT(_value)
      CASE "USHORT"   : vd.vt = VT_UI2  : vd.uiVal = CUSHORT(_value)
      CASE "INT"      : vd.vt = VT_INT  : vd.intVal = CLNG(_value)
      CASE "UINT"     : vd.vt = VT_UINT : vd.uintVal = CULNG(_value)
      CASE "LONG"     : vd.vt = VT_I4   : vd.lVal = CLNG(_value)
      CASE "ULONG"    : vd.vt = VT_UI4  : vd.ulVal = CULNG(_value)
      CASE "LONGINT"  : vd.vt = VT_I8   : vd.llVal = CLNGINT(_value)
      CASE "ULONGINT" : vd.vt = VT_UI8  : vd.ullVal = CULNGINT(_value)
      ' // If you need to assign an ULongInt value greater than a LongInt, use PutUlongInt.
      CASE "NULL"     : vd.vt = VT_NULL
      CASE ELSE       : vd.vt = VT_I8   : vd.llVal = CLNGINT(_value)
   END SELECT
END SUB
' ========================================================================================

' ========================================================================================
' Assigns a float value to a CVAR
' If _vType is wrong, default to VT_R8 (double).
' ========================================================================================
PRIVATE SUB CVar.Put (BYVAL _value AS DOUBLE, BYVAL _vType AS WORD = VT_R8)
   CVAR_DP("CVAR Put - vType")
   VariantClear(@vd)
   IF _value > FLT_MAX OR _value < FLT_MIN THEN _vType = VT_R8
   SELECT CASE _vType
      CASE VT_R4 : vd.vt = VT_R4 : vd.fltVal = CSNG(_value)
      CASE VT_R8 : vd.vt = VT_R8 : vd.dblVal = CDBL(_value)
      CASE VT_BOOL, VT_I1, VT_UI1, VT_I2, VT_UI2, VT_INT, VT_UINT, VT_I4, VT_UI4, VT_I8, VT_UI8, VT_NULL
         this.Put CLNGINT(_value), _vType
      CASE ELSE
         vd.vt = VT_R8 : vd.dblVal = CDBL(_value)
   END SELECT
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB CVar.Put (BYVAL _value AS DOUBLE, BYREF strType AS STRING)
   CVAR_DP("CVAR Put - strType")
   VariantClear(@vd)
   SELECT CASE UCASE(strType)
      CASE "SINGLE" : vd.vt = VT_R4   : vd.fltVal = CSNG(_value)
      CASE "FLOAT"  : vd.vt = VT_R4   : vd.fltVal = CSNG(_value)
      CASE "DOUBLE" : vd.vt = VT_R8   : vd.dblVal = CDBL(_value)
      CASE ELSE     : vd.vt = VT_R8   : vd.dblVal = CDBL(_value)
   END SELECT
END SUB
' ========================================================================================

' ========================================================================================
' Assigns a value by reference (a pointer to a variable).
' ========================================================================================
PRIVATE FUNCTION CVar.PutRef (BYVAL _pvar AS ANY PTR, BYVAL _vType AS WORD) AS HRESULT
   CVAR_DP("CVAR PutRef - vType")
   VariantClear(@vd)
   IF _pvar = NULL THEN RETURN E_INVALIDARG
   SELECT CASE _vType
      CASE VT_BOOL      : vd.vt = VT_BOOL OR VT_BYREF      : vd.pboolVal = _pvar
      CASE VT_I1        : vd.vt = VT_I1 OR VT_BYREF        : vd.pcVal = _pvar
      CASE VT_UI1       : vd.vt = VT_UI1 OR VT_BYREF       : vd.pbVal = _pvar
      CASE VT_I2        : vd.vt = VT_I2 OR VT_BYREF        : vd.piVal = _pvar
      CASE VT_UI2       : vd.vt = VT_UI2 OR VT_BYREF       : vd.puiVal = _pvar
      CASE VT_INT       : vd.vt = VT_INT OR VT_BYREF       : vd.pintVal = _pvar
      CASE VT_UINT      : vd.vt = VT_UINT OR VT_BYREF      : vd.puintVal = _pvar
      CASE VT_I4        : vd.vt = VT_I4 OR VT_BYREF        : vd.plVal = _pvar
      CASE VT_UI4       : vd.vt = VT_UI4 OR VT_BYREF       : vd.pulVal = _pvar
      CASE VT_I8        : vd.vt = VT_I8 OR VT_BYREF        : vd.pllVal = _pvar
      CASE VT_UI8       : vd.vt = VT_UI8 OR VT_BYREF       : vd.pullVal = _pvar
      CASE VT_R4        : vd.vt = VT_R4 OR VT_BYREF        : vd.pfltVal = _pvar
      CASE VT_R8        : vd.vt = VT_R8 OR VT_BYREF        : vd.pdblVal = _pvar
      CASE VT_BSTR      : vd.vt = VT_BSTR OR VT_BYREF      : vd.pbstrVal = _pvar
      CASE VT_UNKNOWN   : vd.vt = VT_UNKNOWN OR VT_BYREF   : vd.ppunkVal = _pvar
      CASE VT_DISPATCH  : vd.vt = VT_DISPATCH OR VT_BYREF  : vd.ppdispVal = _pvar
      CASE VT_DECIMAL   : vd.vt = VT_DECIMAL OR VT_BYREF   : vd.pdecVal = _pvar
      CASE VT_CY        : vd.vt = VT_CY OR VT_BYREF        : vd.pcyVal = _pvar
      CASE VT_DATE      : vd.vt = VT_DATE OR VT_BYREF      : vd.pdate = _pvar
      CASE VT_VARIANT   : vd.vt = VT_VARIANT OR VT_BYREF   : vd.pvarVal = _pvar
      CASE VT_SAFEARRAY : vd.vt = VT_SAFEARRAY OR VT_BYREF : vd.pvarVal = _pvar
      CASE VT_ERROR     : vd.vt = VT_ERROR OR VT_BYREF     : vd.pparray = _pvar
      CASE ELSE         : RETURN E_INVALIDARG
   END SELECT
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CVar.PutRef (BYVAL _pvar AS ANY PTR, BYREF strType AS STRING) AS HRESULT
   CVAR_DP("CVAR PutRef - strType")
   VariantClear(@vd)
   IF _pvar = NULL THEN RETURN E_INVALIDARG
   SELECT CASE UCASE(strType)
      CASE "BOOL"      : vd.vt = VT_BOOL OR VT_BYREF      : vd.pboolVal = _pvar
      CASE "BOOLEAN"   : vd.vt = VT_BOOL OR VT_BYREF      : vd.pboolVal = _pvar
      CASE "BYTE"      : vd.vt = VT_I1 OR VT_BYREF        : vd.pcVal = _pvar
      CASE "UBYTE"     : vd.vt = VT_UI1 OR VT_BYREF       : vd.pbVal = _pvar
      CASE "SHORT"     : vd.vt = VT_I2 OR VT_BYREF        : vd.piVal = _pvar
      CASE "USHORT"    : vd.vt = VT_UI2 OR VT_BYREF       : vd.puiVal = _pvar
      CASE "INT"       : vd.vt = VT_INT OR VT_BYREF       : vd.pintVal = _pvar
      CASE "UINT"      : vd.vt = VT_UINT OR VT_BYREF      : vd.puintVal = _pvar
      CASE "LONG"      : vd.vt = VT_I4 OR VT_BYREF        : vd.plVal = _pvar
      CASE "ULONG"     : vd.vt = VT_UI4 OR VT_BYREF       : vd.pulVal = _pvar
      CASE "LONGINT"   : vd.vt = VT_I8 OR VT_BYREF        : vd.pllVal = _pvar
      CASE "ULONGINT"  : vd.vt = VT_UI8 OR VT_BYREF       : vd.pullVal = _pvar
      CASE "SINGLE"    : vd.vt = VT_R4 OR VT_BYREF        : vd.pfltVal = _pvar
      CASE "FLOAT"     : vd.vt = VT_R4 OR VT_BYREF        : vd.pfltVal = _pvar
      CASE "DOUBLE"    : vd.vt = VT_R8 OR VT_BYREF        : vd.pdblVal = _pvar
      CASE "BSTR"      : vd.vt = VT_BSTR OR VT_BYREF      : vd.pbstrVal = _pvar
      CASE "UNKNOWN"   : vd.vt = VT_UNKNOWN OR VT_BYREF   : vd.ppunkVal = _pvar
      CASE "DISPATCH"  : vd.vt = VT_DISPATCH OR VT_BYREF  : vd.ppdispVal = _pvar
      CASE "DECIMAL"   : vd.vt = VT_DECIMAL OR VT_BYREF   : vd.pdecVal = _pvar
      CASE "CY"        : vd.vt = VT_CY OR VT_BYREF        : vd.pcyVal = _pvar
      CASE "CURRENCY"  : vd.vt = VT_CY OR VT_BYREF        : vd.pcyVal = _pvar
      CASE "DATE"      : vd.vt = VT_DATE OR VT_BYREF      : vd.pdate = _pvar
      CASE "VARIANT"   : vd.vt = VT_VARIANT OR VT_BYREF   : vd.pvarVal = _pvar
      CASE "SAFEARRAY" : vd.vt = VT_SAFEARRAY OR VT_BYREF : vd.pvarVal = _pvar
      CASE "ERROR"     : vd.vt = VT_ERROR OR VT_BYREF     : vd.pparray = _pvar
      CASE ELSE        : RETURN E_INVALIDARG
   END SELECT
END FUNCTION
' ========================================================================================

' ========================================================================================
' Assigns a null value.
' ========================================================================================
PRIVATE SUB CVar.PutNull
   CVAR_DP("CVAR PutNull")
   VariantClear(@vd)
   vd.vt = VT_NULL
END SUB
' ========================================================================================

' ========================================================================================
' Assigns a boolean value.
' ========================================================================================
PRIVATE SUB CVar.PutBool (BYVAL _value AS BOOL)
   CVAR_DP("CVAR PutBool")
   VariantClear(@vd)
   vd.vt = VT_BOOL : vd.boolVal = IIF(_value = 0, 0, -1)
END SUB
' ========================================================================================

' ========================================================================================
' Assigns a boolean value.
' ========================================================================================
PRIVATE SUB CVar.PutBoolean (BYVAL _value AS BOOLEAN)
   CVAR_DP("CVAR PutBoolean")
   VariantClear(@vd)
   vd.vt = VT_BOOL : vd.boolVal = IIF(_value = 0, 0, -1)
END SUB
' ========================================================================================

' ========================================================================================
' Assigns a byte value.
' ========================================================================================
PRIVATE SUB CVar.PutByte (BYVAL _value AS BYTE)
   CVAR_DP("CVAR PutByte")
   VariantClear(@vd)
   vd.vt = VT_I1 : vd.cVal = CBYTE(_value)
END SUB
' ========================================================================================

' ========================================================================================
' Assigns an ubyte value.
' ========================================================================================
PRIVATE SUB CVar.PutUByte (BYVAL _value AS UBYTE)
   CVAR_DP("CVAR PutUByte")
   VariantClear(@vd)
   vd.vt = VT_UI1 : vd.bVal = CUBYTE(_value)
END SUB
' ========================================================================================

' ========================================================================================
' Assigns a short value.
' ========================================================================================
PRIVATE SUB CVar.PutShort (BYVAL _value AS SHORT)
   CVAR_DP("CVAR PutShort")
   VariantClear(@vd)
   vd.vt = VT_I2 : vd.iVal = CSHORT(_value)
END SUB
' ========================================================================================

' ========================================================================================
' Assigns an ushort value.
' ========================================================================================
PRIVATE SUB CVar.PutUShort (BYVAL _value AS USHORT)
   CVAR_DP("CVAR PutUShort")
   VariantClear(@vd)
   vd.vt = VT_UI2 : vd.uiVal = CUSHORT(_value)
END SUB
' ========================================================================================

' ========================================================================================
' Assigns an INT_ value.
' ========================================================================================
PRIVATE SUB CVar.PutInt (BYVAL _value AS INT_)
   CVAR_DP("CVAR PutInt")
   VariantClear(@vd)
   vd.vt = VT_INT : vd.intVal = CLNG(_value)
END SUB
' ========================================================================================

' ========================================================================================
' Assigns an UINT value.
' ========================================================================================
PRIVATE SUB CVar.PutUInt (BYVAL _value AS UINT)
   CVAR_DP("CVAR PutUInt")
   VariantClear(@vd)
   vd.vt = VT_UINT : vd.uintVal = CULNG(_value)
END SUB
' ========================================================================================

' ========================================================================================
' Assigns a LONG value.
' ========================================================================================
PRIVATE SUB CVar.PutLong (BYVAL _value AS LONG)
   CVAR_DP("CVAR PutLong")
   VariantClear(@vd)
   vd.vt = VT_I4 : vd.lVal = CLNG(_value)
END SUB
' ========================================================================================

' ========================================================================================
' Assigns an ULONG value.
' ========================================================================================
PRIVATE SUB CVar.PutULong (BYVAL _value AS ULONG)
   CVAR_DP("CVAR PutLong")
   VariantClear(@vd)
   vd.vt = VT_UI4 : vd.ulVal = CULNG(_value)
END SUB
' ========================================================================================

' ========================================================================================
' Assigns a LONGINT value.
' ========================================================================================
PRIVATE SUB CVar.PutLongInt (BYVAL _value AS LONGINT)
   CVAR_DP("CVAR PutLongInt")
   VariantClear(@vd)
   vd.vt = VT_I8 : vd.llVal = CLNGINT(_value)
END SUB
' ========================================================================================

' ========================================================================================
' Assigns an ULONGINT value.
' ========================================================================================
PRIVATE SUB CVar.PutULongInt (BYVAL _value AS ULONGINT)
   CVAR_DP("CVAR PutULongInt")
   VariantClear(@vd)
   vd.vt = VT_UI8 : vd.ullVal = _value
END SUB
' ========================================================================================

' ========================================================================================
' Assigns a SINGLE value.
' ========================================================================================
PRIVATE SUB CVar.PutSingle (BYVAL _value AS SINGLE)
   CVAR_DP("CVAR PutSingle")
   VariantClear(@vd)
   vd.vt = VT_R4 : vd.fltVal = CSNG(_value)
END SUB
' ========================================================================================

' ========================================================================================
' Assigns a FLOAT value.
' ========================================================================================
PRIVATE SUB CVar.PutFloat (BYVAL _value AS SINGLE)
   CVAR_DP("CVAR PutFloat")
   VariantClear(@vd)
   vd.vt = VT_R4 : vd.fltVal = CSNG(_value)
END SUB
' ========================================================================================

' ========================================================================================
' Assigns a DOUBLE value.
' ========================================================================================
PRIVATE SUB CVar.PutDouble (BYVAL _value AS DOUBLE)
   CVAR_DP("CVAR PutDouble")
   VariantClear(@vd)
   vd.vt = VT_R8 : vd.dblVal = CDBL(_value)
END SUB
' ========================================================================================

' ========================================================================================
' Initializes a CVAR from a buffer. Creates a VT_ARRAY OR VT_UI1 variant.
' ========================================================================================
PRIVATE FUNCTION CVar.PutBuffer (BYVAL pv AS ANY PTR, BYVAL cb AS UINT) AS HRESULT
   CVAR_DP("CVAR PutBuffer")
   VariantClear(@vd)
   RETURN AfxVariantFromBuffer(pv, cb, @vd)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Initializes a CVAR from the contents of an UTF8 string.
' ========================================================================================
PRIVATE FUNCTION CVar.PutUtf8 (BYREF strUtf8 AS STRING) AS HRESULT
   CVAR_DP("CVAR PutUtf8")
   VariantClear(@vd)
   vd.vt = VT_BSTR
   vd.bstrVal = SysAllocString(AfxUcode(strUtf8, CP_UTF8))
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Initializes a CVAR with the contents of a safe array. If fAttach = FALSE, the safe
' array is copied; if fAttach = TRUE, the safe array is attached.
' ========================================================================================
PRIVATE FUNCTION CVar.PutSafeArray (BYVAL parray AS SAFEARRAY PTR, BYVAL fAttach AS BOOLEAN = FALSE) AS HRESULT
   CVAR_DP("CVAR PutSafeArray")
   VariantClear(@vd)
   IF parray = NULL THEN RETURN E_INVALIDARG
   DIM vt AS VARTYPE
   DIM hr AS HRESULT = SafeArrayGetVartype(parray, @vt)
   IF hr <> S_OK THEN RETURN hr
   IF fAttach = FALSE THEN
      DIM pcopy AS SAFEARRAY PTR
      hr = SafeArrayCopy(parray, @pcopy)
      IF hr <> S_OK THEN RETURN hr
      vd.vt = vt OR VT_ARRAY
      vd.parray = pcopy
   ELSE
      vd.vt = vt OR VT_ARRAY
      vd.parray = parray
   END IF
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Initializes a CVAR from a a string resource imbedded in an executable file.
' ========================================================================================
PRIVATE FUNCTION CVar.PutResource (BYVAL hinst AS HINSTANCE, BYVAL id AS UINT) AS HRESULT
   CVAR_DP("CVAR PutResource")
   VariantClear(@vd)
   RETURN AfxVariantFromResource(hinst, id, @vd)
END FUNCTION
' ========================================================================================

' =====================================================================================
' Initializes a CVAR with a reference to an UDT.
' =====================================================================================
PRIVATE FUNCTION CVar.PutRecord (BYVAL pIRecordInfo AS IRecordInfo PTR, BYVAL pRec AS ANY PTR) AS HRESULT
   CVAR_DP("CVAR PutRecord")
   VariantClear(@vd)
   IF pIRecordInfo = NULL OR pRec = NULL THEN RETURN E_INVALIDARG
   vd.vt = VT_RECORD
   vd.pvRecord = pRec
   vd.pRecInfo = pIRecordInfo
   RETURN S_OK
END FUNCTION
' =====================================================================================

' =====================================================================================
' Initializes a CVAR as VT_DATE from a string.
' =====================================================================================
PRIVATE FUNCTION CVar.PutDateString (BYVAL pwszDate AS WSTRING PTR, BYVAL lcid AS LCID = 0, BYVAL dwFlags AS ULONG = 0) AS HRESULT
   CVAR_DP("CVAR PutDateString")
   VariantClear(@vd)
   DIM dateOut AS DOUBLE
   DIM hr AS HRESULT = VarDateFromStr(pwszDate, lcid, dwFlags, @dateOut)
   IF hr = S_OK THEN vd.vt = VT_DATE : vd.date = dateOut
   RETURN hr
END FUNCTION
' =====================================================================================

' =====================================================================================
' Initializes a CVAR as VT_DATE from a DATE type.
' =====================================================================================
PRIVATE FUNCTION CVar.PutVbDate (BYVAL vbDate AS DATE_) AS HRESULT
   CVAR_DP("CVAR PutVbDate")
   VariantClear(@vd)
   vd.vt = VT_DATE
   vd.date = vbDate
   RETURN S_OK
END FUNCTION
' =====================================================================================

' =====================================================================================
' Converts a SYSTEMTIME to a VT_DATE variant.
' =====================================================================================
PRIVATE FUNCTION CVar.PutSystemTime (BYVAL st AS SYSTEMTIME PTR) AS BOOLEAN
   CVAR_DP("CVAR PutSystemTime")
   VariantClear(@vd)
   DIM vbDate AS DATE_
   DIM bRes AS BOOLEAN = SystemTimeToVariantTime(st, @vbDate)
   IF bRes THEN vd.vt = VT_DATE : vd.date = vbDate
   RETURN bRes
END FUNCTION
' =====================================================================================

' =====================================================================================
' Initializes a CVAR based on a GUID. The structure is initialized as a VT_BSTR type.
' =====================================================================================
PRIVATE FUNCTION CVar.PutGuid (BYVAL guid AS IID PTR) AS HRESULT
   CVAR_DP("CVAR PutGuid")
   VariantClear(@vd)
   RETURN AfxVariantFromGUIDAsString(guid, @vd)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Initializes a CVAR structure with the contents of a FILETIME structure.
' =====================================================================================
PRIVATE FUNCTION CVar.PutFileTime (BYVAL pft AS FILETIME PTR) AS HRESULT
   CVAR_DP("CVAR PutFileTime")
   VariantClear(@vd)
   RETURN AfxVariantFromFileTime(pft, @vd)
END FUNCTION
' =====================================================================================

' ========================================================================================
' Initializes a VARIANT structure with an array of FILETIME structures.
' ========================================================================================
PRIVATE FUNCTION CVar.PutFileTimeArray (BYVAL prgft AS FILETIME PTR, BYVAL cElems AS ULONG) AS HRESULT
   CVAR_DP("CVAR PutFileTimeArray")
   VariantClear(@vd)
   RETURN AfxVariantFromFileTimeArray(prgft, cElems, @vd)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Initializes a VARIANT structure with a string stored in a STRRET structure.
' =====================================================================================
PRIVATE FUNCTION CVar.PutStrRet (BYVAL pstrret AS STRRET PTR, BYVAL pidl AS PCUITEMID_CHILD) AS HRESULT
   CVAR_DP("CVAR PutStrRet")
   VariantClear(@vd)
   RETURN AfxVariantFromStrRet(pstrret, pidl, @vd)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Initializes a CVAR as VT_DECIMAL from a DECIMAL structure.
' =====================================================================================
PRIVATE FUNCTION CVar.PutDec (BYVAL dec AS DECIMAL) AS HRESULT
   CVAR_DP("CVAR PutDec")
   VariantClear(@vd)
   vd.vt = VT_DECIMAL
   vd.decVal = dec
   RETURN S_OK
END FUNCTION
' =====================================================================================

' =====================================================================================
' Initializes a CVAR as VT_DECIMAL from a string.
' =====================================================================================
PRIVATE FUNCTION CVar.PutDecFromStr (BYVAL pwszIn AS WSTRING PTR, BYVAL lcid AS LCID = 0, BYVAL dwFlags AS ULONG = 0) AS HRESULT
   CVAR_DP("CVAR PutDecFromStr")
   VariantClear(@vd)
   IF pwszIn = NULL THEN RETURN E_INVALIDARG
   DIM dec AS DECIMAL
   DIM hr AS HRESULT = VarDecFromStr(pwszIn, lcid, dwFlags, @dec)
   IF hr = S_OK THEN vd.vt = VT_DECIMAL : vd.decVal = dec
   RETURN hr
END FUNCTION
' =====================================================================================

' =====================================================================================
' Initializes a CVAR as VT_DECIMAL from a double value.
' =====================================================================================
PRIVATE FUNCTION CVar.PutDecFromDouble (BYVAL dbIn AS DOUBLE) AS HRESULT
   CVAR_DP("CVAR PutDecFromStr")
   VariantClear(@vd)
   DIM dec AS DECIMAL
   DIM hr AS HRESULT = VarDecFromR8(dbIn, @dec)
   IF hr = S_OK THEN vd.vt = VT_DECIMAL : vd.decVal = dec
   RETURN hr
END FUNCTION
' =====================================================================================

' =====================================================================================
' Converts a currency value to a CVAR of type VT_DECIMAL.
' =====================================================================================
PRIVATE FUNCTION CVar.PutDecFromCy (BYVAL cyIn AS CY) AS HRESULT
   CVAR_DP("CVAR PutDecFromStr")
   VariantClear(@vd)
   DIM decOut AS DECIMAL
   DIM hr AS HRESULT = VarDecFromCy(cyIn, @decOut)
   IF hr = S_OK THEN vd.vt = VT_DECIMAL : vd.decVal = decOut
   RETURN hr
END FUNCTION
' =====================================================================================

' =====================================================================================
' Initializes a CVAR structure from an array of Boolean values.
' =====================================================================================
PRIVATE FUNCTION CVar.PutBooleanArray (BYVAL prgf AS WINBOOL PTR, BYVAL cElems AS ULONG) AS HRESULT
   CVAR_DP("CVAR PutBooleanArray")
   VariantClear(@vd)
   RETURN AfxVariantFromBooleanArray(prgf, cElems, @vd)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Initializes a CVAR structure from an array of 16-bit integer values.
' =====================================================================================
PRIVATE FUNCTION CVar.PutShortArray (BYVAL prgf AS SHORT PTR, BYVAL cElems AS ULONG) AS HRESULT
   CVAR_DP("CVAR PutShortArray")
   VariantClear(@vd)
   RETURN AfxVariantFromInt16Array(prgf, cElems, @vd)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Initializes a CVAR structure from an array of unsigned 16-bit integer values.
' =====================================================================================
PRIVATE FUNCTION CVar.PutUShortArray (BYVAL prgf AS USHORT PTR, BYVAL cElems AS ULONG) AS HRESULT
   CVAR_DP("CVAR PutUShortArray")
   VariantClear(@vd)
   RETURN AfxVariantFromUInt16Array(prgf, cElems, @vd)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Initializes a CVAR structure from an array of 32-bit integer values.
' =====================================================================================
PRIVATE FUNCTION CVar.PutLongArray (BYVAL prgn AS LONG PTR, BYVAL cElems AS ULONG) AS CVAR
   CVAR_DP("CVAR PutLongArray")
   VariantClear(@vd)
   RETURN AfxVariantFromInt32Array(prgn, cElems, @vd)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Initializes a CVAR structure from an array of 32-bit integer values.
' =====================================================================================
PRIVATE FUNCTION CVar.PutULongArray (BYVAL prgn AS ULONG PTR, BYVAL cElems AS ULONG) AS HRESULT
   CVAR_DP("CVAR PutULongArray")
   VariantClear(@vd)
   RETURN AfxVariantFromUInt32Array(prgn, cElems, @vd)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Initializes a CVAR structure from an array of 64-bit integer values.
' =====================================================================================
PRIVATE FUNCTION CVar.PutLongIntArray (BYVAL prgn AS LONGINT PTR, BYVAL cElems AS ULONG) AS HRESULT
   CVAR_DP("CVAR PutLongIntArray")
   VariantClear(@vd)
   RETURN AfxVariantFromInt64Array(prgn, cElems, @vd)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Initializes a CVAR structure from an array of unsigned 64-bit integer values.
' =====================================================================================
PRIVATE FUNCTION CVar.PutULongIntArray (BYVAL prgn AS ULONGINT PTR, BYVAL cElems AS ULONG) AS HRESULT
   CVAR_DP("CVAR PutLongIntArray")
   VariantClear(@vd)
   RETURN AfxVariantFromUInt64Array(prgn, cElems, @vd)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Initializes a CVAR structure from an array of values of type DOUBLE.
' =====================================================================================
PRIVATE FUNCTION CVar.PutDoubleArray (BYVAL prgn AS DOUBLE PTR, BYVAL cElems AS ULONG) AS HRESULT
   CVAR_DP("CVAR PutLongIntArray")
   VariantClear(@vd)
   RETURN AfxVariantFromDoubleArray(prgn, cElems, @vd)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Initializes a CVAR structure from an array of unicode strings.
' =====================================================================================
PRIVATE FUNCTION CVar.PutStringArray (BYVAL prgsz AS PCWSTR, BYVAL cElems AS ULONG) AS HRESULT
   CVAR_DP("CVAR PutStringArray")
   VariantClear(@vd)
   RETURN AfxVariantFromStringArray(prgsz, cElems, @vd)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Converts the contents of a PROPVARIANT structure to a CVAR.
' =====================================================================================
PRIVATE FUNCTION CVar.PutPropVariant (BYVAL pPropVar AS PROPVARIANT PTR) AS HRESULT
   CVAR_DP("CVAR PutPropVariant")
   VariantClear(@vd)
   RETURN AfxPropVariantToVariant(pPropVar, @vd)
END FUNCTION
' =====================================================================================

' ========================================================================================
' Initializes a CVAR structure from a specified variant element.
' Called VarGetElem (inline function) in the propvarutil.h header, which I find confusing.
' VariantFromVariantArrayElem is more descriptive.
' ========================================================================================
PRIVATE FUNCTION CVar.PutVariantArrayElem (BYVAL pvarIn AS VARIANT PTR, BYVAL iElem AS ULONG) AS CVAR
   CVAR_DP("CVAR PutVariantArrayElem")
   VariantClear(@vd)
   RETURN AfxVariantFromVariantArrayElem(pvarIn, iElem, @vd)
END FUNCTION
' ========================================================================================

' =====================================================================================
' Returns the number of dimensions for variants of type VT_ARRAY; returns 0 otherwise.
' =====================================================================================
PRIVATE FUNCTION CVar.GetDim () AS ULONG
   CVAR_DP("CVAR GetDim")
   IF (vd.vt AND VT_ARRAY) = VT_ARRAY THEN
      IF vd.parray THEN RETURN SafeArrayGetDim(vd.parray)
   END IF
END FUNCTION
' =====================================================================================

' =====================================================================================
' Returns the lower bound for the specified dimension of the safe array for variants of
' type VT_ARRAY; returns 0 otherwise.
' =====================================================================================
PRIVATE FUNCTION CVar.GetLBound (BYVAL nDim AS UINT = 1) AS LONG
   CVAR_DP("CVAR GetLBound")
   DIM hr AS HRESULT
   IF (vd.vt AND VT_ARRAY) = VT_ARRAY THEN
      DIM nBound AS LONG
      IF vd.parray THEN hr = SafeArrayGetLBound(vd.parray, nDim, @nBound)
      IF hr = S_OK THEN RETURN nBound
   END IF
END FUNCTION
' =====================================================================================

' Returns the upper bound for the specified dimension of the safe array for variants of
' type VT_ARRAY; returns -1 otherwise.
' =====================================================================================
PRIVATE FUNCTION CVar.GetUBound (BYVAL nDim AS UINT = 1) AS LONG
   CVAR_DP("CVAR GetUBound")
   DIM hr AS HRESULT
   IF (vd.vt AND VT_ARRAY) = VT_ARRAY THEN
      DIM nBound AS LONG
      IF vd.parray THEN hr = SafeArrayGetUBound(vd.parray, nDim, @nBound)
      IF hr = S_OK THEN RETURN nBound
   END IF
   RETURN -1
END FUNCTION
' =====================================================================================

' =====================================================================================
' Returns the number of elements for variants of type VT_ARRAY; otherwise, returns 1.
' =====================================================================================
PRIVATE FUNCTION CVar.GetElementCount () AS ULONG
   CVAR_DP("CVAR GetElementCount")
   RETURN AfxVariantGetElementCount(@vd)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Extracts a single Variant element from a safe array of variants.
' =====================================================================================
PRIVATE FUNCTION CVar.GetVariantElem (BYVAL iElem AS ULONG) AS CVAR
   CVAR_DP("CVAR GetVariantElem")
   DIM hr AS HRESULT, vItem AS VARIANT
   IF vd.vt = VT_ARRAY OR VT_VARIANT THEN
      DIM idx AS LONG = iElem
      SetLastError(SafeArrayGetElement(vd.parray, @idx, @vItem))
   END IF
   RETURN vItem
END FUNCTION
' =====================================================================================

' =====================================================================================
' Extracts a single Boolean element from a CVAR.
' =====================================================================================
PRIVATE FUNCTION CVar.GetBooleanElem (BYVAL iElem AS ULONG) AS BOOLEAN
   CVAR_DP("CVAR GetBooleanElem")
   DIM pfVal AS WINBOOL
   SetLastError AfxVariantGetBooleanElem(@vd, iElem, @pfVal)
   RETURN pfVal
END FUNCTION
' =====================================================================================

' =====================================================================================
' Extracts a single Int16 element from a CVAR.
' =====================================================================================
PRIVATE FUNCTION CVar.GetShortElem (BYVAL iElem AS ULONG) AS SHORT
   CVAR_DP("CVAR GetShortElem")
   DIM pnVal AS SHORT
   SetLastError AfxVariantGetInt16Elem(@vd, iElem, @pnVal)
   RETURN pnVal
END FUNCTION
' =====================================================================================

' =====================================================================================
' Extracts a single unsigned Int16 element from a CVAR.
' =====================================================================================
PRIVATE FUNCTION CVar.GetUShortElem (BYVAL iElem AS ULONG) AS USHORT
   CVAR_DP("CVAR GetUShortElem")
   DIM pnVal AS USHORT
   SetLastError AfxVariantGetUInt16Elem(@vd, iElem, @pnVal)
   RETURN pnVal
END FUNCTION
' =====================================================================================

' =====================================================================================
' Extracts a single Int32 element from a CVAR.
' =====================================================================================
PRIVATE FUNCTION CVar.GetLongElem (BYVAL iElem AS ULONG) AS LONG
   CVAR_DP("CVAR GetLongElem")
   DIM pnVal AS LONG
   SetLastError AfxVariantGetInt32Elem(@vd, iElem, @pnVal)
   RETURN pnVal
END FUNCTION
' =====================================================================================

' =====================================================================================
' Extracts a single unsigned Int32 element from a CVAR.
' =====================================================================================
PRIVATE FUNCTION CVar.GetULongElem (BYVAL iElem AS ULONG) AS ULONG
   CVAR_DP("CVAR GetULongElem")
   DIM pnVal AS ULONG
   SetLastError AfxVariantGetUInt32Elem(@vd, iElem, @pnVal)
   RETURN pnVal
END FUNCTION
' =====================================================================================

' =====================================================================================
' Extracts a single Int64 element from a CVAR.
' =====================================================================================
PRIVATE FUNCTION CVar.GetLongIntElem (BYVAL iElem AS ULONG) AS LONGINT
   CVAR_DP("CVAR GetLongIntElem")
   DIM pnVal AS LONGINT
   SetLastError AfxVariantGetInt64Elem(@vd, iElem, @pnVal)
   RETURN pnVal
END FUNCTION
' =====================================================================================

' =====================================================================================
' Extracts a single UInt64 element from a CVAR.
' =====================================================================================
PRIVATE FUNCTION CVar.GetULongIntElem (BYVAL iElem AS ULONG) AS ULONGINT
   CVAR_DP("CVAR GetULongIntElem")
   DIM pnVal AS ULONGINT
   SetLastError AfxVariantGetUInt64Elem(@vd, iElem, @pnVal)
   RETURN pnVal
END FUNCTION
' =====================================================================================

' =====================================================================================
' Extracts one double element from a CVAR.
' =====================================================================================
PRIVATE FUNCTION CVar.GetDoubleElem (BYVAL iElem AS ULONG) AS DOUBLE
   CVAR_DP("CVAR GetDoubleElem")
   DIM pnVal AS DOUBLE
   SetLastError AfxVariantGetDoubleElem(@vd, iElem, @pnVal)
   RETURN pnVal
END FUNCTION
' =====================================================================================

' =====================================================================================
' Extracts an string element from a VARIANT.
' =====================================================================================
PRIVATE FUNCTION CVar.GetStringElem (BYVAL iElem AS ULONG) AS CWSTR
   CVAR_DP("CVAR GetStringElem")
   DIM pwszVal AS WSTRING PTR, cwsVal AS CWSTR
   SetLastError AfxVariantGetStringElem(@vd, iElem, @pwszVal)
   IF pwszVal THEN
      cwsVal = *pwszVal
      CoTaskMemFree pwszVal
   END IF
   RETURN cwsVal
END FUNCTION
' =====================================================================================

' =====================================================================================
' Formats a variant containing numbers into a string form.
' =====================================================================================
PRIVATE FUNCTION CVar.FormatNumber (BYVAL iNumDig AS LONG = -1, BYVAL iIncLead AS LONG = -2, _
   BYVAL iUseParens AS LONG = -2, BYVAL iGroup AS LONG = -2, BYVAL dwFlags AS DWORD = 0) AS CWSTR
   CVAR_DP("CVAR FormatNumber")
   DIM cwsOut AS CWSTR, bstrOut AS AFX_BSTR
   SetLastError VarFormatNumber(@vd, iNumDig, iIncLead, iUseParens, iGroup, dwFlags, @bstrOut)
   cwsOut = *bstrOut
   SysFreeString bstrOut
   RETURN cwsOut
END FUNCTION
' =====================================================================================

' ========================================================================================
' Clears the contents of the variant data and sets the vt field to VT_EMPTY.
' ========================================================================================
PRIVATE SUB CVar.Clear
   CVAR_DP("CVAR Clear")
   VariantClear(@vd)
END SUB
' ========================================================================================

' ========================================================================================
' Attaches a variant to this class. The source variant is marked as empty.
' ========================================================================================
PRIVATE FUNCTION CVar.Attach (BYVAL pvar AS VARIANT PTR) AS HRESULT
   CVAR_DP("CVAR ATTACH - VARIANT PTR")
   IF pvar = NULL THEN RETURN E_INVALIDARG
   VariantClear @vd
   ' // Copy the contents and give control to CVar
   DIM pdest AS ANY PTR = memcpy(@vd, pvar, SIZEOF(VARIANT))
   IF pdest = NULL THEN RETURN E_FAIL
   ' // Mark the source variant as VT_EMPTY instead of clearing it with VariantClear
   ' // because we aren't making a duplicate of the contents, but transfering ownership.
   pvar->vt = VT_EMPTY
   RETURN S_OK
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CVar.Attach (BYREF v AS VARIANT) AS HRESULT
   CVAR_DP("CVAR ATTACH - BYREF VARIANT")
   VariantClear @vd
   ' // Copy the contents and give control to CVar
   DIM pdest AS ANY PTR = memcpy(@vd, @v, SIZEOF(VARIANT))
   IF pdest = NULL THEN RETURN E_FAIL
   ' // Mark the source variant as VT_EMPTY instead of clearing it with VariantClear
   ' // because we aren't making a duplicate of the contents, but transfering ownership.
   v.vt = VT_EMPTY
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Detaches the variant data from this class and returs it as a VARIANT.
' Don't clear vd with VariantClear because we are transfering ownership.
' ========================================================================================
PRIVATE FUNCTION CVar.Detach (BYVAL pvar AS VARIANT PTR) AS HRESULT
   CVAR_DP("CVAR DETACH - VARIANT PTR")
   IF pvar = NULL THEN RETURN E_INVALIDARG
   VariantClear(pvar)
   DIM pdest AS ANY PTR = memcpy(pvar, @vd, SIZEOF(VARIANT))
   IF pdest = NULL THEN RETURN E_FAIL
   vd.vt = VT_EMPTY
   RETURN S_OK
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CVar.Detach (BYREF v AS VARIANT) AS HRESULT
   CVAR_DP("CVAR DETACH - BYREF VARIANT")
   VariantClear(@v)
   DIM pdest AS ANY PTR = memcpy(@v, @vd, SIZEOF(VARIANT))
   IF pdest = NULL THEN RETURN E_FAIL
   vd.vt = VT_EMPTY
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the VARIANT type.
' See: https://msdn.microsoft.com/en-us/library/windows/desktop/ms221170(v=vs.85).aspx
' ========================================================================================
PRIVATE FUNCTION CVar.vType () AS VARTYPE
   CVAR_DP("CVAR vType")
   FUNCTION = vd.vt
END FUNCTION
' ========================================================================================

' =====================================================================================
' Converts the variant from one type to another.
' =====================================================================================
PRIVATE FUNCTION CVar.ChangeType (BYVAL vtNew AS VARTYPE, BYVAL wFlags AS USHORT = 0) AS HRESULT
   CVAR_DP("CVAR ChangeType")
   RETURN VariantChangeType(@vd, @vd, wFlags, vtNew)
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CVar.ChangeTypeEx (BYVAL vtNew AS VARTYPE, BYVAL lcid AS LCID = 0, BYVAL wFlags AS USHORT = 0) AS HRESULT
   CVAR_DP("CVAR ChangeTypeEx")
   RETURN VariantChangeTypeEx(@vd, @vd, lcid, wFlags, vtNew)
END FUNCTION
' =====================================================================================

' =====================================================================================
PRIVATE FUNCTION CVar.Round (BYREF cv AS CVAR, BYVAL cDecimals AS LONG) AS CVAR
   CVAR_DP("CVAR Round")
   DIM cvRes AS CVAR
   VarRound(@cv.vd, cDecimals, cvRes.vptr)
   RETURN cvRes
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CVar.Round (BYVAL cDecimals AS LONG) AS CVAR
   CVAR_DP("CVAR Round=")
   DIM cvRes AS CVAR = vd
   IF VarRound(@vd, cDecimals, cvRes.vptr) = S_OK THEN vd = cvRes.vd
   RETURN cvRes
END FUNCTION
' =====================================================================================

' =====================================================================================
' Converts the string to a 32bit integer
' =====================================================================================
PRIVATE FUNCTION CVar.ValLong () AS LONG
   RETURN .ValInt(**AfxVarToStr(@vd))
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CVar.ValInt () AS LONG
   RETURN .ValInt(**AfxVarToStr(@vd))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Converts the string to an unsigned 32bit integer
' =====================================================================================
PRIVATE FUNCTION CVar.ValULong () AS ULONG
   RETURN .ValUInt(**AfxVarToStr(@vd))
END FUNCTION
' =====================================================================================
PRIVATE FUNCTION CVar.ValUInt () AS ULONG
   RETURN .ValUInt(**AfxVarToStr(@vd))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Converts the string to a 64bit integer
' =====================================================================================
PRIVATE FUNCTION CVar.ValLongInt () AS LONGINT
   RETURN .ValLng(**AfxVarToStr(@vd))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Converts the string to an unsigned 64bit integer
' =====================================================================================
PRIVATE FUNCTION CVar.ValULongInt () AS ULONGINT
   RETURN .ValULng(**AfxVarToStr(@vd))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Converts the string to a floating point number (DOUBLE)
' =====================================================================================
PRIVATE FUNCTION CVar.ValDouble () AS DOUBLE
   RETURN .VAL(**AfxVarToStr(@vd))
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CVar.Value () AS DOUBLE
   RETURN .VAL(**AfxVarToStr(@vd))
END FUNCTION
' =====================================================================================

' ========================================================================================
' Math operators
' ========================================================================================
PRIVATE OPERATOR CVar.+= (BYREF cv AS CVAR)
   IF vd.vt = VT_BSTR AND cv.vd.vt = VT_BSTR THEN
      ' // One of the values is a string, so concatenate
      VarCat(@vd, @cv.vd, @vd)
   ELSE
      ' // Add
      DIM vRes AS VARIANT
      IF VarAdd(@vd, @cv.vd, @vRes) = S_OK THEN VariantCopy(@vd, @vRes)
   END IF
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CVar.&= (BYREF cv AS CVAR)
   IF vd.vt = VT_BSTR AND cv.vd.vt = VT_BSTR THEN
      ' // One of the values is a string, so concatenate
      VarCat(@vd, @cv.vd, @vd)
   ELSE
      ' // Add
      DIM vRes AS VARIANT
      IF VarAdd(@vd, @cv.vd, @vRes) = S_OK THEN VariantCopy(@vd, @vRes)
   END IF
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CVar.-= (BYREF cv AS CVAR)
   IF vd.vt <> VT_BSTR AND cv.vd.vt <> VT_BSTR THEN
      DIM vRes AS VARIANT
      IF VarSub(@vd, @cv.vd, @vRes) = S_OK THEN VariantCopy(@vd, @vRes)
   END IF
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CVar.*= (BYREF cv AS CVAR)
   IF vd.vt <> VT_BSTR AND cv.vd.vt <> VT_BSTR THEN
      DIM vRes AS VARIANT
      IF VarMul(@vd, @cv.vd, @vRes) = S_OK THEN VariantCopy(@vd, @vRes)
   END IF
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CVar./= (BYREF cv AS CVAR)
   IF vd.vt <> VT_BSTR AND cv.vd.vt <> VT_BSTR THEN
      DIM vRes AS VARIANT
      IF VarDiv(@vd, @cv.vd, @vRes) = S_OK THEN VariantCopy(@vd, @vRes)
   END IF
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CVar.\= (BYREF cv AS CVAR)
   IF vd.vt <> VT_BSTR AND cv.vd.vt <> VT_BSTR THEN
      DIM vRes AS VARIANT
      IF VarIDiv(@vd, @cv.vd, @vRes) = S_OK THEN VariantCopy(@vd, @vRes)
   END IF
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CVar.^= (BYREF cv AS CVAR)
   IF vd.vt <> VT_BSTR AND cv.vd.vt <> VT_BSTR THEN
      DIM vRes AS VARIANT
      IF VarPow(@vd, @cv.vd, @vRes) = S_OK THEN VariantCopy(@vd, @vRes)
   END IF
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CVar.And= (BYREF cv AS CVAR)
   IF vd.vt <> VT_BSTR AND cv.vd.vt <> VT_BSTR THEN
      DIM vRes AS VARIANT
      IF VarAnd(@vd, @cv.vd, @vRes) = S_OK THEN VariantCopy(@vd, @vRes)
   END IF
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CVar.Or= (BYREF cv AS CVAR)
   IF vd.vt <> VT_BSTR AND cv.vd.vt <> VT_BSTR THEN
      DIM vRes AS VARIANT
      IF VarOr(@vd, @cv.vd, @vRes) = S_OK THEN VariantCopy(@vd, @vRes)
   END IF
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CVar.Xor= (BYREF cv AS CVAR)
   IF vd.vt <> VT_BSTR AND cv.vd.vt <> VT_BSTR THEN
      DIM vRes AS VARIANT
      IF VarXor(@vd, @cv.vd, @vRes) = S_OK THEN VariantCopy(@vd, @vRes)
   END IF
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CVar.Eqv= (BYREF cv AS CVAR)
   IF vd.vt <> VT_BSTR AND cv.vd.vt <> VT_BSTR THEN
      DIM vRes AS VARIANT
      IF VarEqv(@vd, @cv.vd, @vRes) = S_OK THEN VariantCopy(@vd, @vRes)
   END IF
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CVar.Imp= (BYREF cv AS CVAR)
   IF vd.vt <> VT_BSTR AND cv.vd.vt <> VT_BSTR THEN
      DIM vRes AS VARIANT
      IF VarImp(@vd, @cv.vd, @vRes) = S_OK THEN VariantCopy(@vd, @vRes)
   END IF
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CVar.Mod= (BYREF cv AS CVAR)
   IF vd.vt <> VT_BSTR AND cv.vd.vt <> VT_BSTR THEN
      DIM vRes AS VARIANT
      IF VarMod(@vd, @cv.vd, @vRes) = S_OK THEN VariantCopy(@vd, @vRes)
   END IF
END OPERATOR
' ========================================================================================

' ########################################################################################
'                               *** HELPER FUNCTIONS ***
' ########################################################################################

' ========================================================================================
' Extracts the contents of a CVAR to a CWSTR.
' ========================================================================================
PRIVATE FUNCTION AfxCVarToStr OVERLOAD (BYREF cv AS CVAR) AS CWSTR
   RETURN AfxVarToStr(@cv.vd)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION AfxCVarToStr OVERLOAD (BYVAL pcv AS CVAR PTR) AS CWSTR
   RETURN AfxVarToStr(@pcv->vd)
END FUNCTION
' ========================================================================================
#define CVAR_TOSTR AfxCVarToStr

END NAMESPACE

' ########################################################################################
'                               *** GLOBAL OPERATORS ***
' ########################################################################################

' // Outside a namespace because they are global
using Afx

' ========================================================================================
PRIVATE OPERATOR & (BYREF cv1 AS CVAR, BYREF cv2 AS CVAR) AS CVAR
   DIM cvRes AS CVAR
   VarCat(@cv1.vd, @cv2.vd, cvRes.vptr)
   OPERATOR = cvRes
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Left OVERLOAD (BYREF cv AS CVAR, BYVAL nChars AS INTEGER) AS CWSTR
   DIM cws AS CWSTR = cv.wstr
   RETURN LEFT(**cws, nChars)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Right OVERLOAD (BYREF cv AS CVAR, BYVAL nChars AS INTEGER) AS CWSTR
   DIM cws AS CWSTR = cv.wstr
   RETURN RIGHT(**cws, nChars)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Val OVERLOAD (BYREF cv AS CVAR) AS DOUBLE
   DIM cws AS CWSTR = cv.wstr
   RETURN .VAL(**cws)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Math operators
' ========================================================================================
PRIVATE OPERATOR + (BYREF cv1 AS CVAR, BYREF cv2 AS CVAR) AS CVAR
   DIM cvRes AS CVAR
   IF cv1.vd.vt = VT_BSTR OR cv2.vd.vt = VT_BSTR THEN
      ' // One of the values is a string, so concatenate
      cvRes = cv1 & cv2
   ELSE
      ' // Add
      VarAdd(@cv1.vd, @cv2.vd, cvRes.vptr)
   END IF
   OPERATOR = cvRes
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR - (BYREF cv1 AS CVAR, BYREF cv2 AS CVAR) AS CVAR
   ' // Subtract
   DIM cvRes AS CVAR
   VarSub(@cv1.vd, @cv2.vd, cvRes.vptr)
   OPERATOR = cvRes
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR * (BYREF cv1 AS CVAR, BYREF cv2 AS CVAR) AS CVAR
   ' // Multiply
   DIM cvRes AS CVAR
   VarMul(@cv1.vd, @cv2.vd, cvRes.vptr)
   OPERATOR = cvRes
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR / (BYREF cv1 AS CVAR, BYREF cv2 AS CVAR) AS CVAR
   ' // Divide
   DIM cvRes AS CVAR
   VarDiv(@cv1.vd, @cv2.vd, cvRes.vptr)
   OPERATOR = cvRes
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR \ (BYREF cv1 AS CVAR, BYREF cv2 AS CVAR) AS CVAR
   ' // Integer divide
   DIM cvRes AS CVAR
   VarIDiv(@cv1.vd, @cv2.vd, cvRes.vptr)
   OPERATOR = cvRes
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR = (BYREF cv1 AS CVAR, BYREF cv2 AS CVAR) AS BOOLEAN
   ' // Equality
   RETURN (VarCmp(@cv1.vd, @cv2.vd, GetThreadLocale(), 0) = VARCMP_EQ)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR <> (BYREF cv1 AS CVAR, BYREF cv2 AS CVAR) AS BOOLEAN
   ' // Inequality
   RETURN (VarCmp(@cv1.vd, @cv2.vd, GetThreadLocale(), 0) <> VARCMP_EQ)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR < (BYREF cv1 AS CVAR, BYREF cv2 AS CVAR) AS BOOLEAN
   RETURN (VarCmp(@cv1.vd, @cv2.vd, GetThreadLocale(), 0) = VARCMP_LT)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR > (BYREF cv1 AS CVAR, BYREF cv2 AS CVAR) AS BOOLEAN
   RETURN (VarCmp(@cv1.vd, @cv2.vd, GetThreadLocale(), 0) = VARCMP_GT)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR <= (BYREF cv1 AS CVAR, BYREF cv2 AS CVAR) AS BOOLEAN
   DIM hr AS HRESULT = VarCmp(@cv1.vd, @cv2.vd, GetThreadLocale(), 0)
   RETURN (hr = VARCMP_LT) OR (hr = VARCMP_EQ)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR >= (BYREF cv1 AS CVAR, BYREF cv2 AS CVAR) AS BOOLEAN
   DIM hr AS HRESULT = VarCmp(@cv1.vd, @cv2.vd, GetThreadLocale(), 0)
   RETURN (hr = VARCMP_GT) OR (hr = VARCMP_EQ)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR - (BYREF cv AS CVAR) AS CVAR
   ' // Negate
   DIM cvRes AS CVAR
   VarNeg(@cv.vd, cvRes.vptr)
   OPERATOR = cvRes
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR Not (BYREF cv AS CVAR) AS INTEGER
   ' // Bitwise negation
   DIM cvRes AS CVAR
   VarNot(@cv.vd, cvRes.vptr)
   OPERATOR = VAL(cvRes)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR And (BYREF cv1 AS CVAR, BYREF cv2 AS CVAR) AS INTEGER
   DIM cvRes AS CVAR
   VarAnd(@cv1.vd, @cv2.vd, cvRes.vptr)
   OPERATOR = VAL(cvRes)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR Or (BYREF cv1 AS CVAR, BYREF cv2 AS CVAR) AS INTEGER
   DIM cvRes AS CVAR
   VarOr(@cv1.vd, @cv2.vd, cvRes.vptr)
   OPERATOR = VAL(cvRes)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR Xor (BYREF cv1 AS CVAR, BYREF cv2 AS CVAR) AS INTEGER
   DIM cvRes AS CVAR
   VarXor(@cv1.vd, @cv2.vd, cvRes.vptr)
   OPERATOR = VAL(cvRes)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR Mod (BYREF cv1 AS CVAR, BYREF cv2 AS CVAR) AS INTEGER
   DIM cvRes AS CVAR
   VarMod(@cv1.vd, @cv2.vd, cvRes.vptr)
   OPERATOR = VAL(cvRes)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR Imp (BYREF cv1 AS CVAR, BYREF cv2 AS CVAR) AS INTEGER
   DIM cvRes AS CVAR
   VarImp(@cv1.vd, @cv2.vd, cvRes.vptr)
   OPERATOR = VAL(cvRes)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR Eqv (BYREF cv1 AS CVAR, BYREF cv2 AS CVAR) AS INTEGER
   DIM cvRes AS CVAR
   VarEqv(@cv1.vd, @cv2.vd, cvRes.vptr)
   OPERATOR = VAL(cvRes)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR ^ (BYREF cv1 AS CVAR, BYREF cv2 AS CVAR) AS DOUBLE
   DIM cvRes AS CVAR
   VarPow(@cv1.vd, @cv2.vd, cvRes.vptr)
   OPERATOR = VAL(cvRes)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR Abs (BYREF cv AS CVAR) AS DOUBLE
   DIM cvRes AS CVAR
   VarAbs(@cv.vd, cvRes.vptr)
   OPERATOR = VAL(cvRes)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR Fix (BYREF cv AS CVAR) AS DOUBLE
   DIM cvRes AS CVAR
   VarFix(@cv.vd, cvRes.vptr)
   OPERATOR = VAL(cvRes)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR Int (BYREF cv AS CVAR) AS DOUBLE
   DIM cvRes AS CVAR
   VarInt(@cv.vd, cvRes.vptr)
   OPERATOR = VAL(cvRes)
END OPERATOR
' ========================================================================================
