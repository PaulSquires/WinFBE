' ########################################################################################
' Microsoft Windows
' File: CPropVar.inc
' Contents: Windows PROPVARIANT class.
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2017 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "windows.bi"
#include once "crt/limits.bi"
#include once "win/propidl.bi"
#include once "Afx/AfxCOM.inc"

NAMESPACE Afx

' ========================================================================================
' Macro for debug
' To allow debugging, define _CPROPVAR_DEBUG_ 1 in your application before including this file.
' ========================================================================================
#ifndef _CPROPVAR_DEBUG_
   #define _CPROPVAR_DEBUG_ 0
#ENDIF
#ifndef _CPROPVAR_DP_
   #define _CPROPVAR_DP_ 1
   #MACRO CPROPVAR_DP(st)
      #IF (_CPROPVAR_DEBUG_ = 1)
         OutputDebugStringW(st)
      #ENDIF
   #ENDMACRO
#ENDIF
' ========================================================================================

TYPE PROPVAR_CHANGE_FLAGS AS LONG
ENUM   ' tagPROPVAR_CHANGE_FLAGS
   PVCHF_DEFAULT           = &h00000000
   PVCHF_NOVALUEPROP       = &h00000001   ' // Maps to VARIANT_NOVALUEPROP for VariantChangeType
   PVCHF_ALPHABOOL         = &h00000002   ' // Maps to VARIANT_ALPHABOOL for VariantChangeType
   PVCHF_NOUSEROVERRIDE    = &h00000004   ' // Maps to VARIANT_NOUSEROVERRIDE for VariantChangeType
   PVCHF_LOCALBOOL         = &h00000008   ' // Maps to VARIANT_LOCALBOOL for VariantChangeType
   PVCHF_NOHEXSTRING       = &h00000010   ' // Don't convert a string that looks like hexadecimal (0xABCD) to the numerical equivalent
END ENUM

' // The progress bar property control uses a specially formatted PROPVARIANT to convey the look of the progress bar
' // propvar.vt = VT_UI4
' // propvar.caul.pElems[0] = current progress
' // propvar.caul.pElems[1] = total progress
' // propvar.caul.pElems[2] = DRAWPROGRESSFLAGS (see below);
ENUM DRAWPROGRESSFLAGS
   DPF_NONE                = &h0   ' // No progress flags.
   DPF_MARQUEE             = &h1   ' // The progress bar should draw in marquee mode.
   DPF_MARQUEE_COMPLETE    = &h2   ' // The marquee format progress bar has completed.
   DPF_ERROR               = &h4   ' // The progress bar should be drawn in the error state.
   DPF_WARNING             = &h8   ' // The progress bar should be drawn in the warning state.
   DPF_STOPPED             = &h10  ' // The progress bar is stopped.
END ENUM

' ========================================================================================
' Frees the memory and references used by an array of PROPVARIANT structures stored in an array.
' ========================================================================================
PRIVATE SUB AfxClearPropVariantArray (BYVAL rgPropVar AS PROPVARIANT PTR, BYVAL cVars AS UINT)
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT SUB
   DIM pClearPropVariantArray AS SUB (BYVAL rgPropVar AS PROPVARIANT PTR, BYVAL cVars AS UINT)
   pClearPropVariantArray = DyLibSymbol(pLib, "ClearPropVariantArray")
   IF pClearPropVariantArray THEN pClearPropVariantArray(rgPropVar, cVars)
   DyLibFree(pLib)
END SUB
' ========================================================================================

' ========================================================================================
' Serializes a specified PROPVARIANT structure, creating a SERIALIZEDPROPERTYVALUE structure.
' ========================================================================================
PRIVATE FUNCTION AfxStgSerializePropVariant (BYVAL ppropvar AS PROPVARIANT PTR, BYVAL ppProp AS SERIALIZEDPROPERTYVALUE PTR, BYVAL pcb AS ULONG PTR) AS HRESULT
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pStgSerializePropVariant AS FUNCTION (BYVAL ppropvar AS PROPVARIANT PTR, BYVAL ppProp AS SERIALIZEDPROPERTYVALUE PTR, BYVAL pcb AS ULONG PTR) AS HRESULT
   pStgSerializePropVariant = DyLibSymbol(pLib, "StgSerializePropVariant")
   IF pStgSerializePropVariant THEN FUNCTION = pStgSerializePropVariant(ppropvar, ppProp, pcb)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Serializes a specified PROPVARIANT structure, creating a SERIALIZEDPROPERTYVALUE structure.
' ========================================================================================
PRIVATE FUNCTION AfxStgDeserializePropVariant (BYVAL pProp AS SERIALIZEDPROPERTYVALUE PTR, BYVAL cbMax AS ULONG PTR, BYVAL ppropvar AS PROPVARIANT PTR) AS HRESULT
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pStgDeserializePropVariant AS FUNCTION (BYVAL pProp AS SERIALIZEDPROPERTYVALUE PTR, BYVAL cbMax AS ULONG PTR, BYVAL ppropvar AS PROPVARIANT PTR) AS HRESULT
   pStgDeserializePropVariant = DyLibSymbol(pLib, "StgDeserializePropVariant")
   IF pStgDeserializePropVariant THEN FUNCTION = pStgDeserializePropVariant(pprop, cbMax, ppropvar)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Coerces a value stored as a PROPVARIANT structure to an equivalent value of a different type.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantChangeType (BYVAL ppropvarDest AS PROPVARIANT PTR, BYVAL propvarSrc AS PROPVARIANT PTR, BYVAL wFlags AS USHORT, BYVAL vt AS VARTYPE) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pPropVariantChangeType AS FUNCTION (BYVAL ppropvarDest AS PROPVARIANT PTR, BYVAL propvarSrc AS PROPVARIANT PTR, BYVAL wFlags AS USHORT, BYVAL vt AS VARTYPE) AS HRESULT
   pPropVariantChangeType = DyLibSymbol(pLib, "PropVariantChangeType")
   IF pPropVariantChangeType = NULL THEN FUNCTION = E_FAIL
   FUNCTION = pPropVariantChangeType(ppropvarDest, propvarSrc, wflags, vt)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Initializes a PROPVARIANT structure with the contents of a buffer.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantFromBuffer (BYVAL pv AS VOID PTR, BYVAL cb AS UINT, BYVAL pPropVar AS PROPVARIANT PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pInitPropVariantFromBuffer AS FUNCTION (BYVAL pv AS VOID PTR, BYVAL cb AS UINT, BYVAL pVar AS PROPVARIANT PTR) AS HRESULT
   pInitPropVariantFromBuffer = DyLibSymbol(pLib, "InitPropVariantFromBuffer")
   IF pInitPRopVariantFromBuffer THEN FUNCTION = pInitPropVariantFromBuffer(pv, cb, pPropVar)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts the buffer value from a PROPVARIANT structure of type VT_VECTOR OR VT_UI1 or
' VT_ARRRAY OR VT_UI1.
' Parameters:
' - pPropVarIn : [in] Reference to a source VARIANT structure.
' - pv         : [out] Pointer to a buffer of length cb bytes.
' - cb         : [in] The size of the pv buffer, in bytes. The buffer should be the same size as
'                the data to be extracted.
' Return value:
' Returns one of the following values:
' - S_OK         : Data successfully extracted.
' - E_INVALIDARG : The VARIANT was not of type VT_ARRRAY OR VT_UI1.
' - E_FAIL       : The VARIANT buffer value had fewer than cb bytes.
' Note: Requires Windows XP SP2 or superior.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantToBuffer (BYVAL pPropVarIn AS PROPVARIANT PTR, BYVAL pv AS LPVOID, BYVAL cb AS ULONG) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pPropVariantToBuffer AS FUNCTION (BYVAL pPropVarIn AS PROPVARIANT PTR, BYVAL pv AS LPVOID, BYVAL cb AS ULONG) AS HRESULT
   pPropVariantToBuffer = DyLibSymbol(pLib, "PropVariantToBuffer")
   IF pPropVariantToBuffer = NULL THEN FUNCTION = E_FAIL
   FUNCTION = pPropVariantToBuffer(pPropVarIn, pv, cb)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Initializes a PROPVARIANT structure based on a string resource imbedded in an executable file.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantFromResource (BYVAL hinst AS HINSTANCE, BYVAL id AS UINT, BYVAL pPropVar AS PROPVARIANT PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pInitPropVariantFromResource AS FUNCTION (BYVAL hinst AS HINSTANCE, BYVAL id AS UINT, BYVAL pPropVar AS PROPVARIANT PTR) AS HRESULT
   pInitPropVariantFromResource = DyLibSymbol(pLib, "InitPropVariantFromResource")
   IF pInitPropVariantFromResource THEN FUNCTION = pInitPropVariantFromResource(hinst, id, pPropVar)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Initializes a PROPVARIANT structure based on a GUID. The structure is initialized as a The structure is initialized as VT_LPWSTR.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantFromGUIDAsString (BYVAL guid AS IID PTR, BYVAL pPropVar AS PROPVARIANT PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pInitPropVariantFromGUIDAsString AS FUNCTION (BYVAL guid AS IID PTR, BYVAL pPropVar AS PROPVARIANT PTR) AS HRESULT
   pInitPropVariantFromGUIDAsString = DyLibSymbol(pLib, "InitPropVariantFromGUIDAsString")
   IF pInitPropVariantFromGUIDAsString THEN FUNCTION = pInitPropVariantFromGUIDAsString(guid, pPropVar)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Initializes a PROPVARIANT structure based on a Class identifier.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantFromClsid (BYVAL rclsid AS CLSID PTR, BYVAL pPropVar AS PROPVARIANT PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pInitPropVariantFromCLSID AS FUNCTION (BYVAL rclsid AS CLSID PTR, BYVAL pPropVar AS PROPVARIANT PTR) AS HRESULT
   pInitPropVariantFromCLSID = DyLibSymbol(pLib, "InitPropVariantFromCLSID")
   IF pInitPropVariantFromCLSID THEN FUNCTION = pInitPropVariantFromCLSID(rclsid, pPropVar)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Initializes a PROPVARIANT structure with the contents of a FILETIME structure.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantFromFileTime (BYVAL pft AS FILETIME PTR, BYVAL pPropVar AS PROPVARIANT PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pInitPropVariantFromFileTime AS FUNCTION (BYVAL pft AS FILETIME PTR, BYVAL pPropVar AS PROPVARIANT PTR) AS HRESULT
   pInitPropVariantFromFileTime = DyLibSymbol(pLib, "InitPropVariantFromFileTime")
   IF pInitPropVariantFromFileTime THEN FUNCTION = pInitPropVariantFromFileTime(pft, pPropVar)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Initializes a PROPVARIANT structure with a string stored in a STRRET structure.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantFromStrRet (BYVAL pstrret AS STRRET PTR, BYVAL pidl AS PCUITEMID_CHILD, BYVAL pPropVar AS PROPVARIANT PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pInitPropVariantFromStrRet AS FUNCTION (BYVAL pstrret AS STRRET PTR, BYVAL pidl AS PCUITEMID_CHILD, BYVAL pPropVar AS PROPVARIANT PTR) AS HRESULT
   pInitPropVariantFromStrRet = DyLibSymbol(pLib, "InitPropVariantFromStrRet")
   IF pInitPropVariantFromStrRet THEN FUNCTION = pInitPropVariantFromStrRet(pstrret, pidl, pPropVar)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Initializes a PROPVARIANT structure from an array of Boolean values.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantFromBooleanVector (BYVAL prgf AS WINBOOL PTR, BYVAL cElems AS ULONG, BYVAL pPropVar AS PROPVARIANT PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pInitVariantFromBooleanVector AS FUNCTION (BYVAL prgf AS WINBOOL PTR, BYVAL cElems AS ULONG, BYVAL pPropVar AS PROPVARIANT PTR) AS HRESULT
   pInitVariantFromBooleanVector = DyLibSymbol(pLib, "InitPropVariantFromBooleanVector")
   IF pInitVariantFromBooleanVector THEN FUNCTION = pInitVariantFromBooleanVector(prgf, cElems, pPropVar)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Initializes a PROPVARIANT structure with an array of 16-bit integer values.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantFromInt16Vector (BYVAL prgf AS SHORT PTR, BYVAL cElems AS ULONG, BYVAL pPropVar AS PROPVARIANT PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pInitPropVariantFromInt16Vector AS FUNCTION (BYVAL prgf AS SHORT PTR, BYVAL cElems AS ULONG, BYVAL pPropVar AS PROPVARIANT PTR) AS HRESULT
   pInitPropVariantFromInt16Vector = DyLibSymbol(pLib, "InitPropVariantFromInt16Vector")
   IF pInitPropVariantFromInt16Vector THEN FUNCTION = pInitPropVariantFromInt16Vector(prgf, cElems, pPropVar)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Initializes a PROPVARIANT structure with an array of unsigned 16-bit integer values.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantFromUInt16Vector (BYVAL prgf AS USHORT PTR, BYVAL cElems AS ULONG, BYVAL pPropVar AS PROPVARIANT PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pInitPropVariantFromUInt16Vector AS FUNCTION (BYVAL prgf AS SHORT PTR, BYVAL cElems AS ULONG, BYVAL pPropVar AS PROPVARIANT PTR) AS HRESULT
   pInitPropVariantFromUInt16Vector = DyLibSymbol(pLib, "InitPropVariantFromUInt16Vector")
   IF pInitPropVariantFromUInt16Vector THEN FUNCTION = pInitPropVariantFromUInt16Vector(prgf, cElems, pPropVar)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Initializes a PROPVARIANT structure with an array of 32-bit integer values.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantFromInt32Vector (BYVAL prgn AS LONG PTR, BYVAL cElems AS ULONG, BYVAL pPropVar AS PROPVARIANT PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pInitPropVariantFromInt32Vector AS FUNCTION (BYVAL prgn AS LONG PTR, BYVAL cElems AS ULONG, BYVAL pPRopVar AS PROPVARIANT PTR) AS HRESULT
   pInitPropVariantFromInt32Vector = DyLibSymbol(pLib, "InitPropVariantFromInt32Vector")
   IF pInitPropVariantFromInt32Vector THEN FUNCTION = pInitPropVariantFromInt32Vector(prgn, cElems, pPropVar)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Initializes a VARIANT structure with an array of unsigned 32-bit integer values.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantFromUInt32Vector (BYVAL prgn AS ULONG PTR, BYVAL cElems AS ULONG, BYVAL pPropVar AS PROPVARIANT PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pInitPropVariantFromUInt32Vector AS FUNCTION (BYVAL prgn AS ULONG PTR, BYVAL cElems AS ULONG, BYVAL pPropVar AS PROPVARIANT PTR) AS HRESULT
   pInitPropVariantFromUInt32Vector = DyLibSymbol(pLib, "InitPropVariantFromUInt32Vector")
   IF pInitPropVariantFromUInt32Vector THEN FUNCTION = pInitPropVariantFromUInt32Vector(prgn, cElems, pPropVar)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Initializes a PROPVARIANT structure with an array of 64-bit integer values.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantFromInt64Vector (BYVAL prgn AS LONGINT PTR, BYVAL cElems AS ULONG, BYVAL pPropVar AS PROPVARIANT PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pInitPropVariantFromInt64Vector AS FUNCTION (BYVAL prgn AS LONGINT PTR, BYVAL cElems AS ULONG, BYVAL pPropVar AS PROPVARIANT PTR) AS HRESULT
   pInitPropVariantFromInt64Vector = DyLibSymbol(pLib, "InitPropVariantFromInt64Vector")
   IF pInitPropVariantFromInt64Vector THEN FUNCTION = pInitPropVariantFromInt64Vector(prgn, cElems, pPropVar)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Initializes a PROPVARIANT structure with an array of unsigned 64-bit integer values.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantFromUInt64Vector (BYVAL prgn AS ULONGINT PTR, BYVAL cElems AS ULONG, BYVAL pPropVar AS PROPVARIANT PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pInitPropVariantFromUInt64Vector AS FUNCTION (BYVAL prgn AS ULONGINT PTR, BYVAL cElems AS ULONG, BYVAL pPropVar AS PROPVARIANT PTR) AS HRESULT
   pInitPropVariantFromUInt64Vector = DyLibSymbol(pLib, "InitPropVariantFromUInt64Vector")
   IF pInitPropVariantFromUInt64Vector THEN FUNCTION = pInitPropVariantFromUInt64Vector(prgn, cElems, pPropVar)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Initializes a PROPVARIANT structure with an array of values of type DOUBLE.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantFromDoubleVector (BYVAL prgn AS DOUBLE PTR, BYVAL cElems AS ULONG, BYVAL pPropVar AS PROPVARIANT PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pInitPropVariantFromDoubleVector AS FUNCTION (BYVAL prgn AS DOUBLE PTR, BYVAL cElems AS ULONG, BYVAL pPropVar AS PROPVARIANT PTR) AS HRESULT
   pInitPropVariantFromDoubleVector = DyLibSymbol(pLib, "InitPropVariantFromDoubleVector")
   IF pInitPropVariantFromDoubleVector THEN FUNCTION = pInitPropVariantFromDoubleVector(prgn, cElems, pPropVar)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Initializes a PROPVARIANT structure with an array of FILETIME structures.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantFromFileTimeVector (BYVAL prgft AS FILETIME PTR, BYVAL cElems AS ULONG, BYVAL pPropVar AS PROPVARIANT PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pInitPropVariantFromFileTimeVector AS FUNCTION (BYVAL prgft AS FILETIME PTR, BYVAL cElems AS ULONG, BYVAL pPropVar AS PROPVARIANT PTR) AS HRESULT
   pInitPropVariantFromFileTimeVector = DyLibSymbol(pLib, "InitPropVariantFromFileTimeVector")
   IF pInitPropVariantFromFileTimeVector THEN FUNCTION = pInitPropVariantFromFileTimeVector(prgft, cElems, pPropVar)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Initializes a PROPVARIANT structure with an array of strings.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantFromStringVector (BYVAL prgsz AS PCWSTR, BYVAL cElems AS ULONG, BYVAL pPropVar AS PROPVARIANT PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pInitPropVariantFromStringVector AS FUNCTION (BYVAL prgsz AS PCWSTR, BYVAL cElems AS ULONG, BYVAL pPropVar AS PROPVARIANT PTR) AS HRESULT
   pInitPropVariantFromStringVector = DyLibSymbol(pLib, "InitPropVariantFromStringVector")
   IF pInitPropVariantFromStringVector THEN FUNCTION = pInitPropVariantFromStringVector(prgsz, cElems, pPropVar)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Initializes a PROPVARIANT structure from a specified string. The string is parsed as a
' semi-colon delimited list (for example: "A;B;C").
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantFromStringAsVector (BYVAL pwsz AS WSTRING PTR, BYVAL pPropVar AS PROPVARIANT PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pInitPropVariantFromStringAsVector AS FUNCTION (BYVAL prgsz AS PCWSTR, BYVAL pPropVar AS PROPVARIANT PTR) AS HRESULT
   pInitPropVariantFromStringAsVector = DyLibSymbol(pLib, "InitPropVariantFromStringAsVector")
   IF pInitPropVariantFromStringAsVector THEN FUNCTION = pInitPropVariantFromStringAsVector(pwsz, pPropVar)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Initializes a PROPVARIANT structure based on a specified PROPVARIANT vector element.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantFromPropVariantVectorElem (BYVAL pPropVarIn AS PROPVARIANT PTR, BYVAL iElem AS ULONG, BYVAL pPropVar AS PROPVARIANT PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pInitPropVariantFromPropVariantVectorElem AS FUNCTION (BYVAL pPropVarIn AS PROPVARIANT PTR, BYVAL iElem AS ULONG, BYVAL pPropVar AS PROPVARIANT PTR) AS HRESULT
   pInitPropVariantFromPropVariantVectorElem = DyLibSymbol(pLib, "InitPropVariantFromPropVariantVectorElem")
   IF pInitPropVariantFromPropVariantVectorElem THEN FUNCTION = pInitPropVariantFromPropVariantVectorElem(pPropVarIn, iElem, pPropVar)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the element count of a PROPVARIANT structure.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantGetElementCount (BYVAL pPropVarIn AS PROPVARIANT PTR) AS ULONG
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pPropVariantGetElementCount AS FUNCTION (BYVAL pPropVarIn AS PROPVARIANT PTR) AS ULONG
   pPropVariantGetElementCount = DyLibSymbol(pLib, "PropVariantGetElementCount")
   IF pPropVariantGetElementCount = NULL THEN EXIT FUNCTION
   FUNCTION = pPropVariantGetElementCount(pPropVarIn)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts a single Boolean element from a PROPVARIANT structure.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantGetBooleanElem (BYVAL pPropVar AS PROPVARIANT PTR, BYVAL iElem AS ULONG, BYVAL pfVal AS WINBOOL PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pPropVariantGetBooleanElem AS FUNCTION (BYVAL pPropVar AS PROPVARIANT PTR, BYVAL iElem AS ULONG, BYVAL pfVal AS WINBOOL PTR) AS HRESULT
   pPropVariantGetBooleanElem = DyLibSymbol(pLib, "PropVariantGetBooleanElem")
   IF pPropVariantGetBooleanElem THEN FUNCTION = pPropVariantGetBooleanElem(pPropVar, iElem, pfVal)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts a single Int16 element from a PROPVARIANT structure.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantGetInt16Elem (BYVAL pPropVar AS PROPVARIANT PTR, BYVAL iElem AS ULONG, BYVAL pfVal AS SHORT PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pPropVariantGetInt16Elem AS FUNCTION (BYVAL pPropVar AS PROPVARIANT PTR, BYVAL iElem AS ULONG, BYVAL pfVal AS SHORT PTR) AS HRESULT
   pPropVariantGetInt16Elem = DyLibSymbol(pLib, "PropVariantGetInt16Elem")
   IF pPropVariantGetInt16Elem THEN FUNCTION = pPropVariantGetInt16Elem(pPropVar, iElem, pfVal)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts a single unsigned Int16 element from a PROPVARIANT structure.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantGetUInt16Elem (BYVAL pPropVar AS PROPVARIANT PTR, BYVAL iElem AS ULONG, BYVAL pnVal AS USHORT PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pPropVariantGetUInt16Elem AS FUNCTION (BYVAL pPropVar AS PROPVARIANT PTR, BYVAL iElem AS ULONG, BYVAL pnVal AS USHORT PTR) AS HRESULT
   pPropVariantGetUInt16Elem = DyLibSymbol(pLib, "PropVariantGetUInt16Elem")
   IF pPropVariantGetUInt16Elem THEN FUNCTION = pPropVariantGetUInt16Elem(pPropVar, iElem, pnVal)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts a single Int32 element from a PROPVARIANT structure.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantGetInt32Elem (BYVAL pPropVar AS PROPVARIANT PTR, BYVAL iElem AS ULONG, BYVAL pnVal AS LONG PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pPropVariantGetInt32Elem AS FUNCTION (BYVAL pPropVar AS PROPVARIANT PTR, BYVAL iElem AS ULONG, BYVAL pnVal AS LONG PTR) AS HRESULT
   pPropVariantGetInt32Elem = DyLibSymbol(pLib, "PropVariantGetInt32Elem")
   IF pPropVariantGetInt32Elem THEN FUNCTION = pPropVariantGetInt32Elem(pPropVar, iElem, pnVal)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts a single unsigned Int32 element from a PROPVARIANT structure.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantGetUInt32Elem (BYVAL pPropVar AS PROPVARIANT PTR, BYVAL iElem AS ULONG, BYVAL pnVal AS ULONG PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pPropVariantGetUInt32Elem AS FUNCTION (BYVAL pPropVar AS PROPVARIANT PTR, BYVAL iElem AS ULONG, BYVAL pnVal AS ULONG PTR) AS HRESULT
   pPropVariantGetUInt32Elem = DyLibSymbol(pLib, "PropVariantGetUInt32Elem")
   IF pPropVariantGetUInt32Elem THEN FUNCTION = pPropVariantGetUInt32Elem(pPropVar, iElem, pnVal)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts a single Int64 element from a PROPVARIANT structure.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantGetInt64Elem (BYVAL pPropVar AS PROPVARIANT PTR, BYVAL iElem AS ULONG, BYVAL pnVal AS LONGINT PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pPropVariantGetInt64Elem AS FUNCTION (BYVAL pPropVar AS PROPVARIANT PTR, BYVAL iElem AS ULONG, BYVAL pnVal AS LONGINT PTR) AS HRESULT
   pPropVariantGetInt64Elem = DyLibSymbol(pLib, "PropVariantGetInt64Elem")
   IF pPropVariantGetInt64Elem THEN FUNCTION = pPropVariantGetInt64Elem(pPropVar, iElem, pnVal)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts a single unsigned Int64 element from a PROPVARIANT structure.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantGetUInt64Elem (BYVAL pPropVar AS PROPVARIANT PTR, BYVAL iElem AS ULONG, BYVAL pnVal AS ULONGINT PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pPropVariantGetUInt64Elem AS FUNCTION (BYVAL pPropVar AS PROPVARIANT PTR, BYVAL iElem AS ULONG, BYVAL pnVal AS ULONGINT PTR) AS HRESULT
   pPropVariantGetUInt64Elem = DyLibSymbol(pLib, "PropVariantGetUInt64Elem")
   IF pPropVariantGetUInt64Elem THEN FUNCTION = pPropVariantGetUInt64Elem(pPropVar, iElem, pnVal)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts a double element from a PROPVARIANT structure.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantGetDoubleElem (BYVAL pPropVar AS PROPVARIANT PTR, BYVAL iElem AS ULONG, BYVAL pnVal AS DOUBLE PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pPropVariantGetDoubleElem AS FUNCTION (BYVAL pPropVar AS PROPVARIANT PTR, BYVAL iElem AS ULONG, BYVAL pnVal AS DOUBLE PTR) AS HRESULT
   pPropVariantGetDoubleElem = DyLibSymbol(pLib, "PropVariantGetDoubleElem")
   IF pPropVariantGetDoubleElem THEN FUNCTION = pPropVariantGetDoubleElem(pPropVar, iElem, pnVal)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts a single wide string element from a PROPVARIANT structure.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantGetStringElem (BYVAL pPropVar AS PROPVARIANT PTR, BYVAL iElem AS ULONG, BYVAL ppszVal AS PWSTR PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pPropVariantGetStringElem AS FUNCTION (BYVAL pPropVar AS PROPVARIANT PTR, BYVAL iElem AS ULONG, BYVAL ppszVal AS PWSTR PTR) AS HRESULT
   pPropVariantGetStringElem = DyLibSymbol(pLib, "PropVariantGetStringElem")
   IF pPropVariantGetStringElem THEN FUNCTION = pPropVariantGetStringElem(pPropVar, iElem, ppszVal)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts a FILETIME element from a PROPVARIANT structure.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantGetFileTimeElem (BYVAL pPropVar AS PROPVARIANT PTR, BYVAL iElem AS ULONG, BYVAL pftVal AS FILETIME PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pPropVariantGetFileTimeElem AS FUNCTION (BYVAL pPropVar AS PROPVARIANT PTR, BYVAL iElem AS ULONG, BYVAL pftVal AS FILETIME PTR) AS HRESULT
   pPropVariantGetFileTimeElem = DyLibSymbol(pLib, "PropVariantGetFileTimeElem")
   IF pPropVariantGetFileTimeElem THEN FUNCTION = pPropVariantGetFileTimeElem(pPropVar, iElem, pftVal)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts a Boolean property value of a PROPVARIANT structure.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantToBoolean (BYVAL pPropVarIn AS PROPVARIANT PTR, BYVAL pfRet AS WINBOOL PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pPropVariantToBoolean AS FUNCTION (BYVAL pPropVarIn AS PROPVARIANT PTR, BYVAL pfRet AS WINBOOL PTR) AS HRESULT
   pPropVariantToBoolean = DyLibSymbol(pLib, "PropVariantToBoolean")
   IF pPropVariantToBoolean THEN FUNCTION = pPropVariantToBoolean(pPropVarIn, pfRet)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts an Int16 property value of a PROPVARIANT structure.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantToInt16 (BYVAL pPropVarIn AS PROPVARIANT PTR, BYVAL piRet AS SHORT PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pPropVariantToInt16 AS FUNCTION (BYVAL pPropVarIn AS PROPVARIANT PTR, BYVAL piRet AS SHORT PTR) AS HRESULT
   pPropVariantToInt16 = DyLibSymbol(pLib, "PropVariantToInt16")
   IF pPropVariantToInt16 THEN FUNCTION = pPropVariantToInt16(pPropVarIn, piRet)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts an UInt16 property value of a PROPVARIANT structure.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantToUInt16 (BYVAL pPropVarIn AS PROPVARIANT PTR, BYVAL puiRet AS USHORT PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pPropVariantToUInt16 AS FUNCTION (BYVAL pPropVarIn AS PROPVARIANT PTR, BYVAL puiRet AS USHORT PTR) AS HRESULT
   pPropVariantToUInt16 = DyLibSymbol(pLib, "PropVariantToUInt16")
   IF pPropVariantToUInt16 THEN FUNCTION = pPropVariantToUInt16(pPropVarIn, puiRet)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts an Int32 property value of a PROPVARIANT structure.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantToInt32 (BYVAL pPropVarIn AS PROPVARIANT PTR, BYVAL plRet AS LONG PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pPropVariantToInt32 AS FUNCTION (BYVAL pPropVarIn AS PROPVARIANT PTR, BYVAL plRet AS LONG PTR) AS HRESULT
   pPropVariantToInt32 = DyLibSymbol(pLib, "PropVariantToInt32")
   IF pPropVariantToInt32 THEN FUNCTION = pPropVariantToInt32(pPropVarIn, plRet)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts an UInt32 property value of a PROPVARIANT structure.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantToUInt32 (BYVAL pPropVarIn AS PROPVARIANT PTR, BYVAL pulRet AS ULONG PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pPropVariantToUInt32 AS FUNCTION (BYVAL pPropVarIn AS PROPVARIANT PTR, BYVAL pulRet AS ULONG PTR) AS HRESULT
   pPropVariantToUInt32 = DyLibSymbol(pLib, "PropVariantToUInt32")
   IF pPropVariantToUInt32 THEN FUNCTION = pPropVariantToUInt32(pPropVarIn, pulRet)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts an Int64 property value of a PROPVARIANT structure.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantToInt64 (BYVAL pPropVarIn AS PROPVARIANT PTR, BYVAL pllRet AS LONGLONG PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pPropVariantToInt64 AS FUNCTION (BYVAL pPropVarIn AS PROPVARIANT PTR, BYVAL pllRet AS LONGLONG PTR) AS HRESULT
   pPropVariantToInt64 = DyLibSymbol(pLib, "PropVariantToInt64")
   IF pPropVariantToInt64 THEN FUNCTION = pPropVariantToInt64(pPropVarIn, pllRet)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts an UInt64 property value of a PROPVARIANT structure.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantToUInt64 (BYVAL pPropVarIn AS PROPVARIANT PTR, BYVAL pullRet AS ULONGLONG PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pPropVariantToUInt64 AS FUNCTION (BYVAL pPropVarIn AS PROPVARIANT PTR, BYVAL pullRet AS ULONGLONG PTR) AS HRESULT
   pPropVariantToUInt64 = DyLibSymbol(pLib, "PropVariantToUInt64")
   IF pPropVariantToUInt64 THEN FUNCTION = pPropVariantToUInt64(pPropVarIn, pullRet)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts a Double property value of a PROPVARIANT structure.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantToDouble (BYVAL pPropVarIn AS PROPVARIANT PTR, BYVAL pdblRet AS DOUBLE PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pPropVariantToDouble AS FUNCTION (BYVAL pPropVarIn AS PROPVARIANT PTR, BYVAL pdblRet AS DOUBLE PTR) AS HRESULT
   pPropVariantToDouble = DyLibSymbol(pLib, "PropVariantToDouble")
   IF pPropVariantToDouble THEN FUNCTION = pPropVariantToDouble(pPropVarIn, pdblRet)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts a string value from a PROPVARIANT structure.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantToString (BYVAL pPropVarIn AS PROPVARIANT PTR, BYVAL psz AS ZSTRING PTR, BYVAL cch AS UINT) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pPropVariantToString AS FUNCTION (BYVAL pPropVarIn AS PROPVARIANT PTR, BYVAL psz AS ZSTRING PTR, BYVAL cch AS UINT) AS HRESULT
   pPropVariantToString = DyLibSymbol(pLib, "PropVariantToString")
   IF pPropVariantToString THEN FUNCTION = pPropVariantToString(pPropVarIn, psz, cch)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts the BSTR property value from a PROPVARIANT structure.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantToBSTR (BYVAL pPropVarIn AS PROPVARIANT PTR, BYVAL pbstrOut AS BSTR PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pPropVariantToBSTR AS FUNCTION (BYVAL pPropVarIn AS PROPVARIANT PTR, BYVAL pbstrOut AS BSTR PTR) AS HRESULT
   pPropVariantToBSTR = DyLibSymbol(pLib, "PropVariantToBSTR")
   IF pPropVariantToBSTR THEN FUNCTION = pPropVariantToBSTR(pPropVarIn, pbstrOut)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts a GUID property value of a PROPVARIANT structure.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantToGUID (BYVAL pPropVarIn AS PROPVARIANT PTR, BYVAL pguid AS GUID PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pPropVariantToGUID AS FUNCTION (BYVAL pPropVarIn AS PROPVARIANT PTR, BYVAL pguid AS GUID PTR) AS HRESULT
   pPropVariantToGUID = DyLibSymbol(pLib, "PropVariantToGUID")
   IF pPropVariantToGUID THEN FUNCTION = pPropVariantToGUID(pPropVarIn, pguid)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts a FILETIME structure from a variant structure.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantToFileTime (BYVAL pPropVarIn AS PROPVARIANT PTR, BYVAL stfOut AS AFX_PSTIME_FLAGS, BYVAL pftOut AS FILETIME PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pPropVariantToFileTime AS FUNCTION (BYVAL pPropVarIn AS PROPVARIANT PTR, BYVAL stfOut AS AFX_PSTIME_FLAGS, BYVAL pftOut AS FILETIME PTR) AS HRESULT
   pPropVariantToFileTime = DyLibSymbol(pLib, "PropVariantToFileTime")
   IF pPropVariantToFileTime THEN FUNCTION = pPropVariantToFileTime(pPropVarIn, stfOut, pftOut)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' If the source PROPVARIANT is a VT_BSTR, extracts string and places it into a STRRET structure.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantToStrRet (BYVAL pPropVarIn AS PROPVARIANT PTR, BYVAL pstrret AS STRRET PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pPropVariantToStrRet AS FUNCTION (BYVAL pPropVarIn AS PROPVARIANT PTR, BYVAL pstrret AS STRRET PTR) AS HRESULT
   pPropVariantToStrRet = DyLibSymbol(pLib, "PropVariantToStrRet")
   IF pPropVariantToStrRet THEN FUNCTION = pPropVariantToStrRet(pPropVarIn, pstrret)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts an array of Boolean values from a PROPVARIANT structure.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantToBooleanVector (BYVAL pPropVar AS PROPVARIANT PTR, BYVAL prgf AS WINBOOL PTR, BYVAL crgn AS ULONG, BYVAL pcElem AS ULONG PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pPropVariantToBooleanVector AS FUNCTION (BYVAL pPropVar AS PROPVARIANT PTR, BYVAL prgf AS WINBOOL PTR, BYVAL crgn AS ULONG, BYVAL pcElem AS ULONG PTR) AS HRESULT
   pPropVariantToBooleanVector = DyLibSymbol(pLib, "PropVariantToBooleanVector")
   IF pPropVariantToBooleanVector THEN FUNCTION = pPropVariantToBooleanVector(pPropVar, prgf, crgn, pcElem)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Allocates an array of BOOL values then extracts data from a PROPVARIANT structure into that array.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantToBooleanVectorAlloc (BYVAL pPropVar AS PROPVARIANT PTR, BYVAL pprgf AS WINBOOL PTR PTR, BYVAL pcElem AS ULONG PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pPropVariantToBooleanVectorAlloc AS FUNCTION (BYVAL pPropVar AS PROPVARIANT PTR, BYVAL pprgf AS WINBOOL PTR PTR, BYVAL pcElem AS ULONG PTR) AS HRESULT
   pPropVariantToBooleanVectorAlloc = DyLibSymbol(pLib, "PropVariantToBooleanVectorAlloc")
   IF pPropVariantToBooleanVectorAlloc THEN FUNCTION = pPropVariantToBooleanVectorAlloc(pPropVar, pprgf, pcElem)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts data from a vector structure into an Int16 array.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantToInt16Vector (BYVAL pPropVar AS PROPVARIANT PTR, BYVAL prgn AS SHORT PTR, BYVAL crgn AS ULONG, BYVAL pcElem AS ULONG PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pPropVariantToInt16Vector AS FUNCTION (BYVAL pPropVar AS PROPVARIANT PTR, BYVAL prgn AS SHORT PTR, BYVAL crgn AS ULONG, BYVAL pcElem AS ULONG PTR) AS HRESULT
   pPropVariantToInt16Vector = DyLibSymbol(pLib, "PropVariantToInt16Vector")
   IF pPropVariantToInt16Vector THEN FUNCTION = pPropVariantToInt16Vector(pPropVar, prgn, crgn, pcElem)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Extracts data from a vector structure into a newly-allocated Int16 array.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantToInt16VectorAlloc (BYVAL pPropVar AS PROPVARIANT PTR, BYVAL pprgn AS SHORT PTR PTR, BYVAL pcElem AS ULONG PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pPropVariantToInt16VectorAlloc AS FUNCTION (BYVAL pPropVar AS PROPVARIANT PTR, BYVAL pprgn AS SHORT PTR PTR, BYVAL pcElem AS ULONG PTR) AS HRESULT
   pPropVariantToInt16VectorAlloc = DyLibSymbol(pLib, "PropVariantToInt16VectorAlloc")
   IF pPropVariantToInt16VectorAlloc THEN FUNCTION = pPropVariantToInt16VectorAlloc(pPropVar, pprgn, pcElem)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts data from a vector structure into an unsigned Int16 array.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantToUInt16Vector (BYVAL pPropVar AS PROPVARIANT PTR, BYVAL prgn AS USHORT PTR, BYVAL crgn AS ULONG, BYVAL pcElem AS ULONG PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pPropVariantToUInt16Vector AS FUNCTION (BYVAL pPropVar AS PROPVARIANT PTR, BYVAL prgn AS USHORT PTR, BYVAL crgn AS ULONG, BYVAL pcElem AS ULONG PTR) AS HRESULT
   pPropVariantToUInt16Vector = DyLibSymbol(pLib, "PropVariantToUInt16Vector")
   IF pPropVariantToUInt16Vector THEN FUNCTION = pPropVariantToUInt16Vector(pPropVar, prgn, crgn, pcElem)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Extracts data from a vector structure into a newly-allocated unsigned Int16 array.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantToUInt16VectorAlloc (BYVAL pPropVar AS PROPVARIANT PTR, BYVAL pprgn AS USHORT PTR PTR, BYVAL pcElem AS ULONG PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pPropVariantToUInt16VectorAlloc AS FUNCTION (BYVAL pPropVar AS PROPVARIANT PTR, BYVAL pprgn AS USHORT PTR PTR, BYVAL pcElem AS ULONG PTR) AS HRESULT
   pPropVariantToUInt16VectorAlloc = DyLibSymbol(pLib, "PropVariantToUInt16VectorAlloc")
   IF pPropVariantToUInt16VectorAlloc THEN FUNCTION = pPropVariantToUInt16VectorAlloc(pPropVar, pprgn, pcElem)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts data from a vector structure into an unsigned Int16 array.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantToInt32Vector (BYVAL pPropVar AS PROPVARIANT PTR, BYVAL prgn AS LONG PTR, BYVAL crgn AS ULONG, BYVAL pcElem AS ULONG PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pPropVariantToInt32Vector AS FUNCTION (BYVAL pPropVar AS PROPVARIANT PTR, BYVAL prgn AS LONG PTR, BYVAL crgn AS ULONG, BYVAL pcElem AS ULONG PTR) AS HRESULT
   pPropVariantToInt32Vector = DyLibSymbol(pLib, "PropVariantToInt32Vector")
   IF pPropVariantToInt32Vector THEN FUNCTION = pPropVariantToInt32Vector(pPropVar, prgn, crgn, pcElem)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Extracts data from a vector structure into a newly-allocated unsigned Int16 array.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantToInt32VectorAlloc (BYVAL pPropVar AS PROPVARIANT PTR, BYVAL pprgn AS LONG PTR PTR, BYVAL pcElem AS ULONG PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pPropVariantToInt32VectorAlloc AS FUNCTION (BYVAL pPropVar AS PROPVARIANT PTR, BYVAL pprgn AS LONG PTR PTR, BYVAL pcElem AS ULONG PTR) AS HRESULT
   pPropVariantToInt32VectorAlloc = DyLibSymbol(pLib, "PropVariantToInt32VectorAlloc")
   IF pPropVariantToInt32VectorAlloc THEN FUNCTION = pPropVariantToInt32VectorAlloc(pPropVar, pprgn, pcElem)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts data from a vector structure into an unsigned Int16 array.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantToUInt32Vector (BYVAL pPropVar AS PROPVARIANT PTR, BYVAL prgn AS ULONG PTR, BYVAL crgn AS ULONG, BYVAL pcElem AS ULONG PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pPropVariantToUInt32Vector AS FUNCTION (BYVAL pPropVar AS PROPVARIANT PTR, BYVAL prgn AS ULONG PTR, BYVAL crgn AS ULONG, BYVAL pcElem AS ULONG PTR) AS HRESULT
   pPropVariantToUInt32Vector = DyLibSymbol(pLib, "PropVariantToUInt32Vector")
   IF pPropVariantToUInt32Vector THEN FUNCTION = pPropVariantToUInt32Vector(pPropVar, prgn, crgn, pcElem)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Extracts data from a vector structure into a newly-allocated unsigned Int32 array.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantToUInt32VectorAlloc (BYVAL pPropVar AS PROPVARIANT PTR, BYVAL pprgn AS ULONG PTR PTR, BYVAL pcElem AS ULONG PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pPropVariantToUInt32VectorAlloc AS FUNCTION (BYVAL pPropVar AS PROPVARIANT PTR, BYVAL pprgn AS ULONG PTR PTR, BYVAL pcElem AS ULONG PTR) AS HRESULT
   pPropVariantToUInt32VectorAlloc = DyLibSymbol(pLib, "PropVariantToUInt32VectorAlloc")
   IF pPropVariantToUInt32VectorAlloc THEN FUNCTION = pPropVariantToUInt32VectorAlloc(pPropVar, pprgn, pcElem)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts data from a vector structure into an unsigned Int64 array.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantToInt64Vector (BYVAL pPropVar AS PROPVARIANT PTR, BYVAL prgn AS LONGINT PTR, BYVAL crgn AS ULONG, BYVAL pcElem AS ULONG PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pPropVariantToInt64Vector AS FUNCTION (BYVAL pPropVar AS PROPVARIANT PTR, BYVAL prgn AS LONGINT PTR, BYVAL crgn AS ULONG, BYVAL pcElem AS ULONG PTR) AS HRESULT
   pPropVariantToInt64Vector = DyLibSymbol(pLib, "PropVariantToInt64Vector")
   IF pPropVariantToInt64Vector THEN FUNCTION = pPropVariantToInt64Vector(pPropVar, prgn, crgn, pcElem)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Extracts data from a vector structure into a newly-allocated Int64 array.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantToInt64VectorAlloc (BYVAL pPropVar AS PROPVARIANT PTR, BYVAL pprgn AS LONGINT PTR PTR, BYVAL pcElem AS ULONG PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pPropVariantToInt64VectorAlloc AS FUNCTION (BYVAL pPropVar AS PROPVARIANT PTR, BYVAL pprgn AS LONGINT PTR PTR, BYVAL pcElem AS ULONG PTR) AS HRESULT
   pPropVariantToInt64VectorAlloc = DyLibSymbol(pLib, "PropVariantToInt64VectorAlloc")
   IF pPropVariantToInt64VectorAlloc THEN FUNCTION = pPropVariantToInt64VectorAlloc(pPropVar, pprgn, pcElem)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts data from a vector structure into an unsigned Int64 array.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantToUInt64Vector (BYVAL pPropVar AS PROPVARIANT PTR, BYVAL prgn AS ULONGINT PTR, BYVAL crgn AS ULONG, BYVAL pcElem AS ULONG PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pPropVariantToUInt64Vector AS FUNCTION (BYVAL pPropVar AS PROPVARIANT PTR, BYVAL prgn AS ULONGINT PTR, BYVAL crgn AS ULONG, BYVAL pcElem AS ULONG PTR) AS HRESULT
   pPropVariantToUInt64Vector = DyLibSymbol(pLib, "PropVariantToUInt64Vector")
   IF pPropVariantToUInt64Vector THEN FUNCTION = pPropVariantToUInt64Vector(pPropVar, prgn, crgn, pcElem)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Extracts data from a vector structure into a newly-allocated Int64 array.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantToUInt64VectorAlloc (BYVAL pPropVar AS PROPVARIANT PTR, BYVAL pprgn AS ULONGINT PTR PTR, BYVAL pcElem AS ULONG PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pPropVariantToUInt64VectorAlloc AS FUNCTION (BYVAL pPropVar AS PROPVARIANT PTR, BYVAL pprgn AS ULONGINT PTR PTR, BYVAL pcElem AS ULONG PTR) AS HRESULT
   pPropVariantToUInt64VectorAlloc = DyLibSymbol(pLib, "PropVariantToUInt64VectorAlloc")
   IF pPropVariantToUInt64VectorAlloc THEN FUNCTION = pPropVariantToUInt64VectorAlloc(pPropVar, pprgn, pcElem)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts an array of DOUBLE values from a PROPVARIANT structure.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantToDoubleVector (BYVAL pPropVar AS PROPVARIANT PTR, BYVAL prgn AS DOUBLE PTR, BYVAL crgn AS ULONG, BYVAL pcElem AS ULONG PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pPropVariantToDoubleVector AS FUNCTION (BYVAL pPropVar AS PROPVARIANT PTR, BYVAL prgn AS DOUBLE PTR, BYVAL crgn AS ULONG, BYVAL pcElem AS ULONG PTR) AS HRESULT
   pPropVariantToDoubleVector = DyLibSymbol(pLib, "PropVariantToDoubleVector")
   IF pPropVariantToDoubleVector THEN FUNCTION = pPropVariantToDoubleVector(pPropVar, prgn, crgn, pcElem)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Allocates an array of DOUBLE values then extracts data from a PROPVARIANT structure into that array.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantToDoubleVectorAlloc (BYVAL pPropVar AS PROPVARIANT PTR, BYVAL pprgn AS DOUBLE PTR PTR, BYVAL pcElem AS ULONG PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pPropVariantToDoubleVectorAlloc AS FUNCTION (BYVAL pPropVar AS PROPVARIANT PTR, BYVAL pprgn AS DOUBLE PTR PTR, BYVAL pcElem AS ULONG PTR) AS HRESULT
   pPropVariantToDoubleVectorAlloc = DyLibSymbol(pLib, "PropVariantToDoubleVectorAlloc")
   IF pPropVariantToDoubleVectorAlloc THEN FUNCTION = pPropVariantToDoubleVectorAlloc(pPropVar, pprgn, pcElem)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts data from a vector structure into a string array.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantToStringVector (BYVAL pPropVar AS PROPVARIANT PTR, BYVAL prgsz AS PWSTR, BYVAL crgsz AS ULONG, BYVAL pcElem AS ULONG PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pPropVariantToStringVector AS FUNCTION (BYVAL pPropVar AS PROPVARIANT PTR, BYVAL prgsz AS PWSTR, BYVAL crgsz AS ULONG, BYVAL pcElem AS ULONG PTR) AS HRESULT
   pPropVariantToStringVector = DyLibSymbol(pLib, "PropVariantToStringVector")
   IF pPropVariantToStringVector THEN FUNCTION = pPropVariantToStringVector(pPropVar, prgsz, crgsz, pcElem)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Extracts data from a vector structure into a newly-allocated string array.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantToStringVectorAlloc (BYVAL pPropVar AS PROPVARIANT PTR, BYVAL pprgsz AS PWSTR PTR, BYVAL pcElem AS ULONG PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pPropVariantToStringVectorAlloc AS FUNCTION (BYVAL pPropVar AS PROPVARIANT PTR, BYVAL pprgsz AS PWSTR PTR, BYVAL pcElem AS ULONG PTR) AS HRESULT
   pPropVariantToStringVectorAlloc = DyLibSymbol(pLib, "PropVariantToStringVectorAlloc")
   IF pPropVariantToStringVectorAlloc THEN FUNCTION = pPropVariantToStringVectorAlloc(pPropVar, pprgsz, pcElem)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts data from a vector structure into a FILETIME array.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantToFileTimeVector (BYVAL pPropVar AS PROPVARIANT PTR, BYVAL prgft AS FILETIME PTR, BYVAL crgft AS ULONG, BYVAL pcElem AS ULONG PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pPropVariantToFileTimeVector AS FUNCTION (BYVAL pPropVar AS PROPVARIANT PTR, BYVAL prgft AS FILETIME PTR, BYVAL crgft AS ULONG, BYVAL pcElem AS ULONG PTR) AS HRESULT
   pPropVariantToFileTimeVector = DyLibSymbol(pLib, "PropVariantToFileTimeVector")
   IF pPropVariantToFileTimeVector THEN FUNCTION = pPropVariantToFileTimeVector(pPropVar, prgft, crgft, pcElem)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Extracts data from a vector structure into a newly-allocated FILETIME array.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantToFileTimeVectorAlloc (BYVAL pPropVar AS PROPVARIANT PTR, BYVAL pprgft AS FILETIME PTR PTR, BYVAL pcElem AS ULONG PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pPropVariantToFileTimeVectorAlloc AS FUNCTION (BYVAL pPropVar AS PROPVARIANT PTR, BYVAL pprgft AS FILETIME PTR PTR, BYVAL pcElem AS ULONG PTR) AS HRESULT
   pPropVariantToFileTimeVectorAlloc = DyLibSymbol(pLib, "PropVariantToFileTimeVectorAlloc")
   IF pPropVariantToFileTimeVectorAlloc THEN FUNCTION = pPropVariantToFileTimeVectorAlloc(pPropVar, pprgft, pcElem)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Initializes a vector element in a PROPVARIANT structure with a value stored in another PROPVARIANT.
' This function is used to convert a PROPVARIANT structure that contains a single value
' into a vector value.
' For simple source types, this function initializes the PROPVARIANT as a vector of one element.
' For a source that contains a string, this function initializes the PROPVARIANT with zero or
' more substrings taken from the source string, treating semicolons as delimiters.
' See InitPropVariantFromStringAsVector for more details.
' The following input types are supported:
' VT_I2, VT_UI2, VT_I4, VT_UI4, VT_I8, VT_UI8, VT_R8, VT_BOOL, VT_DATE, VT_FILETIME, VT_BSTR, VT_LPWSTR
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantArrayFromPropVariant (BYVAL propvarSingle AS PROPVARIANT PTR, BYVAL ppropvarVector AS PROPVARIANT PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pInitPropVariantVectorFromPropVariant AS FUNCTION (BYVAL propvarSingle AS PROPVARIANT PTR, BYVAL ppropvarVector AS PROPVARIANT PTR) AS HRESULT
   pInitPropVariantVectorFromPropVariant = DyLibSymbol(pLib, "InitPropVariantVectorFromPropVariant")
   IF pInitPropVariantVectorFromPropVariant THEN FUNCTION = pInitPropVariantVectorFromPropVariant(propvarSingle, ppropvarVector)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts the Boolean property value of a PROPVARIANT structure. If no value exists, then
' the specified default value is returned.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantToBooleanWithDefault (BYVAL pPropVarIn AS PROPVARIANT PTR, BYVAL fDefault AS WINBOOL) AS WINBOOL
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pPropVariantToBooleanWithDefault AS FUNCTION (BYVAL pPropVarIn AS PROPVARIANT PTR, BYVAL fDefault AS WINBOOL) AS WINBOOL
   pPropVariantToBooleanWithDefault = DyLibSymbol(pLib, "PropVariantToBooleanWithDefault")
   IF pPropVariantToBooleanWithDefault THEN FUNCTION = pPropVariantToBooleanWithDefault(pPropVarIn, fDefault)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts the Int16 property value of a PROPVARIANT structure. If no value exists, then
' the specified default value is returned.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantToInt16WithDefault (BYVAL pPropVarIn AS PROPVARIANT PTR, BYVAL iDefault AS SHORT) AS SHORT
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pPropVariantToInt16WithDefault AS FUNCTION (BYVAL pPropVarIn AS PROPVARIANT PTR, BYVAL iDefault AS SHORT) AS SHORT
   pPropVariantToInt16WithDefault = DyLibSymbol(pLib, "PropVariantToInt16WithDefault")
   IF pPropVariantToInt16WithDefault THEN FUNCTION = pPropVariantToInt16WithDefault(pPropVarIn, iDefault)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts the UInt16 property value of a PROPVARIANT structure. If no value exists, then
' the specified default value is returned.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantToUInt16WithDefault (BYVAL pPropVarIn AS PROPVARIANT PTR, BYVAL uiDefault AS USHORT) AS USHORT
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pPropVariantToUInt16WithDefault AS FUNCTION (BYVAL pPropVarIn AS PROPVARIANT PTR, BYVAL uiDefault AS USHORT) AS USHORT
   pPropVariantToUInt16WithDefault = DyLibSymbol(pLib, "PropVariantToUInt16WithDefault")
   IF pPropVariantToUInt16WithDefault THEN FUNCTION = pPropVariantToUInt16WithDefault(pPropVarIn, uiDefault)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts the Int32 property value of a PROPVARIANT structure. If no value exists, then
' the specified default value is returned.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantToInt32WithDefault (BYVAL pPropVarIn AS PROPVARIANT PTR, BYVAL lDefault AS LONG) AS LONG
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pPropVariantToInt32WithDefault AS FUNCTION (BYVAL pPropVarIn AS PROPVARIANT PTR, BYVAL lDefault AS LONG) AS LONG
   pPropVariantToInt32WithDefault = DyLibSymbol(pLib, "PropVariantToInt32WithDefault")
   IF pPropVariantToInt32WithDefault THEN FUNCTION = pPropVariantToInt32WithDefault(pPropVarIn, lDefault)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts the UInt32 property value of a PROPVARIANT structure. If no value exists, then
' the specified default value is returned.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantToUInt32WithDefault (BYVAL pPropVarIn AS PROPVARIANT PTR, BYVAL ulDefault AS ULONG) AS ULONG
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pPropVariantToUInt32WithDefault AS FUNCTION (BYVAL pPropVarIn AS PROPVARIANT PTR, BYVAL ulDefault AS ULONG) AS ULONG
   pPropVariantToUInt32WithDefault = DyLibSymbol(pLib, "PropVariantToUInt32WithDefault")
   IF pPropVariantToUInt32WithDefault THEN FUNCTION = pPropVariantToUInt32WithDefault(pPropVarIn, ulDefault)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts the Int64 property value of a PROPVARIANT structure. If no value exists, then
' the specified default value is returned.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantToInt64WithDefault (BYVAL pPropVarIn AS PROPVARIANT PTR, BYVAL llDefault AS LONGLONG) AS LONGLONG
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pPropVariantToInt64WithDefault AS FUNCTION (BYVAL pPropVarIn AS PROPVARIANT PTR, BYVAL llDefault AS LONGLONG) AS LONGLONG
   pPropVariantToInt64WithDefault = DyLibSymbol(pLib, "PropVariantToInt64WithDefault")
   IF pPropVariantToInt64WithDefault THEN FUNCTION = pPropVariantToInt64WithDefault(pPropVarIn, llDefault)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts the UInt64 property value of a PROPVARIANT structure. If no value exists, then
' the specified default value is returned.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantToUInt64WithDefault (BYVAL pPropVarIn AS PROPVARIANT PTR, BYVAL llDefault AS LONGLONG) AS LONGLONG
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pPropVariantToUInt64WithDefault AS FUNCTION (BYVAL pPropVarIn AS PROPVARIANT PTR, BYVAL llDefault AS LONGLONG) AS LONGLONG
   pPropVariantToUInt64WithDefault = DyLibSymbol(pLib, "PropVariantToUInt64WithDefault")
   IF pPropVariantToUInt64WithDefault THEN FUNCTION = pPropVariantToUInt64WithDefault(pPropVarIn, llDefault)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts the Double property value of a PROPVARIANT structure. If no value exists, then
' the specified default value is returned.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantToDoubleWithDefault (BYVAL pPropVarIn AS PROPVARIANT PTR, BYVAL dblDefault AS DOUBLE) AS DOUBLE
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pPropVariantToDoubleWithDefault AS FUNCTION (BYVAL pPropVarIn AS PROPVARIANT PTR, BYVAL dblDefault AS DOUBLE) AS DOUBLE
   pPropVariantToDoubleWithDefault = DyLibSymbol(pLib, "PropVariantToDoubleWithDefault")
   IF pPropVariantToDoubleWithDefault THEN FUNCTION = pPropVariantToDoubleWithDefault(pPropVarIn, dblDefault)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts the string property value of a PROPVARIANT structure. If no value exists, then the
' specified default value is returned.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantToStringWithDefault (BYVAL pPropVarIn AS PROPVARIANT PTR, BYVAL pszDefault AS LPCWSTR) AS PCWSTR
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pPropVariantToStringWithDefault AS FUNCTION (BYVAL pPropVarIn AS PROPVARIANT PTR, BYVAL pszDefault AS LPCWSTR) AS PCWSTR
   pPropVariantToStringWithDefault = DyLibSymbol(pLib, "PropVariantToStringWithDefault")
   IF pPropVariantToStringWithDefault THEN FUNCTION = pPropVariantToStringWithDefault(pPropVarIn, pszDefault)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts a string property value from a PROPVARIANT structure.
' Parameters:
' - pvarIn   : [in] Reference to a source PROPVARIANT structure.
' - ppwszBuf : [out] Pointer to the extracted property value if one exists; otherwise, empty.
' Return value:
' If this function succeeds, it returns S_OK. Otherwise, it returns an HRESULT error code.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantToStringAlloc (BYVAL pPropVarIn AS PROPVARIANT PTR, BYVAL ppwszBuf AS WSTRING PTR PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pPropVariantToStringAlloc AS FUNCTION (BYVAL pPropVar AS PROPVARIANT PTR, BYVAL ppwszBuf AS WSTRING PTR PTR) AS HRESULT
   pPropVariantToStringAlloc = DyLibSymbol(pLib, "PropVariantToStringAlloc")
   IF pPropVariantToStringAlloc = NULL THEN FUNCTION = E_FAIL
   FUNCTION = pPropVariantToStringAlloc(pPropVarIn, ppwszBuf)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts the contents of a PROPVARIANT and returns them as a CWSTR.
' When pPropVarIn contains an array, each element of the array is appended to the resulting
' string separated with a semicolon and a space.
' For variants that contains an array of bytes, use AfxPropVariantToBuffer.
' Parameters:
' - pPropvarIn = Pointer to the variant.
' - bClear = Clear the contents of the PROPVARIANT (TRUE or FALSE).
' Return value:
' - If the function succeeds, it returns the contents of the variant in string form;
'   if it fails, it returns an empty string and the contents of the PROPVARIANT aren't cleared.
' ========================================================================================
PRIVATE FUNCTION AfxPropVarToStr OVERLOAD (BYVAL pPropVarIn AS PROPVARIANT PTR, BYVAL bClear AS BOOLEAN = FALSE) AS CWSTR

   DIM cws AS CWSTR
   IF pPropVarIn = NULL THEN RETURN ""

   SELECT CASE pPropVarIn->vt

      CASE VT_ERROR : cws = WSTR(pPropVarIn->scode)
      CASE VT_R4 : cws = WSTR(pPropVarIn->fltVal)

      ' // VT_BYREF variants
      CASE VT_LPWSTR OR VT_BYREF : IF pPropVarIn->pwszVal THEN cws = *pPropVarIn->pwszVal
      CASE VT_BSTR OR VT_BYREF : IF pPropVarIn->pbstrVal THEN cws = *pPropVarIn->pbstrVal
      CASE VT_BOOL OR VT_BYREF : IF pPropVarIn->pboolVal THEN cws = WSTR(*pPropVarIn->pboolVal)
      CASE VT_I1 OR VT_BYREF   : IF pPropVarIn->pcVal    THEN cws = WSTR(*pPropVarIn->pcVal)
      CASE VT_UI1 OR VT_BYREF  : IF pPropVarIn->pbVal    THEN cws = WSTR(*pPropVarIn->pbVal)
      CASE VT_I2 OR VT_BYREF   : IF pPropVarIn->piVal    THEN cws = WSTR(*pPropVarIn->piVal)
      CASE VT_UI2 OR VT_BYREF  : IF pPropVarIn->piVal    THEN cws = WSTR(*pPropVarIn->puiVal)
      CASE VT_INT OR VT_BYREF  : IF pPropVarIn->pintVal  THEN cws = WSTR(*pPropVarIn->pintVal)
      CASE VT_UINT OR VT_BYREF : IF pPropVarIn->puintVal THEN cws = WSTR(*pPropVarIn->puintVal)
      CASE VT_I4 OR VT_BYREF   : IF pPropVarIn->plVal    THEN cws = WSTR(*pPropVarIn->plVal)
      CASE VT_UI4 OR VT_BYREF  : IF pPropVarIn->pulVal   THEN cws = WSTR(*pPropVarIn->pulVal)
'      CASE VT_I8 OR VT_BYREF   : IF pPropVarIn->pllVal   THEN cws = WSTR(*pPropVarIn->pllVal
'      CASE VT_UI8 OR VT_BYREF  : IF pPropVarIn->pullVal  THEN cws = WSTR(*pPropVarIn->pullVal)
      CASE VT_R4 OR VT_BYREF   : IF pPropVarIn->pfltVal  THEN cws = WSTR(*pPropVarIn->pfltVal)
      CASE VT_R8 OR VT_BYREF   : IF pPropVarIn->pdblVal  THEN cws = WSTR(*pPropVarIn->pdblVal)

      ' // Other...
      CASE ELSE
         DIM ppwszBuf AS WSTRING PTR
         DIM hr AS HRESULT = AfxPropVariantToStringAlloc(pPropVarIn, @ppwszBuf)
         IF hr <> S_OK OR ppwszBuf = NULL THEN RETURN ""
         cws = *ppwszBuf
         CoTaskMemFree ppwszBuf

   END SELECT

   ' // Clear the passed PropVariant
   IF bClear THEN PropVariantClear(pPropVarIn)
   ' // Return the string
   RETURN cws

END FUNCTION
' ========================================================================================

' ========================================================================================
' Allows to pass the PROPVARIANT by reference, i.e. AfxPropVarToStr(v), instead of using
' AfxPropVarToStr(@v). This also allows to use directly the retult of a function that returns
' a PROPVARIANT or a CPROPVAR, e.g. AfxPropVarToStr(AfxPropVarFromLong(83838), TRUE),
' AfxPropVarToStr(AfxCPropVarFromLong(83838)), while trying to use the first overloaded
' function, e.g. AfxPropVarToStr(@AfxPropVarFromLong(83838), TRUE), gives an apparently
' bogus compiler error: Error 7: Expected ")".
' ========================================================================================
PRIVATE FUNCTION AfxPropVarToStr OVERLOAD (BYREF PropVarIn AS PROPVARIANT, BYVAL bClear AS BOOLEAN = FALSE) AS CWSTR
   RETURN AfxPropVarToStr(@PropVarIn, bClear)
END FUNCTION
' ========================================================================================

' ########################################################################################
' CPropVar - PROPVARIANT class
' ########################################################################################
TYPE CPropVar

Private:
   cws AS CWSTR          ' // To allow to return a WSTRING by reference

Public:
   pvd AS PROPVARIANT         ' // PropVariant data

   ' // Constructors
   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR
   DECLARE CONSTRUCTOR (BYREF cpv AS CPropVar)
   DECLARE CONSTRUCTOR (BYVAL pv AS PROPVARIANT)
   DECLARE CONSTRUCTOR (BYVAL v AS VARIANT PTR)
   DECLARE CONSTRUCTOR (BYREF wsz AS WSTRING)
   DECLARE CONSTRUCTOR (BYREF cws AS CWSTR)
   DECLARE CONSTRUCTOR (BYREF cbs AS CBSTR)
   DECLARE CONSTRUCTOR (BYVAL pvar AS PROPVARIANT PTR)
   DECLARE CONSTRUCTOR (BYVAL cy AS CURRENCY)
   DECLARE CONSTRUCTOR (BYVAL dec AS DECIMAL)
   DECLARE CONSTRUCTOR (BYVAL pdisp AS IDispatch PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   DECLARE CONSTRUCTOR (BYVAL punk AS IUnknown PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   DECLARE CONSTRUCTOR (BYVAL pStream AS IStream PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   DECLARE CONSTRUCTOR (BYVAL pStorage AS IStorage PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   DECLARE CONSTRUCTOR (BYVAL pVersionedStream AS VERSIONEDSTREAM PTR)
   DECLARE CONSTRUCTOR (BYVAL _value AS LONGINT, BYVAL _vType AS WORD = VT_I4)
   DECLARE CONSTRUCTOR (BYVAL _value AS DOUBLE, BYVAL _vType AS WORD = VT_R8)
   DECLARE CONSTRUCTOR (BYVAL _value AS LONGINT, BYREF strType AS STRING)
   DECLARE CONSTRUCTOR (BYVAL _value AS DOUBLE, BYREF strType AS STRING)
   DECLARE CONSTRUCTOR (BYVAL _pPropVar AS ANY PTR, BYVAL _vType AS WORD)
   DECLARE CONSTRUCTOR (BYVAL _pPropVar AS ANY PTR, BYREF strType AS STRING)
   DECLARE CONSTRUCTOR (BYREF _blob AS BLOB)
   DECLARE CONSTRUCTOR (BYREF _bstrBlob AS BSTRBLOB)
   DECLARE CONSTRUCTOR (BYVAL pclipData AS CLIPDATA PTR)
   ' // Casting
'   DECLARE OPERATOR @ () AS PROPVARIANT PTR
   DECLARE FUNCTION vptr () AS PROPVARIANT PTR
   DECLARE FUNCTION sptr () AS PROPVARIANT PTR
   DECLARE FUNCTION wstr () AS CWSTR
   DECLARE FUNCTION bstr () AS CBSTR
   DECLARE OPERATOR CAST () AS PROPVARIANT
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
   DECLARE FUNCTION ToVariant () AS VARIANT
   DECLARE FUNCTION DecToDouble () AS DOUBLE
   DECLARE FUNCTION DecToCy () AS CY
   DECLARE FUNCTION ToVbDate () AS DATE_
   DECLARE FUNCTION ToSystemTime () AS SYSTEMTIME
   DECLARE FUNCTION ToGuid () AS GUID
   DECLARE FUNCTION ToClsid () AS CLSID
   DECLARE FUNCTION ToGuidStr () AS CWSTR
   DECLARE FUNCTION ToGuidWStr () AS CWSTR
   DECLARE FUNCTION ToGuidBStr () AS CBSTR
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
   DECLARE FUNCTION ToFileTimeArray (BYVAL prgft AS FILETIME PTR, BYVAL crgft AS ULONG) AS HRESULT
   DECLARE FUNCTION ToFileTimeArrayAlloc (BYVAL prgft AS FILETIME PTR PTR) AS ULONG
   ' // LET assignments
   DECLARE OPERATOR Let (BYREF cpv AS CPropVar)
   DECLARE OPERATOR Let (BYREF pv AS PROPVARIANT)
   DECLARE OPERATOR Let (BYREF wszStr AS WSTRING)
   DECLARE OPERATOR Let (BYREF cws AS CWSTR)
   DECLARE OPERATOR Let (BYREF cbs AS CBSTR)
   DECLARE OPERATOR Let (BYVAL pPropVar AS PROPVARIANT PTR)
   DECLARE OPERATOR Let (BYVAL cy AS CURRENCY)
   DECLARE OPERATOR Let (BYVAL dec AS DECIMAL)
   DECLARE OPERATOR Let (BYVAL pdisp AS IDispatch PTR)
   DECLARE OPERATOR Let (BYVAL punk AS IUnknown PTR)
   DECLARE OPERATOR Let (BYVAL pStream AS IStream PTR)
   DECLARE OPERATOR Let (BYVAL pStorage AS IStorage PTR)
   DECLARE OPERATOR Let (BYVAL pVersionedStream AS VERSIONEDSTREAM PTR)
   DECLARE OPERATOR Let (BYREF _blob AS BLOB)
   DECLARE OPERATOR Let (BYREF _bstrBlob AS BSTRBLOB)
   DECLARE OPERATOR Let (BYVAL pclipData AS CLIPDATA PTR)
   DECLARE OPERATOR Let (BYREF _cac AS CAC)
   DECLARE OPERATOR Let (BYREF _caub AS CAUB)
   DECLARE OPERATOR Let (BYREF _cai AS CAI)
   DECLARE OPERATOR Let (BYREF _caui AS CAUI)
   DECLARE OPERATOR Let (BYREF _cal AS CAL)
   DECLARE OPERATOR Let (BYREF _caul AS CAUL)
   DECLARE OPERATOR Let (BYREF _cah AS CAH)
   DECLARE OPERATOR Let (BYREF _caflt AS CAFLT)
   DECLARE OPERATOR Let (BYREF _cadbl AS CADBL)
   DECLARE OPERATOR Let (BYREF _cabool AS CABOOL)
   DECLARE OPERATOR Let (BYREF _cascode AS CASCODE)
   DECLARE OPERATOR Let (BYREF _cacy AS CACY)
   DECLARE OPERATOR Let (BYREF _cadate AS CADATE)
   DECLARE OPERATOR Let (BYREF _cafiletime AS CAFILETIME)
   DECLARE OPERATOR Let (BYREF _caclsid AS CACLSID)
   DECLARE OPERATOR Let (BYREF _caclipdata AS CACLIPDATA)
   DECLARE OPERATOR Let (BYREF _cabstr AS CABSTR)
   DECLARE OPERATOR Let (BYREF _cabstrblob AS CABSTRBLOB)
   DECLARE OPERATOR Let (BYREF _calpstr AS CALPSTR)
   DECLARE OPERATOR Let (BYREF _calpwstr AS CALPWSTR)
   DECLARE OPERATOR Let (BYREF _capropvar AS CAPROPVARIANT)
   ' // Assignments
   DECLARE SUB Put (BYREF wszStr AS WSTRING)
   DECLARE SUB Put (BYREF cws AS CWSTR)
   DECLARE SUB Put (BYREF cbs AS CBSTR)
   DECLARE FUNCTION Put (BYVAL pdisp AS IDispatch PTR, BYVAL fAddRef AS BOOLEAN = FALSE) AS HRESULT
   DECLARE FUNCTION Put (BYVAL punk AS IUnknown PTR, BYVAL fAddRef AS BOOLEAN = FALSE) AS HRESULT
   DECLARE FUNCTION Put (BYREF cv AS CPropVar) AS HRESULT
   DECLARE FUNCTION Put (BYREF pv AS PROPVARIANT) AS HRESULT
   DECLARE FUNCTION Put (BYVAL pPropVar AS PROPVARIANT PTR) AS HRESULT
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
   DECLARE FUNCTION PutDateString (BYVAL pwszDate AS WSTRING PTR, BYVAL lcid AS LCID = 0, BYVAL dwFlags AS ULONG = 0) AS HRESULT
   DECLARE FUNCTION PutVbDate (BYVAL vbDate AS DATE_) AS HRESULT
   DECLARE FUNCTION PutSystemTime (BYVAL st AS SYSTEMTIME PTR) AS BOOLEAN
   DECLARE FUNCTION PutGuidAsString (BYVAL guid AS GUID PTR) AS HRESULT
   DECLARE FUNCTION PutGuidAsBuffer (BYVAL guid AS GUID PTR) AS HRESULT
   DECLARE FUNCTION PutClsid (BYVAL rclsid AS CLSID PTR) AS HRESULT
   DECLARE FUNCTION PutFileTime (BYVAL pft AS FILETIME PTR) AS HRESULT
   DECLARE FUNCTION PutFileTimeArray (BYVAL prgft AS FILETIME PTR, BYVAL cElems AS ULONG) AS HRESULT
   DECLARE FUNCTION PutStrRet (BYVAL pstrret AS STRRET PTR, BYVAL pidl AS PCUITEMID_CHILD) AS HRESULT
   DECLARE FUNCTION PutDec (BYREF dec AS DECIMAL) AS HRESULT
   DECLARE FUNCTION PutDecFromStr (BYVAL pwszIn AS WSTRING PTR, BYVAL lcid AS LCID = 0, BYVAL dwFlags AS ULONG = 0) AS HRESULT
   DECLARE FUNCTION PutDecFromDouble (BYVAL dbIn AS DOUBLE) AS HRESULT
   DECLARE FUNCTION PutDecFromCy (BYVAL cyIn AS CY) AS HRESULT
   DECLARE FUNCTION PutBooleanArray (BYVAL prgf AS WINBOOL PTR, BYVAL cElems AS ULONG) AS HRESULT
   DECLARE FUNCTION PutShortArray (BYVAL prgf AS SHORT PTR, BYVAL cElems AS ULONG) AS HRESULT
   DECLARE FUNCTION PutUShortArray (BYVAL prgf AS USHORT PTR, BYVAL cElems AS ULONG) AS HRESULT
   DECLARE FUNCTION PutLongArray (BYVAL prgn AS LONG PTR, BYVAL cElems AS ULONG) AS CPropVar
   DECLARE FUNCTION PutULongArray (BYVAL prgn AS ULONG PTR, BYVAL cElems AS ULONG) AS HRESULT
   DECLARE FUNCTION PutLongIntArray (BYVAL prgn AS LONGINT PTR, BYVAL cElems AS ULONG) AS HRESULT
   DECLARE FUNCTION PutULongIntArray (BYVAL prgn AS ULONGINT PTR, BYVAL cElems AS ULONG) AS HRESULT
   DECLARE FUNCTION PutDoubleArray (BYVAL prgn AS DOUBLE PTR, BYVAL cElems AS ULONG) AS HRESULT
   DECLARE FUNCTION PutStringArray (BYVAL prgsz AS PCWSTR, BYVAL cElems AS ULONG) AS HRESULT
   DECLARE FUNCTION PutStringAsArray (BYREF wsz AS WSTRING) AS HRESULT
   DECLARE FUNCTION PutPropVariant (BYVAL pPropVar AS PROPVARIANT PTR) AS HRESULT
   DECLARE FUNCTION PutPropVariantArrayElem (BYVAL pPropVarIn AS PROPVARIANT PTR, BYVAL iElem AS ULONG) AS CPropVar
   DECLARE FUNCTION PropVariantArrayFromPropVariant (BYVAL propvarSingle AS PROPVARIANT PTR) AS HRESULT
   ' // Assignments by reference
   DECLARE FUNCTION PutRef (BYVAL _value AS ANY PTR, BYVAL _vType AS WORD = VT_I8) AS HRESULT
   DECLARE FUNCTION PutRef (BYVAL _pvar AS ANY PTR, BYREF strType AS STRING) AS HRESULT
   ' // Safe arrays
   DECLARE FUNCTION GetDim () AS ULONG
   DECLARE FUNCTION GetLBound (BYVAL nDim AS UINT = 1) AS LONG
   DECLARE FUNCTION GetUBound (BYVAL nDim AS UINT = 1) AS LONG
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
   DECLARE FUNCTION GetFileTimeElem (BYVAL iElem AS ULONG) AS FILETIME
   DECLARE FUNCTION GetPropVariantElem (BYVAL iElem AS ULONG) AS CPropVar
   ' // Other...
   DECLARE SUB Clear
   DECLARE FUNCTION Attach (BYVAL pPropVar AS PROPVARIANT PTR) AS HRESULT
   DECLARE FUNCTION Attach (BYREF pv AS PROPVARIANT) AS HRESULT
   DECLARE FUNCTION Detach (BYVAL pPropVar AS PROPVARIANT PTR) AS HRESULT
   DECLARE FUNCTION Detach (BYREF pv AS PROPVARIANT) AS HRESULT
   DECLARE FUNCTION vType () AS VARTYPE
   DECLARE FUNCTION ChangeType (BYVAL vtNew AS VARTYPE) AS HRESULT
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

END TYPE
' ########################################################################################

' ========================================================================================
' Note about constructors:
' We can use the constructors to pass values to parameters in procedures without assigning
' them first to a variable, e.g.:
' SUB Foo (BYREF cv AS CPropVar)   ' -or- (BYVAL cv AS CPropVar PTR)
'    PRINT AfxCPropVarToStr(cv)
' END SUB
' Foo CPropVar("Test string")
' Foo CPropVar(12345)
' Foo CPropVar(12345, "LONG")
' ========================================================================================

' ========================================================================================
' * Default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CPropVar
   CPROPVAR_DP("CPropVar CONSTRUCTOR")
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' * Destructor
' ========================================================================================
PRIVATE DESTRUCTOR CPropVar
   CPROPVAR_DP("CPropVar DESTRUCTOR")
   PropVariantClear(@pvd)
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' * Initializes the CPropVar from another CPropVar.
' ========================================================================================
PRIVATE CONSTRUCTOR CPropVar (BYREF cpv AS CPropVar)
   CPROPVAR_DP("CPropVar CONSTRUCTOR - BYREF CPropVar")
'   PropVariantCopy(@pvd, @cpv)   ' // removed @ operator
   PropVariantCopy(@pvd, cpv.sptr)
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' * Initializes the CPropVar from a PROPVARIANT.
' ========================================================================================
PRIVATE CONSTRUCTOR CPropVar (BYVAL pv AS PROPVARIANT)
   CPROPVAR_DP("CPropVar CONSTRUCTOR - BYVAL PROPVARIANT")
   PropVariantCopy(@pvd, @pv)
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' * Initializes the CPropVar from a VARIANT.
' We can also use a CVAR:
' DIM cv AS CVAR = "pepe"
' DIM p AS CPropVar = *cv
' ========================================================================================
PRIVATE CONSTRUCTOR CPropVar (BYVAL v AS VARIANT PTR)
   CPROPVAR_DP("CPropVar CONSTRUCTOR - BYVAL VARIANT")
   AfxVariantToPropVariant(v, @pvd)
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' * Initializes the CPropVar from a WSTRING.
' ========================================================================================
PRIVATE CONSTRUCTOR CPropVar (BYREF wsz AS WSTRING)
   CPROPVAR_DP("CPropVar CONSTRUCTOR - WSTRING")
   pvd.vt = VT_LPWSTR : SHStrDupW(@wsz, @pvd.pwszVal)
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' * Initializes the CPropVar from a CWSTR.
' ========================================================================================
PRIVATE CONSTRUCTOR CPropVar (BYREF cws AS CWSTR)
   CPROPVAR_DP("CPropVar CONSTRUCTOR - CWSTR")
   pvd.vt = VT_LPWSTR : SHStrDupW(cws, @pvd.pwszVal)
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' * Initializes the CPropVar from a CBSTR.
' ========================================================================================
PRIVATE CONSTRUCTOR CPropVar (BYREF cbs AS CBSTR)
   CPROPVAR_DP("CPropVar CONSTRUCTOR - CBSTR")
   pvd.vt = VT_LPWSTR : SHStrDupW(cbs, @pvd.pwszVal)
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' * Initializes a CPropVar from a pointer to a PROPVARIANT.
' DIM pp AS PROPVARIANT
' pp.vt = VT_I4
' pp.lVal = 12345
' DIM p AS CPropVar = @pp
' Ve can also use a CPropVar:
' DIM p AS CPropVar = "Test string"
' DIM p2 AS CPropVar = *p
' ========================================================================================
PRIVATE CONSTRUCTOR CPropVar (BYVAL pPropVar AS PROPVARIANT PTR)
   CPROPVAR_DP("CPropVar CONSTRUCTOR - PROPVARIANT PTR")
   PropVariantCopy(@pvd, pPropVar)
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Initializes a CPropVar from a CURRENCY structure.
' ========================================================================================
PRIVATE CONSTRUCTOR CPropVar (BYVAL cy AS CURRENCY)
   CPROPVAR_DP("CPropVar CONSTRUCTOR - CURRENCY")
   pvd.vt = VT_CY
   pvd.cyVal = cy
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Initializes a CPropVar from a DECIMAL structure.
' ========================================================================================
PRIVATE CONSTRUCTOR CPropVar (BYVAL dec AS DECIMAL)
   CPROPVAR_DP("CPropVar CONSTRUCTOR - DECIMAL")
   pvd.vt = VT_DECIMAL
   pvd.decVal = dec
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Initializes a CPropVar from an IDispatch pointer
' ========================================================================================
PRIVATE CONSTRUCTOR CPropVar (BYVAL pdisp AS IDispatch PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   CPROPVAR_DP("CPropVar CONSTRUCTOR - IDISPATCH PTR")
   IF fAddRef THEN IDispatch_AddRef(pdisp)
   pvd.vt = VT_DISPATCH : pvd.pdispVal = pdisp
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Initializes a CPropVar from an IUnknown pointer
' ========================================================================================
PRIVATE CONSTRUCTOR CPropVar (BYVAL punk AS IUnknown PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   CPROPVAR_DP("CPropVar CONSTRUCTOR - IUNKNOWN PTR")
   IF fAddRef THEN IUnknown_AddRef(punk)
   pvd.vt = VT_UNKNOWN : pvd.punkVal = punk
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Initializes a CPropVar from an IStream pointer
' ========================================================================================
PRIVATE CONSTRUCTOR CPropVar (BYVAL pStream AS IStream PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   CPROPVAR_DP("CPropVar CONSTRUCTOR - IUNKNOWN PTR")
   IF fAddRef THEN IUnknown_AddRef(pStream)
   pvd.vt = VT_STREAM : pvd.pStream = pStream
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Initializes a CPropVar from an IStream pointer
' ========================================================================================
PRIVATE CONSTRUCTOR CPropVar (BYVAL pStorage AS IStorage PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   CPROPVAR_DP("CPropVar CONSTRUCTOR - IUNKNOWN PTR")
   IF fAddRef THEN IUnknown_AddRef(pStorage)
   pvd.vt = VT_STORAGE : pvd.pStorage = pStorage
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Initializes a CPropVar from a versioned stream
' ========================================================================================
PRIVATE CONSTRUCTOR CPropVar (BYVAL pVersionedStream AS VERSIONEDSTREAM PTR)
   CPROPVAR_DP("CPropVar CONSTRUCTOR - VERSIONEDSTREAM PTR")
   pvd.vt = VT_VERSIONED_STREAM : pvd.pVersionedStream = pVersionedStream
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Initializes a CPropVar from a BLOB structure
' ========================================================================================
PRIVATE CONSTRUCTOR CPropVar (BYREF _blob AS BLOB)
   CPROPVAR_DP("CPropVar CONSTRUCTOR - BLOB")
   pvd.vt = VT_BLOB : pvd.blob = _blob
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Initializes a CPropVar from a BSTRBLOB structure
' ========================================================================================
PRIVATE CONSTRUCTOR CPropVar (BYREF _bstrBlob AS BSTRBLOB)
   CPROPVAR_DP("CPropVar CONSTRUCTOR - BSTRBLOB")
   pvd.vt = VT_BSTR_BLOB : pvd.bstrblobVal = _bstrBlob
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Initializes a CPropVar from a CLIPDATA structure
' ========================================================================================
PRIVATE CONSTRUCTOR CPropVar (BYVAL pclipData AS CLIPDATA PTR)
   CPROPVAR_DP("CPropVar CONSTRUCTOR - CLIPDATA PRE")
   pvd.vt = VT_CF : pvd.pclipdata = pclipData
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Initializes a CPropVar from an integer value
' If _vType is wrong, default to VT_I8 (LongInt).
' ========================================================================================
PRIVATE CONSTRUCTOR CPropVar (BYVAL _value AS LONGINT, BYVAL _vType AS WORD = VT_I4)
   CPROPVAR_DP("CPropVar CONSTRUCTOR - value, type")
   IF (_vType AND VT_ARRAY) = VT_ARRAY THEN
      ' // NULL array : Suitable for parameters that, instead of being optional,
      ' // require a variant with a null safearray of a given type.
      pvd.vt = _vType
   ELSE
      ' // If _value exceeds the limits of max long or max ulong, promote to ulongint
      IF SGN(_value) = -1 AND _value > LONG_MAX THEN
         _vType = VT_I8
      ELSEIF _value > ULONG_MAX THEN
         _vType = VT_I8
      END IF
      SELECT CASE _vType
         CASE VT_BOOL  : pvd.vt = VT_BOOL : pvd.boolVal = IIF(_value = 0, 0, -1)
         CASE VT_I1    : pvd.vt = VT_I1   : pvd.cVal = CBYTE(_value)
         CASE VT_UI1   : pvd.vt = VT_UI1  : pvd.bVal = CUBYTE(_value)
         CASE VT_I2    : pvd.vt = VT_I2   : pvd.iVal = CSHORT(_value)
         CASE VT_UI2   : pvd.vt = VT_UI2  : pvd.uiVal = CUSHORT(_value)
         CASE VT_INT   : pvd.vt = VT_INT  : pvd.intVal = CLNG(_value)
         CASE VT_UINT  : pvd.vt = VT_UINT : pvd.uintVal = CULNG(_value)
         CASE VT_I4    : pvd.vt = VT_I4   : pvd.lVal = CLNG(_value)
         CASE VT_UI4   : pvd.vt = VT_UI4  : pvd.ulVal = CULNG(_value)
         CASE VT_I8    : pvd.vt = VT_I8   : pvd.hVal.QuadPart = CLNGINT(_value)
         CASE VT_UI8   : pvd.vt = VT_UI8  : pvd.uhVal.QuadPart = CULNGINT(_value)
         ' // If you need to assign an ULongInt value greater than a LongInt, use PutUlongInt.
         CASE VT_NULL  : pvd.vt = VT_NULL
         CASE ELSE     : pvd.vt = VT_I4   : pvd.hVal.QuadPart = CLNG(_value)
      END SELECT
   END IF
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CPropVar (BYVAL _value AS LONGINT, BYREF strType AS STRING)
   CPROPVAR_DP("CPropVar CONSTRUCTOR - strType")
   DIM vt AS WORD
   SELECT CASE UCASE(strType)
      CASE "BOOL"     : pvd.vt = VT_BOOL : pvd.boolVal = IIF(_value = 0, 0, -1)
      CASE "BOOLEAN"  : pvd.vt = VT_BOOL : pvd.boolVal = IIF(_value = 0, 0, -1)
      CASE "BYTE"     : pvd.vt = VT_I1   : pvd.cVal = CBYTE(_value)
      CASE "UBYTE"    : pvd.vt = VT_UI1  : pvd.bVal = CUBYTE(_value)
      CASE "SHORT"    : pvd.vt = VT_I2   : pvd.iVal = CSHORT(_value)
      CASE "USHORT"   : pvd.vt = VT_UI2  : pvd.uiVal = CUSHORT(_value)
      CASE "INT"      : pvd.vt = VT_INT  : pvd.intVal = CLNG(_value)
      CASE "UINT"     : pvd.vt = VT_UINT : pvd.uintVal = CULNG(_value)
      CASE "LONG"     : pvd.vt = VT_I4   : pvd.lVal = CLNG(_value)
      CASE "ULONG"    : pvd.vt = VT_UI4  : pvd.ulVal = CULNG(_value)
      CASE "LONGINT"  : pvd.vt = VT_I8   : pvd.hVal.QuadPart = CLNGINT(_value)
      CASE "ULONGINT" : pvd.vt = VT_UI8  : pvd.uhVal.QuadPart = CULNGINT(_value)
      ' // If you need to assign an ULongInt value greater than a LongInt, use PutUlongInt.
      CASE "NULL"     : pvd.vt = VT_NULL
      CASE ELSE       : pvd.vt = VT_I4   : pvd.hVal.QuadPart = CLNG(_value)
   END SELECT
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Initializes a CPropVar from a float value
' If _vType is wrong, default to VT_R8 (double).
' ========================================================================================
PRIVATE CONSTRUCTOR CPropVar (BYVAL _value AS DOUBLE, BYVAL _vType AS WORD = VT_R8)
   CPROPVAR_DP("CPropVar CONSTRUCTOR - FLOAT")
   IF _value > FLT_MAX OR _value < FLT_MIN THEN _vType = VT_R8
   SELECT CASE _vType
      CASE VT_R4    : pvd.vt = VT_R4   : pvd.fltVal = CSNG(_value)
      CASE VT_R8    : pvd.vt = VT_R8   : pvd.dblVal = CDBL(_value)
      CASE ELSE     : pvd.vt = VT_R8   : pvd.dblVal = CDBL(_value)
   END SELECT
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CPropVar (BYVAL _value AS DOUBLE, BYREF strType AS STRING)
   CPROPVAR_DP("CPropVar CONSTRUCTOR - strType")
   SELECT CASE UCASE(strType)
      CASE "SINGLE" : pvd.vt = VT_R4   : pvd.fltVal = CSNG(_value)
      CASE "FLOAT"  : pvd.vt = VT_R4   : pvd.fltVal = CSNG(_value)
      CASE "DOUBLE" : pvd.vt = VT_R8   : pvd.dblVal = CDBL(_value)
      CASE ELSE     : pvd.vt = VT_R8   : pvd.dblVal = CDBL(_value)
   END SELECT
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Initializes a CPropVar from a value by reference (a pointer to a variable), e.g.:
' DIM dblVal AS DOUBLE = 123456.12
' DIM cpv AS CPropVar = CPropVar(@dblVal, "DOUBLE")
' print cpv.ToStr   ' // prints 123456.12
' dblVal = 345.12
' print cpv.ToStr   ' // prints 345.12
' ========================================================================================
PRIVATE CONSTRUCTOR CPropVar (BYVAL _pPropVar AS ANY PTR, BYVAL _vType AS WORD)
   CPROPVAR_DP("CPropVar CONSTRUCTOR - ANY PTR - vType")
   SELECT CASE _vType
      CASE VT_BOOL      : pvd.vt = VT_BOOL OR VT_BYREF      : pvd.pboolVal = _pPropVar
      CASE VT_I1        : pvd.vt = VT_I1 OR VT_BYREF        : pvd.pcVal = _pPropVar
      CASE VT_UI1       : pvd.vt = VT_UI1 OR VT_BYREF       : pvd.pbVal = _pPropVar
      CASE VT_I2        : pvd.vt = VT_I2 OR VT_BYREF        : pvd.piVal = _pPropVar
      CASE VT_UI2       : pvd.vt = VT_UI2 OR VT_BYREF       : pvd.puiVal = _pPropVar
      CASE VT_INT       : pvd.vt = VT_INT OR VT_BYREF       : pvd.pintVal = _pPropVar
      CASE VT_UINT      : pvd.vt = VT_UINT OR VT_BYREF      : pvd.puintVal = _pPropVar
      CASE VT_I4        : pvd.vt = VT_I4 OR VT_BYREF        : pvd.plVal = _pPropVar
      CASE VT_UI4       : pvd.vt = VT_UI4 OR VT_BYREF       : pvd.pulVal = _pPropVar
'      CASE VT_I8        : pvd.vt = VT_I8 OR VT_BYREF        : pvd.pllVal = _pPropVar
'      CASE VT_UI8       : pvd.vt = VT_UI8 OR VT_BYREF       : pvd.pullVal = _pPropVar
      CASE VT_R4        : pvd.vt = VT_R4 OR VT_BYREF        : pvd.pfltVal = _pPropVar
      CASE VT_R8        : pvd.vt = VT_R8 OR VT_BYREF        : pvd.pdblVal = _pPropVar
      CASE VT_LPWSTR    : pvd.vt = VT_LPWSTR OR VT_BYREF    : pvd.pwszVal = _pPropVar
      CASE VT_BSTR      : pvd.vt = VT_BSTR OR VT_BYREF      : pvd.pbstrVal = _pPropVar
      CASE VT_UNKNOWN   : pvd.vt = VT_UNKNOWN OR VT_BYREF   : pvd.ppunkVal = _pPropVar
      CASE VT_DISPATCH  : pvd.vt = VT_DISPATCH OR VT_BYREF  : pvd.ppdispVal = _pPropVar
      CASE VT_DECIMAL   : pvd.vt = VT_DECIMAL OR VT_BYREF   : pvd.pdecVal = _pPropVar
      CASE VT_CY        : pvd.vt = VT_CY OR VT_BYREF        : pvd.pcyVal = _pPropVar
      CASE VT_DATE      : pvd.vt = VT_DATE OR VT_BYREF      : pvd.pdate = _pPropVar
      CASE VT_VARIANT   : pvd.vt = VT_VARIANT OR VT_BYREF   : pvd.pvarVal = _pPropVar
      CASE VT_SAFEARRAY : pvd.vt = VT_SAFEARRAY OR VT_BYREF : pvd.pvarVal = _pPropVar
      CASE VT_ERROR     : pvd.vt = VT_ERROR OR VT_BYREF     : pvd.pparray = _pPropVar
   END SELECT
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CPropVar (BYVAL _pPropVar AS ANY PTR, BYREF strType AS STRING)
   CPROPVAR_DP("CPropVar CONSTRUCTOR - ANY PTR - strType")
   IF _pPropVar <> NULL THEN
      SELECT CASE UCASE(strType)
         CASE "BOOL"      : pvd.vt = VT_BOOL OR VT_BYREF      : pvd.pboolVal = _pPropVar
         CASE "BOOLEAN"   : pvd.vt = VT_BOOL OR VT_BYREF      : pvd.pboolVal = _pPropVar
         CASE "BYTE"      : pvd.vt = VT_I1 OR VT_BYREF        : pvd.pcVal = _pPropVar
         CASE "UBYTE"     : pvd.vt = VT_UI1 OR VT_BYREF       : pvd.pbVal = _pPropVar
         CASE "SHORT"     : pvd.vt = VT_I2 OR VT_BYREF        : pvd.piVal = _pPropVar
         CASE "USHORT"    : pvd.vt = VT_UI2 OR VT_BYREF       : pvd.puiVal = _pPropVar
         CASE "INT"       : pvd.vt = VT_INT OR VT_BYREF       : pvd.pintVal = _pPropVar
         CASE "UINT"      : pvd.vt = VT_UINT OR VT_BYREF      : pvd.puintVal = _pPropVar
         CASE "LONG"      : pvd.vt = VT_I4 OR VT_BYREF        : pvd.plVal = _pPropVar
         CASE "ULONG"     : pvd.vt = VT_UI4 OR VT_BYREF       : pvd.pulVal = _pPropVar
'         CASE "LONGINT"   : pvd.vt = VT_I8 OR VT_BYREF        : pvd.pllVal = _pPropVar
'         CASE "ULONGINT"  : pvd.vt = VT_UI8 OR VT_BYREF       : pvd.pullVal = _pPropVar
         CASE "SINGLE"    : pvd.vt = VT_R4 OR VT_BYREF        : pvd.pfltVal = _pPropVar
         CASE "FLOAT"     : pvd.vt = VT_R4 OR VT_BYREF        : pvd.pfltVal = _pPropVar
         CASE "DOUBLE"    : pvd.vt = VT_R8 OR VT_BYREF        : pvd.pdblVal = _pPropVar
         CASE "WSTRING"   : pvd.vt = VT_LPWSTR OR VT_BYREF    : pvd.pwszVal = _pPropVar
         CASE "BSTR"      : pvd.vt = VT_BSTR OR VT_BYREF      : pvd.pbstrVal = _pPropVar
         CASE "UNKNOWN"   : pvd.vt = VT_UNKNOWN OR VT_BYREF   : pvd.ppunkVal = _pPropVar
         CASE "DISPATCH"  : pvd.vt = VT_DISPATCH OR VT_BYREF  : pvd.ppdispVal = _pPropVar
         CASE "DECIMAL"   : pvd.vt = VT_DECIMAL OR VT_BYREF   : pvd.pdecVal = _pPropVar
         CASE "CY"        : pvd.vt = VT_CY OR VT_BYREF        : pvd.pcyVal = _pPropVar
         CASE "CURRENCY"  : pvd.vt = VT_CY OR VT_BYREF        : pvd.pcyVal = _pPropVar
         CASE "DATE"      : pvd.vt = VT_DATE OR VT_BYREF      : pvd.pdate = _pPropVar
         CASE "VARIANT"   : pvd.vt = VT_VARIANT OR VT_BYREF   : pvd.pvarVal = _pPropVar
         CASE "SAFEARRAY" : pvd.vt = VT_SAFEARRAY OR VT_BYREF : pvd.pvarVal = _pPropVar
         CASE "ERROR"     : pvd.vt = VT_ERROR OR VT_BYREF     : pvd.pparray = _pPropVar
      END SELECT
   END IF
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the address of the underlying variant.
' Removed to allow to use @ to get the address of the class.
' After removing it, we can now do:
' SUB Foo (BYVAL cpv AS CPropVar PTR)
'   PRINT AfxCPropVarToStr(cpv)
' END SUB
' DIM cpv AS CPropVar = "Test string"
' Foo @cpv
' --Or--:
' SUB Foo (BYREF pv AS PROPVARIANT)
'    PRINT AfxPropVarToStr(@pv)
' END SUB
' DIM cpv AS CPropVar = "Test string"
' Foo cpv
' ========================================================================================
'PRIVATE OPERATOR CPropVar.@ () AS PROPVARIANT PTR
'   CPROPVAR_DP("CPropVar OPERATOR @ - " & WSTR(@pvd))
'   OPERATOR = @pvd
'END OPERATOR
' ========================================================================================

' ========================================================================================
' Clears the CPropVar and returns the address of the underlying variant.
' To pass the variant to an OUT BYVAL VARIANT PTR parameter.
' If we pass a CPropVar to a function with an OUT variant parameter without first clearing the
' contents of the CPropVar, we may have a memory leak.
' Example:
' SUB Foo (BYVAL pv AS PROPVARIANT PTR)
'    v->vt = VT_I4
'    v->lVal = 12345
' END SUB
' DIM cpv AS CPropVar = "Test string"
' Foo cpv.vptr
' PRINT cpv.ToStr
' Otherwise, you need to clear the underlying variant before passing the CPropVar.
' DIM cpv AS CPropVar = "Test string"
' cpv.Clear
' Foo *cpv
' PRINT cpv.ToStr
' ========================================================================================
PRIVATE FUNCTION CPropVar.vptr () AS PROPVARIANT PTR
   CPROPVAR_DP("CPropVar FUNCTION vptr")
   PropVariantClear @pvd
   RETURN @pvd
END FUNCTION
' ========================================================================================

' ========================================================================================
' To pass the PropVariant to a IN BYVAL PROPVARIANT PTR parameter.
' Usage example:
' SUB Foo (BYVAL pv AS PROPVARIANT PTR)
'    PRINT AfxPropVarToStr(pv)
' END SUB
' Using the pointer syntax:
' DIM pcpv AS CPropVar PTR = NEW CPropVar("Test string")
' Foo pcpv->sptr
' Delete pcpv
' Using the normal syntax:
' DIM cpv AS CPropVar = "Test string"
' Foo cpv.sptr
' But with the normal syntax you can use the * operator instead:
' DIM cpv AS CPropVar = "Test string"
' Foo *cpv
' ========================================================================================
PRIVATE FUNCTION CPropVar.sptr () AS PROPVARIANT PTR
   CPROPVAR_DP("CPropVar FUNCTION sptr")
   RETURN @pvd
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the address of the underlying PROPVARIANT.
' One * returns the address of the underlying PROPVARIANT.
' Two ** deferences the PROPVARIANT data.
' Can be used to pass the PROPVARIANT to a BYVAL VARIANT parameter, e.g.
' SUB Foo (BYVAL pv AS PROPVARIANT)
'    PRINT AfxPropVarToStr(@pv)
' END SUB
' DIM cpv AS CPropVar = "Test string"
' Foo **cpv
' -or-
' Foo *cpv.sptr
' Using the pointer syntax:
' DIM pcpv AS CPropVar PTR  = NEW CPropVar("Test string")
' Foo *pcpv
' Delete pcpv
' Using the constructors_
' Foo **CPropVar(12345, "LONG")
' ========================================================================================
PRIVATE OPERATOR * (BYREF cpv AS CPropVar) AS PROPVARIANT PTR
   CPROPVAR_DP("CPropVar OPERATOR *")
   OPERATOR = @cpv.pvd
END OPERATOR
' ========================================================================================

' ========================================================================================
' The CAST operators allow to transparently pass the underlying VARIANT to a procedure.
' They aren't called directly.
'
' SUB Foo (BYREF pv AS PROPVARIANT)
'    PRINT AfxPropVarToStr(@pv)
' END SUB
' Foo CPropVar(12345, "LONG")
'
' SUB Foo2 (BYVAL pv AS PROPVARIANT)
'    PRINT AfxPropVarToStr(@pv)
' END SUB
' Foo2 CPropVar(12345, "LONG")
'
' SUB Foo3 (BYREF cpv AS CPropVar)
'    PRINT cpv.ToStr
' END SUB
' Foo3 CPropVar(12345, "LONG")
'
' SUB Foo4 (BYVAL cpv AS CPropVar PTR)
'    PRINT cpv->ToStr
' END SUB
' Foo4 CPropVar(12345, "LONG")
' ========================================================================================

' ========================================================================================
' Returns the variant data.
' ========================================================================================
PRIVATE OPERATOR CPropVar.CAST () AS PROPVARIANT
   CPROPVAR_DP("CPropVar CAST VARIANT")
   OPERATOR = pvd
END OPERATOR
' ========================================================================================
' ========================================================================================
' Returns a pointer to the variant data.
' ========================================================================================
PRIVATE OPERATOR CPropVar.CAST () AS ANY PTR
   CPROPVAR_DP("CPropVar CAST ANY PTR")
   OPERATOR = cast(ANY PTR, @pvd)
END OPERATOR
' ========================================================================================

' =====================================================================================
' Extracts the contents of a VARIANT and returns them as a WSTRING.
' =====================================================================================
PRIVATE OPERATOR CPropVar.CAST () BYREF AS WSTRING
   CPROPVAR_DP("CPropVar CAST BYREF AS WSTRING")
   cws = AfxPropVarToStr(@pvd)
   OPERATOR = *cast(WSTRING PTR, cws.m_pBuffer)
END OPERATOR
' =====================================================================================

' =====================================================================================
' Extracts the contents of a PROPVARIANT and returns them as a CWSTR.
' =====================================================================================
PRIVATE FUNCTION CPropVar.ToStr () AS CWSTR
   CPROPVAR_DP("CPropVar ToStr")
   RETURN AfxPropVarToStr(@pvd)
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CPropVar.ToWStr () AS CWSTR
   CPROPVAR_DP("CPropVar ToWStr")
   RETURN AfxPropVarToStr(@pvd)
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CPropVar.wstr () AS CWSTR
   CPROPVAR_DP("CPropVar wstr")
   RETURN AfxPropVarToStr(@pvd)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Extracts the contents of a PROPVARIANT and returns them as a CBSTR.
' =====================================================================================
PRIVATE FUNCTION CPropVar.ToBStr () AS CBSTR
   CPROPVAR_DP("CPropVar ToBStr")
   RETURN AfxPropVarToStr(@pvd)
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CPropVar.bstr () AS CBSTR
   CPROPVAR_DP("CPropVar bstr")
   RETURN AfxPropVarToStr(@pvd)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Extracts an unicode string contained in a variant and maps it to an encoded UTF8 code page.
' =====================================================================================
PRIVATE FUNCTION CPropVar.ToUtf8 () AS STRING
   CPROPVAR_DP("CPropVar ToUtf8")
   IF pvd.vt = VT_BSTR THEN RETURN AfxAcode(pvd.bstrVal, CP_UTF8)
   IF pvd.vt = VT_LPWSTR THEN RETURN AfxAcode(pvd.pwszVal, CP_UTF8)
   RETURN ""
END FUNCTION
' =====================================================================================

' =====================================================================================
' Extracts the contents of a buffer stored in a CPropVar of type VT_ARRRAY OR VT_UI1 or
' VT_VECTOR OR VT_UI1.
' Note: To retrieve the size of the array call GetElementCount.
' =====================================================================================
PRIVATE FUNCTION CPropVar.ToBuffer (BYVAL pv AS ANY PTR, BYVAL cb AS UINT) AS HRESULT
   CPROPVAR_DP("CPropVar ToBuffer")
   RETURN AfxPropVariantToBuffer(@pvd, pv, cb)
END FUNCTION
' =====================================================================================
' =====================================================================================
' Extracts the contents of a buffer stored in a CPropVar of type VT_ARRRAY OR VT_UI1 or
' VT_VECTOR OR VT_UI1 and returns it as string used as a buffer.
' =====================================================================================
PRIVATE FUNCTION CPropVar.ToBuffer () AS STRING
   CPROPVAR_DP("CPropVar ToBuffer - STRING")
   IF (pvd.vt = VT_ARRAY OR VT_UI1) OR (pvd.vt = VT_VECTOR OR VT_UI1) THEN
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
PRIVATE FUNCTION CPropVar.ToUnknown () AS ANY PTR
   IF pvd.vt <> VT_UNKNOWN AND pvd.vt <> VT_DISPATCH THEN RETURN NULL
   IF pvd.vt = VT_UNKNOWN THEN
      IF pvd.punkVal = NULL THEN RETURN NULL
      IUnknown_AddRef(pvd.punkVal)
      RETURN pvd.punkVal
   END IF
   IF pvd.vt = VT_DISPATCH THEN
      IF pvd.pdispVal = NULL THEN RETURN NULL
      DIM punk AS IUnknown PTR
      DIM hr AS HRESULT = pvd.pdispVal->lpvtbl->QueryInterface(pvd.pdispVal, @IID_IUnknown, @punk)
      IF hr = S_OK THEN
         ' // QueryInterface calls IUnknown_AddRef in the pointer it returns
         RETURN punk
      ELSE
         IDispatch_AddRef(pvd.pdispVal)
         RETURN pvd.pdispVal
      END IF
   END IF
END FUNCTION
' =====================================================================================

' =====================================================================================
' If the variant is of the type VT_UNKNOWN or VT_DISPATCH it returns the dispatch pointer.
' It is the responsability of the called of release the returned pointer.
' =====================================================================================
PRIVATE FUNCTION CPropVar.ToDispatch () AS ANY PTR
   IF pvd.vt <> VT_UNKNOWN AND pvd.vt <> VT_DISPATCH THEN RETURN NULL
   IF pvd.vt = VT_DISPATCH THEN
      IF pvd.pdispVal = NULL THEN RETURN NULL
      IDispatch_AddRef(pvd.pdispVal)
      RETURN pvd.pdispVal
   END IF
   IF pvd.vt = VT_UNKNOWN THEN
      IF pvd.punkVal = NULL THEN RETURN NULL
      DIM pdisp AS IDispatch PTR
      DIM hr AS HRESULT = pvd.punkVal->lpvtbl->QueryInterface(pvd.punkVal, @IID_IDispatch, @pdisp)
      IF hr = S_OK THEN
         ' // QueryInterface calls IUnknown_AddRef in the pointer it returns
         RETURN pdisp
      ELSE
         IUnknown_AddRef(pvd.punkVal)
         RETURN pvd.punkVal
      END IF
   END IF
END FUNCTION
' =====================================================================================

' =====================================================================================
' Converts the contents of a PROPVARIANT structure to a VARIANT structure.
' =====================================================================================
PRIVATE FUNCTION CPropVar.ToVariant () AS VARIANT
   CPROPVAR_DP("CPropVar YoVariant")
   DIM v AS VARIANT
   AfxPropVariantToVariant(@pvd, @v)
   RETURN v
END FUNCTION
' =====================================================================================

' =====================================================================================
' Converts a CPropVar of type decimal to a double.
' =====================================================================================
PRIVATE FUNCTION CPropVar.DecToDouble () AS DOUBLE
   CPROPVAR_DP("CPropVar DecToDouble")
   DIM dblOut AS DOUBLE
   IF pvd.vt = VT_DECIMAL THEN VarR8FromDec(@pvd.decVal, @dblOut)
   RETURN dblOut
END FUNCTION
' =====================================================================================

' =====================================================================================
' Converts a variant of type decimal to currency.
' =====================================================================================
PRIVATE FUNCTION CPropVar.DecToCy () AS CY
   CPROPVAR_DP("CPropVar DecToCy")
   DIM cyOut AS CY
   IF pvd.vt = VT_DECIMAL THEN VarCyFromDec(@pvd.decVal, @cyOut)
   RETURN cyOut
END FUNCTION
' =====================================================================================

' =====================================================================================
' Converts a VT_DATE variant to a variant representation (double)
' =====================================================================================
PRIVATE FUNCTION CPropVar.ToVbDate () AS DATE_
   CPROPVAR_DP("CPropVar ToVbDate")
   IF pvd.vt = VT_DATE THEN RETURN pvd.date
END FUNCTION
' =====================================================================================

' =====================================================================================
' Extracts a FILETIME structure from a CPropVar as a SYSTEMTIME structure.
' =====================================================================================
PRIVATE FUNCTION CPropVar.ToSystemTime () AS SYSTEMTIME
   CPROPVAR_DP("CPropVar ToSystemTime")
   DIM st AS SYSTEMTIME
   IF pvd.vt = VT_DATE THEN VariantTimeToSystemTime(pvd.date, @st)
   RETURN st
END FUNCTION
' =====================================================================================

' =====================================================================================
' Extracts a GUID property value of a PROPVARIANT.
' =====================================================================================
PRIVATE FUNCTION CPropVar.ToGuid () AS GUID
   CPROPVAR_DP("CPropVar ToGuid")
   DIM guid_ AS GUID : AfxPropVariantToGUID(@pvd, @guid_) : RETURN guid_
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CPropVar.ToClsid () AS CLSID
   CPROPVAR_DP("CPropVar ToClsid")
   DIM guid_ AS CLSID : AfxPropVariantToGUID(@pvd, @guid_) : RETURN guid_
END FUNCTION
' =====================================================================================

' =====================================================================================
' Extracts a GUID property value as a CWSTR.
' =====================================================================================
PRIVATE FUNCTION CPropVar.ToGuidStr () AS CWSTR
   CPROPVAR_DP("CPropVar ToGuidStr")
   DIM guid_ AS GUID, cwsGuid AS CWSTR, pwsz AS WSTRING PTR
   DIM hr AS HRESULT = AfxPropVariantToGUID(@pvd, @guid_)
   IF hr = S_OK THEN
      StringFromCLSID(@guid_, CAST(LPOLESTR PTR, @pwsz))
      IF pwsz THEN cwsGuid = *pwsz
      CoTaskMemFree(pwsz)
   END IF
   RETURN cwsGuid
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CPropVar.ToGuidWStr () AS CWSTR
   CPROPVAR_DP("CPropVar ToGuidStr")
   DIM guid_ AS GUID, cwsGuid AS CWSTR, pwsz AS WSTRING PTR
   DIM hr AS HRESULT = AfxPropVariantToGUID(@pvd, @guid_)
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
PRIVATE FUNCTION CPropVar.ToGuidBStr () AS CBSTR
   CPROPVAR_DP("CPropVar ToGuidBStr")
   DIM guid_ AS GUID, cbsGuid AS CBSTR, pwsz AS WSTRING PTR
   DIM hr AS HRESULT = AfxPropVariantToGUID(@pvd, @guid_)
   IF hr = S_OK THEN
      StringFromCLSID(@guid_, CAST(LPOLESTR PTR, @pwsz))
      IF pwsz THEN cbsGuid = *pwsz
      CoTaskMemFree(pwsz)
   END IF
   RETURN cbsGuid
END FUNCTION
' =====================================================================================

' =====================================================================================
' Extracts a FILETIME structure from a PROPVARIANT.
' =====================================================================================
PRIVATE FUNCTION CPropVar.ToFileTime (BYVAL stfOut AS AFX_PSTIME_FLAGS) AS FILETIME
   CPROPVAR_DP("CPropVar ToFileTime")
   DIM ft AS FILETIME
   AfxPropVariantToFileTime(@pvd, stfOut, @ft)
   RETURN ft
END FUNCTION
' =====================================================================================

' =====================================================================================
' If the source PROPVARIANT is a VT_BSTR, extracts string and places it into a STRRET structure.
' =====================================================================================
PRIVATE FUNCTION CPropVar.ToStrRet () AS STRRET
   CPROPVAR_DP("CPropVar ToStrRet")
   DIM strret_ AS STRRET
   AfxPropVariantToStrRet(@pvd, @strret_)
   RETURN strret_
END FUNCTION
' =====================================================================================

' =====================================================================================
' Extracts an array of Boolean values from a PROPVARIANT structure.
' Returns the number of elements extracted from the source PROPVARIANT structure.
' GetLastError: Returns S_OK if successful, or an error value otherwise, including the following:
' TYPE_E_BUFFERTOOSMALL: The source PROPVARIANT contained more than crgn values.
' E_INVALIDARG: The PROPVARIANT was not of the appropriate type.
' To retrieve the number of elements in the array call the GetElementCount method.
' =====================================================================================
PRIVATE FUNCTION CPropVar.ToBooleanArray (BYVAL prgf AS WINBOOL PTR, BYVAL crgn AS ULONG) AS ULONG
   CPROPVAR_DP("CPropVar ToBooleanArray")
   DIM pcElem AS ULONG
   SetLastError AfxPropVariantToBooleanVector(@pvd, prgf, crgn, @pcElem)
   RETURN pcElem
END FUNCTION
' =====================================================================================
' =====================================================================================
' Allocates an array of BOOL values then extracts data from a PROPVARIANT into that array.
' GetLastError: Returns S_OK if successful, or an error value otherwise.
' Note: Use CoTaskMemFree to free the memory allocated for the array.
' =====================================================================================
PRIVATE FUNCTION CPropVar.ToBooleanArrayAlloc (BYVAL pprgf AS WINBOOL PTR PTR) AS ULONG
   CPROPVAR_DP("CPropVar ToBooleanArrayAlloc")
   DIM pcElem AS ULONG
   SetLastError AfxPropVariantToBooleanVectorAlloc(@pvd, pprgf, @pcElem)
   RETURN pcElem
END FUNCTION
' =====================================================================================

' =====================================================================================
' Extracts an array of Int16 values from a PROPVARIANT structure.
' Returns the number of elements extracted from the source PROPVARIANT structure.
' GetLastError: Returns S_OK if successful, or an error value otherwise, including the following:
' TYPE_E_BUFFERTOOSMALL: The source PROPVARIANT contained more than crgn values.
' E_INVALIDARG: The PROPVARIANT was not of the appropriate type.
' To retrieve the number of elements in the array call the GetElementCount method.
' =====================================================================================
PRIVATE FUNCTION CPropVar.ToShortArray (BYVAL prgn AS SHORT PTR, BYVAL crgn AS ULONG) AS ULONG
   CPROPVAR_DP("CPropVar ToShortArray")
   DIM pcElem AS ULONG
   SetLastError AfxPropVariantToInt16Vector(@pvd, prgn, crgn, @pcElem)
   RETURN pcElem
END FUNCTION
' =====================================================================================
' =====================================================================================
' Extracts data from a vector structure into a newly-allocated Int16 array.
' GetLastError: Returns S_OK if successful, or an error value otherwise.
' Note Use CoTaskMemFree to release the memory allocated for the array.
' =====================================================================================
PRIVATE FUNCTION CPropVar.ToShortArrayAlloc (BYVAL pprgn AS SHORT PTR PTR) AS ULONG
   CPROPVAR_DP("CPropVar ToShortArrayAlloc")
   DIM pcElem AS ULONG
   SetLastError AfxPropVariantToInt16VectorAlloc(@pvd, pprgn, @pcElem)
   RETURN pcElem
END FUNCTION
' =====================================================================================

' =====================================================================================
' Extracts an array of UInt16 values from a PROPVARIANT structure.
' Returns the number of elements extracted from the source PROPVARIANT structure.
' GetLastError: Returns S_OK if successful, or an error value otherwise, including the following:
' TYPE_E_BUFFERTOOSMALL: The source PROPVARIANT contained more than crgn values.
' E_INVALIDARG: The PROPVARIANT was not of the appropriate type.
' To retrieve the number of elements in the array call the GetElementCount method.
' =====================================================================================
PRIVATE FUNCTION CPropVar.ToUShortArray (BYVAL prgn AS USHORT PTR, BYVAL crgn AS ULONG) AS ULONG
   CPROPVAR_DP("CPropVar ToShortArray")
   DIM pcElem AS ULONG
   SetLastError AfxPropVariantToUInt16Vector(@pvd, prgn, crgn, @pcElem)
   RETURN pcElem
END FUNCTION
' =====================================================================================
' =====================================================================================
' Extracts data from a vector structure into a newly-allocated Int16 array.
' GetLastError: Returns S_OK if successful, or an error value otherwise.
' Note Use CoTaskMemFree to release the memory allocated for the array.
' =====================================================================================
PRIVATE FUNCTION CPropVar.ToUShortArrayAlloc (BYVAL pprgn AS USHORT PTR PTR) AS ULONG
   CPROPVAR_DP("CPropVar ToUShortArrayAlloc")
   DIM pcElem AS ULONG
   SetLastError AfxPropVariantToUInt16VectorAlloc(@pvd, pprgn, @pcElem)
   RETURN pcElem
END FUNCTION
' =====================================================================================

' =====================================================================================
' Extracts an array of Int32 values from a PROPVARIANT structure.
' Returns the number of elements extracted from the source PROPVARIANT structure.
' GetLastError: Returns S_OK if successful, or an error value otherwise, including the following:
' TYPE_E_BUFFERTOOSMALL: The source PROPVARIANT contained more than crgn values.
' E_INVALIDARG: The PROPVARIANT was not of the appropriate type.
' To retrieve the number of elements in the array call the GetElementCount method.
' =====================================================================================
PRIVATE FUNCTION CPropVar.ToLongArray (BYVAL prgn AS LONG PTR, BYVAL crgn AS ULONG) AS ULONG
   CPROPVAR_DP("CPropVar ToLongArray")
   DIM pcElem AS ULONG
   SetLastError AfxPropVariantToInt32Vector(@pvd, prgn, crgn, @pcElem)
   RETURN pcElem
END FUNCTION
' =====================================================================================
' =====================================================================================
' Extracts data from a vector structure into a newly-allocated Int32 array.
' GetLastError: Returns S_OK if successful, or an error value otherwise.
' Note Use CoTaskMemFree to release the memory allocated for the array.
' =====================================================================================
PRIVATE FUNCTION CPropVar.ToLongArrayAlloc (BYVAL pprgn AS LONG PTR PTR) AS ULONG
   CPROPVAR_DP("CPropVar ToLongArrayAlloc")
   DIM pcElem AS ULONG
   SetLastError AfxPropVariantToInt32VectorAlloc(@pvd, pprgn, @pcElem)
   RETURN pcElem
END FUNCTION
' =====================================================================================

' =====================================================================================
' Extracts an array of UInt32 values from a PROPVARIANT structure.
' Returns the number of elements extracted from the source PROPVARIANT structure.
' GetLastError: Returns S_OK if successful, or an error value otherwise, including the following:
' TYPE_E_BUFFERTOOSMALL: The source PROPVARIANT contained more than crgn values.
' E_INVALIDARG: The PROPVARIANT was not of the appropriate type.
' To retrieve the number of elements in the array call the GetElementCount method.
' =====================================================================================
PRIVATE FUNCTION CPropVar.ToULongArray (BYVAL prgn AS ULONG PTR, BYVAL crgn AS ULONG) AS ULONG
   CPROPVAR_DP("CPropVar ToULongArray")
   DIM pcElem AS ULONG
   SetLastError AfxPropVariantToUInt32Vector(@pvd, prgn, crgn, @pcElem)
   RETURN pcElem
END FUNCTION
' =====================================================================================
' =====================================================================================
' Extracts data from a vector structure into a newly-allocated unsigned Int32 array.
' GetLastError: Returns S_OK if successful, or an error value otherwise.
' Note Use CoTaskMemFree to release the memory allocated for the array.
' =====================================================================================
PRIVATE FUNCTION CPropVar.ToULongArrayAlloc (BYVAL pprgn AS ULONG PTR PTR) AS ULONG
   CPROPVAR_DP("CPropVar ToULongArrayAlloc")
   DIM pcElem AS ULONG
   SetLastError AfxPropVariantToUInt32VectorAlloc(@pvd, pprgn, @pcElem)
   RETURN pcElem
END FUNCTION
' =====================================================================================

' =====================================================================================
' Extracts an array of Int64 values from a PROPVARIANT structure.
' Returns the number of elements extracted from the source PROPVARIANT structure.
' GetLastError: Returns S_OK if successful, or an error value otherwise, including the following:
' TYPE_E_BUFFERTOOSMALL: The source PROPVARIANT contained more than crgn values.
' E_INVALIDARG: The PROPVARIANT was not of the appropriate type.
' To retrieve the number of elements in the array call the GetElementCount method.
' =====================================================================================
PRIVATE FUNCTION CPropVar.ToLongIntArray (BYVAL prgn AS LONGINT PTR, BYVAL crgn AS ULONG) AS ULONG
   CPROPVAR_DP("CPropVar ToULongArray")
   DIM pcElem AS ULONG
   SetLastError AfxPropVariantToInt64Vector(@pvd, prgn, crgn, @pcElem)
  RETURN pcElem
END FUNCTION
' =====================================================================================
' =====================================================================================
' Extracts data from a vector structure into a newly-allocated Int64 array.
' GetLastError: Returns S_OK if successful, or an error value otherwise.
' Note Use CoTaskMemFree to release the memory allocated for the array.
' =====================================================================================
PRIVATE FUNCTION CPropVar.ToLongIntArrayAlloc (BYVAL pprgn AS LONGINT PTR PTR) AS ULONG
   CPROPVAR_DP("CPropVar ToLongIntArrayAlloc")
   DIM pcElem AS ULONG
   SetLastError AfxPropVariantToInt64VectorAlloc(@pvd, pprgn, @pcElem)
   RETURN pcElem
END FUNCTION
' =====================================================================================

' =====================================================================================
' Extracts an array of Int64 values from a PROPVARIANT structure.
' Returns the number of elements extracted from the source PROPVARIANT structure.
' GetLastError: Returns S_OK if successful, or an error value otherwise, including the following:
' TYPE_E_BUFFERTOOSMALL: The source PROPVARIANT contained more than crgn values.
' E_INVALIDARG: The VARIANT was not of the appropriate type.
' To retrieve the number of elements in the array call the GetElementCount method.
' =====================================================================================
PRIVATE FUNCTION CPropVar.ToULongIntArray (BYVAL prgn AS ULONGINT PTR, BYVAL crgn AS ULONG) AS ULONG
   CPROPVAR_DP("CPropVar ToULongIntArray")
   DIM pcElem AS ULONG
   SetLastError AfxPropVariantToUInt64Vector(@pvd, prgn, crgn, @pcElem)
   RETURN pcElem
END FUNCTION
' =====================================================================================
' =====================================================================================
' Extracts data from a vector structure into a newly-allocated unsigned Int64 array.
' GetLastError: Returns S_OK if successful, or an error value otherwise.
' Note Use CoTaskMemFree to release the memory allocated for the array.
' =====================================================================================
PRIVATE FUNCTION CPropVar.ToULongIntArrayAlloc (BYVAL pprgn AS ULONGINT PTR PTR) AS ULONG
   CPROPVAR_DP("CPropVar ToULongIntArrayAlloc")
   DIM pcElem AS ULONG
   SetLastError AfxPropVariantToUInt64VectorAlloc(@pvd, pprgn, @pcElem)
  RETURN pcElem
END FUNCTION
' =====================================================================================

' =====================================================================================
' Extracts an array of Double values from a PROPVARIANT structure.
' Returns the number of elements extracted from the source PROPVARIANT structure.
' GetLastError: Returns S_OK if successful, or an error value otherwise, including the following:
' TYPE_E_BUFFERTOOSMALL: The source PROPVARIANT contained more than crgn values.
' E_INVALIDARG: The VARIANT was not of the appropriate type.
' To retrieve the number of elements in the array call the GetElementCount method.
' =====================================================================================
PRIVATE FUNCTION CPropVar.ToDoubleArray (BYVAL prgn AS DOUBLE PTR, BYVAL crgn AS ULONG) AS ULONG
   CPROPVAR_DP("CPropVar ToDoubleArray")
   DIM pcElem AS ULONG
   SetLastError AfxPropVariantToDoubleVector(@pvd, prgn, crgn, @pcElem)
   RETURN pcElem
END FUNCTION
' =====================================================================================
' =====================================================================================
' Allocates an array of DOUBLE values then extracts data from a VARIANT into that array.
' GetLastError: Returns S_OK if successful, or an error value otherwise.
' Note Use CoTaskMemFree to release the memory allocated for the array.
' =====================================================================================
PRIVATE FUNCTION CPropVar.ToDoubleArrayAlloc (BYVAL pprgn AS DOUBLE PTR PTR) AS ULONG
   CPROPVAR_DP("CPropVar ToDoubleArrayAlloc")
   DIM pcElem AS ULONG
   SetLastError AfxPropVariantToDoubleVectorAlloc(@pvd, pprgn, @pcElem)
   RETURN pcElem
END FUNCTION
' =====================================================================================

' =====================================================================================
' Extracts an array of string values from a PROPVARIANT structure.
' Returns the number of elements extracted from the source PROPVARIANT structure.
' GetLastError: Returns S_OK if successful, or an error value otherwise, including the following:
' TYPE_E_BUFFERTOOSMALL: The source VARIANT contained more than crgn values.
' E_INVALIDARG: The PROPVARIANT was not of the appropriate type.
' To retrieve the number of elements in the array call the GetElementCount method.
' =====================================================================================
PRIVATE FUNCTION CPropVar.ToStringArray (BYVAL prgsz AS PWSTR, BYVAL crgsz AS ULONG) AS ULONG
   CPROPVAR_DP("CPropVar ToStringArray")
   DIM pcElem AS ULONG
   SetLastError AfxPropVariantToStringVector(@pvd, prgsz, crgsz, @pcElem)
   RETURN pcElem
END FUNCTION
' =====================================================================================
' =====================================================================================
' Extracts data from a vector structure into a newly-allocated String array.
' GetLastError: Returns S_OK if successful, or an error value otherwise.
' Note Use CoTaskMemFree to free the memory used by each of the strings and the returned array.
' =====================================================================================
PRIVATE FUNCTION CPropVar.ToStringArrayAlloc (BYVAL pprgsz AS PWSTR PTR) AS ULONG
   CPROPVAR_DP("CPropVar ToStringArrayAlloc")
   DIM pcElem AS ULONG
   SetLastError AfxPropVariantToStringVectorAlloc(@pvd, pprgsz, @pcElem)
   RETURN pcElem
END FUNCTION
' =====================================================================================

' =====================================================================================
' Extracts data from a PROPVARIANT structure into a FILETIME vector.
' Returns the number of elements extracted from the source PROPVARIANT structure.
' GetLastError: Returns S_OK if successful, or an error value otherwise, including the following:
' TYPE_E_BUFFERTOOSMALL: The source VARIANT contained more than crgn values.
' E_INVALIDARG: The PROPVARIANT was not of the appropriate type.
' To retrieve the number of elements in the array call the GetElementCount method.
' =====================================================================================
PRIVATE FUNCTION CPropVar.ToFileTimeArray (BYVAL prgft AS FILETIME PTR, BYVAL crgft AS ULONG) AS HRESULT
   CPROPVAR_DP("CPropVar ToFileTimeArray")
   DIM pcElem AS ULONG
   SetLastError AfxPropVariantToFileTimeVector(@pvd, prgft, crgft, @pcElem)
   RETURN pcElem
END FUNCTION
' =====================================================================================
' =====================================================================================
' Extracts data from a vector structure into a newly-allocated String array.
' GetLastError: Returns S_OK if successful, or an error value otherwise.
' Note Use CoTaskMemFree to free the memory used by each of the strings and the returned array.
' =====================================================================================
PRIVATE FUNCTION CPropVar.ToFileTimeArrayAlloc (BYVAL prgft AS FILETIME PTR PTR) AS ULONG
   CPROPVAR_DP("CPropVar ToFileTimeVectorAlloc")
   DIM pcElem AS ULONG
   SetLastError AfxPropVariantToFileTimeVectorAlloc(@pvd, prgft, @pcElem)
   RETURN pcElem
END FUNCTION
' =====================================================================================

' ========================================================================================
' Assigns another CPropVar.
' ========================================================================================
PRIVATE OPERATOR CPropVar.Let (BYREF cpv AS CPropVar)
   CPROPVAR_DP("CPropVar LET = CPropVar")
'   PropVariantClear(@pvd) : PropVariantCopy(@pvd, @cpv)   ' // removed @ operator
   PropVariantClear(@pvd) : PropVariantCopy(@pvd, *cpv)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Assigns a VARIANT.
' ========================================================================================
PRIVATE OPERATOR CPropVar.Let (BYREF pv AS PROPVARIANT)
   CPROPVAR_DP("CPropVar LET = PROPVARIANT")
   PropVariantClear(@pvd) : PropVariantCopy(@pvd, @pv)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Assigns a WSTRING.
' ========================================================================================
PRIVATE OPERATOR CPropVar.Let (BYREF wszStr AS WSTRING)
   CPROPVAR_DP("CPropVar LET = WSTRING")
   PropVariantClear(@pvd) : pvd.vt = VT_LPWSTR : SHStrDupW(@wszStr, @pvd.pwszVal)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Assigns a CWSTR.
' ========================================================================================
PRIVATE OPERATOR CPropVar.Let (BYVAL pPropVar AS PROPVARIANT PTR)
   CPROPVAR_DP("CPropVar LET = VARIANT PTR")
   PropVariantClear(@pvd) : PropVariantCopy(@pvd, pPropVar)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Assigns a CWSTR.
' ========================================================================================
PRIVATE OPERATOR CPropVar.Let (BYREF cws AS CWSTR)
   CPROPVAR_DP("CPropVar LET = CWSTR")
   PropVariantClear(@pvd) : pvd.vt = VT_LPWSTR : SHStrDupW(cws, @pvd.pwszVal)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Assigns a CBSTR.
' ========================================================================================
PRIVATE OPERATOR CPropVar.Let (BYREF cbs AS CBSTR)
   CPROPVAR_DP("CPropVar LET = CBSTR")
   PropVariantClear(@pvd) : pvd.vt = VT_LPWSTR : SHStrDupW(cbs, @pvd.pwszVal)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Assigns a CURRENCY structure.
' ========================================================================================
PRIVATE OPERATOR CPropVar.Let (BYVAL cy AS CURRENCY)
   CPROPVAR_DP("CPropVar LET - CURRENCY")
   PropVariantClear(@pvd) : pvd.vt = VT_CY : pvd.cyVal = cy
END OPERATOR
' ========================================================================================

' ========================================================================================
' Assigns a CURRENCY structure.
' ========================================================================================
PRIVATE OPERATOR CPropVar.Let (BYVAL dec AS DECIMAL)
   CPROPVAR_DP("CPropVar LET - DECIMAL")
   PropVariantClear(@pvd) : pvd.vt = VT_DECIMAL : pvd.decVal = dec
END OPERATOR
' ========================================================================================

' ========================================================================================
' Assigns a Dispatch pointer
' ========================================================================================
PRIVATE OPERATOR CPropVar.Let (BYVAL pdisp AS IDispatch PTR)
   CPROPVAR_DP("CPropVar LET = IDispatch PTR")
   PropVariantClear(@pvd) : pvd.vt = VT_DISPATCH : pvd.pdispVal = pdisp : IDispatch_AddRef(pdisp)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Assigns an IUnknown pointer
' ========================================================================================
PRIVATE OPERATOR CPropVar.Let (BYVAL punk AS IUnknown PTR)
   CPROPVAR_DP("CPropVar LET = IUnknown PTR")
   PropVariantClear(@pvd) : pvd.vt = VT_UNKNOWN : pvd.punkVal = punk : IUnknown_AddRef(punk)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Assigns an IStream pointer
' ========================================================================================
PRIVATE OPERATOR CPropVar.Let (BYVAL pStream AS IStream PTR)
   CPROPVAR_DP("CPropVar LET = IStream PTR")
   PropVariantClear(@pvd) : pvd.vt = VT_STREAM : pvd.pStream = pStream : IUnknown_AddRef(pStream)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Assigns an IStorage pointer
' ========================================================================================
PRIVATE OPERATOR CPropVar.Let (BYVAL pStorage AS IStorage PTR)
   CPROPVAR_DP("CPropVar LET = IStorage PTR")
   PropVariantClear(@pvd) : pvd.vt = VT_STORAGE : pvd.pStorage = pStorage : IUnknown_AddRef(pStorage)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Assigns an IStorage pointer
' ========================================================================================
PRIVATE OPERATOR CPropVar.Let (BYVAL pVersionedStream AS VERSIONEDSTREAM PTR)
   CPROPVAR_DP("CPropVar LET = VERSIONEDSTREAM PTR")
   PropVariantClear(@pvd) : pvd.vt = VT_VERSIONED_STREAM : pvd.pVersionedStream = pVersionedStream
END OPERATOR
' ========================================================================================

' ========================================================================================
' Assigns a BLOB structure
' ========================================================================================
PRIVATE OPERATOR CPropVar.Let (BYREF _blob AS BLOB)
   CPROPVAR_DP("CPropVar CONSTRUCTOR - BLOB")
   PropVariantClear(@pvd) : pvd.vt = VT_BLOB : pvd.blob = _blob
END OPERATOR
' ========================================================================================

' ========================================================================================
' Initializes a CPropVar from a BSTRBLOB structure
' ========================================================================================
PRIVATE OPERATOR CPropVar.Let (BYREF _bstrBlob AS BSTRBLOB)
   CPROPVAR_DP("CPropVar CONSTRUCTOR - BSTRBLOB")
   PropVariantClear(@pvd) : pvd.vt = VT_BSTR_BLOB : pvd.bstrblobVal = _bstrBlob
END OPERATOR
' ========================================================================================

' ========================================================================================
' Initializes a CPropVar from a CLIPDATA structure
' ========================================================================================
PRIVATE OPERATOR CPropVar.Let (BYVAL pclipData AS CLIPDATA PTR)
   CPROPVAR_DP("CPropVar CONSTRUCTOR - CLIPDATA PTR")
   PropVariantClear(@pvd) : pvd.vt = VT_CF : pvd.pclipdata = pclipData
END OPERATOR
' ========================================================================================

' ========================================================================================
' Assigns structures
' ========================================================================================
PRIVATE OPERATOR CPropVar.Let (BYREF _cac AS CAC)
   CPROPVAR_DP("CPropVar CONSTRUCTOR - CAC")
   PropVariantClear(@pvd) : pvd.vt = VT_VECTOR OR VT_I1 : pvd.cac = _cac
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CPropVar.Let (BYREF _caub AS CAUB)
   CPROPVAR_DP("CPropVar CONSTRUCTOR - CAUB")
   PropVariantClear(@pvd) : pvd.vt = VT_VECTOR OR VT_UI1 : pvd.caub = _caub
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CPropVar.Let (BYREF _cai AS CAI)
   CPROPVAR_DP("CPropVar CONSTRUCTOR - CAI")
   PropVariantClear(@pvd) : pvd.vt = VT_VECTOR OR VT_I2 : pvd.cai = _cai
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CPropVar.Let (BYREF _caui AS CAUI)
   CPROPVAR_DP("CPropVar CONSTRUCTOR - CAUI")
   PropVariantClear(@pvd) : pvd.vt = VT_VECTOR OR VT_UI2 : pvd.caui = _caui
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CPropVar.Let (BYREF _cal AS CAL)
   CPROPVAR_DP("CPropVar CONSTRUCTOR - CAL")
   PropVariantClear(@pvd) : pvd.vt = VT_VECTOR OR VT_I4 : pvd.cal = _cal
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CPropVar.Let (BYREF _caul AS CAUL)
   CPROPVAR_DP("CPropVar CONSTRUCTOR - CAUL")
   PropVariantClear(@pvd) : pvd.vt = VT_VECTOR OR VT_UI4 : pvd.caul = _caul
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CPropVar.Let (BYREF _cah AS CAH)
   CPROPVAR_DP("CPropVar CONSTRUCTOR - CAH")
   PropVariantClear(@pvd) : pvd.vt = VT_VECTOR OR VT_I8 : pvd.cah = _cah
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CPropVar.Let (BYREF _caflt AS CAFLT)
   CPROPVAR_DP("CPropVar CONSTRUCTOR - CAFLT")
   PropVariantClear(@pvd) : pvd.vt = VT_VECTOR OR VT_R4 : pvd.caflt = _caflt
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CPropVar.Let (BYREF _cadbl AS CADBL)
   CPROPVAR_DP("CPropVar CONSTRUCTOR - CADBL")
   PropVariantClear(@pvd) : pvd.vt = VT_VECTOR OR VT_R8 : pvd.cadbl = _cadbl
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CPropVar.Let (BYREF _cabool AS CABOOL)
   CPROPVAR_DP("CPropVar CONSTRUCTOR - CABOOL")
   PropVariantClear(@pvd) : pvd.vt = VT_VECTOR OR VT_BOOL : pvd.cabool = _cabool
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CPropVar.Let (BYREF _cascode AS CASCODE)
   CPROPVAR_DP("CPropVar CONSTRUCTOR - CASCODE")
   PropVariantClear(@pvd) : pvd.vt = VT_VECTOR OR VT_ERROR : pvd.cascode = _cascode
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CPropVar.Let (BYREF _cacy AS CACY)
   CPROPVAR_DP("CPropVar CONSTRUCTOR - CACY")
   PropVariantClear(@pvd) : pvd.vt = VT_VECTOR OR VT_CY : pvd.cacy = _cacy
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CPropVar.Let (BYREF _cadate AS CADATE)
   CPROPVAR_DP("CPropVar CONSTRUCTOR - CADATE")
   PropVariantClear(@pvd) : pvd.vt = VT_VECTOR OR VT_DATE : pvd.cadate = _cadate
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CPropVar.Let (BYREF _cafiletime AS CAFILETIME)
   CPROPVAR_DP("CPropVar CONSTRUCTOR - CAFILETIME")
   PropVariantClear(@pvd) : pvd.vt = VT_VECTOR OR VT_FILETIME : pvd.cafiletime = _cafiletime
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CPropVar.Let (BYREF _caclsid AS CACLSID)
   CPROPVAR_DP("CPropVar CONSTRUCTOR - CACLSID")
   PropVariantClear(@pvd) : pvd.vt = VT_VECTOR OR VT_CLSID : pvd.cauuid = _caclsid
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CPropVar.Let (BYREF _caclipdata AS CACLIPDATA)
   CPROPVAR_DP("CPropVar CONSTRUCTOR - CACLIPDATA")
   PropVariantClear(@pvd) : pvd.vt = VT_VECTOR OR VT_CF : pvd.caclipdata = _caclipdata
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CPropVar.Let (BYREF _cabstr AS CABSTR)
   CPROPVAR_DP("CPropVar CONSTRUCTOR - CABSTR")
   PropVariantClear(@pvd) : pvd.vt = VT_VECTOR OR VT_BSTR : pvd.cabstr = _cabstr
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CPropVar.Let (BYREF _cabstrblob AS CABSTRBLOB)
   CPROPVAR_DP("CPropVar CONSTRUCTOR - CABSTRBLOB")
   PropVariantClear(@pvd) : pvd.vt = VT_VECTOR OR VT_BSTR_BLOB : pvd.cabstrblob = _cabstrblob
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CPropVar.Let (BYREF _calpstr AS CALPSTR)
   CPROPVAR_DP("CPropVar CONSTRUCTOR - CALPSTR")
   PropVariantClear(@pvd) : pvd.vt = VT_VECTOR OR VT_LPSTR : pvd.calpstr = _calpstr
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CPropVar.Let (BYREF _calpwstr AS CALPWSTR)
   CPROPVAR_DP("CPropVar CONSTRUCTOR - CALPWSTR")
   PropVariantClear(@pvd) : pvd.vt = VT_VECTOR OR VT_LPWSTR : pvd.calpwstr = _calpwstr
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CPropVar.Let (BYREF _capropvar AS CAPROPVARIANT)
   CPROPVAR_DP("CPropVar CONSTRUCTOR - CAPROPVAR")
   PropVariantClear(@pvd) : pvd.vt = VT_VECTOR OR VT_VARIANT : pvd.capropvar = _capropvar
END OPERATOR
' ========================================================================================

' ========================================================================================
' Assigns another CPropVar.
' ========================================================================================
PRIVATE FUNCTION CPropVar.Put (BYREF cpv AS CPropVar) AS HRESULT
   CPROPVAR_DP("CPropVar Put = CPropVar")
'   PropVariantClear(@pvd) : PropVariantCopy(@pvd, @cpv)   ' // removed @ operator
   PropVariantClear(@pvd)
   RETURN PropVariantCopy(@pvd, *cpv)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Assigns a PROPVARIANT.
' ========================================================================================
PRIVATE FUNCTION CPropVar.Put (BYREF pv AS PROPVARIANT) AS HRESULT
   CPROPVAR_DP("CPropVar Put = PROPVARIANT")
   PropVariantClear(@pvd)
   RETURN PropVariantCopy(@pvd, @pv)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Assigns a WSTRING.
' ========================================================================================
PRIVATE SUB CPropVar.Put (BYREF wszStr AS WSTRING)
   CPROPVAR_DP("CPropVar Put = WSTRING")
   PropVariantClear(@pvd) : pvd.vt = VT_LPWSTR : SHStrDupW(@wszStr, @pvd.pwszVal)
END SUB
' ========================================================================================

' ========================================================================================
' Assigns a PROPVARIANT.
' ========================================================================================
PRIVATE FUNCTION CPropVar.Put (BYVAL pPropVar AS PROPVARIANT PTR) As HRESULT
   CPROPVAR_DP("CPropVar Put = PROPVARIANT PTR")
   PropVariantClear(@pvd)
   RETURN PropVariantCopy(@pvd, pPropVar)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Assigns a CWSTR.
' ========================================================================================
PRIVATE SUB CPropVar.Put (BYREF cws AS CWSTR)
   CPROPVAR_DP("CPropVar Put = CWSTR")
   PropVariantClear(@pvd) : pvd.vt = VT_LPWSTR : SHStrDupW(cws, @pvd.pwszVal)
END SUB
' ========================================================================================

' ========================================================================================
' Assigns a CBSTR.
' ========================================================================================
PRIVATE SUB CPropVar.Put (BYREF cbs AS CBSTR)
   CPROPVAR_DP("CPropVar Put = CBSTR")
   PropVariantClear(@pvd) : pvd.vt = VT_LPWSTR : SHStrDupW(cbs, @pvd.pwszVal)
END SUB
' ========================================================================================

' ========================================================================================
' Assigns a Dispatch pointer
' ========================================================================================
PRIVATE FUNCTION CPropVar.Put (BYVAL pdisp AS IDispatch PTR, BYVAL fAddRef AS BOOLEAN = FALSE) AS HRESULT
   CPROPVAR_DP("CPropVar Put = IDispatch PTR")
   PropVariantClear(@pvd)
   IF pDisp = NULL THEN RETURN E_INVALIDARG
   pvd.vt = VT_DISPATCH
   pvd.pdispVal = pdisp
   IF fAddRef THEN IDispatch_AddRef(pdisp)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Assigns an IUnknown pointer
' ========================================================================================
PRIVATE FUNCTION CPropVar.Put (BYVAL punk AS IUnknown PTR, BYVAL fAddRef AS BOOLEAN = FALSE) AS HRESULT
   CPROPVAR_DP("CPropVar Put = IUnknown PTR")
   PropVariantClear(@pvd)
   IF punk = NULL THEN RETURN E_INVALIDARG
   pvd.vt = VT_UNKNOWN
   pvd.punkVal = punk
   IF fAddRef THEN IUnknown_AddRef(punk)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Assigns an integer value to a CPropVar
' If _vType is wrong, default to VT_I8 (LongInt).
' ========================================================================================
PRIVATE SUB CPropVar.Put (BYVAL _value AS LONGINT, BYVAL _vType AS WORD)
   CPROPVAR_DP("CPropVar Put - vType")
   PropVariantClear(@pvd)
   IF (_vType AND VT_ARRAY) = VT_ARRAY THEN
      ' // NULL array : Suitable for parameters that, instead of being optional,
      ' // require a variant with a null safearray of a given type.
      pvd.vt = _vType
   ELSE
      ' // If _value exceeds the limits of max long or max ulong, promote to ulongint
      IF SGN(_value) = -1 AND _value > LONG_MAX THEN
         _vType = VT_I8
      ELSEIF _value > ULONG_MAX THEN
         _vType = VT_I8
      END IF
      SELECT CASE _vType
         CASE VT_BOOL  : pvd.vt = VT_BOOL : pvd.boolVal = IIF(_value = 0, 0, -1)
         CASE VT_I1    : pvd.vt = VT_I1   : pvd.cVal = CBYTE(_value)
         CASE VT_UI1   : pvd.vt = VT_UI1  : pvd.bVal = CUBYTE(_value)
         CASE VT_I2    : pvd.vt = VT_I2   : pvd.iVal = CSHORT(_value)
         CASE VT_UI2   : pvd.vt = VT_UI2  : pvd.uiVal = CUSHORT(_value)
         CASE VT_INT   : pvd.vt = VT_INT  : pvd.intVal = CLNG(_value)
         CASE VT_UINT  : pvd.vt = VT_UINT : pvd.uintVal = CULNG(_value)
         CASE VT_I4    : pvd.vt = VT_I4   : pvd.lVal = CLNG(_value)
         CASE VT_UI4   : pvd.vt = VT_UI4  : pvd.ulVal = CULNG(_value)
         CASE VT_I8    : pvd.vt = VT_I8   : pvd.hVal.QuadPart = CLNGINT(_value)
         CASE VT_UI8   : pvd.vt = VT_UI8  : pvd.uhVal.QuadPart = CULNGINT(_value)
         ' // If you need to assign an ULongInt value greater than a LongInt, use PutUlongInt.
         CASE VT_NULL  : pvd.vt = VT_NULL
        CASE ELSE      : pvd.vt = VT_I8   : pvd.hVal.QuadPart = CLNGINT(_value)
      END SELECT
   END IF
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB CPropVar.Put (BYVAL _value AS LONGINT, BYREF strType AS STRING)
   CPROPVAR_DP("CPropVar Put - strType")
   PropVariantClear(@pvd)
   DIM vt AS WORD
   SELECT CASE UCASE(strType)
      CASE "BOOL"     : pvd.vt = VT_BOOL : pvd.boolVal = IIF(_value = 0, 0, -1)
      CASE "BOOLEAN"  : pvd.vt = VT_BOOL : pvd.boolVal = IIF(_value = 0, 0, -1)
      CASE "BYTE"     : pvd.vt = VT_I1   : pvd.cVal = CBYTE(_value)
      CASE "UBYTE"    : pvd.vt = VT_UI1  : pvd.bVal = CUBYTE(_value)
      CASE "SHORT"    : pvd.vt = VT_I2   : pvd.iVal = CSHORT(_value)
      CASE "USHORT"   : pvd.vt = VT_UI2  : pvd.uiVal = CUSHORT(_value)
      CASE "INT"      : pvd.vt = VT_INT  : pvd.intVal = CLNG(_value)
      CASE "UINT"     : pvd.vt = VT_UINT : pvd.uintVal = CULNG(_value)
      CASE "LONG"     : pvd.vt = VT_I4   : pvd.lVal = CLNG(_value)
      CASE "ULONG"    : pvd.vt = VT_UI4  : pvd.ulVal = CULNG(_value)
      CASE "LONGINT"  : pvd.vt = VT_I8   : pvd.hVal.QuadPart = CLNGINT(_value)
      CASE "ULONGINT" : pvd.vt = VT_UI8  : pvd.uhVal.QuadPart = CULNGINT(_value)
      ' // If you need to assign an ULongInt value greater than a LongInt, use PutUlongInt.
      CASE "NULL"     : pvd.vt = VT_NULL
      CASE ELSE       : pvd.vt = VT_I8   : pvd.hVal.QuadPart = CLNGINT(_value)
   END SELECT
END SUB
' ========================================================================================

' ========================================================================================
' Assigns a float value to a CPropVar
' If _vType is wrong, default to VT_R8 (double).
' ========================================================================================
PRIVATE SUB CPropVar.Put (BYVAL _value AS DOUBLE, BYVAL _vType AS WORD = VT_R8)
   CPROPVAR_DP("CPropVar Put - vType")
   PropVariantClear(@pvd)
   IF _value > FLT_MAX OR _value < FLT_MIN THEN _vType = VT_R8
   SELECT CASE _vType
      CASE VT_R4 : pvd.vt = VT_R4 : pvd.fltVal = CSNG(_value)
      CASE VT_R8 : pvd.vt = VT_R8 : pvd.dblVal = CDBL(_value)
      CASE VT_BOOL, VT_I1, VT_UI1, VT_I2, VT_UI2, VT_INT, VT_UINT, VT_I4, VT_UI4, VT_I8, VT_UI8, VT_NULL
         this.Put CLNGINT(_value), _vType
      CASE ELSE
         pvd.vt = VT_R8 : pvd.dblVal = CDBL(_value)
   END SELECT
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB CPropVar.Put (BYVAL _value AS DOUBLE, BYREF strType AS STRING)
   CPROPVAR_DP("CPropVar Put - strType")
   PropVariantClear(@pvd)
   SELECT CASE UCASE(strType)
      CASE "SINGLE" : pvd.vt = VT_R4   : pvd.fltVal = CSNG(_value)
      CASE "FLOAT"  : pvd.vt = VT_R4   : pvd.fltVal = CSNG(_value)
      CASE "DOUBLE" : pvd.vt = VT_R8   : pvd.dblVal = CDBL(_value)
      CASE ELSE     : pvd.vt = VT_R8   : pvd.dblVal = CDBL(_value)
   END SELECT
END SUB
' ========================================================================================

' ========================================================================================
' Assigns a value by reference (a pointer to a variable).
' ========================================================================================
PRIVATE FUNCTION CPropVar.PutRef (BYVAL _pvar AS ANY PTR, BYVAL _vType AS WORD) AS HRESULT
   CPROPVAR_DP("CPropVar PutRef - vType")
   PropVariantClear(@pvd)
   IF _pvar = NULL THEN RETURN E_INVALIDARG
   SELECT CASE _vType
      CASE VT_BOOL      : pvd.vt = VT_BOOL OR VT_BYREF      : pvd.pboolVal = _pvar
      CASE VT_I1        : pvd.vt = VT_I1 OR VT_BYREF        : pvd.pcVal = _pvar
      CASE VT_UI1       : pvd.vt = VT_UI1 OR VT_BYREF       : pvd.pbVal = _pvar
      CASE VT_I2        : pvd.vt = VT_I2 OR VT_BYREF        : pvd.piVal = _pvar
      CASE VT_UI2       : pvd.vt = VT_UI2 OR VT_BYREF       : pvd.puiVal = _pvar
      CASE VT_INT       : pvd.vt = VT_INT OR VT_BYREF       : pvd.pintVal = _pvar
      CASE VT_UINT      : pvd.vt = VT_UINT OR VT_BYREF      : pvd.puintVal = _pvar
      CASE VT_I4        : pvd.vt = VT_I4 OR VT_BYREF        : pvd.plVal = _pvar
      CASE VT_UI4       : pvd.vt = VT_UI4 OR VT_BYREF       : pvd.pulVal = _pvar
'      CASE VT_I8        : pvd.vt = VT_I8 OR VT_BYREF        : pvd.pllVal = _pvar
'      CASE VT_UI8       : pvd.vt = VT_UI8 OR VT_BYREF       : pvd.pullVal = _pvar
      CASE VT_R4        : pvd.vt = VT_R4 OR VT_BYREF        : pvd.pfltVal = _pvar
      CASE VT_R8        : pvd.vt = VT_R8 OR VT_BYREF        : pvd.pdblVal = _pvar
      CASE VT_LPWSTR    : pvd.vt = VT_LPWSTR OR VT_BYREF    : pvd.pwszVal = _pvar
      CASE VT_BSTR      : pvd.vt = VT_BSTR OR VT_BYREF      : pvd.pbstrVal = _pvar
      CASE VT_UNKNOWN   : pvd.vt = VT_UNKNOWN OR VT_BYREF   : pvd.ppunkVal = _pvar
      CASE VT_DISPATCH  : pvd.vt = VT_DISPATCH OR VT_BYREF  : pvd.ppdispVal = _pvar
      CASE VT_DECIMAL   : pvd.vt = VT_DECIMAL OR VT_BYREF   : pvd.pdecVal = _pvar
      CASE VT_CY        : pvd.vt = VT_CY OR VT_BYREF        : pvd.pcyVal = _pvar
      CASE VT_DATE      : pvd.vt = VT_DATE OR VT_BYREF      : pvd.pdate = _pvar
      CASE VT_VARIANT   : pvd.vt = VT_VARIANT OR VT_BYREF   : pvd.pvarVal = _pvar
      CASE VT_SAFEARRAY : pvd.vt = VT_SAFEARRAY OR VT_BYREF : pvd.pvarVal = _pvar
      CASE VT_ERROR     : pvd.vt = VT_ERROR OR VT_BYREF     : pvd.pparray = _pvar
      CASE ELSE         : RETURN E_INVALIDARG
   END SELECT
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CPropVar.PutRef (BYVAL _pvar AS ANY PTR, BYREF strType AS STRING) AS HRESULT
   CPROPVAR_DP("CPropVar PutRef - strType")
   PropVariantClear(@pvd)
   IF _pvar = NULL THEN RETURN E_INVALIDARG
   SELECT CASE UCASE(strType)
      CASE "BOOL"      : pvd.vt = VT_BOOL OR VT_BYREF      : pvd.pboolVal = _pvar
      CASE "BOOLEAN"   : pvd.vt = VT_BOOL OR VT_BYREF      : pvd.pboolVal = _pvar
      CASE "BYTE"      : pvd.vt = VT_I1 OR VT_BYREF        : pvd.pcVal = _pvar
      CASE "UBYTE"     : pvd.vt = VT_UI1 OR VT_BYREF       : pvd.pbVal = _pvar
      CASE "SHORT"     : pvd.vt = VT_I2 OR VT_BYREF        : pvd.piVal = _pvar
      CASE "USHORT"    : pvd.vt = VT_UI2 OR VT_BYREF       : pvd.puiVal = _pvar
      CASE "INT"       : pvd.vt = VT_INT OR VT_BYREF       : pvd.pintVal = _pvar
      CASE "UINT"      : pvd.vt = VT_UINT OR VT_BYREF      : pvd.puintVal = _pvar
      CASE "LONG"      : pvd.vt = VT_I4 OR VT_BYREF        : pvd.plVal = _pvar
      CASE "ULONG"     : pvd.vt = VT_UI4 OR VT_BYREF       : pvd.pulVal = _pvar
'      CASE "LONGINT"   : pvd.vt = VT_I8 OR VT_BYREF        : pvd.pllVal = _pvar
'      CASE "ULONGINT"  : pvd.vt = VT_UI8 OR VT_BYREF       : pvd.pullVal = _pvar
      CASE "SINGLE"    : pvd.vt = VT_R4 OR VT_BYREF        : pvd.pfltVal = _pvar
      CASE "FLOAT"     : pvd.vt = VT_R4 OR VT_BYREF        : pvd.pfltVal = _pvar
      CASE "DOUBLE"    : pvd.vt = VT_R8 OR VT_BYREF        : pvd.pdblVal = _pvar
      CASE "WSTRING"   : pvd.vt = VT_LPWSTR OR VT_BYREF    : pvd.pwszVal = _pvar
      CASE "BSTR"      : pvd.vt = VT_BSTR OR VT_BYREF      : pvd.pbstrVal = _pvar
      CASE "UNKNOWN"   : pvd.vt = VT_UNKNOWN OR VT_BYREF   : pvd.ppunkVal = _pvar
      CASE "DISPATCH"  : pvd.vt = VT_DISPATCH OR VT_BYREF  : pvd.ppdispVal = _pvar
      CASE "DECIMAL"   : pvd.vt = VT_DECIMAL OR VT_BYREF   : pvd.pdecVal = _pvar
      CASE "CY"        : pvd.vt = VT_CY OR VT_BYREF        : pvd.pcyVal = _pvar
      CASE "CURRENCY"  : pvd.vt = VT_CY OR VT_BYREF        : pvd.pcyVal = _pvar
      CASE "DATE"      : pvd.vt = VT_DATE OR VT_BYREF      : pvd.pdate = _pvar
      CASE "VARIANT"   : pvd.vt = VT_VARIANT OR VT_BYREF   : pvd.pvarVal = _pvar
      CASE "SAFEARRAY" : pvd.vt = VT_SAFEARRAY OR VT_BYREF : pvd.pvarVal = _pvar
      CASE "ERROR"     : pvd.vt = VT_ERROR OR VT_BYREF     : pvd.pparray = _pvar
      CASE ELSE        : RETURN E_INVALIDARG
   END SELECT
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Assigns a null value.
' ========================================================================================
PRIVATE SUB CPropVar.PutNull
   CPROPVAR_DP("CPropVar PutNull")
   PropVariantClear(@pvd)
   pvd.vt = VT_NULL
END SUB
' ========================================================================================

' ========================================================================================
' * Assigns a boolean value.
' ========================================================================================
PRIVATE SUB CPropVar.PutBool (BYVAL _value AS BOOL)
   CPROPVAR_DP("CPropVar PutBool")
   PropVariantClear(@pvd)
   pvd.vt = VT_BOOL : pvd.boolVal = IIF(_value = 0, 0, -1)
END SUB
' ========================================================================================

' ========================================================================================
' * Assigns a boolean value.
' ========================================================================================
PRIVATE SUB CPropVar.PutBoolean (BYVAL _value AS BOOLEAN)
   CPROPVAR_DP("CPropVar PutBoolean")
   PropVariantClear(@pvd)
   pvd.vt = VT_BOOL : pvd.boolVal = IIF(_value = 0, 0, -1)
END SUB
' ========================================================================================

' ========================================================================================
' * Assigns a byte value.
' ========================================================================================
PRIVATE SUB CPropVar.PutByte (BYVAL _value AS BYTE)
   CPROPVAR_DP("CPropVar PutByte")
   PropVariantClear(@pvd)
   pvd.vt = VT_I1 : pvd.cVal = CBYTE(_value)
END SUB
' ========================================================================================

' ========================================================================================
' * Assigns an ubyte value.
' ========================================================================================
PRIVATE SUB CPropVar.PutUByte (BYVAL _value AS UBYTE)
   CPROPVAR_DP("CPropVar PutUByte")
   PropVariantClear(@pvd)
   pvd.vt = VT_UI1 : pvd.bVal = CUBYTE(_value)
END SUB
' ========================================================================================

' ========================================================================================
' * Assigns a short value.
' ========================================================================================
PRIVATE SUB CPropVar.PutShort (BYVAL _value AS SHORT)
   CPROPVAR_DP("CPropVar PutShort")
   PropVariantClear(@pvd)
   pvd.vt = VT_I2 : pvd.iVal = CSHORT(_value)
END SUB
' ========================================================================================

' ========================================================================================
' * Assigns an ushort value.
' ========================================================================================
PRIVATE SUB CPropVar.PutUShort (BYVAL _value AS USHORT)
   CPROPVAR_DP("CPropVar PutUShort")
   PropVariantClear(@pvd)
   pvd.vt = VT_UI2 : pvd.uiVal = CUSHORT(_value)
END SUB
' ========================================================================================

' ========================================================================================
' * Assigns an INT_ value.
' ========================================================================================
PRIVATE SUB CPropVar.PutInt (BYVAL _value AS INT_)
   CPROPVAR_DP("CPropVar PutInt")
   PropVariantClear(@pvd)
   pvd.vt = VT_INT : pvd.intVal = CLNG(_value)
END SUB
' ========================================================================================

' ========================================================================================
' * Assigns an UINT value.
' ========================================================================================
PRIVATE SUB CPropVar.PutUInt (BYVAL _value AS UINT)
   CPROPVAR_DP("CPropVar PutUInt")
   PropVariantClear(@pvd)
   pvd.vt = VT_UINT : pvd.uintVal = CULNG(_value)
END SUB
' ========================================================================================

' ========================================================================================
' * Assigns a LONG value.
' ========================================================================================
PRIVATE SUB CPropVar.PutLong (BYVAL _value AS LONG)
   CPROPVAR_DP("CPropVar PutLong")
   PropVariantClear(@pvd)
   pvd.vt = VT_I4 : pvd.lVal = CLNG(_value)
END SUB
' ========================================================================================

' ========================================================================================
' * Assigns an ULONG value.
' ========================================================================================
PRIVATE SUB CPropVar.PutULong (BYVAL _value AS ULONG)
   CPROPVAR_DP("CPropVar PutLong")
   PropVariantClear(@pvd)
   pvd.vt = VT_UI4 : pvd.ulVal = CULNG(_value)
END SUB
' ========================================================================================

' ========================================================================================
' * Assigns a LONGINT value.
' ========================================================================================
PRIVATE SUB CPropVar.PutLongInt (BYVAL _value AS LONGINT)
   CPROPVAR_DP("CPropVar PutLongInt")
   PropVariantClear(@pvd)
   pvd.vt = VT_I8 : pvd.hVal.QuadPart = CLNGINT(_value)
END SUB
' ========================================================================================

' ========================================================================================
' * Assigns an ULONGINT value.
' ========================================================================================
PRIVATE SUB CPropVar.PutULongInt (BYVAL _value AS ULONGINT)
   CPROPVAR_DP("CPropVar PutULongInt")
   PropVariantClear(@pvd)
   pvd.vt = VT_UI8 : pvd.uhVal.QuadPart = _value
END SUB
' ========================================================================================

' ========================================================================================
' * Assigns a SINGLE value.
' ========================================================================================
PRIVATE SUB CPropVar.PutSingle (BYVAL _value AS SINGLE)
   CPROPVAR_DP("CPropVar PutSingle")
   PropVariantClear(@pvd)
   pvd.vt = VT_R4 : pvd.fltVal = CSNG(_value)
END SUB
' ========================================================================================

' ========================================================================================
' * Assigns a FLOAT value.
' ========================================================================================
PRIVATE SUB CPropVar.PutFloat (BYVAL _value AS SINGLE)
   CPROPVAR_DP("CPropVar PutFloat")
   PropVariantClear(@pvd)
   pvd.vt = VT_R4 : pvd.fltVal = CSNG(_value)
END SUB
' ========================================================================================

' ========================================================================================
' * Assigns a DOUBLE value.
' ========================================================================================
PRIVATE SUB CPropVar.PutDouble (BYVAL _value AS DOUBLE)
   CPROPVAR_DP("CPropVar PutDouble")
   PropVariantClear(@pvd)
   pvd.vt = VT_R8 : pvd.dblVal = CDBL(_value)
END SUB
' ========================================================================================

' ========================================================================================
' Initializes a CPropVar from a buffer. Creates a VT_ARRAY OR VT_UI1 variant.
' ========================================================================================
PRIVATE FUNCTION CPropVar.PutBuffer (BYVAL pv AS ANY PTR, BYVAL cb AS UINT) AS HRESULT
   CPROPVAR_DP("CPropVar PutBuffer")
   PropVariantClear(@pvd)
   RETURN AfxPropVariantFromBuffer(pv, cb, @pvd)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Initializes a CPropVar from the contents of an UTF8 string.
' ========================================================================================
PRIVATE FUNCTION CPropVar.PutUtf8 (BYREF strUtf8 AS STRING) AS HRESULT
   CPROPVAR_DP("CPropVar PutUtf8")
   PropVariantClear(@pvd)
   pvd.vt = VT_LPWSTR
   SHStrDupW(AfxUcode(strUtf8, CP_UTF8), @pvd.pwszVal)
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Initializes a CPropVar with the contents of a safe array. If fAttach = FALSE, the safe
' array is copied; if fAttach = TRUE, the safe array is attached.
' ========================================================================================
PRIVATE FUNCTION CPropVar.PutSafeArray (BYVAL parray AS SAFEARRAY PTR, BYVAL fAttach AS BOOLEAN = FALSE) AS HRESULT
   CPROPVAR_DP("CPropVar PutSafeArray")
   PropVariantClear(@pvd)
   IF parray = NULL THEN RETURN E_INVALIDARG
   DIM vt AS VARTYPE
   DIM hr AS HRESULT = SafeArrayGetVartype(parray, @vt)
   IF hr <> S_OK THEN RETURN hr
   IF fAttach = FALSE THEN
      DIM pcopy AS SAFEARRAY PTR
      hr = SafeArrayCopy(parray, @pcopy)
      IF hr <> S_OK THEN RETURN hr
      pvd.vt = vt OR VT_ARRAY
      pvd.parray = pcopy
   ELSE
      pvd.vt = vt OR VT_ARRAY
      pvd.parray = parray
   END IF
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Initializes a CPropVar from a a string resource imbedded in an executable file.
' ========================================================================================
PRIVATE FUNCTION CPropVar.PutResource (BYVAL hinst AS HINSTANCE, BYVAL id AS UINT) AS HRESULT
   CPROPVAR_DP("CPropVar PutResource")
   PropVariantClear(@pvd)
   RETURN AfxPropVariantFromResource(hinst, id, @pvd)
END FUNCTION
' ========================================================================================

' =====================================================================================
' Initializes a CPropVar as VT_DATE from a string.
' =====================================================================================
PRIVATE FUNCTION CPropVar.PutDateString (BYVAL pwszDate AS WSTRING PTR, BYVAL lcid AS LCID = 0, BYVAL dwFlags AS ULONG = 0) AS HRESULT
   CPROPVAR_DP("CPropVar PutDateString")
   PropVariantClear(@pvd)
   DIM dateOut AS DOUBLE
   DIM hr AS HRESULT = VarDateFromStr(pwszDate, lcid, dwFlags, @dateOut)
   IF hr = S_OK THEN pvd.vt = VT_DATE : pvd.date = dateOut
   RETURN hr
END FUNCTION
' =====================================================================================

' =====================================================================================
' Initializes a CPropVar as VT_DATE from a DATE type.
' =====================================================================================
PRIVATE FUNCTION CPropVar.PutVbDate (BYVAL vbDate AS DATE_) AS HRESULT
   CPROPVAR_DP("CPropVar PutVbDate")
   PropVariantClear(@pvd)
   pvd.vt = VT_DATE
   pvd.date = vbDate
   RETURN S_OK
END FUNCTION
' =====================================================================================

' =====================================================================================
' Converts a SYSTEMTIME to a VT_DATE variant.
' =====================================================================================
PRIVATE FUNCTION CPropVar.PutSystemTime (BYVAL st AS SYSTEMTIME PTR) AS BOOLEAN
   CPROPVAR_DP("CPropVar PutSystemTime")
   PropVariantClear(@pvd)
   DIM vbDate AS DATE_
   DIM bRes AS BOOLEAN = SystemTimeToVariantTime(st, @vbDate)
   IF bRes THEN pvd.vt = VT_DATE : pvd.date = vbDate
   RETURN bRes
END FUNCTION
' =====================================================================================

' =====================================================================================
' Initializes a CPropVar based on a GUID. The structure is initialized as a VT_LPWSTR type.
' =====================================================================================
PRIVATE FUNCTION CPropVar.PutGuidAsString (BYVAL guid AS GUID PTR) AS HRESULT
   CPROPVAR_DP("CPropVar PutGuid")
   PropVariantClear(@pvd)
   RETURN AfxPropVariantFromGUIDAsString(guid, @pvd)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Initializes a CPropVar based on a GUID. Creates a VT_VECTOR OR VT_UI1 propvariant.
' =====================================================================================
PRIVATE FUNCTION CPropVar.PutGuidAsBuffer (BYVAL guid AS GUID PTR) AS HRESULT
   CPROPVAR_DP("CPropVar PutGuid")
   PropVariantClear(@pvd)
   RETURN AfxPropVariantFromBuffer(@guid, SIZEOF(GUID), @pvd)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Initializes a CPropVar based on a class identifier (CLSID).
' =====================================================================================
PRIVATE FUNCTION CPropVar.PutClsid (BYVAL rclsid AS CLSID PTR) AS HRESULT
   CPROPVAR_DP("CPropVar PutClsid")
   PropVariantClear(@pvd)
   RETURN AfxPropVariantFromCLSID(rclsid, @pvd)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Initializes a CPropVar structure with the contents of a FILETIME structure.
' =====================================================================================
PRIVATE FUNCTION CPropVar.PutFileTime (BYVAL pft AS FILETIME PTR) AS HRESULT
   CPROPVAR_DP("CPropVar PutFileTime")
   PropVariantClear(@pvd)
   RETURN AfxPropVariantFromFileTime(pft, @pvd)
END FUNCTION
' =====================================================================================

' ========================================================================================
' Initializes a VARIANT structure with an array of FILETIME structures.
' ========================================================================================
PRIVATE FUNCTION CPropVar.PutFileTimeArray (BYVAL prgft AS FILETIME PTR, BYVAL cElems AS ULONG) AS HRESULT
   CPROPVAR_DP("CPropVar PutFileTimeArray")
   PropVariantClear(@pvd)
   RETURN AfxPropVariantFromFileTimeVector(prgft, cElems, @pvd)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Initializes a PROPVARIANT structure with a string stored in a STRRET structure.
' =====================================================================================
PRIVATE FUNCTION CPropVar.PutStrRet (BYVAL pstrret AS STRRET PTR, BYVAL pidl AS PCUITEMID_CHILD) AS HRESULT
   CPROPVAR_DP("CPropVar PutStrRet")
   PropVariantClear(@pvd)
   RETURN AfxPropVariantFromStrRet(pstrret, pidl, @pvd)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Initializes a CPropVar as VT_DECIMAL from a DECIMAL structure.
' =====================================================================================
PRIVATE FUNCTION CPropVar.PutDec (BYREF dec AS DECIMAL) AS HRESULT
   CPROPVAR_DP("CPropVar PutDec")
   PropVariantClear(@pvd)
   pvd.vt = VT_DECIMAL
   pvd.decVal = dec
   RETURN S_OK
END FUNCTION
' =====================================================================================

' =====================================================================================
' Initializes a CPropVar as VT_DECIMAL from a string.
' =====================================================================================
PRIVATE FUNCTION CPropVar.PutDecFromStr (BYVAL pwszIn AS WSTRING PTR, BYVAL lcid AS LCID = 0, BYVAL dwFlags AS ULONG = 0) AS HRESULT
   CPROPVAR_DP("CPropVar PutDecFromStr")
   PropVariantClear(@pvd)
   IF pwszIn = NULL THEN RETURN E_INVALIDARG
   DIM dec AS DECIMAL
   DIM hr AS HRESULT = VarDecFromStr(pwszIn, lcid, dwFlags, @dec)
   IF hr = S_OK THEN pvd.vt = VT_DECIMAL : pvd.decVal = dec
   RETURN hr
END FUNCTION
' =====================================================================================

' =====================================================================================
' Initializes a CPropVar as VT_DECIMAL from a double value.
' =====================================================================================
PRIVATE FUNCTION CPropVar.PutDecFromDouble (BYVAL dbIn AS DOUBLE) AS HRESULT
   CPROPVAR_DP("CPropVar PutDecFromStr")
   PropVariantClear(@pvd)
   DIM dec AS DECIMAL
   DIM hr AS HRESULT = VarDecFromR8(dbIn, @dec)
   IF hr = S_OK THEN pvd.vt = VT_DECIMAL : pvd.decVal = dec
   RETURN hr
END FUNCTION
' =====================================================================================

' =====================================================================================
' Converts a currency value to a CPropVar of type VT_DECIMAL.
' =====================================================================================
PRIVATE FUNCTION CPropVar.PutDecFromCy (BYVAL cyIn AS CY) AS HRESULT
   CPROPVAR_DP("CPropVar PutDecFromStr")
   PropVariantClear(@pvd)
   DIM decOut AS DECIMAL
   DIM hr AS HRESULT = VarDecFromCy(cyIn, @decOut)
   IF hr = S_OK THEN pvd.vt = VT_DECIMAL : pvd.decVal = decOut
   RETURN hr
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Initializes a CPropVar from an array of Boolean values.
' DIM rg (2) AS WINBOOL = {-1, 0, -1}
' DIM cpv AS CPropVar
' cpv.PutBooleanArray(@rg(0), 3)
' print cpv.wstr
' =====================================================================================
PRIVATE FUNCTION CPropVar.PutBooleanArray (BYVAL prgf AS WINBOOL PTR, BYVAL cElems AS ULONG) AS HRESULT
   CPROPVAR_DP("CPropVar PutBooleanArray")
   PropVariantClear(@pvd)
   RETURN AfxPropVariantFromBooleanVector(prgf, cElems, @pvd)
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Initializes a CPropVar from an array of 16-bit integer values.
' DIM rg (2) AS SHORT = {123, 234, 345}
' DIM cpv AS CPropVar
' cpv.PutShortArray(@rg(0), 3)
' print cpv.wstr
' =====================================================================================
PRIVATE FUNCTION CPropVar.PutShortArray (BYVAL prgf AS SHORT PTR, BYVAL cElems AS ULONG) AS HRESULT
   CPROPVAR_DP("CPropVar PutShortArray")
   PropVariantClear(@pvd)
   RETURN AfxPropVariantFromInt16Vector(prgf, cElems, @pvd)
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Initializes a CPropVar from an array of unsigned 16-bit integer values.
' DIM rg (2) AS USHORT = {123, 234, 345}
' DIM cpv AS CPropVar
' cpv.PutUShortArray(@rg(0), 3)
' print cpv.wstr
' =====================================================================================
PRIVATE FUNCTION CPropVar.PutUShortArray (BYVAL prgf AS USHORT PTR, BYVAL cElems AS ULONG) AS HRESULT
   CPROPVAR_DP("CPropVar PutUShortArray")
   PropVariantClear(@pvd)
   RETURN AfxPropVariantFromUInt16Vector(prgf, cElems, @pvd)
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Initializes a CPropVar from an array of 32-bit integer values.
' DIM rg (2) AS LONG = {-123, 234, 345}
' DIM cpv AS CPropVar
' cpv.PutLongArray(@rg(0), 3)
' print cpv.wstr
' =====================================================================================
PRIVATE FUNCTION CPropVar.PutLongArray (BYVAL prgn AS LONG PTR, BYVAL cElems AS ULONG) AS CPropVar
   CPROPVAR_DP("CPropVar PutLongArray")
   PropVariantClear(@pvd)
   RETURN AfxPropVariantFromInt32Vector(prgn, cElems, @pvd)
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Initializes a CPropVar from an array of 32-bit integer values.
' DIM rg (2) AS ULONG = {123, 234, 345}
' DIM cpv AS CPropVar
' cpv.PutULongArray(@rg(0), 3)
' print cpv.wstr
' =====================================================================================
PRIVATE FUNCTION CPropVar.PutULongArray (BYVAL prgn AS ULONG PTR, BYVAL cElems AS ULONG) AS HRESULT
   CPROPVAR_DP("CPropVar PutULongArray")
   PropVariantClear(@pvd)
   RETURN AfxPropVariantFromUInt32Vector(prgn, cElems, @pvd)
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Initializes a CPropVar from an array of 64-bit integer values.
' DIM rg (2) AS LONGINT = {123, 234, 345}
' DIM cpv AS CPropVar
' cpv.PutLongIntArray(@rg(0), 3)
' print cpv.wstr
' =====================================================================================
PRIVATE FUNCTION CPropVar.PutLongIntArray (BYVAL prgn AS LONGINT PTR, BYVAL cElems AS ULONG) AS HRESULT
   CPROPVAR_DP("CPropVar PutLongIntArray")
   PropVariantClear(@pvd)
   RETURN AfxPropVariantFromInt64Vector(prgn, cElems, @pvd)
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Initializes a CPropVar from an array of unsigned 64-bit integer values.
' DIM rg (2) AS ULONGINT = {123, 234, 345}
' DIM cpv AS CPropVar
' cpv.PutLongIntArray(@rg(0), 3)
' print cpv.wstr
' =====================================================================================
PRIVATE FUNCTION CPropVar.PutULongIntArray (BYVAL prgn AS ULONGINT PTR, BYVAL cElems AS ULONG) AS HRESULT
   CPROPVAR_DP("CPropVar PutLongIntArray")
   PropVariantClear(@pvd)
   RETURN AfxPropVariantFromUInt64Vector(prgn, cElems, @pvd)
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Initializes a CPropVar from an array of values of type DOUBLE.
' DIM rg (2) AS DOUBLE = {123.12, 234.45, 345.67}
' DIM cpv AS CPropVar
' cpv.PutDoubleArray(@rg(0), 3)
' print cpv.wstr
' =====================================================================================
PRIVATE FUNCTION CPropVar.PutDoubleArray (BYVAL prgn AS DOUBLE PTR, BYVAL cElems AS ULONG) AS HRESULT
   CPROPVAR_DP("CPropVar PutLongIntArray")
   PropVariantClear(@pvd)
   RETURN AfxPropVariantFromDoubleVector(prgn, cElems, @pvd)
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Initializes a CPropVar from an array of unicode strings.
' DIM rg (2) AS WSTRING PTR
' rg(0) = @WSTR("One")
' rg(1) = @WSTR("Two")
' rg(2) = @WSTR("Three")
' DIM cpv AS CPropVar
' cpv.PutStringArray(cast(PCWSTR, @rg(0)), 3)
' print cpv.wstr
' -or-
' DIM rg (2) AS WSTRING PTR
' rg(0) = AfxWstrAlloc("One")
' rg(1) = AfxWstrAlloc("Two")
' rg(2) = AfxWstrAlloc("Three")
' DIM cpv AS CPropVar
' cpv.PutStringArray(cast(PCWSTR, @rg(0)), 3)
' print cpv.wstr
' =====================================================================================
PRIVATE FUNCTION CPropVar.PutStringArray (BYVAL prgsz AS PCWSTR, BYVAL cElems AS ULONG) AS HRESULT
   CPROPVAR_DP("CPropVar PutStringArray")
   PropVariantClear(@pvd)
   RETURN AfxPropVariantFromStringVector(prgsz, cElems, @pvd)
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Initializes a CPropVar from a specified string. The string is parsed as a semi-colon
' delimited list (for example: "A;B;C").
' DIM wsz AS WSTRING * 260 = "One;Two;Three"
' DIM cpv AS CPropVar
' cpv.PutStringAsArray(wsz)
' print cpv.wstr
' =====================================================================================
PRIVATE FUNCTION CPropVar.PutStringAsArray (BYREF wsz AS WSTRING) AS HRESULT
   CPROPVAR_DP("CPropVar PutStringAsArray")
   PropVariantClear(@pvd)
   RETURN AfxPropVariantFromStringAsVector(@wsz, @pvd)
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Converts the contents of a PROPVARIANT structure to a CPropVar.
' DIM pp AS PROPVARIANT
' pp.vt = VT_I4
' pp.lVal = 12345
' DIM p AS CPropVar
' p.PutPropVariant(@pp)
' print p.wstr
' We can also use a CPropVar:
' DIM pp AS CPropVar = CPropVar(12345, "LONG")
' DIM p AS CPropVar
' p.PutPropVariant(*pp)
' print p.wstr
' =====================================================================================
PRIVATE FUNCTION CPropVar.PutPropVariant (BYVAL pPropVar AS PROPVARIANT PTR) AS HRESULT
   CPROPVAR_DP("CPropVar PutPropVariant")
   PropVariantClear(@pvd)
   RETURN PropVariantCopy(@pvd, pPropVar)
END FUNCTION
' =====================================================================================

' ========================================================================================
' * Initializes a CPropVar structure from a specified PROPVARIANT vector element.
' Called PropVariantGetElem (inline function) in the propvarutil.h header, which I find confusing.
' PutPropVariantArrayElem is more descriptive.
' DIM rg(2) AS LONG = {123, 456, 789}
' DIM p1 AS CPropVar
' p1.PutLongArray(@rg(0), 3)
' print p1.wstr
' DIM p2 AS CPropVar
' p2.PutPropVariantArrayElem(p1, 1)
' print p2.wstr
' ========================================================================================
PRIVATE FUNCTION CPropVar.PutPropVariantArrayElem (BYVAL pPropVarIn AS PROPVARIANT PTR, BYVAL iElem AS ULONG) AS CPropVar
   CPROPVAR_DP("CPropVar PutPropVariantArrayElem")
   PropVariantClear(@pvd)
   RETURN AfxPropVariantFromPropVariantVectorElem(pPropVarIn, iElem, @pvd)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Initializes a vector element in a PROPVARIANT structure with a value stored in another PROPVARIANT.
' This function is used to convert a PROPVARIANT structure that contains a single value
' into a vector value.
' For simple source types, this function initializes the PROPVARIANT as a vector of one element.
' For a source that contains a string, this function initializes the PROPVARIANT with zero or
' more substrings taken from the source string, treating semicolons as delimiters.
' See InitPropVariantFromStringAsVector for more details.
' The following input types are supported:
' VT_I2, VT_UI2, VT_I4, VT_UI4, VT_I8, VT_UI8, VT_R8, VT_BOOL, VT_DATE, VT_FILETIME, VT_BSTR, VT_LPWSTR
' DIM p1 AS CPropVar = "Test string"
' DIM p2 AS CPropVar
' p2.PropVariantArrayFromPropVariant(p1)
' print p2.wstr
' ========================================================================================
PRIVATE FUNCTION CPropVar.PropVariantArrayFromPropVariant (BYVAL propvarSingle AS PROPVARIANT PTR) AS HRESULT
   CPROPVAR_DP("CPropVar PropVariantArrayFromPropVariant")
   PropVariantClear(@pvd)
   RETURN AfxPropVariantArrayFromPropVariant(propvarSingle, @pvd)
END FUNCTION
' ========================================================================================

' =====================================================================================
' Returns the number of dimensions for variants of type VT_ARRAY; returns 0 otherwise.
' =====================================================================================
PRIVATE FUNCTION CPropVar.GetDim () AS ULONG
   CPROPVAR_DP("CPropVar GetDim")
   IF (pvd.vt AND VT_ARRAY) = VT_ARRAY THEN
      IF pvd.parray THEN RETURN SafeArrayGetDim(pvd.parray)
   END IF
END FUNCTION
' =====================================================================================

' =====================================================================================
' Returns the lower bound for the specified dimension of the safe array for variants of
' type VT_ARRAY; returns 0 otherwise.
' =====================================================================================
PRIVATE FUNCTION CPropVar.GetLBound (BYVAL nDim AS UINT = 1) AS LONG
   CPROPVAR_DP("CPropVar GetLBound")
   DIM hr AS HRESULT
   IF (pvd.vt AND VT_ARRAY) = VT_ARRAY THEN
      DIM nBound AS LONG
      IF pvd.parray THEN hr = SafeArrayGetLBound(pvd.parray, nDim, @nBound)
      IF hr = S_OK THEN RETURN nBound
   END IF
END FUNCTION
' =====================================================================================

' Returns the upper bound for the specified dimension of the safe array for variants of
' type VT_ARRAY; returns -1 otherwise.
' =====================================================================================
PRIVATE FUNCTION CPropVar.GetUBound (BYVAL nDim AS UINT = 1) AS LONG
   CPROPVAR_DP("CPropVar GetUBound")
   DIM hr AS HRESULT
   IF (pvd.vt AND VT_ARRAY) = VT_ARRAY THEN
      DIM nBound AS LONG
      IF pvd.parray THEN hr = SafeArrayGetUBound(pvd.parray, nDim, @nBound)
      IF hr = S_OK THEN RETURN nBound
   END IF
   RETURN -1
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Returns the element count of a VT_VECTOR or VT_ARRAY value: for single values,
' returns 1; for empty structures, returns 0.
' DIM rg(2) AS LONG = {123, 234, 456}
' DIM p1 AS CPropVar
' p1.PutLongArray(@rg(0), 3)
' print p1.GetElementCount
' =====================================================================================
PRIVATE FUNCTION CPropVar.GetElementCount () AS ULONG
   CPROPVAR_DP("CPropVar GetElementCount")
   RETURN AfxPropVariantGetElementCount(@pvd)
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Extracts a single Boolean element from a CPropVar.
' DIM rg(2) AS WINBOOL = {0, -1, -1}
' DIM p1 AS CPropVar
' p1.PutBooleanArray(@rg(0), 3)
' print p1.GetBooleanElem(1)
' =====================================================================================
PRIVATE FUNCTION CPropVar.GetBooleanElem (BYVAL iElem AS ULONG) AS BOOLEAN
   CPROPVAR_DP("CPropVar GetBooleanElem")
   DIM pfVal AS WINBOOL
   SetLastError AfxPropVariantGetBooleanElem(@pvd, iElem, @pfVal)
   RETURN pfVal
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Extracts a single Int16 element from a CPropVar.
' DIM rg(2) AS SHORT = {-123, 234, 456}
' DIM p1 AS CPropVar
' p1.PutShortArray(@rg(0), 3)
' print p1.GetShortElem(1)
' =====================================================================================
PRIVATE FUNCTION CPropVar.GetShortElem (BYVAL iElem AS ULONG) AS SHORT
   CPROPVAR_DP("CPropVar GetShortElem")
   DIM pnVal AS SHORT
   SetLastError AfxPropVariantGetInt16Elem(@pvd, iElem, @pnVal)
   RETURN pnVal
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Extracts a single unsigned Int16 element from a CPropVar.
' DIM rg(2) AS USHORT = {123, 234, 456}
' DIM p1 AS CPropVar
' p1.PutUShortArray(@rg(0), 3)
' print p1.GetUShortElem(1)
' =====================================================================================
PRIVATE FUNCTION CPropVar.GetUShortElem (BYVAL iElem AS ULONG) AS USHORT
   CPROPVAR_DP("CPropVar GetUShortElem")
   DIM pnVal AS USHORT
   SetLastError AfxPropVariantGetUInt16Elem(@pvd, iElem, @pnVal)
   RETURN pnVal
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Extracts a single Int32 element from a CPropVar.
' DIM rg(2) AS LONG = {123, 234, 456}
' DIM p1 AS CPropVar
' p1.PutLongArray(@rg(0), 3)
' print p1.GetLongElem(1)
' =====================================================================================
PRIVATE FUNCTION CPropVar.GetLongElem (BYVAL iElem AS ULONG) AS LONG
   CPROPVAR_DP("CPropVar GetLongElem")
   DIM pnVal AS LONG
   SetLastError AfxPropVariantGetInt32Elem(@pvd, iElem, @pnVal)
   RETURN pnVal
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Extracts a single unsigned Int32 element from a CPropVar.
' DIM rg(2) AS ULONG = {123, 234, 456}
' DIM p1 AS CPropVar
' p1.PutULongArray(@rg(0), 3)
' print p1.GetULongElem(1)
' =====================================================================================
PRIVATE FUNCTION CPropVar.GetULongElem (BYVAL iElem AS ULONG) AS ULONG
   CPROPVAR_DP("CPropVar GetULongElem")
   DIM pnVal AS ULONG
   SetLastError AfxPropVariantGetUInt32Elem(@pvd, iElem, @pnVal)
   RETURN pnVal
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Extracts a single Int64 element from a CPropVar.
' DIM rg(2) AS LONGINT = {123, 234, 456}
' DIM p1 AS CPropVar
' p1.PutLongIntArray(@rg(0), 3)
' print p1.GetLongIntElem(1)
' =====================================================================================
PRIVATE FUNCTION CPropVar.GetLongIntElem (BYVAL iElem AS ULONG) AS LONGINT
   CPROPVAR_DP("CPropVar GetLongIntElem")
   DIM pnVal AS LONGINT
   SetLastError AfxPropVariantGetInt64Elem(@pvd, iElem, @pnVal)
   RETURN pnVal
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Extracts a single UInt64 element from a CPropVar.
' DIM rg(2) AS ULONGINT = {123, 234, 456}
' DIM p1 AS CPropVar
' p1.PutULongIntArray(@rg(0), 3)
' print p1.GetULongIntElem(1)
' =====================================================================================
PRIVATE FUNCTION CPropVar.GetULongIntElem (BYVAL iElem AS ULONG) AS ULONGINT
   CPROPVAR_DP("CPropVar GetULongIntElem")
   DIM pnVal AS ULONGINT
   SetLastError AfxPropVariantGetUInt64Elem(@pvd, iElem, @pnVal)
   RETURN pnVal
END FUNCTION
' =====================================================================================

' =====================================================================================
'* Extracts a double element from a CPropVar.
' DIM rg(2) AS DOUBLE = {123.12, 234.33, 456.56}
' DIM p1 AS CPropVar
' p1.PutDoubleArray(@rg(0), 3)
' print p1.GetDoubleElem(1)
' =====================================================================================
PRIVATE FUNCTION CPropVar.GetDoubleElem (BYVAL iElem AS ULONG) AS DOUBLE
   CPROPVAR_DP("CPropVar GetDoubleElem")
   DIM pnVal AS DOUBLE
   SetLastError AfxPropVariantGetDoubleElem(@pvd, iElem, @pnVal)
   RETURN pnVal
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Extracts an string element from a PROPVARIANT.
' DIM rg (2) AS WSTRING PTR
' rg(0) = @WSTR("One")
' rg(1) = @WSTR("Two")
' rg(2) = @WSTR("Three")
' DIM cpv AS CPropVar
' cpv.PutStringArray(cast(PCWSTR, @rg(0)), 3)
' print cpv.GetStringElem(1)
' =====================================================================================
PRIVATE FUNCTION CPropVar.GetStringElem (BYVAL iElem AS ULONG) AS CWSTR
   CPROPVAR_DP("CPropVar GetStringElem")
   DIM pwszVal AS WSTRING PTR, cwsVal AS CWSTR
   SetLastError AfxPropVariantGetStringElem(@pvd, iElem, @pwszVal)
   IF pwszVal THEN
      cwsVal = *pwszVal
      CoTaskMemFree pwszVal
   END IF
   RETURN cwsVal
END FUNCTION
' =====================================================================================

' =====================================================================================
' Extracts a FILETIME element from a CPropVar.
' =====================================================================================
PRIVATE FUNCTION CPropVar.GetFileTimeElem (BYVAL iElem AS ULONG) AS FILETIME
   CPROPVAR_DP("CPropVar GetDoubleElem")
   DIM pftVal AS FILETIME
   SetLastError AfxPropVariantGetFileTimeElem(@pvd, iElem, @pftVal)
   RETURN pftVal
END FUNCTION
' =====================================================================================

' =====================================================================================
' Extracts a PROPVARIANT element from a CPropVar.
' =====================================================================================
PRIVATE FUNCTION CPropVar.GetPropVariantElem (BYVAL iElem AS ULONG) AS CPropVar
   CPROPVAR_DP("CPropVar GetDoubleElem")
   DIM ppropvar AS PROPVARIANT
   SetLastError AfxPropVariantFromPropVariantVectorElem(@pvd, iElem, @ppropvar)
   RETURN ppropvar
END FUNCTION
' =====================================================================================

' =====================================================================================
' Formats a variant containing numbers into a string form.
' =====================================================================================
PRIVATE FUNCTION CPropVar.FormatNumber (BYVAL iNumDig AS LONG = -1, BYVAL iIncLead AS LONG = -2, _
   BYVAL iUseParens AS LONG = -2, BYVAL iGroup AS LONG = -2, BYVAL dwFlags AS DWORD = 0) AS CWSTR
   CPROPVAR_DP("CPropVar FormatNumber")
   DIM cwsOut AS CWSTR, bstrOut AS AFX_BSTR
   SetLastError(VarFormatNumber(cast(VARIANT PTR, @pvd), iNumDig, iIncLead, iUseParens, iGroup, dwFlags, @bstrOut))
   cwsOut = *bstrOut
   SysFreeString bstrOut
   RETURN cwsOut
END FUNCTION
' =====================================================================================

' ========================================================================================
' Clears the contents of the PROPVARIANT data and sets the vt field to VT_EMPTY.
' To free arrays of PROVARIANTs, use FreePropVariantArray or ClearPropVariantArray.
' ========================================================================================
PRIVATE SUB CPropVar.Clear
   CPROPVAR_DP("CPropVar Clear")
   PropVariantClear(@pvd)
END SUB
' ========================================================================================

' ========================================================================================
' Attaches a variant to this class. The source PROPVARIANT is marked as empty.
' ========================================================================================
PRIVATE FUNCTION CPropVar.Attach (BYVAL pPropVar AS PROPVARIANT PTR) AS HRESULT
   CPROPVAR_DP("CPropVar ATTACH - PROPVARIANT PTR")
   IF pPropVar = NULL THEN RETURN E_INVALIDARG
   PropVariantClear @pvd
   ' // Copy the contents and give control to CPropVar
   DIM pdest AS ANY PTR = memcpy(@pvd, pPropVar, SIZEOF(PROPVARIANT))
   IF pdest = NULL THEN RETURN E_FAIL
   ' // Mark the source variant as VT_EMPTY instead of clearing it with PropVariantClear
   ' // because we aren't making a duplicate of the contents, but transfering ownership.
   pPropVar->vt = VT_EMPTY
   RETURN S_OK
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CPropVar.Attach (BYREF pv AS PROPVARIANT) AS HRESULT
   CPROPVAR_DP("CPropVar ATTACH - BYREF PROPVARIANT")
   PropVariantClear @pvd
   ' // Copy the contents and give control to CPropVar
   DIM pdest AS ANY PTR = memcpy(@pvd, @pv, SIZEOF(PROPVARIANT))
   IF pdest = NULL THEN RETURN E_FAIL
   ' // Mark the source variant as VT_EMPTY instead of clearing it with VariantClear
   ' // because we aren't making a duplicate of the contents, but transfering ownership.
   pv.vt = VT_EMPTY
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Detaches the variant data from this class and returs it as a PROPVARIANT.
' Don't clear pvd with PropVariantClear because we are transfering ownership.
' ========================================================================================
PRIVATE FUNCTION CPropVar.Detach (BYVAL pPropVar AS PROPVARIANT PTR) AS HRESULT
   CPROPVAR_DP("CPropVar DETACH - PROPVARIANT PTR")
   IF pPropVar = NULL THEN RETURN E_INVALIDARG
   PropVariantClear(pPropVar)
   DIM pdest AS ANY PTR = memcpy(pPropVar, @pvd, SIZEOF(PROPVARIANT))
   IF pdest = NULL THEN RETURN E_FAIL
   pvd.vt = VT_EMPTY
   RETURN S_OK
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CPropVar.Detach (BYREF pv AS PROPVARIANT) AS HRESULT
   CPROPVAR_DP("CPropVar DETACH - BYREF PROPVARIANT")
   PropVariantClear(@pv)
   DIM pdest AS ANY PTR = memcpy(@pv, @pvd, SIZEOF(VARIANT))
   IF pdest = NULL THEN RETURN E_FAIL
   pvd.vt = VT_EMPTY
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the PROPVARIANT type.
' See: https://msdn.microsoft.com/en-us/library/windows/desktop/aa380072(v=vs.85).aspx
' ========================================================================================
PRIVATE FUNCTION CPropVar.vType () AS VARTYPE
   CPROPVAR_DP("CPropVar vType")
   FUNCTION = pvd.vt
END FUNCTION
' ========================================================================================

' =====================================================================================
' * Converts the variant from one type to another.
' Note that the source and destination PROPVARIANT structures must be separate structures.
' You cannot overwrite the source PROPVARIANT data with the new destination data;
' attempting to do so will result in an error.
' =====================================================================================
PRIVATE FUNCTION CPropVar.ChangeType (BYVAL vtNew AS VARTYPE) AS HRESULT
   CPROPVAR_DP("CPropVar ChangeType")
   DIM pv AS PROPVARIANT
   DIM hr AS HRESULT = AfxPropVariantChangeType(@pv, @pvd, 0, vtNew)
   IF hr = S_OK THEN
      PropVariantCopy(@pvd, @pv)
      PropVariantClear @pv
   END IF
   RETURN hr
END FUNCTION
' =====================================================================================

' =====================================================================================
' Converts the string to a 32bit integer
' =====================================================================================
PRIVATE FUNCTION CPropVar.ValLong () AS LONG
   RETURN .ValInt(**AfxPropVarToStr(@pvd))
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CPropVar.ValInt () AS LONG
   RETURN .ValInt(**AfxPropVarToStr(@pvd))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Converts the string to an unsigned 32bit integer
' =====================================================================================
PRIVATE FUNCTION CPropVar.ValULong () AS ULONG
   RETURN .ValUInt(**AfxPropVarToStr(@pvd))
END FUNCTION
' =====================================================================================
PRIVATE FUNCTION CPropVar.ValUInt () AS ULONG
   RETURN .ValUInt(**AfxPropVarToStr(@pvd))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Converts the string to a 64bit integer
' =====================================================================================
PRIVATE FUNCTION CPropVar.ValLongInt () AS LONGINT
   RETURN .ValLng(**AfxPropVarToStr(@pvd))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Converts the string to an unsigned 64bit integer
' =====================================================================================
PRIVATE FUNCTION CPropVar.ValULongInt () AS ULONGINT
   RETURN .ValULng(**AfxPropVarToStr(@pvd))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Converts the string to a floating point number (DOUBLE)
' =====================================================================================
PRIVATE FUNCTION CPropVar.ValDouble () AS DOUBLE
  RETURN .VAL(**AfxPropVarToStr(@pvd))
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CPropVar.Value () AS DOUBLE
   RETURN .VAL(**AfxPropVarToStr(@pvd))
END FUNCTION
' =====================================================================================

' ########################################################################################
'                               *** HELPER FUNCTIONS ***
' ########################################################################################

' ========================================================================================
' Variant types
' ========================================================================================
PRIVATE FUNCTION AfxGetPropVarType (BYVAL pPropVar AS PROPVARIANT PTR) AS VARTYPE
   RETURN pPropVar->vt
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION AfxIsPropVarTypeFloat (BYVAL vt AS VARTYPE) AS BOOLEAN
   RETURN (vt = VT_R4 OR vt = VT_R8)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION AfxIsPropVariantArray (BYVAL pPropVar AS PROPVARIANT PTR) AS BOOLEAN
   IF (pPropVar->vt AND VT_ARRAY = VT_ARRAY) OR (pPropVar->vt AND VT_VECTOR = VT_VECTOR) THEN RETURN TRUE
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION AfxIsPropVariantVector (BYVAL pPropVar AS PROPVARIANT PTR) AS BOOLEAN
   IF (pPropVar->vt AND VT_ARRAY = VT_ARRAY) OR (pPropVar->vt AND VT_VECTOR = VT_VECTOR) THEN RETURN TRUE
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION AfxIsPropVariantString (BYVAL pPropVar AS PROPVARIANT PTR) AS BOOLEAN
   RETURN AfxPropVariantToStringWithDefault(pPropVar, NULL) <> NULL
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION AfxIsPropVarTypeSignedInteger (BYVAL vt AS VARTYPE) AS BOOLEAN
   DIM fRet AS BOOLEAN
   SELECT CASE vt
      CASE VT_I1, VT_I2, VT_I4, VT_I8
         fRet = TRUE
   END SELECT
   RETURN fRet
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION AfxIsPropVarTypeUnsignedInteger (BYVAL vt AS VARTYPE) AS BOOLEAN
   DIM fRet AS BOOLEAN
   SELECT CASE vt
      CASE VT_UI1, VT_UI2, VT_UI4, VT_UI8
         fRet = TRUE
   END SELECT
   RETURN fRet
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION AfxIsPropVarTypeInteger (BYVAL vt AS VARTYPE) AS BOOLEAN
   RETURN AfxIsPropVarTypeSignedInteger(vt) OR AfxIsPropVarTypeUnsignedInteger(vt)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION AfxIsPropVarTypeNumber (BYVAL vt AS VARTYPE) AS BOOLEAN
   RETURN (AfxIsPropVarTypeInteger(vt)) OR (AfxIsPropVarTypeFloat(vt))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts the contents of a CPROPVAR to a CWSTR.
' ========================================================================================
PRIVATE FUNCTION AfxCPropVarToStr OVERLOAD (BYREF cpv AS CPROPVAR) AS CWSTR
   RETURN AfxPropVarToStr(@cpv.pvd)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION AfxCPropVarToStr OVERLOAD (BYVAL pcpv AS CPROPVAR PTR) AS CWSTR
   RETURN AfxPropVarToStr(@pcpv->pvd)
END FUNCTION
' ========================================================================================
#define CPROPVAR_TOSTR(cpv) AfxCPropVarToStr(cpv)
' ========================================================================================
PRIVATE FUNCTION AfxCPropVarOptPrm () AS CPROPVAR
   DIM cpv AS PROPVARIANT
   cpv.vt = VT_ERROR
   cpv.scode = DISP_E_PARAMNOTFOUND
   DIM cpvOpt AS CPROPVAR = cpv
   RETURN cpvOpt
END FUNCTION
' ========================================================================================
#define CPROPVAR_OPTPRM AfxCPropVarOptPrm

' ========================================================================================
' Extracts the contents of a propvariant that contains an array of bytes.
' ========================================================================================
PRIVATE FUNCTION AfxCPropVariantToBuffer (BYREF cvIn AS CPropVar, BYVAL pv AS LPVOID, BYVAL cb AS ULONG) AS ULONG
   RETURN AfxPropVariantToBuffer(@cvIn.pvd, pv, cb)
END FUNCTION
' ========================================================================================

ENUM PROPVAR_COMPARE_UNIT
   PVCU_DEFAULT = 0
   PVCU_SECOND  = 1
   PVCU_MINUTE  = 2
   PVCU_HOUR    = 3
   PVCU_DAY     = 4
   PVCU_MONTH   = 5
   PVCU_YEAR    = 6
END ENUM

TYPE PROPVAR_COMPARE_FLAGS AS LONG
ENUM  ' PROPVAR_COMPARE_FLAGS
   PVCF_DEFAULT                 = &h00000000   ' // When comparing strings, use StrCmpLogical
   PVCF_TREATEMPTYASGREATERTHAN = &h00000001   ' // Empty/null values are greater-than non-empty values
   PVCF_USESTRCMP               = &h00000002   ' // When comparing strings, use StrCmp
   PVCF_USESTRCMPC              = &h00000004   ' // When comparing strings, use StrCmpC
   PVCF_USESTRCMPI              = &h00000008   ' // When comparing strings, use StrCmpI
   PVCF_USESTRCMPIC             = &h00000010   ' // When comparing strings, use StrCmpIC
   PVCF_DIGITSASNUMBERS_CASESENSITIVE = &h00000020   ' // When comparing strings, use CompareStringEx with LOCALE_NAME_USER_DEFAULT and SORT_DIGITSASNUMBERS.  This corresponds to the linguistically correct order for UI lists.
END ENUM

' ========================================================================================
' Compares two PROPVARIANT structures, based on default comparison units and settings.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantCompareEx (BYVAL propvar1 AS PROPVARIANT PTR, BYVAL propvar2 AS PROPVARIANT PTR, BYVAL unit AS PROPVAR_COMPARE_UNIT, BYVAL flags AS PROPVAR_COMPARE_FLAGS) AS HRESULT
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pPropVariantCompareEx AS FUNCTION (BYVAL propvar1 AS PROPVARIANT PTR, BYVAL propvar2 AS PROPVARIANT PTR, BYVAL unit AS PROPVAR_COMPARE_UNIT, BYVAL flags AS PROPVAR_COMPARE_FLAGS) AS HRESULT
   pPropVariantCompareEx = DyLibSymbol(pLib, "PropVariantCompareEx")
   IF pPropVariantCompareEx THEN FUNCTION = pPropVariantCompareEx(propvar1, propvar1, unit, flags)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantCompare (BYVAL propvar1 AS PROPVARIANT PTR, BYVAL propvar2 AS PROPVARIANT PTR) AS HRESULT
   FUNCTION = AfxPropVariantCompareEx(propvar1, propvar2, PVCU_DEFAULT, PVCF_DEFAULT)
END FUNCTION
' ========================================================================================

END NAMESPACE

' ########################################################################################
'                              *** GLOBAL PROCEDURES ***
' ########################################################################################

' // Outside a namespace because they are global
USING Afx

' ========================================================================================
PRIVATE FUNCTION Left OVERLOAD (BYREF cpv AS CPropVar, BYVAL nChars AS INTEGER) AS CWSTR
   DIM cws AS CWSTR = cpv.wstr
   RETURN LEFT(**cws, nChars)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Right OVERLOAD (BYREF cpv AS CPropVar, BYVAL nChars AS INTEGER) AS CWSTR
   DIM cws AS CWSTR = cpv.wstr
   RETURN RIGHT(**cws, nChars)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Val OVERLOAD (BYREF cpv AS CPropVar) AS DOUBLE
   DIM cws AS CWSTR = cpv.wstr
   RETURN .VAL(**cws)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Helper procedure to display messages
' ========================================================================================
PRIVATE FUNCTION AfxMsg OVERLOAD (BYREF cpv AS CPropVar, BYREF wszCaption AS WSTRING = "Message", BYVAL uType AS DWORD = 0) AS LONG
   FUNCTION = MessageBoxW(GetActiveWindow, cpv.wstr, @wszCaption, MB_APPLMODAL OR uType)
END FUNCTION
' ========================================================================================
