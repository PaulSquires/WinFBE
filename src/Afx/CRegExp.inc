' ########################################################################################
' Microsoft Windows
' File: CRegExp.inc
' Contents: Regular expressions
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2017 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "windows.bi"
#include once "Afx/AfxRegExp.bi"
#include once "Afx/AfxCOM.inc"
#include once "Afx/CVar.inc"

NAMESPACE Afx

' ========================================================================================
' CRegExp class
' ========================================================================================
TYPE CRegExp

   Public:
      m_Result AS HRESULT                            ' // HRESULT
      m_pRegExp AS Afx_IRegExp2 PTR                  ' // IRegExp2 interface
      m_pMatches AS Afx_IMatchCollection2 PTR        ' // IMatchCollection2 interface

   Private:
      m_bUninitCOM AS BOOLEAN

   Public:
      DECLARE CONSTRUCTOR
      DECLARE CONSTRUCTOR (BYREF cbsPattern AS CBSTR, BYVAL bIgnoreCase AS BOOLEAN = FALSE, _
        BYVAL bGlobal AS BOOLEAN = FALSE, BYVAL bMultiline AS BOOLEAN = FALSE)
      DECLARE DESTRUCTOR
      DECLARE FUNCTION GetLastResult () AS HRESULT
      DECLARE FUNCTION SetResult (BYVAL Result AS HRESULT) AS HRESULT
      DECLARE PROPERTY Pattern () AS CBSTR
      DECLARE PROPERTY Pattern (BYREF cbsPattern AS CBSTR)
      DECLARE PROPERTY Global () AS BOOLEAN
      DECLARE PROPERTY Global (BYVAL bGlobal AS BOOLEAN)
      DECLARE PROPERTY IgnoreCase () AS BOOLEAN
      DECLARE PROPERTY IgnoreCase (BYVAL bIgnoreCase AS BOOLEAN)
      DECLARE PROPERTY Multiline () AS BOOLEAN
      DECLARE PROPERTY Multiline (BYVAL bMultiline AS BOOLEAN)
      DECLARE FUNCTION Execute (BYREF cbsSourceString AS CBSTR) AS BOOLEAN
      DECLARE FUNCTION Test (BYREF cbsSourceString AS CBSTR) AS BOOLEAN
      DECLARE FUNCTION Replace (BYREF cbsSourceString AS CBSTR, BYREF cvReplaceString AS CVAR) AS CBSTR
      DECLARE FUNCTION MatchCount () AS LONG
      DECLARE FUNCTION MatchValue (BYVAL index AS LONG) AS CBSTR
      DECLARE FUNCTION MatchPos (BYVAL index AS LONG) AS LONG
      DECLARE FUNCTION MatchLen (BYVAL index AS LONG) AS LONG
      DECLARE FUNCTION SubMatchCount (BYVAL index AS LONG) AS LONG
      DECLARE FUNCTION SubMatchValue (BYVAL MatchIndex AS LONG, BYVAL SubMatchIndex AS LONG) AS CVAR

END TYPE
' ========================================================================================

' ========================================================================================
' CRegExp default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CRegExp
   ' // Initialize the COM library
   DIM hr AS HRESULT = CoInitialize(NULL)
   IF hr = S_OK OR hr = S_FALSE THEN m_bUninitCOM = TRUE
   ' // Create an instance of the IRegExp2 interface
   m_pRegExp = AfxNewCom("VBScript.RegExp")
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' * CRegExp constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CRegExp (BYREF cbsPattern AS CBSTR, BYVAL bIgnoreCase AS BOOLEAN = FALSE, _
        BYVAL bGlobal AS BOOLEAN = FALSE, BYVAL bMultiline AS BOOLEAN = FALSE)
   ' // Initialize the COM library
   DIM hr AS HRESULT = CoInitialize(NULL)
   IF hr = S_OK OR hr = S_FALSE THEN m_bUninitCOM = TRUE
   ' // Create an instance of the IRegExp2 interface
   m_pRegExp = AfxNewCom("VBScript.RegExp")
   IF bIgnoreCase THEN m_pRegExp->put_IgnoreCase(TRUE)
   IF bGlobal THEN m_pRegExp->put_Global(TRUE)
   IF bMultiline THEN m_pRegExp->put_Multiline(TRUE)
   m_pRegExp->put_Pattern(cbsPattern)
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CRegExp destructor
' ========================================================================================
PRIVATE DESTRUCTOR CRegExp
   ' // Release the interfaces
   IF m_pMatches THEN m_pMatches->Release
   IF m_pRegExp THEN m_pRegExp->Release
   ' // Free the COM library
   IF m_bUninitCOM THEN CoUninitialize
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the last result code.
' ========================================================================================
PRIVATE FUNCTION CRegExp.GetLastResult () AS HRESULT
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the last result code.
' ========================================================================================
PRIVATE FUNCTION CRegExp.SetResult (BYVAL Result AS HRESULT) AS HRESULT
   m_Result = Result
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' * The regular expression pattern to use.
' ========================================================================================
PROPERTY CRegExp.Pattern () AS CBSTR
   IF m_pRegExp = NULL THEN this.SetResult(E_POINTER): EXIT PROPERTY
   DIM bstrPattern AS AFX_BSTR
   this.SetResult(m_pRegExp->get_Pattern(@bstrPattern))
   PROPERTY = bstrPattern
END PROPERTY
' ========================================================================================
' ========================================================================================
PROPERTY CRegExp.Pattern (BYREF cbsPattern AS CBSTR)
   IF m_pRegExp = NULL THEN this.SetResult(E_POINTER): EXIT PROPERTY
   this.SetResult(m_pRegExp->put_Pattern(cbsPattern))
END PROPERTY
' ========================================================================================

' ========================================================================================
' * Global seach for all occurrences of the pattern.
' ========================================================================================
PROPERTY CRegExp.Global () AS BOOLEAN
   IF m_pRegExp = NULL THEN this.SetResult(E_POINTER): EXIT PROPERTY
   DIM bGlobal AS VARIANT_BOOL
   this.SetResult(m_pRegExp->get_Global(@bGlobal))
   PROPERTY = bGlobal
END PROPERTY
' ========================================================================================
' ========================================================================================
PROPERTY CRegExp.Global (BYVAL bGlobal AS BOOLEAN)
   IF m_pRegExp = NULL THEN this.SetResult(E_POINTER): EXIT PROPERTY
   this.SetResult(m_pRegExp->put_Global(cast(VARIANT_BOOL, bGlobal)))
END PROPERTY
' ========================================================================================

' ========================================================================================
' * Ignore case.
' ========================================================================================
PROPERTY CRegExp.IgnoreCase () AS BOOLEAN
   IF m_pRegExp = NULL THEN this.SetResult(E_POINTER): EXIT PROPERTY
   DIM bIgnoreCase AS VARIANT_BOOL
   this.SetResult(m_pRegExp->get_IgnoreCase(@bIgnoreCase))
   PROPERTY = bIgnoreCase
END PROPERTY
' ========================================================================================
' ========================================================================================
PROPERTY CRegExp.IgnoreCase (BYVAL bIgnoreCase AS BOOLEAN)
   IF m_pRegExp = NULL THEN this.SetResult(E_POINTER): EXIT PROPERTY
   this.SetResult(m_pRegExp->put_IgnoreCase(cast(VARIANT_BOOL, bIgnoreCase)))
END PROPERTY
' ========================================================================================

' ========================================================================================
' Multiline search.
' ========================================================================================
PROPERTY CRegExp.Multiline () AS BOOLEAN
   IF m_pRegExp = NULL THEN this.SetResult(E_POINTER): EXIT PROPERTY
   DIM bMultiline AS VARIANT_BOOL
   this.SetResult(m_pRegExp->get_Multiline(@bMultiline))
   PROPERTY = bMultiline
END PROPERTY
' ========================================================================================
' ========================================================================================
PROPERTY CRegExp.Multiline (BYVAL bMultiline AS BOOLEAN)
   IF m_pRegExp = NULL THEN this.SetResult(E_POINTER): EXIT PROPERTY
   this.SetResult(m_pRegExp->put_Multiline(cast(VARIANT_BOOL, bMultiline)))
END PROPERTY
' ========================================================================================

' ========================================================================================
' * Executes a regular expression search against a specified string.
' ========================================================================================
FUNCTION CRegExp.Execute (BYREF cbsSourceString AS CBSTR) AS BOOLEAN
   IF m_pRegExp = NULL THEN this.SetResult(E_POINTER): RETURN FALSE
   IF m_pMatches THEN AfxSafeRelease(m_pMatches)
   this.SetResult(m_pRegExp->Execute(cbsSourceString, cast(Afx_IDispatch PTR PTR, @m_pMatches)))
   IF m_Result = S_OK THEN RETURN TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Executes a regular expression search against a specified string and returns a Boolean
' value that indicates if a pattern match was found.
' ========================================================================================
FUNCTION CRegExp.Test (BYREF cbsSourceString AS CBSTR) AS BOOLEAN
   IF m_pRegExp = NULL THEN this.SetResult(E_POINTER): RETURN FALSE
   DIM bMatch AS VARIANT_BOOL
   this.SetResult(m_pRegExp->Test(cbsSourceString, @bMatch))
   RETURN bMatch
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns a copy of a string with text replaced using a regular expression or search string.
' ========================================================================================
FUNCTION CRegExp.Replace (BYREF cbsSourceString AS CBSTR, BYREF cvReplaceString AS CVAR) AS CBSTR
   DIM bstrDestString AS AFX_BSTR
   IF m_pRegExp = NULL THEN
      this.SetResult(E_INVALIDARG)
   ELSE
      this.SetResult(m_pRegExp->Replace(cbsSourceString, cvReplaceString, @bstrDestString))
   END IF
   RETURN bstrDestString
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the number of matches.
' ========================================================================================
FUNCTION CRegExp.MatchCount () AS LONG
   IF m_pMatches = NULL THEN this.SetResult(E_POINTER): RETURN 0
   DIM nCount AS LONG
   this.SetResult(m_pMatches->get_Count(@nCount))
   RETURN nCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the value or text of a match found in a search string.
' ========================================================================================
FUNCTION CRegExp.MatchValue (BYVAL index AS LONG) AS CBSTR
   DIM bstrValue AS AFX_BSTR
   IF m_pMatches = NULL THEN
      this.SetResult(E_POINTER)
   ELSE
      DIM pMatch AS Afx_IMatch2 PTR
      this.SetResult(m_pMatches->get_Item(index, cast(Afx_IDispatch PTR PTR, @pMatch)))
      IF pMatch THEN
         pMatch->get_Value(@bstrValue)
         AfxSafeRelease(pMatch)
      END IF
   END IF
   RETURN bstrValue
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the position in a search string where a match occurs.
' ========================================================================================
FUNCTION CRegExp.MatchPos (BYVAL index AS LONG) AS LONG
   DIM nFirstIndex AS LONG
   IF m_pMatches = NULL THEN
      this.SetResult(E_POINTER)
   ELSE
      DIM pMatch AS Afx_IMatch2 PTR
      this.SetResult(m_pMatches->get_Item(index, cast(Afx_IDispatch PTR PTR, @pMatch)))
      IF pMatch THEN
         pMatch->get_FirstIndex(@nFirstIndex)
         AfxSafeRelease(pMatch)
      END IF
   END IF
   RETURN nFirstIndex
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the position in a search string where a match occurs.
' ========================================================================================
FUNCTION CRegExp.MatchLen (BYVAL index AS LONG) AS LONG
   DIM nLen AS LONG
   IF m_pMatches = NULL THEN
      this.SetResult(E_POINTER)
   ELSE
      DIM pMatch AS Afx_IMatch2 PTR
      this.SetResult(m_pMatches->get_Item(index, cast(Afx_IDispatch PTR PTR, @pMatch)))
      IF pMatch THEN
         pMatch->get_Length(@nLen)
         AfxSafeRelease(pMatch)
      END IF
   END IF
   FUNCTION = nLen
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the number of submatches.
' ========================================================================================
FUNCTION CRegExp.SubMatchCount (BYVAL index AS LONG) AS LONG
   DIM nCount AS LONG
   IF m_pMatches = NULL THEN
      this.SetResult(E_POINTER)
   ELSE
      DIM pMatch AS Afx_IMatch2 PTR
      this.SetResult(m_pMatches->get_Item(index, cast(Afx_IDispatch PTR PTR, @pMatch)))
      IF pMatch THEN
         DIM pSubMatches AS Afx_ISubMatches PTR
         this.SetResult(pMatch->get_SubMatches(cast(Afx_IDispatch PTR PTR, @pSubMatches)))
         AfxSafeRelease(pMatch)
         IF pSubMatches THEN
            this.SetResult(pSubMatches->get_Count(@nCount))
            AfxSafeRelease(pSubMatches)
         END IF
      END IF
   END IF
   RETURN nCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the value or text of a submatch found in a search string.
' ========================================================================================
FUNCTION CRegExp.SubMatchValue (BYVAL MatchIndex AS LONG, BYVAL SubMatchIndex AS LONG) AS CVAR
   DIM vValue AS VARIANT
   IF m_pMatches = NULL THEN
      this.SetResult(E_POINTER)
   ELSE
      DIM pMatch AS Afx_IMatch2 PTR
      this.SetResult(m_pMatches->get_Item(MatchIndex, cast(Afx_IDispatch PTR PTR, @pMatch)))
      IF pMatch THEN
         DIM pSubMatches AS Afx_ISubMatches PTR
         this.SetResult(pMatch->get_SubMatches(cast(Afx_IDispatch PTR PTR, @pSubMatches)))
         AfxSafeRelease(pMatch)
         IF pSubMatches THEN
            this.SetResult(pSubMatches->get_Item(SubMatchIndex, @vValue))
            AfxSafeRelease(pSubMatches)
         END IF
      END IF
   END IF
   RETURN vValue
END FUNCTION
' ========================================================================================

END NAMESPACE
