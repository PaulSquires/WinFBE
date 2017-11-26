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
      DECLARE CONSTRUCTOR (BYREF pRegExp AS CRegExp)
      DECLARE CONSTRUCTOR (BYVAL bIgnoreCase AS BOOLEAN, BYVAL bGlobal AS BOOLEAN = TRUE, BYVAL bMultiline AS BOOLEAN = FALSE)
      DECLARE CONSTRUCTOR (BYREF cbsPattern AS CBSTR, BYVAL bIgnoreCase AS BOOLEAN = FALSE, _
        BYVAL bGlobal AS BOOLEAN = FALSE, BYVAL bMultiline AS BOOLEAN = FALSE)
      DECLARE DESTRUCTOR
      DECLARE FUNCTION GetLastResult () AS HRESULT
      DECLARE FUNCTION SetResult (BYVAL Result AS HRESULT) AS HRESULT
      DECLARE PROPERTY RegExpPtr () AS Afx_IRegExp2 PTR
      DECLARE PROPERTY Pattern () AS CBSTR
      DECLARE PROPERTY Pattern (BYREF cbsPattern AS CBSTR)
      DECLARE PROPERTY Global () AS BOOLEAN
      DECLARE PROPERTY Global (BYVAL bGlobal AS BOOLEAN)
      DECLARE PROPERTY IgnoreCase () AS BOOLEAN
      DECLARE PROPERTY IgnoreCase (BYVAL bIgnoreCase AS BOOLEAN)
      DECLARE PROPERTY Multiline () AS BOOLEAN
      DECLARE PROPERTY Multiline (BYVAL bMultiline AS BOOLEAN)
      DECLARE FUNCTION Execute (BYREF cbsSourceString AS CBSTR, BYVAL bIgnoreCase AS BOOLEAN = FALSE, _
              BYVAL bGlobal AS BOOLEAN = TRUE, BYVAL bMultiline AS BOOLEAN = FALSE) AS CRegExp
      DECLARE FUNCTION Execute (BYREF cbsSourceString AS CBSTR, BYREF cbsPattern AS CBSTR, _
              BYVAL bIgnoreCase AS BOOLEAN = FALSE, BYVAL bGlobal AS BOOLEAN = TRUE, BYVAL bMultiline AS BOOLEAN = FALSE) AS CRegExp
      DECLARE FUNCTION MatchesCount () AS LONG
      DECLARE FUNCTION MatchValue (BYVAL index AS LONG = 0) AS CBSTR
      DECLARE FUNCTION MatchPos (BYVAL index AS LONG = 0) AS LONG
      DECLARE FUNCTION MatchLen (BYVAL index AS LONG = 0) AS LONG
      DECLARE FUNCTION SubMatchesCount (BYVAL index AS LONG = 0) AS LONG
      DECLARE FUNCTION SubMatchValue (BYVAL MatchIndex AS LONG = 0, BYVAL SubMatchIndex AS LONG = 0) AS CVAR
      DECLARE FUNCTION Test (BYREF cbsSourceString AS CBSTR, BYVAL bIgnoreCase AS BOOLEAN = FALSE, _
              BYVAL bMultiline AS BOOLEAN = FALSE) AS BOOLEAN
      DECLARE FUNCTION Test (BYREF cbsSourceString AS CBSTR, BYREF cbsPattern AS CBSTR, _
              BYVAL bIgnoreCase AS BOOLEAN = FALSE, BYVAL bMultiline AS BOOLEAN = FALSE) AS BOOLEAN
      DECLARE FUNCTION Find (BYREF cbsSourceString AS CBSTR, BYVAL bIgnoreCase AS BOOLEAN = FALSE) AS LONG
      DECLARE FUNCTION Find (BYREF cbsSourceString AS CBSTR, BYREF cbsPattern AS CBSTR, BYVAL bIgnoreCase AS BOOLEAN = FALSE) AS LONG
      DECLARE FUNCTION Find (BYVAL nStart AS LONG, BYREF cbsSourceString AS CBSTR, BYVAL bIgnoreCase AS BOOLEAN = FALSE) AS LONG
      DECLARE FUNCTION Find (BYVAL nStart AS LONG, BYREF cbsSourceString AS CBSTR, BYREF cbsPattern AS CBSTR, BYVAL bIgnoreCase AS BOOLEAN = FALSE) AS LONG
      DECLARE FUNCTION FindEx (BYREF cbsSourceString AS CBSTR, BYVAL bIgnoreCase AS BOOLEAN = FALSE, _
              BYVAL bGlobal AS BOOLEAN = TRUE, BYVAL bMultiline AS BOOLEAN = TRUE) AS CBSTR
      DECLARE FUNCTION FindEx (BYREF cbsSourceString AS CBSTR, BYREF cbsPattern AS CBSTR, _
              BYVAL bIgnoreCase AS BOOLEAN = FALSE, BYVAL bGlobal AS BOOLEAN = TRUE, BYVAL bMultiline AS BOOLEAN = TRUE) AS CBSTR
      DECLARE FUNCTION Replace (BYREF cbsSourceString AS CBSTR, BYREF cvReplaceString AS CVAR, _
              BYVAL bIgnoreCase AS BOOLEAN = FALSE, BYVAL bGlobal AS BOOLEAN = TRUE, BYVAL bMultiline AS BOOLEAN = FALSE) AS CBSTR
      DECLARE FUNCTION Replace (BYREF cbsSourceString AS CBSTR, BYREF cbsPattern AS CBSTR, BYREF cvReplaceString AS CVAR, _
              BYVAL bIgnoreCase AS BOOLEAN = FALSE, BYVAL bGlobal AS BOOLEAN = TRUE, BYVAL bMultiline AS BOOLEAN = FALSE) AS CBSTR
      DECLARE FUNCTION Remove (BYREF cbsSourceString AS CBSTR, BYVAL bIgnoreCase AS BOOLEAN = FALSE, _
              BYVAL bGlobal AS BOOLEAN = TRUE, BYVAL bMultiline AS BOOLEAN = FALSE) AS CBSTR
      DECLARE FUNCTION Remove (BYREF cbsSourceString AS CBSTR, BYREF cbsPattern AS CBSTR, _
              BYVAL bIgnoreCase AS BOOLEAN = FALSE, BYVAL bGlobal AS BOOLEAN = TRUE, BYVAL bMultiline AS BOOLEAN = FALSE) AS CBSTR
      DECLARE FUNCTION Extract (BYREF cbsSourceString AS CBSTR, BYVAL bIgnoreCase AS BOOLEAN = FALSE, BYVAL bGlobal AS BOOLEAN = FALSE, BYVAL bMultiline AS BOOLEAN = FALSE) AS CBSTR
      DECLARE FUNCTION Extract (BYREF cbsSourceString AS CBSTR, BYREF cbsPattern AS CBSTR, BYVAL bIgnoreCase AS BOOLEAN = FALSE, BYVAL bGlobal AS BOOLEAN = FALSE, BYVAL bMultiline AS BOOLEAN = FALSE) AS CBSTR
      DECLARE FUNCTION Extract (BYVAL nStart AS LONG, BYREF cbsSourceString AS CBSTR, BYVAL bIgnoreCase AS BOOLEAN = FALSE, BYVAL bGlobal AS BOOLEAN = FALSE, BYVAL bMultiline AS BOOLEAN = FALSE) AS CBSTR
      DECLARE FUNCTION Extract (BYVAL nStart AS LONG, BYREF cbsSourceString AS CBSTR, BYREF cbsPattern AS CBSTR, BYVAL bIgnoreCase AS BOOLEAN = FALSE, BYVAL bGlobal AS BOOLEAN = FALSE, BYVAL bMultiline AS BOOLEAN = FALSE) AS CBSTR

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
' Clones the passed CRegExp object.
' ========================================================================================
PRIVATE CONSTRUCTOR CRegExp (BYREF pRegExp AS CRegExp)
   ' // Initialize the COM library
   DIM hr AS HRESULT = CoInitialize(NULL)
   IF hr = S_OK OR hr = S_FALSE THEN m_bUninitCOM = TRUE
   ' // Copy pointers and increment the reference count
   m_pRegExp = pRegExp.m_pRegExp
   AfxSafeAddRef(m_pRegExp)
   m_pMatches = pRegExp.m_pMatches
   AfxSafeAddRef(m_pMatches)
   this.IgnoreCase = pRegExp.IgnoreCase
   this.Global = pRegExp.Global
   this.Multiline = pRegExp.Multiline
   this.Pattern = pRegExp.Pattern
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' * CRegExp constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CRegExp (BYVAL bIgnoreCase AS BOOLEAN, BYVAL bGlobal AS BOOLEAN = TRUE, BYVAL bMultiline AS BOOLEAN = FALSE)
   ' // Initialize the COM library
   DIM hr AS HRESULT = CoInitialize(NULL)
   IF hr = S_OK OR hr = S_FALSE THEN m_bUninitCOM = TRUE
   ' // Create an instance of the IRegExp2 interface
   m_pRegExp = AfxNewCom("VBScript.RegExp")
   IF bIgnoreCase THEN m_pRegExp->put_IgnoreCase(TRUE)
   IF bGlobal THEN m_pRegExp->put_Global(TRUE)
   IF bMultiline THEN m_pRegExp->put_Multiline(TRUE)
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
' Returns a pointer to the Afx_IRegExp2 interface
' ========================================================================================
PRIVATE PROPERTY CRegExp.RegExpPtr () AS Afx_IRegExp2 PTR
   PROPERTY = m_pRegExp
END PROPERTY
' ========================================================================================

' ========================================================================================
' * The regular expression pattern to use.
' ========================================================================================
PRIVATE PROPERTY CRegExp.Pattern () AS CBSTR
   IF m_pRegExp = NULL THEN this.SetResult(E_POINTER): EXIT PROPERTY
   DIM bstrPattern AS AFX_BSTR
   this.SetResult(m_pRegExp->get_Pattern(@bstrPattern))
   PROPERTY = bstrPattern
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CRegExp.Pattern (BYREF cbsPattern AS CBSTR)
   IF m_pRegExp = NULL THEN this.SetResult(E_POINTER): EXIT PROPERTY
   this.SetResult(m_pRegExp->put_Pattern(cbsPattern))
END PROPERTY
' ========================================================================================

' ========================================================================================
' * Global seach for all occurrences of the pattern.
' ========================================================================================
PRIVATE PROPERTY CRegExp.Global () AS BOOLEAN
   IF m_pRegExp = NULL THEN this.SetResult(E_POINTER): EXIT PROPERTY
   DIM bGlobal AS VARIANT_BOOL
   this.SetResult(m_pRegExp->get_Global(@bGlobal))
   PROPERTY = bGlobal
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CRegExp.Global (BYVAL bGlobal AS BOOLEAN)
   IF m_pRegExp = NULL THEN this.SetResult(E_POINTER): EXIT PROPERTY
   this.SetResult(m_pRegExp->put_Global(cast(VARIANT_BOOL, bGlobal)))
END PROPERTY
' ========================================================================================

' ========================================================================================
' * Ignore case.
' ========================================================================================
PRIVATE PROPERTY CRegExp.IgnoreCase () AS BOOLEAN
   IF m_pRegExp = NULL THEN this.SetResult(E_POINTER): EXIT PROPERTY
   DIM bIgnoreCase AS VARIANT_BOOL
   this.SetResult(m_pRegExp->get_IgnoreCase(@bIgnoreCase))
   PROPERTY = bIgnoreCase
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CRegExp.IgnoreCase (BYVAL bIgnoreCase AS BOOLEAN)
   IF m_pRegExp = NULL THEN this.SetResult(E_POINTER): EXIT PROPERTY
   this.SetResult(m_pRegExp->put_IgnoreCase(cast(VARIANT_BOOL, bIgnoreCase)))
END PROPERTY
' ========================================================================================

' ========================================================================================
' Multiline search.
' ========================================================================================
PRIVATE PROPERTY CRegExp.Multiline () AS BOOLEAN
   IF m_pRegExp = NULL THEN this.SetResult(E_POINTER): EXIT PROPERTY
   DIM bMultiline AS VARIANT_BOOL
   this.SetResult(m_pRegExp->get_Multiline(@bMultiline))
   PROPERTY = bMultiline
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CRegExp.Multiline (BYVAL bMultiline AS BOOLEAN)
   IF m_pRegExp = NULL THEN this.SetResult(E_POINTER): EXIT PROPERTY
   this.SetResult(m_pRegExp->put_Multiline(cast(VARIANT_BOOL, bMultiline)))
END PROPERTY
' ========================================================================================

' ========================================================================================
' * Executes a regular expression search against a specified string.
' ========================================================================================
PRIVATE FUNCTION CRegExp.Execute (BYREF cbsSourceString AS CBSTR, BYVAL bIgnoreCase AS BOOLEAN = FALSE, _
   BYVAL bGlobal AS BOOLEAN = TRUE, BYVAL bMultiline AS BOOLEAN = FALSE) AS CRegExp
   IF m_pRegExp = NULL THEN this.SetResult(E_POINTER): RETURN FALSE
   IF m_pMatches THEN AfxSafeRelease(m_pMatches)
   m_pRegExp->put_IgnoreCase(bIgnoreCase)
   m_pRegExp->put_Global(bGlobal)
   m_pRegExp->put_Multiline(bMultiline)
   this.SetResult(m_pRegExp->Execute(cbsSourceString, cast(Afx_IDispatch PTR PTR, @m_pMatches)))
   IF m_Result = S_OK THEN RETURN this
END FUNCTION
' ========================================================================================
' ========================================================================================
' * Executes a regular expression search against a specified string.
' - cbsSourceString : The source string.
' - cbsPattern : The pattern to search.
' - bIgnoreCase : FALSE = case sensitive; TRUE = case insensitive.
' - bGlobal : FALSE = Delete only the first match; TRUE = delete all matches.
' - bMultiline : TRUE = Match at the start and the end of multiple lines separated by line breaks.
' ========================================================================================
PRIVATE FUNCTION CRegExp.Execute (BYREF cbsSourceString AS CBSTR, BYREF cbsPattern AS CBSTR, _
   BYVAL bIgnoreCase AS BOOLEAN = FALSE, BYVAL bGlobal AS BOOLEAN = TRUE, BYVAL bMultiline AS BOOLEAN = FALSE) AS CRegExp
   IF m_pRegExp = NULL THEN this.SetResult(E_POINTER): RETURN FALSE
   IF m_pMatches THEN AfxSafeRelease(m_pMatches)
   m_pRegExp->put_IgnoreCase(bIgnoreCase)
   m_pRegExp->put_Global(bGlobal)
   m_pRegExp->put_Pattern(cbsPattern)
   m_pRegExp->put_Multiline(bMultiline)
   this.SetResult(m_pRegExp->Execute(cbsSourceString, cast(Afx_IDispatch PTR PTR, @m_pMatches)))
   IF m_Result = S_OK THEN RETURN this
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the number of matches.
' ========================================================================================
PRIVATE FUNCTION CRegExp.MatchesCount () AS LONG
   IF m_pMatches = NULL THEN this.SetResult(E_POINTER): RETURN 0
   DIM nCount AS LONG
   this.SetResult(m_pMatches->get_Count(@nCount))
   RETURN nCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the value or text of a match found in a search string.
' ========================================================================================
PRIVATE FUNCTION CRegExp.MatchValue (BYVAL index AS LONG) AS CBSTR
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
PRIVATE FUNCTION CRegExp.MatchPos (BYVAL index AS LONG) AS LONG
   DIM nFirstIndex AS LONG
   IF m_pMatches = NULL THEN
      this.SetResult(E_POINTER)
   ELSE
      DIM pMatch AS Afx_IMatch2 PTR
      this.SetResult(m_pMatches->get_Item(index, cast(Afx_IDispatch PTR PTR, @pMatch)))
      IF pMatch THEN
         ' // FirstIndex property indicates the number of characters in the string to
         ' // the left of the match. If the match was found at the very start of the
         ' // string, FirstIndex will be zero. If the match starts at the second character
         ' // in the string, FirstIndex will be one, etc.
         IF pMatch->get_FirstIndex(@nFirstIndex) = S_OK THEN nFirstIndex += 1
         AfxSafeRelease(pMatch)
      END IF
   END IF
   RETURN nFirstIndex
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the position in a search string where a match occurs.
' ========================================================================================
PRIVATE FUNCTION CRegExp.MatchLen (BYVAL index AS LONG) AS LONG
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
PRIVATE FUNCTION CRegExp.SubMatchesCount (BYVAL index AS LONG) AS LONG
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
PRIVATE FUNCTION CRegExp.SubMatchValue (BYVAL MatchIndex AS LONG, BYVAL SubMatchIndex AS LONG) AS CVAR
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

' ========================================================================================
' * Executes a regular expression search against a specified string and returns a Boolean
' value that indicates if a pattern match was found.
' - cbsSourceString : The source string.
' - cbsPattern : The pattern to search.
' - bIgnoreCase : FALSE = case sensitive; TRUE = case insensitive.
' - bMultiline : TRUE = Match at the start and the end of multiple lines separated by line breaks.
' Note: The Global property has no effect on the Test method.
' ========================================================================================
PRIVATE FUNCTION CRegExp.Test (BYREF cbsSourceString AS CBSTR, BYVAL bIgnoreCase AS BOOLEAN = FALSE, _
   BYVAL bMultiline AS BOOLEAN = FALSE) AS BOOLEAN
   IF m_pRegExp = NULL THEN this.SetResult(E_POINTER): RETURN FALSE
   m_pRegExp->put_IgnoreCase(bIgnoreCase)
   m_pRegExp->put_Multiline(bMultiline)
   DIM bMatch AS VARIANT_BOOL
   this.SetResult(m_pRegExp->Test(cbsSourceString, @bMatch))
   RETURN bMatch
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CRegExp.Test (BYREF cbsSourceString AS CBSTR, BYREF cbsPattern AS CBSTR, _
   BYVAL bIgnoreCase AS BOOLEAN = FALSE, BYVAL bMultiline AS BOOLEAN = FALSE) AS BOOLEAN
   IF m_pRegExp = NULL THEN this.SetResult(E_POINTER): RETURN FALSE
   m_pRegExp->put_IgnoreCase(bIgnoreCase)
   m_pRegExp->put_Pattern(cbsPattern)
   m_pRegExp->put_Multiline(bMultiline)
   DIM bMatch AS VARIANT_BOOL
   this.SetResult(m_pRegExp->Test(cbsSourceString, @bMatch))
   RETURN bMatch
END FUNCTION
' ========================================================================================

' ========================================================================================
' Find function with VBScript regular expressions search patterns.
' Parameters:
' - cbsSourceString = The text to be parsed.
' - cbsPattern = The pattern to match.
' - bIgnoreCase = Ignore case.
' Return value:
' - Returns the position of the match or 0 if not found.
'   The length of the match can be retrieved calling pRegExp.MatchLen.
' Usage example:
'   DIM pRegExp AS CRegExp
'   DIM cbsText AS CBSTR = "blah blah a234 blah blah x345 blah blah"
'   DIM cbsPattern AS CBSTR = "[a-z][0-9][0-9][0-9]"
'   DIM nPos AS LONG = pRegExp.Find(cbsText, cbsPattern)
' ========================================================================================
PRIVATE FUNCTION CRegExp.Find (BYREF cbsSourceString AS CBSTR, BYVAL bIgnoreCase AS BOOLEAN = FALSE) AS LONG
   IF m_pRegExp = NULL THEN this.SetResult(E_POINTER): RETURN 0
   IF m_pMatches THEN AfxSafeRelease(m_pMatches)
   m_pRegExp->put_IgnoreCase(bIgnoreCase)
   m_pRegExp->put_Global(FALSE)
   m_pRegExp->put_Multiline(FALSE)
   this.SetResult(m_pRegExp->Execute(cbsSourceString, cast(Afx_IDispatch PTR PTR, @m_pMatches)))
   RETURN this.MatchPos(0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CRegExp.Find (BYREF cbsSourceString AS CBSTR, BYREF cbsPattern AS CBSTR, BYVAL bIgnoreCase AS BOOLEAN = FALSE) AS LONG
   IF m_pRegExp = NULL THEN this.SetResult(E_POINTER): RETURN 0
   IF m_pMatches THEN AfxSafeRelease(m_pMatches)
   m_pRegExp->put_IgnoreCase(bIgnoreCase)
   m_pRegExp->put_Global(FALSE)
   m_pRegExp->put_Multiline(FALSE)
   m_pRegExp->put_Pattern(cbsPattern)
   this.SetResult(m_pRegExp->Execute(cbsSourceString, cast(Afx_IDispatch PTR PTR, @m_pMatches)))
   RETURN this.MatchPos(0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Find function with VBScript regular expressions search patterns.
' Parameters:
' - nStart: The position in the string at which the search will begin. The first character
'   starts at position 1.
' - cbsSourceString = The text to be parsed.
' - cbsPattern = The pattern to match.
' - bIgnoreCase = Ignore case.
' Return value:
' - Returns the position of the match or 0 if not found.
'   The length of the match can be retrieved calling pRegExp.MatchLen.
' Usage example:
'   DIM pRegExp AS CRegExp
'   DIM cbsText AS CBSTR = "blah blah a234 blah blah x345 blah blah"
'   DIM cbsPattern AS CBSTR = "[a-z][0-9][0-9][0-9]"
'   DIM nPos AS LONG = pRegExp.Find(15, cbsText, cbsPattern)
' ========================================================================================
PRIVATE FUNCTION CRegExp.Find (BYVAL nStart AS LONG, BYREF cbsSourceString AS CBSTR, BYVAL bIgnoreCase AS BOOLEAN = FALSE) AS LONG
   IF m_pRegExp = NULL THEN this.SetResult(E_POINTER): RETURN 0
   IF m_pMatches THEN AfxSafeRelease(m_pMatches)
   m_pRegExp->put_IgnoreCase(bIgnoreCase)
   m_pRegExp->put_Global(FALSE)
   m_pRegExp->put_Multiline(FALSE)
   DIM pbstr AS BSTR
   IF nStart < 0 OR nStart > SysStringLen(cbsSourceString.m_bstr) THEN RETURN 0
   DIM cbs AS CBSTR = cbsSourceString.MidChars(nStart)
   this.SetResult(m_pRegExp->Execute(cbs, cast(Afx_IDispatch PTR PTR, @m_pMatches)))
   RETURN this.MatchPos(0) + nStart - 1
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CRegExp.Find (BYVAL nStart AS LONG, BYREF cbsSourceString AS CBSTR, BYREF cbsPattern AS CBSTR, BYVAL bIgnoreCase AS BOOLEAN = FALSE) AS LONG
   IF m_pRegExp = NULL THEN this.SetResult(E_POINTER): RETURN 0
   IF m_pMatches THEN AfxSafeRelease(m_pMatches)
   m_pRegExp->put_IgnoreCase(bIgnoreCase)
   m_pRegExp->put_Global(FALSE)
   m_pRegExp->put_Multiline(FALSE)
   m_pRegExp->put_Pattern(cbsPattern)
   DIM pbstr AS BSTR
   IF nStart < 0 OR nStart > SysStringLen(cbsSourceString.m_bstr) THEN RETURN 0
   DIM cbs AS CBSTR = cbsSourceString.MidChars(nStart)
   this.SetResult(m_pRegExp->Execute(cbs, cast(Afx_IDispatch PTR PTR, @m_pMatches)))
   RETURN this.MatchPos(0) + nStart - 1
END FUNCTION
' ========================================================================================

' ========================================================================================
' Global, multiline find function with VBScript regular expressions search patterns.
' Parameters:
' - cbsSourceString = The text to be parsed.
' - cbsPattern = The pattern to match.
' - bIgnoreCase = Ignore case.
' - bGlobal : FALSE = Return only the first match; TRUE = return all matches.
' - bMultiline : TRUE = Match at the start and the end of multiple lines separated by line breaks.
' Return value:
' - Returns a list of comma separated "index, length" value pairs. The pairs are separated
'   by a semicolon.
' Usage Example:
'   DIM pRegExp AS CRegExp
'   DIM cbsText AS CBSTR = "blah blah a234 blah blah x345 blah blah"
'   DIM cbsPattern AS CBSTR = "[a-z][0-9][0-9][0-9]"
'   DIM cbsOut AS CBSTR
'   cbsOut = pRegExp.FindEx(cbsText, cbsPattern)
' ========================================================================================
PRIVATE FUNCTION CRegExp.FindEx (BYREF cbsSourceString AS CBSTR, BYVAL bIgnoreCase AS BOOLEAN = FALSE, _
   BYVAL bGlobal AS BOOLEAN = TRUE, BYVAL bMultiline AS BOOLEAN = TRUE) AS CBSTR

   IF m_pRegExp = NULL THEN this.SetResult(E_POINTER): RETURN ""
   m_pRegExp->put_IgnoreCase(bIgnoreCase)
   m_pRegExp->put_Global(bGlobal)
   m_pRegExp->put_Multiline(bMultiline)

   DIM cbsOut AS CBSTR, pMatches AS Afx_IMatchCollection2 PTR
   this.SetResult(m_pRegExp->Execute(cbsSourceString, cast(Afx_IDispatch PTR PTR, @pMatches)))
   IF pMatches THEN
      DIM nCount AS LONG
      pMatches->get_Count(@nCount)
      FOR i AS LONG = 0 TO nCount - 1
         DIM pMatch AS Afx_IMatch2 PTR
         this.SetResult(pMatches->get_Item(i, cast(Afx_IDispatch PTR PTR, @pMatch)))
         IF pMatch THEN
            DIM nFirstIndex AS LONG
            pMatch->get_FirstIndex(@nFirstIndex)
            DIM nLen AS LONG
            pMatch->get_Length(@nLen)
            IF i < nCount - 1 THEN
               cbsOut += STR(nFirstIndex + 1) & "," & STR(nLen) & ";"
            ELSE
               cbsOut += STR(nFirstIndex + 1) & "," & STR(nLen)
            END IF
            AfxSafeRelease(pMatch)
         END IF
      NEXT
      AfxSafeRelease(pMatches)
   END IF
   RETURN cbsOut

END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CRegExp.FindEx (BYREF cbsSourceString AS CBSTR, BYREF cbsPattern AS CBSTR, _
   BYVAL bIgnoreCase AS BOOLEAN = FALSE, BYVAL bGlobal AS BOOLEAN = TRUE, BYVAL bMultiline AS BOOLEAN = TRUE) AS CBSTR

   IF m_pRegExp = NULL THEN this.SetResult(E_POINTER): RETURN ""
   m_pRegExp->put_IgnoreCase(bIgnoreCase)
   m_pRegExp->put_Global(bGlobal)
   m_pRegExp->put_Multiline(bMultiline)
   m_pRegExp->put_Pattern(cbsPattern)

   DIM cbsOut AS CBSTR, pMatches AS Afx_IMatchCollection2 PTR
   this.SetResult(m_pRegExp->Execute(cbsSourceString, cast(Afx_IDispatch PTR PTR, @pMatches)))
   IF pMatches THEN
      DIM nCount AS LONG
      pMatches->get_Count(@nCount)
      FOR i AS LONG = 0 TO nCount - 1
         DIM pMatch AS Afx_IMatch2 PTR
         this.SetResult(pMatches->get_Item(i, cast(Afx_IDispatch PTR PTR, @pMatch)))
         IF pMatch THEN
            DIM nFirstIndex AS LONG
            pMatch->get_FirstIndex(@nFirstIndex)
            DIM nLen AS LONG
            pMatch->get_Length(@nLen)
            IF i < nCount - 1 THEN
               cbsOut += STR(nFirstIndex + 1) & "," & STR(nLen) & ";"
            ELSE
               cbsOut += STR(nFirstIndex + 1) & "," & STR(nLen)
            END IF
            AfxSafeRelease(pMatch)
         END IF
      NEXT
      AfxSafeRelease(pMatches)
   END IF
   RETURN cbsOut

END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns a copy of a string with text replaced using a regular expression as the search string.
' Parameters:
' - cbsSourceString : The source string.
' - cbsPattern : The pattern to search.
' - cvReplaceString : The replacement string.
' - bIgnoreCase : FALSE = case sensitive; TRUE = case insensitive.
' - bGlobal : FALSE = Delete only the first match; TRUE = delete all matches.
' - bMultiline : TRUE = Match at the start and the end of multiple lines separated by line breaks.
' Examples:
' DIM pRegExp AS CRegExp
' PRINT pRegExp.Replace("Hello World", "World", "Earth") - prints "Hello Earth"
' PRINT pRegExp.Replace("abacadabra", "[bac]", "*") - prints "*****d**r*"
' PRINT pRegExp.Replace("555-123-4567", "(\d{3})-(\d{3})-(\d{4})", "($1) $2-$3") - prints "(555) 123-4567"
' PRINT pRegExp.Replace("Squires, Paul", "(\S+), (\S+)", "$2 $1") - prints "Paul Squires"
' PRINT pRegExp.Replace("0000.34500044", $"\b0{1,}\.", ".") - prints ".34500044"
' ========================================================================================
PRIVATE FUNCTION CRegExp.Replace (BYREF cbsSourceString AS CBSTR, BYREF cvReplaceString AS CVAR, _
   BYVAL bIgnoreCase AS BOOLEAN = FALSE, BYVAL bGlobal AS BOOLEAN = TRUE, BYVAL bMultiline AS BOOLEAN = FALSE) AS CBSTR
   DIM bstrDestString AS AFX_BSTR
   IF m_pRegExp = NULL THEN this.SetResult(E_POINTER): RETURN bstrDestString
   m_pRegExp->put_IgnoreCase(bIgnoreCase)
   m_pRegExp->put_Global(bGlobal)
   m_pRegExp->put_Multiline(bMultiline)
   this.SetResult(m_pRegExp->Replace(cbsSourceString, cvReplaceString, @bstrDestString))
   RETURN bstrDestString
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CRegExp.Replace (BYREF cbsSourceString AS CBSTR, BYREF cbsPattern AS CBSTR, BYREF cvReplaceString AS CVAR, _
   BYVAL bIgnoreCase AS BOOLEAN = FALSE, BYVAL bGlobal AS BOOLEAN = TRUE, BYVAL bMultiline AS BOOLEAN = FALSE) AS CBSTR
   DIM bstrDestString AS AFX_BSTR
   IF m_pRegExp = NULL THEN this.SetResult(E_POINTER): RETURN bstrDestString
   m_pRegExp->put_IgnoreCase(bIgnoreCase)
   m_pRegExp->put_Global(bGlobal)
   m_pRegExp->put_Multiline(bMultiline)
   m_pRegExp->put_Pattern(cbsPattern)
   this.SetResult(m_pRegExp->Replace(cbsSourceString, cvReplaceString, @bstrDestString))
   RETURN bstrDestString
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns a copy of a string with text removed using a regular expression as the search string.
' Parameters:
' - cbsSourceString : The source string.
' - cbsPattern : The pattern to search.
' - bIgnoreCase : FALSE = case sensitive; TRUE = case insensitive.
' - bGlobal : FALSE = Delete only the first match; TRUE = delete all matches.
' - bMultiline : TRUE = Match at the start and the end of multiple lines separated by line breaks.
' Examples:
' DIM pRegExp AS CRegExp
' PRINT pRegExp.Remove("abacadabra", "ab") - prints "acadra"
' PRINT pRegExp.Remove("abacadabra", "[bAc]", TRUE) - prints "dr"
' ========================================================================================
PRIVATE FUNCTION CRegExp.Remove (BYREF cbsSourceString AS CBSTR, BYVAL bIgnoreCase AS BOOLEAN = FALSE, _
   BYVAL bGlobal AS BOOLEAN = TRUE, BYVAL bMultiline AS BOOLEAN = FALSE) AS CBSTR
   DIM bstrDestString AS AFX_BSTR
   IF m_pRegExp = NULL THEN this.SetResult(E_POINTER): RETURN bstrDestString
   m_pRegExp->put_IgnoreCase(bIgnoreCase)
   m_pRegExp->put_Global(bGlobal)
   m_pRegExp->put_Multiline(bMultiline)
   this.SetResult(m_pRegExp->Replace(cbsSourceString, CVAR(""), @bstrDestString))
   RETURN bstrDestString
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CRegExp.Remove (BYREF cbsSourceString AS CBSTR, BYREF cbsPattern AS CBSTR, _
   BYVAL bIgnoreCase AS BOOLEAN = FALSE, BYVAL bGlobal AS BOOLEAN = TRUE, BYVAL bMultiline AS BOOLEAN = FALSE) AS CBSTR
   DIM bstrDestString AS AFX_BSTR
   IF m_pRegExp = NULL THEN this.SetResult(E_POINTER): RETURN bstrDestString
   m_pRegExp->put_IgnoreCase(bIgnoreCase)
   m_pRegExp->put_Global(bGlobal)
   m_pRegExp->put_Multiline(bMultiline)
   m_pRegExp->put_Pattern(cbsPattern)
   this.SetResult(m_pRegExp->Replace(cbsSourceString, CVAR(""), @bstrDestString))
   RETURN bstrDestString
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts a substring using VBScript regular expressions search patterns.
' Parameters:
' - cbsSourceString = The text to be parsed.
' - cbsPattern = The pattern to match.
' - bIgnoreCase = Ignore case.
' Return value:
' - Returns the retrieved string on exit or an empty string on failure.
' Usage example:
'   DIM pRegExp AS CRegExp
'   DIM cbsText AS CBSTR = "blah blah a234 blah blah x345 blah blah"
'   DIM cbsPattern AS CBSTR = "[a-z][0-9][0-9][0-9]"
'   DIM cbs AS CBSTR = pRegExp.Extract(cbsText, cbsPattern)
' ========================================================================================
PRIVATE FUNCTION CRegExp.Extract (BYREF cbsSourceString AS CBSTR, BYVAL bIgnoreCase AS BOOLEAN = FALSE, _
   BYVAL bGlobal AS BOOLEAN = FALSE, BYVAL bMultiline AS BOOLEAN = FALSE) AS CBSTR
   IF m_pRegExp = NULL THEN this.SetResult(E_POINTER): RETURN ""
   IF m_pMatches THEN AfxSafeRelease(m_pMatches)
   m_pRegExp->put_IgnoreCase(bIgnoreCase)
   m_pRegExp->put_Global(bGlobal)
   m_pRegExp->put_Multiline(bMultiline)
   this.SetResult(m_pRegExp->Execute(cbsSourceString, cast(Afx_IDispatch PTR PTR, @m_pMatches)))
   DIM nPos AS LONG = this.MatchPos(0)
   DIM nLen AS LONG = this.MatchLen(0)
   DIM cbs AS CBSTR = MID(cbsSourceString, nPos, nLen)
   RETURN cbs
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CRegExp.Extract (BYREF cbsSourceString AS CBSTR, BYREF cbsPattern AS CBSTR, BYVAL bIgnoreCase AS BOOLEAN = FALSE, _
   BYVAL bGlobal AS BOOLEAN = FALSE, BYVAL bMultiline AS BOOLEAN = FALSE) AS CBSTR
   IF m_pRegExp = NULL THEN this.SetResult(E_POINTER): RETURN ""
   IF m_pMatches THEN AfxSafeRelease(m_pMatches)
   m_pRegExp->put_IgnoreCase(bIgnoreCase)
   m_pRegExp->put_Global(bGlobal)
   m_pRegExp->put_Multiline(bMultiline)
   m_pRegExp->put_Pattern(cbsPattern)
   this.SetResult(m_pRegExp->Execute(cbsSourceString, cast(Afx_IDispatch PTR PTR, @m_pMatches)))
   DIM nPos AS LONG = this.MatchPos(0)
   DIM nLen AS LONG = this.MatchLen(0)
   DIM cbs AS CBSTR = MID(cbsSourceString, nPos, nLen)
   RETURN cbs
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts a substring using VBScript regular expressions search patterns.
' Parameters:
' - nStart: The position in the string at which the search will begin. The first character
'   starts at position 1.
' - cbsSourceString = The text to be parsed.
' - cbsPattern = The pattern to match.
' - bIgnoreCase = Ignore case.
' Return value:
' - Returns the position of the match or 0 if not found.
'   The length can be retrieves calling pRegExp.MatchLen.
' Usage example:
'   DIM pRegExp AS CRegExp
'   DIM cbsText AS CBSTR = "blah blah a234 blah blah x345 blah blah"
'   DIM cbsPattern AS CBSTR = "[a-z][0-9][0-9][0-9]"
'   DIM cbs AS CBSTR = pRegExp.Extract(15, cbsText, cbsPattern)
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CRegExp.Extract (BYVAL nStart AS LONG, BYREF cbsSourceString AS CBSTR, BYVAL bIgnoreCase AS BOOLEAN = FALSE, _
   BYVAL bGlobal AS BOOLEAN = FALSE, BYVAL bMultiline AS BOOLEAN = FALSE) AS CBSTR
   IF m_pRegExp = NULL THEN this.SetResult(E_POINTER): RETURN ""
   IF m_pMatches THEN AfxSafeRelease(m_pMatches)
   m_pRegExp->put_IgnoreCase(bIgnoreCase)
   m_pRegExp->put_Global(bGlobal)
   m_pRegExp->put_Multiline(bMultiline)
   DIM pbstr AS BSTR
   IF nStart < 0 OR nStart > SysStringLen(cbsSourceString.m_bstr) THEN RETURN ""
   DIM cbs AS CBSTR = cbsSourceString.MidChars(nStart)
   this.SetResult(m_pRegExp->Execute(cbs, cast(Afx_IDispatch PTR PTR, @m_pMatches)))
   DIM nPos AS LONG = this.MatchPos(0)
   DIM nLen AS LONG = this.MatchLen(0)
   RETURN MID(cbsSourceString, nPos + nStart - 1, nLen)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CRegExp.Extract (BYVAL nStart AS LONG, BYREF cbsSourceString AS CBSTR, BYREF cbsPattern AS CBSTR, BYVAL bIgnoreCase AS BOOLEAN = FALSE, _
   BYVAL bGlobal AS BOOLEAN = FALSE, BYVAL bMultiline AS BOOLEAN = FALSE) AS CBSTR
   IF m_pRegExp = NULL THEN this.SetResult(E_POINTER): RETURN ""
   IF m_pMatches THEN AfxSafeRelease(m_pMatches)
   m_pRegExp->put_IgnoreCase(bIgnoreCase)
   m_pRegExp->put_Global(bGlobal)
   m_pRegExp->put_Multiline(bMultiline)
   m_pRegExp->put_Pattern(cbsPattern)
   DIM pbstr AS BSTR
   IF nStart < 0 OR nStart > SysStringLen(cbsSourceString.m_bstr) THEN RETURN ""
   DIM cbs AS CBSTR = cbsSourceString.MidChars(nStart)
   this.SetResult(m_pRegExp->Execute(cbs, cast(Afx_IDispatch PTR PTR, @m_pMatches)))
   DIM nPos AS LONG = this.MatchPos(0)
   DIM nLen AS LONG = this.MatchLen(0)
   RETURN MID(cbsSourceString, nPos + nStart - 1, nLen)
END FUNCTION
' ========================================================================================

END NAMESPACE
