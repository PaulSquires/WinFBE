' ########################################################################################
' Microsoft Windows
' File: AfxSapi.inc
' Contents: Microsoft SAPI 5.xx.
' Portions Copyright (c) Microsoft Corporation.
' Compiler: Free Basic 32 & 64 bit
' Copyright (c) 2017 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "Afx/AfxSapi.bi"
#include once "Afx/CDispInvoke.inc"
using Afx

NAMESPACE Afx

' ========================================================================================
' Retrieves the collection of voices installed separated with a carriage returns and line feed.
' Returns a CRLF separated list of names. There if not CRLF in the last name of the lst
' (or in the first, if only one voice is installed).
' ========================================================================================
PRIVATE FUNCTION AfxSpGetVoices () AS CWSTR

   DIM cwsVoices AS CWSTR
   ' // Get a reference to an enumerator for the voices collection
   DIM hr AS HRESULT
   DIM pEnum AS Afx_IEnumSpObjectTokens PTR
   DIM pCategory AS Afx_ISpObjectTokenCategory PTR
   pCategory = AfxNewCom(AFX_CLSID_SpObjectTokenCategory)
   IF pCategory = NULL THEN SetLastError(E_NOINTERFACE) : RETURN ""
   DIM wszCategoryId AS WSTRING * 260 = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\Voices"
   hr = pCategory->SetId(@wszCategoryId, FALSE)
   IF FAILED(hr) THEN AfxSafeRelease(pCategory) : RETURN ""
   hr = pCategory->EnumTokens(NULL, NULL, @pEnum)
   AfxSafeRelease(pCategory)
   IF pEnum = NULL THEN RETURN ""

   ' // Parse the collection of tokens
   DIM nCount AS LONG
   pEnum->GetCount(@nCount)
   DIM pToken AS Afx_ISpObjectToken PTR
   DIM wszLangId AS WSTRING * 10 = HEX(GetUserDefaultUILanguage(), 4)
   FOR i AS LONG = 0 TO nCount - 1
      pEnum->Item(i, @pToken)
      IF pToken = NULL THEN EXIT FOR
      DIM pwszDescription AS WSTRING PTR
      hr = pToken->GetStringValue(@wszLangId, @pwszDescription)
      IF hr = SPERR_NOT_FOUND THEN
         hr = pToken->GetStringValue(NULL, @pwszDescription)
      END IF
      AfxSafeRelease(pToken)
      IF hr = S_OK AND pwszDescription <> NULL THEN
         cwsVoices += *pwszDescription & IIF(i < nCount -1, CHR(13, 10), "")
         CoTaskMemFree pwszDescription
      END IF
   NEXT

   ' // Release the enumerator
   AfxSafeRelease(pEnum)
   ' // Returns the retrieved voices
   RETURN cwsVoices

END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the voice used by the speech engine.
' wszVoice = Voice to set, e.g. "Microsoft Mary".
' ========================================================================================
PRIVATE FUNCTION AfxSpSetVoice (BYREF wszVoice AS WSTRING) AS HRESULT

   ' // Create an instance of the SpVoice object
   DIM pSpVoice AS Afx_ISpVoice PTR = AfxNewCom("SAPI.SpVoice")
   IF pSpVoice = NULL THEN RETURN E_NOINTERFACE

   ' // Get a reference to an enumerator for the voices collection
   DIM hr AS HRESULT
   DIM pEnum AS Afx_IEnumSpObjectTokens PTR
   DIM pCategory AS Afx_ISpObjectTokenCategory PTR
   pCategory = AfxNewCom(AFX_CLSID_SpObjectTokenCategory)
   IF pCategory = NULL THEN RETURN E_NOINTERFACE
   DIM wszCategoryId AS WSTRING * 260 = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\Voices"
   hr = pCategory->SetId(@wszCategoryId, FALSE)
   IF FAILED(hr) THEN AfxSafeRelease(pCategory) : RETURN hr
   hr = pCategory->EnumTokens(NULL, NULL, @pEnum)
   AfxSafeRelease(pCategory)
   IF pEnum = NULL THEN RETURN E_NOINTERFACE

   ' // Parse the collection of tokens
   DIM nCount AS LONG
   pEnum->GetCount(@nCount)
   DIM pToken AS Afx_ISpObjectToken PTR
   DIM wszLangId AS WSTRING * 10 = HEX(GetUserDefaultUILanguage(), 4)
   FOR i AS LONG = 0 TO nCount - 1
      pEnum->Item(i, @pToken)
      IF pToken = NULL THEN EXIT FOR
      DIM pwszDescription AS WSTRING PTR
      hr = pToken->GetStringValue(@wszLangId, @pwszDescription)
      IF hr = SPERR_NOT_FOUND THEN
         hr = pToken->GetStringValue(NULL, @pwszDescription)
      END IF
      AfxSafeRelease(pToken)
      IF hr = S_OK AND pwszDescription <> NULL THEN
         IF UCASE(*pwszDescription) = UCASE(wszVoice) THEN
            pSpVoice->SetVoice(pToken)
            CoTaskMemFree pwszDescription
            EXIT FOR
         END IF
         CoTaskMemFree pwszDescription
      END IF
   NEXT

   ' // Release the enumerator
   AfxSafeRelease(pEnum)
   ' // Returns the result code
   RETURN hr

END FUNCTION
' ========================================================================================

END NAMESPACE
