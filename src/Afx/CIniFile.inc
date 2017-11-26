' ########################################################################################
' Microsoft Windows
' File: CIniFile.inc
' Contents: Class to work with Windows .ini files
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2017 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "windows.bi"
#include once "crt/stdio.bi"
#include once "Afx/CSafeArray.inc"
#include once "Afx/CDicObj.inc"
USING Afx

NAMESPACE Afx

' ########################################################################################
' CIniFile class.
' ########################################################################################
TYPE CIniFile

   m_Path AS WSTRING * MAX_PATH   ' // Full path of the .ini file

   DECLARE CONSTRUCTOR (BYREF wszFileName AS WSTRING)
   DECLARE DESTRUCTOR
   DECLARE FUNCTION GetPath () AS CWSTR
   DECLARE FUNCTION WriteValue (BYREF wszSectionName AS WSTRING, BYREF wszKeyName AS WSTRING, BYREF wszValue AS WSTRING) AS BOOLEAN
   DECLARE FUNCTION WriteValue (BYREF wszSectionName AS WSTRING, BYREF wszKeyName AS WSTRING, BYREF dblValue AS DOUBLE) AS BOOLEAN
   DECLARE FUNCTION WriteValue (BYREF wszSectionName AS WSTRING, BYREF wszKeyName AS WSTRING, BYREF intValue AS LONG) AS BOOLEAN
   DECLARE FUNCTION DeleteKey (BYREF wszSectionName AS WSTRING, BYREF wszKeyName AS WSTRING) AS BOOLEAN
   DECLARE FUNCTION DeleteSection (BYREF wszSectionName AS WSTRING) AS BOOLEAN
   DECLARE FUNCTION GetString (BYREF wszSectionName AS WSTRING, BYREF wszKeyName AS WSTRING, BYREF wszDefaultValue AS WSTRING = "") AS CWSTR
   DECLARE FUNCTION GetDouble (BYREF wszSectionName AS WSTRING, BYREF wszKeyName AS WSTRING, BYVAL nDefaultValue AS DOUBLE = 0) AS DOUBLE
   DECLARE FUNCTION GetInt (BYREF wszSectionName AS WSTRING, BYREF wszKeyName AS WSTRING, BYVAL nDefaultValue AS LONG = 0) AS LONG
   DECLARE FUNCTION GetSectionNames () AS CSafeARray
   DECLARE FUNCTION GetKeyNames (BYREF wszSectionName AS WSTRING) AS CSafeARray
   DECLARE FUNCTION GetSectionValues (BYREF wszSectionName AS WSTRING, BYREF pDic AS CDicObj) AS BOOLEAN

END TYPE

' ========================================================================================
' Default constructor
' Initializes a new instance of the CIniFile class.
' - wszFileName: The ini file to read and write from.
' ========================================================================================
PRIVATE CONSTRUCTOR CIniFile (BYREF wszFileName AS WSTRING)
   ' // Convert to the full path. Because of backward compatibility, 
   ' // the win32 functions tend to assume the path should be the 
   ' // root Windows directory if it is not specified. By calling 
   ' // GetFullPath, we make sure we are always passing the full path
   ' // the win32 functions.
   DIM nLen AS LONG, buffer AS WSTRING * 4096
   nLen = .GetFullPathNameW(wszFileName, SIZEOF(buffer) \ 2, buffer, NULL)
   IF nLen THEN m_Path = LEFT(buffer, nLen)
   ' // Even if we use WritePrivateProfileStringW, it will only write unicode text if the
   ' // file has been created using UTF 16 little endian. The solution is to create the
   ' // ini-file with the encoding UTF-16LE before writing an unicode string to it.
   ' // Make sure the file does not already exist
   IF AfxFileExists(m_Path) = FALSE THEN
      ' // create file with encoding UTF-16LE
      DIM fileHandle AS FILE PTR
      DIM wszMode AS WSTRING * 260 = "w, ccs=UTF-16LE"
      fileHandle = _wfopen(@m_Path, @wszMode)
      IF fileHandle THEN fclose(fileHandle)
   END IF
END CONSTRUCTOR
' ========================================================================================
 
' ========================================================================================
' Destructor
' ========================================================================================
PRIVATE DESTRUCTOR CIniFile
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the full path of ini file this object instance is operating on.
' ========================================================================================
PRIVATE FUNCTION CIniFile.GetPath () AS CWSTR
   RETURN m_Path
END FUNCTION
' ========================================================================================

' ========================================================================================
' Copies a value into the specified section of an initialization file.
' - wszSectionName: Name of the section
' - wszKeyName: Name of key
' - wszValue / dblValue / intValue: The value to write
' ========================================================================================
PRIVATE FUNCTION CIniFile.WriteValue (BYREF wszSectionName AS WSTRING, BYREF wszKeyName AS WSTRING, BYREF wszValue AS WSTRING) AS BOOLEAN
   RETURN WritePrivateProfileStringW(wszSectionName, wszKeyName, wszValue, m_path)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CIniFile.WriteValue (BYREF wszSectionName AS WSTRING, BYREF wszKeyName AS WSTRING, BYREF dblValue AS DOUBLE) AS BOOLEAN
   RETURN WritePrivateProfileStringW(wszSectionName, wszKeyName, WSTR(dblValue), m_path)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CIniFile.WriteValue (BYREF wszSectionName AS WSTRING, BYREF wszKeyName AS WSTRING, BYREF intValue AS LONG) AS BOOLEAN
   RETURN WritePrivateProfileStringW(wszSectionName, wszKeyName, WSTR(intValue), m_path)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Retrieves a string from the specified section in an initialization file.
' - wszSectionName: Name of the section
' - wszKeyName: Name of key
' - wszDefaultValue: A default string. If the key key cannot be found in the initialization
'   file, the default string is returned.
'   Avoid specifying a default string with trailing blank characters. The function inserts
'   a null character in the returned buffer to strip any trailing blanks.
' ========================================================================================
PRIVATE FUNCTION CIniFile.GetString (BYREF wszSectionName AS WSTRING, BYREF wszKeyName AS WSTRING, BYREF wszDefaultValue AS WSTRING = "") AS CWSTR
   DIM wsz AS WSTRING * 32767
   DIM dwLen AS DWORD = GetPrivateProfileStringW(wszSectionName, wszKeyName, wszDefaultValue, @wsz, 32767, m_path)
   RETURN LEFT(wsz, dwLen)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CIniFile.GetDouble (BYREF wszSectionName AS WSTRING, BYREF wszKeyName AS WSTRING, BYVAL nDefaultValue AS DOUBLE = 0) AS DOUBLE
   DIM wsz AS WSTRING * 32767
   DIM dwLen AS DWORD = GetPrivateProfileStringW(wszSectionName, wszKeyName, NULL, @wsz, 32767, m_path)
   wsz = LEFT(wsz, dwLen)
   IF VAL(wsz) = 0 THEN RETURN nDefaultValue ELSE RETURN VAL(wsz)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CIniFile.GetInt (BYREF wszSectionName AS WSTRING, BYREF wszKeyName AS WSTRING, BYVAL nDefaultValue AS LONG = 0) AS LONG
   RETURN CLNG(GetPrivateProfileInt(wszSectionName, wszKeyName, nDefaultValue, m_path))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Deletes a key from the specified section of an initialization file.
' - wszSectionName: Name of the section
' - wszKeyName: Name of key
' ========================================================================================
PRIVATE FUNCTION CIniFile.DeleteKey (BYREF wszSectionName AS WSTRING, BYREF wszKeyName AS WSTRING) AS BOOLEAN
   RETURN WritePrivateProfileStringW(wszSectionName, wszKeyName, NULL, m_path)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Deletes a section from an initialization file.
' - wszSectionName: Name of the section
' ========================================================================================
PRIVATE FUNCTION CIniFile.DeleteSection (BYREF wszSectionName AS WSTRING) AS BOOLEAN
   RETURN WritePrivateProfileStringW(wszSectionName, NULL, NULL, m_path)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a safe array with the names of all sections in the ini file.
' Example:
' DIM cIni AS CInifile = "Test.ini"
' DIM csa AS CSafeArray = cIni.GetSectionNames
' FOR i AS LONG = csa.LBound TO csa.UBound
'    print csa.GetStr(i)
' NEXT
' ========================================================================================
PRIVATE FUNCTION CIniFile.GetSectionNames () AS CSafeArray
   DIM csa AS CSafeArray = CSafeArray(VT_BSTR, 0, 1)
   DIM wsz AS WSTRING * 32767
   DIM dwLen AS DWORD = GetPrivateProfileSectionNamesW(@wsz, 32767, m_path)
   IF dwLen = 0 THEN RETURN csa
   DIM pwsz AS WSTRING PTR = @wsz
   DO
      IF pwsz = NULL THEN EXIT DO
      csa.AppendStr(pwsz)
      dwLen = LEN(*pwsz)
      IF dwLen = 0 THEN EXIT DO
      pwsz += dwLen + 1
   LOOP
   RETURN csa
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a safe array with the names of all the keys of the specified section.
' Example:
' DIM cIni AS CInifile = "Test.ini"
' DIM csa AS CSafeArray = cIni.GetKeyNames("Startup")
' FOR i AS LONG = csa.LBound TO csa.UBound
'    print csa.GetStr(i)
' NEXT
' ========================================================================================
PRIVATE FUNCTION CIniFile.GetKeyNames (BYREF wszSectionName AS WSTRING) AS CSafeARray
   DIM csa AS CSafeArray = CSafeArray(VT_BSTR, 0, 1)
   DIM wsz AS WSTRING * 32767
   DIM dwLen AS DWORD = GetPrivateProfileStringW(@wszSectionName, NULL, NULL, @wsz, 32767, m_path)
   IF dwLen = 0 THEN RETURN csa
   DIM pwsz AS WSTRING PTR = @wsz
   DO
      IF pwsz = NULL THEN EXIT DO
      csa.AppendStr(pwsz)
      dwLen = LEN(*pwsz)
      IF dwLen = 0 THEN EXIT DO
      pwsz += dwLen + 1
   LOOP
   RETURN csa
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the keys and values of the specified section as a dictionary object.
' ========================================================================================
PRIVATE FUNCTION CIniFile.GetSectionValues (BYREF wszSectionName AS WSTRING, BYREF pDic AS CDicObj) AS BOOLEAN
   IF pDic.m_pDictionary = NULL THEN RETURN FALSE
   DIM wsz AS WSTRING * 32767
   DIM dwLen AS DWORD = GetPrivateProfileStringW(@wszSectionName, NULL, NULL, @wsz, 32767, m_path)
   IF dwLen = 0 THEN RETURN FALSE
   DIM cwsKeyName AS CWSTR, cwsValue AS CWSTR
   DIM pwsz AS WSTRING PTR = @wsz
   DO
      IF pwsz = NULL THEN EXIT DO
      cwsKeyName = pwsz
      cwsValue = this.GetString(wszSectionName, cwsKeyName)
      pDic.Add(cwsKeyName, cwsValue)
      dwLen = LEN(*pwsz)
      IF dwLen = 0 THEN EXIT DO
      pwsz += dwLen + 1
   LOOP
   RETURN TRUE
END FUNCTION
' ========================================================================================

END NAMESPACE
