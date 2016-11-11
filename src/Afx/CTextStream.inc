' ########################################################################################
' Microsoft Windows
' File: CTextStream.inc
' Contents: Wrapper class for reading and writing text files.
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "windows.bi"
#include once "Afx/AfxScrRun.bi"
#include once "Afx/CWStr.inc"

NAMESPACE Afx

' ########################################################################################
' CTextStream - Class for reading and writing text files.
' ########################################################################################
TYPE CTextStream

   m_pFileSys AS Afx_IFileSystem PTR
   m_pTxtStm AS Afx_ITextStream PTR
   m_Result AS HRESULT
   m_bUninitCOM AS BOOLEAN

   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR
   DECLARE FUNCTION GetLastResult () AS HRESULT
   DECLARE FUNCTION SetResult (BYVAL Result AS HRESULT) AS HRESULT
   DECLARE FUNCTION Create (BYREF cwsFileName AS CWSTR, BYVAL bOverwrite AS BOOLEAN = TRUE, BYVAL bUnicode AS BOOLEAN = FALSE) AS HRESULT
   DECLARE FUNCTION Open (BYREF cwsFileName AS CWSTR, BYVAL IOMode AS LONG = 1, BYVAL bCreate AS BOOLEAN = FALSE, BYVAL bUnicode AS BOOLEAN = FALSE) AS HRESULT
   DECLARE FUNCTION OpenUnicode (BYREF cwsFileName AS CWSTR, BYVAL IOMode AS LONG = 1, BYVAL bCreate AS BOOLEAN = FALSE) AS HRESULT
   DECLARE FUNCTION Close () AS HRESULT
   DECLARE PROPERTY Line () AS LONG
   DECLARE PROPERTY Column () AS LONG
   DECLARE FUNCTION EOS () AS BOOLEAN
   DECLARE FUNCTION EOL () AS BOOLEAN
   DECLARE FUNCTION Read (BYVAL numChars AS LONG) AS CWSTR
   DECLARE FUNCTION ReadLine () AS CWSTR
   DECLARE FUNCTION ReadAll () AS CWSTR
   DECLARE FUNCTION Write (BYREF cwsText AS CWSTR) AS HRESULT
   DECLARE FUNCTION WriteLine (BYREF cwsText AS CWSTR) AS HRESULT
   DECLARE FUNCTION WriteBlankLines (BYVAL numLines AS LONG) AS HRESULT
   DECLARE FUNCTION Skip (BYVAL numChars AS LONG) AS HRESULT
   DECLARE FUNCTION SkipLine () AS HRESULT

END TYPE
' ########################################################################################

' ========================================================================================
' Constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CTextStream
   ' // Initialize the COM library
   DIM hr AS HRESULT = CoInitialize(NULL)
   IF hr = S_OK OR hr = S_FALSE THEN m_bUninitCOM = TRUE
   ' // Create an instance of the Filesystem object
   DIM CLSID_FileSystemObject_ AS GUID = (&h0D43FE01, &hF093, &h11CF, {&h89, &h40, &h00, &hA0, &hC9, &h05, &h42, &h28})
   DIM IID_IFileSystem AS GUID = (&h0AB5A3D0, &hE5B6, &h11D0, {&hAB, &hF5, &h00, &hA0, &hC9, &h0F, &hFF, &hC0})
   SetResult(CoCreateInstance(@CLSID_FileSystemObject_, NULL, CLSCTX_INPROC_SERVER, @IID_IFileSystem, @m_pFileSys))
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Destructor
' ========================================================================================
PRIVATE DESTRUCTOR CTextStream
   IF m_pTxtStm THEN
      m_pTxtStm->Close
      m_pTxtStm->Release
   END IF
   IF m_pFileSys THEN m_pFileSys->Release
   ' // Free the COM library
   IF m_bUninitCOM THEN CoUninitialize
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the last status code.
' ========================================================================================
PRIVATE FUNCTION CTextStream.GetLastResult () AS HRESULT
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the last status code.
' ========================================================================================
PRIVATE FUNCTION CTextStream.SetResult (BYVAL Result AS HRESULT) AS HRESULT
   m_Result = Result
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates a specified file name and returns a TextStream object that can be used to read
' from or write to the file.
' Parameters:
' - bstrFileName: String expression that identifies the file to create.
' - bOverwrite: Boolean value that indicates whether you can overwrite an existing file.
'   The value is true if the file can be overwritten, false if it can't be overwritten.
'   If omitted, existing files are not overwritten.
' - bUnicode: Boolean value that indicates whether the file is created as a Unicode or
'   ASCII file. The value is true if the file is created as a Unicode file, false if it's
'   created as an ASCII file. If omitted, an ASCII file is assumed.
' ========================================================================================
PRIVATE FUNCTION CTextStream.Create (BYREF cwsFileName AS CWSTR, BYVAL bOverwrite AS BOOLEAN = TRUE, BYVAL bUnicode AS BOOLEAN = FALSE) AS HRESULT
   IF m_pTxtStm THEN
      SetResult(m_pTxtStm->Close)
      m_pTxtStm->Release
      m_pTxtStm = NULL
   END IF
   DIM bstrFileName AS AFX_BSTR = SysAllocString(**cwsFileName)
   SetResult(m_pFileSys->CreateTextFile(bstrFileName, bOverwrite, bUnicode, @m_pTxtStm))
   SysFreeString bstrFileName
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Opens a specified file and returns a TextStream object that can be used to read from,
' write to, or append to the file.
' Parameters:
' - bstrFileName: String expression that identifies the file to open.
' - IOMode: Can be one of three constants: IOMODE_ForReading, IOMODE_ForWriting, or IOMODE_ForAppending.
' - bCreate: Value that indicates whether a new file can be created if the specified
'   filename doesn't exist. The value is True if a new file is created, False if it isn't
'   created. If omitted, a new file isn't created.
' - bUnicode: TRUE or FALSE. Indicates the format of the opened file. If omitted, the file
'   is opened as ASCII.
' ========================================================================================
PRIVATE FUNCTION CTextStream.Open (BYREF cwsFileName AS CWSTR, BYVAL IOMode AS LONG = 1, BYVAL bCreate AS BOOLEAN = FALSE, BYVAL bUnicode AS BOOLEAN = FALSE) AS HRESULT
   IF m_pTxtStm THEN
      SetResult(m_pTxtStm->Close)
      m_pTxtStm->Release
      m_pTxtStm = NULL
   END IF
   DIM bstrFileName AS AFX_BSTR = SysAllocString(**cwsFileName)
   SetResult(m_pFileSys->OpenTextFile(bstrFileName, IOMode, bCreate, bUnicode, @m_pTxtStm))
   SysFreeString bstrFileName
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Opens a specified file and returns a TextStream object that can be used to read from,
' write to, or append to the file.
' Parameters:
' - bstrFileName: String expression that identifies the file to open.
' - IOMode: Can be one of three constants: IOMODE_ForReading, IOMODE_ForWriting, or IOMODE_ForAppending.
' - bCreate: Value that indicates whether a new file can be created if the specified
'   filename doesn't exist. The value is True if a new file is created, False if it isn't
'   created. If omitted, a new file isn't created.
' ========================================================================================
PRIVATE FUNCTION CTextStream.OpenUnicode (BYREF cwsFileName AS CWSTR, BYVAL IOMode AS LONG = 1, BYVAL bCreate AS BOOLEAN = FALSE) AS HRESULT
   IF m_pTxtStm THEN
      SetResult(m_pTxtStm->Close)
      m_pTxtStm->Release
      m_pTxtStm = NULL
   END IF
   DIM bstrFileName AS AFX_BSTR = SysAllocString(**cwsFileName)
   SetResult(m_pFileSys->OpenTextFile(bstrFileName, IOMode, bCreate, Tristate_True, @m_pTxtStm))
   SysFreeString bstrFileName
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Closes an open TextStream file.
' ========================================================================================
PRIVATE FUNCTION CTextStream.Close () AS HRESULT
   IF m_pTxtStm THEN
      RETURN(SetResult(m_pTxtStm->Close))
      m_pTxtStm = NULL
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Read-only property that returns the current line number in a TextStream file.
' After a file is initially opened and before anything is written, Line is equal to 1.
' ========================================================================================
PRIVATE PROPERTY CTextStream.Line () AS LONG
   DIM nLine AS LONG
   IF m_pTxtStm THEN SetResult(m_pTxtStm->get_Line(@nLine))
   PROPERTY = nLine
END PROPERTY
' ========================================================================================

' ========================================================================================
' Read-only property that returns the column number of the current character position in a
' TextStream file. After a newline character has been written, but before any other character
' is written, Column is equal to 1.
' ========================================================================================
PRIVATE PROPERTY CTextStream.Column () AS LONG
   DIM nColumn AS LONG
   IF m_pTxtStm THEN SetResult(m_pTxtStm->get_Column(@nColumn))
   PROPERTY = nColumn
END PROPERTY
' ========================================================================================

' ========================================================================================
' Returns TRUE if the file pointer is at the end of a TextStream file; FALSE if it is not. Read-only.
' ========================================================================================
PRIVATE FUNCTION CTextStream.EOS () AS BOOLEAN
   DIM nEOS AS VARIANT_BOOL
   IF m_pTxtStm THEN SetResult(m_pTxtStm->get_AtEndOfStream(@nEOS))
   RETURN nEOS
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns TRUE if the file pointer is positioned immediately before the end-of-line marker
' in a TextStream file; FALSE if it is not. Read-only.
' ========================================================================================
PRIVATE FUNCTION CTextStream.EOL () AS BOOLEAN
   DIM nEOL AS VARIANT_BOOL
   IF m_pTxtStm THEN SetResult(m_pTxtStm->get_AtEndOfLine(@nEOL))
   RETURN nEOL
END FUNCTION
' ========================================================================================

' ========================================================================================
' Reads a specified number of characters from a TextStream file and returns the resulting string.
' After a file is initially opened and before anything is written, Line is equal to 1.
' ========================================================================================
PRIVATE FUNCTION CTextStream.Read (BYVAL numChars AS LONG) AS CWSTR
   DIM bstrText AS AFX_BSTR
   IF m_pTxtStm THEN SetResult(m_pTxtStm->Read(numChars, @bstrText))
   DIM cwsText AS CWSTR = SysAllocString(*bstrText)
   SysFreeString bstrText
   RETURN cwsText
END FUNCTION
' ========================================================================================

' ========================================================================================
' Reads an entire line (up to, but not including, the newline character) from a TextStream
' file and returns the resulting string.
' ========================================================================================
PRIVATE FUNCTION CTextStream.ReadLine () AS CWSTR
   DIM bstrText AS AFX_BSTR
   IF m_pTxtStm THEN SetResult(m_pTxtStm->ReadLine(@bstrText))
   DIM cwsText AS CWSTR = SysAllocString(*bstrText)
   SysFreeString bstrText
   RETURN cwsText
END FUNCTION
' ========================================================================================

' ========================================================================================
' Reads an entire TextStream file and returns the resulting string.
' ========================================================================================
PRIVATE FUNCTION CTextStream.ReadAll () AS CWSTR
   DIM bstrText AS AFX_BSTR
   IF m_pTxtStm THEN SetResult(m_pTxtStm->ReadAll(@bstrText))
   DIM cwsText AS CWSTR = *bstrText
   SysFreeString bstrText
   RETURN cwsText
END FUNCTION
' ========================================================================================

' ========================================================================================
' Writes a specified string to a TextStream file.
' ========================================================================================
PRIVATE FUNCTION CTextStream.Write (BYREF cwsText AS CWSTR) AS HRESULT
   DIM bstrText AS AFX_BSTR = SysAllocString(**cwsText)
   IF m_pTxtStm THEN SetResult(m_pTxtStm->Write(bstrText))
   SysFreeString bstrText
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Writes a specified string and newline character to a TextStream file.
' ========================================================================================
PRIVATE FUNCTION CTextStream.WriteLine (BYREF cwsText AS CWSTR) AS HRESULT
   DIM bstrText AS AFX_BSTR = SysALlocString(**cwsText)
   IF m_pTxtStm THEN SetResult(m_pTxtStm->WriteLine(bstrText))
   SysFreeString bstrText
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Writes a specified number of newline characters to a TextStream file.
' ========================================================================================
PRIVATE FUNCTION CTextStream.WriteBlankLines (BYVAL numLines AS LONG) AS HRESULT
   IF m_pTxtStm THEN RETURN(SetResult(m_pTxtStm->WriteBlankLines(numLines)))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Skips a specified number of characters when reading a TextStream file.
' ========================================================================================
PRIVATE FUNCTION CTextStream.Skip (BYVAL numChars AS LONG) AS HRESULT
   IF m_pTxtStm THEN RETURN(SetResult(m_pTxtStm->Skip(numChars)))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Skips the next line when reading a TextStream file.
' ========================================================================================
PRIVATE FUNCTION CTextStream.SkipLine () AS HRESULT
   IF m_pTxtStm THEN RETURN(SetResult(m_pTxtStm->SkipLine))
END FUNCTION
' ========================================================================================

END NAMESPACE
