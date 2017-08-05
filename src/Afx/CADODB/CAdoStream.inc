' ########################################################################################
' Microsoft Windows
' Contents: ADO Stream class
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once

NAMESPACE Afx

' ########################################################################################
' CAdoStream class
' ########################################################################################
TYPE CAdoStream EXTENDS CAdoBase

   m_Result AS HRESULT
   m_pStream AS Afx_ADOStream PTR

   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYVAL pStream AS Afx_ADOStream PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   DECLARE DESTRUCTOR
'   DECLARE OPERATOR @ () AS Afx_ADOStream PTR PTR
   DECLARE FUNCTION vptr () AS Afx_ADOStream PTR PTR

   ' // Attaches a recordset to the class
   DECLARE SUB Attach (BYVAL pStream AS Afx_ADOStream PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   ' // Indicates the size of the stream in number of bytes.
   DECLARE PROPERTY Size () AS LONG
   ' // Indicates whether the current position is at the end of the stream.
   DECLARE PROPERTY EOS () AS BOOLEAN
   ' // Indicates the current position within a Stream object.
   DECLARE PROPERTY Position () AS LONG
   DECLARE PROPERTY Position (BYVAL nPos AS LONG)
   ' // Indicates the type of data contained in the Stream (binary or text).
   DECLARE PROPERTY Type_ () AS StreamTypeEnum
   DECLARE PROPERTY Type_ (BYVAL nType AS StreamTypeEnum)
   ' // Indicates the binary character to be used as the line separator in text Stream objects.
   DECLARE PROPERTY LineSeparator () AS LineSeparatorEnum
   DECLARE PROPERTY LineSeparator (BYVAL pLS AS LineSeparatorEnum)
   ' // Indicates for whether the state of the Stream object is open or closed.
   DECLARE PROPERTY State () AS ObjectStateEnum
   ' // Indicates the available permissions for modifying data in a Stream object.
   DECLARE PROPERTY Mode () AS ConnectModeEnum
   DECLARE PROPERTY Mode (BYVAL nMode AS ConnectModeEnum)
   ' // Indicates the character set into which the contents of a text Stream should be
   ' // translated for storage in the Stream object's internal buffer.
   DECLARE PROPERTY Charset () AS CBSTR
   DECLARE PROPERTY Charset (BYREF cbsCharset AS CBSTR)
   ' // Reads a specified number of bytes from a binary Stream object.
   DECLARE FUNCTION Read (BYVAL NumBytes AS LONG = adReadAll) AS CVAR
   DECLARE FUNCTION Read (BYVAL NumBytes AS LONG = adReadAll, BYREF cvValue AS CVAR) AS HRESULT
   ' // Opens a connection to a data source.
   DECLARE FUNCTION Open (BYREF cvSource AS CVAR = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND), _
           BYVAL nMode AS ConnectModeEnum = adModeUnknown, BYVAL Options AS StreamOpenOptionsEnum = adOpenStreamUnspecified, _
           BYREF cbsUserName AS CBSTR = "", BYREF cbsPassword AS CBSTR = "") AS HRESULT
   ' // Closes a Stream object and any dependent objects.
   DECLARE FUNCTION Close () AS HRESULT
   ' // Skips one entire line when reading a text stream.
   DECLARE FUNCTION SkipLine () AS HRESULT
   ' // Writes binary data to a Stream object.
   DECLARE FUNCTION Write (BYREF cvBuffer AS CVAR) AS HRESULT
   ' // Writes a string to a Stream object.
   DECLARE FUNCTION WriteText (BYREF cbsData AS CBSTR, BYVAL Options AS StreamWriteEnum = adWriteChar) AS HRESULT
   ' // Sets the position that is the end of the stream.
   DECLARE FUNCTION SetEOS () AS HRESULT
   ' // Copies the specified number of characters or bytes (depending on Type) in the Stream to another Stream object.
   DECLARE FUNCTION CopyTo (BYVAL pDestStream AS Afx_ADOStream PTR, BYVAL CharNumber AS LONG = adReadAll) AS HRESULT
   ' // Forces the contents of the Stream remaining in the ADO buffer to the underlying object with which the Stream is associated.
   DECLARE FUNCTION Flush () AS HRESULT
   ' // Saves the binary contents of a Stream to a file.
   DECLARE FUNCTION SaveToFile (BYREF cbsFileName AS CBSTR, BYVAL Options AS SaveOptionsEnum = adSaveCreateNotExist) AS HRESULT
   ' // Loads the contents of an existing file into a Stream.
   DECLARE FUNCTION LoadFromFile (BYREF cbsFileName AS CBSTR) AS HRESULT
   ' // Reads a specified number of characters, an entire line, or the entire stream from a Stream object and returns the resulting string.
   DECLARE FUNCTION ReadText (BYVAL NumChars AS LONG = adReadAll) AS CBSTR
   DECLARE FUNCTION ReadText (BYVAL NumChars AS LONG = adReadAll, BYREF cbsText AS CBSTR) AS HRESULT
   ' // Cancels execution of a pending, asynchronous method call.
   DECLARE FUNCTION Cancel () AS HRESULT

END TYPE

' ========================================================================================
' Default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CAdoStream
   CADODB_DP("-BEGIN CAdoStream CONSTRUCTOR")
   ' // Create an instance of the Connection object
   m_pStream = AfxNewCom("ADODB.Stream")
   CADODB_DP("-END CAdoStream CONSTRUCTOR " & WSTR(m_pStream))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CAdoStream (BYVAL pStream AS Afx_ADOStream PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   CADODB_DP("CAdoStream CONSTRUCTOR")
   m_pStream = pStream
   IF fAddRef THEN
      IF m_pStream THEN m_pStream->AddRef
   END IF
END CONSTRUCTOR
' ========================================================================================

' ===========================================================================================
' Cleanup
' ===========================================================================================
PRIVATE DESTRUCTOR CAdoStream
   CADODB_DP("CAdoStream DESTRUCTOR " & WSTR(m_pStream))
   IF m_pStream THEN
      IF this.State = adStateOpen THEN this.Close
      m_pStream->Release
   END IF
END DESTRUCTOR
' ===========================================================================================

' ========================================================================================
' Returns the address of the connection pointer
' ========================================================================================
'PRIVATE OPERATOR CAdoStream.@ () AS Afx_ADOStream PTR PTR
'   OPERATOR = @m_pStream
'END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CAdoStream.vptr () AS Afx_ADOStream PTR PTR
   CVAR_DP("CAdoStream vptr")
   IF m_pStream THEN m_pStream->Release
   RETURN @m_pStream
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the encapsulated IDispatch pointer
' ========================================================================================
PRIVATE OPERATOR * (BYREF cStream AS CAdoStream) AS Afx_ADOStream PTR
   OPERATOR = cStream.m_pStream
END OPERATOR
' ========================================================================================

' ========================================================================================
' Attaches an stream to the class.
' Useful to attach to the class a pointer returned by a function.
' ========================================================================================
PRIVATE SUB CAdoStream.Attach (BYVAL pStream AS Afx_ADOStream PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   CADODB_DP("CAdoStream.Attach")
   IF m_pStream THEN m_pStream->Release
   m_pStream = pStream
   IF fAddRef THEN
      IF m_pStream THEN m_pStream->AddRef
   END IF
END SUB
' ========================================================================================

' ===========================================================================================
' Indicates the size of the stream in number of bytes.
' ===========================================================================================
PRIVATE PROPERTY CAdoStream.Size () AS LONG
   CADODB_DP("CAdoStream.Size")
   DIM lSize AS LONG
   IF m_pStream THEN SetResult(m_pStream->get_Size(@lSize))
   PROPERTY = lSize
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates whether the current position is at the end of the stream.
' ===========================================================================================
PRIVATE PROPERTY CAdoStream.EOS () AS BOOLEAN
   CADODB_DP("CAdoStream.EOS")
   DIM bEOS AS VARIANT_BOOL
   IF m_pStream THEN SetResult(m_pStream->get_EOS(@bEOS))
   PROPERTY = bEOS
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the current position within a Stream object.
' ===========================================================================================
PRIVATE PROPERTY CAdoStream.Position () AS LONG
   CADODB_DP("CAdoStream.Position")
   DIM nPos AS LONG
   IF m_pStream THEN SetResult(m_pStream->get_Position(@nPos))
   PROPERTY = nPos
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoStream.Position (BYVAL nPos AS LONG)
   CADODB_DP("CAdoStream.Position")
   IF m_pStream THEN SetResult(m_pStream->put_Position(nPos))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the type of data contained in the Stream (binary or text).
' ===========================================================================================
PRIVATE PROPERTY CAdoStream.Type_ () AS StreamTypeEnum
   CADODB_DP("CAdoStream.Type_")
   DIM nType AS StreamTypeEnum
   IF m_pStream THEN SetResult(m_pStream->get_Type(@nType))
   PROPERTY = nType
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoStream.Type_ (BYVAL nType AS StreamTypeEnum)
   CADODB_DP("CAdoStream.Type_")
   IF m_pStream THEN SetResult(m_pStream->put_Type(nType))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the binary character to be used as the line separator in text Stream objects.
' ===========================================================================================
PRIVATE PROPERTY CAdoStream.LineSeparator () AS LineSeparatorEnum
   CADODB_DP("CAdoStream.LineSeparator")
   DIM LS AS LineSeparatorEnum
   IF m_pStream THEN SetResult(m_pStream->get_LineSeparator(@LS))
   PROPERTY = LS
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoStream.LineSeparator (BYVAL LS AS LineSeparatorEnum)
   CADODB_DP("CAdoStream.LineSeparator")
   IF m_pStream THEN SetResult(m_pStream->put_LineSeparator(LS))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates for whether the state of the Stream object is open or closed.
' ===========================================================================================
PRIVATE PROPERTY CAdoStream.State () AS ObjectStateEnum
   CADODB_DP("CAdoStream.LineSeparator")
   DIM lState AS ObjectStateEnum
   IF m_pStream THEN SetResult(m_pStream->get_State(@lState))
   PROPERTY = lState
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the available permissions for modifying data in a Stream object.
' ===========================================================================================
PRIVATE PROPERTY CAdoStream.Mode () AS ConnectModeEnum
   CADODB_DP("CAdoStream.Mode")
   DIM nMode AS ConnectModeEnum
   IF m_pStream THEN SetResult(m_pStream->get_Mode(@nMode))
   PROPERTY = nMode
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoStream.Mode (BYVAL nMode AS ConnectModeEnum)
   CADODB_DP("CAdoStream.Mode")
   IF m_pStream THEN SetResult(m_pStream->put_Mode(nMode))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the available permissions for modifying data in a Stream object.
' ===========================================================================================
PRIVATE PROPERTY CAdoStream.Charset () AS CBSTR
   CADODB_DP("CAdoStream.Charset")
   DIM bstrCharset AS AFX_BSTR
   IF m_pStream THEN SetResult(m_pStream->get_Charset(@bstrCharset))
   PROPERTY = bstrCharset
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoStream.Charset (BYREF cbsCharset AS CBSTR)
   CADODB_DP("CAdoStream.Charset")
   IF m_pStream THEN SetResult(m_pStream->put_Charset(cbsCharset.m_bstr))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Reads a specified number of bytes from a binary Stream object.
' ===========================================================================================
PRIVATE FUNCTION CAdoStream.Read (BYVAL NumBytes AS LONG = adReadAll) AS CVAR
   CADODB_DP("CAdoStream.Read")
   DIM cvValue AS CVAR
   IF m_pStream THEN SetResult(m_pStream->Read(NumBytes, cvValue.vptr))
   RETURN cvValue
END FUNCTION
' ===========================================================================================
' ===========================================================================================
PRIVATE FUNCTION CAdoStream.Read (BYVAL NumBytes AS LONG = adReadAll, BYREF cvValue AS CVAR) AS HRESULT
   CADODB_DP("CAdoStream.Read")
   IF m_pStream THEN RETURN(SetResult(m_pStream->Read(NumBytes, cvValue.vptr)))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Opens a connection to a data source.
' ===========================================================================================
PRIVATE FUNCTION CAdoStream.Open (BYREF cvSource AS CVAR = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND), _
BYVAL nMode AS ConnectModeEnum = adModeUnknown, BYVAL Options AS StreamOpenOptionsEnum = adOpenStreamUnspecified, _
BYREF cbsUserName AS CBSTR = "", BYREF cbsPassword AS CBSTR = "") AS HRESULT
   CADODB_DP("CAdoStream.Open")
   IF m_pStream THEN RETURN(SetResult(m_pStream->Open(cvSource, nMode, Options, cbsUserName.m_bstr, cbsPassword.m_bstr)))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Closes a Stream object and any dependent objects.
' ===========================================================================================
PRIVATE FUNCTION CAdoStream.Close () AS HRESULT
   CADODB_DP("CAdoStream.Close")
   IF m_pStream THEN RETURN(SetResult(m_pStream->Close))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Skips one entire line when reading a text stream.
' ===========================================================================================
PRIVATE FUNCTION CAdoStream.SkipLine () AS HRESULT
   CADODB_DP("CAdoStream.SkipLine")
   IF m_pStream THEN RETURN(SetResult(m_pStream->SkipLine))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Writes binary data to a Stream object.
' ===========================================================================================
PRIVATE FUNCTION CAdoStream.Write (BYREF cvBuffer AS CVAR) AS HRESULT
   CADODB_DP("CAdoStream.Write")
   IF m_pStream THEN RETURN(SetResult(m_pStream->Write(cvBuffer)))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Writes a string to a Stream object.
' ===========================================================================================
PRIVATE FUNCTION CAdoStream.WriteText (BYREF cbsData AS CBSTR, BYVAL Options AS StreamWriteEnum = adWriteChar) AS HRESULT
   CADODB_DP("CAdoStream.WriteText")
   IF m_pStream THEN RETURN(SetResult(m_pStream->WriteText(cbsData.m_bstr, Options)))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Sets the position that is the end of the stream.
' ===========================================================================================
PRIVATE FUNCTION CAdoStream.SetEOS () AS HRESULT
   CADODB_DP("CAdoStream.SetEOS")
   IF m_pStream THEN RETURN(SetResult(m_pStream->SetEOS))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Copies the specified number of characters or bytes (depending on Type) in the Stream to
' another Stream object.
' ===========================================================================================
PRIVATE FUNCTION CAdoStream.CopyTo (BYVAL pDestStream AS Afx_ADOStream PTR, BYVAL CharNumber AS LONG = adReadAll) AS HRESULT
   CADODB_DP("CAdoStream.CopyTo")
   IF m_pStream THEN RETURN(SetResult(m_pStream->CopyTo(pDestStream, CharNumber)))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Forces the contents of the Stream remaining in the ADO buffer to the underlying object with
' which the Stream is associated.
' ===========================================================================================
PRIVATE FUNCTION CAdoStream.Flush () AS HRESULT
   CADODB_DP("CAdoStream.Flush")
   IF m_pStream THEN RETURN(SetResult(m_pStream->Flush))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Saves the binary contents of a Stream to a file.
' ===========================================================================================
PRIVATE FUNCTION CAdoStream.SaveToFile (BYREF cbsFileName AS CBSTR, BYVAL Options AS SaveOptionsEnum = adSaveCreateNotExist) AS HRESULT
   CADODB_DP("CAdoStream.SaveToFile")
   IF m_pStream THEN RETURN(SetResult(m_pStream->SaveToFile(cbsFileName.m_bstr, Options)))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Loads the contents of an existing file into a Stream.
' ===========================================================================================
PRIVATE FUNCTION CAdoStream.LoadFromFile (BYREF cbsFileName AS CBSTR) AS HRESULT
   CADODB_DP("CAdoStream.LoadFromFile")
   IF m_pStream THEN RETURN(SetResult(m_pStream->LoadFromFile(cbsFileName.m_bstr)))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Reads a specified number of characters, an entire line, or the entire stream from a Stream
' object and returns the resulting string.
' ===========================================================================================
PRIVATE FUNCTION CAdoStream.ReadText (BYVAL NumChars AS LONG = adReadAll) AS CBSTR
   CADODB_DP("CAdoStream.ReadText")
   DIM bstrText AS AFX_BSTR
   IF m_pStream THEN SetResult(m_pStream->ReadText(NumChars, @bstrText))
   RETURN bstrText
END FUNCTION
' ===========================================================================================
' ===========================================================================================
PRIVATE FUNCTION CAdoStream.ReadText (BYVAL NumChars AS LONG = adReadAll, BYREF cbsText AS CBSTR) AS HRESULT
   CADODB_DP("CAdoStream.ReadText")
   IF m_pStream THEN RETURN(SetResult(m_pStream->ReadText(NumChars, cbsText.vptr)))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Cancels execution of a pending, asynchronous method call.
' ===========================================================================================
PRIVATE FUNCTION CAdoStream.Cancel () AS HRESULT
   CADODB_DP("CAdoStream.Cancel")
   IF m_pStream THEN RETURN(SetResult(m_pStream->Cancel))
END FUNCTION
' ===========================================================================================

END NAMESPACE
