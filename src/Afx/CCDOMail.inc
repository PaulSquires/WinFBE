' ########################################################################################
' Microsoft Windows
' File: CCDOMail.inc
' Contents: CDO class to send emails.
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2017 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "windows.bi"
#include once "Afx/AfxCDOSys.bi"
#include once "Afx/AfxCOM.inc"
#include once "Afx/CVar.inc"

NAMESPACE Afx

' ========================================================================================
' CCDOMessage class
' ========================================================================================
TYPE CCDOMessage

   Public:
      m_Result AS HRESULT                   ' // HRESULT
      m_pConfig AS Afx_IConfiguration PTR   ' // Pointer to the Afx_IConfiguration interface
      m_pMessage AS Afx_IMessage PTR        ' // Pointer to the Afx_IMessage interface
      m_pFields AS Afx_AdoFields PTR        ' // Pointer to the Afx_AdoFields interface

   Private:
      m_bUninitCOM AS BOOLEAN

   Public:
      DECLARE CONSTRUCTOR
      DECLARE DESTRUCTOR
      DECLARE FUNCTION GetLastResult () AS HRESULT
      DECLARE FUNCTION SetResult (BYVAL Result AS HRESULT) AS HRESULT
      DECLARE FUNCTION ConfigValue (BYREF cvField AS CVAR, BYREF cvValue AS CVAR) AS HRESULT
      DECLARE FUNCTION ConfigUpdate () AS HRESULT
      DECLARE FUNCTION BCC (BYREF cbsBCC AS CBSTR) AS HRESULT
      DECLARE FUNCTION CC (BYREF cbsCC AS CBSTR) AS HRESULT
      DECLARE FUNCTION From (BYREF cbsFrom AS CBSTR) AS HRESULT
      DECLARE FUNCTION Sender (BYREF cbsSender AS CBSTR) AS HRESULT
      DECLARE FUNCTION Recipients (BYREF cbsTo AS CBSTR) AS HRESULT
      DECLARE FUNCTION ReplyTo (BYREF cbsReplyTo AS CBSTR) AS HRESULT
      DECLARE FUNCTION Keywords (BYREF cbsKeywords AS CBSTR) AS HRESULT
      DECLARE FUNCTION Organization (BYREF cbsOrganization AS CBSTR) AS HRESULT
      DECLARE FUNCTION Subject (BYREF cbsSubject AS CBSTR) AS HRESULT
      DECLARE FUNCTION TextBody (BYREF cbsTextBody AS CBSTR) AS HRESULT
      DECLARE FUNCTION HTMLBody (BYREF cbsTextBody AS CBSTR) AS HRESULT
      DECLARE FUNCTION DSNOptions (BYVAL pDSNOptions AS CdoDSNOptions) AS HRESULT
      DECLARE FUNCTION MimeFormatted (BYVAL pMimeFormatted AS VARIANT_BOOL) AS HRESULT
      DECLARE FUNCTION MDNRequested (BYVAL pMDNRequested AS VARIANT_BOOL) AS HRESULT
      DECLARE FUNCTION AddAttachment (BYREF cbsURL AS CBSTR, BYREF cbsUserName AS CBSTR = "", BYREF cbsPassword AS CBSTR = "") AS HRESULT
      DECLARE FUNCTION Send () AS HRESULT
      DECLARE FUNCTION FollowUpTo (BYREF cbsFollowUpTo AS CBSTR) AS HRESULT
      DECLARE FUNCTION Newsgroups (BYREF cbsNewsgroups AS CBSTR) AS HRESULT
      DECLARE FUNCTION Post () AS HRESULT

END TYPE
' ========================================================================================

' ========================================================================================
' CCDOMessage constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CCDOMessage
   ' // Initialize the COM library
   DIM hr AS HRESULT = CoInitialize(NULL)
   IF hr = S_OK OR hr = S_FALSE THEN m_bUninitCOM = TRUE
   ' // Create an instance of the IMessage interface
   m_pMessage = AfxNewCom("CDO.Message")
   ' // Create an instance of the IConfiguration interface
   m_pConfig = AfxNewCom("CDO.Configuration")
   ' // Get a reference to the Afx_AdoFields interface
   IF m_pConfig THEN m_pConfig->get_Fields(@m_pFields)
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CCDOMessage destructor
' ========================================================================================
PRIVATE DESTRUCTOR CCDOMessage
   ' // Release the interfaces
   IF m_pFields THEN m_pFields->Release
   IF m_pConfig THEN m_pConfig->Release
   IF m_pMessage THEN m_pMessage->Release
   ' // Free the COM library
   IF m_bUninitCOM THEN CoUninitialize
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the last result code.
' ========================================================================================
PRIVATE FUNCTION CCDOMessage.GetLastResult () AS HRESULT
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the last result code.
' ========================================================================================
PRIVATE FUNCTION CCDOMessage.SetResult (BYVAL Result AS HRESULT) AS HRESULT
   m_Result = Result
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the value of the specified property.
' ========================================================================================
PRIVATE FUNCTION CCDOMessage.ConfigValue (BYREF cvField AS CVAR, BYREF cvValue AS CVAR) AS HRESULT
   IF m_pConfig = NULL OR m_pFields = NULL THEN RETURN E_NOINTERFACE
   DIM pField AS Afx_AdoField PTR
   this.SetResult(m_pFields->get_Item(cvField, @pField))
   IF m_Result <> S_OK THEN RETURN m_Result
   this.SetResult(pField->put_Value(cvValue))
   pField->Release
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Saves changes made to the Fields collection.
' ========================================================================================
PRIVATE FUNCTION CCDOMessage.ConfigUpdate () AS HRESULT
   IF m_pFields = NULL THEN RETURN E_NOINTERFACE
   RETURN this.SetResult(m_pFields->Update)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Adds to the list of blind copy (Bcc) recipients.
' ========================================================================================
PRIVATE FUNCTION CCDOMessage.BCC (BYREF cbsBCC AS CBSTR) AS HRESULT
   IF m_pMessage = NULL THEN RETURN E_NOINTERFACE
   RETURN this.SetResult(m_pMessage->put_BCC(cbsBCC.m_bstr))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Adds to the list of copy (Cc) recipients.
' ========================================================================================
PRIVATE FUNCTION CCDOMessage.CC (BYREF cbsCC AS CBSTR) AS HRESULT
   IF m_pMessage = NULL THEN RETURN E_NOINTERFACE
   RETURN this.SetResult(m_pMessage->put_CC(cbsCC.m_bstr))
END FUNCTION
' ========================================================================================

' ========================================================================================
' The messaging address of the principal author of the message.
' ========================================================================================
PRIVATE FUNCTION CCDOMessage.From (BYREF cbsFrom AS CBSTR) AS HRESULT
   IF m_pMessage = NULL THEN RETURN E_NOINTERFACE
   RETURN this.SetResult(m_pMessage->put_From(cbsFrom.m_bstr))
END FUNCTION
' ========================================================================================

' ========================================================================================
' The messaging address of the message submitter.
' ========================================================================================
PRIVATE FUNCTION CCDOMessage.Sender (BYREF cbsSender AS CBSTR) AS HRESULT
   IF m_pMessage = NULL THEN RETURN E_NOINTERFACE
   RETURN this.SetResult(m_pMessage->put_Sender(cbsSender.m_bstr))
END FUNCTION
' ========================================================================================

' ========================================================================================
' The principal (To) recipients for this message.
' Named as Recipients instead of To because To is a reserved word.
' ========================================================================================
PRIVATE FUNCTION CCDOMessage.Recipients (BYREF cbsTo AS CBSTR) AS HRESULT
   IF m_pMessage = NULL THEN RETURN E_NOINTERFACE
   RETURN this.SetResult(m_pMessage->put_To(cbsTo.m_bstr))
END FUNCTION
' ========================================================================================

' ========================================================================================
' The messaging addresses to which replies to this message should be sent.
' ========================================================================================
PRIVATE FUNCTION CCDOMessage.ReplyTo (BYREF cbsReplyTo AS CBSTR) AS HRESULT
   IF m_pMessage = NULL THEN RETURN E_NOINTERFACE
   RETURN this.SetResult(m_pMessage->put_ReplyTo(cbsReplyTo.m_bstr))
END FUNCTION
' ========================================================================================

' ========================================================================================
' The list of keywords for this message.
' ========================================================================================
PRIVATE FUNCTION CCDOMessage.Keywords (BYREF cbsKeywords AS CBSTR) AS HRESULT
   IF m_pMessage = NULL THEN RETURN E_NOINTERFACE
   RETURN this.SetResult(m_pMessage->put_Keywords(cbsKeywords))
END FUNCTION
' ========================================================================================

' ========================================================================================
' The organization of the sender.
' ========================================================================================
PRIVATE FUNCTION CCDOMessage.Organization (BYREF cbsOrganization AS CBSTR) AS HRESULT
   IF m_pMessage = NULL THEN RETURN E_NOINTERFACE
   RETURN this.SetResult(m_pMessage->put_Organization(cbsOrganization.m_bstr))
END FUNCTION
' ========================================================================================

' ========================================================================================
' The message subject.
' ========================================================================================
PRIVATE FUNCTION CCDOMessage.Subject (BYREF cbsSubject AS CBSTR) AS HRESULT
   IF m_pMessage = NULL THEN RETURN E_NOINTERFACE
   RETURN this.SetResult(m_pMessage->put_Subject(cbsSubject.m_bstr))
END FUNCTION
' ========================================================================================

' ========================================================================================
' The plain text representation of the message.
' ========================================================================================
PRIVATE FUNCTION CCDOMessage.TextBody (BYREF cbsTextBody AS CBSTR) AS HRESULT
   IF m_pMessage = NULL THEN RETURN E_NOINTERFACE
   RETURN this.SetResult(m_pMessage->put_TextBody(cbsTextBody.m_bstr))
END FUNCTION
' ========================================================================================

' ========================================================================================
' The plain text representation of the message.
' ========================================================================================
PRIVATE FUNCTION CCDOMessage.HTMLBody (BYREF cbsHTMLBody AS CBSTR) AS HRESULT
   IF m_pMessage = NULL THEN RETURN E_NOINTERFACE
   RETURN this.SetResult(m_pMessage->put_HTMLBody(cbsHTMLBody.m_bstr))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Includes a request for a return report on the delivery status of the message.
' ========================================================================================
PRIVATE FUNCTION CCDOMessage.DSNOptions (BYVAL pDSNOptions AS CdoDSNOptions) AS HRESULT
   IF m_pMessage = NULL THEN RETURN E_NOINTERFACE
   RETURN this.SetResult(m_pMessage->put_DSNOptions(pDSNOptions))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Indicates whether a Message Disposition Notification is requested on a message.
' ========================================================================================
PRIVATE FUNCTION CCDOMessage.MimeFormatted (BYVAL pMimeFormatted AS VARIANT_BOOL) AS HRESULT
   IF m_pMessage = NULL THEN RETURN E_NOINTERFACE
   RETURN this.SetResult(m_pMessage->put_MimeFormatted(pMimeFormatted))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Indicates whether a Message Disposition Notification is requested on a message.
' ========================================================================================
PRIVATE FUNCTION CCDOMessage.MDNRequested (BYVAL pMDNRequested AS VARIANT_BOOL) AS HRESULT
   IF m_pMessage = NULL THEN RETURN E_NOINTERFACE
   RETURN this.SetResult(m_pMessage->put_MDNRequested(pMDNRequested))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Adds an attachment to this message.
' ========================================================================================
PRIVATE FUNCTION CCDOMessage.AddAttachment (BYREF cbsURL AS CBSTR, BYREF cbsUserName AS CBSTR = "", BYREF cbsPassword AS CBSTR = "") AS HRESULT
   IF m_pMessage = NULL THEN RETURN E_NOINTERFACE
   RETURN this.SetResult(m_pMessage->AddAttachment(cbsURL.m_bstr, cbsUserName.m_bstr, cbsPassword.m_bstr, NULL))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sends the message.
' ========================================================================================
PRIVATE FUNCTION CCDOMessage.Send () AS HRESULT
   IF m_pMessage = NULL THEN RETURN E_NOINTERFACE
   ' // Set the configuration
   m_pMessage->put_Configuration(m_pConfig)
   ' // Send the message
   RETURN this.SetResult(m_pMessage->Send)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Newsgroups to which any responses to this message are posted.
' ========================================================================================
PRIVATE FUNCTION CCDOMessage.FollowUpTo (BYREF cbsFollowUpTo AS CBSTR) AS HRESULT
   IF m_pMessage = NULL THEN RETURN E_NOINTERFACE
   RETURN this.SetResult(m_pMessage->put_ReplyTo(cbsFollowUpTo.m_bstr))
END FUNCTION
' ========================================================================================

' ========================================================================================
' The newsgroup recipients for the message.
' ========================================================================================
PRIVATE FUNCTION CCDOMessage.Newsgroups (BYREF cbsNewsgroups AS CBSTR) AS HRESULT
   IF m_pMessage = NULL THEN RETURN E_NOINTERFACE
   RETURN this.SetResult(m_pMessage->put_To(cbsNewsgroups.m_bstr))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Posts the message.
' ========================================================================================
PRIVATE FUNCTION CCDOMessage.Post () AS HRESULT
   IF m_pMessage = NULL THEN RETURN E_NOINTERFACE
   ' // Set the configuration
   m_pMessage->put_Configuration(m_pConfig)
   ' // Post the message
   this.SetResult(m_pMessage->Post)
   RETURN m_Result
END FUNCTION
' ========================================================================================

END NAMESPACE
