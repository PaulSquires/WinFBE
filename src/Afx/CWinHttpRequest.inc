' ########################################################################################
' Microsoft Windows
' File: CWinHttpRequest.inc
' Contents: Wrapper class for Microsoft WinHTTP Services, version 5.1
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "windows.bi"
#include once "Afx/AfxWinHttpRequest.bi"
#include once "Afx/CWSTR.inc"

NAMESPACE Afx

' ========================================================================================
' Determines whether the current platform is supported by this version of Microsoft
' Windows HTTP Services (WinHTTP).
' ========================================================================================
#ifndef __AfxWinHttpCheckPlatform_DEFINED__
#define __AfxWinHttpCheckPlatform_DEFINED__
PRIVATE FUNCTION AfxWinHttpCheckPlatform () AS BOOLEAN
   DIM AS ANY PTR pLib = DyLibLoad("winhttp.dll")
   IF pLib = NULL THEN RETURN FALsE
   DIM pCheckPlatform AS FUNCTION () AS BOOL
   pCheckPlatform = DyLibSymbol(pLib, "WinHttpCheckPlatform")
   DIM bResult AS BOOLEAN
   IF pCheckPlatform THEN bResult = pCheckPlatform()
   IF pLib THEN DyLibFree(pLib)
   RETURN bResult
END FUNCTION
#endif
' ========================================================================================

' ########################################################################################
' CWinHttpRequest - Class for Microsoft WinHTTP Services, version 5.1
' ########################################################################################
TYPE CWinHttpRequest

   m_Result AS HRESULT
   m_pWinHttp AS Afx_IWinHttpRequest PTR
   m_bUninitCOM AS BOOLEAN

   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR
   DECLARE FUNCTION GetLastResult () AS HRESULT
   DECLARE FUNCTION SetResult (BYVAL Result AS HRESULT) AS HRESULT
   DECLARE FUNCTION Open (BYREF cwsMethod AS CWSTR, BYREF cwsUrl AS CWSTR, BYVAL bAsync AS BOOLEAN = FALSE) AS HRESULT
   DECLARE FUNCTION Send (BYVAL vBody AS VARIANT = TYPE(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND)) AS HRESULT
   DECLARE FUNCTION WaitForResponse (BYVAL nTimeout AS LONG = 0) AS BOOLEAN
   DECLARE FUNCTION GetResponseHeader (BYREF cwsHeader AS CWSTR) AS CWSTR
   DECLARE FUNCTION GetAllResponseHeaders () AS CWSTR
   DECLARE FUNCTION GetResponseBody () AS STRING
   DECLARE FUNCTION GetResponseStream () AS STRING
   DECLARE FUNCTION GetResponseText () AS CWSTR
   DECLARE FUNCTION SetProxy (BYVAL ProxySetting AS HTTPREQUEST_PROXY_SETTING, BYVAL vProxyServer AS VARIANT = TYPE(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND), BYVAL vBypassList AS VARIANT = TYPE(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND)) AS HRESULT
   DECLARE FUNCTION SetCredentials (BYREF cwsUserName AS CWSTR, BYREF cwsPassword AS CWSTR, BYVAL Flags AS HTTPREQUEST_SETCREDENTIALS_FLAGS) AS HRESULT
   DECLARE FUNCTION SetRequestHeader (BYREF cwsHeader AS CWSTR, BYREF cwsValue AS CWSTR) AS HRESULT
   DECLARE FUNCTION GetStatus () AS LONG
   DECLARE FUNCTION GetStatusText () AS CWSTR
   DECLARE FUNCTION Abort () AS HRESULT
   DECLARE FUNCTION SetTimeouts (BYVAL ResolveTimeout AS LONG, BYVAL ConnectTimeout AS LONG, BYVAL SendTimeout AS LONG, BYVAL ReceiveTimeout AS LONG) AS HRESULT
   DECLARE FUNCTION SetClientCertificate (BYREF cwsClientCertificate AS CWSTR) AS HRESULT
   DECLARE FUNCTION SetAutoLogonPolicy (BYVAL AutoLogonPolicy AS WinHttpRequestAutoLogonPolicy) AS HRESULT
   DECLARE FUNCTION GetOption (BYVAL nOption AS WinHttpRequestOption) AS VARIANT
   DECLARE FUNCTION SetOption (BYVAL nOption AS WinHttpRequestOption, BYVAL vValue AS VARIANT) AS HRESULT
   DECLARE FUNCTION SetOption (BYVAL nOption AS WinHttpRequestOption, BYREF cwsValue AS CWSTR) AS HRESULT

END TYPE
' ########################################################################################

' ========================================================================================
' Constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CWinHttpRequest
   ' // Initialize the COM library
   DIM hr AS HRESULT = CoInitialize(NULL)
   IF hr = S_OK OR hr = S_FALSE THEN m_bUninitCOM = TRUE
   ' // Create an instance of the WinHttpRequest object
   DIM CLSID_WinHttpRequest_ AS GUID = (&h2087C2F4, &h2CEF, &h4953, {&hA8, &hAB, &h66, &h77, &h9B, &h67, &h04, &h95})
   DIM IID_IWinHttpRequest_ AS GUID = (&h016FE2EC, &hB2C8, &h45F8, {&hB2, &h3B, &h39, &hE5, &h3A, &h75, &h39, &h6B})
   SetResult(CoCreateInstance(@CLSID_WinHttpRequest_, NULL, CLSCTX_INPROC_SERVER, @IID_IWinHttpRequest_, @m_pWinHttp))
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Destructor
' ========================================================================================
PRIVATE DESTRUCTOR CWinHttpRequest
   IF m_pWinHttp THEN m_pWinHttp->Release
   IF m_bUninitCOM THEN CoUninitialize
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the last status code.
' ========================================================================================
PRIVATE FUNCTION CWinHttpRequest.GetLastResult () AS HRESULT
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the last status code.
' ========================================================================================
PRIVATE FUNCTION CWinHttpRequest.SetResult (BYVAL Result AS HRESULT) AS HRESULT
   m_Result = Result
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Opens an HTTP connection to an HTTP resource.
' Parameters:
' - Method [in]: Specifies the HTTP verb used for the Open method, such as "GET" or "PUT".
'   Always use uppercase as some servers ignore lowercase HTTP verbs.
' - Url [in]: Specifies the name of the resource. This must be an absolute URL.
' Async [in, optional]: Indicates whether to open in asynchronous mode.
'   FALSE: Opens the HTTP connection in synchronous mode. A call to Send does not return
'   until WinHTTP has completely received the response.
'   TRUE: Opens the HTTP connection in asynchronous mode.
' The return value is S_OK on success or an error value otherwise.
' ========================================================================================
PRIVATE FUNCTION CWinHttpRequest.Open (BYREF cwsMethod AS CWSTR, BYREF cwsUrl AS CWSTR, BYVAL bAsync AS BOOLEAN = FALSE) AS HRESULT
   DIM vAsync AS VARIANT : V_VT(@vAsync) = VT_BOOL : V_BOOL(@vAsync) = bAsync
   DIM bstrMethod AS AFX_BSTR = SysAllocString(**cwsMethod)
   DIM bstrUrl AS AFX_BSTR = SysAllocString(**cwsUrl)
   IF m_pWinHttp THEN SetResult(m_pWinHttp->Open(bstrMethod, bstrUrl, vAsync))
   SysFreeString bstrMethod
   SysFreeString bstrUrl
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' The Send method sends an HTTP request to an HTTP server.
' Parameter:
' - vBody [in, optional]: Data to be sent to the server.
' The return value is S_OK on success or an error value otherwise.
' Remarks: The request to be sent was defined in a prior call to the Open method. The
' calling application can provide data to be sent to the server through the Body parameter.
' If the HTTP verb of the object's Open is "GET", this method sends the request without
' Body, even if it is provided by the calling application.
' ========================================================================================
PRIVATE FUNCTION CWinHttpRequest.Send (BYVAL vBody AS VARIANT = TYPE(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND)) AS HRESULT
   IF m_pWinHttp THEN RETURN SetResult(m_pWinHttp->Send(vBody))
END FUNCTION
' ========================================================================================

' ========================================================================================
' The WaitForResponse method waits for an asynchronous Send method to complete, with
' optional time-out value, in seconds.
' Parameter:
' - Timeout: Time-out value, in seconds. Default time-out is infinite. To explicitly set
'   time-out to infinite, use the value -1.
' Return value:
' TRUE: A response has been received.
' FALSE: The specified time-out period was exceeded.
' ========================================================================================
PRIVATE FUNCTION CWinHttpRequest.WaitForResponse (BYVAL nTimeout AS LONG = 0) AS BOOLEAN
   DIM vTimeout AS VARIANT, Succeeded AS VARIANT_BOOL
   IF nTimeout = 0 THEN V_VT(@vTimeout) = VT_ERROR : V_ERROR(@vTimeout) = DISP_E_PARAMNOTFOUND
   IF nTimeout THEN V_VT(@vTimeout) = VT_INT : V_INT(@vTimeout) = nTimeout
   IF m_pWinHttp THEN SetResult(m_pWinHttp->WaitForResponse(vTimeout, @Succeeded))
   RETURN Succeeded
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the HTTP response headers.
' Parameters:
' - cwsHeader [in]: Specifies the case-insensitive header name.
' Return value: The resulting header information.
' Result code (GetLastResult):
' The result code is S_OK on success or an error value otherwise.
' Remarks: This method returns the value of the response header named in Header. Invoke
' this method only after the Send method has been called.
' ========================================================================================
PRIVATE FUNCTION CWinHttpRequest.GetResponseHeader (BYREF cwsHeader AS CWSTR) AS CWSTR
   DIM bstrValue AS AFX_BSTR, bstrHeader AS AFX_BSTR = SysAllocString(**cwsHeader)
   IF m_pWinHttp THEN SetResult(m_pWinHttp->GetResponseHeader(bstrHeader, @bstrValue))
   SysFreeString bstrHeader
   DIM cwsValue AS CWSTR = *bstrValue
   SysFreeString bstrValue
   RETURN cwsValue
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves all HTTP response headers.
' Return value: The resulting headers.
' Result code (GetLastResult):
' The result code is S_OK on success or an error value otherwise.
' Remarks: This method returns all of the headers contained in the most recent server
' response. The individual headers are delimited by a carriage return and line feed
' combination (ASCII 13 and 10). The last entry is followed by two delimiters (13, 10, 13, 10).
' Invoke this method only after the Send method has been called.
' ========================================================================================
PRIVATE FUNCTION CWinHttpRequest.GetAllResponseHeaders () AS CWSTR
   DIM bstrHeaders AS AFX_BSTR
   IF m_pWinHttp THEN SetResult(m_pWinHttp->GetAllResponseHeaders(@bstrHeaders))
   DIM cwsHeaders AS CWSTR = *bstrHeaders
   SysFreeString bstrHeaders
   RETURN cwsHeaders
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the response entity body as an array of unsigned bytes. This array contains
' the raw data as received directly from the server.
' Result code (GetLastResult):
' The result code is S_OK on success or an error value otherwise.
' Remarks: This function returns the response data in an array of unsigned bytes. If the
' response does not have a response body, an empty variant is returned. This property can
' only be invoked after the Send method has been called.
' ========================================================================================
PRIVATE FUNCTION CWinHttpRequest.GetResponseBody () AS STRING
   DIM vBody AS VARIANT, buffer AS STRING
   IF m_pWinHttp THEN SetResult(m_pWinHttp->get_ResponseBody(@vBody))
   IF m_Result = S_OK THEN
      DIM pvData AS ANY PTR
      IF vBody.pArray THEN
         DIM lLBound AS LONG, lUBound AS LONG
         SetResult(SafeArrayGetLBound(vBody.pArray, 1, @lLBound))
         SetResult(SafeArrayGetUBound(vBody.pArray, 1, @lUBound))
         DIM cElements AS LONG = lUBound - lLBound + 1
         SafeArrayAccessData(vBody.pArray, @pvData)
         IF pvData THEN
            buffer = SPACE(cElements)
            memcpy STRPTR(buffer), pvData, cElements
         END IF
         SafeArrayUnaccessData(vBody.pArray)
      END IF
   END IF
   VariantClear(@vBody)
   RETURN buffer
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the response entity body as an IStream. This stream returns the raw data as
' received directly from the server. This function can only be invoked after the Send
' method has been called.
' Return value: Pointer to the IStream interface on success, or NULL on failure.
' Result code (GetLastResult):
' The result value is S_OK on success or an error value otherwise.
' It will be E_PENDING if the previous Send operation is not complete.
' ========================================================================================
PRIVATE FUNCTION CWinHttpRequest.GetResponseStream () AS STRING
   DIM vResponse AS VARIANT, pStream AS IStream PTR, tstatstg AS STATSTG, strBuffer AS STRING
   IF m_pWinHttp THEN SetResult(m_pWinHttp->get_ResponseStream(@vResponse))
   IF m_Result = S_OK THEN
      IF vResponse.punkVal THEN
         vResponse.punkVal->lpvtbl->QueryInterface(vResponse.punkVal, @IID_IStream, @pStream)
         IF pStream THEN
            ' // Get the size of the stream
            pStream->lpvtbl->Stat(pStream, @tstatstg, STATFLAG_NONAME)
            DIM nSize AS ULONGLONG = tstatstg.cbSize.QuadPart
            ' // Read the stream
            IF nSize THEN
               DIM cbRead AS ULONG
               strBuffer = SPACE(nSize)
               pStream->lpvtbl->Read(pStream, STRPTR(strBuffer), LEN(strBuffer), @cbRead)
               IF cbRead < nSize THEN strBuffer = LEFT(strBuffer, cbRead)
            END IF
            IUnknown_Release(pStream)
         END IF
      END IF
   END IF
   VariantClear(@vResponse)
   RETURN strBuffer
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the response entity body as text.
' This function can only be invoked after the Send method has been called.
' When using this function in synchronous mode, the limit to the number of characters it
' returns is approximately 2,169,895.
' Result code (GetLastResult):
' The result value is S_OK on success or an error value otherwise.
' ========================================================================================
PRIVATE FUNCTION CWinHttpRequest.GetResponseText () AS CWSTR
   DIM bstrResponse AS AFX_BSTR
   IF m_pWinHttp THEN SetResult(m_pWinHttp->get_ResponseText(@bstrResponse))
   DIM cwsResponse AS CWSTR = *bstrResponse
   SysFreeString bstrResponse
   RETURN cwsResponse
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets proxy server information.
' Parameters:
' - ProxySetting: The flags that control this method. Can be one of the following values.
'   HTTPREQUEST_PROXYSETTING_DEFAULT
'      Default proxy setting. Equivalent to HTTPREQUEST_PROXYSETTING_PRECONFIG.
'   HTTPREQUEST_PROXYSETTING_PRECONFIG
'      Indicates that the proxy settings should be obtained from the registry. This assumes
'      that Proxycfg.exe has been run. If Proxycfg.exe has not been run and HTTPREQUEST_PROXYSETTING_PRECONFIG
'      is specified, then the behavior is equivalent to HTTPREQUEST_PROXYSETTING_DIRECT.
'   HTTPREQUEST_PROXYSETTING_DIRECT
'      Indicates that all HTTP and HTTPS servers should be accessed directly. Use this
'      command if there is no proxy server.
'   HTTPREQUEST_PROXYSETTING_PROXY
'      When HTTPREQUEST_PROXYSETTING_PROXY is specified, varProxyServer should be set to a
'      proxy server string and varBypassList should be set to a domain bypass list string.
'      This proxy configuration applies only to the current instance of the WinHttpRequest object.
' - vProxyServer [in, optional]
'      Set to a proxy server string when ProxySetting equals HTTPREQUEST_PROXYSETTING_PROXY.
' - vBypassList [in, optional]
'      Set to a domain bypass list string when ProxySetting equals HTTPREQUEST_PROXYSETTING_PROXY.
' Return value:
'   The return value is S_OK on success or an error value otherwise.
' Remarks:
' Enables the calling application to specify use of default proxy information (configured
' by the proxy configuration tool) or to override Proxycfg.exe. This method must be called
' before calling the Send method. If this method is called after the Send method, it has no effect.
' ========================================================================================
PRIVATE FUNCTION CWinHttpRequest.SetProxy (BYVAL ProxySetting AS HTTPREQUEST_PROXY_SETTING, BYVAL vProxyServer AS VARIANT = TYPE(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND), BYVAL vBypassList AS VARIANT = TYPE(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND)) AS HRESULT
   IF m_pWinHttp THEN RETURN SetResult(m_pWinHttp->SetProxy(ProxySetting, vProxyServer, vBypassList))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets credentials to be used with an HTTP server, whether it is a proxy server or an originating server.
' Parameters:
' - cbsUserName [in]
'   Specifies the user name for authentication.
' - cbsPassword [in]
'   Specifies the password for authentication. This parameter is ignored if bstrUserName is NULL or missing.
' - Flags [in]
'      Specifies when IWinHttpRequest uses credentials. Can be one of the following values.
'      HTTPREQUEST_SETCREDENTIALS_FOR_SERVER
'         Credentials are passed to a server.
'      HTTPREQUEST_SETCREDENTIALS_FOR_PROXY
'         Credentials are passed to a proxy.
' Return value:
'   The return value is S_OK on success or an error value otherwise.
' Remarks:
' This method returns an error value if a call to Open has not completed successfully. It
' is assumed that some measure of interaction with a proxy server or origin server must
' occur before users can set credentials for the session. Moreover, until users know which
' authentication scheme(s) are supported, they cannot format the credentials.
' To authenticate with both the server and the proxy, the application must call SetCredentials
' twice; first with the Flags parameter set to HTTPREQUEST_SETCREDENTIALS_FOR_SERVER, and
' second, with the Flags parameter set to HTTPREQUEST_SETCREDENTIALS_FOR_PROXY.
' ========================================================================================
PRIVATE FUNCTION CWinHttpRequest.SetCredentials (BYREF cwsUserName AS CWSTR, BYREF cwsPassword AS CWSTR, BYVAL Flags AS HTTPREQUEST_SETCREDENTIALS_FLAGS) AS HRESULT
   DIM bstrUserName AS AFX_BSTR = SysAllocString(**cwsUserName)
   DIM bstrPassword AS AFX_BSTR = SysAllocString(**cwsPassword)
   IF m_pWinHttp THEN SetResult(m_pWinHttp->SetCredentials(bstrUserName, bstrPassword, Flags))
   SysFreeString bstrUserName
   SysFreeString bstrPassword
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Adds, changes, or deletes an HTTP request header.
' cbsHeader [in]
'    Specifies the name of the header to be set, for example, "depth". This parameter should
'    not contain a colon and must be the actual text of the HTTP header.
' cbsValue [in]
'    Specifies the value of the header, for example, "infinity".
' Return value:
'   The return value is S_OK on success or an error value otherwise.
' Remarks:
' Headers are transferred across redirects. This can create a security vulnerability. To
' avoid having headers transferred if a redirect occurs, use the WINHTTP_STATUS_CALLBACK
' callback to correct the specific headers when a redirect occurs.
' The SetRequestHeader method enables the calling application to add or delete an HTTP
' request header prior to sending the request. The header name is given in Header, and the
' header token or value is given in Value. To add a header, supply a header name and value.
' If another header already exists with this name, it is replaced. To delete a header, set
' Header to the name of the header to delete and set Value to NULL.
' The name and value of request headers added with this method are validated. Headers must
' be well formed. For more information about valid HTTP headers, see RFC 2616. If an
' invalid header is used, an error occurs and the header is not added.
' ========================================================================================
PRIVATE FUNCTION CWinHttpRequest.SetRequestHeader (BYREF cwsHeader AS CWSTR, BYREF cwsValue AS CWSTR) AS HRESULT
   DIM bstrHeader AS AFX_BSTR = SysAllocString(**cwsHeader)
   DIM bstrValue AS AFX_BSTR = SysAllocString(**cwsValue)
   IF m_pWinHttp THEN SetResult(m_pWinHttp->SetRequestHeader(bstrHeader, bstrValue))
   SysFreeString bstrHeader
   SysFreeString bstrValue
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the HTTP status code from the last response.
' The results of this property are valid only after the Send method has successfully completed.
' For a list of status codes see HTTP Status Codes.
' https://msdn.microsoft.com/en-us/library/windows/desktop/aa383887(v=vs.85).aspx
' Remarks: This is not the same that the values returned by GetLastResult.
' ========================================================================================
PRIVATE FUNCTION CWinHttpRequest.GetStatus () AS LONG
   DIM nStatus AS LONG
   IF m_pWinHttp THEN SetResult(m_pWinHttp->get_Status(@nStatus))
   RETURN nStatus
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the HTTP status text.
' Remarks: Retrieves the text portion of the server response line, making available the
' "user-friendly" equivalent to the numeric HTTP status code. The results of this property
' are valid only after the Send method has successfully completed.
' Result code (GetLastResult):
' The return value is S_OK on success or an error value otherwise.
' ========================================================================================
PRIVATE FUNCTION CWinHttpRequest.GetStatusText () AS CWSTR
   DIM bstrText AS AFX_BSTR
   IF m_pWinHttp THEN SetResult(m_pWinHttp->get_StatusText(@bstrText))
   DIM cwsText AS CWSTR = *bstrText
   SysFreeString bstrText
   RETURN cwsText
END FUNCTION
' ========================================================================================

' ========================================================================================
' Aborts a WinHTTP Send method.
' The return value is S_OK on success or an error value otherwise.
' Remarks: You can abort both asynchronous and synchronous Send methods. To abort a
' synchronous Send method, you must call Abort from within an IWinHttpRequestEvents event.
' ========================================================================================
PRIVATE FUNCTION CWinHttpRequest.Abort () AS HRESULT
   IF m_pWinHttp THEN RETURN(SetResult(m_pWinHttp->Abort))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Specifies the individual time-out components of a send/receive operation, in milliseconds.
' Parameters
' - ResolveTimeout: [in] Value of type Long integer. Time-out value applied when resolving
'   a host name (such as www.microsoft.com) to an IP address (such as 192.168.131.199),
'   in milliseconds. The default value is zero, meaning no time-out (infinite). If DNS
'   timeout is specified using NAME_RESOLUTION_TIMEOUT, there is an overhead of one thread
'   per request.
' - ConnectTimeout: [in] Value of type Long integer. Time-out value applied when establishing
'   a communication socket with the target server, in milliseconds. The default value is
'   60,000 (60 seconds).
' - SendTimeout: [in] Value of type Long integer. Time-out value applied when sending an
'   individual packet of request data on the communication socket to the target server, in
'   milliseconds. A large request sent to an HTTP server are normally be broken up into
'   multiple packets; the send time-out applies to sending each packet individually. The
'   default value is 30,000 (30 seconds).
' - ReceiveTimeout: [in] Value of type Long integer. Time-out value applied when receiving
'   a packet of response data from the target server, in milliseconds. Large responses are
'   be broken up into multiple packets; the receive time-out applies to fetching each packet
'   of data off the socket. The default value is 30,000 (30 seconds).
' ========================================================================================
PRIVATE FUNCTION CWinHttpRequest.SetTimeouts (BYVAL ResolveTimeout AS LONG, BYVAL ConnectTimeout AS LONG, BYVAL SendTimeout AS LONG, BYVAL ReceiveTimeout AS LONG) AS HRESULT
   IF m_pWinHttp THEN RETURN(SetResult(m_pWinHttp->SetTimeouts(ResolveTimeout, ConnectTimeout, SendTimeout, ReceiveTimeout)))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Selects a client certificate to send to a Secure Hypertext Transfer Protocol (HTTPS) server.
' Parameter:
' - cbsClientCertificate: Specifies the location, certificate store, and subject of a client certificate.
' The return value is S_OK on success or an error value otherwise.
' Remarks:
' The string specified in the ClientCertificate parameter consists of the certificate location,
' certificate store, and subject name delimited by backslashes. For more information about
' the components of the certificate string, see Client Certificates.
' The certificate store name and location are optional. However, if you specify a certificate
' store, you must also specify the location of that certificate store. The default location
' is CURRENT_USER and the default certificate store is "MY". A blank subject indicates that
' the first certificate in the certificate store should be used.
' Call SetClientCertificate to select a certificate before calling Send to send the request.
' Microsoft Windows HTTP Services (WinHTTP) does not provide client certificates to proxy
' servers that request certificates for authentication.
' ========================================================================================
PRIVATE FUNCTION CWinHttpRequest.SetClientCertificate (BYREF cwsClientCertificate AS CWSTR) AS HRESULT
   DIM bstrClientCertificate AS AFX_BSTR = SysAllocString(**cwsClientCertificate)
   IF m_pWinHttp THEN SetResult(m_pWinHttp->SetClientCertificate(bstrClientCertificate))
   SysFreeString bstrClientCertificate
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the current Automatic Logon Policy.
' Parameter:
' - AutoLogonPolicy [in]: Specifies the current automatic logon policy.
' The return value is S_OK on success or an error value otherwise.
' The default policy is AutoLogonPolicy_OnlyIfBypassProxy.
' Call SetAutoLogonPolicy to set the automatic logon policy before calling Send to se
' ========================================================================================
PRIVATE FUNCTION CWinHttpRequest.SetAutoLogonPolicy (BYVAL AutoLogonPolicy AS WinHttpRequestAutoLogonPolicy) AS HRESULT
   IF m_pWinHttp THEN RETURN(SetResult(m_pWinHttp->SetAutoLogonPolicy(AutoLogonPolicy)))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves a Microsoft Windows HTTP Services (WinHTTP) option value.
' ========================================================================================
PRIVATE FUNCTION CWinHttpRequest.GetOption (BYVAL nOption AS WinHttpRequestOption) AS VARIANT
   DIM vValue AS VARIANT
   IF m_pWinHttp THEN SetResult(m_pWinHttp->get_Option(nOption, @vValue))
   RETURN vValue
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets a Microsoft Windows HTTP Services (WinHTTP) option value.
' ========================================================================================
PRIVATE FUNCTION CWinHttpRequest.SetOption (BYVAL nOption AS WinHttpRequestOption, BYVAL vValue AS VARIANT) AS HRESULT
   IF m_pWinHttp THEN RETURN(SetResult(m_pWinHttp->put_Option(nOption, vValue)))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CWinHttpRequest.SetOption (BYVAL nOption AS WinHttpRequestOption, BYREF cwsValue AS CWSTR) AS HRESULT
   DIM vValue AS VARIANT : vValue.vt = VT_BSTR : vValue.bstrVal = SysAllocString(**cwsValue)
   IF m_pWinHttp THEN SetResult(m_pWinHttp->put_Option(nOption, vValue))
   VariantClear @vValue
   RETURN m_Result
END FUNCTION
' ========================================================================================

END NAMESPACE
