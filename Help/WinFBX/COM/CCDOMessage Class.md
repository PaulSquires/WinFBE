# CCDOMessage Class

Collaboration Data Objects (CDO) is a Component Object Model (COM) component designed to simplify writing programs that create or manipulate Internet messages. It supports creating and managing messages formatted and sent using Internet standards such as the Multipurpose Internet Mail Extensions (MIME) standard. CDO supports sending messages using both the SMTP and NNTP protocols, as well as through a local SMTP/NNTP service pickup directory. CDO integrates with the Microsoft® ActiveX® Data Objects (ADO) component to present a consistent mechanism for managing data that comprises a message.

**CCDOMessage** is a Free Basic class that allows to send messages using CDO.

**Include file:** CCDOMail.inc.

#### Example

```
' // Create an instance of the CCdoMessage class
DIM pMsg AS CCDOMessage

' // Configuration
pMsg.ConfigValue(cdoSendUsingMethod, CdoSendUsingPort)
pMsg.ConfigValue(cdoSMTPServer, "smtp.xxxxx.xxx")
pMsg.ConfigValue(cdoSMTPServerPort, 25)
pMsg.ConfigValue(cdoSMTPAuthenticate, 1)
pMsg.ConfigValue(cdoSendUserName, "xxxx@xxxx.xxx")
pMsg.ConfigValue(cdoSendPassword, "xxxxxxxx")
pMsg.ConfigValue(cdoSMTPUseSSL, 1)
pMsg.ConfigUpdate
 
' // Recipient name --> change as needed
pMsg.Recipients("xxxxx@xxxxx")
' // Sender mail address --> change as needed
pMsg.From("xxxxx@xxxxx")
' // Subject --> change as needed
pMsg.Subject("This is a sample subject")
' // Text body --> change as needed
pMsg.TextBody("This is a sample message text")
' // Add the attachment (use absolute paths).
' // Note  By repeating the call you can attach more than one file.
pMsg.AddAttachment ExePath & "\xxxxx.xxx"
' // Send the message
pMsg.Send
IF pMsg.GetLastResult = S_OK THEN PRINT "Message sent" ELSE PRINT "Failure"

To send messages using gmail, simply change the values of the server name and the server port:

pMsg.ConfigValue(cdoSMTPServer, "smtp.gmail.com")
pMsg.ConfigValue(cdoSMTPServerPort, 465)
```

### Methods

| Name       | Description |
| ---------- | ----------- |
| [AddAttachment](#AddAttachment) | Adds an attachment to the message. |
| [BCC](#BCC) | The blind carbon copy (Bcc) recipients for the message. |
| [CC](#CC) | The secondary (carbon copy) recipients for this message. |
| [ConfigUpdate](#ConfigUpdate) | Saves the changes you make to the **Fields** collection of the CDO **IConfiguration** interface. |
| [ConfigValue](#ConfigValue) | Sets the value of the specified configuration field. |
| [DSNOptions](#DSNOptions) | Includes a request for a return report on the delivery status of the message. |
| [FollowUpTo](#FollowUpTo) | Newsgroups to which any responses to this message are posted. |
| [From](#From) | The messaging address of the principal author of the message. |
| [GetErrorInfo](#GetErrorInfo) | Returns the description of the most recent error. |
| [GetLastResult](#GetLastResult) | Returns the last result code. |
| [HTMLBody](#HTMLBody) | The Hypertext Markup Language (HTML) representation of the message. |
| [Keywords](#Keywords) | The list of keywords for the message. |
| [MDNRequested](#MDNRequested) | Indicates whether a Message Disposition Notification is requested on a message. |
| [MimeFormatted](#MimeFormatted) | Indicates whether a Message Disposition Notification is requested on a message. |
| [Newsgroups](#Newsgroups) | The newsgroup recipients for the message. |
| [Organization](#Organization) | The organization of the sender. |
| [Post](#Post) | Submits this message to the specified newsgroups. |
| [Recipients](#Recipients) | The principal (To) recipients for this message. |
| [ReplyTo](#ReplyTo) | The messaging addresses to which replies to this message should be sent. |
| [Send](#Send) | Sends the message. |
| [Sender](#Sender) | The messaging address of the message submitter. |
| [Subject](#Subject) | The message subject. |
| [TextBody](#TextBody) | The plain text representation of the message. |

# <a name="AddAttachment"></a>AddAttachment

Adds an attachment to the message.

```
FUNCTION AddAttachment (BYREF cbsURL AS CBSTR, BYREF cbsUserName AS CBSTR = "", _
   BYREF cbsPassword AS CBSTR = "") AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsURL* | The full path and file name of the message to be attached to this message. |
| *cbsUserName* | An optional user name to use for authentication when retrieving the resource using Hypertext Tranfer Protocol (HTTP). Can be used to set the credentials for both the basic and NTLM authentication mechanisms. |
| *cbsPassword* | An optional password to use for authentication when retrieving the resource using the HTTP. Can be used to set the credentials for the basic and NTLM authentication mechanisms. |

Return value

S_OK or an HRESULT code.

#### Remarks

The **AddAttachment** method adds the attachment by first retrieving the resource specified by the Uniform Resource Locator (URL) and then adding the content to the message's **Attachments** collection within a **BodyPart** object.

The URL prefixes supported in the URL parameter are file://, ftp://, http://, and https://. The default prefix is file://. This facilitates designation of paths starting with drive letters and of universal naming convention (UNC) paths.

The **MIMEFormatted** property determines how the attachment is formatted when the message is serialized for delivery to a Simple Mail Transfer Protocol (SMTP) service. If this property is set to True, the attachment is formatted using Multipurpose Internet Mail Extensions (MIME). If the property is set to False, the attachment is added to the serialized content stream in Uuencoded format.

If you populate the **HTMLBody** property before calling the **AddAttachment** method, any inline images are displayed as part of the message.

Use the **UserName** and **Password** parameters when you are requesting Web pages using Hypertext Transfer Protocol (HTTP) from a server that requires client authentication. If the Web server supports only the basic authentication mechanism, these credentials must be supplied. If the Web server supports the NTLM authentication mechanism, by default, the current process security context is used to authenticate; however, you can specify alternate credentials for NTLM authentication with the **UserName** and **Password** properties.

**Important**: Storing user names and passwords inside source code can lead to security vulnerabilities in your software. Do not store user names and passwords in your production code.

# <a name="BCC"></a>BCC

The blind carbon copy (Bcc) recipients for the message.

```
FUNCTION BCC (BYREF cbsBCC AS CBSTR PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsBCC* | The string in the BCC method can represent a single recipient or multiple recipients. Each address in the list must be a full messaging address, such as "User" \<example@example.com> or example@example.com.<br>In lists of multiple recipients, commas separate addresses, as follows: "User1" \<example1@example.com >, "User 2" \<example2@example.com>, "User3" \<example3@example.com><br>A comma is not valid in any part of a messaging address unless it is enclosed in quotation marks.<br>To maintain the privacy intended by blind copies, **BCC** is regarded as an envelope property rather than a message property; accordingly, the corresponding header field and its contents are removed when the message is delivered, and the **BCC** property is always empty on a received message.<br>The default value of the **BCC** property is an empty string. |

#### Return value

S_OK (0) or an HRESULT code.

# <a name="CC"></a>CC

The secondary (carbon copy) recipients for this message.

```
FUNCTION CC (BYREF cbsCC AS CBSTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsCC* | The string in the BCC method can represent a single recipient or multiple recipients. Each address in the list must be a full messaging address, such as "User" \<example@example.com> or example@example.com.<br>In lists of multiple recipients, commas separate addresses, as follows: "User1" \<example1@example.com >, "User 2" \<example2@example.com>, "User3" \<example3@example.com><br>A comma is not valid in any part of a messaging address unless it is enclosed in quotation marks.<br>To maintain the privacy intended by blind copies, **CC** is regarded as an envelope property rather than a message property; accordingly, the corresponding header field and its contents are removed when the message is delivered, and the **CC** property is always empty on a received message.<br>The default value of the **CC** property is an empty string. |

#### Return value

S_OK (0) or an HRESULT code.

# <a name="ConfigUpdate"></a>ConfigUpdate

Saves the changes you make to the **Fields** collection of the CDO **IConfiguration** interface.

```
FUNCTION ConfigUpdate () AS HRESULT
```

#### Return value

S_OK (0) or an HRESULT code.

# <a name="ConfigValue"></a>ConfigValue

Sets the value of the specified configuration field.

```
FUNCTION ConfigValue (BYREF cbsField AS CBSTR, BYREF cbsValue AS CBSTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsField* | The http://schemas.microsoft.com/cdo/configuration/ namespace defines the majority of fields used to set configurations for various CDO objects. These configuration fields are set using an implementation of the **IConfiguration.Fields** collection.<br><br>Many CDO objects use information stored in an associated **Configuration** object to define configuration settings. One example is the **Message** object, where you use its associated **Configuration** object to set fields such as sendusing. This field defines whether to send the message using the local SMTP service drop directory (if the local machine has the SMTP service installed), an SMTP service directly over the network. If sending over the network, you set smtpserver to specify the IP address or DNS name of the machine hosting the SMTP service, and optionally, smtpserverport to specify a port value. If credentials are required for connecting to an SMTP service, you can specify them by setting the sendusername and sendpassword.<br>A similar set of fields exists for posting messages using either a local NNTP service pickup directory, or over the network. |
| *cbsValue* | The value to set. |

All of the names listed below are also defined as string constants in the file **AfxCDOSys.bi** for convenience.

```
cdoAutoPromoteBodyParts = "http://schemas.microsoft.com/cdo/configuration/autopromotebodyparts"
cdoFlushBuffersOnWrite = "http://schemas.microsoft.com/cdo/configuration/flushbuffersonwrite"
cdoHTTPCookies = "http://schemas.microsoft.com/cdo/configuration/httpcookies"
cdoLanguageCode = "http://schemas.microsoft.com/cdo/configuration/languagecode"
cdoNNTPAccountName = "http://schemas.microsoft.com/cdo/configuration/nntpaccountname"
cdoNNTPAuthenticate = "http://schemas.microsoft.com/cdo/configuration/nntpauthenticate"
cdoNNTPConnectionTimeout = "http://schemas.microsoft.com/cdo/configuration/nntpconnectiontimeout"
cdoNNTPServer = "http://schemas.microsoft.com/cdo/configuration/nntpserver"
cdoNNTPServerPickupDirectory = "http://schemas.microsoft.com/cdo/configuration/nntpserverpickupdirectory"
cdoNNTPServerPort = "http://schemas.microsoft.com/cdo/configuration/nntpserverport"
cdoNNTPUseSSL = "http://schemas.microsoft.com/cdo/configuration/nntpusessl"
cdoPostEmailAddress = "http://schemas.microsoft.com/cdo/configuration/postemailaddress"
cdoPostPassword = "http://schemas.microsoft.com/cdo/configuration/postpassword"
cdoPostUserName = "http://schemas.microsoft.com/cdo/configuration/postusername"
cdoPostUserReplyEmailAddress = "http://schemas.microsoft.com/cdo/configuration/postuserreplyemailaddress"
cdoPostUsingMethod = "http://schemas.microsoft.com/cdo/configuration/postusing"
cdoSaveSentItems = "http://schemas.microsoft.com/cdo/configuration/savesentitems"
cdoSendEmailAddress = "http://schemas.microsoft.com/cdo/configuration/sendemailaddress"
cdoSendPassword = "http://schemas.microsoft.com/cdo/configuration/sendpassword"
cdoSendUserName = "http://schemas.microsoft.com/cdo/configuration/sendusername"
cdoSendUserReplyEmailAddress = "http://schemas.microsoft.com/cdo/configuration/senduserreplyemailaddress"
cdoSendUsingMethod = "http://schemas.microsoft.com/cdo/configuration/sendusing"
cdoSMTPAccountName = "http://schemas.microsoft.com/cdo/configuration/smtpaccountname"
cdoSMTPAuthenticate = "http://schemas.microsoft.com/cdo/configuration/smtpauthenticate"
cdoSMTPConnectionTimeout = "http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout"
cdoSMTPServer = "http://schemas.microsoft.com/cdo/configuration/smtpserver"
cdoSMTPServerPickupDirectory = "http://schemas.microsoft.com/cdo/configuration/smtpserverpickupdirectory"
cdoSMTPServerPort = "http://schemas.microsoft.com/cdo/configuration/smtpserverport"
cdoSMTPUseSSL = "http://schemas.microsoft.com/cdo/configuration/smtpusessl"
cdoURLGetLatestVersion = "http://schemas.microsoft.com/cdo/configuration/urlgetlatestversion"
cdoURLProxyBypass = "http://schemas.microsoft.com/cdo/configuration/urlproxybypass"
cdoURLProxyServer = "http://schemas.microsoft.com/cdo/configuration/urlproxyserver"
cdoUseMessageResponseText = "http://schemas.microsoft.com/cdo/configuration/usemessageresponsetext"
```

#### Return value

S_OK (0) or an HRESULT code.

### Example

```
' // Create an instance of the CCdoMessage class
DIM pMsg AS CCdoMessage
' // Configuration
pMsg.ConfigValue(cdoSendUsingMethod, CdoSendUsingPort)
pMsg.ConfigValue(cdoSMTPServer, "smtp.xxxxx.xxx")
pMsg.ConfigValue(cdoSMTPServerPort, 25)
pMsg.ConfigValue(cdoSMTPAuthenticate, 1)
pMsg.ConfigValue(cdoSendUserName, "xxxx@xxxx.xxx")
pMsg.ConfigValue(cdoSendPassword, "xxxxxxxx")
pMsg.ConfigValue(cdoSMTPUseSSL, 1)
pMsg.ConfigUpdate
```

# <a name="DSNOptions"></a>DSNOptions

Includes a request for a return report on the delivery status of the message.

```
FUNCTION DSNOptions (BYVAL BYVAL pDSNOptions AS CdoDSNOptions) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pDSNOptions* | Delivery Status Notifications are essentially requests for receipts of message delivery status. The request can be for the enumerated delivery conditions. Note that Delivery Status Notifications can pass through numerous message transfer agents and thus return receipts can have different meanings. *DSNOptions* defines the type of Delivery Status Notification requested per Request for Comments (RFC) 1894. The RFC covers the nature and uses of Delivery Status Notifications in depth. |

The **CdoDSNOptions** enumeration is provided for this property.

| Constant   | Value | Description |
| ---------- | ----- | ----------- |
| **cdoDSNDefault** | 0 | No delivery status notifications are issued. |
| **cdoDSNNever** | 1 | No delivery status notifications are issued. |
| **cdoDSNFailure** | 2 | Returns a delivery status notification if delivery fails. |
| **cdoDSNSuccess** | 4 | Returns a delivery status notification if delivery succeeds. |
| **cdoDSNDelay** | 8 | Returns a delivery status notification if delivery is delayed. |
| **cdoDSNSuccessFailOrDelay** | 14 | Returns a delivery status notification if delivery succeeds, fails, or is delayed. |

#### Return value

S_OK (0) or an HRESULT code.

# <a name="FollowUpTo"></a>FollowUpTo

Newsgroups to which any responses to this message are posted.

```
FUNCTION FollowUpTo (BYREF cbsFollowUpTo AS CBSTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsFollowUpTo* | The string in the **FollowUpTo** property can represent a single newsgroup or multiple newsgroups. A newsgroup name is a set of words concatenated by periods, and multiple newsgroups in the list are separated by commas, as shown in the following examples:<br><br>name.public.discussion<br>alt.sample,sample.newsgroup.name.public.chat<br><br>A newsgroup name cannot contain a comma, and only existing newsgroups should be specified. If a newsgroup with a particular name does not exist, posting to that name does not cause the newsgroup to be created. If the **FollowUpTo** property is not set, responses are posted to the newsgroups specified by the **Newsgroups** property. The default value of the **FollowUpTo** property is an empty string. |

#### Return value

S_OK (0) or an HRESULT code.

# <a name="From"></a>From

The messaging address of the principal author of the message.

```
FUNCTION From (BYREF cbsFrom AS CBSTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsFrom* | The string in the From method represents the full messaging addresses of the message author or authors, as shown in the following example: "Jane Doe" \<example@example.com>, \<example@example.com>, example2@example.com<br><br>Commas serve as address separators; however, commas are not parsed as separators when they appear in an address that is enclosed in double quotes, such as the following: "John Jones, Jr." \<example@example.com>, "Jane Doe" \<example@example.com><br><br>If you do not set the **From** property before calling the **Send** method, Microsoft Collaboration Data Objects (CDO) uses the **Sender** property if it is present. If neither the **From** and **Sender** properties are set before calling the **Send** method, an exception is raised.<br><br>The **Sender** property is not used with Network News Transfer Protocol (NTTP).<br><br>The difference between the **From** and **Sender** properties is that the **Sender** property represents the messaging user that actually submits the message, while the **From** property designates its principal author or authors. If the author and the sender are the same, it is only necessary to set the **From** property, and it is recommended to leave the **Sender** property unset in this case.  |

#### Return value

S_OK (0) or an HRESULT code.

# <a name="GetErrorInfo"></a>GetErrorInfo

Returns the description of the most recent OLE error in the current logical thread and clears the error state for the thread. It should be called as soon as possible after calling a method of this class. The numerical error code can obtained calling the **GetLastResult** method.

```
FUNCTION GetErrorInfo() AS CBSTR
```

#### Return value

A localized description of the error.


# <a name="GetLastResult"></a>GetLastResult

Returns the last result code.

```
FUNCTION GetLastResult () AS HRESULT
```

#### Return value

The result code returned by the last executed method. CDO uses the **IErrorInfo** interface to provide error data. A description of the error can be obtained calling the **GetOleErrorInfo** method.

# <a name="HTMLBody"></a>HTMLBody

The Hypertext Markup Language (HTML) representation of the message.

```
FUNCTION HTMLBody (BYREF cbsHTMLBody AS CBSTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsHTMLBody* | The default value of the **HTMLBody** property is an empty string.  |

Example:

```
DIM cbsHTML AS CBSTR
cwsHTML = "<!DOCTYPE HTML PUBLIC ""-//IETF//DTD HTML//EN"">" & $LF
cwsHTML += "<HTML>"
cwsHTML += "  <HEAD>"
cwsHTML += "    <TITLE>Sample GIF</TITLE>"
cwsHTML += "  </HEAD>"
cwsHTML += "  <BODY><P>"
cwsHTML += "   <IMG src=""XYZ.gif""></p><p>Inline graphics</P>"
cwsHTML += "  </BODY>"
cwsHTML += "</HTML>"
pMsg.HtmlBody(cbsHTML)
```

#### Return value

S_OK (0) or an HRESULT code.

# <a name="Keywords"></a>Keywords

The list of keywords for the message.

```
FUNCTION Keywords (BYREF cbsKeywords AS CBSTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsKeywords* | The keywords of a message can be useful for determining if the message is of interest to the reader, or for searching for relevant messages in a collection.<br><br>The string in the **Keywords** property can represent a single keyword or multiple keywords. A keyword string can optionally contain spaces, such as the following:<br><br>"1997 payroll"<br><br>Multiple keywords in the list are separated by commas, as in the following example:<br><br>"operating systems,Windows NT,functions"<br><br>Because the comma is the delimiter of this field, a keyword cannot contain a comma. If you add a keyword that contains a comma, the result will be two keywords that are separated by the comma.<br><br>The default value of the **Keywords** property is an empty string.  |

#### Return value

S_OK (0) or an HRESULT code.

# <a name="MDNRequested"></a>MDNRequested

Indicates whether a Message Disposition Notification is requested on a message.

```
FUNCTION MDNRequested (BYVAL BYVAL pMDNRequested AS VARIANT_BOOL) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pMDNRequested* | A Message Disposition Notificationis a request for information to be returned on the status of this message. The message sender or an intermediary Message Transfer Agent may request Message Disposition Notifications. The message to be returned is nested in a message with a content type of multipart/report. Request for Comments (RFC) 2298 describes Message Disposition Notifications, their function, and their format. |

#### Return value

S_OK (0) or an HRESULT code.

# <a name="MimeFormatted"></a>MimeFormatted

Indicates whether a Message Disposition Notification is requested on a message.

```
FUNCTION MimeFormatted (BYVAL pMimeFormatted AS VARIANT_BOOL) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pMimeFormatted* | If the **MimeFormatted** property is set to False and you set the **HTMLBody** property, Microsoft Collaboration Data Objects (CDO) sets the **MimeFormatted** property to True. If you set the **MimeFormatted** property to False, the **HTMLBody** property is removed from the message and the Hypertext Markup Language (HTML) text is lost.<br><br>If you prepare a message with the **MimeFormatted** property set to True and later change it to False, all of the body parts are made into attachments and marked for encoding with the Uuencode mechanism. Encoding takes place only when you call the **Send** or **Post** method.<br><br>The default value of the **MimeFormatted** property is True. |

#### Return value

S_OK (0) or an HRESULT code.

# <a name="Newsgroups"></a>Newsgroups

The newsgroup recipients for the message.

```
FUNCTION Newsgroups (BYREF cbsNewgroups AS CBSTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsNewgroups* | The string in the **Newsgroups** property can represent a single newsgroup or multiple newsgroups. A newsgroup name is a set of words concatenated by periods, as shown in the following example:<br><br>name.public.discussion<br><br>In a list of multiple newsgroups, commas separate the names of the newsgroups, as follows:<br><br>alt.sample,sample.newsgroup.name,name.public.chat<br><br>A newsgroup name cannot contain a comma, and only existing newsgroups should be specified. If a newsgroup with a particular name does not exist, posting to that name does not cause the newsgroup to be created.<br><br>The **Newsgroups** property is required on news messages. If you do not set the **Newsgroups** property before calling the **Post** method, an error is returned.<br><br>Responses to this message are posted to the same newsgroups unless the **FollowUpTo** property is set.<br><br>The default value of the **Newsgroups** property is an empty string. |

#### Return value

S_OK (0) or an HRESULT code.

# <a name="Organization"></a>Organization

The organization of the sender.

```
FUNCTION Organization (BYREF cbsOrganization AS CBSTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsOrganization* | The **Organization** property is used for the Network News Transfer Protocol (NNTP) Organization header field. This field supplies a short phrase that meaningfully describes the sender's organization, such as "Sample Corporation, Payroll Department". A phrase such as this can be easier to recognize than a cryptic messaging address like "Q1006453@example.com".<br><br>The default value of the **Organization** property is an empty string. |

#### Return value

S_OK (0) or an HRESULT code.

# <a name="Post"></a>Post

Submits this message to the specified newsgroups.

```
FUNCTION Post () AS HRESULT
```

#### Remarks

How the message is posted depends on the current configuration. This configuration is defined in the associated Configuration object that is referenced by the **Configuration** property on this interface. Consult a specific Component Object Model (COM) class implementation for a list of configuration settings used when posting messages.

#### Return value

S_OK (0) or an HRESULT code.

# <a name="Recipients"></a>Recipients

The principal (To) recipients for this message.

```
FUNCTION Recipients (BYVAL cbsRecipients AS CBSTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsRecipients* | The string in the **Recipients** property can represent a single recipient or multiple recipients. For example, each of the following qualifies as a full messaging address:<br><br>"User Address" \<example@example.com><br>\<example@example.com><br>example@example.com<br><br>In lists with multiple recipients, addresses are separated by commas:<br><br>"User 1" \<example1@example.com>, example2@example.com, \<example3@example.com><br><br>A comma is not allowed in any part of a messaging address unless it is contained within quotation marks.<br><br>The default value of the **Recipients** property is an empty string. |

#### Return value

S_OK (0) or an HRESULT code.

# <a name="ReplyTo"></a>ReplyTo

The messaging addresses to which replies to this message should be sent.

```
FUNCTION ReplyTo (BYREF cbsReplyTo AS CBSTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsReplyTo* | The messaging address or addresses to which replies are to be sent. The default value of the **ReplyTo** property is an empty string. |

#### Return value

S_OK (0) or an HRESULT code.

# <a name="Send"></a>Send

Sends the message.

```
FUNCTION Send () AS HRESULT
```

#### Return value

S_OK (0) or an HRESULT code.

# <a name="Sender"></a>Sender

The messaging address of the message submitter.

```
FUNCTION Sender (BYREF cbsSender AS CBSTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsSender* | The difference between the **From** and **Sender** properties is that the **Sender** property identifies the address of the user or entity that actually submits the message, whereas the **From** property designates its principal author or authors. By convention, which is outlined in Request for Comments (RFC) 822, multiple addressees are not identified in the **Sender** property.<br><br>The **Sender** property is normally used under the following circumstances:<br><br>Multiple addressees are listed in the **From** header. The sender is the e-mail address of the user or entity in the From field that actually submitted the message.<br><br>The user or entity that is submitting the message is not included in the **From** field. |

#### Return value

S_OK (0) or an HRESULT code.

# <a name="Subject"></a>Subject

The message subject.

```
FUNCTION Subject (BYREF cbsSubject AS CBSTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsSubject* | The default value of the **Subject** property is an empty string. |

#### Return value

S_OK (0) or an HRESULT code.

# <a name="TextBody"></a>TextBody

The plain text representation of the message.

```
FUNCTION TextBody (BYREF cbsTextBody AS CBSTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsTextBody* | The default value of the **TextBody** property is an empty string. |

#### Return value

S_OK (0) or an HRESULT code.
