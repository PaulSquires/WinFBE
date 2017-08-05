' ########################################################################################
' Microsoft Windows
' File: AfxCDOSys.bi
' Library name: CDO
' Documentation string: Microsoft CDO for Windows 2000 Library
' GUID: {CD000000-8B95-11D1-82DB-00C04FB1625D}
' Version: 1.0, Locale ID = 0
' Path: C:\Windows\SysWOW64\cdosys.dll
' Attributes: 8 [&h00000008]  [HasDiskImage]
' ########################################################################################

#pragma once
#include once "Afx/AfxCDOSys.bi"
#include once "Afx/AfxADO.bi"

NAMESPACE Afx

' // The definition for BSTR in the FreeBASIC headers was inconveniently changed to WCHAR
#ifndef AFX_BSTR
   #define AFX_BSTR WSTRING PTR
#endif

' // ProgIDs (Program Identifiers)
CONST AFX_PROGID_CDO_Configuration_1 = "CDO.Configuration.1"
CONST AFX_PROGID_CDO_DropDirectory_1 = "CDO.DropDirectory.1"
CONST AFX_PROGID_CDO_Message_1 = "CDO.Message.1"
CONST AFX_PROGID_CDO_NNTPEarlyConnector_1 = "CDO.NNTPEarlyConnector.1"
CONST AFX_PROGID_CDO_NNTPFinalConnector_1 = "CDO.NNTPFinalConnector.1"
CONST AFX_PROGID_CDO_NNTPPostConnector_1 = "CDO.NNTPPostConnector.1"
CONST AFX_PROGID_CDO_SMTPConnector_1 = "CDO.SMTPConnector.1"

' // Version independent ProgIDs
CONST AFX_PROGID_CDO_Configuration = "CDO.Configuration"
CONST AFX_PROGID_CDO_DropDirectory = "CDO.DropDirectory"
CONST AFX_PROGID_CDO_Message = "CDO.Message"
CONST AFX_PROGID_CDO_NNTPEarlyConnector = "CDO.NNTPEarlyConnector"
CONST AFX_PROGID_CDO_NNTPFinalConnector = "CDO.NNTPFinalConnector"
CONST AFX_PROGID_CDO_NNTPPostConnector = "CDO.NNTPPostConnector"
CONST AFX_PROGID_CDO_SMTPConnector = "CDO.SMTPConnector"


' // ClsIDs (Class identifiers)
CONST AFX_CLSID_Configuration = "{CD000002-8B95-11D1-82DB-00C04FB1625D}"
CONST AFX_CLSID_DropDirectory = "{CD000004-8B95-11D1-82DB-00C04FB1625D}"
CONST AFX_CLSID_Message = "{CD000001-8B95-11D1-82DB-00C04FB1625D}"
CONST AFX_CLSID_NNTPEarlyConnector = "{CD000011-8B95-11D1-82DB-00C04FB1625D}"
CONST AFX_CLSID_NNTPFinalConnector = "{CD000010-8B95-11D1-82DB-00C04FB1625D}"
CONST AFX_CLSID_NNTPPostConnector = "{CD000009-8B95-11D1-82DB-00C04FB1625D}"
CONST AFX_CLSID_SMTPConnector = "{CD000008-8B95-11D1-82DB-00C04FB1625D}"

' // IIDs (Interface identifiers)
CONST AFX_IID_Afx_IBodyPart = "{CD000021-8B95-11D1-82DB-00C04FB1625D}"
CONST AFX_IID_Afx_IBodyParts = "{CD000023-8B95-11D1-82DB-00C04FB1625D}"
CONST AFX_IID_Afx_IConfiguration = "{CD000022-8B95-11D1-82DB-00C04FB1625D}"
CONST AFX_IID_Afx_IDataSource = "{CD000029-8B95-11D1-82DB-00C04FB1625D}"
CONST AFX_IID_Afx_IDropDirectory = "{CD000024-8B95-11D1-82DB-00C04FB1625D}"
CONST AFX_IID_Afx_IMessage = "{CD000020-8B95-11D1-82DB-00C04FB1625D}"
CONST AFX_IID_Afx_IMessages = "{CD000025-8B95-11D1-82DB-00C04FB1625D}"
CONST AFX_IID_Afx_INNTPEarlyScriptConnector = "{CD000034-8B95-11D1-82DB-00C04FB1625D}"
CONST AFX_IID_Afx_INNTPFinalScriptConnector = "{CD000032-8B95-11D1-82DB-00C04FB1625D}"
CONST AFX_IID_Afx_INNTPOnPost = "{CD000027-8B95-11D1-82DB-00C04FB1625D}"
CONST AFX_IID_Afx_INNTPOnPostEarly = "{CD000033-8B95-11D1-82DB-00C04FB1625D}"
CONST AFX_IID_Afx_INNTPOnPostFinal = "{CD000028-8B95-11D1-82DB-00C04FB1625D}"
CONST AFX_IID_Afx_INNTPPostScriptConnector = "{CD000031-8B95-11D1-82DB-00C04FB1625D}"
CONST AFX_IID_Afx_ISMTPOnArrival = "{CD000026-8B95-11D1-82DB-00C04FB1625D}"
CONST AFX_IID_Afx_ISMTPScriptConnector = "{CD000030-8B95-11D1-82DB-00C04FB1625D}"

' // Enumerations

ENUM CdoConfigSource
   ' // Number of constants: 3
   cdoDefaults = -1   ' (&hFFFFFFFF)
   cdoIIS = 1   ' (&h00000001)
   cdoOutlookExpress = 2   ' (&h00000002)
END ENUM

ENUM CdoDSNOptions
   ' // Number of constants: 6
   cdoDSNDefault = 0   ' (&h00000000)
   cdoDSNNever = 1   ' (&h00000001)
   cdoDSNFailure = 2   ' (&h00000002)
   cdoDSNSuccess = 4   ' (&h00000004)
   cdoDSNDelay = 8   ' (&h00000008)
   cdoDSNSuccessFailOrDelay = 14   ' (&h0000000E)
END ENUM

ENUM CdoEventStatus
   ' // Number of constants: 2
   cdoRunNextSink = 0   ' (&h00000000)
   cdoSkipRemainingSinks = 1   ' (&h00000001)
END ENUM

ENUM cdoImportanceValues
   ' // Number of constants: 3
   cdoLow = 0   ' (&h00000000)
   cdoNormal = 1   ' (&h00000001)
   cdoHigh = 2   ' (&h00000002)
END ENUM

ENUM CdoMessageStat
   ' // Number of constants: 3
   cdoStatSuccess = 0   ' (&h00000000)
   cdoStatAbortDelivery = 2   ' (&h00000002)
   cdoStatBadMail = 3   ' (&h00000003)
END ENUM

ENUM CdoMHTMLFlags
   ' // Number of constants: 7
   cdoSuppressNone = 0   ' (&h00000000)
   cdoSuppressImages = 1   ' (&h00000001)
   cdoSuppressBGSounds = 2   ' (&h00000002)
   cdoSuppressFrames = 4   ' (&h00000004)
   cdoSuppressObjects = 8   ' (&h00000008)
   cdoSuppressStyleSheets = 16   ' (&h00000010)
   cdoSuppressAll = 31   ' (&h0000001F)
END ENUM

ENUM CdoNNTPProcessingField
   ' // Number of constants: 3
   cdoPostMessage = 1   ' (&h00000001)
   cdoProcessControl = 2   ' (&h00000002)
   cdoProcessModerator = 4   ' (&h00000004)
END ENUM

ENUM CdoPostUsing
   ' // Number of constants: 2
   cdoPostUsingPickup = 1   ' (&h00000001)
   cdoPostUsingPort = 2   ' (&h00000002)
END ENUM

ENUM cdoPriorityValues
   ' // Number of constants: 3
   cdoPriorityNonUrgent = -1   ' (&hFFFFFFFF)
   cdoPriorityNormal = 0   ' (&h00000000)
   cdoPriorityUrgent = 1   ' (&h00000001)
END ENUM

ENUM CdoProtocolsAuthentication
   ' // Number of constants: 3
   cdoAnonymous = 0   ' (&h00000000)
   cdoBasic = 1   ' (&h00000001)
   cdoNTLM = 2   ' (&h00000002)
END ENUM

ENUM CdoReferenceType
   ' // Number of constants: 2
   cdoRefTypeId = 0   ' (&h00000000)
   cdoRefTypeLocation = 1   ' (&h00000001)
END ENUM

ENUM CdoSendUsing
   ' // Number of constants: 2
   cdoSendUsingPickup = 1   ' (&h00000001)
   cdoSendUsingPort = 2   ' (&h00000002)
END ENUM

ENUM cdoSensitivityValues
   ' // Number of constants: 4
   cdoSensitivityNone = 0   ' (&h00000000)
   cdoPersonal = 1   ' (&h00000001)
   cdoPrivate = 2   ' (&h00000002)
   cdoCompanyConfidential = 3   ' (&h00000003)
END ENUM

ENUM CdoTimeZoneId
   ' // Number of constants: 83
   cdoUTC = 0   ' (&h00000000)
   cdoGMT = 1   ' (&h00000001)
   cdoSarajevo = 2   ' (&h00000002)
   cdoParis = 3   ' (&h00000003)
   cdoBerlin = 4   ' (&h00000004)
   cdoEasternEurope = 5   ' (&h00000005)
   cdoPrague = 6   ' (&h00000006)
   cdoAthens = 7   ' (&h00000007)
   cdoBrasilia = 8   ' (&h00000008)
   cdoAtlanticCanada = 9   ' (&h00000009)
   cdoEastern = 10   ' (&h0000000A)
   cdoCentral = 11   ' (&h0000000B)
   cdoMountain = 12   ' (&h0000000C)
   cdoPacific = 13   ' (&h0000000D)
   cdoAlaska = 14   ' (&h0000000E)
   cdoHawaii = 15   ' (&h0000000F)
   cdoMidwayIsland = 16   ' (&h00000010)
   cdoWellington = 17   ' (&h00000011)
   cdoBrisbane = 18   ' (&h00000012)
   cdoAdelaide = 19   ' (&h00000013)
   cdoTokyo = 20   ' (&h00000014)
   cdoSingapore = 21   ' (&h00000015)
   cdoBangkok = 22   ' (&h00000016)
   cdoBombay = 23   ' (&h00000017)
   cdoAbuDhabi = 24   ' (&h00000018)
   cdoTehran = 25   ' (&h00000019)
   cdoBaghdad = 26   ' (&h0000001A)
   cdoIsrael = 27   ' (&h0000001B)
   cdoNewfoundland = 28   ' (&h0000001C)
   cdoAzores = 29   ' (&h0000001D)
   cdoMidAtlantic = 30   ' (&h0000001E)
   cdoMonrovia = 31   ' (&h0000001F)
   cdoBuenosAires = 32   ' (&h00000020)
   cdoCaracas = 33   ' (&h00000021)
   cdoIndiana = 34   ' (&h00000022)
   cdoBogota = 35   ' (&h00000023)
   cdoSaskatchewan = 36   ' (&h00000024)
   cdoMexicoCity = 37   ' (&h00000025)
   cdoArizona = 38   ' (&h00000026)
   cdoEniwetok = 39   ' (&h00000027)
   cdoFiji = 40   ' (&h00000028)
   cdoMagadan = 41   ' (&h00000029)
   cdoHobart = 42   ' (&h0000002A)
   cdoGuam = 43   ' (&h0000002B)
   cdoDarwin = 44   ' (&h0000002C)
   cdoBeijing = 45   ' (&h0000002D)
   cdoAlmaty = 46   ' (&h0000002E)
   cdoIslamabad = 47   ' (&h0000002F)
   cdoKabul = 48   ' (&h00000030)
   cdoCairo = 49   ' (&h00000031)
   cdoHarare = 50   ' (&h00000032)
   cdoMoscow = 51   ' (&h00000033)
   cdoFloating = 52   ' (&h00000034)
   cdoCapeVerde = 53   ' (&h00000035)
   cdoCaucasus = 54   ' (&h00000036)
   cdoCentralAmerica = 55   ' (&h00000037)
   cdoEastAfrica = 56   ' (&h00000038)
   cdoMelbourne = 57   ' (&h00000039)
   cdoEkaterinburg = 58   ' (&h0000003A)
   cdoHelsinki = 59   ' (&h0000003B)
   cdoGreenland = 60   ' (&h0000003C)
   cdoRangoon = 61   ' (&h0000003D)
   cdoNepal = 62   ' (&h0000003E)
   cdoIrkutsk = 63   ' (&h0000003F)
   cdoKrasnoyarsk = 64   ' (&h00000040)
   cdoSantiago = 65   ' (&h00000041)
   cdoSriLanka = 66   ' (&h00000042)
   cdoTonga = 67   ' (&h00000043)
   cdoVladivostok = 68   ' (&h00000044)
   cdoWestCentralAfrica = 69   ' (&h00000045)
   cdoYakutsk = 70   ' (&h00000046)
   cdoDhaka = 71   ' (&h00000047)
   cdoSeoul = 72   ' (&h00000048)
   cdoPerth = 73   ' (&h00000049)
   cdoArab = 74   ' (&h0000004A)
   cdoTaipei = 75   ' (&h0000004B)
   cdoSydney2000 = 76   ' (&h0000004C)
   cdoChihuahua = 77   ' (&h0000004D)
   cdoCanberraCommonwealthGames2006 = 78   ' (&h0000004E)
   cdoAdelaideCommonwealthGames2006 = 79   ' (&h0000004F)
   cdoHobartCommonwealthGames2006 = 80   ' (&h00000050)
   cdoTijuana = 81   ' (&h00000051)
   cdoInvalidTimeZone = 82   ' (&h00000052)
END ENUM


' // Modules

' // Module: CdoCalendar
' // Number of constants: 1
CONST cdoTimeZoneIDURN = "urn:schemas:calendar:timezoneid"

' // Module: CdoCharset
' // Number of constants: 21
CONST cdoBIG5 = "big5"
CONST cdoEUC_JP = "euc-jp"
CONST cdoEUC_KR = "euc-kr"
CONST cdoGB2312 = "gb2312"
CONST cdoISO_2022_JP = "iso-2022-jp"
CONST cdoISO_2022_KR = "iso-2022-kr"
CONST cdoISO_8859_1 = "iso-8859-1"
CONST cdoISO_8859_2 = "iso-8859-2"
CONST cdoISO_8859_3 = "iso-8859-3"
CONST cdoISO_8859_4 = "iso-8859-4"
CONST cdoISO_8859_5 = "iso-8859-5"
CONST cdoISO_8859_6 = "iso-8859-6"
CONST cdoISO_8859_7 = "iso-8859-7"
CONST cdoISO_8859_8 = "iso-8859-8"
CONST cdoISO_8859_9 = "iso-8859-9"
CONST cdoKOI8_R = "koi8-r"
CONST cdoShift_JIS = "shift-jis"
CONST cdoUS_ASCII = "us-ascii"
CONST cdoUTF_7 = "utf-7"
CONST cdoUTF_8 = "utf-8"
CONST cdoISO_8859_15 = "iso-8859-15"

' // Module: CdoConfiguration
' // Number of constants: 33
CONST cdoAutoPromoteBodyParts = "http://schemas.microsoft.com/cdo/configuration/autopromotebodyparts"
CONST cdoFlushBuffersOnWrite = "http://schemas.microsoft.com/cdo/configuration/flushbuffersonwrite"
CONST cdoHTTPCookies = "http://schemas.microsoft.com/cdo/configuration/httpcookies"
CONST cdoLanguageCode = "http://schemas.microsoft.com/cdo/configuration/languagecode"
CONST cdoNNTPAccountName = "http://schemas.microsoft.com/cdo/configuration/nntpaccountname"
CONST cdoNNTPAuthenticate = "http://schemas.microsoft.com/cdo/configuration/nntpauthenticate"
CONST cdoNNTPConnectionTimeout = "http://schemas.microsoft.com/cdo/configuration/nntpconnectiontimeout"
CONST cdoNNTPServer = "http://schemas.microsoft.com/cdo/configuration/nntpserver"
CONST cdoNNTPServerPickupDirectory = "http://schemas.microsoft.com/cdo/configuration/nntpserverpickupdirectory"
CONST cdoNNTPServerPort = "http://schemas.microsoft.com/cdo/configuration/nntpserverport"
CONST cdoNNTPUseSSL = "http://schemas.microsoft.com/cdo/configuration/nntpusessl"
CONST cdoPostEmailAddress = "http://schemas.microsoft.com/cdo/configuration/postemailaddress"
CONST cdoPostPassword = "http://schemas.microsoft.com/cdo/configuration/postpassword"
CONST cdoPostUserName = "http://schemas.microsoft.com/cdo/configuration/postusername"
CONST cdoPostUserReplyEmailAddress = "http://schemas.microsoft.com/cdo/configuration/postuserreplyemailaddress"
CONST cdoPostUsingMethod = "http://schemas.microsoft.com/cdo/configuration/postusing"
CONST cdoSaveSentItems = "http://schemas.microsoft.com/cdo/configuration/savesentitems"
CONST cdoSendEmailAddress = "http://schemas.microsoft.com/cdo/configuration/sendemailaddress"
CONST cdoSendPassword = "http://schemas.microsoft.com/cdo/configuration/sendpassword"
CONST cdoSendUserName = "http://schemas.microsoft.com/cdo/configuration/sendusername"
CONST cdoSendUserReplyEmailAddress = "http://schemas.microsoft.com/cdo/configuration/senduserreplyemailaddress"
CONST cdoSendUsingMethod = "http://schemas.microsoft.com/cdo/configuration/sendusing"
CONST cdoSMTPAccountName = "http://schemas.microsoft.com/cdo/configuration/smtpaccountname"
CONST cdoSMTPAuthenticate = "http://schemas.microsoft.com/cdo/configuration/smtpauthenticate"
CONST cdoSMTPConnectionTimeout = "http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout"
CONST cdoSMTPServer = "http://schemas.microsoft.com/cdo/configuration/smtpserver"
CONST cdoSMTPServerPickupDirectory = "http://schemas.microsoft.com/cdo/configuration/smtpserverpickupdirectory"
CONST cdoSMTPServerPort = "http://schemas.microsoft.com/cdo/configuration/smtpserverport"
CONST cdoSMTPUseSSL = "http://schemas.microsoft.com/cdo/configuration/smtpusessl"
CONST cdoURLGetLatestVersion = "http://schemas.microsoft.com/cdo/configuration/urlgetlatestversion"
CONST cdoURLProxyBypass = "http://schemas.microsoft.com/cdo/configuration/urlproxybypass"
CONST cdoURLProxyServer = "http://schemas.microsoft.com/cdo/configuration/urlproxyserver"
CONST cdoUseMessageResponseText = "http://schemas.microsoft.com/cdo/configuration/usemessageresponsetext"

' // Module: CdoContentTypeValues
' // Number of constants: 11
CONST cdoGif = "image/gif"
CONST cdoJpeg = "image/jpeg"
CONST cdoMessageExternalBody = "message/external-body"
CONST cdoMessagePartial = "message/partial"
CONST cdoMessageRFC822 = "message/rfc822"
CONST cdoMultipartAlternative = "multipart/alternative"
CONST cdoMultipartDigest = "multipart/digest"
CONST cdoMultipartMixed = "multipart/mixed"
CONST cdoMultipartRelated = "multipart/related"
CONST cdoTextHTML = "text/html"
CONST cdoTextPlain = "text/plain"

' // Module: CdoDAV
' // Number of constants: 2
CONST cdoContentClass = "DAV:contentclass"
CONST cdoGetContentType = "DAV:getcontenttype"

' // Module: CdoEncodingType
' // Number of constants: 7
CONST cdo7bit = "7bit"
CONST cdo8bit = "8bit"
CONST cdoBase64 = "base64"
CONST cdoBinary = "binary"
CONST cdoMacBinHex40 = "mac-binhex40"
CONST cdoQuotedPrintable = "quoted-printable"
CONST cdoUuencode = "uuencode"

' // Module: CdoErrors
' // Number of constants: 61
CONST CDO_E_UNCAUGHT_EXCEPTION = -2147220991   ' (&h80040201)
CONST CDO_E_NOT_OPENED = -2147220990   ' (&h80040202)
CONST CDO_E_UNSUPPORTED_DATASOURCE = -2147220989   ' (&h80040203)
CONST CDO_E_INVALID_PROPERTYNAME = -2147220988   ' (&h80040204)
CONST CDO_E_PROP_UNSUPPORTED = -2147220987   ' (&h80040205)
CONST CDO_E_INACTIVE = -2147220986   ' (&h80040206)
CONST CDO_E_NO_SUPPORT_FOR_OBJECTS = -2147220985   ' (&h80040207)
CONST CDO_E_NOT_AVAILABLE = -2147220984   ' (&h80040208)
CONST CDO_E_NO_DEFAULT_DROP_DIR = -2147220983   ' (&h80040209)
CONST CDO_E_SMTP_SERVER_REQUIRED = -2147220982   ' (&h8004020A)
CONST CDO_E_NNTP_SERVER_REQUIRED = -2147220981   ' (&h8004020B)
CONST CDO_E_RECIPIENT_MISSING = -2147220980   ' (&h8004020C)
CONST CDO_E_FROM_MISSING = -2147220979   ' (&h8004020D)
CONST CDO_E_SENDER_REJECTED = -2147220978   ' (&h8004020E)
CONST CDO_E_RECIPIENTS_REJECTED = -2147220977   ' (&h8004020F)
CONST CDO_E_NNTP_POST_FAILED = -2147220976   ' (&h80040210)
CONST CDO_E_SMTP_SEND_FAILED = -2147220975   ' (&h80040211)
CONST CDO_E_CONNECTION_DROPPED = -2147220974   ' (&h80040212)
CONST CDO_E_FAILED_TO_CONNECT = -2147220973   ' (&h80040213)
CONST CDO_E_INVALID_POST = -2147220972   ' (&h80040214)
CONST CDO_E_AUTHENTICATION_FAILURE = -2147220971   ' (&h80040215)
CONST CDO_E_INVALID_CONTENT_TYPE = -2147220970   ' (&h80040216)
CONST CDO_E_LOGON_FAILURE = -2147220969   ' (&h80040217)
CONST CDO_E_HTTP_NOT_FOUND = -2147220968   ' (&h80040218)
CONST CDO_E_HTTP_FORBIDDEN = -2147220967   ' (&h80040219)
CONST CDO_E_HTTP_FAILED = -2147220966   ' (&h8004021A)
CONST CDO_E_MULTIPART_NO_DATA = -2147220965   ' (&h8004021B)
CONST CDO_E_INVALID_ENCODING_FOR_MULTIPART = -2147220964   ' (&h8004021C)
CONST CDO_E_UNSAFE_OPERATION = -2147220963   ' (&h8004021D)
CONST CDO_E_PROP_NOT_FOUND = -2147220962   ' (&h8004021E)
CONST CDO_E_INVALID_SEND_OPTION = -2147220960   ' (&h80040220)
CONST CDO_E_INVALID_POST_OPTION = -2147220959   ' (&h80040221)
CONST CDO_E_NO_PICKUP_DIR = -2147220958   ' (&h80040222)
CONST CDO_E_NOT_ALL_DELETED = -2147220957   ' (&h80040223)
CONST CDO_E_NO_METHOD = -2147220956   ' (&h80040224)
CONST CDO_E_PROP_READONLY = -2147220953   ' (&h80040227)
CONST CDO_E_PROP_CANNOT_DELETE = -2147220952   ' (&h80040228)
CONST CDO_E_BAD_DATA = -2147220951   ' (&h80040229)
CONST CDO_E_PROP_NONHEADER = -2147220950   ' (&h8004022A)
CONST CDO_E_INVALID_CHARSET = -2147220949   ' (&h8004022B)
CONST CDO_E_ADOSTREAM_NOT_BOUND = -2147220948   ' (&h8004022C)
CONST CDO_E_CONTENTPROPXML_NOT_FOUND = -2147220947   ' (&h8004022D)
CONST CDO_E_CONTENTPROPXML_WRONG_CHARSET = -2147220946   ' (&h8004022E)
CONST CDO_E_CONTENTPROPXML_PARSE_FAILED = -2147220945   ' (&h8004022F)
CONST CDO_E_CONTENTPROPXML_CONVERT_FAILED = -2147220944   ' (&h80040230)
CONST CDO_E_NO_DIRECTORIES_SPECIFIED = -2147220943   ' (&h80040231)
CONST CDO_E_DIRECTORIES_UNREACHABLE = -2147220942   ' (&h80040232)
CONST CDO_E_BAD_SENDER = -2147220941   ' (&h80040233)
CONST CDO_E_SELF_BINDING = -2147220940   ' (&h80040234)
CONST CDO_E_BAD_ATTENDEE_DATA = -2147220939   ' (&h80040235)
CONST CDO_E_ROLE_NOMORE_AVAILABLE = -2147220938   ' (&h80040236)
CONST CDO_E_BAD_TASKTYPE_ONASSIGN = -2147220937   ' (&h80040237)
CONST CDO_E_NOT_ASSIGNEDTO_USER = -2147220936   ' (&h80040238)
CONST CDO_E_OUTOFDATE = -2147220935   ' (&h80040239)
CONST CDO_E_ARGUMENT1 = -2147205120   ' (&h80044000)
CONST CDO_E_ARGUMENT2 = -2147205119   ' (&h80044001)
CONST CDO_E_ARGUMENT3 = -2147205118   ' (&h80044002)
CONST CDO_E_ARGUMENT4 = -2147205117   ' (&h80044003)
CONST CDO_E_ARGUMENT5 = -2147205116   ' (&h80044004)
CONST CDO_E_NOT_FOUND = -2146644475   ' (&h800CCE05)
CONST CDO_E_INVALID_ENCODING_TYPE = -2146644451   ' (&h800CCE1D)

' // Module: CdoExchange
' // Number of constants: 1
CONST cdoSensitivity = "http://schemas.microsoft.com/exchange/sensitivity"

' // Module: CdoHTTPMail
' // Number of constants: 19
CONST cdoAttachmentFilename = "urn:schemas:httpmail:attachmentfilename"
CONST cdoBcc = "urn:schemas:httpmail:bcc"
CONST cdoCc = "urn:schemas:httpmail:cc"
CONST cdoContentDispositionType = "urn:schemas:httpmail:content-disposition-type"
CONST cdoContentMediaType = "urn:schemas:httpmail:content-media-type"
CONST cdoDate = "urn:schemas:httpmail:date"
CONST cdoDateReceived = "urn:schemas:httpmail:datereceived"
CONST cdoFrom = "urn:schemas:httpmail:from"
CONST cdoHasAttachment = "urn:schemas:httpmail:hasattachment"
CONST cdoHTMLDescription = "urn:schemas:httpmail:htmldescription"
CONST cdoImportance = "urn:schemas:httpmail:importance"
CONST cdoNormalizedSubject = "urn:schemas:httpmail:normalizedsubject"
CONST cdoPriority = "urn:schemas:httpmail:priority"
CONST cdoReplyTo = "urn:schemas:httpmail:reply-to"
CONST cdoSender = "urn:schemas:httpmail:sender"
CONST cdoSubject = "urn:schemas:httpmail:subject"
CONST cdoTextDescription = "urn:schemas:httpmail:textdescription"
CONST cdoThreadTopic = "urn:schemas:httpmail:thread-topic"
CONST cdoTo = "urn:schemas:httpmail:to"

' // Module: CdoInterfaces
' // Number of constants: 6
CONST cdoAdoStream = "_Stream"
CONST cdoIBodyPart = "IBodyPart"
CONST cdoIConfiguration = "IConfiguration"
CONST cdoIDataSource = "IDataSource"
CONST cdoIMessage = "IMessage"
CONST cdoIStream = "IStream"

' // Module: CdoMailHeader
' // Number of constants: 36
CONST cdoApproved = "urn:schemas:mailheader:approved"
CONST cdoComment = "urn:schemas:mailheader:comment"
CONST cdoContentBase = "urn:schemas:mailheader:content-base"
CONST cdoContentDescription = "urn:schemas:mailheader:content-description"
CONST cdoContentDisposition = "urn:schemas:mailheader:content-disposition"
CONST cdoContentId = "urn:schemas:mailheader:content-id"
CONST cdoContentLanguage = "urn:schemas:mailheader:content-language"
CONST cdoContentLocation = "urn:schemas:mailheader:content-location"
CONST cdoContentTransferEncoding = "urn:schemas:mailheader:content-transfer-encoding"
CONST cdoContentType = "urn:schemas:mailheader:content-type"
CONST cdoControl = "urn:schemas:mailheader:control"
CONST cdoDisposition = "urn:schemas:mailheader:disposition"
CONST cdoDispositionNotificationTo = "urn:schemas:mailheader:disposition-notification-to"
CONST cdoDistribution = "urn:schemas:mailheader:distribution"
CONST cdoExpires = "urn:schemas:mailheader:expires"
CONST cdoFollowupTo = "urn:schemas:mailheader:followup-to"
CONST cdoInReplyTo = "urn:schemas:mailheader:in-reply-to"
CONST cdoLines = "urn:schemas:mailheader:lines"
CONST cdoMessageId = "urn:schemas:mailheader:message-id"
CONST cdoMIMEVersion = "urn:schemas:mailheader:mime-version"
CONST cdoNewsgroups = "urn:schemas:mailheader:newsgroups"
CONST cdoOrganization = "urn:schemas:mailheader:organization"
CONST cdoOriginalRecipient = "urn:schemas:mailheader:original-recipient"
CONST cdoPath = "urn:schemas:mailheader:path"
CONST cdoPostingVersion = "urn:schemas:mailheader:posting-version"
CONST cdoReceived = "urn:schemas:mailheader:received"
CONST cdoReferences = "urn:schemas:mailheader:references"
CONST cdoRelayVersion = "urn:schemas:mailheader:relay-version"
CONST cdoReturnPath = "urn:schemas:mailheader:return-path"
CONST cdoReturnReceiptTo = "urn:schemas:mailheader:return-receipt-to"
CONST cdoSummary = "urn:schemas:mailheader:summary"
CONST cdoThreadIndex = "urn:schemas:mailheader:thread-index"
CONST cdoXMailer = "urn:schemas:mailheader:x-mailer"
CONST cdoXref = "urn:schemas:mailheader:xref"
CONST cdoXUnsent = "urn:schemas:mailheader:x-unsent"
CONST cdoXFidelity = "urn:schemas:mailheader:x-cdostreamhighfidelity"

' // Module: CdoNamespace
' // Number of constants: 6
CONST cdoNSConfiguration = "http://schemas.microsoft.com/cdo/configuration/"
CONST cdoNSContacts = "urn:schemas:contacts:"
CONST cdoNSHTTPMail = "urn:schemas:httpmail:"
CONST cdoNSMailHeader = "urn:schemas:mailheader:"
CONST cdoNSNNTPEnvelope = "http://schemas.microsoft.com/cdo/nntpenvelope/"
CONST cdoNSSMTPEnvelope = "http://schemas.microsoft.com/cdo/smtpenvelope/"

' // Module: CdoNNTPEnvelope
' // Number of constants: 2
CONST cdoNewsgroupList = "http://schemas.microsoft.com/cdo/nntpenvelope/newsgrouplist"
CONST cdoNNTPProcessing = "http://schemas.microsoft.com/cdo/nntpenvelope/nntpprocessing"

' // Module: CdoOffice
' // Number of constants: 1
CONST cdoKeywords = "urn:schemas-microsoft-com:office:office#Keywords"

' // Module: CdoSMTPEnvelope
' // Number of constants: 6
CONST cdoArrivalTime = "http://schemas.microsoft.com/cdo/smtpenvelope/arrivaltime"
CONST cdoClientIPAddress = "http://schemas.microsoft.com/cdo/smtpenvelope/clientipaddress"
CONST cdoMessageStatus = "http://schemas.microsoft.com/cdo/smtpenvelope/messagestatus"
CONST cdoPickupFileName = "http://schemas.microsoft.com/cdo/smtpenvelope/pickupfilename"
CONST cdoRecipientList = "http://schemas.microsoft.com/cdo/smtpenvelope/recipientlist"
CONST cdoSenderEmailAddress = "http://schemas.microsoft.com/cdo/smtpenvelope/senderemailaddress"

#ifndef __Afx_IUnknown_INTERFACE_DEFINED__
#define __Afx_IUnknown_INTERFACE_DEFINED__
TYPE Afx_IUnknown AS Afx_IUnknown_
TYPE Afx_IUnknown_ EXTENDS OBJECT
	DECLARE ABSTRACT FUNCTION QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObject AS LPVOID PTR) AS HRESULT
	DECLARE ABSTRACT FUNCTION AddRef () AS ULONG
	DECLARE ABSTRACT FUNCTION Release () AS ULONG
END TYPE
TYPE AFX_LPUNKNOWN AS Afx_IUnknown PTR
#endif

#ifndef __Afx_IDispatch_INTERFACE_DEFINED__
#define __Afx_IDispatch_INTERFACE_DEFINED__
TYPE Afx_IDispatch AS Afx_IDispatch_
TYPE Afx_IDispatch_  EXTENDS Afx_Iunknown
   DECLARE ABSTRACT FUNCTION GetTypeInfoCount (BYVAL pctinfo AS UINT PTR) as HRESULT
   DECLARE ABSTRACT FUNCTION GetTypeInfo (BYVAL iTInfo AS UINT, BYVAL lcid AS LCID, BYVAL ppTInfo AS ITypeInfo PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetIDsOfNames (BYVAL riid AS CONST IID CONST PTR, BYVAL rgszNames AS LPOLESTR PTR, BYVAL cNames AS UINT, BYVAL lcid AS LCID, BYVAL rgDispId AS DISPID PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL dispIdMember AS DISPID, BYVAL riid AS CONST IID CONST PTR, BYVAL lcid AS LCID, BYVAL wFlags AS WORD, BYVAL pDispParams AS DISPPARAMS PTR, BYVAL pVarResult AS VARIANT PTR, BYVAL pExcepInfo AS EXCEPINFO PTR, BYVAL puArgErr AS UINT PTR) AS HRESULT
END TYPE
TYPE AFX_LPDISPATCH AS Afx_IDispatch PTR
#endif

' // Dual interfaces - Forward references
TYPE Afx_IBodyPart AS Afx_IBodyPart_
TYPE Afx_IBodyParts AS Afx_IBodyParts_
TYPE Afx_IConfiguration AS Afx_IConfiguration_
TYPE Afx_IDataSource AS Afx_IDataSource_
TYPE Afx_IDropDirectory AS Afx_IDropDirectory_
TYPE Afx_IMessage AS Afx_IMessage_
TYPE Afx_IMessages AS Afx_IMessages_
TYPE Afx_INNTPEarlyScriptConnector AS Afx_INNTPEarlyScriptConnector_
TYPE Afx_INNTPFinalScriptConnector AS Afx_INNTPFinalScriptConnector_
TYPE Afx_INNTPOnPost AS Afx_INNTPOnPost_
TYPE Afx_INNTPOnPostEarly AS Afx_INNTPOnPostEarly_
TYPE Afx_INNTPOnPostFinal AS Afx_INNTPOnPostFinal_
TYPE Afx_INNTPPostScriptConnector AS Afx_INNTPPostScriptConnector_
TYPE Afx_ISMTPOnArrival AS Afx_ISMTPOnArrival_
TYPE Afx_ISMTPScriptConnector AS Afx_ISMTPScriptConnector_

' // Dual interfaces

' ########################################################################################
' Interface name: IBodyPart
' IID: {CD000021-8B95-11D1-82DB-00C04FB1625D}
' Documentation string: Defines methods, properties, and collections used to manage a message body part.
' Attributes =  4544 [&h000011C0] [Dual] [Nonextensible] [Oleautomation] [Dispatchable]
' Inherited interface = IDispatch
' Number of methods = 22
' ########################################################################################

#ifndef __Afx_IBodyPart_INTERFACE_DEFINED__
#define __Afx_IBodyPart_INTERFACE_DEFINED__

TYPE Afx_IBodyPart_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION get_BodyParts (BYVAL varBodyParts AS Afx_IBodyParts PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ContentTransferEncoding (BYVAL pContentTransferEncoding AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_ContentTransferEncoding (BYVAL pContentTransferEncoding AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ContentMediaType (BYVAL pContentMediaType AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_ContentMediaType (BYVAL pContentMediaType AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Fields (BYVAL varFields AS Afx_AdoFields PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Charset (BYVAL pCharset AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Charset (BYVAL pCharset AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_FileName (BYVAL varFileName AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_DataSource (BYVAL varDataSource AS Afx_IDataSource PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ContentClass (BYVAL pContentClass AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_ContentClass (BYVAL pContentClass AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ContentClassName (BYVAL pContentClassName AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_ContentClassName (BYVAL pContentClassName AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Parent (BYVAL varParent AS Afx_IBodyPart PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION AddBodyPart (BYVAL Index AS LONG = -1, BYVAL ppPart AS Afx_IBodyPart PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION SaveToFile (BYVAL FileName AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetEncodedContentStream (BYVAL ppStream AS Afx_AdoStream PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetDecodedContentStream (BYVAL ppStream AS Afx_AdoStream PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetStream (BYVAL ppStream AS Afx_AdoStream PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetFieldParameter (BYVAL FieldName AS AFX_BSTR, BYVAL Parameter AS AFX_BSTR, BYVAL pbstrValue AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetInterface (BYVAL Interface AS AFX_BSTR, BYVAL ppUnknown AS Afx_IDispatch PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: IBodyParts
' IID: {CD000023-8B95-11D1-82DB-00C04FB1625D}
' Documentation string: Defines methods and properties used to manage a collection of BodyPart objects.
' Attributes =  4544 [&h000011C0] [Dual] [Nonextensible] [Oleautomation] [Dispatchable]
' Inherited interface = IDispatch
' Number of methods = 6
' ########################################################################################

#ifndef __Afx_IBodyParts_INTERFACE_DEFINED__
#define __Afx_IBodyParts_INTERFACE_DEFINED__

TYPE Afx_IBodyParts_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION get_Count (BYVAL varCount AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Item (BYVAL Index AS LONG, BYVAL ppBody AS Afx_IBodyPart PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get__NewEnum (BYVAL retval AS Afx_IUnknown PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Delete_ (BYVAL varBP AS VARIANT) AS HRESULT
   DECLARE ABSTRACT FUNCTION DeleteAll () AS HRESULT
   DECLARE ABSTRACT FUNCTION Add (BYVAL Index AS LONG = -1, BYVAL ppPart AS Afx_IBodyPart PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: IConfiguration
' IID: {CD000022-8B95-11D1-82DB-00C04FB1625D}
' Documentation string: Defines methods, properties, and collections used to manage configuration information for CDO objects.
' Attributes =  4544 [&h000011C0] [Dual] [Nonextensible] [Oleautomation] [Dispatchable]
' Inherited interface = IDispatch
' Number of methods = 3
' ########################################################################################

#ifndef __Afx_IConfiguration_INTERFACE_DEFINED__
#define __Afx_IConfiguration_INTERFACE_DEFINED__

TYPE Afx_IConfiguration_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION get_Fields (BYVAL varFields AS Afx_AdoFields PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Load (BYVAL LoadFrom AS CdoConfigSource, BYVAL URL AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetInterface (BYVAL Interface AS AFX_BSTR, BYVAL ppUnknown AS Afx_IDispatch PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: IDataSource
' IID: {CD000029-8B95-11D1-82DB-00C04FB1625D}
' Documentation string: Defines methods, properties, and collections used to extract messages from or embed messages into other CDO message body parts.
' Attributes =  4544 [&h000011C0] [Dual] [Nonextensible] [Oleautomation] [Dispatchable]
' Inherited interface = IDispatch
' Number of methods = 12
' ########################################################################################

#ifndef __Afx_IDataSource_INTERFACE_DEFINED__
#define __Afx_IDataSource_INTERFACE_DEFINED__

TYPE Afx_IDataSource_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION get_SourceClass (BYVAL varSourceClass AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Source (BYVAL varSource AS Afx_IUnknown PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_IsDirty (BYVAL pIsDirty AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_IsDirty (BYVAL pIsDirty AS VARIANT_BOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_SourceURL (BYVAL varSourceURL AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ActiveConnection (BYVAL varActiveConnection AS Afx_AdoConnection PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION SaveToObject (BYVAL Source AS Afx_IUnknown PTR, BYVAL InterfaceName AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION OpenObject (BYVAL Source AS Afx_IUnknown PTR, BYVAL InterfaceName AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION SaveTo (BYVAL SourceURL AS AFX_BSTR, BYVAL ActiveConnection AS Afx_IDispatch PTR, BYVAL Mode AS ConnectModeEnum, BYVAL CreateOptions AS RecordCreateOptionsEnum, BYVAL Options AS RecordOpenOptionsEnum, BYVAL UserName AS AFX_BSTR, BYVAL Password AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Open (BYVAL SourceURL AS AFX_BSTR, BYVAL ActiveConnection AS Afx_IDispatch PTR, BYVAL Mode AS ConnectModeEnum, BYVAL CreateOptions AS RecordCreateOptionsEnum = -1, BYVAL Options AS RecordOpenOptionsEnum, BYVAL UserName AS AFX_BSTR, BYVAL Password AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Save () AS HRESULT
   DECLARE ABSTRACT FUNCTION SaveToContainer (BYVAL ContainerURL AS AFX_BSTR, BYVAL ActiveConnection AS Afx_IDispatch PTR, BYVAL Mode AS ConnectModeEnum, BYVAL CreateOptions AS RecordCreateOptionsEnum, BYVAL Options AS RecordOpenOptionsEnum, BYVAL UserName AS AFX_BSTR, BYVAL Password AS AFX_BSTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: IDropDirectory
' IID: {CD000024-8B95-11D1-82DB-00C04FB1625D}
' Documentation string: Defines methods, properties, and collections used to manage a collection of messages on the file system.
' Attributes =  4544 [&h000011C0] [Dual] [Nonextensible] [Oleautomation] [Dispatchable]
' Inherited interface = IDispatch
' Number of methods = 1
' ########################################################################################

#ifndef __Afx_IDropDirectory_INTERFACE_DEFINED__
#define __Afx_IDropDirectory_INTERFACE_DEFINED__

TYPE Afx_IDropDirectory_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION GetMessages (BYVAL DirName AS AFX_BSTR, BYVAL Msgs AS Afx_IMessages PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: IMessage
' IID: {CD000020-8B95-11D1-82DB-00C04FB1625D}
' Documentation string: Defines methods, properties, and collections used to manage a message.
' Attributes =  4544 [&h000011C0] [Dual] [Nonextensible] [Oleautomation] [Dispatchable]
' Inherited interface = IDispatch
' Number of methods = 57
' ########################################################################################

#ifndef __Afx_IMessage_INTERFACE_DEFINED__
#define __Afx_IMessage_INTERFACE_DEFINED__

TYPE Afx_IMessage_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION get_BCC (BYVAL pBCC AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_BCC (BYVAL pBCC AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_CC (BYVAL pCC AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_CC (BYVAL pCC AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_FollowUpTo (BYVAL pFollowUpTo AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_FollowUpTo (BYVAL pFollowUpTo AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_From (BYVAL pFrom AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_From (BYVAL pFrom AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Keywords (BYVAL pKeywords AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Keywords (BYVAL pKeywords AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_MimeFormatted (BYVAL pMimeFormatted AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_MimeFormatted (BYVAL pMimeFormatted AS VARIANT_BOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Newsgroups (BYVAL pNewsgroups AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Newsgroups (BYVAL pNewsgroups AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Organization (BYVAL pOrganization AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Organization (BYVAL pOrganization AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ReceivedTime (BYVAL varReceivedTime AS DATE_ PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ReplyTo (BYVAL pReplyTo AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_ReplyTo (BYVAL pReplyTo AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_DSNOptions (BYVAL pDSNOptions AS CdoDSNOptions PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_DSNOptions (BYVAL pDSNOptions AS CdoDSNOptions) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_SentOn (BYVAL varSentOn AS DATE_ PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Subject (BYVAL pSubject AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Subject (BYVAL pSubject AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_To (BYVAL pTo AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_To (BYVAL pTo AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_TextBody (BYVAL pTextBody AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_TextBody (BYVAL pTextBody AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_HTMLBody (BYVAL pHTMLBody AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_HTMLBody (BYVAL pHTMLBody AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Attachments (BYVAL varAttachments AS Afx_IBodyParts PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Sender (BYVAL pSender AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Sender (BYVAL pSender AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Configuration (BYVAL pConfiguration AS Afx_IConfiguration PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Configuration (BYVAL pConfiguration AS Afx_IConfiguration PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION putref_Configuration (BYVAL pConfiguration AS Afx_IConfiguration PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_AutoGenerateTextBody (BYVAL pAutoGenerateTextBody AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_AutoGenerateTextBody (BYVAL pAutoGenerateTextBody AS VARIANT_BOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_EnvelopeFields (BYVAL varEnvelopeFields AS Afx_AdoFields PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_TextBodyPart (BYVAL varTextBodyPart AS Afx_IBodyPart PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_HTMLBodyPart (BYVAL varHTMLBodyPart AS Afx_IBodyPart PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_BodyPart (BYVAL varBodyPart AS Afx_IBodyPart PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_DataSource (BYVAL varDataSource AS Afx_IDataSource PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Fields (BYVAL varFields AS Afx_AdoFields PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_MDNRequested (BYVAL pMDNRequested AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_MDNRequested (BYVAL pMDNRequested AS VARIANT_BOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION AddRelatedBodyPart (BYVAL URL AS AFX_BSTR, BYVAL Reference AS AFX_BSTR, BYVAL ReferenceType AS CdoReferenceType, BYVAL UserName AS AFX_BSTR, BYVAL Password AS AFX_BSTR, BYVAL ppBody AS Afx_IBodyPart PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION AddAttachment (BYVAL URL AS AFX_BSTR, BYVAL UserName AS AFX_BSTR = NULL, BYVAL Password AS AFX_BSTR = NULL, BYVAL ppBody AS Afx_IBodyPart PTR PTR = NULL) AS HRESULT
   DECLARE ABSTRACT FUNCTION CreateMHTMLBody (BYVAL URL AS AFX_BSTR, BYVAL Flags AS CdoMHTMLFlags = 0, BYVAL UserName AS AFX_BSTR = NULL, BYVAL Password AS AFX_BSTR = NULL) AS HRESULT
   DECLARE ABSTRACT FUNCTION Forward (BYVAL ppMsg AS Afx_IMessage PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Post () AS HRESULT
   DECLARE ABSTRACT FUNCTION PostReply (BYVAL ppMsg AS Afx_IMessage PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Reply (BYVAL ppMsg AS Afx_IMessage PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION ReplyAll (BYVAL ppMsg AS Afx_IMessage PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Send () AS HRESULT
   DECLARE ABSTRACT FUNCTION GetStream (BYVAL ppStream AS Afx_AdoStream PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetInterface (BYVAL Interface AS AFX_BSTR, BYVAL ppUnknown AS Afx_IDispatch PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: IMessages
' IID: {CD000025-8B95-11D1-82DB-00C04FB1625D}
' Documentation string: Defines methods and properties used to manage a collection of Message objects on the file system. Returned by Afx_IDropDirectory.GetMessages.
' Attributes =  4544 [&h000011C0] [Dual] [Nonextensible] [Oleautomation] [Dispatchable]
' Inherited interface = IDispatch
' Number of methods = 6
' ########################################################################################

#ifndef __Afx_IMessages_INTERFACE_DEFINED__
#define __Afx_IMessages_INTERFACE_DEFINED__

TYPE Afx_IMessages_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION get_Item (BYVAL Index AS LONG, BYVAL ppMessage AS Afx_IMessage PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Count (BYVAL varCount AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Delete_ (BYVAL Index AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION DeleteAll () AS HRESULT
   DECLARE ABSTRACT FUNCTION get__NewEnum (BYVAL retval AS Afx_IUnknown PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_FileName (BYVAL var AS VARIANT, BYVAL FileName AS AFX_BSTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: INNTPEarlyScriptConnector
' IID: {CD000034-8B95-11D1-82DB-00C04FB1625D}
' Documentation string: Afx_INNTPFinalScriptConnector interface
' Attributes =  4560 [&h000011D0] [Hidden] [Dual] [Nonextensible] [Oleautomation] [Dispatchable]
' Inherited interface = IDispatch
' Number of methods = 0
' ########################################################################################

'#ifndef __Afx_INNTPEarlyScriptConnector_INTERFACE_DEFINED__
'#define __Afx_INNTPEarlyScriptConnector_INTERFACE_DEFINED__

'#endif

' ########################################################################################

' ########################################################################################
' Interface name: INNTPFinalScriptConnector
' IID: {CD000032-8B95-11D1-82DB-00C04FB1625D}
' Documentation string: Afx_INNTPFinalScriptConnector interface
' Attributes =  4560 [&h000011D0] [Hidden] [Dual] [Nonextensible] [Oleautomation] [Dispatchable]
' Inherited interface = IDispatch
' Number of methods = 0
' ########################################################################################

'#ifndef __Afx_INNTPFinalScriptConnector_INTERFACE_DEFINED__
'#define __Afx_INNTPFinalScriptConnector_INTERFACE_DEFINED__

'#endif

' ########################################################################################

' ########################################################################################
' Interface name: INNTPOnPost
' IID: {CD000027-8B95-11D1-82DB-00C04FB1625D}
' Documentation string: Implement when creating NNTP OnPost event sinks.
' Attributes =  4544 [&h000011C0] [Dual] [Nonextensible] [Oleautomation] [Dispatchable]
' Inherited interface = IDispatch
' Number of methods = 1
' ########################################################################################

'#ifndef __Afx_INNTPOnPost_INTERFACE_DEFINED__
'#define __Afx_INNTPOnPost_INTERFACE_DEFINED__

'TYPE Afx_INNTPOnPost_ EXTENDS Afx_IDispatch
'   DECLARE ABSTRACT FUNCTION OnPost (BYVAL Msg AS Afx_IMessage PTR, BYVAL EventStatus AS CdoEventStatus PTR) AS HRESULT
'END TYPE

'#endif

' ########################################################################################

' ########################################################################################
' Interface name: INNTPOnPostEarly
' IID: {CD000033-8B95-11D1-82DB-00C04FB1625D}
' Documentation string: Implement when creating NNTP OnPostEarly event sinks.
' Attributes =  4544 [&h000011C0] [Dual] [Nonextensible] [Oleautomation] [Dispatchable]
' Inherited interface = IDispatch
' Number of methods = 1
' ########################################################################################

'#ifndef __Afx_INNTPOnPostEarly_INTERFACE_DEFINED__
'#define __Afx_INNTPOnPostEarly_INTERFACE_DEFINED__

'TYPE Afx_INNTPOnPostEarly_ EXTENDS Afx_IDispatch
'   DECLARE ABSTRACT FUNCTION OnPostEarly (BYVAL Msg AS Afx_IMessage PTR, BYVAL EventStatus AS CdoEventStatus PTR) AS HRESULT
'END TYPE

'#endif

' ########################################################################################

' ########################################################################################
' Interface name: INNTPOnPostFinal
' IID: {CD000028-8B95-11D1-82DB-00C04FB1625D}
' Documentation string: Implement when creating NNTP OnPostFinal event sinks.
' Attributes =  4544 [&h000011C0] [Dual] [Nonextensible] [Oleautomation] [Dispatchable]
' Inherited interface = IDispatch
' Number of methods = 1
' ########################################################################################

'#ifndef __Afx_INNTPOnPostFinal_INTERFACE_DEFINED__
'#define __Afx_INNTPOnPostFinal_INTERFACE_DEFINED__

'TYPE Afx_INNTPOnPostFinal_ EXTENDS Afx_IDispatch
'   DECLARE ABSTRACT FUNCTION OnPostFinal (BYVAL Msg AS Afx_IMessage PTR, BYVAL EventStatus AS CdoEventStatus PTR) AS HRESULT
'END TYPE

'#endif

' ########################################################################################

' ########################################################################################
' Interface name: INNTPPostScriptConnector
' IID: {CD000031-8B95-11D1-82DB-00C04FB1625D}
' Documentation string: Afx_INNTPPostScriptConnector interface
' Attributes =  4560 [&h000011D0] [Hidden] [Dual] [Nonextensible] [Oleautomation] [Dispatchable]
' Inherited interface = IDispatch
' Number of methods = 0
' ########################################################################################

'#ifndef __Afx_INNTPPostScriptConnector_INTERFACE_DEFINED__
'#define __Afx_INNTPPostScriptConnector_INTERFACE_DEFINED__

'#endif

' ########################################################################################

' ########################################################################################
' Interface name: ISMTPOnArrival
' IID: {CD000026-8B95-11D1-82DB-00C04FB1625D}
' Documentation string: Implement when creating SMTP OnArrival event sinks.
' Attributes =  4544 [&h000011C0] [Dual] [Nonextensible] [Oleautomation] [Dispatchable]
' Inherited interface = IDispatch
' Number of methods = 1
' ########################################################################################

'#ifndef __Afx_ISMTPOnArrival_INTERFACE_DEFINED__
'#define __Afx_ISMTPOnArrival_INTERFACE_DEFINED__

'TYPE Afx_ISMTPOnArrival_ EXTENDS Afx_IDispatch
'   DECLARE ABSTRACT FUNCTION OnArrival (BYVAL Msg AS Afx_IMessage PTR, BYVAL EventStatus AS CdoEventStatus PTR) AS HRESULT
'END TYPE

'#endif

' ########################################################################################

' ########################################################################################
' Interface name: ISMTPScriptConnector
' IID: {CD000030-8B95-11D1-82DB-00C04FB1625D}
' Documentation string: Afx_ISMTPScriptConnector interface
' Attributes =  4560 [&h000011D0] [Hidden] [Dual] [Nonextensible] [Oleautomation] [Dispatchable]
' Inherited interface = IDispatch
' Number of methods = 0
' ########################################################################################

'#ifndef __Afx_ISMTPScriptConnector_INTERFACE_DEFINED__
'#define __Afx_ISMTPScriptConnector_INTERFACE_DEFINED__

'#endif

' ########################################################################################

END NAMESPACE
